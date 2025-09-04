from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import os
import re
from datetime import datetime
import warnings
import io
import logging
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# OpenAI imports (replacing LangChain imports)
from openai import OpenAI

warnings.filterwarnings('ignore')

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

app = Flask(__name__)

# File paths
OPEN_TICKETS_PATH = r"C:\Users\RitabrataRoyChoudhur\OneDrive - GyanSys Inc\Desktop\Python\ServiceNow\open.xlsx"
CLOSED_TICKETS_PATH = r"C:\Users\RitabrataRoyChoudhur\OneDrive - GyanSys Inc\Desktop\Python\ServiceNow\close.xlsx"
DEFAULT_THRESHOLD = 0.75

# Updated required columns to include 'Assigned To' and 'Close Notes'
REQUIRED_COLUMNS = ['Number', 'Short Description', 'Assignment group', 'Customer', 'Created', 'Assigned to']
PREFERRED_COLUMNS = REQUIRED_COLUMNS + ['Close Notes', 'Resolved by']

# Initialize OpenAI client globally
try:
    # Initialize OpenAI client - you can set API key in environment variable or directly here
    openai_client = OpenAI(
        api_key=os.getenv("OPENAI_API_KEY", "#API_KEY#")
    )
    llm = openai_client  # Keep the same variable name for compatibility
    logging.info("OpenAI GPT initialized successfully")
except Exception as e:
    logging.error(f"Failed to initialize OpenAI GPT: {str(e)}")
    llm = None
    openai_client = None

# Updated template for generating root cause and suggested fixes separately
analysis_template = """You are a highly experienced SAP IT support specialist specializing in analyzing closed support tickets to provide actionable solutions for similar open tickets.

Using your deep knowledge of SAP systems, modules, and official SAP Notes, consider the following inputs:

Closing Notes from Similar Tickets:
{closing_notes}

Open Ticket Description:
{open_ticket_description}

Please analyze the above information and provide a structured response with TWO distinct sections:

## ROOT CAUSE ANALYSIS:
Identify and explain the likely root causes based on patterns in the closing notes and your SAP expertise. Focus on:
- Technical reasons for the issue
- System configuration problems
- User access or permission issues
- Data inconsistencies
- Integration problems
- Performance bottlenecks

## SUGGESTED RESOLUTION:
Provide a step-by-step solution that an SAP support technician can follow, including:
1. Immediate troubleshooting steps
2. Configuration changes needed
3. Relevant SAP transaction codes and tools
4. SAP Notes or documentation references
5. Testing procedures
6. Preventive measures for future occurrences
7. Escalation procedures if the steps don't resolve the issue

Format your response clearly with these two sections. Use professional technical language suitable for SAP support staff.

If the closing notes lack sufficient detail, provide generalized SAP troubleshooting guidance applicable to this type of issue."""

def clean_value(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def preprocess_text(text):
    """Enhanced text preprocessing"""
    if pd.isna(text) or text == "":
        return ""
    
    text = str(text).lower()
    # Remove special characters but keep alphanumeric and spaces
    text = re.sub(r'[^a-zA-Z0-9\s]', ' ', text)
    # Remove extra whitespace
    text = re.sub(r'\s+', ' ', text)
    # Remove common ServiceNow terms that don't add semantic value
    common_terms = ['ticket', 'issue', 'problem', 'error', 'failed', 'failure']
    words = text.split()
    words = [word for word in words if word not in common_terms or len(words) <= 3]
    
    return ' '.join(words).strip()

def generate_root_cause_and_fix_with_openai(similar_closed_tickets, open_ticket_description=""):
    """
    Generate separate root cause analysis and suggested fix for open ticket 
    using OpenAI GPT instead of Ollama
    """
    global openai_client
    
    try:
        closing_notes = []
        
        # Extract closing notes from similar tickets
        for ticket in similar_closed_tickets:
            closing_note = ticket.get('Close Notes', '').strip()
            if closing_note and closing_note.lower() not in ['n/a', 'na', '', 'none', 'null', 'not applicable']:
                closing_notes.append(f"• {closing_note}")
        
        # If no meaningful closing notes found
        if not closing_notes:
            return {
                'root_cause': "No meaningful closing notes available from similar tickets for root cause analysis.",
                'suggested_fix': "Please investigate manually or contact the appropriate support team for detailed analysis."
            }
        
        # If only one closing note or OpenAI client is not available, return basic format
        if len(closing_notes) == 1 or openai_client is None:
            basic_notes = "\n".join(closing_notes)
            return {
                'root_cause': f"Based on similar resolved ticket:\n\n{basic_notes}\n\n(Manual analysis required for detailed root cause identification)",
                'suggested_fix': f"Follow the resolution approach from similar ticket:\n\n{basic_notes}" + 
                               ("\n\n(Note: AI enhancement unavailable - OpenAI client not initialized)" if openai_client is None else "")
            }
        
        # Use OpenAI GPT to generate enhanced analysis when multiple closing notes exist
        closing_notes_text = "\n".join(closing_notes)
        
        # Create the prompt
        prompt = analysis_template.format(
            closing_notes=closing_notes_text,
            open_ticket_description=open_ticket_description
        )
        
        logging.info(f"Generating OpenAI GPT-enhanced analysis for ticket with {len(closing_notes)} similar closing notes")
        
        # Generate enhanced analysis using OpenAI
        response = openai_client.chat.completions.create(
            model="gpt-4o-mini",  # You can also use "gpt-4o" or "gpt-3.5-turbo"
            messages=[
                {
                    "role": "system", 
                    "content": "You are an expert SAP IT support specialist. Provide structured, professional analysis of support tickets."
                },
                {
                    "role": "user", 
                    "content": prompt
                }
            ],
            max_tokens=1500,
            temperature=0.3  # Lower temperature for more consistent technical responses
        )
        
        # Extract and clean the response
        analysis_content = response.choices[0].message.content.strip()
        
        # Split the response into root cause and suggested fix
        root_cause = ""
        suggested_fix = ""
        
        if "## ROOT CAUSE ANALYSIS:" in analysis_content and "## SUGGESTED RESOLUTION:" in analysis_content:
            parts = analysis_content.split("## SUGGESTED RESOLUTION:")
            root_cause_part = parts[0].replace("## ROOT CAUSE ANALYSIS:", "").strip()
            suggested_fix_part = parts[1].strip()
            
            root_cause = root_cause_part if root_cause_part else "Root cause analysis not clearly identified in the response."
            suggested_fix = suggested_fix_part if suggested_fix_part else "Suggested resolution not clearly provided in the response."
        else:
            # If the format is not as expected, treat the whole response as suggested fix
            suggested_fix = analysis_content
            root_cause = "Root cause analysis requires further investigation based on the available information."
        
        logging.info("OpenAI GPT-enhanced analysis generated successfully")
        return {
            'root_cause': root_cause,
            'suggested_fix': suggested_fix
        }
        
    except Exception as e:
        logging.error(f"Error generating OpenAI GPT-enhanced analysis: {str(e)}")
        
        # Fallback to original method if OpenAI fails
        closing_notes = []
        for ticket in similar_closed_tickets:
            closing_note = ticket.get('Close Notes', '').strip()
            if closing_note and closing_note.lower() not in ['n/a', 'na', '']:
                closing_notes.append(f"• {closing_note}")
        
        if closing_notes:
            fallback_notes = "\n".join(closing_notes)
            return {
                'root_cause': f"Based on similar resolved tickets (OpenAI analysis failed):\n\n{fallback_notes}\n\nError: {str(e)}",
                'suggested_fix': f"Follow resolution approach from similar tickets:\n\n{fallback_notes}\n\n(Note: AI enhancement failed)"
            }
        else:
            return {
                'root_cause': "No closing notes available for root cause analysis",
                'suggested_fix': "No resolution guidance available from similar tickets"
            }

# Keep the original function name for backward compatibility
def generate_root_cause_and_fix(similar_closed_tickets, open_ticket_description=""):
    """Wrapper function to maintain compatibility"""
    return generate_root_cause_and_fix_with_openai(similar_closed_tickets, open_ticket_description)

def calculate_semantic_similarity(open_tickets_df, closed_tickets_df, assignment_group_filter=None, threshold=0.75):
    """
    Find similar tickets between open and closed tickets using TF-IDF + Cosine similarity
    with Assignment Group matching and semantic understanding
    """
    print("Starting semantic similarity analysis...")
    
    # Validate required columns
    for col in ['Short Description', 'Assignment group']:
        if col not in open_tickets_df.columns:
            raise ValueError(f"'{col}' column not found in open tickets")
        if col not in closed_tickets_df.columns:
            raise ValueError(f"'{col}' column not found in closed tickets")
    
    # Filter by Assignment Group if specified (only filter open tickets)
    open_df = open_tickets_df.copy()
    closed_df = closed_tickets_df.copy()
    
    if assignment_group_filter and assignment_group_filter != 'All':
        open_df = open_df[open_df['Assignment group'] == assignment_group_filter].reset_index(drop=True)
        
        if open_df.empty:
            print(f"No open tickets found with 'Assignment group' = '{assignment_group_filter}'")
            return []
    
    print(f"Open tickets count: {len(open_df)}")
    print(f"Closed tickets count: {len(closed_df)}")
    
    # Preprocess texts (no translation needed)
    open_texts = [preprocess_text(desc) for desc in open_df['Short Description'].tolist()]
    closed_texts = [preprocess_text(desc) for desc in closed_df['Short Description'].tolist()]
    
    # Combine all texts for TF-IDF fitting
    all_texts = open_texts + closed_texts
    
    print("Calculating TF-IDF vectors...")
    # Use TF-IDF with n-grams for better semantic understanding
    vectorizer = TfidfVectorizer(
        stop_words='english',
        ngram_range=(1, 3),  # Include trigrams for better context
        max_features=5000,   # Limit features for performance
        min_df=1,           # Minimum document frequency
        max_df=0.95         # Maximum document frequency to filter common terms
    )
    
    try:
        tfidf_matrix = vectorizer.fit_transform(all_texts)
        open_vectors = tfidf_matrix[:len(open_texts)]
        closed_vectors = tfidf_matrix[len(open_texts):]
    except ValueError:
        print("Error in TF-IDF vectorization")
        return []
    
    print("Calculating similarity matrix...")
    # Calculate similarity between open and closed tickets
    similarity_matrix = cosine_similarity(open_vectors, closed_vectors)
    
    results = []
    
    for i, open_ticket_row in open_df.iterrows():
        open_assignment_group = clean_value(open_ticket_row['Assignment group'])
        
        # Find similar closed tickets with same assignment group
        similar_closed = []
        
        for j, closed_ticket_row in closed_df.iterrows():
            closed_assignment_group = clean_value(closed_ticket_row['Assignment group'])
            
            # Check if Assignment Group matches (case-insensitive)
            if open_assignment_group.lower() != closed_assignment_group.lower():
                continue
            
            # Get similarity score
            similarity_score = similarity_matrix[i][j]
            
            if similarity_score >= threshold:
                # Prepare closed ticket data
                closed_ticket_data = {}
                for col in closed_df.columns:
                    closed_ticket_data[col] = clean_value(closed_ticket_row[col])
                
                closed_ticket_data['Similarity Score'] = round(similarity_score, 3)
                closed_ticket_data['Similarity Percentage'] = f"{similarity_score * 100:.1f}%"
                
                similar_closed.append(closed_ticket_data)
        
        if similar_closed:
            # Sort by similarity score (highest first)
            similar_closed.sort(key=lambda x: x['Similarity Score'], reverse=True)
            
            # Prepare open ticket data
            open_ticket_data = {}
            for col in open_df.columns:
                open_ticket_data[col] = clean_value(open_ticket_row[col])
            
            # Generate root cause and suggested fix with OpenAI GPT enhancement
            open_description = clean_value(open_ticket_row.get('Short Description', ''))
            analysis_result = generate_root_cause_and_fix(similar_closed, open_description)
            
            results.append({
                'open_ticket': open_ticket_data,
                'similar_closed_tickets': similar_closed,
                'best_similarity_score': similar_closed[0]['Similarity Score'],
                'total_similar_closed': len(similar_closed),
                'root_cause': analysis_result['root_cause'],
                'suggested_fix': analysis_result['suggested_fix']
            })
    
    # Sort results by best similarity score
    results.sort(key=lambda x: x['best_similarity_score'], reverse=True)
    
    print(f"Found {len(results)} open tickets with similar closed tickets")
    return results

def get_open_assignment_groups(open_df):
    """Get unique Assignment Groups from open tickets only"""
    if 'Assignment group' in open_df.columns:
        groups = open_df['Assignment group'].dropna().unique().tolist()
        groups = [str(group).strip() for group in groups if str(group).strip()]
        groups.sort()
        return groups
    return []

def validate_columns(df, dataset_name):
    """Validate that required columns are present"""
    missing_required = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    missing_preferred = [col for col in PREFERRED_COLUMNS if col not in df.columns]
    
    return {
        'missing_required': missing_required,
        'missing_preferred': missing_preferred,
        'has_all_required': len(missing_required) == 0
    }

def create_excel_export(results, analysis_params):
    """
    Create Excel file with analysis results including separate root cause and suggested fix
    """
    wb = Workbook()
    
    # Remove default worksheet
    wb.remove(wb.active)
    
    # Create Summary sheet
    summary_ws = wb.create_sheet("Summary")
    
    # Define styles
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    open_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    closed_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Summary data
    summary_data = [
        ["Analysis Summary", ""],
        ["Assignment Group Filter", analysis_params.get('assignment_group_filter', 'All')],
        ["Similarity Threshold", analysis_params.get('threshold_percentage', 'N/A')],
        ["Total Open Tickets Analyzed", analysis_params.get('filtered_open_tickets', 0)],
        ["Open Tickets with Matches", analysis_params.get('open_tickets_with_matches', 0)],
        ["Total Similar Closed Tickets Found", analysis_params.get('total_matches', 0)],
        ["LLM Status", "OpenAI GPT Available" if llm is not None else "Unavailable"],
        ["Analysis Date", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
    ]
    
    # Write summary
    for row_idx, (key, value) in enumerate(summary_data, 1):
        cell_key = summary_ws.cell(row=row_idx, column=1, value=key)
        cell_value = summary_ws.cell(row=row_idx, column=2, value=value)
        
        if row_idx == 1:  # Header row
            cell_key.font = header_font
            cell_key.fill = header_fill
            cell_value.font = header_font
            cell_value.fill = header_fill
        
        cell_key.border = border
        cell_value.border = border
    
    summary_ws.column_dimensions['A'].width = 30
    summary_ws.column_dimensions['B'].width = 40
    
    # Create Detailed Results sheet
    details_ws = wb.create_sheet("Detailed Results")
    
    # Prepare detailed data
    detailed_data = []
    
    # Headers
    headers = [
        "Match #", "Open Ticket Number", "Open Description", "Open Assignment Group", 
        "Open Customer", "Open Created", "Open Assigned To", "Closed Ticket Number", 
        "Closed Description", "Closed Assignment Group", "Closed Resolved By", 
        "Closed Assigned To", "Close Notes", "Similarity Score", "Similarity %"
    ]
    
    detailed_data.append(headers)
    
    # Data rows
    for match_idx, match in enumerate(results, 1):
        open_ticket = match['open_ticket']
        
        for closed_ticket in match['similar_closed_tickets']:
            row = [
                match_idx,
                open_ticket.get('Number', ''),
                open_ticket.get('Short Description', ''),
                open_ticket.get('Assignment group', ''),
                open_ticket.get('Customer', ''),
                open_ticket.get('Created', ''),
                open_ticket.get('Assigned to', ''),
                closed_ticket.get('Number', ''),
                closed_ticket.get('Short Description', ''),
                closed_ticket.get('Assignment group', ''),
                closed_ticket.get('Resolved by', ''),
                closed_ticket.get('Assigned to', ''),
                closed_ticket.get('Close Notes', ''),
                closed_ticket.get('Similarity Score', ''),
                closed_ticket.get('Similarity Percentage', '')
            ]
            detailed_data.append(row)
    
    # Write detailed data
    for row_idx, row_data in enumerate(detailed_data, 1):
        for col_idx, value in enumerate(row_data, 1):
            cell = details_ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border
            
            if row_idx == 1:  # Header row
                cell.font = header_font
                cell.fill = header_fill
            elif col_idx <= 7:  # Open ticket columns
                cell.fill = open_fill
            else:  # Closed ticket columns
                cell.fill = closed_fill
    
    # Adjust column widths
    for col in details_ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
        details_ws.column_dimensions[column].width = adjusted_width
    
    # Create Root Cause Analysis sheet
    root_cause_ws = wb.create_sheet("Root Cause Analysis")
    
    root_cause_data = [["Open Ticket Number", "Root Cause Analysis"]]
    
    for match in results:
        open_ticket = match['open_ticket']
        root_cause = match['root_cause']
        
        root_cause_data.append([
            open_ticket.get('Number', ''),
            root_cause
        ])
    
    # Write root cause data
    for row_idx, row_data in enumerate(root_cause_data, 1):
        for col_idx, value in enumerate(row_data, 1):
            cell = root_cause_ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            
            if row_idx == 1:  # Header row
                cell.font = header_font
                cell.fill = header_fill
    
    root_cause_ws.column_dimensions['A'].width = 20
    root_cause_ws.column_dimensions['B'].width = 80
    
    # Create Suggested Fixes sheet
    fixes_ws = wb.create_sheet("Suggested Fixes")
    
    fixes_data = [["Open Ticket Number", "Suggested Fix"]]
    
    for match in results:
        open_ticket = match['open_ticket']
        suggested_fix = match['suggested_fix']
        
        fixes_data.append([
            open_ticket.get('Number', ''),
            suggested_fix
        ])
    
    # Write fixes data
    for row_idx, row_data in enumerate(fixes_data, 1):
        for col_idx, value in enumerate(row_data, 1):
            cell = fixes_ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            
            if row_idx == 1:  # Header row
                cell.font = header_font
                cell.fill = header_fill
    
    fixes_ws.column_dimensions['A'].width = 20
    fixes_ws.column_dimensions['B'].width = 80
    
    # Save to BytesIO
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    
    return excel_buffer

@app.route('/')
def index():
    try:
        # Check if files exist
        if not os.path.exists(OPEN_TICKETS_PATH):
            return render_template('index.html', error=f"Open tickets file not found: {OPEN_TICKETS_PATH}")
        
        if not os.path.exists(CLOSED_TICKETS_PATH):
            return render_template('index.html', error=f"Closed tickets file not found: {CLOSED_TICKETS_PATH}")
        
        # Load datasets
        open_df = pd.read_excel(OPEN_TICKETS_PATH)
        closed_df = pd.read_excel(CLOSED_TICKETS_PATH)
        
        # Validate columns
        open_validation = validate_columns(open_df, 'open tickets')
        closed_validation = validate_columns(closed_df, 'closed tickets')
        
        # Get Assignment Groups from open tickets only
        assignment_groups = get_open_assignment_groups(open_df)
        
        file_info = {
            'open_tickets_count': len(open_df),
            'closed_tickets_count': len(closed_df),
            'open_columns': open_df.columns.tolist(),
            'closed_columns': closed_df.columns.tolist(),
            'required_columns': REQUIRED_COLUMNS,
            'preferred_columns': PREFERRED_COLUMNS,
            'missing_open_required': open_validation['missing_required'],
            'missing_closed_required': closed_validation['missing_required'],
            'missing_open_preferred': open_validation['missing_preferred'],
            'missing_closed_preferred': closed_validation['missing_preferred'],
            'assignment_groups': assignment_groups,
            'has_closing_note': 'Close Notes' in closed_df.columns,
            'has_resolved_by': 'Resolved by' in closed_df.columns,
            'llm_available': llm is not None,
            'llm_model': 'OpenAI GPT-4o-mini' if llm is not None else 'None'
        }
        
        # Check if we have minimum required columns for analysis
        if not open_validation['has_all_required']:
            error_msg = f"Open tickets missing required columns: {', '.join(open_validation['missing_required'])}"
            return render_template('index.html', error=error_msg, file_info=file_info)
        
        if not closed_validation['has_all_required']:
            error_msg = f"Closed tickets missing required columns: {', '.join(closed_validation['missing_required'])}"
            return render_template('index.html', error=error_msg, file_info=file_info)
        
        return render_template('index.html', file_info=file_info)
        
    except Exception as e:
        return render_template('index.html', error=str(e))

@app.route('/analyze', methods=['POST'])
def analyze():
    try:
        # Check files exist
        if not os.path.exists(OPEN_TICKETS_PATH):
            return jsonify({'success': False, 'error': 'Open tickets file not found'}), 404
        
        if not os.path.exists(CLOSED_TICKETS_PATH):
            return jsonify({'success': False, 'error': 'Closed tickets file not found'}), 404
        
        # Load datasets
        open_df = pd.read_excel(OPEN_TICKETS_PATH)
        closed_df = pd.read_excel(CLOSED_TICKETS_PATH)
        
        # Validate required columns
        for col in ['Short Description', 'Assignment group']:
            if col not in open_df.columns:
                return jsonify({'success': False, 'error': f"'{col}' column missing from open tickets"}), 400
            if col not in closed_df.columns:
                return jsonify({'success': False, 'error': f"'{col}' column missing from closed tickets"}), 400
        
        # Get parameters
        threshold = float(request.form.get('threshold', DEFAULT_THRESHOLD))
        assignment_group_filter = request.form.get('assignment_group', 'All')
        
        print(f"Analysis parameters: threshold={threshold}, assignment_group={assignment_group_filter}")
        
        # Find similar tickets
        results = calculate_semantic_similarity(open_df, closed_df, assignment_group_filter, threshold)
        
        # Calculate statistics
        total_open_tickets = len(open_df)
        total_closed_tickets = len(closed_df)
        
        filtered_open_tickets = total_open_tickets
        
        if assignment_group_filter and assignment_group_filter != 'All':
            filtered_open_tickets = len(open_df[open_df['Assignment group'] == assignment_group_filter])
        
        total_matches = sum(result['total_similar_closed'] for result in results)
        
        response = {
            'success': True,
            'results': results,
            'total_open_tickets': total_open_tickets,
            'total_closed_tickets': total_closed_tickets,
            'filtered_open_tickets': filtered_open_tickets,
            'open_tickets_with_matches': len(results),
            'total_matches': total_matches,
            'threshold': threshold,
            'threshold_percentage': f"{threshold * 100:.0f}%",
            'assignment_group_filter': assignment_group_filter,
            'required_columns': REQUIRED_COLUMNS,
            'preferred_columns': PREFERRED_COLUMNS,
            'open_columns': open_df.columns.tolist(),
            'closed_columns': closed_df.columns.tolist(),
            'has_closing_note': 'Close Notes' in closed_df.columns,
            'has_resolved_by': 'Resolved by' in closed_df.columns,
            'llm_available': llm is not None,
            'ai_enhanced_count': sum(1 for r in results if 'ROOT CAUSE ANALYSIS:' in r.get('root_cause', ''))
        }
        
        return jsonify(response)
        
    except Exception as e:
        print(f"Analysis error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/llm_status')
def llm_status():
    """
    Check if OpenAI GPT is available and working
    """
    global openai_client
    
    if openai_client is None:
        return jsonify({
            'available': False,
            'error': 'OpenAI client not initialized',
            'model': None
        })
    
    try:
        # Test the OpenAI client with a simple message
        test_response = openai_client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": "Test connection"}],
            max_tokens=10
        )
        return jsonify({
            'available': True,
            'status': 'Working',
            'model': 'OpenAI GPT-4o-mini',
            'test_response_length': len(test_response.choices[0].message.content)
        })
    except Exception as e:
        logging.error(f"OpenAI GPT test failed: {str(e)}")
        return jsonify({
            'available': False,
            'error': str(e),
            'model': 'OpenAI GPT-4o-mini'
        })

@app.route('/configure_llm', methods=['POST'])
def configure_llm():
    """
    Endpoint to configure OpenAI settings
    """
    global llm, openai_client
    
    try:
        data = request.get_json()
        api_key = data.get('api_key')
        model_name = data.get('model', 'gpt-4o-mini')
        base_url = data.get('base_url')  # Optional for custom OpenAI endpoints
        
        # Configure OpenAI client
        if api_key:
            if base_url:
                openai_client = OpenAI(api_key=api_key, base_url=base_url)
            else:
                openai_client = OpenAI(api_key=api_key)
        else:
            # Use existing API key from environment or default
            if base_url:
                openai_client = OpenAI(base_url=base_url)
            else:
                openai_client = OpenAI()
        
        llm = openai_client  # Keep compatibility
        
        # Test the OpenAI client
        test_response = openai_client.chat.completions.create(
            model=model_name,
            messages=[{"role": "user", "content": "Test message"}],
            max_tokens=10
        )
        
        logging.info(f"OpenAI client reconfigured successfully with model: {model_name}")
        
        return jsonify({
            'success': True, 
            'message': f'OpenAI GPT configured successfully with model: {model_name}',
            'model': model_name,
            'base_url': base_url
        })
        
    except Exception as e:
        logging.error(f"Failed to configure OpenAI GPT: {str(e)}")
        return jsonify({
            'success': False, 
            'error': f'Failed to configure OpenAI GPT: {str(e)}'
        }), 500

@app.route('/export_excel', methods=['POST'])
def export_excel():
    try:
        # Get the analysis results from the request
        data = request.get_json()
        
        if not data or 'results' not in data:
            return jsonify({'success': False, 'error': 'No analysis results to export'}), 400
        
        results = data['results']
        analysis_params = data.get('analysis_params', {})
        
        # Create Excel file
        excel_buffer = create_excel_export(results, analysis_params)
        
        # Generate filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"ticket_similarity_analysis_{timestamp}.xlsx"
        
        return send_file(
            excel_buffer,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Export error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    # Print OpenAI GPT status on startup
    if llm is not None:
        print("✅ OpenAI GPT initialized successfully - AI-enhanced suggestions available")
        try:
            # Test connection
            test_response = openai_client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": "Hello"}],
                max_tokens=5
            )
            print("✅ OpenAI GPT connection test successful")
        except Exception as e:
            print(f"⚠️ OpenAI GPT connection test failed: {str(e)}")
    else:
        print("⚠️ OpenAI GPT not available - will use basic suggestion format")
    
    app.run(debug=True)
