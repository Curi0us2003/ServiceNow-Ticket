from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import os
import re
from googletrans import Translator
from datetime import datetime
import io

import warnings
warnings.filterwarnings('ignore')

app = Flask(__name__)

EXCEL_FILE_PATH = r"C:\Users\RitabrataRoyChoudhur\OneDrive - GyanSys Inc\Desktop\Python\ServiceNow\tickets.xlsx"
SHEET_NAME = "RawData"
DEFAULT_THRESHOLD = 0.75
REQUIRED_COLUMNS = ['Number', 'Short Description', 'Correct CI', 'Resolved by', 'Created AMER', 'Assignment group', 'Customer']

translator = Translator()

def clean_value(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def translate_to_english(text):
    if pd.isna(text) or text == "":
        return ""
    text = str(text).strip()
    if not text:
        return ""
    try:
        detection = translator.detect(text)
        if detection.lang != 'en' and detection.confidence > 0.7:
            translation = translator.translate(text, dest='en')
            return translation.text
        else:
            return text
    except Exception as e:
        print(f"Translation error: {str(e)}")
        return text

def preprocess_text(text):
    if pd.isna(text):
        return ""
    text = str(text).lower()
    text = re.sub(r'[^a-zA-Z0-9\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def find_similar_tickets(df, assignment_group=None, threshold=0.75):
    print("Starting similarity analysis...")
    if 'Short Description' not in df.columns:
        raise ValueError("'Short Description' column not found")

    # Filter by assignment group if specified
    df_filtered = df.copy()
    if assignment_group and assignment_group != 'All':
        if 'Assignment group' not in df.columns:
            raise ValueError("'Assignment group' column not found")
        df_filtered = df[df['Assignment group'] == assignment_group].reset_index(drop=True)
        if df_filtered.empty:
            print(f"No tickets found with 'Assignment group' = '{assignment_group}'")
            return []

    print(f"Filtered tickets count: {len(df_filtered)}")

    print("Translating descriptions...")
    translated_descriptions = [translate_to_english(desc) for desc in df_filtered['Short Description']]

    processed_texts = [preprocess_text(text) for text in translated_descriptions]

    print("Calculating similarity matrix...")
    vectorizer = CountVectorizer(stop_words='english', ngram_range=(1, 2))
    try:
        bow_matrix = vectorizer.fit_transform(processed_texts)
        similarity_matrix = cosine_similarity(bow_matrix)
    except ValueError:
        similarity_matrix = np.zeros((len(processed_texts), len(processed_texts)))
    
    visited = set()
    groups = []

    for i in range(len(df_filtered)):
        if i in visited:
            continue
        similar_indices = [j for j in range(len(df_filtered)) if similarity_matrix[i][j] >= threshold]
        if len(similar_indices) > 1:
            for idx in similar_indices:
                visited.add(idx)

            group_tickets = []
            for idx in similar_indices:
                ticket_data = {}
                # Add all columns from the dataframe
                for col in df_filtered.columns:
                    ticket_data[col] = clean_value(df_filtered.iloc[idx][col])
                
                ticket_data['Translated Description'] = clean_value(translated_descriptions[idx])
                ticket_data['Original Description'] = clean_value(df_filtered.iloc[idx]['Short Description'])
                group_tickets.append(ticket_data)
            
            similarities = [similarity_matrix[x][y] for x in similar_indices for y in similar_indices if x != y]
            avg_similarity = float(np.mean(similarities)) if similarities else 0.0

            groups.append({
                'tickets': group_tickets,
                'similarity_score': round(avg_similarity, 3),
                'group_size': len(similar_indices),
                'similarity_percentage': f"{avg_similarity * 100:.1f}%"
            })

    groups.sort(key=lambda x: x['similarity_score'], reverse=True)
    print(f"Found {len(groups)} similar groups")
    return groups

def export_to_excel(groups, assignment_group, threshold, total_tickets, filtered_tickets):
    """Export similarity analysis results to Excel format"""
    # Create a BytesIO buffer to store the Excel file
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        workbook = writer.book
        
        # Create a worksheet
        worksheet_name = 'Similarity Analysis Results'
        worksheet = workbook.create_sheet(worksheet_name)
        workbook.active = worksheet
        
        # Styling
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        from openpyxl.utils import get_column_letter
        
        # Define styles
        header_font = Font(bold=True, size=14, color='FFFFFF')
        subheader_font = Font(bold=True, size=12)
        group_header_font = Font(bold=True, size=11, color='FFFFFF')
        regular_font = Font(size=10)
        
        header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        group_fill = PatternFill(start_color='5B9BD5', end_color='5B9BD5', fill_type='solid')
        required_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
        additional_fill = PatternFill(start_color='E7F3FF', end_color='E7F3FF', fill_type='solid')
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        current_row = 1
        
        # Title and summary information
        worksheet.merge_cells(f'A{current_row}:H{current_row}')
        cell = worksheet[f'A{current_row}']
        cell.value = f"Ticket Similarity Analysis Results - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        current_row += 2
        
        # Summary information
        summary_data = [
            ['Assignment Group:', assignment_group if assignment_group != 'All' else 'All Groups'],
            ['Similarity Threshold:', f"{threshold * 100:.0f}%"],
            ['Total Tickets:', total_tickets],
            ['Filtered Tickets:', filtered_tickets],
            ['Similar Groups Found:', len(groups)],
            ['Total Tickets in Groups:', sum(group['group_size'] for group in groups)]
        ]
        
        for item, value in summary_data:
            worksheet[f'A{current_row}'] = item
            worksheet[f'A{current_row}'].font = subheader_font
            worksheet[f'B{current_row}'] = value
            worksheet[f'B{current_row}'].font = regular_font
            current_row += 1
        
        current_row += 2
        
        if not groups:
            worksheet[f'A{current_row}'] = "No similar groups found with the specified criteria."
            worksheet[f'A{current_row}'].font = subheader_font
        else:
            # Process each group
            for group_idx, group in enumerate(groups, 1):
                # Group header
                group_header_row = current_row
                worksheet.merge_cells(f'A{current_row}:H{current_row}')
                cell = worksheet[f'A{current_row}']
                cell.value = f"Group {group_idx} - {group['group_size']} Similar Tickets ({group['similarity_percentage']} Similarity)"
                cell.font = group_header_font
                cell.fill = group_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
                current_row += 2
                
                # Get all unique columns from tickets in this group
                all_columns = set()
                for ticket in group['tickets']:
                    all_columns.update(ticket.keys())
                
                # Remove translation-specific columns from display
                display_columns = [col for col in all_columns if col not in ['Original Description', 'Translated Description']]
                
                # Separate required and additional columns
                required_cols = [col for col in REQUIRED_COLUMNS if col in display_columns]
                additional_cols = [col for col in display_columns if col not in REQUIRED_COLUMNS]
                
                # Write column headers
                col_num = 1
                
                # Required columns header
                if required_cols:
                    start_col = col_num
                    for col in required_cols:
                        cell = worksheet.cell(row=current_row, column=col_num)
                        cell.value = col
                        cell.font = subheader_font
                        cell.fill = required_fill
                        cell.border = thin_border
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        col_num += 1
                    
                    current_row += 1
                    
                    # Write required column data
                    for ticket in group['tickets']:
                        col_num = 1
                        for col in required_cols:
                            cell = worksheet.cell(row=current_row, column=col_num)
                            
                            # Special handling for Short Description
                            if col == 'Short Description':
                                translated_desc = ticket.get('Translated Description', ticket.get(col, 'N/A'))
                                original_desc = ticket.get('Original Description', '')
                                
                                if original_desc and original_desc != translated_desc:
                                    cell.value = f"Translated: {translated_desc}\nOriginal: {original_desc}"
                                else:
                                    cell.value = translated_desc
                            else:
                                cell.value = ticket.get(col, 'N/A')
                            
                            cell.font = regular_font
                            cell.border = thin_border
                            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                            col_num += 1
                        current_row += 1
                    
                    current_row += 1
                
                # Additional columns (if any)
                if additional_cols:
                    worksheet[f'A{current_row}'] = "Additional Details:"
                    worksheet[f'A{current_row}'].font = subheader_font
                    current_row += 1
                    
                    # Additional columns header
                    col_num = 1
                    for col in additional_cols:
                        cell = worksheet.cell(row=current_row, column=col_num)
                        cell.value = col
                        cell.font = subheader_font
                        cell.fill = additional_fill
                        cell.border = thin_border
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        col_num += 1
                    
                    current_row += 1
                    
                    # Write additional column data
                    for ticket in group['tickets']:
                        col_num = 1
                        for col in additional_cols:
                            cell = worksheet.cell(row=current_row, column=col_num)
                            cell.value = ticket.get(col, 'N/A')
                            cell.font = regular_font
                            cell.border = thin_border
                            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                            col_num += 1
                        current_row += 1
                
                current_row += 2  # Space between groups
        
        # Auto-adjust column widths
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
            worksheet.column_dimensions[column_letter].width = adjusted_width
        
        # Set row heights for better readability
        for row in worksheet.iter_rows():
            worksheet.row_dimensions[row[0].row].height = 20
    
    output.seek(0)
    return output

def get_assignment_groups(df):
    """Get unique assignment groups from the dataframe"""
    if 'Assignment group' in df.columns:
        groups = df['Assignment group'].dropna().unique().tolist()
        groups.sort()
        return groups
    return []

@app.route('/')
def index():
    try:
        if not os.path.exists(EXCEL_FILE_PATH):
            return render_template('index.html', error=f"File not found: {EXCEL_FILE_PATH}")

        # Check sheet exists
        xl = pd.ExcelFile(EXCEL_FILE_PATH)
        if SHEET_NAME not in xl.sheet_names:
            return render_template('index.html', error=f"Sheet '{SHEET_NAME}' not found in Excel file.")
        
        # Load all columns to get assignment groups
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=SHEET_NAME)
        
        # Get available columns
        available_columns = df.columns.tolist()
        missing_cols = [col for col in REQUIRED_COLUMNS if col not in available_columns]
        
        # Get assignment groups
        assignment_groups = get_assignment_groups(df)

        file_info = {
            'path': EXCEL_FILE_PATH,
            'sheet_name': SHEET_NAME,
            'total_tickets': len(df),
            'available_columns': available_columns,
            'required_columns': REQUIRED_COLUMNS,
            'missing_columns': missing_cols,
            'assignment_groups': assignment_groups
        }
        return render_template('index.html', file_info=file_info)
    except Exception as e:
        return render_template('index.html', error=str(e))

@app.route('/analyze', methods=['POST'])
def analyze():
    try:
        if not os.path.exists(EXCEL_FILE_PATH):
            return jsonify({'success': False, 'error': 'Excel file not found'}), 404
        
        xl = pd.ExcelFile(EXCEL_FILE_PATH)
        if SHEET_NAME not in xl.sheet_names:
            return jsonify({'success': False, 'error': f"Sheet '{SHEET_NAME}' not found in Excel file."}), 400

        # Load all columns
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=SHEET_NAME)
        
        if 'Short Description' not in df.columns:
            return jsonify({'success': False, 'error': "'Short Description' column is required but not found."}), 400

        threshold = float(request.form.get('threshold', DEFAULT_THRESHOLD))
        assignment_group = request.form.get('assignment_group', 'All')

        groups = find_similar_tickets(df, assignment_group, threshold)

        total_tickets_in_groups = sum(group['group_size'] for group in groups)
        
        # Calculate filtered ticket count
        filtered_ticket_count = len(df)
        if assignment_group and assignment_group != 'All':
            filtered_ticket_count = len(df[df['Assignment group'] == assignment_group])
        
        response = {
            'success': True,
            'groups': groups,
            'total_groups': len(groups),
            'total_tickets': len(df),
            'filtered_tickets': filtered_ticket_count,
            'tickets_in_groups': total_tickets_in_groups,
            'threshold': threshold,
            'threshold_percentage': f"{threshold * 100:.0f}%",
            'assignment_group': assignment_group,
            'required_columns': REQUIRED_COLUMNS,
            'all_columns': df.columns.tolist()
        }
        return jsonify(response)
    except Exception as e:
        print(f"Analysis error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/export', methods=['POST'])
def export():
    try:
        # Get the analysis results from the request
        data = request.get_json()
        
        if not data or 'groups' not in data:
            return jsonify({'success': False, 'error': 'No analysis data provided'}), 400
        
        # Extract export parameters
        groups = data['groups']
        assignment_group = data.get('assignment_group', 'All')
        threshold = data.get('threshold', DEFAULT_THRESHOLD)
        total_tickets = data.get('total_tickets', 0)
        filtered_tickets = data.get('filtered_tickets', 0)
        
        # Generate Excel file
        excel_buffer = export_to_excel(groups, assignment_group, threshold, total_tickets, filtered_tickets)
        
        # Create filename with timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        assignment_part = assignment_group.replace(' ', '_') if assignment_group != 'All' else 'All_Groups'
        filename = f"Ticket_Similarity_Analysis_{assignment_part}_{timestamp}.xlsx"
        
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
    app.run(debug=True)