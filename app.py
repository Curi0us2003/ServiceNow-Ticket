from flask import Flask, render_template, request, jsonify
import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import os
import re
import asyncio
from googletrans import Translator
from datetime import datetime
import warnings

warnings.filterwarnings('ignore')

app = Flask(__name__)

# File paths
OPEN_TICKETS_PATH = r"C:\Users\RitabrataRoyChoudhur\OneDrive - GyanSys Inc\Desktop\Python\ServiceNow\open.xlsx"
CLOSED_TICKETS_PATH = r"C:\Users\RitabrataRoyChoudhur\OneDrive - GyanSys Inc\Desktop\Python\ServiceNow\close.xlsx"
DEFAULT_THRESHOLD = 0.75

# Updated required columns to include 'Assigned To' and 'Closing note'
REQUIRED_COLUMNS = ['Number', 'Short Description', 'Assignment group', 'Customer', 'Created', 'Assigned to']
PREFERRED_COLUMNS = REQUIRED_COLUMNS + ['Closing note', 'Resolved by']

translator = Translator()

def clean_value(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def translate_to_english(descriptions):
    translations = []
    for desc in descriptions:
        if pd.isna(desc) or str(desc).strip() == "":
            translations.append({
                "original": "",
                "detected_lang": "en",
                "english": ""
            })
            continue
        desc_str = str(desc).strip()
        try:
            # Detect language (synchronous)
            detection = translator.detect(desc_str)
            if detection.lang != 'en' and detection.confidence > 0.7:
                # Translate to English (synchronous)
                translated = translator.translate(desc_str, dest="en")
                translations.append({
                    "original": desc_str,
                    "detected_lang": detection.lang,
                    "english": translated.text
                })
                print(f"{desc_str} -> {translated.text} (from {detection.lang})")
            else:
                translations.append({
                    "original": desc_str,
                    "detected_lang": detection.lang,
                    "english": desc_str
                })
        except Exception as e:
            print(f"Translation error for '{desc_str}': {str(e)}")
            translations.append({
                "original": desc_str,
                "detected_lang": "unknown",
                "english": desc_str
            })
    return translations

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

def generate_suggested_fix(similar_closed_tickets):
    """
    Generate suggested fix for open ticket based on closing notes from similar closed tickets
    """
    closing_notes = []
    
    for ticket in similar_closed_tickets:
        closing_note = ticket.get('Closing note', '').strip()
        if closing_note and closing_note.lower() not in ['n/a', 'na', '']:
            closing_notes.append(f"• {closing_note}")
    
    if closing_notes:
        # Concatenate all closing notes
        suggested_fix = "Based on similar resolved tickets:\n\n" + "\n".join(closing_notes)
        return suggested_fix
    else:
        return "No closing notes available from similar tickets"

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
    
    # Translate descriptions to English
    print("Translating open ticket descriptions...")
    open_translations = translate_to_english(open_df['Short Description'].tolist())
    
    print("Translating closed ticket descriptions...")
    closed_translations = translate_to_english(closed_df['Short Description'].tolist())
    
    # Debug translation output
    for i, item in enumerate(open_translations[:5]):  # Show first 5 for debugging
        if item['original'] != item['english']:
            print(f"Open {i}: {item['original']} -> {item['english']} (from {item['detected_lang']})")
    
    # Preprocess texts
    open_texts = [preprocess_text(trans['english']) for trans in open_translations]
    closed_texts = [preprocess_text(trans['english']) for trans in closed_translations]
    
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
                
                # Add translation info
                closed_ticket_data['Translated Description'] = closed_translations[j]['english']
                closed_ticket_data['Original Description'] = closed_translations[j]['original']
                closed_ticket_data['Detected Language'] = closed_translations[j]['detected_lang']
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
            
            # Add translation info for open ticket
            open_ticket_data['Translated Description'] = open_translations[i]['english']
            open_ticket_data['Original Description'] = open_translations[i]['original']
            open_ticket_data['Detected Language'] = open_translations[i]['detected_lang']
            
            # Generate suggested fix based on all similar closed tickets
            suggested_fix = generate_suggested_fix(similar_closed)
            
            results.append({
                'open_ticket': open_ticket_data,
                'similar_closed_tickets': similar_closed,
                'best_similarity_score': similar_closed[0]['Similarity Score'],
                'total_similar_closed': len(similar_closed),
                'suggested_fix': suggested_fix
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
            'has_closing_note': 'Closing note' in closed_df.columns,
            'has_resolved_by': 'Resolved by' in closed_df.columns
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
            'has_closing_note': 'Closing note' in closed_df.columns,
            'has_resolved_by': 'Resolved by' in closed_df.columns
        }
        
        return jsonify(response)
        
    except Exception as e:
        print(f"Analysis error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
