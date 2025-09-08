# Ticket Similarity Analysis Tool

A Flask-based web application that analyzes IT tickets from Excel files and identifies similar tickets using Natural Language Processing (NLP) techniques. The tool automatically groups tickets with similar descriptions and provides detailed similarity metrics.

## ✨ Features

- **📊 Excel File Processing**: Upload and analyze IT tickets from Excel (.xlsx) files
- **🌍 Multi-language Support**: Automatically detects and translates non-English ticket descriptions
- **🧹 Smart Text Preprocessing**: Removes special characters, numbers, and stop words
- **🤖 Advanced NLP**: Uses CountVectorizer with Bag of Words and Bigrams for similarity detection
- **📈 Similarity Scoring**: Implements cosine similarity with configurable thresholds
- **👥 Intelligent Grouping**: Clusters similar tickets with detailed metadata
- **🌐 Web Interface**: Clean, intuitive Flask-based UI
- **🔌 API Endpoint**: JSON API for integration with other tools
- **⚙️ Configurable**: Adjustable similarity thresholds (default: 75%)

## 🛠️ Tech Stack

- **Backend**: Flask (Python)
- **Machine Learning**: scikit-learn, numpy, pandas
- **Translation**: Google Translate API (googletrans)
- **Data Processing**: openpyxl for Excel handling
- **Frontend**: Jinja2 templates

## 📋 Prerequisites

- Python 3.8 or higher
- Excel file with IT ticket data

## 🚀 Installation

### 1. Clone the Repository

```bash
git clone https://github.com/Curi0us2003/ServiceNow-Ticket.git
cd ticket-similarity-analysis
```

### 2. Create Virtual Environment

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

If `requirements.txt` doesn't exist, install manually:

```bash
pip install flask pandas numpy scikit-learn googletrans==4.0.0rc1 openpyxl
```

### 4. Configure Excel File Path

Update the file path in `app.py`:

```python
EXCEL_FILE_PATH = r"FILE_PATH_FOR_EXCEL"
SHEET_NAME = "RawData"  # Your sheet name
```

### 5. Run the Application

```bash
python app.py
```

The application will be available at: `http://127.0.0.1:5000/`

## 📖 Usage

### Web Interface

1. Open your browser and navigate to `http://127.0.0.1:5000/`
2. Review the file details and required columns
3. Click "Analyze" to process tickets with default threshold (75%)
4. View grouped similar tickets with:
   - Similarity percentages
   - Group sizes
   - Original and translated descriptions
   - Key metadata (CI, Assignment Group, Customer, etc.)

### API Usage

Send a POST request to `/analyze` endpoint:

```bash
curl -X POST http://127.0.0.1:5000/analyze \
  -H "Content-Type: application/json" \
  -d '{"threshold": 0.8}'
```

**Response Format:**
```json
{
  "success": true,
  "groups": [
    {
      "tickets": [
        {
          "CI": "SERVER001",
          "Description": "Server not responding",
          "Translated_Description": "Server not responding",
          "Created AMER": "2024-01-15",
          "Assignment Group": "Infrastructure",
          "Customer": "John Doe"
        }
      ],
      "similarity_score": 0.82,
      "group_size": 5,
      "similarity_percentage": "82.0%"
    }
  ],
  "total_groups": 3,
  "total_tickets": 120,
  "tickets_in_groups": 15,
  "threshold": 0.8,
  "threshold_percentage": "80%"
}
```

## 📂 Project Structure

```
ticket-similarity-analysis/
├── app.py                 # Main Flask application
├── templates/
│   └── index.html         # Web interface template
├── static/                # CSS/JS files (optional)
├── requirements.txt       # Python dependencies
└── README.md             # This file
```

## 📊 Expected Excel Format

Your Excel file should contain the following columns in the "RawData" sheet:

| Column | Description |
|--------|-------------|
| CI | Configuration Item |
| Description | Ticket description text |
| Created AMER | Creation date |
| Assignment Group | Team assigned to ticket |
| Customer | Customer name |

## ⚙️ Configuration

### Similarity Threshold
- Default: 0.75 (75%)
- Range: 0.0 to 1.0
- Higher values = more strict similarity matching

### Supported Languages
- Automatic detection of non-English text
- Translation powered by Google Translate
- Supports 100+ languages

## 🚀 Future Enhancements

- [ ] **TF-IDF Integration**: Enhanced similarity detection
- [ ] **Visualization Dashboard**: Interactive ticket cluster graphs  
- [ ] **Batch Processing**: Handle multiple Excel files
- [ ] **File Upload Interface**: Direct file upload capability
- [ ] **Translation Caching**: Optimize API usage
- [ ] **Export Options**: CSV/Excel export of results
- [ ] **Advanced Filters**: Filter by date range, assignment group
- [ ] **Real-time Processing**: WebSocket-based live updates


**⭐ If this project helped you, please give it a star!**

Built with ❤️ using Python and Flask
