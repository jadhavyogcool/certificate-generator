# Certificate Generator - Flask Application

A Flask web application that generates personalized certificates from PowerPoint templates.

## Features

- 📤 **Upload PPTX Template** - Upload a PowerPoint certificate template with placeholders
- 📊 **Upload CSV Data** - Batch process multiple names from CSV files
- 🔄 **Automatic Generation** - Replace `<<Name>>` placeholder with actual names
- 📥 **Merged PDF Output** - Download all certificates as a single PDF file

## Requirements

- Windows OS (for PowerPoint COM automation)
- Microsoft PowerPoint installed
- Python 3.8+

## Installation

1. **Clone or download this repository**

2. **Create a virtual environment**
   ```bash
   python -m venv venv
   ```

3. **Activate the virtual environment**
   ```bash
   .\venv\Scripts\activate
   ```

4. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. **Start the Flask application**
   ```bash
   python app.py
   ```

2. **Open your browser**
   Navigate to: `http://localhost:5000`

3. **Upload your files**
   - Upload a PPTX template containing `<<Name>>` placeholder
   - Upload a CSV file with columns: `Id` and `Name`

4. **Generate certificates**
   Click "Generate Certificates" and wait for processing

5. **Download**
   Download the merged PDF containing all certificates

## CSV Format

Your CSV file should have the following structure:

```csv
Id,Name
1,John Doe
2,Jane Smith
3,Alice Johnson
```

## PPTX Template

Create a PowerPoint slide with the placeholder text `<<Name>>` where you want the person's name to appear. The application will replace this placeholder with actual names from your CSV file.

## Project Structure

```
certificate-generator/
├── app.py              # Main Flask application
├── requirements.txt    # Python dependencies
├── templates/
│   └── index.html     # Upload interface
├── static/
│   └── style.css      # Styling
├── uploads/           # Temporary upload storage
└── output/            # Generated certificates
```

## Technologies Used

- **Flask** - Web framework
- **python-pptx** - PowerPoint manipulation
- **comtypes** - Windows COM interface for PowerPoint
- **PyPDF2** - PDF merging

## Notes

- Generated files are automatically cleaned up after 1 hour
- Maximum file size: 50MB
- Supports PPTX and CSV formats only

## License

MIT License