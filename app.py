import os
import csv
import shutil
from flask import Flask, render_template, request, send_file, jsonify, url_for
from werkzeug.utils import secure_filename
from pptx import Presentation
import win32com.client
import pythoncom
from PyPDF2 import PdfMerger
import time

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
app.config['ALLOWED_EXTENSIONS'] = {'pptx', 'csv'}

# Create necessary folders
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)


def allowed_file(filename, extension):
    """Check if file has allowed extension"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == extension


def cleanup_old_files():
    """Clean up old files from uploads and output folders"""
    folders = [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER']]
    for folder in folders:
        if os.path.exists(folder):
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                try:
                    if os.path.isfile(file_path):
                        # Delete files older than 1 hour
                        if time.time() - os.path.getmtime(file_path) > 3600:
                            os.unlink(file_path)
                except Exception as e:
                    print(f"Error deleting {file_path}: {e}")


def read_csv_data(csv_path):
    """Read CSV file and extract Id and Name columns"""
    data = []
    try:
        with open(csv_path, 'r', encoding='utf-8-sig') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if 'Name' in row and row['Name'].strip():
                    data.append({
                        'id': row.get('Id', '').strip(),
                        'name': row['Name'].strip()
                    })
        return data
    except Exception as e:
        raise Exception(f"Error reading CSV: {str(e)}")


def replace_placeholder_in_pptx(template_path, output_path, name):
    """Replace <<Name>> placeholder in PPTX with actual name"""
    try:
        prs = Presentation(template_path)
        
        # Iterate through all slides
        for slide in prs.slides:
            # Check all shapes in the slide
            for shape in slide.shapes:
                # Check text in shape
                if hasattr(shape, "text_frame"):
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if '<<Name>>' in run.text:
                                run.text = run.text.replace('<<Name>>', name)
                
                # Check tables
                if hasattr(shape, "table"):
                    for row in shape.table.rows:
                        for cell in row.cells:
                            if '<<Name>>' in cell.text:
                                for paragraph in cell.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        if '<<Name>>' in run.text:
                                            run.text = run.text.replace('<<Name>>', name)
        
        prs.save(output_path)
        return True
    except Exception as e:
        raise Exception(f"Error modifying PPTX: {str(e)}")


def convert_pptx_to_pdf(pptx_path, pdf_path):
    """Convert PPTX to PDF using PowerPoint COM automation (Windows only)"""
    # Initialize COM
    pythoncom.CoInitialize()
    
    try:
        # Initialize PowerPoint
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        powerpoint.Visible = 1
        
        # Open presentation
        deck = powerpoint.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
        
        # Export as high-quality PDF for printing
        # FixedFormatType: 2 = ppFixedFormatTypePDF
        # FixedFormatIntent: 1 = ppFixedFormatIntentPrint (high quality)
        # FrameSlides: 0 = msoFalse (no frame around slides)
        # HandoutOrder: 1 = ppPrintHandoutVerticalFirst
        # OutputType: 0 = ppPrintOutputSlides
        # PrintHiddenSlides: 0 = msoFalse
        # PrintRange: None (all slides)
        # RangeType: 1 = ppPrintAll
        # SlideShowName: "" (not applicable)
        # IncludeDocProperties: True
        # KeepIRMSettings: True
        # DocStructureTags: True (for accessibility)
        # BitmapMissingFonts: True (embed fonts as bitmaps if needed)
        # UseISO19005_1: False (PDF/A compliance - set True if needed)
        
        deck.ExportAsFixedFormat(
            os.path.abspath(pdf_path),
            FixedFormatType=2,  # PDF format
            Intent=1,  # Print quality (highest quality)
            FrameSlides=0,  # No frame
            HandoutOrder=1,
            OutputType=0,  # Slides output
            PrintHiddenSlides=0,
            PrintRange=None,
            RangeType=1,  # All slides
            SlideShowName="",
            IncludeDocProperties=True,
            KeepIRMSettings=True,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False  # Set to True for PDF/A compliance if needed
        )
        
        # Close presentation
        deck.Close()
        powerpoint.Quit()
        
        return True
    except Exception as e:
        raise Exception(f"Error converting PPTX to PDF: {str(e)}. Make sure PowerPoint is installed.")
    finally:
        # Uninitialize COM
        pythoncom.CoUninitialize()


def merge_pdfs(pdf_files, output_path):
    """Merge multiple PDF files into one"""
    try:
        merger = PdfMerger()
        
        for pdf in pdf_files:
            merger.append(pdf)
        
        merger.write(output_path)
        merger.close()
        
        return True
    except Exception as e:
        raise Exception(f"Error merging PDFs: {str(e)}")


@app.route('/')
def index():
    """Render the main upload page"""
    cleanup_old_files()
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate_certificates():
    """Process uploads and generate certificates"""
    try:
        # Check if files are present
        if 'template' not in request.files or 'csv_file' not in request.files:
            return jsonify({'error': 'Both template and CSV file are required'}), 400
        
        template_file = request.files['template']
        csv_file = request.files['csv_file']
        
        # Validate files
        if template_file.filename == '' or csv_file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(template_file.filename, 'pptx'):
            return jsonify({'error': 'Template must be a PPTX file'}), 400
        
        if not allowed_file(csv_file.filename, 'csv'):
            return jsonify({'error': 'Data file must be a CSV file'}), 400
        
        # Save uploaded files
        timestamp = str(int(time.time()))
        template_filename = secure_filename(f"template_{timestamp}.pptx")
        csv_filename = secure_filename(f"data_{timestamp}.csv")
        
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_filename)
        csv_path = os.path.join(app.config['UPLOAD_FOLDER'], csv_filename)
        
        template_file.save(template_path)
        csv_file.save(csv_path)
        
        # Read CSV data
        csv_data = read_csv_data(csv_path)
        
        if not csv_data:
            return jsonify({'error': 'No valid data found in CSV file. Make sure it has Id and Name columns'}), 400
        
        # Generate certificates
        pdf_files = []
        
        for idx, person in enumerate(csv_data):
            # Create PPTX with name
            pptx_output = os.path.join(app.config['OUTPUT_FOLDER'], f"cert_{timestamp}_{idx}.pptx")
            replace_placeholder_in_pptx(template_path, pptx_output, person['name'])
            
            # Convert to PDF
            pdf_output = os.path.join(app.config['OUTPUT_FOLDER'], f"cert_{timestamp}_{idx}.pdf")
            convert_pptx_to_pdf(pptx_output, pdf_output)
            
            pdf_files.append(pdf_output)
            
            # Clean up individual PPTX
            os.remove(pptx_output)
        
        # Merge all PDFs
        merged_pdf = f"certificates_{timestamp}.pdf"
        merged_pdf_path = os.path.join(app.config['OUTPUT_FOLDER'], merged_pdf)
        merge_pdfs(pdf_files, merged_pdf_path)
        
        # Clean up individual PDFs
        for pdf in pdf_files:
            os.remove(pdf)
        
        # Clean up uploaded files
        os.remove(template_path)
        os.remove(csv_path)
        
        return jsonify({
            'success': True,
            'message': f'Generated {len(csv_data)} certificates successfully!',
            'download_url': url_for('download_file', filename=merged_pdf),
            'filename': merged_pdf
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/download/<filename>')
def download_file(filename):
    """Download the merged PDF file"""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        return send_file(file_path, as_attachment=True, download_name='certificates.pdf')
    except Exception as e:
        return jsonify({'error': f'File not found: {str(e)}'}), 404


if __name__ == '__main__':
    app.run(debug=True, port=5000)
