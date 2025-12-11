import os
import fitz # PyMuPDF for PDF handling
import io
import sys
from google import genai
from PIL import Image
from docx import Document
from docx.enum.section import WD_SECTION
from flask import Flask, request, render_template, send_file
# Note: Ensure all dependencies are installed: pip install Flask google-genai pymupdf Pillow python-docx python-dotenv

# --- ENVIRONMENT SETUP ---
# Load environment variables from .env file if it exists
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass 
    
# 1. --- FLASK SETUP --- (Gunicorn looks for this object)
# This MUST be defined globally and early.
app = Flask(__name__)
# Temporary folder to store uploaded files
app.config['UPLOAD_FOLDER'] = './tmp/uploads'
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    # Create the directory if it doesn't exist
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# --- GEMINI CONFIGURATION ---
API_KEY = os.environ.get("GEMINI_API_KEY") 
MODEL_ID = "gemini-2.5-flash-lite" 

# REVISED PROMPT for LOGICAL STRUCTURE and CLEAN READABILITY
OCR_PROMPT = (
    "Perform Optical Character Recognition (OCR) on this document. The content is primarily in Hindi, English, Gujarati, or Marathi. "
    
    "Your output MUST be structured logically. Consolidate fragmented text into natural paragraphs where appropriate, but ensure each question, instruction, and its options are clearly separated by new lines. "
    "Use standard double line breaks to separate distinct elements like question headers, question bodies, and individual multiple-choice options (A, B, C, D). "
    
    "Do NOT join entire questions or unrelated blocks of text together. "
    "Prioritize text accuracy and logical separation for readability over replicating the exact visual position of every word. "
    "Do not include any formatting tags (HTML/Markdown)."
)
# --- END GEMINI CONFIGURATION ---


# --- CORE PROCESSING FUNCTION (Memory Optimized) ---
def process_document(input_file_path, prompt, client):
    """
    Processes a PDF or image file by converting pages to in-memory streams 
    (for PDF) or using the file path (for image) and sending them to Gemini for OCR.
    Returns an io.BytesIO stream containing the DOCX file content.
    """
    document = Document()
    pdf_document = None
    pages_to_process = []
    
    if client is None:
        raise Exception("Gemini client is not initialized. API key missing.")
    
    try:
        # 1. Prepare pages for processing
        if input_file_path.lower().endswith(('.pdf')):
            pdf_document = fitz.open(input_file_path) 
            num_pages = len(pdf_document)
            
            # Use 150 DPI scaling factor for good quality
            scale_factor = 150 / 72  
            matrix = fitz.Matrix(scale_factor, scale_factor)
            
            for i in range(num_pages):
                page = pdf_document.load_page(i)
                pix = page.get_pixmap(matrix=matrix) 
                png_bytes = pix.tobytes(output='png')
                image_stream = io.BytesIO(png_bytes)
                pages_to_process.append(image_stream)
            
        elif input_file_path.lower().endswith(('.jpg', '.jpeg', '.png')):
            pages_to_process = [input_file_path]
            
        else:
            raise ValueError("Unsupported file type. Please upload a PDF, JPG, or PNG.")

        # 2. GEMINI PROCESSING LOOP
        for i, page_source in enumerate(pages_to_process):
            page_number = i + 1
            
            with Image.open(page_source) as page_image:
                response = client.models.generate_content(
                    model=MODEL_ID, 
                    contents=[prompt, page_image]
                )
            
            extracted_text = response.text
            
            # ⚠️ ROBUSTNESS CHECK: Check for empty text
            if not extracted_text or extracted_text.strip() == "":
                extracted_text = "\n--- OCR failed to return text for this page. Please review the original image quality. ---"
            
            # 3. Add content to DOCX document
            document.add_paragraph(f"\n--- Page {page_number} ---")
            document.add_paragraph(extracted_text)
            
            # Add a new page break in the DOCX for subsequent pages
            if len(pages_to_process) > 1 and page_number < len(pages_to_process):
                document.add_section(WD_SECTION.NEW_PAGE)
                
        # 4. Save the final DOCX to an in-memory buffer
        doc_io = io.BytesIO()
        document.save(doc_io)
        doc_io.seek(0)
        return doc_io

    finally:
        # CLEANUP: Close PDF resource
        if pdf_document is not None:
            pdf_document.close()
            
        # Cleanup: Remove the original uploaded file
        if os.path.exists(input_file_path):
            try: 
                os.remove(input_file_path)
            except Exception as e:
                print(f"Warning: Could not delete original file {input_file_path}. {e}", file=sys.stderr)


# --- FLASK ROUTES ---

@app.route('/', methods=['GET'])
def index():
    # Renders the HTML form for file upload.
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_and_convert():
    if not API_KEY:
        return "Server Error: Gemini API key is missing. Please set the GEMINI_API_KEY environment variable.", 500
        
    if 'file' not in request.files:
        return 'No file part', 400
    
    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400
    
    filepath = None
    
    if file:
        filename = os.path.basename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        # Save the file temporarily
        try:
            file.save(filepath)
        except Exception as e:
            return "Failed to save the uploaded file.", 500

        try:
            # Initialize client for this request
            gemini_client = genai.Client(api_key=API_KEY)
            doc_stream = process_document(filepath, OCR_PROMPT, gemini_client)
            
            # Create a smart download name
            output_filename = filename.rsplit('.', 1)[0] + '_OCR_Structured.docx'
            
            return send_file(
                doc_stream,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                as_attachment=True,
                download_name=output_filename
            )
            
        except ValueError as e:
            return str(e), 400
        except Exception as e:
            print(f"An unexpected error occurred during processing: {e}", file=sys.stderr)
            return "An internal conversion error occurred. Check the server logs.", 500