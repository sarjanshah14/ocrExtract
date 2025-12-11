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
# Load environment variables from .env file if it exists (for local development only)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass 
    

# --- FLASK SETUP ---
app = Flask(__name__)
# Temporary folder to store uploaded files
app.config['UPLOAD_FOLDER'] = './tmp/uploads'
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# --- GEMINI CONFIGURATION ---
# IMPORTANT: This securely loads the key from the environment (e.g., .env file or production configuration)
API_KEY = os.environ.get("GEMINI_API_KEY") 
MODEL_ID = "gemini-2.5-flash-lite" 

# REVAMPED PROMPT for Strict Line-Fidelity and Multilingual Support
# This prompt prioritizes replicating the exact visual structure (including two-column options)
OCR_PROMPT = (
    "Perform Optical Character Recognition (OCR) on this image. The content is primarily in Hindi, English, Gujarati, or Marathi. "
    
    "Crucially, you **MUST** strictly preserve the visual line breaks and spatial layout of the original source document. "
    "If a line of text in the image starts at one point and breaks to the next line, the output **MUST** replicate that exact line break. "
    "Do NOT join lines into continuous paragraphs, even if they form a complete sentence. "
    
    "Example: If the image has 'The dog is big and' on Line 1, and 'eating food.' on Line 2, your output must retain those two separate lines. "
    
    "It is acceptable if some badly handwritten words are unreadable (output the text as best as possible). "
    "Do not include any formatting tags (HTML/Markdown). "
    "Ensure proper spacing and separate distinct blocks of text like headings and paragraphs."
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
    pdf_document = None # Initialize to None for cleanup safety
    pages_to_process = []
    
    # Check if the API client is initialized before processing
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
                
                # Convert pixmap directly to PNG bytes stream (in memory)
                png_bytes = pix.tobytes(output='png')
                
                # Use io.BytesIO to treat the bytes as a file for PIL
                image_stream = io.BytesIO(png_bytes)
                pages_to_process.append(image_stream)
            
        elif input_file_path.lower().endswith(('.jpg', '.jpeg', '.png')):
            # For images, we just pass the file path to open it later
            pages_to_process = [input_file_path]
            
        else:
            raise ValueError("Unsupported file type. Please upload a PDF, JPG, or PNG.")

        # 2. GEMINI PROCESSING LOOP
        for i, page_source in enumerate(pages_to_process):
            page_number = i + 1
            
            # Open the image from the stream (for PDF) or file path (for image)
            with Image.open(page_source) as page_image:
                response = client.models.generate_content(
                    model=MODEL_ID, 
                    contents=[prompt, page_image]
                )
            
            extracted_text = response.text
            
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
        # CLEANUP
        if pdf_document is not None:
            pdf_document.close()
            
        # Cleanup: Remove the original uploaded file
        if os.path.exists(input_file_path):
            try: 
                os.remove(input_file_path)
            except Exception as e:
                # Print a warning but continue if file cannot be deleted
                print(f"Warning: Could not delete original file {input_file_path}. {e}", file=sys.stderr)


# --- FLASK ROUTES ---

@app.route('/', methods=['GET'])
def index():
    # Renders the HTML form for file upload. You need to create a 'templates/index.html' file.
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_and_convert():
    if not API_KEY:
        return "Server not configured: Gemini API key is missing. Please set the GEMINI_API_KEY environment variable.", 500
        
    if 'file' not in request.files:
        return 'No file part', 400
    
    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400
    
    filepath = None
    
    if file:
        # Use a secure filename to prevent directory traversal
        filename = os.path.basename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        # Save the file temporarily
        try:
            file.save(filepath)
        except Exception as e:
            print(f"Error saving file: {e}", file=sys.stderr)
            return "Failed to save the uploaded file.", 500

        try:
            # Initialize client for this request
            gemini_client = genai.Client(api_key=API_KEY)
            doc_stream = process_document(filepath, OCR_PROMPT, gemini_client)
            
            # Create a smart download name
            output_filename = filename.rsplit('.', 1)[0] + '_OCR.docx'
            
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
        finally:
             # Cleanup handled by process_document, but a robust app would ensure temporary files are managed.
             pass


if __name__ == '__main__':
    # Flask requires an 'index.html' file in a 'templates' folder to run the index route.
    if not os.path.exists('templates/index.html'):
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!", file=sys.stderr)
        print("SETUP WARNING: Create a 'templates/index.html' file to run the web app.", file=sys.stderr)
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!", file=sys.stderr)
        
    print("\nStarting OCR Flask Application...")
    print("API Key Status:", "LOADED" if API_KEY else "MISSING")
    print("Navigate to http://127.0.0.1:5000/")
    app.run(debug=True)