import os
import fitz
from google import genai
from PIL import Image
from docx import Document
from docx.enum.section import WD_SECTION
from flask import Flask, request, render_template, send_file
import io

# --- ENVIRONMENT SETUP ---
# Load environment variables from .env file if it exists (for LOCAL development only)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass # python-dotenv is not strictly required for production
    

# --- FLASK SETUP ---
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = './uploads'
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# --- GEMINI CONFIGURATION ---
# IMPORTANT: This loads the key from the environment (Render UI or local .env)
API_KEY = os.environ.get("GEMINI_API_KEY") 
MODEL_ID = "gemini-2.5-flash-lite" 

OCR_PROMPT = (
    "Extract all text from this single page image, focusing on clarity and document structure. "
    "Do not include any formatting tags like <u>, <s>, or any HTML/Markdown. "
    "Clean the extracted text by intelligently inserting appropriate line breaks, "
    "bullet points (for numbered lists), and proper spacing to make the content "
    "look like a well-formatted, clean document. "
    "Ensure headings are separate from paragraphs."
)
# --- END GEMINI CONFIGURATION ---


# --- CORE PROCESSING FUNCTION (Remains the same as the final version) ---
def process_document(input_file_path, prompt, client):
    document = Document()
    temp_images = []
    pdf_document = None  # Initialize to None for cleanup safety
    
    try:
        if input_file_path.lower().endswith(('.pdf')):
            pdf_document = fitz.open(input_file_path) 
            num_pages = len(pdf_document)
            
            for i in range(num_pages):
                page = pdf_document.load_page(i)
                pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72)) 
                temp_image_path = os.path.join(app.config['UPLOAD_FOLDER'], f"temp_page_{i+1}.png")
                pix.save(temp_image_path)
                temp_images.append(temp_image_path)
            
            pages_to_process = temp_images
            
        elif input_file_path.lower().endswith(('.jpg', '.jpeg', '.png')):
            pages_to_process = [input_file_path]
            
        else:
            raise ValueError("Unsupported file type.")

        # GEMINI PROCESSING LOOP
        for i, page_path in enumerate(pages_to_process):
            page_number = i + 1
            
            with Image.open(page_path) as page_image:
                response = client.models.generate_content(
                    model=MODEL_ID, 
                    contents=[prompt, page_image]
                )
            
            extracted_text = response.text
            
            document.add_paragraph(f"\n--- Page {page_number} ---")
            document.add_paragraph(extracted_text)
            
            if len(pages_to_process) > 1 and page_number < len(pages_to_process):
                document.add_section(WD_SECTION.NEW_PAGE)
                
        doc_io = io.BytesIO()
        document.save(doc_io)
        doc_io.seek(0)
        return doc_io

    finally:
        if pdf_document is not None:
            pdf_document.close()
            
        for temp_img in temp_images:
            if os.path.exists(temp_img):
                try:
                    os.remove(temp_img)
                except PermissionError as e:
                    print(f"Warning: Could not delete temporary image {temp_img}. {e}")

        if os.path.exists(input_file_path):
            try:
                os.remove(input_file_path)
            except PermissionError as e:
                print(f"Warning: Could not delete original file {input_file_path}. {e}")


# --- FLASK ROUTES ---

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_and_convert():
    if not API_KEY:
        return "Server not configured: API key is missing.", 500
        
    if 'file' not in request.files:
        return 'No file part', 400
    
    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400
    
    if file:
        filename = file.filename
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        try:
            # Initialize client here using the secure API_KEY
            gemini_client = genai.Client(api_key=API_KEY)
            doc_stream = process_document(filepath, OCR_PROMPT, gemini_client)
            
            return send_file(
                doc_stream,
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                as_attachment=True,
                download_name='temp_output.docx' 
            )
        
        except ValueError as e:
            if os.path.exists(filepath):
                try: os.remove(filepath)
                except: pass
            return str(e), 400
        except Exception as e:
            if os.path.exists(filepath):
                try: os.remove(filepath)
                except: pass
            return f"An internal error occurred: {e}", 500

if __name__ == '__main__':
    print("Running Flask app. Navigate to http://127.0.0.1:5000/")
    # If the API_KEY is loaded from .env/environment, we can run safely
    app.run(debug=True)