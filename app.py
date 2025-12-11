import os
import fitz # PyMuPDF for PDF handling
import io
import sys
from google import genai
from PIL import Image
from docx import Document
from docx.enum.section import WD_SECTION

# Note: Ensure all dependencies are installed: pip install google-genai pymupdf Pillow python-docx python-dotenv

# --- ENVIRONMENT SETUP ---
# For local testing, ensure you have a .env file with GEMINI_API_KEY="your_key"
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass
    
# --- FILE PATH AND CONFIGURATION ---

# 1. *** UPDATE YOUR INPUT FILE PATH HERE ***
# Set this to the path of your image or PDF file 
INPUT_FILE_PATH = "GUJ.jpg" 
OUTPUT_FILE_NAME = "ocr_output_final_generalized.docx"

# IMPORTANT: This securely loads the key from the environment (Render UI or local .env)
API_KEY = os.environ.get("GEMINI_API_KEY")
MODEL_ID = "gemini-2.5-flash-lite" 

# GENERALIZED PROMPT: Focuses on Logical flow while strictly preserving multi-element visual structure (like A B / C D options).
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
# --- END CONFIGURATION ---


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
    
    print(f"Starting processing for: {input_file_path}")
    
    try:
        if input_file_path.lower().endswith(('.pdf')):
            # PDF Processing: Convert pages to in-memory image streams
            pdf_document = fitz.open(input_file_path) 
            num_pages = len(pdf_document)
            print(f"Found {num_pages} pages in PDF.")
            
            # Use 150 DPI scaling factor for good quality (150/72 = ~2.08)
            scale_factor = 150 / 72  
            matrix = fitz.Matrix(scale_factor, scale_factor)
            
            for i in range(num_pages):
                page = pdf_document.load_page(i)
                pix = page.get_pixmap(matrix=matrix) 
                png_bytes = pix.tobytes(output='png')
                image_stream = io.BytesIO(png_bytes)
                pages_to_process.append(image_stream)
            
        elif input_file_path.lower().endswith(('.jpg', '.jpeg', '.png')):
            # Image Processing
            pages_to_process = [input_file_path]
            print("Found 1 image file.")
            
        else:
            raise ValueError("Unsupported file type. Please use PDF, JPG, or PNG.")

        # GEMINI PROCESSING LOOP
        for i, page_source in enumerate(pages_to_process):
            page_number = i + 1
            print(f"--- Processing Page {page_number}/{len(pages_to_process)} with Gemini ---")
            
            # Open image from stream (for PDF) or file path (for image)
            with Image.open(page_source) as page_image:
                response = client.models.generate_content(
                    model=MODEL_ID, 
                    contents=[prompt, page_image]
                )
            
            extracted_text = response.text
            
            # ⚠️ ROBUSTNESS CHECK: Ensure the model returned meaningful content
            if not extracted_text or extracted_text.strip() == "":
                print(f"Warning: Gemini returned empty or whitespace-only text for Page {page_number}", file=sys.stderr)
                extracted_text = "\n--- OCR failed to return text for this page. Please review the original image quality. ---"

            
            document.add_paragraph(f"\n--- Page {page_number} ---")
            document.add_paragraph(extracted_text)
            
            # Add a new page break in the DOCX for subsequent pages
            if len(pages_to_process) > 1 and page_number < len(pages_to_process):
                document.add_section(WD_SECTION.NEW_PAGE)
                
        # Save the final DOCX to an in-memory buffer
        doc_io = io.BytesIO()
        document.save(doc_io)
        doc_io.seek(0)
        return doc_io

    finally:
        # Cleanup PDF resource
        if pdf_document is not None:
            pdf_document.close()


# --- MAIN EXECUTION ---
if __name__ == '__main__':
    # Initial Check for API Key
    if not API_KEY:
        print("ERROR: GEMINI_API_KEY is missing. Please set it in your environment variables or in a local .env file.", file=sys.stderr)
        sys.exit(1)
    
    # Initial Check for Input File
    if not os.path.exists(INPUT_FILE_PATH):
        print(f"ERROR: Input file not found at '{INPUT_FILE_PATH}'. Please update the INPUT_FILE_PATH variable.", file=sys.stderr)
        sys.exit(1)

    try:
        # Initialize client
        gemini_client = genai.Client(api_key=API_KEY)
        
        # Process the document
        doc_stream = process_document(INPUT_FILE_PATH, OCR_PROMPT, gemini_client)
        
        # Write the DOCX stream to a file on disk
        with open(OUTPUT_FILE_NAME, 'wb') as f:
            f.write(doc_stream.read())
            
        print(f"\nSUCCESS: OCR completed.")
        print(f"Output saved to: {os.path.abspath(OUTPUT_FILE_NAME)}")
        
    except ValueError as e:
        print(f"\nERROR: File Processing failed: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"\nERROR: An unexpected error occurred: {e}", file=sys.stderr)
        # Attempt to provide more detail on API errors
        if "API_KEY" in str(e):
             print("\nSuggestion: Double-check your GEMINI_API_KEY for correctness and ensure it has not expired.", file=sys.stderr)
        sys.exit(1)