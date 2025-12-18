import os
import io
import sys
import fitz  # PyMuPDF
from google import genai
from PIL import Image, ImageEnhance
from docx import Document
from docx.shared import Pt
from docx.enum.section import WD_SECTION
from flask import Flask, request, render_template, send_file

# --------------------------------------------------
# ENV SETUP
# --------------------------------------------------
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "./tmp/uploads"
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

# --------------------------------------------------
# GEMINI CONFIG & RT-CROS PROMPT
# --------------------------------------------------
API_KEY = os.environ.get("GEMINI_API_KEY")
# Using 2.0 Flash for best vision/speed balance
MODEL_ID = "gemini-2.0-flash" 

# RT-CROS Structured Prompt
OCR_PROMPT = """
{
  "ROLE": "Expert Linguistic Forensic OCR Engine",
  "TASK": "Perform a high-fidelity verbatim transcription of the provided image document.",
  "CONTEXT": "The input is a scanned document containing potentially complex scripts (Gujarati, Devanagari, or English). The goal is to digitize this text for a professional DOCX report without any loss of data or linguistic nuance.",
  "REASON": "The user requires an exact digital replica of the text for archival and editing purposes. Accuracy is paramount; hallucinations or 'corrections' of the original text will fail the mission.",
  "OUTPUT_FORMAT": {
    "TYPE": "Plain Text",
    "RULES": [
      "Preserve exact line breaks and paragraph spacing.",
      "Maintain the original script (do not translate).",
      "Do not include any commentary, headers, or metadata in the output.",
      "Capture all punctuation and special characters exactly as they appear."
    ]
  },
  "STOPPING_CONDITION": "Stop immediately once the last visible character at the bottom-right of the image has been transcribed. Do not add conversational filler."
}
"""

# --------------------------------------------------
# IMAGE PREPROCESSING
# --------------------------------------------------
def preprocess_image(image: Image.Image) -> Image.Image:
    """Enhances image for better OCR legibility."""
    if image.mode != "RGB":
        image = image.convert("RGB")

    # Limit size to prevent OOM but keep high enough for small text
    max_dim = 1800
    if max(image.width, image.height) > max_dim:
        image.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)

    # Moderate Contrast & Sharpness for better glyph recognition
    image = ImageEnhance.Contrast(image).enhance(1.2)
    image = ImageEnhance.Sharpness(image).enhance(1.5)

    return image

# --------------------------------------------------
# CORE OCR LOGIC
# --------------------------------------------------
def process_document(input_path: str, prompt: str, client):
    doc = Document()
    
    # Set default style for the Word Document
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)

    if input_path.lower().endswith(".pdf"):
        pdf = fitz.open(input_path)
        try:
            # 144 DPI is a sweet spot for quality vs memory
            matrix = fitz.Matrix(144 / 72, 144 / 72) 

            for page_index in range(len(pdf)):
                page = pdf.load_page(page_index)
                pix = page.get_pixmap(matrix=matrix)
                
                img_data = io.BytesIO(pix.tobytes("png"))
                img = Image.open(img_data)
                img = preprocess_image(img)

                response = client.models.generate_content(
                    model=MODEL_ID,
                    contents=[prompt, img]
                )

                text = response.text.strip() if response.text else "[Page empty or unreadable]"

                # Add content to Word
                if page_index > 0:
                    doc.add_section(WD_SECTION.NEW_PAGE)
                
                p = doc.add_paragraph()
                p.add_run(f"--- PAGE {page_index + 1} ---").bold = True
                doc.add_paragraph(text)

                del img
                del pix
        finally:
            pdf.close()
    else:
        # Process Single Image
        with Image.open(input_path) as img:
            img = preprocess_image(img)
            response = client.models.generate_content(
                model=MODEL_ID,
                contents=[prompt, img]
            )
            text = response.text.strip() if response.text else "[No text detected]"
            doc.add_paragraph(text)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --------------------------------------------------
# FLASK ROUTES
# --------------------------------------------------
@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    if not API_KEY:
        return "Server Error: Gemini API key is missing from environment variables.", 500

    if "file" not in request.files:
        return "No file uploaded", 400

    file = request.files["file"]
    if not file.filename:
        return "No file selected", 400

    temp_path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(temp_path)

    try:
        # Initialize the GenAI Client
        client = genai.Client(api_key=API_KEY)
        
        # Process and generate the .docx
        doc_stream = process_document(temp_path, OCR_PROMPT, client)

        # Format output filename
        base_name = os.path.splitext(file.filename)[0]
        out_name = f"{base_name}_OCR_Structured.docx"

        return send_file(
            doc_stream,
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print(f"CRITICAL OCR ERROR: {str(e)}", file=sys.stderr)
        return f"Processing failed: {str(e)}", 500

    finally:
        # Clean up uploaded file
        if os.path.exists(temp_path):
            os.remove(temp_path)

if __name__ == "__main__":
    # Ensure tmp directory exists
    os.makedirs("./tmp/uploads", exist_ok=True)
    app.run(debug=True, port=5000)