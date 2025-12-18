import os
import io
from google import genai
from PIL import Image, ImageEnhance
from docx import Document
from docx.shared import Pt
from docx.enum.section import WD_SECTION
from flask import Flask, request, render_template, send_file
import fitz  # PyMuPDF

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

API_KEY = os.environ.get("GEMINI_API_KEY")
MODEL_ID = "gemini-2.0-flash" 

# --------------------------------------------------
# MULTILINGUAL "HUMAN SCRIBE" PROMPT
# --------------------------------------------------
OCR_PROMPT = """
{
  "ROLE": "Expert Polyglot Scribe and Paleographer",
  "TASK": "Character-by-character transcription of the provided handwriting.",
  "CONTEXT": "The image contains handwritten text in an Indic script (likely Gujarati, Hindi, or Marathi) or English. Previous machine attempts incorrectly mixed English letters into Indic words.",
  "COGNITIVE_RULES": [
    "Identify the script visually (e.g., Look for the Shirorekha line for Hindi/Marathi, or curved loops for Gujarati).",
    "STRICT SCRIPT ISOLATION: Do not output English/Latin characters unless the word is explicitly written in that alphabet. Never use English letters to represent Indic phonetics (e.g., don't write 'maa' if the script is 'મા').",
    "CONTEXTUAL RECOGNITION: If a stroke is messy, use the vocabulary of the detected language to resolve the word. Do not 'guess' using the English alphabet.",
    "VERBATIM: Transcribe exactly what is written, preserving line breaks and paragraph structure."
  ],
  "OUTPUT_FORMAT": "Return ONLY the transcribed text in its original script. No meta-data, no 'source' tags, and no commentary."
}
"""

# --------------------------------------------------
# ADVANCED PREPROCESSING FOR HANDWRITING
# --------------------------------------------------
def preprocess_image(image: Image.Image) -> Image.Image:
    if image.mode != "RGB":
        image = image.convert("RGB")

    # Increase resolution (300 DPI equivalent) for finer pen-stroke detail
    target_width = 2400 
    ratio = target_width / float(image.width)
    new_height = int(float(image.height) * float(ratio))
    image = image.resize((target_width, new_height), Image.Resampling.LANCZOS)

    # Aggressive contrast to separate faded ink from paper texture
    image = ImageEnhance.Contrast(image).enhance(1.6)
    image = ImageEnhance.Sharpness(image).enhance(1.5)

    return image

# --------------------------------------------------
# CORE LOGIC
# --------------------------------------------------
def process_document(input_path: str, prompt: str, client):
    doc = Document()
    
    # Use Arial Unicode MS or Shruti (for Windows) to support all Indic scripts
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial Unicode MS' 
    font.size = Pt(12)

    if input_path.lower().endswith(".pdf"):
        pdf = fitz.open(input_path)
        try:
            # Use 3x zoom for PDF-to-Image rendering
            matrix = fitz.Matrix(3, 3) 
            for page_index in range(len(pdf)):
                page = pdf.load_page(page_index)
                pix = page.get_pixmap(matrix=matrix)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                img = preprocess_image(img)

                response = client.models.generate_content(
                    model=MODEL_ID,
                    contents=[prompt, img]
                )

                text = response.text.strip() if response.text else ""
                if page_index > 0:
                    doc.add_section(WD_SECTION.NEW_PAGE)
                
                doc.add_paragraph(text)
        finally:
            pdf.close()
    else:
        with Image.open(input_path) as img:
            img = preprocess_image(img)
            response = client.models.generate_content(
                model=MODEL_ID,
                contents=[prompt, img]
            )
            text = response.text.strip() if response.text else ""
            doc.add_paragraph(text)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    if not API_KEY:
        return "API Key Missing", 500
    if "file" not in request.files:
        return "No file", 400

    file = request.files["file"]
    path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(path)

    try:
        client = genai.Client(api_key=API_KEY)
        doc_stream = process_document(path, OCR_PROMPT, client)
        out_name = f"{os.path.splitext(file.filename)[0]}_OCR.docx"

        return send_file(
            doc_stream,
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        return str(e), 500
    finally:
        if os.path.exists(path):
            os.remove(path)

if __name__ == "__main__":
    app.run(debug=True, port=5000)