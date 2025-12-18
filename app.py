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

API_KEY = os.environ.get("GEMINI_API_KEY")
MODEL_ID = "gemini-2.0-flash" 

# --------------------------------------------------
# REFINED RT-CROS PROMPT (HUMAN-CENTRIC)
# --------------------------------------------------
# This prompt uses a JSON structure to ensure the AI follows specific logic gates.
OCR_PROMPT = """
{
  "ROLE": "Expert Polyglot Scribe and Paleographer",
  "TASK": "Verbatim transcription of handwritten documents in any detected language (Gujarati, Hindi, Marathi, or English).",
  "CONTEXT": "You are reading handwritten documents that may switch between scripts or be written entirely in one regional language. You must act as a human would: recognize the language first, then transcribe the characters according to that specific language's rules.",
  "REASON": "The user needs a digital version of handwritten notes that preserves the original language's integrity without machine-generated errors like mixing Latin letters into Indic scripts.",
  "STRATEGY": {
    "Detection_Logic": "Analyze the visual structure of the characters to identify the script (Devanagari, Gujarati, or Latin).",
    "Verbatim_Rules": [
      "Transcribe exactly what is written. If the text is in Gujarati, use only Gujarati characters.",
      "STRICT ALPHABET BOUNDARY: Do not use English characters to represent phonetics in Indic scripts. Only use English if the author actually wrote in the English alphabet.",
      "Maintain the flow and dialect of the writer without 'correcting' it to formal versions.",
      "If a page contains multiple languages, preserve the transition exactly as it appears in the source."
    ]
  },
  "OUTPUT_FORMAT": "Return only the transcribed text in its original script. No conversational filler, no source tags, and no meta-commentary.",
  "STOPPING_CONDITION": "End output immediately after the last character of the document is transcribed."
}
"""

# --------------------------------------------------
# IMAGE PREPROCESSING (BALANCED)
# --------------------------------------------------
def preprocess_image(image: Image.Image) -> Image.Image:
    if image.mode != "RGB":
        image = image.convert("RGB")

    # Increase resolution slightly for better handwriting recognition
    target_width = 2000 
    ratio = target_width / float(image.width)
    new_height = int(float(image.height) * float(ratio))
    image = image.resize((target_width, new_height), Image.Resampling.LANCZOS)

    # Enhance contrast to separate ink from paper background
    image = ImageEnhance.Contrast(image).enhance(1.4)
    image = ImageEnhance.Sharpness(image).enhance(1.2)

    return image

# --------------------------------------------------
# CORE LOGIC
# --------------------------------------------------
def process_document(input_path: str, prompt: str, client):
    doc = Document()
    
    # Modern Document Styling
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial Unicode MS' # Standard Windows Gujarati Font
    font.size = Pt(12)

    if input_path.lower().endswith(".pdf"):
        pdf = fitz.open(input_path)
        try:
            matrix = fitz.Matrix(2, 2) # Higher resolution render

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
                del img
                del pix
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