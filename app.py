import os
import io
import json
from google import genai
from PIL import Image, ImageEnhance
from docx import Document
from docx.shared import Pt
from docx.enum.section import WD_SECTION
from flask import Flask, request, render_template, send_file, jsonify
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
MODEL_ID = "gemini-2.0-flash-lite" 

# --------------------------------------------------
# PROMPTS
# --------------------------------------------------
TRANSCRIPTION_PROMPT = """
{
  "ROLE": "Expert Polyglot Scribe and Paleographer",
  "TASK": "Character-by-character transcription of the provided handwriting.",
  "CONTEXT": "The image contains handwritten text in an Indic script (likely Gujarati, Hindi, or Marathi) or English.",
  "COGNITIVE_RULES": [
    "Identify the script visually (e.g., Look for the Shirorekha line for Hindi/Marathi, or curved loops for Gujarati).",
    "STRICT SCRIPT ISOLATION: Do not output English/Latin characters unless the word is explicitly written in that alphabet.",
    "VERBATIM: Transcribe exactly what is written, preserving line breaks and paragraph structure."
  ],
  "OUTPUT_FORMAT": "Return ONLY the transcribed text in its original script. No meta-data or commentary."
}
"""

STRUCTURED_EXTRACT_PROMPT = """
Extract the handwritten order quantities from this Navneet Indent Form.
The output MUST be a valid JSON list of objects. 

Return ONLY raw JSON. No markdown formatting.

Structure:
[
  {
    "table_title": "Standard/Grade Name (e.g. Std 9)",
    "items": [
      {"code": "N 101", "qty": "25", "subject": "Gujarati"}
    ]
  }
]
RULES:
1. Scan for the 'CODE' and 'QTY' columns.
2. Only include rows where a quantity is handwritten.
3. If the handwritten number is unclear, use the surrounding context to provide the best estimate.
"""

# --------------------------------------------------
# IMAGE PREPROCESSING
# --------------------------------------------------
def preprocess_image(image: Image.Image) -> Image.Image:
    if image.mode != "RGB":
        image = image.convert("RGB")
    target_width = 2400 
    ratio = target_width / float(image.width)
    new_height = int(float(image.height) * float(ratio))
    image = image.resize((target_width, new_height), Image.Resampling.LANCZOS)
    image = ImageEnhance.Contrast(image).enhance(1.8)
    image = ImageEnhance.Sharpness(image).enhance(1.5)
    return image

# --------------------------------------------------
# DOCUMENT PROCESSING LOGIC
# --------------------------------------------------
@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    if not API_KEY:
        return "API Key Missing", 500
    if "file" not in request.files:
        return "No file part", 400

    file = request.files["file"]
    mode = request.form.get("mode", "transcribe")
    path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(path)

    try:
        client = genai.Client(api_key=API_KEY)
        prompt = TRANSCRIPTION_PROMPT if mode == "transcribe" else STRUCTURED_EXTRACT_PROMPT
        
        results = []
        
        # Handle PDF (Multi-page)
        if path.lower().endswith(".pdf"):
            pdf = fitz.open(path)
            matrix = fitz.Matrix(3, 3) 
            for page_index in range(len(pdf)):
                page = pdf.load_page(page_index)
                pix = page.get_pixmap(matrix=matrix)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                img = preprocess_image(img)
                response = client.models.generate_content(model=MODEL_ID, contents=[prompt, img])
                if response.text: results.append(response.text.strip())
            pdf.close()
        # Handle Image
        else:
            with Image.open(path) as img:
                img = preprocess_image(img)
                response = client.models.generate_content(model=MODEL_ID, contents=[prompt, img])
                if response.text: results.append(response.text.strip())

        # MODE 1: STRUCTURED (UI TABLES)
        if mode == "extract":
            final_json = []
            for r in results:
                try:
                    clean = r.replace("```json", "").replace("```", "").strip()
                    final_json.extend(json.loads(clean))
                except: continue
            return jsonify(final_json)

        # MODE 2: TRANSCRIPTION (WORD DOC)
        else:
            doc = Document()
            style = doc.styles['Normal']
            style.font.name = 'Arial Unicode MS'
            style.font.size = Pt(12)
            
            for i, text in enumerate(results):
                if i > 0: doc.add_section(WD_SECTION.NEW_PAGE)
                doc.add_paragraph(text)
            
            out_io = io.BytesIO()
            doc.save(out_io)
            out_io.seek(0)
            return send_file(out_io, as_attachment=True, download_name=f"{os.path.splitext(file.filename)[0]}_OCR.docx")

    except Exception as e:
        return str(e), 500
    finally:
        if os.path.exists(path): os.remove(path)

if __name__ == "__main__":
    app.run(debug=True, port=5000)