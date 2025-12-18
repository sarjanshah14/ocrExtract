import os
import fitz  # PyMuPDF
import io
import sys
from google import genai
from PIL import Image, ImageEnhance
from docx import Document
from docx.enum.section import WD_SECTION
from flask import Flask, request, render_template, send_file

# ----------------------------------------
# ENV SETUP
# ----------------------------------------
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "./tmp/uploads"
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

# ----------------------------------------
# GEMINI CONFIG
# ----------------------------------------
API_KEY = os.environ.get("GEMINI_API_KEY")
MODEL_ID = "gemini-2.5-flash-lite"

OCR_PROMPT = (
    "You are a literal Optical Character Recognition (OCR) engine.\n\n"

    "LANGUAGE & SCRIPT:\n"
    "- Detect language strictly from the script.\n"
    "- Gujarati script → Gujarati output.\n"
    "- Devanagari → keep as-is (Hindi/Marathi).\n"
    "- English stays English.\n"
    "- DO NOT translate.\n\n"

    "ACCURACY RULES:\n"
    "- Output exactly what is visible.\n"
    "- DO NOT guess words.\n"
    "- DO NOT autocorrect spelling.\n"
    "- DO NOT normalize grammar.\n"
    "- If unclear, reproduce visible characters only.\n\n"

    "LAYOUT:\n"
    "- Maintain reading order (top to bottom).\n"
    "- New line for each question, instruction, or option.\n"
    "- ONE blank line between unrelated blocks.\n\n"

    "OUTPUT:\n"
    "- Plain text only.\n"
    "- No explanations.\n"
    "- No HTML or Markdown."
)

# ----------------------------------------
# IMAGE PREPROCESSING (CONSERVATIVE)
# ----------------------------------------
def preprocess_image(image: Image.Image) -> Image.Image:
    if image.mode != "RGB":
        image = image.convert("RGB")

    if image.width > 1600:
        ratio = 1600 / image.width
        image = image.resize(
            (1600, int(image.height * ratio)),
            Image.BICUBIC
        )

    image = ImageEnhance.Contrast(image).enhance(1.15)
    image = ImageEnhance.Sharpness(image).enhance(1.05)
    return image

# ----------------------------------------
# CORE OCR FUNCTION
# ----------------------------------------
def process_document(input_path, prompt, client):
    document = Document()
    pdf_doc = None
    pages = []

    try:
        if input_path.lower().endswith(".pdf"):
            pdf_doc = fitz.open(input_path)
            matrix = fitz.Matrix(150 / 72, 150 / 72)

            for i in range(len(pdf_doc)):
                page = pdf_doc.load_page(i)
                pix = page.get_pixmap(matrix=matrix)
                pages.append(io.BytesIO(pix.tobytes("png")))
        else:
            pages = [input_path]

        for idx, src in enumerate(pages, start=1):
            with Image.open(src) as img:
                img = preprocess_image(img)

                response = client.models.generate_content(
                    model=MODEL_ID,
                    contents=[prompt, img]
                )

            text = response.text or "[NO TEXT RETURNED]"

            document.add_paragraph(f"\n--- Page {idx} ---")
            document.add_paragraph(text)

            if idx < len(pages):
                document.add_section(WD_SECTION.NEW_PAGE)

        output = io.BytesIO()
        document.save(output)
        output.seek(0)
        return output

    finally:
        if pdf_doc:
            pdf_doc.close()
        if os.path.exists(input_path):
            try:
                os.remove(input_path)
            except Exception:
                pass

# ----------------------------------------
# ROUTES
# ----------------------------------------
@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    if not API_KEY:
        return "GEMINI_API_KEY not set", 500

    if "file" not in request.files:
        return "No file uploaded", 400

    file = request.files["file"]
    if file.filename == "":
        return "Empty filename", 400

    path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(path)

    try:
        client = genai.Client(api_key=API_KEY)
        doc = process_document(path, OCR_PROMPT, client)

        out_name = file.filename.rsplit(".", 1)[0] + "_OCR.docx"
        return send_file(
            doc,
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print("OCR ERROR:", e, file=sys.stderr)
        return "OCR failed. Check server logs.", 500
