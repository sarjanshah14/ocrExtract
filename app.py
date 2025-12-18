import os
import io
import sys
import fitz  # PyMuPDF
from google import genai
from PIL import Image, ImageEnhance
from docx import Document
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
# GEMINI CONFIG
# --------------------------------------------------
API_KEY = os.environ.get("GEMINI_API_KEY")
MODEL_ID = "gemini-2.5-flash-lite"

OCR_PROMPT = (
    "You are a visual transcription engine.\n\n"

    "TASK:\n"
    "Transcribe ALL visible text from the ENTIRE image.\n"
    "Continue until no visible text remains.\n\n"

    "LANGUAGE RULES:\n"
    "- Preserve original script exactly.\n"
    "- Gujarati stays Gujarati.\n"
    "- Devanagari stays Devanagari.\n"
    "- English stays English.\n"
    "- Do NOT translate.\n\n"

    "ACCURACY RULES:\n"
    "- Copy characters exactly as seen.\n"
    "- Do NOT guess words.\n"
    "- Do NOT autocorrect spelling.\n"
    "- Do NOT normalize grammar.\n"
    "- If unclear, output closest visible glyph.\n\n"

    "LAYOUT RULES:\n"
    "- Maintain reading order (top to bottom).\n"
    "- Preserve line breaks.\n"
    "- Keep paragraphs separate.\n"
    "- Do NOT merge unrelated text.\n\n"

    "OUTPUT:\n"
    "- Plain text only.\n"
    "- No explanations.\n"
    "- No HTML or Markdown."
)

# --------------------------------------------------
# IMAGE PREPROCESSING (VERY SAFE)
# --------------------------------------------------
def preprocess_image(image: Image.Image) -> Image.Image:
    if image.mode != "RGB":
        image = image.convert("RGB")

    # Keep resolution LOW to avoid hallucination + OOM
    max_width = 1400
    if image.width > max_width:
        ratio = max_width / image.width
        image = image.resize(
            (max_width, int(image.height * ratio)),
            Image.BICUBIC
        )

    # Extremely light enhancement
    image = ImageEnhance.Contrast(image).enhance(1.1)
    image = ImageEnhance.Sharpness(image).enhance(1.02)

    return image

# --------------------------------------------------
# CORE OCR (OOM SAFE)
# --------------------------------------------------
def process_document(input_path: str, prompt: str, client):
    document = Document()

    if input_path.lower().endswith(".pdf"):
        pdf = fitz.open(input_path)

        try:
            matrix = fitz.Matrix(110 / 72, 110 / 72)  # LOW DPI

            for page_index in range(len(pdf)):
                page = pdf.load_page(page_index)

                pix = page.get_pixmap(matrix=matrix)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                img = preprocess_image(img)

                response = client.models.generate_content(
                    model=MODEL_ID,
                    contents=[prompt, img]
                )

                text = response.text or "[NO TEXT RETURNED]"

                document.add_paragraph(f"\n--- Page {page_index + 1} ---")
                document.add_paragraph(text)

                if page_index < len(pdf) - 1:
                    document.add_section(WD_SECTION.NEW_PAGE)

                # 🔥 FREE MEMORY
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

            text = response.text or "[NO TEXT RETURNED]"
            document.add_paragraph(text)

    output = io.BytesIO()
    document.save(output)
    output.seek(0)
    return output

# --------------------------------------------------
# ROUTES
# --------------------------------------------------
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
    if not file.filename:
        return "Empty filename", 400

    path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(path)

    try:
        client = genai.Client(api_key=API_KEY)
        doc_stream = process_document(path, OCR_PROMPT, client)

        out_name = file.filename.rsplit(".", 1)[0] + "_OCR.docx"
        return send_file(
            doc_stream,
            as_attachment=True,
            download_name=out_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        print("OCR ERROR:", e, file=sys.stderr)
        return "OCR failed. Check server logs.", 500

    finally:
        if os.path.exists(path):
            try:
                os.remove(path)
            except Exception:
                pass
