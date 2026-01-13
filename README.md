# Navneet Text Extraction Tool

A powerful web application built with Flask that leverages Google's Gemini 2.0 Flash Lite model to perform advanced OCR and data extraction from images and PDFs.

## Features

- **Handwriting OCR (Transcribe Mode)**:
  - Transcribes handwritten text, specialized for Indic scripts (Gujarati, Hindi, Marathi) and English.
  -Preserves layout and paragraph structure.
  - allows downloading the result as a Microsoft Word (`.docx`) file.

- **Structured Data Extraction (Extract Mode)**:
  - Extracts structured data from two-column forms.
  - Identifies "Code", "Subject", and "Quantity" fields.
  - Displays results in an interactive dashboard with "Left Column" and "Right Column" separation.
  - Handles complex layouts, including footer tables.

- **Multi-Format Support**:
  - Supports image files (JPG, PNG, etc.) and multi-page PDFs.
  - Automatically preprocesses images (contrast enhancement, sharpening, resizing) for optimal OCR results.

## Prerequisites

- Python 3.8+
- A Google Gemini API Key

## Installation

1.  **Clone the repository** (if applicable) or navigate to the project directory.

2.  **Install dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

3.  **Environment Setup**:
    - Create a `.env` file in the root directory.
    - Add your Gemini API key:
      ```env
      GEMINI_API_KEY=your_api_key_here
      ```

## Usage

1.  **Start the application**:
    ```bash
    python app.py
    ```

2.  **Open in Browser**:
    Navigate to `http://localhost:5000` in your web browser.

3.  **Tool Operation**:
    - **Upload**: Drag & drop or select an image/PDF file.
    - **Select Mode**: Choose between "Handwriting OCR" or "Code & Qty Extraction".
    - **Process**: Click "Start Extraction" to view results or download the transcription.

## Technologies Used

- **Backend**: Flask, Python
- **AI/ML**: Google Gemini 2.0 Flash Lite (`google-genai`)
- **Image Processing**: Pillow (PIL), PyMuPDF (`fitz`)
- **Document Generation**: `python-docx`
- **Frontend**: HTML5, CSS3, JavaScript (Vanilla)
