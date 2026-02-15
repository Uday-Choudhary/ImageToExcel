# Image to Excel Pipeline (OCR + Vision AI)

A powerful Python pipeline that converts images (invoices, bills, tables) into structured Excel files. It offers two distinct methods: a high-precision **Vision AI** mode (using Llama 3 via Groq) and a fast, offline **Legacy OCR** mode (using EasyOCR).

## Features

- **Vision AI Mode (Recommended)**:

  - Uses Llama 3 Vision (via Groq API) for superior accuracy.
  - Handles complex layouts, handwritten text, and multi-column tables.
  - Performs logical audits (math validation) on extracted data.
  - Outputs to `output/Extracted_Data_Vision.xlsx`.
- **Legacy OCR Mode**:

  - Uses EasyOCR for completely offline processing.
  - Includes spatial analysis for table structure reconstruction.
  - Best for simple, high-contrast documents.
  - Outputs to `output/Extracted_Data_OCR.xlsx`.
- **Excel Formatting**:

  - Auto-sized columns.
  - Header styling and color-coding.
  - Validation checks for calculated totals.

## Tech Stack

- **Language**: Python 3.10+
- **AI/ML**:
  - [Groq API](https://groq.com/) (Llama 3 Vision)
  - [EasyOCR](https://github.com/JaidedAI/EasyOCR) (PyTorch)
  - [OpenCV](https://opencv.org/) (Image Preprocessing)
- **Data Processing**: Pandas, NumPy
- **Excel Generation**: OpenPyXL

## Setup Guide

### 1. Clone the Repository

```bash
git clone <your-repo-url>
cd ImageToExcel
```

### 2. Create a Virtual Environment

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

*(Note: If `requirements.txt` is missing, install manually: `pip install groq pandas openpyxl easyocr opencv-python-headless`)*

### 4. Configure API Key

Copy the example environment file and add your Groq API key:

```bash
cp .env.example .env
```

Open `.env` and paste your key:

```
GROQ_API_KEY=gsk_...
```

*You can get a free key from [console.groq.com](https://console.groq.com).*

## Usage

### Option 1: Vision AI (Best Quality)

Use this for most documents, especially invoices with complex layouts.

```bash
python run_pipeline.py
```

- **Input**: Images in `input/` folder.
- **Output**: `output/Extracted_Data_Vision.xlsx`

### Option 2: Legacy OCR (Offline)

Use this if you don't have an API key or need offline capability.

```bash
python run_pipeline.py --method ocr
```

- **Input**: Images in `input/` folder.
- **Output**: `output/Extracted_Data_OCR.xlsx`

## Project Structure

```
├── input/                  # Place source images here
├── output/                 # Generated Excel files appear here
├── ocr_data/               # Intermediate OCR JSON data (Legacy Mode)
├── vision_data/            # Intermediate Vision JSON data (Vision Mode)
├── run_pipeline.py         # Main entry point script
├── vision_processor.py     # Llama Vision integration
├── ocr_extraction.py       # EasyOCR integration
├── json_to_excel.py        # Vision JSON -> Excel converter
├── convert_to_excel.py     # OCR JSON -> Excel converter
└── spatial_table_extractor.py # Table logic for Legacy OCR
```
