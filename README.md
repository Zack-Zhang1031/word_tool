# word_tool

A GUI tool for extracting unique English words from PDF, TXT, or Excel files (including image-based PDFs with OCR), filtering out known vocabulary using your own word list, removing invalid/short/blacklisted words, and automatically translating unfamiliar words into Chinese. Output is saved as an Excel file. No coding required!

---

## Features

- Supports PDF (both text and scanned/image-based), TXT, and Excel files as input.
- Automatic OCR recognition for image-based PDFs (requires [Tesseract-OCR](https://github.com/tesseract-ocr/tesseract) and [Poppler](https://github.com/oschwartz10612/poppler-windows/releases)).
- Filter out already learned words using a custom vocabulary list (TXT/Excel).
- Remove invalid, too short, or blacklisted words.
- Batch translate unfamiliar words into Chinese (using Google Translate).
- Export results to Excel.
- Easy-to-use GUI, no coding required.

---

## Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/Zack-Zhang1031/word_tool.git
cd word_tool
