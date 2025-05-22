# word\_tool - English Vocabulary Extractor & Translator

## Project Overview

**word\_tool** is a GUI application (based on PySimpleGUI) for extracting unique English words from PDF, TXT, or Excel files (with OCR support for image-based PDFs), removing already learned words using a custom word list, filtering out invalid/short/blacklisted words, and batch translating unfamiliar words into Chinese. Results can be exported to Excel. No programming experience required.

---

## Features

* Supports extraction from text-based and image-based PDF, TXT, and Excel files
* Automatic OCR for image/scanned PDFs (no manual conversion required)
* Filter out words already in your vocabulary (custom TXT or Excel word list)
* Remove invalid, too-short, or blacklisted words for a cleaner list
* Batch translation to Chinese via Google Translate
* Export results as an Excel file
* User-friendly GUI interface

---

## Environment Setup

### 1. Python Dependencies

Ensure you have Python 3.8 or later installed. Using a virtual environment (venv or conda) is recommended.

Install the required Python packages:

```bash
pip install -r requirements.txt
```

Sample `requirements.txt`:

```
PySimpleGUI
pandas
PyMuPDF
pdf2image
pytesseract
googletrans==4.0.0-rc1
```

---

### 2. External Dependencies (Manual Download Required)

**Note:** The following dependencies cannot be installed with pip. You must download and extract them manually as described below.

#### a) Tesseract-OCR Installation & Configuration

* Official homepage: [https://github.com/tesseract-ocr/tesseract](https://github.com/tesseract-ocr/tesseract)
* For Windows, use the recommended third-party builds:

  * [UB Mannheim Tesseract at GitHub Releases](https://github.com/UB-Mannheim/tesseract/wiki)
  * Download the latest Windows installer, e.g., `tesseract-ocr-w64-setup-5.3.3.20231005.exe`
* Run the installer and install Tesseract to any directory.
* After installation, copy the entire `Tesseract-OCR` folder (containing `tesseract.exe`) to the same directory as `word_tool.py`.
* If the installer does not create a folder, manually create `Tesseract-OCR` and place all files inside.

**For Mac/Linux:**

* Mac (Homebrew):

  ```bash
  brew install tesseract
  ```
* Ubuntu/Debian:

  ```bash
  sudo apt-get install tesseract-ocr
  ```

#### b) Poppler Installation & Configuration

* Official homepage: [https://poppler.freedesktop.org/](https://poppler.freedesktop.org/)
* For Windows, download the latest release from [Poppler for Windows Releases](https://github.com/oschwartz10612/poppler-windows/releases)

  * Download the latest zip archive (e.g., `poppler-23.11.0/Library/bin`)
* Extract the entire `poppler` folder to the same directory as `word_tool.py`.

  * The key binary should be at `poppler/Library/bin/pdftoppm.exe` (along with other tools)
* After extraction, your directory should look like:

```
word_tool/
├── word_tool.py
├── requirements.txt
├── Tesseract-OCR/
│   └── tesseract.exe
├── poppler/
│   └── Library/bin/pdftoppm.exe
```

**For Mac/Linux:**

* Mac:

  ```bash
  brew install poppler
  ```
* Ubuntu/Debian:

  ```bash
  sudo apt-get install poppler-utils
  ```

---

## Quick Start

1. Make sure all dependencies are installed and all folders are placed as shown above.
2. Run:

   ```bash
   python word_tool.py
   ```

   or double-click `word_tool.exe` (if provided).
3. Follow the GUI prompts to select your source file, vocabulary list, and output file name to generate results with one click.

---

## FAQ

* Tesseract/Poppler path not found? Ensure your directory structure matches the example and that the executable files are present.
* Translation errors? These may be due to network issues or Google API rate limits.
* Slow EXE startup? First launch may be slow as dependencies are unpacked—this is normal.

---

## License

This project is open-sourced under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## Credits

* [PySimpleGUI](https://github.com/PySimpleGUI/PySimpleGUI)
* [pytesseract](https://github.com/madmaze/pytesseract)
* [pdf2image](https://github.com/Belval/pdf2image)
* [PyMuPDF](https://github.com/pymupdf/PyMuPDF)
* [googletrans](https://github.com/ssut/py-googletrans)
