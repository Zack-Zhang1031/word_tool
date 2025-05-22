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
# word\_tool 英文词表提取与翻译助手

## 项目简介

word\_tool 是一款基于 PySimpleGUI 的图形界面工具，可自动从 PDF、TXT、Excel 文件中提取英文单词（支持图片型 PDF 的 OCR 识别），对照自定义词表去重，过滤无效词、短词、黑名单词，并自动批量翻译为中文。支持导出结果为 Excel。整个过程无需编程基础，一键可用，极适合英语学习者和词表整理者。

---

## 功能特性

* 支持 PDF（文本型和图片型）、TXT、Excel 文件的英文单词自动提取
* 对图片型 PDF 自动 OCR 识别，无需手动转换
* 支持自定义已学单词词表（TXT 或 Excel），自动去除已学词
* 过滤无效、短词和黑名单词，确保词表纯净
* 批量调用 Google 翻译，输出单词中文释义
* 结果一键导出为 Excel
* 简单易用的 GUI 图形界面

---

## 环境准备

### 1. Python 依赖安装

请确保已安装 Python 3.8 及以上版本。推荐使用虚拟环境（如 venv 或 conda）。

先安装 Python 包依赖：

```bash
pip install -r requirements.txt
```

requirements.txt 示例内容：

```
PySimpleGUI
pandas
PyMuPDF
pdf2image
pytesseract
googletrans==4.0.0-rc1
```

---

### 2. 外部依赖程序（必须手动下载安装）

**注意：下列依赖不能用 pip 安装，需手动下载、解压至指定位置。**

#### a) Tesseract-OCR 安装与配置

* 官网主页：[https://github.com/tesseract-ocr/tesseract](https://github.com/tesseract-ocr/tesseract)
* Windows 64 位安装包推荐第三方维护版（原生项目不维护 Windows 安装包）：

  * [UB Mannheim Tesseract at GitHub Releases](https://github.com/UB-Mannheim/tesseract/wiki)
  * 直接下载最新的 Windows 版安装包，例如 `tesseract-ocr-w64-setup-5.3.3.20231005.exe`
* 双击安装，安装路径可自定义
* 安装后，把整个 `Tesseract-OCR` 文件夹（内含 tesseract.exe）复制到 word\_tool 主程序同级目录
* 如果解压后没有文件夹，自己新建 `Tesseract-OCR` 文件夹并将所有内容放进去即可。

**Mac/Linux 用户**：

* Mac：可以用 Homebrew 安装：

  ```bash
  brew install tesseract
  ```
* Ubuntu/Debian：

  ```bash
  sudo apt-get install tesseract-ocr
  ```

#### b) Poppler 安装与配置

* 官网主页：[https://poppler.freedesktop.org/](https://poppler.freedesktop.org/)
* Windows 推荐下载 [Poppler for Windows Releases](https://github.com/oschwartz10612/poppler-windows/releases)

  * 进入页面，下载最新的压缩包（例如 `poppler-23.11.0/Library/bin`）
* 解压整个 `poppler` 文件夹到 word\_tool 主程序同级目录

  * 关键子路径为 `poppler/Library/bin/pdftoppm.exe` （和其它 bin 工具）
* 解压后结构如下：

```
word_tool/
├── word_tool.py
├── requirements.txt
├── Tesseract-OCR/
│   └── tesseract.exe
├── poppler/
│   └── Library/bin/pdftoppm.exe
```

**Mac/Linux 用户**：

* Mac：

  ```bash
  brew install poppler
  ```
* Ubuntu/Debian：

  ```bash
  sudo apt-get install poppler-utils
  ```

---

## 快速开始

1. 确认上述所有依赖已正确放置和安装。
2. 运行：

   ```bash
   python word_tool.py
   ```

   或双击打包好的 word\_tool.exe（如有）。
3. 按界面提示操作，选择原始文件、词表、输出文件名，即可一键生成结果。

---

## 常见问题

* Tesseract/Poppler 路径找不到？请确保文件夹结构正确，exe 文件实际存在。
* 翻译失败？可能是网络问题或 Google API 限制。
* EXE 打包版运行慢？首次启动解压依赖较慢，属正常现象。

---

## 许可证

本项目基于 MIT 协议开源，详见 LICENSE 文件。

---

## 致谢

* [PySimpleGUI](https://github.com/PySimpleGUI/PySimpleGUI)
* [pytesseract](https://github.com/madmaze/pytesseract)
* [pdf2image](https://github.com/Belval/pdf2image)
* [PyMuPDF](https://github.com/pymupdf/PyMuPDF)
* [googletrans](https://github.com/ssut/py-googletrans)

