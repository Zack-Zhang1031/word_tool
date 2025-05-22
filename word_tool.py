import PySimpleGUI as sg
import pandas as pd
import os
import re
from pdf2image import convert_from_path
import pytesseract
import fitz  # PyMuPDF
from googletrans import Translator

# ===== 工具函数 =====

def get_exe_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

# 配置 Tesseract 和 Poppler 路径
def setup_tesseract_poppler():
    exe_dir = get_exe_dir()
    tess_path = os.path.join(exe_dir, "Tesseract-OCR", "tesseract.exe")
    poppler_path = os.path.join(exe_dir, "poppler", "Library", "bin")
    pytesseract.pytesseract.tesseract_cmd = tess_path
    return poppler_path

def ocr_pdf(pdf_path, poppler_path):
    pages = convert_from_path(pdf_path, poppler_path=poppler_path)
    full_text = ''
    for i, img in enumerate(pages, 1):
        text = pytesseract.image_to_string(img, lang='eng')
        full_text += text + '\n'
    return full_text

def extract_text_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text()
        return text
    except Exception as e:
        return ""

def extract_unique_words(text):
    words = re.findall(r"[a-zA-Z]{3,}", text)  # 长度至少3
    words = [w.lower() for w in words]
    return sorted(set(words))

def load_learned_words(filepath):
    ext = os.path.splitext(filepath)[-1].lower()
    words = set()
    if ext == ".txt":
        # 支持中英混合行，只取英文词
        with open(filepath, 'r', encoding="utf-8-sig") as f:
            for line in f:
                for w in re.findall(r"[a-zA-Z]{3,}", line):
                    words.add(w.lower())
    elif ext in (".xls", ".xlsx"):
        df = pd.read_excel(filepath)
        col = df.columns[0]
        words = set(df[col].astype(str).str.lower())
    return words

def filter_words(all_words, learned_words):
    # 可自定义黑名单
    blacklist = {"aa", "qi", "xi", "za", "zz", "xx", "ll", "rn"}
    valid_words = []
    for w in all_words:
        if (
            w.isalpha() and len(w) >= 3
            and w not in learned_words
            and w not in blacklist
        ):
            valid_words.append(w)
    return valid_words

def batch_translate(words):
    translator = Translator()
    result = []
    for w in words:
        try:
            zh = translator.translate(w, src='en', dest='zh-cn').text
            result.append(zh)
        except Exception:
            result.append("翻译失败")
    return result

# ===== PySimpleGUI 主界面 =====

sg.theme('SystemDefault')
layout = [
    [sg.Text('英文词表提取&翻译助手', font=('微软雅黑', 18), justification='center')],
    [sg.Text('请选择PDF/txt/xlsx原始文件:'), sg.InputText(key='src', readonly=True, size=(50,1)), sg.FileBrowse()],
    [sg.Text('请选择你的词表(已学单词):'), sg.InputText(key='dict', readonly=True, size=(50,1)), sg.FileBrowse()],
    [sg.Text('输出Excel文件名:'), sg.InputText('result.xlsx', key='out')],
    [sg.Button('开始处理', size=(10,1)), sg.Button('退出', size=(8,1))],
    [sg.Output(size=(80,10))]
]

window = sg.Window('词表助手', layout)

while True:
    event, values = window.read()
    if event in (sg.WINDOW_CLOSED, '退出'):
        break
    if event == '开始处理':
        src_file = values['src']
        dict_file = values['dict']
        out_file = values['out']
        if not (src_file and dict_file and out_file):
            print("⚠️ 请先选择所有文件")
            continue

        print("初始化环境...")
        poppler_path = setup_tesseract_poppler()
        print("加载原始文件...")
        ext = os.path.splitext(src_file)[-1].lower()
        # 文本提取
        if ext == ".pdf":
            text = extract_text_pdf(src_file)
            if len(text.strip()) < 20:
                print("正在进行 OCR 识别...")
                text = ocr_pdf(src_file, poppler_path)
        elif ext == ".txt":
            with open(src_file, encoding="utf-8-sig") as f:
                text = f.read()
        elif ext in (".xls", ".xlsx"):
            df = pd.read_excel(src_file)
            col = df.columns[0]
            text = " ".join(str(x) for x in df[col])
        else:
            print("❌ 不支持的文件格式！")
            continue

        words = extract_unique_words(text)
        print(f"共提取唯一英文单词 {len(words)} 个")
        learned = load_learned_words(dict_file)
        print(f"你的词表包含 {len(learned)} 个已学词")
        unfamiliar = filter_words(words, learned)
        print(f"过滤后剩余未学单词 {len(unfamiliar)} 个")
        print("正在自动翻译...")
        chn = batch_translate(unfamiliar)
        result_df = pd.DataFrame({'Word': unfamiliar, 'Chinese': chn})
        result_df.to_excel(out_file, index=False)
        print(f"✅ 已完成！文件保存为 {out_file}")
        sg.popup("处理完成！", f"文件已保存：{out_file}")

window.close()
