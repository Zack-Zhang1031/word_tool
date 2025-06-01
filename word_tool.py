import PySimpleGUI as sg
import pandas as pd
import os
import re
from pdf2image import convert_from_path
import pytesseract
import fitz  # PyMuPDF
from googletrans import Translator
import nltk

# ===== 自动检测/下载 nltk 依赖 =====
def safe_nltk_download(resource):
    try:
        nltk.data.find(resource)
    except LookupError:
        nltk.download(resource.split('/')[-1])

safe_nltk_download('corpora/names')
safe_nltk_download('corpora/words')

from nltk.corpus import names, words as nltk_words

# ===== 基础词库 =====
ENGLISH_NAMES = set(n.lower() for n in names.words())
VALID_WORDS = set(w.lower() for w in nltk_words.words())

PLACE_WORDS = {'london', 'china', 'paris', 'tokyo', 'america', 'india', 'beijing', 'shanghai', 'japan', 'europe', 'africa', 'canada', 'germany', 'france', 'spain', 'russia'}
FILLER_WORDS = {'well', 'uh', 'um', 'hmm', 'like', 'so', 'actually', 'right', 'okay', 'just', 'really', 'basically', 'oh', 'hey'}
USELESS_WORDS = {'etc', 'et', 'al', 'aa', 'bb', 'cc', 'zz', 'rn', 'oo'}

# ====== 工具函数 ======
def get_exe_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

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
    except Exception:
        return ""

def extract_unique_words(text):
    words = re.findall(r"[a-zA-Z]{3,}", text)
    words = [w.lower() for w in words]
    return sorted(set(words))

def load_learned_words(filepath):
    ext = os.path.splitext(filepath)[-1].lower()
    words = set()
    if ext == ".txt":
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
    blacklist = {"aa", "qi", "xi", "za", "zz", "xx", "ll", "rn"}
    valid_words = []
    for w in all_words:
        if (
            w.isalpha() and len(w) >= 3
            and w not in learned_words
            and w not in blacklist
            and w in VALID_WORDS
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

def filter_custom(words, translations, opts):
    result_words, result_trans = [], []
    for w, t in zip(words, translations):
        if opts['rm_name'] and w in ENGLISH_NAMES:
            continue
        if opts['rm_place'] and w in PLACE_WORDS:
            continue
        if opts['rm_filler'] and w in FILLER_WORDS:
            continue
        if opts['rm_failed'] and t == '翻译失败':
            continue
        if opts['rm_useless'] and (len(w) <= 2 or w in USELESS_WORDS):
            continue
        result_words.append(w)
        result_trans.append(t)
    return result_words, result_trans

# ====== PySimpleGUI 主界面 ======

sg.theme('SystemDefault')
layout = [
    [sg.Text('英文词表提取&翻译助手', font=('微软雅黑', 18), justification='center')],
    [sg.Text('请选择PDF/txt/xlsx原始文件:'), sg.InputText(key='src', readonly=True, size=(50,1)), sg.FileBrowse()],
    [sg.Text('请选择你的词表(已学单词):'), sg.InputText(key='dict', readonly=True, size=(50,1)), sg.FileBrowse()],
    [sg.Text('输出目录:'), sg.InputText(key='outdir', readonly=True, size=(40,1)), sg.FolderBrowse(),
     sg.Text('输出Excel文件名:'), sg.InputText('result.xlsx', key='out')],
    [sg.Checkbox('去除人名', default=True, key='rm_name'),
     sg.Checkbox('去除地名', default=True, key='rm_place'),
     sg.Checkbox('去除语气词', default=True, key='rm_filler'),
     sg.Checkbox('去除翻译失败的词', default=True, key='rm_failed'),
     sg.Checkbox('去除无意义短词', default=True, key='rm_useless')],
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
        out_dir = values['outdir'] or os.getcwd()
        if not (src_file and dict_file and out_file):
            print("⚠️ 请先选择所有文件")
            continue
        save_path = os.path.join(out_dir, out_file)

        print("初始化环境...")
        poppler_path = setup_tesseract_poppler()
        print("加载原始文件...")
        ext = os.path.splitext(src_file)[-1].lower()
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

        # 🌐 输出所有未过滤词及翻译
        all_trans_path = os.path.join(out_dir, 'all_translated_words.xlsx')
        pd.DataFrame({'Word': unfamiliar, 'Chinese': chn}).to_excel(all_trans_path, index=False)
        print(f"🌐 所有未过滤词及翻译已保存到：{all_trans_path}")

        opts = {
            'rm_name': values['rm_name'],
            'rm_place': values['rm_place'],
            'rm_filler': values['rm_filler'],
            'rm_failed': values['rm_failed'],
            'rm_useless': values['rm_useless']
        }
        unfamiliar2, chn2 = filter_custom(unfamiliar, chn, opts)
        print(f"应用自定义过滤后剩余单词 {len(unfamiliar2)} 个")
        result_df = pd.DataFrame({'Word': unfamiliar2, 'Chinese': chn2})
        result_df.to_excel(save_path, index=False)
        print(f"✅ 已完成！文件保存为 {save_path}")
        sg.popup("处理完成！", f"文件已保存：{save_path}")

window.close()
