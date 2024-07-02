import fitz  # PyMuPDF
import re
import unicodedata
import pandas as pd
import os
import requests
import random
import json
from hashlib import md5
import time
import tkinter as tk
from tkinter import filedialog, messagebox

api_error_shown = False
selected_pdf_path = None


def sanitize_filename(filename):
    invalid_chars = r'["*<>?\\\/:|]'
    sanitized_filename = re.sub(invalid_chars, '_', filename)
    sanitized_filename = ''.join(c for c in unicodedata.normalize('NFKD', sanitized_filename)
                                 if not unicodedata.combining(c))
    sanitized_filename = re.sub('_+', '_', sanitized_filename)
    sanitized_filename = sanitized_filename.strip('_')
    return sanitized_filename[:200].strip()


def extract_title_from_pdf(pdf_path):
    document = fitz.open(pdf_path)
    title = ""
    max_font_size = 0

    journal_patterns = [
        r'Journal of Statistical Software',
        r'http://www\.jstatsoft\.org',
    ]

    for page_num in range(min(3, len(document))):
        page = document[page_num]
        blocks = page.get_text("dict")["blocks"]

        for block in blocks:
            if block['type'] == 0:
                for line in block["lines"]:
                    for span in line["spans"]:
                        if any(re.search(pattern, span['text'], re.IGNORECASE) for pattern in journal_patterns):
                            continue
                        if span['size'] > max_font_size and span['origin'][1] < 300:
                            max_font_size = span['size']

        for block in blocks:
            if block['type'] == 0:
                for line in block["lines"]:
                    for span in line["spans"]:
                        if span['size'] == max_font_size and span['origin'][1] < 300:
                            if any(re.search(pattern, span['text'], re.IGNORECASE) for pattern in journal_patterns):
                                continue
                            title += span['text'] + " "

        if title:
            break

    title = re.sub(r'\s+', ' ', title).strip()
    title = re.sub(r'\barXiv:\S+', '', title, flags=re.IGNORECASE).strip()

    if not title:
        title = "No title found"

    sanitized_title = sanitize_filename(title)

    return sanitized_title


def extract_abstract_from_pdf(pdf_path):
    document = fitz.open(pdf_path)
    abstract = ""
    abstract_started = False

    common_sections = [
        "Introduction", "Background", "Related Work", "Methodology", "Methods",
        "Experiment", "Results", "Discussion", "Conclusion", "References",
        "Acknowledgements", "Appendix"
    ]

    section_pattern = r'\n(' + '|'.join(re.escape(section) for section in common_sections) + r')\s*\n'

    for page_num in range(min(5, len(document))):
        page = document[page_num]
        text = page.get_text()

        if not abstract_started:
            match = re.search(r'Abstract\s*(.+)', text, re.DOTALL | re.IGNORECASE)
            if match:
                abstract_started = True
                abstract = match.group(1)
        else:
            abstract += "\n" + text

        if abstract_started:
            end_match = re.search(section_pattern, abstract, re.IGNORECASE)
            if end_match:
                abstract = abstract[:end_match.start()]
                break

    abstract = re.sub(r'\s+', ' ', abstract).strip()

    if len(abstract) > 2100:
        abstract = abstract[:2100] + "..."  # Truncate if too long

    if not abstract:
        abstract = "No abstract found"

    return abstract


def extract_conclusion_from_pdf(pdf_path):
    document = fitz.open(pdf_path)
    conclusion = ""
    conclusion_started = False

    common_sections = [
        "Acknowledgements", "References", "Bibliography", "Appendix",
        "Supplementary Material", "Notes", "Funding", "Declarations",
        "Conflict of Interest", "Author Contributions", "Data Availability",
        "Ethics Statement", "Abbreviations", "A Synthetic Experiments Analysis"
    ]

    section_pattern = r'\n(' + '|'.join(re.escape(section) for section in common_sections) + r')\s*\n'

    for page_num in range(len(document)):
        page = document[page_num]
        text = page.get_text()

        if not conclusion_started:
            match = re.search(r'Conclusion\s*(.+)', text, re.DOTALL)
            if match:
                conclusion_started = True
                conclusion = match.group(1)
        else:
            conclusion += "\n" + text

        if conclusion_started:
            end_match = re.search(section_pattern, conclusion, re.IGNORECASE)
            if end_match:
                conclusion = conclusion[:end_match.start()]
                break

    conclusion = re.sub(r'\s+', ' ', conclusion).strip()

    if not conclusion:
        conclusion = "No conclusion found"

    return conclusion


def get_api_credentials():
    global api_error_shown
    try:
        with open('api.txt', 'r') as file:
            appid = file.readline().strip()
            secretKey = file.readline().strip()
        return appid, secretKey
    except FileNotFoundError:
        if not api_error_shown:
            messagebox.showerror("错误", "api.txt文件未找到。请创建一个包含API的文件。")
            api_error_shown = True
        return None, None
    except Exception as e:
        if not api_error_shown:
            messagebox.showerror("错误", f"读取API时出错: {e}")
            api_error_shown = True
        return None, None


def baidu_translate(text, from_lang='auto', to_lang='zh', domain='electronics'):
    appid, secretKey = get_api_credentials()
    if not appid or not secretKey:
        return None

    url = 'https://fanyi-api.baidu.com/api/trans/vip/fieldtranslate'

    def translate_chunk(chunk):
        salt = random.randint(32768, 65536)
        sign = md5((appid + chunk + str(salt) + domain + secretKey).encode('utf-8')).hexdigest()

        payload = {
            'appid': appid,
            'q': chunk,
            'from': from_lang,
            'to': to_lang,
            'salt': salt,
            'sign': sign,
            'domain': domain
        }

        try:
            response = requests.get(url, params=payload)
            response.raise_for_status()

            result = response.json()

            if 'trans_result' in result:
                return result['trans_result'][0]['dst']
            else:
                print("API Error:", result)
                return None
        except requests.exceptions.RequestException as e:
            print("Request Error:", e)
            return None
        except json.JSONDecodeError as e:
            print("JSON Decode Error:", e)
            print("Response content:", response.text)
            return None

    chunk_size = 2000
    chunks = [text[i:i + chunk_size] for i in range(0, len(text), chunk_size)]

    translated_chunks = [translate_chunk(chunk) for chunk in chunks]
    if None in translated_chunks:
        return None
    return ' '.join(translated_chunks)


def process_pdf(pdf_path):
    title = extract_title_from_pdf(pdf_path)
    abstract = extract_abstract_from_pdf(pdf_path)
    conclusion = extract_conclusion_from_pdf(pdf_path)
    return title, abstract, conclusion


def save_to_excel(pdf_path, excel_path):
    global api_error_shown
    title, abstract, conclusion = process_pdf(pdf_path)

    try:
        df = pd.read_excel(excel_path)
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Title', 'Title_Translated', 'Abstract', 'Abstract_Translated', 'Conclusion',
                                   'Conclusion_Translated'])

    if title in df['Title'].values:
        print(f"Skipping {title} as it already exists in the Excel file.")
        return

    title_translated = baidu_translate(title, domain='electronics')
    abstract_translated = baidu_translate(abstract, domain='electronics')
    conclusion_translated = baidu_translate(conclusion, domain='electronics')

    if title_translated is None or abstract_translated is None or conclusion_translated is None:
        if not api_error_shown:
            messagebox.showerror("错误", "翻译失败。请检查API。")
            api_error_shown = True
        title_translated = title_translated if title_translated is not None else ""
        abstract_translated = abstract_translated if abstract_translated is not None else ""
        conclusion_translated = conclusion_translated if conclusion_translated is not None else ""

    new_record = {
        'Title': title,
        'Title_Translated': title_translated,
        'Abstract': abstract,
        'Abstract_Translated': abstract_translated,
        'Conclusion': conclusion,
        'Conclusion_Translated': conclusion_translated
    }
    new_df = pd.DataFrame([new_record])

    df = df.reset_index(drop=True)
    new_df = new_df.reset_index(drop=True)

    df = pd.concat([df, new_df], ignore_index=True)

    df.to_excel(excel_path, index=False)


def batch_extract():
    pdf_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if not pdf_paths:
        return

    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(script_dir, "ReadPaperRecord_v1.0.xlsx")

    for pdf_path in pdf_paths:
        save_to_excel(pdf_path, excel_path)
    messagebox.showinfo("完成", "批量摘录完成并保存到Excel。")


def retranslate_existing():
    global api_error_shown
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(script_dir, "ReadPaperRecord_v1.0.xlsx")

    try:
        df = pd.read_excel(excel_path)
    except FileNotFoundError:
        messagebox.showerror("错误", "Excel文件未找到。请先进行批量摘录。")
        return

    for index, row in df.iterrows():
        updated = False
        if pd.isna(row['Title_Translated']) or row['Title_Translated'] == "":
            title_translated = baidu_translate(row['Title'], domain='electronics')
            if title_translated is not None:
                df.at[index, 'Title_Translated'] = title_translated
                updated = True
        if pd.isna(row['Abstract_Translated']) or row['Abstract_Translated'] == "":
            abstract_translated = baidu_translate(row['Abstract'], domain='electronics')
            if abstract_translated is not None:
                df.at[index, 'Abstract_Translated'] = abstract_translated
                updated = True
        if pd.isna(row['Conclusion_Translated']) or row['Conclusion_Translated'] == "":
            conclusion_translated = baidu_translate(row['Conclusion'], domain='electronics')
            if conclusion_translated is not None:
                df.at[index, 'Conclusion_Translated'] = conclusion_translated
                updated = True

        if updated and (title_translated is None or abstract_translated is None or conclusion_translated is None):
            if not api_error_shown:
                messagebox.showerror("错误", "翻译失败。请检查API。")
                api_error_shown = True

    df.to_excel(excel_path, index=False)
    messagebox.showinfo("完成", "重新翻译完成并保存到Excel。")


def select_note():
    global selected_pdf_path
    selected_pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if not selected_pdf_path:
        return

    title = extract_title_from_pdf(selected_pdf_path)
    selected_title_label.config(text=title)


def add_note():
    global selected_pdf_path
    if not selected_pdf_path:
        messagebox.showinfo("提示", "请先选择论文。")
        return

    title = extract_title_from_pdf(selected_pdf_path)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(script_dir, "ReadPaperRecord_v1.0.xlsx")

    try:
        df = pd.read_excel(excel_path)
    except FileNotFoundError:
        messagebox.showerror("错误", "Excel文件未找到。请先进行批量摘录。")
        return

    if title not in df['Title'].values:
        messagebox.showinfo("提示", "该PDF尚未摘录，请先进行摘录。")
        return

    note = note_entry.get("1.0", tk.END).strip()
    if not note:
        messagebox.showinfo("提示", "请先输入笔记。")
        return

    row_index = df.index[df['Title'] == title].tolist()[0]

    col_index = 6  # 从第7列开始添加笔记
    while col_index < len(df.columns) and pd.notna(df.iloc[row_index, col_index]):
        col_index += 1

    if col_index >= len(df.columns):
        df[f'Note_{col_index - 5}'] = ""

    df.at[row_index, df.columns[col_index]] = note

    df.to_excel(excel_path, index=False)

    note_entry.delete("1.0", tk.END)  # Clear the text entry


def rename_files():
    pdf_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if not pdf_paths:
        return

    for pdf_path in pdf_paths:
        title = extract_title_from_pdf(pdf_path)
        new_filename = sanitize_filename(title) + '.pdf'
        new_filepath = os.path.join(os.path.dirname(pdf_path), new_filename)
        if pdf_path != new_filepath:
            os.rename(pdf_path, new_filepath)
    messagebox.showinfo("完成", "文件重命名完成。")


def create_gui():
    global note_entry, selected_title_label

    root = tk.Tk()
    root.title("论文PDF工具")
    root.geometry("350x350")

    selected_title_label = tk.Label(root, text="None", wraplength=300, justify="left")
    selected_title_label.pack(pady=10)

    note_entry = tk.Text(root, width=40, height=8)
    note_entry.pack(pady=10)

    upper_button_frame = tk.Frame(root)
    upper_button_frame.pack(pady=5)

    batch_button = tk.Button(upper_button_frame, text="批量摘录", command=batch_extract)
    batch_button.grid(row=0, column=0, padx=5)

    rename_button = tk.Button(upper_button_frame, text="改文件名", command=rename_files)
    rename_button.grid(row=0, column=1, padx=5)

    retranslate_button = tk.Button(upper_button_frame, text="重新翻译", command=retranslate_existing)
    retranslate_button.grid(row=0, column=2, padx=5)

    lower_button_frame = tk.Frame(root)
    lower_button_frame.pack(pady=5)

    select_button = tk.Button(lower_button_frame, text="选择笔记", command=select_note)
    select_button.grid(row=0, column=0, padx=5)

    add_button = tk.Button(lower_button_frame, text="添加笔记", command=add_note)
    add_button.grid(row=0, column=1, padx=5)

    root.mainloop()


if __name__ == "__main__":
    create_gui()
