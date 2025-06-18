# ✅ conversion_utils.py
import os
import re
import pandas as pd
from docx.shared import Pt
from docx.oxml.ns import qn
__all__ = [
    "read_data_auto",
    "col_letter_to_index",
    "prepare_assign_map",
    "replace_placeholders",
    "save_doc_with_name",
    ]

def read_data_auto(filepath):
    try:
        with open(filepath, 'rb') as f:
            head = f.read(1024)
    except Exception as e:
        raise ValueError(f"無法開啟檔案: {e}")

    try:
        if head.startswith(b'\xD0\xCF\x11\xE0'):
            return pd.read_excel(filepath, dtype=str, engine='xlrd')
        elif head.startswith(b'PK\x03\x04'):
            return pd.read_excel(filepath, dtype=str, engine='openpyxl')
        elif b'<html' in head.lower() or b'<!doctype' in head.lower():
            return pd.read_html(filepath, encoding='utf-8')[0]
        else:
            return pd.read_csv(filepath, dtype=str, encoding='utf-8')
    except Exception as e:
        raise ValueError(f"讀取檔案失敗: {e}")

def col_letter_to_index(letter):
    return ord(letter.upper()) - ord('A')


def prepare_assign_map(replacements, doc):
    assign_map = {}
    for key, raw_val in replacements.items():
        val = raw_val.strip() if raw_val else ''
        assign_map[key] = val
    return assign_map


def replace_placeholders(doc, assign_map, font_size_rules=None):
    if font_size_rules is None:
        font_size_rules = [(8, 22), (20, 18), (9999, 12)]

    pattern = re.compile(r"{{\s*([a-zA-Z]+)\s*}}")

    for para in doc.paragraphs:
        runs = para.runs
        if not runs:
            continue

        # 合併段落所有 run 的文字，並記錄每個 run 的 (start_idx, end_idx, run_index)
        full_text = ''
        run_positions = []  # (start_idx, end_idx, run_idx)
        idx = 0
        for i, run in enumerate(runs):
            text = run.text
            start_idx = idx
            end_idx = start_idx + len(text)
            run_positions.append((start_idx, end_idx, i))
            full_text += text
            idx = end_idx

        # 找到所有佔位符（從後往前處理避免索引錯亂）
        for match in reversed(list(pattern.finditer(full_text))):
            key = match.group(1).upper()
            replacement = assign_map.get(key, '')

            start, end = match.start(), match.end()

            # 找出被佔位符覆蓋的 runs
            affected = [r for r in run_positions if not (r[1] <= start or r[0] >= end)]
            if not affected:
                continue

            first_run_idx = affected[0][2]

            # 先將第一個 run 的文字替換成替換文字
            runs[first_run_idx].text = replacement

            # 其他被佔位符覆蓋的 run 文字清空
            for _, _, r_idx in affected[1:]:
                runs[r_idx].text = ''

            # 設定字體大小及字型
            length = len(replacement)
            font_size = Pt(12)
            for max_len, size_pt in font_size_rules:
                if length <= max_len:
                    font_size = Pt(size_pt)
                    break

            run = runs[first_run_idx]
            run.font.size = font_size
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def save_doc_with_name(doc, folder, filename):
    os.makedirs(folder, exist_ok=True)
    output_path = os.path.join(folder, filename)
    doc.save(output_path)
    return output_path
