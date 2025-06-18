# ✅ conversion_utils.py
import os
import re
import pandas as pd
from docx import Document
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
    placeholder_counts = {}
    pattern_template = r"{{\s*%s\s*}}"
    for key in replacements:
        pattern = re.compile(pattern_template % key, re.IGNORECASE)
        count = 0
        for para in doc.paragraphs:
            for run in para.runs:
                count += len(pattern.findall(run.text))
        placeholder_counts[key] = count

    assign_map = {}
    for key, raw_val in replacements.items():
        vals = [v.strip() for v in raw_val.split(',') if v.strip()] if raw_val else ['']
        n = placeholder_counts.get(key, 0)
        if n == 0:
            assign_map[key] = []
        else:
            if len(vals) < n:
                vals += [''] * (n - len(vals))
            elif len(vals) > n:
                vals = vals[:n-1] + [','.join(vals[n-1:])]
            assign_map[key] = vals
    return assign_map

def replace_placeholders(doc, assign_map, font_size_rules=None):
    """
    font_size_rules: list of tuples (max_length, font_size_pt)
    例如 [(20,22), (25,16), (9999,12)]
    按字數限制選字體大小，第一個符合條件的套用
    """
    if font_size_rules is None:
        font_size_rules = [(20, 22), (25, 16), (9999, 12)]

    pattern_all = re.compile(r"{{\s*([a-zA-Z])\s*}}")
    key_counter = {k: 0 for k in assign_map}

    for para in doc.paragraphs:
        for run in para.runs:
            matches = list(pattern_all.finditer(run.text))
            if not matches:
                continue
            
            new_text = run.text
            for match in matches:
                key = match.group(1).upper()
                val_list = assign_map.get(key, [])
                idx = key_counter.get(key, 0)
                replacement = val_list[idx] if idx < len(val_list) else ''
                key_counter[key] = idx + 1

                # 根據字數選字體大小
                font_size = Pt(12)  # 預設字體大小
                for max_len, size_pt in font_size_rules:
                    if len(replacement) <= max_len:
                        font_size = Pt(size_pt)
                        break

                new_text = new_text.replace(match.group(0), replacement, 1)

            run.text = new_text
            run.font.size = font_size
            run.font.name = '標楷體'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

            # new_text = run.text
            # for match in matches:
            #     key = match.group(1).upper()
            #     val_list = assign_map.get(key, [])
            #     idx = key_counter.get(key, 0)
            #     replacement = val_list[idx] if idx < len(val_list) else ''
            #     key_counter[key] = idx + 1

            #     if len(replacement) <= 20:
            #         font_size = Pt(22)
            #     elif len(replacement) <= 25:
            #         font_size = Pt(16)
            #     else:
            #         font_size = Pt(12)

            #     new_text = new_text.replace(match.group(0), replacement, 1)

            # run.text = new_text
            # run.font.size = font_size
            # run.font.name = '標楷體'
            # run._element.rPr.rFonts.set(qn('w:eastAsia'), '標楷體')

def save_doc_with_name(doc, folder, filename):
    os.makedirs(folder, exist_ok=True)
    output_path = os.path.join(folder, filename)
    doc.save(output_path)
    return output_path


