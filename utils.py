import json
import pdfplumber

MAPPING = json.load(open("mapping.json", "r", encoding='utf-8'))

# ------------ Global Tools ------------ #

def pretty_save_json(file: str, data: dict) -> None:
    with open(file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def to_float(value: str):
    "Convert string to float"
    return float(value.replace(',', ''))

def to_string(value: float):
    "Convert float to string with 2 decimal places"
    return '{:,.2f}'.format(value)

def correct_words(text: str, mapping: dict) -> str:
    for word in mapping:
        text = text.replace(word, mapping[word])
    return text

# ------------ PDF Tools ------------ #

def ie_extract_text(path: str) -> str:
    text = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text += page.extract_text()
        text = correct_words(text, MAPPING)
    return text