import fitz  # PyMuPDF, imported as fitz for backward compatibility reasons
import cv2
import pytesseract
import os
import re


png_folder_path = r'C:\Users\38595\Luka Munivrana\Desktop\data_science\images'
pdf_folder_path = r'C:\Users\38595\Luka Munivrana\Desktop\data_science\documents'

table_pattern = r'^(?:\d{1,2}(?:st|nd|rd|th)?(?:\.|\b)|\b(?:[ivxlcdm]+|[IVXLCDM]+)(?:\.|\)))\s.*$'

# Read all pdf file paths
def get_pdf_file_paths(folder_path):
    pdf_file_paths = []
    for file in os.listdir(folder_path):
        if file.lower().endswith('.pdf'):
            pdf_file_path = os.path.join(folder_path, file)
            pdf_file_paths.append(pdf_file_path)
    return pdf_file_paths

# Read all png file paths
def read_png_files(folder_path):
    png_files = []
    for file in os.listdir(folder_path):
        if file.lower().endswith('.png'):
            png_file_path = os.path.join(folder_path, file)
            png_files.append(png_file_path)
    return png_files

# Delete all files in a folder
def delete_all_files(folder_path):
    for file in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file)
        if os.path.isfile(file_path):
            os.remove(file_path)

def process_ocr_result(text):
    lines = [line.strip() for line in text.splitlines()]
    return lines

def process_table_data(table_data):
    processed_data = []
    for line in table_data:
        # Remove specified symbols
        line = re.sub(r'[.,:;_\/\\|\(\)\[\]\-—=]', '', line)
        # Remove first encountered numbers
        line = re.sub(r'^\d+\s+', '', line)
        # Remove numbers with more than a single digit
        line = re.sub(r'\b\d{2,}\b', '', line)
        # Add spaces between numbers and digits
        line = re.sub(r'(?<=\d)(?=[a-zA-Z])|(?<=[a-zA-Z])(?=\d)', ' ', line)
        # Remove unnecessary words
        line = re.sub(r'\b(?:dovoljan|dobar|vrlo|vrlodobar|izvrstan|NEE|IH|N)\b', '', line)
        # Remove standalone o and O
        line = re.sub(r'\b[0Oo]\b', '', line)
        # Remove standalone o and O
        line = re.sub(r'(?<!\S)\s{2,}[a-zA-Z]\s{2,}(?!\S)', '', line)
        # Lowercase l into 1
        line = re.sub(r'\bl\b', '1', line)
        # Remove symbol '
        line = re.sub(r"'", '', line)
        # Replace "{standalone letter}{whitespace}{digit}" pattern with digit
        line = re.sub(r'\b[A-Za-z]\s+(\d)', r'\1', line)
        # Replace "{digit}{whitespace}{standalone letter}" pattern with digit
        line = re.sub(r'(\d)\s+\b[A-Za-z]', r'\1', line)
        # Remove all whitespace characters between line end and grade
        line = re.sub(r'\s+\Z', '', re.sub(r'\s+(?!\s*$)', ' ', line))
        # Remove characters between numbers
        matches = re.findall(r'(\b\d+\b).*?(\b\d+\b)', line)
        if len(matches) != 0:
            line = re.sub(r'(\b\d+\b).*?(\b\d+\b)', r'\1 \2', line)
        # Replace "1 1" with "1"
        line = re.sub(r'\b1\s+1\b', '1', line)
        # Remove all lines that don't end in a number
        line = re.sub(r'^(?![\s\S]*\d\s*$).*?$', '', line)
        if line:
            processed_data.append(line)
        
    return processed_data

all_pdf_files = get_pdf_file_paths(pdf_folder_path)

for pdf_file in all_pdf_files:

    delete_all_files(png_folder_path)
    table_data = []

    doc = fitz.open(pdf_file)

    # PDF file pages to png
    for i, page in enumerate(doc):
        zoom = 1
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix = mat, dpi=300)
        pix.save(f"images/page_{i}.png")

    all_png_files = read_png_files(png_folder_path)

    for png_file in all_png_files:
        processed_text = []
        img = cv2.imread(png_file)
        ocr_result = pytesseract.image_to_string(png_file, lang='hrv')
        ocr_result = re.sub(r'[\/\\\-_—|,:;=]', '', ocr_result)
        processed_text = process_ocr_result(ocr_result)
        print(ocr_result)
        for line in processed_text:
            matches = re.findall(table_pattern, line, re.MULTILINE)
            table_data.extend(matches)
        table_data = process_table_data(table_data)

    delete_all_files(png_folder_path)
    print(table_data)