import os
import shutil
import subprocess
import sys
import re
import traceback
import json
import pdfplumber
import pdfkit
from PyPDF2 import PdfReader, PdfWriter
from pdf2docx import Converter
from docx import Document
from colorama import Fore, init
from docx2pdf import convert

def move_files(src_dir):
    text_files_dir = os.path.join(src_dir, 'Text files')
    docx_files_dir = os.path.join(src_dir, 'Word files')
    pdf_files_dir = os.path.join(src_dir, 'PDFs')

    os.makedirs(text_files_dir, exist_ok=True)
    os.makedirs(docx_files_dir, exist_ok=True)
    os.makedirs(pdf_files_dir, exist_ok=True)

    files = os.listdir(src_dir)

    for file in files:
        file_path = os.path.join(src_dir, file)
        if os.path.isfile(file_path):
            if file.endswith('.txt'):
                # Rename the text file
                new_file_name = file.split('_')[0] + '.txt'
                new_file_path = os.path.join(text_files_dir, new_file_name)
                shutil.move(file_path, new_file_path)
            elif file.endswith('.docx'):
                # Rename the docx file
                new_file_name = file.split('_')[0] + '.docx'
                new_file_path = os.path.join(docx_files_dir, new_file_name)
                shutil.move(file_path, new_file_path)
            else:
                # Extract the base file name without the extension
                base_file_name = os.path.splitext(file)[0]
                # Remove anything after the underscore (_) character
                new_file_name = base_file_name.split('_')[0] + os.path.splitext(file)[1]
                new_file_path = os.path.join(pdf_files_dir, new_file_name)
                shutil.move(file_path, new_file_path)

def convert_docx_to_pdf(dir_path):
    docx_files = [f for f in os.listdir(dir_path) if f.endswith('.docx')]
    
    for docx_file in docx_files:
        docx_path = os.path.join(dir_path, docx_file)

        # Create a new path for the 'Stripped PDFs' folder
        stripped_dir = os.path.join(dir_path, 'Stripped PDFs')
        if not os.path.exists(stripped_dir):
            os.makedirs(stripped_dir)

        # Now create the full pdf_path with the new directory
        pdf_path = os.path.join(stripped_dir, docx_file[:-5] + '.pdf')

        convert(docx_path, pdf_path)  # convert the docx file to pdf

def install_and_import(package):
    try:
        __import__(package)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        __import__(package)

packages = ['PyPDF2', 'pdf2docx', 'docx', 'colorama', 'pdfplumber']
for package in packages:
    install_and_import(package)

init(autoreset=True)

def decrypt_pdf(input_pdf_path, output_pdf_path, password):
    reader = PdfReader(input_pdf_path)
    if reader.is_encrypted:
        try:
            reader.decrypt(password)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            writer.write(output_pdf_path)
        except NotImplementedError:
            print(Fore.RED + f"Sorry, the encryption used in this PDF: {input_pdf_path} is not supported.")
    else:
        shutil.copy(input_pdf_path, output_pdf_path)

def convert_pdf_to_docx(input_pdf_path, output_docx_path):
    cv = Converter(input_pdf_path)
    cv.convert(output_docx_path, start=0, end=None)
    cv.close()

def convert_pdf_to_txt(input_pdf_path, output_txt_path):
    with pdfplumber.open(input_pdf_path) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
    text = text.replace('\uf04a', '')
    with open(output_txt_path, 'w', encoding='utf-8') as f:
        f.write(text)

def process_pdf_files(dir_path, password):
    pdf_files = [f for f in os.listdir(dir_path) if f.endswith('.pdf')]
    for pdf_file in pdf_files:
        input_pdf_path = os.path.join(dir_path, pdf_file)
        output_pdf_path = os.path.join(dir_path, 'decrypted', pdf_file[:-4] + '_decrypted.pdf')
        output_docx_path = os.path.join(dir_path, 'decrypted', pdf_file[:-4] + '.docx')
        output_txt_path = os.path.join(dir_path, 'decrypted', pdf_file[:-4] + '.txt')
        if not os.path.exists(os.path.join(dir_path, 'decrypted')):
            os.makedirs(os.path.join(dir_path, 'decrypted'))
        decrypt_pdf(input_pdf_path, output_pdf_path, password)
        convert_pdf_to_docx(output_pdf_path, output_docx_path)
        convert_pdf_to_txt(output_pdf_path, output_txt_path)

def docx_find_replace_text(doc_obj, find_text, replace_text):
    pattern = re.compile(find_text, re.IGNORECASE)
    for p in doc_obj.paragraphs:
        if pattern.search(p.text):
            inline = p.runs
            for i in range(len(inline)):
                if pattern.search(inline[i].text):
                    replace_text = str(replace_text)
                    text = pattern.sub(replace_text, inline[i].text)
                    inline[i].text = text
    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_find_replace_text(cell, find_text, replace_text)


def process_docx_files(dir_path, replacements):
    docx_files = [f for f in os.listdir(dir_path) if f.endswith('.docx')]
    for docx_file in docx_files:
        docx_path = os.path.join(dir_path, docx_file)
        try:
            doc = Document(docx_path)
            for replacement in replacements:
                find_text = replacement['find_text']
                replace_text = replacement['replace_text']
                docx_find_replace_text(doc, find_text, replace_text)
            doc.save(docx_path)
        except PermissionError as e:
            print(f"{Fore.RED}An error occurred while processing the file:")
            print(f"{e}")

def process_txt_files(dir_path, replacements):
    txt_files = [f for f in os.listdir(dir_path) if f.endswith('.txt')]
    for txt_file in txt_files:
        txt_path = os.path.join(dir_path, txt_file)
        try:
            with open(txt_path, 'r', encoding='utf-8') as f:
                text = f.read()
            for replacement in replacements:
                find_text = replacement['find_text']
                replace_text = replacement['replace_text']
                text = re.sub(find_text, replace_text, text, flags=re.IGNORECASE)
            with open(txt_path, 'w', encoding='utf-8') as f:
                f.write(text)
        except PermissionError as e:
            print(f"{Fore.RED}An error occurred while processing the file:")
            print(f"{e}")

def read_find_replace_from_json(json_path):
    try:
        with open(json_path, 'r') as f:
            replacements = json.load(f)
        return replacements
    except FileNotFoundError:
        print(f"JSON file at path {json_path} was not found. Creating a new one.")
        replacements = [{'find_text': '', 'replace_text': ''}]
        with open(json_path, 'w') as f:
            json.dump(replacements, f)
        return replacements

def main():
    print(Fore.GREEN + "Please choose an action: ")
    print("1. Decrypt/unlock + FindReplace + Export to text")
    print("2. Decrypt/Unlock")
    print("3. FindReplace")
    print("4. Export to text")
    choice = int(input("Enter choice: "))

    dir_path = input(Fore.GREEN + "Please enter the absolute path of the PDFs (leave blank for current directory): ")
    if not dir_path:
        dir_path = '.'

    if not os.path.exists(dir_path):
        print(Fore.RED + f"The directory {dir_path} does not exist.")
        return

    pdf_files = [f for f in os.listdir(dir_path) if f.endswith('.pdf')]

    if not pdf_files:
        print(Fore.RED + f"No PDF files found in the directory: {dir_path}")
        return

    password = input(Fore.GREEN + "Please enter the password (leave blank if password is stored in password.txt): ")
    if not password:
        with open('password.txt', 'r') as f:
            password = f.read().strip()

    find_replace_json = input(Fore.GREEN + "Please enter the path of the Find_Replace JSON (leave blank for 'find_replace.json' in current directory): ")
    if not find_replace_json:
        find_replace_json = 'find_replace.json'
    
    replacements = read_find_replace_from_json(find_replace_json)

    try:
        if choice == 1:
            process_pdf_files(dir_path, password)
            process_docx_files(os.path.join(dir_path, 'decrypted'), replacements)
            process_txt_files(os.path.join(dir_path, 'decrypted'), replacements)
            move_files(os.path.join(dir_path, 'decrypted'))
        elif choice == 2:
            process_pdf_files(dir_path, password)
        elif choice == 3:
            process_docx_files(os.path.join(dir_path, 'decrypted'), replacements)
            process_txt_files(os.path.join(dir_path, 'decrypted'), replacements)
        elif choice == 4:
            process_pdf_files(dir_path, password)
            move_files(os.path.join(dir_path, 'decrypted'))
        else:
            print(Fore.RED + "Invalid choice.")
    except Exception as e:
        print(f"{Fore.RED}An error occurred:")
        print(traceback.format_exc())

if __name__ == "__main__":
    main()
    move_files('decrypted')
    convert_docx_to_pdf('decrypted/Word files')
