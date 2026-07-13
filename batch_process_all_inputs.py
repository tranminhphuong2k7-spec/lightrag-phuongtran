import os
import glob
import zipfile
import tempfile
import pandas as pd
import requests
from pypdf import PdfReader

# --- CONFIGURATION ---
INPUT_DIR = r"C:\Users\Admin\LightRAG\inputs"
SERVER_URL = "http://127.0.0.1:9621/documents/text"

def extract_zip_contents(zip_path, extract_to):
    """Unpacks a ZIP archive and returns all extracted file paths."""
    extracted_files = []
    try:
        with zipfile.ZipFile(zip_path, 'r') as z:
            z.extractall(extract_to)
            for root, _, files in os.walk(extract_to):
                for file in files:
                    extracted_files.append(os.path.join(root, file))
    except Exception as e:
        print(f"   ⚠️ Could not unpack ZIP {os.path.basename(zip_path)}: {e}")
    return extracted_files

def extract_embedded_objects(excel_path, temp_dir):
    """Scans inside an XLSX file for embedded PDFs or embedded ZIP files."""
    extracted_paths = []
    try:
        with zipfile.ZipFile(excel_path, 'r') as z:
            for file_info in z.infolist():
                if "embeddings" in file_info.filename:
                    ext = os.path.splitext(file_info.filename)[1].lower()
                    extracted_name = os.path.basename(file_info.filename)
                    target_path = os.path.join(temp_dir, extracted_name)

                    with z.open(file_info.filename) as source, open(target_path, "wb") as target:
                        target.write(source.read())

                    if ext in [".pdf", ".bin"]:
                        extracted_paths.append(target_path)
                    elif ext == ".zip":
                        # If a ZIP is embedded inside Excel, extract its contents
                        unzipped_folder = os.path.join(temp_dir, f"unzipped_{extracted_name}")
                        extracted_paths.extend(extract_zip_contents(target_path, unzipped_folder))
    except zipfile.BadZipFile:
        pass
    return extracted_paths

def parse_pdf_text(pdf_path):
    """Extracts plain text from a PDF file."""
    try:
        reader = PdfReader(pdf_path)
        text = ""
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
        return text
    except Exception as e:
        return f"[Error reading PDF: {e}]"

def process_single_excel(excel_path):
    """Processes a single Excel file and any attachments inside it."""
    file_name = os.path.basename(excel_path)
    print(f"🔄 Processing Excel: {file_name}")

    # 1. Read Table Data
    df = pd.read_excel(excel_path).dropna(how="all")
    excel_markdown = df.to_markdown(index=False)

    # 2. Check for Embedded PDFs or ZIPs inside Excel
    attachments_text = ""
    with tempfile.TemporaryDirectory() as temp_dir:
        extracted_files = extract_embedded_objects(excel_path, temp_dir)

        for file_path in extracted_files:
            fname = os.path.basename(file_path)
            if file_path.lower().endswith(".pdf"):
                print(f"   📦 Found embedded/nested PDF: {fname}")
                text = parse_pdf_text(file_path)
                attachments_text += f"\n\n--- Attachment ({fname}) ---\n{text}"

    # 3. Combine and Upload
    payload = f"Data from Excel Document '{file_name}':\n\n### Table Data:\n{excel_markdown}"
    if attachments_text:
        payload += f"\n\n### Extracted Attachment Data:\n{attachments_text}"

    res = requests.post(SERVER_URL, json={"text": payload, "description": f"Import of {file_name}"})
    if res.status_code == 200:
        print(f"   ✨ Successfully indexed '{file_name}' into Neo4j!")
    else:
        print(f"   ❌ Server error for '{file_name}': {res.status_code}")

def process_all_inputs():
    print(f"📂 Scanning input directory: {INPUT_DIR}\n" + "="*50)

    with tempfile.TemporaryDirectory() as zip_extract_dir:
        # First, unpack any standalone .zip files dropped in the inputs folder
        standalone_zips = glob.glob(os.path.join(INPUT_DIR, "*.zip"))
        for z_path in standalone_zips:
            print(f"📦 Unpacking input ZIP folder: {os.path.basename(z_path)}")
            extract_zip_contents(z_path, zip_extract_dir)

        # Gather all Excel files (from inputs directly + from unpacked ZIPs)
        direct_excels = glob.glob(os.path.join(INPUT_DIR, "*.xlsx")) + glob.glob(os.path.join(INPUT_DIR, "*.xls"))
        unpacked_excels = glob.glob(os.path.join(zip_extract_dir, "**", "*.xlsx"), recursive=True) + \
                          glob.glob(os.path.join(zip_extract_dir, "**", "*.xls"), recursive=True)

        all_excels = direct_excels + unpacked_excels

        if not all_excels:
            print("📭 No Excel files found to process.")
            return

        for excel_path in all_excels:
            try:
                process_single_excel(excel_path)
            except Exception as e:
                print(f"   ❌ Failed to process '{os.path.basename(excel_path)}': {e}")

if __name__ == "__main__":
    process_all_inputs()