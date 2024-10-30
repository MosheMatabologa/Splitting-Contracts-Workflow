import os
import time
from fpdf import FPDF
from PIL import Image
from PyPDF2 import PdfReader, PdfWriter
import win32com.client as win32
import psutil
import streamlit as st




# Folder paths
input_folder = r"C:\Users\Q624157\Desktop\Conversion"
output_folder = r"C:\Users\Q624157\Desktop\Conversion\Converted_PDFs"

# Ensure output folder exists
os.makedirs(output_folder, exist_ok=True)

def close_lingering_processes(app_name):
    """Force close lingering Word or Excel processes."""
    for process in psutil.process_iter(attrs=['pid', 'name']):
        if process.info['name'] and app_name.lower() in process.info['name'].lower():
            try:
                process.terminate()
                process.wait()
                print(f"Closed lingering {app_name} process.")
            except Exception as e:
                print(f"Could not close {app_name} process: {e}")

def convert_image_to_pdf(image_path, output_path):
    """Convert image file to PDF format."""
    try:
        with Image.open(image_path) as img:
            pdf = FPDF()
            pdf.add_page()
            img_width, img_height = img.size
            width, height = 210, (img_height / img_width) * 210
            pdf.image(image_path, x=0, y=0, w=width, h=height)
            pdf.output(output_path)
        print(f"Image {image_path} successfully converted to PDF.")
    except Exception as e:
        print(f"Error converting {image_path} to PDF: {e}")

def convert_word_to_pdf(docx_path, output_path):
    """Convert Word document to PDF format with enhanced error handling."""
    close_lingering_processes("WINWORD")  # Close any lingering Word processes
    word = win32.Dispatch('Word.Application')
    word.Visible = False
    try:
        if "~$" in os.path.basename(docx_path):
            print(f"Skipping temporary file: {docx_path}")
            return
        if not os.access(docx_path, os.R_OK):
            print(f"Cannot access {docx_path}.")
            return

        # Open as read-only and force save if necessary
        doc = word.Documents.Open(docx_path, ReadOnly=True)
        success = False
        for attempt in range(3):
            try:
                doc.SaveAs(output_path, FileFormat=17)  # FileFormat=17 is PDF
                print(f"Word document {docx_path} saved as PDF.")
                success = True
                break
            except Exception as save_error:
                print(f"Attempt {attempt+1} to save {docx_path} failed: {save_error}")
                time.sleep(2)
        
        if not success:
            print(f"Failed to save {docx_path} after multiple attempts.")
        doc.Close(False)
    except Exception as e:
        print(f"Error converting {docx_path}: {e}")
    finally:
        word.Quit()

def convert_excel_to_pdf(excel_path, output_path):
    """Convert Excel workbook to PDF format."""
    close_lingering_processes("EXCEL")  # Close any lingering Excel processes
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        workbook = excel.Workbooks.Open(excel_path, ReadOnly=True)
        for attempt in range(3):
            try:
                workbook.ExportAsFixedFormat(0, output_path)
                print(f"Excel workbook {excel_path} saved as PDF.")
                break
            except Exception as save_error:
                print(f"Attempt {attempt+1} to save {excel_path} failed: {save_error}")
                time.sleep(2)
        workbook.Close(False)
    except Exception as e:
        print(f"Error converting {excel_path} to PDF: {e}")
    finally:
        excel.Quit()

def convert_pdf_to_pdf(input_path, output_path):
    """Copy or process existing PDFs for consistency."""
    try:
        with open(input_path, "rb") as infile:
            reader = PdfReader(infile)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            with open(output_path, "wb") as outfile:
                writer.write(outfile)
        print(f"PDF {input_path} copied to {output_path}.")
    except Exception as e:
        print(f"Error copying PDF {input_path}: {e}")

def process_file(file_path, output_folder):
    """Determine file type and convert to PDF if supported."""
    file_name, ext = os.path.splitext(os.path.basename(file_path))
    output_path = os.path.join(output_folder, f"{file_name}.pdf")
    
    if ext.lower() in [".jpg", ".jpeg", ".png"]:
        convert_image_to_pdf(file_path, output_path)
    elif ext.lower() == ".docx":
        convert_word_to_pdf(file_path, output_path)
    elif ext.lower() == ".xlsx":
        convert_excel_to_pdf(file_path, output_path)
    elif ext.lower() == ".pdf":
        convert_pdf_to_pdf(file_path, output_path)
    else:
        print(f"Unsupported file format for {file_path}")

def traverse_folders(input_folder, output_folder):
    """Process all files in the input folder."""
    for root, _, files in os.walk(input_folder):
        for file in files:
            file_path = os.path.join(root, file)
            print(f"Processing {file_path}")
            process_file(file_path, output_folder)

# Start the conversion process
traverse_folders(input_folder, output_folder)
print("Conversion complete.")
