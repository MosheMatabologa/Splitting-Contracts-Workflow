import os
import time
import tempfile
import psutil
import streamlit as st
from fpdf import FPDF
from PIL import Image
from PyPDF2 import PdfReader, PdfWriter
import win32com.client as win32

# Temporary directories for input and output files in Streamlit
input_folder = tempfile.gettempdir()
output_folder = os.path.join(tempfile.gettempdir(), "Converted_PDFs")
os.makedirs(output_folder, exist_ok=True)

# Function to close lingering processes
def close_lingering_processes(app_name):
    for process in psutil.process_iter(attrs=['pid', 'name']):
        if process.info['name'] and app_name.lower() in process.info['name'].lower():
            try:
                process.terminate()
                process.wait()
                print(f"Closed lingering {app_name} process.")
            except Exception as e:
                print(f"Could not close {app_name} process: {e}")

# Conversion functions
def convert_image_to_pdf(image_path, output_path):
    try:
        with Image.open(image_path) as img:
            pdf = FPDF()
            pdf.add_page()
            img_width, img_height = img.size
            width, height = 210, (img_height / img_width) * 210
            pdf.image(image_path, x=0, y=0, w=width, h=height)
            pdf.output(output_path)
    except Exception as e:
        print(f"Error converting {image_path} to PDF: {e}")

def convert_word_to_pdf(docx_path, output_path):
    close_lingering_processes("WINWORD")
    word = win32.Dispatch('Word.Application')
    word.Visible = False
    try:
        if "~$" in os.path.basename(docx_path):
            return
        doc = word.Documents.Open(docx_path, ReadOnly=True)
        for attempt in range(3):
            try:
                doc.SaveAs(output_path, FileFormat=17)
                break
            except Exception as e:
                print(f"Attempt {attempt+1} failed: {e}")
                time.sleep(2)
        doc.Close(False)
    except Exception as e:
        print(f"Error converting {docx_path}: {e}")
    finally:
        word.Quit()

def convert_excel_to_pdf(excel_path, output_path):
    close_lingering_processes("EXCEL")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        workbook = excel.Workbooks.Open(excel_path, ReadOnly=True)
        for attempt in range(3):
            try:
                workbook.ExportAsFixedFormat(0, output_path)
                break
            except Exception as e:
                print(f"Attempt {attempt+1} failed: {e}")
                time.sleep(2)
        workbook.Close(False)
    except Exception as e:
        print(f"Error converting {excel_path} to PDF: {e}")
    finally:
        excel.Quit()

def convert_pdf_to_pdf(input_path, output_path):
    try:
        with open(input_path, "rb") as infile:
            reader = PdfReader(infile)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            with open(output_path, "wb") as outfile:
                writer.write(outfile)
    except Exception as e:
        print(f"Error copying PDF {input_path}: {e}")

# Processing files based on extension
def process_file(file_path, output_folder):
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

# Streamlit App
st.title("Document Conversion Tool")
st.write("Upload files (.docx, .xlsx, .jpg, .jpeg, .png, .pdf) to convert them to PDF format.")

# Upload files
uploaded_files = st.file_uploader("Choose files", accept_multiple_files=True, type=["docx", "xlsx", "jpg", "jpeg", "png", "pdf"])

# Process and convert files
if uploaded_files:
    for uploaded_file in uploaded_files:
        # Save uploaded file to temporary directory
        file_path = os.path.join(input_folder, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Process and convert to PDF
        process_file(file_path, output_folder)

        # Prepare for download
        converted_pdf_path = os.path.join(output_folder, f"{os.path.splitext(uploaded_file.name)[0]}.pdf")
        if os.path.exists(converted_pdf_path):
            with open(converted_pdf_path, "rb") as pdf_file:
                st.download_button(
                    label=f"Download {os.path.basename(converted_pdf_path)}",
                    data=pdf_file,
                    file_name=os.path.basename(converted_pdf_path),
                    mime="application/pdf"
                )
    st.success("Conversion complete!")
else:
    st.write("Upload a file to start the conversion process.")
