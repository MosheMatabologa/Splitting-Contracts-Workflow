import os
import pandas 
from PyPDF2 import PdfReader, PdfWriter

# List of names (your original names)
names = [
    "676872_Francina_Nkosi_TX-R-45",
    "676874_Ntobeng Meshack_Maila_TX-R-43",
    "676875_Sphiwe_Mentoor_TX-R-43",
    "676877_Khensani Obed_Baloyi_TX-R-43",
    "676878_Tshiamo_Kekana_TX-R-43",
    "676879_Reneilwe Andria_Serakalala_TX-R-43",
    "676880_Patricia Dikgetho_Makokoana_TX-R-43",
    "676881_Mbali Michiel_Mtungwa_TX-R-45",
    "676883_Moeketsi_Seretlo_TX-R-45",
    "676884_Thato Hendrick_Phaho_TX-R-43",
    "676886_Lerato_Kgasi_TX-R-45",
    "676888_Tebogo Magdeline_Mataboge_TX-R-43",
    "676890_Thulisiwe_Shoko_TX-R-43",
    "676891_Mosa Priscilla_Monaheng_TX-R-45",
    "676909_Keamogetswe Calvin_Aphane_TX-R-45",
    "676912_Itumeleng Petrus_Tsetsewa_TX-R-45",
    "676913_Gontse Solomon Victor_Ntladi_TX-R-45",
    "676914_Jane_Mokgola_TX-R-45",
    "676915_Amogelang Mary Alice_Masilo_TX-R-45",
    "676916_Anna Maletjane Adelaide_Mondlane_TX-R-45",
    "676918_Eva Melina_Manhica_TX-R-45",
    "676919_Simiso Sinothile Noxolo_Manqele_TX-R-43",
    "676921_Kevin_Masekela_TX-R-45",
    "676922_Kealeboga Pretty_Komane_TX-R-45",
    "676923_Dudu Maria_Masimula_TX-R-45",
    "676924_Denzel Masesi_Nqina_TX-R-45",
    "676925_David Kamogelo_Motaung_TX-R-45",
    "676926_Precious_Lesufi_TX-R-43",
    "676927_Boikhutso Josiah_Mangoagape_TX-R-45",
    "676928_Andani_Musetsho_TX-R-45"
]









# Number of contracts to split (for this test, we'll split only 5)
total_contracts = 30

# PDF file to split
input_pdf = r"C:\Users\Q624157\Desktop\20241018_PAT_Contract_Mail_Merge_Batch 30 Letters Final.pdf"

# Create a PdfReader object to read the input PDF
pdf_reader = PdfReader(input_pdf)

# Total number of pages for the first 5 contracts (5 contracts * 8 pages = 40 pages)
total_pages = total_contracts * 8

# Specify the output directory
output_directory = r"C:\Users\Q624157\Desktop\Batch_30_Split_Contracts"
os.makedirs(output_directory, exist_ok=True)

# Split the PDF into multiple PDFs (just the first 5 contracts)
for i in range(total_contracts):
    # Create a new PdfWriter for each contract
    pdf_writer = PdfWriter()

    # Add pages to the PdfWriter (each contract has 8 pages)
    start_page = i * 8
    end_page = start_page + 8
    for page_num in range(start_page, end_page):
        pdf_writer.add_page(pdf_reader.pages[page_num])

    # Output PDF file name using the names list
    output_pdf = os.path.join(output_directory, f"{names[i]}.pdf")

    # Write the PdfWriter pages to the output PDF file
    with open(output_pdf, "wb") as out:
        pdf_writer.write(out)

    print(f"Created PDF: {output_pdf}")
    
#print(len(names))
print('Contracts have all succesfully been split, Moshe')
