import os
from PyPDF2 import PdfReader, PdfWriter

# List of names (your original names)
names = [
    "676593_Nawana Hlologelo Hope_Setati_TX-R-16",
    "676594_Sipho Johannes_Makhubela_TX-R-16",
    "676595_Bridged Lebogang_Dhlamini_TX-R-43",
    "676596_Nythel Nkaforoana_Kekana_TX-R-16",
    "676597_Thabo Philemon_Moalosi_TX-R-16",
    "676598_Precious Ramaite_Lekgau_TX-R-43",
    "676599_Kgomotso Precious_Mlangeni_TX-R-43",
    "676600_Ofentse Martinah_Seriti_TX-R-16",
    "676601_Selaelo Kanti_Rasehona_TX-R-43",
    "676602_Kenny Kagiso_Mahlangu_TX-R-16",
    "676603_Thokozile Zuzile_Twala_TX-R-16",
    "676604_Khomotso Brian_Moshobane_TX-R-43",
    "676605_Thabelo Pinky_Manyoka_TX-R-16",
    "676606_Sizzi_Kopotja_TX-R-43",
    "676607_Rose Duduzile_Madhlope_TX-R-43",
    "676608_Timothy Isaac_Baloyi_TX-R-43",
    "676609_Kuhle_Ngceba_TX-R-16",
    "676610_Kesaobaka_Ledikwa_TX-R-43",
    "676611_Mthobisi_Shezi_TX-R-16",
    "676612_Herodian Cavern_Guss_TX-R-16",
    "676613_Karabelo Surprise_Matsepe_TX-R-43",
    "676614_Keamogetswe Precious_Maake_TX-R-43",
    "676615_Mmasetshaba Carol_Molobela_TX-R-16",
    "676616_Gerald Jackie_Matshele_TX-R-43",
    "676617_Kenny Kagiso_Mahlangu_TX-R-16"
]








# Number of contracts to split (for this test, we'll split only 5)
total_contracts = 25

# PDF file to split
input_pdf = r"C:\Users\Q624157\Downloads\20241018_PAT_Contract_Mail_Merge_Batch 26 Letters Final.pdf"

# Create a PdfReader object to read the input PDF
pdf_reader = PdfReader(input_pdf)

# Total number of pages for the first 5 contracts (5 contracts * 8 pages = 40 pages)
total_pages = total_contracts * 8

# Specify the output directory
output_directory = r"C:\Users\Q624157\Desktop\Batch_26_Contracts_split"
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
print(f'There are {len(names)} contracts completed.')
