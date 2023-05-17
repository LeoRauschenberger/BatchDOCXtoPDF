import os
from docx2pdf import convert

# Folder containing the DOCX files
folder_path = "C:/Users/.../location"

# Output folder for PDF files
output_folder = "CC:/Users/.../location/PDF"

# Iterate over files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith(".docx"):
        docx_path = os.path.join(folder_path, filename)
        pdf_filename = filename[:-5] + ".pdf"
        pdf_path = os.path.join(output_folder, pdf_filename)
        convert(docx_path, pdf_path)
        print(f"Converted {filename} to PDF.")

print("Conversion complete.")
