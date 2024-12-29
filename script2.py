from docx import Document, document
import openpyxl
import os

# Paths
template_path = "./ChangeMe.docx"
output_dir = "factures"
os.makedirs(output_dir, exist_ok=True)
excel_file = "word_metadata.xlsx"

# Create or load Excel workbook
if os.path.exists(excel_file):
    wb = openpyxl.load_workbook(excel_file)
    ws = wb.active
else:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["File Name", "Content Summary"])

# Sample data to populate the table
from docx import Document

# Data to populate in the table
data = [
    [
        {"khedma": "faza", "qte": "2", "Prix": "400"},
        {"khedma": "faza2", "qte": "1", "Prix": "400"},
        {"khedma": "faza3", "qte": "1", "Prix": "400"}
    ],
    [
        {"khedma": "efe", "qte": "2", "Prix": "400"},
        {"khedma": "fazfegega2", "qte": "1", "Prix": "400"},
        {"khedma": "eeee", "qte": "1", "Prix": "400"}
    ]
]

template_path = "ChangeMe.docx"
output_dir = "factures"
os.makedirs(output_dir, exist_ok=True)
doc = Document(template_path)
table = doc.tables[1] 


j=1
for factures in data:
    i=1
    for facture in factures:
        values = list(facture.values())
        for col_index, value in enumerate(values):
            print(value,i,col_index)
            table.cell(i, col_index+1).text = value
        i=i+1
    file_name = f"Facture{j}.docx"
    j+=1
    file_path = os.path.join(output_dir, file_name)
    doc.save(file_path)




print("Documents created successfully.")

