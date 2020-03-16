import win32com.client
import os

# Input the path your folder
folder = "C:\\Users\\thhc1\\PycharmProjects\\Docx\\in"

wdToPDF = win32com.client.DispatchEx("Word.Application")
wdFormatPDF = 17
files = os.listdir(folder)
word_files = [f for f in files if f.endswith((".doc", ".docx"))]
for word_file in word_files:
    word_path = os.path.join(folder, word_file)
    pdf_path = word_path.strip('.docx')
    if pdf_path[-3:] != 'pdf':
        pdf_path = pdf_path + ".pdf"

    if os.path.exists(pdf_path):
        os.remove(pdf_path)

    pdfCreate = wdToPDF.Documents.Open(word_path)
    pdfCreate.SaveAs(pdf_path, wdFormatPDF)
