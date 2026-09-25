from pdf2docx import Converter

# Use the exact filename of your uploaded PDF
pdf_file = "PHASE 3 - Site Development Plan and Investment Program.pdf"
docx_file = "PHASE 3 - Site Development Plan and Investment Program.docx"

# Convert PDF to Word
cv = Converter(pdf_file)
cv.convert(docx_file, start=0, end=None)
cv.close()
print("Conversion complete! Your Word document has been generated.")
