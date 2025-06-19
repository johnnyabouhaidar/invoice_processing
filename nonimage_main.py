import fitz  # PyMuPDF

doc = fitz.open("0005-1143.pdf")
for page in doc:
    text = page.get_text()
    print(text)