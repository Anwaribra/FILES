# PDF to Word Converter in Python 
This repository contains a Python script that extracts text from PDF files and converts it into Word documents. The script utilizes the PyMuPDF library for PDF text extraction and the python-docx library for creating and saving Word documents.
## Features                                                                                                    
- Extract text from each page of a PDF file.
- Create a new Word document and add the extracted text.
- Save the Word document to a specified location.
### Requirements 
- Python 3.x
- `PyMuPDF` library
- `python-docx` library
 # Installation 
    pip install PyMuPDF python-docx
# Usage
Clone the repository:
```
git clone https://github.com/Anwaribra/PDF-to-Word-Converter-in-Python.git
 cd pdf-to-word-converter
```
## Script code

```python
import fitz
from docx import Document

def pdf_to_word(pdf_path, word_path):
    # Open the PDF file
    pdf_document = fitz.open(pdf_path)
    
    # Create a new Word document
    doc = Document()

    # Extract text from each page of the PDF
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text = page.get_text()

        # Add the extracted text to the Word document
        doc.add_paragraph(text)

    # Save the Word document
    doc.save(word_path)

# Specify the paths to your PDF and Word files
pdf_path = 'A:/FILES/TBTAE VOL04.pdf'
word_path = 'A:/FILES/TBATE.docx'

# Convert PDF to Word
pdf_to_word(pdf_path, word_path)
