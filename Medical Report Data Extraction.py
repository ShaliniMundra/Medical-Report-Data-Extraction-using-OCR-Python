#!/usr/bin/env python
# coding: utf-8

# In[1]:


pip install img2table pytesseract


# In[2]:


pip install pdf2image


# In[16]:


from pdf2image import convert_from_path
from img2table.ocr import TesseractOCR
from img2table.document import Image
import pytesseract

# Set Tesseract path
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Set Poppler bin folder path
poppler_path = r"C:\Users\HP\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin"

# Convert PDF to images
images = convert_from_path(r"C:\Users\HP\Downloads\Hematology Report Example, Format, Sample and Template - Drlogy Lab Reports.pdf", poppler_path = poppler_path)
images


# In[55]:


# Process each page
import os
ocr = TesseractOCR(n_threads=1, lang="eng")

for i, img in enumerate(images):
    img_path = f"page_{i}.png"
    img.save(img_path, "PNG")
    print(f"✅ Saved image: {img_path}")
    
# Process image with img2table
doc = Image(img_path)
output_file = f"output_page_{i}.xlsx"
print(output_file)
#     doc.to_xlsx(dest=output_file,
#                 ocr=ocr,
#                 implicit_rows=False,
#                 implicit_columns=False,
#                 borderless_tables=True,
#                 min_confidence=50)
#     print(f"✅ Excel file should be generated: {output_file}")
# Extract tables from the page
tables = doc.extract_tables(
    ocr=ocr,
    implicit_rows=False,
    implicit_columns=False,
    borderless_tables=True,
    min_confidence=50
)

