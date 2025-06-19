import pytesseract
from pytesseract import Output
from pdf2image import convert_from_path
import pandas as pd
import cv2
import numpy as np
from PIL import Image

# Convert PDF pages to images
images = convert_from_path("0005-1143.pdf", dpi=300,poppler_path=r"D:\KSA\saudi coffee co\Release-24.08.0-0\poppler-24.08.0\Library\bin")

def extract_header(data):
    # Get header by taking first few top-positioned words
    top_lines = [data['text'][i] for i in range(len(data['text']))
                 if int(data['top'][i]) < 200 and data['text'][i].strip() != '']
    return " ".join(top_lines)

def extract_table_data(image):
    # Convert to OpenCV image
    open_cv_image = np.array(image.convert('RGB'))
    gray = cv2.cvtColor(open_cv_image, cv2.COLOR_RGB2GRAY)
    
    # Binarize image
    _, binary = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)

    # Detect horizontal and vertical lines for table structure
    kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
    kernel_v = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 30))

    detect_horizontal = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel_h)
    detect_vertical = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel_v)

    table_mask = cv2.add(detect_horizontal, detect_vertical)

    # Find contours (table cells)
    contours, _ = cv2.findContours(table_mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

    cell_data = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > 50 and h > 15:  # Filter out small artifacts
            roi = open_cv_image[y:y+h, x:x+w]
            text = pytesseract.image_to_string(roi, config='--psm 6').strip()
            cell_data.append((x, y, w, h, text))

    # Sort by y, then x (row-wise)
    cell_data.sort(key=lambda b: (b[1], b[0]))

    # Group into rows
    rows = []
    current_row = []
    last_y = -1
    for cell in cell_data:
        x, y, w, h, text = cell
        if last_y == -1 or abs(y - last_y) <= 15:
            current_row.append(text)
        else:
            rows.append(current_row)
            current_row = [text]
        last_y = y
    if current_row:
        rows.append(current_row)

    return rows

# Process each page
for i, image in enumerate(images):
    print(f"\n--- Page {i+1} ---")
    
    # OCR full data
    
    pytesseract.pytesseract.tesseract_cmd = r'C:\Users\johnny.abouhaidar\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
    data = pytesseract.image_to_string(image)
    print(data)
    with open('output.txt', 'w') as file:
        file.write(data)
  
    '''
    # Extract header
    header = extract_header(data)
    print(f"\nHeader: {header}")

    # Extract table
    rows = extract_table_data(image)
    df = pd.DataFrame(rows)
    print("\nExtracted Table:")
    print(df)

    # Optional: Save to Excel
    df.to_excel(f"extracted_page_{i+1}.xlsx", index=False, header=False)
    '''