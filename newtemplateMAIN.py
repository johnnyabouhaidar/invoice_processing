from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import pandas as pd
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


pytesseract.pytesseract.tesseract_cmd = r'C:\Users\johnny.abouhaidar\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

# Convert PDF to image



def process_invoice(filepath,outputpath):
    images = convert_from_path(filepath,poppler_path=r"D:\KSA\saudi coffee co\Release-24.08.0-0\poppler-24.08.0\Library\bin")
    
    columnnames = ['Inv No', 'Date']
    header_df = pd.DataFrame(columns=columnnames)

    header_fields= [[(569, 484, 1080, 524),'Inv No'],[(570, 523, 1080, 566),'Date']]
    for field in header_fields:
        cropped = images[0].crop(field[0])
        
        if field[1]!="Date" and "PO" and "Inv" not in field[1]:
            text = pytesseract.image_to_string(cropped,lang="ara")    
        else:
            text = pytesseract.image_to_string(cropped)    
        
        header_df[field[1]]=text.split("|||")
    print(header_df)
    header_df.to_excel(outputpath,"header",header=True,index=False)
    ##############################table

    columnnames = ['SKU', 'QTY', 'unit price','total price']
    df = pd.DataFrame(columns=columnnames)


    columns= [[(1253, 770, 1522, 1463),'SKU'],[(511, 770, 671, 1463),'QTY'],[(286, 770, 501, 1463),'unit price'],[(65, 770, 283, 1463),'total price']]
    for image in images:
        tmpdf = pd.DataFrame(columns=columnnames)
        for col in columns:
            cropped = image.crop(col[0])
            
            text = pytesseract.image_to_string(cropped,config=r"--psm 6 -c tessedit_char_whitelist=0123456789.")
            #print(text.replace("\n","|||"))

            
            tmpdf[col[1]]=(text.replace("\n","|||").split('|||'))
            
            
        df = pd.concat([df, tmpdf], ignore_index=True)
    df.replace("", pd.NA, inplace=True)
    df = df.dropna()
    print(df["total price"].astype(float).sum()*0.15)
    #df.to_excel("output.xlsx","table",header=False,index=False)
    with pd.ExcelWriter(outputpath, mode='a', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Table', index=False)





##############################header

if __name__ == '__main__':
    in_path = sys.argv[1]
    out_path = sys.argv[2]
    process_invoice(in_path,out_path)