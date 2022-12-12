# import functions
from pathlib import Path
from docx2pdf import convert
import os
import pandas as pd  # pip install pandas openpyxl
from docxtpl import DocxTemplate  # pip install docxtpl
base_dir = Path(__file__).parent
excel_path1 = base_dir / "Project RP07 Letter.xlsx"

def convert_pdf():
    path = output_dir
    os.chdir(path)
    # files = glob.glob(path+"/*.docx")
    for f in os.listdir():
        if f.endswith(".docx"):
            convert(f)
            os.remove(f)
            #convert and remove docx documents

path = output_dir
def scenerio_1():
    word_template_path1 = base_dir / "Templates1/Payment Receipt.docx"
    word_template_path2 = base_dir / "Templates1/RP07 Letter Template.docx"
    word_template_path3 = base_dir / "Templates1/Booking Confirmation.docx"
    word_template_path3 = base_dir / "Templates1/Booking Confirmation.docx"
    output_dir = base_dir / "OUTPUT1"
    # Create output folder for the word documents
    output_dir.mkdir(exist_ok=True)
    df = pd.read_excel(excel_path1, sheet_name="Sheet1")

    # Iterate over each row in df and render word document
    for record in df.to_dict(orient="records"):
        if record['Package'] == "Silver" and record['Payment']=="Quarterly":
            record['Price'] = "38.87"
            record['Vat'] = 6.48
        elif record['Package'] == "Bronze" and record['Payment']=="Quarterly":
            record['Price'] = "12.87"
            record['Vat'] = 2.15
        elif record['Package'] == "Gold" and record['Payment']=="Quarterly":
            record['Price'] = "64.87"
            record['Vat'] = 10.18
        elif record['Package'] == "Platinum" and record['Payment']=="Quarterly":
            record['Price'] = "90.87"
            record['Vat'] = 15.15

        if record['Package'] == "Silver" and record['Payment']=="Annual":
            record['Price'] = "120.12"
            record['Vat'] = 20.02
        elif record['Package'] == "Bronze" and record['Payment']=="Annual":
            record['Price'] = "45.76"
            record['Vat'] = 7.63
        elif record['Package'] == "Gold" and record['Payment']=="Annual":
            record['Price'] = "200.20"
            record['Vat'] = 33.37
        elif record['Package'] == "Platinum" and record['Payment']=="Annual":
            record['Price'] = "273.00"
            record['Vat'] = 45.50

        record['Net'] = round(float(record['Price']) - float(record['Vat']), 2)
        if record['Email'] == "nan":
            record['Email']= "  "
        if record['Contact_No'] == "nan":
            record['Contact_No']= "  "

        doc = DocxTemplate(word_template_path1)
        # get_payment()
        doc.render(record)
        output_path = output_dir / f"{record['Company_Name']}-Payment.docx"
        doc.save(output_path)
        doc2 = DocxTemplate(word_template_path2)
        doc2.render(record)
        output_path1 = output_dir / f"{record['Company_Name']}-Letter.docx"
        doc2.save(output_path1)
        doc3 = DocxTemplate(word_template_path3)
        doc3.render(record)
        output_path2 = output_dir / f"{record['Company_Name']}-Booking.docx"
        doc3.save(output_path2)
        # doc4 = DocxTemplate(word_template_path4)
        # doc4.render(record)
        # output_path3 = output_dir / f"{record['Company_Name']}-Booking.docx"
        # doc3.save(output_path3)
        # os.chdir(r"E:\python\OUTPUT")
        os.chdir(f"{output_dir}")
        os.system(f"mkdir {(record['Company_Name'].split())[0]} ")
        convert_pdf()
        os.system(f"move {(record['Company_Name'].split())[0]}*.pdf {output_dir}/{(record['Company_Name'].split())[0]}")






