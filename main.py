# import functions
from pathlib import Path
import PySimpleGUI as sg
from docx2pdf import convert
import shutil
import os
import pandas as pd  # pip install pandas openpyxl
from docxtpl import DocxTemplate  # pip install docxtpl
base_dir = Path(__file__).parent
excel_path1 = base_dir / "Project RP07 Letter.xlsx"
excel_path2 = base_dir / "Project RP07 Letter2 .xlsx"
output_dir = base_dir / "OUTPUT1"
output_dir1 = base_dir / "OUTPUT2"
Temp1 = base_dir / "Templates1"
Temp2 = base_dir / "Templates2"
sg.theme('Dark Grey 13')

def convert_pdf():
    path = output_dir
    os.chdir(path)
    # files = glob.glob(path+"/*.docx")
    for f in os.listdir():
        if f.endswith(".docx"):
            convert(f)
            os.remove(f)
            #convert and remove docx documents
def convert_pdf2():
    path = output_dir1
    os.chdir(path)
    # files = glob.glob(path+"/*.docx")
    for f in os.listdir():
        if f.endswith(".docx"):    
            convert(f)
            os.remove(f)
            #convert and remove docx documents


def scenerio_1():
    word_template_path1 = base_dir / "Templates1/Payment Receipt.docx"
    word_template_path2 = base_dir / "Templates1/RP07 Letter Template.docx"
    word_template_path3 = base_dir / "Templates1/Booking Confirmation.docx"
    word_template_path4 = base_dir / "Templates1/RP07_V1.2a.docx"
    # output_dir = base_dir / "OUTPUT1"
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
        
        x = len(record['RP07_ADDRESS_POSTCODE'])
        if x != 8:
            record['RP07_ADDRESS_POSTCODE'] = record['RP07_ADDRESS_POSTCODE'] + " "*(8-x)

        record['p1'] = (record['RP07_ADDRESS_POSTCODE'])[0]
        record['p2'] = (record['RP07_ADDRESS_POSTCODE'])[1]
        record['p3'] = (record['RP07_ADDRESS_POSTCODE'])[2]
        record['p4'] = (record['RP07_ADDRESS_POSTCODE'])[3]
        record['p5'] = (record['RP07_ADDRESS_POSTCODE'])[4]
        record['p6'] = (record['RP07_ADDRESS_POSTCODE'])[5]
        try:
            record['p7'] = (record['RP07_ADDRESS_POSTCODE'])[6]
        except:
            print("List out of values")
        try:
            record['p8'] = (record['RP07_ADDRESS_POSTCODE'])[7]
        except:
            print("List out of values")

        record['c1'] = (str(record['Company_Number']))[0]
        record['c2'] = (str(record['Company_Number']))[1]
        record['c3'] = (str(record['Company_Number']))[2]
        record['c4'] = (str(record['Company_Number']))[3]
        record['c5'] = (str(record['Company_Number']))[4]
        record['c6'] = (str(record['Company_Number']))[5]
        record['c7'] = (str(record['Company_Number']))[6]
        record['c8'] = (str(record['Company_Number']))[7]

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
        doc4 = DocxTemplate(word_template_path4)
        doc4.render(record)
        # print("Waiting")
        # time.sleep(20)
        output_path3 = output_dir / f"{record['Company_Name']}-RP07-2.docx"
        doc4.save(output_path3)
        # os.chdir(r"E:\python\OUTPUT")
        os.chdir(f"{output_dir}")
        os.system(f"mkdir {record['Company_Number']} ")
        convert_pdf()
        os.system(f"move {(record['Company_Name'].split())[0]}*.pdf {output_dir}/{record['Company_Number']} || mv {(record['Company_Name'].split())[0]}*.pdf {output_dir}/{record['Company_Number']} ")
        

def scenerio_2():
    word_template_path1 = base_dir / "Templates2/RP07 Letter.docx"
    word_template_path2 = base_dir / "Templates2/RP07_V1.2a.docx"
    # output_dir = base_dir / "OUTPUT2"
    # Create output folder for the word documents
    output_dir1.mkdir(exist_ok=True)
    df = pd.read_excel(excel_path2, sheet_name="Sheet1")

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
        if record['Contact_No'] == "":
            record['Contact_No']= "  "
        
        record['p1'] = (record['RP07_ADDRESS_POSTCODE'])[0]
        record['p2'] = (record['RP07_ADDRESS_POSTCODE'])[1]
        record['p3'] = (record['RP07_ADDRESS_POSTCODE'])[2]
        record['p4'] = (record['RP07_ADDRESS_POSTCODE'])[3]
        record['p5'] = (record['RP07_ADDRESS_POSTCODE'])[4]
        record['p6'] = (record['RP07_ADDRESS_POSTCODE'])[5]
        try:
            record['p7'] = (record['RP07_ADDRESS_POSTCODE'])[6]
        except:
            print("List out of values")
        try:
            record['p8'] = (record['RP07_ADDRESS_POSTCODE'])[7]
        except:
            print("List out of values")

        record['c1'] = (str(record['Company_Number']))[0]
        record['c2'] = (str(record['Company_Number']))[1]
        record['c3'] = (str(record['Company_Number']))[2]
        record['c4'] = (str(record['Company_Number']))[3]
        record['c5'] = (str(record['Company_Number']))[4]
        record['c6'] = (str(record['Company_Number']))[5]
        record['c7'] = (str(record['Company_Number']))[6]
        record['c8'] = (str(record['Company_Number']))[7]
        

        doc = DocxTemplate(word_template_path1)
        # get_payment()
        doc.render(record)
        output_path = output_dir1 / f"{record['Company_Name']}-Letter.docx"
        doc.save(output_path)
        doc2 = DocxTemplate(word_template_path2)
        doc2.render(record)
        output_path1 = output_dir1 / f"{record['Company_Name']}-RP07 v1.docx"
        doc2.save(output_path1)

        # os.chdir(r"E:\python\OUTPUT")
        os.chdir(f"{output_dir1}")
        os.system(f"mkdir {record['Company_Number']} ")
        convert_pdf2()
        os.system(f"move {(record['Company_Name'].split())[0]}*.pdf {output_dir1}/{record['Company_Number']} || mv {(record['Company_Name'].split())[0]}*.pdf {output_dir1}/{record['Company_Number']}")
        if "Chadwell" in record['Location']:
            shutil.copy(f"{Temp2}\\Chadwell Heath Address Proof.pdf",f"{record['Company_Number']}")
        if "East" in record['Location']:
            shutil.copy(f"{Temp2}\\East Ham Address Proof.pdf",f"{record['Company_Number']}")
        if "Hainault" in record['Location']:
            shutil.copy(f"{Temp2}\\Hainault Address Proof.pdf",f"{record['Company_Number']}")
        if "Hatton" in record['Location']:
            shutil.copy(f"{Temp2}\\Hatton Garden Lease.pdf",f"{record['Company_Number']}")

layout = [
    [sg.Text("Welcome User", justification="centre")],
    [sg.Button("Scenario 1")],
    [sg.Button("Scenario 2")],
    [sg.Exit()]
]

window = sg.Window("Excel 2 Pdf Maker 1.0", layout , size=(400, 200) , grab_anywhere=True, element_justification='c' )

while True:
    event , values = window.read()
    if event in (sg.WINDOW_CLOSED,"Exit"):
        break
    if event == "Scenario 1":
        scenerio_1()
        sg.popup_no_titlebar(f"Done Scenerio 1")
    if event == "Scenario 2":
        scenerio_2()
        sg.popup_no_titlebar(f"Done Scenerio 2")

window.close()