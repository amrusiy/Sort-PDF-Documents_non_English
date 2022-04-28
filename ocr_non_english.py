# coding: utf-8
import os, glob,re, sys
from os import listdir
from glob import glob
import pandas as pd
from pdf2image import convert_from_path
from PyPDF2 import PdfFileWriter, PdfFileReader
import shutil
import cv2
import matplotlib.pyplot as plt
import datetime
import pathlib
import pytesseract
from PIL import Image
from PIL import Image, ImageEnhance



POPPLER_PATH = r'C:\Users\amrus\Desktop\פרויקטים\כל בו לבניין אמרוסי\scan doc\ocr_non_english\poppler-0.68.0\bin'
PATH_TESSERACT = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
PATH_FILE = r'C:\Users\amrus\Desktop\פרויקטים\כל בו לבניין אמרוסי\scan doc\ocr_non_english'
PATH_DEFAULT = r'C:\Users\amrus\Desktop\פרויקטים\כל בו לבניין אמרוסי\scan doc\ocr_non_english\file\default'
PATH_CUSTOMER = r'C:\Users\amrus\Desktop\פרויקטים\כל בו לבניין אמרוסי\scan doc\ocr_non_english\file\customers'
PATH_SUPPLY_INV = r'C:\Users\amrus\Desktop\פרויקטים\כל בו לבניין אמרוסי\scan doc\ocr_non_english\file\supply\incvoice'
PATH_SUPPLY_SHPPING = r'C:\Users\amrus\Desktop\פרויקטים\כל בו לבניין אמרוסי\scan doc\ocr_non_english\file\supply\shipping certificate'
EXCEL_PATH = r'C:\Users\amrus\Desktop\פרויקטים\כל בו לבניין אמרוסי\scan doc\ocr_non_english\management customer and supply.xlsx'
IMAGE_PATH = r'C:\Users\amrus\Desktop\פרויקטים\כל בו לבניין אמרוסי\scan doc\ocr_non_english\file\default\page_1.png'
SUPPLY = "supply"
CUSTORMER = "customer"
SHIPPING_CERTIFICATE = "shipping certificate"
RECEIPT = "receipt"
CP_AMRUSI = '511527046'

#extract data from excel and return privates_companies number and name, customer_code number and name
def extract_from_excel():
    excel_data = data = pd.read_excel (EXCEL_PATH)
    privates_companies = data["ח.פ"].tolist()
    privates_companies_name = data["שם הספק"].tolist()
    customer_code = data["קוד הלקוח"].tolist()
    customer_code_name = data["שם הלקוח"].tolist()

    return privates_companies, privates_companies_name, customer_code_name, customer_code
#get list of names with all the pdf files in default
def get_names_of_all_the_pdf_file():
    os.chdir(PATH_DEFAULT)
    file_list = glob("*.pdf")
    return file_list
#return the path to file name "filename" in directory default
def fix_path_name(filename,path):
    filename = path+ "\\" +filename
    return filename
#convert the pdf to picture
def Converting_PDF_to_jpeg(files_names):
    "Part #1 : Converting PDF to images"
    os.chdir(PATH_DEFAULT)
    pages = convert_from_path(files_names, 500, poppler_path = POPPLER_PATH)
    for page in pages:
        # Counter to store images of each page of PDF to image
        print(f"converting pdf to png")
        # PDF page n -> page_n.png
        filename = "page_"+'1'+".png"
        # Save the image of the page in system
        page.save(filename, 'png')
#extract the text from the picture and remove the jpeg picture
def improve_contrast():
    im = Image.open("page_1.png")
    enhancer = ImageEnhance.Contrast(im)
    factor = 2  # increase contrast
    im_output = enhancer.enhance(factor)
    im_output.save('page_1.png')

    # enhancer = ImageEnhance.Sharpness(im)
    # factor = 2
    # im_s_1 = enhancer.enhance(factor)
    # im_s_1.save('page_1.png');
def Recognizing_text_from_the_images_using_OCR(type):
    os.chdir(PATH_DEFAULT)
    text = ""
    pytesseract.pytesseract.tesseract_cmd = PATH_TESSERACT
    if type == 'heb':
        text = str((pytesseract.image_to_string(Image.open('page_1.png'),lang = 'heb')))
    if type == 'heb_eng':
        text = str(((pytesseract.image_to_string(Image.open('page_1.png'), lang = 'heb+eng'))))


    # print("deleting temporary png")
    # filename = "page_1.png"
    # if filename.endswith('.png'):
    #     os.remove(filename)

    return str(text)

"find supply information"
def second_call_cp(text):
    cp = ""
    cp = find_cp(text, "מורשה")
    if (cp == "") or (cp == CP_AMRUSI):
        cp = find_cp(text, "ח.פ")
        if (cp == "") or (cp == CP_AMRUSI):
            cp = find_cp(text, "ת.פ")
            if (cp == "") or (cp == CP_AMRUSI):
                cp = find_cp(text, "מספר ע.מ")
                if (cp == "") or (cp == CP_AMRUSI):
                    cp = find_cp(text, "ע.מ")
                    if (cp == "") or (cp == CP_AMRUSI):
                        cp = find_cp(text, "מס'")
                        if (cp == "") or (cp == CP_AMRUSI):
                            cp = find_cp(text, "עמ")
                            if (cp == "") or (cp == CP_AMRUSI):
                                cp = find_cp(text, "חפ")

    return cp
def find_cp(text,word):
    cprivate = ""
    cp = ""
    find_first = text.find(word)
    if find_first != -1:
        cprivate = text[find_first - 13:find_first + 30]
        cprivate = cprivate.replace("-","")
        cp = re.findall(r'(\d+\d+\d+\d+\d+\d+\d+\d+)', cprivate)
        if len(cp) > 0:
            cp = cp[0]
            # for cp with 8 digit
            if cp[0] == '0':
                print(cp[1:])
                cp = cp[1:]
        else:
            cp = ""
    if (cp == CP_AMRUSI) or (cp == ""):
        find_second = text.find(word,text.find(word) + 1)
        if find_first != -1:
            cprivate = text[find_first:find_first + 30]
            cprivate = cprivate.replace("-","")
            cp = re.findall(r'(\d+\d+\d+\d+\d+\d+\d+\d+)', cprivate)
            if len(cp) > 0:
                cp = cp[0]
                # for cp with 8 digit
                if cp[0] == '0':
                    print(cp[1:])
                    cp = cp[1:]
            else:
                cp = ""
    if (cp == CP_AMRUSI) or (cp == ""):
        cp_list = re.findall(r'(\d+\d+\d+\d+\d+\d+\d+\d+\d+)', text)
        for i in cp_list:
            if (i != CP_AMRUSI) and (len(i) == 9) and (i[0] == '5'):
                cp = i
                break

    if (len(cp) == 9) and (cp[0] != '5'):
        temp = list(cp)
        temp[0] = '5'
        cp = "".join(temp)

    return cp
def doc_name_searce_word(text,word,type):
    dn = ""
    find_first = text.find(word)
    if find_first != -1:
        if type == 'heb':
            doc_name = text[find_first:find_first + 34]
        if type == 'heb_eng':
            doc_name = text[find_first:find_first + 34]
        dn = re.findall(r'(\d+\d+\d+\d+\d+)', doc_name)
        if len(dn) > 0:
            dn = dn[0]
        else:
            dn = ""
    return dn
def find_type_and_doc_name(text,type):
    doc_name = ""
    dn = ""
    type_file = ""

    dn = doc_name_searce_word(text, "חשבוני", type)
    if dn != "":
        type_file = RECEIPT
    else:
        dn = doc_name_searce_word(text, "משלו", type)
        if dn != "":
            type_file = SHIPPING_CERTIFICATE

    if dn == "":
        dn = doc_name_searce_word(text, "תעודת", type)
        if dn != "":
            type_file = SHIPPING_CERTIFICATE

    if dn == "":
        dn = doc_name_searce_word(text, "SH", type)
    if dn == "":
        dn = doc_name_searce_word(text,"מספר",type)
    if dn == "":
        dn = doc_name_searce_word(text,"מס",type)

    return type_file, dn
def find_date_from_text(text):
    date = ""
    find_first = text.find("תאריך")
    if find_first != -1:
        date_text = text[find_first :find_first + 30]
        date = re.findall(r'(\d+/\d+/\d+)', date_text)
        if len(date)>0:
            date = date[0]
        else:
            date = ""
    else:
        search_date = re.findall(r'(\d+/\d+/\d+)', text)
        if len(search_date) > 0:
            date = search_date[0]
        else:
            date = ""

    return date
# find from the text - compeny privet number, if its incvoice or shipping certificate
# and the number of the order -> retrun cp, kind,num_order
def find_in_text_information_company_private(text):
     where = ""
     dn = ""
     date = ""
     type = ""

     # find cp in the text
     cp = second_call_cp(text)

     if cp != "":
         # if its RECEIPT or shipping certificate return type
         # find the name of the document
         type, dn = find_type_and_doc_name(text,"heb")
         if type == "":
             type = SHIPPING_CERTIFICATE #until the recipt start to work

         # find date in the the text, if there isnt extract current date
         date = find_date_from_text(text)

     # check the text is supply or customer and return data
     if (cp == "") or (dn == ""):
        where = CUSTORMER
     else:
        where = SUPPLY
        if type == RECEIPT:
            if date == "":
                where = ""


     return cp, where, type, dn, date

"find cutomer information"
def find_cc(text,word):
    cc_text = ""
    cc = ""
    number = [1,2,3,4,5,6,7,8,9,0]

    if text.find("חשבונית") == -1:
        find_first = text.find(word)
        if find_first != -1:
            cc_text = text[find_first:find_first + 20]
            cc = re.findall(r'(\d+\d+\d+\d+\d+)', cc_text)
            if (len(cc) > 0) and (len(str(cc[0])) == 5):
                cc = cc[0]
            else:
                cc = ""
    else: # in case the file is RECEIPT
        cc = re.findall(r'(\d+\d+\d+\d+\d+)', text)
        if (len(cc) > 0):
            for i in cc:
                if len(str(i)) == 5:
                    cc = i
                    break
        if (len(cc) == "") or (len(cc[0]) != 5):
            cc = ""
        else:
            for i in range(len(number)):
                if (cc[0] == str(number[i])) and (i > 6):
                    cc = ""
                    break

    if cc == "":
        find_first = text.find("10000")
        try:
            for i in number:
                if text[find_first+5] == str(i):
                    find_first = -1
                    break
        except:
            pass
        if find_first != -1:
            return 10000

    return cc
def find_type_nameDoc(text,word):
    name_doc = ""
    find_first = text.find(word)
    if find_first != -1:
        name_doc = text[find_first:find_first + 30]
        name_doc = re.findall(r'(\d+\d+\d+\d+\d+\d+)', name_doc)
        if len(name_doc) > 0:
            name_doc = name_doc[0]
        else:
            name_doc = ""
    if name_doc == "":
         find_second = text.find(word, text.find(word) + 1)
         if find_second != -1:
            name_doc = text[find_second:find_second + 30]
            name_doc = re.findall(r'(\d+\d+\d+\d+\d+\d+)', name_doc)
            if len(name_doc) > 0:
                name_doc = name_doc[0]
            else:
                name_doc = ""

    return name_doc
# find from the text customer code number, transfer document and date - return all data
def find_in_text_imformation_cc_type_nameDoc_date(text,date):
    cc = ""
    name_doc = ""

    cc = find_cc(text,"לקוח")
    if cc == "":
        cc = find_cc(text,"מספר")

    if cc != "":
       name_doc = find_type_nameDoc(text,"משלוח")
       if name_doc == "":
           name_doc = find_type_nameDoc(text,"חשבונית")
           if name_doc == "":
               name_doc = find_type_nameDoc(text,"הפרשים")
    if (name_doc != "")  and (date == ""):
        date = find_date_from_text(text)

    if (cc == "") or (name_doc == "") or (date == "") :
        where = ""
    else:
        where = CUSTORMER

    return cc,name_doc,date, where

"create folder and transfer file"
def create_folder(name,directory_name):
    parent_dir = directory_name
    path = os.path.join(parent_dir, name)
    os.mkdir(path)
    return path
#find the folder of the compeny if not create one retrun path
def directory_exist(name,where,type = ""):
    name = name.replace("\"","")
    if where == SUPPLY: #in case it's about supply PDF
        if type == RECEIPT:
            directory_name = PATH_SUPPLY_INV +"\\"
            os.chdir(PATH_SUPPLY_INV)
        else: # type = SHIPPING_CERTIFICATE
            directory_name = PATH_SUPPLY_SHPPING + "\\"
            os.chdir(PATH_SUPPLY_SHPPING)
    else:  #In case it's about customer pdf
        directory_name = PATH_CUSTOMER + "\\"
        os.chdir(PATH_CUSTOMER)

    folder_directory_name = directory_name + name
    if os.path.isdir('./'+ name):
        print("The folder \""+name + "\" is exist")
        return folder_directory_name
    else:
        print("The folder \"" + name + "\" was created")
        return create_folder(name,directory_name)
#create month folder inside the path_directory if itnt exist
def directory_month(path_directory, date, where):
    month = {
        "01": "ינואר",
        "02": "פברואר",
        "03": "מרץ",
        "04": "אפריל",
        "05": "מאי",
        "06": "יוני",
        "07": "יולי",
        "08": "אוגוסט",
        "09": "ספטמבר",
        "10": "אוקטובר",
        "11": "נובמבר",
        "12": "דצמבר"}
    if where == SUPPLY:
        if date != "":
            if (len(date) != ""):
                if date[1] == '/':
                    date = '0' + date
                current_month = str(date[3]) + str(date[4])

        if (date == "") or (current_month.find("/") != -1):
            current_month = datetime.datetime.now().month
            if ((current_month !=10) & (current_month !=11) & (current_month !=12)):
                current_month = '0' + str(current_month)

    else:  # in case where == customer
        if(date != ""):
            if date[1] == '/':
                date = '0' + date
            current_month = str(date[3]) + str(date[4])
        else:
            current_month = datetime.datetime.now().month
            if ((current_month != 10) & (current_month != 11) & (current_month != 12)):
                current_month = '0' + str(current_month)

    month_name = month[str(current_month)]
    os.chdir(path_directory)
    path_month_folder = path_directory + "\\"
    path_month_folder = path_month_folder + month_name
    if os.path.isdir('./'+ month_name):
        print("The folder \""+month_name + "\" is exist")
        return path_month_folder
    else:
        print("The folder \"" + month_name + "\" was created")
        os.mkdir('./'+ month_name)
        return path_month_folder
    #move pdf to path_directory it belong
def move_file_to_directory_and_change_name(file_name,path_directory, name_doc, where):
    if where == SUPPLY:
        name = str(name_doc) + '.pdf'
    else:
        name = str(name_doc) + '.pdf'
    path_file_name = fix_path_name(file_name,PATH_DEFAULT)
    path_current_day = fix_path_name(name, PATH_DEFAULT)
    try:
        if os.path.isfile(path_current_day):
            print("The file already exists")
        else:
            # Rename the file
                os.rename(path_file_name, path_current_day)

        file_name = name
        src_path = fix_path_name(file_name,PATH_DEFAULT)
        dst_path = fix_path_name(file_name,path_directory)
        shutil.move(src_path, dst_path)
        flage = 1
        return  flage
    except:
        print("please close all of the pdf files")
        return 0

"mange the extract imformation from text & mange transform to file and name the documents"
def transfer_file_name_it(file_name,c_p,where,type,name_doc,date,list_code_excel,list_name_excel):
    if where == "":
        text = Recognizing_text_from_the_images_using_OCR('heb_eng')
        if name_doc == "":
            name_doc = find_type_and_doc_name(text,"משלוח","heb_eng")
            if name_doc == "":
                name_doc = find_type_and_doc_name(text, "חשבונית", "heb_eng")


    for i in range(len(list_code_excel)):
        list_code_excel[i] = str(list_code_excel[i]).replace(".0", "")
        if str(list_code_excel[i]) == c_p:
            name_cp = list_name_excel[i]
            name_folder = name_cp + " - " + str(c_p)
            if where == SUPPLY:
                print(where + " - " + "type: " + type + " - Name folder: " + name_folder)
            if where == CUSTORMER:
                print(where + " - " + "Name folder: " + name_folder)
            path_directory = directory_exist(name_folder, where, type)
            path_directory = directory_month(path_directory, date, where)
            flage = move_file_to_directory_and_change_name(file_name, path_directory, name_doc, where)
            return flage
def find_match_in_excel_extract_data_and_transfer_file_name_it(text ,file_name):
    name_cp = ""
    num_cp = ""
    cc = ""
    cp = ""
    name_doc = ""

    # extract private company number and customer code from excel
    privates_companies, privates_companies_name, customer_code_name, customer_code = extract_from_excel()
    # find from text cp that not amrusi cp
    cp, where, type,name_doc,date = find_in_text_information_company_private(text)
    if where != SUPPLY:
        cc,name_doc,date, where = find_in_text_imformation_cc_type_nameDoc_date(text,date)

    if where == "":
        text = Recognizing_text_from_the_images_using_OCR('heb')
        if name_doc == "":
            type,name_doc = find_type_and_doc_name(text,"heb_eng")
        if cp == "":
            cp = second_call_cp(text)
        if cc == "":
            cc = find_cc(text, "לקוח")

        if (name_doc != "") and (cp != "") :
            if type == RECEIPT:
                if date != "":
                    where = SUPPLY
            else:
                where = SUPPLY
        if (name_doc != "") and ((cc != "") and (len(str(cc)) == 5)) and (where == ""):
            where = CUSTORMER

    print(SUPPLY + ": " + "type = " + type + " cp = "+str(cp) + " name_doc = "+ str(name_doc) + " date = " + str(date))
    print(CUSTORMER + ": " +"cc = " + str(cc) + " name_doc = " + str(name_doc) +" date = " + str(date))

    if where == SUPPLY:
        for i in range(len(privates_companies)):
            privates_companies[i] = str(privates_companies[i]).replace(".0", "")
            if str(privates_companies[i]) == cp:
                name_cp = privates_companies_name[i]
                name_folder = name_cp + " - " + str(cp)
                print(where + " - "+"type: "+type + " - Name folder: " + name_folder)
                path_directory = directory_exist(name_folder,where,type)
                path_directory = directory_month(path_directory, date,where)
                flage = move_file_to_directory_and_change_name(file_name, path_directory,name_doc,where)
                return flage

    if where == CUSTORMER:
        for i in range(len(customer_code)):
            if str(customer_code[i]) == cc:
                name_cc = customer_code_name[i]
                name_folder = name_cc + " - " + str(cc)
                print(CUSTORMER + " - " + "Name folder: " + name_folder)
                path_directory = directory_exist(name_folder,where,type)
                path_directory = directory_month(path_directory, date,where)
                flage = move_file_to_directory_and_change_name(file_name, path_directory,name_doc,where)
                return flage

    return 0













