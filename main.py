# coding: utf-8
from ocr_non_english import *


def main():
    count = 0
    files_names = get_names_of_all_the_pdf_file()
    for i in range(len(files_names)):
        flage = 1
        text = ""
        fm = files_names[i]  # Filename
        print("working on - "+fm+" file")


        try:
            Converting_PDF_to_jpeg(fm)  # create image name "page_1.png"
            improve_contrast()
            text = Recognizing_text_from_the_images_using_OCR("heb_eng")

        except:
            print("Error: Converting_PDF_to_jpeg or Recognizing_text_from_the_images_using_OCR")
            flage = 0
            pass

        if flage != 0:
           flage = find_match_in_excel_extract_data_and_transfer_file_name_it(text, files_names[i])
        if flage == 0:
            print("The sort didn't Succeeded please sort manualy")
            print("")
        else:
            print("The sort Succeeded  (^o^)(^_^ ) ")
            count = count +1
            print("")
    print("")
    print("^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^")
    print("Number of PDF: "+str(len(files_names))+" | Number of success: "+ str(count))
    print("^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^")

if __name__ == "__main__":
    main()
