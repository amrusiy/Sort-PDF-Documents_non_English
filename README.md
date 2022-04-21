# Sort-PDF-Documents_non_English

Auto read Hebrew PDF and sort to appropriate folder.


## Getting Started

In general, the script converts a PDF to an image, improves the contrast and extracts the text.

Documents to be sorted:
  - Delivery certifications of suppliers
  - supplie receipts
  - Delivery certifications of customers 
  - Customer receipts
  
To sort, the script looks for a number of parameters:
  - Private company number
  - Customer number (for internal use by the business owner)
  - Name of the document (represented by a number)
  - Date

The types of folders to which the documents will be transferred:
  - Customer - contain all kind of customer file (Delivery certifications and receipts)
  - Supply
      - shipping certificate
      - receipt

In order to achieve high accuracy, I added an Excel document that contains all the information about the suppliers and customers(names, private company number
and customer number).
Before sorting the documents, we will check that the name of the supplier or customer exists inside the Excel document.
If all the parameters are found inside the document, the script creates a folder for the client\supply and month subdolder and transfere the file to the subfolder.
If the folder exists, the script will move the document to the folder under the relevant month subfolder.


### Prerequisites

Things you need to install on your workstation

- python 3
- pip
- tesseract v5.0.1 
- poppler

Here are some references for **poppler** and **tesserect**:

 - [blog.alivate.com.au/poppler-windows](https://blog.alivate.com.au/poppler-windows)
 - [tesseract-ocr.github.io/tessdoc/Home](https://tesseract-ocr.github.io/tessdoc/Home.html)

Some useful reference for pdf2image:
- pdf2image [https://pypi.org/project/pdf2image/](https://pypi.org/project/pdf2image/)


### Installing

A step by step guide to set up a development environment.

- Install **poppler** on your workstation.

- Install **tesseract** on your workstation.


## Code Explanation

The script is divided in 4 main sections.

First at ocr_non_english.py - **change all the path to fit in your enviroments!!!**

### Part #1 : Converting PDF to images and Recognizing text from the images using OCR

The PDF file is converted into jpgs.

Temporary jpg and txt files will be generated per pdf page.

More info is available in the **pdf2image** documentation above.

**Note**: You may need to add full path to **poppler** on some work stations:

With the help of ImageEnhance package we will improve the contrast of the image to improve the accuracy of the text extraction

OCR will extract text from the jpg files to txt files.

### Part #2 : Finding parmeters

Looking in the text for all the parameters I mentioned above

### Part #3 : Search again (option)

In case no parameters were found we will run the OCR again differently in order to further detailsץ

### Part #4 : Checking compatibility with Excel documents, creating a folder and transferring the document to the same folder

Checks whether the details extracted from the pdf exist in the Excel documet, if so, a folder has been created for the customer / supplier
and a subfolder with a suitable month.
then our PDF goes into that subfolder.


**Note**: The Excel documennt and all the PDF files are internal documents of the company thus they didnt attached.


## Authors

* [amrusiy](https://github.com/amrusiy?tab=repositories)





רק
