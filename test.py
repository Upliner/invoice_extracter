#!/usr/bin/python3

import os, sys, xlrd
from pytesseract import image_to_string
from PIL import Image
from PyPDF2 import PdfFileReader

if len(sys.argv) <= 1:
   print("Usege: python test.py file1 [file2...]\n");

def processWords(words):
    for w in words:
        if len(w) == 20 and '810' in w: print(w);


def processImage(filename):
    processWords(image_to_string(Image.open(filename)).split())

def processPDF(filename):
    with open(filename, "rb") as f:
        rdr = PdfFileReader(f)
        for i in range(0,rdr.getNumPages()):
            processWords(rdr.getPage(i).extractText().split())

def processExcel(filename):
    wbk = xlrd.open_workbook(filename)
    for sht in wbk.sheets():
        for row in range(sht.nrows):
            for col in range(sht.ncols):
                processWords(str(sht.cell_value(row,col)).split())

for i in range(1,len(sys.argv)):
    f, ext = os.path.splitext(sys.argv[i])
    f = sys.argv[i]
    ext = ext.lower()
    if (ext in ['.png','.bmp','.jpg','.gif']):
        processImage(f)
    elif (ext == '.pdf'):
        processPDF(f)
    elif (ext in ['.xls', '.xlsx']):
        processExcel(f)
    else:
        sys.stderr.write("%s: unknown extension\n" % f)