#!/usr/bin/python

import os, sys
from pytesseract import image_to_string
from PIL import Image
from PyPDF2 import PdfFileReader

if len(sys.argv) <= 1:
   print("Usege: python test.py file1 [file2...]\n");

def processWords(words):
    for w in words:
        if '8888' in w: print(w);


def processImage(filename):
    processWords(image_to_string(Image.open(filename)).split())

def processPDF(filename):
    with open(filename, "rb") as f:
        rdr = PdfFileReader(f)
        for i in range(0,rdr.getNumPages()):
            processWords(rdr.getPage(i).extractText().split())

for i in range(1,len(sys.argv)):
    f, ext = os.path.splitext(sys.argv[i])
    f = sys.argv[i]
    ext = ext.lower()
    if (ext in ['.png','.bmp','.jpg','.gif']):
        processImage(f)
    elif (ext == '.pdf'):
        processPDF(f)
    else:
        sys.stderr.write("%s: unknown extension\n" % f)