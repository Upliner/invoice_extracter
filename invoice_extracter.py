#!/usr/bin/python2
# -*- coding: utf-8

import os, sys, xlrd
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTTextContainer, LTImage
from pytesseract import image_to_string
from PIL import Image

if len(sys.argv) <= 1:
   print("Usege: python test.py file1 [file2...]\n");

our_inn = 7702818199
our_kpp = 770201001

class InvParseException(Exception):
    def __init__(self, message):
        super(InvParseException, self).__init__(message)
class ParseResult:
    def __init__(self):
        self.inv = None
        self.inn = None
        self.kpp = None
        self.acc = None
        self.acc_cor = None
        self.bic = None
    def prn(self):
        if self.inv != None: print(u"%s" % self.inv);
        if self.inn != None: print(u"ИНН: %s" % self.inn);
        if self.kpp != None: print(u"КПП: %s" % self.kpp);
        if self.acc != None: print(u"Счёт: %s" % self.acc);
        if self.bic != None: print(u"БИК: %s" % self.bic);
        if self.acc_cor != None: print(u"Корсчёт: %s" % self.acc_cor);

def processCell(content, pr):
    if content.startswith(u"ИНН "):
        try:
            inn = content.split()[1]
        except IndexError:
            return False
        if inn == our_inn: return
        if len(inn) != 10 and len(inn) != 12:
            raise InvParseException(u"Найден некорректный ИНН: " + inn)
        if pr.inn != None and pr.inn != inn:
            raise InvParseException(u"Найдено несколько различных ИНН")
        pr.inn = inn
        return True
    if content.startswith(u"КПП "):
        try:
            kpp = content.split()[1]
        except IndexError:
            return False
        if (kpp == our_kpp): return
        if len(kpp) != 9:
            raise InvParseException(u"Найден некорректный КПП: " + kpp)
        if pr.kpp != None and pr.inn != kpp:
            raise InvParseException(u"Найдено несколько различных КПП")
        pr.kpp = kpp
        return True
    if content.startswith(u"Счет на оплату"):
        pr.inv = content
        return True
    return False

def processWords(words, pr):
    for w in words:
        if len(w) == 20 and w[5:8] == "810":
           if w[0] == "4":
               if pr.acc != None and pr.acc != w:
                  raise InvParseException(u"Найдено несколько различных банковских счетов")
               pr.acc = w
           if w[0:5] == "30101":
               if pr.acc != None and pr.acc != w:
                  return
                  #raise InvParseException(u"Найдено несколько корсчетов")
               pr.acc_cor = w
        if len(w) == 9 and w[0:2] == "04":
            pr.bic = w

def processImage(filename, pr):
    processWords(image_to_string(Image.open(filename), lang="rus").split(), pr)

def processPDF(filename, pr):
    with open(filename, "rb") as f:
        parser = PDFParser(f)
        document = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.create_pages(document):
            interpreter.process_page(page)
            layout = device.get_result()
            for obj in layout:
                if isinstance(obj, LTTextBox):
                    txt = obj.get_text()
                    foundInfo = False
                    for line in txt.split("\n"):
                        if processCell(line, pr):
                            foundInfo = True
                    if not foundInfo: processWords(txt.split(), pr)

def processExcel(filename, pr):
    wbk = xlrd.open_workbook(filename)
    for sht in wbk.sheets():
        for row in range(sht.nrows):
            for col in range(sht.ncols):
                cell = unicode(sht.cell_value(row,col));
                if not processCell(cell, pr):
                    processWords(cell.split(), pr)

for i in range(1,len(sys.argv)):
    f, ext = os.path.splitext(sys.argv[i])
    f = sys.argv[i]
    ext = ext.lower()
    p = ParseResult()
    if (ext in ['.png','.bmp','.jpg','.gif']):
        processImage(f, p)
    elif (ext == '.pdf'):
        processPDF(f, p)
    elif (ext in ['.xls', '.xlsx']):
        processExcel(f, p)
    else:
        sys.stderr.write("%s: unknown extension\n" % f)
    p.prn()