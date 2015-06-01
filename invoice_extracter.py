#!/usr/bin/python2
# -*- coding: utf-8

import os, sys, xlrd, re, io
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

# Данные нашей организации, при встрече в документах игнорируем их, ищем только данные контрагентов

our = {
u"ИНН": "7702818199",
u"КПП": "770201001", # КПП может быть одинаковый у нас и у контрагента
}


# Алгоритм работы:
# В первом проходе ищем надписи ИНН, КПП, БИК и переносим числа, следующие за ними в
# соответствующие поля. 
# Если не в первом проходе не найдены ИНН, КПП или БИК, тогда запускаем второй проход:
# Первое десятизначное числоинтерпретируем как ИНН, первое девятизначное с первыми четырьмя 
# цифрами, совпадающими с ИНН - как КПП, первое девятизначное, начинающиеся с 04 -- как БИК
#

class InvParseException(Exception):
    def __init__(self, message):
        super(InvParseException, self).__init__(message)

# Заполняет поле с именем fld и проверяет, чтобы в нём уже не присутствовало другое значение
def fillField(pr, fld, value):
    if value == None or (fld != u"КПП" and value == our.get(fld)): return
    oldVal = pr.get(fld)
    if oldVal != None and oldVal != value and value != our.get(fld) and oldVal != our.get(fld):
       raise InvParseException(u"Найдено несколько различных %s: %s и %s" % (fld, oldVal, value))
    pr[fld] = value

# Проверка полей
def checkInn(val):
    if len(val) != 10 and len(val) != 12:
        print(val)
        raise InvParseException(u"Найден некорректный ИНН: %r" % val)
def checkKpp(val):
    if len(val) != 9:
        raise InvParseException(u"Найден некорректный КПП: %s" % val)
def checkBic(val):
    if len(val) != 9 or not val.startswith("04"):
        raise InvParseException(u"Найден некорректный БИК: %s" % val)

# Находит ближайший LTTextLine справа от указанного
def pdfFindRight(pdf, pl):
    y = (pl.y0 + pl.y1) / 2
    result = None
    for obj in pdf:
        if not isinstance(obj, LTTextBox): continue
        if obj.y0 > y or obj.y1 < y: continue
        for line in obj:
            if not isinstance(line, LTTextLine): continue
            if line.y0 > y or line.y1 < y or line.x0<=pl.x0: continue
            if result != None and result.x0 <= obj.x0: continue
            result = line
    return result

def processPdfLine(pdf, pl, pr):
    content = pl.get_text()
    def getSecondValue():
        try:
            return content.split(None, 2)[1]
        except IndexError:
            # В данном текстбоксе данных не найдено, проверяем текстбокс справа
            ntl = pdfFindRight(pdf, pl)
            if ntl == None: return None
            return ntl.get_text()
    for fld, check in [[u"ИНН", checkInn], [u"КПП", checkKpp], [u"БИК", checkBic]]:
        if re.match(fld + "[: ]", content):
            val = getSecondValue()
            if val == None: return False
            rm = re.match("[0-9]+", val)
            if not rm: return False
            val = rm.group(0)
            if val == our.get(fld): return False
            check(val)
            fillField(pr, fld, val)
            return True
    if re.match((u"Счет *(на оплату)? *№"), content):
        fillField(pr, u"Счет", content)
        return True
    if content.startswith(u"ИНН/КПП:"):
        val = getSecondValue()
        if val == None: return False
        rm = re.match("([0-9]{10}) */ *([0-9]{9})", val)
        if rm == None:
            rm = re.match("([0-9]{12}) */?", val)
            if rm == None: return False
            checkInn(rm.group(1))
            fillField(pr, u"ИНН", rm.group(1))
        checkInn(rm.group(1))
        fillField(pr, u"ИНН", rm.group(1))
        checkKpp(rm.group(2))
        fillField(pr, u"КПП", rm.group(2))
    return False

def findBankAccounts(text, pr):
    for w in text.split():
        if len(w) == 20 and w[5:8] == "810" and re.match("[0-9]{20}", w):
           if w[0] == "4":
               fillField(pr, u"р/с", w)
           if w[0:5] == "30101":
               fillField(pr, u"Корсчет", w)

def processImage(image, pr):
    findBankAccounts(image_to_string(image, lang="rus"), pr)

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
            hasText = False
            for obj in layout:
                if isinstance(obj, LTTextBox):
                    hasText = True
                    txt = obj.get_text()
                    foundInfo = False

                    for line in obj:
                        if isinstance(line, LTTextLine):
                            if processPdfLine(layout, line, pr):
                                foundInfo = True
                    if not foundInfo: findBankAccounts(txt, pr)
            if not hasText:
                # Текста pdf-файл не содержит, возможно содержит картинки, которые можно прогнать через OCR
                for obj in layout:
                    if isinstance(obj, LTImage):
                        processImage(Image.open(io.BytesIO(obj.stream.get_rawdata())))
                        

def processExcel(filename, pr):
    wbk = xlrd.open_workbook(filename)
    for sht in wbk.sheets():
        for row in range(sht.nrows):
            for col in range(sht.ncols):
                cell = unicode(sht.cell_value(row,col));
                if not processCell(cell, pr):
                    findBankAccounts(cell, pr)

def printInvoiceData(pr):
    if u"счет" in pr:
        print(pr[u"счет"])
    for fld in [u"ИНН", u"КПП", u"р/с", u"БИК", u"Корсчет"]:
        val = pr.get(fld)
        if (val != None):
            print("%s: %s" % (fld, val))
            

for i in range(1,len(sys.argv)):
    f, ext = os.path.splitext(sys.argv[i])
    f = sys.argv[i]
    print(f)
    ext = ext.lower()
    pr = {}
    try:
        if (ext in ['.png','.bmp','.jpg','.gif']):
            processImage(Image.open(f), pr)
        elif (ext == '.pdf'):
            processPDF(f, pr)
        elif (ext in ['.xls', '.xlsx']):
            processExcel(f, pr)
        else:
            sys.stderr.write("%s: unknown extension\n" % f)
    except InvParseException as e:
        print(unicode(e))
    printInvoiceData(pr)
