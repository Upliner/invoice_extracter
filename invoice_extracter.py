#!/usr/bin/python2
# -*- coding: utf-8

import os, sys, xlrd, re, io, subprocess
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator, TextConverter
from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTTextContainer, LTImage
from pytesseract import image_to_string
from PIL import Image
from io import BytesIO

if len(sys.argv) <= 1:
   print("Usege: python test.py file1 [file2...]\n");

# Реквизиты нашей организации, при встрече в документах игнорируем их, ищем только данные контрагентов

our = {
u"ИНН": "7702818199",
u"КПП": "770201001", 
u"р/с": "40702810700120030086", 
u"БИК": "044525201", 
u"Корсчет": "30101810000000000201", 
}

# Алгоритм работы:
# В первом проходе ищем надписи ИНН, КПП, БИК и переносим числа, следующие за ними в
# соответствующие поля. 
# Если не в первом проходе не найдены ИНН, КПП или БИК, тогда запускаем второй проход:
# Первое десятизначное число интерпретируем как ИНН, первое девятизначное с первыми четырьмя 
# цифрами, совпадающими с ИНН - как КПП, первое девятизначное, начинающиеся с 04 -- как БИК
#

class InvParseException(Exception):
    def __init__(self, message):
        super(InvParseException, self).__init__(message)

# Заполняет поле с именем fld и проверяет, чтобы в нём уже не присутствовало другое значение
def fillField(pr, fld, value):
    ourVal = our.get(fld)
    if value == None or (fld == u"ИНН" and value == ourVal): return
    oldVal = pr.get(fld)
    if value == ourVal and oldVal != ourVal: return
    if oldVal != None and oldVal != value and value != ourVal and oldVal != ourVal:
       if fld == u"Счет": return
       raise InvParseException(u"Найдено несколько различных %s: %s и %s" % (fld, oldVal, value))
    pr[fld] = value

# Проверка полей
def checkInn(val):
    if len(val) != 10 and len(val) != 12:
        raise InvParseException(u"Найден некорректный ИНН: %r" % val)
def checkKpp(val):
    if len(val) != 9:
        raise InvParseException(u"Найден некорректный КПП: %s" % val)
def checkBic(val):
    if len(val) != 9 or not val.startswith("04"):
        raise InvParseException(u"Найден некорректный БИК: %s" % val)

def processXlsCell(sht, row, col, pr):
    content = unicode(sht.cell_value(row, col))
    def getValueToTheRight(col):
        col = col + 1
        val = None
        while col < sht.ncols:
            val = sht.cell_value(row, col)
            if val != None and val != "" and val != 0: break
            col = col + 1
        if content == u"БИК" and type(val) in [int, float] and 40000000 <= val < 50000000:
            val = "0" + unicode(val) # Исправление БИКа в некоторых xls-файлах
        elif val != None: val = unicode(val)
        return (val, col)
    return processCellContent(content, getValueToTheRight, col, pr)

# Находит ближайший LTTextLine справа от указанного
def pdfFindRight(pdf, pl, distTresh = 10000):
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
    if result != None and result.x0 - pl.x1 > distTresh:
        return None # Ближайший текстбокс слишком далеко от текущего
    return result
    
def processPdfLine(pdf, pl, pr):
    content = pl.get_text()
    #print((("(%i,%i,%i,%i) %s") % (pl.x0, pl.y0, pl.x1, pl.y1, content)).encode("utf-8"))
    def getValueToTheRight(pdfLine):
        pdfLine = pdfFindRight(pdf, pdfLine, 10000 if pdfLine == pl else 216)
        if pdfLine == None: return (None, None)
        return (pdfLine.get_text(), pdfLine)
    return processCellContent(content, getValueToTheRight, pl, pr)

def processCellContent(content, getValueToTheRight, firstCell, pr):
    def getSecondValue():
        # Значение может находиться как в текущей ячейке, так и в следующей
        try:
            return content.split(None, 2)[1]
        except IndexError:
            # В данной ячейке данных не найдено, проверяем ячейки/текстбоксы справа
            return getValueToTheRight(firstCell)[0]
        #if val and len(val)<9 and re.match("[0-9]+$", val):
        #    # Ни одно искомое числовое значение не может иметь меньше девяти знаков
        #    # Предполагаем, что число может продолжаться в соседних ячейках (актуально для PDF)
        #    v2, cell = getValueToTheRight(cell)
        #    while v2 and re.match("[0-9]+$", v2):
        #        val += v2
        #        v2, cell = getValueToTheRight(cell)
        #return val

    for fld, check in [[u"ИНН", checkInn], [u"КПП", checkKpp], [u"БИК", checkBic]]:
        if re.match(u"[^a-zA-Zа-яА-Я]?"  + fld + u"\\b", content, re.UNICODE | re.IGNORECASE):
            val = getSecondValue()
            if val == None: return False
            rm = re.match("[0-9]+", val)
            if not rm: return False
            val = rm.group(0)
            if val == our.get(fld): return False
            check(val)
            fillField(pr, fld, val)
            return True
    if re.match(u" *Сч[её]т *(на оплату|№|.*? от [0-9][0-9]?[\\. ])", content, re.UNICODE | re.IGNORECASE):
        text = content
        val, cell = getValueToTheRight(firstCell)
        while val != None:
            text += " "
            text += val
            val, cell = getValueToTheRight(cell)
        fillField(pr, u"Счет", text.strip().replace("\n"," ").replace("  "," "))
        return True
    if re.match(u"ИНН */ *КПП\\b", content, re.UNICODE | re.IGNORECASE):
        val = getSecondValue()
        if val == None: return False
        rm = re.match("([0-9]{10}) */ *([0-9]{9})\\b", val, re.UNICODE)
        if rm:
            rm = re.match("([0-9]{12}) */?\\b", val, re.UNICODE)
            if rm == None: return False
            checkInn(rm.group(1))
            fillField(pr, u"ИНН", rm.group(1))
        checkInn(rm.group(1))
        fillField(pr, u"ИНН", rm.group(1))
        checkKpp(rm.group(2))
        fillField(pr, u"КПП", rm.group(2))
    return False

def findBankAccounts(text, pr):
    def processAcc(w):
       if w[0] == "4":
           fillField(pr, u"р/с", w)
       if w[0:5] == "30101":
           fillField(pr, u"Корсчет", w)
    hasIncomplete = False
    for w in text.split():
        if w == "3010": hasIncomplete = True
        if len(w) == 20 and re.match(u"[0-9О]{20}", w):
            w = w.replace(u"О", "0") # Многие OCR-движки путают букву О с нулём
            if w[5:8] == "810":
                processAcc(w)
    
    if not u"р/с" in pr or hasIncomplete:
        # В некоторых документах р/с написан с пробелами
        for w in re.finditer(u"[0-9]{4} *[0-9]810 *[0-9]{4} *[0-9]{4} *[0-9]{4}\\b", text, re.UNICODE):
            processAcc(w.group(0).replace(" ", ""))

nap = u"[^a-zA-Zа-яА-Я]?" # Non-alpha prefix
bndry = u"(?:\\b|[a-zA-Zа-яА-Я ])"

def processText(text, pr):
    if not u"р/с" in pr:
        findBankAccounts(text, pr)
    for fld, regexp in [
            [u"ИНН", u"[0-9О]{10}|[0-9О]{12}"],
            [u"КПП", u"[0-9О]{9}"],
            [u"БИК", u"[0О]4[0-9О]{7}"]]:
        for val in re.finditer(nap + fld + u" *(%s)\\b" % regexp, text, re.UNICODE | re.IGNORECASE):
            fillField(pr, fld, val.group(1).replace(u"О", "0"))

    rr = re.search(u"^ *Сч[её]т *(на оплату|№|.*?\\bот[0-9 ].*).*", text, re.UNICODE | re.IGNORECASE | re.MULTILINE)
    if rr: fillField(pr, u"Счет", rr.group(0))

    # Поиск находящихся рядом пар ИНН/КПП с совпадающими первыми четырьмя цифрами
    if u"ИНН" not in pr and u"КПП" not in pr:
        results = re.findall(u"([0-9О]{10}) *[\\\\\\[\\]\\|/ ] *([0-9О]{9})\\b", text, re.UNICODE)
        results = [r for r in [v.replace(u"О", "0") for v in results]]
        for inn, kpp in results:
            if inn[0:4] == kpp[0:4]:
                fillField(pr, u"ИНН", inn)
                fillField(pr, u"КПП", kpp)
        if len(results)>0 and u"ИНН" not in pr and u"КПП" not in pr:
            for inn, kpp in results:
                fillField(pr, u"ИНН", inn)
                fillField(pr, u"КПП", kpp)
            # Пар с совпадающими первыми цифрами не найдено, вставляем любые пары
            

    # Если предыдущие шаги не дали никаких результатов, вставляем как ИНН, КПП и БИК
    # первые подходящие цифры
    if u"ИНН" not in pr:
        rm = re.search(nap + u"\\b([0-9О]{10}|[0-9О]{12})\\b" + bndry, text, re.UNICODE)
        if rm: fillField(pr, u"ИНН", rm.group(1).replace(u"О", "0"))
    #Ищем КПП только если ИНН десятизначный
    if u"КПП" not in pr and (u"ИНН" not in pr or len(pr[u"ИНН"]) == 10):
        rm = re.search(u"\\b([0-9О]{9})" + bndry, text, re.UNICODE)
        if rm: fillField(pr, u"КПП", rm.group(1).replace(u"О", "0"))
    if u"БИК" not in pr:
        rm = re.search(u"\\b([0О]4[0-9О]{7})\\b" + bndry, text, re.UNICODE)
        if rm: fillField(pr, u"БИК", rm.group(1).replace(u"О", "0"))

def processImage(image, pr):
    debug = True
    text = image_to_string(image, lang="rus").decode("utf-8")
    if debug:
        with open("invext-debug.txt","w") as f:
            f.write(text.encode("utf-8"))
    processText(text, pr)

def processPDF(f, pr):
        debug = False
        parser = PDFParser(f)
        document = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        laparams = LAParams()
        daggr = PDFPageAggregator(rsrcmgr, laparams=laparams)
        parsedTextStream = BytesIO()
        dtc = TextConverter(rsrcmgr, parsedTextStream, codec="utf-8", laparams=laparams)
        iaggr = PDFPageInterpreter(rsrcmgr, daggr)
        itc = PDFPageInterpreter(rsrcmgr, dtc)
        for page in PDFPage.create_pages(document):
            iaggr.process_page(page)
            layout = daggr.get_result()
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
                # В pdf-файле отсутствует текст, возможно есть картинки, которые можно прогнать через OCR
                for obj in layout:
                    if isinstance(obj, LTImage):
                        processImage(Image.open(BytesIO(obj.stream.get_rawdata())))
            else:
                if u"р/с" not in pr or u"ИНН" not in pr or u"КПП" not in pr or u"БИК" not in pr or u"Счет" not in pr:
                    # Текст в файле есть, но его не удалось полностью распознать, используем fallback метод
                    itc.process_page(page)
                    text = parsedTextStream.getvalue().decode("utf-8")
                    if debug:
                        with open("invext-debug.txt","w") as f:
                            f.write(text.encode("utf-8"))
                    processText(text, pr)
                    parsedTextStream = BytesIO()

def processExcel(filename, pr):
    wbk = xlrd.open_workbook(filename)
    for sht in wbk.sheets():
        for row in range(sht.nrows):
            for col in range(sht.ncols):
                if not processXlsCell(sht, row, col, pr):
                    findBankAccounts(unicode(sht.cell_value(row,col)), pr)

# TODO: Парсинг из XML, без конвертации в PDF
def processMsWord(filename, pr):
    sp = subprocess.Popen(["antiword", "-a", "a4", filename], stdout=subprocess.PIPE, stderr=sys.stderr)
    stdoutdata, stderrdata = sp.communicate()
    if sp.poll() != 0:
        print("Call to antiword failed, errcode is " + sp.poll())
        return
    processPDF(io.BytesIO(stdoutdata), pr)

def printInvoiceData(pr):
    if u"Счет" in pr:
        print(pr[u"Счет"])
    for fld in [u"ИНН", u"КПП", u"р/с", u"БИК", u"Корсчет"]:
        val = pr.get(fld)
        if (val != None):
            print("%s: %s" % (fld, val))
            

for i in range(1,len(sys.argv)):
    f, ext = os.path.splitext(sys.argv[i])
    f = sys.argv[i]
    print(f)
    ext = ext.lower()
    pr = {} # Parse result
    try:
        if (ext in ['.png','.bmp','.jpg','.gif']):
            processImage(Image.open(f), pr)
        elif (ext == '.pdf'):
            with open(f, "rb") as f: processPDF(f, pr)
        elif (ext in ['.xls', '.xlsx']):
            processExcel(f, pr)
        elif (ext in ['.doc']):
            processMsWord(f, pr)
        elif (ext in ['.txt']):
            with open(f, "rb") as f: processText(f.read().decode("utf-8"), pr)
        else:
            sys.stderr.write("%s: unknown extension\n" % f)
    except InvParseException as e:
        print(unicode(e))
    printInvoiceData(pr)
    sys.stdout.flush()
