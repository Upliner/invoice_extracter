#!/usr/bin/python2
# -*- coding: utf-8

import os, sys, xlrd, re, io, subprocess
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator, TextConverter
from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTTextContainer, LTFigure, LTImage
from pytesseract import image_to_string
from PIL import Image
from io import BytesIO

if len(sys.argv) <= 1:
   print("Usage: python test.py file1 [file2...]\n");

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

# Убрать лишние данные из номера счёта
def stripInvoiceNumber(num):
    m = re.search(ur"\bот:? *[0-3]?\d(\.[0-1]\d\.| *[а-яА-Я]* )(20)?\d\d( *г([г\.]|ода?)?)?", num, re.UNICODE)
    if m: return num[0:m.end()]
    return num

def processXlsCell(sht, row, col, pr):
    try:
        content = unicode(sht.cell_value(row, col))
    except IndexError:
        return
    def getValueToTheRight(col):
        col = col + 1
        val = None
        while col < sht.ncols:
            try:
                val = sht.cell_value(row, col)
            except IndexError:
                return (None, None)
            if val != None and val != "" and val != 0: break
            col = col + 1
        if content == u"БИК" and type(val) in [int, float] and 40000000 <= val < 50000000:
            val = "0" + unicode(val) # Исправление БИКа в некоторых xls-файлах
        elif val != None: val = unicode(val)
        return (val, col)
    processCellContent(content, getValueToTheRight, col, pr)

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
    content = pl.get_text().strip()
    def getValueToTheRight(pdfLine):
        pdfLine = pdfFindRight(pdf, pdfLine)
        if pdfLine == None: return (None, None)
        return (pdfLine.get_text(), pdfLine)
    processCellContent(content, getValueToTheRight, pl, pr)

def parse(num):
    return float(re.sub(ur"руб\.?| ","", num).replace(",",".").strip())

def processCellContent(content, getValueToTheRight, firstCell, pr):
    def getSecondValue(separator = None):
        # Значение может находиться как в текущей ячейке, так и в следующей
        try:
            val = content.split(separator, 2)[1].strip()
        except IndexError:
            val = ""
        if len(val) > 0: return val
        # В данной ячейке данных не найдено, проверяем ячейки/текстбоксы справа
        return getValueToTheRight(firstCell)[0]

    for fld, check in [[u"ИНН", checkInn], [u"КПП", checkKpp], [u"БИК", checkBic]]:
        if re.match(u"[^a-zA-Zа-яА-Я]?"  + fld + u"\\b", content, re.UNICODE | re.IGNORECASE):
            val = getSecondValue()
            if val == None: return
            rm = re.match("[0-9]+", val)
            if not rm: return
            val = rm.group(0)
            if val == our.get(fld): return
            check(val)
            fillField(pr, fld, val)
            return
    if re.match(u"Сч[её]т *(на оплату|№|.*? от [0-9][0-9]?[\\. ])", content, re.UNICODE | re.IGNORECASE):
        text = content
        val, cell = getValueToTheRight(firstCell)
        while val != None:
            text += " "
            text += val
            val, cell = getValueToTheRight(cell)
        fillField(pr, u"Счет", stripInvoiceNumber(text.strip().replace("\n"," ").replace("  "," ")))
        return
    if re.match(u"ИНН */ *КПП\\b", content, re.UNICODE | re.IGNORECASE):
        val = getSecondValue()
        if val == None: return
        rm = re.match("([0-9]{10}) */? *([0-9]{9})\\b", val, re.UNICODE)
        if rm:
            rm = re.match("([0-9]{12}) */?\\b", val, re.UNICODE)
            if rm == None: return
            checkInn(rm.group(1))
            fillField(pr, u"ИНН", rm.group(1))
        checkInn(rm.group(1))
        fillField(pr, u"ИНН", rm.group(1))
        checkKpp(rm.group(2))
        fillField(pr, u"КПП", rm.group(2))
        return
    if (re.match(u"Итого|Всего", content, re.UNICODE | re.IGNORECASE) and
           (u"без НДС" in content or not u"НДС" in content)):
        if ":" in content:
            val = getSecondValue(":")
        else:
            val = getValueToTheRight(firstCell)[0]
        if val == None: return
        try:
            val = parse(val)
        except ValueError:
            return
        oldTotal = pr.get(u"Итого")
        if (val == pr.get(u"ИтогоБезНДС")): return
        if (val == pr.get(u"ИтогоСНДС")): return
        if oldTotal != None:
            if oldTotal != val:
                pr[u"ИтогоБезНДС"] = min(val, oldTotal)
                pr[u"ИтогоСНДС"] = max(val, oldTotal)
                del pr[u"Итого"]
        else:
            pr[u"Итого"] = val
        return
    if u"НДС" in content:
        if content == u"Без НДС" or content == u"Без налога (НДС)":
            pr[u"СтавкаНДС"] = u"БезНДС"
            pr[u"СуммаНДС"] = 0
        if "18%" in content: pr[u"СтавкаНДС"] = u"18%"
        if "10%" in content: pr[u"СтавкаНДС"] = u"10%"
        if "0%"  in content: pr[u"СтавкаНДС"] = u"0%"; pr[u"СуммаНДС"] = 0
        if re.match(ur"(Всего|Итого|Сумма|в т\.ч\.|в том числе).*?", content, re.UNICODE | re.IGNORECASE):
            if ":" in content:
                val = getSecondValue(":")
            else:
                val = getValueToTheRight(firstCell)[0]
            if (re.match(ur"Без", content, re.IGNORECASE)):
                pr[u"СтавкаНДС"] = u"БезНДС"
                pr[u"СуммаНДС"] = 0
                return 
            try:
                val = parse(val)
            except ValueError:
                return
            pr[u"СуммаНДС"] = val
            return
        
    findBankAccounts(content, pr)

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
        for val in re.finditer(nap + fld + u"\\n? *(%s)\\b" % regexp, text, re.UNICODE | re.IGNORECASE):
            fillField(pr, fld, val.group(1).replace(u"О", "0"))

    rr = re.search(u"^ *Сч[её]т *(на оплату|№|.*?\\bот *[0-9]).*", text, re.UNICODE | re.IGNORECASE | re.MULTILINE)
    if rr: fillField(pr, u"Счет", stripInvoiceNumber(rr.group(0)))

    # Поиск находящихся рядом пар ИНН/КПП с совпадающими первыми четырьмя цифрами
    if u"ИНН" not in pr and u"КПП" not in pr:
        results = re.findall(u"([0-9О]{10}) *[\\\\\\[\\]\\|/ ] *([0-9О]{9})\\b", text, re.UNICODE)
        results = [[v.replace(u"О", "0") for v in r] for r in results]
        for inn, kpp in results:
            if inn[0:4] == kpp[0:4]:
                fillField(pr, u"ИНН", inn)
                fillField(pr, u"КПП", kpp)
        if len(results)>0 and u"ИНН" not in pr and u"КПП" not in pr:
            # Пар с совпадающими первыми цифрами не найдено, вставляем любые пары
            for inn, kpp in results:
                fillField(pr, u"ИНН", inn)
                fillField(pr, u"КПП", kpp)
            

    # Если предыдущие шаги не дали никаких результатов, вставляем как ИНН, КПП и БИК
    # первые подходящие цифры
    if u"ИНН" not in pr:
        rm = re.search(nap + u"\\b([0-9О]{10}|[0-9О]{12})\\b" + bndry, text, re.UNICODE)
        if rm: fillField(pr, u"ИНН", rm.group(1).replace(u"О", "0"))
    # Ищем КПП только если ИНН десятизначный
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
        x0, y0, x1, y1 = (sys.maxint, sys.maxint, -sys.maxint, -sys.maxint) # Text bbox
        for obj in layout:
            if isinstance(obj, LTTextBox):
                x0 = min(x0, obj.x0)
                y0 = min(y0, obj.y0)
                x1 = max(x1, obj.x1)
                y1 = max(y1, obj.y1)
                for line in obj:
                    if isinstance(line, LTTextLine):
                        processPdfLine(layout, line, pr)
        if u"р/с" not in pr or u"ИНН" not in pr or u"КПП" not in pr or u"БИК" not in pr or u"Счет" not in pr:
            # Текст в файле есть, но его не удалось полностью распознать
            # Возможно это плохо распознанный PDF, ищем картинки, перекрывающие всю страницу
            for obj in layout:
                if isinstance(obj, LTFigure):
                    for img in obj:
                        if (isinstance(img, LTImage) and
                                img.x0<x0 and img.y0<y0 and img.x1>x1 and img.y1>y1):
                            processImage(Image.open(BytesIO(img.stream.get_rawdata())), pr)
                            break
            else:
                # Подходящих картинок нет, используем fallback метод
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
                processXlsCell(sht, row, col, pr)

def processMsWord(filename, pr):
    debug = False
    sp = subprocess.Popen(["antiword", "-x", "db", filename], stdout=subprocess.PIPE, stderr=sys.stderr)
    stdoutdata, stderrdata = sp.communicate()
    if sp.poll() != 0:
        print("Call to antiword failed, errcode is " + sp.poll())
        return
    if debug:
        with open("invext-debug.xml","w") as f:
            f.write(stdoutdata)
    processText(stdoutdata.decode("utf-8"), pr)

def printInvoiceData(pr):
    if u"Счет" in pr:
        print(pr[u"Счет"])
    for fld in [u"ИНН", u"КПП", u"р/с", u"БИК", u"Корсчет",
            u"ИтогоБезНДС", u"СтавкаНДС", u"СуммаНДС", u"Итого", u"ИтогоСНДС"]:
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
        elif (ext in ['.txt', '.xml']):
            with open(f, "rb") as f: processText(f.read().decode("utf-8"), pr)
        else:
            sys.stderr.write("%s: unknown extension\n" % f)
    except InvParseException as e:
        print(unicode(e))
    printInvoiceData(pr)
    sys.stdout.flush()
