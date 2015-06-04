#!/usr/bin/python2
# -*- coding: utf-8

import os, sys, xlrd, re, io, subprocess, urllib2
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator, TextConverter
from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTTextContainer, LTFigure, LTImage
from pytesseract import image_to_string
from PIL import Image
from io import BytesIO
from mylingv import searchSums

if len(sys.argv) <= 1:
   print("Usage: python test.py file1 [file2...]\n");

verbose = True

# Реквизиты нашей организации, при встрече в документах игнорируем их, ищем только данные контрагентов

our = {
u"ИНН": "7702818199",
u"КПП": "770201001", 
u"р/с": "40702810700120030086", 
u"БИК": "044525201", 
u"Корсчет": "30101810000000000201", 
}

drp = re.UNICODE | re.IGNORECASE # default regexp parameters

# Алгоритм работы:
# В первом проходе ищем надписи ИНН, КПП, БИК и переносим числа, следующие за ними в
# соответствующие поля. 
# Если не в первом проходе не найдены ИНН, КПП или БИК, тогда запускаем второй проход:
# Первое десятизначное число интерпретируем как ИНН, первое девятизначное с первыми четырьмя 
# цифрами, совпадающими с ИНН - как КПП, первое девятизначное, начинающиеся с 04 -- как БИК
#

def epsilonEquals(a,b):
    if a == None or b == None: return False
    test = abs(a-b)
    return test < 0.0001

def fillTotal(pr, val):
    if val == None: return
    if epsilonEquals(val, pr.get(u"ИтогоБезНДС")): return
    if epsilonEquals(val, pr.get(u"ИтогоСНДС")): return
    oldTotal = pr.get(u"Итого")
    if oldTotal != None:
        if oldTotal != val:
            fillField(pr, u"ИтогоБезНДС", min(val, oldTotal))
            fillField(pr, u"ИтогоСНДС", max(val, oldTotal))
            del pr[u"Итого"]
    else:
        pr[u"Итого"] = val

def fillVatType(pr, content):
    if content == u"Без НДС" or content == u"Без налога (НДС)":
        fillField(u"СтавкаНДС", u"БезНДС")
        fillField(u"СуммаНДС", 0)
    if "18%" in content: fillField(pr, u"СтавкаНДС", u"18%")
    if "10%" in content: fillField(pr, u"СтавкаНДС", u"10%")
    if "0%"  in content:
        fillField(pr, u"СтавкаНДС", u"0%") 
        fillField(pr, u"СуммаНДС", 0)

innChk = [3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8]

# Проверка полей
def checkInn(val):
    if (len(val) != 10 and len(val) != 12) or not re.match("[0-9]+$", val): return False
    # Проверка контрольных цифр
    def innControlDigit(idx):
        if int(val[idx]) != sum([int(i)*c for i,c in zip(val[:idx],innChk[-idx:])])%11%10:
            if verbose: sys.stderr.write(u"Неверная контрольная цифра %i в ИНН: %r\n" % (idx+1, val))
            return False
        return True
    if len(val) == 10:
        return innControlDigit(9)
    return innControlDigit(10) and innControlDigit(11)
   
def checkKpp(val):
    if len(val) != 9 or not re.match("[0-9]+$", val): return False
    return True
def checkBic( val):
    if len(val) != 9 or not val.startswith("04") or not re.match("[0-9]+$", val): return False
    return True

checkDict = { u"ИНН": checkInn, u"КПП": checkKpp, u"БИК": checkBic }

class Err: pass

# Заполняет поле с именем fld, проверяет его, и проверяет, чтобы в поле уже не присутствовало другое значение
def fillField(pr, fld, value):
    if value == None: return
    ourVal = our.get(fld)
    if fld == u"ИНН" and value == ourVal: return
    oldVal = pr.get(fld)
    if isinstance(oldVal, Err): return
    if value == ourVal and oldVal != ourVal: return
    if value == oldVal: return
    check = checkDict.get(fld)
    if check != None and not check(value):
        if verbose: sys.stderr.write(u"%s: Найден некорректный %s: %s\n" % (pr["filename"], fld, value))
        return
    if oldVal != None and value != ourVal and oldVal != ourVal:
       if fld == u"Счет": return
       sys.stderr.write(u"%s: Найдено несколько различных %s: %s и %s\n" % (pr["filename"], fld, oldVal, value))
       pr[fld]=Err()
    pr[fld] = value


accChk = [7,1,3]*7+[7,1]
# Проверка ключа банковского счёта по БИКу и номеру счёта
def checkBicAcc(pr):
    if u"Корсчет" in pr:
       prefix = pr[u"БИК"][6:10]
    else:
       prefix = "0" + pr[u"БИК"][4:6]
    fullAcc = prefix + pr[u"р/с"]
    err = u"%s: Некорректный ключ номера счёта: %s\n" % (pr["filename"], fullAcc)
    if sum([int(i)*c for i,c in zip(fullAcc, accChk)]) % 10 != 0:
        sys.stderr.write(err)
        return False
    key = int(fullAcc[11])
    fullAcc = fullAcc[:11] + u"0" + fullAcc[12:]
    if sum([int(i)*c for i,c in zip(fullAcc, accChk)]) * 3 % 10 != key:
        sys.stderr.write(err)
        return False
    return True

# Убрать лишние данные из номера счёта
def stripInvoiceNumber(num):
    m = re.search(ur"\bот:? *[0-3]?\d(\.[0-1]\d\.| *[а-яА-Я]* )(20)?\d\d( *г([г\.]|ода?)?)?", num, drp)
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
    try:
       return float(re.sub(ur"руб\.?| ","", num).replace(",",".").strip())
    except ValueError:
       return None

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

    for fld in [u"ИНН", u"КПП", u"БИК"]:
        if re.match(u"[^a-zA-Zа-яА-Я]?"  + fld + u"\\b", content, drp):
            val = getSecondValue()
            if val == None: return
            rm = re.match("[0-9]+", val)
            if not rm: return
            val = rm.group(0)
            if val == our.get(fld): return
            fillField(pr, fld, val)
            return
    if re.match(u"Сч[её]т *(на оплату|№|.*? от [0-9][0-9]?[\\. ])", content, drp):
        text = content
        val, cell = getValueToTheRight(firstCell)
        while val != None:
            text += " "
            text += val
            val, cell = getValueToTheRight(cell)
        fillField(pr, u"Счет", stripInvoiceNumber(text.strip().replace("\n"," ").replace("  "," ")))
        return
    if re.match(u"ИНН */ *КПП\\b", content, drp):
        val = getSecondValue()
        if val == None: return
        rm = re.match("([0-9]{10}) */? *([0-9]{9})\\b", val, drp)
        if rm:
            rm = re.match("([0-9]{12}) */?\\b", val, drp)
            if rm == None: return
            checkInn(rm.group(1))
            fillField(pr, u"ИНН", rm.group(1))
        checkInn(rm.group(1))
        fillField(pr, u"ИНН", rm.group(1))
        checkKpp(rm.group(2))
        fillField(pr, u"КПП", rm.group(2))
        return
    if u" рубл" in content: findSumsInWords(content, pr)
    if (re.match(u"Итого|Всего", content, drp) and
            (re.search(u"(с|без) *НДС", content, drp) or not u"НДС" in content)):
        if ":" in content: val = getSecondValue(":")
        else: val = getValueToTheRight(firstCell)[0]
        if val == None: return
        fillTotal(pr, parse(val))
        return
    if u"НДС" in content:
        fillVatType(pr, content)
        if re.search(ur"(Всего|Итого|Сумма|Включая|в т\.ч\.|в том числе|НДС *\(?[0-9]*%\)? *:?).*?", content, drp):
            if ":" in content: val = getSecondValue(":")
            else: val = getValueToTheRight(firstCell)[0]
            if val == None: return
            if (re.match(ur"Без", val, drp)):
                fillField(pr, u"СтавкаНДС", u"БезНДС")
                fillField(pr, u"СуммаНДС", 0)
                return 
            fillField(pr, u"СуммаНДС", parse(val))
            return

    findBankAccounts(content, pr)

def findSumsInWords(text, pr):
    for psum in searchSums(text):
        if epsilonEquals(psum, pr.get(u"Итого")):
           del pr[u"Итого"]
        elif epsilonEquals(psum, pr.get(u"СуммаНДС")):
            return
        fillField(pr, u"ИтогоСНДС", psum)
        pr["СуммаПрописью"] = True

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
        for w in re.finditer(u"[0-9]{4} *[0-9]810 *[0-9]{4} *[0-9]{4} *[0-9]{4}\\b", text, drp):
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
        for val in re.finditer(nap + fld + u"\\n? *(%s)\\b" % regexp, text, drp):
            fillField(pr, fld, val.group(1).replace(u"О", "0"))

    rr = re.search(u"^ *Сч[её]т *(на оплату|№|.*?\\bот *[0-9]).*", text, drp | re.MULTILINE)
    if rr: fillField(pr, u"Счет", stripInvoiceNumber(rr.group(0)))

    # Поиск находящихся рядом пар ИНН/КПП с совпадающими первыми четырьмя цифрами
    if u"ИНН" not in pr and u"КПП" not in pr:
        results = re.findall(ur"([0-9О]{10}) *[\\\[\]\|/ ] *([0-9О]{9})\b", text, drp)
        results = [[v.replace(u"О", "0") for v in r] for r in results]
        for inn, kpp in results:
            if inn[0:4] == kpp[0:4]:
                fillField(pr, u"ИНН", inn)
                fillField(pr, u"КПП", kpp)
        if len(results)>0 and u"ИНН" not in pr and u"КПП" not in pr:
            for inn, kpp in results:
               if inn[0:2] == kpp[0:2]:
                   fillField(pr, u"ИНН", inn)
                   fillField(pr, u"КПП", kpp)
        if len(results)>0 and u"ИНН" not in pr and u"КПП" not in pr:
            # Пар с совпадающими первыми цифрами не найдено, вставляем любые пары
            for inn, kpp in results:
                fillField(pr, u"ИНН", inn)
                fillField(pr, u"КПП", kpp)


    # Если предыдущие шаги не дали никаких результатов, вставляем как ИНН, КПП и БИК
    # все подходящие цифры
    if u"ИНН" not in pr:
        for rm in re.finditer(nap + u"\\b([0-9О]{10}|[0-9О]{12})\\b" + bndry, text, drp):
            fillField(pr, u"ИНН", rm.group(1).replace(u"О", "0"))

    if u"БИК" not in pr:
        for rm in re.finditer(u"\\b([0О]4[0-9О]{7})\\b" + bndry, text, drp):
            fillField(pr, u"БИК", rm.group(1).replace(u"О", "0"))

    # Ищем КПП только если ИНН десятизначный
    if u"КПП" not in pr and (u"ИНН" not in pr or len(pr[u"ИНН"]) == 10):
        for rm in re.finditer(u"\\b([0-9О]{9})" + bndry, text, drp):
            val = rm.group(1).replace(u"О", "0")
            if val == pr.get(u"БИК"): continue
            fillField(pr, u"КПП", val)

    # Ищем итоги, ставки и суммы НДС
    for r in re.finditer(ur"Итого( [а-яА-Я ]*)?:?\n? *([0-9\., ]*)", text, drp):
        if r.group(1) == None or (re.match(u"(c|без) *НДС",r.group(1), drp) or not u"НДС" in r.group(1)):
            fillTotal(pr, parse(r.group(2).strip(".,")))
    for r in re.finditer(ur"(?:Всего |Итого |Сумма |в т\.ч\.|в том числе |включая ) *" +
            ur"НДС *(\([0-9%]*)?(?: [а-яА-Я \)]*)?\.?:?\n? *([0-9\., ]*)", text, drp):
        if r.group(1) != None: fillVatType(pr, r.group(1))
        fillField(pr, u"СуммаНДС", parse(r.group(2).strip(".,")))
    if re.search(ur"Без *(налога)? *\(?НДС", text, drp):
        fillField(pr, u"СтавкаНДС", u"БезНДС")

    findSumsInWords(text, pr)

def processImage(image, pr):
    debug = False
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
        sys.stderr.write("%s: Call to antiword failed, errcode is %i\n" % (filename, sp.poll()))
        return
    if debug:
        with open("invext-debug.xml","w") as f:
            f.write(stdoutdata)
    processText(stdoutdata.decode("utf-8"), pr)

def getBicData(bic):
    try:
        f = urllib2.urlopen("http://www.bik-info.ru/bik_%s.html" % bic)
        page = f.read().decode("cp1251")
    except URLError:
        return None
    finally:
        f.close()
    return {
        u"Корсчет": re.search(u"Корреспондентский счет: <b>(.*?)</b>", page).group(1),
        u"Наименование": re.search(u"Наименование банка: <b>(.*?)</b>", page).group(1),
    }

def finalizeAndCheck(pr):
    def deleteBank():
        for fld in [u"БИК", u"Корсчет", u"р/c"]:
            if fld in pr: del pr[fld]
    if u"БИК" in pr:
        bicData = getBicData(pr[u"БИК"])
        if bicData != None:
            if bicData[u"Корсчет"] != bicData.get(u"Корсчет", u""):
                sys.stderr.write(u"%s: Ошибка: не совпадает корсчёт\n" % pr["filename"])
                deleteBank()
            else:
                pr[u"Банк"] = bicData[u"Наименование"]
        else:
            sys.stderr.write(u"%s: Ошибка: не удалось получить данные по БИК\n" % pr["filename"])
            deleteBank()
        if u"р/с" in pr:
            if not checkBicAcc(pr):
                deleteBank()

def printInvoiceData(pr):
    if u"Счет" in pr:
        print(pr[u"Счет"])
    for fld in [u"ИНН", u"КПП", u"р/с", u"Банк", u"БИК", u"Корсчет",
            u"ИтогоБезНДС", u"СтавкаНДС", u"СуммаНДС", u"Итого", u"ИтогоСНДС"]:
        val = pr.get(fld)
        if val != None and not isinstance(val, Err):
            print("%s: %s" % (fld, val))
    if not pr.get("СуммаПрописью"):
        sys.stderr.write(u"%s: Предупреждение: сумма прописью не найдена\n" % pr["filename"])


for i in range(1,len(sys.argv)):
    f, ext = os.path.splitext(sys.argv[i])
    f = sys.argv[i]
    print(f)
    ext = ext.lower()
    pr = { "filename": f} # Parse result
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
    finalizeAndCheck(pr)
    printInvoiceData(pr)
    sys.stdout.flush()
