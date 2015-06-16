#!/usr/bin/python2
# -*- coding: utf-8

import os, sys, xlrd, re, io, subprocess, urllib, urllib2, argparse, datetime, time, math
from pdfminer.psparser import PSLiteral
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator, TextConverter
from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTTextContainer, LTFigure, LTImage
from pytesseract import image_to_string
from PIL import Image, ImageOps
from io import BytesIO
from mylingv import searchSums, searchSumsFiltered
from xml.sax.saxutils import unescape

verbose = False
strict = False

devnull = open(os.devnull, "w")

def errWrite(s):
    sys.stderr.write(s.encode("utf-8"))

def parseArguments():
    parser = argparse.ArgumentParser(description='Extract data from invoices.')

    parser.add_argument("files", metavar="file", type=str, nargs="+", help="file to process")
    parser.add_argument("-v", "--verbose", dest="verbose", action='store_true',
                   help='verbose mode')
    parser.add_argument("-q", "--requisites", type=str, default = None, dest="reqfile",
                   help="specify requisites file")
    parser.add_argument("--strict", action='store_true', dest="strict",
                   help="remove KPP and coracc from output when mismatches with federal database")
    parser.add_argument("--inn", type=str, default = None, help="specify INN code")
    parser.add_argument("--kpp", type=str, default = None, help="specify KPP code")
    parser.add_argument("--acc", type=str, default = None, help="specify bank account")
    parser.add_argument("--bic", type=str, default = None, help="specify bank identification code")
    parser.add_argument("--coracc", type=str, default = None, help="specify bank transit account")

    args = parser.parse_args()

    # Принимаем реквизиты из файла, коммандной строки или берём по умолчанию
    if args.inn or args.kpp or args.acc or args.bic or args.coracc:
        if args.reqfile != None:
            errWrite(u"Ошибка: одновременно заданы реквизиты в файле и в коммандной строке\n")
            sys.exit(1)
        our = {}
        for a, o in [("inn", u"ИНН"), ("kpp", u"КПП"), ("acc", u"р/с"), ("bic", u"БИК"), ("coracc", u"Корсчет")]:
            val = vars(args)[a]
            if val != None: our[o] = val
    elif args.reqfile == None:
        our = {
            u"Наименование": u"ООО Бесконтактные устройства",
            u"ИНН": "7702818199",
            u"КПП": "770201001",
            u"р/с": "40702810700120030086",
            u"Банк": u"ОАО АКБ \"АВАНГАРД\"",
            u"БИК": "044525201",
            u"Корсчет": "30101810000000000201",
        }
    else:
        our = {}
        with open(args.reqfile.decode("utf-8"),"r") as cf: data = cf.read()
        data = data.decode("utf-8")
        for line in data.split("\n"):
            line = line.strip()
            if len(line) > 0:
                key, val = line.split(":", 1)
                our[key] = val.strip()

    return (args, our)

drp = re.UNICODE | re.IGNORECASE # default regexp parameters

# Алгоритм работы:
# В первом проходе ищем надписи ИНН, КПП, БИК и переносим числа, следующие за ними в
# соответствующие поля.
# Если не в первом проходе не найдены ИНН, КПП или БИК, тогда запускаем второй проход:
# Первое десятизначное число интерпретируем как ИНН, первое девятизначное с первыми четырьмя 
# цифрами, совпадающими с ИНН - как КПП, первое девятизначное, начинающиеся с 04 -- как БИК
#

class ParseResult(dict):
    def __init__(self, our, filename = None):
        self.filename = filename
        self.our = our
        self.errs = []

class Err:
    def __repr__(self):
        return ""

def isErr(val): return isinstance(val, Err)

def epsilonEquals(a,b):
    if a == None or b == None: return False
    if isErr(a) or isErr(b): return False
    return abs(a-b) < 0.0001

def parse(num):
    num = re.split(r"\t|\n", num.lstrip())[0] # При парсинге Excel подхватываются числа сразу из нескольких колонок
    num = re.sub(ur"руб(лей)?\.?", "", num, drp).strip(u",. \u00a0\n")
    if len(num)==0: return None
    num = re.sub(u"[',\\.\\s\u00a0]([0-9]{3})", r"\1", num) # Удаляем точки, запятые, апострофы и пробелы, отделяющие тысячи
    num = re.sub(r"[,\-]([0-9][0-9]?)$", r".\1", num)       # Заменяем запятую и дефис, отделяющие десятичные знаки, на точку
    try:
       return float(num)
    except ValueError:
       return None

innChk = [3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8]

# Проверка полей
def checkInn(val):
    if not re.match("[0-9]{10}(?:[0-9]{2})?$", val): return False
    # Проверка контрольных цифр
    def innControlDigit(idx):
        if int(val[idx]) != sum(int(i)*c for i,c in zip(val[:idx],innChk[-idx:]))%11%10:
            if verbose: errWrite(u"Неверная контрольная цифра %i в ИНН: %r\n" % (idx+1, val))
            return False
        return True
    if len(val) == 10:
        return innControlDigit(9)
    return innControlDigit(10) and innControlDigit(11)

def checkKpp(val): return re.match("[0-9]{9}$", val) != None
def checkBic(val): return re.match("04[0-9]{7}$", val) != None
def checkAcc(val): return re.match("[0-9]{20}$", val) != None

checkDict = { u"ИНН": checkInn, u"КПП": checkKpp, u"БИК": checkBic, u"р/с": checkAcc, u"Корсчет": checkAcc }

# Заполняет поле с именем fld, проверяет его, и проверяет, чтобы в поле уже не присутствовало другое значение
def fillField(pr, fld, value):
    if value == None: return
    ourVal = pr.our.get(fld)
    if fld == u"ИНН" and value == ourVal: return
    oldVal = pr.get(fld)
    if isErr(oldVal): return
    if value == oldVal: return
    if value == ourVal and oldVal != None: return
    check = checkDict.get(fld)
    if check != None and not check(value):
        if verbose: pr.errs.append(u"Найден некорректный %s: %s" % (fld, value))
        return
    if fld == u"СтавкаНДС" and oldVal == u"БезНДС" and "%" in value:
        pr[u"СтавкаНДС"] = value
        return
    if oldVal != None and value != ourVal and oldVal != ourVal:
        if fld == u"Счет": return
        pr.errs.append(u"Найдено несколько различных %s: %s и %s" % (fld, oldVal, value))
        pr[fld]=Err()
        return
    pr[fld] = value

def fillTotal(pr, val):
    if val == None: return
    if epsilonEquals(val, pr.get(u"ИтогоБезНДС")): return
    if epsilonEquals(val, pr.get(u"ИтогоСНДС")): return
    if pr.get(u"СуммаПрописью", False): return
    oldTotal = pr.get(u"Итого")
    if oldTotal != None:
        if oldTotal == val: return
        if pr.get(u"СтавкаНДС") in [u"БезНДС", u"0%"]:
            pr.errs.append(u'Ошибка: найдено несколько итоговых сумм и слова "Без НДС"')
            pr[u"СтавкаНДС"] = Err()
            pr[u"СуммаНДС"] = Err()
        if isErr(oldTotal): return
        withoutVat = min(val, oldTotal)
        withVat = max(val, oldTotal)
        if abs(withVat/1.18-withoutVat)>0.1:
            if verbose:
                pr.errs.append(u"Найдено несколько различных числовых сумм Итого, полагаемся только на сумму прописью")
            pr[u"Итого"] = Err()
        else:
            fillField(pr, u"ИтогоБезНДС", withoutVat)
            fillField(pr, u"ИтогоСНДС", withVat)
            del pr[u"Итого"]
    else:
        pr[u"Итого"] = val

# Функции по поиску и обработке строк с данными по НДС
def checkWithoutVat(pr, text):
    rr = re.search(ur"(Сумма|Цена|Итого|Всего)?(\s*|[^:;\n\t]*?)Без\s*(налога)?\s*\(?НДС", text, drp)
    if (rr != None and rr.group(1) == None) or u"НДС не облагается" in text:
        fillField(pr, u"СтавкаНДС", u"БезНДС")

def fillVatType(pr, content):
    if re.search(r"\b18(.00)?%", content, drp): fillField(pr, u"СтавкаНДС", u"18%")
    if re.search(r"\b10(.00)?%", content, drp): fillField(pr, u"СтавкаНДС", u"10%")
    if re.search(r"\b[^\.,\d]0%", content, drp):
        fillField(pr, u"СтавкаНДС", u"0%")

vatIntro     = ur"(Всего|Итого|Сумма|в\sт\.\s?ч\.|в\sтом\sчисле|включая)?\s*НДС\s*"
vatPercentRe = "[^\d\n]?(\d\d?(?:.00)?%)?[^\d\n]?"
def checkVatAmount(pr, text, allowNewlines = False):
    for r in re.finditer(vatIntro +                                                            # Вводные слова
             ur"(?:по\sставке)?\s*%s(\s?[^\d\n]?\s*(?:руб)?([^\d\n]*)\s*)" % vatPercentRe +    # Ставка НДС
             ur"(?:([0-9][0-9\.,'\-\s]*)\s*(?:руб(?:лей)?\.?\s*([0-9][0-9]?)\s*коп(?:еек)?\.?)?)?", # Сумма НДС
             text, drp):
        if r.group(1) == None and r.group(2) == None: continue # Если нет вводного слова и ставки -- пропускаем
        if r.group(2) != None: fillVatType(pr, r.group(2)) # group 2: ставка НДС
        if r.group(4) != None and u"руб" in r.group(4):    # group 4: произвольные слова
            continue # Игнорировать сумму НДС прописью, она обрабатывается в другом месте

        if not allowNewlines and u"\n" in r.group(3): continue # group 3: whitespace и произовольные слова

        # Блокируем случайный подхват нескольких строк с итогами в multiline режиме
        if allowNewlines and re.search(u"Итого|Всего|Сумма", r.group(3), drp) != None: continue

        vat = None
        if r.group(5) != None:  # group 5: Сумма НДС
            vat = parse(r.group(5))
        if r.group(6) != None:  # group 6: Копейки в сумме НДС
            vat += float(r.group(6))/100

        if vat != None and (r.group(1) != None or (r.group(2) != None and pr.get(u"СтавкаНДС") == "18%")):
            fillField(pr, u"СуммаНДС", vat)

accChk = [7,1,3]*7+[7,1]
# Проверка ключа банковского счёта по БИКу и номеру счёта
def checkBicAcc(pr, errs = None):
    def showError():
        errs.append(u"Некорректный ключ номера счёта: %s" % pr[u"р/с"])

    if u"Корсчет" in pr:
       prefix = pr[u"БИК"][6:10]
    else:
       prefix = "0" + pr[u"БИК"][4:6]
    fullAcc = prefix + pr[u"р/с"]

    if sum(int(i)*c for i,c in zip(fullAcc, accChk)) % 10 != 0:
        showError()
        return False
    key = int(fullAcc[11])
    fullAcc = fullAcc[:11] + u"0" + fullAcc[12:]
    if sum(int(i)*c for i,c in zip(fullAcc, accChk)) * 3 % 10 != key:
        showError()
        return False
    return True

def extractPdfImage(pr, img):
    filters = img.stream.get_filters()
    if len(filters)>0 and filters[0][0].name == "DCTDecode":
        return Image.open(BytesIO(img.stream.get_rawdata()))
    if img.bits == 8:
        if isinstance(img.colorspace[0], PSLiteral):
            if img.colorspace[0].name == "DeviceRGB":
                return Image.frombuffer("RGB", img.srcsize, img.stream.get_data(), "raw", "RGB", 0, 1)
            if img.colorspace[0].name == "DeviceGray":
                return Image.frombuffer("L", img.srcsize, img.stream.get_data(), "raw", "L", 0, 1)
    if img.bits == 1:
        if isinstance(img.colorspace[0], PSLiteral) and img.colorspace[0].name == "DeviceGray":
            return Image.frombuffer("1", img.srcsize, img.stream.get_data(), "raw", "1", 0, 1)
        if img.colorspace[0] == None:
            return Image.frombuffer("1", img.srcsize, img.stream.get_data(), "raw", "1;I", 0, -1)
    if verbose: pr.errs.append("image with unknown format found, skipping")

# Убрать лишние данные из номера счёта
def stripInvoiceNumber(num):
    m = re.search(ur"\b(от|дата):? *[0-3]?\d(\.[0-1]\d\.| *[а-яА-Я]* )(20)?\d\d( *г([г\.]|ода?)?)?", num, drp)
    if m: return num[0:m.end()]
    return num

def processXlsCell(sht, row, col, pr):
    try:
        content = sht.cell_value(row, col)
        if isinstance(content, float):
            content = u"%.2f" % content
        else:
            content = unicode(content).strip()
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
        elif isinstance(val, float): val = u"%.2f" % val
        elif val != None: val = unicode(val)
        return (val, col)
    processCellContent(content, getValueToTheRight, col, pr)

# Находит ближайший LTTextLine справа от указанного
def pdfFindRight(pdf, pl):
    y = (pl.y0 + pl.y1) / 2
    result = None
    for obj in pdf:
        if not isinstance(obj, LTTextBox) and not isinstance(obj, LTFigure): continue
        if obj.y0 > y or obj.y1 < y: continue
        for line in obj:
            if not isinstance(line, LTTextLine) and not isinstance(line, LTImage): continue
            if line.y0 > y or line.y1 < y or line.x0<=pl.x0: continue
            if result != None and result.x0 <= obj.x0: continue
            result = line
    return result

def processPdfLine(pdf, pl, pr):
    content = pl.get_text().strip()
    def getValueToTheRight(pdfLine):
        obj = pdfFindRight(pdf, pdfLine)
        if obj == None: return (None, None)
        if isinstance(obj, LTTextLine):
            val = obj.get_text()
        elif isinstance(obj, LTImage):
            if pdfLine != pl: return (None, None)
            img = extractPdfImage(pr, obj)
            if img == None: return (None, None)
            val = image_to_string(img, lang="rus+rusnum").decode("utf-8").strip()
        else: raise AssertionError()
        return (val, obj)
    processCellContent(content, getValueToTheRight, pl, pr)

# Шаблон для поиска номера счёта
inv_base = ur"Сч[её]т\s*(на\sоплату|№|N|.*?\bот\s[0-9][0-9]?[\.\s]|.*?\bДата [0-9]{2}.[0-9]{2}.[0-9]{2}[0-9]{2}?"

def processCellContent(content, getValueToTheRight, firstCell, pr):
    def getSecondValue(separator = None):
        # Значение может находиться как в текущей ячейке, так и в следующей
        try:
            val = content.split(separator, 1)[1].strip()
        except IndexError:
            val = ""
        if len(val) > 0: return val
        # В данной ячейке данных не найдено, проверяем ячейки/текстбоксы справа
        return getValueToTheRight(firstCell)[0]

    rm = re.match(ur"ИНН\s*/\s*КПП\b[:\s]*(.*)", content, drp)
    if rm:
        val = rm.group(1).strip()
        if len(val) == 0: val = getValueToTheRight(firstCell)[0]
        if val == None: return
        rm = re.match(r"([0-9]{10})\s*/?\s*([0-9]{9})\b", val, drp)
        if rm == None:
            rm = re.match(r"([0-9]{12})\b", val, drp)
            if rm == None: return
            checkInn(rm.group(1))
            fillField(pr, u"ИНН", rm.group(1))
            return
        fillField(pr, u"ИНН", rm.group(1))
        fillField(pr, u"КПП", rm.group(2))
        return

    for fld in [u"ИНН", u"КПП", u"БИК"]:
        if re.match(u"[^a-zA-Zа-яА-ЯёЁ]?"  + fld + r"\b", content, drp):
            val = getSecondValue()
            if val == None: return
            rm = re.match("[0-9]+", val)
            if not rm: return
            val = rm.group(0)
            if val == pr.our.get(fld): return
            fillField(pr, fld, val)
            return

    rr = re.search(ur"%s|$)" % inv_base, content, drp)
    if rr:
        text = content[rr.start(0):]
        val, cell = getValueToTheRight(firstCell)
        while val != None:
            text += " "
            text += val
            val, cell = getValueToTheRight(cell)
        fillField(pr, u"Счет", stripInvoiceNumber(text.strip().replace("\n"," ").replace("  "," ")))
        return

    if re.search(ur"\sруб", content, drp): findSumsInWords(content, pr)
    if (re.match(u"Итого|Всего|Сумма", content, drp) and
            (re.search(u"(с|без)\s*НДС", content, drp) or not u"НДС" in content)):
        if ":" in content: val = getSecondValue(":")
        else: val = getValueToTheRight(firstCell)[0]
        if val == None: return
        fillTotal(pr, parse(val))
        return

    if u"НДС" in content:
        checkWithoutVat(pr, content)
        fillVatType(pr, content)
        checkVatAmount(pr, content)
        rr = re.search(vatIntro + vatPercentRe, content, drp)
        if rr and (rr.group(1) != None or rr.group(2) != None): # Возможно, сумма НДС находится в другой ячейке
            val = getValueToTheRight(firstCell)[0]
            if val == None: return
            if (re.match(u"Без", val, drp)):
                fillField(pr, u"СтавкаНДС", u"БезНДС")
                return 
            fillField(pr, u"СуммаНДС", parse(val))
            return

    findBankAccounts(content, pr)

def hasIncompleteFields(pr):
    if u"ИтогоСНДС" not in pr and u"Итого" not in pr: return True
    for i in [u"р/с", u"ИНН", u"КПП", u"БИК", u"Счет", u"СуммаНДС", u"СуммаПрописью"]:
        if i not in pr: return True
    return False

def findSumsInWords(text, pr, func = searchSums):
    for psum in func(text):
        if epsilonEquals(psum, pr.get(u"Итого")):
           del pr[u"Итого"]
        elif epsilonEquals(psum, pr.get(u"СуммаНДС")):
            continue

        # Определяем случаи, когда прописью написана суммма НДС
        oldValue = pr.get(u"ИтогоСНДС")
        if oldValue != None and oldValue > 1 and abs(oldValue/1.18*0.18-psum)<0.05:
            fillField(pr, u"СуммаНДС", psum)
            continue

        if not pr.get(u"СуммаПрописью", False) and u"ИтогоСНДС" in pr:
            # Удаляем предыдущие числовые значения т.к. сумма прописью самая надёжная
            del pr[u"ИтогоСНДС"]

        fillField(pr, u"ИтогоСНДС", psum)
        pr[u"СуммаПрописью"] = True

def findBankAccounts(text, pr):
    for w in re.finditer(u"[0-9О]{4} *[0-9О]81[0О] *[0-9О]{4} *[0-9О]{4} *[0-9О]{4}\\b", text, drp):
        w = w.group(0).replace(" ","").replace(u"О", "0") # Многие OCR-движки путают букву О с нулём
        if w[5:8] == "810":
            if w[0] == "4":
                fillField(pr, u"р/с", w)
            if w[0:5] == "30101":
                fillField(pr, u"Корсчет", w)

bndry = u"(?:\\b|[a-zA-Zа-яА-ЯёЁ ])"

def processText(text, pr, allowNewlines = False):
    if not u"р/с" in pr:
        findBankAccounts(text, pr)
    for fld, regexp in [
            [u"ИНН", u"[0-9О]{10}|[0-9О]{12}"],
            [u"КПП", u"[0-9О]{9}"],
            [u"БИК", u"[0О]4[0-9О]{7}"]]:
        for val in re.finditer("\\b" + fld + u"\\n? *(%s)\\b" % regexp, text, drp):
            fillField(pr, fld, val.group(1).replace(u"О", "0"))

    rr = re.search(ur"^\s*%s).*" % inv_base, text, drp | re.MULTILINE)
    if rr: fillField(pr, u"Счет", re.sub(ur"^(Сч[её]т(?: на оплату)?) (?:не|м[9в]) ", ur"\1 № ",
                     stripInvoiceNumber(rr.group(0).strip()), flags=drp))

    # Поиск находящихся рядом пар ИНН/КПП с совпадающими первыми четырьмя цифрами
    if u"ИНН" not in pr and u"КПП" not in pr:
        results = re.findall(ur"[^0-9О]([0-9О]{10}) *[\\\[\]\|/ ] *([0-9О]{9})\b", text, drp)
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
        for rm in re.finditer(u"\\b([0-9О]{10}|[0-9О]{12})\\b" + bndry, text, drp):
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
    if u"СуммаПрописью" not in pr:
        for r in re.finditer(ur"(?:Всего|Итого)(\s?[^0-9\n]?\s*(?:руб)?([^0-9\n]*)\s*)([0-9][0-9'\-\.,\s]*)", text, drp):
            if (allowNewlines or "\n" not in r.group(1)) and u"руб" not in r.group(2) and (
                    re.match(u"(c|без) *НДС",r.group(2), drp) or not u"НДС" in r.group(2)):
                fillTotal(pr, parse(r.group(3).strip(".,")))

    vat = pr.get(u"СуммаНДС")
    if not isinstance(vat, float): checkVatAmount(pr, text, allowNewlines)
    if u"СтавкаНДС" not in pr: checkWithoutVat(pr, text)

    if u"СуммаПрописью" not in pr: findSumsInWords(text, pr, searchSumsFiltered)

    # Если найдена сумма прописью и не найден НДС, ищем по документу цифру,
    # составляющую 118%/18% от суммы прописью
    if u"СуммаПрописью" in pr and (u"СуммаНДС" not in pr or isErr(pr[u"СуммаНДС"])):
        amtvat = pr[u"ИтогоСНДС"]/1.18*0.18
        for nums in re.finditer(ur"\b([0-9'\-\.,\s]*)\b", text, drp):
            for num in re.split(r"\t|     ", nums.group(0)):
                num = parse(num)
                if num != None and abs(amtvat-num)<0.05:
                    pr[u"СуммаНДС"] = num

def processImage(image, pr):
    debug = False
    if image == None: return
    def doProcess():
        text = image_to_string(image, lang="rus+rusnum").decode("utf-8")
        if debug:
            with open("invext-debug.txt","w") as f:
                f.write(text.encode("utf-8"))
            image.save("invext-debug.png", "PNG")
        processText(text, pr)
    # Увеличиваем маленькие изображения
    if hasIncompleteFields(pr) and image.size[0]*image.size[1] < 8000000:
        multiplier = 3
        image = image.resize(tuple([int(i * multiplier) for i in image.size]), Image.BICUBIC)
    doProcess()
    if hasIncompleteFields(pr) and image.mode == "RGB":
        # Убираем синие подписи и печати
        image = ImageOps.autocontrast(image).convert("L", (-0.5,-0.5,2,0))
        doProcess()

def pdfToTextPoppler(pr):
    debug = False
    sp = subprocess.Popen(["pdftotext", "-layout", pr.filename, "-"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    txtdata, stderrdata = sp.communicate()
    if sp.poll() != 0:
        pr.errs.append("Call to pdftotext failed, errcode is %i" % sp.poll())
        pr.errs.append(stderrdata)
        return ""
    return txtdata.decode("utf-8")

def processPDF(f, pr):
    debug = False
    parser = PDFParser(f)
    document = PDFDocument(parser)
    rsrcmgr = PDFResourceManager()
    daggr = PDFPageAggregator(rsrcmgr, laparams=LAParams())
    interpreter = PDFPageInterpreter(rsrcmgr, daggr)
    for page in PDFPage.create_pages(document):
        interpreter.process_page(page)
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
        if hasIncompleteFields(pr):
            # Файл не удалось полностью распознать, ищем картинки, перекрывающие всю страницу
            for obj in layout:
                if isinstance(obj, LTFigure):
                    for img in obj:
                        if (isinstance(img, LTImage) and
                                img.x0<x0 and img.y0<y0 and img.x1>x1 and img.y1>y1):
                            processImage(extractPdfImage(pr, img), pr)
    if hasIncompleteFields(pr):
        processText(pdfToTextPoppler(pr), pr)

def processExcel(pr):
    wbk = xlrd.open_workbook(pr.filename, logfile = devnull)
    for sht in wbk.sheets():
        for row in range(sht.nrows):
            for col in range(sht.ncols):
                processXlsCell(sht, row, col, pr)
    if hasIncompleteFields(pr):
        text = ""
        for sht in wbk.sheets():
            for row in range(sht.nrows):
                for col in range(sht.ncols):
                    val = sht.cell_value(row, col)
                    if isinstance(val, float): text += u"%.2f\t" % val
                    else: text += unicode(sht.cell_value(row, col).strip()) + "\t"
                text += "\n"
        processText(text, pr)

def processMsWord(pr):
    debug = False
    sp = subprocess.Popen(["antiword", "-x", "db", pr.filename], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    docdata, stderrdata = sp.communicate()
    docdata = unescape(re.sub(r"</?[^>]+>", "\n", docdata))
    if sp.poll() != 0:
        pr.errs.append("Call to antiword failed, errcode is %i" % sp.poll())
        pr.errs.append(stderrdata)
        return
    if debug:
        with open("invext-debug.txt","w") as f:
            f.write(docdata)
    processText(docdata.decode("utf-8"), pr, True)

def getBicData(bic, errs):
    url = "http://www.bik-info.ru/bik_%s.html" % bic
    try:
        f = urllib2.urlopen(url)
    except urllib2.URLError:
        errs.append(u"Ошибка: не удалось загрузить страницу %s" % url)
    page = f.read().decode("cp1251")
    f.close()
    try:
        return {
            u"Корсчет": re.search(u"Корреспондентский счет: <b>(.*?)</b>", page).group(1),
            u"Наименование": re.search(u"Наименование банка: <b>(.*?)</b>", page).group(1),
            u"Город": re.search(u"Расположение банка: <b>(.*?)</b>", page).group(1),
        }
    except AttributeError:
        errs.append(u"Ошибка: не удалось распознать страницу %s" % url)

def requestCompanyInfoFedresurs(inn, errs):
    data = urllib.urlencode({
    "ctl00$MainContent$txtCode": inn,
    "ctl00$MainContent$btnSearch": u"Поиск".encode("utf-8"),
    "__VIEWSTATE":"/wEPDwUJNjcyMTM3MDUwD2QWAmYPZBYOZg8UKwACFCsAAw8WAh4XRW5hYmxlQWpheFNraW5SZW5kZXJpbmdoZGRkZGQCAg9kFg4CDw8PFgIeC05hdmlnYXRlVXJsBRxodHRwOi8vYmFua3JvdC5mZWRyZXN1cnMucnUvZGQCEw8PFgIfAQUVfi9IZWxwcy9wcml2YXRlMS5odG1sZGQCFA8PFgIfAQULL2JhY2tvZmZpY2VkZAIaDxYCHglpbm5lcmh0bWwFgAo8c3Bhbj4NCiAgICDQktC60LvRjtGH0LXQvdC40LUg0YHQstC10LTQtdC90LjQuSDQsiDQldC00LjQvdGL0Lkg0YTQtdC00LXRgNCw0LvRjNC90YvQuSAg0YDQtdC10YHRgtGAINGB0LLQtdC00LXQvdC40Lkg0L4g0YTQsNC60YLQsNGFINC00LXRj9GC0LXQu9GM0L3QvtGB0YLQuCDRjtGA0LjQtNC40YfQtdGB0LrQuNGFINC70LjRhiDQvtGB0YPRidC10YHRgtCy0LvRj9C10YLRgdGPINC90LAg0L7RgdC90L7QstCw0L3QuNC4INGB0YLQsNGC0YzQuCA3LjEg0KTQtdC00LXRgNCw0LvRjNC90L7Qs9C+INC30LDQutC+0L3QsCDQvtGCIDgg0LDQstCz0YPRgdGC0LAgMjAwMSDQs9C+0LTQsCDihJYgMTI5LdCk0JcgItCeINCz0L7RgdGD0LTQsNGA0YHRgtCy0LXQvdC90L7QuSDRgNC10LPQuNGB0YLRgNCw0YbQuNC4INGO0YDQuNC00LjRh9C10YHQutC40YUg0LvQuNGGINC4INC40L3QtNC40LLQuNC00YPQsNC70YzQvdGL0YUg0L/RgNC10LTQv9GA0LjQvdC40LzQsNGC0LXQu9C10LkiICjQsiDRgNC10LTQsNC60YbQuNC4INCk0LXQtNC10YDQsNC70YzQvdC+0LPQviDQt9Cw0LrQvtC90LAg0L7RgiAxOCDQuNGO0LvRjyAyMDExINCz0L7QtNCwIOKEliAyMjgt0KTQlyAi0J4g0LLQvdC10YHQtdC90LjQuCDQuNC30LzQtdC90LXQvdC40Lkg0LIg0L7RgtC00LXQu9GM0L3Ri9C1INC30LDQutC+0L3QvtC00LDRgtC10LvRjNC90YvQtSDQsNC60YLRiyDQoNC+0YHRgdC40LnRgdC60L7QuSDQpNC10LTQtdGA0LDRhtC40Lgg0LIg0YfQsNGB0YLQuCDQv9C10YDQtdGB0LzQvtGC0YDQsCDRgdC/0L7RgdC+0LHQvtCyINC30LDRidC40YLRiyDQv9GA0LDQsiDQutGA0LXQtNC40YLQvtGA0L7QsiDQv9GA0Lgg0YPQvNC10L3RjNGI0LXQvdC40Lgg0YPRgdGC0LDQstC90L7Qs9C+INC60LDQv9C40YLQsNC70LAsINC40LfQvNC10L3QtdC90LjRjyDRgtGA0LXQsdC+0LLQsNC90LjQuSDQuiDRhdC+0LfRj9C50YHRgtCy0LXQvdC90YvQvCDQvtCx0YnQtdGB0YLQstCw0Lwg0LIg0YHQu9GD0YfQsNC1INC90LXRgdC+0L7RgtCy0LXRgtGB0YLQstC40Y8g0YPRgdGC0LDQstC90L7Qs9C+INC60LDQv9C40YLQsNC70LAg0YHRgtC+0LjQvNC+0YHRgtC4INGH0LjRgdGC0YvRhSDQsNC60YLQuNCy0L7QsiIpINGBIDEg0Y/QvdCy0LDRgNGPIDIwMTMg0LPQvtC00LAgKNC/0YPQvdC60YIgMiDRgdGC0LDRgtGM0LggNiDQpNC10LTQtdGA0LDQu9GM0L3QvtCz0L4g0LfQsNC60L7QvdCwINC+0YIgMTgg0LjRjtC70Y8gMjAxMSDQs9C+0LTQsCDihJYgMjI4LdCk0JcpLiANCjwvc3Bhbj4NCmQCLA8PFgQfAQUZbWFpbHRvOiBiaGVscEBpbnRlcmZheC5ydR4EVGV4dAURYmhlbHBAaW50ZXJmYXgucnVkZAItDw8WBB8BBSJ+L0hlbHBzL0VGUlNEVUwtRkFRLTIwMTUtMDQtMDYucGRmHwMFM9Ce0YLQstC10YLRiyDQvdCwINGH0LDRgdGC0YvQtSDQstC+0L/RgNC+0YHRiyAoRkFRKWRkAi8PZBYCZg9kFgJmD2QWAgIBD2QWAmYPZBYCAgUPZBYCZg9kFggCAQ8PFgIeB1Zpc2libGVoZBYCZg8PFgIfAwVB0K3RgtC+0LPQviDQvdC1INC00L7Qu9C20L3QviDQsdGL0YLRjCDQt9C00LXRgdGMINC90LDQv9C40YHQsNC90L5kZAIDDxQrAAJkEBYAFgAWAGQCBQ8UKwACDxYEHgtfIURhdGFCb3VuZGceC18hSXRlbUNvdW50AgFkZBYCZg9kFgICAQ9kFgJmDxUIEy9jb21wYW5pZXMvMTEzNzkzNjU20J7QntCeINCR0JXQodCa0J7QndCi0JDQmtCi0J3Qq9CVINCj0KHQotCg0J7QmdCh0KLQktCQDGlubGluZS1ibG9jawo3NzAyODE4MTk5DTExMzc3NDY1NzY2NDgEbm9uZRbQlNC10LnRgdGC0LLRg9GO0YnQtdC1ajEyNzA1MSwg0JzQvtGB0LrQstCwINCzLCDQodGD0YXQsNGA0LXQstGB0LrQuNC5INCcLiDQv9C10YAsIDksINCh0KLQoC4gMSwg0K3QotCQ0JYgMiDQn9Ce0JwuIEkg0JrQntCcLjU20JBkAgcPFCsAAmQQFgAWABYAZAIDDw8WAh8BBSNodHRwOi8vZm9ydW0tZmVkcmVzdXJzLmludGVyZmF4LnJ1L2RkAgQPDxYCHwEFHmh0dHA6Ly90ZXN0LWZhY3RzLmludGVyZmF4LnJ1L2RkAgUPFgIeBGhyZWYFGWh0dHA6Ly93d3cuZWNvbm9teS5nb3YucnVkAgYPFgIfBwUUaHR0cDovL3d3dy5uYWxvZy5ydS9kAgcPFgIfAgWZBTxzcGFuPg0KICAgINCS0LrQu9GO0YfQtdC90LjQtSDRgdCy0LXQtNC10L3QuNC5INCyINCV0LTQuNC90YvQuSDRhNC10LTQtdGA0LDQu9GM0L3Ri9C5ICDRgNC10LXRgdGC0YAg0YHQstC10LTQtdC90LjQuSDQviDRhNCw0LrRgtCw0YUg0LTQtdGP0YLQtdC70YzQvdC+0YHRgtC4INGO0YDQuNC00LjRh9C10YHQutC40YUg0LvQuNGGINC+0YHRg9GJ0LXRgdGC0LLQu9GP0LXRgtGB0Y8g0L3QsCDQvtGB0L3QvtCy0LDQvdC40Lgg0YHRgtCw0YLRjNC4IDcuMSDQpNC10LTQtdGA0LDQu9GM0L3QvtCz0L4g0LfQsNC60L7QvdCwIA0KICAgINC+0YIgOCDQsNCy0LPRg9GB0YLQsCAyMDAxINCz0L7QtNCwIOKEliAxMjkt0KTQlyAi0J4g0LPQvtGB0YPQtNCw0YDRgdGC0LLQtdC90L3QvtC5INGA0LXQs9C40YHRgtGA0LDRhtC40Lgg0Y7RgNC40LTQuNGH0LXRgdC60LjRhSDQu9C40YYg0Lgg0LjQvdC00LjQstC40LTRg9Cw0LvRjNC90YvRhSDQv9GA0LXQtNC/0YDQuNC90LjQvNCw0YLQtdC70LXQuSIg0YEgMSDRj9C90LLQsNGA0Y8gMjAxMyDQs9C+0LTQsCANCiAgICAo0L/Rg9C90LrRgiAyINGB0YLQsNGC0YzQuCA2INCk0LXQtNC10YDQsNC70YzQvdC+0LPQviDQt9Cw0LrQvtC90LAg0L7RgiAxOCDQuNGO0LvRjyAyMDExINCz0L7QtNCwIOKEliAyMjgt0KTQlykuDQo8L3NwYW4+ZBgEBR5fX0NvbnRyb2xzUmVxdWlyZVBvc3RCYWNrS2V5X18WAQUWY3RsMDAkcmFkV2luZG93TWFuYWdlcgUjY3RsMDAkTWFpbkNvbnRlbnQkdWNCb3R0b21EYXRhUGFnZXIPFCsABGRkAhQCAWQFH2N0bDAwJE1haW5Db250ZW50JGx2Q29tcGFueUxpc3QPFCsADmRkZGRkZGQUKwABZAIBZGRkZgIUZAUgY3RsMDAkTWFpbkNvbnRlbnQkdWNUb3BEYXRhUGFnZXIPFCsABGRkAhQCAWQ=",
    "__EVENTVALIDATION":"/wEdAAgIIDiAdZkwuCBwGFg+Yb3wQAkPfS3ALM1l8HYCRLcTjKUF4enLXO3emfMk8iBi1qvRvDs5OXQ11rod7fgapnnyQ2pdoSqiOAqq4PCYWcCsWwd4wD37xIK/Lo7dzZyKenvHGmy602W3dHJKoVjq4UNjn4j8c9nzo0RlxtfBH2PEDg=="
    })
    url = "http://www.fedresurs.ru/companies/IsSearching"
    req = urllib2.Request(url, data)
    fp = None
    try:
        fp = urllib2.urlopen(req)
        orgId = re.search(r"window.location.assign\('/companies/([0-9]+)'\)", fp.read().decode("utf-8")).group(1)
        fp.close(); fp = None
        url = "http://www.fedresurs.ru/companies/" + orgId
        fp = urllib2.urlopen(url)
        page = fp.read().decode("utf-8")
        fp.close(); fp = None
        inn2 = re.search(ur"ИНН:</td>\s*<td>([0-9]{10})</td>", page, re.UNICODE).group(1)
        if inn2 != inn:
            errs.append(u"Ошибка обращения к сайту fedresurs.ru: ИНН не соответствует запрошенному")
            return None
        return {
            u"КПП": re.search(ur"КПП:</td>\s*<td>([0-9]{9})</td>", page, re.UNICODE).group(1),
            u"Наименование": re.search(ur"<td>Сокращённое фирменное наименование:</td>\s*<td>(.*?)</td>", page, re.UNICODE).group(1),
        }
    except urllib2.URLError:
        errs.append(u"Ошибка: не удалось загрузить страницу %s" % url)
    except AttributeError:
        errs.append(u"Ошибка: не удалось распознать страницу %s" % url)
    finally:
        if fp != None: fp.close()
    return None

def requestCompanyNameIgk(inn, errs):
    url = "http://online.igk-group.ru/ru/home?inn=" + inn
    try:
        f = urllib2.urlopen(url)
    except urllib2.URLError:
        errs.append(u"Ошибка: не удалось загрузить страницу %s" % url)
        return None
    page = f.read().decode("utf-8")
    f.close()
    try:
        inn2 = re.search(ur"<th>ИНН</th>\s*<td>([0-9]{10}(?:[0-9]{2})?)</td>", page, re.UNICODE).group(1)
        if inn2 != inn:
            errs.append(u"Ошибка обращения к сайту igk-group.ru: ИНН не соответствует запрошенному")
            return None
        if len(inn) == 12:
            return u"ИП " + re.search(ur"<th>Руководство</th>\s*<td>\s*(.*?)\s*</td>", page, re.UNICODE).group(1)
        elif len(inn) == 10:
            return re.search(ur"<th>Краткое название</th>\s*<td colspan=\"3\">\s*(.*?)\s*</td>", page, re.UNICODE).group(1)
        errs.append(u"Неверная длина ИНН: " + inn)
    except AttributeError:
        errs.append(u"Ошибка: не удалось распознать страницу %s" % url)
    return None

def finalizeAndCheck(pr):
    if isinstance(pr.get(u"СтавкаНДС"), Err): del pr[u"СуммаНДС"]
    for fld in pr.keys():
        if isinstance(pr[fld], Err): del pr[fld]
    if u"ИтогоСНДС" not in pr and u"Итого" in pr: pr[u"ИтогоСНДС"] = pr[u"Итого"]
    def deleteBank():
        for fld in [u"БИК", u"Корсчет", u"р/c"]:
            if fld in pr: del pr[fld]
    if u"р/с" in pr and "БИК" in pr and not checkBicAcc(pr, pr.errs):
        deleteBank()
    if u"БИК" in pr:
        bicData = getBicData(pr[u"БИК"], pr.errs)
        if bicData:
            if bicData[u"Корсчет"] != pr.get(u"Корсчет", ""):
                if u"Корсчет" in pr and pr[u"Корсчет"] == pr.our.get(u"Корсчет"):
                    # Распознался наш корсчёт, настоящий корсчёт не распознался, либо его нет
                    if len(bicData[u"Корсчет"]) == 0:
                        del pr[u"Корсчет"] # На самом деле корсчёта быть не должно
                    else:
                        # На самом деле корсчёт другой, но он не распознался, возможно из-за OCR
                        if verbose: pr.errs.append(u"Настоящий корсчёт не распознался: в файле %s, в базе %s" % (
                            pr.get(u"Корсчет", u"пусто"), u"пусто" if len(bicData[u"Корсчет"]) == 0 else bicData[u"Корсчет"]))
                        pr[u"Корсчет"] = bicData[u"Корсчет"]
                else:
                    pr.errs.append(u"Ошибка: не совпадает корсчёт: в файле %s, в интернет-базе %s. Оставляю из файла" % (
                        pr.get(u"Корсчет", u"пусто"), u"пусто" if len(bicData[u"Корсчет"]) == 0 else bicData[u"Корсчет"]))
                    if not strict and u"Корсчет" not in pr:
                        pr[u"Корсчет"] = bicData[u"Корсчет"]
                    else:
                        deleteBank()
            if u"БИК" in pr:
                pr[u"Банк"] = bicData[u"Наименование"]
                pr[u"Банк2"] = u"г. " + bicData[u"Город"]
        else:
            pr.errs.append(u"Ошибка: не удалось получить данные по БИК %s" % pr[u"БИК"])
            if strict: deleteBank()

    if u"ИНН" in pr:
        ci = None
        if len(pr[u"ИНН"]) == 10:
            ci = requestCompanyInfoFedresurs(pr[u"ИНН"], pr.errs)
            if ci == None:
                # Иногда fedresurs с первого раза не отвечает, пробуем снова через 1 секунду
                time.sleep(1)
                ci = requestCompanyInfoFedresurs(pr[u"ИНН"], pr.errs)
        if ci != None:
            if ci[u"КПП"] != pr.get(u"КПП", u""):
                pr.errs.append(u"Не совпадает КПП для %s: в файле %s, в интернет-базе %s. Оставляю из файла" % (
                        ci[u"Наименование"], pr.get(u"КПП", u"пусто"), ci[u"КПП"]))
                if strict: del pr[u"КПП"]
                elif u"КПП" not in pr: pr[u"КПП"] = ci[u"КПП"]
            pr[u"Наименование"] = ci[u"Наименование"]
        else:
            pr[u"Наименование"] = requestCompanyNameIgk(pr[u"ИНН"], pr.errs)

    if not pr.get(u"СуммаПрописью"):
        pr.errs.append(u"Предупреждение: сумма прописью не найдена")

    # Проверяем, чтобы сумма НДС не была слишком большой (это значит, что некорректно подхватилось другое число)
    amt = pr.get(u"ИтогоСНДС")
    vat = pr.get(u"СуммаНДС")
    if vat != None and amt != None:
        if vat>(amt*0.18+0.1):
            pr.errs.append(u"Ошибка: некорректная сумма НДС: %r" % vat)
            del pr[u"СуммаНДС"]

    # Автоматическое определение ставки НДС если явно не указано в документе
    vatMatch = amt > 1 and vat and abs(amt/1.18*0.18 - vat)<0.05
    if u"СтавкаНДС" not in pr and vatMatch:
        pr[u"СтавкаНДС"] = "18%"

    # При сумме больше 100 тысяч выводим сумму только если есть сумма прописью либо соответствующий НДС
    if amt > 100000 and not vatMatch and not pr.get(u"СуммаПрописью", False):
        del pr[u"ИтогоСНДС"]

    # Генерируем назначение платежа
    paydetails = pr.get(u"Счет", u"Номер счета неизвестен")

    if u"ИтогоСНДС" in pr:
        paydetails += u" Сумма %.2f" % pr[u"ИтогоСНДС"]
    vatRate = pr.get(u"СтавкаНДС")
    vat = pr.get(u"СуммаНДС")
    if vatRate == u"БезНДС":
        paydetails += u", НДС не облагается"
    elif vat != None:
        paydetails += u", в т.ч. НДС"
        if vatRate != None: paydetails += u" (%s)" % vatRate
        paydetails += u" - %.2f" % pr.get(u"СуммаНДС")
    pr[u"НазначениеПлатежа"] = paydetails

def printMainInvoiceData(pr, fout):
    if u"Счет" in pr: fout.write((pr[u"Счет"] + "\n").encode("utf-8"))
    else: fout.write(u"Номер счёта не найден\n".encode("utf-8"))
    for fld in [u"ИНН", u"р/с", u"БИК", u"НазначениеПлатежа"]:
        fout.write((u"%s: %s\n" % (fld, pr.get(fld, u"не найдено"))).encode("utf-8"))

def processFile(our, f):
    pr = ParseResult(our,f)
    ext = os.path.splitext(f)[1].lower()
    if (ext in ['.png','.bmp','.jpg','.jpeg','.gif','.tif','.tiff','.ppm']):
        processImage(Image.open(f), pr)
    elif (ext == '.pdf'):
        with open(f, "rb") as ff: processPDF(ff, pr)
    elif (ext in ['.xls', '.xlsx']):
        processExcel(pr)
    elif (ext in ['.doc']):
        processMsWord(pr)
    elif (ext in ['.txt', '.xml']):
        with open(f, "rb") as ff: processText(ff.read().decode("utf-8"), pr)
    else:
        pr.errs.append("unknown extension")
    return pr

def checkOur(our, errs):
    result = True
    if len(our[u"ИНН"]) == 12: del our[u"КПП"]
    for fld, check in checkDict.iteritems():
        if our.get(fld, "") != "" and not check(our[fld]):
            errs.append(u"Ошибка: задан некорректный %s нашей организации: %s" % (fld, our[fld]))
            result = False

    if u"БИК" in our and u"р/с" in our and u"Корсчет" in our:
        result &= checkBicAcc(our, errs)

    return result

oneCHeader = u"""1CClientBankExchange
ВерсияФормата=1.02
Кодировка=Windows
ДатаСоздания={0}
ВремяСоздания={1}
ДатаНачала={0}
ДатаКонца={0}
РасчСчет={2}
"""
oneCItemTemplate = (
u"""СекцияДокумент=Платежное поручение
Дата={Дата}
Сумма={ИтогоСНДС}
ПлательщикСчет={our:р/с}
ПлательщикРасчСчет={our:р/с}
Плательщик={our:Наименование}
Плательщик1={our:Наименование}
ПлательщикИНН={our:ИНН}
ПлательщикКПП={our:КПП}
ПлательщикБанк1={our:Банк}
ПлательщикБанк2={our:Банк2}
ПлательщикБИК={our:БИК}
ПлательщикКорсчет={our:Корсчет}
ПолучательСчет={р/с}
ПолучательРасчСчет={р/с}
Получатель={Наименование}
Получатель1={Наименование}
ПолучательИНН={ИНН}
ПолучательКПП={КПП}
ПолучательБанк1={Банк}
ПолучательБанк2={Банк2}
ПолучательБИК={БИК}
ПолучательКорсчет={Корсчет}
ВидОплаты=01
Очередность=5
НазначениеПлатежа={НазначениеПлатежа}
КонецДокумента
""")

class OneCOutput:
    def __init__(self, filename, our):
        dateStr = datetime.date.today().strftime("%d.%m.%Y")
        self.fout = open(filename, "w")
        self.fout.write(oneCHeader.format(dateStr,
            datetime.datetime.now().strftime("%H:%M:%S"),
            our.get(u"р/с", "")).encode("cp1251"))
        self.itemTemplate = oneCItemTemplate.replace(u"{Дата}", dateStr)
        for fld, val in our.iteritems():
            self.itemTemplate = self.itemTemplate.replace(u"{our:%s}" % fld, val)
        self.itemTemplate = re.sub(r"\{our:.*?\}","", self.itemTemplate)

    def writeDocument(self, pr):
        item = self.itemTemplate
        for fld in [u"ИНН", u"КПП", u"Наименование", u"р/с", u"Банк", u"Банк2", u"БИК", u"Корсчет", u"ИтогоСНДС", u"НазначениеПлатежа"]:
            val = pr.get(fld, u"")
            if isinstance(val, float): val = "%.2f" % val
            item = item.replace(u"{%s}" % fld, val)
        item = re.sub(r"\{.*?\}","", item)
        self.fout.write(item.encode("cp1251"))

    def close(self):
        self.fout.write(u"КонецФайла\n".encode("cp1251"))
        self.fout.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()
        return False

if __name__ == '__main__':
    args, our = parseArguments()

    verbose = args.verbose
    strict  = args.strict

    errs = []
    if not checkOur(our, errs):
        for err in errs: sys.stderr.write(err + "\n")
        exit(1)

    oneC = OneCOutput("1c_to_kl.txt", our)
    try:
        for f in args.files:
            print(f)
            pr = processFile(our, f.decode("utf-8"))
            if len(pr) == 0:
                print(u"Не распознано")
                for err in pr.errs: print(err)
            else:
                try:
                    finalizeAndCheck(pr)
                    printMainInvoiceData(pr, sys.stdout)
                finally:
                    for err in pr.errs: print(err)
                oneC.writeDocument(pr)
    finally:
        oneC.close()
