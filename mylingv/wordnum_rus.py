# -*- coding: utf-8
from pyparsing import *

def makeList(lst):
    return Or([CaselessKeyword(v) for v in lst])

dictUnitsCommon = {
    u"три"   :3,
    u"четыре":4,
    u"пять"  :5,
    u"шесть" :6,
    u"семь"  :7,
    u"восемь":8,
    u"девять":9,
}

dictUnitsMasc = {
    u"один":1,
    u"два" :2,
    }


dictUnitsFem = {
    u"одна":1,
    u"две" :2,
    }

unitsCommon = makeList(dictUnitsCommon)
unitsMasc = makeList(dictUnitsMasc) ^ unitsCommon
unitsFem = makeList(dictUnitsFem) ^ unitsCommon

dictTeens = {
    u"десять"      :10,
    u"одиннадцать" :11,
    u"двенадцать"  :12,
    u"тринадцать"  :13,
    u"четырнадцать":14,
    u"пятнадцать"  :15,
    u"шестнадцать" :16,
    u"семнадцать"  :17,
    u"восемнадцать":18,
    u"девятнадцать":19,
    }

teens = makeList(dictTeens)

dictTens = {
    u"двадцать"   :20,
    u"тридцать"   :30,
    u"сорок"      :40,
    u"пятьдесят"  :50,
    u"шестьдесят" :60,
    u"семьдесят"  :70,
    u"восемьдесят":80,
    u"девяносто"  :90,
    }

tens = makeList(dictTens)

dictHundreds = {
    u"сто"      :100,
    u"двести"   :200,
    u"триста"   :300,
    u"четыреста":400,
    u"пятьсот"  :500,
    u"шестьсот" :600,
    u"семьсот"  :700,
    u"восемьсот":800,
    u"девятьсот":900,
    }

hundreds = makeList(dictHundreds)

arrThousands = [u"тысяча", u"тысячи", u"тысяч"]
arrMillions = [u"миллион", u"миллиона", u"миллионов"]

thousands = makeList(arrThousands)
millions = makeList(arrMillions)

zero = CaselessKeyword(u"ноль")

lessHundredMasc = (tens + Optional(unitsMasc)) ^ teens ^ unitsMasc
lessHundredFem = (tens + Optional(unitsFem)) ^ teens ^ unitsFem

lessThousandMasc = Group((hundreds + Optional(lessHundredMasc)) ^ lessHundredMasc)
lessThousandFem = Group((hundreds + Optional(lessHundredFem)) ^ lessHundredFem)

lessMillion = Group((lessThousandFem + thousands + Optional(lessThousandMasc)) ^ lessThousandMasc)

number = (lessThousandMasc + millions + Optional(lessMillion)) ^ lessMillion ^ zero

totalDict = {}
reverseDict = {0: u"ноль"}
for d in [dictHundreds, dictTens, dictTeens, dictUnitsCommon, dictUnitsMasc, dictUnitsFem]:
    totalDict.update(d)

for d in [dictUnitsCommon, dictUnitsMasc, dictTeens]:
    for key, val in d.iteritems():
        reverseDict[val] = key

def chkNum(val, num, errs, arr):
    if not num in arr: return
    val = val.lower()
    if val in [u"один", u"одна"]:
        if num != arr[0]: errs.append(num)
    elif val in [u"два", u"две", u"три", u"четыре"]:
        if num != arr[1]: errs.append(num)
    else:
        if num != arr[2]: errs.append(num)

def parseLessThousand(num):
    result = 0
    for subnum in num:
        result += totalDict[subnum]
    return result

def parseThousands(num, errs):
    if len(num)==1:
        return parseLessThousand(num[0])
    else:
        chkNum(num[0][-1], num[1], errs, arrThousands)
        result = parseLessThousand(num[0]) * 1000
        if len(num)>2: result += parseLessThousand(num[2])
        return result

def parseMillions(num, errs):
    if len(num)==1:
        return parseThousands(num[0], errs)
    else:
        chkNum(num[0][-1], num[1], errs, arrMillions)
        result = parseLessThousand(num[0]) * 1000000
        if len(num)>2: result += parseThousands(num[2], errs)
        return result

def printErrors(errs):
    for err in errs: print(u"Слово \"%s\" стоит в неверном падеже" % err)

def parseNumber(s):
    errs = []
    result = parseMillions(number.parseString(s), errs)
    printErrors(errs)
    return result

arrRub = [u"рубль", u"рубля", u"рублей"]
arrKop = [u"копейка", u"копейки", u"копеек"]

rub = Or([CaselessLiteral(v) for v in arrRub + [u"руб"]])
kop = Or([CaselessLiteral(v) for v in arrRub + [u"коп"]])

sumParse = Group(number) + rub + Optional((Word(srange("[0-9]"), None, 1,2) | Group(lessHundredFem) | zero) + kop)

def parseRubKop(pr):
    errs = []
    result = parseMillions(pr[0], errs)
    lastword = pr[0][-1]
    while isinstance(lastword, ParseResults):
        lastword = lastword[-1]
    chkNum(lastword, pr[1], errs, arrRub)
    if len(pr) > 3:
        if isinstance(pr[2], ParseResults):
            result += float(parseLessThousand(pr[2]))/100
            chkNum(pr[2][-1], pr[3], errs, arrKop)
        elif pr[2] == u"ноль":
            chkNum(pr[2], pr[3], errs, arrKop)
        else:
            kop = int(pr[2])
            result += float(kop)/100
            chkNum(reverseDict[kop % 10 if kop > 19 else kop], pr[3], errs, arrKop)
    printErrors(errs)
    return result

def parseSum(s):
    return parseRubKop(sumParse.parseString(s))

def searchSums(text):
    for pr, start, end in sumParse.scanString(text):
        yield parseRubKop(pr)

def test(expected, s):
    try:
        val = parseSum(s)
    except ParseException, pe:
        print("Parsing failed:")
        print(pe.line)
        print("%s^" % (' '*(pe.col-1)))
        print(pe.msg)
    else:
        print "'%s' -> %r" % (s, val),
        if val == expected:
            print("OK")
        else:
            print("WRONG, expected %r" % expected)

if __name__ == '__main__':
    test(1218660.34, u"Один миллион двести восемнадцать тысяч шестьсот шестьдесят рублей 34 копейки")
    test(1017339.36, u"Один миллион семнадцать тысяч триста тридцать девять рублей 36 копеек")
    test(34130.32, u"Тридцать четыре тысячи сто тридцать рублей 32 копейки")
    test(354000.00, u"Триста пятьдесят четыре тысячи рублей 00 копеек")
    test(6652474.20, u"Шесть миллионов шестьсот пятьдесят две тысячи четыреста семьдесят четыре рубля 20 копеек")
    test(731635.99, u"Семьсот тридцать одна тысяча шестьсот тридцать пять рублей 99 копеек")
    test(194571.38, u"Сто девяносто четыре тысячи пятьсот семьдесят один рубль 38 копеек")
    test(2165851.00, u"Два миллиона сто шестьдесят пять тысяч восемьсот пятьдесят один рубль 00 копеек")
    test(1427328.00, u"Один миллион четыреста двадцать семь тысяч триста двадцать восемь рублей 00 копеек")
    test(696200.00, u"Шестьсот девяносто шесть тысяч двести рублей 00 копеек")
    test(96200.00, u"Девяносто шесть тысяч двести рублей")
    test(7880.00, u"семь тысяч восемьсот восемьдесят рублей 00 копеек")
    test(375.80, u"триста семьдесят пять рублей 80 копеек")
    test(1427311.01, u"Один миллион четыреста двадцать семь тысяч триста одиннадцать рублей 01 копейка")
    test(75.80, u"семьдесят пять рублей 80 копеек")
    test(32000.21, u"Тридцать две тысячи рублей 21 копейка")
    test(32000000.61, u"Тридцать два миллиона рублей 61 копейка")
    test(32001000.81, u"Тридцать два миллиона одна тысяча рублей 81 копейка")
    test(1.32, u"Один рубль тридцать две копейки")
    test(2.00, u"Два рубля ноль копеек")
    test(5.02, u"Пять рублей две копейки")
    test(0.11, u"Ноль рублей одиннадцать копеек")
