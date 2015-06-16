#!/usr/bin/python2
# -*- coding: utf-8
import pyparsing as pp

pp.ParserElement.setDefaultWhitespaceChars(u" .,\t\n\r\u00a0")

def makeList(lst):
    if isinstance(lst, dict):
        return pp.Or([pp.CaselessKeyword(v).setParseAction(pp.replaceWith(lst[v])) for v in lst])
    else:
        return pp.Or([pp.CaselessKeyword(v) for v in lst])

dictUnits = {
    u"один"  :1,
    u"одна"  :1,
    u"два"   :2,
    u"две"   :2,
    u"три"   :3,
    u"четыре":4,
    u"пять"  :5,
    u"шесть" :6,
    u"семь"  :7,
    u"восемь":8,
    u"девять":9,
}

units = makeList(dictUnits)

dictTeens = {
    u"десять"       :10,
    u"одиннадцать"  :11,
    u"двенадцать"   :12,
    u"тринадцать"   :13,
    u"четырнадцать" :14,
    u"пятнадцать"   :15,
    u"шестнадцать"  :16,
    u"семнадцать"   :17,
    u"восемнадцать" :18,
    u"девятнадцать" :19,
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

thousands = makeList([u"тысяча", u"тысячи", u"тысяч"])
millions  = makeList([u"миллион", u"миллиона", u"миллионов"])


asum = lambda arr: sum(arr)
ath  = lambda arr: arr[0] * 1000
amil = lambda arr: arr[0] * 1000000

zero = pp.CaselessKeyword(u"ноль").setParseAction(pp.replaceWith(0))

lessHundred = ((tens + pp.Optional(units)) ^ teens ^ units).setParseAction(asum)

lessThousand = (hundreds + pp.Optional(lessHundred)).setParseAction(asum) ^ lessHundred

lessMillion = ((lessThousand + thousands).setParseAction(ath ) + pp.Optional(lessThousand)).setParseAction(asum) ^ lessThousand

number =      ((lessThousand + millions ).setParseAction(amil) + pp.Optional(lessMillion )).setParseAction(asum) ^ lessMillion ^ zero

rub = makeList([u"рубль", u"рубля", u"рублей", u"руб"])
kop = makeList([u"копейка", u"копейки", u"копеек", u"коп"])

sumParse = number + rub + (pp.Word(pp.srange("[0-9]"), None, 1,2) | lessHundred | zero) + kop

sumParse.setParseAction(lambda arr: arr[0] + float(arr[2])/100)

def parseSum(s):
    return sumParse.parseString(s)[0]

def searchSums(text):
    for pr, start, end in sumParse.scanString(text):
        yield pr[0]

def test(expected, s):
    try:
        val = parseSum(s)
    except pp.ParseException, pe:
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
    test(4658, u"Четыре тысячи шестьсот пятьдесят восемь  руб. 00 коп.")
    test(8460, u"Восемь тысяч четыреста шестьдесят рублей 00 копеек")
    test(5000, u"Пять тысяч рублей ноль копеек.")
