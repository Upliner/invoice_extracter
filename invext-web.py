#!/usr/bin/python2
# -*- coding: utf-8

import os, sys, datetime, traceback
import invoice_extracter as ie

lineNum = 1
dbgFile = None

def safeprint(s):
    global lineNum
    if isinstance(s, float): s = "%.2f\n" % s
    else: s = unicode(s).replace("\n"," ").encode("utf-8") + "\n"
    sys.stdout.write(s)
    if dbgFile != None: dbgFile.write(s)
    lineNum += 1

def finish(our, pr, errs, outfile):
    safeprint(datetime.date.today().strftime("%d.%m.%Y")) #1
    safeprint(our.get(u"ИНН", "")) #2
    safeprint(our.get(u"КПП", "")) #3
    safeprint(our.get(u"Наименование", "")) #4
    safeprint(pr.get(u"ИтогоСНДС", "")) #5
    safeprint(our.get(u"р/с", "")) #6
    safeprint((our.get(u"Банк", "") + ' ' + our.get(u"Банк2","")).strip()) #7
    safeprint(our.get(u"БИК", "")) #8
    safeprint(our.get(u"Корсчет", "")) #9
    safeprint((pr.get(u"Банк", "") + ' ' + pr.get(u"Банк2","")).strip()) #10
    safeprint(pr.get(u"БИК", "")) #11
    safeprint(pr.get(u"Корсчет", "")) #12
    safeprint(pr.get(u"ИНН", "")) #13
    safeprint(pr.get(u"КПП", "")) #14
    safeprint(pr.get(u"Наименование", "")) #15
    safeprint(pr.get(u"р/с", "")) #16
    safeprint(pr.get(u"НазначениеПлатежа", "")) #17
    safeprint(outfile) #18
    assert(lineNum == 19)
    for err in errs:
        safeprint(err)
    sys.exit(0)
errs = []
our = {}
outfile = ""
try:
    # Берём данные из коммандной строки
    our = {
        u'ИНН': sys.argv[1],
        u'КПП': sys.argv[2],
        u'БИК': sys.argv[3],
        u'Корсчет': sys.argv[4],
        u'р/с': sys.argv[5],
    }
    infile = sys.argv[6].decode("utf-8")

    # Проверяем корректность введённых данных
    if not ie.checkOur(our, errs):
        finish(our, {}, errs, "")

    # Создаём директории "1C" и "debug" в директории с входящим файлом
    odir = os.path.join(os.path.dirname(infile), "1C")
    if not os.path.isdir(odir):
        os.mkdir(odir)
    dbgdir = os.path.join(os.path.dirname(infile), "debug")
    if not os.path.isdir(dbgdir):
        os.mkdir(dbgdir)

    # Выводим отладочные данные
    dbgFile = open(os.path.join(dbgdir, "debug.txt"), "a")
    dbgFile.write(" ".join(sys.argv) + "\n")

    # Запрашиваем данные по введённому ИНН в online-базах
    ci = ie.requestCompanyInfoFedresurs(our[u"ИНН"], errs)
    if ci != None:
        our[u"Наименование"] = ci[u"Наименование"]
    else:
        name = ie.requestCompanyNameIgk(our[u"ИНН"], errs)
        if name != None:
            our[u"Наименование"] = name
            errs = []
    bicData = ie.getBicData(our[u"БИК"], errs)
    if bicData != None:
        our[u"Банк"] = bicData[u"Наименование"]
        our[u"Банк2"] = u"г. " + bicData[u"Город"]

    # Обрабатываем файл со счётом
    pr = ie.processFile(our, infile.encode("utf-8"))
    errs += pr.errs
    pr.errs = []
    ie.finalizeAndCheck(pr)
    errs += pr.errs

    # Сохраняем результаты обработки в файл и выводим результаты в stdout
    outfile = os.path.abspath(os.path.join(odir, "1c_to_kl.txt"))
    with ie.OneCOutput(outfile, our) as oneC:
        oneC.writeDocument(pr)
    finish(our, pr, errs, outfile)
except SystemExit: pass
except:
    errs.append("Python exception:")
    errs += [ i.rstrip('\n') for i in traceback.format_exception(*sys.exc_info()) ]
    if lineNum == 1:
        finish(our, {}, errs, outfile)
    else:
        if lineNum < 18: sys.stdout.write("\n"*(19-lineNum))
        for err in errs:
            safeprint(err)
finally:
    if dbgFile != None: dbgFile.close()
