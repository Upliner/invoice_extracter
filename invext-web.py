#!/usr/bin/python2
# -*- coding: utf-8

import os, sys, datetime, traceback
import invoice_extracter as ie

def finish(our, pr, errs, outfile):
    print(datetime.date.today().strftime("%d.%m.%Y")) #1
    print(our.get(u"ИНН", "")) #2
    print(our.get(u"КПП", "")) #3
    print(our.get(u"Наименование", "")) #4
    print(pr.get(u"ИтогоСНДС", "")) #5
    print(our.get(u"р/с", "")) #6
    print((our.get(u"Банк", "") + ' ' + our.get(u"Банк2","")).strip()) #7
    print(our.get(u"БИК", "")) #8
    print(our.get(u"Корсчет", "")) #9
    print((pr.get(u"Банк", "") + ' ' + pr.get(u"Банк2","")).strip()) #10
    print(pr.get(u"БИК", "")) #11
    print(pr.get(u"Корсчет", "")) #12
    print(pr.get(u"ИНН", "")) #13
    print(pr.get(u"КПП", "")) #14
    print(pr.get(u"Наименование", "")) #15
    print(pr.get(u"р/с", "")) #16
    print(pr.get(u"НазначениеПлатежа", "")) #17
    print(outfile) #18
    for err in errs:
        print(err)
    sys.exit(0)
errs = []
our = {}
outfile = ""
try:
    our = {
        u'ИНН': sys.argv[1],
        u'КПП': sys.argv[2],
        u'БИК': sys.argv[3],
        u'Корсчет': sys.argv[4],
        u'р/с': sys.argv[5],
    }
    infile = sys.argv[6].decode("utf-8")
    if not ie.checkOur(our, errs):
        finish(our, {}, errs, "")
    ci = ie.requestCompanyInfoFedresurs(our[u"ИНН"], errs)
    if ci != None:
        our[u"Наименование"] = ci[u"Наименование"]
    else:
        name = ie.requestCompanyNameIgk(our[u"ИНН"], errs)
        if name != None:
            our[u"Наименование"] = name
            errs = []
    if not os.path.isdir("1C"):
        os.mkdir("1C")
    bicData = ie.getBicData(our[u"БИК"], errs)
    if ci != None:
        our[u"Банк"] = bicData[u"Наименование"]
        our[u"Банк2"] = u"г. " + bicData[u"Город"]

    pr = ie.processFile(our, infile)
    errs += pr.errs
    pr.errs = []
    ie.finalizeAndCheck(pr)
    errs += pr.errs
    outfile = os.path.abspath("1C/" + os.path.basename(infile) + ".txt")
    with ie.OneCOutput(outfile, our) as oneC:
        oneC.writeDocument(pr)
    finish(our, pr, errs, outfile)
except SystemExit: pass
except:
    errs.append("Python exception:")
    errs += [ i.rstrip('\n') for i in traceback.format_exception(*sys.exc_info()) ]
    finish(our, {}, errs, outfile)
