#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import sys
import time
sys.path.append(os.path.dirname(__file__))
os.environ['DJANGO_SETTINGS_MODULE'] = 'settings'

import zipfile
from StringIO import StringIO
import xlrd

import urllib2
from BeautifulSoup import BeautifulSoup
from fcimanager.models import Stock, Metric
from datetime import date, datetime

# the parent directory of the project

CNV_BASEURL = "http://www.cnv.gob.ar/"
CNV_REPORTSURL = "InfoFinan/Fondos/Zips.asp?Lang=0&CodiSoc=70086&DescriSoc=Camara%20Argentina%20de%20Fondos%20Comunes%20de%20Inversion&Letra=&Tipoarchivo=1&TipoDocum=26&descripcion=Valores%20Diarios%20de%20Cuotaparte%20%28Valores%20Sujetos%20a%20Revision%29&description=Valores%20Diarios%20de%20Cuotaparte%20%28Valores%20Sujetos%20a%20Revision%29"
CNV_GETZIPREPORT = "Infofinan/fondos/BLOB_Zip.asp?cod_doc="



def CNV_Opener(URL):
    opener = urllib2.build_opener()
    opener.addheaders = [(
        'User-agent',
        'Mozilla/5.0 (Windows NT 5.1; rv:16.0) Gecko/20100101 Firefox/16.0'
        )]
    fp = urllib2.urlopen(CNV_BASEURL + URL)
    page = fp.read()
    fp.close()
    return page


def CNV_FindTable(page):
    soup = BeautifulSoup(page)
    table = soup.find("table", "tablaMenuYContenido")  # Busca tabla clase Text
    return table


def CNV_FindLatestReport(table):
    cells = []
    rowindex = 0

    for row in table.findAll("tr", "text"):  # Busca Registros con Clase Text
        cells.append("")
        cells[rowindex] = [x.text for x in row.findAll("td")]
        rowindex += 1

    Fecha_struct = time.strptime(cells[0][1], "%d %b %Y %H:%M")
    Fecha = datetime.fromtimestamp(time.mktime(Fecha_struct))

    return (Fecha, cells[0][3])


def CNV_GetReports(table):
    cells = []
    rowindex = 0

    for row in table.findAll("tr", "text"):  # Busca Registros con Clase Text
        cells.append("")

        cells[rowindex] = [x.text for x in row.findAll("td")]
        del cells[rowindex][2]  # Delete unwanted fields
        del cells[rowindex][0]  # .
        cells[rowindex][0] = datetime.fromtimestamp(
            time.mktime(
                time.strptime(cells[rowindex][0], "%d %b %Y %H:%M")
                )
            )  # Convert First Field from string to a Datetime object
        rowindex += 1
    return cells


def CNV_GetXLSReport(CNV_Report):
#    URL = CNV_GETZIPREPORT+CNV_Report[1].split("-",)[1]+"&error_page=Error.asp"

    ZipReport = CNV_Opener(CNV_GETZIPREPORT + CNV_Report[1].split("-",)[1])
    fp = StringIO(ZipReport)
    zfp = zipfile.ZipFile(fp, "r")
    xls = zfp.read(zfp.namelist()[0])

    return xls


def CNV_GetStockValueByName(StockName, CNV_Report):
    XLS = CNV_GetXLSReport(CNV_Report)
    wb = xlrd.open_workbook(file_contents=XLS)
    sheet = wb.sheet_by_index(0)

    for row in range(sheet.nrows):
        if (sheet.cell(row, 5).ctype == xlrd.XL_CELL_NUMBER):
            if ((sheet.cell(row, 0).value.strip() == StockName) or
                (sheet.cell(row, 0).value.strip() == StockName + " - Clase A")
               ):
                return sheet.cell(row, 5). value / 1000
    return False

if __name__ == "__main__":
    print "------------------------------------ %s" % str(datetime.now())
    page = CNV_Opener(CNV_REPORTSURL)
    table = CNV_FindTable(page)
    CNV_Report = CNV_FindLatestReport(table)

    #debug -> show LastTrade Date and ReportName from CNV
    stocks = Stock.live.all()
    for stk in stocks:
        print "%s:" % stk.name

        # check if the last metric is equal to the bloomberg last trade
        if Metric.objects.filter(stock=stk).values('date_taken')[:1]:
            lastrec = Metric.objects.filter(
                stock=stk
                ).values('date_taken')[:1][0].values()[0]
        else:
            lastrec = date(1900, 1, 1)

        # debug -> show last DB date
        print "- %s: last metric on DB" % str(lastrec)
        lasttr = CNV_Report[0].date()

        ## debug -> show last trade date
        print "- %s: last trade date" % str(lasttr)
        valor = CNV_GetStockValueByName(stk.name, CNV_Report)

        if not valor:
            print "- !warning! - looks like %s is gone" % (stk.name)
            continue

        if (lastrec != lasttr or lastrec > lasttr) and valor:
            print "- %.6f: updated value" % (valor)
            m = Metric(stock = stk, value = "%.6f" % valor,
                 date_taken = lasttr)
            m.save()

        else:
            print "- value already on DB"


    time.sleep(2)

