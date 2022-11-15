import os
import sys
import webbrowser
from tkinter import StringVar, ttk
import tkinter as tk
from tkinter import *
import tkinter.messagebox
import pandas as pd
import csv
print('Starting the prorgam...')
os.chdir(sys.path[0])


class TableRow():
    def __init__(self, window, row):
        self.window = window
        self.row = row

    def getDataFromWindow(self):
        # print('>getDataFromWindow:')
        drugItem = self.variableOfDrugs.get()
        region = variableOfRegion.get()

        totalOptima = countSalesOptima(drugItem, region)

        self.sumDrugsInRegionOptima.set(int(totalOptima))
        #print(f'Total = {self.sumDrugsInRegion.get()}')

        totalVenta = countSalesVenta(drugItem, region)
        self.sumDrugsInRegionVenta.set(totalVenta)

        totalBadm = countSalesBadm(drugItem, region)
        self.sumDrugsInRegionBADM.set(totalBadm)

        totalSum = int(totalOptima) + int(totalVenta) + int(totalBadm)
        self.sumDrugsInRegions.set(totalSum)

    def createRow(self):
        Label(self.window, text=self.row).grid(column=0, row=self.row)
        self.variableOfDrugs = StringVar()
        self.comboDrags = ttk.Combobox(
            self.window, textvariable=self.variableOfDrugs, values=drugsListOptima, width=40)
        self.comboDrags['state'] = 'readonly'
        self.comboDrags.set('Оберіть препарат')
        self.comboDrags.grid(column=1, row=self.row)

        self.sumDrugsInRegionOptima = StringVar()
        sumLableOptima = Label(
            self.window, textvariable=self.sumDrugsInRegionOptima)
        sumLableOptima.grid(column=2, row=self.row)

        self.sumDrugsInRegionVenta = StringVar()
        sumLableVenta = Label(
            self.window, textvariable=self.sumDrugsInRegionVenta)
        sumLableVenta.grid(column=3, row=self.row)

        self.sumDrugsInRegionBADM = StringVar()
        sumLableBADM = Label(
            self.window, textvariable=self.sumDrugsInRegionBADM)
        sumLableBADM.grid(column=4, row=self.row)

        self.sumDrugsInRegions = StringVar()
        sumLable = Label(
            self.window, textvariable=self.sumDrugsInRegions)
        sumLable.grid(column=5, row=self.row)

        # Bind
        self.comboDrags.bind("<<ComboboxSelected>>",
                             lambda x: self.getDataFromWindow())

    def setDefoltDrugName(self, name):
        # print('>:setDefoltDrugName')
        self.comboDrags.set(name)

    def getDrugFromRaw(self):
        return self.variableOfDrugs.get()

    def getSumOptimaFromRaw(self):
        drugItem = self.variableOfDrugs.get()
        region = variableOfRegion.get()
        return countSalesOptima(drugItem, region)

    def getSumVentaFromRaw(self):
        drugItem = self.variableOfDrugs.get()
        region = variableOfRegion.get()
        return int(countSalesVenta(drugItem, region))

    def getSumBadmFromRaw(self):
        drugItem = self.variableOfDrugs.get()
        region = variableOfRegion.get()
        return countSalesBadm(drugItem, region)

    def getSumFromRaw(self):
        drugItem = self.variableOfDrugs.get()
        region = variableOfRegion.get()
        totalOptima = countSalesOptima(drugItem, region)
        totalVenta = countSalesVenta(drugItem, region)
        totalBadm = countSalesBadm(drugItem, region)
        totalSum = int(totalOptima) + int(totalVenta) + int(totalBadm)
        return totalSum


listDataOptima = []
drugsListOptima = []

listDataVenta = []
listDataBadm = []


regionListOptima = []

# List for UI row objects
listOfRowObjects = []

print('Try to open optima.xlsx')
read_file = pd.read_excel('optima.xlsx')

print('optima.xlsx in opened')

print('Try to convert optima.xlsx to optima.csv')
read_file.to_csv('optima.csv', index=None, header=True)

print('optima.csv is created')

with open('optima.csv', encoding='utf-8') as input:
    csv_reader = csv.DictReader(input, delimiter=',')
    for row in csv_reader:
        listDataOptima.append(row)

for drugItemOptima in listDataOptima:
    drugsListOptima.append(drugItemOptima['Товар'])
drugsListOptima = list(set(drugsListOptima))
drugsListOptima.sort()

for region in listDataOptima:
    regionListOptima.append(region['Область'])

regionListOptima = list(set(regionListOptima))
regionListOptima.sort()

# Read venta.xls and convert to venta.csv
read_fileVenta = pd.read_excel('venta.xls')
read_fileVenta.to_csv('venta.csv', index=None, header=True)
with open('venta.csv', encoding='utf-8') as input:
    csv_reader = csv.DictReader(input, delimiter=',')
    for row in csv_reader:
        listDataVenta.append(row)

# Read badm.xlsx and convert to venta.csv
read_fileBadm = pd.read_excel('badm.xlsx')
read_fileBadm.to_csv('badm.csv', index=None, header=True)
with open('badm.csv', encoding='utf-8') as input:
    csv_reader = csv.DictReader(input, delimiter=',')
    for row in csv_reader:
        listDataBadm.append(row)


def convertNameDrugInVenta():
    # Change drug item name as like in optima names
    for objectInListDataVenta in listDataVenta:
        # print(objectInListDataVenta['Товар'])
        if objectInListDataVenta['Товар'] == 'Аримидекс табл.п/п/о 1мг №28':
            objectInListDataVenta['Товар'] = 'АРИМИДЕКС ТАБ.П/О 1МГ #28'
        elif objectInListDataVenta['Товар'] == 'Беталок Зок табл.п/о 100мг №30':
            objectInListDataVenta['Товар'] = 'БЕТАЛОК ЗОК ТАБ. 100МГ #30'
        elif objectInListDataVenta['Товар'] == 'Беталок Зок табл.п/о 25мг №14':
            objectInListDataVenta['Товар'] = 'БЕТАЛОК ЗОК ТАБ. 25МГ #14'
        elif objectInListDataVenta['Товар'] == 'Беталок Зок табл.п/о 50мг №30':
            objectInListDataVenta['Товар'] = 'БЕТАЛОК ЗОК ТАБ. 50МГ #30'
        elif objectInListDataVenta['Товар'] == 'Беталок р-р д/ин.1мг/мл 5мл амп.№5':
            objectInListDataVenta['Товар'] = 'БЕТАЛОК Д/ИН.1МГ/МЛ 5МЛ АМП.#5'
        elif objectInListDataVenta['Товар'] == 'Брилинта табл. п/о 90мг №56':
            objectInListDataVenta['Товар'] = 'БРИЛИНТА ТАБ.П/О 90МГ#56(14X4)'
        elif objectInListDataVenta['Товар'] == 'Брилинта табл.п/п/о 60мг №56':
            objectInListDataVenta['Товар'] = 'БРИЛИНТА ТАБ.П/О 60МГ#56(14X4)'
        elif objectInListDataVenta['Товар'] == 'Будесонид Астразенека сусп.д/распылен.0.5мг/мл конт.2мл №20':
            objectInListDataVenta['Товар'] = 'БУДЕСОНИД АСТР.0.5МГ/МЛ2МЛ#20'
        elif objectInListDataVenta['Товар'] == 'Золадекс капс.д/подкож.введ.10.8мг шприц-аплик.№1':
            objectInListDataVenta['Товар'] = 'ЗОЛАДЕКС КАПС.10.8МГ#1ШПР-АППЛ'
        elif objectInListDataVenta['Товар'] == 'Золадекс капс.д/подкож.введ.10.8мг шприц-аплик.№1 1+1':
            objectInListDataVenta['Товар'] = 'ЗОЛАДЕКС К.10.8МГ#1ШПР-АППЛ1+1'
        elif objectInListDataVenta['Товар'] == 'Золадекс капс.д/подкож.введ.3.6мг шприц-аплик.№1':
            objectInListDataVenta['Товар'] = 'ЗОЛАДЕКС КАПС.3.6МГ#1 ШПР-АППЛ'
        elif objectInListDataVenta['Товар'] == 'Золадекс капс.д/подкож.введ.3.6мг шприц-аплик.№1 3уп.':
            objectInListDataVenta['Товар'] = 'ЗОЛАДЕКС К.3.6МГ#1 ШПР-АППЛ1+2'
        elif objectInListDataVenta['Товар'] == 'Касодекс табл.п/п/о 150мг №28':
            objectInListDataVenta['Товар'] = 'КАСОДЕКС ТАБ.П/О150МГ#28(14X2)'
        elif objectInListDataVenta['Товар'] == 'Касодекс табл.п/п/о 50мг №28':
            objectInListDataVenta['Товар'] = 'КАСОДЕКС ТАБ.П/О 50МГ #28'
        elif objectInListDataVenta['Товар'] == 'Комбоглиза XR табл.п/п/о 5мг/1000мг №28':
            objectInListDataVenta['Товар'] = 'КОМБОГЛИЗА XR ТАБ.5/1000МГ #28'
        elif objectInListDataVenta['Товар'] == 'Крестор табл.п/о 10мг №28':
            objectInListDataVenta['Товар'] = 'КРЕСТОР ТАБ.П/О 10МГ #28(14X2)'
        elif objectInListDataVenta['Товар'] == 'Крестор табл.п/о 20мг №28':
            objectInListDataVenta['Товар'] = 'КРЕСТОР ТАБ.П/О 20МГ #28(14X2)'
        elif objectInListDataVenta['Товар'] == 'Крестор табл.п/о 5мг №28':
            objectInListDataVenta['Товар'] = 'КРЕСТОР ТАБ.П/О 5МГ #28(14X2)'
        elif objectInListDataVenta['Товар'] == 'Крестор табл.п/п/о 40мг №28':
            objectInListDataVenta['Товар'] = 'КРЕСТОР ТАБ.П/О 40МГ #28(7X4)'
        elif objectInListDataVenta['Товар'] == 'Ксигдуо Пролонг табл.п/п/о пролонг.действ.10/1000мг №28':
            objectInListDataVenta['Товар'] = 'КСИГДУО ПРОЛ. ТАБ 10/1000МГ#28'
        elif objectInListDataVenta['Товар'] == 'Ксигдуо Пролонг табл.п/п/о пролонг.действ.5/1000мг №28':
            objectInListDataVenta['Товар'] = 'КСИГДУО ПРОЛОНГ ТАБ5/1000МГ#28'
        elif objectInListDataVenta['Товар'] == 'Линпарза табл.п/п/о 150мг №56':
            objectInListDataVenta['Товар'] = 'ЛИНПАРЗА ТАБ.П/О150МГ#56(8X7)'
        elif objectInListDataVenta['Товар'] == 'Онглиза табл.п/п/о 5мг №30':
            objectInListDataVenta['Товар'] = 'ОНГЛИЗА ТАБ.П/О 5МГ #30(10X3)'
        elif objectInListDataVenta['Товар'] == 'Пульмикорт сусп.д/распылен.0.25мг/мл 2мл конт.№20':
            objectInListDataVenta['Товар'] = 'ПУЛЬМИКОРТ СУСП0.25МГ/МЛ2МЛ#20'
        elif objectInListDataVenta['Товар'] == 'Пульмикорт сусп.д/распылен.0.25мг/мл 2мл конт.№20 1+1':
            objectInListDataVenta['Товар'] = 'ПУЛЬМИКОРТ 0.25МГ/МЛ2МЛ#20(1+1'
        elif objectInListDataVenta['Товар'] == 'Пульмикорт сусп.д/распылен.0.5мг/мл 2мл конт.№20':
            objectInListDataVenta['Товар'] = 'ПУЛЬМИКОРТ СУСП.0.5МГ/МЛ2МЛ#20'
        elif objectInListDataVenta['Товар'] == 'Пульмикорт Турбухалер пор.д/инг.100мкг/д.200д.№1':
            objectInListDataVenta['Товар'] = 'ПУЛЬМИКОРТ ТУРБ.100МКГ/Д.200Д.'
        elif objectInListDataVenta['Товар'] == 'Пульмикорт Турбухалер пор.д/инг.200мкг/д.100д.№1':
            objectInListDataVenta['Товар'] = 'ПУЛЬМИКОРТ ТУРБ.200МКГ/Д.100Д.'
        elif objectInListDataVenta['Товар'] == 'Сероквель XR табл.пролонг.дейст.200мг №60':
            objectInListDataVenta['Товар'] = 'СЕРОКВЕЛЬ XR ТАБ.200МГ #60'
        elif objectInListDataVenta['Товар'] == 'Сероквель XR табл.пролонг.дейст.50мг №60':
            objectInListDataVenta['Товар'] = 'СЕРОКВЕЛЬ XR ТАБ.50МГ #60'
        elif objectInListDataVenta['Товар'] == 'Симбикорт Турбухалер пор.д/инг.160мкг/4.5мкг/д.60д':
            objectInListDataVenta['Товар'] = 'СИМБИКОРТ ТУРБ.160/4.5/ДОЗА60Д'
        elif objectInListDataVenta['Товар'] == 'Симбикорт Турбухалер пор.д/инг.320мкг/9.0мкг/д.60д':
            objectInListDataVenta['Товар'] = 'СИМБИКОРТ ТУРБ.320/9.0/ДОЗА60Д'
        elif objectInListDataVenta['Товар'] == 'Симбикорт Турбухалер пор.д/инг.80мкг/4.5мкг/д.60д':
            objectInListDataVenta['Товар'] = 'СИМБИКОРТ ТУРБ.80/4.5/ДОЗА60Д'
        elif objectInListDataVenta['Товар'] == 'Тагриссо табл.п/п/о 80мг №30':
            objectInListDataVenta['Товар'] = 'ТАГРИССО ТАБ. П/О 80МГ#30'
        elif objectInListDataVenta['Товар'] == 'Фазлодекс р-р д/ин.250мг/5мл шприц 5мл №2':
            objectInListDataVenta['Товар'] = 'ФАЗЛОДЕКС 250МГ/5МЛ 5МЛ ШПР.#2'
        elif objectInListDataVenta['Товар'] == 'Форксига табл.п/п/о 10мг №30':
            objectInListDataVenta['Товар'] = 'ФОРКСИГА ТАБ.П/О 10МГ#30(10X3)'


def convertNameDrugInBadm():
    # Change drug item name as like in optima names
    for objectInListDataBadm in listDataBadm:
        # print(objectInListDataVenta['Товар'])
        if objectInListDataBadm['Товар'] == 'Аримідекс табл.в / пл.об. 1 мг N28 (14х2) блістер *':
            objectInListDataBadm['Товар'] = 'АРИМИДЕКС ТАБ.П/О 1МГ #28'
        elif objectInListDataBadm['Товар'] == 'Беталок ЗОК табл.в / пл.об. з упов.вив.100мг N30 фл. *':
            objectInListDataBadm['Товар'] = 'БЕТАЛОК ЗОК ТАБ. 100МГ #30'
        elif objectInListDataBadm['Товар'] == 'Беталок ЗОК табл.в / пл.об. з упов.вив.25мг N14 блістер *':
            objectInListDataBadm['Товар'] = 'БЕТАЛОК ЗОК ТАБ. 25МГ #14'
        elif objectInListDataBadm['Товар'] == 'Беталок ЗОК табл.в / пл.об. з упов.вив.50мг N30 фл. *':
            objectInListDataBadm['Товар'] = 'БЕТАЛОК ЗОК ТАБ. 50МГ #30'
        elif objectInListDataBadm['Товар'] == 'Беталок р-н д / ін.1мг / мл 5 мл амп. N5 *':
            objectInListDataBadm['Товар'] = 'БЕТАЛОК Д/ИН.1МГ/МЛ 5МЛ АМП.#5'
        elif objectInListDataBadm['Товар'] == 'Брилінта табл. в / пл.об. 90мг №56 (14х4) блістер':
            objectInListDataBadm['Товар'] = 'БРИЛИНТА ТАБ.П/О 90МГ#56(14X4)'
        elif objectInListDataBadm['Товар'] == 'Брилінта табл. в / пл.об. 60мг №56 (14х4) блістер':
            objectInListDataBadm['Товар'] = 'БРИЛИНТА ТАБ.П/О 60МГ#56(14X4)'
        elif objectInListDataBadm['Товар'] == 'Будесонід Астразенека сусп.д/расп.0.25 мг/мл 2мл конт.конв.№20(5х4) карт уп*':
            objectInListDataBadm['Товар'] = 'БУДЕСОНИД АСТР.0.25МГ/МЛ2МЛ#20'
        elif objectInListDataBadm['Товар'] == 'Будесонід Астразенека сусп.д/расп.0.5 мг/мл конт.конв.№20(5х4) карт уп***':
            objectInListDataBadm['Товар'] = 'БУДЕСОНИД АСТР.0.5МГ/МЛ2МЛ#20'
        elif objectInListDataBadm['Товар'] == 'Золадекс капс.д / підшк.введ.прол.дії10.8мг шпр.-аплік.з зах.мех.N1 *':
            objectInListDataBadm['Товар'] = 'ЗОЛАДЕКС КАПС.10.8МГ#1ШПР-АППЛ'
        elif objectInListDataBadm['Товар'] == 'Золадекс капс.д / підшк.введен.прол.дії 10.8мг шпр.-аплік.з зах.мех.N1 Смотка 2 уп *':
            objectInListDataBadm['Товар'] = 'ЗОЛАДЕКС К.10.8МГ#1ШПР-АППЛ1+1'
        elif objectInListDataBadm['Товар'] == 'Золадекс капс.д / підшк.введен.прол.дії 3.6мг шпр.-аплік.з зах.мех.N1 *':
            objectInListDataBadm['Товар'] = 'ЗОЛАДЕКС КАПС.3.6МГ#1 ШПР-АППЛ'
        elif objectInListDataBadm['Товар'] == 'Золадекс капс.д / підшк.введ.прол.дії 3.6мг шпр.-аплік.з зах.мех.N1 Смотка 3 уп *':
            objectInListDataBadm['Товар'] = 'ЗОЛАДЕКС К.3.6МГ#1 ШПР-АППЛ1+2'
        elif objectInListDataBadm['Товар'] == 'Касодекс® табл. в / пл.об.150мг N28 *' or objectInListDataBadm['Товар'] == 'Касодекс табл. в / пл.об.150мг N28 Смотка 2 уп *':
            objectInListDataBadm['Товар'] = 'КАСОДЕКС ТАБ.П/О150МГ#28(14X2)'
        elif objectInListDataBadm['Товар'] == 'Касодекс табл. в / пл.об. 50мг N28 (14х2) блістер Смотка 2 уп *':
            objectInListDataBadm['Товар'] = 'КАСОДЕКС ТАБ.П/О 50МГ #28'
        elif objectInListDataBadm['Товар'] == 'Касодекс табл. в / пл.об. 50мг N28 (14х2) блістер *':
            objectInListDataBadm['Товар'] = 'КАСОДЕКС ТАБ.П/О 50МГ #28'
        elif objectInListDataBadm['Товар'] == 'Комбогліза XR табл. в / пл.об. 5 мг / 1000мг №28 (7х4)':
            objectInListDataBadm['Товар'] = 'КОМБОГЛИЗА XR ТАБ.5/1000МГ #28'
        elif objectInListDataBadm['Товар'] == 'Крестор табл.в / пл.об.10мг N28 (14х2)':
            objectInListDataBadm['Товар'] = 'КРЕСТОР ТАБ.П/О 10МГ #28(14X2)'
        elif objectInListDataBadm['Товар'] == 'Крестор табл.в / пл.об. 20мг N28 (14х2) блістер':
            objectInListDataBadm['Товар'] = 'КРЕСТОР ТАБ.П/О 20МГ #28(14X2)'
        elif objectInListDataBadm['Товар'] == 'Крестор табл. в / пл.об. 5мг N28 (14х2) блістер':
            objectInListDataBadm['Товар'] = 'КРЕСТОР ТАБ.П/О 5МГ #28(14X2)'
        elif objectInListDataBadm['Товар'] == 'Крестор табл. в / пл.об. 40мг N28 (7х4)':
            objectInListDataBadm['Товар'] = 'КРЕСТОР ТАБ.П/О 40МГ #28(7X4)'
        elif objectInListDataBadm['Товар'] == 'Ксігдуо Пролонг табл.в / пл.об.прол.дії 10 / 1000мг №28 (7х4) блістер':
            objectInListDataBadm['Товар'] = 'КСИГДУО ПРОЛ. ТАБ 10/1000МГ#28'
        elif objectInListDataBadm['Товар'] == 'Ксігдуо Пролонг табл.в / пл.об.прол.дії 5 / 1000мг №28 (7х4) блістер':
            objectInListDataBadm['Товар'] = 'КСИГДУО ПРОЛОНГ ТАБ5/1000МГ#28'
        elif objectInListDataBadm['Товар'] == 'Лінпарза табл.вк/пл.об.150м№56(7х8) блістер' or objectInListDataBadm['Товар'] == 'Лінпарза табл.вк/пл.об.150м№56(7х8) блістер Акція':
            objectInListDataBadm['Товар'] = 'ЛИНПАРЗА ТАБ.П/О150МГ#56(8X7)'
        elif objectInListDataBadm['Товар'] == 'Онгліза табл.в / пл.об. 5мг N30 (10х3) блістер':
            objectInListDataBadm['Товар'] = 'ОНГЛИЗА ТАБ.П/О 5МГ #30(10X3)'
        elif objectInListDataBadm['Товар'] == 'Пульмікорт сусп.д / розп.0.25мг / мл 2 мл N20 (5х4) *':
            objectInListDataBadm['Товар'] = 'ПУЛЬМИКОРТ СУСП0.25МГ/МЛ2МЛ#20'
        elif objectInListDataBadm['Товар'] == 'Пульмикорт сусп.д/распылен.0.25мг/мл 2мл конт.№20 1+1':
            objectInListDataBadm['Товар'] = 'ПУЛЬМИКОРТ 0.25МГ/МЛ2МЛ#20(1+1'
        elif objectInListDataBadm['Товар'] == 'Пульмікорт сусп.д / розп.0.5мг / мл 2 мл N20 (5х4) ***':
            objectInListDataBadm['Товар'] = 'ПУЛЬМИКОРТ СУСП.0.5МГ/МЛ2МЛ#20'
        elif objectInListDataBadm['Товар'] == 'Пульмікорт Турбухалер пор.д / інг. 100 мкг / доза 200 доз ***':
            objectInListDataBadm['Товар'] = 'ПУЛЬМИКОРТ ТУРБ.100МКГ/Д.200Д.'
        elif objectInListDataBadm['Товар'] == 'Пульмікорт Турбухалер пор.д / інг. 200мкг / доза 100 доз ***':
            objectInListDataBadm['Товар'] = 'ПУЛЬМИКОРТ ТУРБ.200МКГ/Д.100Д.'
        elif objectInListDataBadm['Товар'] == 'Сероквель XR табл.пролонг.дейст.200мг №60':
            objectInListDataBadm['Товар'] = 'СЕРОКВЕЛЬ XR ТАБ.200МГ #60'
        elif objectInListDataBadm['Товар'] == 'Сероквель XR табл.пролонг.дейст.50мг №60':
            objectInListDataBadm['Товар'] = 'СЕРОКВЕЛЬ XR ТАБ.50МГ #60'
        elif objectInListDataBadm['Товар'] == 'Симбікорт Турбухалер пор.д / інг.160мкг / 4.5мкг / доза 60 доз ***':
            objectInListDataBadm['Товар'] = 'СИМБИКОРТ ТУРБ.160/4.5/ДОЗА60Д'
        elif objectInListDataBadm['Товар'] == 'Симбікорт Турбухалер пор.д / інг.320мкг / 9мкг / доза 60 доз ***':
            objectInListDataBadm['Товар'] = 'СИМБИКОРТ ТУРБ.320/9.0/ДОЗА60Д'
        elif objectInListDataBadm['Товар'] == 'Симбікорт Турбухалер пор.д / інг.80мкг / 4.5мкг / доза 60 доз ***':
            objectInListDataBadm['Товар'] = 'СИМБИКОРТ ТУРБ.80/4.5/ДОЗА60Д'
        elif objectInListDataBadm['Товар'] == 'Тагриссо табл.п/п/о 80мг №30':
            objectInListDataBadm['Товар'] = 'ТАГРИССО ТАБ. П/О 80МГ#30'
        elif objectInListDataBadm['Товар'] == 'Фазлодекс р-н д / ін 250мг / 5мл предв.заполн.шпріц 5мл N2 з 2-ма стер.ігламі':
            objectInListDataBadm['Товар'] = 'ФАЗЛОДЕКС 250МГ/5МЛ 5МЛ ШПР.#2'
        elif objectInListDataBadm['Товар'] == 'Форксіга табл.в / пл.об.10мг №30 (10х3) блістер' or objectInListDataBadm['Товар'] == 'Форксіга табл.в / пл.об.10мг №30 (10х3) блістер Спец':
            objectInListDataBadm['Товар'] = 'ФОРКСИГА ТАБ.П/О 10МГ#30(10X3)'
    pass


def countSalesOptima(drug, region):
    if drug == 'Оберіть препарат' or region == 'Оберіть область':
        # print('Введіть дані')
        return 0
    else:
        listSumKhmelnytskyi = []
        #total = 0
        for row in listDataOptima:
            if row['Область'] == region and row['Товар'] == drug:
                #total += float(row['Продажі шт'])
                listSumKhmelnytskyi.append(float(row['Продажи шт']))
        #print(f'{drug}, {region}: {sum(listSumKhmelnytskyi)}')
        return sum(listSumKhmelnytskyi)


def countSalesVenta(drug, region):
    # print('>countSalesVenta:')
    if drug == 'Оберіть препарат' or region == 'Оберіть область':
        # print('Введіть дані')
        return 0
    if region == 'М.КИЇВ':
        region = 'КИЇВСЬКА ОБЛАСТЬ'
    sumOfSales = None
    for myDrug in listDataVenta:
        if myDrug['Товар'] == drug:
            sumOfSales = myDrug[region]
    if sumOfSales == None:
        return 0
    return sumOfSales


def countSalesBadm(drug, region):
    # print('>countSalesBadm:')
    #print(drug, region)
    if drug == 'Оберіть препарат' or region == 'Оберіть область':
        # print('Введіть дані')
        return 0
    sum = 0
    for listObject in listDataBadm:
        if listObject['Товар'] == drug and listObject['Область'] == region:
            # print(listObject['Количество'])
            sum = sum + int(listObject['Количество'])
    # print(sum)
    return sum


def tableHead(window):
    # Create table head
    Label(window, text='№').grid(column=0, row=0)
    Label(window, text='Назва препарата').grid(column=1, row=0)
    Label(window, text='Оптіма').grid(column=2, row=0)
    Label(window, text='Вента').grid(column=3, row=0)
    Label(window, text='БаДМ').grid(column=4, row=0)
    Label(window, text='Сума', width=8).grid(column=5, row=0)
    Label(window, text='Область', width=32).grid(column=6, row=0)


def countDataFromDrag(event):
    # print('>getDataFromDrag')
    for listTableObject in listOfRowObjects:
        listTableObject.getDataFromWindow()


def convertRegionKeysInVenta():
    # print(">convertRegionInVenta:")
    for region in listDataVenta:
        new_key = "ВІННИЦЬКА ОБЛАСТЬ"
        old_key = "Вінницька"
        region[new_key] = region.pop(old_key)

        new_key = "ВОЛИНСЬКА ОБЛАСТЬ"
        old_key = "Волинська"
        region[new_key] = region.pop(old_key)

        new_key = "ДНІПРОПЕТРОВСЬКА ОБЛАСТЬ"
        old_key = "Дніпропетровська"
        region[new_key] = region.pop(old_key)

        new_key = "ДОНЕЦЬКА ОБЛАСТЬ"
        old_key = "Донецька"
        region[new_key] = region.pop(old_key)

        new_key = "ЖИТОМИРСЬКА ОБЛАСТЬ"
        old_key = "Житомирська"
        region[new_key] = region.pop(old_key)

        new_key = "ЗАКАРПАТСЬКА ОБЛАСТЬ"
        old_key = "Закарпатська"
        region[new_key] = region.pop(old_key)

        new_key = "ЗАПОРІЗЬКА ОБЛАСТЬ"
        old_key = "Запорізька"
        region[new_key] = region.pop(old_key)

        new_key = "ІВАНО-ФРАНКІВСЬКА ОБЛАСТЬ"
        old_key = "Івано-Франківська"
        region[new_key] = region.pop(old_key)

        new_key = "КИЇВСЬКА ОБЛАСТЬ"
        old_key = "Київська"
        region[new_key] = region.pop(old_key)

        new_key = "КІРОВОГРАДСЬКА ОБЛАСТЬ"
        old_key = "Кіровоградська"
        region[new_key] = region.pop(old_key)

        new_key = "ЛЬВІВСЬКА ОБЛАСТЬ"
        old_key = "Львівська"
        region[new_key] = region.pop(old_key)

        new_key = "МИКОЛАЇВСЬКА ОБЛАСТЬ"
        old_key = "Миколаївська"
        region[new_key] = region.pop(old_key)

        new_key = "ОДЕСЬКА ОБЛАСТЬ"
        old_key = "Одеська"
        region[new_key] = region.pop(old_key)

        new_key = "ПОЛТАВСЬКА ОБЛАСТЬ"
        old_key = "Полтавська"
        region[new_key] = region.pop(old_key)

        new_key = "РІВНЕНСЬКА ОБЛАСТЬ"
        old_key = "Рівненська"
        region[new_key] = region.pop(old_key)

        new_key = "СУМСЬКА ОБЛАСТЬ"
        old_key = "Сумська"
        region[new_key] = region.pop(old_key)

        new_key = "ТЕРНОПІЛЬСЬКА ОБЛАСТЬ"
        old_key = "Тернопільська"
        region[new_key] = region.pop(old_key)

        new_key = "ХАРКІВСЬКА ОБЛАСТЬ"
        old_key = "Харківська"
        region[new_key] = region.pop(old_key)

        new_key = "ХАРКІВСЬКА ОБЛАСТЬ"
        old_key = "Херсонська"
        region[new_key] = region.pop(old_key)

        new_key = "ХМЕЛЬНИЦЬКА ОБЛАСТЬ"
        old_key = "Хмельницька"
        region[new_key] = region.pop(old_key)

        new_key = "ЧЕРКАСЬКА ОБЛАСТЬ"
        old_key = "Черкаська"
        region[new_key] = region.pop(old_key)

        new_key = "ЧЕРНІГІВСЬКА ОБЛАСТЬ"
        old_key = "Чернігівська"
        region[new_key] = region.pop(old_key)

        new_key = "ЧЕРНІВЕЦЬКА ОБЛАСТЬ"
        old_key = "Чернівецька"
        region[new_key] = region.pop(old_key)

        new_key = "КРИМ"
        old_key = "Крим"
        region[new_key] = region.pop(old_key)

        new_key = "ЛУГАНСЬКА ОБЛАСТЬ"
        old_key = "Луганська"
        region[new_key] = region.pop(old_key)


def convertRegionInBadm():
    for region in listDataBadm:
        if region['Область'] == 'ВИННИЦКАЯ':
            region['Область'] = 'ВІННИЦЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ВОЛЫНСКАЯ':
            region['Область'] = 'ВОЛИНСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ДНЕПРОПЕТРОВСКАЯ':
            region['Область'] = 'ДНІПРОПЕТРОВСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ДОНЕЦКАЯ':
            region['Область'] = 'ДОНЕЦЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ЖИТОМИРСКАЯ':
            region['Область'] = 'ЖИТОМИРСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ЗАКАРПАТСКАЯ':
            region['Область'] = 'ЗАКАРПАТСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ЗАПОРОЖСКАЯ':
            region['Область'] = 'ЗАПОРІЗЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ИВАНО-ФРАНКОВСКАЯ':
            region['Область'] = 'ІВАНО-ФРАНКІВСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'КИЕВСКАЯ':
            region['Область'] = 'КИЇВСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'КИРОВОГРАДСКАЯ':
            region['Область'] = 'КІРОВОГРАДСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ЛЬВОВСКАЯ':
            region['Область'] = 'ЛЬВІВСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'НИКОЛАЕВСКАЯ':
            region['Область'] = 'МИКОЛАЇВСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ОДЕССКАЯ':
            region['Область'] = 'ОДЕСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ПОЛТАВСКАЯ':
            region['Область'] = 'ПОЛТАВСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'РОВЕНСКАЯ':
            region['Область'] = 'РІВНЕНСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'СУМСКАЯ':
            region['Область'] = 'СУМСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ТЕРНОПОЛЬСКАЯ':
            region['Область'] = 'ТЕРНОПІЛЬСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ХАРЬКОВСКАЯ':
            region['Область'] = 'ХАРКІВСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ХМЕЛЬНИЦКАЯ':
            region['Область'] = 'ХМЕЛЬНИЦЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ЧЕРКАССКАЯ':
            region['Область'] = 'ЧЕРКАСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ЧЕРНИГОВСКАЯ':
            region['Область'] = 'ЧЕРНІВЕЦЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ЧЕРНОВИЦКАЯ':
            region['Область'] = 'ЧЕРНІВЕЦЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ВІННИЦЬКА ОБЛАСТЬ':
            region['Область'] = 'ВІННИЦЬКА ОБЛАСТЬ'
    pass


def setDefoltDrugsInObject(listOfRowObjects):
    listOfRowObjects[0].setDefoltDrugName('БЕТАЛОК Д/ИН.1МГ/МЛ 5МЛ АМП.#5')
    listOfRowObjects[1].setDefoltDrugName('БЕТАЛОК ЗОК ТАБ. 25МГ #14')
    listOfRowObjects[2].setDefoltDrugName('БЕТАЛОК ЗОК ТАБ. 50МГ #30')
    listOfRowObjects[3].setDefoltDrugName('БЕТАЛОК ЗОК ТАБ. 100МГ #30')
    listOfRowObjects[4].setDefoltDrugName('БРИЛИНТА ТАБ.П/О 90МГ#56(14X4)')
    listOfRowObjects[5].setDefoltDrugName('КОМБОГЛИЗА XR ТАБ.5/1000МГ #28')
    listOfRowObjects[6].setDefoltDrugName('КРЕСТОР ТАБ.П/О 5МГ #28(14X2)')
    listOfRowObjects[7].setDefoltDrugName('КРЕСТОР ТАБ.П/О 10МГ #28(14X2)')
    listOfRowObjects[8].setDefoltDrugName('КРЕСТОР ТАБ.П/О 20МГ #28(14X2)')
    listOfRowObjects[9].setDefoltDrugName('КРЕСТОР ТАБ.П/О 40МГ #28(7X4)')
    listOfRowObjects[10].setDefoltDrugName('КСИГДУО ПРОЛОНГ ТАБ5/1000МГ#28')
    listOfRowObjects[11].setDefoltDrugName('КСИГДУО ПРОЛ. ТАБ 10/1000МГ#28')
    listOfRowObjects[12].setDefoltDrugName('ОНГЛИЗА ТАБ.П/О 5МГ #30(10X3)')
    listOfRowObjects[13].setDefoltDrugName('ФОРКСИГА ТАБ.П/О 10МГ#30(10X3)')
    pass


def openAuthorLink(url):
    webbrowser.open_new(url)


def insertWebLinks():
    authorLink = Label(
        root_tk, text='maksym.protsak@gmail.com V1.1', fg="blue", cursor="hand2")

    authorLink.place(x=500, y=21*16)
    authorLink.bind(
        "<Button-1>", lambda e: openAuthorLink("https://tangerine-youtiao-a51230.netlify.app/"))

    optimaLink = Label(
        root_tk, text='optimapharm.ua', fg="blue", cursor="hand2")
    optimaLink.place(x=500, y=21*12)
    optimaLink.bind(
        "<Button-1>", lambda e: openAuthorLink("https://optimapharm.ua/"))

    ventaLink = Label(root_tk, text='ventaltd.com.ua',
                      fg="blue", cursor="hand2")
    ventaLink.place(x=500, y=21*13)
    ventaLink.bind(
        "<Button-1>", lambda e: openAuthorLink(f"https://www.ventaltd.com.ua/"))

    ventaLink = Label(root_tk, text='badm.ua',
                      fg="blue", cursor="hand2")
    ventaLink.place(x=500, y=21*14)
    ventaLink.bind(
        "<Button-1>", lambda e: openAuthorLink(f"https://www.badm.ua/ua/"))


def exportExcelFile():
    print(">exportExcelFile:")
    nums = []
    drugs = []
    optima = []
    venta = []
    badm = []
    sum = []
    region = []
    i = 0
    print(len(listOfRowObjects))
    for object in listOfRowObjects:
        if object.getDrugFromRaw() == 'Оберіть препарат':
            pass
        else:
            i += 1
            nums.append(i)
            drugs.append(object.getDrugFromRaw())
            optima.append(object.getSumOptimaFromRaw())
            venta.append(object.getSumVentaFromRaw())
            badm.append(object.getSumBadmFromRaw())
            sum.append(object.getSumFromRaw())
            region.append(variableOfRegion.get())
    data = {'№': nums, 'Препарат': drugs,
            'Оптіма': optima, 'Вента': venta, 'Бадм': badm, 'Сума': sum, 'Регіон': region}

    df = pd. DataFrame(data)
    df.to_excel('Розрахунок.xlsx', index=False)
    tkinter.messagebox.showinfo(
        "Експорт в Excel файл",  'Файл "Розрахунок.xlsx" збережено')


convertNameDrugInVenta()
convertRegionKeysInVenta()

convertNameDrugInBadm()
convertRegionInBadm()


####UI####
root_tk = tk.Tk()

# Main settings
root_tk.config(bg='#EBEBEB')
root_tk.title('Підрахунок продажів')
root_tk.geometry('700x360+400+100')
root_tk.resizable(False, False)

tableHead(root_tk)

# Creating comboRegion
variableOfRegion = StringVar()
comboRegion = ttk.Combobox(
    root_tk, textvariable=variableOfRegion, values=regionListOptima, width=30)
comboRegion['state'] = 'readonly'
comboRegion.set('Оберіть область')
comboRegion.grid(column=6, row=1)

for i in range(1, 16):
    tableRowObject = TableRow(root_tk, i)
    listOfRowObjects.append(tableRowObject)
    tableRowObject.createRow()

# Create button for export in Excel file
buttonExportExcel = Button(
    root_tk, text='Експорт в Excel файл', command=exportExcelFile)
buttonExportExcel.grid(column=6, row=3)
setDefoltDrugsInObject(listOfRowObjects)

comboRegion.bind("<<ComboboxSelected>>", countDataFromDrag)

insertWebLinks()

root_tk.mainloop()
