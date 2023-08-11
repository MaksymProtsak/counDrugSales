import os
import sys
from time import strftime
import webbrowser
from tkinter import StringVar, ttk, filedialog
import tkinter as tk
from tkinter import *
import tkinter.messagebox
import pandas as pd
import csv
from shutil import copy2

print('Starting the prorgam...')
os.chdir(sys.path[0])

drugList = []  # From txt file

listDataOptima = []
listDataVenta = []
listDataBadm = []
listDataKoneks = []

regionListOptima = []

# List for UI row objects
listOfRowObjects = []


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
        # print(f'Total = {self.sumDrugsInRegion.get()}')

        totalVenta = countSalesVenta(drugItem, region)
        self.sumDrugsInRegionVenta.set(totalVenta)

        totalBadm = countSalesBadm(drugItem, region)
        self.sumDrugsInRegionBADM.set(totalBadm)

        totalKoneks = countSalesKoneks(drugItem, region)
        self.sumDrugsInRegionKoneks.set(int(totalKoneks))
        #print(f'Total = {self.sumDrugsInRegionKoneks.get()}')

        totalSum = int(totalOptima) + int(totalVenta) + \
            int(totalBadm) + int(totalKoneks)
        self.sumDrugsInRegions.set(totalSum)

    def createRow(self):
        Label(self.window, text=self.row).grid(column=0, row=self.row)
        self.variableOfDrugs = StringVar()
        self.comboDrags = ttk.Combobox(
            self.window, textvariable=self.variableOfDrugs, values=drugList, width=40)
        self.comboDrags['state'] = 'readonly'
        #self.comboDrags.set('Оберіть препарат')
        self.comboDrags.grid(column=1, row=self.row)

        self.sumDrugsInRegionBADM = StringVar()
        sumLableBADM = Label(
            self.window, textvariable=self.sumDrugsInRegionBADM, width=6)
        sumLableBADM.grid(column=2, row=self.row)

        self.sumDrugsInRegionOptima = StringVar()
        sumLableOptima = Label(
            self.window, textvariable=self.sumDrugsInRegionOptima, width=6)
        sumLableOptima.grid(column=3, row=self.row)

        self.sumDrugsInRegionKoneks = StringVar()
        sumLableKoneks = Label(
            self.window, textvariable=self.sumDrugsInRegionKoneks, width=7)
        sumLableKoneks.grid(column=4, row=self.row)

        self.sumDrugsInRegionVenta = StringVar()
        sumLableVenta = Label(
            self.window, textvariable=self.sumDrugsInRegionVenta, width=6)
        sumLableVenta.grid(column=5, row=self.row)

        self.sumDrugsInRegions = StringVar()
        sumLable = Label(
            self.window, textvariable=self.sumDrugsInRegions, width=6)
        sumLable.grid(column=6, row=self.row)

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

    def getSumBadmFromKoneks(self):
        drugItem = self.variableOfDrugs.get()
        region = variableOfRegion.get()
        return countSalesKoneks(drugItem, region)

    def getSumFromRaw(self):
        drugItem = self.variableOfDrugs.get()
        region = variableOfRegion.get()
        totalOptima = countSalesOptima(drugItem, region)
        totalVenta = countSalesVenta(drugItem, region)
        totalBadm = countSalesBadm(drugItem, region)
        totalKoneks = countSalesKoneks(drugItem, region)
        totalSum = int(totalOptima) + int(totalVenta) + \
            int(totalBadm)+int(totalKoneks)
        return totalSum


# Create a list of drugs from txt file
my_file = open("data\\drugList.txt", "r", encoding='utf-8')
data = my_file.read()
drugList = data.split("\n")

# print("Read optima.xls and convert to optima.csv: START")
# Read optima.xls and convert to optima.csv
read_fileOptima = pd.read_excel('data\optima.xlsx')
read_fileOptima.to_csv('data\optima.csv', index=None, header=True)
# print("Read optima.xls and convert to optima.csv: FINISH")


with open('data\optima.csv', encoding='utf-8') as input:
    csv_reader = csv.DictReader(input, delimiter=',')
    for row in csv_reader:
        listDataOptima.append(row)

for drugItemOptima in listDataOptima:
    drugList.append(drugItemOptima['Товар'])
drugList = list(set(drugList))
drugList.sort()


for region in listDataOptima:
    regionListOptima.append(region['Область'])

regionListOptima = list(set(regionListOptima))
regionListOptima.sort()

# Read venta.xls and convert to venta.csv
read_fileVenta = pd.read_excel('data\\venta.xls')
read_fileVenta.to_csv('data\\venta.csv', index=None, header=True)
with open('data\\venta.csv', encoding='utf-8') as input:
    csv_reader = csv.DictReader(input, delimiter=',')
    for row in csv_reader:
        listDataVenta.append(row)

# Read badm.xlsx and convert to badm.csv
read_fileBadm = pd.read_excel('data\\badm.xlsx')
read_fileBadm.to_csv('data\\badm.csv', index=None, header=True)
with open('data\\badm.csv', encoding='utf-8') as input:
    csv_reader = csv.DictReader(input, delimiter=',')
    for row in csv_reader:
        listDataBadm.append(row)

# Read koneks.xls and convert to koneks.csv
read_fileKoneks = pd.read_excel(
    'data\\koneks.xls', skiprows=5, usecols=[1, 4, 16, 17])
read_fileKoneks.to_csv('data\\koneks.csv', index=None, header=True)

with open('data\\koneks.csv', encoding='utf-8') as input:
    csv_reader = csv.DictReader(input, delimiter=',')
    for row in csv_reader:
        listDataKoneks.append(row)


def tableData():
    w = LabelFrame(root_tk, text='Завантажені файли:', padx=22, pady=6)
    w.place(x=480+100, y=50)
    label = Label(w, text="badm.xlsx", padx=10)
    label.grid(column=0, row=0, ipady=4)
    button = Button(w, text='Оновити',
                    command=lambda: updateTableFile('badm.xlsx', '*.xlsx'))
    button.grid(column=1, row=0)

    label = Label(w, text="optima.xlsx", padx=10)
    label.grid(column=0, row=1, ipady=4)
    button = Button(w, text='Оновити',
                    command=lambda: updateTableFile("optima.xlsx", '*.xlsx'))
    button.grid(column=1, row=1)

    label = Label(w, text="venta.xls", padx=10)
    label.grid(column=0, row=2, ipady=4)
    button = Button(w, text='Оновити',
                    command=lambda: updateTableFile("venta.xls", '*.xls'))
    button.grid(column=1, row=2)

    label = Label(w, text="koneks.xls", padx=10)
    label.grid(column=0, row=3, ipady=4)
    button = Button(w, text='Оновити',
                    command=lambda: updateTableFile("koneks.xls", '*.xls'))
    button.grid(column=1, row=3)


def updateTableFile(fileName, typeType):
    # print(fileName)
    source = filedialog.askopenfilename(
        initialdir="/",
        title=f'Оновити таблицю "{fileName}"',
        filetypes=(("Таблиця Excel", f"{typeType}"),
                   ("all files", "*.*")))

    copy2(source, f'data\\{fileName}')

    tkinter.messagebox.showinfo(
        "Оновити таблицю",  f'Файл {fileName} оновлено!')

    if fileName == 'badm.xlsx':
        updateDataBaDM()
    elif fileName == "optima.xlsx":
        updateDataOptima()
    elif fileName == "venta.xls":
        updateDataVenta()
    elif fileName == "koneks.xls":
        updateDataKoneks()

    updateTableGUIFrame()
    # print('updateTableFile:DONE')


def updateDataBaDM():
    # print('updateDataBaDM:')

    listDataBadm.clear()
    # print(listDataBadm)

    read_fileBadm = pd.read_excel('data\\badm.xlsx')
    read_fileBadm.to_csv('data\\badm.csv', index=None, header=True)
    with open('data\\badm.csv', encoding='utf-8') as input:
        csv_reader = csv.DictReader(input, delimiter=',')
        for row in csv_reader:
            listDataBadm.append(row)

    convertNameDrugInBadm()
    convertRegionInBadm()


def updateDataOptima():
    # print('updateDataOptima:')
    listDataOptima.clear()
    # print(len(listDataOptima))

    read_fileBadm = pd.read_excel('data\\optima.xlsx')
    read_fileBadm.to_csv('data\\optima.csv', index=None, header=True)
    with open('data\\optima.csv', encoding='utf-8') as input:
        csv_reader = csv.DictReader(input, delimiter=',')
        for row in csv_reader:
            listDataOptima.append(row)
    pass


def updateDataVenta():
    # print('updateDataVenta:')

    listDataVenta.clear()
    # print(listDataVenta)

    read_fileBadm = pd.read_excel('data\\venta.xls')
    read_fileBadm.to_csv('data\\venta.csv', index=None, header=True)
    with open('data\\venta.csv', encoding='utf-8') as input:
        csv_reader = csv.DictReader(input, delimiter=',')
        for row in csv_reader:
            listDataVenta.append(row)
    # print(len(listDataVenta))
    convertNameDrugInVenta()
    convertRegionKeysInVenta()


def updateDataKoneks():
    print('updateDataKoneks:')

    listDataKoneks.clear()
    print(listDataKoneks)

    read_fileKoneks = pd.read_excel(
        'data\\koneks.xls', skiprows=5, usecols=[1, 4, 16, 17])
    read_fileKoneks.to_csv('data\\koneks.csv', index=None, header=True)
    with open('data\\koneks.csv', encoding='utf-8') as input:
        csv_reader = csv.DictReader(input, delimiter=',')
        for row in csv_reader:
            listDataKoneks.append(row)
    print(len(listDataKoneks))
    convertNameDrugInKoneks()
    convertRegionInKoneks()
    print(listDataKoneks[:5])
    print('updateDataKoneks:DONE')
    pass


def updateTableGUIFrame():
    for object in listOfRowObjects:
        object.getDataFromWindow()
    # setDefoltDrugsInObject(listOfRowObjects)


def convertNameDrugInVenta():
    # Change drug item name as like in optima names
    for objectInListDataVenta in listDataVenta:
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
        elif objectInListDataVenta['Товар'] == 'Будесонид Астразенека сусп.д/распылен.0.25мг/мл конт.2мл №20':
            objectInListDataVenta['Товар'] = 'БУДЕСОНИД АСТР.0.25МГ/МЛ2МЛ#20'
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
        elif objectInListDataVenta['Товар'] == 'Линпарза табл.п/п/о 100мг №56':
            objectInListDataVenta['Товар'] = 'ЛИНПАРЗА ТАБ.П/О100МГ#56(8X7)'


def convertNameDrugInBadm():
    # Change drug item name as like in badm names
    # print("convertNameDrugInBadm: START")
    for objectInListDataBadm in listDataBadm:
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
        elif objectInListDataBadm['Товар'] == "Лінпарза табл.вк/пл.об.150мг №56(7х8) блістер Акція":
            objectInListDataBadm['Товар'] = 'ЛИНПАРЗА ТАБ.П/О150МГ#56(8X7)'
        elif objectInListDataBadm['Товар'] == "Лінпарза табл.вк/пл.об.150мг \№56(7х8) блістер":
            objectInListDataBadm['Товар'] = 'ЛИНПАРЗА ТАБ.П/О150МГ#56(8X7)'
        elif objectInListDataBadm['Товар'] == "Лінпарза табл.вк/пл.об.100мг №56(7х8) блістер":
            objectInListDataBadm['Товар'] = "ЛИНПАРЗА ТАБ.П/О100МГ#56(8X7)"
        elif objectInListDataBadm['Товар'] == "Лінпарза .табл.вк/пл.об.100мг №56(7х8) блістер Акція":
            objectInListDataBadm['Товар'] = "ЛИНПАРЗА ТАБ.П/О100МГ#56(8X7)"
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
        elif objectInListDataBadm['Товар'] == "Сероквель XR табл.прол.дії 200мг N60(10х6) блістер":
            objectInListDataBadm['Товар'] = "СЕРОКВЕЛЬ XR ТАБ.200МГ #60"
        elif objectInListDataBadm['Товар'] == 'Сероквель XR табл.пролонг.дейст.200мг №60':
            objectInListDataBadm['Товар'] = 'СЕРОКВЕЛЬ XR ТАБ.200МГ #60'
        elif objectInListDataBadm['Товар'] == 'Сероквель XR табл.прол.дії 200мг N60 (10х6) блістер':
            objectInListDataBadm['Товар'] = "СЕРОКВЕЛЬ XR ТАБ.200МГ #60"
        elif objectInListDataBadm['Товар'] == 'Сероквель XR табл.пролонг.дейст.50мг №60':
            objectInListDataBadm['Товар'] = 'СЕРОКВЕЛЬ XR ТАБ.50МГ #60'
        elif objectInListDataBadm['Товар'] == 'Сероквель XR табл.прол.дії50мг N60 (10х6) блістер':
            objectInListDataBadm['Товар'] = 'СЕРОКВЕЛЬ XR ТАБ.50МГ #60'
        elif objectInListDataBadm['Товар'] == 'Симбікорт Турбухалер пор.д / інг.160мкг / 4.5мкг / доза 60 доз ***':
            objectInListDataBadm['Товар'] = 'СИМБИКОРТ ТУРБ.160/4.5/ДОЗА60Д'
        elif objectInListDataBadm['Товар'] == 'Симбікорт Турбухалер пор.д / інг.320мкг / 9мкг / доза 60 доз ***':
            objectInListDataBadm['Товар'] = 'СИМБИКОРТ ТУРБ.320/9.0/ДОЗА60Д'
        elif objectInListDataBadm['Товар'] == 'Симбікорт Турбухалер пор.д / інг.80мкг / 4.5мкг / доза 60 доз ***':
            objectInListDataBadm['Товар'] = 'СИМБИКОРТ ТУРБ.80/4.5/ДОЗА60Д'
        elif objectInListDataBadm['Товар'] == 'Тагриссо табл.п/п/о 80мг №30':
            objectInListDataBadm['Товар'] = 'ТАГРИССО ТАБ. П/О 80МГ#30'
        elif objectInListDataBadm['Товар'] == 'Тагріссо табл. п / в / пл.об.б. 80мг №30 (10х3) блістер Акція':
            objectInListDataBadm['Товар'] = 'ТАГРИССО ТАБ. П/О 80МГ#30'
        elif objectInListDataBadm['Товар'] == 'Тагріссо табл. в / пл.об. 80мг №30 (10х3) блістер':
            objectInListDataBadm['Товар'] = 'ТАГРИССО ТАБ. П/О 80МГ#30'
        elif objectInListDataBadm['Товар'] == 'Фазлодекс р-н д / ін 250мг / 5мл предв.заполн.шпріц 5мл N2 з 2-ма стер.ігламі':
            objectInListDataBadm['Товар'] = 'ФАЗЛОДЕКС 250МГ/5МЛ 5МЛ ШПР.#2'
        elif objectInListDataBadm['Товар'] == "Фазлодекс р-н д / ін 250мг / 5мл предв.заполн.шпріц 5мл N2 з 2-ма стер.ігламі*":
            objectInListDataBadm['Товар'] = "ФАЗЛОДЕКС 250МГ/5МЛ 5МЛ ШПР.#2"
        elif objectInListDataBadm['Товар'] == "Форксіга табл.в / пл.об.10мг №30 (10х3) блістер*":
            objectInListDataBadm['Товар'] = "ФОРКСИГА ТАБ.П/О 10МГ#30(10X3)"
        elif objectInListDataBadm['Товар'] == 'Форксіга табл.в / пл.об.10мг №30 (10х3) блістер' or objectInListDataBadm['Товар'] == 'Форксіга табл.в / пл.об.10мг №30 (10х3) блістер Спец':
            objectInListDataBadm['Товар'] = 'ФОРКСИГА ТАБ.П/О 10МГ#30(10X3)'
        elif objectInListDataBadm['Товар'] == 'Онгліза табл.в / пл.об. 2.5мг N30 (10х3) блістер':
            objectInListDataBadm['Товар'] = 'ОНГЛИЗА ТАБ.П/О 2.5МГ #30(10X3)'

    # for object in listDataBadm:
    #     print(object["Товар"])
    # print("convertNameDrugInBadm: START")


def convertNameDrugInKoneks():
    # Change drug item name as like in koneks names
    for objectInListDataKoneks in listDataKoneks:
        if objectInListDataKoneks['Назва продукту'] == 'Беталок Зок таб. в/о 100мг №30':
            objectInListDataKoneks['Назва продукту'] = 'БЕТАЛОК ЗОК ТАБ. 100МГ #30'
        elif objectInListDataKoneks['Назва продукту'] == 'Беталок Зок таб. в/о 25мг №14':
            objectInListDataKoneks['Назва продукту'] = 'БЕТАЛОК ЗОК ТАБ. 25МГ #14'
        elif objectInListDataKoneks['Назва продукту'] == 'Беталок Зок таб. в/о 50мг №30':
            objectInListDataKoneks['Назва продукту'] = 'БЕТАЛОК ЗОК ТАБ. 50МГ #30'
        elif objectInListDataKoneks['Назва продукту'] == 'Брилінта таб. в/о 60мг №56 (14х4) бліст.':
            objectInListDataKoneks['Назва продукту'] = 'БРИЛИНТА ТАБ.П/О 60МГ#56(14X4)'
        elif objectInListDataKoneks['Назва продукту'] == 'Брилінта таб. п/о 90мг №56 (14х4) бліст.':
            objectInListDataKoneks['Назва продукту'] = 'БРИЛИНТА ТАБ.П/О 90МГ#56(14X4)'
        elif objectInListDataKoneks['Назва продукту'] == 'Будесонід Астразенека сусп.д/росп 0,5мг/мл 2мл №20':
            objectInListDataKoneks['Назва продукту'] = 'БУДЕСОНИД АСТР.0.5МГ/МЛ2МЛ#20'
        elif objectInListDataKoneks['Назва продукту'] == 'Будесонід Астразенека сусп.д/роз.0,5мг/мл 2мл №20':
            objectInListDataKoneks['Назва продукту'] = 'БУДЕСОНИД АСТР.0.5МГ/МЛ2МЛ#20'
        elif objectInListDataKoneks['Назва продукту'] == 'Будесонід Астразенека сусп.д/роз.0,25мг/мл 2мл №20':
            objectInListDataKoneks['Назва продукту'] = 'БУДЕСОНИД АСТР.0.25МГ/МЛ2МЛ#20'
        elif objectInListDataKoneks['Назва продукту'] == 'Золадекс капс. 3.6мг шприц-аплікатор №1':
            objectInListDataKoneks['Назва продукту'] = 'ЗОЛАДЕКС К.3.6МГ#1 ШПР-АППЛ1+2'
        elif objectInListDataKoneks['Назва продукту'] == 'Комбогліза XR таб. в/о 5мг/1000мг №28 (7х4)бліст':
            objectInListDataKoneks['Назва продукту'] = 'КОМБОГЛИЗА XR ТАБ.5/1000МГ #28'
        elif objectInListDataKoneks['Назва продукту'] == 'Крестор таб. в/о 10мг №28 (14х2) бліст.':
            objectInListDataKoneks['Назва продукту'] = 'КРЕСТОР ТАБ.П/О 10МГ #28(14X2)'
        elif objectInListDataKoneks['Назва продукту'] == 'Крестор таб. в/о 20мг №28 (14х2) бліст.':
            objectInListDataKoneks['Назва продукту'] = 'КРЕСТОР ТАБ.П/О 20МГ #28(14X2)'
        elif objectInListDataKoneks['Назва продукту'] == 'Крестор таб. в/о 40мг №28 (7х4) бліст.':
            objectInListDataKoneks['Назва продукту'] = 'КРЕСТОР ТАБ.П/О 40МГ #28(7X4)'
        elif objectInListDataKoneks['Назва продукту'] == 'Крестор таб. в/о 5мг №28 (14х2) бліст.':
            objectInListDataKoneks['Назва продукту'] = 'КРЕСТОР ТАБ.П/О 5МГ #28(14X2)'
        elif objectInListDataKoneks['Назва продукту'] == 'Ксігдуо Пролонг таб. в/о 5мг/1000мг №28(4х7)бліст.':
            objectInListDataKoneks['Назва продукту'] = 'КСИГДУО ПРОЛОНГ ТАБ5/1000МГ#28'
        elif objectInListDataKoneks['Назва продукту'] == 'Ксігдуо Пролонг таб.в/о 10мг/1000мг №28(4х7)бліст.':
            objectInListDataKoneks['Назва продукту'] = 'КСИГДУО ПРОЛ. ТАБ 10/1000МГ#28'
        elif objectInListDataKoneks['Назва продукту'] == 'Онгліза 5мг таб.№30':
            objectInListDataKoneks['Назва продукту'] = 'ОНГЛИЗА ТАБ.П/О 5МГ #30(10X3)'
        elif objectInListDataKoneks['Назва продукту'] == 'Пульмікорт сусп.д/розп. 0,25мг/мл фл. 2мл №20 п/е':
            objectInListDataKoneks['Назва продукту'] = 'ПУЛЬМИКОРТ 0.25МГ/МЛ2МЛ#20(1+1'
        elif objectInListDataKoneks['Назва продукту'] == 'Пульмікорт сусп.д/розп. 0,5мг/мл фл. 2мл №20 п/е':
            objectInListDataKoneks['Назва продукту'] = 'ПУЛЬМИКОРТ СУСП.0.5МГ/МЛ2МЛ#20'
        elif objectInListDataKoneks['Назва продукту'] == 'Пульмікорт турбухалер пор. д/інг.200мкг 100доз':
            objectInListDataKoneks['Назва продукту'] = 'ПУЛЬМИКОРТ ТУРБ.200МКГ/Д.100Д.'
        elif objectInListDataKoneks['Назва продукту'] == 'Пульмікорт турбухалер пор. д/інг.100мкг 200доз':
            objectInListDataKoneks['Назва продукту'] = 'ПУЛЬМИКОРТ ТУРБ.100МКГ/Д.200Д.'
        elif objectInListDataKoneks['Назва продукту'] == 'Симбікорт турбух. пор. дл/інг. 160мкг+4,5мкг 60д':
            objectInListDataKoneks['Назва продукту'] = 'СИМБИКОРТ ТУРБ.160/4.5/ДОЗА60Д'
        elif objectInListDataKoneks['Назва продукту'] == 'Симбікорт турбух. пор. дл/інг. 320мкг+9мкг 60д':
            objectInListDataKoneks['Назва продукту'] = 'СИМБИКОРТ ТУРБ.320/9.0/ДОЗА60Д'
        elif objectInListDataKoneks['Назва продукту'] == 'Симбікорт турбух. пор. дл/інг. 80мкг+4,5мкг 60д':
            objectInListDataKoneks['Назва продукту'] = 'СИМБИКОРТ ТУРБ.80/4.5/ДОЗА60Д'
        elif objectInListDataKoneks['Назва продукту'] == 'Форксіга таб. в/о 10мг №30 (10х3) бліст.':
            objectInListDataKoneks['Назва продукту'] = 'ФОРКСИГА ТАБ.П/О 10МГ#30(10X3)'


def countSalesOptima(drug, region):
    if drug == 'Оберіть препарат' or region == 'Оберіть область':
        # print('Введіть дані')
        return 0
    else:
        listSumKhmelnytskyi = []
        # total = 0
        for row in listDataOptima:
            if row['Область'] == region and row['Товар'] == drug:
                # total += float(row['Продажі шт'])
                listSumKhmelnytskyi.append(float(row['Продажи шт']))
        # print(f'{drug}, {region}: {sum(listSumKhmelnytskyi)}')
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
    if drug == 'Оберіть препарат' or region == 'Оберіть область':
        # print('Введіть дані')
        return 0
    regionSum = 0
    for listObject in listDataBadm:
        if listObject['Товар'] == drug and listObject['Область'] == region:
            # print(listObject['Количество'])
            regionSum += int(listObject['Количество'])
    return regionSum


def countSalesKoneks(drug, region):
    if drug == 'Оберіть препарат' or region == 'Оберіть область':
        # print('Введіть дані')
        return 0
    regionSum = 0

    for listObject in listDataKoneks:
        if listObject['Назва продукту'] == drug and listObject['Область України'] == region:
            regionSum += float(listObject['Всього, уп'])
    return regionSum


def exportExcelFile():
    print(">exportExcelFile:")
    nums = []
    drugs = []
    badm = []
    optima = []
    koneks = []
    venta = []
    sum = []
    region = []
    i = 0
    # print(len(listOfRowObjects))
    for object in listOfRowObjects:
        if object.getDrugFromRaw() == 'Оберіть препарат':
            pass
        else:
            i += 1
            nums.append(i)
            drugs.append(object.getDrugFromRaw())
            badm.append(object.getSumBadmFromRaw())
            optima.append(object.getSumOptimaFromRaw())
            koneks.append(object.getSumBadmFromKoneks())
            venta.append(object.getSumVentaFromRaw())
            sum.append(object.getSumFromRaw())
            region.append(variableOfRegion.get())
    data = {
        '№': nums,
        'Препарат': drugs,
        'Бадм': badm,
        'Оптіма': optima,
        'Конекс': koneks,

        'Вента': venta,
        'Сума': sum,
        'Регіон': region}
    df = pd. DataFrame(data)
    df.to_excel(f'Розрахунок {variableOfRegion.get()}.xlsx', index=False)
    tkinter.messagebox.showinfo(
        "Експорт в Excel файл",  f'Файл "Розрахунок {variableOfRegion.get()}.xlsx" збережено')


def generateEmail():
    # print('generateMail:')

    with open('data\\emailTemplate\\mailTo.txt', 'r') as myfile:
        mailTo = myfile.read()

    with open('data\\emailTemplate\\mailCopy.txt', 'r') as myfile:
        mailCC = myfile.read()

    with open('data\\emailTemplate\\mailTemplate.html', 'rb') as myfile:
        data = myfile.read().decode("UTF-8")
    subject = f"Виконання плану на {strftime('%d.%m.%Y')}"

    import win32com.client as win32
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = mailTo
    mail.CC = mailCC
    mail.Subject = subject
    mail.HtmlBody = data
    mail.Display(True)


def mainGUISettings():
    # Main settings
    screen_width = root_tk.winfo_screenwidth()
    screen_height = root_tk.winfo_screenheight()
    sizeWindowW = 800
    sizeWindowH = 360

    marginWindowLeft = int(screen_width/2-sizeWindowW/2)
    marginWindowTOP = int(screen_height/3-sizeWindowH/2)

    root_tk.config(bg='#EBEBEB')
    root_tk.title('Підрахунок продажів')
    root_tk.geometry(
        f'{sizeWindowW}x{sizeWindowH}+{marginWindowLeft}+{marginWindowTOP}')
    root_tk.resizable(False, False)


def tableHead(window):
    # Create table head
    Label(window, text='№').place(x=0, y=0)
    Label(window, text='Назва препарата').place(x=100, y=0)
    Label(window, text='БаДМ').place(x=287, y=0)
    Label(window, text='Оптіма').place(x=330, y=0)
    Label(window, text='Конекс').place(x=383, y=0)
    Label(window, text='Вента').place(x=437, y=0)
    Label(window, text='Сума', width=5).place(x=485, y=0)
    Label(window, text='Область', width=25).place(x=580, y=0)


def countDataFromDrag(event):
    # print('>getDataFromDrag')
    for listTableObject in listOfRowObjects:
        listTableObject.getDataFromWindow()


def convertRegionKeysInVenta():
    # print("convertRegionKeysInVenta: START")
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
    # print(listDataVenta)
    # print("convertRegionKeysInVenta: FINISH")


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
        elif region['Область'] == "ЛУГАНСКАЯ":
            region['Область'] = "ЛУГАНСЬКА ОБЛАСТЬ"
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
        elif region['Область'] == 'ХЕРСОНСКАЯ':
            region['Область'] = 'ХЕРСОНСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ХМЕЛЬНИЦКАЯ':
            region['Область'] = 'ХМЕЛЬНИЦЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ЧЕРКАССКАЯ':
            region['Область'] = 'ЧЕРКАСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ЧЕРНИГОВСКАЯ':
            region['Область'] = 'ЧЕРНІГІВСЬКА ОБЛАСТЬ'
        elif region['Область'] == 'ЧЕРНОВИЦКАЯ':
            region['Область'] = 'ЧЕРНІВЕЦЬКА ОБЛАСТЬ'


def convertRegionInKoneks():
    for region in listDataKoneks:
        if region['Область України'] == 'Вінницька':
            region['Область України'] = 'ВІННИЦЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Волинська':
            region['Область України'] = 'ВОЛИНСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Дніпропетровська':
            region['Область України'] = 'ДНІПРОПЕТРОВСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Донецька':
            region['Область України'] = 'ДОНЕЦЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Житомирська':
            region['Область України'] = 'ЖИТОМИРСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Закарпатська':
            region['Область України'] = 'ЗАКАРПАТСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Івано-Франківська':
            region['Область України'] = 'ІВАНО-ФРАНКІВСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Запоріжська':
            region['Область України'] = 'ЗАПОРІЗЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Київська':
            region['Область України'] = 'КИЇВСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Кіровоградська':
            region['Область України'] = 'КІРОВОГРАДСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Львівська':
            region['Область України'] = 'ЛЬВІВСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Миколаївська':
            region['Область України'] = 'МИКОЛАЇВСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Одеська':
            region['Область України'] = 'ОДЕСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Полтавська':
            region['Область України'] = 'ПОЛТАВСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Рівненська':
            region['Область України'] = 'РІВНЕНСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Тернопільська':
            region['Область України'] = 'ТЕРНОПІЛЬСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Харківська':
            region['Область України'] = 'ХАРКІВСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Хмельницька':
            region['Область України'] = 'ХМЕЛЬНИЦЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Черкаська':
            region['Область України'] = 'ЧЕРКАСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Чернівецька':
            region['Область України'] = 'ЧЕРНІВЕЦЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Чернігівська':
            region['Область України'] = 'ЧЕРНІГІВСЬКА ОБЛАСТЬ'
        elif region['Область України'] == 'Київ':
            region['Область України'] = 'М.КИЇВ'
        elif region['Область України'] == 'Дніпро':
            region['Область України'] = 'ДНІПРОПЕТРОВСЬКА ОБЛАСТЬ'


def findCityInAddressKoneks(string):
    city = ''
    cityList = ['Київ', 'Дніпро', 'Киів']
    for i in range(len(cityList)):
        index = string.find(cityList[i])
        if index > 0:
            city = string[index:index+len(cityList[i])]
            #print('findCityInAddressKineks:', city)
    if city == 'Киів':
        city = 'Київ'
    return city


def fillRegionInKoneks():
    for listObject in listDataKoneks:
        if listObject['Область України'] == '':
            #print('Адреса клієнта-одержувача:', listObject['Адреса клієнта-одержувача'])
            #print('Область України:', listObject['Область України'])
            address = listObject['Адреса клієнта-одержувача']
            # print(address)
            listObject['Область України'] = findCityInAddressKoneks(address)
            #print('Область України:', listObject['Область України'])
            # print()

    # print('-------------------')
    for listObject in listDataKoneks:
        if listObject['Область України'] == '':
            print('ERROR IN STRING:', listObject['Адреса клієнта-одержувача'])


def setDefoltDrugsInObject(listOfRowObjects):

    listOfRowObjects[0].setDefoltDrugName('БРИЛИНТА ТАБ.П/О 90МГ#56(14X4)')
    listOfRowObjects[1].setDefoltDrugName('БРИЛИНТА ТАБ.П/О 60МГ#56(14X4)')
    listOfRowObjects[2].setDefoltDrugName('КРЕСТОР ТАБ.П/О 5МГ #28(14X2)')
    listOfRowObjects[3].setDefoltDrugName('КРЕСТОР ТАБ.П/О 10МГ #28(14X2)')
    listOfRowObjects[4].setDefoltDrugName('КРЕСТОР ТАБ.П/О 20МГ #28(14X2)')
    listOfRowObjects[5].setDefoltDrugName('КРЕСТОР ТАБ.П/О 40МГ #28(7X4)')
    listOfRowObjects[6].setDefoltDrugName('БЕТАЛОК ЗОК ТАБ. 100МГ #30')
    listOfRowObjects[7].setDefoltDrugName('БЕТАЛОК ЗОК ТАБ. 50МГ #30')
    listOfRowObjects[8].setDefoltDrugName('БЕТАЛОК ЗОК ТАБ. 25МГ #14')
    listOfRowObjects[9].setDefoltDrugName('БЕТАЛОК Д/ИН.1МГ/МЛ 5МЛ АМП.#5')
    listOfRowObjects[10].setDefoltDrugName('ФОРКСИГА ТАБ.П/О 10МГ#30(10X3)')
    listOfRowObjects[11].setDefoltDrugName('КСИГДУО ПРОЛ. ТАБ 10/1000МГ#28')
    listOfRowObjects[12].setDefoltDrugName('КСИГДУО ПРОЛОНГ ТАБ5/1000МГ#28')
    listOfRowObjects[13].setDefoltDrugName('ОНГЛИЗА ТАБ.П/О 2.5МГ #30(10X3)')
    listOfRowObjects[14].setDefoltDrugName('ОНГЛИЗА ТАБ.П/О 5МГ #30(10X3)')
    listOfRowObjects[15].setDefoltDrugName('КОМБОГЛИЗА XR ТАБ.5/1000МГ #28')


def openAuthorLink(url):
    webbrowser.open_new(url)


def insertWebLinks():

    authorLink = Label(
        root_tk, text='maksym.protsak@gmail.com V1.3', fg="blue", cursor="hand2")

    authorLink.place(x=600, y=21*16)
    authorLink.bind(
        "<Button-1>", lambda e: openAuthorLink("https://tangerine-youtiao-a51230.netlify.app/"))

    optimaLink = Label(
        root_tk, text='optimapharm.ua', fg="blue", cursor="hand2")
    optimaLink.place(x=600, y=21*13)
    optimaLink.bind(
        "<Button-1>", lambda e: openAuthorLink("https://optimapharm.ua/"))

    ventaLink = Label(root_tk, text='ventaltd.com.ua',
                      fg="blue", cursor="hand2")
    ventaLink.place(x=600, y=21*14)
    ventaLink.bind(
        "<Button-1>", lambda e: openAuthorLink(f"https://www.ventaltd.com.ua/"))

    ventaLink = Label(root_tk, text='badm.ua',
                      fg="blue", cursor="hand2")
    ventaLink.place(x=600, y=21*15)
    ventaLink.bind(
        "<Button-1>", lambda e: openAuthorLink(f"https://www.badm.ua/ua/"))


convertNameDrugInVenta()
convertRegionKeysInVenta()
convertNameDrugInBadm()
convertRegionInBadm()
convertNameDrugInKoneks()
fillRegionInKoneks()
convertRegionInKoneks()

####UI####
root_tk = tk.Tk()

mainGUISettings()
tableHead(root_tk)

# Creating comboRegion
variableOfRegion = StringVar()
comboRegion = ttk.Combobox(
    root_tk, textvariable=variableOfRegion, values=regionListOptima, width=30)
comboRegion['state'] = 'readonly'
comboRegion.set('Оберіть область')
comboRegion.place(x=580, y=20)

tableGUIFrame = Frame(root_tk)
tableGUIFrame.place(x=0, y=20)

for i in range(1, 17):
    tableRowObject = TableRow(tableGUIFrame, i)
    listOfRowObjects.append(tableRowObject)
    tableRowObject.createRow()
setDefoltDrugsInObject(listOfRowObjects)

# Create button for export in Excel file
buttonExportExcel = Button(
    root_tk, text='Експорт в Excel файл', command=exportExcelFile)
buttonExportExcel.place(x=600, y=200)

# Create button to generate email in OutLook
buttonGenerateEmail = Button(
    root_tk, text='Сформувати лист', command=generateEmail)
buttonGenerateEmail.place(x=600, y=230)

comboRegion.bind("<<ComboboxSelected>>", countDataFromDrag)

insertWebLinks()

tableData()

root_tk.mainloop()
