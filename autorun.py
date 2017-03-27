
# -*- coding: utf-8 -*-
"""
Lass die andere sich veaendern und bleib so wie du bist
Esse script converte um arquivo xls em csv.
Feito por: Saulo Henrique Dos Santos
"""
import xlrd
import csv
import pandas as pd
import datetime
import os
import ctypes

os.path.dirname(os.path.abspath('autorun.py'))
datapath1 = os.path.dirname(os.path.abspath('autorun.py')) 
datapath = os.path.join(datapath1,r'Produção2016MotoresFire.xlsm')
datapath2= ((datapath1+'\Resultados'))
#path1 = r'C:\Users\saulo.santos\Desktop'
wb = xlrd.open_workbook(datapath,'r')
for cont in range(0,(wb.nsheets - 15)):
    sh = wb.sheet_by_index(cont)   
    name = wb.sheet_names()
    planilha = name[cont] 
    your_csv_file = open('your_csv_file.csv', 'w')
    wr = csv.writer(your_csv_file,delimiter=',', quoting=csv.QUOTE_MINIMAL)
    for rownum in range(7,sh.nrows):
        now = datetime.datetime(*xlrd.xldate_as_tuple(sh.cell(rownum,1).value,wb.datemode))
        wr.writerow([sh.row_values(rownum,3,7),now.strftime('%Y.%m.%d')])
   
    your_csv_file.close()
    df = pd.read_csv(r'your_csv_file.csv', sep=',') 
    df.rename(columns={df.columns[0]: 'Date' }, inplace=True)
    df['Date']=df['Date'].str.lstrip('[ ]  "').str.rstrip('[ ]  "')
    df.to_csv(os.path.join(datapath2,(planilha+'.csv')), sep=',', index=False, header=None)
    os.remove('your_csv_file.csv')
ctypes.windll.user32.MessageBoxW(0, "Programa Executado com Sucesso", "XLS para CSV", 1)
