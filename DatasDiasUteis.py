#!/usr/bin/env python
# coding: utf-8
print('''by Fábio Guimarães\n© Copyright 2020. All Rights Reserved.''')
import urllib.request, re, openpyxl
from urllib.request import Request

def Luga(v1, v2):
    l = sheet.max_row + 1
    sheet['A' + str(l)] = v1
    sheet['B' + str(l)] = v2

criar_excel = openpyxl.Workbook()
sheet = criar_excel.active
sheet.title = 'A' + str(1)

dados = re.compile(r'title=\".*?,\s+?(\d+)\s+?([^\n]+),\s+?(\d+).*?;.*?;([^\n&]+)', re.I)

for ano in range(1990, 2031):
    url = 'https://www.dias-uteis.com/calendario_dias_uteis_'+ str(ano) +'.htm'
    url = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    print('https://www.dias-uteis.com/calendario_dias_uteis_'+ str(ano) +'.htm')
    AbrLink = urllib.request.urlopen(url).read().decode('utf-8')
    for BB in dados.findall(AbrLink):
        #print(BB)
        Luga(BB[0] + "/" + BB[1] + "/" + BB[2], BB[3])
    guia = + 1
    sheet = criar_excel.create_sheet('A' + str(guia), -1)

criar_excel.save(r'/content/drive/My Drive/Dados/DIASUTEIS.xlsx')
criar_excel = openpyxl.load_workbook(r'/content/drive/My Drive/Dados/DIASUTEIS.xlsx', keep_vba=True)
criar_excel.save(r'/content/drive/My Drive/Dados/DIASUTEIS.xlsm')
os.remove(r'/content/drive/My Drive/Dados/DIASUTEIS.xlsx')