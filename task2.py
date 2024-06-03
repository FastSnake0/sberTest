import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from bs4 import BeautifulSoup
import requests


exel_path = 'test_input.xlsx' 

excel_data = pd.read_excel(exel_path, skiprows=4) # читаем данные из эксельки

data_head = pd.read_excel(exel_path, header=None, nrows=3, usecols=[0, 1]) 
file_name = data_head[1].loc[data_head.index[2]] 

# Делаем дерево

root=ET.Element('CERTDATA') 

filename = ET.SubElement(root, 'FILENAME')
filename.text = file_name 

m = ET.SubElement(root, 'ENVELOPE')

def format_value(value): # Округляем до сотых, убираем лишние нули
    temp = round(value,2)
    return f"{temp:.2f}".rstrip('00').rstrip('0').rstrip('.')

def find_in_table(table, lcode): # Ищем в табличке 
    for row in table.find_all('tr'):
            cells = row.find_all('td')
            if len(cells) >= 2:
                if cells[1].get_text(strip=True) == lcode:
                    value = cells[4].get_text(strip=True)
                    return float(value.replace(',', '.'))
                    break # Этот брэйк здесь на память

def get_usd(date): # Парсинг по дате
    url = 'https://www.cbr.ru/currency_base/daily/?UniDbQuery.Posted=True&UniDbQuery.To=' + date.strftime('%d.%m.%Y')
    response = requests.get(url)
    html_content = response.content
    soup = BeautifulSoup(html_content, 'html.parser')
    table = soup.find('table', {'class': 'data'})
    rate = find_in_table(table, 'USD')
    return rate

for index, row in excel_data.iterrows():
    ecert = ET.SubElement(m, 'ECERT') # Формируем ECERT 
    ET.SubElement(ecert, 'CERTNO').text = row['Ref no']
    ET.SubElement(ecert, 'CERTDATE').text = row['Issuance Date'].strftime('%Y-%m-%d') # Форматируем дату
    ET.SubElement(ecert, 'STATUS').text = row['Status']

    iec=str(row['IE Code']) # Весёлый формат
    while(len(iec) < 10):
        iec = "0" + iec 
    
    ET.SubElement(ecert, 'IEC').text = iec
    ET.SubElement(ecert, 'EXPNAME').text = '"' + row['Client'] + '"'
    ET.SubElement(ecert, 'BILLID').text = row['Bill Ref no']
    ET.SubElement(ecert, 'SDATE').text = row['SB Date'].strftime('%Y-%m-%d') # Форматируем дату (ещё раз)
    ET.SubElement(ecert, 'SCC').text = row['SB Currency']
    ET.SubElement(ecert, 'SVALUE').text = str(format_value(row['SB Amount']))

    rate = get_usd(row['SB Date']) # Получаем курс долларов

    ET.SubElement(ecert, 'SVALUEUSD').text = str(format_value(row['SB Amount']/rate))

tree = ET.ElementTree(root)
tree.write('output_xml_v2.xml', encoding='utf-8') # запись в файл