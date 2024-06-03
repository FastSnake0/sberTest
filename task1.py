import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom

exel_path = 'test_input.xlsx'


excel_data = pd.read_excel(exel_path, skiprows=4) # читаем данные из эксельки

data_head = pd.read_excel(exel_path, header=None, nrows=3, usecols=[0, 1])
file_name = data_head[1].loc[data_head.index[2]]

# Делаем дерево

root=ET.Element('CERTDATA')
filename = ET.SubElement(root, 'FILENAME')
filename.text = file_name
m = ET.SubElement(root, 'ENVELOPE')

def format_value(value):
    return f"{value:.2f}".rstrip('00').rstrip('0').rstrip('.')

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
    ET.SubElement(ecert, 'SDATE').text = row['SB Date'].strftime('%Y-%m-%d') # Форматируем дату
    ET.SubElement(ecert, 'SCC').text = row['SB Currency']
    ET.SubElement(ecert, 'SVALUE').text = str(format_value(row['SB Amount']))

tree = ET.ElementTree(root)
tree.write('output_xml.xml', encoding='utf-8') # Запись в файл
