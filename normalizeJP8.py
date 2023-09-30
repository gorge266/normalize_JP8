import pandas as pd
import re
from openpyxl import Workbook

input_file_name = '目録第8版公開用リスト20230930.xlsx'
output_file_name = 'jp8out.xlsx'
target_sheet_name = '和名・学名リスト'

original_JP8 = pd.ExcelFile(input_file_name)
original_sheet_df = original_JP8.parse(target_sheet_name)

output_list = []

wb = Workbook()
ws = wb.active
id = 0
ws.append(["ID", "ORDER", "目", "FAMILY", "科", "GENUS", "属", "SPECIES", "種"])

for index, row in original_sheet_df.iterrows():
   # if re.match(r'.*[\u30A1-\u30FF]+',row[0]):
   if isinstance(row[0], str):
    m = re.match(r'.*Order\s+([A-Za-z]*)\s+([\u30A1-\u30FF]+)',row[0])
    if m:
        order_s = m.groups()[0]
        order_j = m.groups()[1]
    m = re.match(r'.*Family\s+([A-Za-z]*)\s+([\u30A1-\u30FF]+)',row[0])
    if m:
        family_s = m.groups()[0]
        family_j = m.groups()[1]
    m = re.match(r'^([A-Za-z]+)\s+([\u30A1-\u30FF]+)',row[0])
    if m:
        genus_s = m.groups()[0]
        genus_j = m.groups()[1]
    m = re.match(r'^\s*\d+\.\s*([A-Za-z ]+)\s+([\u30A1-\u30FF]+)',row[0])
    if m:
        spec_s = m.groups()[0]
        spec_j = m.groups()[1]
        id = id+1
        ws.append([id, order_s, order_j, family_s, family_j, genus_s, genus_j, spec_s, spec_j])


wb.save(output_file_name)
wb.close()