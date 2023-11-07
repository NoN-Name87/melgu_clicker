import pandas as pd
from openpyxl import Workbook
from openpyxl import load_workbook
import os
import re

def parse_person(file_path):
    print("ALERT")
    file_format = re.search(r'\.[a-z]+', file_path).group()
    if(file_format == '.xlsx'):
        result = pd.read_excel(file_path)
    else:
        result = pd.read_csv(file_path)
    result_list = []
    for row in result.iloc:
        result_list.append(row.to_dict())
    return result_list

def add_row(person):
    path = os.path.join('test', 'dump.xlsx')
    row = pd.DataFrame({"ID":[person["ID"]], "Имя":[person["Имя"]], "Фамилия":[person["Фамилия"]], "Отчество":[person["Отчество"]], "Почта":[person["Почта"]]})
    df_excel = pd.read_excel(path)
    result = pd.concat([df_excel, row], ignore_index=True)
    result.to_excel(path, index=False)
        