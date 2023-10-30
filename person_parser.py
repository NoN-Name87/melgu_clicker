import pandas as pd
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