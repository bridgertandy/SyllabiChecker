import re,os,numbers
from pathlib import Path
from tkinter import filedialog
import xlwings as xw

def select_directory():
    return filedialog.askdirectory(title="SELECT FOLDER TO SORT")

def cell_to_string(v):
    if v is None:
        return ''
    if isinstance(v,numbers.Real): #Checks if its a number
        return str(int(v))
    return str(v).strip()

if __name__=="__main__":
    folder_path = Path(select_directory())
    MANUAL_REVIEW_PATH = folder_path/"MANUAL REVIEW"
    # input("Please open the excel file AND the worksheet you wish to check. When finished, press enter to continue. ")
    
    print('test a')
    # master_list = xw.books.open(r"C:\Users\laxhe\OneDrive - College of Western Idaho\Projects\SyllabiChecker\00 MasterList-Syllabi Archival WS  102925 - Copy.xlsx")
    # worksheet = xw.sheets['2020']
    master_list = xw.books.active
    worksheet = xw.sheets.active
    print('test b')
    data_range = worksheet.used_range.value
    print('test c')
    #=================CREATE EXCEL DATA LIST===========
    print('test 1')
    if data_range is None:
        data = []
    else:
        if not isinstance(data_range,list):
            data_range = [[data_range]]
        elif data_range and not isinstance(data_range[0],(list,tuple)):
            data_range = [data_range]
    data = [[cell_to_string(v) for v in row] for row in data_range]
    
    #=================CREATE FILE LIST==================
    print('test 2')
    file_list = []
    for f in folder_path.iterdir():
        if f.is_file() and f.suffix.lower()=='.pdf':
            file_list.append(f.name)
            # print(f.name)


    #==================START SORT=======================
    print('test 3')
    for index, row in enumerate(data):
        for f in file_list:
            parts = f.split()
            if row[0] != parts[0]:
                print('wrong term')
                break
            match = (row[0]==parts[0])and(row[1]==parts[1])and(row[2]==parts[2])and(row[3]==parts[3])
            if match:
                print(f"Found match: {f}")
                worksheet.range(f'G{index+1}').value='X'
            continue