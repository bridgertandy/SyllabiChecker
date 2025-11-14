import re,numbers,time,sys
from pathlib import Path
from tkinter import filedialog
from colored import bg,fg,attr
import xlwings as xw

def select_directory(dialog_title):
    return filedialog.askdirectory(title=dialog_title)

def cell_to_string(v):
    if v is None:
        return ''
    # treat booleans separately because bool is a subclass of int
    if isinstance(v, bool):
        return str(v)
    # try to coerce numeric-like values to float (handles int, float, numpy types, numeric strings)
    try:
        fv = float(v)
    except Exception:
        # not numeric: normalize whitespace and invisible characters
        s = str(v)
        s = s.replace('\u00A0', ' ').replace('\u200B', '').replace('\uFEFF', '')
        return re.sub(r'\s+', ' ', s).strip()
    else:
        # drop trailing .0 for whole numbers, otherwise keep float representation
        return str(int(fv)) if fv.is_integer() else str(fv)

if __name__=="__main__":
    folder_path = Path(select_directory("SELECT FOLDER TO SORT"))
    FOUND_MATCH_PATH = folder_path/"FOUND MATCH"

    if not FOUND_MATCH_PATH.is_dir():
        print(bg('red')+'The "FOUND MATCH" folder could not be found. Please make sure you run SyllabiChecker.py before using this script.')
        print('This program will automatically close in 30 seconds'+attr('reset'))
        time.sleep(30)
        sys.exit()

    while (True):
        print(bg('blue')+'Make sure the excel file AND the correct workbook are open AND that it is formatted correctly')   
        input('When ready, press enter to continue. You may have to click inside the window'+attr('reset'))
        try:
            master_list = xw.books.active
            worksheet = xw.sheets.active
            data_range = worksheet.used_range.value
            break
        except:
            print(bg('red')+'ERROR, TRY AGAIN'+attr('reset'))
            continue
    
    #=================CREATE EXCEL DATA LIST===========
    if data_range is None:
        data = []
    else:
        if not isinstance(data_range,list):
            data_range = [[data_range]]
        elif data_range and not isinstance(data_range[0],(list,tuple)):
            data_range = [data_range]
    data = [[cell_to_string(v).lstrip('0') for v in row] 
                                for row in data_range]
    
    #=================CREATE FILE LIST==================
    file_list = []
    for f in folder_path.iterdir():
        if f.is_file() and f.suffix.lower()=='.pdf':
            file_list.append(f)

    #==================START SORT=======================
    match_counter = 0
    for index, row in enumerate(data):
        for f in file_list:
            parts = f.name.split()
            if row[0] != parts[0]:
                # print('wrong term')
                break
            match = (row[0]==parts[0])and(row[1]==parts[1])and(row[2]==parts[2].lstrip('0'))and(row[3]==parts[3].lstrip('0'))
            if match:
                match_counter+=1
                try:
                    print(f"Found match: {f.name}")
                    worksheet.range(f'G{index+1}').value='X'
                    f.rename(FOUND_MATCH_PATH/f.name)
                except Exception as e:
                    print(f'\nError: {e}\tFile: {f.name}')
                    print(bg('red')+f'Check the master list for proper formatting, duplicate rows (Different instructors is OK), etc.'+attr('reset')+'\n')
            continue

    print(bg('blue')+f'{match_counter} matches found!')
    input('All done! Press enter to exit or just close the window'+attr('reset'))