import re,sys,os#,openpyxl
from pathlib import Path
from tkinter import filedialog

current_directory = os.getcwd()
master_list_path = Path(r"C:\Users\Bridger\OneDrive - College of Western Idaho\Coding Projects\SyllabiChecker\SyllabiChecker\00 MasterList-Syllabi Archival WS  102925 - Copy.xlsx")
#folder_path = Path(r"C:\Users\btandy\OneDrive - College of Western Idaho\Projects\SyllabiChecker\Thomas Bu - COMPLETED-MOVE TO ARCHIVE")
file_name_pattern = r'^[A-Z]{2}\d{2}\s[A-Z]{4}\s\d{3}[A-Z]?\s\d{3}[A-Z]?\sSyllabus\.pdf$'
#file name example         SU22      MATH        123P         001W      Syllabus .pdf

def name_checker(file_name):
    #print(file_name)
    #print(file_name_pattern)
    try:
        if re.fullmatch(file_name_pattern,file_name):
            return False
        else:
            return True
    except:
        return True

def schedule_finder(file_name):
    if "Schedule" in file_name:
        return True
    else:
        return False

def select_file():
    return filedialog.askopenfilename(defaultextension='.xlsx',initialdir=current_directory,title="SELECT THE MASTER LISE (EXCEL FILE)")
    
def select_directory():
    return filedialog.askdirectory(title="SELECT FOLDER TO SORT")

if __name__ == "__main__":
    folder_path = Path(select_directory())

    NAME_ERROR_PATH = folder_path/"NAME ERROR"
    SCHEDULE_PATH = folder_path/"SCHEDULES"
    MANUAL_REVIEW_PATH = folder_path/"MANUAL REVIEW"

    NAME_ERROR_PATH.mkdir(exist_ok=True)
    SCHEDULE_PATH.mkdir(exist_ok=True)
    MANUAL_REVIEW_PATH.mkdir(exist_ok=True)

    for f in folder_path.iterdir():
        if f.is_dir():
            continue
        else:
            try:
                if schedule_finder(f.name):
                    f.rename(SCHEDULE_PATH/f.name)
                    continue
                if name_checker(f.name):
                    f.rename(NAME_ERROR_PATH/f.name)
                    continue
            except FileExistsError:
                continue