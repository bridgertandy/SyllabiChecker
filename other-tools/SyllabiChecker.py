import re,sys,os
from pathlib import Path
from tkinter import filedialog
from colored import bg,fg,attr

def name_checker(pattern,input): #takes a string and compares it to a string pattern
    try:
        if re.fullmatch(pattern,input):
            return False
        else:
            return True
    except Exception as e:
        print(e)
        return True

def schedule_finder(string): #takes a lowercase string and looks for "schedule"
    if "schedule" not in string:
        return False
    else:
        return True

#=================================================================================================================#
    
if __name__ == "__main__":
    file_name_pattern = r'^[A-Z]{2}\d{2}\s[A-Z]{3,4}\s\d{3}[A-Z]?\s\d{3}[A-Z]?\sSyllabus\.pdf$'
    #file name example         SU22      MATH        123P         001W      Syllabus .pdf

    folder_path = Path(filedialog.askdirectory(title="SELECT FOLDER TO SORT"))

    NAME_ERROR_PATH = folder_path/"NAME ERROR"
    SCHEDULE_PATH = folder_path/"SCHEDULES"
    FOUND_MATCH_PATH = folder_path/"FOUND MATCH"

    NAME_ERROR_PATH.mkdir(exist_ok=True)
    SCHEDULE_PATH.mkdir(exist_ok=True)
    

    schedule_count = 0
    name_count = 0 #TODO sometimes this counter will increment when it should not, has no major effects other than the counter being off by one when printed to the terminal

    for f in folder_path.iterdir():
        if f.is_dir():
            continue
        else:
            try:
                if schedule_finder(f.name.lower()):
                    f.rename(SCHEDULE_PATH/f.name)
                    schedule_count+= 1
                    continue
                if name_checker(file_name_pattern,f.name):
                    f.rename(NAME_ERROR_PATH/f.name)
                    name_count+= 1
                    continue
            except FileExistsError:
                continue
            
    print(bg('blue')+f"{schedule_count} Schedules found")
    print(f"{name_count} Name errors found"+attr('reset'))
                
    input(bg('green')+"All done! Press enter to quit or just close the window"+attr('reset'))
