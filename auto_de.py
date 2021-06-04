import subprocess
import sys
import time
from PIL.ImageOps import grayscale
import pyautogui
import csv
import os
import datetime

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

def screen_wait(ssname, required=True, search_time=10):
    result = pyautogui.locateOnScreen(ssname,minSearchTime=search_time, grayscale=True)
    if result == None and required==True:
        raise pyautogui.ImageNotFoundException
    return result

def open_xml(xml_file):
    pyautogui.press('alt')
    pyautogui.press('f')
    pyautogui.press('o')
    screen_wait(r".\screenshot\DE_open.png")

    pyautogui.write(xml_file)
    pyautogui.press('enter')
    return

def export_csv(xml_file, destination):
    # alt, e, e - CSV Export
    pyautogui.press('alt')
    pyautogui.press('e')
    pyautogui.press('e')
    screen_wait(r".\screenshot\DE_csv_export.png")

    export_file = build_fname(xml_file, destination)
    save_file(export_file)
    return

def build_fname(xml_file, destination, ext=''):
    fname = destination
    fname += os.path.splitext(os.path.basename(xml_file))[0]
    fname += datetime.datetime.now().strftime("_%a_%G-%m-%d").upper() # Today's date
    if ext!='': 
        fname += '.' + ext
    return fname

def export_stata(xml_file, destination):
    # alt, e, s - Stata Export
    pyautogui.press('alt')
    pyautogui.press('e')
    pyautogui.press('s')
    export_file = build_fname(xml_file, destination)
    while screen_wait(r".\screenshot\DE_stata_export.png",required=False):
        save_file(export_file)
    
def save_file(fname):
    # Write the name in the File Name field
    pyautogui.write(fname) 
    pyautogui.press('enter')
   
    #If the file already exists, confirm overwrite existing file
    if screen_wait(r".\screenshot\DE_save_as.png",search_time=0, required=False):
        pyautogui.press('y')

def main():
    subprocess.Popen(r"Z:\SLMS_CTU_IT_SOFTWARE\Uat\MACRO Data Exporter\MacroDataExporterV4.0 (slmscctusql01)\MacroDataExporter.exe")
    screen_wait(r".\screenshot\DE_login_button.png")  

    pyautogui.click(r".\screenshot\DE_trusted_connection.png")
    pyautogui.click(r".\screenshot\DE_login_button.png")

    with open('.\DE_xml_list.csv', mode='rt') as csv_file:
        csv_reader = csv.DictReader(csv_file, delimiter=',')
        weekday=datetime.datetime.now().strftime("%a").lower()
        for row in csv_reader:           
            if row['frequency'].strip() != '' and row['frequency'].lower().find(weekday) == -1:
                continue

            screen_wait(r".\screenshot\DE_file.png")

            open_xml(row['xml'])
            screen_wait(r".\screenshot\DE_global_options.png")

            if row['export_format'].lower()=='csv':
                export_csv(row['xml'], row['destination'])
                continue

            if row['export_format'].lower()=='stata':
                export_stata(row['xml'], row['destination'])
                continue
            
        pyautogui.hotkey('alt','f4') # Exit the application
    return



if __name__ == '__main__':
    main()
