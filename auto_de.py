import subprocess
import time
import pyautogui
import csv
import os
import datetime


pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

def screen_wait(ssname, loop_limit=10):
    loop_count = 0
    result = None
    while (result==None and loop_count < loop_limit ) :
        result = pyautogui.locateOnScreen(ssname)
        loop_count += 1
    if result==None:
        print(f"{ssname} not found")
    return result

def open_xml(xml_file):
    pyautogui.press('alt')
    pyautogui.press('f')
    pyautogui.press('o')
    if screen_wait(r".\screenshot\DE_open.png")==None:
        return False
    pyautogui.write(xml_file)
    pyautogui.press('enter')
    return True

def export_csv(xml_file, destination):
    pyautogui.press('alt')
    pyautogui.press('e')
    pyautogui.press('e')
    if screen_wait(r".\screenshot\DE_csv_export.png")==None:
        return False
    export_file = destination
    export_file += os.path.splitext(os.path.basename(xml_file))[0]
    today = datetime.datetime.now().strftime("_%a_%G-%m-%d").upper()
    export_file += today + '.csv'
    pyautogui.write(export_file)
    pyautogui.press('enter')
    return True

if __name__ == '__main__':
    subprocess.Popen(r"Z:\SLMS_CTU_IT_SOFTWARE\Uat\MACRO Data Exporter\MacroDataExporterV4.0 (slmscctusql01)\MacroDataExporter.exe")
    if screen_wait(r".\screenshot\DE_login_button.png")==None:
        exit

    pyautogui.click(r".\screenshot\DE_trusted_connection.png")
    pyautogui.click(r".\screenshot\DE_login_button.png")

    with open('.\DE_xml_list.csv', mode='rt') as csv_file:
        csv_reader = csv.DictReader(csv_file, delimiter=',')
        for row in csv_reader:
            print(row.keys())
            print(row['trial'],row['xml'], row['frequency'], row['export_format'], row['destination'])
            
            weekday=datetime.datetime.now().strftime("%a").lower()

            if row['frequency'].strip() != '' and row['frequency'].lower().find(weekday) == -1:
                continue

            if screen_wait(r".\screenshot\DE_file.png")==None:
                exit
            open_xml(row['xml'])
            if screen_wait(r".\screenshot\DE_global_options.png")==None:
                exit
            if row['export_format'].lower()=='csv':
                export_csv(row['xml'], row['destination'])
            pyautogui.hotkey('alt','f4')

    # alt, e, s - Stata Export
    # alt, e, e - CSV Export

