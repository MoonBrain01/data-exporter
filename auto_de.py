import subprocess
import sys
import time
import csv
import os
import datetime
from argparse import ArgumentParser

import pyautogui
from PIL.ImageOps import grayscale
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

import logging
logging.basicConfig(filename='.\_logs\DE'+datetime.datetime.now().strftime("_%G-%m-%d")+'.log', filemode='w',format='%(asctime)s - %(message)s',level=logging.INFO)





def screen_wait(ssname, required=True, search_time=10):
    result = pyautogui.locateOnScreen(ssname,minSearchTime=search_time, grayscale=True)
    if result == None and required==True:
        raise pyautogui.ImageNotFoundException
    return result


def open_xml(xml_file):
    # Open the DE XML file in the DE application
    pyautogui.press('alt')
    pyautogui.press('f')
    pyautogui.press('o')
    # Wait for the Open File dialog to appear
    screen_wait(r".\screenshot\DE_open.png")

    # "Type" the XML file path and name into the File Name field and "press" Enter
    pyautogui.write(xml_file)
    pyautogui.press('enter')
    return

def export_csv(xml_file, destination):
    # alt, e, e - CSV Export
    logging.info('Export to CSV')
    pyautogui.press('alt')
    pyautogui.press('e')
    pyautogui.press('e')
    screen_wait(r".\screenshot\DE_csv_export.png")

    export_file = build_fname(xml_file, destination)
    logging.info(f"Filename:{export_file}")
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
    # alt, e, s - Export STATA
    logging.info('Export to STATA')
    pyautogui.press('alt')
    pyautogui.press('e')
    pyautogui.press('s')
    export_file = build_fname(xml_file, destination)
    logging.info(f"Filename:{export_file}")
    while screen_wait(r".\screenshot\DE_stata_export.png", required=False):
        save_file(export_file)
    return

def save_file(fname):
    # Write the filenamename into the File Name field of the Save dialog
    pyautogui.write(fname) 
    pyautogui.press('enter')
   
    #If the file already exists, confirm overwrite existing file
    if screen_wait(r".\screenshot\DE_save_as.png",search_time=0, required=False):
        pyautogui.press('y')
    return

def main():
    try:
        logging.info('Launch Data Exporter')
        subprocess.Popen(r"Z:\SLMS_CTU_IT_SOFTWARE\Uat\MACRO Data Exporter\MacroDataExporterV4.0 (slmscctusql01)\MacroDataExporter.exe")
        logging.info('Login to Data Exporter')
        screen_wait(r".\screenshot\DE_login_button.png")  
        pyautogui.click(r".\screenshot\DE_trusted_connection.png")
        pyautogui.click(r".\screenshot\DE_login_button.png")
    except:
        logging.exception("Exception occurred")

    try:
        logging.info('Read list of Data Exporter XML files')
        with open('.\DE_xml_list.csv', mode='rt') as csv_file:
            csv_reader = csv.DictReader(csv_file, delimiter=',')
            weekday=datetime.datetime.now().strftime("%a").lower()
            
            for row in csv_reader:
                #Frequency is used to indicate if a export is to be run on the same day each week (WED)
                if row['frequency'].strip() != '' and row['frequency'].lower().find(weekday) == -1:
                    continue
                
                logging.info(f"Processing:{row['xml']}")
                screen_wait(r".\screenshot\DE_file.png", search_time=30)

                open_xml(row['xml'])
                screen_wait(r".\screenshot\DE_global_options.png")

                if row['export_format'].lower()=='csv':
                    export_csv(row['xml'], row['destination'])
                    continue

                if row['export_format'].lower()=='stata':
                    export_stata(row['xml'], row['destination'])
                    continue

            logging.info('Exit Data Exporter')
            pyautogui.hotkey('alt','f4') # Exit the application
    except:
        logging.exception("Exception occurred")
    return

if __name__ == '__main__':
    # Code for processing command-line arguments
    # Copied from https://srcco.de/posts/writing-python-command-line-scripts.html
    
    parser = ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group()
    group.add_argument('-v', '--verbose', help='Verbose (debug) logging', action='store_const', const=logging.DEBUG,
                       dest='loglevel')
    group.add_argument('-q', '--quiet', help='Silent mode, only log warnings', action='store_const',
                       const=logging.WARN, dest='loglevel')
    parser.add_argument('--dry-run', help='Noop, do not write anything', action='store_true')
    parser.add_argument('config', nargs='+', help='Configuation File')
    args = parser.parse_args()
    main(args)
