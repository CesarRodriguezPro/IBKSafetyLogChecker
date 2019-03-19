# !/usr/bin/env python3
# -*- coding: utf8 -*-

import webbrowser, platform, shutil, subprocess, os, csv
import openpyxl
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
from gsheets import Sheets
from openpyxl.styles import Alignment, Color, PatternFill, Font


today = datetime.today().strftime('%Y-%m-%d %H:%M')
today_year = datetime.today().strftime('%Y')
current_os = platform.system()
global locations

###################################### settings app ########################################################

with open("Google_url.txt", "r") as file_to_open:
    GOOGLE_URL = file_to_open.read() 
with open('TimeStation_Key.txt', "r") as file_open:
    API_KEY = file_open.read()

CURRENT_EMPLOYEES_DATA = 'CURRENT_EMPLOYEES_DATA.txt'
G_BEFORE_PD = "G_BEFORE_PD.csv"
sheets = Sheets.from_files('credentials.json')
CODE = 37
GOOGLE_DATA = 'google_data.xlsx'
############################################################################################################


################################## MultiL OS Support #######################################################
DIR_PATH = os.path.abspath(os.path.dirname(__file__))
GOOGLE_PATH = os.path.join(DIR_PATH, GOOGLE_DATA)
clean_screen = os.system('cls') if current_os == "Windows" else os.system('clear')

def open_file(path):
    if platform.system() == "Windows":
        os.startfile(path)
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])
############################################################################################################


class GettingDataForReport:

    def __init__(self):
        self.google_sheet_process()
        self.get_data(CURRENT_EMPLOYEES_DATA)

    def google_sheet_process(self):

        s = sheets.get(GOOGLE_URL)
        s.sheets[0].to_csv(G_BEFORE_PD, encoding='utf-8', dialect='excel')
        google_data = pd.read_csv(G_BEFORE_PD)
        google_data["Name"] = google_data['Name'].apply(lambda x: x.lower())
        google_data.to_excel(GOOGLE_DATA)

    def time_return(self, x):

        def convert_time(x):
            v = datetime.strptime(x, '%Y-%m-%d')
            return v

        try:
            years_ago = x - relativedelta(years=-5)
            return years_ago.strftime('%Y-%m-%d')
        except:
            conve = convert_time(x)
            years_ago = conve - relativedelta(years=-5)
            return years_ago.strftime('%Y-%m-%d')

    def get_data(self, out_file, code=37):
        #  this is preconfigure for current employee status  for the whole company
        #  by we can use any other code that not require a date input.

        try:
            ship_api_url = "https://api.mytimestation.com/v0.1/reports/?api_key={}&id={}&exportformat=csv".format(API_KEY, CODE)
            url_responds = requests.get(ship_api_url)
            CURRENT_EMPLOYEES_DATA = url_responds.text
            save_text = open(out_file, 'w', encoding='utf8')
            save_text.write(CURRENT_EMPLOYEES_DATA)
            save_text.close()
            print('--connection made')
        except:
            print('there was a problem downloading the information from the server')
            print('please check the internet connection.')
            input('')

    def orginazing_timesation_data(self):

        """ this get the information from timestation current and give back 
        a list of names with devices"""

        item_to_return = {}
        with open(CURRENT_EMPLOYEES_DATA, 'r', newline="") as file_to_open:
            reader = csv.DictReader(file_to_open, delimiter=',')
            for row in reader:
                if row['Status'] == 'In' and locations in row['Current Department']:
                    item_to_return.setdefault(row["Device"], []).append([row['Name'].lower(), row['Current Department']])

        print('---sort finish')
        return item_to_return

    def read_info_google(self):

        file_excel = openpyxl.load_workbook(GOOGLE_PATH)
        sheet_names = file_excel.sheetnames
        sheet_name = sheet_names[0]
        active_sheet = file_excel[sheet_name]
        rows_number = active_sheet.max_row

        number = 3
        db = {}
        for x in range(rows_number-2):
            name = active_sheet['c{}'.format(number)].value
            license_number = active_sheet['d{}'.format(number)].value
            issued_date = active_sheet['e{}'.format(number)].value
            expiration_date = active_sheet['f{}'.format(number)].value
            db[name.lower()] = [license_number, issued_date, expiration_date]
            number += 1

        return db

    def run(self):

        dict_employees = self.orginazing_timesation_data()
        db = self.read_info_google()

        dict_names = [items for items in dict_employees.values()]

        dict_location = {}
        for items in dict_names:
            for item in items:
                dict_location[item[0].lower()] = item[1]


        dict_full_data = {}   # will be fill in processing data
        for device, names in dict_employees.items():
            for name in names:
                if name[0] not in db.keys():
                    no_osha = {}
                    no_osha[name[0]] = ["none", "None", "None",name[1]]
                    dict_full_data.setdefault(device, []).append(no_osha)  # this was declare in the settings

        for g_name, g_info in db.items():
            for device, list_names in dict_employees.items():
                if g_name in [name[0] for name in list_names]:  
                    g_dict = {}
                    g_info.append(dict_location[g_name])
                    g_dict[g_name] = g_info
                    dict_full_data.setdefault(device, []).append(g_dict)
        
        return dict_full_data


class CreatedReport:

    def __init__(self):

        self.active = GettingDataForReport()
        self.dict_full_data = self.active.run()

    def display_info(self):
        ''' is created to see the infomation recolected without creating the sheets'''

        number = 1
        for device_name, list_items in self.dict_full_data.items():
            for dict_with_names in list_items:
                for names, values in dict_with_names.items():
                    print(f"{number:5} - {device_name:26} -- {names:25} -- {values}")
                    number += 1

    def styling_cells(self, cell):
        ''' this will created the black background on the text and white font. also you can modified
        the text zice and font type in here'''

        cell.fill = PatternFill(patternType='solid', fill_type='solid', fgColor=Color('000000'))
        cell.font = Font(color=Color('ffffff'))
        cell.alignment = Alignment(horizontal="center")

    def write_heather(self, ws, device_name):
        ''' writes the heather of the page and format the heather '''

        ws.merge_cells('A1:D1')
        ws.merge_cells('A2:D2')
        ws.merge_cells('A3:D3')

        ws["A1"] = "IBK CONSTRUCTIONG GROUP"
        ws["A2"] = "30 Hours Licenses"
        ws['A3'] = f"Device : {device_name}"

        ws["A5"] = "Name"
        ws["B5"] = "Date Issued"
        ws["C5"] = "Expiration Date"
        ws["D5"] = "Location"

        self.styling_cells(ws["A1"])

        self.styling_cells(ws['A5'])
        self.styling_cells(ws['B5'])
        self.styling_cells(ws['C5'])
        self.styling_cells(ws['D5'])

    def run(self):

        for device_name, list_items in self.dict_full_data.items():

            None if os.path.isdir(locations) else os.mkdir(locations) #  check for creating folders
            wb = openpyxl.Workbook()
            ws = wb.active
            file_out_name = os.path.join(f"{locations}", f'{device_name}.xlsx')
            name_for_out = "{}".format(os.path.join(DIR_PATH, file_out_name))

            self.write_heather(ws=ws, device_name=device_name)

            number = 6
            for dict_with_names in list_items:
                for names, values in dict_with_names.items():
                    ws["A{}".format(number)] = names.title()
                    ws["B{}".format(number)] = values[1]
                    ws["C{}".format(number)] = values[2]
                    ws["D{}".format(number)] = values[3]

                    print(f"------- write > {names} {values[1]} {values[2]}")
                    number += 1

            ws.column_dimensions['A'].width = 22
            ws.column_dimensions['B'].width = 12
            ws.column_dimensions['C'].width = 14
            ws.column_dimensions['D'].width = 22


            wb.save(name_for_out)
            wb.close()
            print(f'--> form device created for {device_name}\n')

        open_file(locations)


if __name__ == '__main__':

    print("____________ Welcome ____________")
    locations = input("Please introduce your location \n\n --->  ")

    active = CreatedReport()
    active.run()