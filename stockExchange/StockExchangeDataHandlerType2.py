'''
Created on 09-Apr-2018
@author: Rohit Kumar
Purpose: This class will download the data files from website
         and extract the data from the downloaded file and
         convert the time series data in the required format
         and generates the output files.
'''
from config import *
from requests import get
from xlrd import open_workbook
import csv

class StockExchangeDataHandlerType2:
    def __init__(self):
        """
        This dictionary will used to get month number for given month name
        """
        self.month = {
            'JAN':1, 'FEB':2, 'MAR':3, 'APR':4, 'MAY':5, 'JUN':6,
             'JUL':7, 'AUG':8, 'SEP':9, 'OCT':10, 'NOV':11, 'DEC':12
             }
    
    def saveDownloadedFile(self, response, fname):
        """
        This function will save the downloaded response.
        Output file path is given in config file.
        closing the file object to save the downloaded response.
        """
        filename = OUTPUT_FILE_PATH + fname
        with open(filename, "wb") as file:
            file.write(response.content)
        file.close()
    
    def downloadDataFiles(self):
        """
        This method will start the process to download the required file.
        """
        response = get(INPUT_URL + TYPE2_INPUT_FILE)
        self.saveDownloadedFile(response, TYPE2_INPUT_FILE)
        
    def getMonthNumber(self, month):
        """
        This method will return the month number againest month name
        """
        return self.month[month]
    
    def getYear(self, yearStr):
        """
        This function will extract the year from the 1st column and return
        """
        temp_year = 0
        try:
            if (str(int(float(yearStr))).strip()).isdigit():
                temp_year = int(float(yearStr))
        except:
            if (str(yearStr).strip()).isdigit():
                temp_year = int(yearStr)
        return temp_year
    
    def getMonth(self, monthStr):
        """
        This method will extract the month from 2nd column and return
        """
        temp_month = 0
        try:
            if not (str(int(float(monthStr))).strip()).isdigit():
                temp_month_str = (str(int(float(monthStr))).strip()[:3]).upper()
                if temp_month_str in self.month:
                    temp_month = self.getMonthNumber(temp_month_str)
        except:
            if not (str(monthStr).strip()).isdigit():
                temp_month_str = (str(monthStr).strip()[:3]).upper()
                if temp_month_str in self.month:
                    temp_month = self.getMonthNumber(temp_month_str)
        return temp_month
    
    def generateOutputFile(self):
        """
        This function will create the output file in csv format and save it
        at the path which is mention in the config file
        """
        headers = ['Date', 'BCB_FX_Position']
        wb = open_workbook(OUTPUT_FILE_PATH + TYPE2_INPUT_FILE)
        sheet = wb.sheet_by_index(0)
        """
        Opening file to write
        """
        with open(OUTPUT_FILE_PATH + TYPE2_OUTPUT_FILE, "w") as file:
            writer = csv.writer(file, delimiter = ",")
            writer.writerow(headers)
            last_year, last_month, last_day = 0, 0, 1
            temp_date_list = LAST_UPDATED_DATA_TYPE2.split("/")
            last_update_year = int(temp_date_list[2])
            last_update_month = int(temp_date_list[0])
            last_update_day = int(temp_date_list[1])
            """
            Starting iteration over downloaded sheet
            """
            for row_idx in range(13, sheet.nrows):
                #print(sheet.row(row_idx)[0].value, sheet.row(row_idx)[1].value)
                if not str(sheet.row(row_idx)[1].value).strip() or sheet.row(row_idx)[1].value is None:
                    continue
                """
                Extracting Year and updating last year variable with the larger value
                """
                temp_year = self.getYear(str(sheet.row(row_idx)[0].value).strip())
                if temp_year >= last_year:
                    last_year = temp_year
                    last_month = 0
                """
                Extracting Month and updating last month variable with the larger value
                """
                temp_month = self.getMonth(str(sheet.row(row_idx)[1].value).strip())
                if temp_month >= last_month:
                    last_month = temp_month
                temp_date = str(last_month)+"/"+str(last_day)+"/"+str(last_year)
                """
                Starting Comparision whether particular row has to be 
                written in the output file
                """
                if last_year > last_update_year:
                    row = [temp_date] + [cell.value for cell in sheet.row(row_idx)[2:]]
                    writer.writerow(row)
                elif last_year == last_update_year:
                    if last_month >= last_update_month:
                        row = [temp_date] + [cell.value for cell in sheet.row(row_idx)[2:]]
                        writer.writerow(row)
                """
                Written row data in output file.
                Output file path is mentioned in config file.
                """
        
sedht2 = StockExchangeDataHandlerType2()
sedht2.downloadDataFiles()
sedht2.generateOutputFile()