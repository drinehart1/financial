#!/usr/bin/env python3

# CREATED: 29-OCT-2020
# LAST EDIT: 03-NOV-2020
# AUTHORS: DUANE RINEHART, MBA (duane.rinehart@gmail.com)

# REQUIRES:
# - PYTHON 3.5+

# IMPLEMENTS API CONNECTION TO STOCK INFORMATION SERVICE FOR PURPOSES OF STOCK SCREENING

# LOAD PREREQUISITES
try:
    import constants
    import os
    from urllib.request import urlopen
    import json
    #import csv
    import xlsxwriter
except ImportError:
    print('ERROR LOADING PREREQUISITES')

#from pathlib import Path
#import subprocess
#import sys
#import time

api_base_url = 'https://financialmodelingprep.com/api/v3/profile/'

class conn:
    def __init__(self):
        # CREATE DESTINATION PATH IF NOT EXISTS
        if not os.path.exists(outpath):
            os.makedirs(outpath)

    def pull_data(self, symbol):
        # CREATE CONNECTION IF NOT EXISTS
        data_src = api_base_url + symbol + '?apikey=' + constants.api_key
        response = urlopen(data_src)
        data = response.read().decode("utf-8")
        return json.loads(data)


def main():
    # RETRIEVE STOCK INTEREST LIST (IF DEFINED)
    stock_list = ['ET', 'MSFT', 'CFG', 'T', 'UVE', 'SJI', 'UBS', 'CSCO', 'ZNGA', 'AGI', 'XOM', 'OPK', 'RKT', 'XOM', 'OKE', 'LUMN', 'SPH', 'LNC', 'KR', 'ORA']
    stock_list.sort()
    #print(stock_list)
    connect = conn()

    # OUTPUT RESULTS TO SPREADSHEET
    workbook = xlsxwriter.Workbook(outpath + 'output.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0
    col = 0

    for symbol in stock_list:
        stock_info = connect.pull_data(symbol)
        print(json.dumps(stock_info))
        try:
            for key in stock_info[0].keys():
            #header = stock_info[0].keys()

                row += 1
            #for key in header:
                print(key)
                worksheet.write(row, col, json.dumps(stock_info))
                # for item in stock_info[0][header]:
                #     worksheet.write(row, col + 1, json.dumps(item))
                #     row += 1
        except:
            print('LOOKUP FAILED FOR: ', symbol)

    workbook.close()

    # data_file = open(outpath + 'output.csv', 'w')
    # csv_writer = csv.writer(data_file, lineterminator = '\n')
    # count = 0

    # for symbol in stock_list:
    #     stock_info = connect.pull_data(symbol)
    #
    #     try:
    #         if count == 0:
    #             # Writing headers of CSV file
    #             header = stock_info[0].keys()
    #             csv_writer.writerow(header)
    #             count += 1
    #
    #         # Writing data to CSV file
    #         csv_writer.writerow(stock_info[0].values())
    #     except:
    #         print('LOOKUP FAILED FOR: ', symbol)
    #
    # data_file.close()

if __name__ == '__main__':
    print('SCRIPT START')
    print('OUTPUT DESTINATION: ', outpath + 'output.csv')
    main()
    print('SCRIPT END')