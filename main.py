#!/usr/bin/env python3

# CREATED: 29-OCT-2020
# LAST EDIT: 30-OCT-2020
# AUTHORS: DUANE RINEHART, MBA (duane.rinehart@gmail.com)

# REQUIRES:
# - PYTHON 3.5+

# IMPLEMENTS API CONNECTION TO STOCK INFORMATION SERVICE FOR PURPOSES OF STOCK SCREENING

# LOAD PREREQUISITES
#import requests
try:
    import os
    from urllib.request import urlopen
    import json
    import csv
except ImportError:
    print('ERROR LOADING PREREQUISITES')

#from pathlib import Path
#import subprocess
#import sys
#import time

# *** DEFINE CONSTANTS ***
# OUTPUT DATA DESTINATION (WINDOWS FORMAT)
outpath = "D://financial"

# SOURCE: https://financialmodelingprep.com/
api_key = '6049b76c7f5fc22e84c6691343c9975f'
api_base_url = 'https://financialmodelingprep.com/api/v3/profile/'

class conn:
    def __init__(self):
        # CREATE DESTINATION PATH IF NOT EXISTS
        if not os.path.exists(outpath):
            os.makedirs(outpath)

    def pull_data(self, symbol):
        # CREATE CONNECTION IF NOT EXISTS
        data_src = api_base_url + symbol + '?apikey=' + api_key
        response = urlopen(data_src)
        data = response.read().decode("utf-8")
        return json.loads(data)


def main():
    # RETRIEVE STOCK INTEREST LIST (IF DEFINED)
    stock_list = ['CX', 'XOM']

    connect = conn()
    stock_info = []
    for symbol in stock_list:
        data = connect.pull_data(symbol)
        stock_info.append(data)
        print(data)
        #stock_info[] = data[0]['symbol']

    data_file = open(outpath+'output.csv', 'w')
    csv_writer = csv.writer(data_file)
    count = 0

    for emp in employee_data:
        if count == 0:
            # Writing headers of CSV file
            header = emp.keys()
            csv_writer.writerow(header)
            count += 1

        # Writing data of CSV file
        csv_writer.writerow(emp.values())

    data_file.close()
    print(stock_info)

if __name__ == '__main__':
    status = 'NO ERRORS'
    print('SCRIPT START')
    main()
    print(f'SCRIPT END - STATUS:, {status}')