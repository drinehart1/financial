# SRC: https://github.com/mariostoev/finviz

# CREATED: 30-OCT-2020
# LAST EDIT: 07-NOV-2020
# AUTHORS: DUANE RINEHART, MBA (duane.rinehart@gmail.com)

# IMPLEMENTS API CONNECTION TO STOCK INFORMATION SERVICE FINVIZ FOR PURPOSES OF STOCK SCREENING

# REQUIRES:
# - PYTHON 3.5+

# LOAD PREREQUISITES
try:
    import constants
    import finviz
    import os
    import time
    import concurrent.futures
    from datetime import datetime
    import pandas as pd
    import xlsxwriter

    # from finviz.screener import Screener
    #from urllib.request import urlopen
    #import json
    #import csv

except ImportError:
    print('ERROR LOADING PREREQUISITES')

# GLOBAL VARIABLES
stock_data = []

def lookup(stock):
    stock_data.append(finviz.get_stock(stock))

def main(output):

    stock_list = ['ET', 'MSFT', 'CFG', 'T', 'UVE', 'SJI', 'UBS', 'CSCO', 'ZION', 'AGI', 'XOM', 'OPK', 'RKT', 'XOM', 'OKE', 'LUMN', 'SPH', 'LNC', 'KR', 'ORA']
    stock_list = list(dict.fromkeys(stock_list)) #DEDUPLICATE

    with concurrent.futures.ThreadPoolExecutor() as executor: #AUTO-JOINS PROCESSES (WAITS FOR PROCESSES TO COMPLETE BEFORE CONTINUING WITH SCRIPT)
       for stock in stock_list:
          f1 = executor.submit(lookup(stock)) #SCHEDULES METHOD FOR EXECUTION AND RETURNS FUTURE OBJECT

    df = pd.DataFrame(data=stock_data)
    df.insert(loc=0, column='symbol', value=stock_list) #ADD SYMBOLS TO BEGINNING OF DATAFRAME

    # EXTRACT SPECIFIC COLUMNS FOR ANALYSIS
    df2 = df[['symbol', 'Price', 'Dividend', 'Dividend %']]

    #SAVE RAW STOCK DATA IN WORKSHEET 'raw_data'
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='raw_data')
    workbook = writer.bookworksheet = writer.sheets['raw_data']

    #SAVE ANALYSIS DATA IN WORKSHEET 'analysis'
    #for item in df:
    #    print(df['symbol'], df['Price'], df['Dividend'], df['Dividend %'])



    #df2 = pd.DataFrame(data=stock_list)
    #df2.columns = ["symbol"]
    df2.to_excel(writer, sheet_name='analysis', index=False)
    writer.save()

    #finviz.get_analyst_price_targets('AAPL')
    #filters = ['fa_div_pos', #POSITIVE DIVIDEND YIELD
    #           'fa_payoutratio_pos', #POSITIVE DIVIDENT PAYOUT RATIO
    #           'sh_avgvol_o200', #AVG VOLUME >200K
    #           'sh_curvol_o200'] #CUR_VOLUME >200K

    #stock_list = Screener(filters=filters, table='Performance', order='change')  # Get the performance table and sort it by price ascending

    # Export the screener results to .csv
    #stock_list.to_csv("stock.csv")

    # Create a SQLite database
    #stock_list.to_sqlite("stock.sqlite3")

    #for stock in stock_list[0:50]:  # Loop through 10th - 20th stocks
    #    print(stock['Ticker'], stock['Price'], stock['Change']) # Print symbol and price

    # Add more filters
    #stock_list.add(filters=['fa_div_high'])  # Show stocks with high dividend yield
    # or just stock_list(filters=['fa_div_high'])

    # Print the table into the console
    #print(stock_list)

if __name__ == '__main__':
    start = time.perf_counter()
    print('SCRIPT NAME:', os.path.basename(__file__))
    now = datetime.now() # datetime object containing current date and time
    dt_string = now.strftime("%m/%d/%Y %H:%M:%S")
    print("SCRIPT START TIMESTAMP:", dt_string)
    print('OUTPUT DESTINATION:', constants.outpath + constants.outfile)

    main(constants.outpath + constants.outfile)

    now = datetime.now()  # datetime object containing current date and time
    dt_string = now.strftime("%m-%d-%Y %H:%M:%S")
    print('SCRIPT END TIMESTAMP:', dt_string)
    finish = time.perf_counter()
    print(f'EXECUTION TIME: {round(finish-start,2)}s')