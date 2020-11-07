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
    from datetime import datetime
    import pandas as pd
    import xlsxwriter

    # from finviz.screener import Screener
    #from urllib.request import urlopen
    #import json
    #import csv

except ImportError:
    print('ERROR LOADING PREREQUISITES')


def main():
    stock_data = []

    stock_list = ['ET', 'MSFT', 'CFG', 'T', 'UVE', 'SJI', 'UBS', 'CSCO', 'ZION', 'AGI', 'XOM', 'OPK', 'RKT', 'XOM', 'OKE', 'LUMN', 'SPH', 'LNC', 'KR', 'ORA']
    stock_list = list(dict.fromkeys(stock_list)) #DEDUPLICATE
    stock_list.sort()

    for stock in stock_list:
        stock_data.append(finviz.get_stock(stock))

    df = pd.DataFrame(data=stock_data)
    df.insert(loc=0, column='symbol', value=stock_list) #ADD SYMBOLS TO BEGINNING OF DATAFRAME

    #SAVE RAW STOCK DATA IN WORKSHEET 'raw_stocks'
    writer = pd.ExcelWriter(constants.outpath + 'stock_data_output.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='raw_stocks')
    workbook = writer.bookworksheet = writer.sheets['raw_stocks']

    #SAVE ANALYSIS DATA IN WORKSHEET 'analysis'
    df2 = pd.DataFrame(data=stock_list)
    df2.columns = ["symbol"]
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
    print('SCRIPT START:', os.path.basename(__file__))
    now = datetime.now() # datetime object containing current date and time
    dt_string = now.strftime("%m/%d/%Y %H:%M:%S")
    print("SCRIPT START TIMESTAMP:", dt_string)
    print('OUTPUT DESTINATION:', constants.outpath + 'output.csv')

    main()

    now = datetime.now()  # datetime object containing current date and time
    dt_string = now.strftime("%m-%d-%Y %H:%M:%S")
    print('SCRIPT END TIMESTAMP:', dt_string)
    finish = time.perf_counter()
    print(f'EXECUTION TIME: {round(finish-start,2)}s')