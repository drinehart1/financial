# SRC: https://github.com/mariostoev/finviz

# CREATED: 18-NOV-2020
# LAST EDIT: 25-NOV-2020
# AUTHORS: DUANE RINEHART, MBA (duane.rinehart@gmail.com)

# READ LIST OF VETTED STOCKS
# TECHNICAL ANALYSIS, OPTIMIZE FOR BEST OPPORTUNITY (IMPLEMENTS FINVIZ FOR STOCK SCREENING)

# REQUIRES:
# - PYTHON 3.5+

# LOAD PREREQUISITES
try:
    import sys
    import constants
    import finviz
    import os
    import time
    import concurrent.futures
    from datetime import datetime
    import pandas as pd
    import xlsxwriter

    # from finviz.screener import Screener
    # from urllib.request import urlopen
    # import json
    # import csv

except ImportError:
    print('ERROR LOADING PREREQUISITES')
    exit()

# GLOBAL VARIABLES
stock_data = []
analyst_data = []


def load_stock_list(worksheet):
    infile = os.path.join(constants.outpath, 'blotter.xlsx')
    return pd.read_excel(infile, sheet_name=worksheet)


def lookup(stock):
    stock_data.append(finviz.get_stock(stock))

    analyst_raw = {}
    analyst_raw['symbol'] = stock
    analyst_raw['info'] = finviz.get_analyst_price_targets(stock)
    analyst_data.append(analyst_raw)
    # analyst_data.update(finviz.get_analyst_price_targets(stock)[0])
    # print(str(analyst_data))


def main(output):
    input_worksheet = 'prospects'
    xl_df = load_stock_list(input_worksheet)
    stock_list = xl_df['SYMBOL'].values.tolist()
    stock_list = list(dict.fromkeys(stock_list))  # DEDUPLICATE

    with concurrent.futures.ThreadPoolExecutor() as executor:  # AUTO-JOINS PROCESSES (WAITS FOR PROCESSES TO COMPLETE BEFORE CONTINUING WITH SCRIPT)
        for stock in stock_list:
            f1 = executor.submit(lookup(stock))  # SCHEDULES METHOD FOR EXECUTION AND RETURNS FUTURE OBJECT

    df = pd.DataFrame(data=stock_data)
    df.insert(loc=0, column='symbol', value=stock_list)  # ADD SYMBOLS TO BEGINNING OF DATAFRAME

    # SAVE SCREENER DATA IN OUTPUT WORKSHEET
    output_worksheet = 'raw_data'
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name=output_worksheet)
    workbook = writer.bookworksheet = writer.sheets[output_worksheet]

    # analyst = pd.DataFrame(data=analyst_data)
    # analyst.to_excel(writer, index=False, sheet_name='analysts')
    # workbook = writer.bookworksheet = writer.sheets['analysts']

    # MERGE DATAFRAMES PRIOR TO SAVE
    df2 = df[['symbol', 'Dividend', 'Payout', 'ROE']]  # EXTRACT SPECIFIC COLUMNS FOR ANALYSIS
    input_stock_list = pd.DataFrame(xl_df, columns=['SYMBOL', 'DIV_FREQUENCY', 'DIV_GROWTH_RATE'])
    mergeDf = pd.merge(input_stock_list, df2, left_on='SYMBOL', right_on='symbol')
    del mergeDf['symbol']

    # CAST VARIABLES FOR ANALYSIS
    mergeDf['Dividend'] = mergeDf['Dividend'].replace('%', '', regex=True)
    mergeDf['Dividend'] = mergeDf['Dividend'].replace('-', 0, regex=True)
    mergeDf['Dividend'] = mergeDf['Dividend'].astype(float)

    mergeDf['Payout'] = mergeDf['Payout'].replace('%', '', regex=True)
    mergeDf['Payout'] = mergeDf['Payout'].replace('-', 0, regex=True)
    mergeDf['Payout'] = mergeDf['Payout'].astype(float)

    mergeDf['ROE'] = mergeDf['ROE'].replace('%', '', regex=True)
    mergeDf['ROE'] = mergeDf['ROE'].replace('-', 0, regex=True)
    mergeDf['ROE'] = mergeDf['ROE'].astype(float)

    mergeDf["SUSTAINABLE_GROWTH_RATE"] = ((mergeDf["ROE"] * (1 - mergeDf["Payout"]))).astype(float)
    mergeDf["DIVIDEND_INTRENSIC_VALUE"] = ((mergeDf["Dividend"] * (1 + mergeDf["DIV_GROWTH_RATE"]))/(constants.discount_rate - mergeDf["DIV_GROWTH_RATE"])).astype(float)

    # RENAME COLUMN FOR CLARITY
    mergeDf.rename(columns={'Dividend': 'DIVIDEND(YEARLY)'}, inplace=True)
    mergeDf.rename(columns={'DIV_GROWTH_RATE': 'DIV_GROWTH_RATE(5YR PREV)'}, inplace=True)

    #mergeDf = mergeDf.sort_values(by='ROI', ascending=False)
    # SAVE ANALYSIS DATA IN WORKSHEET 'analysis'
    mergeDf.to_excel(writer, sheet_name='analysis', index=False)

    # RESIZE COLUMNS FOR VISUALS
    worksheet = writer.sheets['analysis']  # pull worksheet object
    for idx, col in enumerate(mergeDf):  # loop through all columns
        series = mergeDf[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
        )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width

    # DEBUG
    writer.save()
    sys.exit()

    # ADDITIONAL COLUMN FORMATTING ['analysis' WORKSHEET]
    workbook = writer.book
    worksheet = writer.sheets['analysis']
    format2 = workbook.add_format({'num_format': '0.00%'})
    worksheet.set_column('H:I', None, format2)

    # PORTFOLIO ANALYSIS
    port1 = df[['symbol', 'Sector', 'Price']]  # EXTRACT SPECIFIC COLUMNS FOR ANALYSIS
    port2 = pd.DataFrame(xl_df, columns=['SYMBOL', 'SHARES'])
    mergePort = pd.merge(port2, port1, left_on='SYMBOL', right_on='symbol')
    del mergePort['symbol']

    mergePort['Price'] = mergePort['Price'].astype(float)
    mergePort["NAV"] = (mergePort["Price"] * mergePort["SHARES"]).astype(float)
    mergePort.to_excel(writer, sheet_name='portfolio', index=False)

    format1 = workbook.add_format({'num_format': '0.00'})
    workbook = writer.book
    worksheet = writer.sheets['portfolio']
    worksheet.set_column('D:E', None, format1)
    writer.save()

    # finviz.get_analyst_price_targets('AAPL')
    # filters = ['fa_div_pos', #POSITIVE DIVIDEND YIELD
    #           'fa_payoutratio_pos', #POSITIVE DIVIDENT PAYOUT RATIO
    #           'sh_avgvol_o200', #AVG VOLUME >200K
    #           'sh_curvol_o200'] #CUR_VOLUME >200K

    # stock_list = Screener(filters=filters, table='Performance', order='change')  # Get the performance table and sort it by price ascending

    # Export the screener results to .csv
    # stock_list.to_csv("stock.csv")

    # Create a SQLite database
    # stock_list.to_sqlite("stock.sqlite3")

    # for stock in stock_list[0:50]:  # Loop through 10th - 20th stocks
    #    print(stock['Ticker'], stock['Price'], stock['Change']) # Print symbol and price

    # Add more filters
    # stock_list.add(filters=['fa_div_high'])  # Show stocks with high dividend yield
    # or just stock_list(filters=['fa_div_high'])

    # Print the table into the console
    # print(stock_list)


if __name__ == '__main__':
    start = time.perf_counter()
    print('SCRIPT NAME:', os.path.basename(__file__))
    now = datetime.now()  # datetime object containing current date and time
    dt_string = now.strftime("%m/%d/%Y %H:%M:%S")
    print("SCRIPT START TIMESTAMP:", dt_string)
    os.chdir(constants.outpath)
    out_dir = os.path.join(constants.outpath, constants.out_screener)
    print('CURRENT PATH:', os.getcwd())
    print('OUTPUT DESTINATION:', os.path.join(constants.outpath, constants.out_screener))

    main(out_dir)

    now = datetime.now()  # datetime object containing current date and time
    dt_string = now.strftime("%m-%d-%Y %H:%M:%S")
    print('SCRIPT END TIMESTAMP:', dt_string)
    finish = time.perf_counter()
    print(f'EXECUTION TIME: {round(finish - start, 2)}s')