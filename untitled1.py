from bsedata.bse import BSE
from datetime import datetime
import time
import os
import pandas as pd
from csv import writer



while True:
    today = datetime.now()
    dt_string = today.strftime("%d/%b/%Y"+' | '+ "%H:%M:%S")
    hms = today.strftime("%H:%M:%S") #hour min sec
    b = BSE()

    quote = b.getQuote('500325') #get quote name
    quote['Date & Time'] = dt_string
    # pprint(quote)
    current = quote['currentValue']
#    print(quote)

    #creating excel set
    def CreatingNewExcelFile():
        # writer = pd.ExcelWriter("YourBseExcelFile.xlsx",engine='xlsxwriter')

        df = pd.DataFrame.from_dict(quote)
        new_df = df[["Date & Time",'companyName','previousOpen','dayLow','dayHigh','currentValue']]

        # Convert the dataframe to an XlsxWriter Excel object.
        new_df.to_csv('YourBseCsvFile.csv', index=False)

        # # Close the Pandas Excel writer and output the Excel file.
        # writer.save()
        # writer.close()

    def AppendingToExistingExcelFile(message):

        # writer = pd.ExcelWriter('YourBseExcelFile.xlsx', engine='openpyxl')
        #df = pd.DataFrame.from_dict(quote)
        companyName = message['companyName']
        previousOpen = message['previousOpen']
        dayLow = message['dayLow']
        dayHigh = message['dayHigh']
        #new_df = df[["Date & Time", 'companyName', 'previousOpen', 'dayLow', 'dayHigh','currentValue']]

        with open('YourBseCsvFile.csv', 'a+') as csv_file:
            # Pass this file object to csv.writer()
            # and get a writer object
            writer_object = writer(csv_file)

            # Pass the list as an argument into
            # the writerow()
            writer_object.writerow([dt_string,companyName,previousOpen,dayLow,dayHigh,current])

            # Close the file object
            csv_file.close()


    if os.path.exists("YourBseCsvFile.csv"):
        print("Appeding new Data...")
        AppendingToExistingExcelFile(quote)
        time.sleep(10)
    else:
        print("creating new record ...")
        CreatingNewExcelFile()
