"""
Name: Hitesh Punjabi
Date: 09/01/2023

This python script uses TWS API to finds the realised PnL of every trade and the exchange rate for a particular account. The data is then stored onto a google sheet.
This is the normal account.
"""

from ibapi.client import EClient
from ibapi.wrapper import EWrapper
from ibapi.contract import Contract
from threading import Timer
import pandas as pd
from datetime import datetime
import os
import pygsheets

class TestApp(EWrapper, EClient):
    def __init__(self):
        EClient.__init__(self, self)
        self.df = pd.DataFrame(columns=['accountName', 'contract', 'position','marketPrice', 'marketValue', 'averageCost', 'unrealizedPNL','realizedPNL'])
        self.df1 = pd.DataFrame(columns=['ExchangeRate', 'value', 'currency'])

    def error(self, reqId, errorCode, errorString):
        print("Error: ", reqId, " ", errorCode, " ", errorString)

    def nextValidId(self,orderId):
        self.start()

    def updatePortfolio(self, contract: Contract, position: float, marketPrice: float, marketValue: float,
                        averageCost: float, unrealizedPNL: float, realizedPNL: float, accountName: str):
        self.df.loc[len(self.df)] = [accountName, contract.symbol,position,marketPrice, marketValue,averageCost, unrealizedPNL,realizedPNL]

    def updateAccountValue(self, key: str, val: str, currency: str, accountName: str):
        # returns the cash balance, the required margin for the account, or the net liquidity

        if key == 'ExchangeRate':
            self.df1.loc[len(self.df1)] = ['ExchangeRate', val, currency]  # 'ExchangeRate'
        else:
            pass

        print("UpdateAccountValue. Key:", key, "Value:", val, "Currency:", currency, "AccountName:", accountName)

    def updateAccountTime(self, timeStamp: str):
        print("UpdateAccountTime. Time:", timeStamp)

    def accountDownloadEnd(self, accountName: str):
        print("AccountDownloadEnd. Account:", accountName)

    def start(self):
        # ReqAccountUpdates - This function causes both position and account information to be returned for a specified account
        # Invoke ReqAccountUpdates with true to start a subscription
        self.reqAccountUpdates(True, "U10919202") # <----- Change account number here "U6084453"

    def stop(self):
        self.reqAccountUpdates(False,"U10919202") # <----- Change account number here "U6084453"
        self.done = True
        self.disconnect()


def main():
    app = TestApp()
    app.connect("127.0.0.1", 7496, 0)

    Timer(5, app.stop).start()
    app.run()

    #Relative path to google sheet json key
    cwd = os.getcwd()

    # service file is the ib_data json file key
    service_file_path = cwd + '\ib-data-373823-e173b2a01a5f.json'
    #service_file_path = "/ib-data-373823-e173b2a01a5f.json"

    gc = pygsheets.authorize(service_file= service_file_path )

    # google sheet: U6084453_ib_data spreadsheet ID
    spreadsheet_id = '1wRpKjwTwDoYLMYGTRLXG7_ItU5ozwJN4SGTZpIW_AKo' # consolidated spreadsheet ID

        #'1APe0iftDJx6tQBMAltAzRCWFR_w9mJPI6l13Eq8S3Mk' - individual spreadsheet id

    sh = gc.open_by_key(spreadsheet_id)

    # inserting date and time to the data
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")

    app.df['Date_Time'] = dt_string
    app.df1['Date_Time'] = dt_string

    #check if the google sheet has previous data

    wk_sheet_stock = gc.open('Howard Hedge Fund Account').worksheet_by_title('U10919202_stock')
    cells = wk_sheet_stock.get_all_values(include_tailing_empty_rows=False, include_tailing_empty=False, returnas='matrix')

    # checks the number of cells, if there is no data in the gsheet then default value for len(cells) = 1

    # if there is no data
    if len(cells) < 2:
        print('no data in the file')
        data_write = sh.worksheet_by_title('U10919202_stock')
        data_write.clear('A1',None,'*')

        data_write.set_dataframe(app.df, (1,1), encoding='utf-8', fit=True)
        data_write.frozen_rows = 1

        data_write = sh.worksheet_by_title('U10919202_currency')
        data_write.clear('A1', None, '*')
        data_write.set_dataframe(app.df1, (1, 1), encoding='utf-8', fit=True)
        data_write.frozen_rows = 1

    else:
        print('adding data to the existing file')

        stock_df_values = app.df.values.tolist()
        currency_df_values = app.df1.values.tolist()

        gc.sheet.values_append(spreadsheet_id, stock_df_values, "ROWS", "U10919202_stock")
        gc.sheet.values_append(spreadsheet_id, currency_df_values, "ROWS", "U10919202_currency")

if __name__ == "__main__":
   main()