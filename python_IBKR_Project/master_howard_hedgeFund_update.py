"""
Name: Hitesh Punjabi
Date: 09/01/2023

This python script uses TWS API to finds the realised PnL of every trade and the exchange rate for a particular account. The data is then stored onto a google sheet.
"""

import U6084453_data_to_gsheet
import U9417448_data_to_gsheet
import U9426503_data_to_gsheet
import U10758302_data_to_gsheet
import U10919202_data_to_gsheet

print('Updating account U6084453')
U6084453_data_to_gsheet.main()
print('Finish updating account U6084453')

print('Updating hedge accounts: U9417448, U9426503, U10758302, U10919202')

U9417448_data_to_gsheet.main()
print('Finish updating U9417448')
print('Currently updating U9426503')

U9426503_data_to_gsheet.main()
print('Finish updating U9426503')
print('Currently updating U10758302')

U10758302_data_to_gsheet.main()
print('Finish updating U10758302')
print('Currently updating U10919202')

U10919202_data_to_gsheet.main()
print('Finish updating U10919202')
print('All accounts are updated on google sheet')
