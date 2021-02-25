'''
http://www.workout-wednesday.com/pbi-2021-w01/
This week uses a data set that breaks down NCAA athletic department expenses and revenues by year.

- Requirements -
• Create cleaned data set with only teams from the Football Bowl Division. 
• In order to do this, in the FBS Conferences field exclude the FBA Totals and null values.
• Within NCAA Subdivision, remove all Conference Medians
• Create data model properly relating two dim tables to the fact table. 
• Try and use as least amount of steps as possible.

'''
import pandas as pd

# get data
fact = pd.read_excel('NCAA Profit and Losses.xlsx', sheet_name=0)
school_dim = pd.read_excel('NCAA Profit and Losses.xlsx', sheet_name=1)
conference_dim = pd.read_excel('NCAA Profit and Losses.xlsx', sheet_name=2)

# exclude FBA Totals and NULLs in the FBS Conferences field
fact = fact[fact['FBS Conference'].notna()]
fact = fact[fact['FBS Conference'] != 'FBS Total']

# Within NCAA Subdivision, remove all Conference Medians in the NCAA Subdivision column
fact = fact[~fact['NCAA Subdivision'].str.contains('Median')]

# print(fact['NCAA Subdivision'].unique())
# fact.info()
print(fact['FBS Conference'])

# wrtie mutiple dfs to an excel file
writer = pd.ExcelWriter('cleaned_data.xlsx', engine='xlsxwriter')
fact.to_excel(writer, sheet_name='fact', index=None)
school_dim.to_excel(writer, sheet_name='school_dim', index=None)
conference_dim.to_excel(writer, sheet_name='conference_dim', index=None)
writer.save()
