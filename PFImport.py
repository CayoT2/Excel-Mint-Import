
# coding: utf-8

# In[90]:

import xlwings as xw
import numpy as np
import pandas as pd

# Update with names: Categories, Transactions, Expenses

    
def transaction_import():

    book_name = 'PF Transaction Import.xlsm'  ## EDIT ##
    cat_sheet_name = 'Categories'
    expense_sheet_name = 'Expenses'
    transaction_sheet_name = 'Transactions'
    
    trans_sheet = xw.Book(book_name).sheets[transaction_sheet_name]
    transactions_data = trans_sheet.range('A1').options(pd.DataFrame, header=1, expand='table',index=False).value
    transactions_data.drop(['Original Description','Account Name','Labels','Notes'], axis=1, inplace=True)

    # Turn Credits into Negative Amount

    for index, row in transactions_data.iterrows():
        if row['Transaction Type'] == 'credit':
            transactions_data.loc[index, 'Amount'] = -row['Amount']

    # Group by and Pivot

    transactions_data = transactions_data.set_index('Date').groupby('Category').resample('M').sum().fillna(0).reset_index().reindex(columns=['Date', 'Category', 'Amount'])

    transactions_data = transactions_data.pivot(index='Category', columns='Date', values='Amount').fillna(0)
    transactions_data = transactions_data[transactions_data.columns[::-1]]

    # Pull Categories from Excel sheet

    cats_sheet = xw.Book(book_name).sheets[cat_sheet_name]
    cat_parent = cats_sheet.range('A2:B2').expand('down').options(dict).value
    categories = cat_parent.keys()
    parent = cats_sheet.range('B2').expand('down').value

    noDupes = []
    [noDupes.append(i) for i in parent if not noDupes.count(i)]
    parent = noDupes

    # Drop Categories not in Excel sheet

    transactions_data = transactions_data.ix[transactions_data.index.isin(categories)]
    transactions_data = transactions_data.reindex(categories).fillna(0)

    # Sum Categories to Parent and Reindex

    transactions_data['Expenses'] = transactions_data.reset_index()['Category'].map(cat_parent).values
    df_sum = transactions_data.groupby('Expenses').sum()
    df_sum = df_sum.reindex(parent)
    
    # Transpose and add Total column
    
    df_sum = df_sum.transpose()
    df_sum['Total'] = df_sum.sum(axis=1)
    df_sum.sort_index(inplace=True)
	

    # Paste to sheet

    expense_sheet = xw.Book(book_name).sheets[expense_sheet_name]
    expense_sheet.clear_contents()
    expense_sheet.range('A1').value = df_sum

