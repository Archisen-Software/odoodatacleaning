import pandas as pd
import numpy as np
import os
import datetime as dt

# Please tell me which column are disappeared in input file
# Task is to generate the new xlsx fiel from twoo sheets ?
#  i need to generate data from 2 sheets, into the same format. Now, i only have data from sheet2 which is fairprice, data from sheet 1 redmart is missing. 

def channelConverter ():
    print(f'This function is used to preprocess data from Fairprice and Redmart.')
    input_file = input("Indicate the full file name here. I.e. If file name is filename.xlsx, input 'filename.xlsx'.")
    # os.chdir("/Users/leon/Documents/Github/bd-sales-analysis/salesCompiler") # Set absolute path where to the file is located
    file = pd.ExcelFile(input_file) 


    sheet_names = {} # Create Empty dictionary
    for sheet in file.sheet_names: # Iterate through the 
        print("Sheet", sheet)
        sheet_names[f'{sheet}'] = pd.read_excel(file,sheet_name = sheet)

    for key, value in sheet_names.items():
        sheet_names[key]['Channel'] = key
        sheet_names[key]['Customer'] = key


    df = pd.concat(sheet_names.values(), ignore_index=True)

    df.drop(columns={'Month','Week Number'},inplace=True) # Remove columns 'Month' and 'Week Number'
    df.to_excel("test1.xlsx")


    df1 = pd.melt(df, id_vars=['Date','Channel','Customer'], var_name='SKU', value_name='Qty') # Convert data from wide to long format, using Date as the ID
    df1.dropna(subset=['Qty'],inplace=True) # Remove all NA values in the 'Qty' column

    ## MAP SKU NAMES 
    # Create dictionary to align all SKU names
    skudict = {
        'Crunchy Classics (100g)':'Just Mesclun: Crunchy Classics (100g)',
        'Peppery Mizuna (100g)':'Just Mesclun: Peppery Mizuna (100g)',
        'Zesty Mustard (100g)':'Just Mesclun: Zesty Mustard (100g)',
        'Tangy Sorrel (100g)':'Just Mesclun: Tangy Sorrel (100g)',
        'Crunchy Classics (500g)': 'Just Mesclun: Crunchy Classics (500g)'
    }

    df1.replace({'SKU':skudict},inplace=True) # Replace the wrong skus with inputs from the dictionary

    # Create dictionary that maps RSP to the SKU
    price_dict = {
        'Just Mesclun: Crunchy Classics (100g)': 4, 
        'Just Mesclun: Peppery Mizuna (100g)': 4.5,
        'Just Mesclun: Zesty Mustard (100g)': 4.5, 
        'Just Mesclun: Tangy Sorrel (100g)': 4,
        'Just Crunchy Lettuce (100g)': 2.2,
        'Just Mustard (20g)': 4,
        'Just Sorrel (20g)': 4, 
        'Just Ice Plant (50g)': 4.5,
        'Just Mustard (250g)': 15,
        'Just Sorrel (250g)': 15,
        'Just Ice Plant (500g)': 30
    }

    # Create dictionary that maps weight to the SKU
    weightdict = {
        'Just Mesclun: Tangy Sorrel (100g)':100,
        'Just Mesclun: Peppery Mizuna (100g)': 100,
        'Just Mesclun: Crunchy Classics (100g)':100,
        'Just Mesclun: Zesty Mustard (100g)': 100,
        'Just Mesclun: Crunchy Classics (500g)':500, 
        'Just Ice Plant (500g)':500,
        'Just Kale (250g)':250, 
        'Just Sorrel (250g)': 250, 
        'Just Mizuna (250g)':250,
        'Just Chard (250g)':250, 
        'Just Mustard (250g)':250,
        'Just Crystal Lettuce (500g)':500,
        'Just Coral Lettuce (500g)':500, 
        'Just Ice Plant (50g)':50,
        'Just Sorrel (20g)':20,
        'Just Mustard (20g)':20,
        'Just Crunchy Lettuce (100g)':100
    }

    # Create dictionary that maps type to SKU
    typedict = {'Just Mesclun: Tangy Sorrel (100g)':'Retail',
        'Just Mesclun: Peppery Mizuna (100g)':'Retail',
        'Just Mesclun: Crunchy Classics (100g)':'Retail',
        'Just Mesclun: Zesty Mustard (100g)':'Retail',
        'Just Mesclun: Crunchy Classics (500g)':'Foodservice', 
        'Just Coral Lettuce (500g)':'Foodservice',
        'Just Ice Plant (500g)':'Foodservice',
        'Just Kale (250g)':'Foodservice', 
        'Just Sorrel (250g)':'Foodservice', 
        'Just Mizuna (250g)':'Foodservice',
        'Just Chard (250g)':'Foodservice', 
        'Just Mustard (250g)':'Foodservice',
        'Just Crystal Lettuce (500g)':'Foodservice', 
        'Just Ice Plant (50g)':'Retail',
        'Just Sorrel (20g)':'Retail',
        'Just Mustard (20g)':'Retail',
        'Just Crunchy Lettuce (100g)':'Retail'}

    df1['Weight(g)'] = df1['SKU'].map(weightdict) # Map product weight according to the SKU
    df1['Sales Price'] = df1['SKU'].map(price_dict) # Map sales price according to the SKU
    df1['Type'] = df1['SKU'].map(typedict) # Map product Type according to the SKU

    print(df1.dtypes)
    # print(df1['Qty'])
    df1['Qty'] = pd.to_numeric(df1['Qty'], errors='coerce')
    df1['Amount'] = df1['Qty']*df1['Sales Price'] # Create new column for 'Amount' to indicate total value obtained for that transaction (Qty * RSP)

    df1 = df1.sort_values(by=['Date'],).reset_index(drop=True)

    # print(df1['Date'])
    # Convert 'date' type from timestamp to datetime
    df1['Date'] = df1['Date'].apply(lambda x: x.date())
    # Convert 'date' type from datetime to str
    
    df1['Date'] = df1['Date'].apply(lambda x: pd.NaT if pd.isnull(x) else dt.datetime.strftime(x, '%d/%m/%Y'))

    df1.to_excel("FPRM_clean.xlsx",sheet_name='Sales by Customer Detail$')

    return (df1)
