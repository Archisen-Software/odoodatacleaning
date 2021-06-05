# Clean up sales data from Quickbooks for analysis in Tableau

import pandas as pd
import numpy as np
import os


def quickbooksCleaner ():
    print(f'This function is used to preprocess data from Quickbooks.')
    # Load file
    file = input("Indicate file name here. I.e. If file name is filename.xlsx, input 'filename'.")
    df = pd.read_excel('{}.xlsx'.format(file)) 
    df = df.drop(df.columns[0],axis=1) # Drop first empty column
    df = df.rename(columns={'Memo/Description':'SKU'}).dropna(subset=['SKU'])

    # Exclude irrelevant SKU observations 
    excludeSku = ['RedMart Commission','Sample','Farm Tower','Farm Services','RedMart Discount','Other Vegetables','Grab Service Fee','Diff on invoice', 'make up for difference adjusted under Inv 140 Delivery hero Jurong','minor diff'] # List of partial strings in SKU to exclude
    cleandf = df[~df['SKU'].str.contains('|'.join(excludeSku))].reset_index(drop=True) # Remove rows that contain these strings

    # Exclude irrelevant Product/Services
    excludePdt = ['Payment received']
    cleandf = cleandf[~cleandf['Product/Service'].str.contains('|'.join(excludePdt))].reset_index(drop=True) # Remove rows that contain these strings
    
    # Exclude irrelevant Customer observations
    excludeCust = ['Cash','Archisen Pte Ltd','RedMart','NTUC Fairprice']
    cleandf = cleandf[~cleandf['Customer'].str.contains('|'.join(excludeCust))].reset_index(drop=True) # Remove rows that contain these strings

    # Replacing wrong SKUs with newly finalised SKU names 
    # Setting dictionary for SKUs mapping
    skudict = {'Just Mesclun + Mizuna (100g)':'Just Mesclun: Peppery Mizuna (100g)',
        'Just Mesclun + Sorrel (100g)':'Just Mesclun: Tangy Sorrel (100g)',
        'Just Mesclun (100g)':'Just Mesclun: Crunchy Classics (100g)',
        'Ice Plant, Himalayan Pink Salt (500g)':'Just Ice Plant (500g)',
        'Mesclun (500g)': 'Just Mesclun: Crunchy Classics (500g)',
        'Just Mesclun (500g)': 'Just Mesclun: Crunchy Classics (500g)',
        'Kale Leaves, Tuscan (250g)':'Just Kale (250g)',
        'Kale Leaves, Tuscan (500g)': 'Just Kale (250g)',
        'Just Produce + Mizuna (100g)': 'Just Mesclun: Peppery Mizuna (100g)',
        'Just Mesclun + Mustard': 'Just Mesclun: Zesty Mustard (100g)',
        'Just Mesclun + Mustard (100g)': 'Just Mesclun: Zesty Mustard (100g)',
        'Ice Plant, Himalayan Pink Salt (1000g)': 'Just Ice Plant (500g)',
        'Mizuna, Green (1000g)': 'Just Mizuna (250g)',
        'Sorrel, Red-Veined (1000g)': 'Just Sorrel (250g)',
        'Swiss Chard, Rainbow (1000g)': 'Just Chard (250g)',
        'Mustard, Green Wave (1000g)': 'Just Mustard (250g)',
        'Kale, Tuscan (250g)': 'Just Kale (250g)',
        'Lettuce, Crystal (500g)': 'Just Crystal Lettuce (500g)',
        'Mizuna, Green (250g)': 'Just Mizuna (250g)',
        'Sorrel, Red-Veined (250g)':'Just Sorrel (250g)',
        'Mustard, Green Wave (250g)':'Just Mustard (250g)'}

    cleandf.replace({'SKU':skudict},inplace=True) # Replace the wrong skus with inputs from the dictionary

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

    weightdict = {'Just Mesclun: Tangy Sorrel (100g)':100,
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
    'Just Crunchy Lettuce (100g)':100}

    channeldict = {
        # Old
        'Sheng Siong Supermarket Pte Ltd (HQ):Sheng Siong (200 Upper Thomson Road)':'Sheng Siong',
        'Sheng Siong Supermarket Pte Ltd (HQ):Sheng Siong (301 Punggol Central)':'Sheng Siong',
        'Sheng Siong Supermarket Pte Ltd (HQ):Sheng Siong (623 Elias Mall)':'Sheng Siong',
        'Sheng Siong Supermarket Pte Ltd (HQ):Sheng Siong (85 Dawson Road)':'Sheng Siong',
        'Sheng Siong Supermarket Pte Ltd (HQ):Sheng Siong (527D Pasir Ris St 51)':'Sheng Siong',
        'Sheng Siong Supermarket Pte Ltd (HQ):Sheng Siong (25 Ghim Moh Link)':'Sheng Siong',
        'Sheng Siong Supermarket Pte Ltd (HQ):Sheng Siong (739A Bedok Reservoir Road)':'Sheng Siong',
        'Sheng Siong Supermarket Pte Ltd (HQ):Sheng Siong (720, Clementi West St 2)':'Sheng Siong',
        'Delivery Hero Stores:PandaMart (Redhill)':'PandaMart',
        'Delivery Hero Stores:PandaMart (Yishun)':'PandaMart',
        'Delivery Hero Stores:PandaMart (Woodlands)':'PandaMart',
        'Delivery Hero Stores:PandaMart (Whampoa)':'PandaMart',
        'Delivery Hero Stores:PandaMart (Sims)':'PandaMart',
        'Delivery Hero Stores:PandaMart (Outram)':'PandaMart',
        'Delivery Hero Stores:PandaMart (Jurong)':'PandaMart',
        'Delivery Hero Stores:PandaMart (Bukit Batok)':'PandaMart',
        'Delivery Hero Stores:PandaMart (Tampines)':'PandaMart',
        'Delivery Hero Stores:PandaMart (Bedok)':'PandaMart',
        'Delivery Hero Stores:PandaMart (Mandai)':'PandaMart',
        'Delivery Hero Stores:PandaMart (Ang Mo Kio)':'PandaMart',
        'Delivery Hero Stores:PandaMart (Serangoon)':'PandaMart',
        'Delivery Hero Stores:PandaMart (Punggol)':'PandaMart',
        'At Fresh Pte Ltd':'AtFresh',
        'Urban Tiller Pte Ltd':'Urban Tiller', 
        "Bootle's Pte Ltd": "Bootles",
        'Urban Origins Pte Ltd':'Urban Origins',
        'Arktitude Pte Ltd': 'Wholesome Farms',
        
        # New
        'PandaMart:PandaMart (Redhill)':'PandaMart',
        'PandaMart:PandaMart (Yishun)':'PandaMart',
        'PandaMart:PandaMart (Woodlands)':'PandaMart',
        'PandaMart:PandaMart (Whampoa)':'PandaMart',
        'PandaMart:PandaMart (Sims)':'PandaMart',
        'PandaMart:PandaMart (Outram)':'PandaMart',
        'PandaMart:PandaMart (Jurong)':'PandaMart',
        'PandaMart:PandaMart (Bukit Batok)':'PandaMart',
        'PandaMart:PandaMart (Tampines)':'PandaMart',
        'PandaMart:PandaMart (Bedok)':'PandaMart',
        'PandaMart:PandaMart (Mandai)':'PandaMart',
        'PandaMart:PandaMart (Ang Mo Kio)':'PandaMart',
        'PandaMart:PandaMart (Serangoon)':'PandaMart',
        'PandaMart:PandaMart (Punggol)':'PandaMart',
        'Arktitude': 'Wholesome Farms'
        }


    cleandf['Type'] = cleandf['SKU'].map(typedict) # Map product Type according to the SKU
    cleandf['Weight(g)'] = cleandf['SKU'].map(weightdict) # Map product weight according to the SKU
    cleandf['Channel'] = cleandf['Customer'] 
    cleandf.replace({'Channel': channeldict},inplace=True) # Replace the wrong skus with inputs from the dictionary
    cleandf.drop(columns={'Product/Service','Balance'},inplace=True)


    cleandf.to_excel(("{}_clean.xlsx").format(file),sheet_name='Sales by Customer Detail$')
    return (cleandf)
