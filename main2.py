import pandas as pd
import re
import numpy as np
import sys

# Custom script
sys.path.append(f"./utils")
from fileselect import getFile

input_file = getFile()
df = pd.read_excel(input_file) 
print(df.head())

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
    'Just Ice Plant (200g)':'Retail',
    'Just Crunchy Lettuce (100g)':'Retail',
    'Little Farms: Mesclun Mix (100g)': 'Retail'
    }

channel_dict = {
    "RedMart Limited":	"RedMart",
    "(Changi) NTUC Fairprice Co-operative Ltd": "Fairprice Darkstores",
    "(Sports Hub) NTUC Fairprice Co-operative Ltd":"Fairprice Darkstores",
    "(JEM) NTUC Fairprice Co-operative Ltd": "Fairprice Darkstores",
    "(Orchid Country Club) NTUC Fairprice Co-operative Ltd": "Fairprice Darkstores",
    "(VivoCity) NTUC Fairprice Co-operative Ltd": "Fairprice Darkstores",
    "NTUC Fairprice Co-operative Ltd":	"Fairprice",
    "(Whampoa) Delivery Hero Stores (Singapore) Pte Ltd": 	"PandaMart",
    "(Outram) Delivery Hero Stores (Singapore) Pte Ltd":	"PandaMart",
    "(Serangoon) Delivery Hero Stores (Singapore) Pte Ltd":	"PandaMart",
    "(Tampines) Delivery Hero Stores (Singapore) Pte Ltd":	"PandaMart", 
    "(Redhill) Delivery Hero Stores (Singapore) Pte Ltd":	"PandaMart",
    "Urban Tiller Pte Ltd":	"Urban Tiller",
    "AVE23 Pte Ltd, OpenTaste":	"OpenTaste",
    "Agrivo Mycosciences Pte Ltd, Mushroom Kingdom":	"Mushroom Kingdom",
    "MoguShop Pte Ltd":	"MoguShop",
    "SuperFresh Grocer Pte Ltd":	"SuperFresh Grocer",
    "Water Tiger Investments Pte Ltd":	"Mmmm!",
    "Urban Origins Pte Ltd":	"Urban Origins",
    "Jin Global Ptd Ltd":	"Sing Sing Mart",
    "Fruitso-Mania Pte Ltd":	"Fruitso-Mania",
    "FoodXervices Inc Pte Ltd":	"FoodXervices",
    "Arktitude Pte Ltd":	"Wholesome Farm",
    "Bootle's Pte Ltd":	"Bootle's",
    "Little Farms Pte Ltd":	"Little Farms",
    "AT Fresh Pte Ltd":	"AT Fresh",
    "The Gut's Factory Pte Ltd":	"The Gut's Feeling",
    "MEOD Pte Ltd":	"MEOD",
    "Glife Technologies Pte Ltd, Greenies": "Greenies",
    "Archisen (Internal Orders)": "Internal Orders",
    "Archisen (TGF)": "The Gut's Feeling",
    "Fairprice (Jem)": "Fairprice Darkstores",
    "Bootle's Pte Ltd.": "Bootle's",
    "Jin Global Ptd Ltd, Sing Sing Mart":"Sing Sing Mart",
    "Internal Orders": "Internal Orders"


    }


df['Weight'] = 0
df['Type'] = ''
df['Channel'] = ''

# process columns one by one
prev_date = None
prev_customer = None
for index, row in df.iterrows():
    # Delivery Date
    date = row['Delivery Date']
    if pd.isnull(date):        
        df.at[index, 'Delivery Date'] = prev_date
    else:
        prev_date = date
        

    # Customer
    customer = row['Customer']
    if pd.isnull(customer):
        df.at[index, 'Customer'] = prev_customer
    else:
        prev_customer = customer

    # Extract weight
    order_line = row['Order Lines/Product']
    
    if type(order_line) == str:
        weight_list = re.findall(r"([\d]+)g", order_line)
        if len(weight_list) > 0:
            df.at[index, 'Weight'] = int(weight_list[0])
    else:
        df.at[index, 'Weight'] = 0

    # type
    if order_line in typedict:
        df.at[index, 'Type'] = typedict[order_line]
    else:
        df.at[index, 'Type'] = 'Promo'

    #  channel
    if prev_customer in channel_dict:
        df.at[index, 'Channel'] = channel_dict[prev_customer]
    
df = df.drop(['Total', 'Invoice Status'], 1)
df['Delivery Date'] = df['Delivery Date'].dt.strftime('%m/%d/%Y')
df.loc[df['Type'] != 'Promo', 'Amount'] = df['Order Lines/Delivered Quantity'] * df['Order Lines/Unit Price'] 
df.loc[df['Type'] == 'Promo', 'Amount'] = df['Order Lines/Unit Price']
df['Sum of Weight'] = df['Order Lines/Delivered Quantity'] * df['Weight'] / 1000

print(df.head())


df.to_excel("sale.order.clean.xlsx")
