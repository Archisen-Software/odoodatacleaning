import pandas as pd
import re
import numpy as np
from datetime import datetime

input_file = 'C:\\Users\\User\\Downloads\\sale.order (15).xlsx'
df = pd.read_excel(input_file) 
# print(df.head())

variantdict = {'Just Mesclun: Tangy Sorrel (100g)':'Sorrel, Red-Veined|Lettuce, Green Romaine|Lettuce, Green Crystal',
    'Just Mesclun: Crunchy Classics (100g)':'Lettuce, Green Romaine|Lettuce, Green Crystal',
    'Just Mesclun: Zesty Mustard (100g)':'Mustard, Green Wave|Lettuce, Green Romaine|Lettuce Green Crystal',
    'Just Mesclun: Crunchy Classics (500g)':'Lettuce, Green Romaine|Lettuce, Green Crystal', 
    'Just Coral Lettuce (500g)':'Lettuce, Red Coral',
    'Just Ice Plant (500g)':'Ice Plant, Himalayan Pink Salt',
    'Just Sorrel (250g)':'Sorrel, Red-Veined', 
    'Just Chard (250g)':'Swiss Chard, Rainbow', 
    'Just Mustard (250g)':'Mustard, Green Wave',
    'Just Crystal Lettuce (500g)':'Lettuce, Green Crystal,', 
    'Just Ice Plant (50g)':'Ice Plant, Himalayan Pink Salt',
    'Just Sorrel (20g)':'Sorrel, Red-Veined',
    'Just Mustard (20g)':'Mustard, Green Wave',
    'Just Ice Plant (200g)':'Ice Plant, Himalayan Pink Salt',
    'Just Crunchy Lettuce (100g)':'Lettuce, Green Crystal',
    'Little Farms: Mesclun Mix (100g)': 'Lettuce, Green Crystal|Lettuce, Green Romaine|Lettuce, Red Coral'
    }

# get all products
summary_data = {'Week': []}
product_data = {}
for key in variantdict:
    val = variantdict[key]
    val_list = val.split("|")
    for item in val_list:
        summary_data[item] = []
        product_data[item] = 0

print(summary_data)
print(product_data)

prev_week = -1


for index, row in df.iterrows():
    date = row['Order Date']
    qty = int(row['Order Lines/Delivered Quantity'])
    if pd.isnull(date):        
        pass
    else:
        prev_date = date

    order_line = row['Order Lines/Description']

    weight = 0
    
    if type(order_line) == str:
        weight_list = re.findall(r"([\d]+)g", order_line)
        if len(weight_list) > 0:
            weight = float(weight_list[0])    

            # calculate the number of week
            week_num = prev_date.to_pydatetime().isocalendar()[1]

            if prev_week != week_num:
                if prev_week >= 0:
                    summary_data['Week'].append(prev_week)
                    for key in product_data:
                        summary_data[key].append(product_data[key])
                for key in product_data:                   
                    product_data[key] = 0

                prev_week = week_num
                
            # get product
            val_list = variantdict[order_line].split("|")
            prod_count = len(val_list)

            unit_weight = weight / prod_count / 1000

            print(week_num, prev_date, weight, qty)

            for item in val_list:
                product_data[item] += unit_weight * qty
            

if prev_week >= 0:
    summary_data['Week'].append(prev_week)
    for key in product_data:
        summary_data[key].append(product_data[key])


print(summary_data)

result_df = pd.DataFrame(summary_data)
result_df.to_excel("salesconverter.clean.xlsx")
