import pandas as pd
import numpy as np
import os


# os.chdir("/Users/leon/Documents/Github/bd-sales-analysis/salesCompiler")

# Absolute path for the preprocessing scripts
from preprocess1 import channelConverter
from preprocess2 import quickbooksCleaner

# import importlib
# import quickbooksCleaner
# importlib.reload(quickbooksCleaner)

# Run Channel Converter Script
demand_df = channelConverter()
# Run Quickbooks Converter Script
qb_df = quickbooksCleaner()


# Concatenate both dataframes
df = pd.concat([demand_df,qb_df], ignore_index=True).sort_values(by='Date').reset_index(drop=True)
df.dropna(subset=['Amount'],inplace=True)

# Export final dataframe as an excel document, with sheet named as Quickbooks default sheet name
df.to_excel(("monthmasterCompiler.xlsx"),sheet_name='Sales by Customer Detail$')
