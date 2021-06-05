import pandas as pd
import numpy as np
import os


# os.chdir("/Users/leon/Documents/Github/bd-sales-analysis/salesCompiler")

from preprocess2 import quickbooksCleaner

# import importlib
# import quickbooksCleaner
# importlib.reload(quickbooksCleaner)

# Run Quickbooks Converter Script
qb_df = quickbooksCleaner()


# Concatenate both dataframes
# Export final dataframe as an excel document, with sheet named as Quickbooks default sheet name
qb_df.to_excel((".xlsx"),sheet_name='Sales by Customer Detail$')
