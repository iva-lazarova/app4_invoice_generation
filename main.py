import pandas as pd
import glob

# Creating a python list with the filepaths as we have multiple files
filepaths = glob.glob("Invoices/*.xlsx")
for filepath in filepaths:
    # Sheet name required as argument for excel files
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)