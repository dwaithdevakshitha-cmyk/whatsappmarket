import pandas as pd
import os
path = r"C:\Users\HP\Documents\Trip flux marketing\USA work Permit_1_to_40_06_03_2026.xlsx"
if os.path.exists(path):
    df = pd.read_excel(path, header=None)
    print(df.head(10).to_string())
else:
    print("File not found")
