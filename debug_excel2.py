import pandas as pd
from io import BytesIO

filename = 'VII SW IIT Test - 1 Result 10.05.2025.xlsx'
print(f"Opening {filename}...")
try:
    with open(filename, 'rb') as f:
        content = f.read()
    
    xl = pd.ExcelFile(BytesIO(content), engine='openpyxl')
    print(f'Sheets: {xl.sheet_names}')
    
    df_raw = pd.read_excel(xl, sheet_name=xl.sheet_names[0], header=None)
    print(f'Raw rows: {len(df_raw)}')
    
    # Print the first 15 rows to see the header structure
    for i in range(min(15, len(df_raw))):
        print(f'Row {i}: {df_raw.iloc[i].tolist()}')
except Exception as e:
    import traceback
    traceback.print_exc()
