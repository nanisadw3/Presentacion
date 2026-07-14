import os
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

import pandas as pd
import sys

filename = "2026 Datos para MV Jun comparativo con rend.xlsm"
downloads_path = os.path.join(os.path.expanduser("~"), "Downloads", filename)
wsl_path = os.path.join("/mnt/c/Users/10900096799/Downloads", filename)
file_path = downloads_path if os.path.exists(downloads_path) else wsl_path
try:
    xls = pd.ExcelFile(file_path)
    print("Sheets:", xls.sheet_names)
    
    sheet_to_use = None
    for sheet in xls.sheet_names:
        normalizado = sheet.lower().replace(" ", "").replace("_", "")
        if "envio" in normalizado or "calculo" in normalizado or "promedio" in normalizado:
            sheet_to_use = sheet
            break
            
    if not sheet_to_use:
        sheet_to_use = xls.sheet_names[0]
        
    print("Sheet selected:", sheet_to_use)
    
    # Let's read Table 3 (AE:AF rows 21-40)
    df_prod = pd.read_excel(file_path, sheet_name=sheet_to_use, usecols="AE:AF", skiprows=20, nrows=20, header=None)
    print("\n--- Table 3 (AE:AF, Rows 21-40) ---")
    print(df_prod)

except Exception as e:
    print("Error:", e)
