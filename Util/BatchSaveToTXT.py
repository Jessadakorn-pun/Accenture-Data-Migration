import os
import pandas as pd

def save_load_sheets_to_txt(
    excel_path: str,
    out_dir1: str,
    sep: str = '\t',
):
    """
    Reads all sheets in `excel_path` whose name begins with "Load_"
    and writes each to two text files (one in out_dir1, one in out_dir2)
    as UTF-8 with BOM, tab-delimited by default.

    Parameters
    ----------
    excel_path : str
        Path to the source .xlsx file.
    out_dir1 : str
        First output directory for the .txt files.
    out_dir2 : str
        Second output directory for the .txt files.
    sep : str, default '\t'
        Column separator for the text files.
    """

    
    # Load Excel file (to get sheet names without reading all data at once)
    xls = pd.ExcelFile(excel_path)
    
    # Process each sheet whose name starts with "batch_"
    for sheet_name in xls.sheet_names:
        if sheet_name.startswith('batch_'):
            # Read sheet into DataFrame (all columns as strings)
            df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str)
            
            # Construct a base filename (preserve the sheet name)
            filename = f"{sheet_name}.txt"
            
            # Full output paths
            path1 = os.path.join(out_dir1, filename)
            
            # Write both files as UTF-8 with BOM
            df.to_csv(path1, sep=sep, index=False, encoding='utf-8-sig')
            
            print(f"Saved {sheet_name!r} →")
            print(f"  • {path1}")
            print("All sheets saved successfully.")
            
if __name__ == "__main__":
    save_load_sheets_to_txt(
        excel_path=r'C:\Users\j.a.vorathammaporn\OneDrive - Accenture\Desktop\PTT-WorkSpace\3_LSMW_Load\TM-2011\Data_Preload.xlsx',
        out_dir1=r'C:\Users\j.a.vorathammaporn\OneDrive - Accenture\Desktop\PTT-WorkSpace\3_LSMW_Load\TM-2011',
    )
    