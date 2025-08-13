import os
import re
import pandas as pd

def save_batch_sheets_to_txt(
    excel_path: str,
    out_dir: str,
    sep: str = '\t',
):
    """
    Reads all sheets in `excel_path` whose name ends with '_batchN'
    (where N is one or more digits), and writes each to a .txt file
    in out_dir as UTF-8 with BOM, tab-delimited by default.
    """
    # make sure output folder exists
    os.makedirs(out_dir, exist_ok=True)

    # regex: matches any name ending in "_batch" + digits
    batch_re = re.compile(r'_batch\d+$')

    xls = pd.ExcelFile(excel_path)
    for sheet_name in xls.sheet_names:
        if batch_re.search(sheet_name):
            # read as all strings
            df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=str)

            # output filename = same as sheet
            out_path = os.path.join(out_dir, f"{sheet_name}.txt")

            # write with UTF-8 BOM
            df.to_csv(out_path, sep=sep, index=False, encoding='utf-8-sig')

            print(f"Saved batch sheet: {sheet_name} → {out_path}")

    print("All batch sheets exported.")

            
if __name__ == "__main__":
    save_load_sheets_to_txt(
        excel_path=r'C:\Users\j.a.vorathammaporn\OneDrive - Accenture\Desktop\PTT-WorkSpace\3_LSMW_Load\TM-2011\test\test.xlsx',
        out_dir1=r'C:\Users\j.a.vorathammaporn\OneDrive - Accenture\Desktop\PTT-WorkSpace\3_LSMW_Load\TM-2011\test',
    )
    