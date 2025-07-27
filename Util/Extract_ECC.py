import os
import win32com.client
import sys
import subprocess
import time
import pandas as pd
import pyautogui
from tqdm import tqdm

# extract all data from sap

table_name = """PA0041""".split('\t')  # ใส่ชื่อ table แยกด้วย 1 tab
export_text_path = r"C:\Users\seenlawat.muensuwan\OneDrive - Accenture\Documents\PTT power\SAP automate\Recon M3"  # ใส่ path ที่อยากเก็บ files
limit_record = ""
columns = []

def validate_config(limit, table_list, export, columns):
    if limit != "":
        assert limit.isnumeric(), f"ค่า limit_record ที่ระบุไว้'{limit}' ไม่เป็นตัวเลข ถ้าจะเอาทุก record ใช้ empty string แทนได้"
    assert isinstance(table_list, list), f"ค่า table_list ที่ระบุไว้'{table_list}' ไม่ได้เป็น data type = List กลับไปเช็คว่าค่า table_list ว่าระบุถูกมั้ย?"
    assert isinstance(columns, list), f"ค่า columns ที่ระบุไว้ '{columns}' ไม่ได้เป็น data type = List กลับไปเช็คว่าค่า columns ว่าระบุถูกมั้ย?"
    assert os.path.exists(export), f"ไม่มีfolder export_text_path นี้: {export} กรุณาสร้างใหม่หรืออาจจะใส่พาทผิด"

def get_record_count(session):
    # Get the number of rows in the results table
    results_table = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
    record_count = results_table.RowCount
    return record_count

def extract_data(table_name, path_dir, limit_record, columns=[]):
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    session = connection.Children(0)
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "se16n"
    session.findById("wnd[0]").sendVKey(0)

    for tbl_name in table_name:
        start_time = time.time()  # Start timing

        session.findById("wnd[0]/usr/ctxtGD-TAB").text = tbl_name
        session.findById("wnd[0]").sendVKey(0)
        if limit_record != "":
            session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = f"{limit_record}"
        else:
            session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = ""
        if len(columns) > 0:
            session.findById("wnd[0]").sendVKey(18)
            for column_name in columns:
                session.findById("wnd[0]").sendVKey(71)
                session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").text = column_name
                session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").caretPosition = 5
                session.findById("wnd[1]").sendVKey(0)
                session.findById("wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").selected = True
        session.findById("wnd[0]").sendVKey(8)

        # Get the record count
        record_count = get_record_count(session)
        print(f"Table {tbl_name} has {record_count} records.")

        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").selectContextMenuItem("&PC")
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"{}".format(path_dir)  # ใส่ที่วางไฟล์
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"{tbl_name}.txt"  # ใส่ชื่อไฟล์
        session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "4110"  # set encoding UTF-8
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

        end_time = time.time()  # End timing
        elapsed_time = end_time - start_time
        print(f"Time taken to process table {tbl_name}: {elapsed_time:.2f} seconds")

if __name__ == '__main__':
    validate_config(limit_record, table_name, export_text_path, columns)
    extract_data(table_name, export_text_path, limit_record, columns)
 