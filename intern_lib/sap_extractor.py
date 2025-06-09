# sap_extractor.py

import os
import time
import win32com.client
from typing import List, Dict


class FilterSpec:
    """
    Represents a single filter condition for an SAP table extraction.
    element_id: the GUI element path in SAP GUI scripting.
    value: the selection value (e.g. '>10', '20.10.2021').
    """
    def __init__(self, element_id: str, value: str):
        self.element_id = element_id
        self.value = value


class SAPExtractor:
    """
    A class to extract table data from SAP using GUI scripting.

    Attributes:
        table_names (List[str]): Tables to extract.
        export_path  (str): Directory where exported .txt files go.
        limit_record (str): Max records to fetch (empty = no limit).
        columns      (List[str]): List of field names to include (empty = all).
        filters      (Dict[str, List[FilterSpec]]): Per-table filter specs.
    """

    def __init__(
        self,
        table_names: List[str],
        export_path: str,
        limit_record: str = "",
        columns: List[str] = None,
        filters: Dict[str, List[FilterSpec]] = None,
    ):
        self.table_names = table_names
        self.export_path  = export_path
        self.limit_record = limit_record
        self.columns      = columns or []
        self.filters      = filters or {}

        self._validate_config()

    def _validate_config(self):
        if self.limit_record and not self.limit_record.isnumeric():
            raise ValueError(f"limit_record '{self.limit_record}' must be numeric or empty")
        if not isinstance(self.table_names, list):
            raise TypeError(f"table_names must be a list, got {type(self.table_names)}")
        if not isinstance(self.columns, list):
            raise TypeError(f"columns must be a list, got {type(self.columns)}")
        if not os.path.isdir(self.export_path):
            raise FileNotFoundError(f"Export path does not exist: {self.export_path}")

    def add_filter(self, table: str, element_id: str, value: str):
        """
        Dynamically add a FilterSpec for a given table.
        """
        self.filters.setdefault(table, []).append(FilterSpec(element_id, value))

    def _connect(self):
        SapGuiAuto  = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection  = application.Children(0)
        session     = connection.Children(0)
        session.findById("wnd[0]").maximize()
        return session

    def get_record_count(self, session) -> int:
        results = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
        return results.RowCount

    def extract(self):
        """
        Loop through each table, apply filters, columns, limits, and export to .txt.
        """
        session = self._connect()

        # Jump into SE16N
        session.findById("wnd[0]/tbar[0]/okcd").text = "se16n"
        session.findById("wnd[0]").sendVKey(0)

        for tbl in self.table_names:
            start = time.time()
            self._process_table(session, tbl)

            count = self.get_record_count(session)
            print(f"Table {tbl} has {count} records.")

            # Export via context menu
            shell = session.findById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell")
            shell.pressToolbarContextButton("&MB_EXPORT")
            shell.selectContextMenuItem("&PC")
            session.findById(
                "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/"
                "sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
            ).select()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

            session.findById("wnd[1]/usr/ctxtDY_PATH").text      = self.export_path
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text  = f"{tbl}.txt"
            session.findById("wnd[1]/usr/ctxtDY_FILE_ENCODING").text = "4110"
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
            session.findById("wnd[0]/tbar[0]/btn[3]").press()

            print(f"Processed {tbl} in {time.time() - start:.2f}s")

    def _process_table(self, session, table_name: str):
        # Enter table name and go
        session.findById("wnd[0]/usr/ctxtGD-TAB").text = table_name
        session.findById("wnd[0]").sendVKey(0)

        # Apply any filters
        for spec in self.filters.get(table_name, []):
            session.findById(spec.element_id).text = spec.value

        # Apply record limit
        session.findById("wnd[0]/usr/txtGD-MAX_LINES").text = (
            self.limit_record if self.limit_record else ""
        )

        # Select specific columns, if requested
        if self.columns:
            session.findById("wnd[0]").sendVKey(18)
            for col in self.columns:
                session.findById("wnd[0]").sendVKey(71)
                fld = session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]")
                fld.text = col
                fld.caretPosition = len(col)
                session.findById("wnd[1]").sendVKey(0)
                session.findById(
                    "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]"
                ).selected = True

        # Execute
        session.findById("wnd[0]").sendVKey(8)
