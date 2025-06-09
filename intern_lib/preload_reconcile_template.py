# excel_priority_processor.py

import os
from pathlib import Path
import pandas as pd
from typing import Dict, List


class PrioritySheetProcessor:
    """
    Processor for reading priority Excel sheets from a folder and exporting
    cleaned data as UTF-8 BOM-encoded .txt files with tab delimiters.

    Priority order for sheet selection (case-insensitive):
      1. 'delta'
      2. 'm2'
      3. 'm3'

    Handles two cleaning workflows:
      - Delta sheets (sheet names containing 'delta', case-sensitive)
      - Standard sheets (all other priority sheets)
    """

    def __init__(
        self,
        folder_path: str,
        output_folder: str,
        extensions: tuple = (".xlsx", ".xls")
    ):
        """
        Initialize the processor.

        Args:
            folder_path (str):
                Directory containing Excel files to process.
            output_folder (str):
                Directory where processed .txt files will be saved.
            extensions (tuple, optional):
                File extensions to treat as Excel workbooks.
        """
        self.folder = Path(folder_path)
        self.output_folder = Path(output_folder)
        self.extensions = extensions
        self.priority_keywords = ["delta", "m2", "m3"]

        os.makedirs(self.output_folder, exist_ok=True)

    def process_all(self) -> Dict[str, Dict[str, pd.DataFrame]]:
        """
        Process every Excel file in `folder`, selecting sheets by priority,
        cleaning them, saving to .txt, and returning the results.

        Returns:
            Dict[str, Dict[str, pd.DataFrame]]:
                Mapping of filename → {sheet_name → cleaned DataFrame}.
        """
        result: Dict[str, Dict[str, pd.DataFrame]] = {}
        files = [
            p for p in self.folder.iterdir()
            if p.is_file() and p.suffix.lower() in self.extensions
        ]

        for file_path in files:
            processed = self._process_file(file_path)
            if processed:
                result[file_path.name] = processed

        return result

    def _process_file(self, file_path: Path) -> Dict[str, pd.DataFrame]:
        """
        Read an Excel file, pick sheets by priority, clean each, and save.

        Args:
            file_path (Path): Path to the Excel workbook.

        Returns:
            Dict[str, pd.DataFrame]: Cleaned DataFrames keyed by sheet name.
        """
        xls = pd.ExcelFile(file_path, engine="openpyxl")
        sheets = self._select_priority_sheets(xls.sheet_names)
        sheet_dict: Dict[str, pd.DataFrame] = {}

        for sheet in sheets:
            df_raw = pd.read_excel(
                file_path, sheet_name=sheet, dtype=str, engine="openpyxl"
            )
            # Choose cleaning based on sheet name
            if 'delta' in sheet:
                df_clean = self._clean_delta(df_raw)
            else:
                df_clean = self._clean_standard(df_raw)

            sheet_dict[sheet] = df_clean
            self._save_dataframe(df_clean, sheet)

        return sheet_dict

    def _select_priority_sheets(self, sheet_names: List[str]) -> List[str]:
        """
        From a list of sheet names, pick those matching the highest-priority
        keyword.

        Args:
            sheet_names (List[str]): All sheet names in a workbook.

        Returns:
            List[str]: Names of sheets to process.
        """
        lower_map = {name: name.lower() for name in sheet_names}
        for key in self.priority_keywords:
            matches = [name for name, lname in lower_map.items() if key in lname]
            if matches:
                return matches
        return []

    def _clean_delta(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Cleaning workflow for sheets containing 'delta':
          1. Drop columns 1–6.
          2. Drop rows 0,1,4,5,6.
          3. Drop empty columns after first.
          4. Drop columns with 'as-is' in row1 (case-insensitive).
          5. Remove row1.
          6. Promote row0 to header.
          7. Keep only rows where Status == 'Complete'.
          8. Drop first column.
          9. Rename 'PA*-…' to 'Preload-…' by extracting text after last dash.
        """
        df = df.drop(df.columns[1:7], axis=1)
        df = df.drop(index=[0, 1, 4, 5, 6], errors="ignore").reset_index(drop=True)
        df = self._drop_empty_after_first(df)
        df = self._drop_as_is_columns(df)
        df = df.drop(index=1, errors="ignore").reset_index(drop=True)
        df.columns = df.iloc[0]
        df = df.drop(index=0).reset_index(drop=True)
        if "Status" in df.columns:
            df = df[df["Status"] == "Complete"].reset_index(drop=True)
        df = df.drop(df.columns[0], axis=1)
        # Extract suffix after last dash for rename
        df.columns = [
            f"Preload-{col.rsplit('-',1)[1]}" if col.startswith("PA") and "-" in col else col
            for col in df.columns
        ]
        return df

    def _clean_standard(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Cleaning workflow for non-delta priority sheets:
          1. Drop columns 1–2.
          2. Drop rows 0,1,4,5,6.
          3. Drop empty columns after first.
          4. Drop columns with 'as-is' in row1 (case-insensitive).
          5. Remove row1.
          6. Promote row0 to header.
          7. Keep only rows where Status == 'Complete'.
          8. Drop first column.
          9. Rename 'PA*-…' to 'Preload-…' by extracting text after last dash.
        """
        df = df.drop(df.columns[1:2], axis=1)
        df = df.drop(index=[0, 1, 4, 5, 6], errors="ignore").reset_index(drop=True)
        df = self._drop_empty_after_first(df)
        df = self._drop_as_is_columns(df)
        df = df.drop(index=1, errors="ignore").reset_index(drop=True)
        df.columns = df.iloc[0]
        df = df.drop(index=0).reset_index(drop=True)
        if "Status" in df.columns:
            df = df[df["Status"] == "Complete"].reset_index(drop=True)
        df = df.drop(df.columns[0], axis=1)
        # Extract suffix after last dash for rename
        df.columns = [
            f"Preload-{col.rsplit('-',1)[1]}" if col.startswith("PA") and "-" in col else col
            for col in df.columns
        ]
        return df

    def _drop_empty_after_first(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Drop any column (after the first) where row1 is null or empty.
        """
        cols = df.columns.tolist()
        keep = [cols[0]]
        for col in cols[1:]:
            val = df.at[1, col] if 1 in df.index else None
            if pd.notna(val) and val != "":
                keep.append(col)
        return df[keep]

    def _drop_as_is_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Drop columns where row1 contains 'as-is' (case-insensitive).
        """
        to_drop = [
            col for col, cell in df.iloc[1].items()
            if isinstance(cell, str) and 'as-is' in cell.lower()
        ]
        return df.drop(columns=to_drop, errors="ignore")

    def _save_dataframe(self, df: pd.DataFrame, sheet: str) -> None:
        """
        Save a cleaned DataFrame to a .txt file (UTF-8 BOM, tab-sep).

        For delta sheets: extract suffix after last space in sheet name
        and save as "preload_<suffix>.txt".

        For others: use sheet name as-is (spaces replaced) prefixed by
        "preload_".
        """
        suffix = sheet.split()[-1]
        filename = f"preload_{suffix}.txt"
        out_path = self.output_folder / filename
        df.to_csv(out_path, sep="	", index=False, encoding="utf-8-sig")


if __name__ == "__main__":
    processor = PrioritySheetProcessor("input_folder", "output_folder")
    all_data = processor.process_all()
    print(f"Processed {len(all_data)} files.")
