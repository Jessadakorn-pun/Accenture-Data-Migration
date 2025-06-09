# What the VBA Does

## 1. Read Parameters
From the “Parameters” sheet, it grabs:
- **firstRowOfData** – where your data begins
- **mockColumn** – which column holds the “mock” flag (1 = old, 2 = new)
- **deltaIndicatorColumn** – where to write “Delete” / “New” / “Change”
- **concatenatedKeyColumn** – a unique key per row
- **startColumnCheckData** – first column for change comparisons
- **headerRow**, **statusFromPreviousMockColumn**, **recordLimit**

---

## 2. Loop Through Files
On the “File List” sheet:
- For each (path, filename) pair, opens that workbook.
- Calls `ProcessWorkbook` on each file.

---

## 3. Process Each Sheet
- Skips any sheet whose name appears in the “Exception Sheet Name” sheet.
- For each remaining sheet:
  - **Optionally sorts** all rows by the key column (if `Sort Record!A2 = "Yes"`).
  - Runs three routines in order:
    - **DeltaDelete**
    - **DeltaNew**
    - **DeltaChange**

### DeltaDelete
- Builds a set of keys where **mock = 2** (new records).
- For each **mock = 1** (old) row **not flagged "Add"**:
  - If its key is **not in the set**, mark it **"Delete"**.

### DeltaNew
- Builds a set of keys where **mock = 1** (old records).
- For each **mock = 2** row **not in the set**, mark it **"New"**.

### DeltaChange
- Scans **bottom-up**.
- Whenever a **mock = 2** row immediately follows a **mock = 1** row with the same key:
  - **Compare every data column** (skip headers containing `"__"`, `"tobe"`, etc.).
  - If any cell differs:
    - Both rows are marked **"Change"**.
    - Those cells get **pink fill**.
  - If all cells are identical:
    - **Delete** the new row.

---

## 4. Timing & Message
- Wraps the entire run in a **timer**.
- Pops up a **message box** showing **minutes and seconds elapsed** after finishing.

---

# Step-by-Step Flow Inside `DeltaProcessor`

## Initialization (`__init__`)
- Input: **Master-control workbook path**.
- Sets up a **pink fill style** for "Change" highlighting.
- Calls `_load_master()` to pull configuration.

## Loading Master Workbook (`_load_master`)
- Reads “Parameters” into a **dict of ints**:
  - `first_row`, `mock_col`, `delta_col`, `key_col`, `status_prev_col`, `record_limit`, `header_row`, `start_compare_col`
- Reads “File List” into a **DataFrame**.
- Reads “Exception Sheet Name” into a **list** of sheet names to skip.
- Reads `Sort Record!A2` to set **sort_required** flag.

## Entry Point (`run`)
- Loops every row in **File List**:
  - Build `full_path = os.path.join(path, filename)`.
  - If file exists, call `_process_file(full_path)`.

---

## Per-File Processing (`_process_file`)
- Opens the workbook with `openpyxl.load_workbook()`.
- For each sheet:
  - Skip if it’s in the **exception list**.
  - Read sheet into a **pandas DataFrame**, using **first_row** as the header offset.
  - **Sort** by the **key column** if **sort_required**.
  - Calls:
    - `_delta_delete(df)`
    - `_delta_new(df)`
    - `_delta_change(df)` (returns new DataFrame + set of cell-coordinates to color)
  - Write results **back** into the openpyxl sheet:
    - **Values** from the DataFrame
    - **Pink fill** on cells flagged by `_delta_change`
- Save the workbook **back to disk**.

---

# Delta Routines

## `_delta_delete(df)`
- Builds `new_keys = { key | mock = 2 }` within top `record_limit` rows.
- Marks any row where:
  - `mock = 1`
  - `key ∉ new_keys`
  - `previous-status ≠ "Add"`
- Result: Write **"Delete"** into `delta column`.

## `_delta_new(df)`
- Builds `old_keys = { key | mock = 1 }` within top `record_limit` rows.
- Marks any row where:
  - `mock = 2`
  - `key ∉ old_keys`
- Result: Write **"New"** into `delta column`.

## `_delta_change(df)`
- Scans **bottom-up**.
- Whenever a **mock = 2** row immediately follows a **mock = 1** row with same key:
  - Compare all **data columns** from `start_compare_col` onward.
  - If any cell differs:
    - Set both rows’ `delta column` to **"Change"**.
    - Record the differing cell coordinates for **pink fill**.
  - Else:
    - **Drop** the new row (no change).
- Resets **DataFrame index** before returning.

