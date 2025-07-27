# using for convert string to sortable date format and back
def to_sortable(date_str: str) -> str:
    if pd.isna(date_str):
        return np.nan
    # date_str is "DD.MM.YYYY"
    else:
        dd, mm, yyyy = date_str.split('.')
        return f"{yyyy}{mm}{dd}"      # "YYYYMMDD"

def to_original(code: str) -> str:
    if pd.isna(code):
        return np.nan
    else:
        # code is "YYYYMMDD"
        yyyy, mm, dd = code[:4], code[4:6], code[6:8]
        return f"{dd}.{mm}.{yyyy}"    # "DD.MM.YYYY"
