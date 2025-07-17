# using for convert string to sortable date format and back
def to_sortable(date_str: str) -> str:
    # date_str is "DD.MM.YYYY"
    dd, mm, yyyy = date_str.split('.')
    return f"{yyyy}{mm}{dd}"      # "YYYYMMDD"

def to_original(code: str) -> str:
    # code is "YYYYMMDD"
    yyyy, mm, dd = code[:4], code[4:6], code[6:8]
    return f"{dd}.{mm}.{yyyy}"    # "DD.MM.YYYY"
