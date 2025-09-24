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
    
    
# saving to multiple sheet 
# # Save to one Excel file
with pd.ExcelWriter('PA_0001_PERNR_UNIQUE_ALL.xlsx', engine='openpyxl') as writer:
    pernr_bukrs_eq_10.to_excel(writer, sheet_name='PA_0001_PERNR_UNIQUE_BUKRS_10', index=False)
    pernr_all.to_excel(writer, sheet_name='PA_0001_PERNR_UNIQUE_ALL', index=False)


# padding PERNR 
df_PA_0302['PERNR'] = df_PA_0302['PERNR'].astype(str).str.zfill(8)


# remove whitespace from all string entries
Mapping_AUSBI = Mapping_AUSBI.applymap(lambda x: x.strip() if isinstance(x, str) else x)

# reading mapping which handle NA to NAN case
org_Mapping_SLTP1 = pd.read_csv(path_Mapping_SLTP1, sep="\t", encoding="utf-8", dtype=str, keep_default_na=False, na_values=[])


# grouping and forward fill by specific columns
consolidate_filled[cols_to_fill] = (
    consolidate_filled
    .groupby(['PERNR'])[cols_to_fill]
    .transform(lambda g: g.ffill())
)

# adding row number for each PERNR and BEGDA : ROW_NUMBER(OVER PARTITION BY PERNR, BEGDA ORDER BY PERNR, BEGDA)
result_2.values_sort(by=['0302_PERNR', '0302_BEGDA'], assending=True)
result_2['ROW_NUMBER'] = result_2.groupby(['0302_PERNR', '0302_BEGDA']).cumcount() + 1


def intersects(start1, end1, start2, end2):
    """Return a boolean Series indicating if [start1–end1] overlaps [start2–end2]."""
    return (
        ((start1 >= start2) & (start1 <= end2)) |
        ((start2 >= start1) & (start2 <= end1))
    )