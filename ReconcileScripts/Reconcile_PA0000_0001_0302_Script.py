# import package
import os
import pandas as pd
import numpy as np
from pathlib import Path


# utility function

def to_sortable_date(date_str: str) -> str:
    """
    This function will convert date format of text from SAP to sortable date
    
    Args:
        date_str (str): text date from SAP format DD.MM.YYYY
        
    Returns:
        string: sortable text date format YYYYMMDD
        
    Raise:
        ValueError: if the string format is not in expect format       
    """
    try:
        dd, mm, yyyy = date_str.split('.')
        # validate date format
        if len(dd) != 2 and len(mm) != 2 and len(yyyy) != 4:
            raise ValueError("Invalid date format: Expect format to be DD.MM.YYYY")
        
        return f"{yyyy}{mm}{dd}"
    
    except ValueError as err:
        print(f"Invalid date string: {date_str}. Error: {err}")


def to_original_date(date_str: str) -> str:
    """
    This function will convert text date to the original SAP format
    
    Args:
        date_str(str): text date format of YYYYMMDD
    
    Returns:
        string: original SAP date format DD.MM.YYYY
    
    Raise:
        ValueError: if the string format is not in expect format
    """
    
    if len(date_str) != 8 and date_str.isdigit():
        raise ValueError(f"Invalid date format: Expect format to be YYYYMMDD, got: {date_str}")
    
    return f"{date_str[6:]}.{date_str[4:6]}.{date_str[:4]}"

def validate_file_path(root_path: str, file_name: str, allowed_format: list[str] = [".txt"]) -> Path:
    """
    validate the data path and type of file that using in the program
    
    Args:
        root_path(str): root folder of the source file
        file_name: file name that using the program
        allowed_ext (list[str], optional): List of allowed file extensions. Default is [".txt"].
        
    Returns:
        Path: validated full file path
    
    Raise:
        ValueError: if the file is not exist or not a .txt file
    
    """
    print(f"Checking path of file: {file_name}")
    
    # concat path file
    full_path = Path(root_path) / file_name
    
    if not full_path.exists():
        raise FileExistsError(f"File not found : {full_path}")
    
    if full_path.suffix.lower() not in allowed_format:
        raise ValueError(f"The input file is not .txt format, got {full_path.suffix}")
    
    return full_path

# function for main process
def prepare_data_0000_0302(**kwargs)-> pd.DataFrame:
    """
    Prepare Action and Additional Action data using for join with Organization data including 
        - read in the data
        - select only the relevant columns
        - change type and format
        - remove white space of columns names and values in column PERNR
        - adding prefix including object code for each columns
        - inner join PA0302 with 0000 on PERNR and BEGDA

    Args:
        kwargs:
            path_0000 (str): Path to PA0000 file.
            path_0302 (str): Path to PA0302 file.
            pa_0000_columns (list[str]): Columns to select from PA0000.
            pa_0302_columns (list[str]): Columns to select from PA0302.

        
    Returns:
        pd.DataFrame: cleaned and joined dataframe between object PA000 and PA0302
        
    Raise:
        ValueError: if the error when reading text file and unexpect column in the text file
        
    """

    try:
        # Prepare data for object PA-0000
        df_0000 = pd.read_csv(kwargs.get("path_0000"), sep='\t', encoding='utf-8', dtype=str)
        df_0000.columns = df_0000.columns.str.strip()
        
        if not set(kwargs.get("columns_0000")).issubset(df_0000.columns):
            raise ValueError(f"Columns not found, expect {kwargs.get("columns_0000")}")
            
        df_0000 = df_0000[kwargs.get("columns_0000")]
        df_0000["PERNR"] = df_0000["PERNR"].str.strip()
        df_0000['PERNR'] = df_0000['PERNR'].astype(str).str.zfill(8)
        df_0000 = df_0000.add_prefix("0000_")
        
        # Prepare data for object PA-0302
        df_0302 = pd.read_csv(kwargs.get("path_0302"), sep='\t', encoding='utf-8', dtype=str)
        df_0302.columns = df_0302.columns.str.strip()
        
        if not set(kwargs.get("columns_0302")).issubset(df_0302.columns):
            raise ValueError(f"Columns not found, expect {kwargs.get("columns_0302")}")
            
        df_0302 = df_0302[kwargs.get("columns_0302")]
        df_0302["PERNR"] = df_0302["PERNR"].str.strip()
        df_0302['PERNR'] = df_0302['PERNR'].astype(str).str.zfill(8)
        df_0302 = df_0302.add_prefix("0302_")
        
        # Join PA-0000 with PA-0302 on PERNR and BEGDA
        df_merge = df_0302.merge(
            df_0000,
            left_on=['0302_PERNR', '0302_BEGDA'], 
            right_on=['0000_PERNR', '0000_BEGDA'],
            how='inner'
        )
        
        # select the relevant columns
        output_columns = [
            "0302_PERNR", "0302_BEGDA", "0000_ENDDA", "0302_SEQNR", "0302_MASSN", "0302_MASSG", "0000_STAT2"
        ]
        
        df_merge = df_merge[output_columns]
        
        # create row number over partition by PERNR and BEGDA
        df_merge['row_number'] = df_merge.groupby(['0302_PERNR', '0302_BEGDA']).cumcount() + 1
        
        return df_merge
    
    except Exception as err:
        raise ValueError(f"Error preparing data: {err}")
    
def prepare_data_0001_hrp_1001_sp(**kwargs)-> pd.DataFrame:
    """
    Prepare Organization and Relationship between employee and Organization data using for join with Action data including 
        - read in the data
        - select only the relevant columns
        - change type and format
        - HRP1001-SP join PA0001 (on "PERNER" "OBJID" / "PERNER" "PLANS" (add index before left join)) 
        - select HRP1001-SP record if "BEGDA" and "ENDDA" of PA0001 in range PA0001 using left join and carry the result which not match record in order to find the PROZT of each record
        
        kwargs:
            path_0001 (str): Path to PA0001 file.
            path_1001_sp (str): Path to HRP1001-SP file.
            pa_0001_columns (list[str]): Columns to select from PA001.
            hrp_1001_sp_columns (list[str]): Columns to select from HRP0001-SP.

        
    Returns:
        pd.DataFrame: cleaned and joined dataframe between object PA001 and HRP0001-SP
        
    Raise:
        ValueError: if the error when reading text file and unexpect column in the text file
        
    """
    
    try:
        # prepare data for object PA-0001
        df_0001 = pd.read_csv(kwargs.get("path_0001"), sep="\t", encoding="utf-8", dtype=str)
        df_0001.columns = df_0001.columns.str.strip()
        
        if not set(kwargs.get("columns_0001")).issubset(df_0001.columns):
            raise ValueError(f"Columns not found, expect {kwargs.get("columns_0001")}")
        df_0001 = df_0001[kwargs.get("columns_0001")]
        
        # reset the index of df_0001
        df_0001.reset_index(inplace=True) 
        # increment the index by 1
        df_0001['index'] = df_0001['index'] + 1
        df_0001["PERNR"] = df_0001["PERNR"].str.strip()
        df_0001["PERNR"] = df_0001["PERNR"].astype(str).str.zfill(8)
        df_0001 = df_0001.add_prefix("0001_")
        
        # prepare data for object HRP-1001-SP
        df_1001_SP = pd.read_csv(kwargs.get("path_1001_sp"), sep="\t", encoding="utf-8", dtype=str)
        df_1001_SP.columns = df_1001_SP.columns.str.strip()
        
        if not set(kwargs.get("columns_1001_sp")).issubset(df_1001_SP.columns):
            raise ValueError(f"Columns not found, expect {kwargs.get("columns_1001_sp")}")
        df_1001_SP = df_1001_SP[kwargs.get("columns_1001_sp")]
        
        df_1001_SP = df_1001_SP[kwargs.get("hrp_1001_sp_columns")]
        df_1001_SP = df_1001_SP.add_prefix("1001_")
        
        # Join df_PA_0001 with df_1001_SP on PERNR and PLANS
        # using inner join to find matching records
        leftMerge = df_0001.merge(
            df_1001_SP, 
            left_on=['0001_PERNR', '0001_PLANS'], 
            right_on=['1001_SOBID', '1001_OBJID'],
            # how='inner'
            how='left'
        )
        
        # convert date columns to sortable format
        leftMerge['0001_SORT_BEGDA'] = leftMerge['0001_BEGDA'].apply(to_sortable_date)
        leftMerge['0001_SORT_ENDDA'] = leftMerge['0001_ENDDA'].apply(to_sortable_date)
        leftMerge['1001_SORT_BEGDA'] = leftMerge['1001_BEGDA'].apply(to_sortable_date)
        leftMerge['1001_SORT_ENDDA'] = leftMerge['1001_ENDDA'].apply(to_sortable_date)

        # create filter mask for overlapping date ranges of PA-0001 and HRP-1001-SP
        mask = (
            (leftMerge['1001_SORT_BEGDA'].isna() &  leftMerge['1001_SORT_ENDDA'].isna() ) | # including null records
            (leftMerge['0001_SORT_ENDDA'] >= leftMerge['1001_SORT_BEGDA']) &
            (leftMerge['0001_SORT_BEGDA'] <= leftMerge['1001_SORT_ENDDA'])
        )
            
        # apply the mask to filter the DataFrame
        leftMerge_filtered = leftMerge[mask]
        
        # row_number over(partition by index order by PROZT)
        df_0001_find_prozt = leftMerge.copy()
        df_0001_find_prozt = df_0001_find_prozt[['0001_index', '0001_PERNR', '0001_BEGDA', '0001_ENDDA', '1001_PROZT']]
        df_0001_find_prozt = df_0001_find_prozt.sort_values(['0001_index','1001_PROZT'])
        df_0001_find_prozt['row_number'] = df_0001_find_prozt.groupby('0001_index').cumcount() + 1
        df_0001_find_prozt = df_0001_find_prozt[df_0001_find_prozt['row_number'] == 1]
        df_0001_find_prozt = df_0001_find_prozt.add_prefix("findProzt_")
        
        # Join df_PA_0001 with df_HRP_1001_SP on PERNR and PLANS
        # using left join to find matching records
        df_PA_0001_final = df_0001.merge(
            df_0001_find_prozt, 
            left_on=['0001_index'], 
            right_on=['findProzt_0001_index'],
            how='left'
        )
        
        col_PA_0001_merge = [
            "0001_PERNR", "0001_BEGDA", "0001_ENDDA", "0001_BUKRS", "0001_WERKS", "0001_BTRTL", "0001_KOSTL", "0001_GSBER", 
            "0001_PERSG", "0001_ABKRS", "0001_PERSK", "0001_ANSVH", "findProzt_1001_PROZT", "0001_PLANS", "0001_STELL", "0001_ORGEH", 
            "0001_SBMOD", "0001_SACHP", "0001_SACHZ", "0001_SACHA", "0001_MSTBR"
        ]
        
        df_PA_0001_final = df_PA_0001_final[col_PA_0001_merge].copy()
        df_PA_0001_final.rename(columns={'findProzt_1001_PROZT': '1001_PROZT'}, inplace=True)

        # fill NA with 0 in columns : STELL ORGEH
        df_PA_0001_final['0001_STELL'].fillna("00000000", inplace=True)
        df_PA_0001_final['0001_ORGEH'].fillna("00000000", inplace=True)
        
        return df_PA_0001_final
        
        
    except Exception as err:
        raise ValueError(f"Error preparing data: {err}")

def consolidate_postload(action_data: pd.DataFrame, oranization_data: pd.DataFrame)-> tuple[pd.DataFrame, list[str]]:
    """
    Join Action data and Oranization data for preparing Postload Reconcilation
    
    Args:
        action_data(pd.Dataframe): cleaned action data from prepare_data_0000_0302
        oranization_data(pd.Dataframe): cleaned oranization data from prepare_data_0001_hrp_1001_sp
    
    Returns:
        pd.DataFrame: cleaned and joined dataframe between Action and Organization
        List[str]: Standardized column names
        
    Raise:
        ValueError: if the error when reading text file and unexpect column in the text file
    
    """
    try:
        # create tables with prefixes for merging
        pa0000_table = action_data.add_prefix('0_')
        pa0001_table = oranization_data.add_prefix('1_')

        # fill pa0302 sequence NA with 0
        pa0000_table['0_0302_SEQNR'] = pa0000_table['0_0302_SEQNR'].fillna(0)

        # merge the pa0000 and pa0001 tables on PERNR and BEGDA
        consolidate = pd.merge(
                            pa0000_table, pa0001_table, 
                            left_on=['0_0302_PERNR', '0_0302_BEGDA'],
                            right_on=['1_0001_PERNR', '1_0001_BEGDA'],
                            how='outer'
        )

        # create pernr column using for sort
        consolidate['PERNR'] = np.where(
            consolidate['1_0001_PERNR'].isna() | (consolidate['1_0001_PERNR'] == ''),
            consolidate['0_0302_PERNR'],
            consolidate['1_0001_PERNR']
        )

        consolidate['BEGDA'] = np.where(
            consolidate['1_0001_BEGDA'].isna() | (consolidate['1_0001_BEGDA'] == ''),
            consolidate['0_0302_BEGDA'],
            consolidate['1_0001_BEGDA']
        )

        # convert BEGDA to sortable format for sorting
        consolidate['BEGDA_SORT'] = consolidate['BEGDA'].apply(to_sortable_date)

        # sort the consolidated DataFrame by PERNR and BEGDA_SORT
        consolidate.sort_values(by=['PERNR', 'BEGDA_SORT', '0_0302_SEQNR'], inplace=True)

        # adding sequence number
        consolidate['SEQUENCE'] = consolidate['0_0302_SEQNR'].fillna(0)

        # define columns to fill down
        cols_to_fill = [c for c in consolidate.columns if c not in ['PERNR', 'BEGDA', 'BEGDA_SORT', '0_0302_SEQNR','SEQUENCE']]

        # group by PERNR and fill down the values in the specified columns
        consolidate_filled = consolidate.copy()

        consolidate_filled[cols_to_fill] = (
            consolidate_filled
            .groupby(['PERNR'])[cols_to_fill]
            .transform(lambda g: g.ffill())
        )

        # reset the index of the filled DataFrame
        consolidate_filled.reset_index(drop=True, inplace=True)
        
        # copy final DataFrame for postload
        postload = consolidate_filled.copy()
        # drop unnecessary columns
        postload = postload.drop(columns=['PERNR', 'BEGDA', 'BEGDA_SORT', '0_ROW_NUMBER', 'SEQUENCE'])

        # rename columns to match preload format
        standardized_columns = postload.columns.str.replace(r'\d{4}_', '', regex=True)

        postload.columns = standardized_columns

        postload = postload.add_prefix('postload_')

        # strip whitespace from all object columns
        postload = postload.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        
        return postload, standardized_columns
        
    except Exception as err:
        raise ValueError(f"Error preparing postload data: {err}")
    
def prepare_preload_data(**kwargs)-> pd.DataFrame:
    """
    Prepare Preload Data
    
    Args:
        data_type(str): the expected read in datatype of the dataframe
        encoding(str): encoding code for reading data
        separator: separator of the reading text file 
        kwargs:
            path_preload (str): Path to Preload file.
            preload_columns (list[str]): Columns to select from Preloaddata.
            standardized_columns (list[str])
    
    Returns:
        pd.DataFrame: cleaned and Preload dataframe 
        
    Raise:
        ValueError: if the error when reading text file and unexpect column in the text file
    
    """
    try:
        # prepare data for object PA-0001
        preload  = pd.read_csv(kwargs.get("path_preload"), sep="\t", encoding="utf-8", dtype=str)
        preload.columns = preload.columns.str.strip()
        
        if not set(kwargs.get("columns_preload")).issubset(preload.columns):
            raise ValueError(f"Columns not found, expect {kwargs.get("columns_preload")}")
        preload = preload[kwargs.get("columns_preload")]
        
        # adding prefix to column names
        preload.columns = kwargs.get("standardized_columns")
        preload = preload.add_prefix('preload_')

        # strip whitespace from all object columns
        preload = preload.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        
        return preload
    
    except Exception as err:
        raise ValueError(f"Error preparing preload data: {err}")
    
def create_reconcilation_dataframe(preload: pd.DataFrame, postload: pd.DataFrame, **kwargs) -> pd.DataFrame:
    """
    Full outer join preload data and postload data
    
    Args:
        preload(pd.DataFrame): preload data
        postload(pd.DataFrame): postload data
        kwargs:
            left_on_key (List[str]): Preload key on join
            right_on_key (List[str]): Postload key on join
            join_type (str): join type

    Returns:
        pd.DtaFrame: Reconcilation template ready for export
    
    Raise:
        ValueError: if the error when joining the dataframe
    """
    try:
        reconcile = pd.merge(
            preload, 
            postload, 
            left_on=kwargs.get("left_on_key"), 
            right_on=kwargs.get("right_on_key"),
            how=kwargs.get("join_type")
        )
        
        reconcile['preload_1_SBMOD'] = reconcile['preload_1_WERKS']
        
        print(f"Preload Record: {len(preload)}")
        print(f"Postload Records: {len(postload)}")
        print(f"Reconcilation Records (Record after full join): {len(reconcile)}")
        
        return reconcile
    
    except Exception as err:
        raise ValueError(f"Error create reconcilation dataframe: {err}")

def export_reconcilation_template(preload: pd.DataFrame, postload: pd.DataFrame, reconcile: pd.DataFrame, export_path: Path, split_reconcile: bool=True) -> None:
    """
    Exporting reconcilation template to the folder
    
    Args:
        preload(pd.DataFrame): preload data
        postlaod(pd.DataFrame): postload data
        reconcile(pd.DataFrame): recocile data
        export_path(Path): exporting path
        split_reconcile(bool): if ture will receate extra reconcilation template for action and organization data
    
    Return:
        Excel File: file reconciliation to the desination path
    Raise:
        ValueError: if the error when joining data or exporting file
        
    """
    try:
        if split_reconcile:
            pre_col_0_split = [
                'preload_0_PERNR', 'preload_0_BEGDA', 'preload_0_ENDDA',
                'preload_0_SEQNR', 'preload_0_MASSN', 'preload_0_MASSG',
                'preload_0_STAT2'
            ]

            post_col_0_split = [
                'postload_0_PERNR', 'postload_0_BEGDA', 'postload_0_ENDDA',
                'postload_0_SEQNR', 'postload_0_MASSN', 'postload_0_MASSG',
                'postload_0_STAT2'
            ]
            
            # select only from 0 columns
            preload_0_split = preload[pre_col_0_split]
            postload_0_split = postload[post_col_0_split]

            # full join the preload_0 and postload_0 DataFrames
            reconcile_0_split = pd.merge(
                preload_0_split, postload_0_split, 
                left_on=['preload_0_PERNR', 'preload_0_BEGDA', 'preload_0_ENDDA', 'preload_0_MASSN', 'preload_0_MASSG'], 
                right_on=['postload_0_PERNR', 'postload_0_BEGDA', 'postload_0_ENDDA', 'postload_0_MASSN', 'postload_0_MASSG'],
                how='outer'
            )
            
            pre_col_1_split = [
                'preload_1_PERNR', 'preload_1_BEGDA',
                'preload_1_ENDDA', 'preload_1_BUKRS', 'preload_1_WERKS',
                'preload_1_BTRTL', 'preload_1_KOSTL', 'preload_1_GSBER',
                'preload_1_PERSG', 'preload_1_ABKRS', 'preload_1_PERSK',
                'preload_1_ANSVH', 'preload_1_PROZT', 'preload_1_PLANS',
                'preload_1_STELL', 'preload_1_ORGEH', 'preload_1_SBMOD',
                'preload_1_SACHP', 'preload_1_SACHZ', 'preload_1_SACHA',
                'preload_1_MSTBR'
            ]


            post_col_1_split = [
                'postload_1_PERNR', 'postload_1_BEGDA',
                'postload_1_ENDDA', 'postload_1_BUKRS', 'postload_1_WERKS',
                'postload_1_BTRTL', 'postload_1_KOSTL', 'postload_1_GSBER',
                'postload_1_PERSG', 'postload_1_ABKRS', 'postload_1_PERSK',
                'postload_1_ANSVH', 'postload_1_PROZT', 'postload_1_PLANS',
                'postload_1_STELL', 'postload_1_ORGEH', 'postload_1_SBMOD',
                'postload_1_SACHP', 'postload_1_SACHZ', 'postload_1_SACHA',
                'postload_1_MSTBR'
            ]

            # select only from 1 columns
            preload_1_split = preload[pre_col_1_split]
            postload_1_split = postload[post_col_1_split]

            # full join the preload_1 and postload_1 DataFrames
            reconcile_1_split = pd.merge(
                preload_1_split, postload_1_split, 
                left_on=['preload_1_PERNR', 'preload_1_BEGDA'], 
                right_on=['postload_1_PERNR', 'postload_1_BEGDA'],
                how='outer'
            )
            
            # concat path file
            file_name = r"Recocile_0000_0001_Split.xlsx"
            full_path = Path(export_path) / file_name
            
            with pd.ExcelWriter(path=full_path, engine='openpyxl') as writer:
                reconcile.to_excel(writer, sheet_name='Reconcile_PA-0000', index=False)
                reconcile_0_split.to_excel(writer, sheet_name='Reconcile_PA-0000_Split', index=False)
                reconcile_1_split.to_excel(writer, sheet_name='Reconcile_PA-0001_Split', index=False)
            
        else:
            # concat path file
            file_name = r"Recocile_0000_0001.xlsx"
            full_path = Path(export_path) / file_name
            
            with pd.ExcelWriter(path=full_path, engine='openpyxl') as writer:
                reconcile.to_excel(writer, sheet_name='Reconcile_PA-0000', index=False)
    except Exception as err:
        print(f"Error exporting reconcilation template: {err}")
    
    
# main process of create reconcile template
def main(root_path: Path, postload_file_names: list[str], preload_file_name: str, export_path: Path) -> None:
    """
    Create reconcile template using for reconcile VBA program formatted in Excel file, the function including processes of
        1. Validate in put file path
        2. Prepare postload data
            2.1 Prepare Object PA-0000 (Action) and PA-0302 (Additional action)
            2.2 Prepare Object PA-0001 (Oranization) and HRP-1001-SP (Relationship Between Oranization and Person)
        3. Join the Action data and Ornization data inorder to create Postload Reconcilation
        4. Prepare Preload data
        5. Create Reconcilation Template by join Preload Data and Postload Data
        6. Export the final reconcile dataframe to excel file for reconcilation process
    
    Args:
        root_path(Path): root diractiory of the working file
        postload_file_names(list[str]): list of postload file name
        preload_file_name(str): string of preload file name
    
    Return:
        Excel File: file reconciliation to the desination path
    
    """
    print("Starting Process")
    
    # 1.validate the input file
    print("Process: Validate Input Path")
    preload_path = validate_file_path(root_path, preload_file_name)
    postload_path_list = {
        Path(name).stem.split('_')[1]  # -> "PA-0000"
        : validate_file_path(root_path, name)
        for name in postload_file_names
        if '_' in Path(name).stem
    }
    
    # 2.Prepare postload data
    # 2.1 Prepare Object PA-0000 (Action) and PA-0302 (Additional action)
    print("Process: Prepare Object PA-0000 (Action) and PA-0302 (Additional action)")
    
    pa_0000_columns = ["PERNR" ,"BEGDA" ,"ENDDA" ,"SEQNR" ,"MASSN" ,"MASSG" ,"STAT2"]
    pa_0302_columns = ["PERNR" ,"BEGDA" ,"ENDDA" ,"SEQNR" ,"MASSN" ,"MASSG"]
    
    kwargs_pa_0000_0302 = {
        "path_0000": postload_path_list.get("PA_0000"),
        "path_0302": postload_path_list.get("PA_0302"),
        "columns_0000": pa_0000_columns,
        "columns_0302": pa_0302_columns 
    }
    
    df_action = prepare_data_0000_0302(kwargs=kwargs_pa_0000_0302)
    
    # 2.2 Prepare Object PA-0001 (Organization) and HRP-1001-SP (Relation of Organization and Employee)
    print("Prepare Object PA-0001 (Organization) and HRP-1001-SP (Relation of Organization and Employee)")
    
    pa_0001_columns = [
        "PERNR", "BEGDA", "ENDDA", "BUKRS", "WERKS", 
        "BTRTL", "KOSTL", "GSBER", "PERSG", "ABKRS",
        "PERSK", "ANSVH", "PLANS", "STELL",
        "ORGEH", "SBMOD", "SACHP", "SACHZ", "SACHA", "MSTBR"
    ]

    hrp_1001_sp_columns = [
        'OTYPE', 'OBJID', 'BEGDA', 'ENDDA', 'SOBID', 'PROZT'
    ]
    
    kwargs_pa_0001_1001 = {
        "path_0001": postload_path_list.get("PA_0001"),
        "path_1001": postload_path_list.get("PA_0000"),
        "columns_0001": pa_0001_columns,
        "columns_1001_sp": hrp_1001_sp_columns 
    }
    
    df_organization = prepare_data_0001_hrp_1001_sp(kwargs=kwargs_pa_0001_1001)
    
    # 3. Join the Action data and Ornization data inorder to create Postload Reconcilation
    print("Join the Action data and Ornization data inorder to create Postload Reconcilation")
    postload, standardized_columns = consolidate_postload(df_action, df_organization)
    
    # 4. Prepare Preload data
    print("Prepare Preload data")
    filter_preload_col = [
        'PERNR', 'BEGDA', 'ENDDA', 'SEQNR', 
        'MASSN', 'MASSG', 'STAT2','PA0001.PERNR',
        'PA0001.BEGDA', 'PA0001.ENDDA', 'PA0001.BUKRS', 'TOBE_WERKS',
        'TOBE_BTRTL', 'TOBE_KOSTL', 'PA0001.GSBER', 'TOBE_PERSG',
        'TOBE_ABKRS','TOBE_PERSK', 'PA0001.ANSVH', 'PA0001.PROZT',
        'TOBE_PLANS','PA0001.STELL', 'TOBE_ORGEH', 'PA0001.SBMOD',
        'PA0001.TOBE_SACHP','PA0001.SACHZ', 'PA0001.SACHA', 'PA0001.MSTBR'
    ]
    
    kwargs_preload = {
        "path_preload": preload_path,
        "columns_preload": filter_preload_col,
        "stadardized_column_names": standardized_columns
    }
    
    preload = prepare_preload_data(kwargs_preload)
    
    # 5. Create Reconcilation Template by join Preload Data and Postload Data
    print("Create Reconcilation Template by join Preload Data and Postload Data")
    left_on_key = ['preload_0_PERNR', 'preload_0_BEGDA', 'preload_0_ENDDA', 'preload_0_MASSN', 'preload_0_MASSG', 'preload_1_BEGDA', 'preload_1_ENDDA'], 
    right_on_key = ['postload_0_PERNR', 'postload_0_BEGDA', 'postload_0_ENDDA', 'postload_0_MASSN', 'postload_0_MASSG', 'postload_1_BEGDA', 'postload_1_ENDDA']
    
    kwargs_reconcilation = {
        "left_on_key" : left_on_key,
        "right_on_key" : right_on_key,
        "join_type" : "outer"
    }
    reconcile = create_reconcilation_dataframe(preload, postload, kwargs_reconcilation)
    
    # 6. Export the final reconcile dataframe to excel file for reconcilation process 
    print("Export the final reconcile dataframe to excel file for reconcilation process")
    export_reconcilation_template(preload=preload, postload=postload, reconcile=reconcile, export_path=export_path, split_reconcile=True)
    
    print("Ending Process")


if __name__ == "__main__":
    
    # define root of data source
    ROOT_PATH   = r"C:\Users\wasurat.boonnan\OneDrive - Accenture\Desktop\Data Trnasfomation\Transformed Data\Mock4 Cut Over\M4 Reconcile\0000-0001"
    
    # define export directory
    EXPOT_PATH  =  r"C:\Users\wasurat.boonnan\OneDrive - Accenture\Desktop\Data Trnasfomation\Transformed Data\Mock4 Cut Over\M4 Reconcile\0000-0001"
    
    # define post-load data source 
    PA_0000     = "PostLoad_PA-0000.txt"
    PA_0001     = "PostLoad_PA-0001.txt"
    PA_0302     = "PostLoad_PA-0302.txt"
    HRP_1001_SP = "PostLoad_HRP-1001-SP.txt"
    
    # define pre-load data source
    PRELOAD_FILE_NAME     = "Preload_PA-0000.txt"
    
    # define pre-load data source 
    POSTLOAD_FILE_NAMES = list(PA_0000, PA_0001, PA_0302, HRP_1001_SP)
    
    
    main(ROOT_PATH, POSTLOAD_FILE_NAMES, PRELOAD_FILE_NAME, EXPOT_PATH)

    