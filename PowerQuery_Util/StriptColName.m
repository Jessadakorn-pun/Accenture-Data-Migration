// **Trim the *column names* so " PERNR" becomes "PERNR"**
    #"Trimmed Column Names" =
        Table.TransformColumnNames(
            #"Reordered Columns",
            each Text.Trim(_)
        ),