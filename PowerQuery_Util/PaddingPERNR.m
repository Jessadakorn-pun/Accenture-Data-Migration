// **Now pad PERNR (no leading space) to length 8**
    #"Padded to 8" =
        Table.TransformColumns(
            #"Trimmed Column Names",
            {
                "PERNR",
                each Text.PadStart(_, 8, "0"),
                type text
            }
        )

Text.PadStart(Text.From([A1]), 8, "0")