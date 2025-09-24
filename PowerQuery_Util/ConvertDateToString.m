// assuming your column is named [A1]
each
  Text.End   ( Text.From([A1]), 4 ) &
  Text.Middle( Text.From([A1]), 3, 2 ) &
  Text.Start ( Text.From([A1]), 2 )
