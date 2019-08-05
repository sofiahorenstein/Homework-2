Attribute VB_Name = "Module2"
Sub Multiple_year_stock_data():
  
  ' Set an initial variable for holding the ticker
  Dim Ticker As String
  ' Set an initial variable for holding the total per total stock volume for each ticker
  Dim Total_Ticker As Double
  Total_Ticker = 0
  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  ' Loop through all stock volumes
  For i = 2 To 760192
    ' Check if we are still within the same ticker name, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      ' Set the ticker name
      Ticker = Cells(i, 1).Value
      ' Add to the Total stock volume
      Total_Ticker = Total_Ticker + Cells(i, 7).Value
      ' Print the ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker
      ' Print the total stock volume to the Summary Table
      Range("J" & Summary_Table_Row).Value = Total_Ticker
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the ticker total
        Total_Ticker = 0
    ' If the cell immediately following a row has the same ticker name...
    Else
      ' Add to the Total stock volume
     Total_Ticker = Total_Ticker + Cells(i, 7).Value
    End If
  Next i
End Sub
