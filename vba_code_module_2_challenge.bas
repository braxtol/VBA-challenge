Attribute VB_Name = "Module1"
Sub smartAs()
 
Dim ticker As String
Dim finish As Double
Dim finishvol As Double
Dim begin As Double
Dim beginvol As Double
Dim yearchange As Double
Dim percentchange As Double
Dim totalvol As Double
Dim Last_Row As Long
Dim Summary_Table_Row As Integer
Dim tickerincrease As Double
Dim tickerdecrease As Double
Dim tickergreatestvol As Double
Dim ws As Worksheet



'initialize variables
Summary_Table_Row = 2
totalvol = 0
Range("I1") = "Ticker"
Range("J1") = "Yearly Change ($)"
Range("k1") = "% Change"
Range("L1") = "Total Stock Volume"
Columns(9).AutoFit
Columns(10).AutoFit
Columns(11).AutoFit
Columns(15).AutoFit
Columns(16).AutoFit
Columns(17).AutoFit
Columns("J").ColumnWidth = 13
Columns("K").ColumnWidth = 8
Columns("L").ColumnWidth = 15
Columns("O").ColumnWidth = 17
Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = "Greatest Total Volume"
Range("P1") = "Ticker"
Range("Q1") = "Value"

begin = Cells(2, 3).Value
  'Loop through rows in the column to get it done.
Last_Row = Cells(Rows.Count, 1).End(xlUp).Row

For Each ws In Worksheets

For i = 2 To Last_Row

     If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            ticker = Cells(i, 1).Value
            totalvol = totalvol + Cells(i, 7).Value
            yearchange = Cells(i, 6).Value - begin
            percentchange = (yearchange / Cells(i, 6).Value)
            
            Range("I" & Summary_Table_Row).Value = ticker
            Range("J" & Summary_Table_Row).Value = yearchange
                    If Range("J" & Summary_Table_Row).Value > 0 Then
                            Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    Else
                            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
            Range("K" & Summary_Table_Row).Value = percentchange
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            Range("L" & Summary_Table_Row).Value = totalvol
           
           'reset variables
           totalvol = 0
           begin = Cells(i + 1, 3).Value
           Summary_Table_Row = Summary_Table_Row + 1
       Else
            totalvol = totalvol + Cells(i, 7).Value
       End If
   
   Next i
       
   Range("Q2").Value = WorksheetFunction.Max(Range("K2:K301"))
   Range("Q2").NumberFormat = "0.00%"
   tickerincrease = WorksheetFunction.Match(Range("Q2").Value, Range("K2:K301"), 0)
   Range("P2").Value = Range("I" & tickerincrease + 1).Value
   
   Range("Q3").Value = WorksheetFunction.Min(Range("K2:K301"))
   Range("Q3").NumberFormat = "0.00%"
   tickerdecrease = WorksheetFunction.Match(Range("Q3").Value, Range("K2:K301"), 0)
   Range("P3").Value = Range("I" & tickerdecrease + 1).Value
   
   Range("Q4").Value = WorksheetFunction.Max(Range("L2:L301"))
   tickergreatestvol = WorksheetFunction.Match(Range("Q4").Value, Range("L2:L301"), 0)
   Range("P4").Value = Range("I" & tickergreatestvol + 1).Value

  Next ws

End Sub


