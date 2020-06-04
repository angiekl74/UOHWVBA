Attribute VB_Name = "Module1"
Sub StockSummary():
'This module creates a Summary Table for the following: Yearly Change, Percent Change, Total Stock Volumne for each ticker symbol by year

'Loop through worksheets
Dim ws As Worksheet
For Each ws In Worksheets
    ws.Activate

'Variables for main section of code
    Dim YearChange As Double
    Dim PercentChange As Double
    Dim TotalStkVol As Double
    Dim SummTableRow As Long
    SummTableRow = 2
    
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim StartCell As Double
    Dim EndCell As Double
    Dim StartRow As Double

'Variables to find Min, Max values from summary table
    Min = 10000000
    Max = -1000000
  
   
'Create Column Header names for summary table and Column Header/Row Names for Final Summary Table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
   
    Range("N1").Value = " "
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Decrease"
    Range("N3").Value = "Greatest % Increase"
    Range("N4").Value = "Greatest Total Volume"

'Start the code to summarize and calculate: Yearly Change, Percent Change, Total Stock Volumne
    StartCell = Cells(2, 3).Value
    StartRow = 2
    TotalStkVol = 0

    For i = 2 To (lastrow + 1)

        If Cells(i, 1).Value <> Cells(StartRow, 1).Value Then
   
            TickerName = Cells(i - 1, 1).Value
            TotalStkVol = TotalStkVol
   
            EndCell = Cells(i - 1, 6).Value
           
            YearChange = (EndCell) - (StartCell)
                If StartCell <> 0 Then
                PercentChange = YearChange / StartCell
                Else
                    PercentChange = 0
                End If
           
            StartCell = Cells(i, 3).Value
            StartRow = i
                If PercentChange < Min Then
                    Min = PercentChange
                ElseIf PercentChange > Max Then
                    Max = PercentChange
                End If
               
            'Prints the summary table
            Range("I" & SummTableRow).Value = TickerName
            Range("L" & SummTableRow).Value = TotalStkVol
            Range("J" & SummTableRow).Value = YearChange
            Range("J" & SummTableRow).NumberFormat = "0.00"
                If Range("J" & SummTableRow) >= 0 Then
                    Range("J" & SummTableRow).Interior.ColorIndex = 4
                Else
                    Range("J" & SummTableRow).Interior.ColorIndex = 3
                End If
            Range("K" & SummTableRow).Value = PercentChange
            Range("K" & SummTableRow).NumberFormat = "0.00%"
            SummTableRow = SummTableRow + 1
            TotalStkVol = Cells(i, 7).Value
       
        Else
            TotalStkVol = TotalStkVol + Cells(i, 7).Value
        
        End If
    Next i
   
'print outside the loop the Min, Max, Greatest totalstkVol
    Cells(2, "P").Value = Min
    Cells(2, "P").NumberFormat = "0.00%"
    Cells(2, "O").Value = WorksheetFunction.Index(Range("I2:I1000000"), WorksheetFunction.Match(Range("P2").Value, Range("K2:K1000000"), 0))
             
    Cells(3, "P").Value = Max
    Cells(3, "P").NumberFormat = "0.00%"
    Cells(3, "O").Value = WorksheetFunction.Index(Range("I2:I1000000"), WorksheetFunction.Match(Range("P3").Value, Range("K2:K1000000"), 0))
     
    maxTotalStkVol = WorksheetFunction.Max(ws.Range("L2:L" & SummTableRow))
    Cells(4, "P").Value = maxTotalStkVol
    Cells(4, "O").Value = WorksheetFunction.Index(Range("I2:I1000000"), WorksheetFunction.Match(Range("P4").Value, Range("L2:L1000000"), 0))
    
   
Next
End Sub

