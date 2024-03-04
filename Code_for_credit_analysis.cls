VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CreditCalc()
    
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim TotalStockVolume As Double
    Dim Volume As Double
    Dim row As Integer
    Dim total As Double
    Dim i As Long
    Dim j As Integer
    Dim LastRow As Long
    Dim percentChange As Double
    Dim beginning_of_year As Double
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        YearlyChange = 0
        Volume = 0
        row = 2
        
        For i = 2 To LastRow
        Ticker = ws.Cells(i, 1).Value
        Volume = Volume + ws.Cells(i, 7).Value
            If (ws.Cells(i + 1, 1).Value <> Ticker) Then
                ws.Range("I" & row) = Ticker
                ws.Range("J" & row) = YearlyChange
                ws.Range("K" & row) = PercentageChange
                ws.Range("L" & row) = Volume
                Volume = 0
                row = row + 1
                beginning_of_year = ws.Cells(i + 1, 3).Value
                
            Else
            
                   If (row <= 2) Then
                    YearlyChange = (ws.Cells(i + 1, 6).Value - (ws.Cells(2, 3).Value))
                    PercentageChange = (100 / ws.Cells(2, 3).Value) * YearlyChange & "%"
                Else
                    YearlyChange = (ws.Cells(i + 1, 6).Value - (beginning_of_year))
                    PercentageChange = (100 / beginning_of_year) * YearlyChange & "%"
                End If
            End If
        Next i
        For i = 2 To LastRow
        For j = 10 To 10
                If (ws.Cells(i, j).Value > 0) Then
                    ws.Cells(i, j).Interior.ColorIndex = 4
                ElseIf (ws.Cells(i, j).Value < 0) Then
                    ws.Cells(i, j).Interior.ColorIndex = 3
                Else
                    ws.Cells(i, j).Interior.ColorIndex = 0
                End If
           Next j
           Next i
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & LastRow)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & LastRow)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
        i_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
        d_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
        max_volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & LastRow)), ws.Range("L2:L" & LastRow), 0)
        ws.Range("P2") = ws.Cells(i_number + 1, 9)
        ws.Range("P3") = ws.Cells(d_number + 1, 9)
        ws.Range("P4") = ws.Cells(max_volume + 1, 9)
Next ws
End Sub

