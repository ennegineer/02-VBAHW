Attribute VB_Name = "Module1"
Sub HW():
' Use test data to develop scripts
' Run scripts on the stock data to generate final report

' Create a script that will loop through all the stocks for one year and output the following:
' -ticker symbol
' -yearly change from opening price to ending price
' -percent change from opening to closing price
' -total stock volume of the stock
' -highlight positive change in green, negative change in red

' -return the stock with the "Greatest % increase", "Greatest % decrease", "Greatest total volume"

' For submission: screenshot for each year of results on Multi-Year Stock Data
' VBA scripts as separate files

''''''''''''''''''''''''

    Dim Ticker As String
    Dim TotalVolume As LongLong
    Dim YrChange As Double
    Dim PcChange As Double
    Dim SummaryRow As Integer
    Dim StartingRow As Long
        
    For Each ws In Worksheets
    ' set as active worksheet to begin
    ws.Activate
        'set up summary headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    SummaryRow = 2
    StartingRow = 2
    TotalVolume = 0
    
    'find the last row
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        ' Check if we are still within the same ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        ' Set the ticker name
            Ticker = Cells(i, 1).Value
            
            ' find Yearly Change
            YrChange = Cells(i, 6).Value - ws.Cells(StartingRow, 3).Value
            If YrChange > 0 Then
                Range("J" & SummaryRow).Interior.ColorIndex = 10
                Range("J" & SummaryRow).Font.ColorIndex = 2
            ElseIf YrChange < 0 Then
                Range("J" & SummaryRow).Interior.ColorIndex = 9
                Range("J" & SummaryRow).Font.ColorIndex = 2
            End If
            
            ' find Percent Change
            If Cells(StartingRow, 3).Value <> 0 Then
                PcChange = (Cells(i, 6).Value - Cells(StartingRow, 3).Value) / Cells(StartingRow, 3).Value
                Range("K" & SummaryRow).NumberFormat = "0.00%"
            Else
                PcChange = 0
            End If
            
            ' find next StartingRow
            StartingRow = i + 1
            
            ' Add to the total stock volume
            TotalVolume = TotalVolume + Cells(i, 7).Value

            ' Print the ticker and total stock volume to the Summary
            Range("I" & SummaryRow).Value = Ticker
            Range("J" & SummaryRow).Value = YrChange
            Range("K" & SummaryRow).Value = PcChange
            Range("L" & SummaryRow).Value = TotalVolume

            ' Add one to the Summary table row
            SummaryRow = SummaryRow + 1

            ' Set the volume back to zero for the next ticker
            TotalVolume = 0
        Else
            'Cells(1, 8).Value = TotalVolume

            TotalVolume = TotalVolume + Cells(i, 7).Value
        End If
    Next i
    
    'find the summary table length
    SummLength = Cells(Rows.Count, 9).End(xlUp).Row
    
    ' use below to print SummLength to test calculation
       ' Range("R2").Value = SummLength
    
    ' Calculate greatest increase, decrease, total volume
    ' Find the max and min
    Dim Max As Double
    Dim Min As Double
    Dim MaxVol As LongLong
    
    Min = 0
    Max = 0
    MaxVol = 0
    ' Loop through the summary table to find data for the bonus table
    For x = 2 To SummLength
        If Cells(x, 11).Value > Max Then
            Max = Cells(x, 11).Value
            Range("P2").Value = Cells(x, 9).Value
            Range("Q2").Value = Cells(x, 11).Value
        ElseIf Cells(x, 11).Value < Min Then
            Min = Cells(x, 11).Value
            Range("P3").Value = Cells(x, 9).Value
            Range("Q3").Value = Cells(x, 11).Value
        End If
        If Cells(x, 12).Value > MaxVol Then
            MaxVol = Cells(x, 12).Value
            Range("P4").Value = Cells(x, 9).Value
            Range("Q4").Value = Cells(x, 12).Value
        End If
    Next x
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").NumberFormat = "0.00%"
    Columns("I:Q").AutoFit

    Next ws

MsgBox ("Done")

End Sub

