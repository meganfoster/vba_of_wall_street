Attribute VB_Name = "Module1"
Sub ticker()

    'Loop through all worksheets
    For Each ws In Worksheets
    
        ' Set Column Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        ' Create Variables
        Dim rowcount As Integer
        rowcount = 1
        Dim openvalue As Double
        Dim closevalue As Double
        Dim difference As Double
        Dim ticker As String
        Dim tsv As Double
        tsv = 0
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Identify ticker, open value, close value, and percent change using for loop
        Dim i As Long
        For i = 2 To lastrow
        
            'Checks the first line of the group
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                openvalue = ws.Cells(i, 3).Value
                tsv = tsv + ws.Cells(i, 7).Value
            'Checks the last line of the group
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                rowcount = rowcount + 1
                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
                ticker = ws.Cells(rowcount, 9).Value
                closevalue = ws.Cells(i, 6).Value
                ws.Cells(rowcount, 10).Value = closevalue - openvalue
                difference = ws.Cells(rowcount, 10).Value
                'Format difference as red or green
                    If difference > 0 Then
                        ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                    ElseIf difference < 0 Then
                        ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                    End If
                'Check for 0 in denominator
                    If openvalue = 0 Then
                        ws.Cells(rowcount, 11).Value = 0
                    Else
                        ws.Cells(rowcount, 11).Value = difference / openvalue
                    End If
                'Cells(rowcount, 11).Style = "Percent"
                ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                tsv = tsv + ws.Cells(i, 7).Value
                ws.Cells(rowcount, 12).Value = tsv
                tsv = 0
            'Checks all lines in between of group
            Else
                tsv = tsv + ws.Cells(i, 7).Value
    
            End If
    
        Next i
        
        'Find the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
        'Set Challenge Chart
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        'Set Percent Range
        Dim percentrange As Range
        Set percentrange = ws.Range("K2:K" & rowcount)
        'Find max
        Dim percentmax As Double
        percentmax = WorksheetFunction.Max(percentrange)
        'Find min
        Dim percentmin As Double
        percentmin = WorksheetFunction.Min(percentrange)
        'Set TSV Range
        Dim tsvrange As Range
        Set tsvrange = ws.Range("L2:L" & rowcount)
        'Find max
        Dim tsvmax As Double
        tsvmax = WorksheetFunction.Max(tsvrange)
        'Run Loop to fill in challenge chart
         For i = 2 To rowcount
            If ws.Cells(i, 11).Value = percentmax Then
               ws.Range("P2").Value = ws.Cells(i, 9).Value
               ws.Range("Q2").Value = percentmax
            ElseIf ws.Cells(i, 11).Value = percentmin Then
               ws.Range("P3").Value = ws.Cells(i, 9).Value
               ws.Range("Q3").Value = percentmin
            ElseIf ws.Cells(i, 12).Value = tsvmax Then
               ws.Range("P4").Value = ws.Cells(i, 9).Value
               ws.Range("Q4").Value = tsvmax
            End If
            
        Next i
        
        ws.Columns("A:Q").AutoFit
        
    Next ws
    
End Sub
