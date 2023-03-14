Attribute VB_Name = "Module1"
Sub annual_report()

'For Each ws In Worksheets

'Declare variables for report
Dim i As Integer
Dim yearlyChange As Double
Dim precentChange As Double
Dim totalVolume As Double
Dim currentTicker As String
Dim openingStock As Double
Dim closingStock As Double
Dim tickerHeader As String
Dim yearlyChangeHeader As String
Dim percentChangeHeader As String
Dim totalStockVolumeHeader As String
Dim greatestIncreaseHeader As String
Dim greatestIdecreaseHeader As String
Dim greatestvolumeHeader As String
Dim valueHeader As String
Dim customIndex As Integer
Dim rowCount As Integer
Dim ws As Worksheet
Dim lastRow As Double

tickerHeader = "Ticker"
yearlyChangeHeader = "Yearly Change"
percentChangeHeader = "Percent Change"
totalStockVolumeHeader = "Total Stock Volume"
greatestIncreaseHeader = "Greatest % increase"
greatestIdecreaseHeader = "Greatest % decrease"
greatestvolumeHeader = "Greatest Total Volume"
valueHeader = "Value"
customIndex = 2



For Each ws In Worksheets

        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        ws.Range("I1").value = tickerHeader
        ws.Range("J1").value = yearlyChangeHeader
        ws.Range("K1").value = percentChangeHeader
        ws.Range("L1").value = totalStockVolumeHeader
        ws.Range("O2").value = greatestIncreaseHeader
        ws.Range("O3").value = greatestIdecreaseHeader
        ws.Range("O4").value = greatestvolumeHeader
        ws.Range("P1").value = tickerHeader
        ws.Range("Q1").value = valueHeader
        
        
        
        For i = 2 To lastRow
                
                    If ws.Cells(i - 1, 1).value <> ws.Cells(i, 1).value Then
                        openingStock = ws.Cells(i, 3).value
                        
                    End If
                    
                    If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
                        
                        currentTicker = ws.Cells(i, 1).value
                        ws.Cells(customIndex, 9) = currentTicker
                        closingStock = ws.Cells(i, 6).value
                        yearlyChange = closingStock - openingStock
                        ws.Cells(customIndex, 10) = yearlyChange
                        percentChange = yearlyChange / openingStock
                        ws.Cells(customIndex, 11) = percentChange
                        totalVolume = totalVolume + ws.Cells(i, 7).value
                        ws.Cells(customIndex, 12) = totalVolume
                        customIndex = customIndex + 1
                        totalVolume = 0
                    
                    Else
        
                        totalVolume = totalVolume + ws.Cells(i, 7).value
                    
                    End If
                    
                    If ws.Cells(customIndex, 10).value < 0# Then
                        ws.Cells(customIndex, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(customIndex, 10).Interior.ColorIndex = 4
                    End If
                    
        Next i
        
        ws.Range("Q2").value = "%" & WorksheetFunction.Max(Range("K2:K" & Rows.Count)) * 100
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & Rows.Count)), Range("K2:K" & Rows.Count), 0)
        ws.Range("P2") = Cells(increase_number + 1, 9)
        ws.Range("Q3").value = "%" & WorksheetFunction.Min(Range("K2:K" & Rows.Count)) * 100
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & Rows.Count)), Range("K2:K" & Rows.Count), 0)
        ws.Range("P3") = Cells(decrease_number + 1, 9)
        ws.Range("Q4").value = WorksheetFunction.Max(Range("L2:L" & Rows.Count))
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & Rows.Count)), Range("L2:L" & Rows.Count), 0)
        ws.Range("P4") = Cells(decrease_number + 1, 9)
        'ws.("A:F").Columns("I:Q").AutoFit

Next

End Sub

Sub Reset()

    ws.Columns("I:Q").Clear
    
    ws.Columns("I:Q").ColumnWidth = 8.43
    
End Sub
