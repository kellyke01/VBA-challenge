Attribute VB_Name = "Module1"
Sub Module_hw2():

  

    Dim totalVolume As LongLong
    Dim row As Long
    Dim lastRow As Long
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim summaryTableRow As Double
    Dim firstOpenRow As Long
    Dim findValue As Long
    
    
    For Each ws In Worksheets
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"


        totalStockVolume = 0
        summaryTableRow = 0
        yearlyChange = 0
        firstOpenRow = 2

        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        For row = 2 To lastRow

          If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
        
            totalStockVolume = totalStockVolume + ws.Cells(row, 7).Value

                If totalVolume = 0 Then

                    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                    ws.Range("J" & 2 + summaryTableRow).Value = 0
                    ws.Range("K" & 2 + summaryTableRow).Value = 0 & "%"
                    ws.Range("L" & 2 + summaryTableRow).Value = 0
                Else
           
                    If ws.Cells(firstOpenRow, 3).Value = 0 Then
                        For findValue = firstOpenRow To row
                            If ws.Cells(findValue, 3).Value <> 0 Then
                                firstOpenRow = findValue
                            Exit For
                            End If
                        Next findValue
                    End If

                yearlyChange = ws.Cells(row, 6).Value - ws.Cells(firstOpenRow, 3).Value
                percentChange = yearlyChange / ws.Cells(firstOpenRow, 3).Value


                ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                ws.Range("J" & 2 + summaryTableRow).Value = yearlyChange
                ws.Range("K" & 2 + summaryTableRow).Value = percentChange
                ws.Range("L" & 2 + summaryTableRow).Value = totalVolume
                ws.Range("J" & 2 + summaryTableRow).NumberFormat = "0.00"
                ws.Range("K" & 2 + summaryTableRow).NumberFormat = "0.00%"
                ws.Range("L" & 2 + summaryTableRow).NumberFormat = "#.###"

                'Conditional Formatting for Yearly Change
                If yearlyChange > 0 Then
                     ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4
                ElseIf yearlyChange < 0 Then
                    ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0
                End If
        

                End If
                
            totalVolume = 0
            yearlyChange = 0
            summaryTableRow = summaryTableRow + 1
        Else
            totalVolume = totalVolume + ws.Cells(row, 7).Value
        End If
    

    Next row
    ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastRow)) * 100
    ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastRow)) * 100
    ws.Range("Q4").Value = "%" & WorksheetFunction.Max(ws.Range("L2:L" & lastRow)) * 100
    ws.Range("Q4").NumberFormat = "#.###"

    increaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
    ws.Range("P2").Value = ws.Cells(increaseNumber + 1, 9)
    decreaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
    ws.Range("P3").Value = ws.Cells(decreaseNumber + 1, 9)

    maxTickerVolume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastRow)), ws.Range("L2:L" & lastRow), 0)
    ws.Range("P4").Value = ws.Cells(maxTickerVolume + 1, 9)

      'autofit columns
      ws.Columns("A:Q").Columns.AutoFit

    Next ws
    
End Sub


