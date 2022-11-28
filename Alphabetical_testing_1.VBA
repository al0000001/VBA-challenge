'Extract Open and Close
Sub open_close_price()
    For Each ws In Worksheets
        ws.Cells(1, 19).Value = "Ticker"
        ws.Cells(1, 20).Value = "Openprice"
        ws.Cells(1, 21).Value = "Closing price"
        ws.Cells(1, 22).Value = "Yearly Change"
        ws.Cells(1, 23).Value = "Percent Yearly Change"
            Dim calculation_table_row  As Integer
                calculation_table_row = 2
            Dim openprice As Double
                    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                        For i = 2 To lastrow
                            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value And ws.Cells(i, 2).Value = 20200102 Then
                                ws.Range("S" & calculation_table_row).Value = ws.Cells(i, 1).Value
                                ws.Range("T" & calculation_table_row).Value = ws.Cells(i, 3).Value
                                calculation_table_row = calculation_table_row + 1
                                    Else
                                        ws.Range("T" & calculation_table_row).Value = ws.Cells(i, 3).Value
                             End If
                                     If ws.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                                         ws.Range("U" & calculation_table_row - 1).Value = ws.Cells(i, 6).Value
                                         Else
                                             ws.Range("U" & calculation_table_row - 1).Value = ws.Cells(i, 6).Value
                                     End If
                        Next i
    Next ws
End Sub
'Calculate changes
Sub calculation()
    For Each ws In Worksheets
        Dim openprice As Double
        Dim closeprice As Double
        Dim yearly As Double
        Dim percent As Double
        Dim calculation_table_row  As Integer
            calculation_table_row = 2
            lastrow = ws.Cells(Rows.Count, 19).End(xlUp).Row
                For i = 2 To lastrow
                    openprice = ws.Cells(i, 20).Value
                    closeprice = ws.Cells(i, 21).Value
                    yearly = closeprice - openprice
                    percent = (openprice - closeprice) / openprice
                    ws.Range("V" & calculation_table_row).Value = yearly
                    ws.Range("W" & calculation_table_row).Value = percent
                    calculation_table_row = calculation_table_row + 1
                Next i
    Next ws
End Sub

'Creating summary table'
Sub summary_table()
    For Each ws In Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
            Dim ticker As String
            Dim TotalVolume As Double
                TotalVolume = 0
            Dim yearlychange As Double
                yearlychange = 0
            Dim percentchange As Double
                percentchange = 0
            Dim Summary_Table_Row As Integer
                Summary_Table_Row = 2
            Dim openprice As Double
                    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                        For i = 2 To lastrow
                    
                            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                                ticker = ws.Cells(i, 1).Value
                                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                                ws.Range("I" & Summary_Table_Row).Value = ticker
                                ws.Range("L" & Summary_Table_Row).Value = TotalVolume
                                Summary_Table_Row = Summary_Table_Row + 1
                                TotalVolume = 0
                                    Else
                                        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                            End If
                        Next i
     Next ws
End Sub
'Transferring summary
Sub transfer()
    For Each ws In Worksheets
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            For i = 2 To lastrow
                ws.Cells(i, 10).Value = ws.Cells(i, 22).Value
                ws.Cells(i, 11).Value = ws.Cells(i, 23).Value
            Next i
    Next ws
End Sub

'Formating
Sub Condition_formatting()
    For Each ws In Worksheets
      
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            For i = 2 To lastrow
             ws.Cells(i, 11).NumberFormat = "0.00%"
                    If ws.Cells(i, 10).Value >= 0 Then
                        ws.Cells(i, 10).Interior.ColorIndex = 4
                            Else
                                ws.Cells(i, 10).Interior.ColorIndex = 3
                End If
            Next i
    Next ws
End Sub


'Bonus max%increase
Sub bonusmax()
    For Each ws In Worksheets
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        ws.Range("m2").Value = "Greatest % Increase"
        ws.Range("m3").Value = "Greatest % decrease"
        ws.Range("m4").Value = "Greatest Total Value"
        ws.Range("n1").Value = "Ticker"
        ws.Range("o1").Value = "Value"
            Dim max As Double
            Dim ticker As String
                For i = 2 To lastrow
                    If ws.Cells(i, 11).Value > max Then
                        max = ws.Cells(i, 11).Value
                        ticker = ws.Cells(i, 9).Value
                    End If
                Next i
                    ws.Range("o2").Value = max
                    ws.Range("n2").Value = ticker
    Next ws
End Sub

'Max%decrease
Sub bonusmin()
    For Each ws In Worksheets
        Dim min As Double
        Dim ticker As String
            lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
                For i = 2 To lastrow
                    If ws.Cells(i, 11).Value < min Then
                        min = ws.Cells(i, 11).Value
                        ticker = ws.Cells(i, 9).Value
                    End If
                Next i
           ws.Range("o3").Value = min
           ws.Range("n3").Value = ticker
     Next ws
End Sub
'Max total
Sub bonusmaxtotal()
    For Each ws In Worksheets
        Dim maxV As Double
        Dim ticker As String
            lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
                For i = 2 To lastrow
                    If ws.Cells(i, 12).Value > maxV Then
                        maxV = ws.Cells(i, 12).Value
                        ticker = ws.Cells(i, 9).Value
                    End If
                Next i
           ws.Range("o4").Value = maxV
           ws.Range("n4").Value = ticker
    Next ws
End Sub

'Formating 2
Sub Condition_formatting2()
    For Each ws In Worksheets
            For i = 2 To 3
                ws.Cells(i, 15).NumberFormat = "0.00%"
            Next i
    Next ws
End Sub
