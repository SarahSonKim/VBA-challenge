
Sub WorksheetLoop()

Dim WS_Count As Integer
Dim j As Integer

         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
         For j = 1 To WS_Count
            ActiveWorkbook.Worksheets(j).Activate
            Dim i, stocksum, sumrow As Long
            Dim openingprice, closingprice, yearly, percentage As Double
            Dim currentticker As String
            Dim greatestIncrease, greatestDecrease, greatestVolume As Double
            Dim greatestIncreaseTicker, greatestDecreaseTicker, greatestVolumeTicker As String
    
            stocksum = 0
            sumrow = 2
            greatestIncrease = 0
            greatestDecrease = 0
            greatestVolume = 0
    
            Range("I1").Value = "Ticker"
            Range("J1").Value = "Yearly Change"
            Range("K1").Value = "Percent Change"
            Range("L1").Value = "Total Stock Volume"
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
            Range("O2").Value = "Greatest % Increase"
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatest Total Volume"
    
            For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
                stocksum = stocksum + Cells(i, 7).Value
                currentticker = Cells(i, 1).Value
        
                If currentticker <> Cells(i - 1, 1).Value Then
                    openingprice = Cells(i, 3).Value
                End If
        
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    Cells(sumrow, 9).Value = Cells(i, 1).Value
                    Cells(sumrow, 12).Value = stocksum
                    closingprice = Cells(i, 6).Value
            
                    yearly = Round((openingprice - closingprice), 2)
                    percentage = Round(((yearly / openingprice) * 100), 2)
            
                    Cells(sumrow, 10).NumberFormat = "0.00"
                    Cells(sumrow, 11).NumberFormat = "0.00%"
            
                 
                If yearly > 0 Then
                    Cells(sumrow, 10).Value = "-" & Abs(yearly)
                    Cells(sumrow, 11).Value = "-" & Abs(percentage) & "%"
                    Cells(sumrow, 10).Interior.Color = RGB(255, 0, 0)
                Else
                    Cells(sumrow, 10).Value = Abs(yearly)
                    Cells(sumrow, 11).Value = Abs(percentage) & "%"
                    Cells(sumrow, 10).Interior.Color = RGB(0, 255, 0)
                End If
            
                If greatestIncrease > percentage Then
                    greatestIncrease = percentage
                    greatestIncreaseTicker = currentticker
                End If
            
                If greatestDecrease < percentage Then
                    greatestDecrease = percentage
                    greatestDecreaseTicker = currentticker
                End If
            
                If greatestVolume < stocksum Then
                    greatestVolume = stocksum
                    greatestVolumeTicker = currentticker
                End If
            
                sumrow = sumrow + 1
                stocksum = 0
            End If
        Next i
    
        Range("P2").Value = greatestIncreaseTicker
        Range("Q2").Value = Abs(greatestIncrease) & "%"
        Range("P3").Value = greatestDecreaseTicker
        Range("Q3").Value = "-" & Abs(greatestDecrease) & "%"
        Range("P4").Value = greatestVolumeTicker
        Range("Q4").Value = greatestVolume
        Range("Q4").NumberFormat = "0.00E+00"
        ActiveSheet.Columns.AutoFit

    Next j

End Sub

