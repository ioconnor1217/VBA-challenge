Sub QuarterlyAnalysis()

    ' Declare variables
    Dim ws As Worksheet
    Dim i As Long
    Dim Ticker As String
    Dim CurrentTicker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim QuarterlyChange As Double
    Dim PercentageChange As Double
    Dim TotalVolume As Double
    Dim OutputRow As Long
    Dim DateValue As Double
    Dim VolumeValue As Double
    Dim RowCount As Long
    Dim GreatestIncrease As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecrease As Double
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolume As Double
    Dim GreatestVolumeTicker As String
    
    ' Initialize variables
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestVolume = 0
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "Q*" Then
        
            ' Clear previous output
            ws.Range("H1:P1").ClearContents

            ' Dynamically set row count
            RowCount = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

            ' Set output row
            OutputRow = 2

            ' Initialize variables
            TotalVolume = 0
            CurrentTicker = ""

            ' Loop through each row
            For i = 2 To RowCount

                ' Get current row
                Ticker = ws.Cells(i, 1).Value
                OpenPriceCell = ws.Cells(i, 3).Value
                ClosePriceCell = ws.Cells(i, 6).Value
                VolumeValue = ws.Cells(i, 7).Value

                ' Check for new ticker
                If Ticker <> CurrentTicker Then
                    
                    ' Else calculate previous ticker results
                    If CurrentTicker <> "" Then
                        QuarterlyChange = ClosePrice - OpenPrice
                        
                        ' Calculate PercentageChange
                        PercentageChange = Round(QuarterlyChange / OpenPrice, 4)

                        ' Output result to columns I:L
                        ws.Cells(1, 9).Value = "Ticker"
                        ws.Cells(1, 10).Value = "Quarterly Change"
                        ws.Cells(1, 11).Value = "Percent Change"
                        ws.Cells(1, 12).Value = "Total Volume"
                        
                        ws.Cells(OutputRow, 9).Value = CurrentTicker
                        ws.Cells(OutputRow, 10).Value = QuarterlyChange
                        ws.Cells(OutputRow, 11).Value = PercentageChange
                        ws.Cells(OutputRow, 12).Value = TotalVolume

                        ' Format PercentChange as percent
                        ws.Cells(OutputRow, 11).NumberFormat = "0.00%"

                        ' Track increase, decrease, and volume
                        If PercentageChange > GreatestIncrease Then
                            GreatestIncrease = PercentageChange
                            GreatestIncreaseTicker = CurrentTicker
                        End If
                        If PercentageChange < GreatestDecrease Then
                            GreatestDecrease = PercentageChange
                            GreatestDecreaseTicker = CurrentTicker
                        End If
                        If TotalVolume > GreatestVolume Then
                            GreatestVolume = TotalVolume
                            GreatestVolumeTicker = CurrentTicker
                        End If

                        ' Move to next output row
                        OutputRow = OutputRow + 1
                    End If

                    ' Update ticker
                    CurrentTicker = Ticker
                    OpenPrice = OpenPriceCell
                    TotalVolume = 0
                End If

                ' Update ClosePrice
                ClosePrice = ClosePriceCell

                ' Accumulate TotalVolume
                If IsNumeric(VolumeValue) Then
                    TotalVolume = TotalVolume + VolumeValue
                End If

            Next i

            ' QuarterlyChange calculation
            If CurrentTicker <> "" Then
                QuarterlyChange = ClosePrice - OpenPrice
                
                ' Calculate PercentageChange
                PercentageChange = Round(QuarterlyChange / OpenPrice, 4)

                ' Output final results
                ws.Cells(OutputRow, 9).Value = CurrentTicker
                ws.Cells(OutputRow, 10).Value = QuarterlyChange
                ws.Cells(OutputRow, 11).Value = PercentageChange
                ws.Cells(OutputRow, 12).Value = TotalVolume

                ' Format PercentChange as percent
                ws.Cells(OutputRow, 11).NumberFormat = "0.00%"

                ' Calculate greatest increase, decrease, and volume
                If PercentageChange > GreatestIncrease Then
                    GreatestIncrease = PercentageChange
                    GreatestIncreaseTicker = CurrentTicker
                End If
                If PercentageChange < GreatestDecrease Then
                    GreatestDecrease = PercentageChange
                    GreatestDecreaseTicker = CurrentTicker
                End If
                If TotalVolume > GreatestVolume Then
                    GreatestVolume = TotalVolume
                    GreatestVolumeTicker = CurrentTicker
                End If
                
                ' Output greatest increase, decrease, and volume
                ws.Cells(1, 15).Value = "Ticker"
                ws.Cells(1, 16).Value = "Value"
                
                ws.Cells(2, 14).Value = "Greatest % Increase"
                ws.Cells(3, 14).Value = "Greatest % Decrease"
                ws.Cells(4, 14).Value = "Greatest Volume"
    
                ws.Cells(2, 15).Value = GreatestIncreaseTicker
                ws.Cells(3, 15).Value = GreatestDecreaseTicker
                ws.Cells(4, 15).Value = GreatestVolumeTicker
                
                ' Format increase/decrease as percent
                ws.Cells(2, 16).Value = GreatestIncrease
                ws.Cells(2, 16).NumberFormat = "0.00%"
                
                ws.Cells(3, 16).Value = GreatestDecrease
                ws.Cells(3, 16).NumberFormat = "0.00%"
                
                ws.Cells(4, 16).Value = GreatestVolume
            End If

        End If
        
    Next ws
    
    MsgBox "Completed successfully"

End Sub

