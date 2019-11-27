' Constant declaration on Header position
    Const HEADER_POS As Integer = 1
    Const TICKER_POS As Integer = 1
    Const DATE_POS As Integer = TICKER_POS + 1
    Const OPEN_POS As Integer = DATE_POS + 1
    Const HIGH_POS As Integer = OPEN_POS + 1
    Const LOW_POS As Integer = HIGH_POS + 1
    Const CLOSE_POS As Integer = LOW_POS + 1
    Const VOL_POS As Integer = CLOSE_POS + 1
    Const TICKER_SUMM_POS As Integer = VOL_POS + 2
    Const YEARLY_CHANGE_POS As Integer = TICKER_SUMM_POS + 1
    Const PERCENT_CHANGE_POS As Integer = YEARLY_CHANGE_POS + 1
    Const TOTAL_STOCK_VOLUME_POS As Integer = PERCENT_CHANGE_POS + 1
    Const GREATEST_VALUES_POS As Integer = TOTAL_STOCK_VOLUME_POS + 3

Sub VBAHomework():
' Variable declaration
    Dim tickerVal As String
    Dim totalVolume, openPrice, closePrice, deltaYear, deltaPercentage As Double
    Dim lastRow, lastColumn As Integer
    Dim summaryRow, summaryColPosition As Integer
    Dim hasChanged As Integer

For Each currentWorksheet In Worksheets
    currentWorksheet.Activate

    Call LabelColumns(currentWorksheet)
    lastRow = Cells(Rows.Count, HEADER_POS).End(xlUp).Row
    
    ' Initialize variables
    summaryRow = 2
    hasChanged = 1


    For r = (HEADER_POS + 1) To lastRow
        
        If Cells(r, TICKER_POS) <> Cells(r + 1, TICKER_POS) Then
            
            hasChanged = 1
            ' Add the last of its kindto the running total
            totalVolume = totalVolume + Cells(r, VOL_POS).Value
            closePrice = Cells(r, CLOSE_POS).Value
            
            ' Flush all the values to the summary panel
            Cells(summaryRow, TICKER_SUMM_POS) = Cells(r, TICKER_POS)
            Cells(summaryRow, YEARLY_CHANGE_POS) = closePrice - openPrice
            
            'If the closePrice is Zero (division by zero)
            If openPrice <> 0 Then
                Cells(summaryRow, PERCENT_CHANGE_POS) = Round(((closePrice - openPrice) / openPrice), 4)
            Else
                Cells(summaryRow, PERCENT_CHANGE_POS) = 0
            End If
            
            Cells(summaryRow, TOTAL_STOCK_VOLUME_POS) = totalVolume
            summaryRow = summaryRow + 1
            
            ' then reset Volume to Zero
            totalVolume = 0
        Else
            If hasChanged = 1 Then
                openPrice = Cells(r, OPEN_POS).Value
            End If
            totalVolume = totalVolume + Cells(r, VOL_POS).Value
            hasChanged = 0
        End If
    Next r
           
    Call GreatestSummarizer(currentWorksheet, summaryRow)
    Call Formatter(currentWorksheet, summaryRow)

Next currentWorksheet

End Sub



Function LabelColumns(currentWorksheet):
    currentWorksheet.Activate
' Name Headers and Labels
    Cells(HEADER_POS, TICKER_SUMM_POS).Value = "Ticker"
    Cells(HEADER_POS, YEARLY_CHANGE_POS).Value = "Yearly Change"
    Cells(HEADER_POS, PERCENT_CHANGE_POS).Value = "Percent Change"
    Cells(HEADER_POS, TOTAL_STOCK_VOLUME_POS).Value = "Total Stock Volume"
    Cells(HEADER_POS, GREATEST_VALUES_POS + 1).Value = "Ticker"
    Cells(HEADER_POS, GREATEST_VALUES_POS + 2).Value = "Value"
    Cells(HEADER_POS + 1, GREATEST_VALUES_POS).Value = "Greatest % Increase"
    Cells(HEADER_POS + 2, GREATEST_VALUES_POS).Value = "Greatest % Derease"
    Cells(HEADER_POS + 3, GREATEST_VALUES_POS).Value = "Greatest Total Volume"

End Function

Function Formatter(currentWorksheet, summaryRow):
    currentWorksheet.Activate

    With currentWorksheet.Range(Cells(HEADER_POS + 1, YEARLY_CHANGE_POS), Cells(summaryRow, YEARLY_CHANGE_POS)).FormatConditions _
        .Add(xlCellValue, xlGreater, 0)
        .Interior.ColorIndex = 4
    End With
    
     With currentWorksheet.Columns(YEARLY_CHANGE_POS).FormatConditions _
        .Add(xlCellValue, xlLess, 0)
        .Interior.ColorIndex = 3
    End With

    currentWorksheet.Columns(PERCENT_CHANGE_POS).NumberFormat = "0.00%"
    currentWorksheet.Cells(HEADER_POS + 1, GREATEST_VALUES_POS + 2).NumberFormat = "0.00%"
    currentWorksheet.Cells(HEADER_POS + 2, GREATEST_VALUES_POS + 2).NumberFormat = "0.00%"
    currentWorksheet.Columns(TICKER_SUMM_POS).Resize(, GREATEST_VALUES_POS + 2).AutoFit

End Function


Function GreatestSummarizer(currentWorksheet, summaryRow):
    currentWorksheet.Activate
    
    Dim maxIncrease, maxDecrease, maxVolume As Double
    Dim maxIncreaseT, maxDecreaseT, maxVolumeT As String
    
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0
    
    For summR = (HEADER_POS + 1) To summaryRow
            If (Cells(summR, PERCENT_CHANGE_POS).Value > maxIncrease) Then
                    maxIncrease = Cells(summR, PERCENT_CHANGE_POS).Value
                    maxIncreaseT = Cells(summR, TICKER_SUMM_POS).Value
            ElseIf (Cells(summR, PERCENT_CHANGE_POS).Value < maxDecrease) Then
                    maxDecrease = Cells(summR, PERCENT_CHANGE_POS).Value
                    maxDecreaseT = Cells(summR, TICKER_SUMM_POS).Value
            ElseIf (Cells(summR, TOTAL_STOCK_VOLUME_POS).Value > maxVolume) Then
                    maxVolume = Cells(summR, TOTAL_STOCK_VOLUME_POS).Value
                    maxVolumeT = Cells(summR, TICKER_SUMM_POS).Value
            End If
        Next summR
    
        Cells(HEADER_POS + 1, GREATEST_VALUES_POS + 1).Value = maxIncreaseT
        Cells(HEADER_POS + 2, GREATEST_VALUES_POS + 1).Value = maxDecreaseT
        Cells(HEADER_POS + 3, GREATEST_VALUES_POS + 1).Value = maxVolumeT
        Cells(HEADER_POS + 1, GREATEST_VALUES_POS + 2).Value = maxIncrease
        Cells(HEADER_POS + 2, GREATEST_VALUES_POS + 2).Value = maxDecrease
        Cells(HEADER_POS + 3, GREATEST_VALUES_POS + 2).Value = maxVolume

End Function