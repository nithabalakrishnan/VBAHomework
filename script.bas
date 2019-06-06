Attribute VB_Name = "Module1"



Sub CalcTotalStockVolume()

    'Variable decalarations
    
    Dim lastRow As Long
    Dim totalVolume As Double
    
    Dim ws_count As Integer
    Dim AW_sheets As Worksheet


   ws_count = ActiveWorkbook.Worksheets.Count
     'Looping through sheets in workbook
        For ws_iterator = 1 To ws_count
            Set AW_sheets = ActiveWorkbook.Worksheets(ws_iterator)
            Call processSheet(AW_sheets)
            AW_sheets.Range("p4").Value = "Ticker"
            AW_sheets.Range("Q4").Value = "Value"
            Call greatestIncrease(AW_sheets)
            Call greatestDecrease(AW_sheets)
            Call greatestVolume(AW_sheets)
        Next ws_iterator
End Sub

Function processSheet(ByRef AW_sheets)
            Dim placeHolder As Integer
            Dim ticker As String
            Dim currentYear As Integer
            Dim openValue As Double
            Dim endValue As Double
            Dim flag As Boolean
            Dim yearlyChangeRange As Range
            Dim percentageRnage As Range
            
            flag = True
            placeHolder = 1
            lastRow = AW_sheets.Cells(Rows.Count, 1).End(xlUp).Row
            
            AW_sheets.Cells(1, 10).Value = "Ticker"
            
            AW_sheets.Cells(1, 11).Value = "Yearly Change"
            Set yearlyChangeRange = AW_sheets.Range("K:K")
            yearlyChangeRange.Cells.NumberFormat = "0.000000"
            
            AW_sheets.Cells(1, 12).Value = "Percentage Change"
            Set percentageRange = AW_sheets.Range("L:L")
            percentageRange.Cells.NumberFormat = "0.00%"
            
            AW_sheets.Cells(1, 13).Value = "Total Volume"
            
            Set positiveConditionalFormatter = yearlyChangeRange.FormatConditions.Add(xlCellValue, xlGreater, 0)
            Set zeroConditionalFormatter = yearlyChangeRange.FormatConditions.Add(xlCellValue, xlEqual, 0)
            Set negativeConditionalFormatter = yearlyChangeRange.FormatConditions.Add(xlCellValue, xlLess, 0)
            
            'Looipng throgh each sheet data
            For iterator = 2 To lastRow
                Dim increase As Double
                increase = 0
                'conditional check for ticker in next cell
                If AW_sheets.Cells(iterator + 1, 1).Value <> AW_sheets.Cells(iterator, 1).Value Then
                    Dim percentIncrease As Double
                
                    placeHolder = placeHolder + 1
                    AW_sheets.Cells(placeHolder, 10).Value = ticker
                    endValue = AW_sheets.Cells(iterator, 6)
                    increase = endValue - openValue
                    totalVolume = totalVolume + AW_sheets.Cells(iterator, 7).Value
                    AW_sheets.Cells(placeHolder, 13).Value = totalVolume
                    AW_sheets.Cells(placeHolder, 11).Value = increase
                    
                    If openValue = 0 Then
                        percentIncrease = -9999999.99999 ' Need to check how to represent infinity for calculations
                    Else
                        percentIncrease = ((endValue - openValue) / openValue)
                    End If
                    If percentIncrease <> -9999999.99999 Then
                        AW_sheets.Cells(placeHolder, 12).Value = percentIncrease
                    End If
                    totalVolume = 0
                    openValue = 0
                   flag = True
                Else
                   endValue = 0
                   If (flag = True) Then
                        openValue = AW_sheets.Cells(iterator, 3).Value
                    End If
                    
                    ticker = AW_sheets.Cells(iterator, 1).Value
                   totalVolume = totalVolume + AW_sheets.Cells(iterator, 7).Value
                   flag = False
               End If
            Next iterator
            
            With positiveConditionalFormatter
                .Interior.Color = vbGreen
                .Font.Color = vbBlack
            End With
            
            With zeroConditionalFormatter
                .Interior.Color = vbGreen
                .Font.Color = vbBlack
            End With
            
            With negativeConditionalFormatter
                .Interior.Color = vbRed
                .Font.Color = vbBlack
            End With
End Function

Function greatestIncrease(ByRef AW_sheets)
    Dim greatestPercentageIncrease As Double
    Dim greatestIncreaseTicker As String
    

    lastRow = AW_sheets.Cells(Rows.Count, 10).End(xlUp).Row
    greatestPercentageIncrease = AW_sheets.Range("L2").Value
    greatestIncreaseTicker = AW_sheets.Range("j2").Value
    For iterator = 2 To lastRow
        nextPercentageValue = AW_sheets.Cells(iterator, 12).Value
        If (greatestPercentageIncrease < nextPercentageValue) Then
            greatestPercentageIncrease = nextPercentageValue
            greatestIncreaseTicker = AW_sheets.Cells(iterator, 10).Value
            
        End If
        
    Next iterator

    AW_sheets.Range("o5").Value = "Greatest % Increase"
    
    AW_sheets.Range("p5").Value = greatestIncreaseTicker
    AW_sheets.Cells(5, 17).NumberFormat = "0.00%"
    AW_sheets.Range("q5").Value = greatestPercentageIncrease
    
    
End Function

Function greatestDecrease(ByRef AW_sheets)
    Dim greatestPercentageDecrease As Double
    Dim greatestDecreaseTicker As String
 

    lastRow = AW_sheets.Cells(Rows.Count, 10).End(xlUp).Row
    greatestPercentageDecrease = AW_sheets.Range("L2").Value
    greatestDecreaseTicker = AW_sheets.Range("j2").Value
    For iterator = 2 To lastRow
        nextPercentageValue = AW_sheets.Cells(iterator, 12).Value
        If (greatestPercentageDecrease > nextPercentageValue) Then
            greatestPercentageDecrease = nextPercentageValue
            greatestDecreaseTicker = AW_sheets.Cells(iterator, 10).Value
        End If
        
    Next iterator

    AW_sheets.Range("o6").Value = "Greatest % Decrease"
    AW_sheets.Range("p6").Value = greatestDecreaseTicker
    AW_sheets.Cells(6, 17).NumberFormat = "0.00%"
    AW_sheets.Range("q6").Value = greatestPercentageDecrease
    
    
End Function

Function greatestVolume(ByRef AW_sheets)
   
    Dim greatestVolumeTicker As String
    Dim greatestVolumeValue As Double

    lastRow = AW_sheets.Cells(Rows.Count, 10).End(xlUp).Row
    greatestVolumeValue = AW_sheets.Range("M2").Value
    greatestVolumeTicker = AW_sheets.Range("j2").Value
    For iterator = 2 To lastRow
        nextVolumeValue = AW_sheets.Cells(iterator, 13).Value
        If (greatestVolumeValue < nextVolumeValue) Then
            greatestVolumeValue = nextVolumeValue
            greatestVolumeTicker = AW_sheets.Cells(iterator, 10).Value
            
            
        End If
        
    Next iterator

    AW_sheets.Range("o7").Value = "Greatest Total Volume"
    AW_sheets.Range("p7").Value = greatestVolumeTicker
    AW_sheets.Range("q7").Value = greatestVolumeValue
    
End Function

