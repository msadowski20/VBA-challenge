Attribute VB_Name = "Module1"
Sub stock_checker()

Application.ScreenUpdating = False

Dim stockLookup As Worksheet, lastrowLookup As Long, lookupRange As Range

Set stockLookup = ThisWorkbook.Sheets("Stock Lookup")
Let lastrowLookup = stockLookup.Range("A1").End(xlDown).Row
Set lookupRange = stockLookup.Range("A1:A" & lastrowLookup)

    Dim cell As Range, ws As Worksheet
    
    For Each ws In Worksheets
        
        ws.Activate
        
        If ws.Visible = True Then

            Dim lastrow As Double, summaryRow As Integer, yearlyChange As Double, stockVolume As Double, _
                stockOpen As Double, stockClose As Double, percentChange As Double, visibleLR As Double, visibleFR As Double, _
                dataRange As Range
            
            Let lastrow = ws.Range("A1").End(xlDown).Row
            Set dataRange = ws.Range("A1:G" & lastrow)
            
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"

            summaryRow = 2
            
                For Each cell In lookupRange
                    
                    dataRange.AutoFilter field:=1, Criteria1:=cell
                    
                    If dataRange.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count > 1 Then
                        
                        ws.Range("I" & summaryRow).Value = cell
                        
                        With dataRange.SpecialCells(xlCellTypeVisible)
                            visibleFR = Range("C2:C" & lastrow).SpecialCells(xlCellTypeVisible).Row
                            stockOpen = ws.Range("C" & visibleFR).Value
                            visibleLR = Range("F" & Rows.Count).End(xlUp).Row
                            stockClose = ws.Range("F" & visibleLR).Value
                        End With
                        
                        yearlyChange = stockClose - stockOpen
                        
                            If stockOpen = 0 Then
                                
                                percentChange = "0"
        
                            Else
                                
                                percentChange = (stockClose - stockOpen) / stockOpen
                                
                            End If
                            
                        stockVolume = Application.WorksheetFunction.Sum(ws.Range("G1:G" & lastrow).SpecialCells(xlCellTypeVisible))
                        
                        dataRange.AutoFilter
                        
                        ws.Range("J" & summaryRow).Value = yearlyChange
                        ws.Range("K" & summaryRow).Value = Format(percentChange, "Percent")
                        ws.Range("L" & summaryRow).Value = stockVolume
                        
                        With ws.Range("J" & summaryRow).FormatConditions.Add(xlCellValue, xlLess, "=0")
                            .Interior.ColorIndex = 3
                        End With
    
                        With ws.Range("J" & summaryRow).FormatConditions.Add(xlCellValue, xlGreater, "=0")
                            .Interior.ColorIndex = 4
                        End With
                   
                        summaryRow = summaryRow + 1
                  
                    Else: dataRange.AutoFilter
                    
                    End If
         
                Next cell
           
            Dim summaryLR As Double, changeMax As Double, changeMin As Double, volumeMax As Double, _
                minRange As Variant, maxRange As Variant, volumeRange As Variant

            summaryLR = ws.Range("I1").End(xlDown).Row

            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            
            changeMax = Application.WorksheetFunction.Max(ws.Range("K2:K" & summaryLR))
            changeMin = Application.WorksheetFunction.Min(ws.Range("K2:K" & summaryLR))
            volumeMax = Application.WorksheetFunction.Max(ws.Range("L2:L" & summaryLR))

            ws.Range("Q2").Value = changeMax
            ws.Range("Q3").Value = changeMin
            ws.Range("Q4").Value = volumeMax
   
            maxRange = Application.WorksheetFunction.Match(changeMax, ws.Range("K:K"), 0)
            minRange = Application.WorksheetFunction.Match(changeMin, ws.Range("K:K"), 0)
            volumeRange = Application.WorksheetFunction.Match(volumeMax, ws.Range("L:L"), 0)

            ws.Range("P2") = Cells(maxRange, 9).Value
            ws.Range("P3") = Cells(minRange, 9).Value
            ws.Range("P4") = Cells(volumeRange, 9).Value
   
            ws.Columns("I:Q").AutoFit
            ws.Range("Q2:Q3").NumberFormat = "0.00%"

        End If
    
    Next ws

Application.ScreenUpdating = True

End Sub








