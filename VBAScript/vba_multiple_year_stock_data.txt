Sub MultiYearStockDataSubRoutine():
    
    'Looping through all worksheets in the excel file.
    For Each Sheet In Worksheets
        
        'Declared variables to use in the script.
        Dim TickerLastRow As Long
        Dim YearlyLastClosePrice As Double
        Dim YearlyFirstOpenPrice As Double
        Dim NewTickerLastRow As Long
        Dim NewTickerColumnCount As String
        Dim PercentChange As Double
        Dim TotalStockVolume As Long
        Dim GreatestPercentageIncrease As Double
        Dim GreatestPercentageDecrease As Double
        Dim GreatestTotalVolume As Double
              
        'Ticker column creation - Add column header for Ticker column.
        Sheet.Range("I1").Value = "Ticker"
        'Yearly Change column creation - Add column header for Yearly Change column.
        Sheet.Range("J1").Value = "Yearly Change"
        'Percent Change column creation - Add column header for Percent Change column.
        Sheet.Range("K1").Value = "Percent Change"
        'Total Stock Volume column creation - Add column header for Total Stock Volume.
        Sheet.Range("L1").Value = "Total Stock Volume"
        'Ticker for greatest totals - Add column header for Ticker.
        Sheet.Range("P1").Value = "Ticker"
        'Value for greatest totals - Add column header for value.
        Sheet.Range("Q1").Value = "Value"
        
        'Added row based keys to add greatest totals.
        Sheet.Cells(2, 15).Value = "Greatest % Increase"
        Sheet.Cells(3, 15).Value = "Greatest % Decrease"
        Sheet.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Add default value of the row number after the header.
        NewTickerColumnCount = 2
        j = 2
        
        'Get the last row details from the ticker column.
        TickerLastRow = Sheet.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Perform a loop from second row in the ticker column to the last row.
        For i = 2 To TickerLastRow
            
            'Check if values in preceding column value matches, if not insert the value in new column.
            If Sheet.Cells(i + 1, 1).Value <> Sheet.Cells(i, 1).Value Then
                
                'Get value from the raw ticker column.
                TickerSymbol = Sheet.Cells(i, 1).Value
                
                'Add value in newly created ticker column and format it to $ currency.
                Sheet.Cells(NewTickerColumnCount, 9).Value = Format(TickerSymbol, "$#,##0.00")
                
                'Subtract open value at the beginnning of the year with close value at end of the year.
                 YearlyLastClosePrice = Sheet.Cells(i, 6).Value
                 YearlyFirstOpenPrice = Sheet.Cells(j, 3).Value
                 YearlyChangePrice = YearlyLastClosePrice - YearlyFirstOpenPrice
                 
                 Sheet.Cells(NewTickerColumnCount, 10).Value = YearlyChangePrice
                
                'Add colors to Yearly Change column cells.
                If Sheet.Cells(NewTickerColumnCount, 10).Value > 0 Then
                    'If value is greater than 0, add red else add green for positive value. Index for green is 4 and 3 for red.
                    Sheet.Cells(NewTickerColumnCount, 10).Interior.ColorIndex = 4
                    
                Else
                    Sheet.Cells(NewTickerColumnCount, 10).Interior.ColorIndex = 3
                    
                End If
                
                'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
                PercentChange = YearlyChangePrice / YearlyFirstOpenPrice
                
                Sheet.Cells(NewTickerColumnCount, 11).Value = Format(PercentChange, "Percent")
                
                
                'Add colors to Yearly Change column cells.
                If Sheet.Cells(NewTickerColumnCount, 11).Value > 0 Then
                    'If value is greater than 0, add red else add green for positive value. Index for green is 4 and 3 for red.
                    Sheet.Cells(NewTickerColumnCount, 11).Interior.ColorIndex = 4
                    
                Else
                    Sheet.Cells(NewTickerColumnCount, 11).Interior.ColorIndex = 3
                    
                End If
                'Caluculate the total stock volume and add to the cell
                Sheet.Cells(NewTickerColumnCount, 12).Value = WorksheetFunction.Sum(Range(Sheet.Cells(j, 7), Sheet.Cells(i, 7)))
                
                'Increment the count, so it goes to next cell to add data.
                NewTickerColumnCount = NewTickerColumnCount + 1
                j = i + 1
                
            End If
        'Go to the next row value in the loop.
        Next i
        
        'Get the last row details from the new ticker column.
        NewTickerLastRow = Sheet.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Add default values before the loop.
        GreatestTotalVolume = Sheet.Cells(2, 12).Value
        GreatestPercentageIncrease = Sheet.Cells(2, 11).Value
        GreatestPercentageDecrease = Sheet.Cells(2, 11).Value
        
        'Loop for summary
        For i = 2 To NewTickerLastRow
            
            'Check the next greatest and assign the greatest value.
            If Sheet.Cells(i, 12).Value > GreatestTotalVolume Then
                GreatestTotalVolume = Sheet.Cells(i, 12).Value
                Sheet.Cells(4, 16).Value = Sheet.Cells(i, 9).Value
                
            Else
                GreatestTotalVolume = GreatestTotalVolume
                
            End If
            
            'Check the next greatest percentage increase and assign the greatest value.
            If Sheet.Cells(i, 11).Value > GreatestPercentageIncrease Then
                GreatestPercentageIncrease = Sheet.Cells(i, 11).Value
                Sheet.Cells(2, 16).Value = Sheet.Cells(i, 9).Value
                
            Else
                GreatestPercentageIncrease = GreatestPercentageIncrease
                
            End If
            
            'Check the next greatest percentage decrease and assign the greatest value.
            If Sheet.Cells(i, 11).Value < GreatestPercentageDecrease Then
                GreatestPercentageDecrease = Sheet.Cells(i, 11).Value
                Sheet.Cells(3, 16).Value = Sheet.Cells(i, 9).Value
                
            Else
                GreatestPercentageDecrease = GreatestPercentageDecrease
                
            End If
            
            'Add values to respective cells and format accordingly.
            Sheet.Cells(2, 17).Value = Format(GreatestPercentageIncrease, "Percent")
            Sheet.Cells(3, 17).Value = Format(GreatestPercentageDecrease, "Percent")
            Sheet.Cells(4, 17).Value = Format(GreatestTotalVolume, "Scientific")
            
        'Go to the next row value in the loop.
        Next i
        
        'Fix column and row widths based on values in rows and columns.
        Worksheets(Sheet.Name).Cells.EntireColumn.AutoFit
        Worksheets(Sheet.Name).Cells.EntireRow.AutoFit
        
    'Go to the next sheet in the excel.
    Next Sheet
    
End Sub