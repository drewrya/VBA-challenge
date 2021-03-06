Attribute VB_Name = "Module1"
Sub Stock():

    'Find the last row in column A
    Dim lastTickerRow As Long
    lastTickerRow = Cells(Rows.Count, 1).End(xlUp).Row

    'Set column headers
    Cells(1, 9) = "<ticker>"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    
    'Find all of the unique ticker values in Column A and place them in column I
    ActiveSheet.Range("A1:A" & lastTickerRow).AdvancedFilter , _
    Action:=xlFilterCopy, _
    CopyToRange:=ActiveSheet.Range("I1"), _
    Unique:=True

    'Find the last row in the unique ticker column
    Dim lastUniqueTickerRow As Long
    lastUniqueTickerRow = Cells(Rows.Count, 9).End(xlUp).Row
    
    greatestPercentIncrease = 0
    Dim greatestPercentIncreaseTicker
    greatestPercentDecrease = 0
    Dim greatestPercentDecreaseTicker
    greatestTotalVolume = 0
    Dim greatestTotalVolumeTicker
    
    'Outer loop to go through each unique ticker symbol
    Dim i As Long
    For i = 2 To lastUniqueTickerRow
        
        'Stores the current ticker symbol that is being calculated in the loop
        currentUniqueTicker = Cells(i, 9)
        
        'Declare opening and closing price outside of inner loop as 0 to begin
        openingPrice = 0
        closingPrice = 0
        yearlyChange = 0
        totalStockVolume = 0
        openingPrice = Null
        
            'Inner loop to go through each row and calculate the data if the symbols match
            Dim j As Long
            For j = 2 To lastTickerRow
            
                
                'Stores the current ticker symbol that is being calculated in the loop
                currentRowTicker = Cells(j, 1)
                
                'If the current ticker value being calculated matches the row then...
                If (currentUniqueTicker = currentRowTicker) Then
                
                    'Stores the first row value of the open column for the opening price
                    If IsNull(openingPrice) Then
                        openingPrice = Cells(j, 3)
                    End If
                
                    'Will store each closing price until the last matching ticker row and the last item
                    closingPrice = Cells(j, 6)
                    
                    'Determine total stock volume
                    totalStockVolume = totalStockVolume + Cells(j, 7)
                
                End If
            
            Next j
        
        'Calculate the year change and output to appropriate column
        yearlyChange = closingPrice - openingPrice
        Cells(i, 10) = yearlyChange
        
        'Format the yearly change column to show postitive values green and negative red
        
        If (yearlyChange > 0) Then
            Cells(i, 10).Interior.ColorIndex = 4
        Else
            Cells(i, 10).Interior.ColorIndex = 3
        End If
        
        'Determine the percent change
        percentageChange = yearlyChange / openingPrice
        
        'Output the percent change in the appropriate column and format as percent
        Cells(i, 11) = Format(percentageChange, "Percent")

        'Output the total stock volume to the appropriate cell
        Cells(i, 12) = totalStockVolume
        
        '---------------------------------------------
        '               CHALLENGE
        '---------------------------------------------
        
        'Store the percentage change in a variable to track the lowest and highest
        If (percentageChange > greatestPercentIncrease) Then
            greatestPercentIncrease = percentageChange
            greatestPercentIncreaseTicker = currentUniqueTicker
        End If
        
        If (percentageChange < greatestPercentDecrease) Then
            greatestPercentDecrease = percentageChange
            greatestPercentDecreaseTicker = currentUniqueTicker
        End If
        
        If (totalStockVolume > greatestTotalVolume) Then
            greatestTotalVolume = totalStockVolume
            greatestTotalVolumeTicker = currentUniqueTicker
        End If
        
    Next i
    
    'Output the greatest increase and decrease and their values
    'Set column headers and row header
    Cells(1, 16) = "<ticker>"
    Cells(1, 17) = "Value"
    Cells(2, 15) = "Greatest % Increase"
    Cells(2, 16) = greatestPercentIncreaseTicker
    Cells(2, 17) = Format(greatestPercentIncrease, "Percent")
    Cells(3, 15) = "Greatest % Decrease"
    Cells(3, 16) = greatestPercentDecreaseTicker
    Cells(3, 17) = Format(greatestPercentDecrease, "Percent")
    Cells(4, 15) = "Greatest Total Volume"
    Cells(4, 16) = greatestTotalVolumeTicker
    Cells(4, 17) = greatestTotalVolume
    
End Sub
