Attribute VB_Name = "Module1"
Sub Stock():

'Find how many worksheets are in workbook and store the value
Dim WS_Count As Long
WS_Count = ActiveWorkbook.Worksheets.Count

'Loop through each sheet of the workbook based on the count of sheets variable
Dim o As Long
For o = 1 To WS_Count
    
    'Store the current worksheet in a variable
    Dim activeWS As Worksheet
    Set activeWS = ActiveWorkbook.Worksheets(o)


    'Find the last row in column A
    Dim lastTickerRow As Long
    lastTickerRow = ActiveWorkbook.Worksheets(o).Cells(Rows.Count, 1).End(xlUp).Row

    'Set column headers
    activeWS.Cells(1, 9) = "<ticker>"
    activeWS.Cells(1, 10) = "Yearly Change"
    activeWS.Cells(1, 11) = "Percent Change"
    activeWS.Cells(1, 12) = "Total Stock Volume"
    
    'Find all of the unique ticker values in Column A and place them in column I
    ActiveWorkbook.Worksheets(o).Range("A1:A" & lastTickerRow).AdvancedFilter , _
    Action:=xlFilterCopy, _
    CopyToRange:=ActiveWorkbook.Worksheets(o).Range("I1"), _
    Unique:=True

    'Find the last row in the unique ticker column
    Dim lastUniqueTickerRow As Long
    lastUniqueTickerRow = ActiveWorkbook.Worksheets(o).Cells(Rows.Count, 9).End(xlUp).Row
    
    'Declare percent change and total volume variables
    'Clear certain variables (set equal to 0) for each loop iteration
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
        currentUniqueTicker = activeWS.Cells(i, 9)
        
        'Declare opening and closing price outside of inner loop
        'Set equal to 0 to clear each item for each loop iteration
        openingPrice = 0
        closingPrice = 0
        yearlyChange = 0
        totalStockVolume = 0
        openingPrice = Null

            'Inner loop to go through each row and calculate the data if the ticker symbols match
            Dim j As Long
            For j = 2 To lastTickerRow
            
                'Stores the current ticker symbol that is being calculated in the loop
                currentRowTicker = activeWS.Cells(j, 1)
                
                'If the current ticker value being calculated matches the row then...
                If (currentUniqueTicker = currentRowTicker) Then
                
                    'Stores the first row value of the open column for the opening price
                    If IsNull(openingPrice) Then
                        openingPrice = activeWS.Cells(j, 3)
                    End If
                
                    'Will store each closing price until the last matching ticker row and the last item
                    closingPrice = activeWS.Cells(j, 6)
                    
                    'Determine total stock volume
                    totalStockVolume = totalStockVolume + activeWS.Cells(j, 7)
                
                End If
            
            Next j
        
        'Calculate the year change and output to appropriate column
        yearlyChange = closingPrice - openingPrice
        activeWS.Cells(i, 10) = yearlyChange
        
        'Format the yearly change column to show postitive values green and negative red
        
        If (yearlyChange > 0) Then
            activeWS.Cells(i, 10).Interior.ColorIndex = 4
        Else
            activeWS.Cells(i, 10).Interior.ColorIndex = 3
        End If
        
        'Determine the percent change
        If (yearlyChange = 0) Then
            percentageChange = 0
        Else
            percentageChange = yearlyChange / openingPrice
        End If
            
        'Output the percent change in the appropriate column and format as percent
        activeWS.Cells(i, 11) = Format(percentageChange, "Percent")

        'Output the total stock volume to the appropriate cell
        activeWS.Cells(i, 12) = totalStockVolume
        
        '---------------------------------------------------------------
        '                        CHALLENGE
        '---------------------------------------------------------------
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
    
    'Output the greatest increase, decrease, and greatest total volume and their values
    
    'Output column headers, rows, and their appropriate values
    activeWS.Cells(1, 16) = "<ticker>"
    activeWS.Cells(1, 17) = "Value"
    activeWS.Cells(2, 15) = "Greatest % Increase"
    activeWS.Cells(2, 16) = greatestPercentIncreaseTicker
    activeWS.Cells(2, 17) = Format(greatestPercentIncrease, "Percent")
    activeWS.Cells(3, 15) = "Greatest % Decrease"
    activeWS.Cells(3, 16) = greatestPercentDecreaseTicker
    activeWS.Cells(3, 17) = Format(greatestPercentDecrease, "Percent")
    activeWS.Cells(4, 15) = "Greatest Total Volume"
    activeWS.Cells(4, 16) = greatestTotalVolumeTicker
    activeWS.Cells(4, 17) = greatestTotalVolume
    
 Next o
    
End Sub
