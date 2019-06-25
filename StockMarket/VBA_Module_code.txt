Sub tickers()

    'create variables to run loops and store values
    Dim counter As Integer
    Dim summaryCounter As Integer
    Dim lastSheetRow As Double
    Dim ticVol As Double
    
    Dim priceMod As Double
    Dim pricePercentChange As Double
    Dim ticker As String
    
    Dim greatestIncrease As Double
    Dim giTicker As String
    Dim greatestDecrease As Double
    Dim gdTicker As String
    Dim greatestTotalVol As Double
    Dim gtvTicker As String
    
    'initializing variables
    counter = 2
    summaryCounter = 2
    ticVol = 0
    greatestIncrease = Sheet1.Range("K2").Value
    greatestDecrease = Sheet1.Range("K2").Value
    greatestTotalVol = Sheet1.Range("L2").Value
    giTicker = Sheet1.Range("I2").Value
    gdTicker = Sheet1.Range("I2").Value
    gtvTicker = Sheet1.Range("I2").Value
    
    'creating headers for the columns the ticker summary will go in
    
    'Year
    Sheet1.Range("H1").Value = "Year"
    Sheet1.Range("H1").Font.Bold = True
    Sheet1.Range("H1").WrapText = True
    
    'Ticker Name
    Sheet1.Range("I1").Value = "Ticker"
    Sheet1.Range("I1").Font.Bold = True
    Sheet1.Range("I1").WrapText = True
    
    'Ticker price difference over the course of one year
    Sheet1.Range("J1").Value = "Annual Change In Stock Price"
    Sheet1.Range("J1").Font.Bold = True
    Sheet1.Range("J1").WrapText = True
    
    'Percent change from opening price to closing price
    Sheet1.Range("K1").Value = "Annual Percent Change In Stock Price"
    Sheet1.Range("K1").Font.Bold = True
    Sheet1.Range("K1").WrapText = True
    
    'Ticker total volume
    Sheet1.Range("L1").Value = "Total Stock Volume"
    Sheet1.Range("L1").Font.Bold = True
    Sheet1.Range("L1").WrapText = True
    
    For Each ws In Worksheets
    
    'name a cell with the year being analyzed next to the start of that year's data
    
    Sheet1.Range("H" & counter) = ws.Name
    Sheet1.Range("H" & counter).Font.Bold = True
    
    'determine the length of the data in the sheet for use in the ticker evaluation loop

        lastSheetRow = ws.Range("A" & Rows.Count).End(xlUp).Row
    
    'For loop to calculate the total ticker volume
        For i = 2 To lastSheetRow
            
    'Storing the initial stock value
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                priceMod = ws.Range("C" & i).Value
                
            End If
    
    'evaluating for the case where the cell entry does not match the subsequent entry so as to determine the total volume of one ticker at a time
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'Identifying and storing the name of the ticker group we will be summing
                ticker = ws.Cells(i, 1).Value
    
    'Creating a final total volume for the ticker
                ticVol = ticVol + ws.Cells(i, 7).Value
                
    'Finding the change in stock price for the given ticker
                pricePercentChange = priceMod
                priceMod = ws.Range("F" & i).Value - priceMod
                                
                If priceMod <> 0 And pricePercentChange <> 0 Then
                
                    pricePercentChange = (priceMod / pricePercentChange) * 100
                    
                
                End If
    
    'input the determined ticker
                Sheet1.Range("I" & counter).Value = ticker
                Sheet1.Range("J" & counter).Font.ColorIndex = 1
                
    'format the color of the change in stock price cells positive as green, negative as red
                    If priceMod >= 0 Then
                        
                        Sheet1.Range("J" & counter).Interior.ColorIndex = 4
                    
                    Else
                        
                        Sheet1.Range("J" & counter).Interior.ColorIndex = 3
                        
                    End If
                    
    'input the annual price change
                Sheet1.Range("J" & counter).Value = priceMod
                
    'input the annual percent price change
                Sheet1.Range("K" & counter).Value = pricePercentChange
                Sheet1.Range("K" & counter).NumberFormat = "0.00\%"

    'input the total ticker volume
                Sheet1.Range("L" & counter).Value = ticVol
    
    'Compare new % increase cell with previously greatest generated % increase cell
                If greatestIncrease > Sheet1.Range("K" & counter).Value Then
                    
                    greatestIncrease = greatestIncrease
                    giTicker = giTicker
                
                ElseIf greatestIncrease < Sheet1.Range("K" & counter).Value Then
                    
                    greatestIncrease = Sheet1.Range("K" & counter).Value
                    giTicker = Sheet1.Range("I" & counter).Value
                
                End If
                
    'Compare new % decrease cell with previously greatest generated % decrease cell
                If greatestDecrease < Sheet1.Range("K" & counter).Value Then
                    
                    greatestDecrease = greatestDecrease
                    gdTicker = gdTicker
                
                ElseIf greatestDecrease > Sheet1.Range("K" & counter).Value Then
                    
                    greatestDecrease = Sheet1.Range("K" & counter).Value
                    gdTicker = Sheet1.Range("I" & counter).Value
                
                End If
                
    'Compare new volume cell with previously greatest generated volume cell
                If greatestTotalVol > Sheet1.Range("L" & counter).Value Then
                    
                    greatestTotalVol = greatestTotalVol
                    gtvTicker = gtvTicker
                
                Else
                    
                    greatestTotalVol = Sheet1.Range("L" & counter).Value
                    gtvTicker = Sheet1.Range("I" & counter).Value
                
                End If
        
    'increasing the position-holder for the summary table so the next output is put directly below the most recent entry
                counter = counter + 1
    
    'resetting the ticker values for the next ticker
                ticVol = 0
                priceMod = 0
                pricePercentChange = 0
    
    'evaluating for the case the cell entry and subsequent entry of ticker names do match in order to work towards determining the total volume for that ticker
            Else
        
                ticVol = ticVol + ws.Cells(i, 7).Value
    
            End If
    
        Next i
        
    'return a message box as well as a new summary to let the user know which stock had the greatest % increase, greatest % decrease, and greatest total volume
       
        Sheet1.Range("N" & summaryCounter).Value = "In " & ws.Name & " , the stock with the greatest percent increase of price was: " & giTicker & ":"
        Sheet1.Range("O" & summaryCounter).Value = Round(greatestIncrease, 2) & "%"
        Sheet1.Range("N" & summaryCounter).Font.Bold = True
        Sheet1.Range("N" & summaryCounter).WrapText = True
        Sheet1.Range("O" & summaryCounter).WrapText = True
                
        Sheet1.Range("N" & (summaryCounter + 1)).Value = "In " & ws.Name & " , the stock with the greatest percent decrease of price was: " & gdTicker & ":"
        Sheet1.Range("O" & (summaryCounter + 1)).Value = Round(greatestDecrease, 2) & "%"
        Sheet1.Range("N" & (summaryCounter + 1)).Font.Bold = True
        Sheet1.Range("N" & (summaryCounter + 1)).WrapText = True
        Sheet1.Range("O" & (summaryCounter + 1)).WrapText = True
                
        Sheet1.Range("N" & (summaryCounter + 2)).Value = "In " & ws.Name & " , the stock with the greatest total volume was: " & gtvTicker & ":"
        Sheet1.Range("O" & (summaryCounter + 2)).Value = greatestTotalVol
        Sheet1.Range("N" & (summaryCounter + 2)).Font.Bold = True
        Sheet1.Range("N" & (summaryCounter + 2)).WrapText = True
        Sheet1.Range("O" & (summaryCounter + 2)).WrapText = True
        
    'increasing the position-holder for the annual summary table so the next output is put directly below the most recent entry and resetting the variables for the next count
        summaryCounter = summaryCounter + 3
        greatestIncrease = 0
        greatestDecrease = 0
        greatestTotalVol = 0
        
    Next ws
    
    'adjusting the summary columns to fit the data
    Sheet1.Columns("I:O").AutoFit
    
    
End Sub
