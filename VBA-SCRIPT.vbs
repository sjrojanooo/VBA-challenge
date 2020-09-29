Option Explicit

Sub StockCalculator()

'Declaring ticker as a string variable
Dim ticker As String

'Declaring i as long variable
Dim i As Long

'Declaring tickCount as a counter vairable
Dim tickCount As Long

'Declaring lastrow as long, wil be used to find the last row in each sheet
Dim lastrow As Long

'Declaring ws as a worksheet, will be used when intitiating my for each loop
'that will loop through each sheet in the workbook
Dim ws As Worksheet

'Declaring openp as a double variable that will hold the stocks open price
Dim openp As Double

'Declaring closep as a double variable that will hold the stocks close price
Dim closep As Double

'Declaring yearlypc as a double variable that will hold my yearly price change value
Dim yearlypc As Double

'Declaring total_volume as AboveAverage Variant to hold the values of the volume of stock trade
'I chose the variant over the long because as a long you tend to get an overflow error
'Longs hold can hold a max value of 20 digit values and volume of stocks traded for some groups exceeded that amount
'Variant have the capability of holding 327 digit value values
Dim total_volume As Variant

'Declaring perc as double to hold the percent change value
Dim perc As Double

'Decaring my totalv as an integer that will be used to store the total_volume values
Dim totalv As Integer

'Declaring greatinc as the greatest percentage increase variable
Dim greatinc As Double

'Declaring greatinc as the greatest percentage decrease variable
Dim greatdec As Double

'Declaring a variable that will hold the greatest volume
Dim greatvol As Variant

'Declaring a string variable that hold the ticker value for the greatest percentage increase value
Dim greatinctic As String

'Declaring a string variable that will hold my great decrease ticker value
Dim greatdectic As String




'Initiating for each loop that will run the script/procedure in every sheet of the worksheet
For Each ws In Worksheets

    'this is activating the variable worksheet varisble ws
    ws.Activate
    
    'Setting lasrow to the initial amount of 0
    lastrow = 0
    
    'Setting openp to an initial amount of 0
    openp = 0
    
    'Setting closep to an initial amount of 0
    closep = "0"
    
    'Setting the counter tickCount to an initial amount of 0
    tickCount = 0
    
    'Setting the string variable to an empty string so that it can take in the value i assign it in the procedure
    ticker = ""
    
    'Variable is set to 2 so that I store my variable in the second row of any given column in the data set
    totalv = "2"
    
    'Wrtiting my header values in the cell address on  each worksheet
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Price Change"
    ws.Cells(1, 11) = "Percentage Change"
    ws.Cells(1, 12) = "Total Volume"
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    
    
    'functions that finds the last row of data in dynamic ranges and storing that value in the lastrow variable
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
        
        'Initiating my for loop that will begin at
        For i = 2 To lastrow
            
            'ticker string variable equal to all string values withint column A or 1, from row 2 to the last row containing data
            ticker = ws.Cells(i, 1).Value
            
            'Since my openp variable is already 0 the if statement will run and store the value of the beginning of the year
            'I will reset this variable back to zero towared the end of this procedure
            If openp = "0" Then
                
                openp = ws.Cells(i, 3).Value
                
            End If
            
            'If statement will initiate once ticker value does not equal to the next rows ticker value
            If ws.Cells(i + 1, 1).Value <> ticker Then
                
                'stores the running total of stocks trading volume for the specific group of stock ticker
                total_volume = total_volume + ws.Cells(i, 7).Value
                
                'places the running sum of total_volume in the address L2
                ws.Range("L" & totalv).Value = total_volume
                
                'Counter variable increments by one each time it runs through the loop
                tickCount = tickCount + 1
                
                'places the ticker string value in the second
                ws.Cells(tickCount + 1, 9).Value = ticker
                
                'close price is equal to the closing stock price at the end of the year
                closep = ws.Cells(i, 6).Value
                
                'yearly price change willl be equal to the open price at the beginnig of the year minues the closing
                'price at the end of the year
                yearlypc = openp - closep
                
                'places the values in the second row of column 10 or ("J")
                ws.Cells(tickCount + 1, 10).Value = yearlypc
                
                'conditional if statement that will shade the color of a cel
                'green for a positive increase
                'red for a negative increase
                'yellow no change
                If yearlypc > 0 Then
                
                    ws.Cells(tickCount + 1, 10).Interior.ColorIndex = 35
                    
                ElseIf yearlypc < 0 Then
                
                    ws.Cells(tickCount + 1, 10).Interior.ColorIndex = 38
                    
                Else
                    
                     ws.Cells(tickCount + 1, 10).Interior.ColorIndex = 34
                     
                End If
                
                'if open price is a value greater than zero then the percentage will be
                'yearly price change divided by the open price
                If openp <> 0 Then
                
                    perc = (yearlypc / openp)
                    
                End If
        
            'add one to the volume summary row
            totalv = totalv + 1
            
            'resets the total volume of stock back to zero
            total_volume = 0
            
            'formats all values in column 11 or ("K")
            ws.Cells(tickCount + 1, 11).Value = Format(perc, "Percent")
            
            'resets the value of my open price to 0
            openp = 0
            
            Else
                
                'add the total volume of stocks within a ticker group
                total_volume = total_volume + ws.Cells(i, 7).Value
                
            'end if statement
            End If
            
            
        
            
        'closes for loop for i
        Next i
        
    'Empty string variable for the greatest increase tic
    greatinctic = ""
    
    'Empty string variable for the greatest increase tic
    greatdectic = ""
    
    'setting my greatest increase percentage equalt to the first percentage value in column 11
    greatinc = ws.Cells(2, 11).Value
    
    'setting my greatest decrease percentage equalt to the first percentage value in column 11
    greatdec = ws.Cells(2, 11).Value
    
    'setting my greatest volume equalt to the first percentage value in column 12
    greatvol = ws.Cells(2, 12).Value
    
        'initiative my for loop that will start at row 2 and finish at the columns last row of data
        For i = 2 To lastrow
            
            'if there is a cell value greater than the variable great vol within the iterating procedure,
            'then greatvol will be replaced with that value.
            If ws.Cells(i, 12).Value > greatvol Then
                
                'string variable will be equal to the same address as the cell value that has the greatest volume.
                greatinctic = ws.Cells(i, 9)
                
                'double variable will be equal to the cell value that has the greatest volume.
                greatvol = ws.Cells(i, 12).Value
                
                'assing the cell location for the double variable with the greatest volume
                ws.Cells(4, 17).Value = greatvol
                
                'assing the cell location for the string variable with the greatest volume
                ws.Cells(4, 16) = greatinctic
                
                'set my string variable back to an empty string.
                greatinctic = ""
                
            End If
            
            'this is the same instance but to find the greatest increase in percentage
            If ws.Cells(i, 11).Value > greatinc Then
                
                    greatinctic = ws.Cells(i, 9)
                    
                    greatinc = ws.Cells(i, 11).Value
                    
                    'formats that cell range to a percentage
                    ws.Cells(2, 17).Value = Format(greatinc, "Percent")

                    ws.Cells(2, 16) = greatinctic
                    
                    greatinctic = ""
                    
            End If
            
            'this is the same instance but to find the greatest increase in percentage
            If ws.Cells(i, 11).Value < greatdec And ws.Cells(i, 11).Value < 0 Then
            
                greatdectic = ws.Cells(i, 9)
                
                greatdec = ws.Cells(i, 11)
                
                'formats that cell range to a percentage
                ws.Cells(3, 17).Value = Format(greatdec, "Percent")
                
                ws.Cells(3, 16) = greatdectic
                
                greatdectic = ""
                
            End If
            
            
        Next i
        
    'Autofits all columns so that each header and value are legible to the user
    ws.Columns("A:Q").AutoFit

    'closes for each loop for my worksheet
    Next ws
    
    
    
'End of procedure
End Sub

