'Create a script that loops through all the stocks for one year and outputs the following information:
'
'The ticker symbol
'
'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
'
'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'
'The total stock volume of the stock.
'
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
'
'Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.
'--------------------------------------------------------------------------------------------------------------------------------

Sub stock_data()

Dim ws As Worksheet

'LOOP THROUGH EACH WORKSHEET (for each)
For Each ws In ThisWorkbook.Worksheets
ws.Activate

    'declare variables
    Dim greatesttotalticker, topticker, bottomticker As String
    Dim toptotal, increasechange, decreasechange, totalchange, totalvol, percentagechange, openprice, closeprice, yearlychange As Double
    Dim i, totalrows, outputrowtracker As Integer
    
    
    'some variables need a default value to prevent errors
    
    increasechange = 0
    decreasechange = 0
    outputrowtracker = 2
    totalvol = 0
    
    'add headers to columns where needed
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    'find totalrows
    
    totalrows = Cells(Rows.Count, 1).End(xlUp).Row
    
        'for loop to totalrows
            
            For i = 2 To totalrows
            
            'no need to loop for columns
                
                'add vol to totalvol (needs to be done every time)
                
                totalvol = totalvol + Cells(i, 7).Value
                
                'if ticker is not equal to previous value, we are at the start of that set
                                
                If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                
                    'store value in col 3 as openprice
                
                    openprice = Cells(i, 3).Value
                    
                'if ticker is not equal to the next line, we are at the close of that set
                    
                ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                
                    'paste ticker value into (outputrowtracker, 9)
                    'outputrowtracker is a seperate integer to track the row the total is stored on
                    
                    Cells(outputrowtracker, 9).Value = Cells(i, 1).Value
                    
                    'store value in col 6 as closeprice, this is the final price of the year
                    
                    closeprice = Cells(i, 6).Value
                    
                    'calculate flat difference
                    
                    yearlychange = closeprice - openprice
                    
                    'calculate percent difference
                    
                    percentagechange = (closeprice / openprice) - 1
                    
                        If percentagechange > increasechange Then
                
                            'set increasechange to totalchange
                            'set topticker to col 1
                            
                            increasechange = percentagechange
                            topticker = Cells(i, 1).Value
                        
                            'elseif totalchange is less than than decreasechange
                        
                        ElseIf percentagechange < decreasechange Then
                        
                            'if percentage change is lower than the existing lowest, store new value and ticker
                            
                            decreasechange = percentagechange
                            bottomticker = Cells(i, 1).Value
                            
                            'else move on
                        
                        Else
                        
                        End If
                    
                    'paste totalvol into (outputrowtracker, 12) for Total Stock Volume
                    
                    Cells(outputrowtracker, 12).Value = totalvol
                    
                    'paste percentagechange into (outputrowtracker, 11) for Percent Change
                    
                    Cells(outputrowtracker, 11).Value = percentagechange
                    Cells(outputrowtracker, 11).NumberFormat = "0.00%"
                    
                    'paste yearlychange into (outputrowtracker, 10) for Yearly Change
                    
                    Cells(outputrowtracker, 10).Value = yearlychange
                    
                    'conditional format yearlychange output cell
                    
                    If yearlychange > 0 Then
                        Cells(outputrowtracker, 10).Interior.ColorIndex = 4
                    
                    ElseIf yearlychange < 0 Then
                        Cells(outputrowtracker, 10).Interior.ColorIndex = 3
                        
                    Else
                    
                    End If
                    
                    'add one to outputrowtracker
                    
                    outputrowtracker = outputrowtracker + 1
                
                'else ticker is equal to next ticker value, move along
                Else
                
                End If
            
            'if totalvol is greater than toptotal
            
            If totalvol > toptotal Then
                
                'if the volume is the new highest, store the value and ticker
                
                toptotal = totalvol
                greatesttotalticker = Cells(i, 1).Value
                
            End If
        
        Next i
        
    'Paste values into table now I have compared every row
    
    Cells(2, 16).Value = topticker
    Cells(2, 17).Value = increasechange
    Cells(3, 16).Value = bottomticker
    Cells(3, 17).Value = decreasechange
    Cells(4, 16).Value = greatesttotalticker
    Cells(4, 17).Value = toptotal
    
    'format range (Q2:Q3) to percentage
    
    Range("Q2:Q3").NumberFormat = "0.00%"
    
'move to next worksheet

Next ws

End Sub
