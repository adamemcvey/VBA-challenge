Sub stockmarket()
                                                                                                                        #create loop to go through each worksheet in the workbook (every function will now have to feature ws)
For Each ws In Worksheets

#create variables for the name of each stock, the amount and percent it changed, and its total volum at the end of the >Dim ticker As String                                                                                                    Dim yearlyChange As Double
Dim percentChange As Double
Dim volume As Double
    volume = 0                                                                                                          
#create a placeholder for the row that the current ticker is on
Dim tickerRow As Integer
    tickerRow = 2
                                                                                                                        #create a variable to hold the opening price on opening day, as this is the only opening price we'll need               Dim openingPrice As Double
    openingPrice = ws.Cells(2, 3).Value
                                                                                                                        #the closing price will be determined when a new ticker is found
Dim closingPrice As Double

#create the headers to the columns
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

#create a placeholder for the last row on any given worksheet
last = ws.Cells(Rows.Count, 1).End(xlUp).Row

#for loop that will go through the entire worksheet
For i = 2 To last

#checks to see if the cell after the current cell has a different ticker, if it does:
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        #copy the new ticker into the ticker value and print it on the next cell in the results section
            ticker = ws.Cells(i, 1).Value
            ws.Range("I" & tickerRow).Value = ticker
            #update the closing price
            closingPrice = ws.Cells(i, 6).Value
            #calculate the yearly change based on the opening price subtracted from the new opening price, then is prin>            yearlyChange = (closingPrice - openingPrice)
                ws.Range("J" & tickerRow).Value = yearlyChange
        #a quick check to see if opening price is 0, which would make division impossible: if so, it is automatically s>            If (openingPrice = 0) Then
                percentChange = 0

            Else
        #otherwise the percent change is calculated as the yearly change divided by the opening price
                percentChange = yearlyChange / openingPrice

            End If
        #and that value is printed in the results section and set to be formatted and displayed as a percentage
                ws.Range("K" & tickerRow).Value = percentChange
                ws.Range("K" & tickerRow).NumberFormat = "0.00%"

            #the volume of the stock in the given cell is added to the volume that already exists, and is printed in th>            volume = volume + ws.Cells(i, 7).Value
                ws.Range("L" & tickerRow).Value = volume

        #change the placeholder values by moving the ticker row and the opening price down one row, and resetting the v>            tickerRow = tickerRow + 1
            volume = 0
            openingPrice = ws.Cells(i + 1, 3)

#if the two cells have the same ticker, the only thing that is changed is that the volume is added to the value already>    Else
        volume = volume + ws.Cells(i, 7).Value


    End If

Next i

#create a second placeholder for the last row in the worksheet, this time on column I
last2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

#loop through this column and use conditional formatting to fill the color of the cell, red if negative, green if posit>For i = 2 To last2
#the conditional fomatting must be done as its own loop, apparently the color of each cell cannot be successfully changed in the first for loop
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4

    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3

    End If

Next i
#create the header columns for the challenge portion of the assignment
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

#use the new last row value to loop through the yearly change column again
For i = 2 To last2
#using the worksheet function max to look at the entire column and select the given row if it is the max value, and then print the corresponding value from the percent change column, formatting it as a percentage of course
    If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & last2)) Then
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
        ws.Cells(2, 17).NumberFormat = "0.00%"
#use the similar worksheet function min to look for the minimum yearly change value
    ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & last2)) Then
#print the corresponding percent change in the challenge section as a percentage
        ws.Cells(3, 16).Value = Cells(i, 9).Value
        ws.Cells(3, 17).Value = Cells(i, 11).Value
        ws.Cells(3, 17).NumberFormat = "0.00%"
#use the worksheet function max to look through the total volume column and select the highest one and print it in the challenge section         
    ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & last2)) Then
        ws.Cells(4, 16).Value = Cells(i, 9).Value
        ws.Cells(4, 17).Value = Cells(i, 12).Value
        
    End If
    
Next i
        
Next ws
    
End Sub

