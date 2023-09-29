# VBA-CHALLENGE
ws.Cells(1, 9).Value = "Ticker" - this chunk of coding is saying that "these cells on the worksheet = this value."

Dim lastRow as Long stops the macro at the last row/column on the spreadsheet.

Dim ticker as string, Dim opening/closingPrice as Double, and Dim totalVolume as Double are all declarations of variables in the worksheet.

greatestIncrease declares 0 as the starting point for the stock so that we can calculate the changes from there and indicate the yearly change within each stock.

For i = 2 to lastRow  + If / End If means that a loop was created and stopped after all cells had been completed.

If and ElseIf statements say whether each stock increased or decreased from the start of the year to the end, as well as the percentage gained or lost.

I used outputRow to do conditional formatting on all three spreadsheets based on whether they increased or decreased.

I also used outputRow to gauge the greatest % increase, decrease, and total volume for each stock. The code worked for this and the numbers were correct, but I was unable to put the answers in the correct boxes. They are found starting in row 3003.
