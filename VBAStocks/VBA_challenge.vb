Sub Ticker_Tracking()

	'Loop through all worksheets
	For each ws in Worksheets

		'Set variable to hold Ticker Symbol
		Dim Ticker_Symbol as String
    
		'Set variable to hold total volume per ticker symbol
		Dim Ticker_Total as Long
		Ticker_Total = 0

		'Track ticker symbol location for summary
		Dim Summary_Table_Row as Integer
		Summary_Table_Row = 2

		'create variables for open, close, yrly change and % change
		Dim year_open as Double
		year_open = ws.Cells(2,3).value
		Dim year_close as Double
		Dim yearly_change as Double
		Dim percent_change as Double

		'Determine last row of sheet
		LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

		'Add column headers for summary to match headers from A1:G1 and add Ticker, Yearly Change, Percent Change and Total Stock Volume
		ws.Cells(1,9).value = "Ticker"
		ws.Cells(1,10).value = "Yearly Change"
		ws.Cells(1,11).value = "Percent Change"
		ws.Cells(1,12).value = "Total Stock Volume"

		'Loop through rows in column
		For i = 2 to LastRow
	
			'search for when value of next cell is different than current cell
			If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

				'set ticker symbol
				Ticker_Symbol = ws.Cells(i, 1).value

				'add to ticker total
				Ticker_Total = Ticker_Total + ws.Cells(i, 7).value

				'set close value
				year_close = ws.Cells(i, 6).value

				'Calculate yearly_change
				yearly_change = year_close - year_open

				'Calculate percent_change
				percent_change = Round((year_close - year_open) / year_open * 100,2)

				'print ticker symbol in summary table
				ws.Range("I" & Summary_Table_Row).value = Ticker_Symbol

				'print yearly_change in summary table
				ws.Range("J" & Summary_Table_Row).value = yearly_change

					'format yearly change column colors
					If yearly_change >= 0 Then
	
						ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

					Else

						ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

					End If

				'print percent_change in summary table
				ws.Range("K" & Summary_Table_Row).value = (percent_change & "%")

				'print ticker total amount in summary table
				ws.Range("L" & Summary_Table_Row).value = Ticker_Total

				'add one to summary table row
				Summary_Table_Row = Summary_Table_Row + 1

				'reset ticker total
				Ticker_Total = 0

				'Get new open price for next ticker
				year_open = ws.Cells(i+1,3).Value
		
			'if cell following row is the same ticker symbol
			Else

				'add to ticker total
				Ticker_Total = Ticker_Total + Cells(i, 7).value
				On Error Resume Next

        	End if


    	Next i

	
	Next ws

End Sub
