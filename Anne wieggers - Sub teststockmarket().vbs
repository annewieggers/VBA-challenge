Sub teststockmarket()

' Set up summary table
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly change"
Range("K1").Value = "Percent change"
Range("L1").Value = "Total stock volume"

' Set and define a variable for holding the ticker name
Dim ticker As String
Dim ticker_no As Double
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' Set start row and keep track of the location of each ticker in the summary table
Dim Summary_Table_Row As Long
Summary_Table_Row = 2

' Set and define variable for the opening value, closing value, yearly change and percentchange
Dim start As Long
Dim opening As Double
Dim closing As Double
Dim yearlychange As Double
Dim percentchange As Single

' Set and define a variable for the total stock volume
Dim totvol As Double
totvol = 0

'Set up variables and set up table for greatest % increase, decrease and volume
Dim max_increase As Double
Dim max_decrease As Double
Dim max_volume As Double
Dim match1 As Double
Dim match2 As Double
Dim match3 As Double

Range("O1") = "Ticker"
Range("P1") = "Value"
Range("N2").Value = "Greatest % increase"
Range("N3").Value = "Greatest % decrease"
Range("N4").Value = "Greatest total volume"

'Loop through all stock data
For I = 2 To lastrow

' Determine ticker and opening value from the start
        If Cells(I - 1, 1).Value <> Cells(I, 1).Value Then

            ' Set the ticker name
            ticker = Cells(I, 1).Value
            
            ' Print/save the tickername
            Range("I" & Summary_Table_Row).Value = ticker
            
            ' Determine the opening value
            opening = Cells(I, 3).Value
            
        End If

' Check if we are still within the same ticker, if it is not
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
    
    'Determine the yearly change

            ' Determine the closing value
            closing = Cells(I, 6).Value
            
            'Determine the yearly change
            yearlychange = closing - opening
                
            ' Print the yearly change in the summary table
            Range("J" & Summary_Table_Row).Value = yearlychange
        
            ' Format yearlychange to green for a positive change and red for a negative change
        
            If Range("J" & Summary_Table_Row).Value >= 0 Then
               Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
               Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
    ' Determine the percentchange
            
            ' division error fix
            If opening = 0 Then
                Range("K" & Summary_Table_Row).Value = 0
            Else

            ' calculate percent change
                percentchange = ((closing - opening) / opening)
        
            ' Print the percentchange in the Summary Table
                Range("K" & Summary_Table_Row).Value = FormatPercent(percentchange, 2)
                
            End If

            ' Format percentchange to green for a positive change and red for a negative change
             
            If Range("K" & Summary_Table_Row).Value >= 0 Then
                Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            End If

    ' Determine total stock volume per ticker

            ' Add to total stock volume
            totvol = totvol + Cells(I, 7).Value
    
            ' Print the ticker in the Summary Table
            Range("I" & Summary_Table_Row).Value = ticker

            ' Print the totvol to summary table
            Range("L" & Summary_Table_Row).Value = totvol

            ' Reset the totvol
            totvol = 0
            
    ' move to the next ticker
    
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
        
 ' If the cell immediately following a row is the same ticker
 
     Else
            ' Add to total stock volume
            totvol = totvol + Cells(I, 7).Value
            
     End If

Next I
  

' Determine Greatest % increase, greated % decrease and greatest total volume

    ' Greatest % increase
        ' Determine the max increase
        max_increase = WorksheetFunction.Max(Range("K2:k290"))
        ' Determine the position in the list of rows
        match1 = (WorksheetFunction.Match(max_increase, Range("K2:K290"), 0))
        ' Save the value of the tickername and max increase in the summary table
        Range("O2").Value = Cells(match1 + 1, 1).Value
        Range("P2").Value = max_increase
    ' Greatest % decrease
        ' Determine the max decrease
        max_decrease = WorksheetFunction.Min(Range("K2:K290"))
        ' Determine the position of the max decrease in the list of rows
        match2 = (WorksheetFunction.Match(max_decrease, Range("K2:K290"), 0))
        ' Save the value of the tickername and max decrease in the summary table
        Range("O3").Value = Cells(match2 + 1, 1).Value
        Range("P3").Value = max_decrease
    ' Greatest total volume
        ' Determine the max volume
        max_volume = WorksheetFunction.Max(Range("L2:L290"))
        ' Determine the position of the max volume in the list of rows
        match3 = (WorksheetFunction.Match(max_volume, Range("L2:L290"), 0))
        ' Save the value of the tickername and max volume in the summary table
        Range("O4").Value = Cells(match3 + 1, 1).Value
        Range("P4").Value = max_volume


End Sub


