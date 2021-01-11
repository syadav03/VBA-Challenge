Sub Stocks()



'  Define Variables
Dim Ticker_Symbol

Dim Yearly_Change

Dim Percent_Change

Dim Total_Stock_Volume As Double

Dim Open_Price

Dim Close_Price


Dim The_Table_Row

Dim ws_Total As Integer

Dim GreatestIncTicker
GreatestIncTicker = " "
Dim GreatestDecTicker
GreatestDecTicker = " "
Dim GreatestVolTicker
GreatestVolTicker = " "
Dim GreatestIncValue
GreatestIncValue = 0
Dim GreatestDecValue
GreatestDecValue = 0
Dim GreatestVolValue
GreatestVolValue = 0



'Variable to loop through worksheets
ws_Total = ActiveWorkbook.Worksheets.Count

'Begin Worksheet loop
For i = 1 To ws_Total
    ' Creating Headers for Columns 
    Worksheets(i).Range("I1") = "Ticker"
    Worksheets(i).Range("J1") = "Yearly Change"
    Worksheets(i).Range("K1") = "Percent Change"
    Worksheets(i).Range("L1") = "Total Stock Volume"
    
    
     Worksheets(i).Range("P1") = "Ticker"
     Worksheets(i).Range("Q1") = "Value"
     Worksheets(i).Range("O2") = "Greatest % Increase"
     Worksheets(i).Range("O3") = "Greatest % Decrease"
     Worksheets(i).Range("O4") = "Greatest Total Volume"
     Worksheets(i).Range("Q2:Q3").NumberFormat = "0.00%"
     BottomRow = Worksheets(i).Cells(Rows.Count, 1).End(xlUp).Row
    The_Table_Row = 2
    ' Begin Loops for Rows 
    For x = 2 To BottomRow
        ' set ticker 
        Ticker_Symbol = Worksheets(i).Cells(x, 1)
        Total_Stock_Volume = Total_Stock_Volume + Cells(x, 7)
        
        ' Opening Price
        If Open_Price = "" Then
            Open_Price = Worksheets(i).Cells(x, 3)
        End If
        
        
        If Ticker_Symbol <> Worksheets(i).Cells((x + 1), 1) Then
        
        'Closing  price
        Close_Price = Worksheets(i).Cells(x, 6)
        'Yearly  change
        Yearly_Change = Open_Price - Close_Price
        
        ' Ticker_Symbol 
        Worksheets(i).Range("I" & The_Table_Row).Value = Ticker_Symbol
        
        
        Worksheets(i).Range("J" & The_Table_Row).Value = Yearly_Change
        If Yearly_Change < 0 Then
	' Making Cell Red if negative 
        Worksheets(i).Range("J" & The_Table_Row).Interior.ColorIndex = 3 
        Else
	'Green if Positive 
        Worksheets(i).Range("J" & The_Table_Row).Interior.ColorIndex = 4 
        End If
        
        
        If BeginPrice <> Close_Price Then
            Percent_Change = Yearly_Change / Close_Price
        Else
            Percent_Change = 0
        End If
        Worksheets(i).Range("K" & The_Table_Row).Value = Percent_Change
        Worksheets(i).Range("K" & The_Table_Row).NumberFormat = "0.00%"
        
        
        Worksheets(i).Range("L" & The_Table_Row).Value = Total_Stock_Volume
        
        
        
        If Total_Stock_Volume > GreatestVolValue Then
        GreatestVolValue = Total_Stock_Volume
        GreatestVolTicker = Ticker_Symbol
        End If
        
        
        The_Table_Row = The_Table_Row + 1
        Total_Stock_Volume = 0
        
        End If
        
        
        If Percent_Change > GreatestIncValue Then
        GreatestIncValue = Percent_Change
        GreatestIncTicker = Ticker_Symbol
        ElseIf Percent_Change < GreatestDecValue Then
        GreatestDecValue = Percent_Change
        GreatestDecTicker = Ticker_Symbol
        End If
    
    Next x
    
    
    Worksheets(i).Range("P2") = GreatestIncTicker
    Worksheets(i).Range("Q2") = GreatestIncValue
    Worksheets(i).Range("P3") = GreatestDecTicker
    Worksheets(i).Range("Q3") = GreatestDecValue
    Worksheets(i).Range("P4") = GreatestVolTicker
    Worksheets(i).Range("Q4") = GreatestVolValue

    
    GreatestIncTicker = " "
    GreatestDecTicker = " "
    GreatestVolTicker = " "
    GreatestIncValue = 0
    GreatestDecValue = 0
    GreatestVolValue = 0
    
Next i
End Sub
