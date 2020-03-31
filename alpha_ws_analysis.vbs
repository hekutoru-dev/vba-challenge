Sub alphabetical_worksheets():

    Dim r, j As Integer
    Dim open_value, close_value, yearly_change, percent_change As Double
    Dim start, rowCount As Double
    Dim total As Double

    ' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Initializations
    total = 0   'Sumador de total stock value
    j = 0       'Contador para tabla summary
    start = 2
    yearly_change = 0
    
    'Ubicar el final de la tabla
    'rowCount = Cells(Row.Counts, "A").End(xlUp).Row
    rowCount = Range("A1").End(xlDown).Row
    
    'Iterate all rows with information.
    For r = 2 To rowCount
    
        'Ticker changes.
        If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
        
            'Give total value
            total = total + Cells(r, 7).Value
            Range("I" & 2 + j).Value = Cells(r, 1).Value
            Range("L" & 2 + j).Value = total
            
            'Get yearly change
            open_value = Cells(start, 3).Value
            close_value = Cells(r, 6).Value
            yearly_change = close_value - open_value
            Range("J" & 2 + j).Value = yearly_change
            
            'Ger percent change
            percent_change = Round((yearly_change / Cells(start, 3).Value * 100), 2)
            Range("K" & 2 + j).Value = percent_change & "%"
            
            'Restart values for new ticker
            total = 0
            j = j + 1
            start = r + 1

        Else
            total = total + Cells(r, 7).Value        
        End If       
    
    Next r

End Sub