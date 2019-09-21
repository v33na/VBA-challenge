Sub StockData()
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        
    ' Determine the Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

    ' Add Heading for summary
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
    'Create Variable to hold Value
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
    'Set Initial Open Price
        Open_Price = Cells(2, Column + 2).Value
         
     ' To set ticker symbol
        
        For i = 2 To LastRow
         
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
               
                Ticker_Name = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Ticker_Name
       'To Set Close Price
                Close_Price = Cells(i, Column + 5).Value
       'To Add Yearly Change
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, Column + 9).Value = Yearly_Change
      'To Add Percent Change
                If (Yearly_Change = 0) Then
                    Percent_Change = 0
                
                ElseIf (Open_Price = 0 And Yearly_Change <> 0) Then
                    Percent_Change = 1
                
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, Column + 10).Value = Percent_Change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
            
    ' Add Total Volume
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
    
    ' Add one to the summary table row
                Row = Row + 1
    ' reset the Open Price
                Open_Price = Cells(i + 1, Column + 2)
    ' reset the Volumn Total
                Volume = 0
    'if cells are the same ticker
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            
            End If
        
        Next i
    
    ' Determine the Last Row of Yearly Change per WS
        YearlyChangeLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
    
    ' Set the Cell Colors
        For j = 2 To YearlyChangeLastRow
            If (Cells(j, Column + 9).Value >= 0) Then
                Cells(j, Column + 9).Interior.Color = vbGreen
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.Color = vbRed
            End If
        Next j
        
    ' Set Greatest % Increase, % Decrease, and Total Volume
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 17).Value = "Ticker"
        Cells(1, 18).Value = "Value"
    
    'To find the greatest value and its associate ticker
        For Z = 2 To YearlyChangeLastRow
            If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YearlyChangeLastRow)) Then
                Cells(2, Column + 16).Value = Cells(Z, Column + 8).Value
                Cells(2, Column + 17).Value = Cells(Z, Column + 10).Value
                Cells(2, Column + 17).NumberFormat = "0.00%"
      
      'To find the Minimum value and its associate ticker
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YearlyChangeLastRow)) Then
                Cells(3, Column + 16).Value = Cells(Z, Column + 8).Value
                Cells(3, Column + 17).Value = Cells(Z, Column + 10).Value
                Cells(3, Column + 17).NumberFormat = "0.00%"
        
        'To find the greatest total volume and its associate ticker
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YearlyChangeLastRow)) Then
                Cells(4, Column + 16).Value = Cells(Z, Column + 8).Value
                Cells(4, Column + 17).Value = Cells(Z, Column + 11).Value
            End If
            
        Next Z
 Next WS
        
End Sub
