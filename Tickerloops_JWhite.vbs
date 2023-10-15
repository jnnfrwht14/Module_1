Sub Copywkst()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call tickerloops
    Next xSh
    Application.ScreenUpdating = True
End Sub
Sub tickerloops()

    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Summary_Table_Row As Integer
    Dim Change_Frac As Double
    Dim Total_Stock_Volume As LongLong
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim input_row_num As Long
    Dim Last_Data_Row As Long
    Dim Greatest_Volume As Double
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    
    'summary table names
    Range("i1").Value = "Ticker"
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total Stock Volume"
    Range("o2").Value = "Greatest % Increase"
    Range("o3").Value = "Greatest % Decrease"
    Range("o4").Value = "Greatest Total Volume"
    Range("p1").Value = "Ticker"
    Range("q1").Value = "Value"
    
        Summary_Table_Row = 2
        Total_Stock_Volume = 0
        Open_Price = Cells(2, 3).Value
        
        Last_Data_Row = Cells(Rows.Count, 1).End(xlUp).Row
        
        For input_row_num = 2 To Last_Data_Row
            Ticker = Cells(input_row_num, 1).Value
            Total_Stock_Volume = Total_Stock_Volume + Cells(input_row_num, 7).Value
        If Cells(input_row_num + 1, 1).Value <> Ticker Then
            Close_Price = Cells(input_row_num, 6).Value
            Yearly_Change = Close_Price - Open_Price
            Change_Frac = Yearly_Change / Open_Price
            
            'add some color!
            
            Range("i" & Summary_Table_Row).Value = Ticker
            Range("j" & Summary_Table_Row).Value = Yearly_Change
            Range("k" & Summary_Table_Row).Value = FormatPercent(Change_Frac)
            Range("l" & Summary_Table_Row).Value = Total_Stock_Volume
            If Range("j" & Summary_Table_Row).Value < 0 Then
                Range("j" & Summary_Table_Row).Interior.ColorIndex = 3
            ElseIf Range("j" & Summary_Table_Row).Value > 0 Then
                Range("j" & Summary_Table_Row).Interior.ColorIndex = 14
            End If
            
            Summary_Table_Row = Summary_Table_Row + 1
            Open_Price = Cells((input_row_num + 1), 3).Value
            Total_Stock_Volume = 0
        End If
        
    Next input_row_num
    
        'analyze the data
            'max change
        For input_row_num = 2 To Last_Data_Row
            If Cells(input_row_num, 11).Value = Application.WorksheetFunction.Max(Range("k2:k" & Last_Data_Row)) Then
            Cells(2, 16).Value = Cells(input_row_num, 9).Value
            Cells(2, 17).Value = Cells(input_row_num, 11).Value
            Cells(2, 17).NumberFormat = "0.00%"
            
            'minimum change
        ElseIf Cells(input_row_num, 11).Value = Application.WorksheetFunction.Min(Range("k2:k" & Last_Data_Row)) Then
            Cells(3, 16).Value = Cells(input_row_num, 9).Value
            Cells(3, 17).Value = Cells(input_row_num, 11).Value
            Cells(3, 17).NumberFormat = "0.00%"
            
            'greatest total volume
        ElseIf Cells(input_row_num, 12).Value = Application.WorksheetFunction.Max(Range("l2:l" & Last_Data_Row)) Then
            Cells(4, 16).Value = Cells(input_row_num, 9).Value
            Cells(4, 17).Value = Cells(input_row_num, 12).Value
            
        End If
        
    Next input_row_num
           
    
End Sub