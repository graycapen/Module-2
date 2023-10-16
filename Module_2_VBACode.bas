Attribute VB_Name = "Module1"
Sub Module_2():
    
    'Establishing Variables
    Dim Ticker As String
    Dim Total_Stock_Volume As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Extremes As Double
    Dim lastrow2 As Double
    
    Dim Summary_Table_Row As Integer
    
    Dim Start As Double
    
    Dim ws As Worksheet
    
    
    For Each ws In ThisWorkbook.Worksheets
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Total_Stock_Volume = 0
    Summary_Table_Row = 2
    Start = 2
     
    'Programming Loop
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        Total_Stock_Volume = 0
        
        Yearly_Change = (ws.Cells(i, 6).Value - ws.Cells(Start, 3).Value)
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
        
        Percent_Change = ((ws.Cells(i, 6).Value - ws.Cells(Start, 3).Value) / ws.Cells(Start, 3).Value)
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        
        Summary_Table_Row = Summary_Table_Row + 1
        Start = i + 1
        
        Else
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        End If
    Next i
    
   Next ws
   
    Extremes = Application.WorksheetFunction.Max(Range("K2:K" & lastrow))
    Cells(2, 16).Value = Extremes
    Ticker = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
    Cells(2, 15).Value = Cells(Ticker + 1, 9).Value
    
    Extremes = Application.WorksheetFunction.Min(Range("K2:K" & lastrow))
    Cells(3, 16).Value = Extremes
    Ticker = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
    Cells(3, 15).Value = Cells(Ticker + 1, 9).Value
    
    Extremes = Application.WorksheetFunction.Max(Range("L2:L" & lastrow))
    Cells(4, 16).Value = Extremes
    Ticker = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)
    Cells(4, 15).Value = Cells(Ticker + 1, 9).Value
    
    
    
End Sub
