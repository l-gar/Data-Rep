Sub DelCols()
 
Worksheets("2016").Columns("I:L").Clear
 
Worksheets(Array("2016", "2015", "2014")).FillAcrossSheets Worksheets("2016").Columns("I:L")
 
End Sub

Sub Stock_Testing()
    
    Dim ticker As String
    Dim Vol As Double
    Dim ws As Worksheet
    Dim WS_Count As Integer
    Dim NumRows As Long
    
    Vol = 0
    WS_Count = ActiveWorkbook.Worksheets.Count
        MsgBox (WS_Count)
    Dim Summary_Table_Row As Integer
    
    
    
    
  
    'LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
  ' Set an initial variable for holding the ticker
  
       For x = 1 To WS_Count
            NumRows = ActiveWorkbook.Worksheets(x).Range("A1", ActiveWorkbook.Worksheets(x).Range("A1").End(xlDown)).Rows.Count
            MsgBox ActiveWorkbook.Worksheets(x).Name
            Summary_Table_Row = 2
            ActiveWorkbook.Worksheets(x).Cells(1, 9).Value = "Ticker"
            ActiveWorkbook.Worksheets(x).Cells(1, 10).Value = "Volume"
  ' Loop through all tickers
            For I = 2 To NumRows
                

    ' Check if we ticker is the same ticker, if it is not...
                If ActiveWorkbook.Worksheets(x).Cells(I + 1, 1).Value <> ActiveWorkbook.Worksheets(x).Cells(I, 1).Value Then
                    ticker = ActiveWorkbook.Worksheets(x).Cells(I, 1).Value
                    Vol = Vol + ActiveWorkbook.Worksheets(x).Cells(I, 7).Value
        
                    ActiveWorkbook.Worksheets(x).Range("I" & Summary_Table_Row).Value = ticker
                    ActiveWorkbook.Worksheets(x).Range("J" & Summary_Table_Row).Value = Vol
        
                    Summary_Table_Row = Summary_Table_Row + 1
        
                    Vol = 0
        
                    Else
        
                    Vol = Vol + ActiveWorkbook.Worksheets(x).Cells(I, 7).Value
      
        'MsgBox (Vol)
                End If
            Next I
            
        Next x
End Sub



