Attribute VB_Name = "Module1"
Sub Ticker()
   
    'Set variables
        Dim ws As Worksheet
        Dim Ticker As String
        Dim Volume_Total As Double
        Dim Table_Row As Integer
        Dim LastRow As Long
        Dim Year_Open As Double
        Dim Year_Close As Double
        Dim Year_Change As Double
        Dim Percent_Change As Double
    
 
    
    'Set Headers
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
    
   'Initial Variable
    
    Table_Row = 2
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
    For i = 2 To LastRow
        'Run for Ticker
       If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

         Year_Open = ws.Cells(i, 3).Value
         
        End If
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                
        'Run for Yearly and Percent Change
            
           
                Year_Close = ws.Cells(i, 6).Value
                
                Year_Change = Year_Close - Year_Open
                
                If Year_Open = 0 Then
                    Percent_Change = 0
                Else
                    Percent_Change = (Year_Close - Year_Open) / Year_Open
                End If
            
        'Run for Volume Total
        
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                
        'Summary Table
        
                ws.Range("I" & Table_Row).Value = Ticker
                ws.Range("J" & Table_Row).Value = Year_Change
                    If Year_Change >= 0 Then
                        ws.Range("J" & Table_Row).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & Table_Row).Interior.ColorIndex = 3
                    End If
                ws.Range("K" & Table_Row).Value = Percent_Change
                ws.Range("L" & Table_Row).Value = Volume_Total
                Table_Row = Table_Row + 1

                
                Volume_Total = 0
            Else
            
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                
        End If
    
    Next i
    
    'Format Columns
    ws.Columns("J").NumberFormat = "0.00"
    ws.Columns("K").NumberFormat = "0.00%"


    Next ws
    
End Sub

