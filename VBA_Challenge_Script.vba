Sub VBA_Challenge():


'Loop through each worksheet in the workbook:
    Dim ws As Worksheet
    
        For Each ws In Worksheets
    
            'Generate and define type for all vars:
                Dim Ticker As String
                Dim Summary_Table_Row_1 As Integer
                Dim LastRow As Long
                Dim Tot_Stock_Vol As Double
                Dim Open_Price As Double
                Dim Close_Price As Double
                Dim Yearly_Change As Double
                Dim Percent_Change As Double
                Dim Greatest_Increase As Double
                Dim Greatest_Decrease As Double
                Dim Greatest_Vol As Double
                Dim Prev_Stock_Price As Long
        
            'Assign var names to column headings:
                ws.Cells(1, 9).Value = "Ticker"
                ws.Cells(1, 16).Value = "Ticker"
                ws.Cells(1, 10).Value = "Yearly Change"
                ws.Cells(1, 11).Value = "Percent Change"
                ws.Cells(1, 12).Value = "Total Stock Volume"
                ws.Cells(1, 15).Value = "Bonus"
                ws.Cells(1, 17).Value = "Value"
                ws.Cells(2, 15).Value = "Greatest % Increase"
                ws.Cells(3, 15).Value = "Greatest % Decrease"
                ws.Cells(4, 15).Value = "Greatest Total Volume"

            'For each Stock Ticker, find Yearly Change, Percent Change, and Total Stock Volumne Traded:
                'Instruct summary table to store the data extracted from colA to colG, starting in Row2:
                    Summary_Table_Row_1 = 2
            
                'Assign starting values for all vars:
                    Tot_Stock_Vol = 0
                    Prev_Stock_Price = 2
    
                'Find the last row of the dataset:
                    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
            
                'Loop through the stocks:
                    For i = 2 To LastRow
                    
                        'Add to the total stock volume:
                            Tot_Stock_Vol = Tot_Stock_Vol + ws.Cells(i, 7).Value
                                    
                        
                        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                            'For each new stock, provide the ticker:
                                Ticker = ws.Cells(i, 1).Value

                            'Print the ticker name in the summary table:
                                ws.Range("I" & Summary_Table_Row_1).Value = Ticker

                            'Print the total stock volume in the summary table:
                                 ws.Range("L" & Summary_Table_Row_1).Value = Tot_Stock_Vol

                            'Reset the total stock volume to zero for the next iteration:
                                Tot_Stock_Vol = 0
                        
                            'Find the opening price, closing price, and calculate yearly change:
                                Open_Price = ws.Range("C" & Prev_Stock_Price)
                                Close_Price = ws.Range("F" & i)
                                Yearly_Change = (Close_Price - Open_Price)

                            'Print thes yearly change value in the summary table:
                                ws.Range("J" & Summary_Table_Row_1).Value = Yearly_Change
                
                                'Apply conditional formatting such that cells with negative change are highlighted red, and those with positive change are highlighted green:
                                    If Yearly_Change < 0 Then
                                        ws.Range("J" & Summary_Table_Row_1).Interior.ColorIndex = 3 'Red
                    
                                    Else
                                        ws.Range("J" & Summary_Table_Row_1).Interior.ColorIndex = 4 'Green
                    
                                    End If
                      
                            'Calculate and print the percentage change in the summary table:
                                'For the avoidance of any issues caused by dividing by 0:
                                    If Open_Price = 0 Then
                                        Percent_Change = 0
                            
                                    Else
                                        Percent_Change = (Yearly_Change / Open_Price)
                            
                                    End If
                                        ws.Range("K" & Summary_Table_Row_1).NumberFormat = "0.00%" 'Format column as % with 2 decimal places
                                        ws.Range("K" & Summary_Table_Row_1).Value = Percent_Change
                        
                            'Advance the summary table by one row:
                                Summary_Table_Row_1 = Summary_Table_Row_1 + 1

                            'Advance the previous stock price var by one two
                                Prev_Stock_Price = i + 1
            
                        End If
    
                    Next i


'---------------------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------

    'Bonus:

        'Set starting values for bonus vars:
            Greatest_Increase = 0
            Greatest_Decrease = 0
            Greatest_Vol = 0

        ' Greatest % Increase:
            'Loop through the stocks to identify that with the largest % increase:
                For i = 2 To LastRow
                    If ws.Cells(i, 11).Value > Greatest_Increase Then
                        Greatest_Increase = ws.Cells(i, 11).Value
                        Ticker = ws.Cells(i, 9).Value
            
                    End If
        
                Next i
        
        'Print the greatest % increase and it's associated ticker in the summary table:
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q2").Value = Greatest_Increase
            ws.Range("P2").Value = Ticker
                   
        'Greatest % Decrease:
            'Loop through the stocks to identify that with the largest % decrease:
                For i = 2 To LastRow
                    If ws.Cells(i, 11).Value < Greatest_Decrease Then
                        Greatest_Decrease = ws.Cells(i, 11).Value
                        Ticker = ws.Cells(i, 9).Value
            
                    End If
        
                Next i
    
        'Print the greatest % decrease and it's associated ticker in the summary table:
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("Q3").Value = Greatest_Decrease
            ws.Range("P3").Value = Ticker


        'Greatest Total Stock Volume:
            'Loop through the stocks to identify that with the largest total volume traded:
                For i = 2 To LastRow
                    If ws.Cells(i, 12).Value > Greatest_Vol Then
                        Greatest_Vol = ws.Cells(i, 12).Value
                        Ticker = ws.Cells(i, 9).Value
            
                    End If
        
                Next i
            
        'Print the greatest total volume and it's associated ticker in the summary table:
            ws.Range("Q4").Value = Greatest_Vol
            ws.Range("P4").Value = Ticker
    
    Next ws
    
End Sub
