# HW-2--VBA

@@ -0,0 +1,157 @@
Attribute VB_Name = "Module1"
        'Unit 2 VBA Homework - The VBA of Wall Street
        
        Sub Stock_Market_Analysis()
    
    ' Declaration of Initial Variables And Setting of Default Variables
            Dim Ticker_name As String
            Dim First_open_price As Double
            Dim Last_close_price As Double
            Dim Yearly_change As Double
            Dim Percentage_change As Double
            Dim Last_price As Double
            Last_price = 2
            Dim Ticker_total_volume As Double
            Ticker_total_volume = 0
            Dim Summary_table_row As Long
            Summary_table_row = 2
            Dim LastRow As Long
            Dim LastRow1 As Long
            Dim Rng As Range
            Dim Rng2 As Range
            Dim ws As Worksheet
            
            For Each ws In Worksheets
            
            ' Column Headers
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = " Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            
            
            ' Determine the last row of the sheet
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            For i = 2 To LastRow
            
                     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                                 ' Set Ticker name
                                 Ticker_name = ws.Cells(i, 1).Value
                                 ' Set Ticker's total volume
                                 Ticker_total_volume = Ticker_total_volume + ws.Cells(i, 7).Value
                                 
                                  ' Print the Ticker name to the Summary Table
                                 ws.Range("I" & Summary_table_row).Value = Ticker_name
                                 
                                  'Print the Ticker's Total Volume to the Summary Table
                                 ws.Range("L" & Summary_table_row).Value = Ticker_total_volume
                                 
                                 ' ' Set Yearly Open, Yearly Close and determine the yearly price change
                                 First_open_price = ws.Range("C" & Last_price).Value
                                 Last_close_price = ws.Range("F" & i).Value
                                 Yearly_change = Last_close_price - First_open_price
                                 
                                 'Print the Ticker 's yearly price change to the Summary Table
                                 ws.Range("J" & Summary_table_row).Value = Yearly_change
                                 
                                 ' Determine the percentage change
                               If First_open_price <> 0 Then
                               
                                    Percentage_change = Yearly_change / First_open_price
                                    
                                   ' Print the Ticker 's yearly percentage change to the Summary Table
                                   ws.Range("k" & Summary_table_row).Value = Percentage_change
                                   
                               End If
                               
                               
                                ' Conditional Formatting Highlight Positive Green and Negative Red
                               If Yearly_change >= 0 Then
                                   ws.Range("j" & Summary_table_row).Interior.ColorIndex = 4
                                Else
                                 ws.Range("j" & Summary_table_row).Interior.ColorIndex = 3
                            End If
                                ws.Range("k" & Summary_table_row).NumberFormat = "0.00%"
                                                    
                                  ' Add One To The Summary Table Row , reset volume counter
                                  Summary_table_row = Summary_table_row + 1
                                  Last_price = i + 1
                                 Ticker_total_volume = 0
                
                     Else
                               ' Add to Total volume
                                Ticker_total_volume = Ticker_total_volume + ws.Cells(i, 7).Value
                                 
                End If
    
                
                        Next i
                           
                            'Set range from which to determine smallest and largest values
                            LastRow1 = ws.Cells(Rows.Count, 11).End(xlUp).Row
                            Set Rng = ws.Range("K2:K" & LastRow1)
                            Set Rng2 = ws.Range("l2:l" & LastRow1)
                            
                            
                            'Worksheet function MAX returns the largest value in a range
                            ws.Range("Q2").Value = Application.WorksheetFunction.Max(Rng)
                            
                             'Worksheet function MIN returns the smallest value in a range
                            ws.Range("Q3").Value = Application.WorksheetFunction.Min(Rng)
                            
                             'Worksheet function MAX returns the largest value in a range
                            ws.Range("Q4").Value = Application.WorksheetFunction.Max(Rng2)
                            ws.Range("Q2:Q3").NumberFormat = "0.00%"
                            
                            ' Set variable to exit loop when value found
                            foundMin = 1
                            foundMax = 1
                            foundMaxVol = 1
                            
                            For j = 2 To LastRow1
                                 
                                 If ws.Range("Q2").Value = ws.Range("k" & j).Value And foundMax = 1 Then
                                    
                                    'print value to table
                                    ws.Range("P2").Value = ws.Range("I" & j).Value
                                    
                                    
                                    foundMax = 0
                                End If
                                 
                                 If ws.Range("Q3").Value = ws.Range("k" & j).Value And foundMin = 1 Then
                                    ws.Range("P3").Value = ws.Range("I" & j).Value
                                    foundMin = 0
                                End If
                                
                                If ws.Range("Q4").Value = ws.Range("L" & j).Value And foundMaxVol = 1 Then
                                    ws.Range("P4").Value = ws.Range("I" & j).Value
                                    foundMaxVol = 0
                                End If
                                
                                'Exit loop once value found
                                
                                If foundMin = 0 And foundMax = 0 Then
                                    Exit For
                                End If
                                
                            Next j
                            
             ' Reset variables for next sheet
              Summary_table_row = 2
              Last_price = 2
             Ticker_total_volume = 0
             
             ' Format Table Columns To Auto Fit
            ws.Columns("I:Q").AutoFit
        
    Next ws
   
 End Sub
