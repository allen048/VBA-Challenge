Sub Multiple_Year_Stock()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Summary_Table_Row As Integer
    Dim input_row As Long
    Dim Ticker_Name As String
    Dim Total_Stock_Volume As LongLong
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim Ticker_Title As String
    Dim Max_Percentage_Change As Double
    Dim Min_Percentage_Change As Double
    Dim Max_Total_Volume As LongLong
    Dim Percentage_Change_InputRow As Long
    Dim Percentage_Change_LastRow As Long
    Dim Total_Stock_InputRow As Long
    Dim Total_Stock_LastRow As Long
    
    
   ' Loop through all sheets in the workbook
    For Each ws In Worksheets
    
       ' Keep track of the location for each ticker in the summary table
        Summary_Table_Row = 2
    
        ' Set an initial variable for holding the total stock volume per ticker
        Total_Stock_Volume = 0
        
        ' Find the last row for the Total Stock Volume
        Total_Stock_LastRow = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row
        
        ' Set an initial variable for holding the opening price
        Opening_Price = 0
        
        ' Set an initial variable for holding the yearly change
        Yearly_Change = 0
        
        ' Find the last row for the Percentage Change
        Percentage_Change_LastRow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
    
        ' Find the last row in the current sheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Insert Ticker title into cells on each worksheet
        ws.Cells(1, 9).Value = "Ticker"
        
        ' Insert Yearly Change title into cells on each worksheet
        ws.Cells(1, 10).Value = "Yearly Change"
        
        ' Insert Percentage Change title into cells on each worksheet
        ws.Cells(1, 11).Value = "Percentage Change"
        
        ' Insert Total Stock Volume title into cells on each worksheet
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Insert Greatest % Increase title into cells on each worksheet
        ws.Cells(2, 15).Value = "Greatest % Increase"
        
        ' Insert Greatest % Decrease title into cells on each worksheet
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        
        ' Insert Greatest Total Volume title into cells on each worksheet
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Insert Ticker title into cells on each worksheet
        ws.Cells(1, 16).Value = "Ticker"
        
        ' Insert Value title into cells on each worksheet
        ws.Cells(1, 17).Value = "Value"
        
        ' Finds Greatest % Increase
        Max_Percentage_Change = (WorksheetFunction.Max(ws.Range("K2:K3001"))) * 100
        
        ' Print Greatest % Increase
        ws.Cells(2, 17).Value = Max_Percentage_Change & "%"
        
        ' Finds Greatest % Decrease
        Min_Percentage_Change = (WorksheetFunction.Min(ws.Range("K2:K3001"))) * 100
        
        ' Print Greatest % Increase
        ws.Cells(3, 17).Value = Min_Percentage_Change & "%"
        
        ' Finds Greatest Total Volume
        Max_Total_Volume = WorksheetFunction.Max(ws.Range("L2:L3001"))
        
        ' Print Greatest Total Volume
        ws.Cells(4, 17).Value = Max_Total_Volume
                 
        ' Loop through all stock data
        For input_row = 2 To LastRow
        
              ' Check if we are still within the same ticker, if it is not...
               If ws.Cells(input_row + 1, 1).Value <> ws.Cells(input_row, 1).Value Then
              
              ' Set the Ticker Name
              Ticker_Name = ws.Cells(input_row, 1).Value

              ' Add to the Total Stock Volume Total
              ' Total_Stock_Volume = Total_Stock_Volume + ws.Cells(input_row, 7).Value
      
               ' Print the Ticker Name in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
               
               ' Subtract closing price from first opening price
                Yearly_Change = Round((ws.Cells(input_row, 6).Value - Opening_Price), 3)
                
                ' Print the change in stock price in the Summary Table
                 ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                 
                ' Conditional Formatting for Yearly Change
                    If Yearly_Change > 0 Then
                    
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                            
                    ElseIf Yearly_Change < 0 Then
                            
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                            
                    End If
                 
                 'Calculate Percentage Change
                 Percentage_Change = Round((Yearly_Change / Opening_Price) * 100, 2)
                
                ' Print the Percentage Change to the Summary Table
                 ws.Range("K" & Summary_Table_Row).Value = Percentage_Change & "%"

                ' Conditional Formatting for Percentage Change
                    If Percentage_Change > 0 Then
                    
                            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                            
                    ElseIf Percentage_Change < 0 Then
                            
                            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                            
                    End If
                 
               
                ' Print the Total Stock Volume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
               
                 ' Reset the Total Stock Volume
                 Total_Stock_Volume = 0
                              
              
              'Title:<Multiple_year_stock_data.xlsm>
              'Author:<Peter Sun>
              'Date:<2023>
              'Code Version:<1.0>
              'Availability:<Multiple_year_stock_data.xlsm>
              ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
               ElseIf ws.Cells(input_row, 1) <> ws.Cells(input_row - 1, 1) Then
            
               Opening_Price = ws.Cells(input_row, 3)
               ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
                Else
            
                ' Add to the Total Stock Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(input_row, 7).Value
                End If
                
        Next input_row
        
        
       ' Loops through all Percentage Change values in each worksheet
        For Percentage_Change_InputRow = 2 To Percentage_Change_LastRow
        
        ' Prints ticker values for Greatest % Increase in each worksheet
            If ws.Cells(Percentage_Change_InputRow, 11).Value = ws.Cells(2, 17).Value Then
        
            ws.Cells(2, 16).Value = ws.Cells(Percentage_Change_InputRow, 9).Value
        
            Else
        
            End If
            
            ' Prints ticker values for  Greatest % Decrease in each worksheet
            If ws.Cells(Percentage_Change_InputRow, 11).Value = ws.Cells(3, 17).Value Then
        
            ws.Cells(3, 16).Value = ws.Cells(Percentage_Change_InputRow, 9).Value
        
            Else
        
            End If
              
        Next Percentage_Change_InputRow
        
        ' Loops through all Total Stock Volume values in each worksheet
        For Total_Stock_InputRow = 2 To Total_Stock_LastRow
        
         ' Add to the Total Stock Volume
                ' Total_Stock_Volume = Total_Stock_Volume + ws.Cells(Total_Stock_InputRow, 7).Value
        
        ' Prints ticker values for Total Stock Volume in each worksheet
            If ws.Cells(Total_Stock_InputRow, 12).Value = ws.Cells(4, 17).Value Then
            
            ws.Cells(4, 16).Value = ws.Cells(Total_Stock_InputRow, 9).Value
            
            Else
        
            End If
        
        Next Total_Stock_InputRow
        
    Next ws
    
End Sub

