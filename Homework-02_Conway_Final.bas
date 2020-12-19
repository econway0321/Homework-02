Attribute VB_Name = "Module1"
Sub Test_Data():

'Set initial variables
    
    Dim ws As Worksheet
    
    
    Dim Ticker_Name As String
    Dim i As Long
    Dim Vol_Total As Double
    Dim last_row As Double
    Dim Open_Amt As Double
    Dim Close_Amt As Double
    Dim Yr_Change As Double
    Dim Pct_Change As Double
    Dim Summary_Table_Row As Integer
  
  For Each ws In Worksheets

    'Create Headers
    ws.Cells(1, 9).Value = "Ticker" 'I
    ws.Cells(1, 10).Value = "Yearly Change" 'J
    ws.Cells(1, 11).Value = "Pct Change" 'K
    ws.Cells(1, 12).Value = "Volume Total" 'L

    'Find Last Row
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Keep track of the location for the data in the Summary_Table

    'Initialize open_amount to be the value at C2
    Open_Amt = ws.Range("C2").Value 'Storing the value for future use, purpose of a variable
    Vol_Total = 0 'Starting point, i.e. counter
    Summary_Table_Row = 2
    Ticker_Name = 0 'Starting point, i.e. counter
    
    
    For i = 2 To last_row
            'Check if ticker symbols are the same, if it is not
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then ' <> i.e. Address not equal to the first address
                'Set the Ticker_Name
                Ticker_Name = ws.Cells(i, 1).Value
                'Print the unique ticker symbol in the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

                'Set Vol Total
                Vol_Total = Vol_Total + ws.Cells(i, 7).Value

                'Print the Vol_Total in the summary table
                ws.Range("L" & Summary_Table_Row).Value = Vol_Total

                'Establish Close Amount
                Close_Amt = ws.Cells(i, 6).Value

                'Set the Yr_Change
                Yr_Change = Open_Amt - Close_Amt
                'Print the Yr Change in the summary table
                ws.Range("J" & Summary_Table_Row).Value = Round(Yr_Change, 2)

                If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4 'green
                End If
                If ws.Range("J" & Summary_Table_Row).Value < 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3 'red
                End If
                If ws.Range("J" & Summary_Table_Row).Value = 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 0 'white
                End If



'               Set the Pct Change
                If Open_Amt = 0 Then
                    Pct_Change = 0
                Else
                    Pct_Change = Round(((Open_Amt - Close_Amt) / Open_Amt * 100), 2)
                End If
              
              'Print the Pct_Change in summary table
                    ws.Range("K" & Summary_Table_Row).Value = "%" & Pct_Change
               
                'Set the Open Amount
                Open_Amt = ws.Cells(i + 1, 3).Value 'Coordinates

                 'Reset Ticker Total
                Summary_Table_Row = Summary_Table_Row + 1
                Vol_Total = 0

        Else
                'Add to the Vol Total
                Vol_Total = Vol_Total + ws.Cells(i, 7).Value
                
        
        End If
        
  
   
               
     Next i
     Next ws

End Sub







