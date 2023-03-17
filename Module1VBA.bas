Attribute VB_Name = "Module1"
Sub stonks()

'set initial variables
Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim vol As Long
Dim Summary_Table_Row As Integer

'there is an overflow error that I cannot figure out
On Error Resume Next

For Each ws In Worksheets



LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"


    'setup integers for loop
    Summary_Table_Row = 2

    'loop
    For i = 2 To LastRow
             
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ' Set the ticker
            ticker = ws.Cells(i, 1).Value
        
            ' Add vol
             vol = vol + ws.Cells(i, 7).Value
             
            'set the open price
            open_price = open_price + ws.Cells(i, 3).Value
            'set the close price
            close_price = close_price + ws.Cells(i, 6).Value
            
            'set the yearly difference
            yearly_change = open_price - close_price
            
            'set the percent change
            percent_change = (close_price / open_price) * 100
            
        
            ' Print the ticker
            ws.Range("I" & Summary_Table_Row).Value = ticker
        
            ' Print the volume
            ws.Range("L" & Summary_Table_Row).Value = vol
            
            ' Print the yearly change
            ws.Range("J" & Summary_Table_Row).Value = yearly_change
            
            'print the percent change
            ws.Range("K" & Summary_Table_Row).Value = percent_change
        
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
              
            ' Reset the variables
            vol = 0
            open_price = 0
            close_price = 0
              
        Else

            ' Add to the variables
            vol = vol + ws.Cells(i, 7).Value
            open_price = open_price + ws.Cells(i, 3).Value
            close_price = close_price + ws.Cells(i, 6).Value
        
        End If


    Next i
    
    
    
'formatting
ws.Columns("K").NumberFormat = "0.00%"

Dim rnge As Range
    Dim k As Long
    Dim color_cell As Range
    
    Set rnge = ws.Range("J2", Range("J2").End(xlDown))
    
    
    For k = 1 To LastRow
    Set color_cell = rnge(k)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
            End With
       End Select
    Next k
    
    

'this section of code I used the sytax from this website(https://www.excelanytime.com) about how to pull the min and max from columns
Dim rng As Range
Dim rng2 As Range
Dim dblMin As Double
Dim dblMax As Double
Dim dblMax2 As Double
Dim Summary_Table2_Row As Integer

 
'Set range from which to determine smallest value
Set rng = ws.Range("K")

'Worksheet function MIN
dblMin = Application.WorksheetFunction.Min(rng)
'Worksheet function MAX
dblMax = Application.WorksheetFunction.Max(rng)
    
Set rng2 = ws.Range("K")
dblMax2 = Application.WorksheetFunction.Max(rng)
    
ws.Cells(5, 14).Value = "Greatest % Decrease"
ws.Cells(6, 14).Value = "Greatest % Increase"
ws.Cells(7, 14).Value = "Greatest Volume"
    

ws.Cells(5, 15).Value = dblMin
ws.Cells(6, 15).Value = dblMx
ws.Cells(7, 15).Value = dblMax2

Next ws

End Sub

