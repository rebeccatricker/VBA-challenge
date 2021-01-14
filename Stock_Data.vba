Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Data()

' Loop through all the sheets
For Each ws In Worksheets

    ' Determine the Last Row
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
' Create headers in each sheet
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

' Set an initial variable for holding the ticker name
Dim Ticker_Name As String

' Set an initial variable for opening price
Dim Opening_Price As Double
Opening_Price = 0

' Set an initial variable for closing price
Dim Closing_Price As Double
Closing_Price = 0

' Set an initial variable for holding the yearly change
Dim Yearly_Change As Double
Yearly_Change = 0

' Set an initial variable for the previous amount
Dim Previous_Amount As Long
Previous_Amount = 2

' Set an initial variable for holding the percent change
Dim Percent_Change As Double
Percent_Change = 0

' Set an initial variable for holding the total of stock volume
Dim Stock_Volume As Double
Stock_Volume = 0

' Track each ticker name in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
    
' Loop through all ticker names
For i = 2 To Last_Row
    
    ' Add Total Stock Volume
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    
    ' Check if we are still within the same ticker name, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ' Name Ticker cell range
        Ticker_Name = ws.Cells(i, 1).Value
        
        'Print the Ticker Name in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
          
        'Print the Total Stock Volume
        ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
            
        'Reset the Total Stock Volume
        Stock_Volume = 0
        
        ' Name Opening Price cell range
        Opening_Price = ws.Range("C" & Previous_Amount)
    
        ' Name Closing Price cell range
        Closing_Price = ws.Range("F" & i)
    
        ' Add to the Yearly change
        Yearly_Change = Closing_Price - Opening_Price
        
        'Print the Yearly Change
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
    
        ' Add Percent Change
            If Opening_Price = 0 Then
            
            Percent_Change = 0
        
            Else
            Opening_Price = ws.Range("C" & Previous_Amount)
            Percent_Change = (Yearly_Change / Opening_Price)
        
            End If
        
        'Print the Percent Change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        
        ' Conditional Formatting for Yearly Change
        
        Green = 4
        Red = 3
        
            If Yearly_Change >= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = Green
                
            Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = Red
                
            End If
            
        'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        ' Set previous amount
        Previous_Amount = i + 1
        
    End If
    
  Next i
 
 Next ws
    
End Sub

