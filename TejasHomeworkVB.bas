Attribute VB_Name = "Module1"
Sub Mulitsheet()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Stock
    Next
    Application.ScreenUpdating = True
End Sub

Sub Stock()

  ' Set a variable for specifying the column of interest
  Dim Ticker_Symbol As String
  Dim Ticker_Volume As Double
  
  Dim Opening_Price As Double
  Dim Closing_Price As Double
  
  Dim Initial_Opening_Price As Double

  Dim Total_Stock_Volume As Double
  Dim Yearly_Change As Double
  Dim Percent_Change As Double
  
  Dim LRow As Long
  
  ' Find Last Row
  LRow = Cells(Rows.Count, 1).End(xlUp).Row
  MsgBox "Last Row: " & LRow
  
  Ticker_Volume = 0
  Total_Stock_Volume = 2
  Initial_Opening_Price = 2

  ' Loop through rows in the column till last row
    For i = 2 To LRow
    
    ' Searches for when the value of the next cell is different than that of the current cell
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker_Symbol = Cells(i, 1).Value
        
        Opening_Price = Cells(Initial_Opening_Price, 3).Value
        
        Initial_Opening_Price = i + 1
        
        Closing_Price = Cells(i, 6).Value
        
        Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
        
        Range("I" & Total_Stock_Volume).Value = Ticker_Symbol
        Range("J" & Total_Stock_Volume).Value = Ticker_Volume
        
        ' Finds diffrence between Closing Price and Opening Price for particular stock for that year
        
        Yearly_Change = Closing_Price - Opening_Price
        Range("K" & Total_Stock_Volume).Value = Yearly_Change
        
            If Yearly_Change >= 0 Then
            Range("K" & Total_Stock_Volume).Interior.Color = vbGreen
            Else
            Range("K" & Total_Stock_Volume).Interior.Color = vbRed
            End If
        
                If Opening_Price > 0 Then
                Range("L" & Total_Stock_Volume).Value = Round((Yearly_Change / Opening_Price) * 100, 3)
                Else
                Range("L" & Total_Stock_Volume) = 0
                End If
       
        Percent_Change = Range("L" & Total_Stock_Volume).Value
       
        Total_Stock_Volume = Total_Stock_Volume + 1
      
        Ticker_Volume = 0
        
      Else
      
     Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
  
    End If

  Next i
  
End Sub

