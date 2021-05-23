Attribute VB_Name = "Module1"
Sub StockData()

    For Each ws In Worksheets


 
  ' Set an initial variable for holding the brand name
  Dim Stock As String

  ' Set an initial variable for holding the total per credit card brand
  Dim Volume As Double
  Volume = 0
  
  Dim StartPrice, EndPrice, Change, Percent As Double
  StartPrice = ws.Cells(2, 3).Value

  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  Dim I As Long

  I = 1

  Dim MoreToDo As Boolean
  MoreToDo = True
  
  ' Loop through all credit card purchases
  While MoreToDo = True

    I = I + 1
    
    ' Check if we are still within the same credit card brand, if it is not...
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

      ' Set the Brand name
      Stock = ws.Cells(I, 1).Value
    

      ' Add to the Brand Total
      Volume = Volume + CDbl(ws.Cells(I, 7).Value)

     EndPrice = ws.Cells(I, 6).Value
     Change = EndPrice - StartPrice
     Percent = Change / StartPrice
     
      ' Print the Credit Card Brand in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Stock

      ' Print the Brand Amount to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Change
      If Change < 0 Then
           ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
      Else
           ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
      End If

      ' Print the Brand Amount to the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Percent
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
     
      ' Print the Brand Amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      Volume = 0
      
    ' Debug.Print (Str((I)))
      If IsEmpty(ws.Cells(I + 1, 1).Value) Then
          MoreToDo = False
          ' MsgBox ("test")
      End If

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
     Volume = Volume + CDbl(ws.Cells(I, 7).Value)
     
    

    End If

  Wend
  
  Next ws
   

End Sub
