Sub Testing()


  ' Set an initial variable for holding the ticker symbol
  Dim Tick As String


  ' Set an initial variable for holding the total ticker vomume
  Dim Volume_Total As Double
  Volume_Total = 0

  ' Keep track of the location for each Ticker Symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  'Set a Variable for Stock price at begining of year
  Dim Begin_Price As Double
  Dim Begin_Price1 As Double
  
  'Set a Variable for the stock price at the end of the year
  Dim End_Price As Double
  
  'Set Variable for Yearly Change
  Dim Yearly_Change As Double
  'Call last row
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
  Begin_Price1 = Range("c2").Value
  Range("k" & Summary_Table_Row).Value = Begin_Price1
  

  ' Loop through all Ticker Symbols
  For i = 2 To LastRow

    ' Check if we are still within the same Ticker Symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
      Tick = Cells(i, 1).Value

        'Set the first stock price
      Begin_Price = Cells(i + 1, 3)
      
      'Set the last stock price
      End_Price = Cells(i, 6)
        
      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value

      ' Print the Ticker Name in the Summary Table
      Range("i" & Summary_Table_Row).Value = Tick

      ' Print the Volume Total to the Summary Table
      Range("J" & Summary_Table_Row).Value = Volume_Total
      
      'Print the Begin Price
      Range("k" & Summary_Table_Row + 1).Value = Begin_Price
      
      'Print the End Price
      Range("l" & Summary_Table_Row).Value = End_Price
      

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Volume Total
      Volume_Total = 0

    ' If the cell immediately following a row is the same ticker..
    Else

      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value

    End If

  Next i



End Sub

Sub Calculations()

'Create Variables for Yearly Change and Yearly % Change
Dim Yearly_Change As Double
Dim Yearly_pChange As Variant

'Create variable for summary table
 Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

'create variable for last row for loop
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'begin for loop
For i = 2 To LastRow
    'condition to limit calculations to cells that are populated with data
    If Not IsEmpty(Cells(i, 12).Value) Then

        'calculate yearly change
        Yearly_Change = Cells(i, 12).Value - Cells(i, 11).Value
        'acount for errors when dividing by 0'
            If Cells(i, 11).Value <> 0 Then
                Yearly_pChange = Cells(i, 12).Value / Cells(i, 11).Value
            ElseIf Cells(i, 11).Value = 0 Then
                Yearly_pChange = "Undefined"
                
            Else
            End If
    'print the data
Range("m" & Summary_Table_Row).Value = Yearly_Change
    If Yearly_Change >= 0 Then
        Range("m" & Summary_Table_Row).Interior.ColorIndex = 4
    Else
        Range("m" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
Range("n" & Summary_Table_Row).Value = Yearly_pChange
    Range("n" & Summary_Table_Row).NumberFormat = "0.00%"

Summary_Table_Row = Summary_Table_Row + 1

End If

Next i


End Sub



