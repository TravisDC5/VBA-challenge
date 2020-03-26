Attribute VB_Name = "Module1"

Sub StockCalculations()

    ' Declare Variables
    Dim j As Integer
    Dim i As Long
    Dim Ticker As String
    Dim lastValue As Long
    Dim tempValueOpen As Double
    Dim tempValueClose As Double
   
    
        ' Intialize Variables
        j = 2
        lastValue = ActiveSheet.UsedRange.Rows.Count
        Cells(2, 8).Value = Cells(2, 1).Value
        tempValueOpen = Cells(2, 3).Value
        Cells(1, 8).Value = "Ticker"
        Cells(1, 9).Value = "Yearly Change"
        Cells(1, 10).Value = "Percentage Change"
        Cells(1, 11).Value = "Total Stock Volume"
        

        ' Iterate and Compare
        For i = 2 To lastValue
    
        Ticker = Cells(i, 1).Value
  
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                tempValueClose = Cells(i, 6).Value
                totals = totals + Cells(i, 7).Value
                Cells(j, 11).Value = totals
                Cells(j, 8).Value = Ticker
                totals = 0
        
                    If tempValueOpen <> 0 And tempValueClose <> 0 Then
                        Cells(j, 9).Value = tempValueClose - tempValueOpen
                        Cells(j, 10).Value = ((tempValueClose - tempValueOpen) / tempValueOpen)
                        tempValueClose = 0
                        tempValueOpen = 0
                        j = j + 1
                
                    End If

            ElseIf Cells(i - 1, 2) > Cells(i, 2) Then
                tempValueOpen = Cells(i, 3).Value

            ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value Then
                totals = totals + Cells(i, 7).Value
        
            End If
        Next
    
        ' Formatting Cell Loop
        For i = 2 To j
    
            Cells(i, 10).NumberFormat = "0.00%"
        
            If Cells(i, 9).Value > 0 Then
                Cells(i, 9).Interior.ColorIndex = 4

            ElseIf Cells(i, 9).Value < 0 Then
                Cells(i, 9).Interior.ColorIndex = 3
        
            End If
        
        Next

End Sub

