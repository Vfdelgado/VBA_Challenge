Attribute VB_Name = "Module1"
Sub Stock_Data()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock As Double
    Dim Summary_Table_Row As Integer
    Dim Greatest_Increase As Double
    Dim Increase_Ticker As String
    Dim Greatest_Decrease As Double
    Dim Decrease_Ticker As String
    Dim Greatest_Total As Double
    Dim Total_Ticker As String
    Dim Final_Last_Row As Long

    For Each ws In Worksheets
        With ws
       
            Total_Stock = 0
            Summary_Table_Row = 2

         
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row

       
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"

           
            For i = 2 To lastRow
                If .Cells(i - 1, 1).Value <> .Cells(i, 1).Value Then
                    Opening_Price = .Cells(i, 3)
                ElseIf .Cells(i + 1, 1).Value <> .Cells(i, 1).Value Then
                    Closing_Price = .Cells(i, 6).Value
                    Ticker = .Cells(i, 1).Value
                    Yearly_Change = Closing_Price - Opening_Price
                    Percent_Change = Yearly_Change / Opening_Price
                    Total_Stock = Total_Stock + .Cells(i, 7).Value

                  
                    .Cells(Summary_Table_Row, 9).Value = Ticker
                    .Cells(Summary_Table_Row, 10).Value = Yearly_Change
                    .Cells(Summary_Table_Row, 11).Value = Percent_Change
                    .Cells(Summary_Table_Row, 12).Value = Total_Stock
                    Summary_Table_Row = Summary_Table_Row + 1

                 
                    Total_Stock = 0
                Else
                    Total_Stock = Total_Stock + .Cells(i, 7).Value
                End If
            Next i

          
            .Range("O2").Value = "Greatest % Increase"
            .Range("O3").Value = "Greatest % Decrease"
            .Range("O4").Value = "Greatest Total Volume"
            .Range("P1").Value = "Ticker"
            .Range("Q1").Value = "Value"

      
            Final_Last_Row = .Cells(.Rows.Count, 9).End(xlUp).Row

           
            For j = 2 To Final_Last_Row
                If Greatest_Increase < .Cells(j, 11).Value Then
                    Greatest_Increase = .Cells(j, 11).Value
                    Increase_Ticker = .Cells(j, 9).Value
                End If

                If Greatest_Decrease > .Cells(j, 11).Value Then
                    Greatest_Decrease = .Cells(j, 11).Value
                    Decrease_Ticker = .Cells(j, 9).Value
                End If

                If Greatest_Total < .Cells(j, 12).Value Then
                    Greatest_Total = .Cells(j, 12).Value
                    Total_Ticker = .Cells(j, 9).Value
                End If
            Next j

         
            .Range("P2").Value = Increase_Ticker
            .Range("Q2").Value = Format(Greatest_Increase, "0.00%") ' Format as percentage with 2 decimal places
            .Range("P3").Value = Decrease_Ticker
            .Range("Q3").Value = Format(Greatest_Decrease, "0.00%") ' Format as percentage with 2 decimal places
            .Range("P4").Value = Total_Ticker
            .Range("Q4").Value = Greatest_Total

            
            For k = 2 To Final_Last_Row
                If .Cells(k, 10) > 0 Then
                    .Cells(k, 10).Interior.ColorIndex = 4
                Else
                    .Cells(k, 10).Interior.ColorIndex = 3
                End If
            Next k

       
            .Range("I1:Q1").Font.Bold = True
            .Range("O2:O4").Font.Bold = True
            .Columns("I:Q").AutoFit
        End With
    Next ws
End Sub


