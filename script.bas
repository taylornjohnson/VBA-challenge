Attribute VB_Name = "Module1"
Sub Stock_Volume_Count()
Dim Total_Stock_Volume As Double
Dim i As Long
Dim j As Integer
Dim Change As Double
Dim Start As Long
Dim RowCount As Long
Dim Days As Integer
Dim DailyChange As Double
Dim Percent_Change As Double
j = 0
Total_Stock_Volume = 0
Change = 0
Start = 2
   Range("I1").Value = "Ticker"
   Range("J1").Value = "Yearly Change"
   Range("K1").Value = "Percent Change"
   Range("L1").Value = "Total Stock Volume"
   LastRow = Cells(Rows.Count, "A").End(xlUp).Row
   For i = 2 To LastRow
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
           If Total_Stock_Volume = 0 Then
               Range("I" & 2 + j).Value = Cells(i, 1).Value
               Range("J" & 2 + j).Value = 0
               Range("K" & 2 + j).Value = "%" & 0
               Range("L" & 2 + j).Value = 0
           Else
               If Cells(Start, 3) = 0 Then
                   For find_value = Start To i
                       If Cells(find_value, 3).Value <> 0 Then
                           Start = find_value
                           Exit For
                       End If
                    Next find_value
               End If
               Change = (Cells(i, 6) - Cells(Start, 3))
               Percent_Change = Round((Change / Cells(Start, 3) * 100), 2)
               Start = i + 1
               Range("I" & 2 + j).Value = Cells(i, 1).Value
               Range("J" & 2 + j).Value = Round(Change, 2)
               Range("K" & 2 + j).Value = "%" & Percent_Change
               Range("L" & 2 + j).Value = Total_Stock_Volume
               Select Case Change
                   Case Is > 0
                       Range("J" & 2 + j).Interior.ColorIndex = 4
                   Case Is < 0
                       Range("J" & 2 + j).Interior.ColorIndex = 3
                   Case Else
                       Range("J" & 2 + j).Interior.ColorIndex = 0
               End Select
           End If
           Total_Stock_Volume = 0
           Change = 0
           j = j + 1
           Days = 0
       Else
           Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
       End If
   Next i
End Sub
