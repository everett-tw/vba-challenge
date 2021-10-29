VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Function_Name():
    'MsgBox ("Hello World")
    'alright lets start with the ticker symbols
    'the plan is to create a for loop that only
    'prints when the current symbol doesn't match
    'the previous symbol
    Dim total As Double
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    'found this counting upwards idea on stack exchange
    'MsgBox (RowCount)
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"
    For i = 2 To RowCount:
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Range("I" & 2 + j).Value = Cells(i, 1).Value
            Range("J" & 2 + j).Value = total
            total = 0
        
            j = j + 1
            Else
                total = total + Cells(i, 7).Value
        End If
    Next i
    'I was lost for a very long time on how to iterate on the rows
    'apparently i don't have to declare j
    'which saves me a lot of for loop nonsense
    'but i still think this makes way more sense in python
End Sub

