Sub Credit_Cards()

'Brand Name
'Brand Total
'Summary Table Row

Dim bn As String
Dim bt As Double
Dim str As Integer

str = 2

'To do: add sorting code

    For i = 2 To 101
    
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
        bn = Cells(i, 1).Value
        bt = bt + Cells(i, 3).Value
        Cells(str, 7).Value = bn
        Cells(str, 8).Value = bt
        
        str = str + 1
        
        bt = 1
        
        Else
        
        bt = bt + Cells(i, 3).Value
    
        End If
        
    Next i

End Sub
