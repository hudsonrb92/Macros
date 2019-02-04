Attribute VB_Name = "RetiraVazio"
Sub retiraVazio()
    Dim i As Integer
    Dim ultimaLinha As Integer
    
    ultimaLinha = Range("a1048576").End(xlUp).Row
    
    
    For i = 1 To ultimaLinha
        If Cells(i, 6).Value = "" And Cells(i, 8) = "CR" Then
        Cells(i, 6).Value = "RX"
        End If
        
    Next
    
End Sub
