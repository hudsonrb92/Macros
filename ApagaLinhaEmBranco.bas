Attribute VB_Name = "ApagaLinhaEmBranco"
Sub excluirLinhaEmBranco()
    Dim i As Integer
    Dim ult As Integer
    
    ult = Range("a1048576").End(xlUp).Row
    
    For i = 1 To ult
    
        If Cells(i, 1) = "" Then
        Cells(i, 1).EntireRow.Delete
        End If
    
    Next
End Sub
