Attribute VB_Name = "SomaValores"
Sub somaValores()
    Dim valor As Double
    Dim nomePaciente As String
    Dim roda As Integer
    Dim ult As Integer
    
    ult = Range("a1048576").End(xlUp).Row
    
    Range("b1").Select
    
    
    For roda = 1 To ult
        nomePaciente = Cells(roda, 2)
        If Cells(roda, 2) = Cells(roda + 1, 2) Then
        Cells(roda, 5) = Cells(roda, 5) + Cells(roda + 1, 5)
        Cells(roda + 1, 1).EntireRow.Delete
        End If
        
    Next

End Sub
