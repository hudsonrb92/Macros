Attribute VB_Name = "ContaEstabele"
Sub contaEstabelecimento()
    Dim medico As String
    Dim tipoExame As String
    Dim cont As Integer
    Dim clinica As String
    Dim contClinica As Integer
    Dim valor As Double
    Dim ultLin As Integer
    Dim ultLin2 As Integer
    Dim i As Integer
    Dim i2 As Integer
    ultLin2 = Range("a1048576").End(xlUp).Row
    
    For i = 2 To ultLin2
        If Cells(i, 7) = clinica And Cells(i, 8) = tipoExame Then
        GoTo proximo
        End If
        clinica = Cells(i, 7)
        tipoExame = Cells(i, 8)
        ultLin = fimColuna(4, "A")
        Sheets(4).Select
        
        For i2 = 1 To ultLin
            Sheets(4).Select
            If clinica = Cells(i2, 1) And tipoExame = Cells(i2, 2) Then
            Sheets(1).Select
            GoTo proximo
            End If
         Next
        Sheets(1).Select
        
        For i2 = 2 To ultLin2
            If Cells(i2, 7) = clinica And Cells(i2, 8) = tipoExame Then
            cont = cont + Cells(i2, 10)
            End If
        Next
        
        
        Sheets(4).Select
        Cells(ultLin + 1, 1) = clinica
        Cells(ultLin + 1, 2) = tipoExame
        Cells(ultLin + 1, 3) = cont
        Sheets(1).Select
        cont = 0

proximo:
        
    Next
    
End Sub
