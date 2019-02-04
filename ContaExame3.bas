Attribute VB_Name = "ContaExame3"
Sub contaExame3()
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
        If Cells(i, 9) = medico And Cells(i, 8) = tipoExame And Cells(i, 7) <> "UMC IMAGEM" Then
        GoTo proximo
        End If
        medico = Cells(i, 9)
        tipoExame = Cells(i, 8)
        ultLin = fimColuna(3, "A")
        Sheets(3).Select
        
        For i2 = 1 To ultLin
            Sheets(3).Select
            If medico = Cells(i2, 1) And tipoExame = Cells(i2, 2) Then
            Sheets(1).Select
            GoTo proximo
            End If
         Next
        Sheets(1).Select
        
        For i2 = 2 To ultLin2
            If Cells(i2, 9) = medico And Cells(i2, 8) = tipoExame And Cells(i2, 7) <> "UMC IMAGEM" Then
            cont = cont + Cells(i2, 10)
            End If
        Next
        
        
        Sheets(3).Select
        Cells(ultLin + 1, 1) = medico
        Cells(ultLin + 1, 2) = tipoExame
        Cells(ultLin + 1, 3) = cont
        Sheets(1).Select
        cont = 0

proximo:
        
    Next
    
End Sub

