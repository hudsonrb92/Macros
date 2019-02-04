Attribute VB_Name = "ContaExameFinal"
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
    Dim colunaMedico As Integer
    Dim colunaTipoExame As Integer
    Dim colunaEstabelece As Integer
    Dim colunaResultado As Integer
    Dim colunaContagemExame As Integer
    Dim colunaResultadoExame As Integer
    Dim colunaResultadoMedico As Integer
    Dim folhaMain As Integer
    Dim folhaResultados As Integer
    Dim folhaAbdENaoAbd As Integer
    Dim exameRealizados As Integer
    
    
    exameRealizados = 10
    folhaMain = 1
    folhaResultados = 3
    folhaAbdENaoAbd = 2
    colunaResultado = 3
    colunaResultadoExame = 2
    colunaResultadoMedico = 1
    colunaMedico = 9
    colunaTipoExame = 8
    colunaEstabelece = 7
    
    ultLin2 = Range("a1048576").End(xlUp).Row
    
    
    For i = 2 To ultLin2
        If Cells(i, colunaMedico) = medico And Cells(i, colunaTipoExame) = tipoExame And Cells(i, colunaEstabelece) <> "UMC IMAGEM" Then
        GoTo proximo
        End If
        medico = Cells(i, colunaMedico)
        tipoExame = Cells(i, colunaTipoExame)
        ultLin = fimColuna(folhaResultados, "A")
        Sheets(folhaResultados).Select
        
        For i2 = 1 To ultLin
            Sheets(folhaResultados).Select
            If medico = Cells(i2, colunaResultadoMedico) And tipoExame = Cells(i2, colunaResultadoExame) Then
            Sheets(folhaMain).Select
            GoTo proximo
            End If
         Next
        Sheets(folhaMain).Select
        
        For i2 = 2 To ultLin2
            If Cells(i2, colunaMedico) = medico And Cells(i2, colunaTipoExame) = tipoExame And Cells(i2, colunaEstabelece) <> "UMC IMAGEM" Then
            cont = cont + Cells(i2, exameRealizados)
            End If
        Next
        
        
        Sheets(folhaResultados).Select
        Cells(ultLin + 1, colunaResultadoMedico) = medico
        Cells(ultLin + 1, colunaResultadoExame) = tipoExame
        Cells(ultLin + 1, colunaResultado) = cont
        Sheets(folhaMain).Select
        cont = 0

proximo:
        
    Next
    
End Sub

