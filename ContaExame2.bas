Attribute VB_Name = "ContaExame2"
Sub contaexame2()
    Dim medico As String
    Dim tipoExame As String
    Dim contExame As Integer
    Dim i As Integer
    Dim ultimaLinha As Integer
    Dim ultimaLinhaExames As Integer
    Dim i2 As Integer
    Dim i3 As Integer
    Dim contInterno As Integer

    
    ultimaLinha = Range("a1048576").End(xlUp).Row
    ultimaLinhaExames = Range("m1048576").End(xlUp).Row
    
    Range("h2").Select
    
    For i = 2 To ultimaLinha
        If tipoExame = Cells(i, 8) And medico = Cells(i, 9) Then
        i = i + 1
        End If
        tipoExame = Cells(i, 8)
        medico = Cells(i, 9)
        
        For i2 = 1 To ultimaLinha
            If Cells(i2, 8) = tipoExame And Cells(i2, 9) = medico And Cells(i2, 7) <> "UMC IMAGEM" Then
                contExame = contExame + (Cells(i2, 10))
            End If
        Next
        ultimaLinhaExames = Range("m1048576").End(xlUp).Row
        For i3 = 2 To ultimaLinhaExames + 1
            If tipoExame = Cells(i3, 13) And medico = Cells(i3, 14) Then
            contInterno = contInterno + 1
            End If
            
        Next
        If contInterno = 0 Then
        Cells(ultimaLinhaExames + 1, 12) = contExame
        Cells(ultimaLinhaExames + 1, 13) = tipoExame
        Cells(ultimaLinhaExames + 1, 14) = medico
        End If
        contInterno = 0
        contExame = 0
    Next
End Sub
