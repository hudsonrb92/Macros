Attribute VB_Name = "ContaExame"
Sub contaExame()
    Dim contaExame As Integer
    Dim tipoDeExame As String
    Dim linha As Integer
    Dim linha1 As Integer
    Dim linha2 As Integer
    Dim tipo As String
    Dim soma As Integer
    Dim estabelecimento As String
    Dim medico As String
    Dim u As Integer
    u = 2
    
    linha2 = 1
    linha = 0
    linha1 = 2
    ultimaLinha = Range("h1048576").End(xlUp).Row
    Range("h2").Select
    
    Do Until linha1 > 15
    'Linha 6 devido ao numero de tipode de exames que são 5 dx,cr,ct,mr,mg
    
    'Definir o tipo de exame
    'Observação o tipo de exame nao pode ser repitir aos anteriores
        tipoDeExame = ActiveCell
        
            For linha = 2 To ultimaLinha
            If Cells(linha, 9) = medico Then
            If Cells(linha, 8) = tipoDeExame Then
            'Coluna que recebe o valor de exames é diferente da contagem
            If Cells(linha, 7) <> "UMC IMAGEM" Then
            contaExame = contaExame + Cells(linha, 10)
            'Contagem de exames feita
            End If
            End If
            End If
            Next
        'Tipo de exame a ser escrito
        Cells(linha1, 13) = tipoDeExame
        'escrever a contagem
        Cells(linha1, 14) = contaExame
        Cells(linha1, 15) = medico
        'Zeram a contagem para nao ficar cumulativa
        contaExame = 0
        'soma para continuar o laço
        linha1 = linha1 + 1
        'linha dois resetado para recomeçar a contagem
        linha2 = 2
        'laço para comparar os tipos de exames
        'fazer enquanto tiver espaço em branco
        Do While Cells(linha2, 13) <> ""
        'caso o tipo de exame for diferente da primeira linha da escrita dos exames entao ele vai pra proxima linha
        If tipoDeExame <> Cells(linha2, 13) Then
        linha2 = linha2 + 1
        'caso contrario caso seja igual ele vai pra proxima linha no conteudo
        'laço para contagem de exames e para nao repetir
        Else
        Do While ActiveCell = tipoDeExame
        ActiveCell.Offset(1, 0).Select
        linha2 = 1
        Loop
        tipoDeExame = ActiveCell
        linha2 = linha2 + 1
        End If
        Loop
    Loop
    
    Range("h2").Select
End Sub


