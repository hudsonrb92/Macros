Attribute VB_Name = "CalculaValoresMedicos"
Sub calculaValoresMedicos()
    Dim valor As Double
    Dim exame As String
    Dim i As Integer
    Dim ultimaLinha As Integer
    Dim estabe As String
    
    ultimaLinha = fimColuna(3, "A")
    Sheets(3).Select
    For i = 2 To ultimaLinha
    exame = Cells(i, 2)
    Select Case exame
        Case "CR"
        Cells(i, 4) = Cells(i, 3) * 4
        Case "MR"
        Cells(i, 4) = Cells(i, 3) * 38
        Case "DX"
        Cells(i, 4) = Cells(i, 3) * 4
        Case "MG"
        Cells(i, 4) = Cells(i, 3) * 8
        Case "CT"
        Cells(i, 4) = Cells(i, 3) * 25
        Case "CTA"
        Cells(i, 4) = Cells(i, 3) * 40
        Case "MRA"
        Cells(i, 4) = Cells(i, 3) * (38 * 2)
        
    End Select
    Next
    
    ultimaLinha = fimColuna(4, "A")
    Sheets(4).Select
    
    For i = 2 To ultimaLinha
    exame = Cells(i, 2)
    estabe = Cells(i, 1)
    
    Select Case exame
    
    Case "CR"
    If Cells(i, 1) Like "*Materninadade Santa*" Then
    Cells(i, 4) = Cells(i, 3) * 8.3
    End If
    If Cells(i, 1) = "Hospital Santa Marta" Then
    Cells(i, 4) = Cells(i, 3) * 7.8
    End If
    If Cells(i, 1) = "ULTRAMED SANTA JULIANA" Then
    Cells(i, 4) = Cells(i, 3) * 8.5
    End If
    If Cells(i, 1) = "Uberlândia Medical Center" Then
    Cells(i, 4) = Cells(i, 3) * 10
    Else
    Cells(i, 4) = Cells(i, 3) * 9
    End If
    
    Case "DX"
    If Cells(i, 1) Like "*Materninadade Santa*" Then
    Cells(i, 4) = Cells(i, 3) * 8.3
    End If
    If Cells(i, 1) = "Hospital Santa Marta" Then
    Cells(i, 4) = Cells(i, 3) * 7.8
    End If
    If Cells(i, 1) = "ULTRAMED SANTA JULIANA" Then
    Cells(i, 4) = Cells(i, 3) * 8.5
    End If
    If Cells(i, 1) = "Uberlândia Medical Center" Then
    Cells(i, 4) = Cells(i, 3) * 10
    Else
    Cells(i, 4) = Cells(i, 3) * 9
    End If
    
    Case "CT"
    If Cells(i, 1) Like "*Materninadade Santa*" Then
    Cells(i, 4) = Cells(i, 3) * 49
    End If
    If Cells(i, 1) = "Hospital Santa Marta" Then
    Cells(i, 4) = Cells(i, 3) * 55
    End If
    If Cells(i, 1) = "Diagnóstico Centro de Medicina Avançada" Then
    Cells(i, 4) = Cells(i, 3) * 47
    End If
    If Cells(i, 1) = "Vital Imagem" Then
    Cells(i, 4) = Cells(i, 3) * 40
    End If
    If Cells(i, 1) = "Tomografia Santa Helena" Then
    Cells(i, 4) = Cells(i, 3) * 53
    End If
    If Cells(i, 1) = "Med-Center" Then
    Cells(i, 4) = Cells(i, 3) * 55
    End If
    
    Case "CTA"
    If Cells(i, 1) Like "*Materninadade Santa*" Then
    Cells(i, 4) = Cells(i, 3) * 90
    End If
    If Cells(i, 1) = "Hospital Santa Marta" Then
    Cells(i, 4) = Cells(i, 3) * 90
    End If
    If Cells(i, 1) = "Diagnóstico Centro de Medicina Avançada" Then
    Cells(i, 4) = Cells(i, 3) * 94
    End If
    If Cells(i, 1) = "Vital Imagem" Then
    Cells(i, 4) = Cells(i, 3) * 80
    End If
    If Cells(i, 1) = "Tomografia Santa Helena" Then
    Cells(i, 4) = Cells(i, 3) * 83
    End If
    If Cells(i, 1) Like "Med-Center" Then
    Cells(i, 4) = Cells(i, 3) * 88
    End If
    
    Case "MG"
    Cells(i, 4) = Cells(i, 3) * 20
    
    Case "MR"
    
    Cells(i, 4) = Cells(i, 3) * 55
    
    Case "MRA"
    
    Cells(i, 4) = Cells(i, 3) * 110
    End Select
    Next
    
End Sub
