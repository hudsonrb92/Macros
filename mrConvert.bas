Attribute VB_Name = "mrConvert"
Sub convert2()
    Dim abd As String
    Dim nAbd As String
    Dim ultimaLinha As Integer
    Dim i As Integer
    Dim fimA As Integer
    Dim fimB As Integer
    Dim i2 As Integer
    Dim pergunta As VbMsgBoxResult
    
    
    fimA = fimColunaA("2")
    fimB = fimColunaB("2")
    Sheets(1).Select
    ultimaLinha = Range("f1048576").End(xlUp).Row
    
        For i = 2 To ultimaLinha
        abd = Cells(i, 6)
            If Cells(i, 6) Like UCase("*a*b*d*t*") Then
            If Cells(i, 8) = "MR" Then
inicio:
                For i2 = 2 To fimA
                Sheets(2).Select
                If abd = Cells(i2, 3) Then
                Sheets(1).Select
                Cells(i, 8) = "MRA"
                Cells(i, 6) = "ABDOMETOTAL"
                fimA = fimColunaA("2")
                Sheets(1).Select
                GoTo proximo
                End If
                Next
                
                For i2 = 2 To fimB
                Sheets(2).Select
                If abd = Cells(i2, 4) Then
                fimB = fimColunaB("2")
                Sheets(1).Select
                GoTo proximo
                End If
                Next
                
            Sheets(1).Select
            
            pergunta = MsgBox(abd, vbYesNo, "Abd Sim ou Nao")
            If pergunta = vbYes Then
            fimA = fimColunaA("2")
            Sheets(2).Select
            Cells(fimA + 1, 1) = abd
            Sheets(1).Select
            GoTo inicio
            Else
            fimB = fimColunaB("2")
            Sheets(2).Select
            Cells(fimB + 1, 2) = abd
            Sheets(1).Select
            GoTo inicio
            End If
            End If
            End If
proximo:
          
        Next
    
End Sub
