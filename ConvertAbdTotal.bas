Attribute VB_Name = "ConvertAbdTotal"
Sub convert2()
    Dim abd As String
    Dim nAbd As String
    Dim ultimaLinha As Integer
    Dim i As Integer
    Dim fimA As Integer
    Dim fimB As Integer
    Dim i2 As Integer
    Dim pergunta As VbMsgBoxResult
    
    
    fimA = fimColuna(2, "A")
    fimB = fimColuna(2, "B")
    Sheets(1).Select
    ultimaLinha = Range("f1048576").End(xlUp).Row
    
        For i = 2 To ultimaLinha
        abd = Cells(i, 6)
            If Cells(i, 6) Like UCase("*a*b*d*") Or Cells(i, 6) Like UCase("*uro*") Or Cells(i, 6) Like UCase("*vias*") Then
            If Cells(i, 8) = "CT" Then
inicio:
                fimA = fimColuna(2, "A")
                For i2 = 2 To fimA
                Sheets(2).Select
                If abd = Cells(i2, 1) Then
                Sheets(1).Select
                Cells(i, 8) = "CTA"
                Cells(i, 6) = "ABDOMETOTAL"
                Sheets(1).Select
                GoTo proximo
                End If
                Next
                
                fimB = fimColuna(2, "B")
                For i2 = 2 To fimB
                Sheets(2).Select
                If abd = Cells(i2, 2) Then
                Sheets(1).Select
                GoTo proximo
                End If
                Next
                
            Sheets(1).Select
            
            pergunta = MsgBox(abd, vbYesNo, "CT Abd Sim ou Nao")
            If pergunta = vbYes Then
            fimA = fimColuna(2, "A")
            Sheets(2).Select
            Cells(fimA + 1, 1) = abd
            Sheets(1).Select
            GoTo inicio
            Else
            fimB = fimColuna(2, "B")
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
