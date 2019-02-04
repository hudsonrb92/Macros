Attribute VB_Name = "PegaMedico"
Sub pegadr()
    Dim i As Integer
    Dim i2 As Integer
    Dim colunaReq As Integer
    Dim colunaReq2 As Integer
    Dim colunaMed As Integer
    Dim colunaMed2 As Integer
    Dim req As Double
    Dim med As String
    Dim ultLin1 As Integer
    Dim ultLin2 As Integer
    
    colunaMed = 14
    colunaMed2 = 5
    colunaReq = 3
    colunaReq2 = 2
    ultLin1 = fimColuna(1, "A")
    ultLin2 = fimColuna(2, "A")
    
    For i = 2 To ultLin1
        Sheets(1).Select
        req = Cells(i, colunaReq)
        
        For i2 = 2 To ultLin2
            Sheets(2).Select
            If Cells(i2, colunaReq2) = req Then
                med = Cells(i2, colunaMed2)
                Sheets(1).Select
                Cells(i, colunaMed) = med
                GoTo proximo
            End If
        Next
proximo:
    Next
    
End Sub
