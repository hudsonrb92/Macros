Attribute VB_Name = "UltimaLinhaPlanilha2"
Function fimColuna(numPlanilha As Integer, colunaDaPlanilha As String) As Integer
    Sheets(numPlanilha).Select
    fimColuna = Range(colunaDaPlanilha & "1048576").End(xlUp).Row
    Sheets(1).Select
End Function
