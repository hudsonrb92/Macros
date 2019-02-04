Attribute VB_Name = "ImportaPlanilhas"
Sub importa()
    Dim arquivo As String
    
    arquivo = Application.GetOpenFilename
    Workbooks.Open arquivo
    Sheets(1).Select
    Range("a2").Select
    Range(ActiveCell, ActiveCell.End(xlDown).End(xlToRight)).Copy
    Workbooks("Main.xlsm").Activate
    Range("a1048576").Select
    ActiveCell.End(xlUp).Select
    ActiveCell.PasteSpecial xlValues
    
    
End Sub
