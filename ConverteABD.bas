Attribute VB_Name = "ConverteABD"
Sub coverteABDT()
    Dim varProc As Integer
    Dim procu As String
    Dim respos As VbMsgBoxResult
    ultimaLinha = Range("a1048576").End(xlUp).Row
    Dim abdt As String
    Dim i As Integer
        For varProc = 2 To ultimaLinha
        Range("F" & varProc).Select
        procu = UCase(Range("f" & varProc).Value) Like UCase("*a*b*d*t*")
        porcu2 = UCase(Range("f" & varProc).Value) Like UCase("*uro*")
            If procu = True Or procu2 = True Then
                If Cells(varProc, 8) <> "CTA" And Cells(varProc, 8) = "CT" Then
                abdt = ActiveCell
                repos = MsgBox(abdt, vbYesNo, "É abd T?")
                    If repos = vbYes Then
                    For i = 2 To ultimaLinha
                        If Cells(i, 6) = UCase(Range("f" & varProc).Value) Then
                        Cells(i, 6).Value = "ABDTOTAL"
                        Cells(i, 8).Value = "CTA"
                        End If
                    Next
                    End If
                End If
            End If
        Next
        Range("h2").Select
    
End Sub
