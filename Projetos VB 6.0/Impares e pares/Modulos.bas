Attribute VB_Name = "Module1"
Function Modex(Valor As String) As Boolean
If IsNumeric(Valor) = True Then
    If (Valor = Form1.txtN1.Text And Valor > 0 And Valor < 100) Or (Valor = Form1.txtN2.Text And Valor > 1 And Valor < 101 And Valor > Form1.txtN1.Text) Then
        Modex = True
    Else
        Modex = False
    End If
Else
    Modex = False
End If
End Function
Function Parimpa(ByVal Num As Integer) As Boolean
Do Until (Num < 10)
    Num = Num - 10
Loop
If Num = 8 Or Num = 6 Or Num = 4 Or Num = 2 Or Num = 0 Then
    Parimpa = True
Else
    Parimpa = False
End If
End Function
