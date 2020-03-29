Attribute VB_Name = "Module1"
Global Conexao As New ADODB.Connection
Sub limp(formulario As Form)
    For Each controle In formulario
        If TypeOf controle Is TextBox Then
            controle.Text = ""
        ElseIf TypeOf controle Is ComboBox Then
            controle.Clear
        ElseIf TypeOf controle Is ListBox Then
            controle.Clear
        End If
    Next
End Sub
Sub estado(formulario As Form)
sql = "Select * from Estados"
Set pkestado = Conexao.Execute(sql)
If Not pkestado.EOF Then
    Do While Not pkestado.EOF
        formulario.cmbEstado.AddItem pkestado("UF")
        pkestado.MoveNext
    Loop
End If
End Sub
Sub carregarlist(tabela As String, list As ListBox, campo1 As String, campo2 As String)
sql = "Select * from " & tabela
    Set tabeladinamica = Conexao.Execute(sql)
    If tabeladinamica.EOF Then
    Else
        list.Clear
        Do While Not tabeladinamica.EOF
            list.AddItem tabeladinamica(campo1) & " - " & tabeladinamica(campo2)
            tabeladinamica.MoveNext
        Loop
    End If
End Sub
Function VerCod(tabela As String, campo As String, cod As String) As Boolean
sql = "Select * from " & tabela & " where " & cod & " = '" & campo & "'"
Set tabelax = Conexao.Execute(sql)
If tabelax.EOF Then
    VerCod = True
Else
    VerCod = False
End If
End Function


