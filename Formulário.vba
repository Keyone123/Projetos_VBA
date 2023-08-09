Private Sub CommandButton1_Click()

' Declara as variáveis
Dim vNome As String
Dim vEmail As String
Dim vTelefone As String

' Pega os dados do usuário
vNome = InputBox("Digite o seu nome:")
vEmail = InputBox("Digite o seu email:")
vTelefone = InputBox("Digite o seu telefone:")

' Valida os dados
If vNome = "" Then
MsgBox "O nome é obrigatório."
Exit Sub
End If

If Not IsNumeric(vTelefone) Then
MsgBox "O telefone deve ser um número."
Exit Sub
End If

If Len(vTelefone) < 10 Then
MsgBox "O telefone deve ter pelo menos 10 dígitos."
Exit Sub
End If

If Not IsEmail(vEmail) Then
MsgBox "O email é inválido."
Exit Sub
End If

' Salva os dados no banco de dados
Dim vConn As New ADODB.Connection
Dim vCmd As New ADODB.Command

vConn.Open ""
vCmd.ActiveConnection = vConn
vCmd.CommandText = "INSERT INTO cadastro (nome, email, telefone) VALUES ('" & vNome & "', '" & vEmail & "', '" & vTelefone & "')"
vCmd.Execute

vConn.Close

' Limpa os campos
vNome = ""
vEmail = ""
vTelefone = ""

' Mostra uma mensagem de confirmação
MsgBox "Os dados foram salvos com sucesso!"

End Sub

Private Function IsEmail(ByVal strEmail As String) As Boolean

' Valida o formato do email
IsEmail = InStr(1, strEmail, "@") > 0 And InStr(1, strEmail, ".") > 0

End Function
