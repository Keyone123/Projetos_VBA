Sub GeradorDeSenhas()

' Declara as variáveis
Dim vSenha As String
Dim vComprimento As Integer
Dim vTiposDeCaracteres As String

' Pega os critérios do usuário
InputBox "Digite o comprimento da senha:", vComprimento
InputBox "Digite os tipos de caracteres para a senha (ex: letras, números, símbolos):", vTiposDeCaracteres

' Gera a senha
Randomize
vSenha = ""
For i = 1 To vComprimento
vSenha = vSenha & Mid(vTiposDeCaracteres, Int(Rnd * Len(vTiposDeCaracteres)), 1)
Next i

' Mostra a senha para o usuário
MsgBox "A senha é: " & vSenha

End Sub
