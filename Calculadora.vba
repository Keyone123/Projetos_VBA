Sub Calc()

Dim v1 As Double
Dim v2 As Double
Dim resultado As Double


InputBox "Digite o primeiro número:", v1
InputBox "Digite o segundo número:", v2


MsgBox "Escolha a operação que deseja realizar:" & vbNewLine & vbNewLine _
& "1. Soma" & vbNewLine & "2. Subtração" & vbNewLine & "3. Multiplicação" & vbNewLine & "4. Divisão"

Dim opcao As Integer
InputBox "Qual operação você deseja realizar?", opcao


Select Case opcao
Case 1: resultado = v1 + v2
Case 2: resultado = v1 - v2
Case 3: resultado = v1 * v2
Case 4: resultado = v1 / v2
End Select

MsgBox "O resultado é: " & resultado

End Sub
