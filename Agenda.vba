Sub Agenda()

Dim vData As Date
Dim vHora As Time
Dim vDescricao As String
Dim vRepetir As Boolean
Dim vIntervalo As Integer
Dim vPeriodo As String
Dim vPrioridade As Integer
Dim vLembrete As Boolean
Dim vTempoAntes As Integer

InputBox "Digite a data do compromisso:", vData
InputBox "Digite a hora do compromisso:", vHora
InputBox "Digite a descrição do compromisso:", vDescricao
InputBox "Deseja repetir o compromisso?", vRepetir
If vRepetir = True Then
InputBox "Digite o intervalo de repetição (em dias):", vIntervalo
InputBox "Digite o período de repetição (dia da semana):", vPeriodo
End If
InputBox "Digite a prioridade do compromisso (1-5):", vPrioridade
InputBox "Deseja receber um lembrete do compromisso?", vLembrete
If vLembrete = True Then
InputBox "Digite o tempo antes do compromisso para receber o lembrete (em minutos):", vTempoAntes
End If

Dim vCompromisso As New Appointment
vCompromisso.Date = vData
vCompromisso.Time = vHora
vCompromisso.Description = vDescricao
vCompromisso.Repeat = vRepetir
vCompromisso.Interval = vIntervalo
vCompromisso.Period = vPeriodo
vCompromisso.Priority = vPrioridade
vCompromisso.Reminder = vLembrete
vCompromisso.ReminderTime = vTempoAntes

Application.Save

If vLembrete = True Then
Dim vDataAtual As Date
Dim vHoraAtual As Time
vDataAtual = Date
vHoraAtual = Time
If vDataAtual = vCompromisso.Date And vHoraAtual = vCompromisso.Time Then
MsgBox "O seu compromisso está próximo!"
Else
Dim vTempoFaltante As Double
vTempoFaltante = DateDiff("n", vDataAtual, vCompromisso.Date) * 24 * 60 * 60 + DateDiff("s", vHoraAtual, vCompromisso.Time)
If vTempoFaltante < vTempoAntes Then
MsgBox "O seu compromisso está a " & vTempoFaltante & " minutos de começar!"
End If
End If
End If

End Sub
