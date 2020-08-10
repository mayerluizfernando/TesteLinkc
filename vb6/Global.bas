Attribute VB_Name = "Global"
Public ListaInOut As New Collection
Public ListaCarro As New Collection
Public oListaInOut As New cListaInOut
Public oListaCarro As New cListaCarro
Public oEstacionamento As New cEstacionamento

Public Enum StatusProcessamento
    'Entrada
    AguardandoIniProcEntrada = 0 'WaitBeginProcess = 0
    ProcEntradaEmAndamento = 1 'RunProcess = 1
    ProcEntradaFinalizado = 2  'EndProcess = 2
    'Saida
    AguardandoIniProcSaida = 3
    ProcSaidaEmAndamento = 4
    ProcSaidaFinalizado = 5
    
End Enum


'Public arrayIntOut() As cInOut

Public Function DescribeStatus(pStatus As Variant) As String
    DescribeStatus = vbNullString
    Select Case pStatus
        'Entrada
        Case StatusProcessamento.AguardandoIniProcEntrada
            DescribeStatus = "Aguardando Inicio do Processamento de Entrada"
        Case StatusProcessamento.ProcEntradaEmAndamento
            DescribeStatus = "Processo entrada em andamento"
        Case StatusProcessamento.ProcEntradaFinalizado
            DescribeStatus = "Processo entrada finalizado"
        'Saida
        Case StatusProcessamento.AguardandoIniProcSaida
            DescribeStatus = "Aguardando Inicio do Processamento de Saída"
        Case StatusProcessamento.ProcSaidaEmAndamento
            DescribeStatus = "Processo saída em andamento"
        Case StatusProcessamento.ProcSaidaFinalizado
            DescribeStatus = "Processo saída finalizado"
            
        Case Else
            DescribeStatus = "Erro ###Status Não Definido###"
    End Select
End Function

Public Sub InsereLog(rtf As RichTextBox, sMensagem As String)
    'rtf.Text = sMensagem & vbCrLf & rtf.Text
    rtf.Text = rtf.Text & vbCrLf & sMensagem
End Sub

Public Function RandomNumber(ByVal MaxValue As Long, Optional _
ByVal MinValue As Long = 0)

  On Error Resume Next
  Randomize Timer
  RandomNumber = Int((MaxValue - MinValue + 1) * Rnd) + MinValue

End Function

