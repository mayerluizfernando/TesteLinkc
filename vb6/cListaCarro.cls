VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListaCarro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Public Function QtdeFilaEntrada() As Long
    QtdeFilaEntrada = 0
    For Each c In ListaCarro
        QtdeFilaEntrada = QtdeFilaEntrada + c.FilaEntrada
    Next
    '####
End Function

Public Function QtdeFilaSaida() As Long
    QtdeFilaSaida = 0
    For Each c In ListaCarro
        QtdeFilaSaida = QtdeFilaSaida + c.FilaSaida
    Next
End Function

Public Sub AddListaCarro(Carros As Long)
Dim i As Long
Dim oCarro As cCarro

'Set ListaCarro = New Collection
'i = ListaCarro.Count
For i = ListaCarro.Count + 1 To Carros + ListaCarro.Count
    Set oCarro = New cCarro
    oCarro.ID = "C" & CStr(i)
    oCarro.Tempo = Now
    '############ Entrada ############
    'oCarro.TempoInicioEntrada = DateAdd("s", 10, CDate(oCarro.Tempo))
    oCarro.TempoInicioEntrada = DateAdd("s", RandomNumber(1, 5), CDate(oCarro.Tempo))
    'oCarro.TempoFimProcessamentoEntrada = DateAdd("s", 5, CDate(oCarro.TempoInicioEntrada))
    oCarro.TempoFimProcessamentoEntrada = DateAdd("s", RandomNumber(3, 5), CDate(oCarro.TempoInicioEntrada))
    
    '############ Saída ############
    'oCarro.TempoInicioSaida = DateAdd("s", 10, CDate(oCarro.TempoFimProcessamentoEntrada))
    oCarro.TempoInicioSaida = DateAdd("s", RandomNumber(10, 30), CDate(oCarro.TempoFimProcessamentoEntrada))
    
    'oCarro.TempoFimProcessamentoSaida = DateAdd("s", 5, CDate(oCarro.TempoInicioSaida))
    oCarro.TempoFimProcessamentoSaida = DateAdd("s", RandomNumber(3, 5), CDate(oCarro.TempoInicioSaida))
    
    '############ Status Inicial ############
    oCarro.StatusProcessamento = StatusProcessamento.AguardandoIniProcEntrada
    
    ListaCarro.Add oCarro, oCarro.ID
    
    Debug.Print "AddListaCarro Id:" & oCarro.ID & " TempoInicioEntrada:" & oCarro.TempoInicioEntrada & _
        " TempoFimProcessamentoEntrada:" & oCarro.TempoFimProcessamentoEntrada
    
Next i

End Sub

