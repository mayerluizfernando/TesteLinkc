VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCarro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'local variable(s) to hold property value(s)
Private mvarID As String 'local copy
Private mvarTempoInicioEntrada As Date 'local copy
Private mvarTempoInicioSaida As Date 'local copy
Private mvarTempoFimProcessamentoEntrada As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarTempo As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarStatusProcessamento As String 'local copy
'local variable(s) to hold property value(s)
Private mvarTempoFimProcessamentoSaida As Date 'local copy
'local variable(s) to hold property value(s)
Private mvarKeyEntrada As String 'local copy
Private mvarKeySaida As String 'local copy
'local variable(s) to hold property value(s)
Private mvarFilaEntrada As Integer 'local copy
Private mvarFilaSaida As Integer 'local copy
Public Property Let FilaSaida(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FilaSaida = 5
    mvarFilaSaida = vData
End Property


Public Property Get FilaSaida() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FilaSaida
    FilaSaida = mvarFilaSaida
End Property



Public Property Let FilaEntrada(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FilaEntrada = 5
    mvarFilaEntrada = vData
End Property


Public Property Get FilaEntrada() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FilaEntrada
    FilaEntrada = mvarFilaEntrada
End Property




Public Property Let KeySaida(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.KeySaida = 5
    mvarKeySaida = vData
End Property


Public Property Get KeySaida() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.KeySaida
    KeySaida = mvarKeySaida
End Property

Public Property Let KeyEntrada(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.KeyEntrada = 5
    mvarKeyEntrada = vData
End Property


Public Property Get KeyEntrada() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.KeyEntrada
    KeyEntrada = mvarKeyEntrada
End Property



Public Property Let TempoFimProcessamentoSaida(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TempoFimProcessamentoSaida = 5
    mvarTempoFimProcessamentoSaida = vData
End Property


Public Property Get TempoFimProcessamentoSaida() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TempoFimProcessamentoSaida
    TempoFimProcessamentoSaida = mvarTempoFimProcessamentoSaida
End Property

Public Property Let StatusProcessamento(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StatusProcessamento = 5
    mvarStatusProcessamento = vData
End Property


Public Property Get StatusProcessamento() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StatusProcessamento
    StatusProcessamento = mvarStatusProcessamento
End Property



Public Property Let Tempo(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tempo = 5
    mvarTempo = vData
End Property


Public Property Get Tempo() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tempo
    Tempo = mvarTempo
End Property



Public Property Let TempoFimProcessamentoEntrada(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TempoFimProcessamentoEntrada = 5
    mvarTempoFimProcessamentoEntrada = vData
End Property


Public Property Get TempoFimProcessamentoEntrada() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TempoFimProcessamentoEntrada
    TempoFimProcessamentoEntrada = mvarTempoFimProcessamentoEntrada
End Property



Public Property Let TempoInicioSaida(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TempoInicioSaida = 5
    mvarTempoInicioSaida = vData
End Property


Public Property Get TempoInicioSaida() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TempoInicioSaida
    TempoInicioSaida = mvarTempoInicioSaida
End Property



Public Property Let TempoInicioEntrada(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TempoInicioEntrada = 5
    mvarTempoInicioEntrada = vData
End Property


Public Property Get TempoInicioEntrada() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TempoInicioEntrada
    TempoInicioEntrada = mvarTempoInicioEntrada
End Property



Public Property Let ID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID = 5
    mvarID = vData
End Property


Public Property Get ID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property



