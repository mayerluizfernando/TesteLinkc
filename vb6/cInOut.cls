VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'local variable(s) to hold property value(s)
Private mvarTipo As String 'local copy
Private mvarID As String 'local copy
Private mvarStatus As String 'local copy
Private mvarQtde As Long 'local copy

Public Sub InitInOut(QtEntradas As Long, QtSaidas As Long)

    ReDim arrayIntOut(0)
    ReDim Preserve arrayIntOut(QtEntradas + QtSaidas)
    
    Dim oInOut As cInOut
    Set oInOut = New cInOut


End Sub



Public Property Let Qtde(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Qtde = 5
    mvarQtde = vData
End Property


Public Property Get Qtde() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Qtde
    Qtde = mvarQtde
End Property



Public Property Let Status(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Status = 5
    mvarStatus = vData
End Property


Public Property Get Status() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Status
    Status = mvarStatus
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



Public Property Let Tipo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tipo = 5
    mvarTipo = vData
End Property


Public Property Get Tipo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tipo
    Tipo = mvarTipo
End Property



