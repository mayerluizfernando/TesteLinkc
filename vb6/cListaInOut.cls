VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListaInOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Event Change()

Private Enum Status
    Liberada = 1    'Liberada
    Ocupada = 2     'Ocupada
End Enum

Public Function ESDisponivel(sTipo As String) As String
Dim i As Long

ESDisponivel = vbNullString
For Each es In ListaInOut
    If es.Tipo = sTipo Then
        If es.Status = Status.Liberada Then
            ESDisponivel = es.ID
            es.Qtde = es.Qtde + 1
            Exit Function
        End If
    End If
Next

End Function

Public Sub InitListaInOut(Entradas As Long, Saidas As Long)
Dim i As Long
Dim oInOut As cInOut

Set ListaInOut = New Collection

For i = 1 To Entradas
    Set oInOut = New cInOut
    oInOut.ID = "E" & CStr(i)
    oInOut.Tipo = "E"
    'oInOut.Status = ""
    oInOut.Status = Status.Liberada
    oInOut.Qtde = 0
    ListaInOut.Add oInOut, oInOut.ID
Next i

For i = 1 To Saidas
    Set oInOut = New cInOut
    oInOut.ID = "S" & CStr(i)
    oInOut.Tipo = "S"
    'oInOut.Status = ""
    oInOut.Status = Status.Liberada
    oInOut.Qtde = 0
    ListaInOut.Add oInOut, oInOut.ID
Next i

End Sub

Public Function ObjCollection(col As Collection, sKey As String) As Object

Dim bTest As Boolean

    On Error Resume Next
    
    bTest = IsObject(col(sKey))
    If (Err = 0) Then
        'InCollection = True
        Set ObjCollection = col(sKey)
    Else
        Err.Clear
    End If

End Function

Public Sub OcupaES(col As Collection, sKey As String)
    On Error Resume Next
    Dim oInOut As cInOut
    Set oInOut = col(sKey)
    'oInOut.Status = "OCUPADA"
    oInOut.Status = Status.Ocupada
End Sub

Public Sub LiberaES(col As Collection, sKey As String)
    On Error Resume Next
    Dim oInOut As cInOut
    Set oInOut = col(sKey)
    'oInOut.Status = "LIBERADA"
    oInOut.Status = Status.Liberada
End Sub

Public Sub AddMovimento(col As Collection, sKey As String)
    On Error Resume Next
    Dim oInOut As cInOut
    Set oInOut = col(sKey)
    oInOut.Qtde = oInOut.Qtde + 1
    'RaiseEvent Change
End Sub

