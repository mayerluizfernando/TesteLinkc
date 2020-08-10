VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEstacionamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarQtde As Long 'local copy

Public Function Soma() As Long
    mvarQtde = mvarQtde + 1
    Soma = mvarQtde
End Function

Public Function Subtrai() As Long
    mvarQtde = mvarQtde - 1
    Subtrai = mvarQtde
End Function

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



