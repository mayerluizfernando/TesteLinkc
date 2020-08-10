VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTeste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private FText As String

Public Event OnChange(ByVal Text As String)

Private Sub Change(ByVal Text As String)
  RaiseEvent OnChange(Text)
End Sub

Public Property Get Text() As String
  Text = FText
End Property


Public Property Let Text(ByVal Value As String)
  FText = Value
  Call Change(Value)
End Property
