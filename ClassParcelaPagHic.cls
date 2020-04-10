VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassParcelaPagHic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private mvarsFavorecido As String

Public Property Let sFavorecido(ByVal vData As String)
    mvarsFavorecido = vData
End Property

Public Property Get sFavorecido() As String
    sFavorecido = mvarsFavorecido
End Property

