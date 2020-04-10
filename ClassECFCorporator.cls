VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassECFCorporator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsMD5PAFECF As String
Private mvarcolECFAutorizado As New Collection

Public Property Set colECFAutorizado(ByVal vData As Collection)
    Set mvarcolECFAutorizado = vData
End Property

Public Property Get colECFAutorizado() As Collection
    Set colECFAutorizado = mvarcolECFAutorizado
End Property

Public Property Get sMD5PAFECF() As String
    sMD5PAFECF = mvarsMD5PAFECF
End Property

Public Property Let sMD5PAFECF(ByVal vData As String)
    mvarsMD5PAFECF = vData
End Property





