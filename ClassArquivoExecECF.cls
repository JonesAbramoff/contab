VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassECFAutorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNumFabricacao As String
Private mvarsGT As String

Public Property Get sGT() As String
    sGT = mvarsGT
End Property

Public Property Let sGT(ByVal vData As String)
    mvarsGT = vData
End Property

Public Property Get sNumFabricacao() As String
    sNumFabricacao = mvarsNumFabricacao
End Property

Public Property Let sNumFabricacao(ByVal vData As String)
    mvarsNumFabricacao = vData
End Property

