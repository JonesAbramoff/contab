VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItensDeContratoSrvProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntItemContratoSrv As Long
Private mvarsProduto As String

Public Property Let lNumIntItemContratoSrv(ByVal vData As Long)
    mvarlNumIntItemContratoSrv = vData
End Property

Public Property Get lNumIntItemContratoSrv() As Long
    lNumIntItemContratoSrv = mvarlNumIntItemContratoSrv
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

