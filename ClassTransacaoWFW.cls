VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTransacaoWFW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsSigla As String
Private mvarsTransacao As String
Private mvarsTransacaoTela As String
Private mvarsOrigem As String
Private mvarsObservacao As String
Private mvariCodigo As Integer

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let sOrigem(ByVal vData As String)
    mvarsOrigem = vData
End Property

Public Property Get sOrigem() As String
    sOrigem = mvarsOrigem
End Property

Public Property Let sTransacaoTela(ByVal vData As String)
    mvarsTransacaoTela = vData
End Property

Public Property Get sTransacaoTela() As String
    sTransacaoTela = mvarsTransacaoTela
End Property

Public Property Let sTransacao(ByVal vData As String)
    mvarsTransacao = vData
End Property

Public Property Get sTransacao() As String
    sTransacao = mvarsTransacao
End Property

Public Property Let sSigla(ByVal vData As String)
    mvarsSigla = vData
End Property

Public Property Get sSigla() As String
    sSigla = mvarsSigla
End Property

