VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTabPrecoItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlCodTabela As Long
Private mvarsProduto As String
Private mvardtDataVigencia As Date
Private mvardPreco As Double
Private mvarsDescricao As String

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let dPreco(ByVal vData As Double)
    mvardPreco = vData
End Property

Public Property Get dPreco() As Double
    dPreco = mvardPreco
End Property

Public Property Let dtDataVigencia(ByVal vData As Date)
    mvardtDataVigencia = vData
End Property

Public Property Get dtDataVigencia() As Date
    dtDataVigencia = mvardtDataVigencia
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lCodTabela(ByVal vData As Long)
    mvarlCodTabela = vData
End Property

Public Property Get lCodTabela() As Long
    lCodTabela = mvarlCodTabela
End Property



