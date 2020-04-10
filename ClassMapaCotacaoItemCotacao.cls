VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMapaCotacaoItemCotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlCodMapaCotacao As Long
Private mvariFilialEmpresa As Integer
Private mvarlNumIntItemCotacao As Long
Private mvarlPedidoCotacao As Long
Private mvarsProduto As String
Private mvarsDescricao As String
Private mvardQuantidade As Double
Private mvarsFornecedor As String
Private mvarsUM As String
Private mvariFilialForn As Integer
Private mvarsNomeFilialForn As String
Public Property Let sNomeFilialForn(ByVal vData As String)
    mvarsNomeFilialForn = vData
End Property

Public Property Get sNomeFilialForn() As String
    sNomeFilialForn = mvarsNomeFilialForn
End Property

Public Property Let iFilialForn(ByVal vData As Integer)
    mvariFilialForn = vData
End Property

Public Property Get iFilialForn() As Integer
    iFilialForn = mvariFilialForn
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let sFornecedor(ByVal vData As String)
    mvarsFornecedor = vData
End Property

Public Property Get sFornecedor() As String
    sFornecedor = mvarsFornecedor
End Property

Public Property Let lNumIntItemCotacao(ByVal vData As Long)
    mvarlNumIntItemCotacao = vData
End Property

Public Property Get lNumIntItemCotacao() As Long
    lNumIntItemCotacao = mvarlNumIntItemCotacao
End Property

Public Property Let lCodMapaCotacao(ByVal vData As Long)
    mvarlCodMapaCotacao = vData
End Property

Public Property Get lCodMapaCotacao() As Long
    lCodMapaCotacao = mvarlCodMapaCotacao
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lPedidoCotacao(ByVal vData As Long)
    mvarlPedidoCotacao = vData
End Property

Public Property Get lPedidoCotacao() As Long
    lPedidoCotacao = mvarlPedidoCotacao
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property
