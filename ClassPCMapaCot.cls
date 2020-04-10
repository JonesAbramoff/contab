VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPCMapaCot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodigo As Long
Private mvarsFornecedor As String
Private mvardtData As Date
Private mvardtDataEntrega As Date
Private mvardQuantidade As Double
Private mvardTaxa As Double
Private mvardPreco As Double
Private mvardIPI As Double
Private mvardDesconto As Double
Private mvarsCondPagto As String
Private mvarsStatus As String
Private mvardFrete As Double
Private mvarsFreteTipo As String

Public Property Let sFreteTipo(ByVal vData As String)
    mvarsFreteTipo = vData
End Property

Public Property Get sFreteTipo() As String
    sFreteTipo = mvarsFreteTipo
End Property

Public Property Let dFrete(ByVal vData As Double)
    mvardFrete = vData
End Property

Public Property Get dFrete() As Double
    dFrete = mvardFrete
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let sFornecedor(ByVal vData As String)
    mvarsFornecedor = vData
End Property

Public Property Get sFornecedor() As String
    sFornecedor = mvarsFornecedor
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dtDataEntrega(ByVal vData As Date)
    mvardtDataEntrega = vData
End Property

Public Property Get dtDataEntrega() As Date
    dtDataEntrega = mvardtDataEntrega
End Property

Public Property Let dDesconto(ByVal vData As Double)
    mvardDesconto = vData
End Property

Public Property Get dDesconto() As Double
    dDesconto = mvardDesconto
End Property

Public Property Let sCondPagto(ByVal vData As String)
    mvarsCondPagto = vData
End Property

Public Property Get sCondPagto() As String
    sCondPagto = mvarsCondPagto
End Property

Public Property Let sStatus(ByVal vData As String)
    mvarsStatus = vData
End Property

Public Property Get sStatus() As String
    sStatus = mvarsStatus
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let dPreco(ByVal vData As Double)
    mvardPreco = vData
End Property

Public Property Get dPreco() As Double
    dPreco = mvardPreco
End Property

Public Property Let dIPI(ByVal vData As Double)
    mvardIPI = vData
End Property

Public Property Get dIPI() As Double
    dIPI = mvardIPI
End Property

Public Property Let dTaxa(ByVal vData As Double)
    mvardTaxa = vData
End Property

Public Property Get dTaxa() As Double
    dTaxa = mvardTaxa
End Property

