VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFilFornFilEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlCodFornecedor As Long
Private mvariCodFilial As Integer
Private mvarlNumeroCompras As Long
Private mvardMediaCompra As Double
Private mvardtDataPrimeiraCompra As Date
Private mvardtDataUltimaCompra As Date
Private mvardValorAcumuladoCompras As Double
Private mvardValorAcumuladoDevolucoes As Double
Private mvarlNumTotalDevolucoes As Long
Private mvardtDataUltDevolucao As Date
Private mvarlMaiorAtraso As Long
Private mvarlPedidosEmAberto As Long
Private mvarlAtrasoAcumulado As Long
Private mvarlPedidosRecebidos As Long
Private mvarlItensPedidosRecebidos As Long

Public Property Let lItensPedidosRecebidos(ByVal vData As Long)
    mvarlItensPedidosRecebidos = vData
End Property

Public Property Get lItensPedidosRecebidos() As Long
    lItensPedidosRecebidos = mvarlItensPedidosRecebidos
End Property

Public Property Let lPedidosRecebidos(ByVal vData As Long)
    mvarlPedidosRecebidos = vData
End Property

Public Property Get lPedidosRecebidos() As Long
    lPedidosRecebidos = mvarlPedidosRecebidos
End Property

Public Property Let lAtrasoAcumulado(ByVal vData As Long)
    mvarlAtrasoAcumulado = vData
End Property

Public Property Get lAtrasoAcumulado() As Long
    lAtrasoAcumulado = mvarlAtrasoAcumulado
End Property

Public Property Let lPedidosEmAberto(ByVal vData As Long)
    mvarlPedidosEmAberto = vData
End Property

Public Property Get lPedidosEmAberto() As Long
    lPedidosEmAberto = mvarlPedidosEmAberto
End Property

Public Property Let lMaiorAtraso(ByVal vData As Long)
    mvarlMaiorAtraso = vData
End Property

Public Property Get lMaiorAtraso() As Long
    lMaiorAtraso = mvarlMaiorAtraso
End Property

Public Property Let dValorAcumuladoDevolucoes(ByVal vData As Double)
    mvardValorAcumuladoDevolucoes = vData
End Property

Public Property Get dValorAcumuladoDevolucoes() As Double
    dValorAcumuladoDevolucoes = mvardValorAcumuladoDevolucoes
End Property

Public Property Let lNumTotalDevolucoes(ByVal vData As Long)
    mvarlNumTotalDevolucoes = vData
End Property

Public Property Get lNumTotalDevolucoes() As Long
    lNumTotalDevolucoes = mvarlNumTotalDevolucoes
End Property

Public Property Let dtDataUltDevolucao(ByVal vData As Date)
    mvardtDataUltDevolucao = vData
End Property

Public Property Get dtDataUltDevolucao() As Date
    dtDataUltDevolucao = mvardtDataUltDevolucao
End Property

Public Property Let dValorAcumuladoCompras(ByVal vData As Double)
    mvardValorAcumuladoCompras = vData
End Property

Public Property Get dValorAcumuladoCompras() As Double
    dValorAcumuladoCompras = mvardValorAcumuladoCompras
End Property

Public Property Let dtDataUltimaCompra(ByVal vData As Date)
    mvardtDataUltimaCompra = vData
End Property

Public Property Get dtDataUltimaCompra() As Date
    dtDataUltimaCompra = mvardtDataUltimaCompra
End Property

Public Property Let dtDataPrimeiraCompra(ByVal vData As Date)
    mvardtDataPrimeiraCompra = vData
End Property

Public Property Get dtDataPrimeiraCompra() As Date
    dtDataPrimeiraCompra = mvardtDataPrimeiraCompra
End Property

Public Property Let dMediaCompra(ByVal vData As Double)
    mvardMediaCompra = vData
End Property

Public Property Get dMediaCompra() As Double
    dMediaCompra = mvardMediaCompra
End Property

Public Property Let lNumeroCompras(ByVal vData As Long)
    mvarlNumeroCompras = vData
End Property

Public Property Get lNumeroCompras() As Long
    lNumeroCompras = mvarlNumeroCompras
End Property

Public Property Let lCodFornecedor(ByVal vData As Long)
    mvarlCodFornecedor = vData
End Property

Public Property Get lCodFornecedor() As Long
    lCodFornecedor = mvarlCodFornecedor
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
    mvariCodFilial = vData
End Property

Public Property Get iCodFilial() As Integer
    iCodFilial = mvariCodFilial
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property


