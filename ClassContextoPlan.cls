VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContextoPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public colMnemonicoValor As New ClassColMnemonicoValor

Private mvariFilialFaturamento As Integer
Private mvariAno As Integer
Private mvarsProduto As String
Private mvarlCliente As Long
Private mvariFilialCli As Integer
Private mvardQuantidade As Double
Private mvariVendedor As Integer
Private mvardtDataCustoFixo As Date
Private mvariTabelaPreco As Integer
Private mvardTaxaDescPadrao As Double
Private mvardTaxaValFut As Double
Private mvardDiasValFut As Double
Private mvariRotinaOrigem As Integer

Private mvarsNomeRedCliente As String
Private mvarsUM As String
Private mvardPrecoUnitario As Double

Private mvarobjTelaComissoes As Object

Private mvarcolComissoes As Collection
Private mvardPrecoPraticado As Double
Private mvardQtdePrevTotal As Double

Public Property Let dQtdePrevTotal(ByVal vData As Double)
    mvardQtdePrevTotal = vData
End Property

Public Property Get dQtdePrevTotal() As Double
    dQtdePrevTotal = mvardQtdePrevTotal
End Property

Public Property Let dPrecoPraticado(ByVal vData As Double)
    mvardPrecoPraticado = vData
End Property

Public Property Get dPrecoPraticado() As Double
    dPrecoPraticado = mvardPrecoPraticado
End Property

Public Property Set colComissoes(ByVal vData As Object)
    Set mvarcolComissoes = vData
End Property

Public Property Get colComissoes() As Collection
    Set colComissoes = mvarcolComissoes
End Property

Public Property Let dPrecoUnitario(ByVal vData As Double)
    mvardPrecoUnitario = vData
End Property

Public Property Get dPrecoUnitario() As Double
    dPrecoUnitario = mvardPrecoUnitario
End Property

Public Property Let sNomeRedCliente(ByVal vData As String)
    mvarsNomeRedCliente = vData
End Property

Public Property Get sNomeRedCliente() As String
    sNomeRedCliente = mvarsNomeRedCliente
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Set objTelaComissoes(ByVal vData As Object)
    Set mvarobjTelaComissoes = vData
End Property

Public Property Get objTelaComissoes() As Object
    Set objTelaComissoes = mvarobjTelaComissoes
End Property

Public Property Let iRotinaOrigem(ByVal vData As Integer)
    mvariRotinaOrigem = vData
End Property

Public Property Get iRotinaOrigem() As Integer
    iRotinaOrigem = mvariRotinaOrigem
End Property

Public Property Let dTaxaDescPadrao(ByVal vData As Double)
    mvardTaxaDescPadrao = vData
End Property

Public Property Get dTaxaDescPadrao() As Double
    dTaxaDescPadrao = mvardTaxaDescPadrao
End Property

Public Property Let dTaxaValFut(ByVal vData As Double)
    mvardTaxaValFut = vData
End Property

Public Property Get dTaxaValFut() As Double
    dTaxaValFut = mvardTaxaValFut
End Property

Public Property Let dDiasValFut(ByVal vData As Double)
    mvardDiasValFut = vData
End Property

Public Property Get dDiasValFut() As Double
    dDiasValFut = mvardDiasValFut
End Property

Public Property Let iTabelaPreco(ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco() As Integer
    iTabelaPreco = mvariTabelaPreco
End Property

Public Property Let dtDataCustoFixo(ByVal vData As Date)
    mvardtDataCustoFixo = vData
End Property

Public Property Get dtDataCustoFixo() As Date
    dtDataCustoFixo = mvardtDataCustoFixo
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let iFilialCli(ByVal vData As Integer)
    mvariFilialCli = vData
End Property

Public Property Get iFilialCli() As Integer
    iFilialCli = mvariFilialCli
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

Public Property Let iFilialFaturamento(ByVal vData As Integer)
    mvariFilialFaturamento = vData
End Property

Public Property Get iFilialFaturamento() As Integer
    iFilialFaturamento = mvariFilialFaturamento
End Property

Private Sub Class_Initialize()
    mvardQtdePrevTotal = -1
End Sub
