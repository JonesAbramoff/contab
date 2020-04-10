VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOrcamentoLoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlNumIntDoc As Long
Private mvariStatus As Integer
Private mvariECF As Integer
Private mvardtDataEmissao As Date
Private mvardHoraEmissao As Double
Private mvarlNumero As Long
Private mvarsCPFCGC As String
Private mvarlNumOrcamento As Long
Private mvardValorTotal As Double
Private mvardValorProdutos As Double
Private mvardValorDesconto As Double
Private mvardValorAcrescimo As Double
Private mvariTabelaPreco As Integer
Private mvarsNaturezaOp As String
Private mvarlGerenteCancel As Long
Private mvarsMotivoCancel As String
Private mvarcolItens As New Collection
Private mvariVendedor As Integer
Private mvardValorTroco As Double
Private mvarlMovCaixa As Long
Private mvarlDuracao As Long
Private mvarsClienteNomeRed As String
Private mvarsNumeroDAV As String

Public IdExterno As String 'ex.: order id da NTK
Public NumCaixa As Integer 'do caixa que pegou o orçamento para si
Public StatusExterno As Integer 'confirmado, despachado, entregue,...
Public objPedidoExterno As Object

Public Property Let sNumeroDAV(ByVal vData As String)
    mvarsNumeroDAV = vData
End Property

Public Property Get sNumeroDAV() As String
    sNumeroDAV = mvarsNumeroDAV
End Property


Public Property Let sClienteNomeRed(ByVal vData As String)
    mvarsClienteNomeRed = vData
End Property

Public Property Get sClienteNomeRed() As String
    sClienteNomeRed = mvarsClienteNomeRed
End Property

Public Property Let lDuracao(ByVal vData As Long)
    mvarlDuracao = vData
End Property

Public Property Get lDuracao() As Long
    lDuracao = mvarlDuracao
End Property

Public Property Let lMovCaixa(ByVal vData As Long)
    mvarlMovCaixa = vData
End Property

Public Property Get lMovCaixa() As Long
    lMovCaixa = mvarlMovCaixa
End Property

Public Property Let dValorTroco(ByVal vData As Double)
    mvardValorTroco = vData
End Property

Public Property Get dValorTroco() As Double
    dValorTroco = mvardValorTroco
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let sMotivoCancel(ByVal vData As String)
    mvarsMotivoCancel = vData
End Property

Public Property Get sMotivoCancel() As String
    sMotivoCancel = mvarsMotivoCancel
End Property

Public Property Let lGerenteCancel(ByVal vData As Long)
    mvarlGerenteCancel = vData
End Property

Public Property Get lGerenteCancel() As Long
    lGerenteCancel = mvarlGerenteCancel
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    mvarsNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOp
End Property

Public Property Let iTabelaPreco(ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco() As Integer
    iTabelaPreco = mvariTabelaPreco
End Property

Public Property Let dValorAcrescimo(ByVal vData As Double)
    mvardValorAcrescimo = vData
End Property

Public Property Get dValorAcrescimo() As Double
    dValorAcrescimo = mvardValorAcrescimo
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorProdutos(ByVal vData As Double)
    mvardValorProdutos = vData
End Property

Public Property Get dValorProdutos() As Double
    dValorProdutos = mvardValorProdutos
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let lNumOrcamento(ByVal vData As Long)
    mvarlNumOrcamento = vData
End Property

Public Property Get lNumOrcamento() As Long
    lNumOrcamento = mvarlNumOrcamento
End Property

Public Property Let sCPFCGC(ByVal vData As String)
    mvarsCPFCGC = vData
End Property

Public Property Get sCPFCGC() As String
    sCPFCGC = mvarsCPFCGC
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let dHoraEmissao(ByVal vData As Double)
    mvardHoraEmissao = vData
End Property

Public Property Get dHoraEmissao() As Double
    dHoraEmissao = mvardHoraEmissao
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let iECF(ByVal vData As Integer)
    mvariECF = vData
End Property

Public Property Get iECF() As Integer
    iECF = mvariECF
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property




