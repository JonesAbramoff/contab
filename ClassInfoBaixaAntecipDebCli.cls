VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInfoBaixaAntecipDebCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDocumento As Long
Private mvariExcluido As Integer
Private mvarlNumMovto As Long
Private mvardSaldoNaoApropriado As Double
Private mvarlCliente As Long
Private mvariFilial_Cliente As Integer
Private mvariFilial As Integer
Private mvardtDataEmissao As Date
Private mvardValor As Double
Private mvariTipoMeioPagto As Integer
Private mvarlNumero As Long
Private mvarsNomeReduzidoConta As String
Private mvariCodConta As Integer
Private mvariSelecionado As Integer
Private mvariStatus As Integer
Private mvarsSiglaDocumento As String
Private mvardValorSeguro As Double
Private mvardValorFrete As Double
Private mvardValorOutrasDespesas As Double
Private mvardValorProdutos As Double
Private mvardValorICMS As Double
Private mvardValorICMSSubst As Double
Private mvardValorIPI As Double
Private mvardValorIRRF As Double
Private mvarsObservacao As String
Private mvarlNumIntBaixa As Long



Public Property Let lNumIntDocumento(ByVal vData As Long)
    mvarlNumIntDocumento = vData
End Property

Public Property Get lNumIntDocumento() As Long
    lNumIntDocumento = mvarlNumIntDocumento
End Property

Public Property Let iExcluido(ByVal vData As Integer)
    mvariExcluido = vData
End Property

Public Property Get iExcluido() As Integer
    iExcluido = mvariExcluido
End Property

Public Property Let lNumMovto(ByVal vData As Long)
    mvarlNumMovto = vData
End Property

Public Property Get lNumMovto() As Long
    lNumMovto = mvarlNumMovto
End Property

Public Property Let dSaldoNaoApropriado(ByVal vData As Double)
    mvardSaldoNaoApropriado = vData
End Property

Public Property Get dSaldoNaoApropriado() As Double
    dSaldoNaoApropriado = mvardSaldoNaoApropriado
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iFilial_Cliente(ByVal vData As Integer)
    mvariFilial_Cliente = vData
End Property

Public Property Get iFilial_Cliente() As Integer
    iFilial_Cliente = mvariFilial_Cliente
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let iTipoMeioPagto(ByVal vData As Integer)
    mvariTipoMeioPagto = vData
End Property

Public Property Get iTipoMeioPagto() As Integer
    iTipoMeioPagto = mvariTipoMeioPagto
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let sNomeReduzidoConta(ByVal vData As String)
    mvarsNomeReduzidoConta = vData
End Property

Public Property Get sNomeReduzidoConta() As String
    sNomeReduzidoConta = mvarsNomeReduzidoConta
End Property

Public Property Let iCodConta(ByVal vData As Integer)
    mvariCodConta = vData
End Property

Public Property Get iCodConta() As Integer
    iCodConta = mvariCodConta
End Property

Public Property Let iSelecionado(ByVal vData As Integer)
    mvariSelecionado = vData
End Property

Public Property Get iSelecionado() As Integer
    iSelecionado = mvariSelecionado
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let sSiglaDocumento(ByVal vData As String)
    mvarsSiglaDocumento = vData
End Property

Public Property Get sSiglaDocumento() As String
    sSiglaDocumento = mvarsSiglaDocumento
End Property

Public Property Let dValorSeguro(ByVal vData As Double)
    mvardValorSeguro = vData
End Property

Public Property Get dValorSeguro() As Double
    dValorSeguro = mvardValorSeguro
End Property

Public Property Let dValorFrete(ByVal vData As Double)
    mvardValorFrete = vData
End Property

Public Property Get dValorFrete() As Double
    dValorFrete = mvardValorFrete
End Property

Public Property Let dValorOutrasDespesas(ByVal vData As Double)
    mvardValorOutrasDespesas = vData
End Property

Public Property Get dValorOutrasDespesas() As Double
    dValorOutrasDespesas = mvardValorOutrasDespesas
End Property

Public Property Let dValorProdutos(ByVal vData As Double)
    mvardValorProdutos = vData
End Property

Public Property Get dValorProdutos() As Double
    dValorProdutos = mvardValorProdutos
End Property

Public Property Let dValorICMS(ByVal vData As Double)
    mvardValorICMS = vData
End Property

Public Property Get dValorICMS() As Double
    dValorICMS = mvardValorICMS
End Property

Public Property Let dValorICMSSubst(ByVal vData As Double)
    mvardValorICMSSubst = vData
End Property

Public Property Get dValorICMSSubst() As Double
    dValorICMSSubst = mvardValorICMSSubst
End Property

Public Property Let dValorIPI(ByVal vData As Double)
    mvardValorIPI = vData
End Property

Public Property Get dValorIPI() As Double
    dValorIPI = mvardValorIPI
End Property

Public Property Let dValorIRRF(ByVal vData As Double)
    mvardValorIRRF = vData
End Property

Public Property Get dValorIRRF() As Double
    dValorIRRF = mvardValorIRRF
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let lNumIntBaixa(ByVal vData As Long)
    mvarlNumIntBaixa = vData
End Property

Public Property Get lNumIntBaixa() As Long
    lNumIntBaixa = mvarlNumIntBaixa
End Property

