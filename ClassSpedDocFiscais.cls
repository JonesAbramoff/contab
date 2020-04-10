VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSpedDocFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntSped As Long
Private mvariTipoDoc As Integer
Private mvarlNumIntDoc As Long
Private mvarsBloco As String
Private mvarsRegistro As String
Private mvariFilialEmpresa As Integer
Private mvarlCliente As Long
Private mvariFilialCli As Integer
Private mvarlFornecedor As Long
Private mvariFilialForn As Integer
Private mvardtDataEmissao As Date
Private mvardtDataES As Date
Private mvardValorTotal As Double
Private mvarsOperacao As String
Private mvarsEmitente As String
Private mvarsModelo As String
Private mvariSituacao As Integer
Private mvarsSerie As String
Private mvarlNumDocumento As Long
Private mvarsChaveEletronica As String
Private mvarsFrete As String
Private mvarsPagamento As String
Private mvardValorDesconto As Double
Private mvardValorFrete As Double
Private mvardValorSeguro As Double
Private mvardValorDespesas As Double
Private mvardValorServico As Double
Private mvardValorNaoTributado As Double
Private mvardValorBaseICMS As Double
Private mvardValorICMS As Double
Private mvardValorBaseICMSST As Double
Private mvardValorICMSST As Double
Private mvardValorBaseIPI As Double
Private mvardValorIPI As Double
Private mvardValorBasePIS As Double
Private mvardValorPIS As Double
Private mvardValorBaseCofins As Double
Private mvardValorCofins As Double
Private mvardValorPisRetido As Double
Private mvardValorPisST As Double
Private mvardValorCofinsRetido As Double
Private mvardValorCofinsST As Double
Private mvardValorISS As Double
Private mvarsCGC As String
Private mvarcolItens As New Collection
Private mvarlCodMensagem As Long
Private mvardValorMercadoria As Double
Private mvarsIndNatFrtPis As String
Private mvarsIndNatFrtCofins As String
Private mvardValorRecebido As Double
Private mvardValorNoCR As Double
Private mvardValorRecContrPrev As Double

'EFD ICMS 2019
Public sCodMunicServ As String
Public dValorBaseISS As Double
Public dValorBaseISSRet As Double
Public dValorISSRet As Double

Public Property Let dValorRecContrPrev(ByVal vData As Double)
    mvardValorRecContrPrev = vData
End Property

Public Property Get dValorRecContrPrev() As Double
    dValorRecContrPrev = mvardValorRecContrPrev
End Property

Public Property Let dValorNoCR(ByVal vData As Double)
    mvardValorNoCR = vData
End Property

Public Property Get dValorNoCR() As Double
    dValorNoCR = mvardValorNoCR
End Property

Public Property Let dValorRecebido(ByVal vData As Double)
    mvardValorRecebido = vData
End Property

Public Property Get dValorRecebido() As Double
    dValorRecebido = mvardValorRecebido
End Property

Public Property Let sIndNatFrtCofins(ByVal vData As String)
    mvarsIndNatFrtCofins = vData
End Property

Public Property Get sIndNatFrtCofins() As String
    sIndNatFrtCofins = mvarsIndNatFrtCofins
End Property

Public Property Let sIndNatFrtPis(ByVal vData As String)
    mvarsIndNatFrtPis = vData
End Property

Public Property Get sIndNatFrtPis() As String
    sIndNatFrtPis = mvarsIndNatFrtPis
End Property

Public Property Let dValorMercadoria(ByVal vData As Double)
    mvardValorMercadoria = vData
End Property

Public Property Get dValorMercadoria() As Double
    dValorMercadoria = mvardValorMercadoria
End Property

Public Property Let lCodMensagem(ByVal vData As Long)
    mvarlCodMensagem = vData
End Property

Public Property Get lCodMensagem() As Long
    lCodMensagem = mvarlCodMensagem
End Property

Public Property Let sCgc(ByVal vData As String)
    mvarsCGC = vData
End Property

Public Property Get sCgc() As String
    sCgc = mvarsCGC
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let lNumIntSped(ByVal vData As Long)
    mvarlNumIntSped = vData
End Property

Public Property Get lNumIntSped() As Long
    lNumIntSped = mvarlNumIntSped
End Property

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sBloco(ByVal vData As String)
    mvarsBloco = vData
End Property

Public Property Get sBloco() As String
    sBloco = mvarsBloco
End Property

Public Property Let sRegistro(ByVal vData As String)
    mvarsRegistro = vData
End Property

Public Property Get sRegistro() As String
    sRegistro = mvarsRegistro
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iFilialCli(ByVal vData As Integer)
    mvariFilialCli = vData
End Property

Public Property Get iFilialCli() As Integer
    iFilialCli = mvariFilialCli
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let iFilialForn(ByVal vData As Integer)
    mvariFilialForn = vData
End Property

Public Property Get iFilialForn() As Integer
    iFilialForn = mvariFilialForn
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dtDataES(ByVal vData As Date)
    mvardtDataES = vData
End Property

Public Property Get dtDataES() As Date
    dtDataES = mvardtDataES
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let sOperacao(ByVal vData As String)
    mvarsOperacao = vData
End Property

Public Property Get sOperacao() As String
    sOperacao = mvarsOperacao
End Property

Public Property Let sEmitente(ByVal vData As String)
    mvarsEmitente = vData
End Property

Public Property Get sEmitente() As String
    sEmitente = mvarsEmitente
End Property

Public Property Let sModelo(ByVal vData As String)
    mvarsModelo = vData
End Property

Public Property Get sModelo() As String
    sModelo = mvarsModelo
End Property

Public Property Let iSituacao(ByVal vData As Integer)
    mvariSituacao = vData
End Property

Public Property Get iSituacao() As Integer
    iSituacao = mvariSituacao
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let lNumDocumento(ByVal vData As Long)
    mvarlNumDocumento = vData
End Property

Public Property Get lNumDocumento() As Long
    lNumDocumento = mvarlNumDocumento
End Property

Public Property Let sChaveEletronica(ByVal vData As String)
    mvarsChaveEletronica = vData
End Property

Public Property Get sChaveEletronica() As String
    sChaveEletronica = mvarsChaveEletronica
End Property

Public Property Let sFrete(ByVal vData As String)
    mvarsFrete = vData
End Property

Public Property Get sFrete() As String
    sFrete = mvarsFrete
End Property

Public Property Let sPagamento(ByVal vData As String)
    mvarsPagamento = vData
End Property

Public Property Get sPagamento() As String
    sPagamento = mvarsPagamento
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorFrete(ByVal vData As Double)
    mvardValorFrete = vData
End Property

Public Property Get dValorFrete() As Double
    dValorFrete = mvardValorFrete
End Property

Public Property Let dValorSeguro(ByVal vData As Double)
    mvardValorSeguro = vData
End Property

Public Property Get dValorSeguro() As Double
    dValorSeguro = mvardValorSeguro
End Property

Public Property Let dValorDespesas(ByVal vData As Double)
    mvardValorDespesas = vData
End Property

Public Property Get dValorDespesas() As Double
    dValorDespesas = mvardValorDespesas
End Property

Public Property Let dValorServico(ByVal vData As Double)
    mvardValorServico = vData
End Property

Public Property Get dValorServico() As Double
    dValorServico = mvardValorServico
End Property

Public Property Let dValorNaoTributado(ByVal vData As Double)
    mvardValorNaoTributado = vData
End Property

Public Property Get dValorNaoTributado() As Double
    dValorNaoTributado = mvardValorNaoTributado
End Property

Public Property Let dValorBaseICMS(ByVal vData As Double)
    mvardValorBaseICMS = vData
End Property

Public Property Get dValorBaseICMS() As Double
    dValorBaseICMS = mvardValorBaseICMS
End Property

Public Property Let dValorICMS(ByVal vData As Double)
    mvardValorICMS = vData
End Property

Public Property Get dValorICMS() As Double
    dValorICMS = mvardValorICMS
End Property

Public Property Let dValorBaseICMSST(ByVal vData As Double)
    mvardValorBaseICMSST = vData
End Property

Public Property Get dValorBaseICMSST() As Double
    dValorBaseICMSST = mvardValorBaseICMSST
End Property

Public Property Let dValorICMSST(ByVal vData As Double)
    mvardValorICMSST = vData
End Property

Public Property Get dValorICMSST() As Double
    dValorICMSST = mvardValorICMSST
End Property

Public Property Let dValorBaseIPI(ByVal vData As Double)
    mvardValorBaseIPI = vData
End Property

Public Property Get dValorBaseIPI() As Double
    dValorBaseIPI = mvardValorBaseIPI
End Property

Public Property Let dValorIPI(ByVal vData As Double)
    mvardValorIPI = vData
End Property

Public Property Get dValorIPI() As Double
    dValorIPI = mvardValorIPI
End Property

Public Property Let dValorBasePIS(ByVal vData As Double)
    mvardValorBasePIS = vData
End Property

Public Property Get dValorBasePIS() As Double
    dValorBasePIS = mvardValorBasePIS
End Property

Public Property Let dValorPIS(ByVal vData As Double)
    mvardValorPIS = vData
End Property

Public Property Get dValorPIS() As Double
    dValorPIS = mvardValorPIS
End Property

Public Property Let dValorBaseCofins(ByVal vData As Double)
    mvardValorBaseCofins = vData
End Property

Public Property Get dValorBaseCofins() As Double
    dValorBaseCofins = mvardValorBaseCofins
End Property

Public Property Let dValorCofins(ByVal vData As Double)
    mvardValorCofins = vData
End Property

Public Property Get dValorCofins() As Double
    dValorCofins = mvardValorCofins
End Property

Public Property Let dValorPisRetido(ByVal vData As Double)
    mvardValorPisRetido = vData
End Property

Public Property Get dValorPisRetido() As Double
    dValorPisRetido = mvardValorPisRetido
End Property

Public Property Let dValorPisST(ByVal vData As Double)
    mvardValorPisST = vData
End Property

Public Property Get dValorPisST() As Double
    dValorPisST = mvardValorPisST
End Property

Public Property Let dValorCofinsRetido(ByVal vData As Double)
    mvardValorCofinsRetido = vData
End Property

Public Property Get dValorCofinsRetido() As Double
    dValorCofinsRetido = mvardValorCofinsRetido
End Property

Public Property Let dValorCofinsST(ByVal vData As Double)
    mvardValorCofinsST = vData
End Property

Public Property Get dValorCofinsST() As Double
    dValorCofinsST = mvardValorCofinsST
End Property

Public Property Let dValorISS(ByVal vData As Double)
    mvardValorISS = vData
End Property

Public Property Get dValorISS() As Double
    dValorISS = mvardValorISS
End Property
