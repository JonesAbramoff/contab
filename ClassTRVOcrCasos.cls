VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVOcrCasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarsCodigo As String
Private mvarsTipVou As String
Private mvarsSerie As String
Private mvarlNumVou As Long
Private mvarsPaxNome As String
Private mvarsTitularNome As String
Private mvarlClienteVou As Long
Private mvardtDataEmissao As Date
Private mvardtDataIda As Date
Private mvardtDataVolta As Date
Private mvarsProduto As String
Private mvariQtdPax As Integer
Private mvarlEndereco As Long
Private mvardtDataAbertura As Date
Private mvardtDataDocsRec As Date
Private mvardtDataEnvioAnalise As Date
Private mvarlCGAnalise As Long
Private mvarlCGStatus As Long
Private mvarlCGAutorizadoPor As Long
Private mvardValorAutorizadoTotalRS As Double
Private mvardValorAutorizadoTotalUS As Double
Private mvardtDataLimite As Date
Private mvardtDataEnvioFinac As Date
Private mvardtDataProgFinanc As Date
Private mvardtDataPagtoPax As Date
Private mvariJudicial As Integer
Private mvarsNumProcesso As String
Private mvariCondenado As Integer
Private mvardValorCondenacao As Double
Private mvarsComarca As String
Private mvardtDataFimProcesso As Date
Private mvardtDataPagtoCond As Date
Private mvardValorAutorizadoSeguroRS As Double
Private mvardValorAutorizadoSeguroUS As Double
Private mvardValorAutorizadoAssistRS As Double
Private mvardValorAutorizadoAssistUS As Double
Private mvardCambio As Double
Private mvariAnteciparPagtoSeguro As Integer
Private mvariBanco As Integer
Private mvarsAgencia As String
Private mvarsContaCorrente As String
Private mvarsNomeFavorecido As String
Private mvarlCodFornFavorecido As Long
Private mvarsFavorecidoCGC As String
Private mvardValorInvoicesTotal As Double
Private mvardValorInvoicesTotalUS As Double
Private mvardValorDespesasTotalRS As Double
Private mvardValorDespesasTotalUS As Double
Private mvarlNumIntDocTitPagCobertura As Long
Private mvarlNumIntDocTitPagProcesso As Long
Private mvarlNumFatCobertura As Long
Private mvarlNumFatProcesso As Long
Private mvarsObs As String
Private mvarcolCoberturas As New Collection
Private mvarcolInvoices As New Collection
Private mvarcolOutrasFaturas As New Collection
Private mvarcolHistorico As New Collection
Private mvarcolAnotacoes As New Collection
Private mvarcolParcProcesso As New Collection
Private mvardtDataPriEvento As Date
Private mvarobjEndereco As New ClassEndereco
Private mvarcolEventos As New Collection
Private mvarobjPreReceber As New ClassTRVOcrCasosPreRec
Private mvardValorGastosAdvRS As Double
Private mvarcolDocs As New Collection
Private mvarcolGastosAdvs As New Collection
Private mvarsCcl As String
Private mvardtDataIniProcesso As Date
Private mvardProcessoDanoMaterial As Double
Private mvardProcessoDanoMoral As Double
Private mvariProcon As Integer
Private mvariPerdaTipo As Integer
Private mvarsNumVouTexto As String
Private mvardValorAutoSegRespTrvRS As Double

Public Property Let dValorAutoSegRespTrvRS(ByVal vData As Double)
    mvardValorAutoSegRespTrvRS = vData
End Property

Public Property Get dValorAutoSegRespTrvRS() As Double
    dValorAutoSegRespTrvRS = mvardValorAutoSegRespTrvRS
End Property

Public Property Let sNumVouTexto(ByVal vData As String)
    mvarsNumVouTexto = vData
End Property

Public Property Get sNumVouTexto() As String
    sNumVouTexto = mvarsNumVouTexto
End Property

Public Property Let iPerdaTipo(ByVal vData As Integer)
    mvariPerdaTipo = vData
End Property

Public Property Get iPerdaTipo() As Integer
    iPerdaTipo = mvariPerdaTipo
End Property

Public Property Let iProcon(ByVal vData As Integer)
    mvariProcon = vData
End Property

Public Property Get iProcon() As Integer
    iProcon = mvariProcon
End Property

Public Property Let dProcessoDanoMoral(ByVal vData As Double)
    mvardProcessoDanoMoral = vData
End Property

Public Property Get dProcessoDanoMoral() As Double
    dProcessoDanoMoral = mvardProcessoDanoMoral
End Property

Public Property Let dProcessoDanoMaterial(ByVal vData As Double)
    mvardProcessoDanoMaterial = vData
End Property

Public Property Get dProcessoDanoMaterial() As Double
    dProcessoDanoMaterial = mvardProcessoDanoMaterial
End Property

Public Property Let dtDataIniProcesso(ByVal vData As Date)
    mvardtDataIniProcesso = vData
End Property

Public Property Get dtDataIniProcesso() As Date
    dtDataIniProcesso = mvardtDataIniProcesso
End Property

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Public Property Set colGastosAdvs(ByVal vData As Collection)
    Set mvarcolGastosAdvs = vData
End Property

Public Property Get colGastosAdvs() As Collection
    Set colGastosAdvs = mvarcolGastosAdvs
End Property

Public Property Set colDocs(ByVal vData As Collection)
    Set mvarcolDocs = vData
End Property

Public Property Get colDocs() As Collection
    Set colDocs = mvarcolDocs
End Property

Public Property Let dValorGastosAdvRS(ByVal vData As Double)
    mvardValorGastosAdvRS = vData
End Property

Public Property Get dValorGastosAdvRS() As Double
    dValorGastosAdvRS = mvardValorGastosAdvRS
End Property

Public Property Set objPreReceber(ByVal vData As ClassTRVOcrCasosPreRec)
    Set mvarobjPreReceber = vData
End Property

Public Property Get objPreReceber() As ClassTRVOcrCasosPreRec
    Set objPreReceber = mvarobjPreReceber
End Property

Public Property Set colEventos(ByVal vData As Collection)
    Set mvarcolEventos = vData
End Property

Public Property Get colEventos() As Collection
    Set colEventos = mvarcolEventos
End Property

Public Property Set objEndereco(ByVal vData As ClassEndereco)
    Set mvarobjEndereco = vData
End Property

Public Property Get objEndereco() As ClassEndereco
    Set objEndereco = mvarobjEndereco
End Property

Public Property Let dtDataPriEvento(ByVal vData As Date)
    mvardtDataPriEvento = vData
End Property

Public Property Get dtDataPriEvento() As Date
    dtDataPriEvento = mvardtDataPriEvento
End Property

Public Property Set colParcProcesso(ByVal vData As Collection)
    Set mvarcolParcProcesso = vData
End Property

Public Property Get colParcProcesso() As Collection
    Set colParcProcesso = mvarcolParcProcesso
End Property

Public Property Set colAnotacoes(ByVal vData As Collection)
    Set mvarcolAnotacoes = vData
End Property

Public Property Get colAnotacoes() As Collection
    Set colAnotacoes = mvarcolAnotacoes
End Property

Public Property Set colHistorico(ByVal vData As Collection)
    Set mvarcolHistorico = vData
End Property

Public Property Get colHistorico() As Collection
    Set colHistorico = mvarcolHistorico
End Property

Public Property Set colOutrasFaturas(ByVal vData As Collection)
    Set mvarcolOutrasFaturas = vData
End Property

Public Property Get colOutrasFaturas() As Collection
    Set colOutrasFaturas = mvarcolOutrasFaturas
End Property

Public Property Set colInvoices(ByVal vData As Collection)
    Set mvarcolInvoices = vData
End Property

Public Property Get colInvoices() As Collection
    Set colInvoices = mvarcolInvoices
End Property

Public Property Set colCoberturas(ByVal vData As Collection)
    Set mvarcolCoberturas = vData
End Property

Public Property Get colCoberturas() As Collection
    Set colCoberturas = mvarcolCoberturas
End Property

Public Property Let sObs(ByVal vData As String)
    mvarsObs = vData
End Property

Public Property Get sObs() As String
    sObs = mvarsObs
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property

Public Property Let sTipVou(ByVal vData As String)
    mvarsTipVou = vData
End Property

Public Property Get sTipVou() As String
    sTipVou = mvarsTipVou
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let lNumVou(ByVal vData As Long)
    mvarlNumVou = vData
End Property

Public Property Get lNumVou() As Long
    lNumVou = mvarlNumVou
End Property

Public Property Let spaxnome(ByVal vData As String)
    mvarsPaxNome = vData
End Property

Public Property Get spaxnome() As String
    spaxnome = mvarsPaxNome
End Property

Public Property Let sTitularNome(ByVal vData As String)
    mvarsTitularNome = vData
End Property

Public Property Get sTitularNome() As String
    sTitularNome = mvarsTitularNome
End Property

Public Property Let lClienteVou(ByVal vData As Long)
    mvarlClienteVou = vData
End Property

Public Property Get lClienteVou() As Long
    lClienteVou = mvarlClienteVou
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dtDataIda(ByVal vData As Date)
    mvardtDataIda = vData
End Property

Public Property Get dtDataIda() As Date
    dtDataIda = mvardtDataIda
End Property

Public Property Let dtDataVolta(ByVal vData As Date)
    mvardtDataVolta = vData
End Property

Public Property Get dtDataVolta() As Date
    dtDataVolta = mvardtDataVolta
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iQtdPax(ByVal vData As Integer)
    mvariQtdPax = vData
End Property

Public Property Get iQtdPax() As Integer
    iQtdPax = mvariQtdPax
End Property

Public Property Let lEndereco(ByVal vData As Long)
    mvarlEndereco = vData
End Property

Public Property Get lEndereco() As Long
    lEndereco = mvarlEndereco
End Property

Public Property Let dtDataAbertura(ByVal vData As Date)
    mvardtDataAbertura = vData
End Property

Public Property Get dtDataAbertura() As Date
    dtDataAbertura = mvardtDataAbertura
End Property

Public Property Let dtDataDocsRec(ByVal vData As Date)
    mvardtDataDocsRec = vData
End Property

Public Property Get dtDataDocsRec() As Date
    dtDataDocsRec = mvardtDataDocsRec
End Property

Public Property Let dtDataEnvioAnalise(ByVal vData As Date)
    mvardtDataEnvioAnalise = vData
End Property

Public Property Get dtDataEnvioAnalise() As Date
    dtDataEnvioAnalise = mvardtDataEnvioAnalise
End Property

Public Property Let lCGAnalise(ByVal vData As Long)
    mvarlCGAnalise = vData
End Property

Public Property Get lCGAnalise() As Long
    lCGAnalise = mvarlCGAnalise
End Property

Public Property Let lCGStatus(ByVal vData As Long)
    mvarlCGStatus = vData
End Property

Public Property Get lCGStatus() As Long
    lCGStatus = mvarlCGStatus
End Property

Public Property Let lCGAutorizadoPor(ByVal vData As Long)
    mvarlCGAutorizadoPor = vData
End Property

Public Property Get lCGAutorizadoPor() As Long
    lCGAutorizadoPor = mvarlCGAutorizadoPor
End Property

Public Property Let dValorAutorizadoTotalRS(ByVal vData As Double)
    mvardValorAutorizadoTotalRS = vData
End Property

Public Property Get dValorAutorizadoTotalRS() As Double
    dValorAutorizadoTotalRS = mvardValorAutorizadoTotalRS
End Property

Public Property Let dValorAutorizadoTotalUS(ByVal vData As Double)
    mvardValorAutorizadoTotalUS = vData
End Property

Public Property Get dValorAutorizadoTotalUS() As Double
    dValorAutorizadoTotalUS = mvardValorAutorizadoTotalUS
End Property

Public Property Let dtDataLimite(ByVal vData As Date)
    mvardtDataLimite = vData
End Property

Public Property Get dtDataLimite() As Date
    dtDataLimite = mvardtDataLimite
End Property

Public Property Let dtDataEnvioFinac(ByVal vData As Date)
    mvardtDataEnvioFinac = vData
End Property

Public Property Get dtDataEnvioFinac() As Date
    dtDataEnvioFinac = mvardtDataEnvioFinac
End Property

Public Property Let dtDataProgFinanc(ByVal vData As Date)
    mvardtDataProgFinanc = vData
End Property

Public Property Get dtDataProgFinanc() As Date
    dtDataProgFinanc = mvardtDataProgFinanc
End Property

Public Property Let dtDataPagtoPax(ByVal vData As Date)
    mvardtDataPagtoPax = vData
End Property

Public Property Get dtDataPagtoPax() As Date
    dtDataPagtoPax = mvardtDataPagtoPax
End Property

Public Property Let iJudicial(ByVal vData As Integer)
    mvariJudicial = vData
End Property

Public Property Get iJudicial() As Integer
    iJudicial = mvariJudicial
End Property

Public Property Let sNumProcesso(ByVal vData As String)
    mvarsNumProcesso = vData
End Property

Public Property Get sNumProcesso() As String
    sNumProcesso = mvarsNumProcesso
End Property

Public Property Let iCondenado(ByVal vData As Integer)
    mvariCondenado = vData
End Property

Public Property Get iCondenado() As Integer
    iCondenado = mvariCondenado
End Property

Public Property Let dValorCondenacao(ByVal vData As Double)
    mvardValorCondenacao = vData
End Property

Public Property Get dValorCondenacao() As Double
    dValorCondenacao = mvardValorCondenacao
End Property

Public Property Let sComarca(ByVal vData As String)
    mvarsComarca = vData
End Property

Public Property Get sComarca() As String
    sComarca = mvarsComarca
End Property

Public Property Let dtDataFimProcesso(ByVal vData As Date)
    mvardtDataFimProcesso = vData
End Property

Public Property Get dtDataFimProcesso() As Date
    dtDataFimProcesso = mvardtDataFimProcesso
End Property

Public Property Let dtDataPagtoCond(ByVal vData As Date)
    mvardtDataPagtoCond = vData
End Property

Public Property Get dtDataPagtoCond() As Date
    dtDataPagtoCond = mvardtDataPagtoCond
End Property

Public Property Let dValorAutorizadoSeguroRS(ByVal vData As Double)
    mvardValorAutorizadoSeguroRS = vData
End Property

Public Property Get dValorAutorizadoSeguroRS() As Double
    dValorAutorizadoSeguroRS = mvardValorAutorizadoSeguroRS
End Property

Public Property Let dValorAutorizadoSeguroUS(ByVal vData As Double)
    mvardValorAutorizadoSeguroUS = vData
End Property

Public Property Get dValorAutorizadoSeguroUS() As Double
    dValorAutorizadoSeguroUS = mvardValorAutorizadoSeguroUS
End Property

Public Property Let dValorAutorizadoAssistRS(ByVal vData As Double)
    mvardValorAutorizadoAssistRS = vData
End Property

Public Property Get dValorAutorizadoAssistRS() As Double
    dValorAutorizadoAssistRS = mvardValorAutorizadoAssistRS
End Property

Public Property Let dValorAutorizadoAssistUS(ByVal vData As Double)
    mvardValorAutorizadoAssistUS = vData
End Property

Public Property Get dValorAutorizadoAssistUS() As Double
    dValorAutorizadoAssistUS = mvardValorAutorizadoAssistUS
End Property

Public Property Let dCambio(ByVal vData As Double)
    mvardCambio = vData
End Property

Public Property Get dCambio() As Double
    dCambio = mvardCambio
End Property

Public Property Let iAnteciparPagtoSeguro(ByVal vData As Integer)
    mvariAnteciparPagtoSeguro = vData
End Property

Public Property Get iAnteciparPagtoSeguro() As Integer
    iAnteciparPagtoSeguro = mvariAnteciparPagtoSeguro
End Property

Public Property Let iBanco(ByVal vData As Integer)
    mvariBanco = vData
End Property

Public Property Get iBanco() As Integer
    iBanco = mvariBanco
End Property

Public Property Let sAgencia(ByVal vData As String)
    mvarsAgencia = vData
End Property

Public Property Get sAgencia() As String
    sAgencia = mvarsAgencia
End Property

Public Property Let sContaCorrente(ByVal vData As String)
    mvarsContaCorrente = vData
End Property

Public Property Get sContaCorrente() As String
    sContaCorrente = mvarsContaCorrente
End Property

Public Property Let sNomeFavorecido(ByVal vData As String)
    mvarsNomeFavorecido = vData
End Property

Public Property Get sNomeFavorecido() As String
    sNomeFavorecido = mvarsNomeFavorecido
End Property

Public Property Let lCodFornFavorecido(ByVal vData As Long)
    mvarlCodFornFavorecido = vData
End Property

Public Property Get lCodFornFavorecido() As Long
    lCodFornFavorecido = mvarlCodFornFavorecido
End Property

Public Property Let sFavorecidoCGC(ByVal vData As String)
    mvarsFavorecidoCGC = vData
End Property

Public Property Get sFavorecidoCGC() As String
    sFavorecidoCGC = mvarsFavorecidoCGC
End Property

Public Property Let dValorInvoicesTotal(ByVal vData As Double)
    mvardValorInvoicesTotal = vData
End Property

Public Property Get dValorInvoicesTotal() As Double
    dValorInvoicesTotal = mvardValorInvoicesTotal
End Property

Public Property Let dValorInvoicesTotalUS(ByVal vData As Double)
    mvardValorInvoicesTotalUS = vData
End Property

Public Property Get dValorInvoicesTotalUS() As Double
    dValorInvoicesTotalUS = mvardValorInvoicesTotalUS
End Property

Public Property Let dValorDespesasTotalRS(ByVal vData As Double)
    mvardValorDespesasTotalRS = vData
End Property

Public Property Get dValorDespesasTotalRS() As Double
    dValorDespesasTotalRS = mvardValorDespesasTotalRS
End Property

Public Property Let dValorDespesasTotalUS(ByVal vData As Double)
    mvardValorDespesasTotalUS = vData
End Property

Public Property Get dValorDespesasTotalUS() As Double
    dValorDespesasTotalUS = mvardValorDespesasTotalUS
End Property

Public Property Let lNumIntDocTitPagCobertura(ByVal vData As Long)
    mvarlNumIntDocTitPagCobertura = vData
End Property

Public Property Get lNumIntDocTitPagCobertura() As Long
    lNumIntDocTitPagCobertura = mvarlNumIntDocTitPagCobertura
End Property

Public Property Let lNumIntDocTitPagProcesso(ByVal vData As Long)
    mvarlNumIntDocTitPagProcesso = vData
End Property

Public Property Get lNumIntDocTitPagProcesso() As Long
    lNumIntDocTitPagProcesso = mvarlNumIntDocTitPagProcesso
End Property

Public Property Let lNumFatCobertura(ByVal vData As Long)
    mvarlNumFatCobertura = vData
End Property

Public Property Get lNumFatCobertura() As Long
    lNumFatCobertura = mvarlNumFatCobertura
End Property

Public Property Let lNumFatProcesso(ByVal vData As Long)
    mvarlNumFatProcesso = vData
End Property

Public Property Get lNumFatProcesso() As Long
    lNumFatProcesso = mvarlNumFatProcesso
End Property

Private Sub Class_Initialize()
    dtDataAbertura = DATA_NULA
    dtDataDocsRec = DATA_NULA
    dtDataEmissao = DATA_NULA
    dtDataEnvioAnalise = DATA_NULA
    dtDataEnvioFinac = DATA_NULA
    dtDataFimProcesso = DATA_NULA
    dtDataIda = DATA_NULA
    dtDataLimite = DATA_NULA
    dtDataPagtoCond = DATA_NULA
    dtDataPagtoPax = DATA_NULA
    dtDataPriEvento = DATA_NULA
    dtDataProgFinanc = DATA_NULA
    dtDataVolta = DATA_NULA
End Sub
