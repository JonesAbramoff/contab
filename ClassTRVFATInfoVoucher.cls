VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVFATInfoVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumVou As Long
Private mvarsTipoDoc As String
Private mvarsTipoVou As String
Private mvarsSerie As String
Private mvarlCliente As Long
Private mvarlFatura As Long
Private mvardtDataEmissao As Date
Private mvardValorBruto As Double
Private mvardValorComissao As Double
Private mvardValorAporte As Double
Private mvarlNumIntDocAporte As Long
Private mvariFilialEmpresa As Integer
Private mvariCondPagto As Integer
Private mvardValor As Double
Private mvarlNumIntDoc As Long
Private mvarlNumero As Long
Private mvarsCodGrupo As String
Private mvariMarcado As Integer
Private mvarlNumIntDocAportePagto As Long
Private mvarlClienteAporte As Long
Private mvariAporte As Integer
Private mvarsNome As String
Private mvarlLinha As Long
Private mvarsProduto As String
Private mvarsControle As String
Private mvarsMoeda As String
Private mvariPax As Integer
Private mvarsTitular As String
Private mvardValorCambio As Double
Private mvardCambio As Double
Private mvarlCodigoAporte As Long
Private mvariTipoDoc As Integer
Private mvarsEmissor As String
Private mvariParcelaAporte As Integer
Private mvarsNomeCliVou As String
Private mvarlCliVou As Long
Private mvarcolAportes As New Collection
Private mvariCartao As Integer
Private mvarsPassageiroSobreNome As String
Private mvarsPassageiroNome As String
Private mvarlCliPassageiro As Long
Private mvariFilialFat As Integer
Private mvarlEmissor As Long
Private mvardValorAporteVou As Double
Private mvarsTitularVou As String
Private mvariFaturaJuntoOcr As Integer
Private mvariUmVouPorFat As Integer
Private mvariTipoCliente As Integer
Private mvariTipoGeracao As Integer
Private mvariTipoOCR As Integer
Private mvardValorVou As Double
Private mvarsEmissorCPF As String
Private mvarsEmissorCartao As String
Private mvarsEmissorNomeCompl As String
Private mvariPF As Integer
Private mvardValorComissaoAux As Double

Public Property Let dValorComissaoAux(ByVal vData As Double)
    mvardValorComissaoAux = vData
End Property

Public Property Get dValorComissaoAux() As Double
    dValorComissaoAux = mvardValorComissaoAux
End Property

Public Property Let iPF(ByVal vData As Integer)
    mvariPF = vData
End Property

Public Property Get iPF() As Integer
    iPF = mvariPF
End Property

Public Property Let sEmissorNomeCompl(ByVal vData As String)
    mvarsEmissorNomeCompl = vData
End Property

Public Property Get sEmissorNomeCompl() As String
    sEmissorNomeCompl = mvarsEmissorNomeCompl
End Property

Public Property Let sEmissorCartao(ByVal vData As String)
    mvarsEmissorCartao = vData
End Property

Public Property Get sEmissorCartao() As String
    sEmissorCartao = mvarsEmissorCartao
End Property

Public Property Let sEmissorCPF(ByVal vData As String)
    mvarsEmissorCPF = vData
End Property

Public Property Get sEmissorCPF() As String
    sEmissorCPF = mvarsEmissorCPF
End Property

Public Property Let dValorVou(ByVal vData As Double)
    mvardValorVou = vData
End Property

Public Property Get dValorVou() As Double
    dValorVou = mvardValorVou
End Property

Public Property Let iTipoOCR(ByVal vData As Integer)
    mvariTipoOCR = vData
End Property

Public Property Get iTipoOCR() As Integer
    iTipoOCR = mvariTipoOCR
End Property

Public Property Let iTipoGeracao(ByVal vData As Integer)
    mvariTipoGeracao = vData
End Property

Public Property Get iTipoGeracao() As Integer
    iTipoGeracao = mvariTipoGeracao
End Property

Public Property Let iTipoCliente(ByVal vData As Integer)
    mvariTipoCliente = vData
End Property

Public Property Get iTipoCliente() As Integer
    iTipoCliente = mvariTipoCliente
End Property

Public Property Let iUmVouPorFat(ByVal vData As Integer)
    mvariUmVouPorFat = vData
End Property

Public Property Get iUmVouPorFat() As Integer
    iUmVouPorFat = mvariUmVouPorFat
End Property

Public Property Let iFaturaJuntoOcr(ByVal vData As Integer)
    mvariFaturaJuntoOcr = vData
End Property

Public Property Get iFaturaJuntoOcr() As Integer
    iFaturaJuntoOcr = mvariFaturaJuntoOcr
End Property

Public Property Let sTitularVou(ByVal vData As String)
    mvarsTitularVou = vData
End Property

Public Property Get sTitularVou() As String
    sTitularVou = mvarsTitularVou
End Property

Public Property Let dValorAporteVou(ByVal vData As Double)
    mvardValorAporteVou = vData
End Property

Public Property Get dValorAporteVou() As Double
    dValorAporteVou = mvardValorAporteVou
End Property

Public Property Let lEmissor(ByVal vData As Long)
    mvarlEmissor = vData
End Property

Public Property Get lEmissor() As Long
    lEmissor = mvarlEmissor
End Property

Public Property Let iFilialFat(ByVal vData As Integer)
    mvariFilialFat = vData
End Property

Public Property Get iFilialFat() As Integer
    iFilialFat = mvariFilialFat
End Property

Public Property Let iCartao(ByVal vData As Integer)
    mvariCartao = vData
End Property

Public Property Get iCartao() As Integer
    iCartao = mvariCartao
End Property

Public Property Let sPassageiroSobreNome(ByVal vData As String)
    mvarsPassageiroSobreNome = vData
End Property

Public Property Get sPassageiroSobreNome() As String
    sPassageiroSobreNome = mvarsPassageiroSobreNome
End Property

Public Property Let sPassageiroNome(ByVal vData As String)
    mvarsPassageiroNome = vData
End Property

Public Property Get sPassageiroNome() As String
    sPassageiroNome = mvarsPassageiroNome
End Property

Public Property Let lCliPassageiro(ByVal vData As Long)
    mvarlCliPassageiro = vData
End Property

Public Property Get lCliPassageiro() As Long
    lCliPassageiro = mvarlCliPassageiro
End Property
Public Property Set colAportes(ByVal vData As Collection)
    Set mvarcolAportes = vData
End Property

Public Property Get colAportes() As Collection
    Set colAportes = mvarcolAportes
End Property

Public Property Let lCliVou(ByVal vData As Long)
    mvarlCliVou = vData
End Property

Public Property Get lCliVou() As Long
    lCliVou = mvarlCliVou
End Property

Public Property Let sNomeCliVou(ByVal vData As String)
    mvarsNomeCliVou = vData
End Property

Public Property Get sNomeCliVou() As String
    sNomeCliVou = mvarsNomeCliVou
End Property

Public Property Let iParcelaAporte(ByVal vData As Integer)
    mvariParcelaAporte = vData
End Property

Public Property Get iParcelaAporte() As Integer
    iParcelaAporte = mvariParcelaAporte
End Property

Public Property Let sEmissor(ByVal vData As String)
    mvarsEmissor = vData
End Property

Public Property Get sEmissor() As String
    sEmissor = mvarsEmissor
End Property

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Let lCodigoAporte(ByVal vData As Long)
    mvarlCodigoAporte = vData
End Property

Public Property Get lCodigoAporte() As Long
    lCodigoAporte = mvarlCodigoAporte
End Property

Public Property Let dCambio(ByVal vData As Double)
    mvardCambio = vData
End Property

Public Property Get dCambio() As Double
    dCambio = mvardCambio
End Property

Public Property Let dValorCambio(ByVal vData As Double)
    mvardValorCambio = vData
End Property

Public Property Get dValorCambio() As Double
    dValorCambio = mvardValorCambio
End Property

Public Property Let sTitular(ByVal vData As String)
    mvarsTitular = vData
End Property

Public Property Get sTitular() As String
    sTitular = mvarsTitular
End Property

Public Property Let iPax(ByVal vData As Integer)
    mvariPax = vData
End Property

Public Property Get iPax() As Integer
    iPax = mvariPax
End Property

Public Property Let sMoeda(ByVal vData As String)
    mvarsMoeda = vData
End Property

Public Property Get sMoeda() As String
    sMoeda = mvarsMoeda
End Property

Public Property Let sControle(ByVal vData As String)
    mvarsControle = vData
End Property

Public Property Get sControle() As String
    sControle = mvarsControle
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lLinha(ByVal vData As Long)
    mvarlLinha = vData
End Property

Public Property Get lLinha() As Long
    lLinha = mvarlLinha
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iAporte(ByVal vData As Integer)
    mvariAporte = vData
End Property

Public Property Get iAporte() As Integer
    iAporte = mvariAporte
End Property

Public Property Let lClienteAporte(ByVal vData As Long)
    mvarlClienteAporte = vData
End Property

Public Property Get lClienteAporte() As Long
    lClienteAporte = mvarlClienteAporte
End Property

Public Property Let lNumIntDocAportePagto(ByVal vData As Long)
    mvarlNumIntDocAportePagto = vData
End Property

Public Property Get lNumIntDocAportePagto() As Long
    lNumIntDocAportePagto = mvarlNumIntDocAportePagto
End Property

Public Property Let iMarcado(ByVal vData As Integer)
    mvariMarcado = vData
End Property

Public Property Get iMarcado() As Integer
    iMarcado = mvariMarcado
End Property

Public Property Let sCodGrupo(ByVal vData As String)
    mvarsCodGrupo = vData
End Property

Public Property Get sCodGrupo() As String
    sCodGrupo = mvarsCodGrupo
End Property

Public Property Let lnumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lnumero() As Long
    lnumero = mvarlNumero
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let iCondPagto(ByVal vData As Integer)
    mvariCondPagto = vData
End Property

Public Property Get iCondPagto() As Integer
    iCondPagto = mvariCondPagto
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDocAporte(ByVal vData As Long)
    mvarlNumIntDocAporte = vData
End Property

Public Property Get lNumIntDocAporte() As Long
    lNumIntDocAporte = mvarlNumIntDocAporte
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lFatura(ByVal vData As Long)
    mvarlFatura = vData
End Property

Public Property Get lFatura() As Long
    lFatura = mvarlFatura
End Property

Public Property Let dValorAporte(ByVal vData As Double)
    mvardValorAporte = vData
End Property

Public Property Get dValorAporte() As Double
    dValorAporte = mvardValorAporte
End Property

Public Property Let dValorComissao(ByVal vData As Double)
    mvardValorComissao = vData
End Property

Public Property Get dValorComissao() As Double
    dValorComissao = mvardValorComissao
End Property

Public Property Let lNumVou(ByVal vData As Long)
    mvarlNumVou = vData
End Property

Public Property Get lNumVou() As Long
    lNumVou = mvarlNumVou
End Property

Public Property Let sTipoDoc(ByVal vData As String)
    mvarsTipoDoc = vData
End Property

Public Property Get sTipoDoc() As String
    sTipoDoc = mvarsTipoDoc
End Property

Public Property Let sTipoVou(ByVal vData As String)
    mvarsTipoVou = vData
End Property

Public Property Get sTipoVou() As String
    sTipoVou = mvarsTipoVou
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dValorBruto(ByVal vData As Double)
    mvardValorBruto = vData
End Property

Public Property Get dValorBruto() As Double
    dValorBruto = mvardValorBruto
End Property

