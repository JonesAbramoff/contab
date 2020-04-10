VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPVouchers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumVou As Long
Private mvarsSerie As String
Private mvarlClienteFat As Long
Private mvardValor As Double
Private mvardtData As Date
Private mvarsTipVou As String
Private mvarlNumFat As Long
Private mvariTipoDocDestino As Integer
Private mvarlNumIntDocDestino As Long
Private mvariCartao As Integer
Private mvariStatus As Integer
Private mvarlNumIntDocNVL As Long
Private mvardValorAporte As Double
Private mvardValorComissao As Double
Private mvarsTitular As String
Private mvarsProduto As String
Private mvariPax As Integer
Private mvariMoeda As Integer
Private mvardValorCambio As Double
Private mvardCambio As Double
Private mvarsControle As String
Private mvarsCiaCart As String
Private mvarsNumCCred As String
Private mvarlNumAuto As Long
Private mvariQuantParc As Integer
Private mvariDiasAntc As Integer
Private mvariKit As Integer
Private mvardValorOcr As Double
Private mvariTemOcr As Integer
Private mvarlRepresentante As Long
Private mvardComissaoRep As Double
Private mvarlCorrentista As Long
Private mvardComissaoCorr As Double
Private mvardComissaoProm As Double
Private mvarlEmissor As Long
Private mvardComissaoEmissor As Double
Private mvardComissaoAg As Double
Private mvardValorBruto As Double
Private mvarsPassageiroNome As String
Private mvarsPassageiroSobrenome As String
Private mvarlClienteVou As Long
Private mvarlCliPassageiro As Long
Private mvardtDataCanc As Date
Private mvardtDataVigenciaDe As Date
Private mvardtDataVigenciaAte As Date
Private mvardHoraCanc As Double
Private mvarsUsuarioCanc As String
Private mvarlClienteComissao As Long
Private mvarsCiaAerea As String
Private mvarsAeroportos As String
Private mvarlEnderecoPaxTitular As Long
Private mvariTemQueContabilizar As Integer
Private mvariPromotor As Integer
Private mvariDestino As Integer
Private mvarsTitularCPF As String
Private mvardTarifaUnitaria As Double
Private mvariVigencia As Integer
Private mvardTarifaUnitariaFolheto As Double
Private mvarcolContatos As New Collection
Private mvarcolPassageiros As New Collection
Private mvarsPassageiroCGC As String
Private mvardtPassageiroDataNasc As Date
Private mvariGeraComissao As Integer
Private mvariCancelaComissao As Integer
Private mvariIdioma As Integer
Private mvariGrupo As Integer
Private mvarcolVendedores As New Collection
Private mvardtDataUltimaManut As Date
Private mvardHoraUltimaManut As Double
Private mvarsUsuarioManut As String
Private mvardtDataAutoCC As Date
Private mvarsValidadeCC As String
Private mvariImprimirValor As Integer
Private mvariCodSegurancaCC As Integer
Private mvardValorBrutoComOCR As Double
Private mvardValorCMAComOCR As Double
Private mvardValorCMC As Double
Private mvardValorCMR As Double
Private mvardValorCMCC As Double
Private mvardValorCME As Double
Private mvarsObservacao As String
Private mvardtDataLibManut As Date
Private mvardHoraLibManut As Double
Private mvarsUsuarioLibManut As String

'AUXILIAR
Private mvarobjEnderecoPax As New ClassEndereco
Private mvarcolTRPVoucherInfo As New Collection
Private mvarcolTRPGerComiIntDet As New Collection
Private mvardAporteValor As Double
Private mvariAporteMoeda As Integer
Private mvariTipoCliente As Integer
Private mvardValorAcum As Double
Private mvardValorDolarAcum As Double
Private mvardValorAcumAporte As Double
Private mvardValorDolarAcumAporte As Double
Private mvardPercComissaoProd As Double
Private mvariVendProd As Integer
Private mvariVendCli As Integer
Private mvarlNumIntDocPagtoAporteFat As Long
Private mvariTipoPagtoAporte As Integer

Public Property Let lNumIntDocPagtoAporteFat(ByVal vData As Long)
    mvarlNumIntDocPagtoAporteFat = vData
End Property

Public Property Get lNumIntDocPagtoAporteFat() As Long
    lNumIntDocPagtoAporteFat = mvarlNumIntDocPagtoAporteFat
End Property

Public Property Let iTipoPagtoAporte(ByVal vData As Integer)
    mvariTipoPagtoAporte = vData
End Property

Public Property Get iTipoPagtoAporte() As Integer
    iTipoPagtoAporte = mvariTipoPagtoAporte
End Property

Public Property Let iVendCli(ByVal vData As Integer)
    mvariVendCli = vData
End Property

Public Property Get iVendCli() As Integer
    iVendCli = mvariVendCli
End Property

Public Property Let iVendProd(ByVal vData As Integer)
    mvariVendProd = vData
End Property

Public Property Get iVendProd() As Integer
    iVendProd = mvariVendProd
End Property

Public Property Let dPercComissaoProd(ByVal vData As Double)
    mvardPercComissaoProd = vData
End Property

Public Property Get dPercComissaoProd() As Double
    dPercComissaoProd = mvardPercComissaoProd
End Property

Public Property Let dValorDolarAcumAporte(ByVal vData As Double)
    mvardValorDolarAcumAporte = vData
End Property

Public Property Get dValorDolarAcumAporte() As Double
    dValorDolarAcumAporte = mvardValorDolarAcumAporte
End Property

Public Property Let dValorAcumAporte(ByVal vData As Double)
    mvardValorAcumAporte = vData
End Property

Public Property Get dValorAcumAporte() As Double
    dValorAcumAporte = mvardValorAcumAporte
End Property

Public Property Let dValorDolarAcum(ByVal vData As Double)
    mvardValorDolarAcum = vData
End Property

Public Property Get dValorDolarAcum() As Double
    dValorDolarAcum = mvardValorDolarAcum
End Property

Public Property Let dValorAcum(ByVal vData As Double)
    mvardValorAcum = vData
End Property

Public Property Get dValorAcum() As Double
    dValorAcum = mvardValorAcum
End Property

Public Property Let iTipoCliente(ByVal vData As Integer)
    mvariTipoCliente = vData
End Property

Public Property Get iTipoCliente() As Integer
    iTipoCliente = mvariTipoCliente
End Property

Public Property Let iAporteMoeda(ByVal vData As Integer)
    mvariAporteMoeda = vData
End Property

Public Property Get iAporteMoeda() As Integer
    iAporteMoeda = mvariAporteMoeda
End Property

Public Property Let dAporteValor(ByVal vData As Double)
    mvardAporteValor = vData
End Property

Public Property Get dAporteValor() As Double
    dAporteValor = mvardAporteValor
End Property

Public Property Set colTRPGerComiIntDet(ByVal vData As Collection)
    Set mvarcolTRPGerComiIntDet = vData
End Property

Public Property Get colTRPGerComiIntDet() As Collection
    Set colTRPGerComiIntDet = mvarcolTRPGerComiIntDet
End Property

Public Property Set colTRPVoucherInfo(ByVal vData As Collection)
    Set mvarcolTRPVoucherInfo = vData
End Property

Public Property Get colTRPVoucherInfo() As Collection
    Set colTRPVoucherInfo = mvarcolTRPVoucherInfo
End Property

Public Property Set objEnderecoPax(ByVal vData As ClassEndereco)
    Set mvarobjEnderecoPax = vData
End Property

Public Property Get objEnderecoPax() As ClassEndereco
    Set objEnderecoPax = mvarobjEnderecoPax
End Property

'FIM AUXILIAR

Public Property Let iImprimirValor(ByVal vData As Integer)
    mvariImprimirValor = vData
End Property

Public Property Get iImprimirValor() As Integer
    iImprimirValor = mvariImprimirValor
End Property

Public Property Let iCodSegurancaCC(ByVal vData As Integer)
    mvariCodSegurancaCC = vData
End Property

Public Property Get iCodSegurancaCC() As Integer
    iCodSegurancaCC = mvariCodSegurancaCC
End Property

Public Property Let dValorBrutoComOCR(ByVal vData As Double)
    mvardValorBrutoComOCR = vData
End Property

Public Property Get dValorBrutoComOCR() As Double
    dValorBrutoComOCR = mvardValorBrutoComOCR
End Property

Public Property Let dValorCMAComOCR(ByVal vData As Double)
    mvardValorCMAComOCR = vData
End Property

Public Property Get dValorCMAComOCR() As Double
    dValorCMAComOCR = mvardValorCMAComOCR
End Property

Public Property Let dValorCMC(ByVal vData As Double)
    mvardValorCMC = vData
End Property

Public Property Get dValorCMC() As Double
    dValorCMC = mvardValorCMC
End Property

Public Property Let dValorCMR(ByVal vData As Double)
    mvardValorCMR = vData
End Property

Public Property Get dValorCMR() As Double
    dValorCMR = mvardValorCMR
End Property

Public Property Let dValorCMCC(ByVal vData As Double)
    mvardValorCMCC = vData
End Property

Public Property Get dValorCMCC() As Double
    dValorCMCC = mvardValorCMCC
End Property

Public Property Let dValorCME(ByVal vData As Double)
    mvardValorCME = vData
End Property

Public Property Get dValorCME() As Double
    dValorCME = mvardValorCME
End Property

Public Property Let sValidadeCC(ByVal vData As String)
    mvarsValidadeCC = vData
End Property

Public Property Get sValidadeCC() As String
    sValidadeCC = mvarsValidadeCC
End Property

Public Property Let dtDataAutoCC(ByVal vData As Date)
    mvardtDataAutoCC = vData
End Property

Public Property Get dtDataAutoCC() As Date
    dtDataAutoCC = mvardtDataAutoCC
End Property

Public Property Let dtDataUltimaManut(ByVal vData As Date)
    mvardtDataUltimaManut = vData
End Property

Public Property Get dtDataUltimaManut() As Date
    dtDataUltimaManut = mvardtDataUltimaManut
End Property

Public Property Let dHoraUltimaManut(ByVal vData As Double)
    mvardHoraUltimaManut = vData
End Property

Public Property Get dHoraUltimaManut() As Double
    dHoraUltimaManut = mvardHoraUltimaManut
End Property

Public Property Let sUsuarioManut(ByVal vData As String)
    mvarsUsuarioManut = vData
End Property

Public Property Get sUsuarioManut() As String
    sUsuarioManut = mvarsUsuarioManut
End Property

Public Property Set colVendedores(ByVal vData As Collection)
    Set mvarcolVendedores = vData
End Property

Public Property Get colVendedores() As Collection
    Set colVendedores = mvarcolVendedores
End Property

Public Property Let iIdioma(ByVal vData As Integer)
    mvariIdioma = vData
End Property

Public Property Get iIdioma() As Integer
    iIdioma = mvariIdioma
End Property

Public Property Let iGrupo(ByVal vData As Integer)
    mvariGrupo = vData
End Property

Public Property Get iGrupo() As Integer
    iGrupo = mvariGrupo
End Property

Public Property Let iGeraComissao(ByVal vData As Integer)
    mvariGeraComissao = vData
End Property

Public Property Get iGeraComissao() As Integer
    iGeraComissao = mvariGeraComissao
End Property

Public Property Let iCancelaComissao(ByVal vData As Integer)
    mvariCancelaComissao = vData
End Property

Public Property Get iCancelaComissao() As Integer
    iCancelaComissao = mvariCancelaComissao
End Property

Public Property Let dtPassageiroDataNasc(ByVal vData As Date)
    mvardtPassageiroDataNasc = vData
End Property

Public Property Get dtPassageiroDataNasc() As Date
    dtPassageiroDataNasc = mvardtPassageiroDataNasc
End Property

Public Property Set colPassageiros(ByVal vData As Collection)
    Set mvarcolPassageiros = vData
End Property

Public Property Get colPassageiros() As Collection
    Set colPassageiros = mvarcolPassageiros
End Property

Public Property Set colContatos(ByVal vData As Collection)
    Set mvarcolContatos = vData
End Property

Public Property Get colContatos() As Collection
    Set colContatos = mvarcolContatos
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumVou(ByVal vData As Long)
    mvarlNumVou = vData
End Property

Public Property Get lNumVou() As Long
    lNumVou = mvarlNumVou
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let lClienteFat(ByVal vData As Long)
    mvarlClienteFat = vData
End Property

Public Property Get lClienteFat() As Long
    lClienteFat = mvarlClienteFat
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sTipVou(ByVal vData As String)
    mvarsTipVou = vData
End Property

Public Property Get sTipVou() As String
    sTipVou = mvarsTipVou
End Property

Public Property Let lNumFat(ByVal vData As Long)
    mvarlNumFat = vData
End Property

Public Property Get lNumFat() As Long
    lNumFat = mvarlNumFat
End Property

Public Property Let iTipoDocDestino(ByVal vData As Integer)
    mvariTipoDocDestino = vData
End Property

Public Property Get iTipoDocDestino() As Integer
    iTipoDocDestino = mvariTipoDocDestino
End Property

Public Property Let lNumIntDocDestino(ByVal vData As Long)
    mvarlNumIntDocDestino = vData
End Property

Public Property Get lNumIntDocDestino() As Long
    lNumIntDocDestino = mvarlNumIntDocDestino
End Property

Public Property Let iCartao(ByVal vData As Integer)
    mvariCartao = vData
End Property

Public Property Get iCartao() As Integer
    iCartao = mvariCartao
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let lNumIntDocNVL(ByVal vData As Long)
    mvarlNumIntDocNVL = vData
End Property

Public Property Get lNumIntDocNVL() As Long
    lNumIntDocNVL = mvarlNumIntDocNVL
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

Public Property Let sTitular(ByVal vData As String)
    mvarsTitular = vData
End Property

Public Property Get sTitular() As String
    sTitular = mvarsTitular
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iPax(ByVal vData As Integer)
    mvariPax = vData
End Property

Public Property Get iPax() As Integer
    iPax = mvariPax
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let dValorCambio(ByVal vData As Double)
    mvardValorCambio = vData
End Property

Public Property Get dValorCambio() As Double
    dValorCambio = mvardValorCambio
End Property

Public Property Let dCambio(ByVal vData As Double)
    mvardCambio = vData
End Property

Public Property Get dCambio() As Double
    dCambio = mvardCambio
End Property

Public Property Let sControle(ByVal vData As String)
    mvarsControle = vData
End Property

Public Property Get sControle() As String
    sControle = mvarsControle
End Property

Public Property Let sCiaCart(ByVal vData As String)
    mvarsCiaCart = vData
End Property

Public Property Get sCiaCart() As String
    sCiaCart = mvarsCiaCart
End Property

Public Property Let sNumCCred(ByVal vData As String)
    mvarsNumCCred = vData
End Property

Public Property Get sNumCCred() As String
    sNumCCred = mvarsNumCCred
End Property

Public Property Let lNumAuto(ByVal vData As Long)
    mvarlNumAuto = vData
End Property

Public Property Get lNumAuto() As Long
    lNumAuto = mvarlNumAuto
End Property

Public Property Let iQuantParc(ByVal vData As Integer)
    mvariQuantParc = vData
End Property

Public Property Get iQuantParc() As Integer
    iQuantParc = mvariQuantParc
End Property

Public Property Let iDiasAntc(ByVal vData As Integer)
    mvariDiasAntc = vData
End Property

Public Property Get iDiasAntc() As Integer
    iDiasAntc = mvariDiasAntc
End Property

Public Property Let iKit(ByVal vData As Integer)
    mvariKit = vData
End Property

Public Property Get iKit() As Integer
    iKit = mvariKit
End Property

Public Property Let dValorOcr(ByVal vData As Double)
    mvardValorOcr = vData
End Property

Public Property Get dValorOcr() As Double
    dValorOcr = mvardValorOcr
End Property

Public Property Let iTemOcr(ByVal vData As Integer)
    mvariTemOcr = vData
End Property

Public Property Get iTemOcr() As Integer
    iTemOcr = mvariTemOcr
End Property

Public Property Let lRepresentante(ByVal vData As Long)
    mvarlRepresentante = vData
End Property

Public Property Get lRepresentante() As Long
    lRepresentante = mvarlRepresentante
End Property

Public Property Let dComissaoRep(ByVal vData As Double)
    mvardComissaoRep = vData
End Property

Public Property Get dComissaoRep() As Double
    dComissaoRep = mvardComissaoRep
End Property

Public Property Let lCorrentista(ByVal vData As Long)
    mvarlCorrentista = vData
End Property

Public Property Get lCorrentista() As Long
    lCorrentista = mvarlCorrentista
End Property

Public Property Let dComissaoProm(ByVal vData As Double)
    mvardComissaoProm = vData
End Property

Public Property Get dComissaoProm() As Double
    dComissaoProm = mvardComissaoProm
End Property

Public Property Let dComissaoCorr(ByVal vData As Double)
    mvardComissaoCorr = vData
End Property

Public Property Get dComissaoCorr() As Double
    dComissaoCorr = mvardComissaoCorr
End Property

Public Property Let lEmissor(ByVal vData As Long)
    mvarlEmissor = vData
End Property

Public Property Get lEmissor() As Long
    lEmissor = mvarlEmissor
End Property

Public Property Let dComissaoEmissor(ByVal vData As Double)
    mvardComissaoEmissor = vData
End Property

Public Property Get dComissaoEmissor() As Double
    dComissaoEmissor = mvardComissaoEmissor
End Property

Public Property Let dComissaoAg(ByVal vData As Double)
    mvardComissaoAg = vData
End Property

Public Property Get dComissaoAg() As Double
    dComissaoAg = mvardComissaoAg
End Property

Public Property Let dValorBruto(ByVal vData As Double)
    mvardValorBruto = vData
End Property

Public Property Get dValorBruto() As Double
    dValorBruto = mvardValorBruto
End Property

Public Property Let sPassageiroNome(ByVal vData As String)
    mvarsPassageiroNome = vData
End Property

Public Property Get sPassageiroNome() As String
    sPassageiroNome = mvarsPassageiroNome
End Property

Public Property Let sPassageiroCGC(ByVal vData As String)
    mvarsPassageiroCGC = vData
End Property

Public Property Get sPassageiroCGC() As String
    sPassageiroCGC = mvarsPassageiroCGC
End Property

Public Property Let sPassageiroSobreNome(ByVal vData As String)
    mvarsPassageiroSobrenome = vData
End Property

Public Property Get sPassageiroSobreNome() As String
    sPassageiroSobreNome = mvarsPassageiroSobrenome
End Property

Public Property Let lClienteVou(ByVal vData As Long)
    mvarlClienteVou = vData
End Property

Public Property Get lClienteVou() As Long
    lClienteVou = mvarlClienteVou
End Property

Public Property Let lCliPassageiro(ByVal vData As Long)
    mvarlCliPassageiro = vData
End Property

Public Property Get lCliPassageiro() As Long
    lCliPassageiro = mvarlCliPassageiro
End Property

Public Property Let dtDataCanc(ByVal vData As Date)
    mvardtDataCanc = vData
End Property

Public Property Get dtDataCanc() As Date
    dtDataCanc = mvardtDataCanc
End Property

Public Property Let dtDataVigenciaDe(ByVal vData As Date)
    mvardtDataVigenciaDe = vData
End Property

Public Property Get dtDataVigenciaDe() As Date
    dtDataVigenciaDe = mvardtDataVigenciaDe
End Property

Public Property Let dtDataVigenciaAte(ByVal vData As Date)
    mvardtDataVigenciaAte = vData
End Property

Public Property Get dtDataVigenciaAte() As Date
    dtDataVigenciaAte = mvardtDataVigenciaAte
End Property

Public Property Let dHoraCanc(ByVal vData As Double)
    mvardHoraCanc = vData
End Property

Public Property Get dHoraCanc() As Double
    dHoraCanc = mvardHoraCanc
End Property

Public Property Let sUsuarioCanc(ByVal vData As String)
    mvarsUsuarioCanc = vData
End Property

Public Property Get sUsuarioCanc() As String
    sUsuarioCanc = mvarsUsuarioCanc
End Property

Public Property Let lClienteComissao(ByVal vData As Long)
    mvarlClienteComissao = vData
End Property

Public Property Get lClienteComissao() As Long
    lClienteComissao = mvarlClienteComissao
End Property

Public Property Let sCiaaerea(ByVal vData As String)
    mvarsCiaAerea = vData
End Property

Public Property Get sCiaaerea() As String
    sCiaaerea = mvarsCiaAerea
End Property

Public Property Let sAeroportos(ByVal vData As String)
    mvarsAeroportos = vData
End Property

Public Property Get sAeroportos() As String
    sAeroportos = mvarsAeroportos
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let lEnderecoPaxTitular(ByVal vData As Long)
    mvarlEnderecoPaxTitular = vData
End Property

Public Property Get lEnderecoPaxTitular() As Long
    lEnderecoPaxTitular = mvarlEnderecoPaxTitular
End Property

Public Property Let iTemQueContabilizar(ByVal vData As Integer)
    mvariTemQueContabilizar = vData
End Property

Public Property Get iTemQueContabilizar() As Integer
    iTemQueContabilizar = mvariTemQueContabilizar
End Property

Public Property Let iPromotor(ByVal vData As Integer)
    mvariPromotor = vData
End Property

Public Property Get iPromotor() As Integer
    iPromotor = mvariPromotor
End Property

Public Property Let iDestino(ByVal vData As Integer)
    mvariDestino = vData
End Property

Public Property Get iDestino() As Integer
    iDestino = mvariDestino
End Property

Public Property Let sTitularCPF(ByVal vData As String)
    mvarsTitularCPF = vData
End Property

Public Property Get sTitularCPF() As String
    sTitularCPF = mvarsTitularCPF
End Property

Public Property Let dTarifaUnitaria(ByVal vData As Double)
    mvardTarifaUnitaria = vData
End Property

Public Property Get dTarifaUnitaria() As Double
    dTarifaUnitaria = mvardTarifaUnitaria
End Property

Public Property Let iVigencia(ByVal vData As Integer)
    mvariVigencia = vData
End Property

Public Property Get iVigencia() As Integer
    iVigencia = mvariVigencia
End Property

Public Property Let dTarifaUnitariaFolheto(ByVal vData As Double)
    mvardTarifaUnitariaFolheto = vData
End Property

Public Property Get dTarifaUnitariaFolheto() As Double
    dTarifaUnitariaFolheto = mvardTarifaUnitariaFolheto
End Property

Public Property Let dtDataLibManut(ByVal vData As Date)
    mvardtDataLibManut = vData
End Property

Public Property Get dtDataLibManut() As Date
    dtDataLibManut = mvardtDataLibManut
End Property

Public Property Let dHoraLibManut(ByVal vData As Double)
    mvardHoraLibManut = vData
End Property

Public Property Get dHoraLibManut() As Double
    dHoraLibManut = mvardHoraLibManut
End Property

Public Property Let sUsuarioLibManut(ByVal vData As String)
    mvarsUsuarioLibManut = vData
End Property

Public Property Get sUsuarioLibManut() As String
    sUsuarioLibManut = mvarsUsuarioLibManut
End Property
