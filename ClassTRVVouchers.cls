VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVVouchers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumVou As Long
Private mvarsTipoDoc As String
Private mvarsSerie As String
Private mvarlNumIntArq As Long
Private mvarlNumRegArqTexto As Long
Private mvarlCliente As Long
Private mvarlRecnsica As Long
Private mvardValor As Double
Private mvardtData As Date
Private mvariCondPagto As Integer
Private mvarsTipVou As String
Private mvarsCodGrupo As String
Private mvarlNumFatCoinfo As Long
Private mvariTipoDocDestino As Integer
Private mvarlNumIntDocDestino As Long
Private mvariCartao As Integer
Private mvariStatus As Integer
Private mvarlNumIntDocNVL As Long
Private mvariPax As Integer
Private mvarsProduto As String
Private mvarsControle As String
Private mvarsMoeda As String
Private mvarsTitular As String
Private mvarsTitularCPF As String
Private mvardValorCambio As Double
Private mvardCambio As Double
Private mvarsCiaCart As String
Private mvarsNumCCred As String
Private mvarsValidadeCC As String
'Private mvarlNumAuto As Long
Private mvarsNumAuto As String
Private mvariQuantParc As Integer
Private mvariDiasAntc As Integer
Private mvariKit As Integer
Private mvariTemOcr As Integer
Private mvardValorOcr As Double
Private mvarcolTRVVoucherInfo As New Collection
Private mvarcolTRVGerComiIntDet As New Collection
Private mvarlRepresentante As Long
Private mvarlCorrentista As Long
Private mvarlEmissor As Long
Private mvardComissaoRep As Double
Private mvardComissaoCorr As Double
Private mvardComissaoEmissor As Double
Private mvardComissaoAg As Double
Private mvardValorBruto As Double
Private mvardValorAporte As Double
Private mvarlCodigoAporte As Long
Private mvariParcelaAporte As Integer
Private mvarlNumIntDocPagtoAporteFat As Long
Private mvariTipoPagtoAporte As Integer
Private mvarsPassageiroSobreNome As String
Private mvarsPassageiroNome As String
Private mvarlCliPassageiro As Long
Private mvarlClienteVou As Long
Private mvardAporteValor As Double
Private mvariAporteMoeda As Integer
Private mvariTipoCliente As Integer
Private mvardValorAcum As Double
Private mvardValorDolarAcum As Double
Private mvardValorAcumAporte As Double
Private mvardValorDolarAcumAporte As Double
Private mvarlClienteComissao As Long
Private mvarlPromotor As Long
Private mvarsDestino As String
Private mvarsUsuarioCanc As String
Private mvardValorComissao As Double
Private mvariExtraiInfoSigav As Integer
Private mvardtDataCanc As Date
Private mvardHoraCanc As Double
Private mvardTarifaUnitaria As Double
Private mvardPercComissaoProd As Double
Private mvariVendProd As Integer
Private mvariVendCli As Integer
Private mvardTarifaUNNormal As Double
Private mvariVigencia As Integer
Private mvarsUsuarioWeb As String
Private mvarcolVendedores As New Collection
Private mvardtDataFat As Date
Private mvariAporte As Integer
Private mvarsNomeRedCli As String
Private mvarsNomeRedProd As String
Private mvardValorBaseComis As Double
Private mvardValorBrutoComOCR As Double
Private mvardValorCMAComOCR As Double
Private mvardValorCMC As Double
Private mvardValorCMCC As Double
Private mvardValorCMR As Double
Private mvardValorCME As Double
Private mvarlNumIntDoc As Long
Private mvarcolEmiCI As New Collection
Private mvarsUsuRespCallCenter As String
Private mvarlNumBoleto As Long
Private mvardtDataVencBoleto As Date
Private mvardValorBoleto As Double
Private mvariTrataBoleto As Integer

Public Property Let iTrataBoleto(ByVal vData As Integer)
    mvariTrataBoleto = vData
End Property

Public Property Get iTrataBoleto() As Integer
    iTrataBoleto = mvariTrataBoleto
End Property

Public Property Let dValorBoleto(ByVal vData As Double)
    mvardValorBoleto = vData
End Property

Public Property Get dValorBoleto() As Double
    dValorBoleto = mvardValorBoleto
End Property

Public Property Let dtDataVencBoleto(ByVal vData As Date)
    mvardtDataVencBoleto = vData
End Property

Public Property Get dtDataVencBoleto() As Date
    dtDataVencBoleto = mvardtDataVencBoleto
End Property

Public Property Let lNumBoleto(ByVal vData As Long)
    mvarlNumBoleto = vData
End Property

Public Property Get lNumBoleto() As Long
    lNumBoleto = mvarlNumBoleto
End Property

Public Property Let sUsuRespCallCenter(ByVal vData As String)
    mvarsUsuRespCallCenter = vData
End Property

Public Property Get sUsuRespCallCenter() As String
    sUsuRespCallCenter = mvarsUsuRespCallCenter
End Property

Public Property Set colEmiCI(ByVal vData As Collection)
    Set mvarcolEmiCI = vData
End Property

Public Property Get colEmiCI() As Collection
    Set colEmiCI = mvarcolEmiCI
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sNumAuto(ByVal vData As String)
    mvarsNumAuto = vData
End Property

Public Property Get sNumAuto() As String
    sNumAuto = mvarsNumAuto
End Property

Public Property Let dValorCME(ByVal vData As Double)
    mvardValorCME = vData
End Property

Public Property Get dValorCME() As Double
    dValorCME = mvardValorCME
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

Public Property Let dValorCMC(ByVal vData As Double)
    mvardValorCMC = vData
End Property

Public Property Get dValorCMC() As Double
    dValorCMC = mvardValorCMC
End Property

Public Property Let dValorCMAComOCR(ByVal vData As Double)
    mvardValorCMAComOCR = vData
End Property

Public Property Get dValorCMAComOCR() As Double
    dValorCMAComOCR = mvardValorCMAComOCR
End Property

Public Property Let dValorBrutoComOCR(ByVal vData As Double)
    mvardValorBrutoComOCR = vData
End Property

Public Property Get dValorBrutoComOCR() As Double
    dValorBrutoComOCR = mvardValorBrutoComOCR
End Property

Public Property Let dValorBaseComis(ByVal vData As Double)
    mvardValorBaseComis = vData
End Property

Public Property Get dValorBaseComis() As Double
    dValorBaseComis = mvardValorBaseComis
End Property

Public Property Let sNomeRedProd(ByVal vData As String)
    mvarsNomeRedProd = vData
End Property

Public Property Get sNomeRedProd() As String
    sNomeRedProd = mvarsNomeRedProd
End Property

Public Property Let sNomeRedCli(ByVal vData As String)
    mvarsNomeRedCli = vData
End Property

Public Property Get sNomeRedCli() As String
    sNomeRedCli = mvarsNomeRedCli
End Property

Public Property Let iAporte(ByVal vData As Integer)
    mvariAporte = vData
End Property

Public Property Get iAporte() As Integer
    iAporte = mvariAporte
End Property

Public Property Let dtDataFat(ByVal vData As Date)
    mvardtDataFat = vData
End Property

Public Property Get dtDataFat() As Date
    dtDataFat = mvardtDataFat
End Property

Public Property Set colVendedores(ByVal vData As Collection)
    Set mvarcolVendedores = vData
End Property

Public Property Get colVendedores() As Collection
    Set colVendedores = mvarcolVendedores
End Property

Public Property Let sUsuarioWeb(ByVal vData As String)
    mvarsUsuarioWeb = vData
End Property

Public Property Get sUsuarioWeb() As String
    sUsuarioWeb = mvarsUsuarioWeb
End Property

Public Property Let iVigencia(ByVal vData As Integer)
    mvariVigencia = vData
End Property

Public Property Get iVigencia() As Integer
    iVigencia = mvariVigencia
End Property

Public Property Let dTarifaUNNormal(ByVal vData As Double)
    mvardTarifaUNNormal = vData
End Property

Public Property Get dTarifaUNNormal() As Double
    dTarifaUNNormal = mvardTarifaUNNormal
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

Public Property Let dTarifaUnitaria(ByVal vData As Double)
    mvardTarifaUnitaria = vData
End Property

Public Property Get dTarifaUnitaria() As Double
    dTarifaUnitaria = mvardTarifaUnitaria
End Property

Public Property Let dHoraCanc(ByVal vData As Double)
    mvardHoraCanc = vData
End Property

Public Property Get dHoraCanc() As Double
    dHoraCanc = mvardHoraCanc
End Property

Public Property Let dtDataCanc(ByVal vData As Date)
    mvardtDataCanc = vData
End Property

Public Property Get dtDataCanc() As Date
    dtDataCanc = mvardtDataCanc
End Property

Public Property Let iExtraiInfoSigav(ByVal vData As Integer)
    mvariExtraiInfoSigav = vData
End Property

Public Property Get iExtraiInfoSigav() As Integer
    iExtraiInfoSigav = mvariExtraiInfoSigav
End Property

Public Property Let dValorComissao(ByVal vData As Double)
    mvardValorComissao = vData
End Property

Public Property Get dValorComissao() As Double
    dValorComissao = mvardValorComissao
End Property

Public Property Let sUsuarioCanc(ByVal vData As String)
    mvarsUsuarioCanc = vData
End Property

Public Property Get sUsuarioCanc() As String
    sUsuarioCanc = mvarsUsuarioCanc
End Property

Public Property Let sDestino(ByVal vData As String)
    mvarsDestino = vData
End Property

Public Property Get sDestino() As String
    sDestino = mvarsDestino
End Property

Public Property Let lPromotor(ByVal vData As Long)
    mvarlPromotor = vData
End Property

Public Property Get lPromotor() As Long
    lPromotor = mvarlPromotor
End Property

Public Property Let lClienteComissao(ByVal vData As Long)
    mvarlClienteComissao = vData
End Property

Public Property Get lClienteComissao() As Long
    If mvarlClienteComissao = 0 Then
        lClienteComissao = mvarlClienteVou
    Else
        lClienteComissao = mvarlClienteComissao
    End If
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

Public Property Let lClienteVou(ByVal vData As Long)
    mvarlClienteVou = vData
End Property

Public Property Get lClienteVou() As Long
    lClienteVou = mvarlClienteVou
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

Public Property Let iTipoPagtoAporte(ByVal vData As Integer)
    mvariTipoPagtoAporte = vData
End Property

Public Property Get iTipoPagtoAporte() As Integer
    iTipoPagtoAporte = mvariTipoPagtoAporte
End Property

Public Property Let lNumIntDocPagtoAporteFat(ByVal vData As Long)
    mvarlNumIntDocPagtoAporteFat = vData
End Property

Public Property Get lNumIntDocPagtoAporteFat() As Long
    lNumIntDocPagtoAporteFat = mvarlNumIntDocPagtoAporteFat
End Property

Public Property Let iParcelaAporte(ByVal vData As Integer)
    mvariParcelaAporte = vData
End Property

Public Property Get iParcelaAporte() As Integer
    iParcelaAporte = mvariParcelaAporte
End Property

Public Property Let lCodigoAporte(ByVal vData As Long)
    mvarlCodigoAporte = vData
End Property

Public Property Get lCodigoAporte() As Long
    lCodigoAporte = mvarlCodigoAporte
End Property

Public Property Let dValorAporte(ByVal vData As Double)
    mvardValorAporte = vData
End Property

Public Property Get dValorAporte() As Double
    dValorAporte = mvardValorAporte
End Property

Public Property Let dValorBruto(ByVal vData As Double)
    mvardValorBruto = vData
End Property

Public Property Get dValorBruto() As Double
    dValorBruto = mvardValorBruto
End Property

Public Property Let dComissaoAg(ByVal vData As Double)
    mvardComissaoAg = vData
End Property

Public Property Get dComissaoAg() As Double
    dComissaoAg = mvardComissaoAg
End Property

Public Property Let dComissaoEmissor(ByVal vData As Double)
    mvardComissaoEmissor = vData
End Property

Public Property Get dComissaoEmissor() As Double
    dComissaoEmissor = mvardComissaoEmissor
End Property

Public Property Let dComissaoCorr(ByVal vData As Double)
    mvardComissaoCorr = vData
End Property

Public Property Get dComissaoCorr() As Double
    dComissaoCorr = mvardComissaoCorr
End Property

Public Property Let dComissaoRep(ByVal vData As Double)
    mvardComissaoRep = vData
End Property

Public Property Get dComissaoRep() As Double
    dComissaoRep = mvardComissaoRep
End Property

Public Property Let lEmissor(ByVal vData As Long)
    mvarlEmissor = vData
End Property

Public Property Get lEmissor() As Long
    lEmissor = mvarlEmissor
End Property

Public Property Let lCorrentista(ByVal vData As Long)
    mvarlCorrentista = vData
End Property

Public Property Get lCorrentista() As Long
    lCorrentista = mvarlCorrentista
End Property

Public Property Let lRepresentante(ByVal vData As Long)
    mvarlRepresentante = vData
End Property

Public Property Get lRepresentante() As Long
    lRepresentante = mvarlRepresentante
End Property

Public Property Set colTRVGerComiIntDet(ByVal vData As Collection)
    Set mvarcolTRVGerComiIntDet = vData
End Property

Public Property Get colTRVGerComiIntDet() As Collection
    Set colTRVGerComiIntDet = mvarcolTRVGerComiIntDet
End Property

Public Property Set colTRVVoucherInfo(ByVal vData As Collection)
    Set mvarcolTRVVoucherInfo = vData
End Property

Public Property Get colTRVVoucherInfo() As Collection
    Set colTRVVoucherInfo = mvarcolTRVVoucherInfo
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

Public Property Let iKit(ByVal vData As Integer)
    mvariKit = vData
End Property

Public Property Get iKit() As Integer
    iKit = mvariKit
End Property

Public Property Let idiasantc(ByVal vData As Integer)
    mvariDiasAntc = vData
End Property

Public Property Get idiasantc() As Integer
    idiasantc = mvariDiasAntc
End Property

Public Property Let iQuantParc(ByVal vData As Integer)
    mvariQuantParc = vData
End Property

Public Property Get iQuantParc() As Integer
    iQuantParc = mvariQuantParc
End Property

'Public Property Let lNumAuto(ByVal vData As Long)
'    mvarlNumAuto = vData
'End Property
'
'Public Property Get lNumAuto() As Long
'    lNumAuto = mvarlNumAuto
'End Property

Public Property Let sNumCCred(ByVal vData As String)
    mvarsNumCCred = vData
End Property

Public Property Get sNumCCred() As String
    sNumCCred = mvarsNumCCred
End Property

Public Property Let sValidadeCC(ByVal vData As String)
    mvarsValidadeCC = vData
End Property

Public Property Get sValidadeCC() As String
    sValidadeCC = mvarsValidadeCC
End Property

Public Property Let sCiaCart(ByVal vData As String)
    mvarsCiaCart = vData
End Property

Public Property Get sCiaCart() As String
    sCiaCart = mvarsCiaCart
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

Public Property Let sTitularCPF(ByVal vData As String)
    mvarsTitularCPF = vData
End Property

Public Property Get sTitularCPF() As String
    sTitularCPF = mvarsTitularCPF
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

Public Property Let lNumIntDocNVL(ByVal vData As Long)
    mvarlNumIntDocNVL = vData
End Property

Public Property Get lNumIntDocNVL() As Long
    lNumIntDocNVL = mvarlNumIntDocNVL
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
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

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let lNumIntArq(ByVal vData As Long)
    mvarlNumIntArq = vData
End Property

Public Property Get lNumIntArq() As Long
    lNumIntArq = mvarlNumIntArq
End Property

Public Property Let lNumRegArqTexto(ByVal vData As Long)
    mvarlNumRegArqTexto = vData
End Property

Public Property Get lNumRegArqTexto() As Long
    lNumRegArqTexto = mvarlNumRegArqTexto
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lrecnsica(ByVal vData As Long)
    mvarlRecnsica = vData
End Property

Public Property Get lrecnsica() As Long
    lrecnsica = mvarlRecnsica
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

Public Property Let iCondPagto(ByVal vData As Integer)
    mvariCondPagto = vData
End Property

Public Property Get iCondPagto() As Integer
    iCondPagto = mvariCondPagto
End Property

Public Property Let sTipVou(ByVal vData As String)
    mvarsTipVou = vData
End Property

Public Property Get sTipVou() As String
    sTipVou = mvarsTipVou
End Property

Public Property Let sCodGrupo(ByVal vData As String)
    mvarsCodGrupo = vData
End Property

Public Property Get sCodGrupo() As String
    sCodGrupo = mvarsCodGrupo
End Property

Public Property Let lNumFatCoinfo(ByVal vData As Long)
    mvarlNumFatCoinfo = vData
End Property

Public Property Get lNumFatCoinfo() As Long
    lNumFatCoinfo = mvarlNumFatCoinfo
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

Private Sub Class_Initialize()
    mvardtDataVencBoleto = DATA_NULA
End Sub
