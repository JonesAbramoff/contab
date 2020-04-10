VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocCab As Long
Private mvarlNumIntDocNF As Long
Private mvarlNumNFe As Long
Private mvardtData As Date
Private mvardHora As Double
Private mvarsCodVerificacao As String
Private mvarsTipoRPS As String
Private mvarsSerieRPS As String
Private mvarlNumeroRPS As Long
Private mvardtDataEmissaoRPS As Date
Private mvarsInscricaoMunicialPrest As String
Private mvariTipoCGCPrest As Integer
Private mvarsCGCPrest As String
Private mvarsRazaoSocialPrest As String
Private mvarsTipoEnderecoPrest As String
Private mvarsEnderecoPrest As String
Private mvarsEndNumeroPrest As String
Private mvarsEndComplementoPrest As String
Private mvarsBairroPrest As String
Private mvarsCidadePrest As String
Private mvarsUFPrest As String
Private mvarsCEPPrest As String
Private mvarsEmailPrest As String
Private mvariOPTSimples As Integer
Private mvarsSituacaoNF As String
Private mvardtDataCancelamento As Date
Private mvarsNumGuia As String
Private mvardtDataQuitacaoGuia As Date
Private mvardValorServicos As Double
Private mvardValorDeducoes As Double
Private mvarlCodServico As Long
Private mvardAliquota As Double
Private mvardValorISS As Double
Private mvardValorCredito As Double
Private mvarsISSRetido As String
Private mvariTipoCGCTom As Integer
Private mvarsCGCTom As String
Private mvarsInscricaoMunicipalTom As String
Private mvarsInscricaoEstadualTom As String
Private mvarsRazaoSocialTom As String
Private mvarsTipoEnderecoTom As String
Private mvarsEnderecoTom As String
Private mvarsEndNumeroTom As String
Private mvarsEndComplementoTom As String
Private mvarsBairroTom As String
Private mvarsCidadeTom As String
Private mvarsUFTom As String
Private mvarsCEPTom As String
Private mvarsEmailTom As String
Private mvarsDiscriminacao As String
Private mvariFilialEmpresa As Integer
Private mvardValorCofins As Double
Private mvardValorCSLL As Double
Private mvardValorINSS As Double
Private mvardValorIRPJ As Double
Private mvardValorPIS As Double

Public Property Let dValorPIS(ByVal vData As Double)
    mvardValorPIS = vData
End Property

Public Property Get dValorPIS() As Double
    dValorPIS = mvardValorPIS
End Property

Public Property Let dValorIRPJ(ByVal vData As Double)
    mvardValorIRPJ = vData
End Property

Public Property Get dValorIRPJ() As Double
    dValorIRPJ = mvardValorIRPJ
End Property

Public Property Let dValorINSS(ByVal vData As Double)
    mvardValorINSS = vData
End Property

Public Property Get dValorINSS() As Double
    dValorINSS = mvardValorINSS
End Property

Public Property Let dValorCSLL(ByVal vData As Double)
    mvardValorCSLL = vData
End Property

Public Property Get dValorCSLL() As Double
    dValorCSLL = mvardValorCSLL
End Property

Public Property Let dValorCofins(ByVal vData As Double)
    mvardValorCofins = vData
End Property

Public Property Get dValorCofins() As Double
    dValorCofins = mvardValorCofins
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDocCab(ByVal vData As Long)
    mvarlNumIntDocCab = vData
End Property

Public Property Get lNumIntDocCab() As Long
    lNumIntDocCab = mvarlNumIntDocCab
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocNF(ByVal vData As Long)
    mvarlNumIntDocNF = vData
End Property

Public Property Get lNumIntDocNF() As Long
    lNumIntDocNF = mvarlNumIntDocNF
End Property

Public Property Let lNumNFe(ByVal vData As Long)
    mvarlNumNFe = vData
End Property

Public Property Get lNumNFe() As Long
    lNumNFe = mvarlNumNFe
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dHora(ByVal vData As Double)
    mvardHora = vData
End Property

Public Property Get dHora() As Double
    dHora = mvardHora
End Property

Public Property Let sCodVerificacao(ByVal vData As String)
    mvarsCodVerificacao = vData
End Property

Public Property Get sCodVerificacao() As String
    sCodVerificacao = mvarsCodVerificacao
End Property

Public Property Let sTipoRPS(ByVal vData As String)
    mvarsTipoRPS = vData
End Property

Public Property Get sTipoRPS() As String
    sTipoRPS = mvarsTipoRPS
End Property

Public Property Let sSerieRPS(ByVal vData As String)
    mvarsSerieRPS = vData
End Property

Public Property Get sSerieRPS() As String
    sSerieRPS = mvarsSerieRPS
End Property

Public Property Let lNumeroRPS(ByVal vData As Long)
    mvarlNumeroRPS = vData
End Property

Public Property Get lNumeroRPS() As Long
    lNumeroRPS = mvarlNumeroRPS
End Property

Public Property Let dtDataEmissaoRPS(ByVal vData As Date)
    mvardtDataEmissaoRPS = vData
End Property

Public Property Get dtDataEmissaoRPS() As Date
    dtDataEmissaoRPS = mvardtDataEmissaoRPS
End Property

Public Property Let sInscricaoMunicialPrest(ByVal vData As String)
    mvarsInscricaoMunicialPrest = vData
End Property

Public Property Get sInscricaoMunicialPrest() As String
    sInscricaoMunicialPrest = mvarsInscricaoMunicialPrest
End Property

Public Property Let iTipoCGCPrest(ByVal vData As Integer)
    mvariTipoCGCPrest = vData
End Property

Public Property Get iTipoCGCPrest() As Integer
    iTipoCGCPrest = mvariTipoCGCPrest
End Property

Public Property Let sCGCPrest(ByVal vData As String)
    mvarsCGCPrest = vData
End Property

Public Property Get sCGCPrest() As String
    sCGCPrest = mvarsCGCPrest
End Property

Public Property Let sRazaoSocialPrest(ByVal vData As String)
    mvarsRazaoSocialPrest = vData
End Property

Public Property Get sRazaoSocialPrest() As String
    sRazaoSocialPrest = mvarsRazaoSocialPrest
End Property

Public Property Let sTipoEnderecoPrest(ByVal vData As String)
    mvarsTipoEnderecoPrest = vData
End Property

Public Property Get sTipoEnderecoPrest() As String
    sTipoEnderecoPrest = mvarsTipoEnderecoPrest
End Property

Public Property Let sEnderecoPrest(ByVal vData As String)
    mvarsEnderecoPrest = vData
End Property

Public Property Get sEnderecoPrest() As String
    sEnderecoPrest = mvarsEnderecoPrest
End Property

Public Property Let sEndNumeroPrest(ByVal vData As String)
    mvarsEndNumeroPrest = vData
End Property

Public Property Get sEndNumeroPrest() As String
    sEndNumeroPrest = mvarsEndNumeroPrest
End Property

Public Property Let sEndComplementoPrest(ByVal vData As String)
    mvarsEndComplementoPrest = vData
End Property

Public Property Get sEndComplementoPrest() As String
    sEndComplementoPrest = mvarsEndComplementoPrest
End Property

Public Property Let sBairroPrest(ByVal vData As String)
    mvarsBairroPrest = vData
End Property

Public Property Get sBairroPrest() As String
    sBairroPrest = mvarsBairroPrest
End Property

Public Property Let sCidadePrest(ByVal vData As String)
    mvarsCidadePrest = vData
End Property

Public Property Get sCidadePrest() As String
    sCidadePrest = mvarsCidadePrest
End Property

Public Property Let sUFPrest(ByVal vData As String)
    mvarsUFPrest = vData
End Property

Public Property Get sUFPrest() As String
    sUFPrest = mvarsUFPrest
End Property

Public Property Let sCEPPrest(ByVal vData As String)
    mvarsCEPPrest = vData
End Property

Public Property Get sCEPPrest() As String
    sCEPPrest = mvarsCEPPrest
End Property

Public Property Let sEmailPrest(ByVal vData As String)
    mvarsEmailPrest = vData
End Property

Public Property Get sEmailPrest() As String
    sEmailPrest = mvarsEmailPrest
End Property

Public Property Let iOPTSimples(ByVal vData As Integer)
    mvariOPTSimples = vData
End Property

Public Property Get iOPTSimples() As Integer
    iOPTSimples = mvariOPTSimples
End Property

Public Property Let sSituacaoNF(ByVal vData As String)
    mvarsSituacaoNF = vData
End Property

Public Property Get sSituacaoNF() As String
    sSituacaoNF = mvarsSituacaoNF
End Property

Public Property Let dtDataCancelamento(ByVal vData As Date)
    mvardtDataCancelamento = vData
End Property

Public Property Get dtDataCancelamento() As Date
    dtDataCancelamento = mvardtDataCancelamento
End Property

Public Property Let sNumGuia(ByVal vData As String)
    mvarsNumGuia = vData
End Property

Public Property Get sNumGuia() As String
    sNumGuia = mvarsNumGuia
End Property

Public Property Let dtDataQuitacaoGuia(ByVal vData As Date)
    mvardtDataQuitacaoGuia = vData
End Property

Public Property Get dtDataQuitacaoGuia() As Date
    dtDataQuitacaoGuia = mvardtDataQuitacaoGuia
End Property

Public Property Let dValorServicos(ByVal vData As Double)
    mvardValorServicos = vData
End Property

Public Property Get dValorServicos() As Double
    dValorServicos = mvardValorServicos
End Property

Public Property Let dValorDeducoes(ByVal vData As Double)
    mvardValorDeducoes = vData
End Property

Public Property Get dValorDeducoes() As Double
    dValorDeducoes = mvardValorDeducoes
End Property

Public Property Let lCodServico(ByVal vData As Long)
    mvarlCodServico = vData
End Property

Public Property Get lCodServico() As Long
    lCodServico = mvarlCodServico
End Property

Public Property Let dAliquota(ByVal vData As Double)
    mvardAliquota = vData
End Property

Public Property Get dAliquota() As Double
    dAliquota = mvardAliquota
End Property

Public Property Let dValorISS(ByVal vData As Double)
    mvardValorISS = vData
End Property

Public Property Get dValorISS() As Double
    dValorISS = mvardValorISS
End Property

Public Property Let dValorCredito(ByVal vData As Double)
    mvardValorCredito = vData
End Property

Public Property Get dValorCredito() As Double
    dValorCredito = mvardValorCredito
End Property

Public Property Let sISSRetido(ByVal vData As String)
    mvarsISSRetido = vData
End Property

Public Property Get sISSRetido() As String
    sISSRetido = mvarsISSRetido
End Property

Public Property Let iTipoCGCTom(ByVal vData As Integer)
    mvariTipoCGCTom = vData
End Property

Public Property Get iTipoCGCTom() As Integer
    iTipoCGCTom = mvariTipoCGCTom
End Property

Public Property Let sCGCTom(ByVal vData As String)
    mvarsCGCTom = vData
End Property

Public Property Get sCGCTom() As String
    sCGCTom = mvarsCGCTom
End Property

Public Property Let sInscricaoMunicipalTom(ByVal vData As String)
    mvarsInscricaoMunicipalTom = vData
End Property

Public Property Get sInscricaoMunicipalTom() As String
    sInscricaoMunicipalTom = mvarsInscricaoMunicipalTom
End Property

Public Property Let sInscricaoEstadualTom(ByVal vData As String)
    mvarsInscricaoEstadualTom = vData
End Property

Public Property Get sInscricaoEstadualTom() As String
    sInscricaoEstadualTom = mvarsInscricaoEstadualTom
End Property

Public Property Let sRazaoSocialTom(ByVal vData As String)
    mvarsRazaoSocialTom = vData
End Property

Public Property Get sRazaoSocialTom() As String
    sRazaoSocialTom = mvarsRazaoSocialTom
End Property

Public Property Let sTipoEnderecoTom(ByVal vData As String)
    mvarsTipoEnderecoTom = vData
End Property

Public Property Get sTipoEnderecoTom() As String
    sTipoEnderecoTom = mvarsTipoEnderecoTom
End Property

Public Property Let sEnderecoTom(ByVal vData As String)
    mvarsEnderecoTom = vData
End Property

Public Property Get sEnderecoTom() As String
    sEnderecoTom = mvarsEnderecoTom
End Property

Public Property Let sEndNumeroTom(ByVal vData As String)
    mvarsEndNumeroTom = vData
End Property

Public Property Get sEndNumeroTom() As String
    sEndNumeroTom = mvarsEndNumeroTom
End Property

Public Property Let sEndComplementoTom(ByVal vData As String)
    mvarsEndComplementoTom = vData
End Property

Public Property Get sEndComplementoTom() As String
    sEndComplementoTom = mvarsEndComplementoTom
End Property

Public Property Let sBairroTom(ByVal vData As String)
    mvarsBairroTom = vData
End Property

Public Property Get sBairroTom() As String
    sBairroTom = mvarsBairroTom
End Property

Public Property Let sCidadeTom(ByVal vData As String)
    mvarsCidadeTom = vData
End Property

Public Property Get sCidadeTom() As String
    sCidadeTom = mvarsCidadeTom
End Property

Public Property Let sUFTom(ByVal vData As String)
    mvarsUFTom = vData
End Property

Public Property Get sUFTom() As String
    sUFTom = mvarsUFTom
End Property

Public Property Let sCEPTom(ByVal vData As String)
    mvarsCEPTom = vData
End Property

Public Property Get sCEPTom() As String
    sCEPTom = mvarsCEPTom
End Property

Public Property Let sEmailTom(ByVal vData As String)
    mvarsEmailTom = vData
End Property

Public Property Get sEmailTom() As String
    sEmailTom = mvarsEmailTom
End Property

Public Property Let sDiscriminacao(ByVal vData As String)
    mvarsDiscriminacao = vData
End Property

Public Property Get sDiscriminacao() As String
    sDiscriminacao = mvarsDiscriminacao
End Property
