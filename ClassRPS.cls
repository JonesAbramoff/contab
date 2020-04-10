VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

Public gobjNF As ClassNFiscal
Public gobjTitRec As ClassTituloReceber
Public gobjCli As ClassCliente
Public gobjEnd As ClassEndereco
Public gobjFilCli As ClassFilialCliente

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarsSerie As String
Private mvarsTipo As String
Private mvarlNumero As Long
Private mvardtDataEmissao As Date
Private mvarsSituacao As String
Private mvardValorServico As Double
Private mvardValorDeducao As Double
Private mvarlCodigoServico As Long
Private mvardAliquota As Double
Private mvariISSRetido As Integer
Private mvariTipoCGC As Integer
Private mvarsCGC As String
Private mvarsInscricaoMunicipal As String
Private mvarsInscricaoEstadual As String
Private mvarsRazaoSocial As String
Private mvarsEndereco As String
Private mvarsEndNumero As String
Private mvarsEndComplemento As String
Private mvarsBairro As String
Private mvarsCidade As String
Private mvarsUF As String
Private mvarsCEP As String
Private mvarsEmail As String
Private mvarsDiscriminacao As String
Private mvariFilialEmpresa As Integer
Private mvarlNumIntDocCab As Long
Private mvarlNumIntDocNF As Long
Private mvarlCliente As Long
Private mvariFilialCliente As Integer
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

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Let lNumIntDocNF(ByVal vData As Long)
    mvarlNumIntDocNF = vData
End Property

Public Property Get lNumIntDocNF() As Long
    lNumIntDocNF = mvarlNumIntDocNF
End Property

Public Property Let lNumIntDocCab(ByVal vData As Long)
    mvarlNumIntDocCab = vData
End Property

Public Property Get lNumIntDocCab() As Long
    lNumIntDocCab = mvarlNumIntDocCab
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let sTipo(ByVal vData As String)
    mvarsTipo = vData
End Property

Public Property Get sTipo() As String
    sTipo = mvarsTipo
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let sSituacao(ByVal vData As String)
    mvarsSituacao = vData
End Property

Public Property Get sSituacao() As String
    sSituacao = mvarsSituacao
End Property

Public Property Let dValorServico(ByVal vData As Double)
    mvardValorServico = vData
End Property

Public Property Get dValorServico() As Double
    dValorServico = mvardValorServico
End Property

Public Property Let dValorDeducao(ByVal vData As Double)
    mvardValorDeducao = vData
End Property

Public Property Get dValorDeducao() As Double
    dValorDeducao = mvardValorDeducao
End Property

Public Property Let lCodigoServico(ByVal vData As Long)
    mvarlCodigoServico = vData
End Property

Public Property Get lCodigoServico() As Long
    lCodigoServico = mvarlCodigoServico
End Property

Public Property Let dAliquota(ByVal vData As Double)
    mvardAliquota = vData
End Property

Public Property Get dAliquota() As Double
    dAliquota = mvardAliquota
End Property

Public Property Let iISSRetido(ByVal vData As Integer)
    mvariISSRetido = vData
End Property

Public Property Get iISSRetido() As Integer
    iISSRetido = mvariISSRetido
End Property

Public Property Let iTipoCGC(ByVal vData As Integer)
    mvariTipoCGC = vData
End Property

Public Property Get iTipoCGC() As Integer
    iTipoCGC = mvariTipoCGC
End Property

Public Property Let sCgc(ByVal vData As String)
    mvarsCGC = vData
End Property

Public Property Get sCgc() As String
    sCgc = mvarsCGC
End Property

Public Property Let sInscricaoMunicipal(ByVal vData As String)
    mvarsInscricaoMunicipal = vData
End Property

Public Property Get sInscricaoMunicipal() As String
    sInscricaoMunicipal = mvarsInscricaoMunicipal
End Property

Public Property Let sInscricaoEstadual(ByVal vData As String)
    mvarsInscricaoEstadual = vData
End Property

Public Property Get sInscricaoEstadual() As String
    sInscricaoEstadual = mvarsInscricaoEstadual
End Property

Public Property Let sRazaoSocial(ByVal vData As String)
    mvarsRazaoSocial = vData
End Property

Public Property Get sRazaoSocial() As String
    sRazaoSocial = mvarsRazaoSocial
End Property

Public Property Let sEndereco(ByVal vData As String)
    mvarsEndereco = vData
End Property

Public Property Get sEndereco() As String
    sEndereco = mvarsEndereco
End Property

Public Property Let sEndNumero(ByVal vData As String)
    mvarsEndNumero = vData
End Property

Public Property Get sEndNumero() As String
    sEndNumero = mvarsEndNumero
End Property

Public Property Let sEndComplemento(ByVal vData As String)
    mvarsEndComplemento = vData
End Property

Public Property Get sEndComplemento() As String
    sEndComplemento = mvarsEndComplemento
End Property

Public Property Let sBairro(ByVal vData As String)
    mvarsBairro = vData
End Property

Public Property Get sBairro() As String
    sBairro = mvarsBairro
End Property

Public Property Let sCidade(ByVal vData As String)
    mvarsCidade = vData
End Property

Public Property Get sCidade() As String
    sCidade = mvarsCidade
End Property

Public Property Let sUF(ByVal vData As String)
    mvarsUF = vData
End Property

Public Property Get sUF() As String
    sUF = mvarsUF
End Property

Public Property Let sCEP(ByVal vData As String)
    mvarsCEP = vData
End Property

Public Property Get sCEP() As String
    sCEP = mvarsCEP
End Property

Public Property Let sEmail(ByVal vData As String)
    mvarsEmail = vData
End Property

Public Property Get sEmail() As String
    sEmail = mvarsEmail
End Property

Public Property Let sDiscriminacao(ByVal vData As String)
    mvarsDiscriminacao = vData
End Property

Public Property Get sDiscriminacao() As String
    sDiscriminacao = mvarsDiscriminacao
End Property
