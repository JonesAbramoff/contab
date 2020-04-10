VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmFiliais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'??? atencao: se incluir algum novo campo alterar o metodo Copiar
Private mvariCodFilial As Integer
Private mvarlCodEmpresa As Long
Private mvarsNome As String
Private mvarsCGC As String
Private mvariAlmoxarifadoPadrao As Integer
Private mvarsInscricaoEstadual As String
Private mvariICMSPorEstimativa As Integer
Private mvarsInscricaoMunicipal As String
Private mvardISSPercPadrao As Double
Private mvarsISSCodigoPadrao As String
Private mvariISSIncluso As Integer
Private mvariContribuinteIPI As Integer
Private mvarobjEndereco As New ClassEndereco
Private mvardIRPercPadrao As Double
Private mvarsRamoEmpresa As String
Private mvarsJucerja As String
Private mvardtDataJucerja As Date
Private mvarsContador As String
Private mvarsCRCContador As String
Private mvarsCPFContador As String
Private mvarobjEnderecoEntrega As New ClassEndereco
Private mvarcolModulos As New Collection
Private mvariInativa As Integer
Private mvariTipoTribAtividade As Integer
Private mvariSimplesFederal As Integer
Private mvardSimplesFederalAliq As Double
Private mvardSimplesFederalTeto As Double
Private mvariSuperSimples As Integer
Private mvariPISNaoCumulativo As Integer
Private mvariCOFINSNaoCumulativo As Integer
Private mvariLucroPresumido As Integer
Private mvarsCertificadoA1A3 As String
Private mvariNFeAmbiente As Integer
Private mvarsCNAE As String
Private mvarsSignatarioCTB As String
Private mvarsCodQualiSigCTB As String
Private mvarsCPFSignatarioCTB As String
Private mvarcolAdmCad As New Collection
Private mvarcolContrib As New Collection
Private mvarsSpedFiscalPerfil As String
Private mvarsNomeReduzido As String
Private mvarobjEnderecoContador As New ClassEndereco
Private mvarsCNPJContador As String
Private mvariRPSAmbiente As Integer
Private mvariPisCofinsRegCumTipo As Integer
Private mvariContribPrevPorRecBruta As Integer
Private mvardAliquotaPadraoContribPrev As Double
Private mvariPisCofinsIndAproCred As Integer

'nfe 3.10
Public sNFCECSC As String
Public sidNFCECSC As String
Public iindSincPadrao As Integer
Public iRegimeTrib As Integer
Public iRegimeEspecialTrib As Integer

'??? atencao: se incluir algum novo campo alterar o metodo Copiar
'??? atencao: se incluir algum novo campo alterar o metodo Copiar
'??? atencao: se incluir algum novo campo alterar o metodo Copiar
'??? atencao: se incluir algum novo campo alterar o metodo Copiar

Public Property Let iPisCofinsIndAproCred(ByVal vData As Integer)
    mvariPisCofinsIndAproCred = vData
End Property

Public Property Get iPisCofinsIndAproCred() As Integer
    iPisCofinsIndAproCred = mvariPisCofinsIndAproCred
End Property

Public Property Let dAliquotaPadraoContribPrev(ByVal vData As Double)
    mvardAliquotaPadraoContribPrev = vData
End Property

Public Property Get dAliquotaPadraoContribPrev() As Double
    dAliquotaPadraoContribPrev = mvardAliquotaPadraoContribPrev
End Property

Public Property Let iContribPrevPorRecBruta(ByVal vData As Integer)
    mvariContribPrevPorRecBruta = vData
End Property

Public Property Get iContribPrevPorRecBruta() As Integer
    iContribPrevPorRecBruta = mvariContribPrevPorRecBruta
End Property

Public Property Let iPisCofinsRegCumTipo(ByVal vData As Integer)
    mvariPisCofinsRegCumTipo = vData
End Property

Public Property Get iPisCofinsRegCumTipo() As Integer
    iPisCofinsRegCumTipo = mvariPisCofinsRegCumTipo
End Property

Public Property Let iRPSAmbiente(ByVal vData As Integer)
    mvariRPSAmbiente = vData
End Property

Public Property Get iRPSAmbiente() As Integer
    iRPSAmbiente = mvariRPSAmbiente
End Property

Public Property Let sCNPJContador(ByVal vData As String)
    mvarsCNPJContador = vData
End Property

Public Property Get sCNPJContador() As String
    sCNPJContador = mvarsCNPJContador
End Property

Public Property Set objEnderecoContador(ByVal vData As ClassEndereco)
    Set mvarobjEnderecoContador = vData
End Property

Public Property Get objEnderecoContador() As ClassEndereco
    Set objEnderecoContador = mvarobjEnderecoContador
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let sSpedFiscalPerfil(ByVal vData As String)
    mvarsSpedFiscalPerfil = vData
End Property

Public Property Get sSpedFiscalPerfil() As String
    sSpedFiscalPerfil = mvarsSpedFiscalPerfil
End Property

Public Property Set colContrib(ByVal vData As Object)
    Set mvarcolContrib = vData
End Property

Public Property Get colContrib() As Collection
    Set colContrib = mvarcolContrib
End Property

Public Property Set colAdmCad(ByVal vData As Object)
    Set mvarcolAdmCad = vData
End Property

Public Property Get colAdmCad() As Collection
    Set colAdmCad = mvarcolAdmCad
End Property

Public Property Let sCPFSignatarioCTB(ByVal vData As String)
    mvarsCPFSignatarioCTB = vData
End Property

Public Property Get sCPFSignatarioCTB() As String
    sCPFSignatarioCTB = mvarsCPFSignatarioCTB
End Property

Public Property Let sCodQualiSigCTB(ByVal vData As String)
    mvarsCodQualiSigCTB = vData
End Property

Public Property Get sCodQualiSigCTB() As String
    sCodQualiSigCTB = mvarsCodQualiSigCTB
End Property

Public Property Let sSignatarioCTB(ByVal vData As String)
    mvarsSignatarioCTB = vData
End Property

Public Property Get sSignatarioCTB() As String
    sSignatarioCTB = mvarsSignatarioCTB
End Property

Public Property Let sCNAE(ByVal vData As String)
    mvarsCNAE = vData
End Property

Public Property Get sCNAE() As String
    sCNAE = mvarsCNAE
End Property

Public Property Let iNFeAmbiente(ByVal vData As Integer)
    mvariNFeAmbiente = vData
End Property

Public Property Get iNFeAmbiente() As Integer
    iNFeAmbiente = mvariNFeAmbiente
End Property

'??? ATENCAO: VEJA O METODO COPIAR DESTA CLASSE SE FOR INCLUIR MAIS ATRIBUTOS

Public Property Let sCertificadoA1A3(ByVal vData As String)
    mvarsCertificadoA1A3 = vData
End Property

Public Property Get sCertificadoA1A3() As String
    sCertificadoA1A3 = mvarsCertificadoA1A3
End Property


Public Property Let iSuperSimples(ByVal vData As Integer)
    mvariSuperSimples = vData
End Property

Public Property Get iSuperSimples() As Integer
    iSuperSimples = mvariSuperSimples
End Property

Public Property Let iPISNaoCumulativo(ByVal vData As Integer)
    mvariPISNaoCumulativo = vData
End Property

Public Property Get iPISNaoCumulativo() As Integer
    iPISNaoCumulativo = mvariPISNaoCumulativo
End Property

Public Property Let iLucroPresumido(ByVal vData As Integer)
    mvariLucroPresumido = vData
End Property

Public Property Get iLucroPresumido() As Integer
    iLucroPresumido = mvariLucroPresumido
End Property

Public Property Let iCOFINSNaoCumulativo(ByVal vData As Integer)
    mvariCOFINSNaoCumulativo = vData
End Property

Public Property Get iCOFINSNaoCumulativo() As Integer
    iCOFINSNaoCumulativo = mvariCOFINSNaoCumulativo
End Property

Public Property Let dSimplesFederalAliq(ByVal vData As Double)
    mvardSimplesFederalAliq = vData
End Property

Public Property Get dSimplesFederalAliq() As Double
    dSimplesFederalAliq = mvardSimplesFederalAliq
End Property

Public Property Let dSimplesFederalTeto(ByVal vData As Double)
    mvardSimplesFederalTeto = vData
End Property

Public Property Get dSimplesFederalTeto() As Double
    dSimplesFederalTeto = mvardSimplesFederalTeto
End Property

Public Property Let iSimplesFederal(ByVal vData As Integer)
    mvariSimplesFederal = vData
End Property

Public Property Get iSimplesFederal() As Integer
    iSimplesFederal = mvariSimplesFederal
End Property

Public Property Let iTipoTribAtividade(ByVal vData As Integer)
    mvariTipoTribAtividade = vData
End Property

Public Property Get iTipoTribAtividade() As Integer
    iTipoTribAtividade = mvariTipoTribAtividade
End Property

Public Property Let iInativa(ByVal vData As Integer)
    mvariInativa = vData
End Property

Public Property Get iInativa() As Integer
    iInativa = mvariInativa
End Property

Public Property Set colModulos(ByVal vData As Object)
    Set mvarcolModulos = vData
End Property

Public Property Get colModulos() As Collection
    Set colModulos = mvarcolModulos
End Property

Public Property Set objEnderecoEntrega(ByVal vData As ClassEndereco)
    Set mvarobjEnderecoEntrega = vData
End Property

Public Property Get objEnderecoEntrega() As ClassEndereco
    Set objEnderecoEntrega = mvarobjEnderecoEntrega
End Property

Public Property Let sCPFContador(ByVal vData As String)
    mvarsCPFContador = vData
End Property

Public Property Get sCPFContador() As String
    sCPFContador = mvarsCPFContador
End Property

Public Property Let sCRCContador(ByVal vData As String)
    mvarsCRCContador = vData
End Property

Public Property Get sCRCContador() As String
    sCRCContador = mvarsCRCContador
End Property

Public Property Let sContador(ByVal vData As String)
    mvarsContador = vData
End Property

Public Property Get sContador() As String
    sContador = mvarsContador
End Property

Public Property Let dtDataJucerja(ByVal vData As Date)
    mvardtDataJucerja = vData
End Property

Public Property Get dtDataJucerja() As Date
    dtDataJucerja = mvardtDataJucerja
End Property

Public Property Let sJucerja(ByVal vData As String)
    mvarsJucerja = vData
End Property

Public Property Get sJucerja() As String
    sJucerja = mvarsJucerja
End Property

Public Property Let sRamoEmpresa(ByVal vData As String)
    mvarsRamoEmpresa = vData
End Property

Public Property Get sRamoEmpresa() As String
    sRamoEmpresa = mvarsRamoEmpresa
End Property

Public Property Let dIRPercPadrao(ByVal vData As Double)
    mvardIRPercPadrao = vData
End Property

Public Property Get dIRPercPadrao() As Double
    dIRPercPadrao = mvardIRPercPadrao
End Property

Public Property Set objEndereco(ByVal vData As ClassEndereco)
    Set mvarobjEndereco = vData
End Property

Public Property Get objEndereco() As ClassEndereco
    Set objEndereco = mvarobjEndereco
End Property

Public Property Let iContribuinteIPI(ByVal vData As Integer)
    mvariContribuinteIPI = vData
End Property

Public Property Get iContribuinteIPI() As Integer
    iContribuinteIPI = mvariContribuinteIPI
End Property

Public Property Let iISSIncluso(ByVal vData As Integer)
    mvariISSIncluso = vData
End Property

Public Property Get iISSIncluso() As Integer
    iISSIncluso = mvariISSIncluso
End Property

Public Property Let sISSCodigoPadrao(ByVal vData As String)
    mvarsISSCodigoPadrao = vData
End Property

Public Property Get sISSCodigoPadrao() As String
    sISSCodigoPadrao = mvarsISSCodigoPadrao
End Property

Public Property Let dISSPercPadrao(ByVal vData As Double)
    mvardISSPercPadrao = vData
End Property

Public Property Get dISSPercPadrao() As Double
    dISSPercPadrao = mvardISSPercPadrao
End Property

Public Property Let sInscricaoMunicipal(ByVal vData As String)
    mvarsInscricaoMunicipal = vData
End Property

Public Property Get sInscricaoMunicipal() As String
    sInscricaoMunicipal = mvarsInscricaoMunicipal
End Property

Public Property Let iICMSPorEstimativa(ByVal vData As Integer)
    mvariICMSPorEstimativa = vData
End Property

Public Property Get iICMSPorEstimativa() As Integer
    iICMSPorEstimativa = mvariICMSPorEstimativa
End Property

Public Property Let sInscricaoEstadual(ByVal vData As String)
    mvarsInscricaoEstadual = vData
End Property

Public Property Get sInscricaoEstadual() As String
    sInscricaoEstadual = mvarsInscricaoEstadual
End Property

Public Property Let iAlmoxarifadoPadrao(ByVal vData As Integer)
    mvariAlmoxarifadoPadrao = vData
End Property

Public Property Get iAlmoxarifadoPadrao() As Integer
    iAlmoxarifadoPadrao = mvariAlmoxarifadoPadrao
End Property

Public Property Let sCgc(ByVal vData As String)
    mvarsCGC = vData
End Property

Public Property Get sCgc() As String
    sCgc = mvarsCGC
End Property

Public Property Let sNome(ByVal vData As String)
Attribute sNome.VB_Description = "texto que identificará a empresa numa lista"
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let lCodEmpresa(ByVal vData As Long)
    mvarlCodEmpresa = vData
End Property

Public Property Get lCodEmpresa() As Long
    lCodEmpresa = mvarlCodEmpresa
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
    mvariCodFilial = vData
End Property

Public Property Get iCodFilial() As Integer
    iCodFilial = mvariCodFilial
End Property

Public Function Ativo(sSigla As String) As Long
'recebe a sigla do modulo e retorna a indicacao se o modulo está ativo ou inativo

Dim vSiglaModulo As Variant
Dim iAchou As Integer
    
    For Each vSiglaModulo In mvarcolModulos
    
        If vSiglaModulo = sSigla Then
            Ativo = MODULO_ATIVO
            iAchou = 1
            Exit For
        End If
    Next
    
    If iAchou = 0 Then
        Ativo = -1
    End If
    
End Function

Public Sub Copiar(objFilialOrig As AdmFiliais)
'atualiza o proprio conteudo copiando dados de outra instancia

Dim objFilialAdmCadOrig As AdmFiliaisAdmCad
Dim objFilialAdmCad As AdmFiliaisAdmCad
Dim objFilialContribOrig As AdmFiliaisContrib
Dim objFilialContrib As AdmFiliaisContrib

    mvardISSPercPadrao = objFilialOrig.dISSPercPadrao
    mvariAlmoxarifadoPadrao = objFilialOrig.iAlmoxarifadoPadrao
    mvariCodFilial = objFilialOrig.iCodFilial
    mvariContribuinteIPI = objFilialOrig.iContribuinteIPI
    mvariICMSPorEstimativa = objFilialOrig.iICMSPorEstimativa
    mvariISSIncluso = objFilialOrig.iISSIncluso
    mvarlCodEmpresa = objFilialOrig.lCodEmpresa
    mvarsCGC = objFilialOrig.sCgc
    mvarsInscricaoEstadual = objFilialOrig.sInscricaoEstadual
    mvarsInscricaoMunicipal = objFilialOrig.sInscricaoMunicipal
    mvarsISSCodigoPadrao = objFilialOrig.sISSCodigoPadrao
    mvarsNome = objFilialOrig.sNome
    mvardIRPercPadrao = objFilialOrig.dIRPercPadrao
    mvarsRamoEmpresa = objFilialOrig.sRamoEmpresa
    mvarsJucerja = objFilialOrig.sJucerja
    mvardtDataJucerja = objFilialOrig.dtDataJucerja
    mvarsContador = objFilialOrig.sContador
    mvarsCRCContador = objFilialOrig.sCRCContador
    mvarsCPFContador = objFilialOrig.sCPFContador
    mvariTipoTribAtividade = objFilialOrig.iTipoTribAtividade
    mvariSimplesFederal = objFilialOrig.iSimplesFederal
    mvardSimplesFederalAliq = objFilialOrig.dSimplesFederalAliq
    mvardSimplesFederalTeto = objFilialOrig.dSimplesFederalTeto
    mvariSuperSimples = objFilialOrig.iSuperSimples
    mvariPISNaoCumulativo = objFilialOrig.iPISNaoCumulativo
    mvariCOFINSNaoCumulativo = objFilialOrig.iCOFINSNaoCumulativo
    mvariLucroPresumido = objFilialOrig.iLucroPresumido
    mvarsCertificadoA1A3 = objFilialOrig.sCertificadoA1A3
    mvarsCNAE = objFilialOrig.sCNAE
    mvarsSignatarioCTB = objFilialOrig.sSignatarioCTB
    mvarsCodQualiSigCTB = objFilialOrig.sCodQualiSigCTB
    mvarsCPFSignatarioCTB = objFilialOrig.sCPFSignatarioCTB
    mvariNFeAmbiente = objFilialOrig.iNFeAmbiente
    mvariRPSAmbiente = objFilialOrig.iRPSAmbiente
    mvarsSpedFiscalPerfil = objFilialOrig.sSpedFiscalPerfil
    mvarsNomeReduzido = objFilialOrig.sNomeReduzido
    mvarsCNPJContador = objFilialOrig.sCNPJContador
    mvariPisCofinsRegCumTipo = objFilialOrig.iPisCofinsRegCumTipo
    mvariContribPrevPorRecBruta = objFilialOrig.iContribPrevPorRecBruta
    mvardAliquotaPadraoContribPrev = objFilialOrig.dAliquotaPadraoContribPrev
    mvariPisCofinsIndAproCred = objFilialOrig.iPisCofinsIndAproCred
    
    'nfe 3.10
    sNFCECSC = objFilialOrig.sNFCECSC
    sidNFCECSC = objFilialOrig.sidNFCECSC
    iindSincPadrao = objFilialOrig.iindSincPadrao
    iRegimeTrib = objFilialOrig.iRegimeTrib
    iRegimeEspecialTrib = objFilialOrig.iRegimeEspecialTrib
    
    Call mvarobjEndereco.Copiar(objFilialOrig.objEndereco)
    Call mvarobjEnderecoEntrega.Copiar(objFilialOrig.objEnderecoEntrega)
    Call mvarobjEnderecoContador.Copiar(objFilialOrig.objEnderecoContador)
    
    Set mvarcolAdmCad = New Collection
    For Each objFilialAdmCadOrig In objFilialOrig.colAdmCad
        Set objFilialAdmCad = New AdmFiliaisAdmCad
        Call objFilialAdmCad.Copiar(objFilialAdmCadOrig)
        mvarcolAdmCad.Add objFilialAdmCad
    Next
    
    Set mvarcolContrib = New Collection
    For Each objFilialContribOrig In objFilialOrig.colContrib
        Set objFilialContrib = New AdmFiliaisContrib
        Call objFilialContrib.Copiar(objFilialContribOrig)
        mvarcolContrib.Add objFilialContrib
    Next
    
End Sub
