VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassClienteHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvardtDataAtualizacao As Date
Private mvarlCodigo As Long
Private mvariFilialCliente As Integer
Private mvarsCGC As String
Private mvarsInscricaoEstadual As String
Private mvarsInscricaoMunicipal As String
Private mvarsRazaoSocial As String
Private mvarsEndereco As String
Private mvarsBairro As String
Private mvarsCidade As String
Private mvarsUF As String
Private mvarsPais As String
Private mvarsCEP As String
Private mvariAtivo As Integer
Private mvarlCodigoLoja As Long
Private mvariFilialEmpresaLoja As Integer
Private mvariCodFilialLoja As Integer
Private mvarsInscricaoSuframa As String
Public sNomeReduzido As String
Public iTipo As Integer
Public dLimiteCredito As Double
Public sObservacao As String
Public iCondicaoPagto As Integer
Public iTabelaPreco As Integer
Public sUsuarioCobrador As String
Public sUsuRespCallCenter As String
Public iCodTransportadora As Integer
Public iVendedor As Integer
Public sRG As String
Public iRegimeTributario As Integer
Public sTelefone1 As String
Public sTelefone2 As String
Public sEmail As String
Public sEmail2 As String
Public sMensagemNF As String
Public iRegiao As Integer
Public dDesconto As Double
Public dComissaoVendas As Double
Public iVersaoControle As Integer
Public sUsuario As String
Public dtDataReg As Date
Public dHoraReg As Double


Public Property Let sInscricaoSuframa(ByVal vData As String)
    mvarsInscricaoSuframa = vData
End Property

Public Property Get sInscricaoSuframa() As String
    sInscricaoSuframa = mvarsInscricaoSuframa
End Property

Public Property Let iCodFilialLoja(ByVal vData As Integer)
    mvariCodFilialLoja = vData
End Property

Public Property Get iCodFilialLoja() As Integer
    iCodFilialLoja = mvariCodFilialLoja
End Property

Public Property Let iFilialEmpresaLoja(ByVal vData As Integer)
    mvariFilialEmpresaLoja = vData
End Property

Public Property Get iFilialEmpresaLoja() As Integer
    iFilialEmpresaLoja = mvariFilialEmpresaLoja
End Property

Public Property Let lCodigoLoja(ByVal vData As Long)
    mvarlCodigoLoja = vData
End Property

Public Property Get lCodigoLoja() As Long
    lCodigoLoja = mvarlCodigoLoja
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let sCEP(ByVal vData As String)
    mvarsCEP = vData
End Property

Public Property Get sCEP() As String
    sCEP = mvarsCEP
End Property

Public Property Let sPais(ByVal vData As String)
    mvarsPais = vData
End Property

Public Property Get sPais() As String
    sPais = mvarsPais
End Property

Public Property Let sUF(ByVal vData As String)
    mvarsUF = vData
End Property

Public Property Get sUF() As String
    sUF = mvarsUF
End Property

Public Property Let sCidade(ByVal vData As String)
    mvarsCidade = vData
End Property

Public Property Get sCidade() As String
    sCidade = mvarsCidade
End Property

Public Property Let sBairro(ByVal vData As String)
    mvarsBairro = vData
End Property

Public Property Get sBairro() As String
    sBairro = mvarsBairro
End Property

Public Property Let sEndereco(ByVal vData As String)
    mvarsEndereco = vData
End Property

Public Property Get sEndereco() As String
    sEndereco = mvarsEndereco
End Property

Public Property Let sRazaoSocial(ByVal vData As String)
    mvarsRazaoSocial = vData
End Property

Public Property Get sRazaoSocial() As String
    sRazaoSocial = mvarsRazaoSocial
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

Public Property Let sCgc(ByVal vData As String)
    mvarsCGC = vData
End Property

Public Property Get sCgc() As String
    sCgc = mvarsCGC
End Property

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let dtDataAtualizacao(ByVal vData As Date)
    mvardtDataAtualizacao = vData
End Property

Public Property Get dtDataAtualizacao() As Date
    dtDataAtualizacao = mvardtDataAtualizacao
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

