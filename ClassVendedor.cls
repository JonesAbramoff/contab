VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsNome As String
Private mvarsNomeReduzido As String
Private mvarlEndereco As Long
Private mvariTipo As Integer
Private mvarsMatricula As String
Private mvariCodRegiao As Integer
Private mvardSaldoComissao As Double
Private mvardPercComissao As Double
Private mvardPercComissaoBaixa As Double
Private mvardPercComissaoEmissao As Double
Private mvariComissaoSobreTotal As Integer
Private mvariComissaoFrete As Integer
Private mvariComissaoICM As Integer
Private mvariComissaoIPI As Integer
Private mvariBanco As Integer
Private mvarsAgencia As String
Private mvarsContaCorrente As String
Private mvariComissaoSeguro As Integer
Private mvardtDataUltimaVenda As Date 'local copy
Private mvarcolFiliaisLoja As New Collection 'local copy
Private mvariAtivo As Integer
Private mvariVinculo As Integer
Private mvarsRazaoSocial As String
Private mvarsCGC As String
Private mvarsInscricaoEstadual As String
Private mvarobjInfoUsu As Object
Private mvariCargo As Integer
Private mvariSuperior As Integer
Private mvarsCodUsuario As String
Private mvarsRG As String

Public Property Let sRG(ByVal vData As String)
    mvarsRG = vData
End Property

Public Property Get sRG() As String
    sRG = mvarsRG
End Property

Public Property Let sCodUsuario(ByVal vData As String)
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property

Public Property Set objInfoUsu(ByVal vData As Object)
    Set mvarobjInfoUsu = vData
End Property

Public Property Get objInfoUsu() As Object
    Set objInfoUsu = mvarobjInfoUsu
End Property

Public Property Let sRazaoSocial(ByVal vData As String)
    mvarsRazaoSocial = vData
End Property

Public Property Get sRazaoSocial() As String
    sRazaoSocial = mvarsRazaoSocial
End Property

Public Property Let sCgc(ByVal vData As String)
    mvarsCGC = vData
End Property

Public Property Get sCgc() As String
    sCgc = mvarsCGC
End Property

Public Property Let sInscricaoEstadual(ByVal vData As String)
    mvarsInscricaoEstadual = vData
End Property

Public Property Get sInscricaoEstadual() As String
    sInscricaoEstadual = mvarsInscricaoEstadual
End Property

Public Property Let iVinculo(ByVal vData As Integer)
    mvariVinculo = vData
End Property

Public Property Get iVinculo() As Integer
    iVinculo = mvariVinculo
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Set colFiliaisLoja(ByVal vData As Collection)
    Set mvarcolFiliaisLoja = vData
End Property

Public Property Get colFiliaisLoja() As Collection
Attribute colFiliaisLoja.VB_UserMemId = 0
    Set colFiliaisLoja = mvarcolFiliaisLoja
End Property

Public Property Let dtDataUltimaVenda(ByVal vData As Date)
    mvardtDataUltimaVenda = vData
End Property

Public Property Get dtDataUltimaVenda() As Date
    dtDataUltimaVenda = mvardtDataUltimaVenda
End Property

Public Property Let iComissaoSeguro(ByVal vData As Integer)
    mvariComissaoSeguro = vData
End Property

Public Property Get iComissaoSeguro() As Integer
    iComissaoSeguro = mvariComissaoSeguro
End Property

Public Property Let sContaCorrente(ByVal vData As String)
    mvarsContaCorrente = vData
End Property

Public Property Get sContaCorrente() As String
    sContaCorrente = mvarsContaCorrente
End Property

Public Property Let sAgencia(ByVal vData As String)
    mvarsAgencia = vData
End Property

Public Property Get sAgencia() As String
    sAgencia = mvarsAgencia
End Property

Public Property Let iBanco(ByVal vData As Integer)
    mvariBanco = vData
End Property

Public Property Get iBanco() As Integer
    iBanco = mvariBanco
End Property

Public Property Let iComissaoIPI(ByVal vData As Integer)
    mvariComissaoIPI = vData
End Property

Public Property Get iComissaoIPI() As Integer
    iComissaoIPI = mvariComissaoIPI
End Property

Public Property Let iComissaoICM(ByVal vData As Integer)
    mvariComissaoICM = vData
End Property

Public Property Get iComissaoICM() As Integer
    iComissaoICM = mvariComissaoICM
End Property

Public Property Let iComissaoFrete(ByVal vData As Integer)
    mvariComissaoFrete = vData
End Property

Public Property Get iComissaoFrete() As Integer
    iComissaoFrete = mvariComissaoFrete
End Property

Public Property Let iComissaoSobreTotal(ByVal vData As Integer)
    mvariComissaoSobreTotal = vData
End Property

Public Property Get iComissaoSobreTotal() As Integer
    iComissaoSobreTotal = mvariComissaoSobreTotal
End Property

Public Property Let dPercComissaoEmissao(ByVal vData As Double)
    mvardPercComissaoEmissao = vData
End Property

Public Property Get dPercComissaoEmissao() As Double
    dPercComissaoEmissao = mvardPercComissaoEmissao
End Property

Public Property Let dPercComissaoBaixa(ByVal vData As Double)
    mvardPercComissaoBaixa = vData
End Property

Public Property Get dPercComissaoBaixa() As Double
    dPercComissaoBaixa = mvardPercComissaoBaixa
End Property

Public Property Let dPercComissao(ByVal vData As Double)
    mvardPercComissao = vData
End Property

Public Property Get dPercComissao() As Double
    dPercComissao = mvardPercComissao
End Property

Public Property Let dSaldoComissao(ByVal vData As Double)
    mvardSaldoComissao = vData
End Property

Public Property Get dSaldoComissao() As Double
    dSaldoComissao = mvardSaldoComissao
End Property

Public Property Let iCodRegiao(ByVal vData As Integer)
    mvariCodRegiao = vData
End Property

Public Property Get iCodRegiao() As Integer
    iCodRegiao = mvariCodRegiao
End Property

Public Property Let sMatricula(ByVal vData As String)
    mvarsMatricula = vData
End Property

Public Property Get sMatricula() As String
    sMatricula = mvarsMatricula
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let lEndereco(ByVal vData As Long)
    mvarlEndereco = vData
End Property

Public Property Get lEndereco() As Long
    lEndereco = mvarlEndereco
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let iSuperior(ByVal vData As Integer)
    mvariSuperior = vData
End Property

Public Property Get iSuperior() As Integer
    iSuperior = mvariSuperior
End Property

Public Property Let iCargo(ByVal vData As Integer)
    mvariCargo = vData
End Property

Public Property Get iCargo() As Integer
    iCargo = mvariCargo
End Property



