VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAdmMeioPagto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsNome As String
Private mvariTipoMeioPagto As Integer
Private mvariDefasagemPagtoVista As Integer
Private mvarlEndereco As Long
Private mvariCodBanco As Integer
Private mvarsAgencia As String
Private mvariContaCorrenteInterna As Integer
Private mvarcolCondPagtoLoja As New Collection
Private mvardTaxaVista As Double
Private mvardTaxaParcelado As Double
Private mvarobjEndereco As New ClassEndereco
Private mvariFilialEmpresa As Integer
Private mvariGeraTituloRec As Integer
Private mvariRede As Integer
Private mvarsNomeRede As String
Private mvariTipoLog As Integer
Private mvariAtivo As Integer
Private mvardtDataLog As Date
Private mvariProxCodParcelamento As Integer

Public Property Let iProxCodParcelamento(ByVal vData As Integer)
    mvariProxCodParcelamento = vData
End Property

Public Property Get iProxCodParcelamento() As Integer
    iProxCodParcelamento = mvariProxCodParcelamento
End Property

Public Property Let dtDataLog(ByVal vData As Date)
    mvardtDataLog = vData
End Property

Public Property Get dtDataLog() As Date
    dtDataLog = mvardtDataLog
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let iTipoLog(ByVal vData As Integer)
    mvariTipoLog = vData
End Property

Public Property Get iTipoLog() As Integer
    iTipoLog = mvariTipoLog
End Property

Public Property Let sNomeRede(ByVal vData As String)
    mvarsNomeRede = vData
End Property

Public Property Get sNomeRede() As String
    sNomeRede = mvarsNomeRede
End Property

Public Property Let iRede(ByVal vData As Integer)
    mvariRede = vData
End Property

Public Property Get iRede() As Integer
    iRede = mvariRede
End Property

Public Property Let iGeraTituloRec(ByVal vData As Integer)
    mvariGeraTituloRec = vData
End Property

Public Property Get iGeraTituloRec() As Integer
    iGeraTituloRec = mvariGeraTituloRec
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Set objEndereco(ByVal vData As ClassEndereco)
    Set mvarobjEndereco = vData
End Property

Public Property Get objEndereco() As ClassEndereco
    Set objEndereco = mvarobjEndereco
End Property

Public Property Let dTaxaParcelado(ByVal vData As Double)
    mvardTaxaParcelado = vData
End Property

Public Property Get dTaxaParcelado() As Double
    dTaxaParcelado = mvardTaxaParcelado
End Property

Public Property Let dTaxaVista(ByVal vData As Double)
    mvardTaxaVista = vData
End Property

Public Property Get dTaxaVista() As Double
    dTaxaVista = mvardTaxaVista
End Property

Public Property Set colCondPagtoLoja(ByVal vData As Collection)
    Set mvarcolCondPagtoLoja = vData
End Property

Public Property Get colCondPagtoLoja() As Collection
    Set colCondPagtoLoja = mvarcolCondPagtoLoja
End Property

Public Property Let iContaCorrenteInterna(ByVal vData As Integer)
    mvariContaCorrenteInterna = vData
End Property

Public Property Get iContaCorrenteInterna() As Integer
    iContaCorrenteInterna = mvariContaCorrenteInterna
End Property

Public Property Let sAgencia(ByVal vData As String)
    mvarsAgencia = vData
End Property

Public Property Get sAgencia() As String
    sAgencia = mvarsAgencia
End Property

Public Property Let iCodBanco(ByVal vData As Integer)
    mvariCodBanco = vData
End Property

Public Property Get iCodBanco() As Integer
    iCodBanco = mvariCodBanco
End Property

Public Property Let lEndereco(ByVal vData As Long)
    mvarlEndereco = vData
End Property

Public Property Get lEndereco() As Long
    lEndereco = mvarlEndereco
End Property

Public Property Let iDefasagemPagtoVista(ByVal vData As Integer)
    mvariDefasagemPagtoVista = vData
End Property

Public Property Get iDefasagemPagtoVista() As Integer
    iDefasagemPagtoVista = mvariDefasagemPagtoVista
End Property

Public Property Let iTipoMeioPagto(ByVal vData As Integer)
    mvariTipoMeioPagto = vData
End Property

Public Property Get iTipoMeioPagto() As Integer
    iTipoMeioPagto = mvariTipoMeioPagto
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


