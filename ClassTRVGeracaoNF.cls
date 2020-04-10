VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVGeracaoNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolTitulos As New Collection
Private mvardtDataDe As Date
Private mvardtDataAte As Date
Private mvarcolFiliais As New Collection
Private mvariTipo As Integer
Private mvarsTipoDoc As String
Private mvarcolItens As New Collection
Private mvarcolNF As New Collection
Private mvarcolNFFilialO As New Collection
Private mvarcolNFFilialD As New Collection
Private mvarcolNFClientes As New Collection
Private mvarcolSelecao As New Collection
Private mvariExtrairDados As Integer
Private mvarsSenhaSigav As String
Private mvarlTituloDe As Long
Private mvarlTituloAte As Long
Private mvardtDataBaixaDe As Date
Private mvardtDataBaixaAte As Date
Private mvarsFiltro As String
Private mvardtDataEmissao As Date
Private mvariGerarNFParaCadaFat As Integer
Private mvariEmpresa As Integer

Public Property Let iEmpresa(ByVal vData As Integer)
    mvariEmpresa = vData
End Property

Public Property Get iEmpresa() As Integer
    iEmpresa = mvariEmpresa
End Property

Public Property Let iGerarNFParaCadaFat(ByVal vData As Integer)
    mvariGerarNFParaCadaFat = vData
End Property

Public Property Get iGerarNFParaCadaFat() As Integer
    iGerarNFParaCadaFat = mvariGerarNFParaCadaFat
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let sFiltro(ByVal vData As String)
    mvarsFiltro = vData
End Property

Public Property Get sFiltro() As String
    sFiltro = mvarsFiltro
End Property

Public Property Let dtDataBaixaDe(ByVal vData As Date)
    mvardtDataBaixaDe = vData
End Property

Public Property Get dtDataBaixaDe() As Date
    dtDataBaixaDe = mvardtDataBaixaDe
End Property

Public Property Let dtDataBaixaAte(ByVal vData As Date)
    mvardtDataBaixaAte = vData
End Property

Public Property Get dtDataBaixaAte() As Date
    dtDataBaixaAte = mvardtDataBaixaAte
End Property

Public Property Let lTituloDe(ByVal vData As Long)
    mvarlTituloDe = vData
End Property

Public Property Get lTituloDe() As Long
    lTituloDe = mvarlTituloDe
End Property

Public Property Let lTituloAte(ByVal vData As Long)
    mvarlTituloAte = vData
End Property

Public Property Get lTituloAte() As Long
    lTituloAte = mvarlTituloAte
End Property

Public Property Let sSenhaSigav(ByVal vData As String)
    mvarsSenhaSigav = vData
End Property

Public Property Get sSenhaSigav() As String
    sSenhaSigav = mvarsSenhaSigav
End Property

Public Property Let iExtrairDados(ByVal vData As Integer)
    mvariExtrairDados = vData
End Property

Public Property Get iExtrairDados() As Integer
    iExtrairDados = mvariExtrairDados
End Property

Public Property Set colNF(ByVal vData As Collection)
    Set mvarcolNF = vData
End Property

Public Property Get colNF() As Collection
    Set colNF = mvarcolNF
End Property

Public Property Set colNFFilialO(ByVal vData As Collection)
    Set mvarcolNFFilialO = vData
End Property

Public Property Get colNFFilialO() As Collection
    Set colNFFilialO = mvarcolNFFilialO
End Property

Public Property Set colNFFilialD(ByVal vData As Collection)
    Set mvarcolNFFilialD = vData
End Property

Public Property Get colNFFilialD() As Collection
    Set colNFFilialD = mvarcolNFFilialD
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let sTipoDoc(ByVal vData As String)
    mvarsTipoDoc = vData
End Property

Public Property Get sTipoDoc() As String
    sTipoDoc = mvarsTipoDoc
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Set colFiliais(ByVal vData As Collection)
    Set mvarcolFiliais = vData
End Property

Public Property Get colFiliais() As Collection
    Set colFiliais = mvarcolFiliais
End Property

Public Property Let dtDataDe(ByVal vData As Date)
    mvardtDataDe = vData
End Property

Public Property Get dtDataDe() As Date
    dtDataDe = mvardtDataDe
End Property

Public Property Let dtDataAte(ByVal vData As Date)
    mvardtDataAte = vData
End Property

Public Property Get dtDataAte() As Date
    dtDataAte = mvardtDataAte
End Property

Public Property Set colTitulos(ByVal vData As Collection)
    Set mvarcolTitulos = vData
End Property

Public Property Get colTitulos() As Collection
    Set colTitulos = mvarcolTitulos
End Property

Public Property Set colSelecao(ByVal vData As Collection)
    Set mvarcolSelecao = vData
End Property

Public Property Get colSelecao() As Collection
    Set colSelecao = mvarcolSelecao
End Property

Public Property Set colNFClientes(ByVal vData As Collection)
    Set mvarcolNFClientes = vData
End Property

Public Property Get colNFClientes() As Collection
    Set colNFClientes = mvarcolNFClientes
End Property
