VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCompServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntDocOrigem As Long
Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvarlCliente As Long
Private mvardtDataEmissao As Date
Private mvarsProduto As String
Private mvarsMaterial As String
Private mvardQuantMaterial As Double
Private mvarsUM As String
Private mvardValorMercadoria As Double
Private mvariTipoEmbalagem As Integer
Private mvardFretePeso As Double
Private mvardPedagio As Double
Private mvardAdValoren As Double
Private mvariDespachante As Integer
Private mvarlCodProgNavio As Long
Private mvardtDataDemurrage As Date
Private mvariTipoContainer As Integer
Private mvarsCodigoContainer As String
Private mvardTara As Double
Private mvarsLacre As String
Private mvarsObservacao As String
Private mvarcolCompServItem As New Collection
Private mvarlNumIntNota As Long
Private mvariSituacao As Integer
Private mvardQuantidade As Double
Private mvardValorContainer As Double
Private mvariTipoOperacao As Integer
Private mvarsPorto As String
Private mvarsAlfandega As String
Private mvarsEmbalagem As String
Private mvarsCidadeDestino As String
Private mvarsCidadeOrigem As String
Private mvarsUFOrigem As String
Private mvarsUFDestino As String

Public Property Let sUFDestino(ByVal vData As String)
    mvarsUFDestino = vData
End Property

Public Property Get sUFDestino() As String
    sUFDestino = mvarsUFDestino
End Property

Public Property Let sUFOrigem(ByVal vData As String)
    mvarsUFOrigem = vData
End Property

Public Property Get sUFOrigem() As String
    sUFOrigem = mvarsUFOrigem
End Property

Public Property Let sCidadeOrigem(ByVal vData As String)
    mvarsCidadeOrigem = vData
End Property

Public Property Get sCidadeOrigem() As String
    sCidadeOrigem = mvarsCidadeOrigem
End Property

Public Property Let sCidadeDestino(ByVal vData As String)
    mvarsCidadeDestino = vData
End Property

Public Property Get sCidadeDestino() As String
    sCidadeDestino = mvarsCidadeDestino
End Property

Public Property Let sEmbalagem(ByVal vData As String)
    mvarsEmbalagem = vData
End Property

Public Property Get sEmbalagem() As String
    sEmbalagem = mvarsEmbalagem
End Property

Public Property Let sAlfandega(ByVal vData As String)
    mvarsAlfandega = vData
End Property

Public Property Get sAlfandega() As String
    sAlfandega = mvarsAlfandega
End Property

Public Property Let sPorto(ByVal vData As String)
    mvarsPorto = vData
End Property

Public Property Get sPorto() As String
    sPorto = mvarsPorto
End Property

Public Property Let iTipoOperacao(ByVal vData As Integer)
    mvariTipoOperacao = vData
End Property

Public Property Get iTipoOperacao() As Integer
    iTipoOperacao = mvariTipoOperacao
End Property

Public Property Let dValorContainer(ByVal vData As Double)
    mvardValorContainer = vData
End Property

Public Property Get dValorContainer() As Double
    dValorContainer = mvardValorContainer
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let iSituacao(ByVal vData As Integer)
    mvariSituacao = vData
End Property

Public Property Get iSituacao() As Integer
    iSituacao = mvariSituacao
End Property

Public Property Let lNumIntNota(ByVal vData As Long)
    mvarlNumIntNota = vData
End Property

Public Property Get lNumIntNota() As Long
    lNumIntNota = mvarlNumIntNota
End Property

Public Property Set colCompServItem(ByVal vData As Object)
    Set mvarcolCompServItem = vData
End Property

Public Property Get colCompServItem() As Collection
    Set colCompServItem = mvarcolCompServItem
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let sLacre(ByVal vData As String)
    mvarsLacre = vData
End Property

Public Property Get sLacre() As String
    sLacre = mvarsLacre
End Property

Public Property Let dTara(ByVal vData As Double)
    mvardTara = vData
End Property

Public Property Get dTara() As Double
    dTara = mvardTara
End Property

Public Property Let lCodProgNavio(ByVal vData As Long)
    mvarlCodProgNavio = vData
End Property

Public Property Get lCodProgNavio() As Long
    lCodProgNavio = mvarlCodProgNavio
End Property

Public Property Let sCodigoContainer(ByVal vData As String)
    mvarsCodigoContainer = vData
End Property

Public Property Get sCodigoContainer() As String
    sCodigoContainer = mvarsCodigoContainer
End Property

Public Property Let iTipoContainer(ByVal vData As Integer)
    mvariTipoContainer = vData
End Property

Public Property Get iTipoContainer() As Integer
    iTipoContainer = mvariTipoContainer
End Property

Public Property Let dtDataDemurrage(ByVal vData As Date)
    mvardtDataDemurrage = vData
End Property

Public Property Get dtDataDemurrage() As Date
    dtDataDemurrage = mvardtDataDemurrage
End Property

Public Property Let iDespachante(ByVal vData As Integer)
    mvariDespachante = vData
End Property

Public Property Get iDespachante() As Integer
    iDespachante = mvariDespachante
End Property

Public Property Let dAdValoren(ByVal vData As Double)
    mvardAdValoren = vData
End Property

Public Property Get dAdValoren() As Double
    dAdValoren = mvardAdValoren
End Property

Public Property Let dPedagio(ByVal vData As Double)
    mvardPedagio = vData
End Property

Public Property Get dPedagio() As Double
    dPedagio = mvardPedagio
End Property

Public Property Let dFretePeso(ByVal vData As Double)
    mvardFretePeso = vData
End Property

Public Property Get dFretePeso() As Double
    dFretePeso = mvardFretePeso
End Property

Public Property Let iTipoEmbalagem(ByVal vData As Integer)
    mvariTipoEmbalagem = vData
End Property

Public Property Get iTipoEmbalagem() As Integer
    iTipoEmbalagem = mvariTipoEmbalagem
End Property

Public Property Let dValorMercadoria(ByVal vData As Double)
    mvardValorMercadoria = vData
End Property

Public Property Get dValorMercadoria() As Double
    dValorMercadoria = mvardValorMercadoria
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let dQuantMaterial(ByVal vData As Double)
    mvardQuantMaterial = vData
End Property

Public Property Get dQuantMaterial() As Double
    dQuantMaterial = mvardQuantMaterial
End Property

Public Property Let sMaterial(ByVal vData As String)
    mvarsMaterial = vData
End Property

Public Property Get sMaterial() As String
    sMaterial = mvarsMaterial
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDocOrigem(ByVal vData As Long)
    mvarlNumIntDocOrigem = vData
End Property

Public Property Get lNumIntDocOrigem() As Long
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

