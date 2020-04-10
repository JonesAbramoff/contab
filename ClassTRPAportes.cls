VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPAportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlCodigo As Long
Private mvarlCliente As Long
Private mvariFilialCliente As Integer
Private mvardtDataEmissao As Date
Private mvarsObservacao As String
Private mvarsHistorico As String
Private mvariTipo As Integer
Private mvariMoeda As Integer
Private mvardPrevValor As Double
Private mvardtPrevDataDe As Date
Private mvardtPrevDataAte As Date
Private mvarcolPagtoDireto As New Collection
Private mvarcolPagtoSobreFatura As New Collection
Private mvarcolPagtoCondicionados As New Collection
Private mvariTipoPagto As Integer
Private mvarlNumIntDocPagto As Long
Private mvarcolPagtoSobreFaturaCond As New Collection
Private mvariProxParcela As Integer

Public Property Let iProxParcela(ByVal vData As Integer)
    mvariProxParcela = vData
End Property

Public Property Get iProxParcela() As Integer
    iProxParcela = mvariProxParcela
End Property

Public Property Get colPagtoSobreFaturaCond() As Collection
    Set colPagtoSobreFaturaCond = mvarcolPagtoSobreFaturaCond
End Property

Public Property Let lNumIntDocPagto(ByVal vData As Long)
    mvarlNumIntDocPagto = vData
End Property

Public Property Get lNumIntDocPagto() As Long
    lNumIntDocPagto = mvarlNumIntDocPagto
End Property

Public Property Let iTipoPagto(ByVal vData As Integer)
    mvariTipoPagto = vData
End Property

Public Property Get iTipoPagto() As Integer
    iTipoPagto = mvariTipoPagto
End Property

Public Property Set colPagtoCondicionados(ByVal vData As Collection)
    Set mvarcolPagtoCondicionados = vData
End Property

Public Property Get colPagtoCondicionados() As Collection
    Set colPagtoCondicionados = mvarcolPagtoCondicionados
End Property

Public Property Set colPagtoSobreFatura(ByVal vData As Collection)
    Set mvarcolPagtoSobreFatura = vData
End Property

Public Property Get colPagtoSobreFatura() As Collection
    Set colPagtoSobreFatura = mvarcolPagtoSobreFatura
End Property

Public Property Set colPagtoDireto(ByVal vData As Collection)
    Set mvarcolPagtoDireto = vData
End Property

Public Property Get colPagtoDireto() As Collection
    Set colPagtoDireto = mvarcolPagtoDireto
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
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

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let sHistorico(ByVal vData As String)
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let dPrevValor(ByVal vData As Double)
    mvardPrevValor = vData
End Property

Public Property Get dPrevValor() As Double
    dPrevValor = mvardPrevValor
End Property

Public Property Let dtPrevDataDe(ByVal vData As Date)
    mvardtPrevDataDe = vData
End Property

Public Property Get dtPrevDataDe() As Date
    dtPrevDataDe = mvardtPrevDataDe
End Property

Public Property Let dtPrevDataAte(ByVal vData As Date)
    mvardtPrevDataAte = vData
End Property

Public Property Get dtPrevDataAte() As Date
    dtPrevDataAte = mvardtPrevDataAte
End Property
