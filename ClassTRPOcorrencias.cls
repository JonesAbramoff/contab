VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPOcorrencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumVou As Long
Private mvarsTipoDoc As String
Private mvarsSerie As String
Private mvarlCodigo As Long
Private mvarlCliente As Long
Private mvariFilialCliente As Integer
Private mvardtDataEmissao As Date
Private mvarsObservacao As String
Private mvariStatus As Integer
Private mvariOrigem As Integer
Private mvarsHistorico As String
Private mvariFormaPagto As Integer
Private mvarlNumIntDocDestino As Long
Private mvardValorTotal As Double
Private mvarcolDetalhamento As New Collection
Private mvariTipoDocDestino As Integer
Private mvarlNumDocDestino As Long
Private mvardValorOCRBrutoVou As Double
Private mvardValorOCRCMAVou As Double

Public Property Let dValorOCRCMAVou(ByVal vData As Double)
    mvardValorOCRCMAVou = vData
End Property

Public Property Get dValorOCRCMAVou() As Double
    dValorOCRCMAVou = mvardValorOCRCMAVou
End Property

Public Property Let dValorOCRBrutoVou(ByVal vData As Double)
    mvardValorOCRBrutoVou = vData
End Property

Public Property Get dValorOCRBrutoVou() As Double
    dValorOCRBrutoVou = mvardValorOCRBrutoVou
End Property

Public Property Let lNumDocDestino(ByVal vData As Long)
    mvarlNumDocDestino = vData
End Property

Public Property Get lNumDocDestino() As Long
    lNumDocDestino = mvarlNumDocDestino
End Property

Public Property Let iTipoDocDestino(ByVal vData As Integer)
    mvariTipoDocDestino = vData
End Property

Public Property Get iTipoDocDestino() As Integer
    iTipoDocDestino = mvariTipoDocDestino
End Property

Public Property Set colDetalhamento(ByVal vData As Collection)
    Set mvarcolDetalhamento = vData
End Property

Public Property Get colDetalhamento() As Collection
    Set colDetalhamento = mvarcolDetalhamento
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
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

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let iOrigem(ByVal vData As Integer)
    mvariOrigem = vData
End Property

Public Property Get iOrigem() As Integer
    iOrigem = mvariOrigem
End Property

Public Property Let sHistorico(ByVal vData As String)
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let iFormaPagto(ByVal vData As Integer)
    mvariFormaPagto = vData
End Property

Public Property Get iFormaPagto() As Integer
    iFormaPagto = mvariFormaPagto
End Property

Public Property Let lNumIntDocDestino(ByVal vData As Long)
    mvarlNumIntDocDestino = vData
End Property

Public Property Get lNumIntDocDestino() As Long
    lNumIntDocDestino = mvarlNumIntDocDestino
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property
