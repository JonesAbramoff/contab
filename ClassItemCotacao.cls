VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemCotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariCondPagto As Integer
Private mvardtDataReferencia As Date
Private mvardPrecoUnitario As Double
Private mvardOutrasDespesas As Double
Private mvardValorSeguro As Double
Private mvardValorDesconto As Double
Private mvardValorTotal As Double
Private mvardValorIPI As Double
Private mvardAliquotaIPI As Double
Private mvardAliquotaICMS As Double
Private mvariPrazoEntrega As Integer
Private mvardQuantEntrega As Double
Private mvarlObservacao As Long
Private mvardValorFrete As Double
Private mvarsObservacao As String
Private mvariMoeda As Integer
Private mvardTaxa As Double

Public Property Let dTaxa(ByVal vData As Double)
    mvardTaxa = vData
End Property

Public Property Get dTaxa() As Double
    dTaxa = mvardTaxa
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dValorFrete(ByVal vData As Double)
    mvardValorFrete = vData
End Property

Public Property Get dValorFrete() As Double
    dValorFrete = mvardValorFrete
End Property

Public Property Let lObservacao(ByVal vData As Long)
    mvarlObservacao = vData
End Property

Public Property Get lObservacao() As Long
    lObservacao = mvarlObservacao
End Property

Public Property Let dQuantEntrega(ByVal vData As Double)
    mvardQuantEntrega = vData
End Property

Public Property Get dQuantEntrega() As Double
    dQuantEntrega = mvardQuantEntrega
End Property

Public Property Let iPrazoEntrega(ByVal vData As Integer)
    mvariPrazoEntrega = vData
End Property

Public Property Get iPrazoEntrega() As Integer
    iPrazoEntrega = mvariPrazoEntrega
End Property

Public Property Let dAliquotaICMS(ByVal vData As Double)
    mvardAliquotaICMS = vData
End Property

Public Property Get dAliquotaICMS() As Double
    dAliquotaICMS = mvardAliquotaICMS
End Property

Public Property Let dAliquotaIPI(ByVal vData As Double)
    mvardAliquotaIPI = vData
End Property

Public Property Get dAliquotaIPI() As Double
    dAliquotaIPI = mvardAliquotaIPI
End Property

Public Property Let dValorIPI(ByVal vData As Double)
    mvardValorIPI = vData
End Property

Public Property Get dValorIPI() As Double
    dValorIPI = mvardValorIPI
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorSeguro(ByVal vData As Double)
    mvardValorSeguro = vData
End Property

Public Property Get dValorSeguro() As Double
    dValorSeguro = mvardValorSeguro
End Property

Public Property Let dOutrasDespesas(ByVal vData As Double)
    mvardOutrasDespesas = vData
End Property

Public Property Get dOutrasDespesas() As Double
    dOutrasDespesas = mvardOutrasDespesas
End Property

Public Property Let dPrecoUnitario(ByVal vData As Double)
    mvardPrecoUnitario = vData
End Property

Public Property Get dPrecoUnitario() As Double
    dPrecoUnitario = mvardPrecoUnitario
End Property

Public Property Let dtDataReferencia(ByVal vData As Date)
    mvardtDataReferencia = vData
End Property

Public Property Get dtDataReferencia() As Date
    dtDataReferencia = mvardtDataReferencia
End Property

Public Property Let iCondPagto(ByVal vData As Integer)
    mvariCondPagto = vData
End Property

Public Property Get iCondPagto() As Integer
    iCondPagto = mvariCondPagto
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property
