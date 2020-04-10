VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBaixaAntecipDebCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCliente As Long
Private mvariFilialCliente As Integer
Private mvariTipoBaixa As Integer
Private mvardtDataEmissaoDe As Date
Private mvardtDataEmissaoAte As Date
Private mvariCtaCorrenteSelecionada As Integer
Private mvarlNumDebitoDe As Long
Private mvarlNumDebitoAte As Long
Private mvarsTipoDebito As String
Private mvarcolAntecipDebCliente As New Collection
Private mvariCtaCorrenteTodas As Integer
Private mvariTipoDebitoTodos As Integer
Private mvarlNumIntBaixa As Long
Private mvardtDataBaixa As Date
Private mvariCtaCorrenteBaixa As Integer
Private mvardValorBaixa As Double
Private mvarsHistoricoBaixa As String
Private mvarlNumMovtoCCI As Long

Public Property Let lNumMovtoCCI(ByVal vData As Long)
    mvarlNumMovtoCCI = vData
End Property

Public Property Get lNumMovtoCCI() As Long
    lNumMovtoCCI = mvarlNumMovtoCCI
End Property

Public Property Get colAntecipDebCliente() As Collection
    Set colAntecipDebCliente = mvarcolAntecipDebCliente
End Property

Public Property Set colAntecipDebCliente(ByVal vData As Collection)
    Set mvarcolAntecipDebCliente = vData
End Property

Public Property Let sTipoDebito(ByVal vData As String)
    mvarsTipoDebito = vData
End Property

Public Property Get sTipoDebito() As String
    sTipoDebito = mvarsTipoDebito
End Property

Public Property Let lNumDebitoAte(ByVal vData As Long)
    mvarlNumDebitoAte = vData
End Property

Public Property Get lNumDebitoAte() As Long
    lNumDebitoAte = mvarlNumDebitoAte
End Property

Public Property Let lNumDebitoDe(ByVal vData As Long)
    mvarlNumDebitoDe = vData
End Property

Public Property Get lNumDebitoDe() As Long
    lNumDebitoDe = mvarlNumDebitoDe
End Property

Public Property Let iCtaCorrenteSelecionada(ByVal vData As Integer)
    mvariCtaCorrenteSelecionada = vData
End Property

Public Property Get iCtaCorrenteSelecionada() As Integer
    iCtaCorrenteSelecionada = mvariCtaCorrenteSelecionada
End Property

Public Property Let dtDataEmissaoAte(ByVal vData As Date)
    mvardtDataEmissaoAte = vData
End Property

Public Property Get dtDataEmissaoAte() As Date
    dtDataEmissaoAte = mvardtDataEmissaoAte
End Property

Public Property Let dtDataEmissaoDe(ByVal vData As Date)
    mvardtDataEmissaoDe = vData
End Property

Public Property Get dtDataEmissaoDe() As Date
    dtDataEmissaoDe = mvardtDataEmissaoDe
End Property

Public Property Let iTipoBaixa(ByVal vData As Integer)
    mvariTipoBaixa = vData
End Property

Public Property Get iTipoBaixa() As Integer
    iTipoBaixa = mvariTipoBaixa
End Property

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iCtaCorrenteTodas(ByVal vData As Integer)
    mvariCtaCorrenteTodas = vData
End Property

Public Property Get iCtaCorrenteTodas() As Integer
    iCtaCorrenteTodas = mvariCtaCorrenteTodas
End Property

Public Property Let iTipoDebitoTodos(ByVal vData As Integer)
    mvariTipoDebitoTodos = vData
End Property

Public Property Get iTipoDebitoTodos() As Integer
    iTipoDebitoTodos = mvariTipoDebitoTodos
End Property

Public Property Let iCtaCorrenteBaixa(ByVal vData As Integer)
    mvariCtaCorrenteBaixa = vData
End Property

Public Property Get iCtaCorrenteBaixa() As Integer
    iCtaCorrenteBaixa = mvariCtaCorrenteBaixa
End Property

Public Property Let dtDataBaixa(ByVal vData As Date)
    mvardtDataBaixa = vData
End Property

Public Property Get dtDataBaixa() As Date
    dtDataBaixa = mvardtDataBaixa
End Property

Public Property Let dValorBaixa(ByVal vData As Double)
    mvardValorBaixa = vData
End Property

Public Property Get dValorBaixa() As Double
    dValorBaixa = mvardValorBaixa
End Property

Public Property Let sHistoricoBaixa(ByVal vData As String)
    mvarsHistoricoBaixa = vData
End Property

Public Property Get sHistoricoBaixa() As String
    sHistoricoBaixa = mvarsHistoricoBaixa
End Property

Public Property Let lNumIntBaixa(ByVal vData As Long)
    mvarlNumIntBaixa = vData
End Property

Public Property Get lNumIntBaixa() As Long
    lNumIntBaixa = mvarlNumIntBaixa
End Property

