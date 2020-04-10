VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBaixaRecCancelar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntBaixaParcRecCanc As Long
Private mvarlCliente As Long
Private mvariFilialCliente As Integer
Private mvardtDataCancelamento As Date
Private mvardtDataBaixaInicial As Date
Private mvardtDataBaixaFinal As Date
Private mvardtDataVenctoInicial As Date
Private mvardtDataVenctoFinal As Date
Private mvarlTituloInicial As Long
Private mvarlTituloFinal As Long
Private mvariCtaCorrenteFiltro As Integer
Private mvariCtaCorrenteTodas As Integer
Private mvariTipoBaixas As Integer
Private mvariTipoCancelamento As Integer
Private mvardValorBaixasCancelar As Double
Private mvarsHistoricoCancelamento As String
Private mvarcolInfoBaixaRecCancelar As New Collection
Private mvarcolParcelasMovCCI As New Collection

Public Property Get colParcelasMovCCI() As Collection
    Set colParcelasMovCCI = mvarcolParcelasMovCCI
End Property

Public Property Set colParcelasMovCCI(ByVal vData As Collection)
    Set mvarcolParcelasMovCCI = vData
End Property

Public Property Let iTipoCancelamento(ByVal vData As Integer)
    mvariTipoCancelamento = vData
End Property

Public Property Get iTipoCancelamento() As Integer
    iTipoCancelamento = mvariTipoCancelamento
End Property

Public Property Let iTipoBaixas(ByVal vData As Integer)
    mvariTipoBaixas = vData
End Property

Public Property Get iTipoBaixas() As Integer
    iTipoBaixas = mvariTipoBaixas
End Property

Public Property Let dValorBaixasCancelar(ByVal vData As Double)
    mvardValorBaixasCancelar = vData
End Property

Public Property Get dValorBaixasCancelar() As Double
    dValorBaixasCancelar = mvardValorBaixasCancelar
End Property

Public Property Let iCtaCorrenteFiltro(ByVal vData As Integer)
    mvariCtaCorrenteFiltro = vData
End Property

Public Property Get iCtaCorrenteFiltro() As Integer
    iCtaCorrenteFiltro = mvariCtaCorrenteFiltro
End Property

Public Property Let lTituloFinal(ByVal vData As Long)
    mvarlTituloFinal = vData
End Property

Public Property Get lTituloFinal() As Long
    lTituloFinal = mvarlTituloFinal
End Property

Public Property Let lTituloInicial(ByVal vData As Long)
    mvarlTituloInicial = vData
End Property

Public Property Get lTituloInicial() As Long
    lTituloInicial = mvarlTituloInicial
End Property

Public Property Let dtDataVenctoFinal(ByVal vData As Date)
    mvardtDataVenctoFinal = vData
End Property

Public Property Get dtDataVenctoFinal() As Date
    dtDataVenctoFinal = mvardtDataVenctoFinal
End Property

Public Property Let dtDataVenctoInicial(ByVal vData As Date)
    mvardtDataVenctoInicial = vData
End Property

Public Property Get dtDataVenctoInicial() As Date
    dtDataVenctoInicial = mvardtDataVenctoInicial
End Property

Public Property Let dtDataBaixaFinal(ByVal vData As Date)
    mvardtDataBaixaFinal = vData
End Property

Public Property Get dtDataBaixaFinal() As Date
    dtDataBaixaFinal = mvardtDataBaixaFinal
End Property

Public Property Let dtDataBaixaInicial(ByVal vData As Date)
    mvardtDataBaixaInicial = vData
End Property

Public Property Get dtDataBaixaInicial() As Date
    dtDataBaixaInicial = mvardtDataBaixaInicial
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

Public Property Get colInfoBaixaRecCancelar() As Collection
    Set colInfoBaixaRecCancelar = mvarcolInfoBaixaRecCancelar
End Property

Public Property Set colInfoBaixaRecCancelar(ByVal vData As Collection)
    Set mvarcolInfoBaixaRecCancelar = vData
End Property

Public Property Let dtDataCancelamento(ByVal vData As Date)
    mvardtDataCancelamento = vData
End Property

Public Property Get dtDataCancelamento() As Date
    dtDataCancelamento = mvardtDataCancelamento
End Property

Public Property Let sHistoricoCancelamento(ByVal vData As String)
    mvarsHistoricoCancelamento = vData
End Property

Public Property Get sHistoricoCancelamento() As String
    sHistoricoCancelamento = mvarsHistoricoCancelamento
End Property

Public Property Let lNumIntBaixaParcRecCanc(ByVal vData As Long)
    mvarlNumIntBaixaParcRecCanc = vData
End Property

Public Property Get lNumIntBaixaParcRecCanc() As Long
    lNumIntBaixaParcRecCanc = mvarlNumIntBaixaParcRecCanc
End Property

Public Property Let iCtaCorrenteTodas(ByVal vData As Integer)
    mvariCtaCorrenteTodas = vData
End Property

Public Property Get iCtaCorrenteTodas() As Integer
    iCtaCorrenteTodas = mvariCtaCorrenteTodas
End Property

