VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassReducaoZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvardtDataHora As Date
Private mvariNumeroECF As Integer
Private mvarsRazaoSocial As String
Private mvarsCNPJ As String
Private mvarsInscEstadual As String
Private mvarlNumOperacao As Long
Private mvarlReducao As Long
Private mvardTotalAcumulado As Double
Private mvardCancelamento As Double
Private mvardDesconto As Double
Private mvardVendaBrutalDia As Double
Private mvardVendaLiquidaDia As Double
Private mvardVendaSubstTributaria As Double
Private mvardVendasIsentas As Double
Private mvardVendasNaoTributadas As Double
Private mvardVendasTributadas As Double
Private mvarcolValorIncidente As New Collection
Private mvarcolAliquotaICMS As New Collection
Private mvarcolICMSDebitado As New Collection
Public Property Set colICMSDebitado(ByVal vData As Collection)
    Set mvarcolICMSDebitado = vData
End Property


Public Property Get colICMSDebitado() As Collection
    Set colICMSDebitado = mvarcolICMSDebitado
End Property



Public Property Set colAliquotaICMS(ByVal vData As Collection)
    Set mvarcolAliquotaICMS = vData
End Property


Public Property Get colAliquotaICMS() As Collection
    Set colAliquotaICMS = mvarcolAliquotaICMS
End Property



Public Property Set colValorIncidente(ByVal vData As Collection)
    Set mvarcolValorIncidente = vData
End Property


Public Property Get colValorIncidente() As Collection
    Set colValorIncidente = mvarcolValorIncidente
End Property



Public Property Let dVendasTributadas(ByVal vData As Double)
    mvardVendasTributadas = vData
End Property


Public Property Get dVendasTributadas() As Double
    dVendasTributadas = mvardVendasTributadas
End Property



Public Property Let dVendasNaoTributadas(ByVal vData As Double)
    mvardVendasNaoTributadas = vData
End Property


Public Property Get dVendasNaoTributadas() As Double
    dVendasNaoTributadas = mvardVendasNaoTributadas
End Property



Public Property Let dVendasIsentas(ByVal vData As Double)
    mvardVendasIsentas = vData
End Property


Public Property Get dVendasIsentas() As Double
    dVendasIsentas = mvardVendasIsentas
End Property



Public Property Let dVendaSubstTributaria(ByVal vData As Double)
    mvardVendaSubstTributaria = vData
End Property


Public Property Get dVendaSubstTributaria() As Double
    dVendaSubstTributaria = mvardVendaSubstTributaria
End Property



Public Property Let dVendaLiquidaDia(ByVal vData As Double)
    mvardVendaLiquidaDia = vData
End Property


Public Property Get dVendaLiquidaDia() As Double
    dVendaLiquidaDia = mvardVendaLiquidaDia
End Property



Public Property Let dVendaBrutalDia(ByVal vData As Double)
    mvardVendaBrutalDia = vData
End Property


Public Property Get dVendaBrutalDia() As Double
    dVendaBrutalDia = mvardVendaBrutalDia
End Property



Public Property Let dDesconto(ByVal vData As Double)
    mvardDesconto = vData
End Property


Public Property Get dDesconto() As Double
    dDesconto = mvardDesconto
End Property



Public Property Let dCancelamento(ByVal vData As Double)
    mvardCancelamento = vData
End Property


Public Property Get dCancelamento() As Double
    dCancelamento = mvardCancelamento
End Property



Public Property Let dTotalAcumulado(ByVal vData As Double)
    mvardTotalAcumulado = vData
End Property


Public Property Get dTotalAcumulado() As Double
    dTotalAcumulado = mvardTotalAcumulado
End Property



Public Property Let lReducao(ByVal vData As Long)
    mvarlReducao = vData
End Property


Public Property Get lReducao() As Long
    lReducao = mvarlReducao
End Property



Public Property Let lNumOperacao(ByVal vData As Long)
    mvarlNumOperacao = vData
End Property


Public Property Get lNumOperacao() As Long
    lNumOperacao = mvarlNumOperacao
End Property



Public Property Let sInscEstadual(ByVal vData As String)
    mvarsInscEstadual = vData
End Property


Public Property Get sInscEstadual() As String
    sInscEstadual = mvarsInscEstadual
End Property



Public Property Let sCNPJ(ByVal vData As String)
    mvarsCNPJ = vData
End Property


Public Property Get sCNPJ() As String
    sCNPJ = mvarsCNPJ
End Property



Public Property Let sRazaoSocial(ByVal vData As String)
    mvarsRazaoSocial = vData
End Property


Public Property Get sRazaoSocial() As String
    sRazaoSocial = mvarsRazaoSocial
End Property



Public Property Let iNumeroECF(ByVal vData As Integer)
    mvariNumeroECF = vData
End Property


Public Property Get iNumeroECF() As Integer
    iNumeroECF = mvariNumeroECF
End Property



Public Property Let dtDataHora(ByVal vData As Date)
    mvardtDataHora = vData
End Property


Public Property Get dtDataHora() As Date
    dtDataHora = mvardtDataHora
End Property



