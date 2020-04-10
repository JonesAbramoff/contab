VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassUltimaReducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private mvarlCOOReducaoZ As Long
Private mvardSubstTribICMS As Double
Private mvardIsentoICMS As Double
Private mvardNaoIncideICMS As Double
Private mvardTotalNaoFiscal As Double
Private mvardSubstTribISS As Double
Private mvardIsentoISS As Double
Private mvardNaoIncideISS As Double
Private mvardDescontoICMS As Double
Private mvardDescontoISS As Double
Private mvardDescontoNaoFiscal As Double
Private mvardAcrescimoICMS As Double
Private mvardAcrescimoISS As Double
Private mvardAcrescimoNaoFiscal As Double
Private mvardCancelamentoICMS As Double
Private mvardCancelamentoISS As Double
Private mvardCancelamentoNaoFiscal As Double
Private mvardtDataReducao As Date
Private mvarcolAliquotas As New Collection
Private mvardVendaBruta As Double
Private mvarsIncidenciaDescontoISS As String
Private mvarsHoraReducao As String
Private mvardtDataMovimento As Date

Public Property Let dtDataMovimento(ByVal vData As Date)
    mvardtDataMovimento = vData
End Property

Public Property Get dtDataMovimento() As Date
    dtDataMovimento = mvardtDataMovimento
End Property

Public Property Let sHoraReducao(ByVal vData As String)
    mvarsHoraReducao = vData
End Property

Public Property Get sHoraReducao() As String
    sHoraReducao = mvarsHoraReducao
End Property

Public Property Let sIncidenciaDescontoISS(ByVal vData As String)
    mvarsIncidenciaDescontoISS = vData
End Property

Public Property Get sIncidenciaDescontoISS() As String
    sIncidenciaDescontoISS = mvarsIncidenciaDescontoISS
End Property

Public Property Let dVendaBruta(ByVal vData As Double)
    mvardVendaBruta = vData
End Property

Public Property Get dVendaBruta() As Double
    dVendaBruta = mvardVendaBruta
End Property

Public Property Set colAliquotas(ByVal vData As Collection)
    Set mvarcolAliquotas = vData
End Property

Public Property Get colAliquotas() As Collection
    Set colAliquotas = mvarcolAliquotas
End Property

Public Property Let dtDataReducao(ByVal vData As Date)
    mvardtDataReducao = vData
End Property

Public Property Get dtDataReducao() As Date
    dtDataReducao = mvardtDataReducao
End Property

Public Property Let dCancelamentoNaoFiscal(ByVal vData As Double)
    mvardCancelamentoNaoFiscal = vData
End Property

Public Property Get dCancelamentoNaoFiscal() As Double
    dCancelamentoNaoFiscal = mvardCancelamentoNaoFiscal
End Property

Public Property Let dCancelamentoISS(ByVal vData As Double)
    mvardCancelamentoISS = vData
End Property

Public Property Get dCancelamentoISS() As Double
    dCancelamentoISS = mvardCancelamentoISS
End Property

Public Property Let dCancelamentoICMS(ByVal vData As Double)
    mvardCancelamentoICMS = vData
End Property

Public Property Get dCancelamentoICMS() As Double
    dCancelamentoICMS = mvardCancelamentoICMS
End Property

Public Property Let dAcrescimoNaoFiscal(ByVal vData As Double)
    mvardAcrescimoNaoFiscal = vData
End Property

Public Property Get dAcrescimoNaoFiscal() As Double
    dAcrescimoNaoFiscal = mvardAcrescimoNaoFiscal
End Property

Public Property Let dAcrescimoISS(ByVal vData As Double)
    mvardAcrescimoISS = vData
End Property

Public Property Get dAcrescimoISS() As Double
    dAcrescimoISS = mvardAcrescimoISS
End Property

Public Property Let dAcrescimoICMS(ByVal vData As Double)
    mvardAcrescimoICMS = vData
End Property

Public Property Get dAcrescimoICMS() As Double
    dAcrescimoICMS = mvardAcrescimoICMS
End Property

Public Property Let dDescontoNaoFiscal(ByVal vData As Double)
    mvardDescontoNaoFiscal = vData
End Property

Public Property Get dDescontoNaoFiscal() As Double
    dDescontoNaoFiscal = mvardDescontoNaoFiscal
End Property

Public Property Let dDescontoISS(ByVal vData As Double)
    mvardDescontoISS = vData
End Property

Public Property Get dDescontoISS() As Double
    dDescontoISS = mvardDescontoISS
End Property

Public Property Let dDescontoICMS(ByVal vData As Double)
    mvardDescontoICMS = vData
End Property

Public Property Get dDescontoICMS() As Double
    dDescontoICMS = mvardDescontoICMS
End Property

Public Property Let dNaoIncideISS(ByVal vData As Double)
    mvardNaoIncideISS = vData
End Property

Public Property Get dNaoIncideISS() As Double
    dNaoIncideISS = mvardNaoIncideISS
End Property


Public Property Let dIsentoISS(ByVal vData As Double)
    mvardIsentoISS = vData
End Property

Public Property Get dIsentoISS() As Double
    dIsentoISS = mvardIsentoISS
End Property


Public Property Let dSubstTribISS(ByVal vData As Double)
    mvardSubstTribISS = vData
End Property

Public Property Get dSubstTribISS() As Double
    dSubstTribISS = mvardSubstTribISS
End Property

Public Property Let dTotalNaoFiscal(ByVal vData As Double)
    mvardTotalNaoFiscal = vData
End Property

Public Property Get dTotalNaoFiscal() As Double
    dTotalNaoFiscal = mvardTotalNaoFiscal
End Property

Public Property Let dNaoIncideICMS(ByVal vData As Double)
    mvardNaoIncideICMS = vData
End Property

Public Property Get dNaoIncideICMS() As Double
    dNaoIncideICMS = mvardNaoIncideICMS
End Property

Public Property Let dIsentoICMS(ByVal vData As Double)
    mvardIsentoICMS = vData
End Property

Public Property Get dIsentoICMS() As Double
    dIsentoICMS = mvardIsentoICMS
End Property

Public Property Let dSubstTribICMS(ByVal vData As Double)
    mvardSubstTribICMS = vData
End Property

Public Property Get dSubstTribICMS() As Double
    dSubstTribICMS = mvardSubstTribICMS
End Property

Public Property Let lCOOReducaoZ(ByVal vData As Long)
    mvarlCOOReducaoZ = vData
End Property

Public Property Get lCOOReducaoZ() As Long
    lCOOReducaoZ = mvarlCOOReducaoZ
End Property

            
Private Sub Class_Initialize()
    mvardtDataReducao = DATA_NULA
    mvardtDataMovimento = DATA_NULA
End Sub
