VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelMvDiaCtaRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel As Long
Private mvariSeqCtaPer As Integer
Private mvariSeqCtaDia As Integer
Private mvardtData As Date
Private mvardCredito As Double
Private mvardDebito As Double
Private mvardSaldoAnterior As Double
Private mvardSaldoAtual As Double
Private mvarcolLcto As New Collection

Public Property Set colLcto(ByVal vData As Object)
    Set mvarcolLcto = vData
End Property

Public Property Get colLcto() As Collection
    Set colLcto = mvarcolLcto
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let iSeqCtaPer(ByVal vData As Integer)
    mvariSeqCtaPer = vData
End Property

Public Property Get iSeqCtaPer() As Integer
    iSeqCtaPer = mvariSeqCtaPer
End Property

Public Property Let iSeqCtaDia(ByVal vData As Integer)
    mvariSeqCtaDia = vData
End Property

Public Property Get iSeqCtaDia() As Integer
    iSeqCtaDia = mvariSeqCtaDia
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dCredito(ByVal vData As Double)
    mvardCredito = vData
End Property

Public Property Get dCredito() As Double
    dCredito = mvardCredito
End Property

Public Property Let dDebito(ByVal vData As Double)
    mvardDebito = vData
End Property

Public Property Get dDebito() As Double
    dDebito = mvardDebito
End Property

Public Property Let dSaldoAnterior(ByVal vData As Double)
    mvardSaldoAnterior = vData
End Property

Public Property Get dSaldoAnterior() As Double
    dSaldoAnterior = mvardSaldoAnterior
End Property

Public Property Let dSaldoAtual(ByVal vData As Double)
    mvardSaldoAtual = vData
End Property

Public Property Get dSaldoAtual() As Double
    dSaldoAtual = mvardSaldoAtual
End Property
