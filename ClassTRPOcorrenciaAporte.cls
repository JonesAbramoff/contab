VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPOcorrenciaAporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocOCR As Long
Private mvariTipoPagtoAporte As Integer
Private mvarlNumIntDocPagtoAporte As Long
Private mvarlCodigoAporte As Long
Private mvariParcelaAporte As Integer
Private mvardValorAporte As Double
Private mvardValorAporteUSS As Double
Private mvariMoedaAporte As Integer
Private mvariPagtoParcial As Integer
Private mvariNumParcelas As Integer

Public Property Let iNumParcelas(ByVal vData As Integer)
    mvariNumParcelas = vData
End Property

Public Property Get iNumParcelas() As Integer
    iNumParcelas = mvariNumParcelas
End Property

Public Property Let iMoedaAporte(ByVal vData As Integer)
    mvariMoedaAporte = vData
End Property

Public Property Get iMoedaAporte() As Integer
    iMoedaAporte = mvariMoedaAporte
End Property

Public Property Let dValorAporteUSS(ByVal vData As Double)
    mvardValorAporteUSS = vData
End Property

Public Property Get dValorAporteUSS() As Double
    dValorAporteUSS = mvardValorAporteUSS
End Property

Public Property Let lNumIntDocOCR(ByVal vData As Long)
    mvarlNumIntDocOCR = vData
End Property

Public Property Get lNumIntDocOCR() As Long
    lNumIntDocOCR = mvarlNumIntDocOCR
End Property

Public Property Let iTipoPagtoAporte(ByVal vData As Integer)
    mvariTipoPagtoAporte = vData
End Property

Public Property Get iTipoPagtoAporte() As Integer
    iTipoPagtoAporte = mvariTipoPagtoAporte
End Property

Public Property Let lNumIntDocPagtoAporte(ByVal vData As Long)
    mvarlNumIntDocPagtoAporte = vData
End Property

Public Property Get lNumIntDocPagtoAporte() As Long
    lNumIntDocPagtoAporte = mvarlNumIntDocPagtoAporte
End Property

Public Property Let lCodigoAporte(ByVal vData As Long)
    mvarlCodigoAporte = vData
End Property

Public Property Get lCodigoAporte() As Long
    lCodigoAporte = mvarlCodigoAporte
End Property

Public Property Let iParcelaAporte(ByVal vData As Integer)
    mvariParcelaAporte = vData
End Property

Public Property Get iParcelaAporte() As Integer
    iParcelaAporte = mvariParcelaAporte
End Property

Public Property Let dValorAporte(ByVal vData As Double)
    mvardValorAporte = vData
End Property

Public Property Get dValorAporte() As Double
    dValorAporte = mvardValorAporte
End Property
