VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPAportePagtoCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocAporte As Long
Private mvariBase As Integer
Private mvariStatus As Integer
Private mvardtDataPagto As Date
Private mvarlNumIntDocDestino As Long
Private mvariFormaPagto As Integer
Private mvariTipoDocDestino As Integer
Private mvardPercentual As Double
Private mvardValor As Double
Private mvariSeq As Integer

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocAporte(ByVal vData As Long)
    mvarlNumIntDocAporte = vData
End Property

Public Property Get lNumIntDocAporte() As Long
    lNumIntDocAporte = mvarlNumIntDocAporte
End Property

Public Property Let iBase(ByVal vData As Integer)
    mvariBase = vData
End Property

Public Property Get iBase() As Integer
    iBase = mvariBase
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let dtDataPagto(ByVal vData As Date)
    mvardtDataPagto = vData
End Property

Public Property Get dtDataPagto() As Date
    dtDataPagto = mvardtDataPagto
End Property

Public Property Let lNumIntDocDestino(ByVal vData As Long)
    mvarlNumIntDocDestino = vData
End Property

Public Property Get lNumIntDocDestino() As Long
    lNumIntDocDestino = mvarlNumIntDocDestino
End Property

Public Property Let iFormaPagto(ByVal vData As Integer)
    mvariFormaPagto = vData
End Property

Public Property Get iFormaPagto() As Integer
    iFormaPagto = mvariFormaPagto
End Property

Public Property Let iTipoDocDestino(ByVal vData As Integer)
    mvariTipoDocDestino = vData
End Property

Public Property Get iTipoDocDestino() As Integer
    iTipoDocDestino = mvariTipoDocDestino
End Property

Public Property Let dPercentual(ByVal vData As Double)
    mvardPercentual = vData
End Property

Public Property Get dPercentual() As Double
    dPercentual = mvardPercentual
End Property
