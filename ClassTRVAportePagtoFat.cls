VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVAportePagtoFat"
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
Private mvardValor As Double
Private mvardtValidadeDe As Date
Private mvardtValidadeAte As Date
Private mvardSaldo As Double
Private mvardPercentual As Double
Private mvariSeq As Integer

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
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

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtValidadeDe(ByVal vData As Date)
    mvardtValidadeDe = vData
End Property

Public Property Get dtValidadeDe() As Date
    dtValidadeDe = mvardtValidadeDe
End Property

Public Property Let dtValidadeAte(ByVal vData As Date)
    mvardtValidadeAte = vData
End Property

Public Property Get dtValidadeAte() As Date
    dtValidadeAte = mvardtValidadeAte
End Property

Public Property Let dSaldo(ByVal vData As Double)
    mvardSaldo = vData
End Property

Public Property Get dSaldo() As Double
    dSaldo = mvardSaldo
End Property

Public Property Let dPercentual(ByVal vData As Double)
    mvardPercentual = vData
End Property

Public Property Get dPercentual() As Double
    dPercentual = mvardPercentual
End Property
