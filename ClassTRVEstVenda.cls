VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVEstVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvardtDataFatDe As Date
Private mvardtDataFatAte As Date
Private mvardtDataEmiDe As Date
Private mvardtDataEmiAte As Date
Private mvarlClienteDe As Long
Private mvarlClienteAte As Long
Private mvardValorBrutoVouRS As Double
Private mvardValorBrutoVouRSCanc As Double
Private mvardValorBrutoVouRSFat As Double
Private mvardValorBrutoVouUSS As Double
Private mvardValorBrutoVouUSSCanc As Double
Private mvardValorBrutoVouUSSFat As Double
Private mvardValorFatVou As Double
Private mvardValorFatVouCanc As Double
Private mvardValorFatVouFat As Double
Private mvardQtdeEmiVou As Double
Private mvardQtdeEmiVouCanc As Double
Private mvardQtdeEmiVouFat As Double
Private mvardQtdeEmiPax As Double
Private mvardQtdeEmiPaxCanc As Double
Private mvardQtdeEmiPaxFat As Double
Private mvardQtdeFatCR As Double
Private mvardQtdeFatCP As Double
Private mvardValorFatCR As Double
Private mvardValorFatCP As Double
Private mvardQtdeFatCRCanc As Double
Private mvardQtdeFatCPCanc As Double
Private mvardValorFatCRCanc As Double
Private mvardValorFatCPCanc As Double

Public Property Let dValorFatCPCanc(ByVal vData As Double)
    mvardValorFatCPCanc = vData
End Property

Public Property Get dValorFatCPCanc() As Double
    dValorFatCPCanc = mvardValorFatCPCanc
End Property

Public Property Let dValorFatCRCanc(ByVal vData As Double)
    mvardValorFatCRCanc = vData
End Property

Public Property Get dValorFatCRCanc() As Double
    dValorFatCRCanc = mvardValorFatCRCanc
End Property

Public Property Let dQtdeFatCPCanc(ByVal vData As Double)
    mvardQtdeFatCPCanc = vData
End Property

Public Property Get dQtdeFatCPCanc() As Double
    dQtdeFatCPCanc = mvardQtdeFatCPCanc
End Property

Public Property Let dQtdeFatCRCanc(ByVal vData As Double)
    mvardQtdeFatCRCanc = vData
End Property

Public Property Get dQtdeFatCRCanc() As Double
    dQtdeFatCRCanc = mvardQtdeFatCRCanc
End Property

Public Property Let dValorFatCP(ByVal vData As Double)
    mvardValorFatCP = vData
End Property

Public Property Get dValorFatCP() As Double
    dValorFatCP = mvardValorFatCP
End Property

Public Property Let dValorFatCR(ByVal vData As Double)
    mvardValorFatCR = vData
End Property

Public Property Get dValorFatCR() As Double
    dValorFatCR = mvardValorFatCR
End Property

Public Property Let dQtdeFatCP(ByVal vData As Double)
    mvardQtdeFatCP = vData
End Property

Public Property Get dQtdeFatCP() As Double
    dQtdeFatCP = mvardQtdeFatCP
End Property

Public Property Let dQtdeFatCR(ByVal vData As Double)
    mvardQtdeFatCR = vData
End Property

Public Property Get dQtdeFatCR() As Double
    dQtdeFatCR = mvardQtdeFatCR
End Property

Public Property Let dQtdeEmiPaxFat(ByVal vData As Double)
    mvardQtdeEmiPaxFat = vData
End Property

Public Property Get dQtdeEmiPaxFat() As Double
    dQtdeEmiPaxFat = mvardQtdeEmiPaxFat
End Property

Public Property Let dQtdeEmiPaxCanc(ByVal vData As Double)
    mvardQtdeEmiPaxCanc = vData
End Property

Public Property Get dQtdeEmiPaxCanc() As Double
    dQtdeEmiPaxCanc = mvardQtdeEmiPaxCanc
End Property

Public Property Let dQtdeEmiPax(ByVal vData As Double)
    mvardQtdeEmiPax = vData
End Property

Public Property Get dQtdeEmiPax() As Double
    dQtdeEmiPax = mvardQtdeEmiPax
End Property

Public Property Let dQtdeEmiVouFat(ByVal vData As Double)
    mvardQtdeEmiVouFat = vData
End Property

Public Property Get dQtdeEmiVouFat() As Double
    dQtdeEmiVouFat = mvardQtdeEmiVouFat
End Property

Public Property Let dQtdeEmiVouCanc(ByVal vData As Double)
    mvardQtdeEmiVouCanc = vData
End Property

Public Property Get dQtdeEmiVouCanc() As Double
    dQtdeEmiVouCanc = mvardQtdeEmiVouCanc
End Property

Public Property Let dQtdeEmiVou(ByVal vData As Double)
    mvardQtdeEmiVou = vData
End Property

Public Property Get dQtdeEmiVou() As Double
    dQtdeEmiVou = mvardQtdeEmiVou
End Property

Public Property Let dValorFatVouFat(ByVal vData As Double)
    mvardValorFatVouFat = vData
End Property

Public Property Get dValorFatVouFat() As Double
    dValorFatVouFat = mvardValorFatVouFat
End Property

Public Property Let dValorFatVouCanc(ByVal vData As Double)
    mvardValorFatVouCanc = vData
End Property

Public Property Get dValorFatVouCanc() As Double
    dValorFatVouCanc = mvardValorFatVouCanc
End Property

Public Property Let dValorFatVou(ByVal vData As Double)
    mvardValorFatVou = vData
End Property

Public Property Get dValorFatVou() As Double
    dValorFatVou = mvardValorFatVou
End Property

Public Property Let dValorBrutoVouUSSFat(ByVal vData As Double)
    mvardValorBrutoVouUSSFat = vData
End Property

Public Property Get dValorBrutoVouUSSFat() As Double
    dValorBrutoVouUSSFat = mvardValorBrutoVouUSSFat
End Property

Public Property Let dValorBrutoVouUSSCanc(ByVal vData As Double)
    mvardValorBrutoVouUSSCanc = vData
End Property

Public Property Get dValorBrutoVouUSSCanc() As Double
    dValorBrutoVouUSSCanc = mvardValorBrutoVouUSSCanc
End Property

Public Property Let dValorBrutoVouUSS(ByVal vData As Double)
    mvardValorBrutoVouUSS = vData
End Property

Public Property Get dValorBrutoVouUSS() As Double
    dValorBrutoVouUSS = mvardValorBrutoVouUSS
End Property

Public Property Let dValorBrutoVouRSFat(ByVal vData As Double)
    mvardValorBrutoVouRSFat = vData
End Property

Public Property Get dValorBrutoVouRSFat() As Double
    dValorBrutoVouRSFat = mvardValorBrutoVouRSFat
End Property

Public Property Let dValorBrutoVouRSCanc(ByVal vData As Double)
    mvardValorBrutoVouRSCanc = vData
End Property

Public Property Get dValorBrutoVouRSCanc() As Double
    dValorBrutoVouRSCanc = mvardValorBrutoVouRSCanc
End Property

Public Property Let dValorBrutoVouRS(ByVal vData As Double)
    mvardValorBrutoVouRS = vData
End Property

Public Property Get dValorBrutoVouRS() As Double
    dValorBrutoVouRS = mvardValorBrutoVouRS
End Property

Public Property Let lClienteAte(ByVal vData As Long)
    mvarlClienteAte = vData
End Property

Public Property Get lClienteAte() As Long
    lClienteAte = mvarlClienteAte
End Property

Public Property Let lClienteDe(ByVal vData As Long)
    mvarlClienteDe = vData
End Property

Public Property Get lClienteDe() As Long
    lClienteDe = mvarlClienteDe
End Property

Public Property Let dtDataFatDe(ByVal vData As Date)
    mvardtDataFatDe = vData
End Property

Public Property Get dtDataFatDe() As Date
    dtDataFatDe = mvardtDataFatDe
End Property

Public Property Let dtDataFatAte(ByVal vData As Date)
    mvardtDataFatAte = vData
End Property

Public Property Get dtDataFatAte() As Date
    dtDataFatAte = mvardtDataFatAte
End Property

Public Property Let dtDataEmiDe(ByVal vData As Date)
    mvardtDataEmiDe = vData
End Property

Public Property Get dtDataEmiDe() As Date
    dtDataEmiDe = mvardtDataEmiDe
End Property

Public Property Let dtDataEmiAte(ByVal vData As Date)
    mvardtDataEmiAte = vData
End Property

Public Property Get dtDataEmiAte() As Date
    dtDataEmiAte = mvardtDataEmiAte
End Property
