VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassRelTRVDesviosVend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel AS Long
Private mvarlSeq AS Long
Private mvarlCliente AS Long
Private mvarsNomeCliente AS String
Private mvardValorFatNoMes AS Double
Private mvardValorBrutoNoMes AS Double
Private mvardValorLiqNoMes AS Double
Private mvardValorFatMesAnt AS Double
Private mvardValorBrutoMesAnt AS Double
Private mvardValorLiqMesAnt AS Double
Private mvardValorFatMesAnoAnt AS Double
Private mvardValorBrutoMesAnoAnt AS Double
Private mvardValorLiqMesAnoAnt AS Double
Private mvardTotalValorFat AS Double
Private mvardTotalValorBruto AS Double
Private mvardTotalValorLiq AS Double
Private mvardValorFatMedio AS Double
Private mvardValorBrutoMedio AS Double
Private mvardValorLiqMedio AS Double
Private mvardDesvioValorMes AS Double
Private mvardDesvioValorAno AS Double
Private mvarlQtdVouNoMes AS Long
Private mvarlQtdVouNoMesAnt AS Long
Private mvarlQtdVouNoMesAnoAnt AS Long
Private mvardDesvioQtdMes AS Double
Private mvardDesvioQtdAno AS Double
Private mvarlTotalQtdVou AS Long
Private mvardQtdVouMedio AS Double
Private mvardtPrimeiraCompra AS Date
Private mvardtUltimaCompra AS Date
Private mvarsRespSetor AS String
Private mvarsRespFunc AS String

Public Property Let lNumIntRel (ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel () AS Long
    lNumIntRel= mvarlNumIntRel
End Property

Public Property Let lSeq (ByVal vData As Long)
    mvarlSeq = vData
End Property

Public Property Get lSeq () AS Long
    lSeq= mvarlSeq
End Property

Public Property Let lCliente (ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente () AS Long
    lCliente= mvarlCliente
End Property

Public Property Let sNomeCliente (ByVal vData As String)
    mvarsNomeCliente = vData
End Property

Public Property Get sNomeCliente () AS String
    sNomeCliente= mvarsNomeCliente
End Property

Public Property Let dValorFatNoMes (ByVal vData As Double)
    mvardValorFatNoMes = vData
End Property

Public Property Get dValorFatNoMes () AS Double
    dValorFatNoMes= mvardValorFatNoMes
End Property

Public Property Let dValorBrutoNoMes (ByVal vData As Double)
    mvardValorBrutoNoMes = vData
End Property

Public Property Get dValorBrutoNoMes () AS Double
    dValorBrutoNoMes= mvardValorBrutoNoMes
End Property

Public Property Let dValorLiqNoMes (ByVal vData As Double)
    mvardValorLiqNoMes = vData
End Property

Public Property Get dValorLiqNoMes () AS Double
    dValorLiqNoMes= mvardValorLiqNoMes
End Property

Public Property Let dValorFatMesAnt (ByVal vData As Double)
    mvardValorFatMesAnt = vData
End Property

Public Property Get dValorFatMesAnt () AS Double
    dValorFatMesAnt= mvardValorFatMesAnt
End Property

Public Property Let dValorBrutoMesAnt (ByVal vData As Double)
    mvardValorBrutoMesAnt = vData
End Property

Public Property Get dValorBrutoMesAnt () AS Double
    dValorBrutoMesAnt= mvardValorBrutoMesAnt
End Property

Public Property Let dValorLiqMesAnt (ByVal vData As Double)
    mvardValorLiqMesAnt = vData
End Property

Public Property Get dValorLiqMesAnt () AS Double
    dValorLiqMesAnt= mvardValorLiqMesAnt
End Property

Public Property Let dValorFatMesAnoAnt (ByVal vData As Double)
    mvardValorFatMesAnoAnt = vData
End Property

Public Property Get dValorFatMesAnoAnt () AS Double
    dValorFatMesAnoAnt= mvardValorFatMesAnoAnt
End Property

Public Property Let dValorBrutoMesAnoAnt (ByVal vData As Double)
    mvardValorBrutoMesAnoAnt = vData
End Property

Public Property Get dValorBrutoMesAnoAnt () AS Double
    dValorBrutoMesAnoAnt= mvardValorBrutoMesAnoAnt
End Property

Public Property Let dValorLiqMesAnoAnt (ByVal vData As Double)
    mvardValorLiqMesAnoAnt = vData
End Property

Public Property Get dValorLiqMesAnoAnt () AS Double
    dValorLiqMesAnoAnt= mvardValorLiqMesAnoAnt
End Property

Public Property Let dTotalValorFat (ByVal vData As Double)
    mvardTotalValorFat = vData
End Property

Public Property Get dTotalValorFat () AS Double
    dTotalValorFat= mvardTotalValorFat
End Property

Public Property Let dTotalValorBruto (ByVal vData As Double)
    mvardTotalValorBruto = vData
End Property

Public Property Get dTotalValorBruto () AS Double
    dTotalValorBruto= mvardTotalValorBruto
End Property

Public Property Let dTotalValorLiq (ByVal vData As Double)
    mvardTotalValorLiq = vData
End Property

Public Property Get dTotalValorLiq () AS Double
    dTotalValorLiq= mvardTotalValorLiq
End Property

Public Property Let dValorFatMedio (ByVal vData As Double)
    mvardValorFatMedio = vData
End Property

Public Property Get dValorFatMedio () AS Double
    dValorFatMedio= mvardValorFatMedio
End Property

Public Property Let dValorBrutoMedio (ByVal vData As Double)
    mvardValorBrutoMedio = vData
End Property

Public Property Get dValorBrutoMedio () AS Double
    dValorBrutoMedio= mvardValorBrutoMedio
End Property

Public Property Let dValorLiqMedio (ByVal vData As Double)
    mvardValorLiqMedio = vData
End Property

Public Property Get dValorLiqMedio () AS Double
    dValorLiqMedio= mvardValorLiqMedio
End Property

Public Property Let dDesvioValorMes (ByVal vData As Double)
    mvardDesvioValorMes = vData
End Property

Public Property Get dDesvioValorMes () AS Double
    dDesvioValorMes= mvardDesvioValorMes
End Property

Public Property Let dDesvioValorAno (ByVal vData As Double)
    mvardDesvioValorAno = vData
End Property

Public Property Get dDesvioValorAno () AS Double
    dDesvioValorAno= mvardDesvioValorAno
End Property

Public Property Let lQtdVouNoMes (ByVal vData As Long)
    mvarlQtdVouNoMes = vData
End Property

Public Property Get lQtdVouNoMes () AS Long
    lQtdVouNoMes= mvarlQtdVouNoMes
End Property

Public Property Let lQtdVouNoMesAnt (ByVal vData As Long)
    mvarlQtdVouNoMesAnt = vData
End Property

Public Property Get lQtdVouNoMesAnt () AS Long
    lQtdVouNoMesAnt= mvarlQtdVouNoMesAnt
End Property

Public Property Let lQtdVouNoMesAnoAnt (ByVal vData As Long)
    mvarlQtdVouNoMesAnoAnt = vData
End Property

Public Property Get lQtdVouNoMesAnoAnt () AS Long
    lQtdVouNoMesAnoAnt= mvarlQtdVouNoMesAnoAnt
End Property

Public Property Let dDesvioQtdMes (ByVal vData As Double)
    mvardDesvioQtdMes = vData
End Property

Public Property Get dDesvioQtdMes () AS Double
    dDesvioQtdMes= mvardDesvioQtdMes
End Property

Public Property Let dDesvioQtdAno (ByVal vData As Double)
    mvardDesvioQtdAno = vData
End Property

Public Property Get dDesvioQtdAno () AS Double
    dDesvioQtdAno= mvardDesvioQtdAno
End Property

Public Property Let lTotalQtdVou (ByVal vData As Long)
    mvarlTotalQtdVou = vData
End Property

Public Property Get lTotalQtdVou () AS Long
    lTotalQtdVou= mvarlTotalQtdVou
End Property

Public Property Let dQtdVouMedio (ByVal vData As Double)
    mvardQtdVouMedio = vData
End Property

Public Property Get dQtdVouMedio () AS Double
    dQtdVouMedio= mvardQtdVouMedio
End Property

Public Property Let dtPrimeiraCompra (ByVal vData As Date)
    mvardtPrimeiraCompra = vData
End Property

Public Property Get dtPrimeiraCompra () AS Date
    dtPrimeiraCompra= mvardtPrimeiraCompra
End Property

Public Property Let dtUltimaCompra (ByVal vData As Date)
    mvardtUltimaCompra = vData
End Property

Public Property Get dtUltimaCompra () AS Date
    dtUltimaCompra= mvardtUltimaCompra
End Property

Public Property Let sRespSetor (ByVal vData As String)
    mvarsRespSetor = vData
End Property

Public Property Get sRespSetor () AS String
    sRespSetor= mvarsRespSetor
End Property

Public Property Let sRespFunc (ByVal vData As String)
    mvarsRespFunc = vData
End Property

Public Property Get sRespFunc () AS String
    sRespFunc= mvarsRespFunc
End Property
