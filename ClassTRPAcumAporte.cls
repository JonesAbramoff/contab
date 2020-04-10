VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPAcumAporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardValorDe As Double
Private mvardValorAte As Double
Private mvardValorAcumAporte As Double
Private mvardValorDolarAcumAporte As Double
Private mvardPercentual As Double
Private mvariMoeda As Integer
Private mvardCambioMedio As Double

Public Property Let dCambioMedio(ByVal vData As Double)
    mvardCambioMedio = vData
End Property

Public Property Get dCambioMedio() As Double
    dCambioMedio = mvardCambioMedio
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let dPercentual(ByVal vData As Double)
    mvardPercentual = vData
End Property

Public Property Get dPercentual() As Double
    dPercentual = mvardPercentual
End Property

Public Property Let dValorDolarAcumAporte(ByVal vData As Double)
    mvardValorDolarAcumAporte = vData
End Property

Public Property Get dValorDolarAcumAporte() As Double
    dValorDolarAcumAporte = mvardValorDolarAcumAporte
End Property

Public Property Let dValorAcumAporte(ByVal vData As Double)
    mvardValorAcumAporte = vData
End Property

Public Property Get dValorAcumAporte() As Double
    dValorAcumAporte = mvardValorAcumAporte
End Property

Public Property Let dValorAte(ByVal vData As Double)
    mvardValorAte = vData
End Property

Public Property Get dValorAte() As Double
    dValorAte = mvardValorAte
End Property

Public Property Let dValorDe(ByVal vData As Double)
    mvardValorDe = vData
End Property

Public Property Get dValorDe() As Double
    dValorDe = mvardValorDe
End Property

