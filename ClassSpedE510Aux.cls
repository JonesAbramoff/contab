VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSpedE510Aux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsCFOP As String
Private mvariCSTIPI As Integer
Private mvardValorItemNF As Double
Private mvardValorBCIPI As Double
Private mvardValorIPI As Double

Public Property Let dValorIPI(ByVal vData As Double)
    mvardValorIPI = vData
End Property

Public Property Get dValorIPI() As Double
    dValorIPI = mvardValorIPI
End Property

Public Property Let dValorBCIPI(ByVal vData As Double)
    mvardValorBCIPI = vData
End Property

Public Property Get dValorBCIPI() As Double
    dValorBCIPI = mvardValorBCIPI
End Property

Public Property Let dValorItemNF(ByVal vData As Double)
    mvardValorItemNF = vData
End Property

Public Property Get dValorItemNF() As Double
    dValorItemNF = mvardValorItemNF
End Property

Public Property Let iCSTIPI(ByVal vData As Integer)
    mvariCSTIPI = vData
End Property

Public Property Get iCSTIPI() As Integer
    iCSTIPI = mvariCSTIPI
End Property

Public Property Let sCFOP(ByVal vData As String)
    mvarsCFOP = vData
End Property

Public Property Get sCFOP() As String
    sCFOP = mvarsCFOP
End Property


