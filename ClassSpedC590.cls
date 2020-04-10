VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSpedC590"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariCSTICMS As Integer
Private mvarsCFOP As String
Private mvardAliquotaICMS As Double
Private mvardValorOper As Double
Private mvardValorBCICMS As Double
Private mvardValorICMS As Double
Private mvardValorBCICMSST As Double
Private mvardValorICMSST As Double
Private mvardValorRedBC As Double
Private mvarsCodObs As String

Public Property Let sCodObs(ByVal vData As String)
    mvarsCodObs = vData
End Property

Public Property Get sCodObs() As String
    sCodObs = mvarsCodObs
End Property

Public Property Let dValorRedBC(ByVal vData As Double)
    mvardValorRedBC = vData
End Property

Public Property Get dValorRedBC() As Double
    dValorRedBC = mvardValorRedBC
End Property

Public Property Let dValorICMSST(ByVal vData As Double)
    mvardValorICMSST = vData
End Property

Public Property Get dValorICMSST() As Double
    dValorICMSST = mvardValorICMSST
End Property

Public Property Let dValorBCICMSST(ByVal vData As Double)
    mvardValorBCICMSST = vData
End Property

Public Property Get dValorBCICMSST() As Double
    dValorBCICMSST = mvardValorBCICMSST
End Property

Public Property Let dValorICMS(ByVal vData As Double)
    mvardValorICMS = vData
End Property

Public Property Get dValorICMS() As Double
    dValorICMS = mvardValorICMS
End Property

Public Property Let dValorBCICMS(ByVal vData As Double)
    mvardValorBCICMS = vData
End Property

Public Property Get dValorBCICMS() As Double
    dValorBCICMS = mvardValorBCICMS
End Property

Public Property Let dValorOper(ByVal vData As Double)
    mvardValorOper = vData
End Property

Public Property Get dValorOper() As Double
    dValorOper = mvardValorOper
End Property

Public Property Let dAliquotaICMS(ByVal vData As Double)
    mvardAliquotaICMS = vData
End Property

Public Property Get dAliquotaICMS() As Double
    dAliquotaICMS = mvardAliquotaICMS
End Property

Public Property Let sCFOP(ByVal vData As String)
    mvarsCFOP = vData
End Property

Public Property Get sCFOP() As String
    sCFOP = mvarsCFOP
End Property

Public Property Let iCSTICMS(ByVal vData As Integer)
    mvariCSTICMS = vData
End Property

Public Property Get iCSTICMS() As Integer
    iCSTICMS = mvariCSTICMS
End Property


