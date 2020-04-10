VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSpedC590UF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsCFOP As String
Private mvardValorBCICMSST As Double
Private mvardValorICMSST As Double
Private mvarsUF As String

Public Property Let sUF(ByVal vData As String)
    mvarsUF = vData
End Property

Public Property Get sUF() As String
    sUF = mvarsUF
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

Public Property Let sCFOP(ByVal vData As String)
    mvarsCFOP = vData
End Property

Public Property Get sCFOP() As String
    sCFOP = mvarsCFOP
End Property


