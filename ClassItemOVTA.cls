VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOVTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardPrecoBase As Double
Private mvardPercAcresc As Double
Private mvardAcrescimo As Double

Public Property Let dAcrescimo(ByVal vData As Double)
    mvardAcrescimo = vData
End Property

Public Property Get dAcrescimo() As Double
    dAcrescimo = mvardAcrescimo
End Property

Public Property Let dPercAcresc(ByVal vData As Double)
    mvardPercAcresc = vData
End Property

Public Property Get dPercAcresc() As Double
    dPercAcresc = mvardPercAcresc
End Property

Public Property Let dPrecoBase(ByVal vData As Double)
    mvardPrecoBase = vData
End Property

Public Property Get dPrecoBase() As Double
    dPrecoBase = mvardPrecoBase
End Property

