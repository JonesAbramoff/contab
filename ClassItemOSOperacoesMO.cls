VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOSOperacoesMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDocOper As Long
Private mvarlCodigoMO As Long
Private mvardHoras As Double

Public Property Let dHoras(ByVal vData As Double)
    mvardHoras = vData
End Property

Public Property Get dHoras() As Double
    dHoras = mvardHoras
End Property

Public Property Let lCodigoMO(ByVal vData As Long)
    mvarlCodigoMO = vData
End Property

Public Property Get lCodigoMO() As Long
    lCodigoMO = mvarlCodigoMO
End Property

Public Property Let lNumIntDocOper(ByVal vData As Long)
    mvarlNumIntDocOper = vData
End Property

Public Property Get lNumIntDocOper() As Long
    lNumIntDocOper = mvarlNumIntDocOper
End Property


