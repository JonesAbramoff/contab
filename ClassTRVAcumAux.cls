VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVAcumAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntInfoGer As Long
Private mvarlNumIntRelComiInt As Long
Private mvardAcumValorBase As Double
Private mvardPercComissao As Double

Public Property Let dPercComissao(ByVal vData As Double)
    mvardPercComissao = vData
End Property

Public Property Get dPercComissao() As Double
    dPercComissao = mvardPercComissao
End Property

Public Property Let dAcumValorBase(ByVal vData As Double)
    mvardAcumValorBase = vData
End Property

Public Property Get dAcumValorBase() As Double
    dAcumValorBase = mvardAcumValorBase
End Property

Public Property Let lNumIntRelComiInt(ByVal vData As Long)
    mvarlNumIntRelComiInt = vData
End Property

Public Property Get lNumIntRelComiInt() As Long
    lNumIntRelComiInt = mvarlNumIntRelComiInt
End Property

Public Property Let lNumIntInfoGer(ByVal vData As Long)
    mvarlNumIntInfoGer = vData
End Property

Public Property Get lNumIntInfoGer() As Long
    lNumIntInfoGer = mvarlNumIntInfoGer
End Property


