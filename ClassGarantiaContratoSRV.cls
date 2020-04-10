VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGarantiaContratoSRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntItensOrcSRV As Long
Private mvarlNumIntItensContratoSRV As Long
Private mvarlNumIntGarantia As Long
Private mvardQuantidade As Double
Private mvarlGarantiaCod As Long
Private mvarsContratoCod As String
Private mvarsServicoPecaSRV As String

Public Property Let sServicoPecaSRV(ByVal vData As String)
    mvarsServicoPecaSRV = vData
End Property

Public Property Get sServicoPecaSRV() As String
    sServicoPecaSRV = mvarsServicoPecaSRV
End Property

Public Property Let sContratoCod(ByVal vData As String)
    mvarsContratoCod = vData
End Property

Public Property Get sContratoCod() As String
    sContratoCod = mvarsContratoCod
End Property

Public Property Let lGarantiaCod(ByVal vData As Long)
    mvarlGarantiaCod = vData
End Property

Public Property Get lGarantiaCod() As Long
    lGarantiaCod = mvarlGarantiaCod
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let lNumIntGarantia(ByVal vData As Long)
    mvarlNumIntGarantia = vData
End Property

Public Property Get lNumIntGarantia() As Long
    lNumIntGarantia = mvarlNumIntGarantia
End Property


Public Property Let lNumIntItensContratoSRV(ByVal vData As Long)
    mvarlNumIntItensContratoSRV = vData
End Property

Public Property Get lNumIntItensContratoSRV() As Long
    lNumIntItensContratoSRV = mvarlNumIntItensContratoSRV
End Property

Public Property Let lNumIntItensOrcSRV(ByVal vData As Long)
    mvarlNumIntItensOrcSRV = vData
End Property

Public Property Get lNumIntItensOrcSRV() As Long
    lNumIntItensOrcSRV = mvarlNumIntItensOrcSRV
End Property



