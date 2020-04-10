VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGarantiaNumSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntGarantia As Long
Private mvarsNumSerie As String

Public Property Let sNumSerie(ByVal vData As String)
    mvarsNumSerie = vData
End Property

Public Property Get sNumSerie() As String
    sNumSerie = mvarsNumSerie
End Property

Public Property Let lNumIntGarantia(ByVal vData As Long)
    mvarlNumIntGarantia = vData
End Property

Public Property Get lNumIntGarantia() As Long
    lNumIntGarantia = mvarlNumIntGarantia
End Property



