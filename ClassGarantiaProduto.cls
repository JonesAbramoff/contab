VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGarantiaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntGarantia As Long
Private mvarsProduto As String
Private mvariPrazo As Integer

Public Property Let iPrazo(ByVal vData As Integer)
    mvariPrazo = vData
End Property

Public Property Get iPrazo() As Integer
    iPrazo = mvariPrazo
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lNumIntGarantia(ByVal vData As Long)
    mvarlNumIntGarantia = vData
End Property

Public Property Get lNumIntGarantia() As Long
    lNumIntGarantia = mvarlNumIntGarantia
End Property


