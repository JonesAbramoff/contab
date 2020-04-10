VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdutoSRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntProdSolicSRV As Long
Private mvarsProduto As String
Private mvardQuantidade As Double
Private mvarlGarantia As Long
Private mvarsContrato As String

Public Property Let sContrato(ByVal vData As String)
    mvarsContrato = vData
End Property

Public Property Get sContrato() As String
    sContrato = mvarsContrato
End Property

Public Property Let lGarantia(ByVal vData As Long)
    mvarlGarantia = vData
End Property

Public Property Get lGarantia() As Long
    lGarantia = mvarlGarantia
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lNumIntProdSolicSRV(ByVal vData As Long)
    mvarlNumIntProdSolicSRV = vData
End Property

Public Property Get lNumIntProdSolicSRV() As Long
    lNumIntProdSolicSRV = mvarlNumIntProdSolicSRV
End Property



