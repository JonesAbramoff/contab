VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdSolicSRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private mvarlNumIntDoc As Long
Private mvarlNumIntItensOrcSRV As Long
Private mvardQuantidade As Double
Private mvarsProduto As String
Private mvarsServicoOrcSRV As String
Private mvarsLote As String
Private mvariFilialOP As Integer
Private mvarlGarantia As Long
Private mvarsContrato As String
Private mvarcolProdutoSRV As New Collection
'Private mvarlNumIntItensContratoSRV As Long
'Private mvarlNumIntGarantia As Long

'Public Property Let lNumIntGarantia(ByVal vData As Long)
'    mvarlNumIntGarantia = vData
'End Property
'
'Public Property Get lNumIntGarantia() As Long
'    lNumIntGarantia = mvarlNumIntGarantia
'End Property
'
'Public Property Let lNumIntItensContratoSRV(ByVal vData As Long)
'    mvarlNumIntItensContratoSRV = vData
'End Property
'
'Public Property Get lNumIntItensContratoSRV() As Long
'    lNumIntItensContratoSRV = mvarlNumIntItensContratoSRV
'End Property

Public Property Set colProdutoSRV(ByVal vData As Collection)
    Set mvarcolProdutoSRV = vData
End Property

Public Property Get colProdutoSRV() As Collection
    Set colProdutoSRV = mvarcolProdutoSRV
End Property

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

Public Property Let iFilialOP(ByVal vData As Integer)
    mvariFilialOP = vData
End Property

Public Property Get iFilialOP() As Integer
    iFilialOP = mvariFilialOP
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let sServicoOrcSRV(ByVal vData As String)
    mvarsServicoOrcSRV = vData
End Property

Public Property Get sServicoOrcSRV() As String
    sServicoOrcSRV = mvarsServicoOrcSRV
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let lNumIntItensOrcSRV(ByVal vData As Long)
    mvarlNumIntItensOrcSRV = vData
End Property

Public Property Get lNumIntItensOrcSRV() As Long
    lNumIntItensOrcSRV = mvarlNumIntItensOrcSRV
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

