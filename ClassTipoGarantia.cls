VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlCodigo As Long
Private mvarsDescricao As String
Private mvariPrazoPadrao As Integer
Private mvariGarantiaTotal As Integer
Private mvariGarantiaTotalPrazo As Integer
Private mvarcolTipoGarantiaProduto As New Collection
Private mvarobjTela As Object

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Set colTipoGarantiaProduto(ByVal vData As Collection)
    Set mvarcolTipoGarantiaProduto = vData
End Property

Public Property Get colTipoGarantiaProduto() As Collection
    Set colTipoGarantiaProduto = mvarcolTipoGarantiaProduto
End Property

Public Property Let iGarantiaTotalPrazo(ByVal vData As Integer)
    mvariGarantiaTotalPrazo = vData
End Property

Public Property Get iGarantiaTotalPrazo() As Integer
    iGarantiaTotalPrazo = mvariGarantiaTotalPrazo
End Property

Public Property Let iGarantiaTotal(ByVal vData As Integer)
    mvariGarantiaTotal = vData
End Property

Public Property Get iGarantiaTotal() As Integer
    iGarantiaTotal = mvariGarantiaTotal
End Property

Public Property Let iPrazoPadrao(ByVal vData As Integer)
    mvariPrazoPadrao = vData
End Property

Public Property Get iPrazoPadrao() As Integer
    iPrazoPadrao = mvariPrazoPadrao
End Property


Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property



