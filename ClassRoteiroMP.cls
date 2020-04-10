VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRoteiroMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsProduto As String
Private mvarsUM As String
Private mvardQuantidade As Double
Private mvardCustoTotal As Double
Private mvardCustoUnitario As Double
Private mvarsVersao As String
Private mvarobjProduto As New ClassProduto

Public Property Set objProduto(ByVal vData As ClassProduto)
    Set mvarobjProduto = vData
End Property

Public Property Get objProduto() As ClassProduto
    Set objProduto = mvarobjProduto
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
End Property

Public Property Let dCustoUnitario(ByVal vData As Double)
    mvardCustoUnitario = vData
End Property

Public Property Get dCustoUnitario() As Double
    dCustoUnitario = mvardCustoUnitario
End Property

Public Property Let dCustoTotal(ByVal vData As Double)
    mvardCustoTotal = vData
End Property

Public Property Get dCustoTotal() As Double
    dCustoTotal = mvardCustoTotal
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property


