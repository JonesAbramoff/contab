VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSldDiaFat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarsProduto As String
Private mvardtData As Date
Private mvardValorFaturado As Double
Private mvardTotalDescontos As Double
Private mvardQuantPedida As Double
Private mvardQuantFaturada As Double
Private mvardValorDevolvido As Double
Private mvardQuantDevolvida As Double
Private mvardQuantPedidaSRV As Double

Public Property Let dQuantPedidaSRV(ByVal vData As Double)
    mvardQuantPedidaSRV = vData
End Property

Public Property Get dQuantPedidaSRV() As Double
    dQuantPedidaSRV = mvardQuantPedidaSRV
End Property

Public Property Let dQuantDevolvida(ByVal vData As Double)
    mvardQuantDevolvida = vData
End Property

Public Property Get dQuantDevolvida() As Double
    dQuantDevolvida = mvardQuantDevolvida
End Property

Public Property Let dValorDevolvido(ByVal vData As Double)
    mvardValorDevolvido = vData
End Property

Public Property Get dValorDevolvido() As Double
    dValorDevolvido = mvardValorDevolvido
End Property

Public Property Let dQuantFaturada(ByVal vData As Double)
    mvardQuantFaturada = vData
End Property

Public Property Get dQuantFaturada() As Double
    dQuantFaturada = mvardQuantFaturada
End Property

Public Property Let dQuantPedida(ByVal vData As Double)
    mvardQuantPedida = vData
End Property

Public Property Get dQuantPedida() As Double
    dQuantPedida = mvardQuantPedida
End Property

Public Property Let dTotalDescontos(ByVal vData As Double)
    mvardTotalDescontos = vData
End Property

Public Property Get dTotalDescontos() As Double
    dTotalDescontos = mvardTotalDescontos
End Property

Public Property Let dValorFaturado(ByVal vData As Double)
    mvardValorFaturado = vData
End Property

Public Property Get dValorFaturado() As Double
    dValorFaturado = mvardValorFaturado
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

