VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSldMesFat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariAno As Integer
Private mvariFilialEmpresa As Integer
Private mvarsProduto As String
Private mvardQuantFaturada(1 To 12) As Double
Private mvardValorFaturado(1 To 12) As Double
Private mvardQuantDevolvida(1 To 12) As Double
Private mvardValorDevolvido(1 To 12) As Double
Private mvardTotalDescontos(1 To 12) As Double
Private mvardQuantPedida(1 To 12) As Double
Private mvardQuantPedidaSRV(1 To 12) As Double

Public Property Let dQuantPedidaSRV(ByVal iIndice As Integer, ByVal vData As Double)
    mvardQuantPedidaSRV(iIndice) = vData
End Property

Public Property Get dQuantPedidaSRV(ByVal iIndice As Integer) As Double
    dQuantPedidaSRV = mvardQuantPedidaSRV(iIndice)
End Property

Public Property Let dQuantPedida(ByVal iIndice As Integer, ByVal vData As Double)
    mvardQuantPedida(iIndice) = vData
End Property

Public Property Get dQuantPedida(ByVal iIndice As Integer) As Double
    dQuantPedida = mvardQuantPedida(iIndice)
End Property

Public Property Let dTotalDescontos(ByVal iIndice As Integer, ByVal vData As Double)
    mvardTotalDescontos(iIndice) = vData
End Property

Public Property Get dTotalDescontos(ByVal iIndice As Integer) As Double
    dTotalDescontos = mvardTotalDescontos(iIndice)
End Property

Public Property Let dValorDevolvido(ByVal iIndice As Integer, ByVal vData As Double)
    mvardValorDevolvido(iIndice) = vData
End Property

Public Property Get dValorDevolvido(ByVal iIndice As Integer) As Double
    dValorDevolvido = mvardValorDevolvido(iIndice)
End Property

Public Property Let dQuantDevolvida(ByVal iIndice As Integer, ByVal vData As Double)
    mvardQuantDevolvida(iIndice) = vData
End Property

Public Property Get dQuantDevolvida(ByVal iIndice As Integer) As Double
    dQuantDevolvida = mvardQuantDevolvida(iIndice)
End Property

Public Property Let dValorFaturado(ByVal iIndice As Integer, ByVal vData As Double)
    mvardValorFaturado(iIndice) = vData
End Property

Public Property Get dValorFaturado(ByVal iIndice As Integer) As Double
    dValorFaturado = mvardValorFaturado(iIndice)
End Property

Public Property Let dQuantFaturada(ByVal iIndice As Integer, ByVal vData As Double)
    mvardQuantFaturada(iIndice) = vData
End Property

Public Property Get dQuantFaturada(ByVal iIndice As Integer) As Double
    dQuantFaturada = mvardQuantFaturada(iIndice)
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

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

