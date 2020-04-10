VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEvolFat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardtData As Date
Private mvarlNumVendas As Long
Private mvardNumProdutos As Double
Private mvardValorVendas As Double
Private mvariColVlrVenda As Integer

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lNumVendas(ByVal vData As Long)
    mvarlNumVendas = vData
End Property

Public Property Get lNumVendas() As Long
    lNumVendas = mvarlNumVendas
End Property

Public Property Let dNumProdutos(ByVal vData As Double)
    mvardNumProdutos = vData
End Property

Public Property Get dNumProdutos() As Double
    dNumProdutos = mvardNumProdutos
End Property

Public Property Let dValorVendas(ByVal vData As Double)
    mvardValorVendas = vData
End Property

Public Property Get dValorVendas() As Double
    dValorVendas = mvardValorVendas
End Property

Public Property Let iColVlrVenda(ByVal vData As Integer)
    mvariColVlrVenda = vData
End Property

Public Property Get iColVlrVenda() As Integer
    iColVlrVenda = mvariColVlrVenda
End Property

