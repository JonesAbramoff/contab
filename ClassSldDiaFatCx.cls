VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSldDiaFatCx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvariCaixa As Integer
Private mvardtData As Date
Private mvarsProduto As String
Private mvardValorFaturado As Double
Private mvardQuantFaturada As Double
Private mvarlNumVendas As Long

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iCaixa(ByVal vData As Integer)
    mvariCaixa = vData
End Property

Public Property Get iCaixa() As Integer
    iCaixa = mvariCaixa
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dValorFaturado(ByVal vData As Double)
    mvardValorFaturado = vData
End Property

Public Property Get dValorFaturado() As Double
    dValorFaturado = mvardValorFaturado
End Property

Public Property Let dQuantFaturada(ByVal vData As Double)
    mvardQuantFaturada = vData
End Property

Public Property Get dQuantFaturada() As Double
    dQuantFaturada = mvardQuantFaturada
End Property


Public Property Let lNumVendas(ByVal vData As Long)
    mvarlNumVendas = vData
End Property

Public Property Get lNumVendas() As Long
    lNumVendas = mvarlNumVendas
End Property




