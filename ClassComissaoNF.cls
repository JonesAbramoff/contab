VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassComissaoNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariCodVendedor As Integer
Private mvardValorBase As Double
Private mvardPercentual As Double
Private mvardValor As Double
Private mvardPercentualEmissao As Double
Private mvardValorEmissao As Double
Private mvariIndireta As Integer
Public iSeq As Integer

Public iLinha As Integer 'linha do grid de itens, se houver
Public colLinhas As New Collection

Public Property Let iIndireta(ByVal vData As Integer)
    mvariIndireta = vData
End Property

Public Property Get iIndireta() As Integer
    iIndireta = mvariIndireta
End Property

Public Property Let dValorEmissao(ByVal vData As Double)
    mvardValorEmissao = vData
End Property

Public Property Get dValorEmissao() As Double
    dValorEmissao = mvardValorEmissao
End Property

Public Property Let dPercentualEmissao(ByVal vData As Double)
    mvardPercentualEmissao = vData
End Property

Public Property Get dPercentualEmissao() As Double
    dPercentualEmissao = mvardPercentualEmissao
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dPercentual(ByVal vData As Double)
    mvardPercentual = vData
End Property

Public Property Get dPercentual() As Double
    dPercentual = mvardPercentual
End Property

Public Property Let dValorBase(ByVal vData As Double)
    mvardValorBase = vData
End Property

Public Property Get dValorBase() As Double
    dValorBase = mvardValorBase
End Property

Public Property Let iCodVendedor(ByVal vData As Integer)
    mvariCodVendedor = vData
End Property

Public Property Get iCodVendedor() As Integer
    iCodVendedor = mvariCodVendedor
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property



