VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassComissoesCF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlNumIntDoc As Long
Private mvariCodVendedor As Integer
Private mvardValorBase As Double
Private mvardValorComissao As Double
Private mvariIndireta As Integer

Public Property Let iIndireta(ByVal vData As Integer)
    mvariIndireta = vData
End Property

Public Property Get iIndireta() As Integer
    iIndireta = mvariIndireta
End Property

Public Property Let dValorComissao(ByVal vData As Double)
    mvardValorComissao = vData
End Property

Public Property Get dValorComissao() As Double
    dValorComissao = mvardValorComissao
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

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property


