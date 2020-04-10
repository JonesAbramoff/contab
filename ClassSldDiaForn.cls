VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSldDiaForn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlFornecedor As Long
Private mvariFilialForn As Integer
Private mvarsProduto As String
Private mvardtData As Date
Private mvardQuantCompra As Double
Private mvardValorCompra As Double

Public Property Let dValorCompra(ByVal vData As Double)
    mvardValorCompra = vData
End Property

Public Property Get dValorCompra() As Double
    dValorCompra = mvardValorCompra
End Property

Public Property Let dQuantCompra(ByVal vData As Double)
    mvardQuantCompra = vData
End Property

Public Property Get dQuantCompra() As Double
    dQuantCompra = mvardQuantCompra
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

Public Property Let iFilialForn(ByVal vData As Integer)
    mvariFilialForn = vData
End Property

Public Property Get iFilialForn() As Integer
    iFilialForn = mvariFilialForn
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

