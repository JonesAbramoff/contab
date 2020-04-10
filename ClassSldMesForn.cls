VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSldMesForn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlFornecedor As Long
Private mvariFilialForn As Integer
Private mvariAno As Integer
Private mvarsProduto As String
Private mvaradQuantCompras(1 To 12) As Double
Private mvaradValorCompras(1 To 12) As Double

Public Property Let adValorCompras(ByVal iIndice As Integer, ByVal vData As Double)
    mvaradValorCompras(iIndice) = vData
End Property

Public Property Get adValorCompras(ByVal iIndice As Integer) As Double
    adValorCompras = mvaradValorCompras(iIndice)
End Property

Public Property Let adQuantCompras(ByVal iIndice As Integer, ByVal vData As Double)
    mvaradQuantCompras(iIndice) = vData
End Property

Public Property Get adQuantCompras(ByVal iIndice As Integer) As Double
    adQuantCompras = mvaradQuantCompras(iIndice)
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
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


