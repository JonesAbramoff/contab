VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFDevBenefItemOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsOP As String
Private mvarsProdutoOP As String
Private mvarsUM As String
Private mvardQuantProdDisp As Double
Private mvardQuantSaldo As Double
Private mvarcolInsumos As New Collection
Private mvarlNumIntItemOP As Long
Private mvardFatorProd As Double

Public Property Let dFatorProd(ByVal vData As Double)
    mvardFatorProd = vData
End Property

Public Property Get dFatorProd() As Double
    dFatorProd = mvardFatorProd
End Property

Public Property Let lNumIntItemOP(ByVal vData As Long)
    mvarlNumIntItemOP = vData
End Property

Public Property Get lNumIntItemOP() As Long
    lNumIntItemOP = mvarlNumIntItemOP
End Property

Public Property Set colInsumos(ByVal vData As Object)
    Set mvarcolInsumos = vData
End Property

Public Property Get colInsumos() As Collection
    Set colInsumos = mvarcolInsumos
End Property

Public Property Let dQuantProdDisp(ByVal vData As Double)
    mvardQuantProdDisp = vData
End Property

Public Property Get dQuantProdDisp() As Double
    dQuantProdDisp = mvardQuantProdDisp
End Property

Public Property Let dQuantSaldo(ByVal vData As Double)
    mvardQuantSaldo = vData
End Property

Public Property Get dQuantSaldo() As Double
    dQuantSaldo = mvardQuantSaldo
End Property

Public Property Let sProdutoOP(ByVal vData As String)
    mvarsProdutoOP = vData
End Property

Public Property Get sProdutoOP() As String
    sProdutoOP = mvarsProdutoOP
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let sOP(ByVal vData As String)
    mvarsOP = vData
End Property

Public Property Get sOP() As String
    sOP = mvarsOP
End Property


