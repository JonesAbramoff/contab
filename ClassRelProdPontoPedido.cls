VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelProdPontoPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntRel As Long
Private mvarsProduto As String
Private mvardQuantEstoque As Double
Private mvardQuantReservada As Double
Private mvardQuantEmpenhada As Double
Private mvardQuantPedVenda As Double
Private mvardQuantPedCompra As Double
Private mvardQuantReqCompra As Double
Private mvardSaldoAtual As Double
Private mvardPontoPedido As Double
Private mvardEstoqueMinimo As Double
Private mvardEstoqueSeguranca As Double

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let dQuantEstoque(ByVal vData As Double)
    mvardQuantEstoque = vData
End Property

Public Property Get dQuantEstoque() As Double
    dQuantEstoque = mvardQuantEstoque
End Property

Public Property Let dQuantReservada(ByVal vData As Double)
    mvardQuantReservada = vData
End Property

Public Property Get dQuantReservada() As Double
    dQuantReservada = mvardQuantReservada
End Property

Public Property Let dQuantEmpenhada(ByVal vData As Double)
    mvardQuantEmpenhada = vData
End Property

Public Property Get dQuantEmpenhada() As Double
    dQuantEmpenhada = mvardQuantEmpenhada
End Property

Public Property Let dQuantPedVenda(ByVal vData As Double)
    mvardQuantPedVenda = vData
End Property

Public Property Get dQuantPedVenda() As Double
    dQuantPedVenda = mvardQuantPedVenda
End Property

Public Property Let dQuantPedCompra(ByVal vData As Double)
    mvardQuantPedCompra = vData
End Property

Public Property Get dQuantPedCompra() As Double
    dQuantPedCompra = mvardQuantPedCompra
End Property

Public Property Let dQuantReqCompra(ByVal vData As Double)
    mvardQuantReqCompra = vData
End Property

Public Property Get dQuantReqCompra() As Double
    dQuantReqCompra = mvardQuantReqCompra
End Property

Public Property Let dSaldoAtual(ByVal vData As Double)
    mvardSaldoAtual = vData
End Property

Public Property Get dSaldoAtual() As Double
    dSaldoAtual = mvardSaldoAtual
End Property

Public Property Let dPontoPedido(ByVal vData As Double)
    mvardPontoPedido = vData
End Property

Public Property Get dPontoPedido() As Double
    dPontoPedido = mvardPontoPedido
End Property

Public Property Let dEstoqueMinimo(ByVal vData As Double)
    mvardEstoqueMinimo = vData
End Property

Public Property Get dEstoqueMinimo() As Double
    dEstoqueMinimo = mvardEstoqueMinimo
End Property

Public Property Let dEstoqueSeguranca(ByVal vData As Double)
    mvardEstoqueSeguranca = vData
End Property

Public Property Get dEstoqueSeguranca() As Double
    dEstoqueSeguranca = mvardEstoqueSeguranca
End Property


