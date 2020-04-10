VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelABCComprasTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardtDataDe As Date
Private mvardtDataAte As Date
Private mvariFilialEmpresaDe As Integer
Private mvariFilialEmpresaAte As Integer
Private mvarsProdutoDe As String
Private mvarsProdutoAte As String
Private mvariProdutosTop As Integer
Private mvariTipoProduto As Integer
Private mvarsCategoria As String
Private mvarcolItensCategoria As New Collection

Public Property Let dtDataDe(ByVal vData As Date)
    mvardtDataDe = vData
End Property

Public Property Get dtDataDe() As Date
    dtDataDe = mvardtDataDe
End Property

Public Property Let dtDataAte(ByVal vData As Date)
    mvardtDataAte = vData
End Property

Public Property Get dtDataAte() As Date
    dtDataAte = mvardtDataAte
End Property

Public Property Let iFilialEmpresaDe(ByVal vData As Integer)
    mvariFilialEmpresaDe = vData
End Property

Public Property Get iFilialEmpresaDe() As Integer
    iFilialEmpresaDe = mvariFilialEmpresaDe
End Property

Public Property Let iFilialEmpresaAte(ByVal vData As Integer)
    mvariFilialEmpresaAte = vData
End Property

Public Property Get iFilialEmpresaAte() As Integer
    iFilialEmpresaAte = mvariFilialEmpresaAte
End Property

Public Property Let sProdutoDe(ByVal vData As String)
    mvarsProdutoDe = vData
End Property

Public Property Get sProdutoDe() As String
    sProdutoDe = mvarsProdutoDe
End Property

Public Property Let sProdutoAte(ByVal vData As String)
    mvarsProdutoAte = vData
End Property

Public Property Get sProdutoAte() As String
    sProdutoAte = mvarsProdutoAte
End Property

Public Property Let iProdutosTop(ByVal vData As Integer)
    mvariProdutosTop = vData
End Property

Public Property Get iProdutosTop() As Integer
    iProdutosTop = mvariProdutosTop
End Property

Public Property Let iTipoProduto(ByVal vData As Integer)
    mvariTipoProduto = vData
End Property

Public Property Get iTipoProduto() As Integer
    iTipoProduto = mvariTipoProduto
End Property

Public Property Let sCategoria(ByVal vData As String)
    mvarsCategoria = vData
End Property

Public Property Get sCategoria() As String
    sCategoria = mvarsCategoria
End Property

Public Property Set colItensCategoria(ByVal vData As Collection)
    Set mvarcolItensCategoria = vData
End Property

Public Property Get colItensCategoria() As Collection
    Set colItensCategoria = mvarcolItensCategoria
End Property
