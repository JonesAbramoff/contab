VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelABCFornecedoresTela"
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
Private mvarlFornecedorDe As Long
Private mvarlFornecedorAte As Long
Private mvariTipoProduto As Integer
Private mvariFornecedorTop As Integer
Private mvariDetalharFilial As Integer
Private mvarsCategoriaProdutos As String
Private mvarcolItensCategoriaProdutos As New Collection
Private mvarsCategoriaFornecedores As String
Private mvarcolItensCategoriaFornecedores As New Collection

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

Public Property Let lFornecedorDe(ByVal vData As Long)
    mvarlFornecedorDe = vData
End Property

Public Property Get lFornecedorDe() As Long
    lFornecedorDe = mvarlFornecedorDe
End Property

Public Property Let lFornecedorAte(ByVal vData As Long)
    mvarlFornecedorAte = vData
End Property

Public Property Get lFornecedorAte() As Long
    lFornecedorAte = mvarlFornecedorAte
End Property

Public Property Let iFornecedorTop(ByVal vData As Integer)
    mvariFornecedorTop = vData
End Property

Public Property Get iFornecedorTop() As Integer
    iFornecedorTop = mvariFornecedorTop
End Property

Public Property Let iDetalharFilial(ByVal vData As Integer)
    mvariDetalharFilial = vData
End Property

Public Property Get iDetalharFilial() As Integer
    iDetalharFilial = mvariDetalharFilial
End Property

Public Property Let iTipoProduto(ByVal vData As Integer)
    mvariTipoProduto = vData
End Property

Public Property Get iTipoProduto() As Integer
    iTipoProduto = mvariTipoProduto
End Property

Public Property Let sCategoriaProdutos(ByVal vData As String)
    mvarsCategoriaProdutos = vData
End Property

Public Property Get sCategoriaProdutos() As String
    sCategoriaProdutos = mvarsCategoriaProdutos
End Property

Public Property Set colItensCategoriaProdutos(ByVal vData As Collection)
    Set mvarcolItensCategoriaProdutos = vData
End Property

Public Property Get colItensCategoriaProdutos() As Collection
    Set colItensCategoriaProdutos = mvarcolItensCategoriaProdutos
End Property

Public Property Let sCategoriaFornecedores(ByVal vData As String)
    mvarsCategoriaFornecedores = vData
End Property

Public Property Get sCategoriaFornecedores() As String
    sCategoriaFornecedores = mvarsCategoriaFornecedores
End Property

Public Property Set colItensCategoriaFornecedores(ByVal vData As Collection)
    Set mvarcolItensCategoriaFornecedores = vData
End Property

Public Property Get colItensCategoriaFornecedores() As Collection
    Set colItensCategoriaFornecedores = mvarcolItensCategoriaFornecedores
End Property

