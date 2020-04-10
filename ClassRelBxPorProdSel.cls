VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelBxPorProdSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntRel As Long
Private mvarlClienteDe As Long
Private mvarlClienteAte As Long
Private mvarcolRegioes As New Collection
Private mvariTipoVend As Integer
Private mvariVendedor As Integer
Private mvariDevolucoes As Integer
Private mvariExibirVend As Integer
Private mvariExibirDet As Integer
Private mvariTipoDe As Integer
Private mvariTipoAte As Integer
Private mvarsProdutoDe As String
Private mvarsProdutoAte As String
Private mvarsCidade As String
Private mvariFilialEmpresa As Integer
Private mvardtDataEmiDe As Date
Private mvardtDataEmiAte As Date
Private mvardtDataBxDe As Date
Private mvardtDataBxAte As Date

Public Property Let dtDataBxAte(ByVal vData As Date)
    mvardtDataBxAte = vData
End Property

Public Property Get dtDataBxAte() As Date
    dtDataBxAte = mvardtDataBxAte
End Property

Public Property Let dtDataBxDe(ByVal vData As Date)
    mvardtDataBxDe = vData
End Property

Public Property Get dtDataBxDe() As Date
    dtDataBxDe = mvardtDataBxDe
End Property

Public Property Let dtDataEmiAte(ByVal vData As Date)
    mvardtDataEmiAte = vData
End Property

Public Property Get dtDataEmiAte() As Date
    dtDataEmiAte = mvardtDataEmiAte
End Property

Public Property Let dtDataEmiDe(ByVal vData As Date)
    mvardtDataEmiDe = vData
End Property

Public Property Get dtDataEmiDe() As Date
    dtDataEmiDe = mvardtDataEmiDe
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sCidade(ByVal vData As String)
    mvarsCidade = vData
End Property

Public Property Get sCidade() As String
    sCidade = mvarsCidade
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

Public Property Let iTipoAte(ByVal vData As Integer)
    mvariTipoAte = vData
End Property

Public Property Get iTipoAte() As Integer
    iTipoAte = mvariTipoAte
End Property

Public Property Let iTipoDe(ByVal vData As Integer)
    mvariTipoDe = vData
End Property

Public Property Get iTipoDe() As Integer
    iTipoDe = mvariTipoDe
End Property

Public Property Let iExibirDet(ByVal vData As Integer)
    mvariExibirDet = vData
End Property

Public Property Get iExibirDet() As Integer
    iExibirDet = mvariExibirDet
End Property

Public Property Let iExibirVend(ByVal vData As Integer)
    mvariExibirVend = vData
End Property

Public Property Get iExibirVend() As Integer
    iExibirVend = mvariExibirVend
End Property

Public Property Let iDevolucoes(ByVal vData As Integer)
    mvariDevolucoes = vData
End Property

Public Property Get iDevolucoes() As Integer
    iDevolucoes = mvariDevolucoes
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let iTipoVend(ByVal vData As Integer)
    mvariTipoVend = vData
End Property

Public Property Get iTipoVend() As Integer
    iTipoVend = mvariTipoVend
End Property

Public Property Set colRegioes(ByVal vData As Collection)
    Set mvarcolRegioes = vData
End Property

Public Property Get colRegioes() As Collection
    Set colRegioes = mvarcolRegioes
End Property

Public Property Let lClienteDe(ByVal vData As Long)
    mvarlClienteDe = vData
End Property

Public Property Get lClienteDe() As Long
    lClienteDe = mvarlClienteDe
End Property

Public Property Let lClienteAte(ByVal vData As Long)
    mvarlClienteAte = vData
End Property

Public Property Get lClienteAte() As Long
    lClienteAte = mvarlClienteAte
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property
