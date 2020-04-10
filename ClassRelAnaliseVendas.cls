VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelAnaliseVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntRel As Long
Private mvariFilialEmpresa As Integer
Private mvardtDataEmissaoDe As Date
Private mvardtDataEmissaoAte As Date
Private mvariVendedorDe As Integer
Private mvariVendedorAte As Integer
Private mvarlClienteDe As Long
Private mvarlClienteAte As Long
Private mvarsProdutoDe As String
Private mvarsProdutoAte As String
Private mvardLucroDe As Double
Private mvardLucroAte As Double
Private mvarsLucroDe As String
Private mvarsLucroAte As String
Private mvariTipoFiltroProduto As Integer

Public Property Let iTipoFiltroProduto(ByVal vData As Integer)
    mvariTipoFiltroProduto = vData
End Property

Public Property Get iTipoFiltroProduto() As Integer
    iTipoFiltroProduto = mvariTipoFiltroProduto
End Property

Public Property Let sLucroDe(ByVal vData As String)
    mvarsLucroDe = vData
End Property

Public Property Get sLucroDe() As String
    sLucroDe = mvarsLucroDe
End Property

Public Property Let sLucroAte(ByVal vData As String)
    mvarsLucroAte = vData
End Property

Public Property Get sLucroAte() As String
    sLucroAte = mvarsLucroAte
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

Public Property Let dLucroDe(ByVal vData As Double)
    mvardLucroDe = vData
End Property

Public Property Get dLucroDe() As Double
    dLucroDe = mvardLucroDe
End Property

Public Property Let dLucroAte(ByVal vData As Double)
    mvardLucroAte = vData
End Property

Public Property Get dLucroAte() As Double
    dLucroAte = mvardLucroAte
End Property

Public Property Let dtDataEmissaoDe(ByVal vData As Date)
    mvardtDataEmissaoDe = vData
End Property

Public Property Get dtDataEmissaoDe() As Date
    dtDataEmissaoDe = mvardtDataEmissaoDe
End Property

Public Property Let dtDataEmissaoAte(ByVal vData As Date)
    mvardtDataEmissaoAte = vData
End Property

Public Property Get dtDataEmissaoAte() As Date
    dtDataEmissaoAte = mvardtDataEmissaoAte
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
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

Public Property Let iVendedorDe(ByVal vData As Integer)
    mvariVendedorDe = vData
End Property

Public Property Get iVendedorDe() As Integer
    iVendedorDe = mvariVendedorDe
End Property

Public Property Let iVendedorAte(ByVal vData As Integer)
    mvariVendedorAte = vData
End Property

Public Property Get iVendedorAte() As Integer
    iVendedorAte = mvariVendedorAte
End Property


