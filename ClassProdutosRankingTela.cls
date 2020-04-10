VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdutosRankingTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardtDataDe As Date
Private mvardtDataAte As Date
Private mvariClassificacao As Integer
Private mvariFilialEmpresa As Integer
Private mvarsProdutoDe As String
Private mvarsProdutoAte As String
Private mvariTipoProdutoDe As Integer
Private mvariTipoProdutoAte As Integer
Private mvarsCategoria As String
Private mvarsTituloGrafico As String
Private mvarsNomeGrafico As String
Private mvarsNomePlanilha As String

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

Public Property Let iClassificacao(ByVal vData As Integer)
    mvariClassificacao = vData
End Property

Public Property Get iClassificacao() As Integer
    iClassificacao = mvariClassificacao
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iTipoProdutoAte(ByVal vData As Integer)
    mvariTipoProdutoAte = vData
End Property

Public Property Get iTipoProdutoAte() As Integer
    iTipoProdutoAte = mvariTipoProdutoAte
End Property

Public Property Let iTipoProdutoDe(ByVal vData As Integer)
    mvariTipoProdutoDe = vData
End Property

Public Property Get iTipoProdutoDe() As Integer
    iTipoProdutoDe = mvariTipoProdutoDe
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

Public Property Let sCategoria(ByVal vData As String)
    mvarsCategoria = vData
End Property

Public Property Get sCategoria() As String
    sCategoria = mvarsCategoria
End Property

Public Property Let sTituloGrafico(ByVal vData As String)
    mvarsTituloGrafico = vData
End Property

Public Property Get sTituloGrafico() As String
    sTituloGrafico = mvarsTituloGrafico
End Property

Public Property Let sNomeGrafico(ByVal vData As String)
    mvarsNomeGrafico = vData
End Property

Public Property Get sNomeGrafico() As String
    sNomeGrafico = mvarsNomeGrafico
End Property

Public Property Let sNomePlanilha(ByVal vData As String)
    mvarsNomePlanilha = vData
End Property

Public Property Get sNomePlanilha() As String
    sNomePlanilha = mvarsNomePlanilha
End Property

