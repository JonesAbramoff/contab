VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTabelaPrecoGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariTabela As Integer
Private mvariFilialEmpresa As Integer
Private mvardtDataVigencia As Date
Private mvarcolItens As New Collection
Private mvariTipoNovoPreco As Integer
Private mvardPrecoNovoRS As Double
Private mvardPrecoNovoPerc As Double

'*******************
'Filtros
Private mvardtDataRef As Date
Private mvarcolCategorias As New Collection
Private mvariTipoDeProduto As Integer
Private mvarsProdutoPai As String
Private mvarsCodigoLike As String
Private mvarsDescricaoLike As String
Private mvarsNomeRedLike As String
Private mvarsReferenciaLike As String
Private mvarsModeloLike As String
Private mvariAnaliticoSemGrade As Integer
Private mvariAnaliticoComGrade As Integer
Private mvariGradeKitVenda As Integer
Private mvarsProdutoDe As String
Private mvarsProdutoAte As String

Public Property Let dtDataRef(ByVal vData As Date)
    mvardtDataRef = vData
End Property

Public Property Get dtDataRef() As Date
    dtDataRef = mvardtDataRef
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

Public Property Let iGradeKitVenda(ByVal vData As Integer)
    mvariGradeKitVenda = vData
End Property

Public Property Get iGradeKitVenda() As Integer
    iGradeKitVenda = mvariGradeKitVenda
End Property

Public Property Let iAnaliticoComGrade(ByVal vData As Integer)
    mvariAnaliticoComGrade = vData
End Property

Public Property Get iAnaliticoComGrade() As Integer
    iAnaliticoComGrade = mvariAnaliticoComGrade
End Property

Public Property Let iAnaliticoSemGrade(ByVal vData As Integer)
    mvariAnaliticoSemGrade = vData
End Property

Public Property Get iAnaliticoSemGrade() As Integer
    iAnaliticoSemGrade = mvariAnaliticoSemGrade
End Property

Public Property Let sModeloLike(ByVal vData As String)
    mvarsModeloLike = vData
End Property

Public Property Get sModeloLike() As String
    sModeloLike = mvarsModeloLike
End Property

Public Property Let sReferenciaLike(ByVal vData As String)
    mvarsReferenciaLike = vData
End Property

Public Property Get sReferenciaLike() As String
    sReferenciaLike = mvarsReferenciaLike
End Property

Public Property Let sNomeRedLike(ByVal vData As String)
    mvarsNomeRedLike = vData
End Property

Public Property Get sNomeRedLike() As String
    sNomeRedLike = mvarsNomeRedLike
End Property

Public Property Let sDescricaoLike(ByVal vData As String)
    mvarsDescricaoLike = vData
End Property

Public Property Get sDescricaoLike() As String
    sDescricaoLike = mvarsDescricaoLike
End Property

Public Property Let sCodigoLike(ByVal vData As String)
    mvarsCodigoLike = vData
End Property

Public Property Get sCodigoLike() As String
    sCodigoLike = mvarsCodigoLike
End Property

Public Property Let sProdutoPai(ByVal vData As String)
    mvarsProdutoPai = vData
End Property

Public Property Get sProdutoPai() As String
    sProdutoPai = mvarsProdutoPai
End Property

Public Property Let iTipoDeProduto(ByVal vData As Integer)
    mvariTipoDeProduto = vData
End Property

Public Property Get iTipoDeProduto() As Integer
    iTipoDeProduto = mvariTipoDeProduto
End Property

Public Property Set colCategorias(ByVal vData As Collection)
    Set mvarcolCategorias = vData
End Property

Public Property Get colCategorias() As Collection
    Set colCategorias = mvarcolCategorias
End Property

'*******************
Public Property Let dPrecoNovoPerc(ByVal vData As Double)
    mvardPrecoNovoPerc = vData
End Property

Public Property Get dPrecoNovoPerc() As Double
    dPrecoNovoPerc = mvardPrecoNovoPerc
End Property

Public Property Let dPrecoNovoRS(ByVal vData As Double)
    mvardPrecoNovoRS = vData
End Property

Public Property Get dPrecoNovoRS() As Double
    dPrecoNovoRS = mvardPrecoNovoRS
End Property

Public Property Let iTipoNovoPreco(ByVal vData As Integer)
    mvariTipoNovoPreco = vData
End Property

Public Property Get iTipoNovoPreco() As Integer
    iTipoNovoPreco = mvariTipoNovoPreco
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iTabela(ByVal vData As Integer)
    mvariTabela = vData
End Property

Public Property Get iTabela() As Integer
    iTabela = mvariTabela
End Property

Public Property Let dtDataVigencia(ByVal vData As Date)
    mvardtDataVigencia = vData
End Property

Public Property Get dtDataVigencia() As Date
    dtDataVigencia = mvardtDataVigencia
End Property



