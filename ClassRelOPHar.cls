VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelOPHar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCategoria As String
Private mvarsDescCategoria As String
Private mvarsDescProduto As String
Private mvarsSigla As String
Private mvariSeq As Integer
Private mvarsProduto As String
Private mvarsProdutoPai As String
Private mvarsUM As String
Private mvardQuantidade As Double
Private mvarsCor As String
Private mvarsTonalidade As String
Private mvarsTipoDePintura As String
Private mvarcolItemOP As New Collection
Private mvarobjProduto As New ClassProduto
Private mvarsDescColecao As String

Public Property Let sDescColecao(ByVal vData As String)
    mvarsDescColecao = vData
End Property

Public Property Get sDescColecao() As String
    sDescColecao = mvarsDescColecao
End Property

Public Property Set objProduto(ByVal vData As ClassProduto)
    Set mvarobjProduto = vData
End Property

Public Property Get objProduto() As ClassProduto
    Set objProduto = mvarobjProduto
End Property

Public Property Set colItemOP(ByVal vData As Collection)
    Set mvarcolItemOP = vData
End Property

Public Property Get colItemOP() As Collection
    Set colItemOP = mvarcolItemOP
End Property

Public Property Let sTipoDePintura(ByVal vData As String)
    mvarsTipoDePintura = vData
End Property

Public Property Get sTipoDePintura() As String
    sTipoDePintura = mvarsTipoDePintura
End Property

Public Property Let sTonalidade(ByVal vData As String)
    mvarsTonalidade = vData
End Property

Public Property Get sTonalidade() As String
    sTonalidade = mvarsTonalidade
End Property

Public Property Let sCor(ByVal vData As String)
    mvarsCor = vData
End Property

Public Property Get sCor() As String
    sCor = mvarsCor
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let sProdutoPai(ByVal vData As String)
    mvarsProdutoPai = vData
End Property

Public Property Get sProdutoPai() As String
    sProdutoPai = mvarsProdutoPai
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let sSigla(ByVal vData As String)
    mvarsSigla = vData
End Property

Public Property Get sSigla() As String
    sSigla = mvarsSigla
End Property

Public Property Let sDescCategoria(ByVal vData As String)
    mvarsDescCategoria = vData
End Property

Public Property Get sDescCategoria() As String
    sDescCategoria = mvarsDescCategoria
End Property

Public Property Let sDescProduto(ByVal vData As String)
    mvarsDescProduto = vData
End Property

Public Property Get sDescProduto() As String
    sDescProduto = mvarsDescProduto
End Property

Public Property Let sCategoria(ByVal vData As String)
    mvarsCategoria = vData
End Property

Public Property Get sCategoria() As String
    sCategoria = mvarsCategoria
End Property



