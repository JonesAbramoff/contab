VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemPedCotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarsProduto As String
Private mvardQuantidade As Double
Private mvarsUM As String
Private mvarcolItensCotacao As New Collection
Private mvarlCotacaoProduto As Long
Private mvariExclusivo As Integer
Private mvarlLinkProv As Long
Public sObservacao As String

Public Property Let lLinkProv(ByVal vData As Long)
    mvarlLinkProv = vData
End Property

Public Property Get lLinkProv() As Long
    lLinkProv = mvarlLinkProv
End Property

Public Property Let iExclusivo(ByVal vData As Integer)
    mvariExclusivo = vData
End Property

Public Property Get iExclusivo() As Integer
    iExclusivo = mvariExclusivo
End Property

Public Property Let lCotacaoProduto(ByVal vData As Long)
    mvarlCotacaoProduto = vData
End Property


Public Property Get lCotacaoProduto() As Long
    lCotacaoProduto = mvarlCotacaoProduto
End Property





Public Property Set colItensCotacao(ByVal vData As Object)
    Set mvarcolItensCotacao = vData
End Property


Public Property Get colItensCotacao() As Collection
    Set colItensCotacao = mvarcolItensCotacao
End Property


Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property


Public Property Get sUM() As String
    sUM = mvarsUM
End Property



Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property


Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property



Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property


Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property



Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property


Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property



