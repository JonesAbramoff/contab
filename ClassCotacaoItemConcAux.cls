VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCotacaoItemConcAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarobjCotacaoItemConc As New ClassCotacaoItemConc
Private mvarsProduto As String
Private mvarsDescricao As String
Private mvarsCondPagto As String
Private mvarsFornecedor As String
Private mvarsFilial As String
Private mvardtDataNecessidade As Date

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property
Public Property Let dtDataNecessidade(ByVal vData As Date)
    mvardtDataNecessidade = vData
End Property

Public Property Get dtDataNecessidade() As Date
    dtDataNecessidade = mvardtDataNecessidade
End Property

Public Property Let sFilial(ByVal vData As String)
    mvarsFilial = vData
End Property

Public Property Get sFilial() As String
    sFilial = mvarsFilial
End Property

Public Property Let sFornecedor(ByVal vData As String)
    mvarsFornecedor = vData
End Property

Public Property Get sFornecedor() As String
    sFornecedor = mvarsFornecedor
End Property

Public Property Let sCondPagto(ByVal vData As String)
    mvarsCondPagto = vData
End Property

Public Property Get sCondPagto() As String
    sCondPagto = mvarsCondPagto
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Set objCotacaoItemConc(ByVal vData As ClassCotacaoItemConc)
    Set mvarobjCotacaoItemConc = vData
End Property

Public Property Get objCotacaoItemConc() As ClassCotacaoItemConc
    Set objCotacaoItemConc = mvarobjCotacaoItemConc
End Property
