VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBaixaPedCotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private mvardtDataEmissaoAte As Date
Private mvardtDataEmissaoDe As Date
Private mvarlFornecedorAte As Long
Private mvarlFornecedorDe As Long
Private mvarlPedCotacaoAte As Long
Private mvarlPedCotacaoDe As Long
Private mvardtDataValidadeDe As Date
Private mvardtDataValidadeAte As Date
Private mvarcolPedCotacao As Collection
Public Property Set colPedCotacao(ByVal vData As Collection)
    Set mvarcolPedCotacao = vData
End Property


Public Property Get colPedCotacao() As Collection
    Set colPedCotacao = mvarcolPedCotacao
End Property



Public Property Let dtDataValidadeAte(ByVal vData As Date)
    mvardtDataValidadeAte = vData
End Property


Public Property Get dtDataValidadeAte() As Date
    dtDataValidadeAte = mvardtDataValidadeAte
End Property



Public Property Let dtDataValidadeDe(ByVal vData As Date)
    mvardtDataValidadeDe = vData
End Property


Public Property Get dtDataValidadeDe() As Date
    dtDataValidadeDe = mvardtDataValidadeDe
End Property



Public Property Let lPedCotacaoDe(ByVal vData As Long)
    mvarlPedCotacaoDe = vData
End Property


Public Property Get lPedCotacaoDe() As Long
    lPedCotacaoDe = mvarlPedCotacaoDe
End Property



Public Property Let lPedCotacaoAte(ByVal vData As Long)
    mvarlPedCotacaoAte = vData
End Property


Public Property Get lPedCotacaoAte() As Long
    lPedCotacaoAte = mvarlPedCotacaoAte
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



