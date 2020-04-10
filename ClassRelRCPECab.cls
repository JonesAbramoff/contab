VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelRCPECab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarobjProduto As Object
Private mvarlFolha As Long
Private mvarobjEstoqueProduto As Object
Private mvarobjProdutoFilial As Object
Private mvarColDetalhe As New Collection
Private mvardtDataDe As Date
Private mvardtDataAte As Date

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

Public Property Set ColDetalhe(ByVal vData As Collection)
    Set mvarColDetalhe = vData
End Property

Public Property Get ColDetalhe() As Collection
    Set ColDetalhe = mvarColDetalhe
End Property

Public Property Set objProdutoFilial(ByVal vData As Object)
    Set mvarobjProdutoFilial = vData
End Property

Public Property Get objProdutoFilial() As Object
    Set objProdutoFilial = mvarobjProdutoFilial
End Property

Public Property Set objEstoqueProduto(ByVal vData As Object)
    Set mvarobjEstoqueProduto = vData
End Property

Public Property Get objEstoqueProduto() As Object
    Set objEstoqueProduto = mvarobjEstoqueProduto
End Property

Public Property Let lFolha(ByVal vData As Long)
    mvarlFolha = vData
End Property

Public Property Get lFolha() As Long
    lFolha = mvarlFolha
End Property

Public Property Set objProduto(ByVal vData As Object)
    Set mvarobjProduto = vData
End Property

Public Property Get objProduto() As Object
    Set objProduto = mvarobjProduto
End Property



