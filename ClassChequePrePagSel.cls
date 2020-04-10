VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChequePrePagSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumeroDe As Long
Private mvardtDataEmissaoDe As Date
Private mvardtDataBomParaDe As Date
Private mvarlNumeroAte As Long
Private mvardtDataEmissaoAte As Date
Private mvardtDataBomParaAte As Date
Private mvarcolCheques As New Collection
Private mvarlFornecedor As Long
Private mvariFilial As Integer

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Get colCheques() As Collection
    Set colCheques = mvarcolCheques
End Property

Public Property Set colCheques(ByVal vData As Collection)
    Set mvarcolCheques = vData
End Property

Public Property Let lNumeroDe(ByVal vData As Long)
    mvarlNumeroDe = vData
End Property

Public Property Get lNumeroDe() As Long
    lNumeroDe = mvarlNumeroDe
End Property

Public Property Let dtDataEmissaoDe(ByVal vData As Date)
    mvardtDataEmissaoDe = vData
End Property

Public Property Get dtDataEmissaoDe() As Date
    dtDataEmissaoDe = mvardtDataEmissaoDe
End Property

Public Property Let dtDataBomParaDe(ByVal vData As Date)
    mvardtDataBomParaDe = vData
End Property

Public Property Get dtDataBomParaDe() As Date
    dtDataBomParaDe = mvardtDataBomParaDe
End Property

Public Property Let lNumeroAte(ByVal vData As Long)
    mvarlNumeroAte = vData
End Property

Public Property Get lNumeroAte() As Long
    lNumeroAte = mvarlNumeroAte
End Property

Public Property Let dtDataEmissaoAte(ByVal vData As Date)
    mvardtDataEmissaoAte = vData
End Property

Public Property Get dtDataEmissaoAte() As Date
    dtDataEmissaoAte = mvardtDataEmissaoAte
End Property

Public Property Let dtDataBomParaAte(ByVal vData As Date)
    mvardtDataBomParaAte = vData
End Property

Public Property Get dtDataBomParaAte() As Date
    dtDataBomParaAte = mvardtDataBomParaAte
End Property


