VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassComissoesRegras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariVendedor As Integer
Private mvarsRegra As String
Private mvarlOrdenacao As Long
Private mvarsValorBase As String
Private mvarsPercComissao As String
Private mvariVendedorIndireto As Integer
Private mvarlNumIntDocOrigem As Long
Private mvardPercComissaoEmiss As Double

Public Property Let dPercComissaoEmiss(ByVal vData As Double)
    mvardPercComissaoEmiss = vData
End Property

Public Property Get dPercComissaoEmiss() As Double
    dPercComissaoEmiss = mvardPercComissaoEmiss
End Property

Public Property Let lNumIntDocOrigem(ByVal vData As Long)
    mvarlNumIntDocOrigem = vData
End Property

Public Property Get lNumIntDocOrigem() As Long
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property

Public Property Let iVendedorIndireto(ByVal vData As Integer)
    mvariVendedorIndireto = vData
End Property

Public Property Get iVendedorIndireto() As Integer
    iVendedorIndireto = mvariVendedorIndireto
End Property

Public Property Let sPercComissao(ByVal vData As String)
    mvarsPercComissao = vData
End Property

Public Property Get sPercComissao() As String
    sPercComissao = mvarsPercComissao
End Property

Public Property Let sValorBase(ByVal vData As String)
    mvarsValorBase = vData
End Property

Public Property Get sValorBase() As String
    sValorBase = mvarsValorBase
End Property

Public Property Let lOrdenacao(ByVal vData As Long)
    mvarlOrdenacao = vData
End Property

Public Property Get lOrdenacao() As Long
    lOrdenacao = mvarlOrdenacao
End Property

Public Property Let sRegra(ByVal vData As String)
    mvarsRegra = vData
End Property

Public Property Get sRegra() As String
    sRegra = mvarsRegra
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property
