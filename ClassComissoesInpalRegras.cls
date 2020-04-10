VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassComisInpalRegras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlCodPlanilha As Long
Private mvariRegiaoVenda As Integer
Private mvarlCliente As Long
Private mvariFilialCliente As Integer
Private mvarsCategoriaProduto As String
Private mvarsItemCatProduto As String
Private mvardPercTabelaA As Double
Private mvardPercTabelaB As Double

Public Property Let dPercTabelaB(ByVal vData As Double)
    mvardPercTabelaB = vData
End Property

Public Property Get dPercTabelaB() As Double
    dPercTabelaB = mvardPercTabelaB
End Property

Public Property Let dPercTabelaA(ByVal vData As Double)
    mvardPercTabelaA = vData
End Property

Public Property Get dPercTabelaA() As Double
    dPercTabelaA = mvardPercTabelaA
End Property

Public Property Let sItemCatProduto(ByVal vData As String)
    mvarsItemCatProduto = vData
End Property

Public Property Get sItemCatProduto() As String
    sItemCatProduto = mvarsItemCatProduto
End Property

Public Property Let sCategoriaProduto(ByVal vData As String)
    mvarsCategoriaProduto = vData
End Property

Public Property Get sCategoriaProduto() As String
    sCategoriaProduto = mvarsCategoriaProduto
End Property

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iRegiaoVenda(ByVal vData As Integer)
    mvariRegiaoVenda = vData
End Property

Public Property Get iRegiaoVenda() As Integer
    iRegiaoVenda = mvariRegiaoVenda
End Property

Public Property Let lCodPlanilha(ByVal vData As Long)
    mvarlCodPlanilha = vData
End Property

Public Property Get lCodPlanilha() As Long
    lCodPlanilha = mvarlCodPlanilha
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property
