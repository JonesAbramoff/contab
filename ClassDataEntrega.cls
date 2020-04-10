VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDataEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntPVNF As Long
Private mvarlNumIntItemPVNF As Long
Private mvardtDataEntrega As Date
Private mvardQuantidade As Double
Private mvarsPedidoCliente As String
Private mvardQuantidadeEntregue As Double

Public Property Let dQuantidadeEntregue(ByVal vData As Double)
    mvardQuantidadeEntregue = vData
End Property

Public Property Get dQuantidadeEntregue() As Double
    dQuantidadeEntregue = mvardQuantidadeEntregue
End Property

Public Property Let sPedidoCliente(ByVal vData As String)
    mvarsPedidoCliente = vData
End Property

Public Property Get sPedidoCliente() As String
    sPedidoCliente = mvarsPedidoCliente
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let dtDataEntrega(ByVal vData As Date)
    mvardtDataEntrega = vData
End Property

Public Property Get dtDataEntrega() As Date
    dtDataEntrega = mvardtDataEntrega
End Property

Public Property Let lNumIntItemPVNF(ByVal vData As Long)
    mvarlNumIntItemPVNF = vData
End Property

Public Property Get lNumIntItemPVNF() As Long
    lNumIntItemPVNF = mvarlNumIntItemPVNF
End Property

Public Property Let lNumIntPVNF(ByVal vData As Long)
    mvarlNumIntPVNF = vData
End Property

Public Property Get lNumIntPVNF() As Long
    lNumIntPVNF = mvarlNumIntPVNF
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

