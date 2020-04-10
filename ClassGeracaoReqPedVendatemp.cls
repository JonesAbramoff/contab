VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeracaoReqPedVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarPedidoDe As Long
Private mvarPedidoAte As Long
Private mvarClienteDe As Long
Private mvarClienteAte As Long
Private mvarDataEmissaoDe As Date
Private mvarDataEmissaoAte As Date
Private mvarDataEntregaDe As Date
Private mvarDataEntregaAte As Date
Private mvarcolPedVenda As New Collection
Private mvarsOrdenacao As String
Private mvariExibeTodos As Integer
Public Property Let iExibeTodos(ByVal vData As Integer)
    mvariExibeTodos = vData
End Property


Public Property Get iExibeTodos() As Integer
    iExibeTodos = mvariExibeTodos
End Property



Public Property Let sOrdenacao(ByVal vData As String)
    mvarsOrdenacao = vData
End Property


Public Property Get sOrdenacao() As String
    sOrdenacao = mvarsOrdenacao
End Property



Public Property Set colPedVenda(ByVal vData As Collection)
    Set mvarcolPedVenda = vData
End Property


Public Property Get colPedVenda() As Collection
    Set colPedVenda = mvarcolPedVenda
End Property



Public Property Let DataEntregaAte(ByVal vData As Date)
    mvarDataEntregaAte = vData
End Property


Public Property Get DataEntregaAte() As Date
    DataEntregaAte = mvarDataEntregaAte
End Property



Public Property Let DataEntregaDe(ByVal vData As Date)
    mvarDataEntregaDe = vData
End Property


Public Property Get DataEntregaDe() As Date
    DataEntregaDe = mvarDataEntregaDe
End Property



Public Property Let DataEmissaoAte(ByVal vData As Date)
    mvarDataEmissaoAte = vData
End Property


Public Property Get DataEmissaoAte() As Date
    DataEmissaoAte = mvarDataEmissaoAte
End Property



Public Property Let DataEmissaoDe(ByVal vData As Date)
    mvarDataEmissaoDe = vData
End Property


Public Property Get DataEmissaoDe() As Date
    DataEmissaoDe = mvarDataEmissaoDe
End Property



Public Property Let ClienteAte(ByVal vData As Long)
    mvarClienteAte = vData
End Property


Public Property Get ClienteAte() As Long
    ClienteAte = mvarClienteAte
End Property



Public Property Let ClienteDe(ByVal vData As Long)
    mvarClienteDe = vData
End Property


Public Property Get ClienteDe() As Long
    ClienteDe = mvarClienteDe
End Property



Public Property Let PedidoAte(ByVal vData As Long)
    mvarPedidoAte = vData
End Property


Public Property Get PedidoAte() As Long
    PedidoAte = mvarPedidoAte
End Property



Public Property Let PedidoDe(ByVal vData As Long)
    mvarPedidoDe = vData
End Property


Public Property Get PedidoDe() As Long
    PedidoDe = mvarPedidoDe
End Property



