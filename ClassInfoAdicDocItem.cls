VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInfoAdicDocItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariTipoDoc As Integer
Private mvarlNumIntDocItem As Long
Private mvardtDataLimiteFaturamento As Date
Private mvarsNumPedidoCompra As String
Private mvarlItemPedCompra As Long
Private mvariIncluiValorTotal As Integer
Private mvarsMsg As String
Private mvariItem As Integer
Private mvarsProduto As String
Private mvarsDescProd As String
Private mvarcolDetExportacao As New Collection
Private mvardValorTotal As Double
Private mvardPrecoUnitario As Double
Private mvardQuantidade As Double
Private mvarsUM As String
Private mvarlNumIntDE As Long
Private mvarsNumRE As String

Public Property Let sNumRE(ByVal vData As String)
    mvarsNumRE = vData
End Property

Public Property Get sNumRE() As String
    sNumRE = mvarsNumRE
End Property

Public Property Let lNumIntDE(ByVal vData As Long)
    mvarlNumIntDE = vData
End Property

Public Property Get lNumIntDE() As Long
    lNumIntDE = mvarlNumIntDE
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

Public Property Let dPrecoUnitario(ByVal vData As Double)
    mvardPrecoUnitario = vData
End Property

Public Property Get dPrecoUnitario() As Double
    dPrecoUnitario = mvardPrecoUnitario
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Set colDetExportacao(ByVal vData As Object)
    Set mvarcolDetExportacao = vData
End Property

Public Property Get colDetExportacao() As Collection
    Set colDetExportacao = mvarcolDetExportacao
End Property

Public Property Let sDescProd(ByVal vData As String)
    mvarsDescProd = vData
End Property

Public Property Get sDescProd() As String
    sDescProd = mvarsDescProd
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Let lNumIntDocItem(ByVal vData As Long)
    mvarlNumIntDocItem = vData
End Property

Public Property Get lNumIntDocItem() As Long
    lNumIntDocItem = mvarlNumIntDocItem
End Property

Public Property Let dtDataLimiteFaturamento(ByVal vData As Date)
    mvardtDataLimiteFaturamento = vData
End Property

Public Property Get dtDataLimiteFaturamento() As Date
    dtDataLimiteFaturamento = mvardtDataLimiteFaturamento
End Property

Public Property Let sNumPedidoCompra(ByVal vData As String)
    mvarsNumPedidoCompra = vData
End Property

Public Property Get sNumPedidoCompra() As String
    sNumPedidoCompra = mvarsNumPedidoCompra
End Property

Public Property Let iIncluiValorTotal(ByVal vData As Integer)
    mvariIncluiValorTotal = vData
End Property

Public Property Get iIncluiValorTotal() As Integer
    iIncluiValorTotal = mvariIncluiValorTotal
End Property

Public Property Let lItemPedCompra(ByVal vData As Long)
    mvarlItemPedCompra = vData
End Property

Public Property Get lItemPedCompra() As Long
    lItemPedCompra = mvarlItemPedCompra
End Property

Public Property Let sMsg(ByVal vData As String)
    mvarsMsg = vData
End Property

Public Property Get sMsg() As String
    sMsg = mvarsMsg
End Property

Private Sub Class_Initialize()
    mvardtDataLimiteFaturamento = DATA_NULA
    mvariIncluiValorTotal = MARCADO
End Sub
