VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemPedCompraInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ClassItemPedCotacao"
Option Explicit

Public objInfoUsu As Object 'Inserido por Wagner 14/07/2006

Private mvarlNumIntDoc As Long
Private mvarsProduto As String
Private mvarsDescProduto As String
Private mvardQuantRecebida As Double
Private mvarsUM As String
Private mvardQuantReceber As Double
Private mvarlPedCompra As Long
Private mvariItem As Integer
Private mvardPercentMaisReceb As Double
Private mvariRecebForaFaixa As Integer
Private mvardPrecoUnitario As Double
Private mvardAliquotaIPI As Double
Private mvardAliquotaICMS As Double
Private mvardValorDesconto As Double
Private mvarcolLocalizacao As New Collection
Private mvariMoeda As Integer
Private mvardTaxa As Double
Private mvarbTaxaPedido As Boolean
Private mvardQuantPedida As Double

Public Property Let dQuantPedida(ByVal vData As Double)
    mvardQuantPedida = vData
End Property

Public Property Get dQuantPedida() As Double
    dQuantPedida = mvardQuantPedida
End Property

Public Property Set colLocalizacao(ByVal vData As Object)
    Set mvarcolLocalizacao = vData
End Property

Public Property Get colLocalizacao() As Collection
    Set colLocalizacao = mvarcolLocalizacao
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let iRecebForaFaixa(ByVal vData As Integer)
    mvariRecebForaFaixa = vData
End Property

Public Property Get iRecebForaFaixa() As Integer
    iRecebForaFaixa = mvariRecebForaFaixa
End Property

Public Property Let dPercentMaisReceb(ByVal vData As Double)
    mvardPercentMaisReceb = vData
End Property

Public Property Get dPercentMaisReceb() As Double
    dPercentMaisReceb = mvardPercentMaisReceb
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let lPedCompra(ByVal vData As Long)
Attribute lPedCompra.VB_Description = "Código (externo) do Pedido de Compras."
    mvarlPedCompra = vData
End Property

Public Property Get lPedCompra() As Long
    lPedCompra = mvarlPedCompra
End Property

Public Property Let dQuantReceber(ByVal vData As Double)
    mvardQuantReceber = vData
End Property

Public Property Get dQuantReceber() As Double
    dQuantReceber = mvardQuantReceber
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let dQuantRecebida(ByVal vData As Double)
    mvardQuantRecebida = vData
End Property

Public Property Get dQuantRecebida() As Double
    dQuantRecebida = mvardQuantRecebida
End Property

Public Property Let sDescProduto(ByVal vData As String)
    mvarsDescProduto = vData
End Property

Public Property Get sDescProduto() As String
    sDescProduto = mvarsDescProduto
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

Public Property Let dAliquotaICMS(ByVal vData As Double)
    mvardAliquotaICMS = vData
End Property

Public Property Get dAliquotaICMS() As Double
    dAliquotaICMS = mvardAliquotaICMS
End Property

Public Property Let dAliquotaIPI(ByVal vData As Double)
    mvardAliquotaIPI = vData
End Property

Public Property Get dAliquotaIPI() As Double
    dAliquotaIPI = mvardAliquotaIPI
End Property

Public Property Let dPrecoUnitario(ByVal vData As Double)
    mvardPrecoUnitario = vData
End Property

Public Property Get dPrecoUnitario() As Double
    dPrecoUnitario = mvardPrecoUnitario
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let dTaxa(ByVal vData As Double)
    mvardTaxa = vData
End Property

Public Property Get dTaxa() As Double
    dTaxa = mvardTaxa
End Property

Public Property Let bTaxaPedido(ByVal vData As Boolean)
    mvarbTaxaPedido = vData
End Property

Public Property Get bTaxaPedido() As Boolean
    bTaxaPedido = mvarbTaxaPedido
End Property
