VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPackNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntNF As Long
Private mvariSeq As Integer
Private mvarsDescricaoItem As String
Private mvarsProduto As String
Private mvarsUnidadeMed As String
Private mvardQuantidade As Double
Private mvardPrecoUnitario As Double

Private mvardPercDesc As Double
Private mvardValorDesconto As Double

Private mvarlNumIntPedVenda As Long
Private mvarlNumIntItemPedVenda As Long
Private mvariAlmoxarifado As Integer
Private mvarsAlmoxarifadoNomeRed As String
Private mvariControleEstoque As Integer
Private mvarsUMEstoque As String
Private mvariClasseUM As Integer
Private mvarsUMVenda As String
Private mvardQuantUMVenda As Double
Private mvariApropriacaoProd As Integer 'apropriacao de custo do produto
Private mvardCusto As Double
Private mvardValorTotal As Double

Private mvarcolAlocacoes As New ColAlocacoesItemNF
Private mvarcolRastreamento As New Collection

Public Property Set colRastreamento(ByVal vData As Collection)
    Set mvarcolRastreamento = vData
End Property

Public Property Get colRastreamento() As Collection
    Set colRastreamento = mvarcolRastreamento
End Property

Public Property Let dCusto(ByVal vData As Double)
    mvardCusto = vData
End Property

Public Property Get dCusto() As Double
    dCusto = mvardCusto
End Property

Public Property Get iApropriacaoProd() As Integer
    iApropriacaoProd = mvariApropriacaoProd
End Property

Public Property Let iApropriacaoProd(ByVal vData As Integer)
    mvariApropriacaoProd = vData
End Property

Public Property Let dQuantUMVenda(ByVal vData As Double)
    mvardQuantUMVenda = vData
End Property

Public Property Get dQuantUMVenda() As Double
    dQuantUMVenda = mvardQuantUMVenda
End Property

Public Property Let sUMVenda(ByVal vData As String)
    mvarsUMVenda = vData
End Property

Public Property Get sUMVenda() As String
    sUMVenda = mvarsUMVenda
End Property

Public Property Let iClasseUM(ByVal vData As Integer)
    mvariClasseUM = vData
End Property

Public Property Get iClasseUM() As Integer
    iClasseUM = mvariClasseUM
End Property

Public Property Let sUMEstoque(ByVal vData As String)
    mvarsUMEstoque = vData
End Property

Public Property Get sUMEstoque() As String
    sUMEstoque = mvarsUMEstoque
End Property

Public Property Let iControleEstoque(ByVal vData As Integer)
    mvariControleEstoque = vData
End Property

Public Property Get iControleEstoque() As Integer
    iControleEstoque = mvariControleEstoque
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
    '???
End Property

Public Property Let sAlmoxarifadoNomeRed(ByVal vData As String)
    mvarsAlmoxarifadoNomeRed = vData
End Property

Public Property Get sAlmoxarifadoNomeRed() As String
    sAlmoxarifadoNomeRed = mvarsAlmoxarifadoNomeRed
End Property

Public Property Let iAlmoxarifado(ByVal vData As Integer)
    mvariAlmoxarifado = vData
End Property

Public Property Get iAlmoxarifado() As Integer
    iAlmoxarifado = mvariAlmoxarifado
End Property

Public Property Set ColAlocacoes(ByVal vData As Object)
    Set mvarcolAlocacoes = vData
End Property

Public Property Get ColAlocacoes() As ColAlocacoesItemNF
    Set ColAlocacoes = mvarcolAlocacoes
End Property

Public Property Let lNumIntItemPedVenda(ByVal vData As Long)
    mvarlNumIntItemPedVenda = vData
End Property

Public Property Get lNumIntItemPedVenda() As Long
    lNumIntItemPedVenda = mvarlNumIntItemPedVenda
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntPedVenda(ByVal vData As Long)
    mvarlNumIntPedVenda = vData
End Property

Public Property Get lNumIntPedVenda() As Long
    lNumIntPedVenda = mvarlNumIntPedVenda
End Property

Public Property Let sDescricaoItem(ByVal vData As String)
    mvarsDescricaoItem = vData
End Property

Public Property Get sDescricaoItem() As String
    sDescricaoItem = mvarsDescricaoItem
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dPercDesc(ByVal vData As Double)
    mvardPercDesc = vData
End Property

Public Property Get dPercDesc() As Double
    dPercDesc = mvardPercDesc
End Property

Public Property Let dPrecoUnitario(ByVal vData As Double)
    mvardPrecoUnitario = vData
End Property

Public Property Get dPrecoUnitario() As Double
    dPrecoUnitario = mvardPrecoUnitario
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sUnidadeMed(ByVal vData As String)
    mvarsUnidadeMed = vData
End Property

Public Property Get sUnidadeMed() As String
    sUnidadeMed = mvarsUnidadeMed
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let lNumIntNF(ByVal vData As Long)
    mvarlNumIntNF = vData
End Property

Public Property Get lNumIntNF() As Long
    lNumIntNF = mvarlNumIntNF
End Property




