VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Public objInfoUsu As Object 'Inserido por Wagner 20/03/2006

Private mvarobjInfoAdicDocItem As New ClassInfoAdicDocItem

'dValorTotal é property calculada (read-only)
Private mvarlNumIntNF As Long
Private mvariItem As Integer
Private mvarsProduto As String
Private mvarsUnidadeMed As String
Private mvardQuantidade As Double
Private mvardPrecoUnitario As Double
Private mvardPercDesc As Double
Private mvardValorDesconto As Double
Private mvardValorDescGlobal As Double
Private mvardtDataEntrega As Date
Private mvarsDescricaoItem As String
Private mvardValorAbatComissao As Double
Private mvarlNumIntPedVenda As Long
Private mvarlNumIntItemPedVenda As Long
Private mvarlNumIntDoc As Long
Private mvarlNumIntTrib As Long
Private mvarcolAlocacoes As New ColAlocacoesItemNF
Private mvariAlmoxarifado As Integer
Private mvarsAlmoxarifadoNomeRed As String
Private mvariStatus As Integer
Private mvarlNumIntDocOrig As Long
Private mvarsSerieNFOrig As String
Private mvarlNumNFOrig As Long
Private mvariItemNFOrig As Integer
Private mvarobjTribItemNF As New ClassTributacaoDocItem
Private mvariControleEstoque As Integer
Private mvarsUMEstoque As String
Private mvariClasseUM As Integer
Private mvarsUMVenda As String
Private mvardQuantUMVenda As Double
Private mvarsCcl As String
Private mvariApropriacaoProd As Integer 'apropriacao de custo do produto
Private mvarcolItemNFItemPC As New Collection
Private mvarcolItemNFItemRC As New Collection
Private mvardCusto As Double
Private mvarcolRastreamento As New Collection
Private mvardValorTotal As Double
'Indica se o Item de Pedido de Venda é um pai de grade
Private mvariPossuiGrade As Integer
'Guarda os itens filhos de grade
Private mvarcolItensRomaneioGrade As New Collection
Private mvardPrecoUnitarioMoeda As Double
Private mvarobjCobrItensContrato As New ClassCobrItensContrato 'local copy
Private mvarcolDataEntrega As New Collection
Private mvarobjItensDeContrato As New ClassItensDeContrato
Private mvariRecebForaFaixa As Integer
Private mvardPercentMaisReceb As Double
Private mvardPercentMenosReceb As Double
Private mvariSeqPack As Integer
Private mvarlNumIntPack As Long
Public dDespImpValorRateado As Double
Private mvariTipo As Integer

Private mvarcolDataEntregaPV As New Collection

Private mvarcolItensNFDevBenefRem As New Collection
Private mvarcolItensNFDevBenefItemOP As New Collection
Private mvardQuantidadeAloc As Double
Private mvarsProdutoXml As String
Private mvarsDescricaoProd As String
Private mvarsUnidadeMedXml As String
Private mvarsEANXml As String
Private mvarsEAN As String
Private mvarsCFOPXml As String
Private mvardValorTotalXml As Double

Public dComissao As Double
Public iTabelaPreco As Integer
Private mvarcolEmbalagens As New Collection

Public Property Set colEmbalagens(ByVal vData As Collection)
    Set mvarcolEmbalagens = vData
End Property

Public Property Get colEmbalagens() As Collection
    Set colEmbalagens = mvarcolEmbalagens
End Property

Public Property Let dValorTotalXml(ByVal vData As Double)
    mvardValorTotalXml = vData
End Property

Public Property Get dValorTotalXml() As Double
    dValorTotalXml = mvardValorTotalXml
End Property

Public Property Let sCFOPXml(ByVal vData As String)
    mvarsCFOPXml = vData
End Property

Public Property Get sCFOPXml() As String
    sCFOPXml = mvarsCFOPXml
End Property

Public Property Let sEAN(ByVal vData As String)
    mvarsEAN = vData
End Property

Public Property Get sEAN() As String
    sEAN = mvarsEAN
End Property

Public Property Let sEANXml(ByVal vData As String)
    mvarsEANXml = vData
End Property

Public Property Get sEANXml() As String
    sEANXml = mvarsEANXml
End Property

Public Property Let sUnidadeMedXml(ByVal vData As String)
    mvarsUnidadeMedXml = vData
End Property

Public Property Get sUnidadeMedXml() As String
    sUnidadeMedXml = mvarsUnidadeMedXml
End Property

Public Property Let sDescricaoProd(ByVal vData As String)
    mvarsDescricaoProd = vData
End Property

Public Property Get sDescricaoProd() As String
    sDescricaoProd = mvarsDescricaoProd
End Property

Public Property Let sProdutoXml(ByVal vData As String)
    mvarsProdutoXml = vData
End Property

Public Property Get sProdutoXml() As String
    sProdutoXml = mvarsProdutoXml
End Property

Public Property Let dQuantidadeAloc(ByVal vData As Double)
    mvardQuantidadeAloc = vData
End Property

Public Property Get dQuantidadeAloc() As Double
    dQuantidadeAloc = mvardQuantidadeAloc
End Property

Public Property Set objInfoAdicDocItem(ByVal vData As Object)
    Set mvarobjInfoAdicDocItem = vData
End Property

Public Property Get objInfoAdicDocItem() As ClassInfoAdicDocItem
    Set objInfoAdicDocItem = mvarobjInfoAdicDocItem
End Property

Public Property Set colItensNFDevBenefItemOP(ByVal vData As Object)
    Set mvarcolItensNFDevBenefItemOP = vData
End Property

Public Property Get colItensNFDevBenefItemOP() As Collection
    Set colItensNFDevBenefItemOP = mvarcolItensNFDevBenefItemOP
End Property

Public Property Set colItensNFDevBenefRem(ByVal vData As Object)
    Set mvarcolItensNFDevBenefRem = vData
End Property

Public Property Get colItensNFDevBenefRem() As Collection
    Set colItensNFDevBenefRem = mvarcolItensNFDevBenefRem
End Property

Public Property Set colDataEntregaPV(ByVal vData As Object)
    Set mvarcolDataEntregaPV = vData
End Property

Public Property Get colDataEntregaPV() As Collection
    Set colDataEntregaPV = mvarcolDataEntregaPV
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let lNumIntPack(ByVal vData As Long)
    mvarlNumIntPack = vData
End Property

Public Property Get lNumIntPack() As Long
    lNumIntPack = mvarlNumIntPack
End Property

Public Property Let iSeqPack(ByVal vData As Integer)
    mvariSeqPack = vData
End Property

Public Property Get iSeqPack() As Integer
    iSeqPack = mvariItem
End Property

Public Property Get iRecebForaFaixa() As Integer
    iRecebForaFaixa = mvariRecebForaFaixa
End Property

Public Property Let iRecebForaFaixa(ByVal vData As Integer)
    mvariRecebForaFaixa = vData
End Property

Public Property Get dPercentMaisReceb() As Double
    dPercentMaisReceb = mvardPercentMaisReceb
End Property

Public Property Let dPercentMaisReceb(ByVal vData As Double)
    mvardPercentMaisReceb = vData
End Property

Public Property Get dPercentMenosReceb() As Double
    dPercentMenosReceb = mvardPercentMenosReceb
End Property

Public Property Let dPercentMenosReceb(ByVal vData As Double)
    mvardPercentMenosReceb = vData
End Property

Public Property Set objItensDeContrato(ByVal vData As ClassItensDeContrato)
    Set mvarobjItensDeContrato = vData
End Property

Public Property Get objItensDeContrato() As ClassItensDeContrato
    Set objItensDeContrato = mvarobjItensDeContrato
End Property

Public Property Set colDataEntrega(ByVal vData As Object)
    Set mvarcolDataEntrega = vData
End Property

Public Property Get colDataEntrega() As Collection
    Set colDataEntrega = mvarcolDataEntrega
End Property

Public Property Set objCobrItensContrato(ByVal vData As ClassCobrItensContrato)
    Set mvarobjCobrItensContrato = vData
End Property
Public Property Get objCobrItensContrato() As ClassCobrItensContrato
    Set objCobrItensContrato = mvarobjCobrItensContrato
End Property
'FIM
'###########################################

Public Property Let dPrecoUnitarioMoeda(ByVal vData As Double)
    mvardPrecoUnitarioMoeda = vData
End Property

Public Property Get dPrecoUnitarioMoeda() As Double
    dPrecoUnitarioMoeda = mvardPrecoUnitarioMoeda
End Property

Public Property Set colItensRomaneioGrade(ByVal vData As Collection)
    Set mvarcolItensRomaneioGrade = vData
End Property

Public Property Get colItensRomaneioGrade() As Collection
    Set colItensRomaneioGrade = mvarcolItensRomaneioGrade
End Property

Public Property Get iPossuiGrade() As Integer
    iPossuiGrade = mvariPossuiGrade
End Property

Public Property Let iPossuiGrade(ByVal vData As Integer)
    mvariPossuiGrade = vData
End Property

Public Property Let dValorDescGlobal(ByVal vData As Double)
    mvardValorDescGlobal = vData
End Property

Public Property Get dValorDescGlobal() As Double
    dValorDescGlobal = mvardValorDescGlobal
End Property

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

Public Property Set colItemNFItemRC(ByVal vData As Object)
    Set mvarcolItemNFItemRC = vData
End Property

Public Property Get colItemNFItemRC() As Collection
    Set colItemNFItemRC = mvarcolItemNFItemRC
End Property

Public Property Set colItemNFItemPC(ByVal vData As Object)
    Set mvarcolItemNFItemPC = vData
End Property

Public Property Get colItemNFItemPC() As Collection
    Set colItemNFItemPC = mvarcolItemNFItemPC
End Property

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
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

Public Property Set objTributacaoItemNF(ByVal vData As Object)
    Set mvarobjTribItemNF = vData
End Property

Public Property Get objTributacaoItemNF() As ClassTributacaoDocItem
    Set objTributacaoItemNF = mvarobjTribItemNF
End Property

Public Property Let iItemNFOrig(ByVal vData As Integer)
    mvariItemNFOrig = vData
End Property

Public Property Get iItemNFOrig() As Integer
    iItemNFOrig = mvariItemNFOrig
End Property

Public Property Let sSerieNFOrig(ByVal vData As String)
Attribute sSerieNFOrig.VB_Description = "Série da Nota Fiscal."
    mvarsSerieNFOrig = vData
End Property

Public Property Get sSerieNFOrig() As String
    sSerieNFOrig = mvarsSerieNFOrig
End Property

Public Property Let lNumNFOrig(ByVal vData As Long)
    mvarlNumNFOrig = vData
End Property

Public Property Get lNumNFOrig() As Long
    lNumNFOrig = mvarlNumNFOrig
End Property

Public Property Let lNumIntDocOrig(ByVal vData As Long)
    mvarlNumIntDocOrig = vData
End Property

Public Property Get lNumIntDocOrig() As Long
    lNumIntDocOrig = mvarlNumIntDocOrig
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
Attribute dValorTotal.VB_Description = "Valor Total do ítem de NFiscal."

Dim dValorTotalAux As Double
Dim lErro As Long
Dim objItemNF As ClassItemNF

    Set objItemNF = Me
    
    'Inserido por Wagner 27/04/2006
    lErro = CF("ItemNF_Get_ValorTotal", objItemNF, dValorTotalAux)
    If lErro = SUCESSO Then
        dValorTotal = dValorTotalAux
    Else
        dValorTotal = mvardQuantidade * mvardPrecoUnitario - mvardValorDesconto
    End If
    
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

Public Property Let lNumIntTrib(ByVal vData As Long)
    mvarlNumIntTrib = vData
End Property

Public Property Get lNumIntTrib() As Long
    lNumIntTrib = mvarlNumIntTrib
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

Public Property Let dValorAbatComissao(ByVal vData As Double)
    mvardValorAbatComissao = vData
End Property

Public Property Get dValorAbatComissao() As Double
    dValorAbatComissao = mvardValorAbatComissao
End Property

Public Property Let sDescricaoItem(ByVal vData As String)
    mvarsDescricaoItem = vData
End Property

Public Property Get sDescricaoItem() As String
    sDescricaoItem = mvarsDescricaoItem
End Property

Public Property Let dtDataEntrega(ByVal vData As Date)
    mvardtDataEntrega = vData
End Property

Public Property Get dtDataEntrega() As Date
    dtDataEntrega = mvardtDataEntrega
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

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let lNumIntNF(ByVal vData As Long)
    mvarlNumIntNF = vData
End Property

Public Property Get lNumIntNF() As Long
    lNumIntNF = mvarlNumIntNF
End Property

Sub Inicializa_Tributacao()
'
    Set objTributacaoItemNF = New ClassTributacaoDocItem
    
End Sub

Private Sub Class_Initialize()
    mvariRecebForaFaixa = 1
    mvarobjInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMNF
End Sub

Public Property Set objTributacao(ByVal vData As Object)
    Set mvarobjTribItemNF = vData
End Property

Public Property Get objTributacao() As ClassTributacaoDocItem
    Set objTributacao = mvarobjTribItemNF
End Property
