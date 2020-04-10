VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalEntradaCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjAnotacao As ClassAnotacoes
Public gobjInfoAdic As ClassInfoAdic
Dim dValorDescontoItensAnt As Double
Dim dPercDescontoItensAnt As Double
Dim iDescontoAlterado As Integer

Public gobjInfoUsu As Object

Dim m_objUserControl As Object

Dim gbLimpaTaxa As Boolean

'Property Variables:
Dim m_Caption As String
Event Unload()

'IDHs para HELP
Const IDH_ENTRADA_NFCOM_SIMPLES_DADOS_PRINCIPAIS = 1
Const IDH_ENTRADA_NFCOM_SIMPLES_ITENS = 2
Const IDH_ENTRADA_NFCOM_SIMPLES_COMPLEMENTOS = 3
Const IDH_ENTRADA_NFCOM_SIMPLES_PEDIDOCOMPRA = 4
Const IDH_ENTRADA_NFCOM_SIMPLES_REQUISICAOCOMPRA = 5
Const IDH_ENTRADA_NFCOM_SIMPLES_TRIBUTACAO = 6
Const IDH_ENTRADA_NFCOM_SIMPLES_CONTABILIZACAO = 7

'inicio contabilidade
Public objGrid1 As AdmGrid
Public objContabil As New ClassContabil

Public WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Public WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

Public gobjTelaProjetoInfo As ClassTelaPRJInfo

'Flag que indica se a tela está sendo preenchida.
Public gbCarregandoTela As Boolean
Public gbCarregandoPC As Boolean
Public giPosCargaOk As Integer
Dim giComplemento As Integer

'mnemonicos
Private Const TIPO_NOTA_FISCAL As String = "Tipo_Nota_Fiscal"
Private Const NATUREZA_OP As String = "Natureza_OP"
Private Const FORNECEDOR1 As String = "Fornecedor"
Private Const FILIAL1 As String = "Filial_Fornecedor"
Private Const SERIE1 As String = "Serie"
Private Const NFISCAL1 As String = "Nota_Fiscal"
Private Const DATA_EMISSAO As String = "Data_Emissao"
Private Const DATA_ENTRADA As String = "Data_Entrada"
Private Const PRODUTO1 As String = "Produto_Codigo"
Private Const UNIDADE_MED As String = "Unidade_Medida"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const VALOR_UNITARIO As String = "Valor_Unitario_Prod"
Private Const VALOR_TOTAL As String = "Valor_Total_Produto"
Private Const ALMOXARIFADO1 As String = "Almoxarifado"
Private Const DESCONTO1 As String = "Desconto_Produto"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"
Private Const PRODUTO_ITENS  As String = "Produto_Itens"
Private Const PRODUTO_DIST  As String = "Produto_Dist"
Private Const QUANT_DIST As String = "Quant_Dist"
Private Const FATOR_ITEM_EST As String = "F1"
Private Const PRODUTO_TIPO_EST As String = "Prod_Dist_Tipo"
Private Const PRODUTO_TIPO As String = "Prod_Tipo"
'#######Valores####################
Private Const ICMS As String = "ICMS_Valor"
Private Const ICMSSUBST As String = "ICMSSubst_Valor"
Private Const SUBTOTAL1 As String = "Valor_Produtos"
Private Const VALOR_FRETE As String = "Valor_Frete"
Private Const VALOR_SEGURO As String = "Valor_Seguro"
Private Const VALOR_DESPESAS As String = "Valor_Despesas"
Private Const IPI As String = "IPI_Valor"
Private Const VALOR_DESCONTO As String = "Valor_Desconto_Glb"
Private Const TOTAL1 As String = "Total_Nota_Fiscal"
'###########Tributação##############
Private Const ISS_VALOR As String = "ISS_Valor"
Private Const ISS_INCLUSO As String = "ISS_Incluso"
Private Const VALOR_IRRF As String = "Valor_IRRF"
Private Const VALOR_INSS As String = "Valor_INSS"
Private Const INSS_RETIDO As String = "INSS_Retido"
Private Const CTACONTABILEST1 As String = "ContaContabilEst"
Private Const CTAFILIALFORNECEDOR1 As String = "FilialForn_Conta_Ctb"
Private Const ICMS_CREDITO_VALOR As String = "ICMS_Credito_Valor"
Private Const IPI_CREDITO_VALOR As String = "IPI_Credito_Valor"
Private Const ICMS_CREDITO_FRETE As String = "ICMS_Credito_Frete"
Private Const ICMS_CREDITO_SEGURO As String = "ICMS_Credito_Seguro"
Private Const ICMS_CREDITO_OUTRASDESP As String = "ICMS_Cred_OutrasDesp"
Private Const ICMS_CREDITO_DESCONTO As String = "ICMS_Cred_Desconto"
Private Const IPI_PRODUTO As String = "IPI_Produto"
Private Const IPI_CREDITO_FRETE As String = "IPI_Credito_Frete"
Private Const IPI_CREDITO_SEGURO As String = "IPI_Credito_Seguro"
Private Const IPI_CREDITO_OUTRASDESP As String = "IPI_Cred_OutrasDesp"
Private Const IPI_CREDITO_DESCONTO As String = "IPI_Cred_Desconto"
Private Const ICMS_CREDITO_TOTAL As String = "ICMS_Credito_Total"
Private Const IPI_CREDITO_TOTAL As String = "IPI_Credito_Total"
Private Const PIS_CREDITO_TOTAL As String = "PIS_Credito_Total"
Private Const COFINS_CREDITO_TOTAL As String = "COFINS_Credito_Total"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const ISS_RETIDO As String = "ISS_Retido"
Private Const COFINS_RETIDO As String = "COFINS_Retido"
Private Const CSLL_RETIDO As String = "CSLL_Retido"
Private Const CTACONTABILDIST As String = "ContaContabilDist"
Private Const VALOR_UNIT_DIST As String = "Valor_Unit_Dist"
Private Const ICMS_CRED_VALOR_DIST As String = "ICMS_Cred_Valor_Dist"
Private Const IPI_CRED_VALOR_DIST As String = "IPI_Cred_Valor_Dist"
Private Const IPI_PRODUTO_DIST As String = "IPI_Produto_Dist"
Private Const QUANT_DIST1 As String = "Quant_Dist1"
Private Const CONTROLA_EST As String = "Controla_Est"
Private Const CONTA_DESP_EST_FORN As String = "Conta_Desp_Est_Forn"
Private Const CONTACONTABILAPLIC As String = "ContaContabilAplic"
Private Const CONTACONTABILCATPROD As String = "ContaContabilCatProd"
'fim da contabilidade


'Variáveis Globais
Public iAlterado As Integer
Dim iLinhaAntiga As Integer
Public iFornecedorAlterado As Integer
Public iNaturezaOpAlterada As Integer
Public iFrameAtual As Integer
Private iEmitenteNF As Integer
Public iValorDescontoAlterado As Integer
Public iValorFreteAlterado As Integer
Public iValorSeguroAlterado As Integer
Public iValorDespesasAlterado As Integer
Dim iTipoNFAlterado As Integer
Dim iFornecedorBenefAlterado As Integer
Public iDataEntradaAlterada As Integer
Public iSerieAlterada As Integer

Dim gFornecedorAnterior As String
Dim gcolPedidoCompra As Collection
Dim gcolItemPedCompraInfo As Collection
Dim gcolItemReqCompraInfo As Collection

'Rastreamento
'Public gcolItensNF As Collection
'Public objGenerico As AdmGenerico
Public gobjRastreamento As ClassRastreamento

'distribuicao
Public gobjDistribuicao As Object

'************ GRADE **********************
Public gobjNFiscal As ClassNFiscal
'*****************************************

Dim iFilialAnterior As Integer
Dim iFilialCompraAnterior As Integer
Dim sFornecedorAnterior As String
'Dim lNumeroNotaAnterior As Long

Dim giRecebimento As Integer
Dim lNumIntDocNF As Long

'GridItens
Public objGridItens As AdmGrid
Public iGrid_Produto_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
'distribuicao
'Public iGrid_Almoxarifado_Col As Integer
Public iGrid_Ccl_Col As Integer
Public iGrid_ValorUnitario_Col As Integer
Public iGrid_ValorUnitarioMoeda_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PrecoTotal_Col As Integer
Public iGrid_PrecoTotalB_Col As Integer

'GridItensPC
Public objGridItensPC As AdmGrid
Public iGrid_PedCompra_Col As Integer
Public iGrid_Item_Col As Integer
Public iGrid_Prod_Col As Integer
Public iGrid_DescProdutoPC_Col As Integer
Public iGrid_UM_Col As Integer
Public iGrid_AReceber_Col As Integer
Public iGrid_Recebido_Col As Integer
Public iGrid_Recebido_RS_Col As Integer
Public iGrid_PrecoUnitario_Col As Integer
Public iGrid_Moeda_Col As Integer
Public iGrid_Taxa_Col As Integer

'GridItensRC
Dim objGridItensRC As AdmGrid
Dim iGrid_PedCompraRC_Col As Integer
Dim iGrid_ItemPC_Col As Integer
Dim iGrid_ProdutoRC_Col As Integer
Dim iGrid_DescProdutoRC_Col As Integer
Dim iGrid_UMRC_Col As Integer
Dim iGrid_QuantRecebidaPC_Col As Integer
Dim iGrid_FilialReq_Col As Integer
Dim iGrid_ReqCompra_Col As Integer
Dim iGrid_Urgente_Col As Integer
Dim iGrid_ItemRC_Col As Integer
Dim iGrid_AReceberRC_Col As Integer
Dim iGrid_RecebidoRC_Col As Integer

'Eventos dos Browses
Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Public WithEvents objEventoNFiscal As AdmEvento
Attribute objEventoNFiscal.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoRecebimento As AdmEvento
Attribute objEventoRecebimento.VB_VarHelpID = -1
Private WithEvents objEventoSerieOriginal As AdmEvento
Attribute objEventoSerieOriginal.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoNFiscalOriginal As AdmEvento
Attribute objEventoNFiscalOriginal.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorBenef As AdmEvento
Attribute objEventoFornecedorBenef.VB_VarHelpID = -1
Private WithEvents objEventoChvNFe As AdmEvento
Attribute objEventoChvNFe.VB_VarHelpID = -1

Dim gsChvAnt As String

'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_Itens = 2
Private Const TAB_PedidoCompras = 3
Private Const TAB_RequisicaoCompras = 4
Private TAB_Complemento As Integer
Private TAB_Distribuicao As Integer
Private TAB_Tributacao As Integer
Private Const TAB_Contabilizacao = 8
Private Const TAB_RASTRO = 9

''*** incluidos p/tratamento de tributacao *******************************
Public gobjTribTab As ClassTribTab

'--- inicio dos properties get dos controles da tela
Public Property Get ChvNFe() As Object
     Set ChvNFe = objUserControl.Controls("ChvNFe")
End Property

Public Property Get ValorProdutos2() As Object
     Set ValorProdutos2 = objUserControl.Controls("ValorProdutos2")
End Property

Public Property Get ValorDescontoItens() As Object
     Set ValorDescontoItens = objUserControl.Controls("ValorDescontoItens")
End Property

Public Property Get PercDescontoItens() As Object
     Set PercDescontoItens = objUserControl.Controls("PercDescontoItens")
End Property

Public Property Get ISSBase1() As Object
     Set ISSBase1 = objUserControl.Controls("ISSBase1")
End Property

Public Property Get PrecoTotalB() As Object
     Set PrecoTotalB = objUserControl.Controls("PrecoTotalB")
End Property

Public Property Get MsgAutomatica() As Object
     Set MsgAutomatica = objUserControl.Controls("MsgAutomatica")
End Property

Public Property Get MensagemCorpo() As Object
     Set MensagemCorpo = objUserControl.Controls("MensagemCorpo")
End Property

Public Property Get EletronicaFed() As Object
     Set EletronicaFed = objUserControl.Controls("EletronicaFed")
End Property

Public Property Get ComboPedidoCompras() As Object
     Set ComboPedidoCompras = objUserControl.Controls("ComboPedidoCompras")
End Property

Public Property Get Moeda() As Object
     Set Moeda = objUserControl.Controls("Moeda")
End Property

Public Property Get Taxa() As Object
     Set Taxa = objUserControl.Controls("Taxa")
End Property

Public Property Get LabelTaxa() As Object
     Set LabelTaxa = objUserControl.Controls("LabelTaxa")
End Property

Public Property Get PrecoUnitario() As Object
     Set PrecoUnitario = objUserControl.Controls("PrecoUnitario")
End Property

Public Property Get MoedaGrid() As Object
     Set MoedaGrid = objUserControl.Controls("MoedaGrid")
End Property

Public Property Get TaxaGrid() As Object
     Set TaxaGrid = objUserControl.Controls("TaxaGrid")
End Property

Public Property Get ValorRecebido() As Object
     Set ValorRecebido = objUserControl.Controls("ValorRecebido")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get Serie() As Object
     Set Serie = objUserControl.Controls("Serie")
End Property

Public Property Get TipoNFiscal() As Object
     Set TipoNFiscal = objUserControl.Controls("TipoNFiscal")
End Property

Public Property Get BotaoRecebimentos() As Object
     Set BotaoRecebimentos = objUserControl.Controls("BotaoRecebimentos")
End Property

Public Property Get BotaoLimparNF() As Object
     Set BotaoLimparNF = objUserControl.Controls("BotaoLimparNF")
End Property

Public Property Get NFiscal() As Object
     Set NFiscal = objUserControl.Controls("NFiscal")
End Property

Public Property Get NaturezaOp() As Object
     Set NaturezaOp = objUserControl.Controls("NaturezaOp")
End Property

Public Property Get NumRecebimento() As Object
     Set NumRecebimento = objUserControl.Controls("NumRecebimento")
End Property

Public Property Get NaturezaLabel() As Object
     Set NaturezaLabel = objUserControl.Controls("NaturezaLabel")
End Property

Public Property Get SerieLabel() As Object
     Set SerieLabel = objUserControl.Controls("SerieLabel")
End Property

Public Property Get NFiscalLabel() As Object
     Set NFiscalLabel = objUserControl.Controls("NFiscalLabel")
End Property

Public Property Get NFiscalInterna() As Object
     Set NFiscalInterna = objUserControl.Controls("NFiscalInterna")
End Property

Public Property Get RecebimentoLabel() As Object
     Set RecebimentoLabel = objUserControl.Controls("RecebimentoLabel")
End Property

Public Property Get Status() As Object
     Set Status = objUserControl.Controls("Status")
End Property

Public Property Get Frame14() As Object
     Set Frame14 = objUserControl.Controls("Frame14")
End Property

Public Property Get UpDownEmissao() As Object
     Set UpDownEmissao = objUserControl.Controls("UpDownEmissao")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get UpDownEntrada() As Object
     Set UpDownEntrada = objUserControl.Controls("UpDownEntrada")
End Property

Public Property Get DataEntrada() As Object
     Set DataEntrada = objUserControl.Controls("DataEntrada")
End Property

'horaentrada
Public Property Get HoraEntrada() As Object
     Set HoraEntrada = objUserControl.Controls("HoraEntrada")
End Property

Public Property Get UpDownVencimento() As Object
     Set UpDownVencimento = objUserControl.Controls("UpDownVencimento")
End Property

Public Property Get DataVencimento() As Object
     Set DataVencimento = objUserControl.Controls("DataVencimento")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Fornecedor() As Object
     Set Fornecedor = objUserControl.Controls("Fornecedor")
End Property

Public Property Get FornecedorLabel() As Object
     Set FornecedorLabel = objUserControl.Controls("FornecedorLabel")
End Property

Public Property Get BotaoMarcarTodos() As Object
     Set BotaoMarcarTodos = objUserControl.Controls("BotaoMarcarTodos")
End Property

Public Property Get BotaoDesmarcarTodos() As Object
     Set BotaoDesmarcarTodos = objUserControl.Controls("BotaoDesmarcarTodos")
End Property

Public Property Get FilialCompra() As Object
     Set FilialCompra = objUserControl.Controls("FilialCompra")
End Property

Public Property Get PedidosCompra() As Object
     Set PedidosCompra = objUserControl.Controls("PedidosCompra")
End Property

Public Property Get BotaoCcls() As Object
     Set BotaoCcls = objUserControl.Controls("BotaoCcls")
End Property

Public Property Get BotaoConsultaNFPag() As Object
     Set BotaoConsultaNFPag = objUserControl.Controls("BotaoConsultaNFPag")
End Property

Public Property Get DescricaoItem() As Object
     Set DescricaoItem = objUserControl.Controls("DescricaoItem")
End Property

Public Property Get UnidadeMed() As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

Public Property Get Ccl() As Object
     Set Ccl = objUserControl.Controls("Ccl")
End Property

Public Property Get Desconto() As Object
     Set Desconto = objUserControl.Controls("Desconto")
End Property

Public Property Get PercentDesc() As Object
     Set PercentDesc = objUserControl.Controls("PercentDesc")
End Property

Public Property Get ValorUnitario() As Object
     Set ValorUnitario = objUserControl.Controls("ValorUnitario")
End Property

Public Property Get Quantidade() As Object
     Set Quantidade = objUserControl.Controls("Quantidade")
End Property

Public Property Get ValorTotal() As Object
     Set ValorTotal = objUserControl.Controls("ValorTotal")
End Property

Public Property Get GridItens() As Object
     Set GridItens = objUserControl.Controls("GridItens")
End Property

Public Property Get Total() As Object
     Set Total = objUserControl.Controls("Total")
End Property

Public Property Get ValorFrete() As Object
     Set ValorFrete = objUserControl.Controls("ValorFrete")
End Property

Public Property Get ValorDesconto() As Object
     Set ValorDesconto = objUserControl.Controls("ValorDesconto")
End Property

Public Property Get ValorDespesas() As Object
     Set ValorDespesas = objUserControl.Controls("ValorDespesas")
End Property

Public Property Get ValorSeguro() As Object
     Set ValorSeguro = objUserControl.Controls("ValorSeguro")
End Property

Public Property Get LabelTotais() As Object
     Set LabelTotais = objUserControl.Controls("LabelTotais")
End Property

Public Property Get IPIValor1() As Object
     Set IPIValor1 = objUserControl.Controls("IPIValor1")
End Property

Public Property Get ISSValor1() As Object
     Set ISSValor1 = objUserControl.Controls("ISSValor1")
End Property

Public Property Get SubTotal() As Object
     Set SubTotal = objUserControl.Controls("SubTotal")
End Property

Public Property Get ICMSBase1() As Object
     Set ICMSBase1 = objUserControl.Controls("ICMSBase1")
End Property

Public Property Get ICMSValor1() As Object
     Set ICMSValor1 = objUserControl.Controls("ICMSValor1")
End Property

Public Property Get ICMSSubstBase1() As Object
     Set ICMSSubstBase1 = objUserControl.Controls("ICMSSubstBase1")
End Property

Public Property Get ICMSSubstValor1() As Object
     Set ICMSSubstValor1 = objUserControl.Controls("ICMSSubstValor1")
End Property

Public Property Get BotaoPedidoCompra() As Object
     Set BotaoPedidoCompra = objUserControl.Controls("BotaoPedidoCompra")
End Property

Public Property Get DescProdutoPC() As Object
     Set DescProdutoPC = objUserControl.Controls("DescProdutoPC")
End Property

Public Property Get QuantRecebidaPC() As Object
     Set QuantRecebidaPC = objUserControl.Controls("QuantRecebidaPC")
End Property

Public Property Get UMPC() As Object
     Set UMPC = objUserControl.Controls("UMPC")
End Property

Public Property Get QuantAReceberPC() As Object
     Set QuantAReceberPC = objUserControl.Controls("QuantAReceberPC")
End Property

Public Property Get ProdutoPC() As Object
     Set ProdutoPC = objUserControl.Controls("ProdutoPC")
End Property

Public Property Get CodigoPC() As Object
     Set CodigoPC = objUserControl.Controls("CodigoPC")
End Property

Public Property Get ItemPC() As Object
     Set ItemPC = objUserControl.Controls("ItemPC")
End Property

Public Property Get GridItensPC() As Object
     Set GridItensPC = objUserControl.Controls("GridItensPC")
End Property

Public Property Get RequisicoesCompra() As Object
     Set RequisicoesCompra = objUserControl.Controls("RequisicoesCompra")
End Property

Public Property Get Urgente() As Object
     Set Urgente = objUserControl.Controls("Urgente")
End Property

Public Property Get ItemRC() As Object
     Set ItemRC = objUserControl.Controls("ItemRC")
End Property

Public Property Get FilialReqRC() As Object
     Set FilialReqRC = objUserControl.Controls("FilialReqRC")
End Property

Public Property Get CodigoReq() As Object
     Set CodigoReq = objUserControl.Controls("CodigoReq")
End Property

Public Property Get QuantRecebidaRC() As Object
     Set QuantRecebidaRC = objUserControl.Controls("QuantRecebidaRC")
End Property

Public Property Get QuantAReceberRC() As Object
     Set QuantAReceberRC = objUserControl.Controls("QuantAReceberRC")
End Property

Public Property Get DescProdutoRC() As Object
     Set DescProdutoRC = objUserControl.Controls("DescProdutoRC")
End Property

Public Property Get QuantRecebidaPCRC() As Object
     Set QuantRecebidaPCRC = objUserControl.Controls("QuantRecebidaPCRC")
End Property

Public Property Get ItemPCRC() As Object
     Set ItemPCRC = objUserControl.Controls("ItemPCRC")
End Property

Public Property Get UMRC() As Object
     Set UMRC = objUserControl.Controls("UMRC")
End Property

Public Property Get ProdutoRC() As Object
     Set ProdutoRC = objUserControl.Controls("ProdutoRC")
End Property

Public Property Get CodigoPCRC() As Object
     Set CodigoPCRC = objUserControl.Controls("CodigoPCRC")
End Property

Public Property Get GridItensRC() As Object
     Set GridItensRC = objUserControl.Controls("GridItensRC")
End Property

Public Property Get BotaoReqCompra() As Object
     Set BotaoReqCompra = objUserControl.Controls("BotaoReqCompra")
End Property

Public Property Get FilialDeCompra() As Object
     Set FilialDeCompra = objUserControl.Controls("FilialDeCompra")
End Property

Public Property Get SerieNFiscalOriginal() As Object
     Set SerieNFiscalOriginal = objUserControl.Controls("SerieNFiscalOriginal")
End Property

Public Property Get NFiscalOriginal() As Object
     Set NFiscalOriginal = objUserControl.Controls("NFiscalOriginal")
End Property

Public Property Get SerieOriginalLabel() As Object
     Set SerieOriginalLabel = objUserControl.Controls("SerieOriginalLabel")
End Property

Public Property Get NFiscalOriginalLabel() As Object
     Set NFiscalOriginalLabel = objUserControl.Controls("NFiscalOriginalLabel")
End Property

Public Property Get VolumeMarca() As Object
     Set VolumeMarca = objUserControl.Controls("VolumeMarca")
End Property

Public Property Get VolumeEspecie() As Object
     Set VolumeEspecie = objUserControl.Controls("VolumeEspecie")
End Property

Public Property Get VolumeQuant() As Object
     Set VolumeQuant = objUserControl.Controls("VolumeQuant")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get Mensagem() As Object
     Set Mensagem = objUserControl.Controls("Mensagem")
End Property

Public Property Get PesoLiquido() As Object
     Set PesoLiquido = objUserControl.Controls("PesoLiquido")
End Property

Public Property Get PesoBruto() As Object
     Set PesoBruto = objUserControl.Controls("PesoBruto")
End Property

Public Property Get Transportadora() As Object
     Set Transportadora = objUserControl.Controls("Transportadora")
End Property

Public Property Get Placa() As Object
     Set Placa = objUserControl.Controls("Placa")
End Property

Public Property Get PlacaUF() As Object
     Set PlacaUF = objUserControl.Controls("PlacaUF")
End Property

Public Property Get TransportadoraLabel() As Object
     Set TransportadoraLabel = objUserControl.Controls("TransportadoraLabel")
End Property

Public Property Get CTBLancAutomatico() As Object
     Set CTBLancAutomatico = objUserControl.Controls("CTBLancAutomatico")
End Property

Public Property Get CTBBotaoImprimir() As Object
     Set CTBBotaoImprimir = objUserControl.Controls("CTBBotaoImprimir")
End Property

Public Property Get CTBFrame7() As Object
     Set CTBFrame7 = objUserControl.Controls("CTBFrame7")
End Property

Public Property Get CTBCclDescricao() As Object
     Set CTBCclDescricao = objUserControl.Controls("CTBCclDescricao")
End Property

Public Property Get CTBContaDescricao() As Object
     Set CTBContaDescricao = objUserControl.Controls("CTBContaDescricao")
End Property

Public Property Get CTBLabel7() As Object
     Set CTBLabel7 = objUserControl.Controls("CTBLabel7")
End Property

Public Property Get CTBCclLabel() As Object
     Set CTBCclLabel = objUserControl.Controls("CTBCclLabel")
End Property

Public Property Get CTBModelo() As Object
     Set CTBModelo = objUserControl.Controls("CTBModelo")
End Property

Public Property Get CTBBotaoLimparGrid() As Object
     Set CTBBotaoLimparGrid = objUserControl.Controls("CTBBotaoLimparGrid")
End Property

Public Property Get CTBBotaoModeloPadrao() As Object
     Set CTBBotaoModeloPadrao = objUserControl.Controls("CTBBotaoModeloPadrao")
End Property

Public Property Get CTBListHistoricos() As Object
     Set CTBListHistoricos = objUserControl.Controls("CTBListHistoricos")
End Property

Public Property Get CTBHistorico() As Object
     Set CTBHistorico = objUserControl.Controls("CTBHistorico")
End Property

Public Property Get CTBAglutina() As Object
     Set CTBAglutina = objUserControl.Controls("CTBAglutina")
End Property

Public Property Get CTBConta() As Object
     Set CTBConta = objUserControl.Controls("CTBConta")
End Property

Public Property Get CTBDebito() As Object
     Set CTBDebito = objUserControl.Controls("CTBDebito")
End Property

Public Property Get CTBCredito() As Object
     Set CTBCredito = objUserControl.Controls("CTBCredito")
End Property

Public Property Get CTBCcl() As Object
     Set CTBCcl = objUserControl.Controls("CTBCcl")
End Property

Public Property Get CTBUpDown() As Object
     Set CTBUpDown = objUserControl.Controls("CTBUpDown")
End Property

Public Property Get CTBDataContabil() As Object
     Set CTBDataContabil = objUserControl.Controls("CTBDataContabil")
End Property

Public Property Get CTBLote() As Object
     Set CTBLote = objUserControl.Controls("CTBLote")
End Property

Public Property Get CTBDocumento() As Object
     Set CTBDocumento = objUserControl.Controls("CTBDocumento")
End Property

Public Property Get CTBGridContabil() As Object
     Set CTBGridContabil = objUserControl.Controls("CTBGridContabil")
End Property

Public Property Get CTBTvwCcls() As Object
     Set CTBTvwCcls = objUserControl.Controls("CTBTvwCcls")
End Property

Public Property Get CTBTvwContas() As Object
     Set CTBTvwContas = objUserControl.Controls("CTBTvwContas")
End Property

Public Property Get CTBLabelLote() As Object
     Set CTBLabelLote = objUserControl.Controls("CTBLabelLote")
End Property

Public Property Get CTBLabelDoc() As Object
     Set CTBLabelDoc = objUserControl.Controls("CTBLabelDoc")
End Property

Public Property Get CTBLabel8() As Object
     Set CTBLabel8 = objUserControl.Controls("CTBLabel8")
End Property

Public Property Get CTBTotalCredito() As Object
     Set CTBTotalCredito = objUserControl.Controls("CTBTotalCredito")
End Property

Public Property Get CTBTotalDebito() As Object
     Set CTBTotalDebito = objUserControl.Controls("CTBTotalDebito")
End Property

Public Property Get CTBLabelTotais() As Object
     Set CTBLabelTotais = objUserControl.Controls("CTBLabelTotais")
End Property

Public Property Get CTBLabel1() As Object
     Set CTBLabel1 = objUserControl.Controls("CTBLabel1")
End Property

Public Property Get CTBLabelCcl() As Object
     Set CTBLabelCcl = objUserControl.Controls("CTBLabelCcl")
End Property

Public Property Get CTBLabelContas() As Object
     Set CTBLabelContas = objUserControl.Controls("CTBLabelContas")
End Property

Public Property Get CTBLabelHistoricos() As Object
     Set CTBLabelHistoricos = objUserControl.Controls("CTBLabelHistoricos")
End Property

Public Property Get CTBLabel5() As Object
     Set CTBLabel5 = objUserControl.Controls("CTBLabel5")
End Property

Public Property Get CTBLabel13() As Object
     Set CTBLabel13 = objUserControl.Controls("CTBLabel13")
End Property

Public Property Get CTBExercicio() As Object
     Set CTBExercicio = objUserControl.Controls("CTBExercicio")
End Property

Public Property Get CTBPeriodo() As Object
     Set CTBPeriodo = objUserControl.Controls("CTBPeriodo")
End Property

Public Property Get CTBLabel14() As Object
     Set CTBLabel14 = objUserControl.Controls("CTBLabel14")
End Property

Public Property Get CTBOrigem() As Object
     Set CTBOrigem = objUserControl.Controls("CTBOrigem")
End Property

Public Property Get CTBLabel21() As Object
     Set CTBLabel21 = objUserControl.Controls("CTBLabel21")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

Public Property Get CTBSeqContraPartida() As Object
     Set CTBSeqContraPartida = objUserControl.Controls("CTBSeqContraPartida")
End Property

Public Property Get TipoFrete() As Object
     Set TipoFrete = objUserControl.Controls("TipoFrete")
End Property

Public Property Get VolumeNumero() As Object
     Set VolumeNumero = objUserControl.Controls("VolumeNumero")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

'Rastreamento
Public Property Get GridRastro() As Object
     Set GridRastro = objUserControl.Controls("GridRastro")
End Property

'Distribuicao
Public Property Get GridDist() As Object
     Set GridDist = objUserControl.Controls("GridDist")
End Property

Public Property Get FilialFornBenef() As Object
     Set FilialFornBenef = objUserControl.Controls("FilialFornBenef")
End Property

Public Property Get FornecedorBenef() As Object
     Set FornecedorBenef = objUserControl.Controls("FornecedorBenef")
End Property

Public Property Get FornecedorBenefLabel() As Object
     Set FornecedorBenefLabel = objUserControl.Controls("FornecedorBenefLabel")
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

Public Property Get NumNFe() As Object
     Set NumNFe = objUserControl.Controls("NumNFe")
End Property

'--- fim dos properties get dos controles da tela

Public Function CargaPosFormLoad(Optional bTrazendoDoc As Boolean = False) As Long

Dim iSubTipo As Integer
Dim lErro As Long
Dim sMascaraCcl As String
Dim iIndice As Integer
Dim objFiliais As AdmFiliais
Dim colFiliais As New Collection
Dim iCodigo As Integer
Dim objConfiguraCOM As New ClassConfiguraCOM
Dim bCancel As Boolean

On Error GoTo Erro_CargaPosFormLoad

    If (giPosCargaOk = 0) Then

        'p/permitir o redesenho da tela
        DoEvents
        
        gbCarregandoTela = True

        ValorUnitario.Format = FORMATO_PRECO_UNITARIO_EXTERNO
        
        'Inicializa as variáveis globais da tela(objGrid, objEventos)
        Set objGridItens = New AdmGrid
        Set objEventoSerie = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoNFiscal = New AdmEvento
        Set objEventoNatureza = New AdmEvento
        Set objEventoFornecedor = New AdmEvento
        Set objEventoRecebimento = New AdmEvento
        Set objEventoSerieOriginal = New AdmEvento
        Set objEventoNFiscalOriginal = New AdmEvento
        Set objEventoTransportadora = New AdmEvento
        Set objEventoCcl = New AdmEvento
        Set objEventoFornecedorBenef = New AdmEvento
        Set objEventoChvNFe = New AdmEvento
    
        Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
        Set gobjTelaProjetoInfo.objUserControl = objUserControl
        Set gobjTelaProjetoInfo.objTela = Me
     
        'Verifica se módulo de Compras não está ativo
        If gcolModulo.Ativo(MODULO_COMPRAS) <> MODULO_ATIVO Then gError 74847
        
        'Verifica se o módulo de Estoque não está ativo
        If gcolModulo.Ativo(MODULO_ESTOQUE) <> MODULO_ATIVO Then gError 74848
    
        ''*** incluidos p/tratamento de tributacao *******************************
        Set gobjTribTab = New ClassTribTab
        lErro = gobjTribTab.Ativar(Me, 1)
        If lErro <> SUCESSO Then gError 65615
        '*** fim tributacao
    
        'Rastreamento
        Set gobjRastreamento = New ClassRastreamento
        Set gobjRastreamento.objTela = Me
    
        'distribuicao
        Set gobjDistribuicao = CreateObject("RotinasMat.ClassMATDist")
        Set gobjDistribuicao.objTela = Me
        Set gobjDistribuicao.objRastreamento = gobjRastreamento
    
        'Carrega os Tipos de Documentos relacionadas à tela
        lErro = Carrega_TiposDocInfo()
        If lErro <> SUCESSO Then gError 65616
    
        'Carrega as Combos que armazenam as Séries
        lErro = Carrega_Serie()
        If lErro <> SUCESSO Then gError 65617
    
        'Carrega a combo PlacaUF com as siglas das Estados
        lErro = Carrega_PlacaUF()
        If lErro <> SUCESSO Then gError 65618
    
        'Carrega a combo com o Código e o Nome Reduzido das Transportadoras do BD
        lErro = Carrega_Transportadora()
        If lErro <> SUCESSO Then gError 65619
    
        'Carrega a combo VolumeEspecie
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
        If lErro <> SUCESSO Then gError 102452
    
        'Carrega a combo VolumeMarca
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
        If lErro <> SUCESSO Then gError 102453
    
        'Carrega a combo de FiliaisCompra com as Filiais Empresa
        lErro = CF("FiliaisEmpresas_Le_Empresa", glEmpresa, colFiliais)
        If lErro <> SUCESSO Then gError 65621
    
        For Each objFiliais In colFiliais
            If objFiliais.iCodFilial <> 0 Then
                FilialCompra.AddItem CStr(objFiliais.iCodFilial) & SEPARADOR & objFiliais.sNome
                FilialCompra.ItemData(FilialCompra.NewIndex) = objFiliais.iCodFilial
            End If
        Next
    
        'Limpa a combo de pedidos de compra
        PedidosCompra.Clear
    
        'Leitura da tabela de ComprasConfig
        lErro = CF("ComprasConfig_Le", objConfiguraCOM)
        If lErro <> SUCESSO Then gError 65622
    
        'Inicializa a mascara do CCL
        sMascaraCcl = String(STRING_CCL, 0)
    
        'le a mascara dos centros de custo/lucro
        lErro = MascaraCcl(sMascaraCcl)
        If lErro <> SUCESSO Then gError 65623
    
        Ccl.Mask = sMascaraCcl
    
        'Formata a Quantidade para o Formato de Estoque
        Quantidade.Format = FORMATO_ESTOQUE
    
        'Inicializa a Máscara de Produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoPC)
        If lErro <> SUCESSO Then gError 89168
        
        'Inicializa a Máscara de Produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoRC)
        If lErro <> SUCESSO Then gError 89169
        
        Set objGridItens = New AdmGrid
        Set objGridItensPC = New AdmGrid
        Set objGridItensRC = New AdmGrid
    
        'Inicializa GridItens
        lErro = CF2(Me, "NFEntCom_Inicializa_Grid_Itens", objGridItens)
        If lErro <> SUCESSO Then gError 65624
    
        'Inicializa GridItensPC
        lErro = Inicializa_Grid_ItensPC(objGridItensPC)
        If lErro <> SUCESSO Then gError 65625
    
        'Inicializa GridItensRC
        lErro = Inicializa_Grid_ItensRC(objGridItensRC)
        If lErro <> SUCESSO Then gError 65626
        
        
        iFilialCompraAnterior = Codigo_Extrai(FilialCompra.Text)
    
        Set gcolPedidoCompra = New Collection
        Set gcolItemPedCompraInfo = New Collection
        Set gcolItemReqCompraInfo = New Collection
        
        'Coloca FilialCompra Default na tela
        If objConfiguraCOM.iFilialCompra > 0 Then
            FilialCompra.Text = objConfiguraCOM.iFilialCompra
        Else
            FilialCompra.Text = giFilialEmpresa
        End If
    
        Call FilialCompra_Validate(bCancel)
        
        'Inicializa a contabilidade carregando apenas os modelos contábeis referentes ao TipoDocInfo default
        If objContabil.objGridInt Is Nothing Then
    
            'Obtém o código do subtipo contábil referente ao TipoDocInfo default
            lErro = CF("Obtem_Subtipo_Contabil", "TiposDocInfo", DOCINFO_NFEENCO, iSubTipo)
            If lErro <> SUCESSO And lErro <> 79671 Then gError 79694
            
            'Se não encontrou o subtipo => erro
            If lErro = 79671 Then gError 79695
                
            'Inicialização da parte de contabilidade
            lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_ESTOQUE, iSubTipo)
            If lErro <> SUCESSO Then gError 79370
    
        End If
        
        
        'Inicializa o grid de Rastreamento
        lErro = gobjRastreamento.Inicializa_Grid_Rastreamento()
        If lErro <> SUCESSO Then gError 83274
    
        'Inicializa o grid de Distribuicao
        lErro = gobjDistribuicao.Inicializa_GridDist()
        If lErro <> SUCESSO Then gError 89541
    
        gbCarregandoTela = False
        
        ComboPedidoCompras.ListIndex = 0
        
        iAlterado = 0
        iDescontoAlterado = 0
    
        giPosCargaOk = 1
    
        Call ValorTotal_Calcula

    End If
    
    CargaPosFormLoad = SUCESSO
    
    Exit Function
     
Erro_CargaPosFormLoad:

    gbCarregandoTela = False
    
    CargaPosFormLoad = gErr
    
    Select Case gErr

        Case 65615, 65616, 65617, 65618, 65619, 65621, 65622, 65623, 65624, 65625, 65626, 65627, 79370, 79694, 83274, 89168, 89169, 89541, 102452, 102453

        Case 74847
            Call Rotina_Erro(vbOKOnly, "ERRO_MODULO_COMPRAS_INATIVO", gErr)
                
        Case 74848
            Call Rotina_Erro(vbOKOnly, "ERRO_MODULO_ESTOQUE_INATIVO", gErr)
            
        Case 79695
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPODOCINFO_NAO_ENCONTRADO", gErr, DOCINFO_NFEENCO)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156653)
     
    End Select
     
    iAlterado = 0
     
    Exit Function

End Function

Public Sub FornecedorBenef_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorBenefAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FornecedorBenef_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedorBenef As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_FornecedorBenef_Validate

    If iFornecedorBenefAlterado = 1 Then

        If Len(Trim(FornecedorBenef.Text)) > 0 Then
    
            'Tenta ler o FornecedorBenef (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(FornecedorBenef, objFornecedorBenef, iCodFilial)
            If lErro <> SUCESSO Then Error 52489
    
            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedorBenef, colCodigoNome)
            If lErro <> SUCESSO Then Error 52490
    
            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", FilialFornBenef, colCodigoNome)
    
            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", FilialFornBenef, iCodFilial)
    
    
        ElseIf Len(Trim(FornecedorBenef.Text)) = 0 Then
    
            FilialFornBenef.Clear
    
        End If
        
        iFornecedorBenefAlterado = 0

    End If
    
    Exit Sub

Erro_FornecedorBenef_Validate:

    Cancel = True

    Select Case Err

        Case 52489, 52490

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156654)

    End Select

    Exit Sub

End Sub

Public Sub FilialFornBenef_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub FilialFornBenef_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialFornBenef_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialFornBenef.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialFornBenef.Text = FilialFornBenef.List(FilialFornBenef.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialFornBenef, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 89173

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(FornecedorBenef.Text)) = 0 Then gError 89174

        sFornecedor = FornecedorBenef.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 89175

        If lErro = 18272 Then gError 89176

        'coloca na tela
        FilialFornBenef.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 89177

    Exit Sub

Erro_FilialFornBenef_Validate:

    Cancel = True

    Select Case gErr

        Case 89173, 89175
        
        Case 89174
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 89176
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, FornecedorBenef.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 89177
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, FilialFornBenef.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156655)

    End Select

    Exit Sub

End Sub

Public Sub FornecedorBenefLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = FornecedorBenef.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorBenef)

    Exit Sub

End Sub

Public Sub Destinatario_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub


Public Sub TipoFrete_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub BotaoDesmarcarTodos_Click(Index As Integer)

Dim iIndice As Integer
    
    'Se foi clicado no Botão de Pedido de Compras
    If Index = 0 Then
        For iIndice = 0 To PedidosCompra.ListCount - 1
            PedidosCompra.Selected(iIndice) = False
        Next
    
    'Se foi clicado no Botão de Requisição de Compras
    ElseIf Index = 1 Then
        For iIndice = 0 To RequisicoesCompra.ListCount - 1
            RequisicoesCompra.Selected(iIndice) = False
        Next
    End If
    
    ComboPedidoCompras.ListIndex = 0

End Sub

Public Sub BotaoLimparNF_Click()

    NFiscalInterna.Caption = ""

End Sub

Public Sub BotaoCcls_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As ClassCcl

On Error GoTo Erro_BotaoCcls_Click

    'Se nenhuma linha foi selecionada, Erro
    If GridItens.Row = 0 Then gError 65605

    'Se o Produto da linha selecionada não foi preenchida, erro
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then gError 65606
    
    'Chama a tela "CclLista"
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

    Exit Sub
    
Erro_BotaoCcls_Click:
    
    Select Case gErr
    
        Case 65605
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 65606
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156656)

    End Select
    
    Exit Sub

End Sub

Public Sub BotaoMarcarTodos_Click(Index As Integer)
    
Dim iIndice As Integer

    'Se foi clicado em Marcar Todos Pedidos de Compras
    If Index = 0 Then
        For iIndice = 0 To PedidosCompra.ListCount - 1
            PedidosCompra.Selected(iIndice) = True
        Next
    
    'Se foi clicado em Marcar Todas Requisições de Compras
    ElseIf Index = 1 Then
        For iIndice = 0 To RequisicoesCompra.ListCount - 1
            RequisicoesCompra.Selected(iIndice) = True
        Next
    End If
    
    ComboPedidoCompras.ListIndex = 0
    
End Sub

Public Sub BotaoPedidoCompra_Click()

Dim objPedidoCompra As New ClassPedidoCompras
Dim lErro As Long

On Error GoTo Erro_BotaoPedidoCompra_Click

    'Se nenhuma linha do Grid de Pedido de Compras foi selecionada, ==> avisar
    If GridItensPC.Row = 0 Then gError 89185

    objPedidoCompra.lCodigo = StrParaLong(GridItensPC.TextMatrix(GridItensPC.Row, iGrid_PedCompra_Col))
    objPedidoCompra.iFilialEmpresa = Codigo_Extrai(FilialCompra.Text)

    'Chama a tela "PedComprasCons"
    Call Chama_Tela("PedComprasCons", objPedidoCompra)

    Exit Sub
    
Erro_BotaoPedidoCompra_Click:

    Select Case gErr
    
        Case 89185
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156658)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoReqCompra_Click()

Dim objRequisicaoCompra As New ClassRequisicaoCompras
Dim lErro As Long, iIndice As Integer, objItemRCInfoAux As ClassItemReqComprasInfo

On Error GoTo Erro_BotaoReqCompra_Click

    'Se nenhuma linha do Grid de Requisição de Compras foi selecionada, sai da rotina
    If GridItensRC.Row = 0 Then gError 89187

    objRequisicaoCompra.lCodigo = CLng(GridItensRC.TextMatrix(GridItensRC.Row, iGrid_ReqCompra_Col))
    iIndice = GridItensRC.Row
    If iIndice <= gcolItemReqCompraInfo.Count Then
        Set objItemRCInfoAux = gcolItemReqCompraInfo(iIndice)
        objRequisicaoCompra.iFilialEmpresa = objItemRCInfoAux.iFilialEmpresa
    Else
        objRequisicaoCompra.iFilialEmpresa = Codigo_Extrai(FilialCompra.Text)
    End If

    'Chama a tela "ReqComprasCons"
    Call Chama_Tela("ReqComprasCons", objRequisicaoCompra)

    Exit Sub

Erro_BotaoReqCompra_Click:

    Select Case gErr
    
        Case 89187
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156659)
    
    
    End Select
    
    Exit Sub

End Sub

Public Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Public Sub DataEntrada_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEntrada, iAlterado)

End Sub

Public Sub DataVencimento_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataVencimento, iAlterado)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub FilialCompra_Change()

    If Len(Trim(TipoNFiscal.Text)) > 0 And (sFornecedorAnterior <> Trim(Fornecedor.Text) Or iFilialAnterior <> StrParaInt(Codigo_Extrai(Filial.Text)) _
        Or iFilialCompraAnterior <> Codigo_Extrai(FilialCompra.Text)) Then

        Call Atualiza_ListaPedidos

    End If
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub FilialCompra_Click()
    
    FilialDeCompra.Caption = FilialCompra.Text

    If Len(Trim(TipoNFiscal.Text)) > 0 And (sFornecedorAnterior <> Trim(Fornecedor.Text) Or iFilialAnterior <> StrParaInt(Codigo_Extrai(Filial.Text)) _
        Or iFilialCompraAnterior <> Codigo_Extrai(FilialCompra.Text)) Then

        Call Atualiza_ListaPedidos

    End If
    
End Sub

Public Sub FilialCompra_GotFocus()
    
    iFilialCompraAnterior = Codigo_Extrai(FilialCompra.Text)

End Sub

Public Sub NaturezaOp_GotFocus()

Dim iNaturezaAux As Integer
    
    iNaturezaAux = iNaturezaOpAlterada
    Call MaskEdBox_TrataGotFocus(NaturezaOp, iAlterado)
    iNaturezaOpAlterada = iNaturezaAux

End Sub

Public Sub NFiscal_GotFocus()
    Call MaskEdBox_TrataGotFocus(NFiscal, iAlterado)
End Sub

Public Sub NFiscalOriginal_GotFocus()
    Call MaskEdBox_TrataGotFocus(NFiscalOriginal, iAlterado)
End Sub

Public Sub NumRecebimento_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumRecebimento, iAlterado)
End Sub

Private Sub Class_Terminate()
    'If giDebug = 1 Then MsgBox ("Saiu")

End Sub

Public Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    'Se o produto da linha selecionada estiver preenchido
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) <> 0 And GridItens.Row <> 0 Then

        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 65607

        'Coloca o valor do Ccl na coluna correspondente
        GridItens.TextMatrix(GridItens.Row, iGrid_Ccl_Col) = sCclMascarado
        
        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 65607 'Tratado na rotina chamadora
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156660)

    End Select

    Exit Sub

End Sub

Public Sub RecebimentoLabel_Click()

Dim colSelecao As Collection
Dim objNFiscal As New ClassNFiscal
Dim objFornecedor As New ClassFornecedor
Dim lErro As Long

On Error GoTo Erro_RecebimentoLabel_Click

    'Se o Recebimento estiver preenchido
    If Len(Trim(NumRecebimento.Text)) > 0 Then
        objNFiscal.lNumRecebimento = CLng(NumRecebimento.Text)
    Else
        objNFiscal.lNumRecebimento = 0
    End If
    
    'Verifica preenchimento de Fornecedor
    If Len(Trim(Fornecedor.Text)) <> 0 Then

        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 89309

        'Se não achou o Fornecedor --> erro
        If lErro = 6681 Then gError 89310

        objNFiscal.lFornecedor = objFornecedor.lCodigo

    End If

    If Len(Trim(Filial.Text)) <> 0 Then objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa

    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)

    If Len(Trim(NFiscal.Text)) <> 0 Then objNFiscal.lNumNotaFiscal = CLng(NFiscal.Text)

    objNFiscal.dtDataEntrada = MaskedParaDate(DataEntrada)

    'Chama a tela de browse RecebMaterialFLista
    Call Chama_Tela("RecebMaterialFComLista", colSelecao, objNFiscal, objEventoRecebimento)
    
    Exit Sub
    
Erro_RecebimentoLabel_Click:

    Select Case gErr
    
        Case 89309

        Case 89310
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, Fornecedor.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156661)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Serie_Click()

    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = REGISTRO_ALTERADO

End Sub

Public Sub Total_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dTotal As Double

On Error GoTo Erro_Total_Validate

    'Se estiver preenchido
    If Len(Trim(Total.Text)) > 0 Then
    
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(Total.Text)
        If lErro <> SUCESSO Then gError 65608

        dTotal = CDbl(Total.Text)
        
        'Coloca o valor formatado na tela
        Total.Text = Format(dTotal, "Standard")

    End If

    Exit Sub

Erro_Total_Validate:

    Cancel = True

    Select Case gErr

        Case 65608

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156662)

    End Select

    Exit Sub

End Sub

Public Sub TipoNFiscal_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TipoNFiscal_Click
    
    'Se foi selecionado um Tipo de Nota Fiscal
    If TipoNFiscal.ListIndex = -1 Then
        
        'Atualiza a lista de Pedidos de Compras
        Call Atualiza_ListaPedidos
        Exit Sub
    
    End If
    
    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Se o tipo NF não foi alterado, sai.
    If iTipoNFAlterado = objTipoDocInfo.iCodigo Then Exit Sub

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 65609

    'se não estiver cadastrado ==> erro
    If lErro = 31415 Then gError 65610
    
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then
        giComplemento = MARCADO
    Else
        giComplemento = DESMARCADO
    End If
    
    gobjRastreamento.iCodigo = objTipoDocInfo.iCodigo
    
    iEmitenteNF = objTipoDocInfo.iEmitente
    
    If iEmitenteNF <> DOCINFO_EMPRESA Then
    
        NaturezaOp.Text = NaturezaOp_Conv34(objTipoDocInfo.sNatOpExtPadrao, IIf(MaskedParaDate(DataEntrada) = DATA_NULA, MaskedParaDate(DataEmissao), MaskedParaDate(DataEntrada)))
                
    End If
    
    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
        NFiscal.Visible = False
        NFiscalInterna.Visible = True
        BotaoLimparNF.Visible = True
        'ChvNFe.Enabled = False
        'Cor Preta para label pois não é obrigatorio
        NFiscalLabel.ForeColor = vbBlack
    Else
        NFiscal.Visible = True
        NFiscalInterna.Visible = False
        BotaoLimparNF.Visible = False
        'ChvNFe.Enabled = True
        'Cor Vermelha para label pois é obrigatorio
        NFiscalLabel.ForeColor = &H80&
    End If
    
    If NFiscalInterna.Visible Then
        'Seleciona série padrão
        lErro = Seleciona_Serie_Padrao()
        If lErro <> SUCESSO Then gError 65609
    End If
        
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
    If IsDate(DataEntrada.Text) Then Call gobjTribTab.DataEntrada_Alterada(StrParaDate(DataEntrada.Text))
    Call gobjTribTab.Serie_Alterada(Serie.Text)
'*** fim tributacao
    
    'tratamento dado ao subtipo contábil: de acordo com o DOCINFO selecionado, serão carregados os modelos pertencentes ao subtipo contábil correspondente
    lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_ESTOQUE, objTipoDocInfo.iSubTipoContabil)
    If lErro <> SUCESSO Then gError 79372
    
    'Se o tipo de destino dos pedido de compra não mudou, sai da função
    If (iTipoNFAlterado <> DOCINFO_NFEEBFCOM And iTipoNFAlterado <> DOCINFO_NFIEBFCOM) And objTipoDocInfo.iTipoMovtoEstoque <> MOV_EST_MAT_NOSSO_PARA_BENEF_ENTRADA Then Exit Sub
    If (iTipoNFAlterado = DOCINFO_NFEEBFCOM Or iTipoNFAlterado = DOCINFO_NFIEBFCOM) And objTipoDocInfo.iTipoMovtoEstoque = MOV_EST_MAT_NOSSO_PARA_BENEF_ENTRADA Then Exit Sub
    
    iTipoNFAlterado = objTipoDocInfo.iCodigo
    
    Call Atualiza_ListaPedidos
    
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 79372
    End If
    
    Exit Sub

Erro_TipoNFiscal_Click:

    Select Case gErr

        Case 65609, 79372

        Case 65610
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156663)

    End Select

    Exit Sub

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Call CF2(Me, "NFEnt_Form_Unload", Cancel)
    Set gobjInfoUsu = Nothing
    
    Set gcolPedidoCompra = Nothing
    Set gcolItemPedCompraInfo = Nothing
    Set gcolItemReqCompraInfo = Nothing
    
    Set gobjTelaProjetoInfo = Nothing
    Set gobjInfoAdic = Nothing
    
    'Eventos dos Browses
    Set objEventoSerie = Nothing
    Set objEventoProduto = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoNatureza = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoRecebimento = Nothing
    Set objEventoSerieOriginal = Nothing
    Set objEventoNFiscalOriginal = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoCcl = Nothing
    Set objEventoFornecedorBenef = Nothing
    Set objEventoChvNFe = Nothing

    Set objGrid1 = Nothing
    Set objContabil = Nothing

    'Rastreamento
    Set gobjRastreamento = Nothing
    
    'distribuicao
    Set gobjDistribuicao = Nothing
    
    'Grids
    Set objGridItens = Nothing
    Set objGridItensPC = Nothing
    Set objGridItensRC = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

''*** incluidos p/tratamento de tributacao *******************************
    If Not (gobjTribTab Is Nothing) Then Call gobjTribTab.Desativar
    Set gobjTribTab = Nothing
'*** fim tributacao

    Set gobjNFiscal = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)
    
    Set gobjAnotacao = Nothing
    
    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156664)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLocalizacaoDist_Click()
'distribuicao
    Call gobjDistribuicao.BotaoLocalizacaoDist_Click
End Sub

Public Sub BotaoLotes_Click()
'Chama a tela de Lote de Rastreamento

Dim lErro As Long

On Error GoTo Erro_BotaoLotes_Click
    
    Call gobjRastreamento.BotaoLotes_Click
                    
    Exit Sub

Erro_BotaoLotes_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156665)
    
    End Select
    
    Exit Sub

End Sub

Public Sub ItemNFRastro_Change()
'Rastreamento
    Call gobjRastreamento.ItemNFRastro_Change
End Sub

Public Sub ItemNFRastro_GotFocus()
'Rastreamento
    Call gobjRastreamento.ItemNFRastro_GotFocus
End Sub

Public Sub ItemNFRastro_KeyPress(KeyAscii As Integer)
'Rastreamento
    Call gobjRastreamento.ItemNFRastro_KeyPress(KeyAscii)
End Sub

Public Sub ItemNFRastro_Validate(Cancel As Boolean)
'Rastreamento
    Call gobjRastreamento.ItemNFRastro_Validate(Cancel)
End Sub

Public Sub AlmoxRastro_Change()
'Rastreamento
    Call gobjRastreamento.AlmoxRastro_Change
End Sub

Public Sub AlmoxRastro_GotFocus()
'Rastreamento
    Call gobjRastreamento.AlmoxRastro_GotFocus
End Sub

Public Sub AlmoxRastro_KeyPress(KeyAscii As Integer)
'Rastreamento
    Call gobjRastreamento.AlmoxRastro_KeyPress(KeyAscii)
End Sub

Public Sub AlmoxRastro_Validate(Cancel As Boolean)
'Rastreamento
    Call gobjRastreamento.AlmoxRastro_Validate(Cancel)
End Sub

Public Sub LoteRastro_Change()
'Rastreamento
    Call gobjRastreamento.LoteRastro_Change
End Sub

Public Sub LoteRastro_GotFocus()
'Rastreamento
    Call gobjRastreamento.LoteRastro_GotFocus
End Sub

Public Sub LoteRastro_KeyPress(KeyAscii As Integer)
'Rastreamento
    Call gobjRastreamento.LoteRastro_KeyPress(KeyAscii)
End Sub

Public Sub LoteRastro_Validate(Cancel As Boolean)
'Rastreamento
    Call gobjRastreamento.LoteRastro_Validate(Cancel)
End Sub

Public Sub FilialOPRastro_Change()
'Rastreamento
    Call gobjRastreamento.FilialOPRastro_Change
End Sub

Public Sub FilialOPRastro_GotFocus()
'Rastreamento
    Call gobjRastreamento.FilialOPRastro_GotFocus
End Sub

Public Sub FilialOPRastro_KeyPress(KeyAscii As Integer)
'Rastreamento
    Call gobjRastreamento.FilialOPRastro_KeyPress(KeyAscii)
End Sub

Public Sub FilialOPRastro_Validate(Cancel As Boolean)
'Rastreamento
    Call gobjRastreamento.FilialOPRastro_Validate(Cancel)
End Sub

Public Sub QuantLoteRastro_Change()
'Rastreamento
    Call gobjRastreamento.QuantLoteRastro_Change
End Sub

Public Sub QuantLoteRastro_GotFocus()
'Rastreamento
    Call gobjRastreamento.QuantLoteRastro_GotFocus
End Sub

Public Sub QuantLoteRastro_KeyPress(KeyAscii As Integer)
'Rastreamento
    Call gobjRastreamento.QuantLoteRastro_KeyPress(KeyAscii)
End Sub

Public Sub QuantLoteRastro_Validate(Cancel As Boolean)
'Rastreamento
    Call gobjRastreamento.QuantLoteRastro_Validate(Cancel)
End Sub

Public Sub ItemNFDist_Change()
'distribuicao
    Call gobjDistribuicao.ItemNFDist_Change
End Sub

Public Sub ItemNFDist_GotFocus()
'distribuicao
    Call gobjDistribuicao.ItemNFDist_GotFocus
End Sub

Public Sub ItemNFDist_KeyPress(KeyAscii As Integer)
'distribuicao
    Call gobjDistribuicao.ItemNFDist_KeyPress(KeyAscii)
End Sub

Public Sub ItemNFDist_Validate(Cancel As Boolean)
'distribuicao
    Call gobjDistribuicao.ItemNFDist_Validate(Cancel)
End Sub

Public Sub AlmoxDist_Change()
'distribuicao
    Call gobjDistribuicao.AlmoxDist_Change
End Sub

Public Sub AlmoxDist_GotFocus()
'distribuicao
    Call gobjDistribuicao.AlmoxDist_GotFocus
End Sub

Public Sub AlmoxDist_KeyPress(KeyAscii As Integer)
'distribuicao
    Call gobjDistribuicao.AlmoxDist_KeyPress(KeyAscii)
End Sub

Public Sub AlmoxDist_Validate(Cancel As Boolean)
'distribuicao
    Call gobjDistribuicao.AlmoxDist_Validate(Cancel)
End Sub

Public Sub QuantDist_Change()
'distribuicao
    Call gobjDistribuicao.QuantDist_Change
End Sub

Public Sub QuantDist_GotFocus()
'distribuicao
    Call gobjDistribuicao.QuantDist_GotFocus
End Sub

Public Sub QuantDist_KeyPress(KeyAscii As Integer)
'distribuicao
    Call gobjDistribuicao.QuantDist_KeyPress(KeyAscii)
End Sub

Public Sub QuantDist_Validate(Cancel As Boolean)
'distribuicao
    Call gobjDistribuicao.QuantDist_Validate(Cancel)
End Sub

Public Sub NaturezaOp_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNaturezaOp As New ClassNaturezaOp
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NaturezaOp_Validate
   
    If iNaturezaOpAlterada = 0 Then Exit Sub
    
    'Se a natureza não estiver preenchida, sai da rotina
    If Len(Trim(NaturezaOp.ClipText)) = 0 Then Exit Sub
    
    objNaturezaOp.sCodigo = NaturezaOp.Text
    
    lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
    If lErro <> SUCESSO And lErro <> 17958 Then gError 65611
    
    'Se não encontrou a natureza de operação, erro
    If lErro = 17958 Then gError 65612
    
    'se for nota interna garantir que tenha codigo < 500
    If iEmitenteNF = DOCINFO_EMPRESA Then
    
        'Se o código da natureza de operação for maior ou igual a 500
        If objNaturezaOp.sCodigo >= NATUREZA_SAIDA_COD_INICIAL Then gError 65613
        
        If (gobjTribTab.sNatOpInterna <> objNaturezaOp.sCodigo) Then
        
            gobjTribTab.sNatOpInterna = objNaturezaOp.sCodigo
            Call NatOpInterna_Validate(Cancel)
            
        End If

    End If
    
    'se for nota externa garantir que tenha codigo > 500
    If iEmitenteNF <> DOCINFO_EMPRESA And objNaturezaOp.sCodigo <= NATUREZA_SAIDA_COD_INICIAL Then gError 65614

    iNaturezaOpAlterada = 0
    
    Exit Sub

Erro_NaturezaOp_Validate:

    Cancel = True

    Select Case gErr

        Case 65611

        Case 65612
            'Pergunta se deseja criar natureza de operação
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_NATUREZA_OPERACAO", NaturezaOp.Text)
            If vbMsgRes = vbYes Then
                
                Call Chama_Tela("NaturezaOperacao", objNaturezaOp)
            End If
        
        Case 65613
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_ENTRADA", gErr)
        
        Case 65614
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156666)

    End Select

    Exit Sub

End Sub

Public Sub NatOpInterna_Validate(Cancel As Boolean)
    Call gobjTribTab.NatOpInterna_Validate(Cancel)
End Sub

Public Sub Form_Activate()

Dim lErro As Long

On Error GoTo Erro_Form_Activate

    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError 89320
        
    Call TelaIndice_Preenche(Me)

    Exit Sub
     
Erro_Form_Activate:

    Select Case gErr
          
        Case 89320
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156667)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    gbLimpaTaxa = True

    TAB_Complemento = 5
    TAB_Distribuicao = 6
    TAB_Tributacao = 7
    
    If giTipoVersao = VERSAO_LIGHT Then
        
        BotaoCcls.Visible = False
        
        'sumir com abas de contabilidade e rastro
        TabStrip1.Tabs.Remove (TAB_RASTRO)
        TabStrip1.Tabs.Remove (TAB_Contabilizacao)
        TabStrip1.Tabs.Remove (TAB_RequisicaoCompras)
        
        TAB_Complemento = TAB_Complemento - 1
        TAB_Distribuicao = TAB_Distribuicao - 1
        TAB_Tributacao = TAB_Tributacao - 1
    
    End If
    
    'Coloca gdtDataAtual em DataEntrada
    DataEntrada.PromptInclude = False
    DataEntrada.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEntrada.PromptInclude = True
    
    lErro = Carrega_Moeda()
    If lErro <> SUCESSO Then gError 114500
    
    'Adiciona a opcao TODOS na combo
    ComboPedidoCompras.Clear
    ComboPedidoCompras.AddItem "TODOS"
    
    MsgAutomatica.Value = vbChecked
    Mensagem.MaxLength = MENSAGEM_NF_MAXLENGTH
    MensagemCorpo.MaxLength = MENSAGEM_NF_MAXLENGTH
    
    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207460
    
     Set gobjNFiscal = New ClassNFiscal
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 114500, 207460
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156668)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Function Carrega_TiposDocInfo() As Long
'Carrega na os Tipo de Documentos relacionados com a tela de Nota Fiscal de Entrada

Dim lErro As Long
Dim colTipoDocInfo As New colTipoDocInfo
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Carrega_TiposDocInfo

    Set colTipoDocInfo = gobjCRFAT.colTiposDocInfo

    'Carrega na combo só os Tipos ligados essa tela
    For Each objTipoDocInfo In colTipoDocInfo
        If objTipoDocInfo.sNomeTelaNFiscal = Me.Name Then
            TipoNFiscal.AddItem CStr(objTipoDocInfo.iCodigo) & SEPARADOR & objTipoDocInfo.sNomeReduzido
            TipoNFiscal.ItemData(TipoNFiscal.NewIndex) = objTipoDocInfo.iCodigo
            'se for o tipo padrao, seleciona-o
            If objTipoDocInfo.iPadrao = Padrao Then TipoNFiscal.ListIndex = TipoNFiscal.NewIndex
        End If
    Next

    Carrega_TiposDocInfo = SUCESSO

    Exit Function

Erro_Carrega_TiposDocInfo:

    Carrega_TiposDocInfo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156669)

    End Select

    Exit Function

End Function

Public Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie
Dim sSerieAnt As String

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 65628

    'Carrega na combo
    For Each objSerie In colSerie
        If UCase(sSerieAnt) <> UCase(Desconverte_Serie_Eletronica(objSerie.sSerie)) Then Serie.AddItem Desconverte_Serie_Eletronica(objSerie.sSerie)
        SerieNFiscalOriginal.AddItem objSerie.sSerie
        sSerieAnt = Desconverte_Serie_Eletronica(objSerie.sSerie)
    Next
    
    If gobjEST.iSerieEletronicaPadrao = MARCADO Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If

    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case 65628

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156670)

    End Select

    Exit Function

End Function

Private Function Carrega_PlacaUF() As Long
'Lê as Siglas dos Estados e alimenta a list da Combobox PlacaUF

Dim lErro As Long
Dim colSiglasUF As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_PlacaUF

    Set colSiglasUF = gcolUFs
    
    'Adiciona na Combo PlacaUF
    For iIndice = 1 To colSiglasUF.Count
        PlacaUF.AddItem colSiglasUF.Item(iIndice)
    Next

    Carrega_PlacaUF = SUCESSO

    Exit Function

Erro_Carrega_PlacaUF:

    Carrega_PlacaUF = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156671)

    End Select

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim colTabelaPrecoItem As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "NFiscalTipoDocInfo"

    'Lê os dados da Tela Notas Fiscais a Pagar
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 65629

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objNFiscal.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "TipoNFiscal", objNFiscal.iTipoNFiscal, 0, "TipoNFiscal"
    colCampoValor.Add "NaturezaOp", objNFiscal.sNaturezaOp, STRING_BUFFER_MAX_TEXTO, "NaturezaOp"
    colCampoValor.Add "Fornecedor", objNFiscal.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "FilialForn", objNFiscal.iFilialForn, 0, "FilialForn"
    colCampoValor.Add "Serie", objNFiscal.sSerie, STRING_BUFFER_MAX_TEXTO, "Serie"
    colCampoValor.Add "NumNotaFiscal", objNFiscal.lNumNotaFiscal, 0, "NumNotaFiscal"
    colCampoValor.Add "DataEmissao", objNFiscal.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "DataEntrada", objNFiscal.dtDataEntrada, 0, "DataEntrada"
'horaentrada
    colCampoValor.Add "HoraEntrada", CDbl(objNFiscal.dtHoraEntrada), 0, "HoraEntrada"
    colCampoValor.Add "DataVencimento", objNFiscal.dtDataVencimento, 0, "DataVencimento"
    colCampoValor.Add "ValorFrete", objNFiscal.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorSeguro", objNFiscal.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "ValorOutrasDespesas", objNFiscal.dValorOutrasDespesas, 0, "ValorOutrasDespesas"
    colCampoValor.Add "ValorDesconto", objNFiscal.dValorDesconto, 0, "ValorDesconto"
    colCampoValor.Add "ValorTotal", objNFiscal.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "CodTransportadora", objNFiscal.iCodTransportadora, 0, "CodTransportadora"
    colCampoValor.Add "Placa", objNFiscal.sPlaca, STRING_BUFFER_MAX_TEXTO, "Placa"
    colCampoValor.Add "PlacaUF", objNFiscal.sPlacaUF, STRING_BUFFER_MAX_TEXTO, "PlacaUF"
    colCampoValor.Add "VolumeQuant", objNFiscal.lVolumeQuant, 0, "VolumeQuant"
    colCampoValor.Add "VolumeEspecie", objNFiscal.lVolumeEspecie, 0, "VolumeEspecie" 'Alterado por Luiz Nogueira em 21/08/03
    colCampoValor.Add "VolumeMarca", objNFiscal.lVolumeMarca, 0, "VolumeMarca" 'Alterado por Luiz Nogueira em 21/08/03
    colCampoValor.Add "MensagemNota", objNFiscal.sMensagemNota, STRING_BUFFER_MAX_TEXTO, "MensagemNota"
    colCampoValor.Add "PesoLiq", objNFiscal.dPesoLiq, 0, "Pesoliq"
    colCampoValor.Add "PesoBruto", objNFiscal.dPesoBruto, 0, "PesoBruto"
    colCampoValor.Add "NumIntNotaOriginal", objNFiscal.lNumIntNotaOriginal, 0, "NumIntNotaOriginal"
    colCampoValor.Add "ValorProdutos", objNFiscal.dValorProdutos, 0, "ValorProdutos"
    colCampoValor.Add "FornecedorBenef", objNFiscal.lFornecedorBenef, 0, "FornecedorBenef"
    colCampoValor.Add "FilialFornBenef", objNFiscal.iFilialFornBenef, 0, "FilialFornBenef"
    colCampoValor.Add "FreteRespons", objNFiscal.iFreteRespons, 0, "FreteRespons"
    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_BUFFER_MAX_TEXTO, "VolumeNumero"
    colCampoValor.Add "Observacao", objNFiscal.sObservacao, STRING_NFISCAL_OBSERVACAO, "Observacao"
    colCampoValor.Add "Status", objNFiscal.iStatus, 0, "Status"
    colCampoValor.Add "FilialPedido", objNFiscal.iFilialPedido, 0, "FilialPedido"
    
    'Adiciona filtro para o sistema de setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    colSelecao.Add "NomeTelaNFiscal", OP_IGUAL, Me.Name
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 65629

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156672)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Tela_Preenche

    objNFiscal.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objNFiscal.lNumIntDoc <> 0 Then

        'Coloca os dados da NFiscal na tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 65630

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 65630

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156673)

    End Select

    Exit Sub

End Sub

Public Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    Fornecedor.Text = objFornecedor.lCodigo
    Call Fornecedor_Validate(bCancel)

    Me.Show

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    ComboPedidoCompras.ListIndex = 0
    
    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 65631

    'Limpa a Tela
    Call Limpa_Tela_NFiscalCom1

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 65631

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156674)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 65632

    'Limpa a Tela
    Call Limpa_Tela_NFiscalCom

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 65632

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156675)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objNFiscal Is Nothing) Then

        lErro = CargaPosFormLoad(True)
        If lErro <> SUCESSO Then gError 89321

        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 65633
        
        'Se não encontrou a nota fiscal, erro
        If lErro <> SUCESSO Then gError 65634

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

        'Lê o Tipo de Documento com o Código de objNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 65635
        
        'Se não encontrar --> Erro
        If lErro = 31415 Then gError 65636
                
        'Traz a nota para a tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 65637

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 65633, 65635, 65637, 89317

        Case 65634
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)
            Call Limpa_Tela_NFiscalCom
            iAlterado = 0

        Case 65636
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156676)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO

    Call Fornecedor_Preenche

End Sub

Public Sub Mensagem_Change()
    iAlterado = REGISTRO_ALTERADO
    MsgAutomatica.Value = vbUnchecked
End Sub

Public Sub MensagemCorpo_Change()
    iAlterado = REGISTRO_ALTERADO
    MsgAutomatica.Value = vbUnchecked
End Sub

Public Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NaturezaLabel_Click()

Dim objNaturezaOp As New ClassNaturezaOp
Dim colSelecao As New Collection
Dim dtDataref As Date

    'Se NaturezaOP estiver preenchida coloca no Obj
    If Len(Trim(NaturezaOp.ClipText)) > 0 Then objNaturezaOp.sCodigo = NaturezaOp.Text

    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        dtDataref = MaskedParaDate(DataEmissao)
    Else
        dtDataref = DATA_NULA
    End If
        
    If dtDataref < DATA_INICIO_CFOP4 Then
    
        If Len(Trim(DataEntrada)) > 0 Then
            dtDataref = MaskedParaDate(DataEntrada)
        End If
    
    End If
            
    'Adiciona os limites de Natureza
    Select Case iEmitenteNF
    
        Case DOCINFO_EMPRESA
            colSelecao.Add NATUREZA_ENTRADA_COD_INICIAL
            colSelecao.Add NATUREZA_ENTRADA_COD_FINAL
            
        Case Else
            colSelecao.Add NATUREZA_SAIDA_COD_INICIAL
            colSelecao.Add NATUREZA_SAIDA_COD_FINAL
            
    End Select

    'Chama a Tela de browse de NaturezaOp
    Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, objEventoNatureza, "{fn LENGTH(Codigo) } = " & IIf(dtDataref < DATA_INICIO_CFOP4, "3", "4"))

    Exit Sub

End Sub

Public Sub NaturezaOp_Change()

    iAlterado = REGISTRO_ALTERADO
    iNaturezaOpAlterada = 1
    
End Sub

Public Sub NFiscal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NFiscalLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As Collection
Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_NFiscalLabel_Click

    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 89315

    If lErro = SUCESSO Then objNFiscal.iTipoDocInfo = objTipoDocInfo.iCodigo

    'Recolhe os dados da NFiscal que estão na tela
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 65638

    'Chama a tela de Brow sw Nota Fiscal de Entrada
    Call Chama_Tela("NFiscalEntradaComLista", colSelecao, objNFiscal, objEventoNFiscal)

    Exit Sub

Erro_NFiscalLabel_Click:

    Select Case gErr

        Case 65638, 89315

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156677)

    End Select

    Exit Sub

End Sub

Public Sub NFiscalOriginal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NFiscalOriginalLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As Collection
Dim lErro As Long

On Error GoTo Erro_NFiscalOriginalLabel_Click

    'Recolhe os dados da Nota Fiscal Entrada
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 65639

    'Guarda a Serie e o Número da Nota Fiscal Original da Tela
    objNFiscal.sSerie = SerieNFiscalOriginal.Text
    
    'Se a nota fiscal original foi preenchida
    If Len(Trim(NFiscalOriginal.ClipText)) > 0 Then
        objNFiscal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
    Else
        objNFiscal.lNumNotaFiscal = 0
    End If

    'Chama a Tela NFiscalEntradaNormalLista
    Call Chama_Tela("NFiscalEntradaNormalLista", colSelecao, objNFiscal, objEventoNFiscalOriginal)

    Exit Sub

Erro_NFiscalOriginalLabel_Click:

    Select Case gErr

        Case 65639

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156678)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedorBenef_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    FornecedorBenef.Text = objFornecedor.lCodigo
    Call FornecedorBenef_Validate(bCancel)

    Me.Show

End Sub

Private Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim objNaturezaOp As ClassNaturezaOp

    Set objNaturezaOp = obj1

    'Coloca a natureza Operação na tela
    NaturezaOp.Text = objNaturezaOp.sCodigo

    Me.Show

    Call NaturezaOp_Validate(bSGECancelDummy)
    
    Exit Sub

End Sub

Private Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objNFiscal = obj1

    'Traz a NotaFiscal de entrada para a a tela
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 65640

    'Fecha o Comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case gErr

        Case 65640

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156679)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNFiscalOriginal_evSelecao(obj1 As Object)

Dim objNFiscal As ClassNFiscal

    Set objNFiscal = obj1

    'Preenche a Série e o Número da Nota Fiscal Original
    SerieNFiscalOriginal.Text = objNFiscal.sSerie
    NFiscalOriginal.Text = objNFiscal.lNumNotaFiscal

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    'Coloca a Série na Tela
    Serie.Text = Desconverte_Serie_Eletronica(objSerie.sSerie)
    If ISSerieEletronica(objSerie.sSerie) Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    Call Serie_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoSerieOriginal_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    'Coloca a Série da Nota Fiscal Original na tela
    SerieNFiscalOriginal.Text = objSerie.sSerie

    Me.Show

End Sub

Private Sub objEventoTransportadora_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora

    Set objTransportadora = obj1

    'Coloca a Transportdora na Tela
    Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    Me.Show

    Exit Sub

End Sub

Public Sub PesoBruto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PesoBruto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoBruto_Validate

    'Verifica se está preenchido
    If Len(Trim(PesoBruto.ClipText)) = 0 Then Exit Sub

    'Criticao valor
    lErro = Valor_NaoNegativo_Critica(PesoBruto.Text)
    If lErro <> SUCESSO Then gError 65643

    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True


    Select Case gErr

        Case 65643

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156680)

    End Select

    Exit Sub

End Sub

Public Sub PesoLiquido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PesoLiquido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoLiquido_Validate

    'Verifica se foi preenchido
    If Len(Trim(PesoLiquido.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(PesoLiquido.Text)
    If lErro <> SUCESSO Then gError 65644

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True

    Select Case gErr

        Case 65644

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156681)

    End Select

    Exit Sub

End Sub

Public Sub Placa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PlacaUF_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PlacaUF_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PlacaUF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PlacaUF_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(PlacaUF.Text)) = 0 Then Exit Sub

    'Verifica se existe o ítem na combo
    lErro = Combo_Item_Igual(PlacaUF)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 65645

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 65646

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True


    Select Case gErr

        Case 65645

        Case 65646
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156682)

    End Select

    Exit Sub

End Sub

Public Sub BotaoRecebimentos_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_BotaoRecebimentos_Click

    'Se o Número do Recebimento estiver preenchido
    If Len(Trim(NumRecebimento.Text)) > 0 Then
        
        objNFiscal.lNumRecebimento = NumRecebimento.Text

        'Le o Recebimento
        lErro = CF("NFiscal_Le_Recebimento", objNFiscal)
        If lErro <> SUCESSO And lErro <> 61050 Then gError 65647
        
        'Se não encontrou --> Erro
        If lErro = 61050 Then gError 65648
        
        If objNFiscal.iTipoNFiscal = DOCINFO_NRCP Or objNFiscal.iTipoNFiscal = DOCINFO_NRCC Then gError 65649
                
        'Se a nota fiscal foi baixada
        If objNFiscal.iStatus = STATUS_BAIXADO Then gError 65650
        
        'Lê o Tipo de documento
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 27623 Then gError 75417
        
        'Se o Tipo de documento não é de compras
        If objTipoDocInfo.sNomeTelaNFiscal <> RECEBIMENTO_MATERIAL_FCOM And objTipoDocInfo.sNomeTelaNFiscal <> NOTA_FISCAL_COMPRAS And objTipoDocInfo.sNomeTelaNFiscal <> NOTA_FISCAL_FATURA_COMPRAS Then gError 75418
        
    Else
        
        'Verifica se o tipo está Preenchido
        If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 65651
        
        objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
    
        'Lê o Tipo de documento
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 27623 Then gError 65652
        
        'Se o Tipo de documento não é de compras
        If objTipoDocInfo.sNomeTelaNFiscal <> RECEBIMENTO_MATERIAL_FCOM And objTipoDocInfo.sNomeTelaNFiscal <> NOTA_FISCAL_COMPRAS And objTipoDocInfo.sNomeTelaNFiscal <> NOTA_FISCAL_FATURA_COMPRAS Then gError 75419
        
        'Se não encontrar --> erro
        If lErro = 27623 Then gError 65653
        
        objNFiscal.iTipoDocInfo = objTipoDocInfo.iTipo
        
        'Se for Nota Fiscal Interna então o NumRecebimento tem que está Preenchido
        If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then gError 65654

        'Verifica se os campos essenciais estão preenchidos
        If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 65655
        If Len(Trim(Filial.Text)) = 0 Then gError 65656
        If Len(Trim(Serie.Text)) = 0 Then gError 65657
        If Len(Trim(NFiscal.ClipText)) = 0 Then gError 65658
        If Len(Trim(FilialCompra.Text)) = 0 Then gError 66099
        
        'Recolhe os dados da Tela
        lErro = Move_NF_Memoria(objNFiscal)
        If lErro <> SUCESSO Then gError 65659

        If objNFiscal.iTipoNFiscal = DOCINFO_NFIENCO Then
            objNFiscal.iTipoNFiscal = DOCINFO_NRFPCO
        ElseIf objNFiscal.iTipoNFiscal = DOCINFO_NFEENCO Then
            objNFiscal.iTipoNFiscal = DOCINFO_NRFFCO
        End If

        If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then gError 65660

        'Lê o Recebimento com estas características
        lErro = CF("NFiscal_Le_Num_Forn1", objNFiscal)
        If lErro <> SUCESSO And lErro <> 55885 Then gError 65661
    
        'Se não encontrar -->Erro
        If lErro = 55885 Then gError 65662

        If objNFiscal.iStatus = STATUS_BAIXADO Then gError 65663
    
    End If
    
    giRecebimento = True
    
    'Traz para a tela
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 65664

    giRecebimento = False
    
    'Recalcula o valor da nota fiscal
    Call ValorTotal_Calcula
    
    Exit Sub

Erro_BotaoRecebimentos_Click:

    giRecebimento = False
    
    Select Case gErr

        Case 65648
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NAO_CADASTRADO", gErr, objNFiscal.lNumRecebimento)
        
        Case 65649
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_CLIENTE", gErr, objNFiscal.lNumRecebimento)
        
        Case 65650, 65663
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_VINCULADO_NF", gErr, objNFiscal.lNumRecebimento, objNFiscal.lNumNotaFiscal)
        
        Case 65651
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
        
        Case 65655
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 65656
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_CHQPRE_NAO_PREENCHIDA", gErr)

        Case 65657
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 65658
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)

        Case 65660
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_NORMAL", gErr)

        Case 65661, 65664, 65659, 65652, 65647, 75417

        Case 65662
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_MATERIAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 65653
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO2", gErr, objTipoDocInfo.iCodigo)
                        
        Case 65654
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NFINT_NAO_PREENCHIDO", gErr)
                
        Case 66099
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCOMPRA_NAO_PREENCHIDA", gErr)
            
        Case 75418, 75419
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_DIFERENTE_COMPRAS", gErr, objTipoDocInfo.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156683)

    End Select

    Exit Sub

End Sub

Public Sub Serie_Change()

    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = REGISTRO_ALTERADO
    
End Sub

Public Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Serie_Validate

    'Se a série não foi alterada, sai da rotina
    If iSerieAlterada = 0 Then Exit Sub
    
    'Verifica se foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub
    
    'Verifica se foi selecionada
    If Serie.ListIndex = -1 Then

        'Tenta selecionar a serie
        lErro = Combo_Item_Igual(Serie)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 65665
        
        'Se a serie não está cadastrada --> verifica se é uma nota interna, se for ==> erro, pois a serie não está cadastrada
        If lErro <> SUCESSO Then
    
            'Se o tipo está preenchido
            If Len(Trim(TipoNFiscal.Text)) > 0 Then
        
                objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
                
                'Lê o Tipo de Documento
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO Then gError 65666
        
                'se a nota interna e a serie não está cadastrada ==> erro
                If objTipoDocInfo.iTipo = DOCINFO_NF_INT_ENTRADA Then gError 65667
        
            End If
    
            'Verifica se tamanho da série é maior do que o espaço no bd ==> erro
            If Len(Trim(Serie.Text)) > STRING_SERIE Then gError 65668
    
        End If

    End If
    
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.Serie_Alterada(Serie.Text)
'*** fim tributacao

    iSerieAlterada = 0
    
    Exit Sub

Erro_Serie_Validate:

    Cancel = True

    Select Case gErr

        Case 65665, 65666

        Case 65667
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, Serie.Text)

        Case 65668
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_SERIE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156684)

    End Select

    Exit Sub

End Sub

Public Sub SerieLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'recolhe a serie da tela
    objSerie.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie)

    Exit Sub

End Sub

Public Sub SerieNFiscalOriginal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub SerieNFiscalOriginal_Click()

Dim lErro As Long
Dim lNumAuto As Long

On Error GoTo Erro_SerieNFiscalOriginal_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Verifica se alguma Série foi slecionada
    If SerieNFiscalOriginal.ListIndex = -1 Then Exit Sub
    
    'Pesquisa o Próximo Número disponível para a Nota Fiscal
    lErro = CF("NFiscalEntNum_Automatico", SerieNFiscalOriginal.Text, lNumAuto)
    If lErro <> SUCESSO Then gError 65669
    
    'Coloca o Númeor na tela
    NFiscalOriginal.Text = lNumAuto
    
    Exit Sub

Erro_SerieNFiscalOriginal_Click:

    Select Case gErr
    
        Case 65669
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156685)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub SerieNFiscalOriginal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lTamanho As Long

On Error GoTo Erro_SerieNFiscalOriginal_Validate

    'Verifica se o tamanho da Serie é maior que 3
    If Len(Trim(SerieNFiscalOriginal.Text)) > STRING_SERIE Then gError 65670

    Exit Sub

Erro_SerieNFiscalOriginal_Validate:

    Cancel = True

    Select Case gErr

        Case 65670
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_SERIE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156686)

    End Select

    Exit Sub

End Sub

Public Sub SerieOriginalLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'Recolhe a Série da Nota Fiscal Original da tela
    objSerie.sSerie = SerieNFiscalOriginal.Text

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerieOriginal)

End Sub

Public Sub SubTotal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function IndTab_ObterIndFrame(ByVal iIndTab As Integer) As Integer
'obtem indice nos frames a partir do indice na tabstrip

    IndTab_ObterIndFrame = iIndTab
    
    If giTipoVersao = VERSAO_LIGHT And iIndTab >= TAB_RequisicaoCompras Then
    
        IndTab_ObterIndFrame = iIndTab + 1

    End If

End Function

Public Sub TabStrip1_Click()

Dim lErro As Long, iFrameSelecionado As Integer, iTabSelecionado As Integer

On Error GoTo Erro_TabStrip1_Click

    iTabSelecionado = TabStrip1.SelectedItem.Index
    iFrameSelecionado = IndTab_ObterIndFrame(iTabSelecionado)

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If iFrameSelecionado <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(IndTab_ObterIndFrame(iFrameAtual), TabStrip1, Me) <> SUCESSO Then Exit Sub

''*** incluidos p/tratamento de tributacao *******************************

        'se abriu o tab de tributacao
        If iTabSelecionado = TAB_Tributacao Then
            lErro = gobjTribTab.TabClick
            If lErro <> SUCESSO Then gError 65671
        End If
        
'*** fim tributacao

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(iFrameSelecionado).Visible = True
        'Torna Frame atual invisivel
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = iFrameSelecionado
        
         'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If TabStrip1.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao
        
        Select Case iTabSelecionado
        
            Case TAB_DadosPrincipais
                Parent.HelpContextID = IDH_ENTRADA_NFCOM_SIMPLES_DADOS_PRINCIPAIS
                
            Case TAB_Itens
                Parent.HelpContextID = IDH_ENTRADA_NFCOM_SIMPLES_ITENS
                        
            Case TAB_Complemento
                Call Calcula_Msg
                Parent.HelpContextID = IDH_ENTRADA_NFCOM_SIMPLES_COMPLEMENTOS
                
            Case TAB_PedidoCompras
                Parent.HelpContextID = IDH_ENTRADA_NFCOM_SIMPLES_PEDIDOCOMPRA
            
            Case TAB_RequisicaoCompras
                Parent.HelpContextID = IDH_ENTRADA_NFCOM_SIMPLES_REQUISICAOCOMPRA
            
            Case TAB_Tributacao
                Parent.HelpContextID = IDH_ENTRADA_NFCOM_SIMPLES_TRIBUTACAO
            
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_ENTRADA_NFCOM_SIMPLES_CONTABILIZACAO
                        
        End Select

    End If
    
    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 65671

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156687)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Transportadora() As Long
'Carrega a combo de Transportadoras

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_Transportadora

    'Lê o Código e o NOme de Todas as Transportadoras do BD
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then gError 65672

    'Carrega a combo de Transportadoras
    For Each objCodigoNome In colCodigoNome
        Transportadora.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Transportadora.ItemData(Transportadora.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_Transportadora = SUCESSO

    Exit Function

Erro_Carrega_Transportadora:

    Carrega_Transportadora = gErr

    Select Case gErr

        Case 65672

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156688)

    End Select

    Exit Function

End Function

Public Sub TipoNFiscal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoNFiscal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TipoNFiscal_Validate

    'Verifica se o tipo está preenchido
    If Len(Trim(TipoNFiscal.Text)) = 0 Then
        Call Atualiza_ListaPedidos
        Exit Sub
    End If
    'Verifica se foi selecionado
    If TipoNFiscal.List(TipoNFiscal.ListIndex) = TipoNFiscal.Text Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(TipoNFiscal, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 65673
    
    'Se não conseguir --> Erro
    If lErro <> SUCESSO Then gError 65674
   
    'tenta ler a natureza de operacao
    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
    
    'Lê o Tipo da NF
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 65675

    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 65676
   
    'Se a Natureza de Operação ainda não foi preenchida
    If Len(Trim(NaturezaOp.ClipText)) = 0 Then

        NaturezaOp.Text = NaturezaOp_Conv34(objTipoDocInfo.sNaturezaOperacaoPadrao, IIf(MaskedParaDate(DataEntrada) = DATA_NULA, MaskedParaDate(DataEmissao), MaskedParaDate(DataEntrada)))

    End If
   
    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
        NFiscal.Visible = False
        NFiscalInterna.Visible = True
        BotaoLimparNF.Visible = True
        'ChvNFe.Enabled = False
        'Cor Preta para label pois não é obrigatorio
        NFiscalLabel.ForeColor = vbBlack
    Else
        NFiscal.Visible = True
        NFiscalInterna.Visible = False
        BotaoLimparNF.Visible = False
        'ChvNFe.Enabled = True
        'Cor Vermelha para label pois é obrigatorio
        NFiscalLabel.ForeColor = &H80&
    End If
    
    If NFiscalInterna.Visible Then
        'Seleciona série padrão
        lErro = Seleciona_Serie_Padrao()
        If lErro <> SUCESSO Then gError 65675
    End If
    
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 65675
    End If
   
    Exit Sub

Erro_TipoNFiscal_Validate:

    Cancel = True

    Select Case gErr

        Case 65673, 65675

        Case 65674
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case 65676
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156689)

    End Select

    Exit Sub

End Sub

Public Sub Total_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Transportadora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Transportadora_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Transportadora_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTransportadora As New ClassTransportadora
Dim iCodigo As Integer

On Error GoTo Erro_Transportadora_Validate

    'Verifica se foi preenchida a ComboBox Transportadora
    If Len(Trim(Transportadora.Text)) = 0 Then Exit Sub
    
    'Verifica se está preenchida com o ítem selecionado na ComboBox Transportadora
    If Transportadora.Text = Transportadora.List(Transportadora.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Transportadora, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 65677
    
    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo
        
        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 65678
        
        ''Não encontrou Transportadora no BD
        If lErro = 19250 Then gError 65679

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    End If

    'Não existe o ítem com a String na List da Combo
    If lErro = 6731 Then gError 65680

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case gErr

        Case 65677, 65678

        Case 65679
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr, objTransportadora.iCodigo)
             
        Case 65680
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156690)

    End Select

    Exit Sub

End Sub

Public Sub TransportadoraLabel_Click()

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As Collection

    'Se a transportadora estiver preenchida recolhe.
    If Len(Trim(Transportadora.Text)) > 0 Then objTransportadora.iCodigo = Codigo_Extrai(Transportadora.Text)

    'Chama a Tela de browse de Transportadora
    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTransportadora)

    Exit Sub

End Sub


Public Sub FilialCompra_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialCompra_Validate

    'Verifica se a FilialEmpresa foi preenchida
    If Len(Trim(FilialCompra.Text)) = 0 Then
        FilialDeCompra.Caption = ""
        Call Limpa_Tela_PCRC
        Exit Sub
    End If
    
    'Verifica se é uma FilialEmpresa selecionada
    If FilialCompra.Text = FilialCompra.List(FilialCompra.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialCompra, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 66003

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        objFilialEmpresa.iCodFilial = iCodigo

        'Pesquisa se existe FilialEmpresa com o codigo extraido
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 66004

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then gError 66005

        'coloca na tela
        FilialCompra.Text = iCodigo & SEPARADOR & objFilialEmpresa.sNome
        FilialDeCompra.Caption = iCodigo & SEPARADOR & objFilialEmpresa.sNome
    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 66006

    'Se foi preenchido um tipo de nota fiscal e o Fornecedor ou a FilialFornecedor ou a Filial de Compra foram alterados
    If Len(Trim(TipoNFiscal.Text)) > 0 And (sFornecedorAnterior <> Trim(Fornecedor.Text) Or iFilialAnterior <> Codigo_Extrai(Filial.Text) Or iFilialCompraAnterior <> Codigo_Extrai(FilialCompra.Text)) Then
        
        'Atualiza a lista de Pedido de Compras
        Call Atualiza_ListaPedidos

    End If
        
    Exit Sub

Erro_FilialCompra_Validate:

    Cancel = True
    
    Select Case gErr

        Case 66004, 66003
            
        Case 66005
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialCompra.Text)

        Case 66006
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialCompra.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156691)

    End Select

    Exit Sub

End Sub

Public Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    'Se o Fornecedor foi alterado
    If iFornecedorAlterado = 1 Then

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le3(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 65685

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 65686

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
                If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
                'Seleciona filial na Combo Filial
                Call CF("Filial_Seleciona", Filial, iCodFilial)
                
            End If
            

''*** incluidos p/tratamento de tributacao *******************************
            If iCodFilial <> 0 Then Call gobjTribTab.FilialFornecedor_Alterada(objFornecedor.lCodigo, iCodFilial)
'*** fim tributacao

            If gobjInfoAdic Is Nothing Then Set gobjInfoAdic = New ClassInfoAdic
            gobjInfoAdic.sNatureza = objFornecedor.sNatureza
            
            lErro = CF2(Me, "NFCF_Fornecedor_Validate", objFornecedor.lCodigo, iCodFilial)
            If lErro <> SUCESSO Then gError 131940

        'Se o fornecedor não foi preenchido
        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

            Filial.Clear
            ComboPedidoCompras.Clear
            ComboPedidoCompras.AddItem "TODOS"
            Moeda.ListIndex = -1
            Taxa.Text = ""
            Taxa.Enabled = False
            LabelTaxa.Enabled = False

            Call CF2(Me, "NFCF_Limpa_Tela")

        End If

        'Se foi preenchido um tipo de nota fiscal e o Fornecedor ou a FilialFornecedor ou a Filial de Compra foram alterados
        If Len(Trim(TipoNFiscal.Text)) > 0 And (sFornecedorAnterior <> Trim(Fornecedor.Text) Or iFilialAnterior <> StrParaInt(Codigo_Extrai(Filial.Text)) _
            Or iFilialCompraAnterior <> Codigo_Extrai(FilialCompra.Text)) Then

            'Atualiza a lista de Pedidos de Compras
            Call Atualiza_ListaPedidos

''*** incluidos p/tratamento de tributacao *******************************
            If Len(Trim(Fornecedor.Text)) > 0 Then Call gobjTribTab.FilialFornecedor_Alterada(objFornecedor.lCodigo, iCodFilial)
'*** fim tributacao
        
        End If
        
        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 65685, 65686, 131940 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156692)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Recolhe os dados da nota fiscal
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 65687
    
    'Recolhe os dados do Grid
    If objGridItens.iLinhasExistentes > 0 Then
        lErro = Move_GridItens_Memoria(objNFiscal)
        If lErro <> SUCESSO Then gError 65688
    End If
           
    'distribuicao
    lErro = gobjDistribuicao.Move_GridDist_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 89543
           
    'mover a parte de tributacao
    lErro = gobjTribTab.PreencherNF(objNFiscal)
    If lErro <> SUCESSO Then gError 65690

    lErro = CF2(Me, "NFEnt_Move_Tela_Memoria", objNFiscal)
    If lErro <> SUCESSO Then gError 124253
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 65687, 65688, 65690, 89543, 124253

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156693)

    End Select

    Exit Function

End Function

Private Function Move_GridItens_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os Itens do Grid para a Memória

Dim iIndice As Integer
Dim lErro As Long
Dim objItemNF As ClassItemNF, objItemNFNovo As ClassItemNF
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lTamanho As Long
Dim colAlocacoes As ColAlocacoesItemNF
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objItemNFGlb As ClassItemNF

On Error GoTo Erro_Move_GridItens_Memoria

    'Para cada linha existente do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemNF = New ClassItemNF

        'Verifica se o Produto está preenchido
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 65691
        
        'Armazena produto
        objItemNF.sProduto = sProdutoFormatado
        
        'Guarda os demais campos do Grid em objItemNF
        objItemNF.sDescricaoItem = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
        objItemNF.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then objItemNF.dQuantidade = CDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) > 0 Then objItemNF.dPrecoUnitario = CDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))
        lTamanho = Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col)))
        
        If lTamanho > 0 Then objItemNF.dPercDesc = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col))

        'Se o desconto foi preenchido
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))) > 0 Then objItemNF.dValorDesconto = CDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
        
        'Se o Ccl foi preenchido
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Ccl_Col))) > 0 Then

            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", GridItens.TextMatrix(iIndice, iGrid_Ccl_Col), sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError 65694
            
            objItemNF.sCcl = sCclFormatada
            
        Else

            objItemNF.sCcl = ""

        End If

        'Adiciona na coleção de Ítens
        With objItemNF
            Set objItemNFNovo = objNFiscal.ColItensNF.Add(0, iIndice, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, DATA_NULA, .sDescricaoItem, 0, 0, 0, 0, 0, colAlocacoes, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .sCcl, STATUS_LANCADO, 0, "", 0, 0, 0, "", 0, 0)
        End With

        lErro = CF2(Me, "NFEntCom_Move_GridItem_Memoria", objNFiscal, objItemNFNovo, iIndice)
        If lErro <> SUCESSO Then gError 65694
        
        Set objItemNFNovo.objInfoAdicDocItem = gobjNFiscal.ColItensNF(iIndice).objInfoAdicDocItem
    
    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 65691, 65692, 65694

        Case 65693
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156694)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim sUnidadeMed As String
Dim iIndice As Integer
Dim objItemPedCompraInfo As ClassItemPedCompraInfo
Dim sProdutoMascarado As String
Dim iCont As Integer
Dim objPedidoCompra As New ClassPedidoCompras
Dim objItemPC As ClassItemPedCompra

On Error GoTo Erro_Rotina_Grid_Enable
    
    'Se foi chamada do Saida de Celula, sai da rotina
    If iLocalChamada = 3 Then Exit Sub
    
    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name

        'Produto
        Case Produto.Name
            
            'Se o produto foi preenchido
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Produto_Col))) > 0 Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
                
                'Limpa a combo de Produtos
                Produto.Clear
                
                'Para cada Pedido de Compras da lista
                For iIndice = 0 To PedidosCompra.ListCount - 1

                    'Se o Pedido estiver selecionado
                    If PedidosCompra.Selected(iIndice) = True Then
                        
                        'Guarda o Pedido de Compras relacionado a lista de Pedido de Compras
                        Set objPedidoCompra = gcolPedidoCompra.Item(iIndice + 1)
                    
                        'Para cada itemPC do Pedido de Compras
                        For Each objItemPC In objPedidoCompra.colItens
    
                            'Mascara o produto
                            lErro = Mascara_RetornaProdutoEnxuto(objItemPC.sProduto, sProdutoMascarado)
                            If lErro <> SUCESSO Then gError 65695

                            ProdutoPC.PromptInclude = False
                            ProdutoPC.Text = sProdutoMascarado
                            ProdutoPC.PromptInclude = True
                    
                            sProdutoMascarado = ProdutoPC.Text

                            'Para cada produto da combo de Produtos da linha do Grid atual
                            For iCont = 0 To Produto.ListCount - 1
                                If sProdutoMascarado = Produto.List(iCont) Then
                                    Exit For
                                End If
                            Next
                            
                            'Se não encontrou o Produto na combo de Produtos
                            If iCont = Produto.ListCount Then
                                'Adiciona ele
                                Produto.AddItem sProdutoMascarado
                            End If
                            
                        Next

                    End If

                Next

                'Remove da Combo os Produtos que já estão em GridItens
                For iIndice = Produto.ListCount - 1 To 0 Step -1
                    For iCont = 1 To objGridItens.iLinhasExistentes
                        If GridItens.TextMatrix(iCont, iGrid_Produto_Col) = Produto.List(iIndice) Then
                            Produto.RemoveItem (iIndice)
                        End If
                    Next
                Next

                'Verifica se ainda existe algum Pedido de Compras selecionado
                For iIndice = 0 To PedidosCompra.ListCount - 1
                    If PedidosCompra.Selected(iIndice) = True Then
                        Exit For
                    End If
                Next
            
                'Se não existir, limpa a combo de Produtos
                If PedidosCompra.ListCount = iIndice Then Produto.Clear
            
            End If
            
        'Unidade de Medida
        Case UnidadeMed.Name

            'Limpa a combo de UM
            UnidadeMed.Clear

            'Guarda a UM que está no Grid
            sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 65696

            'Se o produto não foi preenchido
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                UnidadeMed.Enabled = False
            Else
                UnidadeMed.Enabled = True

                objProduto.sCodigo = sProdutoFormatado

                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 65697

                'Não achou o Produto
                If lErro = 28030 Then gError 65698

                objClasseUM.iClasse = objProduto.iClasseUM

                'Lâ as Unidades de Medidas da Classe do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 65699

                'Carrega a combo de UM
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next

                'Seleciona na UM que está preenchida
                UnidadeMed.Text = sUM

                If Len(Trim(sUM)) > 0 Then
                    lErro = Combo_Item_Igual(UnidadeMed)
                    If lErro <> SUCESSO And lErro <> 12253 Then gError 65700
                End If

            End If

        'Nas demais
'distribuicao
'        Case ValorUnitario.Name, PercentDesc.Name, Desconto.Name, Quantidade.Name, Almoxarifado.Name, DescricaoItem.Name, Ccl.Name
        Case ValorUnitario.Name, PercentDesc.Name, Desconto.Name, Quantidade.Name, DescricaoItem.Name, Ccl.Name

            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 65701
             
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
            'Se foi configurado para não aceitar valor unitário do Produto do ItemNF diferente do ItemPC
            If objControl.Name = ValorUnitario.Name And gobjCOM.iNFDiferentePC = NFISCAL_NAO_ACEITA_DIFERENCA_PC Then
                objControl.Enabled = False
            End If
            
        Case Else
            
            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 65701
             
            lErro = CF2(Me, "NFEntCom_Rotina_Grid_Enable", iLinha, objControl, iLocalChamada, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 65701
    
    End Select

    'Rastreamento
    lErro = gobjRastreamento.Rotina_Grid_Enable(iLinha, objControl, iLocalChamada)
    If lErro <> SUCESSO Then gError 83277

    'distribuicao
    lErro = gobjDistribuicao.Rotina_Grid_Enable_Dist(iLinha, objControl, iLocalChamada)
    If lErro <> SUCESSO Then gError 89546

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 65695, 65696, 65697, 65699, 65700, 65701, 83277, 89546

        Case 65698
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156695)

    End Select

    Exit Sub

End Sub

Public Function Inicializa_Grid_ItensPC(objGridInt As AdmGrid) As Long
    Inicializa_Grid_ItensPC = CF2(Me, "NFEntCom_Inicializa_Grid_ItensPC", objGridInt)
End Function

Public Function NFEntCom_Inicializa_Grid_ItensPC(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

Dim bExibeTodos As Boolean
    
    If ComboPedidoCompras.Text = "TODOS" Then bExibeTodos = True
    
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("PedCompra")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("A Receber")
    objGridInt.colColuna.Add ("Recebido")
    objGridInt.colColuna.Add ("Preço")
    If bExibeTodos Then
        objGridInt.colColuna.Add ("Moeda")
        objGridInt.colColuna.Add ("Taxa")
    End If
    objGridInt.colColuna.Add ("Unitário R$")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (CodigoPC.Name)
    objGridInt.colCampo.Add (ItemPC.Name)
    objGridInt.colCampo.Add (ProdutoPC.Name)
    objGridInt.colCampo.Add (DescProdutoPC.Name)
    objGridInt.colCampo.Add (UMPC.Name)
    objGridInt.colCampo.Add (QuantAReceberPC.Name)
    objGridInt.colCampo.Add (QuantRecebidaPC.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    If bExibeTodos Then
        objGridInt.colCampo.Add (MoedaGrid.Name)
        objGridInt.colCampo.Add (TaxaGrid.Name)
    End If
    objGridInt.colCampo.Add (ValorRecebido.Name)

    iGrid_PedCompra_Col = 1
    iGrid_Item_Col = 2
    iGrid_Prod_Col = 3
    iGrid_DescProdutoPC_Col = 4
    iGrid_UM_Col = 5
    iGrid_AReceber_Col = 6
    iGrid_Recebido_Col = 7
    iGrid_PrecoUnitario_Col = 8
    
    If bExibeTodos Then
        
        iGrid_Moeda_Col = 9
        iGrid_Taxa_Col = 10
        iGrid_Recebido_RS_Col = 11
        
    Else
    
        iGrid_Recebido_RS_Col = 9
        
    End If

    'Grid do GridInterno
    objGridInt.objGrid = GridItensPC

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_PEDIDO_COMPRAS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridItensPC.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NFEntCom_Inicializa_Grid_ItensPC = SUCESSO

    Exit Function

End Function

Public Function Inicializa_Grid_ItensRC(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Urgente")
    objGridInt.colColuna.Add ("PedCompra")
    objGridInt.colColuna.Add ("ItemPC")
    objGridInt.colColuna.Add ("ReqCompra")
    objGridInt.colColuna.Add ("ItemRC")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("FilialEmpresa")
    objGridInt.colColuna.Add ("Recebido Pedido")
    objGridInt.colColuna.Add ("A Receber")
    objGridInt.colColuna.Add ("Recebido")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Urgente.Name)
    objGridInt.colCampo.Add (CodigoPCRC.Name)
    objGridInt.colCampo.Add (ItemPCRC.Name)
    objGridInt.colCampo.Add (CodigoReq.Name)
    objGridInt.colCampo.Add (ItemRC.Name)
    objGridInt.colCampo.Add (ProdutoRC.Name)
    objGridInt.colCampo.Add (DescProdutoRC.Name)
    objGridInt.colCampo.Add (UMRC.Name)
    objGridInt.colCampo.Add (FilialReqRC.Name)
    objGridInt.colCampo.Add (QuantRecebidaPCRC.Name)
    objGridInt.colCampo.Add (QuantAReceberRC.Name)
    objGridInt.colCampo.Add (QuantRecebidaRC.Name)

    iGrid_Urgente_Col = 1
    iGrid_PedCompraRC_Col = 2
    iGrid_ItemPC_Col = 3
    iGrid_ReqCompra_Col = 4
    iGrid_ItemRC_Col = 5
    iGrid_ProdutoRC_Col = 6
    iGrid_DescProdutoRC_Col = 7
    iGrid_UMRC_Col = 8
    iGrid_FilialReq_Col = 9
    iGrid_QuantRecebidaPC_Col = 10
    iGrid_AReceberRC_Col = 11
    iGrid_RecebidoRC_Col = 12

    'Grid do GridInterno
    objGridInt.objGrid = GridItensRC

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_REQUISICAO + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridItensRC.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ItensRC = SUCESSO

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 65704

        'GridItens
        If objGridInt.objGrid.Name = GridItens.Name Then

            lErro = Saida_Celula_GridItens(objGridInt)
            If lErro <> SUCESSO Then gError 65705

        'GridItensPC
        ElseIf objGridInt.objGrid.Name = GridItensPC.Name Then

            lErro = Saida_Celula_GridItensPC(objGridInt)
            If lErro <> SUCESSO Then gError 65706

        'GridItensRC
        ElseIf objGridInt.objGrid.Name = GridItensRC.Name Then

            lErro = Saida_Celula_GridItensRC(objGridInt)
            If lErro <> SUCESSO Then gError 65707

        'Rastreamento
        ElseIf objGridInt.objGrid.Name = GridRastro.Name Then

            lErro = gobjRastreamento.Saida_Celula()
            If lErro <> SUCESSO Then gError 83278

        'distribuicao
        ElseIf objGridInt.objGrid.Name = GridDist.Name Then

            lErro = gobjDistribuicao.Saida_Celula_Dist()
            If lErro <> SUCESSO Then gError 89547

        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 65708

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 65704, 65705, 65706, 65707, 83278, 89547

        Case 65708
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156696)

    End Select

    Exit Function

End Function

Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItens

    'Verifica qual a coluna do Grid em questão
    Select Case objGridInt.objGrid.Col

        'Produto
        Case iGrid_Produto_Col
            lErro = Saida_Celula_Produto(objGridInt)
            If lErro <> SUCESSO Then gError 65709

        'DescricaoItem
        Case iGrid_DescProduto_Col
            lErro = Saida_Celula_DescricaoItem(objGridInt)
            If lErro <> SUCESSO Then gError 65710

        'Unidade Medida
        Case iGrid_UnidadeMed_Col
            lErro = Saida_Celula_UnidadeMed(objGridInt)
            If lErro <> SUCESSO Then gError 65711

        'Quantidade
        Case iGrid_Quantidade_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 65712

        'ValorUnitario
        Case iGrid_ValorUnitario_Col
            lErro = Saida_Celula_ValorUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 65714

        'PercDesc
        Case iGrid_PercDesc_Col
            lErro = Saida_Celula_PercDesc(objGridInt)
            If lErro <> SUCESSO Then gError 65715

        'Desconto
        Case iGrid_Desconto_Col
            lErro = Saida_Celula_Desconto(objGridInt)
            If lErro <> SUCESSO Then gError 65716

        'CCL
        Case iGrid_Ccl_Col
            lErro = Saida_Celula_Ccl(objGridInt)
            If lErro <> SUCESSO Then gError 65717

        Case Else
            lErro = CF2(Me, "NFEntCom_Saida_Celula_GridItens", objGridInt)
            If lErro <> SUCESSO Then gError 65717
            
    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 65709, 65710, 65711, 65712, 65713, 65714, 65715, 65716, 65717
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156697)

    End Select

    Exit Function

End Function

Function Saida_Celula_GridItensPC(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItensPC

    'Verifica qual a coluna do Grid em questão
    Select Case objGridInt.objGrid.Col

        'Recebido
        Case iGrid_Recebido_Col
            lErro = Saida_Celula_Recebido(objGridInt)
            If lErro <> SUCESSO Then gError 65718
    
    End Select

    Saida_Celula_GridItensPC = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItensPC:

    Saida_Celula_GridItensPC = gErr
    
    Select Case gErr

        Case 65718
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156698)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Recebido(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim objItemPCInfo As New ClassItemPedCompraInfo
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sProduto As String
Dim dQuantAnterior As Double
Dim dQuantPosterior As Double
Dim dQuantPosteriorUMPC As Double
Dim dQuantReceber As Double
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantRecebidaAnterior As Double
Dim bTeste As Boolean

On Error GoTo Erro_Saida_Celula_Recebido

    Set objGridInt.objControle = QuantRecebidaPC

    'Guarda o ItemPC relacionada a linha do Grid atual
    Set objItemPCInfo = gcolItemPedCompraInfo.Item(GridItensPC.Row)
    
    'GUARDA A QUANTIDADE ANTERIORMENTE DIGITADA NO GRID
    dQuantRecebidaAnterior = StrParaDbl(GridItensPC.TextMatrix(GridItensPC.Row, iGrid_Recebido_Col))
    
    'Se quantidade estiver preenchida
    If Len(Trim(QuantRecebidaPC.ClipText)) > 0 Then
            
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(QuantRecebidaPC.Text)
        If lErro <> SUCESSO Then gError 65719

        dQuantidade = StrParaDbl(QuantRecebidaPC.Text)
        
        'Coloca o valor Formatado na tela
        QuantRecebidaPC.Text = Formata_Estoque(dQuantidade)
            
    End If
    
    'CALCULA QUANTO FALTA RECEBER LEVANDO EM CONTA A TOLERÂNCIA A MAIS
    'A tolerância a mais só deve levar em conta a quantidade pedida.
    dQuantReceber = objItemPCInfo.dQuantReceber + objItemPCInfo.dQuantPedida * objItemPCInfo.dPercentMaisReceb
    
    If objItemPCInfo.iRecebForaFaixa = ERRO_E_REJEITA_RECEBIMENTO Then
    
        'Se a quantidade recebida for maior que a quantidade a receber com tolerancia --> erro
        If StrParaDbl(QuantRecebidaPC.Text) > dQuantReceber Then gError 65720
        
    End If
    
    'Guarda a quantidade Recebita atual
    objItemPCInfo.dQuantRecebida = dQuantidade
    
    sProduto = GridItensPC.TextMatrix(GridItensPC.Row, iGrid_Prod_Col)
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iPreenchido)
    If lErro <> SUCESSO Then gError 89295
    
    objProduto.sCodigo = sProdutoFormatado

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 89296
    
    If lErro = 28030 Then gError 89297
    
    'Atualiza a informação de Quantidade Recebida do Pedido no Grid de Requisições
    For iIndice = 1 To objGridItensRC.iLinhasExistentes
        If GridItensRC.TextMatrix(iIndice, iGrid_PedCompraRC_Col) = GridItensPC.TextMatrix(GridItensPC.Row, iGrid_PedCompra_Col) And _
            GridItensRC.TextMatrix(iIndice, iGrid_ProdutoRC_Col) = GridItensPC.TextMatrix(GridItensPC.Row, iGrid_Prod_Col) And _
            GridItensRC.TextMatrix(iIndice, iGrid_ItemPC_Col) = GridItensPC.TextMatrix(GridItensPC.Row, iGrid_Item_Col) Then
            
            'Converte a UM de GridItensPC para a UM do GridItensRC
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMCompra, GridItensRC.TextMatrix(iIndice, iGrid_UMRC_Col), dFator)
            If lErro <> SUCESSO Then gError 106866
                                
            GridItensRC.TextMatrix(iIndice, iGrid_QuantRecebidaPC_Col) = StrParaDbl(QuantRecebidaPC.Text) * dFator
            
        End If
    Next
    
    'Descobre o Indice do produto alterado no grid de itens
    For iIndice1 = 1 To objGridItens.iLinhasExistentes
    
        bTeste = True
        Call CF2(Me, "NFEntCom_Produto_Igual_Customizado2", iIndice1, GridItensPC.Row, bTeste)

        If sProduto = GridItens.TextMatrix(iIndice1, iGrid_Produto_Col) And bTeste Then Exit For
    Next
    
    'Acumula a quantidade recebida do Produto no GridPC
    For Each objItemPCInfo In gcolItemPedCompraInfo
    
        bTeste = True
        Call CF2(Me, "NFEntCom_Produto_Igual_Customizado4", objItemPCInfo.objInfoUsu, GridItensPC.Row, bTeste)
        
        If sProdutoFormatado = objItemPCInfo.sProduto And bTeste Then
        
            'Converte a UM de GridItensPC para a UM do GridItens
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPCInfo.sUM, GridItens.TextMatrix(iIndice1, iGrid_UnidadeMed_Col), dFator)
            If lErro <> SUCESSO Then gError 89303
        
            dQuantPosterior = dQuantPosterior + objItemPCInfo.dQuantRecebida * dFator
            dQuantPosteriorUMPC = dQuantPosteriorUMPC + objItemPCInfo.dQuantRecebida
        
        End If
    Next
                
    'Guarda a a quantidade total recebida do produto antes da mudança
    dQuantAnterior = StrParaDbl(GridItens.TextMatrix(iIndice1, iGrid_Quantidade_Col))
    
    If dQuantPosterior = 0 Then
        GridItens.TextMatrix(iIndice1, iGrid_Quantidade_Col) = ""
    Else
        GridItens.TextMatrix(iIndice1, iGrid_Quantidade_Col) = Formata_Estoque(dQuantPosterior)
    End If
                
    lErro = CF2(Me, "NFEntCom_Calcula_PesoKg", iIndice1)
    If lErro <> SUCESSO Then gError 181292
    
    'Recalcula os Valores da Nota FIscal
    lErro = Calcula_Valores(iIndice1)
    If lErro <> SUCESSO Then gError 66636
    
    'Atualiza Tributação
    lErro = gobjTribTab.Alteracao_Item_Grid(iIndice1)
    If lErro <> SUCESSO Then gError 67633

    'Recalcula o total da nota fiscal
    Call ValorTotal_Calcula
    
    'Atualiza a Quantidade Recebida no GridItensRC
    Call Atualiza_QuantidadeRC(dQuantAnterior / dFator, dQuantPosteriorUMPC, GridItensPC.Row)
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 65721

    If dQuantidade <> dQuantRecebidaAnterior Then

        lErro = gobjDistribuicao.Preenche_GridDistribuicaoPC1(gcolItemPedCompraInfo)
        If lErro <> SUCESSO Then gError 89630

    End If
    
    Saida_Celula_Recebido = SUCESSO

    Exit Function

Erro_Saida_Celula_Recebido:

    Saida_Celula_Recebido = gErr

    Select Case gErr

        Case 65719, 65721, 66114, 66636, 67633, 89295, 89296, 89298, 89299, 89630, 181292
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 65720
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTRECEBIDA_MAIOR_QUANTRECEBER", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 89297
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156699)

    End Select

    Exit Function

End Function

Function Saida_Celula_GridItensRC(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItensRC

    'Verifica qual a coluna do Grid em questão
    Select Case objGridInt.objGrid.Col

        'QuantRecebidaRC
        Case iGrid_RecebidoRC_Col
            lErro = Saida_Celula_RecebidoRC(objGridInt)
            If lErro <> SUCESSO Then gError 65722
    
    End Select

    Saida_Celula_GridItensRC = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItensRC:

    Saida_Celula_GridItensRC = gErr
    
    Select Case gErr

        Case 65722
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156700)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_RecebidoRC(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim iIndice As Integer
Dim iItem As Integer
Dim sProduto As String
Dim lPedCompra As Long
Dim objItemRCInfo As New ClassItemReqComprasInfo
Dim dQuantAnterior As Double, dFator As Double
Dim dQuantPosterior As Double, dQuantPosteriorUMItem As Double
Dim dQuantRecebidaAnterior As Double, objItemRCInfoAux As ClassItemReqComprasInfo

On Error GoTo Erro_Saida_Celula_RecebidoRC

    Set objGridInt.objControle = QuantRecebidaRC

    'Guarda o ItemRC relacionada a linha do grid em questão
    Set objItemRCInfo = gcolItemReqCompraInfo(GridItensRC.Row)
    
    dQuantRecebidaAnterior = StrParaDbl(GridItensRC.TextMatrix(GridItensRC.Row, iGrid_RecebidoRC_Col))
    
    'Se quantidade estiver preenchida
    If Len(Trim(QuantRecebidaRC.ClipText)) > 0 Then
    
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(QuantRecebidaRC.Text)
        If lErro <> SUCESSO Then gError 65723

        dQuantidade = CDbl(QuantRecebidaRC.Text)
        objItemRCInfo.dQuantRecebida = dQuantidade
        
        'Coloca o valor Formatado na tela
        QuantRecebidaRC.Text = Formata_Estoque(dQuantidade)
                
        sProduto = GridItensRC.TextMatrix(GridItensRC.Row, iGrid_ProdutoRC_Col)
        iItem = CInt(GridItensRC.TextMatrix(GridItensRC.Row, iGrid_ItemPC_Col))
        lPedCompra = CLng(GridItensRC.TextMatrix(GridItensRC.Row, iGrid_PedCompraRC_Col))
        
        'Critica a quantidade
        If dQuantidade > StrParaDbl(GridItensRC.TextMatrix(GridItensRC.Row, iGrid_QuantRecebidaPC_Col)) Then gError 65724
        If dQuantidade > StrParaDbl(GridItensRC.TextMatrix(GridItensRC.Row, iGrid_AReceberRC_Col)) Then gError 65725
        
        'Acumula a quantidade recebida total atual do Produto
        For iIndice = 1 To objGridItensRC.iLinhasExistentes
            If GridItensRC.TextMatrix(iIndice, iGrid_ProdutoRC_Col) = GridItensRC.TextMatrix(GridItensRC.Row, iGrid_ProdutoRC_Col) And iIndice <> GridItensRC.Row Then
                
                Set objItemRCInfoAux = gcolItemReqCompraInfo(iIndice)
                dQuantPosterior = dQuantPosterior + (StrParaDbl(GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col)) / objItemRCInfoAux.dFator)
            
            End If
        Next
        dQuantPosterior = dQuantPosterior + (dQuantidade / objItemRCInfo.dFator)
        
        'Atualiza a quantidade do Grid de itens NF
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = sProduto Then
                
                'Guarda a quantidade total recebida antes da mudança
                dQuantAnterior = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
                
                lErro = CF("UM_Conversao_Trans", objItemRCInfo.iClasseUM, objItemRCInfo.sSiglaUMCompra, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), dFator)
                If lErro <> SUCESSO Then gError 106866
                
                dQuantPosteriorUMItem = dQuantPosterior * dFator
                
                If dQuantPosteriorUMItem = 0 Then
                    GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = ""
                Else
                    GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(dQuantPosteriorUMItem)
                End If
                                
                'Recalcula Valores da nota fiscal
                lErro = Calcula_Valores(iIndice)
                If lErro <> SUCESSO Then gError 67260
            
                'Atualiza Tributação
                lErro = gobjTribTab.Alteracao_Item_Grid(iIndice)
                If lErro <> SUCESSO Then gError 67261
    
                'Recalcula o valor total da nota fiscal
                Call ValorTotal_Calcula
                
                'Atualiza a quantidade recebida do GridItensPC
                Call Atualiza_QuantidadePCRC(dQuantAnterior / dFator, dQuantPosteriorUMItem / dFator, GridItensRC.Row)
                
                Exit For
                
            End If
        Next
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 65726

    If dQuantidade <> dQuantRecebidaAnterior Then

        lErro = gobjDistribuicao.Preenche_GridDistribuicaoPC1(gcolItemPedCompraInfo)
        If lErro <> SUCESSO Then gError 89631

    End If
    

    Saida_Celula_RecebidoRC = SUCESSO

    Exit Function

Erro_Saida_Celula_RecebidoRC:

    Saida_Celula_RecebidoRC = gErr

    Select Case gErr

        Case 65723, 65726, 67260, 67261, 89631
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 65724
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTRECEBIDARC_MAIOR_QUANTRECEBIDAPC", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 65725
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTRECEBIDARC_MAIOR_QUANTRECEBERRC", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156701)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl

    'Se o Ccl foi preenchido
    If Len(Trim(Ccl.ClipText)) <> 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 65727

        'Se não encontrou Ccl, erro
        If lErro = 5703 Then gError 65728

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 65729

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = gErr

    Select Case gErr

        Case 65727, 65729
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 65728
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, Ccl.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156702)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iIndice As Integer, dFator As Double

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    'Verifica se a Data esta preenchida
    If Len(Trim(Produto.Text)) > 0 Then

        lErro = CF("Produto_Critica", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 65730
        
        If lErro = 25041 Then gError 65731

        'Critica se o produto é comprável
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            'não pode ser um produto não comprável
            If objProduto.iCompras = PRODUTO_NAO_COMPRAVEL Then gError 65732
                        
        End If
        
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 65734
        
        For iIndice = 0 To Produto.ListCount - 1
            If Produto.List(iIndice) = sProduto Then
                Produto.ListIndex = iIndice
            End If
        Next
        
        'Coloca as demais características do produto na tela
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then gError 65735
        
            'Inclui Item no GridPC e GridRC
            lErro = Inclui_ItensPCRC(objProduto, True)
            If lErro <> SUCESSO Then gError 63781

            'Atualiza as Quantidades de GridItensPC
            For iIndice = 1 To objGridItensPC.iLinhasExistentes
                If GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) = Produto.Text Then
                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = Formata_Estoque(StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col)))
                End If
            Next
        
            'Atualiza as Quantidades de GridItensRC
            For iIndice = 1 To objGridItensRC.iLinhasExistentes
                If GridItensRC.TextMatrix(iIndice, iGrid_ProdutoRC_Col) = Produto.Text Then
                    'Converte a UM de GridItensPC para a UM do GridItensRC
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col), GridItensRC.TextMatrix(iIndice, iGrid_UMRC_Col), dFator)
                    If lErro <> SUCESSO Then gError 106866
                    
                    GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col) = Formata_Estoque(dFator * StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col)))
                    GridItensRC.TextMatrix(iIndice, iGrid_QuantRecebidaPC_Col) = Formata_Estoque(dFator * StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col)))
                End If
            Next
            
        End If
    
    End If
    
    sProduto = Produto.Text
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 65736

    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = sProduto
    
    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 63781, 65730, 65735, 65736
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 65731
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 65732, 65733
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 65734
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156703)

    End Select

    Exit Function

End Function

Function Inclui_ItensPCRC(objProduto As ClassProduto, Optional ByVal bIgnoraQtdReceb As Boolean = False) As Long
'Preenche linha do Grid de Pedido de Compras e o Grid de Requisição de Compras

Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim objItemRCInfo As ClassItemReqComprasInfo
Dim objPedidoCompras As New ClassPedidoCompras
Dim objRequisicaoCompras As New ClassRequisicaoCompras
Dim objItemPC As ClassItemPedCompra
Dim objItemRC As ClassItemReqCompras
Dim lErro As Long
Dim dFator As Double
Dim colItensRC As New Collection

On Error GoTo Erro_Inclui_ItensPCRC

    'Para cada pedido de compras
    For iIndice = 0 To PedidosCompra.ListCount - 1

        'Se o pedido estiver selecionado
        If PedidosCompra.Selected(iIndice) = True Then
            
            Set objPedidoCompras = gcolPedidoCompra.Item(iIndice + 1)

            iIndice2 = 0

            'Para cada item do pedido de compras
            For Each objItemPC In objPedidoCompras.colItens

                iIndice2 = iIndice2 + 1

                'Se o produto do item for igual ao do objProduto
                If objItemPC.sProduto = objProduto.sCodigo Then

                    Set objItemPCInfo = New ClassItemPedCompraInfo

                    objItemPCInfo.dPercentMaisReceb = objItemPC.dPercentMaisReceb
                    objItemPCInfo.iItem = iIndice2
                    objItemPCInfo.iRecebForaFaixa = objItemPC.iRebebForaFaixa
                    objItemPCInfo.lNumIntDoc = objItemPC.lNumIntDoc
                    objItemPCInfo.lPedCompra = objPedidoCompras.lCodigo
                    objItemPCInfo.sDescProduto = objItemPC.sDescProduto
                    objItemPCInfo.sProduto = objItemPC.sProduto

                    Set objItemPCInfo.objInfoUsu = objItemPC.objInfoUsu

                    'Lê dados do Produto
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 65737

                    'Se não encontrou o produto, erro
                    If lErro = 28030 Then gError 65738

                    'Converte a UM
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPC.sUM, objProduto.sSiglaUMCompra, dFator)
                    If lErro <> SUCESSO Then gError 65739

                    objItemPCInfo.dQuantReceber = (objItemPC.dQuantidade - objItemPC.dQuantRecebida) * dFator
                    objItemPCInfo.dQuantRecebida = (objItemPC.dQuantidade - objItemPC.dQuantRecebida) * dFator
                    objItemPCInfo.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
                    
                    'Adiciona o item em gColItemPedCompraInfo
                    Call InsereOrdenadaPC(objItemPCInfo, bIgnoraQtdReceb)

                End If

            Next

        End If

    Next

    'Limpa o GridItensPC
    Call Grid_Limpa(objGridItensPC)
    
    'Preenche o GridItensPC
    For iIndice = 1 To gcolItemPedCompraInfo.Count
        lErro = PreencheLinha_ItensPC(gcolItemPedCompraInfo(iIndice), bIgnoraQtdReceb)
        If lErro <> SUCESSO Then gError 66634
    Next
    
    'Para cada Requisição de compras
    For iIndice = 0 To RequisicoesCompra.ListCount - 1

        'Se a Requisição estiver selecionada
        If RequisicoesCompra.Selected(iIndice) = True Then

            'Lê os Itens da Requisição de Compras
            objRequisicaoCompras.lCodigo = CLng(RequisicoesCompra.List(iIndice))
            objRequisicaoCompras.iFilialEmpresa = Codigo_Extrai(FilialCompra.Text)
            lErro = CF("ItensRequisicao_Le_PedidoCompra", objRequisicaoCompras, colItensRC)
            If lErro <> SUCESSO Then gError 65741
                        
            'Para cada item da Requisição de compras
            For Each objItemRCInfo In colItensRC
                                
                'Para os Itens da Requisicao em questao
                If objRequisicaoCompras.lCodigo = objItemRCInfo.lReqCompra Then
                
                    'Se o produto do item for igual ao do objProduto
                    If objItemRCInfo.sProduto = objProduto.sCodigo Then

                        'Adiciona o item em gColItemReqCompraInfo
                        Call InsereOrdenadaRC(objItemRCInfo)
                    
                    End If
                    
                End If

            Next

        End If

    Next
    
    Call PreencheGrid_ItensRC

    Inclui_ItensPCRC = SUCESSO
    
    Exit Function

Erro_Inclui_ItensPCRC:

    Inclui_ItensPCRC = gErr

      Select Case gErr

        Case 65737, 65739, 65741, 65742, 65744, 66634
        
        Case 65738, 65743
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156704)

    End Select

    Exit Function

End Function

Function PreencheLinha_ItensPC(objItemPCInfo As ClassItemPedCompraInfo, Optional ByVal bIgnoraQtdReceb As Boolean = False) As Long
'Preenche Linha Corrente do Grid de Pedido de Compras

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoMascarado As String
Dim objInfoPC As ClassItemPedCompraInfo

On Error GoTo Erro_PreencheLinha_ItensPC

    'Se o Item possui quantidade a receber
    If (objItemPCInfo.dQuantReceber > 0) Or bIgnoraQtdReceb Then
        
        'Pedido de Compra
        GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_PedCompra_Col) = CStr(objItemPCInfo.lPedCompra)

        'Item
        GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_Item_Col) = CStr(objItemPCInfo.iItem)

        lErro = Mascara_RetornaProdutoEnxuto(objItemPCInfo.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 65746

        ProdutoPC.PromptInclude = False
        ProdutoPC.Text = sProdutoMascarado
        ProdutoPC.PromptInclude = True

        sProdutoMascarado = ProdutoPC.Text

        'Produto
        GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_Prod_Col) = sProdutoMascarado

        'Descrição do Produto
        GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_DescProdutoPC_Col) = objItemPCInfo.sDescProduto

        'Unidade de Medida
        GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_UM_Col) = objItemPCInfo.sUM

        'Quantidade a receber
        GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_AReceber_Col) = Formata_Estoque(objItemPCInfo.dQuantReceber)

        'Quantidade recebida
        GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_Recebido_Col) = Formata_Estoque(objItemPCInfo.dQuantRecebida)
        
        If ComboPedidoCompras.Text = "TODOS" Then
        
            For iIndice = 0 To Moeda.ListCount - 1
                If Moeda.ItemData(iIndice) = objItemPCInfo.iMoeda Then
                    GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_Moeda_Col) = Moeda.List(iIndice)
                End If
            Next
            
            GridItensPC.TextMatrix(objGridItensPC.iLinhasExistentes + 1, iGrid_Taxa_Col) = Format(objItemPCInfo.dTaxa, FORMATO_TAXA_CONVERSAO_MOEDA)
            
        End If
        
        lErro = CF2(Me, "NFEntCom_PreencheLinha_GridPC", objItemPCInfo.objInfoUsu, objGridItensPC.iLinhasExistentes + 1)
        If lErro <> SUCESSO Then gError 181287

        'ALTERAÇÃO DE LINHAS EXISTENTES
        objGridItensPC.iLinhasExistentes = objGridItensPC.iLinhasExistentes + 1
    
    End If

    PreencheLinha_ItensPC = SUCESSO

    Exit Function

Erro_PreencheLinha_ItensPC:

    PreencheLinha_ItensPC = gErr

    Select Case gErr

        Case 65746, 181287

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156705)

    End Select

    Exit Function

End Function

Function Preenche_Linha_ItensRC(objItemRCInfo As ClassItemReqComprasInfo) As Long
'Preenche Linha Corrente do Grid de Requisição de Compras

Dim lErro As Long
Dim sProdutoMascarado As String
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Linha_ItensRC

    'Se o Item possui quantidade a receber
    If objItemRCInfo.dQuantReceber > 0 Then
        
        'Pedido de Compra
        GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_PedCompraRC_Col) = CStr(objItemRCInfo.lPedCompra)
        
        'ItemPC
        GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_ItemPC_Col) = CStr(objItemRCInfo.iItemPC)

        lErro = Mascara_RetornaProdutoEnxuto(objItemRCInfo.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 66008

        ProdutoPC.PromptInclude = False
        ProdutoPC.Text = sProdutoMascarado
        ProdutoPC.PromptInclude = True

        sProdutoMascarado = ProdutoPC.Text

        'Produto
        GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_ProdutoRC_Col) = sProdutoMascarado

        'Descrição do Produto
        GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_DescProdutoRC_Col) = objItemRCInfo.sDescProduto

        'Unidade de Medida
        GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_UMRC_Col) = objItemRCInfo.sUM

        'Quantidade do Pedido
        For iIndice = 1 To objGridItensPC.iLinhasExistentes
        
            If GridItensPC.TextMatrix(iIndice, iGrid_PedCompra_Col) = CStr(objItemRCInfo.lPedCompra) And GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) = sProdutoMascarado _
            And GridItensPC.TextMatrix(iIndice, iGrid_Item_Col) = objItemRCInfo.iItemPC Then
                        
                GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_QuantRecebidaPC_Col) = objItemRCInfo.dFator * GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col)
                Exit For
            End If
            
        Next
        
        'FilialEmpresa
        GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_FilialReq_Col) = objItemRCInfo.iFilialEmpresa
        
        'Requisição de Compras
        GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_ReqCompra_Col) = objItemRCInfo.lReqCompra
        
        'Urgente
        GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_Urgente_Col) = objItemRCInfo.iUrgente
        
        'ItemRC
        GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_ItemRC_Col) = objItemRCInfo.iItemRC
        
        'Quantidade a receber
        GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_AReceberRC_Col) = Formata_Estoque(objItemRCInfo.dQuantReceber)

        'Quantidade recebida
        GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_RecebidoRC_Col) = Formata_Estoque(objItemRCInfo.dQuantRecebida)

        lErro = Grid_Refresh_Checkbox(objGridItensRC)
        If lErro <> SUCESSO Then gError 66009

        'ALTERAÇÃO DE LINHAS EXISTENTES
        objGridItensRC.iLinhasExistentes = objGridItensRC.iLinhasExistentes + 1

    End If

    Preenche_Linha_ItensRC = SUCESSO

    Exit Function

Erro_Preenche_Linha_ItensRC:

    Preenche_Linha_ItensRC = gErr

    Select Case gErr

        Case 66008, 66009

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156706)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescricaoItem(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoItem

    Set objGridInt.objControle = DescricaoItem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 65747
    
    Saida_Celula_DescricaoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoItem:

    Saida_Celula_DescricaoItem = gErr

    Select Case gErr

        Case 65747
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156707)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
 
Dim lErro As Long
Dim iLinha As Integer
Dim objItemPCInfo As New ClassItemPedCompraInfo
Dim dQuantAnterior As Double
Dim dQuantPosterior As Double
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantTotal As Double
'distribuicao
Dim dQuantidadeAnterior As Double
Dim dQuantidadeAtual As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    'RECOLHE A QUANTIDADE Q ESTAVA INFORMADA ANTERIORMENTE DO GRID
    'distribuicao
    dQuantidadeAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    'fim  distribuicao

    'Se quantidade estiver preenchida
    If Len(Trim(Quantidade.ClipText)) > 0 Then
    
        'Critica o valor
        If giComplemento = DESMARCADO Then
            lErro = Valor_Positivo_Critica(Quantidade.Text)
        Else
            lErro = Valor_NaoNegativo_Critica(Quantidade.Text)
        End If
        If lErro <> SUCESSO Then gError 65748
    
        dQuantPosterior = StrParaDbl(Quantidade.Text)
        dQuantAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    
        'Coloca o valor Formatado na tela
        Quantidade.Text = Formata_Estoque(dQuantPosterior)
            
        'distribuicao
        dQuantidadeAtual = StrParaDbl(Quantidade.Text)
        'fim  distribuicao
            
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 89305

        objProduto.sCodigo = sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 89306
        
        If lErro = 28030 Then gError 89307

        dQuantTotal = 0
        
        'PARA CADA ITEM DE PEDIDO DE COMPRAS
        For Each objItemPCInfo In gcolItemPedCompraInfo
                    
            'SE FOR UM ITEM DESSE PRODUTO
            If objItemPCInfo.sProduto = objProduto.sCodigo Then
            
                'Converte a UM de GridItensPC para a UM do GridItens
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPCInfo.sUM, GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col), dFator)
                If lErro <> SUCESSO Then gError 89308
            
                'ACUMULA A QUANTIDADE A RECEBER MAIS A TOLERÂNCIA PERMITIDA
                dQuantTotal = dQuantTotal + (objItemPCInfo.dQuantReceber + (objItemPCInfo.dQuantPedida * objItemPCInfo.dPercentMaisReceb)) * dFator
            
            End If
        Next
            
        'SE A QUANTIDADE ULTRAPASSAR O TOTAL COM O % A MAIS PERMITIDO --> ERRO
        If (StrParaDbl(Quantidade.Text) - dQuantTotal) > QTDE_ESTOQUE_DELTA Then gError 65749
               
        'Atualiza a quantidade Recebida do GriditensPC
        Call Atualiza_QuantidadePC(dQuantAnterior, dQuantPosterior, GridItens.Row)

    End If
            
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 65750

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 65751
    
    '*** incluidos p/tratamento de tributacao ****************
    If GridItens.Row <= objGridItens.iLinhasExistentes Then
       lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
       If lErro <> SUCESSO Then gError 65752
    End If
    '*** fim tributacao
    
    'Recalcula o total da nota fiscal
    Call ValorTotal_Calcula
    
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 181262
    End If

    If Abs(dQuantidadeAnterior - dQuantidadeAtual) > QTDE_ESTOQUE_DELTA Then
        lErro = CF2(Me, "NFEntCom_Calcula_PesoKg", GridItens.Row)
        If lErro <> SUCESSO Then gError 181262
    End If
    
    'inicio distribuicao
    If dQuantidadeAnterior <> dQuantidadeAtual Then
        
        lErro = gobjDistribuicao.Preenche_GridDistribuicaoPC1(gcolItemPedCompraInfo)
        If lErro <> SUCESSO Then gError 89629
        
    End If
    'fim distribuicao
    
    If GridItens.Row <= objGridItens.iLinhasExistentes Then
        lErro = gobjRastreamento.Atualiza_Troca_Quantidade(GridItens.Row)
        If lErro <> SUCESSO Then gError 83279
    End If
    
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 65748, 65750, 65751, 65752, 83279, 89305, 89306, 89308, 89548, 89629, 181262
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 65749
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_MAIOR_TOTALRECEBER", gErr, dQuantPosterior, dQuantTotal)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 89307
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156708)

    End Select

    Exit Function

End Function

Sub Atualiza_QuantidadePC(dQuantAnterior As Double, dQuantPosterior As Double, iLinhaGridNF As Integer)
'Atualiza a quantidade Recebida do GridItensPC a partir da Quantidade que foi modificada do GridItens NF

'ATENCAO: as qtdes estao na UM do item da NF

Dim sProduto As String, objProduto As New ClassProduto
Dim dQuantDiferenca As Double, dQuantDiferencaUMPC As Double
Dim iIndice As Integer, dQuantAnteriorRecUMPC As Double
Dim iIndiceRC As Integer, dQuantDifPC As Double
Dim iItemPC As Integer, lErro As Long
Dim lPedCompra As Long, dFator As Double, dFatorItemNFPC As Double
Dim objItemPC As ClassItemPedCompra, objItemRCInfoAux As ClassItemReqComprasInfo
Dim objItemPCInfo As ClassItemPedCompraInfo, sUMItemNF As String
Dim bTeste As Boolean

    'Guarda o código do Produto do GridItens
    sProduto = GridItens.TextMatrix(iLinhaGridNF, iGrid_Produto_Col)
    sUMItemNF = GridItens.TextMatrix(iLinhaGridNF, iGrid_UnidadeMed_Col)
    
    'Guarda QuantDiferenca = Quantidade recebida total que estava no Grid - quantidade recebida total atual
    dQuantDiferenca = dQuantPosterior - dQuantAnterior
    
    'Se a quantidade aumentou
    If dQuantDiferenca > 0 Then
    
        'Procura por mesmo Produto no Grid de Pedidos de Compras
        For iIndice = 1 To objGridItensPC.iLinhasExistentes
            
            bTeste = True
            Call CF2(Me, "NFEntCom_Produto_Igual_Customizado2", iLinhaGridNF, iIndice, bTeste)
            
            'Se encontrou
            If GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) = sProduto And bTeste Then
                
                dQuantAnteriorRecUMPC = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                
                Set objItemPCInfo = gcolItemPedCompraInfo.Item(iIndice)
                        
                objProduto.sCodigo = objItemPCInfo.sProduto
        
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 89452
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUMItemNF, objItemPCInfo.sUM, dFatorItemNFPC)
                If lErro <> SUCESSO Then gError 106866
                
                dQuantDiferencaUMPC = dQuantDiferenca * dFatorItemNFPC
                
                'Guarda o número do item e o código do Pedido de Compras
                iItemPC = CInt(GridItensPC.TextMatrix(iIndice, iGrid_Item_Col))
                lPedCompra = CLng(GridItensPC.TextMatrix(iIndice, iGrid_PedCompra_Col))
                
                'Se a quantide a Receber da linha atual é maior que a quantidade total Recebida mais a quantidade que aumentou
                If StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_AReceber_Col)) >= dQuantAnteriorRecUMPC + dQuantDiferencaUMPC Then
                    
                    'Atualiza a Quantidade Recebida no GridItensPC
                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = Formata_Estoque(StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col)) + dQuantDiferencaUMPC)
                                        
                    objItemPCInfo.dQuantRecebida = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                                        
                    'Atualiza a quantidade Recebida do GridItensRC que está vinculada ao ItemPC da linha do GridPC em questão
                    Call Atualiza_QuantidadeRC(dQuantAnteriorRecUMPC, objItemPCInfo.dQuantRecebida, iIndice)
                                        
                    'Atualiza a Quantidade Recebida do Pedido no GridItensRC
                    For iIndiceRC = 1 To objGridItensRC.iLinhasExistentes
                        If GridItensRC.TextMatrix(iIndiceRC, iGrid_ProdutoRC_Col) = sProduto And GridItensRC.TextMatrix(iIndiceRC, iGrid_PedCompraRC_Col) = CStr(lPedCompra) And GridItensRC.TextMatrix(iIndiceRC, iGrid_ItemRC_Col) = CStr(iItemPC) Then
                        
                            Set objItemRCInfoAux = gcolItemReqCompraInfo(iIndiceRC)
                            
                            GridItensRC.TextMatrix(iIndiceRC, iGrid_QuantRecebidaPC_Col) = GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) * objItemRCInfoAux.dFator
                            
                        End If
                    Next
                                
                    Exit For 'já absorveu toda a diferenca
                
                'Se a quantidade a receber não é maior
                Else
                    
                    'Diminui de QuantDiferenca a quantidade que vai entrar no GridPC como quantidade Recebida
                    dQuantDifPC = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_AReceber_Col)) - StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                    dQuantDiferencaUMPC = dQuantDiferencaUMPC - dQuantDifPC
                    dQuantDiferenca = dQuantDiferenca - (dQuantDifPC / dFatorItemNFPC)
                    
                    'Atualiza a quantidade Recebida da linha do GridPC em questão
'                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = GridItensPC.TextMatrix(iIndice, iGrid_AReceber_Col)
                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = Format(StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_AReceber_Col)) + dQuantDiferenca, "STANDARD")
                    objItemPCInfo.dQuantRecebida = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                            
                    'Atualiza a quantidade Recebida do GridItensRC
                    Call Atualiza_QuantidadeRC(dQuantAnteriorRecUMPC, objItemPCInfo.dQuantRecebida, iIndice)
                            
                    'Atualiza a Quantidade Recebida do Pedido no GridItensRC
                    For iIndiceRC = 1 To objGridItensRC.iLinhasExistentes
                        If GridItensRC.TextMatrix(iIndiceRC, iGrid_ProdutoRC_Col) = sProduto And GridItensRC.TextMatrix(iIndiceRC, iGrid_PedCompraRC_Col) = CStr(lPedCompra) And _
                           GridItensRC.TextMatrix(iIndiceRC, iGrid_ItemRC_Col) = CStr(iItemPC) Then
                    
                            Set objItemRCInfoAux = gcolItemReqCompraInfo(iIndiceRC)
                            
                            GridItensRC.TextMatrix(iIndiceRC, iGrid_QuantRecebidaPC_Col) = GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) * objItemRCInfoAux.dFator
                                                
                        End If
                        
                    Next
                End If
            End If
        Next
    
    'Se a quantidade diminuiu
    ElseIf dQuantDiferenca < 0 Then
    
        'Atualiza a quantidade no GridItensPC procurando pelo mesmo Produto do final do Grid para o inicio
        For iIndice = objGridItensPC.iLinhasExistentes To 1 Step -1
        
            bTeste = True
            Call CF2(Me, "NFEntCom_Produto_Igual_Customizado2", iLinhaGridNF, iIndice, bTeste)
       
            If GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) = sProduto And bTeste Then
                
                dQuantAnteriorRecUMPC = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                
                Set objItemPCInfo = gcolItemPedCompraInfo.Item(iIndice)
                        
                objProduto.sCodigo = objItemPCInfo.sProduto
        
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 89452
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUMItemNF, objItemPCInfo.sUM, dFatorItemNFPC)
                If lErro <> SUCESSO Then gError 106866
                
                dQuantDiferencaUMPC = dQuantDiferenca * dFatorItemNFPC
                
                'Guarda o número do item e o código do Pedido de Compras
                lPedCompra = CLng(GridItensPC.TextMatrix(iIndice, iGrid_PedCompra_Col))
                iItemPC = CInt(GridItensPC.TextMatrix(iIndice, iGrid_Item_Col))

                If dQuantAnteriorRecUMPC + dQuantDiferencaUMPC >= 0 Then
                    
                    'Atualiza a quantidade do GridPC
                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = Formata_Estoque(dQuantAnteriorRecUMPC + dQuantDiferencaUMPC)
                                    
                    objItemPCInfo.dQuantRecebida = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                                    
                    'Atualiza a quantidade Recebida do GridItensRC que está vinculada ao ItemPC da linha do GridPC em questão
                    Call Atualiza_QuantidadeRC(dQuantAnteriorRecUMPC, objItemPCInfo.dQuantRecebida, iIndice)
                                    
                    'Atualiza a Quantidade Recebida do Pedido no GridItensRC
                    For iIndiceRC = 1 To objGridItensRC.iLinhasExistentes
                        If GridItensRC.TextMatrix(iIndiceRC, iGrid_ProdutoRC_Col) = sProduto And GridItensRC.TextMatrix(iIndiceRC, iGrid_PedCompraRC_Col) = CStr(lPedCompra) And GridItensRC.TextMatrix(iIndiceRC, iGrid_ItemRC_Col) = CStr(iItemPC) Then
                            
                            Set objItemRCInfoAux = gcolItemReqCompraInfo(iIndiceRC)
                            
                            GridItensRC.TextMatrix(iIndiceRC, iGrid_QuantRecebidaPC_Col) = GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) * objItemRCInfoAux.dFator
                        End If
                    Next
                    
                    Exit For 'já absorveu toda a diferenca
                
                Else
                
                    dQuantDifPC = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                    dQuantDiferencaUMPC = dQuantDiferencaUMPC + dQuantDifPC
                    dQuantDiferenca = dQuantDiferenca + (dQuantDifPC / dFatorItemNFPC)
                    
                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = Formata_Estoque(0)
                
                    objItemPCInfo.dQuantRecebida = 0
                
                    'Atualiza a quantidade Recebida do GridItensRC que está vinculada ao ItemPC da linha do GridPC em questão
                    Call Atualiza_QuantidadeRC(dQuantAnteriorRecUMPC, objItemPCInfo.dQuantRecebida, iIndice)
                
                    'Atualiza a Quantidade Recebida do Pedido no GridItensRC
                    For iIndiceRC = 1 To objGridItensRC.iLinhasExistentes
                        If GridItensRC.TextMatrix(iIndiceRC, iGrid_ProdutoRC_Col) = sProduto And GridItensRC.TextMatrix(iIndiceRC, iGrid_PedCompraRC_Col) = CStr(lPedCompra) And GridItensRC.TextMatrix(iIndiceRC, iGrid_ItemRC_Col) = CStr(iItemPC) Then
                    
                            Set objItemRCInfoAux = gcolItemReqCompraInfo(iIndiceRC)
                            
                            GridItensRC.TextMatrix(iIndiceRC, iGrid_QuantRecebidaPC_Col) = GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) * objItemRCInfoAux.dFator
                        End If
                    Next
                
                End If
            
            End If
        Next
        
    End If

End Sub

Sub Atualiza_QuantidadeRC(dQuantAnterior As Double, dQuantPosterior As Double, iLinhaGridPC As Integer)
'Atualiza a quantidade recebida do Item RC a partir da linha do GridItensPC passada
'ATENCAO: as qdtes vem na UM do item do pedido de compras

Dim sProduto As String
Dim dQuantDiferenca As Double
Dim iIndice As Integer, objItemRCInfoAux As ClassItemReqComprasInfo
Dim sPedCompra As String, dQuantDiferencaUMRC As Double
    
    'Guarda o código do Produto e o código do Pedido do GridItensPC
    sProduto = GridItensPC.TextMatrix(iLinhaGridPC, iGrid_Prod_Col)
    sPedCompra = GridItensPC.TextMatrix(iLinhaGridPC, iGrid_PedCompra_Col)
    
    'Guarda QuantDiferenca = Quantidade recebida que estava no Grid - quantidade recebida atual
    dQuantDiferenca = dQuantPosterior - dQuantAnterior

    'Atualiza as quantidades de ItensRequisiçãoCompras amarrando pelo Produto e Pedido de Compras
    If dQuantDiferenca > 0 Then
    
        'Se a quantidade tiver aumentado, procura primeira linha do GridRC com o Produto em questão e atualiza a quantidade
        For iIndice = 1 To objGridItensRC.iLinhasExistentes
        
            If GridItensRC.TextMatrix(iIndice, iGrid_ProdutoRC_Col) = sProduto And sPedCompra = GridItensRC.TextMatrix(iIndice, iGrid_PedCompraRC_Col) Then
                
                Set objItemRCInfoAux = gcolItemReqCompraInfo(iIndice)
                dQuantDiferencaUMRC = dQuantDiferenca * objItemRCInfoAux.dFator
                
                If StrParaDbl(GridItensRC.TextMatrix(iIndice, iGrid_AReceberRC_Col)) >= StrParaDbl(GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col)) + dQuantDiferencaUMRC Then
                    GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col) = Formata_Estoque(StrParaDbl(StrParaDbl(GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col)) + dQuantDiferencaUMRC))
                    Exit For
                Else
                    dQuantDiferenca = dQuantDiferenca - ((StrParaDbl(GridItensRC.TextMatrix(iIndice, iGrid_AReceberRC_Col)) - StrParaDbl(GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col))) / objItemRCInfoAux.dFator)
                    GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col) = GridItensRC.TextMatrix(iIndice, iGrid_AReceberRC_Col)
                End If
            End If
            
        Next
        
        'Se mesmo redistribuindo as requisições ao máximo requisitado sobrou, é porque é um recebimento a maior
        'Vou colocar tudo na primeira requisição
        If dQuantDiferenca > 0 Then
            For iIndice = 1 To objGridItensRC.iLinhasExistentes
                If GridItensRC.TextMatrix(iIndice, iGrid_ProdutoRC_Col) = sProduto And sPedCompra = GridItensRC.TextMatrix(iIndice, iGrid_PedCompraRC_Col) Then
                    Set objItemRCInfoAux = gcolItemReqCompraInfo(iIndice)
                    dQuantDiferencaUMRC = dQuantDiferenca * objItemRCInfoAux.dFator
                    GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col) = Formata_Estoque(StrParaDbl(StrParaDbl(GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col)) + dQuantDiferencaUMRC))
                    dQuantDiferenca = 0
                    Exit For
                End If
            Next
        End If
    
    'Se a quantidade for menor que zero, procura da última linha até a primeira o ItemRC com o mesmo produto e atualiza a quantidade
    ElseIf dQuantDiferenca < 0 Then
    
        For iIndice = objGridItensRC.iLinhasExistentes To 1 Step -1
            If GridItensRC.TextMatrix(iIndice, iGrid_ProdutoRC_Col) = sProduto And sPedCompra = GridItensRC.TextMatrix(iIndice, iGrid_PedCompraRC_Col) Then
                
                Set objItemRCInfoAux = gcolItemReqCompraInfo(iIndice)
                dQuantDiferencaUMRC = dQuantDiferenca * objItemRCInfoAux.dFator
                
                If StrParaDbl(GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col)) + dQuantDiferencaUMRC >= 0 Then
                    GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col) = Formata_Estoque(StrParaDbl(StrParaDbl(GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col))) + dQuantDiferencaUMRC)
                    Exit For
                Else
                    dQuantDiferenca = dQuantDiferenca + ((StrParaDbl(GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col))) / objItemRCInfoAux.dFator)
                    GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col) = Formata_Estoque(0)
                End If
            End If
        Next
        
    End If

End Sub

Sub Atualiza_QuantidadePCRC(dQuantAnterior As Double, dQuantPosterior As Double, iLinhaGridRC As Integer)
'Atualiza a quantidade do GridItensPC a partir da Quantidade Recebida que foi modificada do GridItensRC
'ATENCAO: as duas qtdes estao na UM do grid de itens

Dim sProduto As String
Dim dQuantDiferenca As Double
Dim iIndice As Integer
Dim iIndiceRC As Integer
Dim iItemPC As Integer
Dim lPedCompra As Long, objItemRCInfoAux As ClassItemReqComprasInfo
Dim objItemPC As ClassItemPedCompra
Dim objItemPCInfo As ClassItemPedCompraInfo

    'Guarda o código do Produto e código do Pedido de Compras do GridItensRC
    sProduto = GridItensRC.TextMatrix(iLinhaGridRC, iGrid_ProdutoRC_Col)
    lPedCompra = GridItensRC.TextMatrix(iLinhaGridRC, iGrid_PedCompraRC_Col)
    
    'Guarda QuantDiferenca = Quantidade recebida total que estava no Grid - quantidade recebida total atual
    dQuantDiferenca = dQuantPosterior - dQuantAnterior
    
    'Atualiza as quantidades de ItensPedidoCompra
    If dQuantDiferenca > 0 Then
    
        For iIndice = 1 To objGridItensPC.iLinhasExistentes
            If GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) = sProduto And GridItensPC.TextMatrix(iIndice, iGrid_PedCompra_Col) = CStr(lPedCompra) Then
                
                'Guarda o número do item
                iItemPC = CInt(GridItensPC.TextMatrix(iIndice, iGrid_Item_Col))
                
                'Atualiza a Quantidade Recebida no GridItensPC
                If StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_AReceber_Col)) >= StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col)) + dQuantDiferenca Then
                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = Formata_Estoque(StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col)) + dQuantDiferenca)
                                        
                    Set objItemPCInfo = gcolItemPedCompraInfo.Item(iIndice)
                    objItemPCInfo.dQuantRecebida = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                                        
                    'Atualiza a Quantidade Recebida do Pedido no GridItensRC
                    For iIndiceRC = 1 To objGridItensRC.iLinhasExistentes
                        If GridItensRC.TextMatrix(iIndiceRC, iGrid_ProdutoRC_Col) = sProduto And GridItensRC.TextMatrix(iIndiceRC, iGrid_PedCompraRC_Col) = CStr(lPedCompra) And GridItensRC.TextMatrix(iIndiceRC, iGrid_ItemRC_Col) = CStr(iItemPC) Then
                    
                            Set objItemRCInfoAux = gcolItemReqCompraInfo(iIndiceRC)

                            GridItensRC.TextMatrix(iIndiceRC, iGrid_QuantRecebidaPC_Col) = GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) * objItemRCInfoAux.dFator
                        End If
                    Next
                
                    Exit For
                
                Else
                    dQuantDiferenca = dQuantDiferenca - (StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_AReceber_Col)) - StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col)))
                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = GridItensPC.TextMatrix(iIndice, iGrid_AReceber_Col)
                            
                    Set objItemPCInfo = gcolItemPedCompraInfo.Item(iIndice)
                    objItemPCInfo.dQuantRecebida = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                            
                    'Atualiza a Quantidade Recebida do Pedido no GridItensRC
                    For iIndiceRC = 1 To objGridItensRC.iLinhasExistentes
                        If GridItensRC.TextMatrix(iIndiceRC, iGrid_ProdutoRC_Col) = sProduto And GridItensRC.TextMatrix(iIndiceRC, iGrid_PedCompraRC_Col) = CStr(lPedCompra) And _
                        GridItensRC.TextMatrix(iIndiceRC, iGrid_ItemRC_Col) = CStr(iItemPC) Then
                    
                            Set objItemRCInfoAux = gcolItemReqCompraInfo(iIndiceRC)
                            
                            GridItensRC.TextMatrix(iIndiceRC, iGrid_QuantRecebidaPC_Col) = GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) * objItemRCInfoAux.dFator
                        End If
                        
                    Next
                End If
            End If
        Next
    
    ElseIf dQuantDiferenca < 0 Then
    
        For iIndice = objGridItensPC.iLinhasExistentes To 1 Step -1
            If GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) = sProduto And GridItensPC.TextMatrix(iIndice, iGrid_PedCompra_Col) = CStr(lPedCompra) Then
                
                'Guarda o número do item
                iItemPC = CInt(GridItensPC.TextMatrix(iIndice, iGrid_Item_Col))

                If StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col)) + dQuantDiferenca >= 0 Then
                    
                    
                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = Formata_Estoque(StrParaDbl(StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col)) + dQuantDiferenca))
                                    
                    Set objItemPCInfo = gcolItemPedCompraInfo.Item(iIndice)
                    objItemPCInfo.dQuantRecebida = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                                    
                    'Atualiza a Quantidade Recebida do Pedido no GridItensRC
                    For iIndiceRC = 1 To objGridItensRC.iLinhasExistentes
                        If GridItensRC.TextMatrix(iIndiceRC, iGrid_ProdutoRC_Col) = sProduto And GridItensRC.TextMatrix(iIndiceRC, iGrid_PedCompraRC_Col) = CStr(lPedCompra) And GridItensRC.TextMatrix(iIndiceRC, iGrid_ItemRC_Col) = CStr(iItemPC) Then
                    
                            Set objItemRCInfoAux = gcolItemReqCompraInfo(iIndiceRC)
                            
                            GridItensRC.TextMatrix(iIndiceRC, iGrid_QuantRecebidaPC_Col) = GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) * objItemRCInfoAux.dFator
                        End If
                    Next
                    
                    Exit For
                
                Else
                    dQuantDiferenca = dQuantDiferenca + StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                    GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) = Formata_Estoque(0)
                
                    Set objItemPCInfo = gcolItemPedCompraInfo.Item(iIndice)
                    objItemPCInfo.dQuantRecebida = 0
                
                    'Atualiza a Quantidade Recebida do Pedido no GridItensRC
                    For iIndiceRC = 1 To objGridItensRC.iLinhasExistentes
                        If GridItensRC.TextMatrix(iIndiceRC, iGrid_ProdutoRC_Col) = sProduto And GridItensRC.TextMatrix(iIndiceRC, iGrid_PedCompraRC_Col) = CStr(lPedCompra) And GridItensRC.TextMatrix(iIndiceRC, iGrid_ItemRC_Col) = CStr(iItemPC) Then
                    
                            Set objItemRCInfoAux = gcolItemReqCompraInfo(iIndiceRC)
                            
                            GridItensRC.TextMatrix(iIndiceRC, iGrid_QuantRecebidaPC_Col) = GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col) * objItemRCInfoAux.dFator
                        End If
                    Next
                
                End If
            
            End If
        Next
        
    End If

End Sub

Public Sub ValorReal_Calcula(dQuantidade As Double, dValorUnitario As Double, dPercentDesc As Double, dDesconto As Double, dValorReal As Double)
'Calcula o Valor Real

Dim dValorTotal As Double
Dim dPercDesc1 As Double
Dim dPercDesc2 As Double

    dValorTotal = dValorUnitario * dQuantidade

     'Se o Percentual Desconto estiver preenchido e ValorTotal positivo
    If dPercentDesc > 0 And dValorTotal > 0 Then
        
        'Testa se o desconto está preenchido
        If dDesconto = 0 Then
            dPercDesc2 = 0
        Else
            'Calcula o Percentual em cima dos valores passados
            dPercDesc2 = (dDesconto / dValorTotal)
            dPercDesc2 = CDbl(Format(dPercDesc2, "0.0000"))
        End If
        'se os percentuais (passado e calulado) forem diferentes calcula-se o desconto
        If dPercentDesc <> dPercDesc2 Then dDesconto = (dPercentDesc * dValorTotal)

    End If

    dValorReal = dValorTotal - dDesconto

End Sub

Private Function SubTotal_Calcula() As Long
'Soma a coluna de Valor Total e acumula em SubTotal

Dim lErro As Long
Dim dSubTotal As Double
Dim dValorProdutos As Double, dValorDescontoItens As Double
Dim iIndice As Integer
Dim dValorDesconto As Double
Dim objItemNF As ClassItemNF

On Error GoTo Erro_SubTotal_Calcula

    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemNF = gobjNFiscal.colItens.Item(iIndice)
        
        If objItemNF.objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
            
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) <> 0 Then
            
                dSubTotal = dSubTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
    
            End If
            dValorProdutos = dValorProdutos + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
            dValorDescontoItens = dValorDescontoItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
        End If
    Next

'    If Len(Trim(ValorDesconto.ClipText)) > 0 And IsNumeric(ValorDesconto) Then dValorDesconto = CDbl(ValorDesconto.Text)
'
'    dSubTotal = Arredonda_Moeda(dSubTotal - dValorDesconto, 2)
    
    SubTotal.Caption = Format(CStr(dSubTotal), "Standard")
    ValorProdutos2.Caption = Format(CStr(dValorProdutos), "Standard")

    dValorDescontoItensAnt = dValorDescontoItens
    If dValorProdutos > 0 Then
        PercDescontoItens.Text = Format(dValorDescontoItens / dValorProdutos * 100, "FIXED")
        dPercDescontoItensAnt = Arredonda_Moeda(dValorDescontoItens / dValorProdutos, 4)
    Else
        PercDescontoItens.Text = Format(0, "FIXED")
        dPercDescontoItensAnt = 0
    End If
    
    SubTotal_Calcula = SUCESSO

    Exit Function

Erro_SubTotal_Calcula:

    SubTotal_Calcula = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156709)

    End Select

    Exit Function

End Function

Public Function Total_Calcula(Optional iTesta As Integer = 1) As Long

Dim lErro As Long
Dim dTotal As Double

On Error GoTo Erro_Total_Calcula

    lErro = gobjTribTab.ValorTotal_Calcula(dTotal, iTesta)
    If lErro <> SUCESSO Then gError 43786

    Total.Text = Format(dTotal, "Standard")

    Total_Calcula = SUCESSO

    Exit Function

Erro_Total_Calcula:

    Total_Calcula = gErr

    Select Case gErr

        Case 43786

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156459)

    End Select

    Exit Function

End Function

Public Function Calcula_Valores(iLinha As Integer) As Long
'recalcula os valores de desconto, percentual de desconto e valor total

Dim sProduto As String
Dim lErro As Long
Dim lTamanho As Long
Dim dPercentDesc As Double
Dim dValorUnitario As Double
Dim dDesconto As Double
Dim dValorReal As Double
Dim dQuantidade As Double

On Error GoTo Erro_Calcula_Valores

    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
    
    lErro = CF2(Me, "NFEntCom_PrecoTotal_Calcula", iLinha, dQuantidade)
    If lErro <> SUCESSO Then gError 181259

    'Recolhe os valores Quantidade, Desconto, PerDesc e Valor Unitário da tela
    If dQuantidade = 0 Or Len(Trim(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))) = 0 Then

        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = ""
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(0, "Standard")
        
    Else

        dValorUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))
        dDesconto = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col))
        
        lTamanho = Len(Trim(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col)))

        If lTamanho > 0 Then
            dPercentDesc = PercentParaDbl(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col))
        Else
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        End If

        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dValorUnitario, dPercentDesc, dDesconto, dValorReal)

        'Coloca o Desconto calculado na tela
        If dDesconto > 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
        Else
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        End If
        
        'Coloca o valor Real em Valor Total
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(dValorReal, "Standard")

        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dValorUnitario, 0, 0, dValorReal)

        'Coloca o valor Real em Valor Total
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(dValorReal, "Standard")
    End If

    lErro = SubTotal_Calcula()
    If lErro <> SUCESSO Then gError 65754

    Calcula_Valores = SUCESSO
    
    Exit Function
    
Erro_Calcula_Valores:

    Calcula_Valores = gErr
    
    Select Case gErr

        Case 65754, 181259

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156711)

    End Select

    Exit Function
    
End Function

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Saida_Celula_ValorUnitario(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorUnitario As Double

On Error GoTo Erro_Saida_Celula_ValorUnitario

    Set objGridInt.objControle = ValorUnitario

    'Se estiver preenchido
    If Len(Trim(ValorUnitario.ClipText)) > 0 Then
        
        'Faz a crítica do valor
        If giComplemento = MARCADO Then
            lErro = Valor_NaoNegativo_Critica(ValorUnitario.Text)
        Else
            lErro = Valor_Positivo_Critica(ValorUnitario.Text)
        End If
        If lErro <> SUCESSO Then gError 65761

        dValorUnitario = CDbl(ValorUnitario.Text)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 65762
    
    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 65763

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 65764
'*** fim tributacao
    
    Call ValorTotal_Calcula
    
    Saida_Celula_ValorUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorUnitario:

    Saida_Celula_ValorUnitario = gErr

    Select Case gErr

        Case 65761, 65762, 65763, 65764
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156713)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercDesc(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim lTamanho As Long
Dim dPercentDescAnterior As Double

On Error GoTo Erro_Saida_Celula_PercDesc

    Set objGridInt.objControle = PercentDesc

    dPercentDesc = StrParaDbl(PercentDesc.Text)
    
    lTamanho = Len(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col))
    If lTamanho > 0 Then dPercentDescAnterior = StrParaDbl(left(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col), lTamanho - 1))

    'verifica se o percentual está preenchido
    If Len(Trim(PercentDesc.Text)) > 0 Then
    
        'Critica a procentagem
        lErro = Porcentagem_Critica(PercentDesc.Text)
        If lErro <> SUCESSO Then gError 65765

'        dPercentDesc = CDbl(PercentDesc.Text)
'
'        lTamanho = Len(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col))
'        If lTamanho > 0 Then dPercentDescAnterior = PercentParaDbl(left(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col), lTamanho - 1))

        If dPercentDesc <> dPercentDescAnterior Then
        
            'Verifica se o percentual é de 100%
            If dPercentDesc = 100 Then gError 65766
    
            PercentDesc.Text = Format(dPercentDesc, "Fixed")
    
            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) > 0 Then dPrecoUnitario = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))
            
        End If
    Else
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
        
    End If
    
    If Abs(dPercentDesc - dPercentDescAnterior) > DELTA_VALORMONETARIO2 Then
        iDescontoAlterado = REGISTRO_ALTERADO
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
     If lErro <> SUCESSO Then gError 65767

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 65768

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 65769
'*** fim tributacao

    Call ValorTotal_Calcula
    
    Saida_Celula_PercDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercDesc:

    Saida_Celula_PercDesc = gErr

    Select Case gErr

        Case 65765, 65767, 65768, 65769
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 65766
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156714)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercentDesc As Double
Dim dValorTotal As Double
Dim dDesconto As Double
Dim dValorUnitario As Double
Dim dValorReal As Double
Dim dQuantidade As Double
Dim dSubTotal As Double
Dim lTamanho As Long

On Error GoTo Erro_Saida_Celula_Desconto

    Set objGridInt.objControle = Desconto

    'Se estiver preenchido
    If Len(Trim(Desconto.ClipText)) > 0 Then
    
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(Desconto.Text)
        If lErro <> SUCESSO Then gError 65770

        dDesconto = CDbl(Desconto.Text)

        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) > 0 Then dQuantidade = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
        
        lErro = CF2(Me, "NFEntCom_PrecoTotal_Calcula", GridItens.Row, dQuantidade)
        If lErro <> SUCESSO Then gError 181270
        
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) > 0 Then dValorUnitario = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))
        
        dValorTotal = dValorUnitario * dQuantidade

        If dValorTotal > 0 Then
        
            dPercentDesc = dDesconto / (dValorTotal)
           
            If dPercentDesc >= 1 Then gError 65771
            
        End If

        GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")

    Else
    
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) <> 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) <> 0 Then

            GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = ""
        
        End If

    End If

    'Verifica se o valor de desconto foi alterado
    If Abs(StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col)) - dDesconto) > DELTA_VALORMONETARIO Then
        iDescontoAlterado = REGISTRO_ALTERADO
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 65772
   
    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 65773

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 65774
'*** fim tributacao
    
    Call ValorTotal_Calcula
    
    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = gErr

    Select Case gErr

        Case 65770, 65772, 65773, 65774, 181270
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 65771
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156715)

    End Select

    Exit Function

End Function

Public Sub ValorSeguro_Validate(Cancel As Boolean)

Dim dValorSeguro As Double
Dim lErro As Long

On Error GoTo Erro_Valorseguro_Validate

    If iValorSeguroAlterado = 0 Then Exit Sub

    'Se estiver preenchido
    If Len(Trim(ValorSeguro.Text)) > 0 Then
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(ValorSeguro.Text)
        If lErro <> SUCESSO Then gError 65775

        dValorSeguro = CDbl(ValorSeguro.Text)
        'Coloca o valor formatado na tela
        ValorSeguro.Text = Format(dValorSeguro, "Standard")

    End If

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorSeguro_Validate(Cancel, dValorSeguro)
'*** fim tributacao

    Call ValorTotal_Calcula

    iValorSeguroAlterado = 0
    
    Exit Sub

Erro_Valorseguro_Validate:

    Cancel = True


    Select Case gErr

        Case 65775

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156716)

    End Select

    Exit Sub

End Sub

Public Sub ValorDespesas_Validate(Cancel As Boolean)

Dim dValorDespesas As Double
Dim lErro As Long

On Error GoTo Erro_ValorDespesas_Validate

    If iValorDespesasAlterado = 0 Then Exit Sub

    'Se  estiver preenchido
    If Len(Trim(ValorDespesas.Text)) > 0 Then

        'Faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorDespesas.Text)
        If lErro <> SUCESSO Then gError 65776

        dValorDespesas = CDbl(ValorDespesas.Text)

        'coloca o valor formatado na tela
        ValorDespesas.Text = Format(dValorDespesas, "Standard")

    End If

    Call gobjTribTab.ValorDespesas_Validate(Cancel, dValorDespesas)
    
    Call ValorTotal_Calcula

    iValorDespesasAlterado = 0
        
    Exit Sub

Erro_ValorDespesas_Validate:

    Cancel = True


    Select Case gErr

        Case 65776

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156717)

    End Select

    Exit Sub

End Sub

Public Sub ValorDesconto_Validate(Cancel As Boolean)

Dim dValorDesconto As Double
Dim lErro As Long
Dim dSubTotal As Double
Dim dValorFrete As Double
Dim dValorSeguro As Double

On Error GoTo Erro_ValorDesconto_Validate

    'Verifica se o valor foi alterado
    If iValorDescontoAlterado = 0 Then Exit Sub

    'Se estiver preenchido
    If Len(Trim(ValorDesconto.Text)) > 0 Then
        
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(ValorDesconto.Text)
        If lErro <> SUCESSO Then gError 65777

        dValorDesconto = CDbl(ValorDesconto.Text)
        'Coloca o valor formatado na tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")

    End If

    Call SubTotal_Calcula
    
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDesconto_Validate(Cancel, dValorDesconto)
'*** fim tributacao
   
    Call ValorTotal_Calcula

    iValorDescontoAlterado = 0

    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True

    Select Case gErr

        Case 65777

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156718)

    End Select

    Exit Sub

End Sub

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iAlmoxarifado As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iCont As Integer
Dim objItemNF As ClassItemNF

On Error GoTo Erro_ProdutoLinha_Preenche

    'Preenche no Grid a Descrição do Produto e a Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMCompra
    GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objProduto.sDescricao

    'Preço Unitário
    For iCont = 1 To gcolItemPedCompraInfo.Count
        If objProduto.sCodigo = gcolItemPedCompraInfo(iCont).sProduto Then
            GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(gcolItemPedCompraInfo(iCont).dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
            Exit For
        End If
    Next
    
    lErro = CF2(Me, "NFEntCom_ProdutoLinha_Preenche", objProduto)
    If lErro <> SUCESSO Then gError 181258
    
    'Se necessário cria uma nova linha no Grid
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1

        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItens.Row)

''*** incluidos p/tratamento de tributacao *******************************
        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 65781
'*** fim tributacao
    
        Set objItemNF = New ClassItemNF
        
        objItemNF.sProduto = objProduto.sCodigo
        objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
        objItemNF.iItem = GridItens.Row
        objItemNF.lNumIntDoc = 0
        objItemNF.sDescricaoItem = objProduto.sDescricao
        
        gobjNFiscal.ColItensNF.Add1 objItemNF

    End If

    ProdutoLinha_Preenche = SUCESSO
    
    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr
    
    Select Case gErr

        Case 65778, 65779, 65781, 181258

        Case 65780
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156719)

    End Select

    Exit Function

End Function

Public Sub Fornecedor_GotFocus()

    sFornecedorAnterior = Trim(Fornecedor.Text)

End Sub

Function Atualiza_ListaPedidos() As Long
'Atualiza a Lista de Pedidos de Compra com os códigos desses pedidos

Dim lErro As Long
Dim iFilial As Integer
Dim iFilialCompra As Integer
Dim objPedidoCompras As ClassPedidoCompras
Dim objFornecedor As New ClassFornecedor
Dim iIndice As Integer

On Error GoTo Erro_Atualiza_ListaPedidos

    If gbCarregandoTela = True Then Exit Function
            
    'Limpa a tela de Pedidos e Requisições
    Call Limpa_Tela_PCRC

    'Se Fornecedor, Filial e Filial de Compra estão preenchidos
    If Len(Trim(TipoNFiscal)) > 0 And Len(Trim(Fornecedor.ClipText)) > 0 And Len(Trim(Filial.Text)) > 0 And Len(Trim(FilialCompra.Text)) > 0 Then

        sFornecedorAnterior = Fornecedor.Text
        iFilialAnterior = Codigo_Extrai(Filial.Text)
        iFilialCompraAnterior = Codigo_Extrai(FilialCompra.Text)
        
        objFornecedor.sNomeReduzido = Fornecedor.Text
        iFilial = Codigo_Extrai(Filial.Text)
        iFilialCompra = Codigo_Extrai(FilialCompra.Text)

        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 65782
        If lErro = 6681 Then gError 65783

        'Lê Pedidos de Compra com o Fornecedor, Filial e FilialCompra da tela
        lErro = CF("PedidosCompraEnv_Le_Recebimento", gcolPedidoCompra, objFornecedor.lCodigo, iFilial, iFilialCompra)
        If lErro <> SUCESSO And lErro <> 65845 Then gError 65784
        If lErro = 65845 Then gError 65785

        'Se não está trazendo um Recebimento para a tela
        If gbCarregandoTela = False Then
        
            'Verifica para cada Pedido lido se existe pelo menos um itemPC com quantidade a receber
            lErro = CF("ItensPedidosCompra_VerificaQuantidade", gcolPedidoCompra)
            If lErro <> SUCESSO Then gError 67250
        
        End If

        'se esta coleção estiver vazia ==> as quantidades que se queria comprar já foram totalmente recebidas
        If gcolPedidoCompra.Count = 0 Then gError 92134

        'Para cada Pedido Lido
        For iIndice = gcolPedidoCompra.Count To 1 Step -1
        
            Set objPedidoCompras = gcolPedidoCompra(iIndice)
        
            'Se for uma nota de beneficiamento
            If iTipoNFAlterado = DOCINFO_NFEEBFCOM Or iTipoNFAlterado = DOCINFO_NFIEBFCOM Then
                'Remove os Pedidos c\ destino <> Fornecedor
                If objPedidoCompras.iTipoDestino <> TIPO_DESTINO_FORNECEDOR Then gcolPedidoCompra.Remove (iIndice)
            Else
            
                If objPedidoCompras.iTipoDestino = TIPO_DESTINO_FORNECEDOR And objFornecedor.lCodigo = objPedidoCompras.lFornCliDestino And iFilial = objPedidoCompras.iFilialDestino Then
                
                    'if iTipoNFAlterado <> DOCINFO.... then gcolPedidoCompra.Remove (iIndice)
                    
                Else
                
                    'Remove os PCs com destino <> empresa ou outra filialempresa
                    lErro = CF("Remove_Pedido_Compra", objPedidoCompras, gcolPedidoCompra, iIndice)
                    If lErro <> SUCESSO Then gError 126715
            
                End If
                
            End If
        
        Next

        'Preenche a ListBox de pedidos com os pedidos lidos do BD
        For Each objPedidoCompras In gcolPedidoCompra
            PedidosCompra.AddItem objPedidoCompras.lCodigo
        Next

    End If

    Atualiza_ListaPedidos = SUCESSO

    Exit Function

Erro_Atualiza_ListaPedidos:

    Atualiza_ListaPedidos = gErr

    Select Case gErr

        Case 65782, 65784, 67250, 126715

        Case 65783
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", objFornecedor.sNomeReduzido)
        
        Case 65785
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PEDIDOCOMPRAS", gErr, objFornecedor.sNomeReduzido, iFilial, iFilialCompra)

        Case 92134
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOSCOMPRA_JA_RECEBIDOS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156720)

    End Select

    Exit Function

End Function

Public Sub Filial_GotFocus()

    iFilialAnterior = Codigo_Extrai(Filial.Text)

End Sub

Public Sub PedidosCompra_ItemCheck(Item As Integer)
 
Dim lErro As Long
Dim objPedidoCompra As New ClassPedidoCompras
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim objItemPC As ClassItemPedCompra
Dim objProduto As New ClassProduto
Dim colReqCompras As New Collection
Dim iIndice As Integer, iIndice2 As Integer
Dim iCont As Integer
Dim iProdutoEncontrado As Integer
Dim dFator As Double
Dim dQuantidade As Double
Dim dQuantAtual As Double
Dim sProdutoMascarado As String
Dim iItem As Integer
Dim lReqCompra As Long
Dim sProduto As String
Dim dDescontoAtual As Double
Dim dValorUnitario As Double
Dim dValorTotal As Double
Dim bTeste As Boolean
Dim dQuantidadeAux As Double
Dim dPrecoAux As Double, bJaMarcadaRC As Boolean
Dim objPedCompraAux As ClassPedidoCompras
Dim objItemPCInfoAux As ClassItemPedCompraInfo
Dim dtDataMaisRecente As Date, objProjetoInfo As Object

On Error GoTo Erro_PedidosCompra_ItemCheck
    
    'Se não existem pedidos na list --> Sai
    If PedidosCompra.ListCount = 0 Then Exit Sub
    
    'Atualiza as quantidades das coleções globais
    Call Atualiza_ItensPC_ItensRC
    
    'Se o pedido clicado na Lista de Pedidos de Compra está marcado
    If PedidosCompra.Selected(Item) = True Then

        'Recolhe o pedido da Coleção global
        Set objPedidoCompra = gcolPedidoCompra.Item(Item + 1)

        'Adiciona na combo de pedido do frame 3
        ComboPedidoCompras.AddItem PedidosCompra.List(PedidosCompra.ListIndex)
        
        'Se não está trazendo uma Nota Fiscal para a tela
        If gbCarregandoTela = False Then
        
            'Lê os itens do Pedido de Compra passado por objPedidoCompra
            lErro = CF("ItensPC_Le_Codigo", objPedidoCompra)
            If lErro <> SUCESSO Then gError 65786
            
            lErro = CF("PCItens_Le_Customizado", objPedidoCompra)
            If lErro <> SUCESSO Then gError 181263
            
            'Para cada Item do Pedido de Compras
            For Each objItemPC In objPedidoCompra.colItens
    
                iItem = iItem + 1
    
                If objItemPC.iMoeda = MOEDA_REAL Then objItemPC.dTaxa = 1
                
                Set objItemPCInfo = New ClassItemPedCompraInfo
    
                sProdutoMascarado = String(STRING_PRODUTO, 0)
                
                lErro = Mascara_RetornaProdutoEnxuto(objItemPC.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 65787
                
                ProdutoPC.PromptInclude = False
                ProdutoPC.Text = sProdutoMascarado
                ProdutoPC.PromptInclude = True
        
                sProdutoMascarado = ProdutoPC.Text
                
                objProduto.sCodigo = objItemPC.sProduto

                'Lê o Produto passado por objProduto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 65788
                If lErro = 28030 Then gError 65789
    
                'Se não é permitido que o preço unitário e as alíquotas do Item NF sejam diferentes do ItemPC
                If gobjCOM.iNFDiferentePC = NFISCAL_NAO_ACEITA_DIFERENCA_PC Then

                    'Verifica se o Preço Unitário de objItemPC é o mesmo do Produto que está no Grid. Se não for, erro.
                    For iIndice = 1 To objGridItens.iLinhasExistentes
                    
                        bTeste = True
                        lErro = CF2(Me, "NFEntCom_Produto_Igual_Customizado", objItemPC.objInfoUsu, iIndice, bTeste)
                        If lErro <> SUCESSO Then gError 181284

                        If sProdutoMascarado = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) And bTeste Then
                        
                            'Atualiza a Quantidade do Grid de Recebimento
                            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPC.sUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), dFator)
                            If lErro <> SUCESSO Then gError 89183
                        
                            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) > 0 And Abs(StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col)) - objItemPC.dPrecoUnitario * dFator * objItemPC.dTaxa) > DELTA_VALORMONETARIO And gbCarregandoTela = False Then gError 66000
                        End If
                    Next

                    'Verifica se a AliquotaIPI e a AliquotaICMS do Produto de objItePC são iguais às dos Produtos que já estão no Grid.
                    For Each objItemPCInfo In gcolItemPedCompraInfo
                    
                        bTeste = True
                        lErro = CF2(Me, "NFEntCom_Produto_Igual_Customizado3", objItemPC.objInfoUsu, objItemPCInfo.objInfoUsu, bTeste)
                        If lErro <> SUCESSO Then gError 181284

                        If objItemPCInfo.sProduto = objItemPC.sProduto And bTeste Then
                            If objItemPCInfo.dAliquotaIPI <> objItemPC.dAliquotaIPI Then gError 66002
                            If objItemPCInfo.dAliquotaICMS <> objItemPC.dAliquotaICMS Then gError 66001
                        End If
                    Next

                End If
                
                iProdutoEncontrado = 0
    
                Set objItemPCInfo = New ClassItemPedCompraInfo
                
                'Preenche o objItemPCInfo a partir do objItemPC
                objItemPCInfo.dPercentMaisReceb = objItemPC.dPercentMaisReceb
                objItemPCInfo.dQuantReceber = objItemPC.dQuantidade - objItemPC.dQuantRecebida
                objItemPCInfo.dQuantRecebida = objItemPC.dQuantidade - objItemPC.dQuantRecebida
                objItemPCInfo.dQuantPedida = objItemPC.dQuantidade
                objItemPCInfo.iItem = iItem
                objItemPCInfo.iRecebForaFaixa = objItemPC.iRebebForaFaixa
                objItemPCInfo.lNumIntDoc = objItemPC.lNumIntDoc
                objItemPCInfo.lPedCompra = objPedidoCompra.lCodigo
                objItemPCInfo.sDescProduto = objItemPC.sDescProduto
                objItemPCInfo.sProduto = objItemPC.sProduto
                objItemPCInfo.sUM = objItemPC.sUM
                objItemPCInfo.dPrecoUnitario = objItemPC.dPrecoUnitario
                objItemPCInfo.dAliquotaIPI = objItemPC.dAliquotaIPI
                objItemPCInfo.dAliquotaICMS = objItemPC.dAliquotaICMS
                objItemPCInfo.iMoeda = objItemPC.iMoeda
                objItemPCInfo.dTaxa = objItemPC.dTaxa
                
                Set objItemPCInfo.objInfoUsu = objItemPC.objInfoUsu

                If objPedidoCompra.dTaxa = 0 Then
                    objItemPCInfo.bTaxaPedido = True
                Else
                    objItemPCInfo.bTaxaPedido = False
                End If
                
                objItemPCInfo.dPrecoUnitario = objItemPC.dPrecoUnitario
                objItemPCInfo.dValorDesconto = objItemPC.dValorDesconto

                
                'Verifica se o produto está presente no grid de itens de recebimento
                For iIndice = 1 To objGridItens.iLinhasExistentes
                
                    bTeste = True
                    lErro = CF2(Me, "NFEntCom_Produto_Igual_Customizado", objItemPC.objInfoUsu, iIndice, bTeste)
                    If lErro <> SUCESSO Then gError 181284
    
                    'Se encontrou o produto no grid de Recebimento
                    If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado And bTeste Then
    
                        'Atualiza a Quantidade do Grid de Recebimento
                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPC.sUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), dFator)
                        If lErro <> SUCESSO Then gError 65790
    
                        dQuantidade = (objItemPC.dQuantidade - objItemPC.dQuantRecebida) * dFator
                        
                        'Calcula a quantidade Atual levanod em conta o novo item inserido.
                        dQuantAtual = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)) + dQuantidade
                                               
                        If gobjEST.iFormaCalcPrecoEntCom = FORMA_CALC_PESO_ENT_COM_MEDIA Then
                            'Calcula por proporção um novo preço unitário
                            dPrecoAux = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)) * StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))
                            dPrecoAux = dPrecoAux + objItemPCInfo.dQuantReceber * objItemPC.dPrecoUnitario * objItemPC.dTaxa
                            dPrecoAux = dPrecoAux / dQuantAtual
                            GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(dPrecoAux, IIf(gobjCOM.sFormatoPrecoUnitario <> "", gobjCOM.sFormatoPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO))
                        End If
                        
                        If gobjEST.iFormaCalcPrecoEntCom = FORMA_CALC_PESO_ENT_COM_ULT Then
                            For Each objItemPCInfoAux In gcolItemPedCompraInfo
                            
                                bTeste = True
                                lErro = CF2(Me, "NFEntCom_Produto_Igual_Customizado3", objItemPC.objInfoUsu, objItemPCInfoAux.objInfoUsu, bTeste)
                                If lErro <> SUCESSO Then gError 181284
                                
                                dtDataMaisRecente = DATA_NULA
                                
                                If objItemPCInfoAux.sProduto = objItemPC.sProduto And bTeste Then
                                    For Each objPedCompraAux In gcolPedidoCompra
                                        If objItemPCInfoAux.lPedCompra = objPedCompraAux.lCodigo Then
                                            If dtDataMaisRecente = DATA_NULA Or dtDataMaisRecente < objPedCompraAux.dtDataEnvio Then
                                                dtDataMaisRecente = objPedCompraAux.dtDataEnvio
                                                dPrecoAux = objItemPCInfoAux.dPrecoUnitario
                                            End If
                                            Exit For
                                        End If
                                    Next
                                End If
                            Next
                            If objPedidoCompra.dtDataEnvio > dtDataMaisRecente Or dtDataMaisRecente = DATA_NULA Then
                                dPrecoAux = objItemPCInfo.dPrecoUnitario
                            End If
                        End If
                                               
                        'Atualiza a quantidade do item de Recebimento do produto
                        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(dQuantAtual)
    
                        'calcula o novo percentual e valor de desconto
                        If objItemPC.dValorDesconto <> 0 Then
                        
                            dDescontoAtual = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col)) + (objItemPC.dValorDesconto * objItemPC.dTaxa)
                            
                            GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(dDescontoAtual, "Standard")
                            
                            If dPrecoAux <> 0 Then
                                GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dDescontoAtual / (dQuantAtual * dPrecoAux), "Percent")
                            End If
                            
                        End If
                        
                        lErro = CF2(Me, "NFEntCom_Calcula_PesoKg", iIndice)
                        If lErro <> SUCESSO Then gError 181288
    
                        lErro = Calcula_Valores(iIndice)
                        If lErro <> SUCESSO Then gError 89179
    
                        ''*** incluidos p/tratamento de tributacao *******************************
                        lErro = gobjTribTab.Alteracao_Item_Grid(iIndice)
                        If lErro <> SUCESSO Then gError 89179
                        '*** fim tributacao
                        
                        Exit For
    
                    End If
    
                Next
    
                'Se não encontrou o Produto
                If iIndice > objGridItens.iLinhasExistentes Then
                    
                    'Se o produto possui quantidade a receber
                    If objItemPCInfo.dQuantReceber > 0 Then
                        
                        'Preenche uma linha do GridItens
                        GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_Produto_Col) = sProdutoMascarado
    
                        For iCont = 0 To Produto.ListCount - 1
                            If Produto.List(iCont) = sProdutoMascarado Then
                                Produto.ListIndex = iCont
                                Exit For
                            End If
                        Next
                    
                        lErro = PreencheLinha_GridItens(objItemPC)
                        If lErro <> SUCESSO Then gError 65791
                
                    'Se esse item não tem quantidade a Receber
                    Else
                        'O Próximo reinicia a procura
                        iIndice = 0
                
                    End If
                
                End If
        
                Call InsereOrdenadaPC(objItemPCInfo)
                
                'distribuicao - busca a distribuição dos produtos oriunda do pedido de compra
                lErro = CF("LocalizacaoItemPC_Le", objItemPC)
                If lErro <> SUCESSO And lErro <> 56361 Then gError 89614
                
                Set objItemPCInfo.colLocalizacao = objItemPC.colLocalizacao
                
            Next
        
            'Limpa o GridItensPC
            Call Grid_Limpa(objGridItensPC)
            
            'Preenche o GridItensPC
            For iIndice = 1 To gcolItemPedCompraInfo.Count
                lErro = PreencheLinha_ItensPC(gcolItemPedCompraInfo(iIndice))
                If lErro <> SUCESSO Then gError 65792
            Next
            
            'Lê as Requisições de Compra vinculadas com estes itens do Pedido
            lErro = CF("PedidoCompra_Le_ReqComprasVinculada", objPedidoCompra, colReqCompras)
            If lErro <> SUCESSO And lErro <> 65796 Then gError 65798
        
        'Se está trazendo uma Nota Fiscal para a tela
        Else
            
            'Lê as Requisições de Compra (baixadas ou não) vinculadas com estes itens do Pedido
            lErro = CF("ReqComprasTodas_Le_PedidoCompra", objPedidoCompra, colReqCompras)
            If lErro <> SUCESSO Then gError 66566

        End If
                                                             
        objPedidoCompra.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_PC
        objPedidoCompra.objInfoAdic.lNumIntDoc = objPedidoCompra.lNumIntDoc
        lErro = CF("InfoAdicionais_Le", objPedidoCompra.objInfoAdic)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 66566
        Set gobjInfoAdic = objPedidoCompra.objInfoAdic
               
        lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objProjetoInfo, objPedidoCompra.lNumIntDoc, PRJ_CR_TIPO_PC)
        If lErro <> SUCESSO Then gError 66566
        
        lErro = gobjTelaProjetoInfo.Recalcula_Dados_Tela
        If lErro <> SUCESSO Then gError 66566
                                                             
        'Preenche a lista de Requisições
        For iIndice = 1 To colReqCompras.Count
            
            bJaMarcadaRC = False
            iIndice2 = 0
            Do While iIndice2 < RequisicoesCompra.ListCount
                If RequisicoesCompra.List(iIndice2) = CStr(colReqCompras.Item(iIndice)) Then
                    bJaMarcadaRC = True
                    Exit Do
                End If
                iIndice2 = iIndice2 + 1
            Loop
            
            If bJaMarcadaRC = False Then
            
                RequisicoesCompra.AddItem colReqCompras.Item(iIndice)
                gbCarregandoPC = True
                RequisicoesCompra.Selected(RequisicoesCompra.NewIndex) = True
                gbCarregandoPC = False
            
            Else
            
                gbCarregandoPC = True
                Call RequisicoesCompra_ItemCheck(iIndice2)
                gbCarregandoPC = False
            
            End If
        
        Next
    
        'Trata a tributação dos Itens do pedido
        lErro = Trata_Tributacao()
        If lErro <> SUCESSO Then gError 66121
        
    'Se o pedido foi desselecionado
    Else
                
        'Guarda o Pedido de compras que foi desselecionado
        Set objPedidoCompra = gcolPedidoCompra.Item(Item + 1)
        
        'Atualiza o Grid de NF e o Grid de Requisição de compras
        lErro = Atualiza_GridItens(objPedidoCompra, Item)
        If lErro <> SUCESSO Then gError 65801
               
        'Remove da combo
        For iIndice = 1 To ComboPedidoCompras.ListCount - 1
            If ComboPedidoCompras.List(iIndice) = objPedidoCompra.lCodigo Then
                ComboPedidoCompras.RemoveItem (iIndice)
                Exit For
            End If
        Next

        Set objPedidoCompra.colItens = New Collection

    End If
    
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 181260
    End If
    
    lErro = CF2(Me, "NFEntCom_Calcula_PesoKg_Total")
    If lErro <> SUCESSO Then gError 181260

    Call ValorTotal_Calcula
       
    If ComboPedidoCompras.ListIndex = 0 Then ComboPedidoCompras_Click
    
    ComboPedidoCompras.ListIndex = 0
       
    lErro = gobjDistribuicao.Preenche_GridDistribuicaoPC1(gcolItemPedCompraInfo)
    If lErro <> SUCESSO Then gError 89628
    
    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_PedidosCompra_ItemCheck:

    Select Case gErr

        Case 65786, 65787, 65788, 65790, 65791, 65792, 65798, 65801, 66121, 66566, 89179, 89183, 89614, 89615, 89628, 181263, 181284, 181288, 181260

        Case 65789
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case 66000
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_PC_PRECOUNITARIO_DIFERENTE", gErr, objItemPC.sProduto)
            PedidosCompra.Selected(Item) = False
        
        Case 66001
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_PC_ALIQUOTAICMS_DIFERENTE", gErr, objItemPC.sProduto)
            PedidosCompra.Selected(Item) = False
            
        Case 66002
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_PC_ALIQUOTAIPI_DIFERENTE", gErr, objItemPC.sProduto)
            PedidosCompra.Selected(Item) = False
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156721)

    End Select

    Exit Sub

End Sub

Function Trata_Tributacao() As Long

Dim iIndice As Integer
Dim iCont As Integer
Dim dQuantidade As Double
Dim lReqCompra As Long
Dim sProduto As String
Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objItemRCInfo As New ClassItemReqComprasInfo
Dim iTipoTributacao As Integer

On Error GoTo Erro_Trata_Tributacao

    'Para cada linha do GridItens
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        dQuantidade = 0
        
        'Procura no Grid de Requisições o item de mesmo produto que possui o maior valor
        For iCont = 1 To objGridItensRC.iLinhasExistentes
            
            If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = GridItensRC.TextMatrix(iCont, iGrid_ProdutoRC_Col) Then
            
                Set objItemRCInfo = gcolItemReqCompraInfo(iCont)
                
                If StrParaDbl(GridItensRC.TextMatrix(iCont, iGrid_AReceberRC_Col)) > dQuantidade And objItemRCInfo.iTipoTributacao <> 0 Then
                    
                    'Guarda o código do Produto de maior valor
                    dQuantidade = StrParaDbl(GridItensRC.TextMatrix(iCont, iGrid_AReceberRC_Col))
                    sProduto = GridItensRC.TextMatrix(iCont, iGrid_ProdutoRC_Col)
                    lReqCompra = CLng(GridItensRC.TextMatrix(iCont, iGrid_ReqCompra_Col))
                    iTipoTributacao = objItemRCInfo.iTipoTributacao
                
                End If
            
            End If
                      
        Next
    
        'Se achou o Produto
        If Len(Trim(sProduto)) > 0 Then
            
            'Formata o Produto
            lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 66126
                    
            'Procura esse produto na coleção de itens de gobjTribTab
            For iCont = 1 To gobjTribTab.objNFTributacao.colItens.Count
                
                If gobjTribTab.objNFTributacao.colItens(iCont).sProduto = sProdutoFormatado Then
                
                    'Coloca o Tipo de Tributação deste item no item de Nota Fiscal
                    gobjTribTab.objNFTributacao.colItens(iCont).iTipoTributacao = iTipoTributacao
                    gobjTribTab.objNFTributacao.colItens(iCont).iTipoTributacaoManual = VAR_PREENCH_AUTOMATICO
                    
                End If
            
            Next
        
        End If
    
    Next

    Trata_Tributacao = SUCESSO
    
    Exit Function
    
Erro_Trata_Tributacao:

    Trata_Tributacao = gErr
    
    Select Case gErr

        Case 66126
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156722)
    
    End Select
    
    Exit Function
    
End Function

Function InsereOrdenadaPC(objItemPCInfo As ClassItemPedCompraInfo, Optional ByVal bIgnoraQtdReceb As Boolean = False) As Long
'Insere novo item de Pedido de compras na coleção global de ItensPC de forma ordenada

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_InsereOrdenadaPC

    'Se o produto possui quantidade a receber
    If (objItemPCInfo.dQuantReceber > 0) Or bIgnoraQtdReceb Then

        iIndice = 1
        
        'Se a coleção está vazia
        If gcolItemPedCompraInfo.Count = 0 Then
            'Insere o ItemPC na coleção
            gcolItemPedCompraInfo.Add objItemPCInfo
        
        'Se não
        Else
        
            'Para cada Item da coleção de ItensPC
            Do While iIndice <= gcolItemPedCompraInfo.Count
                
                'Se o código do Pedido de Compra do Item passado for maior ou igual do que o que está na coleção
                If objItemPCInfo.lPedCompra >= gcolItemPedCompraInfo(iIndice).lPedCompra Then
                    'Busca próximo Item da coleção
                    iIndice = iIndice + 1
                'Se não
                Else
                    'Adiciona o Item na coleção de ItensPC antes do Item da coleção que possui código PC maior
                    gcolItemPedCompraInfo.Add objItemPCInfo, , iIndice
                    Exit Do
                End If
                
            Loop
                        
            'Se o Pedido de Compra do Item é maior do que todos que estão na coleção
            If iIndice > gcolItemPedCompraInfo.Count Then
                'Adiciona o Item no final coleção
                gcolItemPedCompraInfo.Add objItemPCInfo
            End If
        
        End If
        
    End If
        
    InsereOrdenadaPC = SUCESSO
    
    Exit Function
    
Erro_InsereOrdenadaPC:

    InsereOrdenadaPC = gErr
    
    Select Case gErr
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156723)
    
    End Select
    
    Exit Function
    
End Function

Function Atualiza_GridItens(objPedidoCompra As ClassPedidoCompras, iItem As Integer) As Long
'Remove do Grids os produtos de Pedidos de compras desselecionados da lista

Dim objItemPC As ClassItemPedCompra
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim iCont As Integer
Dim sProdutoMascarado As String
Dim dQuantAtual As Double
Dim lErro As Long
Dim lReqCompra As Long
Dim iPossuiRelacionamento As Integer
Dim iIndice As Integer
Dim iLinha As Integer
Dim colPedCompras As New Collection
Dim iIndice2 As Integer, iQtdeRCs As Integer
Dim bTeste As Boolean

On Error GoTo Erro_Atualiza_GridItens
            
    'Retira do Grid de Pedido de Compras os itens do pedido de compras desselecionado
    For iCont = objGridItensPC.iLinhasExistentes To 1 Step -1
        If PedidosCompra.List(iItem) = GridItensPC.TextMatrix(iCont, iGrid_PedCompra_Col) Then
            lErro = Grid_Exclui_Linha(objGridItensPC, iCont)
            If lErro <> SUCESSO Then gError 65800
        End If
    Next

    'Retira da coleção os itens do Pedido de compras desselecionado
    For iCont = gcolItemPedCompraInfo.Count To 1 Step -1
        If gcolItemPedCompraInfo.Item(iCont).lPedCompra = CLng(PedidosCompra.List(iItem)) Then
            gcolItemPedCompraInfo.Remove (iCont)
        End If
    Next
    
    'Para cada Item do Pedido de Compras
    For Each objItemPC In objPedidoCompra.colItens
    
        iPossuiRelacionamento = 0
        
        sProdutoMascarado = String(STRING_PRODUTO, 0)
        lErro = Mascara_RetornaProdutoEnxuto(objItemPC.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 65802

        ProdutoPC.PromptInclude = False
        ProdutoPC.Text = sProdutoMascarado
        ProdutoPC.PromptInclude = True

        sProdutoMascarado = ProdutoPC.Text

        dQuantAtual = 0

        'Para cada Item de Pedido de Compras que foi selecionado
        For Each objItemPCInfo In gcolItemPedCompraInfo

            bTeste = True
            Call CF2(Me, "NFEntCom_Produto_Igual_Customizado3", objItemPC.objInfoUsu, objItemPCInfo.objInfoUsu, bTeste)

            'Se o Item de Pedido de Compras faz parte de outro pedido selecionado
            If objItemPCInfo.sProduto = objItemPC.sProduto And bTeste Then

                dQuantAtual = 0

                'Guarda a nova quantidade recebida desse produto
                For iIndice = 1 To objGridItensPC.iLinhasExistentes
                
                    bTeste = True
                    Call CF2(Me, "NFEntCom_Produto_Igual_Customizado4", objItemPC.objInfoUsu, iIndice, bTeste)

                    If sProdutoMascarado = GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) And bTeste Then
                        dQuantAtual = dQuantAtual + StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
                    End If
                Next

                'Atualiza o Grid de Itens
                For iIndice = 1 To objGridItens.iLinhasExistentes
                    
                    bTeste = True
                    Call CF2(Me, "NFEntCom_Produto_Igual_Customizado", objItemPC.objInfoUsu, iIndice, bTeste)
                    
                    If sProdutoMascarado = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) And bTeste Then
                        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(dQuantAtual)

                        lErro = CF2(Me, "NFEntCom_Calcula_PesoKg", iIndice)
                        If lErro <> SUCESSO Then gError 181288

                        lErro = Calcula_Valores(iIndice)
                        If lErro <> SUCESSO Then gError 89181

                        ''*** incluidos p/tratamento de tributacao *******************************
                        lErro = gobjTribTab.Alteracao_Item_Grid(iIndice)
                        If lErro <> SUCESSO Then gError 89181
                        '*** fim tributacao
                        
                        Exit For

                    End If
                Next

                iPossuiRelacionamento = 1
                Exit For

            End If

        Next
    
        'Se o Item do Pedido que foi desselecionado não está presente em nenhuma outra linha de GridItensPC
        If iPossuiRelacionamento = 0 Then
            
            'Exclui as linhas onde esse Produto aparece
            For iIndice = objGridItens.iLinhasExistentes To 1 Step -1
            
                bTeste = True
                Call CF2(Me, "NFEntCom_Produto_Igual_Customizado", objItemPC.objInfoUsu, iIndice, bTeste)

                If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado And bTeste Then
                    lErro = Grid_Exclui_Linha(objGridItens, iIndice)
                    If lErro <> SUCESSO Then gError 66078
                
                    'Exclui Tributação relacionada ao item
                    lErro = gobjTribTab.Exclusao_Item_Grid(iIndice)
                    If lErro <> SUCESSO Then gError 66115

                End If
            Next
        End If
        
    Next
                
    'Remove do GridItensRC os itens relacionados ao Pedido de Compras Desselecionado
    For iIndice = objGridItensRC.iLinhasExistentes To 1 Step -1
        
        If GridItensRC.TextMatrix(iIndice, iGrid_PedCompraRC_Col) = CStr(objPedidoCompra.lCodigo) Then
            
            lReqCompra = CLng(GridItensRC.TextMatrix(iIndice, iGrid_ReqCompra_Col))
            lErro = Grid_Exclui_Linha(objGridItensRC, iIndice)
            If lErro <> SUCESSO Then gError 66079
            
            'Procura no Grid por outro Pedido associado a mesma requisição da linha que foi excluida
            For iLinha = 1 To objGridItensRC.iLinhasExistentes
                If GridItensRC.TextMatrix(iLinha, iGrid_ReqCompra_Col) = CStr(lReqCompra) And GridItensRC.TextMatrix(iLinha, iGrid_PedCompra_Col) <> CStr(objPedidoCompra.lCodigo) Then
                    Exit For
                End If
            Next
            
            'Se não encontrar
            If iLinha > objGridItensRC.iLinhasExistentes Then
                    
                'Remove a Requisição da lista de Requisições
                For iCont = 0 To RequisicoesCompra.ListCount - 1
                    If RequisicoesCompra.List(iCont) = CStr(lReqCompra) Then
                        RequisicoesCompra.RemoveItem (iCont)
                    End If
                Next
                
            End If
        
        End If
    
    Next
        
    'Remove de gcolItemReqCompraInfo os itens relacionados ao Pedido de Compras Desselecionado
    For iIndice = gcolItemReqCompraInfo.Count To 1 Step -1
        If gcolItemReqCompraInfo.Item(iIndice).lPedCompra = objPedidoCompra.lCodigo Then
            gcolItemReqCompraInfo.Remove (iIndice)
        End If
    Next
    
    Atualiza_GridItens = SUCESSO

    Exit Function

Erro_Atualiza_GridItens:

    Atualiza_GridItens = gErr
    
    Select Case gErr

        Case 65800, 65802, 66078, 66079, 66115, 66129, 89181, 181288

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156724)

    End Select

    Exit Function

End Function

Function PreencheLinha_GridItens(objItemPC As ClassItemPedCompra) As Long
'Preenche uma linha no Grid de Recebimento com os dados passados em objItemPC

Dim iAlmoxarifadoPadrao As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lErro As Long
Dim dValorTotal As Double
Dim dTaxa As Double
Dim objItemNF As ClassItemNF

On Error GoTo Erro_PreencheLinha_GridItens

    Set objItemNF = New ClassItemNF
    
    objItemNF.sProduto = objItemPC.sProduto
    objItemNF.sUnidadeMed = objItemPC.sUM
    objItemNF.iItem = objGridItens.iLinhasExistentes + 1
    objItemNF.lNumIntDoc = 0
    objItemNF.sDescricaoItem = objItemPC.sDescProduto

    gobjNFiscal.ColItensNF.Add1 objItemNF

    'Descrição do produto
    GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_DescProduto_Col) = objItemPC.sDescProduto

    'Unidade de Medida
    GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_UnidadeMed_Col) = objItemPC.sUM
   
    'Quantidade
    GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_Quantidade_Col) = Formata_Estoque(objItemPC.dQuantidade - objItemPC.dQuantRecebida)

    'Preço unitário
    GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_ValorUnitario_Col) = Format(objItemPC.dPrecoUnitario * objItemPC.dTaxa, FORMATO_PRECO_UNITARIO_EXTERNO)
    
    'calcula o percentual e valor de desconto
    If objItemPC.dValorDesconto <> 0 Then

        dValorTotal = objItemPC.dPrecoUnitario * (objItemPC.dQuantidade - objItemPC.dQuantRecebida) * objItemPC.dTaxa

        GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_Desconto_Col) = Format(objItemPC.dValorDesconto * objItemPC.dTaxa, "Standard")

        If dValorTotal <> 0 Then
            
            GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, iGrid_PercDesc_Col) = Format(objItemPC.dValorDesconto * objItemPC.dTaxa / dValorTotal, "Percent")
        
        End If

    End If
    
    lErro = CF2(Me, "NFEntCom_PreencheLinha_GridItens", objItemPC, objGridItens.iLinhasExistentes + 1)
    If lErro <> SUCESSO Then gError 181262
    
    objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    
    Call gobjTelaProjetoInfo.Adiciona_Linha(objGridItens.iLinhasExistentes)
    
    lErro = Calcula_Valores(objGridItens.iLinhasExistentes)
    If lErro <> SUCESSO Then gError 66573

''*** incluidos p/tratamento de tributacao *******************************
        'permite que a tributacao reflita a inclusao de uma linha no grid
    lErro = gobjTribTab.Inclusao_Item_Grid(objGridItens.iLinhasExistentes, objItemPC.sProduto)
    If lErro <> SUCESSO Then gError 65781
    
    '??? jones 31/10/2005
    lErro = gobjTribTab.Alteracao_Item_Grid_ItemPC(objGridItens.iLinhasExistentes, objItemPC)
    If lErro <> SUCESSO Then gError 65781
'*** fim tributacao

    PreencheLinha_GridItens = SUCESSO

    Exit Function

Erro_PreencheLinha_GridItens:

    PreencheLinha_GridItens = gErr

    Select Case gErr

        Case 65781, 66010, 66011, 66573, 181262

        Case 66012
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156725)

    End Select

    Exit Function

End Function

Public Sub RequisicoesCompra_ItemCheck(Item As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim iLinha As Integer
Dim objItemReqCompraInfo As ClassItemReqComprasInfo
Dim colItemRCInfo As New Collection
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim objRequisicaoCompra As New ClassRequisicaoCompras
Dim sProdutoMascarado As String, dFator As Double, objProduto As New ClassProduto
Dim dTotalReq As Double, dDif As Double
Dim colItemPedCompraInfo As Collection

On Error GoTo Erro_RequisicoesCompra_ItemCheck

    'Se não há Requisições na lista de Requisições, sai da rotina
    If RequisicoesCompra.ListCount = 0 Then Exit Sub
    
    'Atualiza as quantidades das coleções globais
    Call Atualiza_ItensPC_ItensRC
    
    'Se estiver selecionando
    If RequisicoesCompra.Selected(Item) = True Then

        objRequisicaoCompra.lCodigo = CLng(RequisicoesCompra.List(Item))
        objRequisicaoCompra.iFilialEmpresa = Codigo_Extrai(FilialCompra.Text)
        
        'Se está trazendo uma Nota Fiscal para tela e não é um Recebimento
        If gbCarregandoTela = True And giRecebimento = False Then
        
            'Lê todos os Itens de Requisição de Compras associadas ao Pedido de Compras
            lErro = CF("ItensRequisicaoTodos_Le_PedidoCompra", objRequisicaoCompra, colItemRCInfo, lNumIntDocNF)
            If lErro <> SUCESSO Then gError 66101
        
        'Se está trazendo um Recebimento para a tela
        ElseIf gbCarregandoTela = True And giRecebimento = True Then
                        
            'Lê os Itens de Requisição de Compras não baixados associados ao Pedido de Compras e a Nota Fiscal
            lErro = CF("ItensRequisicao_Le_PedidoCompra_NF", objRequisicaoCompra, colItemRCInfo, lNumIntDocNF)
            If lErro <> SUCESSO And lErro <> 65851 Then gError 65803
            
            If lErro <> SUCESSO Then
            
            'Se não achou a requisição gravada  lê todas as requisições do pedido
                lErro = CF("ItensRequisicao_Le_PedidoCompra", objRequisicaoCompra, colItemRCInfo)
                If lErro <> SUCESSO Then gError 65803
                
               Set colItemPedCompraInfo = New Collection
            
               'Lê apenas os Itens de Pedido de Compras não baixados relacionados a Nota Fiscal
               lErro = CF("ItensPedCompra_Le_NFiscalReceb", gobjNFiscal, colItemPedCompraInfo)
               If lErro <> SUCESSO Then gError 65803
                
                For Each objItemPCInfo In colItemPedCompraInfo
                    
                    dTotalReq = 0
                    
                    For Each objItemReqCompraInfo In colItemRCInfo
                        If objItemPCInfo.lNumIntDoc = objItemReqCompraInfo.lNumIntDocItemPC Then
                            dTotalReq = dTotalReq + objItemReqCompraInfo.dQuantReceber * objItemReqCompraInfo.dFator
                        End If
                    Next
                    
                    If objItemPCInfo.dQuantRecebida < dTotalReq Then
                        
                        dDif = dTotalReq - objItemPCInfo.dQuantRecebida
                        
                        For iIndice = colItemRCInfo.Count To 1 Step -1
                            Set objItemReqCompraInfo = colItemRCInfo.Item(iIndice)
                            If objItemPCInfo.lNumIntDoc = objItemReqCompraInfo.lNumIntDocItemPC Then
                                If objItemReqCompraInfo.dQuantRecebida > dDif Then
                                    objItemReqCompraInfo.dQuantRecebida = objItemReqCompraInfo.dQuantRecebida - dDif
                                    dDif = 0
                                    Exit For
                                Else
                                    dDif = dDif - objItemReqCompraInfo.dQuantRecebida
                                    objItemReqCompraInfo.dQuantRecebida = 0
                                End If
                            End If
                        Next
                    
                    End If
                        
                Next
            
            End If
        
        'Se apenas marcou uma Requisição e não está trazendo uma nota fiscal para a tela
        Else
            
            'Lê os Itens de Requisição de Compras não baixados associados ao Pedido de Compras
            lErro = CF("ItensRequisicao_Le_PedidoCompra", objRequisicaoCompra, colItemRCInfo)
            If lErro <> SUCESSO Then gError 66143
        
        End If
        
        'Para cada Item de Requisição de Compras
        For Each objItemReqCompraInfo In colItemRCInfo

            'Verifica se o ItemRC está vinculado a um ItemPC cujo Pedido de Compras foi selecionado
            For iIndice = 0 To PedidosCompra.ListCount - 1
                If CLng(PedidosCompra.List(iIndice)) = objItemReqCompraInfo.lPedCompra And PedidosCompra.Selected(iIndice) = True Then
                    Call InsereOrdenadaRC(objItemReqCompraInfo)
                End If
            Next
        
        Next
        
        'Preenche de Itens de Requisição de Compras
        Call PreencheGrid_ItensRC
        
        'Só atualiza as quantidades quando não estiver trazendo uma NF para a tela
        If gbCarregandoTela = False And gbCarregandoPC = False Then
            'Atualiza as quantidades Recebidas
            Call Atualiza_Quantidades(Item)
        End If
        
    'Se estiver deselecionando
    Else

        'Atualiza as quantidades Recebidas
        Call Atualiza_Quantidades(Item)
        
        'Exclui as linhas do GridItensRC que possuem a Requisição que foi desselecionada
        For iIndice = objGridItensRC.iLinhasExistentes To 1 Step -1
            If GridItensRC.TextMatrix(iIndice, iGrid_ReqCompra_Col) = RequisicoesCompra.List(Item) Then
                lErro = Grid_Exclui_Linha(objGridItensRC, iIndice)
                If lErro <> SUCESSO Then gError 66080
            End If
        Next
        
        'Remove da coleção os ItensRC que possuem a Requisição que foi desselecionada
        For iIndice = gcolItemReqCompraInfo.Count To 1 Step -1
            If gcolItemReqCompraInfo.Item(iIndice).lReqCompra = CLng(RequisicoesCompra.List(Item)) Then
                gcolItemReqCompraInfo.Remove (iIndice)
            End If
        Next

        'Para cada Requisição da Lista de Requisições
        For iIndice = 0 To RequisicoesCompra.ListCount - 1

            'Verifica se a Requisição possui pelo menos um item no Grid de Requisições
            For iLinha = 1 To objGridItensRC.iLinhasExistentes
                If RequisicoesCompra.List(iIndice) = GridItensRC.TextMatrix(iLinha, iGrid_ReqCompra_Col) Then
                    Exit For
                End If
            Next

            'Se não possuir, desseleciona a Requisição
            If iLinha > objGridItensRC.iLinhasExistentes Then
                RequisicoesCompra.Selected(iIndice) = False
            End If

        Next
            
    End If
    
    lErro = gobjDistribuicao.Preenche_GridDistribuicaoPC1(gcolItemPedCompraInfo)
    If lErro <> SUCESSO Then gError 89632
    
    Exit Sub

Erro_RequisicoesCompra_ItemCheck:

    Select Case gErr

        Case 65803, 66080, 66101, 66143, 66568, 89632

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156726)

    End Select

    Exit Sub

End Sub

Sub Atualiza_Quantidades(iItem As Integer)
'Atualiza as quantidades do GridItensPC e de GridItens quando
'uma Requisição é marcada ou desmarcada
    
Dim iIndice As Integer
Dim iLinha As Integer
Dim sProduto As String, dFator As Double
Dim dQuantPosterior As Double
Dim dQuantAnterior As Double
Dim lErro As Long, objItemRCInfoAux As ClassItemReqComprasInfo

On Error GoTo Erro_Atualiza_Quantidades

    'Procura os Produtos que entraram o que vão sair do Grid
    For iIndice = 1 To objGridItensRC.iLinhasExistentes
        If GridItensRC.TextMatrix(iIndice, iGrid_ReqCompra_Col) = RequisicoesCompra.List(iItem) Then
            sProduto = GridItensRC.TextMatrix(iIndice, iGrid_ProdutoRC_Col)
            
            dQuantPosterior = 0
            
            'Se uma nova Requisição foi selecinada
            If RequisicoesCompra.Selected(iItem) = True Then
    
                'Acumula a nova quantidade desse produto somando a quantidade que entrou
                For iLinha = 1 To objGridItensRC.iLinhasExistentes
                    If GridItensRC.TextMatrix(iLinha, iGrid_ProdutoRC_Col) = sProduto Then
                            
                        Set objItemRCInfoAux = gcolItemReqCompraInfo(iLinha)
                    
                        dQuantPosterior = dQuantPosterior + (StrParaDbl(GridItensRC.TextMatrix(iLinha, iGrid_RecebidoRC_Col)) / objItemRCInfoAux.dFator)
                    End If
                Next
            
            Else
            
                'Acumula a nova quantidade desse produto diminuindo a quantidade do produto que saiu
                For iLinha = 1 To objGridItensRC.iLinhasExistentes
                    If GridItensRC.TextMatrix(iLinha, iGrid_ProdutoRC_Col) = sProduto And GridItensRC.TextMatrix(iLinha, iGrid_ReqCompra_Col) <> RequisicoesCompra.List(iItem) Then
                        Set objItemRCInfoAux = gcolItemReqCompraInfo(iLinha)
                    
                        dQuantPosterior = dQuantPosterior + (StrParaDbl(GridItensRC.TextMatrix(iLinha, iGrid_RecebidoRC_Col)) / objItemRCInfoAux.dFator)
                    End If
                Next
            
            End If
            
            'Procura esse Produto no GridItens da Nota Fiscal
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If GridItens.TextMatrix(iLinha, iGrid_Produto_Col) = sProduto Then
                    dQuantAnterior = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
                                            
                    'converter a qtdposterior da UM de compras p/a um do item da nf
                    Set objItemRCInfoAux = gcolItemReqCompraInfo(iIndice)
                    lErro = CF("UM_Conversao_Trans", objItemRCInfoAux.iClasseUM, objItemRCInfoAux.sSiglaUMCompra, GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), dFator)
                    If lErro <> SUCESSO Then gError 106866
                    
                    dQuantPosterior = dQuantPosterior * dFator
                    
                    If dQuantPosterior = 0 Then
                        GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = ""
                    Else
                        GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(dQuantPosterior)
                    End If
                    
                    lErro = CF2(Me, "NFEntCom_Calcula_PesoKg", iLinha)
                    If lErro <> SUCESSO Then gError 181293
                    
                    'Recalcula o total por produto
                    Call Calcula_Valores(iLinha)
                                        
                    lErro = gobjTribTab.Alteracao_Item_Grid(iLinha)
                    If lErro <> SUCESSO Then gError 67631
                    
                    Exit For
                End If
            Next
        
            Call Atualiza_QuantidadePCRC(dQuantAnterior / dFator, dQuantPosterior / dFator, iIndice)
        
        End If
    
    Next
            
    'Recalcula o total da Nota Fiscal
    Call Total_Calcula(0)
    
    Exit Sub
    
Erro_Atualiza_Quantidades:
    
    Select Case gErr
    
        Case 67631, 181293
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156727)
        
    End Select
    
    Exit Sub
    
End Sub

Function InsereOrdenadaRC(objItemRCInfo As ClassItemReqComprasInfo) As Long
'Insere ItemRC na coleção de Itens de Requisição de Compra ordenando por lPedCompra e Urgência

Dim iIndice As Integer, bJaExiste As Boolean
Dim lErro As Long, objItemRCInfoAux As ClassItemReqComprasInfo

On Error GoTo Erro_InsereOrdenadaRC

    iIndice = 1
    
    'Se a coleção de ItensRC está vazia
    If gcolItemReqCompraInfo.Count = 0 Then
        'Insere o novo item na coleção
        gcolItemReqCompraInfo.Add objItemRCInfo
    
    'Se não
    Else
    
        bJaExiste = False
        For Each objItemRCInfoAux In gcolItemReqCompraInfo
            If objItemRCInfoAux.lNumIntDoc = objItemRCInfo.lNumIntDoc And objItemRCInfoAux.lNumIntDocItemPC = objItemRCInfo.lNumIntDocItemPC Then
                bJaExiste = True
            End If
        Next
    
        If bJaExiste = False Then
        
            'Para cada Item da coleção de ItensRC
            Do While iIndice <= gcolItemReqCompraInfo.Count
                
                'Se a Requisição da coleção não é urgente, e a que está sendo passada é urgente
                If gcolItemReqCompraInfo(iIndice).iUrgente = 0 And objItemRCInfo.iUrgente = 1 Then
                    'Insere a Requisição urgente antes da que não é urgente
                    gcolItemReqCompraInfo.Add objItemRCInfo, , iIndice
                    Exit Do
                End If
                
                'Se o código do pedido de Compra do ItemRC passado for maior que o do item da coleção
                If objItemRCInfo.lPedCompra >= gcolItemReqCompraInfo(iIndice).lPedCompra Then
                                    
                    'Busca próximo Item
                    iIndice = iIndice + 1
                
                'Se o código do Pedido For menor e a Requisição Não for urgente
                ElseIf objItemRCInfo.iUrgente = 0 And gcolItemReqCompraInfo(iIndice).iUrgente = 1 Then
                    
                    'Busca próximo Item
                    iIndice = iIndice + 1
                                
                Else
                
                    'Insere o itemRC passado antes do Item da coleção que possui maior código de Pedido de Compras
                    gcolItemReqCompraInfo.Add objItemRCInfo, , iIndice
                    Exit Do
                End If
                
            Loop
                        
            'Se o código PC do ItemRC é maior que todos os outros da coleção
            If iIndice > gcolItemReqCompraInfo.Count Then
                'Insere o ItemRC no final da coleção
                gcolItemReqCompraInfo.Add objItemRCInfo
            End If
    
        End If
        
    End If
                
    InsereOrdenadaRC = SUCESSO
    
    Exit Function
    
Erro_InsereOrdenadaRC:
    
    InsereOrdenadaRC = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156728)

    End Select
    
    Exit Function

End Function

Function PreencheGrid_ItensRC()
'Preenche GridItensRC a partir da coleção global de ItensRC

Dim lErro As Long
Dim objItemRCInfo As ClassItemReqComprasInfo
Dim sProdutoMascarado As String
Dim objFilialEmpresa As New AdmFiliais
Dim iIndice As Integer

On Error GoTo Erro_PreencheGrid_ItensRC

    'Limpa o Grid
    Call Grid_Limpa(objGridItensRC)
        
    'Para cada Item da coleção de ItensRC
    For Each objItemRCInfo In gcolItemReqCompraInfo
    
        'Mascara o Produo
        lErro = Mascara_RetornaProdutoEnxuto(objItemRCInfo.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 65810
                
        ProdutoPC.PromptInclude = False
        ProdutoPC.Text = sProdutoMascarado
        ProdutoPC.PromptInclude = True

        sProdutoMascarado = ProdutoPC.Text
                
        'Verifica se o ItemRC está vinculado a um ItemPC cujo Pedido de Compras foi selecionado
        For iIndice = 0 To PedidosCompra.ListCount - 1
            If CLng(PedidosCompra.List(iIndice)) = objItemRCInfo.lPedCompra And PedidosCompra.Selected(iIndice) = True Then
                Exit For
            End If
        Next
        
        'Se encontrou o Pedido de Compras marcado com o mesmo código do ItemRC
        If iIndice <= PedidosCompra.ListCount - 1 Then
        
            'Preenche o Grid com os dados do Item
            GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_PedCompraRC_Col) = CStr(objItemRCInfo.lPedCompra)
            GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_ItemPC_Col) = CStr(objItemRCInfo.iItemPC)
            GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_ProdutoRC_Col) = sProdutoMascarado
            GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_DescProdutoRC_Col) = objItemRCInfo.sDescProduto
            GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_UMRC_Col) = objItemRCInfo.sUM
            GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_QuantRecebidaPC_Col) = Formata_Estoque(objItemRCInfo.dQuantRecebidoPedido)
            
            'FilialEmpresa
            objFilialEmpresa.iCodFilial = objItemRCInfo.iFilialEmpresa
    
            'Lê a FilialEmpresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 27378 Then gError 67248
    
            'Se não encontrou ==>erro
            If lErro = 27378 Then gError 67249
            
            GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_FilialReq_Col) = CStr(objItemRCInfo.iFilialEmpresa & SEPARADOR & objFilialEmpresa.sNome)
            GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_ReqCompra_Col) = CStr(objItemRCInfo.lReqCompra)
            GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_Urgente_Col) = CStr(objItemRCInfo.iUrgente)
            GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_ItemRC_Col) = CStr(objItemRCInfo.iItemRC)
            GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_AReceberRC_Col) = Formata_Estoque(objItemRCInfo.dQuantReceber)
            
            'Se a quantidade Recebida do pedido é menor que a recebida da requisição
            If objItemRCInfo.dQuantRecebidoPedido < objItemRCInfo.dQuantRecebida Then
                GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_RecebidoRC_Col) = Formata_Estoque(objItemRCInfo.dQuantRecebidoPedido)
            Else
                GridItensRC.TextMatrix(objGridItensRC.iLinhasExistentes + 1, iGrid_RecebidoRC_Col) = Formata_Estoque(objItemRCInfo.dQuantRecebida)
            End If
            
            lErro = Grid_Refresh_Checkbox(objGridItensRC)
            If lErro <> SUCESSO Then gError 65804
        
            'Incrementa número de linhas existentes do GridItensRC
            objGridItensRC.iLinhasExistentes = objGridItensRC.iLinhasExistentes + 1
        
        End If
    
    Next
    
    PreencheGrid_ItensRC = SUCESSO
    
    Exit Function
    
Erro_PreencheGrid_ItensRC:
    
    PreencheGrid_ItensRC = gErr

    Select Case gErr
        
        Case 65804, 65810, 67248
        
        Case 67249
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156729)

    End Select
    
    Exit Function

End Function

Private Sub objEventoRecebimento_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoRecebimento_evSelecao

    Set objNFiscal = obj1

    giRecebimento = True

    'Coloca a Nota Fiscal na Tela1
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 65805

    giRecebimento = False
    
    'Fecha o Comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show


    Exit Sub

Erro_objEventoRecebimento_evSelecao:

    giRecebimento = False
    
    Select Case gErr

        Case 65805

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156730)

    End Select

    Exit Sub

End Sub

Public Function Traz_NFiscal_Tela(objNFiscal As ClassNFiscal) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim objNFiscalOriginal As New ClassNFiscal
Dim iIndice As Integer
Dim bCancel As Boolean
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objPedidoCompra As ClassPedidoCompras
Dim colPedidos As New Collection
Dim colPedidosRec As New Collection
Dim objItemNF As ClassItemNF
Dim dPercDesc As Double

On Error GoTo Erro_Traz_NFiscal_Tela

    gbCarregandoTela = True

    'Limpa a tela
    Call Limpa_Tela_NFiscalCom

    If objNFiscal.lNumRecebimento = 0 Then
        NumRecebimento.Text = ""
    Else
        NumRecebimento.Text = objNFiscal.lNumRecebimento
    End If

    'Tenta ler a nota Fiscal passada por parametro
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 65806
    
    'Lê os Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 65806
    
    'Lê as mensagens
    lErro = CF("NFiscal_Le_Msg", objNFiscal)
    If lErro <> SUCESSO Then gError 35705
    
    'distribuicao
    'Lê a Distribuição dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 89544
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Le o Tipo para saber se tem Movimento de Estoque
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO Then gError 83288
    
    If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
        'Carrega ItensNF com Rastreamentos
        lErro = gobjRastreamento.Carrega_RastroItensNF(objNFiscal)
        If lErro <> SUCESSO Then gError 75956
    End If
    
    'Guarda o número interno da Nota Fiscal
    lNumIntDocNF = objNFiscal.lNumIntDoc
    
    Set gobjNFiscal = objNFiscal
    
    lErro = CF("NFiscal_Le_ChvNFe", objNFiscal)
    If lErro <> SUCESSO Then gError 75956
    
    ChvNFe.PromptInclude = False
    ChvNFe.Text = objNFiscal.sChvNFe
    ChvNFe.PromptInclude = True
    gsChvAnt = objNFiscal.sChvNFe
    
    'preenche o Fornecedor
    Fornecedor.Text = objNFiscal.lFornecedor
    Call Fornecedor_Validate(bCancel)

    'Preenche a Filial do Fornecedor
    Filial.Text = objNFiscal.iFilialForn
    Call Filial_Validate(bCancel)

    lErro = CF2(Me, "NFCF_Fornecedor_Traz_NF_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 131938

    If objNFiscal.iTipoNFiscal = DOCINFO_NRFFCO Then
        objNFiscal.iTipoNFiscal = DOCINFO_NFEENCO
    ElseIf objNFiscal.iTipoNFiscal = DOCINFO_NRFPCO Then
        objNFiscal.iTipoNFiscal = DOCINFO_NFIENCO
    End If

    TipoNFiscal.Text = objNFiscal.iTipoNFiscal
    Call TipoNFiscal_Validate(bCancel)

    If objNFiscal.lNumNotaFiscal = 0 Then
        NFiscalInterna.Caption = ""
        NFiscal.Text = ""
    Else
        NFiscalInterna.Caption = CStr(objNFiscal.lNumNotaFiscal)
        NFiscal.Text = CStr(objNFiscal.lNumNotaFiscal)
    End If

    If objNFiscal.sNaturezaOp <> "" Then NaturezaOp.Text = objNFiscal.sNaturezaOp

    If Len(Trim(objNFiscal.sSerie)) > 0 Then
        Serie.Text = Desconverte_Serie_Eletronica(objNFiscal.sSerie)
        If ISSerieEletronica(objNFiscal.sSerie) Then
            EletronicaFed.Value = vbChecked
        Else
            EletronicaFed.Value = vbUnchecked
        End If
    Else
        Serie.Text = ""
    End If
    Call Serie_Validate(bSGECancelDummy)
    
    'Se não trouxer para tela ao regravar vai excluir a NF anterior e apagar
    'qualquer rastro de uso do recebimento deixando-o perdido conforme ocorreu na DLE
    If objNFiscal.lNumRecebimento <> 0 Then NumRecebimento.Text = objNFiscal.lNumRecebimento

    'Se a data não for nula coloca na Tela
    DataEmissao.PromptInclude = False
    If objNFiscal.dtDataEmissao <> DATA_NULA Then DataEmissao.Text = Format(objNFiscal.dtDataEmissao, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    DataEntrada.PromptInclude = False
    If objNFiscal.dtDataEntrada <> DATA_NULA Then DataEntrada.Text = Format(objNFiscal.dtDataEntrada, "dd/mm/yy")
    DataEntrada.PromptInclude = True

'horaentrada
    HoraEntrada.PromptInclude = False
    'este teste está correto
    If objNFiscal.dtDataEntrada <> DATA_NULA Then HoraEntrada.Text = Format(objNFiscal.dtHoraEntrada, "hh:mm:ss")
    HoraEntrada.PromptInclude = True

    DataVencimento.PromptInclude = False
    If objNFiscal.dtDataVencimento <> DATA_NULA Then DataVencimento.Text = Format(objNFiscal.dtDataVencimento, "dd/mm/yy")
    DataVencimento.PromptInclude = True

    If objNFiscal.dValorDesconto > 0 Then ValorDesconto.Text = Format(objNFiscal.dValorDesconto, "Standard")
    If objNFiscal.dValorOutrasDespesas > 0 Then ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Standard")
    If objNFiscal.dValorFrete > 0 Then ValorFrete.Text = Format(objNFiscal.dValorFrete, "Standard")
    If objNFiscal.dValorSeguro > 0 Then ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Standard")
    If objNFiscal.dValorTotal > 0 Then Total.Text = Format(objNFiscal.dValorTotal, "Standard")
    If objNFiscal.dValorProdutos > 0 Then SubTotal.Caption = Format(objNFiscal.dValorProdutos, "Standard")

    For iIndice = 0 To TipoFrete.ListCount - 1
        If TipoFrete.ItemData(iIndice) = objNFiscal.iFreteRespons Then
            TipoFrete.ListIndex = iIndice
            Exit For
        End If
    Next

    If objNFiscal.iCodTransportadora <> 0 Then
        Transportadora.Text = objNFiscal.iCodTransportadora
        Call Transportadora_Validate(bCancel)
    End If
    
    If objNFiscal.iTipoNFiscal = DOCINFO_NFEEBFCOM Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEBFCOM Or objNFiscal.iTipoNFiscal = DOCINFO_NFECCOPC Then

        'Preenche o Fornecedor Benef
        If objNFiscal.lFornecedorBenef <> 0 Then
            FornecedorBenef.Text = objNFiscal.lFornecedorBenef
        Else
            FornecedorBenef.Text = ""
        End If
        Call FornecedorBenef_Validate(bCancel)
        
        'Preenche a Filial do Fornecedor
        If objNFiscal.iFilialFornBenef <> 0 Then
            FilialFornBenef.Text = objNFiscal.iFilialFornBenef
        Else
            FilialFornBenef.Text = ""
        End If
        Call FilialFornBenef_Validate(bCancel)
        
    End If

    Placa.Text = objNFiscal.sPlaca
    PlacaUF.Text = objNFiscal.sPlacaUF
    
    'Traz a espécie dos volumes do pedido
    If objNFiscal.lVolumeEspecie > 0 Then
        VolumeEspecie.Text = objNFiscal.lVolumeEspecie
        Call VolumeEspecie_Validate(bSGECancelDummy)
    Else
        VolumeEspecie.Text = ""
    End If
    
    'Traz a marca dos volumes do pedido
    If objNFiscal.lVolumeMarca > 0 Then
        VolumeMarca.Text = objNFiscal.lVolumeMarca
        Call VolumeMarca_Validate(bSGECancelDummy)
    Else
        VolumeMarca.Text = ""
    End If
    
    If objNFiscal.lVolumeQuant > 0 Then VolumeQuant.Text = objNFiscal.lVolumeQuant

    Mensagem.Text = objNFiscal.sMensagemNota
    MensagemCorpo.Text = objNFiscal.sMensagemCorpoNota
    Observacao.Text = objNFiscal.sObservacao

    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = CStr(objNFiscal.dPesoLiq)
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = CStr(objNFiscal.dPesoBruto)

    'Se o código da Nota Fiscal Original estiver preenchido
    If objNFiscal.lNumIntNotaOriginal <> 0 Then

        objNFiscalOriginal.lNumIntDoc = objNFiscal.lNumIntNotaOriginal

        'Lê a Nota Fiscal original
        lErro = CF("NFiscal_Le", objNFiscalOriginal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 65808

        If lErro = 35132 Then gError 65809

        'Coloca os dados relativos  Nota Fiscal Original na Tela
        SerieNFiscalOriginal.Text = objNFiscalOriginal.sSerie
        NFiscalOriginal.Text = objNFiscalOriginal.lNumNotaFiscal

    End If

    FilialCompra.Text = objNFiscal.iFilialPedido
    Call FilialCompra_Validate(bCancel)
        
    Set gcolPedidoCompra = New Collection

    Call Grid_Limpa(objGridItens)
    
    'Preenche o Grid com os Ítens da Nota Fiscal
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 65807

    'distribuicao
    'Preenche o Grid com as Distribuições dos itens da Nota Fiscal
    lErro = gobjDistribuicao.Preenche_GridDistribuicao(objNFiscal)
    If lErro <> SUCESSO Then gError 89545

    Set gcolItemPedCompraInfo = New Collection
    
    'Se estiver trazendo um Recebimento para a tela
    If giRecebimento = True Then
        
        'Le os pedidos de compra enviados para o fornecedor em questão com quantidade a receber
        lErro = CF("PedidoCompra_Le_EnvComQuantReceber", colPedidos, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.iFilialPedido)
        If lErro <> SUCESSO Then gError 66131
        
        'Lê os Pedidos de Compra relacionados ao Recebimento
        lErro = CF("PedidoCompra_Le_Recebimento", objNFiscal, colPedidosRec)
        If lErro <> SUCESSO Then gError 67000
                
        lErro = CF("Aglutina_Pedidos", gcolPedidoCompra, colPedidos, colPedidosRec)
        If lErro <> SUCESSO Then gError 89272
        
        'Se é um Recebimento, habilita a lista de Pedidos de Compras e a lista de Requisições
        PedidosCompra.Enabled = True
        RequisicoesCompra.Enabled = True
                
        'Adiciona na lista de Pedidos de Compras todos os Pedidos mesmo Fornecedor, Filial e FilialPedido da NFiscal
        For Each objPedidoCompra In gcolPedidoCompra
            
            PedidosCompra.AddItem objPedidoCompra.lCodigo
    
            'Lê os itens do Pedido de Compra passado por objPedidoCompra
            lErro = CF("ItensPCTodos_Le_Codigo", objPedidoCompra)
            If lErro <> SUCESSO Then gError 66116
        
        Next
        
        'Marca os Pedidos de Compras associados ao Recebimento
        For Each objPedidoCompra In colPedidosRec
            For iIndice = 0 To PedidosCompra.ListCount - 1
                If objPedidoCompra.lCodigo = CLng(PedidosCompra.List(iIndice)) Then
                    PedidosCompra.Selected(iIndice) = True
                End If
            Next
        Next
        
        'Lê apenas os Itens de Pedido de Compras não baixados relacionados a Nota Fiscal
        lErro = CF("ItensPedCompra_Le_NFiscalReceb", objNFiscal, gcolItemPedCompraInfo)
        If lErro <> SUCESSO Then gError 66132
        
    Else
    
        'Lê todos os Pedidos de Compras relacionados a nota fiscal
        lErro = CF("PedidoCompraTodos_Le_NFiscal", objNFiscal, gcolPedidoCompra)
        If lErro <> SUCESSO Then gError 65811

        PedidosCompra.ListIndex = -1
        RequisicoesCompra.ListIndex = -1
    
        iIndice = 0
            
        'Seleciona todos Pedidos de Compra relacionados a Nota Fiscal
        For Each objPedidoCompra In gcolPedidoCompra
            
            PedidosCompra.AddItem objPedidoCompra.lCodigo
            PedidosCompra.Selected(iIndice) = True
            iIndice = iIndice + 1
    
            'Lê os itens do Pedido de Compra passado por objPedidoCompra
            lErro = CF("ItensPCTodos_Le_Codigo", objPedidoCompra)
            If lErro <> SUCESSO Then gError 66116
        
        Next
    
        'Lê todos os Itens de Pedidos de Compras relacionados a Nota Fiscal
        lErro = CF("ItensPedCompraTodos_Le_NFiscal", objNFiscal, gcolItemPedCompraInfo)
        If lErro <> SUCESSO Then gError 65812
    
    End If
    
    lErro = CF2(Me, "NFEntCom_Le_PedInfo_Customizado", gcolItemPedCompraInfo)
    If lErro <> SUCESSO Then gError 181293
        
    Call Grid_Limpa(objGridItensPC)
                
    'Preenche Grid de Itens de Pedido de Compras
    lErro = Preenche_GridItensPC()
    If lErro <> SUCESSO Then gError 65813

    If Not giRecebimento Then

        'Traz os dados contábeis para a tela (contabilidade)
        lErro = objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
        If lErro <> SUCESSO And lErro <> 36326 Then gError 65814

        'Preenche o Status da Nota Fiscal
        If objNFiscal.iStatus = STATUS_LANCADO Then
            Status.Caption = STRING_STATUS_LANCADO
        ElseIf objNFiscal.iStatus = STATUS_BAIXADO Then
            Status.Caption = STRING_STATUS_BAIXADO
        ElseIf objNFiscal.iStatus = STATUS_CANCELADO Then
            Status.Caption = STRING_STATUS_CANCELADO
        End If

        'Lê a parte de Tributação
        lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
        If lErro <> SUCESSO Then gError 65815

        'Lê a parte de tributacao dos Ítens da Nota Fiscal
        lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
        If lErro <> SUCESSO Then gError 65816

    ''*** incluidos p/tratamento de tributacao *******************************
        lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 65817
    '*** fim tributacao

    Else

        For iIndice = 1 To objNFiscal.ColItensNF.Count
            lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, objNFiscal.ColItensNF.Item(iIndice).sProduto)
            If lErro <> SUCESSO Then gError 65818
        Next

        Call gobjTribTab.TributacaoRecalcular_Click
        
        Call objContabil.Contabil_Carga_Modelo_Padrao
        
    End If

    lErro = CF2(Me, "NFEnt_Traz_NFiscal_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 124252

    lErro = CF2(Me, "NFEntCom_Traz_NFiscal_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 140323

    NumNFe.Caption = objNFiscal.sStatusNFeFed
    
    MsgAutomatica.Value = vbUnchecked
    
    If objNFiscal.dValorItens = 0 Then
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If objNFiscal.ColItensNF.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
                objNFiscal.dValorItens = objNFiscal.dValorItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
                objNFiscal.dValorDescontoItens = objNFiscal.dValorDescontoItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
            End If
        Next
    End If
    
    ValorProdutos2.Caption = Format(objNFiscal.dValorItens, "Standard")
    ValorDescontoItens.Text = Format(objNFiscal.dValorDescontoItens, "Standard")
    If objNFiscal.dValorItens > 0 Then
        dPercDesc = objNFiscal.dValorDescontoItens / objNFiscal.dValorItens
    Else
        dPercDesc = 0
    End If
    PercDescontoItens.Text = Format(dPercDesc * 100, "Fixed")
    
    dValorDescontoItensAnt = objNFiscal.dValorDescontoItens
    dPercDescontoItensAnt = dPercDesc
     
    Set gobjInfoAdic = objNFiscal.objInfoAdic

    iAlterado = 0

    gbCarregandoTela = False

    Traz_NFiscal_Tela = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela:

    gbCarregandoTela = False

    Traz_NFiscal_Tela = gErr

    Select Case gErr

        Case 65806, 65807, 65808, 65811, 65812, 65813, 65814, 65815, 65816, 65817, 65818, 66116, 66131, 66132, 67000, 75956, 83288, 89272, 89544, 89545, 124252, 131938, 140323, 181293 'Inserido por Wagner

        Case 65809
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntNotaOriginal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156731)

    End Select

    Exit Function

End Function

Public Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO

    'Verifica se alguma filial foi selecionada
    If Filial.ListIndex = -1 Then Exit Sub
    
    If Len(Trim(Filial.Text)) > 0 Then
    
        If Len(Trim(TipoNFiscal.Text)) > 0 And (sFornecedorAnterior <> Trim(Fornecedor.Text) Or iFilialAnterior <> StrParaInt(Codigo_Extrai(Filial.Text)) _
            Or iFilialCompraAnterior <> Codigo_Extrai(FilialCompra.Text)) Then
    
            Call Atualiza_ListaPedidos
    
        End If
    
    End If
    
    Call Trata_FilialForn
    
End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then
        Call Limpa_Tela_PCRC
        Exit Sub
    End If
    
    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 65819

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 65823

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 65820

        If lErro = 18272 Then gError 65821

        'coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

        Call gobjTribTab.FilialFornecedor_Alterada(objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)
    
    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 65822

    If Len(Trim(TipoNFiscal.Text)) > 0 And (sFornecedorAnterior <> Trim(Fornecedor.Text) Or iFilialAnterior <> StrParaInt(Codigo_Extrai(Filial.Text)) _
        Or iFilialCompraAnterior <> Codigo_Extrai(FilialCompra.Text)) Then

        Call Atualiza_ListaPedidos

    End If

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 65819, 65820

        Case 65821
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If
        
        Case 65822
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, Filial.Text)

        Case 65823
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156732)

    End Select

    Exit Sub

End Sub

Public Function Preenche_GridItens(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com os ítens da Nota Fiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim sProdutoMascarado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sCclMascarado As String
Dim iIndice2 As Integer

On Error GoTo Erro_Preenche_GridItens

    'Para cada ítem da Coleção
    For Each objItemNF In objNFiscal.ColItensNF

        iIndice = iIndice + 1

        'Mascara o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 65824
        
        ProdutoPC.PromptInclude = False
        ProdutoPC.Text = sProdutoMascarado
        ProdutoPC.PromptInclude = True

        sProdutoMascarado = ProdutoPC.Text
        
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado
        For iIndice2 = 0 To Produto.ListCount - 1
            If Produto.List(iIndice) = sProdutoMascarado Then
                Produto.ListIndex = iIndice
                Exit For
            End If
        Next
        
        If iIndice2 = Produto.ListCount Then
            Produto.AddItem sProdutoMascarado
            Produto.ListIndex = iIndice2
        End If
                
        sCclMascarado = ""
        
        'Formata Ccl
        If Trim(objItemNF.sCcl) <> "" Then

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_RetornaCclEnxuta(objItemNF.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 65825
    
            'Preenche o campo Ccl
            Ccl.PromptInclude = False
            Ccl.Text = sCclMascarado
            Ccl.PromptInclude = True
            
            'Joga o Ccl no grid
            GridItens.TextMatrix(iIndice, iGrid_Ccl_Col) = Ccl.Text

        End If

        'Preenche o Grid
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItemNF.sDescricaoItem
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemNF.sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemNF.dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(objItemNF.dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
        
        If objItemNF.dPercDesc <> 0 Then GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(objItemNF.dPercDesc, "Percent")
        If objItemNF.dValorDesconto <> 0 Then GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objItemNF.dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objItemNF.dValorTotal, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col) = Format(objItemNF.dValorTotal + objItemNF.dValorDesconto, "Standard")
    
    Next

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice

    lErro = CF2(Me, "NFEntCom_Preenche_GridItens", objNFiscal)
    If lErro <> SUCESSO Then gError 65826
    
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objNFiscal.objProjetoInfo, objNFiscal.lNumIntDoc, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181676
    
    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr
        
        Case 65824
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)
        
        Case 65825
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, objItemNF.sCcl)

        Case 65826, 181676
        
        Case 65827
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156733)

    End Select

    Exit Function

End Function

Public Sub Limpa_Tela_NFiscalCom()
'Limpa a Tela NFiscalEntradaCom
    
Dim objConfiguraCOM As New ClassConfiguraCOM
Dim lErro As Long
Dim bCancel As Boolean

On Error GoTo Erro_Limpa_Tela_NFiscalCom

    'Chama o Limpa tela
    Call Limpa_Tela(Me)
    
    Set gobjNFiscal = New ClassNFiscal

    Set gobjAnotacao = Nothing
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    Set gobjInfoAdic = Nothing
    
    'Limpa o Grid
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridItensPC)
    Call Grid_Limpa(objGridItensRC)
    
    ComboPedidoCompras.Clear
    ComboPedidoCompras.AddItem "TODOS"
    
    Moeda.ListIndex = -1
    ComboPedidoCompras.ListIndex = 0
    
    Taxa.Text = ""
    Taxa.Enabled = False
    LabelTaxa.Enabled = False
    
    PedidosCompra.Clear
    RequisicoesCompra.Clear
    PedidosCompra.Enabled = True
    RequisicoesCompra.Enabled = True
   
    Produto.Clear
    
    Serie.Text = ""
    Call Serie_Validate(bSGECancelDummy)
    'EletronicaFed.Value = vbUnchecked
    If gobjEST.iSerieEletronicaPadrao = MARCADO Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    
    'Limpa os campos que não são limpos pelo Limpa_Tela
    NFiscal.Text = ""
    NFiscalInterna.Caption = ""
    Status.Caption = ""
    Filial.Clear
    SerieNFiscalOriginal.Text = ""
    SubTotal.Caption = ""
    ValorProdutos2.Caption = ""
    Transportadora.Text = ""
    PlacaUF.Text = ""
    TipoFrete.ListIndex = 0
        
    VolumeMarca.Text = ""
    VolumeEspecie.Text = ""
    
    'Recarrega a combo VolumeEspecie e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)

    'Recarrega a combo VolumeMarca e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
        
    Call gobjTribTab.DataEmissao_Alterada(DATA_NULA)
    
    'Preenche a Data de Entrada com a Data Atual
    DataEntrada.PromptInclude = False
    DataEntrada.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEntrada.PromptInclude = True
    Call DataEntrada_Validate(bSGECancelDummy)

''*** incluidos p/tratamento de tributacao *******************************

    Call gobjTribTab.Limpa_Tela

'*** fim tributacao
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    Set gcolItemPedCompraInfo = New Collection
    Set gcolItemReqCompraInfo = New Collection
    Set gcolPedidoCompra = New Collection
    
'    'Rastreamento
'    Set gcolItensNF = New Collection

    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento

    'distribuicao
    Call gobjDistribuicao.Limpa_Tela_Distribuicao

    'Leitura da tabela de ComprasConfig
    lErro = CF("ComprasConfig_Le", objConfiguraCOM)
    If lErro <> SUCESSO Then gError 65622
    
    'Coloca FilialCompra Default na tela
    If objConfiguraCOM.iFilialCompra > 0 Then
        FilialCompra.Text = objConfiguraCOM.iFilialCompra
    Else
        FilialCompra.Text = giFilialEmpresa
    End If

    Call FilialCompra_Validate(bCancel)

    Call CF2(Me, "NFCF_Limpa_Tela")
    
    MsgAutomatica.Value = vbChecked

    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0

    Exit Sub
    
Erro_Limpa_Tela_NFiscalCom:

    Select Case gErr
    
        Case 65622
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156734)

    End Select
    
    Exit Sub
    
End Sub

Public Sub Limpa_Tela_NFiscalCom1()
'Limpa a Tela NFiscalEntrada, mas mantém a natureza e o tipo
    
Dim sNatureza As String
Dim sTipoNFiscal As String

    sNatureza = NaturezaOp.Text
    sTipoNFiscal = TipoNFiscal.Text
    
    Call Limpa_Tela_NFiscalCom

    NaturezaOp.Text = sNatureza
    TipoNFiscal.Text = sTipoNFiscal

End Sub

Public Sub Limpa_Tela_PCRC()
'Limpa da tela os Pedidos de Compras e as Requisições de Compras

    gsChvAnt = ""

    'Limpa os Grids
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridItensPC)
    Call Grid_Limpa(objGridItensRC)

    'Limpa a Lista de Pedidos de Compra
    PedidosCompra.Clear

    'Limpa as coleções
    Set gcolPedidoCompra = New Collection
    Set gcolItemPedCompraInfo = New Collection
    Set gcolItemReqCompraInfo = New Collection
    
    'Limpa a Lista de Requisições
    RequisicoesCompra.Clear

    'Limpa tela de Tributação
    gobjTribTab.Limpa_Tela
''*** incluidos p/tratamento de tributacao *******************************
    If IsDate(DataEntrada.Text) Then Call gobjTribTab.DataEntrada_Alterada(StrParaDate(DataEntrada.Text))
    Call gobjTribTab.Serie_Alterada(Serie.Text)
    Call Trata_FilialForn(True)
'*** fim tributacao
    
    PedidosCompra.Enabled = True
    RequisicoesCompra.Enabled = True
    
    ComboPedidoCompras.ListIndex = 0
    
End Sub

Public Sub DataEntrada_Change()

    iAlterado = REGISTRO_ALTERADO
    iDataEntradaAlterada = REGISTRO_ALTERADO

End Sub

Public Sub DataEntrada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntrada_Validate

    If iDataEntradaAlterada <> 0 Then
        
        'Verifica se a data de Entrada foi digitada
        If Len(Trim(DataEntrada.ClipText)) = 0 Then Exit Sub
    
        'Critica a data digitada
        lErro = Data_Critica(DataEntrada.Text)
        If lErro <> SUCESSO Then gError 65828

''*** incluidos p/tratamento de tributacao *******************************
        Call gobjTribTab.DataEntrada_Alterada(StrParaDate(DataEntrada.Text))
'*** fim tributacao
        
        iDataEntradaAlterada = 0

    End If
    
    Exit Sub

Erro_DataEntrada_Validate:

    Cancel = True

    Select Case gErr

        Case 65828

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156735)

    End Select

    Exit Sub

End Sub

'horaentrada
Public Sub HoraEntrada_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(HoraEntrada, iAlterado)

End Sub

'horaentrada
Public Sub HoraEntrada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'horaentrada
Public Sub HoraEntrada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HoraEntrada_Validate

    'Verifica se a hora de Entrada foi digitada
    If Len(Trim(HoraEntrada.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(HoraEntrada.Text)
    If lErro <> SUCESSO Then gError 89780

    Exit Sub

Erro_HoraEntrada_Validate:

    Cancel = True

    Select Case gErr

        Case 89780

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156736)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a data de emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) <> 0 Then

        'Critica a data digitada
        lErro = Data_Critica(DataEmissao.Text)
        If lErro <> SUCESSO Then gError 65829
    
        Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
        
    End If
        
    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 65829

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156737)

    End Select

    Exit Sub

End Sub

Public Sub DataVencimento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataVencimento_Validate(Cancel As Boolean)

Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim lErro As Long

On Error GoTo Erro_DataVencimento_Validate

    'Verifica se a data de vencimento foi digitada
    If Len(Trim(DataVencimento.ClipText)) = 0 Then Exit Sub

    'Faz a critica da data digitada
    lErro = Data_Critica(DataVencimento.Text)
    If lErro <> SUCESSO Then gError 65830

    'Verifica se a data de emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    dtDataEmissao = CDate(DataEmissao.Text)
    dtDataVencimento = CDate(DataVencimento.Text)

    'Verifica se a data de vencimento é menor que a data de emissão
    If dtDataVencimento < dtDataEmissao Then gError 65831

    Exit Sub

Erro_DataVencimento_Validate:

    Cancel = True


    Select Case gErr

        Case 65830

        Case 65831
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_MENOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156738)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntrada_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntrada_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEntrada, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 65832

    Exit Sub

Erro_UpDownEntrada_DownClick:

    Select Case gErr

        Case 65832

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156739)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntrada_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntrada_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEntrada, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 65833

    Exit Sub

Erro_UpDownEntrada_UpClick:

    Select Case gErr

        Case 65833

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156740)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 65834

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 65834

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156741)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 65835

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 65835

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156742)

    End Select

    Exit Sub

End Sub

Public Sub ValorDesconto_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorDescontoAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorDespesas_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorDespesasAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorFrete_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorFreteAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorFrete_Validate(Cancel As Boolean)

Dim dValorFrete As Double
Dim lErro As Long

On Error GoTo Erro_ValorFrete_Validate

    If iValorFreteAlterado = 0 Then Exit Sub

    'Verifica se o Valor Frete foi preenchido
    If Len(Trim(ValorFrete.Text)) > 0 Then
    
        'faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
        If lErro <> SUCESSO Then gError 65836

        dValorFrete = CDbl(ValorFrete.Text)
        
        'Coloca o Valor Formatado na tela
        ValorFrete.Text = Format(dValorFrete, "Standard")

    End If

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorFrete_Validate(Cancel, dValorFrete)
'*** fim tributacao

    Call ValorTotal_Calcula

    iValorFreteAlterado = 0

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True


    Select Case gErr

        Case 65836

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156743)

    End Select

    Exit Sub

End Sub

Public Sub ValorSeguro_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorSeguroAlterado = REGISTRO_ALTERADO

End Sub

Function Preenche_GridItensPC() As Long
'Preenche GridItensPC

Dim objItemPedCompraInfo As ClassItemPedCompraInfo
Dim iLinha As Integer
Dim iItem As Integer
Dim lErro As Long
Dim lPedCompra As Long
Dim sProdutoMascarado As String
Dim iIndice As Integer
Dim bTodos As Boolean
Dim objPedidoCompras As ClassPedidoCompras
Dim dFrete As Double
Dim dSeguro As Double
Dim dDesconto As Double
Dim dTotal As Double
Dim dDespesa As Double
Dim dAliquotaIPI As Double
Dim bPedidoMarcado As Boolean

On Error GoTo Erro_Preenche_GridItensPC

    Call Grid_Limpa(objGridItensPC)
    
    Set objGridItensPC = New AdmGrid
    
    Call Inicializa_Grid_ItensPC(objGridItensPC)
    
    If ComboPedidoCompras.List(ComboPedidoCompras.ListIndex) = "TODOS" Then bTodos = True
    
    'Para cada Item da coleção de ItensPC
    For Each objItemPedCompraInfo In gcolItemPedCompraInfo
        
        If Not bTodos Then
        
            If objItemPedCompraInfo.lPedCompra = ComboPedidoCompras.List(ComboPedidoCompras.ListIndex) Then
    
                'Incrementa o número de linhas
                iLinha = iLinha + 1
                iItem = iItem + 1
                
                'Se o Pedido de Compras mudou, Inicia novamente o contador de Itens
                If objItemPedCompraInfo.lPedCompra <> lPedCompra And iLinha > 1 Then iItem = 1
                
                GridItensPC.TextMatrix(iLinha, iGrid_PedCompra_Col) = CStr(objItemPedCompraInfo.lPedCompra)
                GridItensPC.TextMatrix(iLinha, iGrid_Item_Col) = iItem
                
                'Mascara o Produto
                lErro = Mascara_RetornaProdutoEnxuto(objItemPedCompraInfo.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 54675
                
                ProdutoPC.PromptInclude = False
                ProdutoPC.Text = sProdutoMascarado
                ProdutoPC.PromptInclude = True
                
                sProdutoMascarado = ProdutoPC.Text
                
                GridItensPC.TextMatrix(iLinha, iGrid_Prod_Col) = sProdutoMascarado
                GridItensPC.TextMatrix(iLinha, iGrid_DescProdutoPC_Col) = objItemPedCompraInfo.sDescProduto
                GridItensPC.TextMatrix(iLinha, iGrid_UM_Col) = objItemPedCompraInfo.sUM
                GridItensPC.TextMatrix(iLinha, iGrid_AReceber_Col) = Formata_Estoque(objItemPedCompraInfo.dQuantReceber)
                GridItensPC.TextMatrix(iLinha, iGrid_Recebido_Col) = Formata_Estoque(objItemPedCompraInfo.dQuantRecebida)
                GridItensPC.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = Format(objItemPedCompraInfo.dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
                
                lErro = CF2(Me, "NFEntCom_PreencheLinha_GridPC", objItemPedCompraInfo.objInfoUsu, iLinha)
                If lErro <> SUCESSO Then gError 181287
                
                If gbLimpaTaxa Then Taxa.Text = ""
                
                If (objItemPedCompraInfo.dTaxa = 0 And StrParaDbl(Taxa.Text) > 0) Or (objItemPedCompraInfo.bTaxaPedido = True And Not gbLimpaTaxa) Then
                    objItemPedCompraInfo.dTaxa = StrParaDbl(Taxa.Text)
                    objItemPedCompraInfo.bTaxaPedido = True
                End If
                
                If ComboPedidoCompras.Text <> "TODOS" And objItemPedCompraInfo.dTaxa > 0 Then
                    Taxa.Text = Format(objItemPedCompraInfo.dTaxa, FORMATO_TAXA_CONVERSAO_MOEDA)
                End If
                
                If StrParaDbl(Taxa.Text) > 0 Then GridItensPC.TextMatrix(iLinha, iGrid_Recebido_RS_Col) = Format(objItemPedCompraInfo.dPrecoUnitario * objItemPedCompraInfo.dTaxa, "STANDARD")
                
                'Guarda o último Pedido de Compras
                lPedCompra = objItemPedCompraInfo.lPedCompra
                
                'Seleciona a moeda e a taxa (caso exista ...)
                For iIndice = 0 To Moeda.ListCount - 1
                    If Moeda.ItemData(iIndice) = objItemPedCompraInfo.iMoeda Then
                        Moeda.ListIndex = iIndice
                        Exit For
                    End If
                Next
                    
                'Se a taxa foi preenchida => Mostra na tela ...
                If objItemPedCompraInfo.dTaxa > 0 Then
                    
                    Taxa.Text = Format(objItemPedCompraInfo.dTaxa, FORMATO_TAXA_CONVERSAO_MOEDA)
                    
                    If objItemPedCompraInfo.bTaxaPedido Then
                        Taxa.Enabled = True
                        LabelTaxa.Enabled = True
                    Else
                        Taxa.Enabled = False
                        LabelTaxa.Enabled = False
                    End If
                
                'Apenas por Seguranca ...
                ElseIf objItemPedCompraInfo.dTaxa = 0 Then
                    
                    Taxa.Text = ""
                    Taxa.Enabled = True
                    LabelTaxa.Enabled = True
                    
                End If
                
                If objItemPedCompraInfo.iMoeda = MOEDA_REAL Then
                    Taxa.Text = ""
                    Taxa.Enabled = False
                    LabelTaxa.Enabled = False
                End If
                
            End If
                
        Else
            
            'Incrementa o número de linhas
            iLinha = iLinha + 1
            iItem = iItem + 1
            
            'Se o Pedido de Compras mudou, Inicia novamente o contador de Itens
            If objItemPedCompraInfo.lPedCompra <> lPedCompra And iLinha > 1 Then iItem = 1
            
            GridItensPC.TextMatrix(iLinha, iGrid_PedCompra_Col) = CStr(objItemPedCompraInfo.lPedCompra)
            GridItensPC.TextMatrix(iLinha, iGrid_Item_Col) = iItem
            
            'Mascara o Produto
            lErro = Mascara_RetornaProdutoEnxuto(objItemPedCompraInfo.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 54675
            
            ProdutoPC.PromptInclude = False
            ProdutoPC.Text = sProdutoMascarado
            ProdutoPC.PromptInclude = True
            
            sProdutoMascarado = ProdutoPC.Text
            
            GridItensPC.TextMatrix(iLinha, iGrid_Prod_Col) = sProdutoMascarado
            GridItensPC.TextMatrix(iLinha, iGrid_DescProdutoPC_Col) = objItemPedCompraInfo.sDescProduto
            GridItensPC.TextMatrix(iLinha, iGrid_UM_Col) = objItemPedCompraInfo.sUM
            GridItensPC.TextMatrix(iLinha, iGrid_AReceber_Col) = Formata_Estoque(objItemPedCompraInfo.dQuantReceber)
            GridItensPC.TextMatrix(iLinha, iGrid_Recebido_Col) = Formata_Estoque(objItemPedCompraInfo.dQuantRecebida)
            GridItensPC.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = Format(objItemPedCompraInfo.dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
            
            lErro = CF2(Me, "NFEntCom_PreencheLinha_GridPC", objItemPedCompraInfo.objInfoUsu, iLinha)
            If lErro <> SUCESSO Then gError 181287
            
            For iIndice = 0 To Moeda.ListCount - 1
                If Moeda.ItemData(iIndice) = objItemPedCompraInfo.iMoeda Then
                    GridItensPC.TextMatrix(iLinha, iGrid_Moeda_Col) = Moeda.List(iIndice)
                    Exit For
                End If
            Next
            
            If objItemPedCompraInfo.dTaxa > 0 Then GridItensPC.TextMatrix(iLinha, iGrid_Taxa_Col) = Format(objItemPedCompraInfo.dTaxa, FORMATO_TAXA_CONVERSAO_MOEDA)
            
            Taxa.Text = ""
            Taxa.Enabled = False
            LabelTaxa.Enabled = False
                
            If objItemPedCompraInfo.dTaxa = 0 And StrParaDbl(Taxa.Text) > 0 Then
                objItemPedCompraInfo.dTaxa = StrParaDbl(Taxa.Text)
                objItemPedCompraInfo.bTaxaPedido = True
            End If
            
            If ComboPedidoCompras.Text <> "TODOS" Then
                Taxa.Text = Format(objItemPedCompraInfo.dTaxa, FORMATO_TAXA_CONVERSAO_MOEDA)
            End If
            
            If objItemPedCompraInfo.dTaxa > 0 Then GridItensPC.TextMatrix(iLinha, iGrid_Recebido_RS_Col) = Format(objItemPedCompraInfo.dPrecoUnitario * objItemPedCompraInfo.dTaxa, "STANDARD")
            
            'Guarda o último Pedido de Compras
            lPedCompra = objItemPedCompraInfo.lPedCompra
            
        End If
        
    Next

    If Not gbCarregandoTela Then

        For Each objPedidoCompras In gcolPedidoCompra
        
            bPedidoMarcado = False
        
            'Para cada Item da coleção de ItensPC
            For Each objItemPedCompraInfo In gcolItemPedCompraInfo
        
                If objPedidoCompras.lCodigo = objItemPedCompraInfo.lPedCompra Then
        
                    bPedidoMarcado = True
                    Exit For
                End If
        
            Next
            
            If bPedidoMarcado Then
                dFrete = dFrete + objPedidoCompras.dValorFrete
                dSeguro = dSeguro + objPedidoCompras.dValorSeguro
                dDesconto = dDesconto + objPedidoCompras.dValorDesconto
                dDespesa = dDespesa + objPedidoCompras.dOutrasDespesas
            End If
        
        Next
        
        ValorFrete.Text = dFrete
        Call ValorFrete_Validate(bSGECancelDummy)
        
        ValorSeguro.Text = dSeguro
        Call ValorSeguro_Validate(bSGECancelDummy)
        
        ValorDesconto.Text = dDesconto
        Call ValorDesconto_Validate(bSGECancelDummy)
        
        ValorDespesas.Text = dDespesa
        Call ValorDespesas_Validate(bSGECancelDummy)
        
    End If
    
    'Alteração das linhas existentes
    objGridItensPC.iLinhasExistentes = gcolItemPedCompraInfo.Count
    
    If Not gbCarregandoTela Then

        lErro = CalculaPrecoUnitario_GridItens
        If lErro <> SUCESSO Then gError 114511
        
    End If

    Preenche_GridItensPC = SUCESSO

    Exit Function

Erro_Preenche_GridItensPC:

    Preenche_GridItensPC = gErr

    Select Case gErr

        Case 54675, 181287, 114511
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156744)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iIndice As Integer
Dim dTotal As Double
Dim objNFiscal As New ClassNFiscal
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim dQuantidade As Double, dQuantidadeComMargem As Double
Dim iLinha As Integer
Dim sCodProduto As String
Dim colPedCompras As New Collection
Dim colReqCompras As New Collection
Dim dQuantLote As Double
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim dFator As Double
Dim objItemNF As ClassItemNF
Dim vbMsg As VbMsgBoxResult
Dim bTeste As Boolean
Dim objNFiscalBD As New ClassNFiscal

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 65863
    If Len(Trim(NaturezaOp.Text)) = 0 Then gError 65864
    If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 65865
    If Len(Trim(Filial.Text)) = 0 Then gError 65866
    If Len(Trim(Serie.Text)) = 0 Then gError 65867
    If Len(Trim(FilialCompra.Text)) = 0 Then gError 65868
        
    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 65869

    If lErro = 31415 Then gError 65870 'Não encontrou

    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
        If Len(Trim(NFiscalInterna.Caption)) > 0 Then
            objNFiscal.lNumNotaFiscal = CLng(NFiscalInterna.Caption)
        Else
            objNFiscal.lNumNotaFiscal = 0
        End If

    Else
        If Len(Trim(NFiscal.ClipText)) = 0 Then gError 65871
    End If

    objNFiscal.iTipoDocInfo = objTipoDocInfo.iTipo

    If Len(Trim(DataEntrada.ClipText)) = 0 Then gError 65872
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 65873

    'Verifica se a Data de Entrada e menor que a Data de Emissão
    If CDate(DataEntrada.Text) < CDate(DataEmissao.Text) Then gError 65876

    'Verifica de a Data de Vencimento e Menor que a de Emissão
    If Len(Trim(DataVencimento.ClipText)) <> 0 Then
        If CDate(DataVencimento.Text) < CDate(DataEmissao.Text) Then gError 65877
    End If

    If objTipoDocInfo.iComplementar = DOCINFO_NORMAL Then
        If objGridItens.iLinhasExistentes = 0 Then gError 65878
    End If
    
    'se o tipo de Nfiscal for do tipo "Externa - Mat. A ser beneficiado" entao obtem o Fornecedor a ser beneficiado
    If TipoNFiscal.ListIndex <> -1 Then
        
        If TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFEEBFCOM Or TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFIEBFCOM Or TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFECCOPC Then
            
            If Len(Trim(FornecedorBenef.Text)) = 0 Then gError 52359
            If Len(Trim(FilialFornBenef.List(FilialFornBenef.ListIndex))) = 0 Then gError 52360
        
        End If
    End If

    'Valida os dados do Grid
    If objGridItens.iLinhasExistentes > 0 Then
    
        For iIndice = 1 To objGridItens.iLinhasExistentes

            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 65879

            'Verifica se a Unidade de Medida foi preenchida
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) = 0 Then gError 65885

            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) = 0 Then gError 65886

        Next
    End If

    dTotal = CDbl(IIf(Len(Trim(Total.Text)) > 0, Total.Text, 0))

    'Se o total for negativo --> Erro
    If dTotal < 0 Then gError 65887
   
    If Len(Trim(PesoLiquido.ClipText)) > 0 And Len(Trim(PesoBruto.ClipText)) > 0 Then
        If CDbl(PesoLiquido.Text) > CDbl(PesoBruto.Text) Then gError 65888
    End If

    lErro = Total_Calcula()
    If lErro <> SUCESSO Then gError 65889

    lErro = CF2(Me, "NFEntCom_Valida")
    If lErro <> SUCESSO Then gError 140325

    lErro = gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
    If lErro <> SUCESSO Then gError 83281

    Set objNFiscal.objRastreamento = gobjRastreamento

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 65890

    objNFiscalBD.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objNFiscalBD.sSerie = objNFiscal.sSerie
    objNFiscalBD.lNumNotaFiscal = objNFiscal.lNumNotaFiscal
    objNFiscalBD.iTipoNFiscal = objNFiscal.iTipoNFiscal
    objNFiscalBD.dtDataEmissao = objNFiscal.dtDataEmissao
    objNFiscalBD.lCliente = objNFiscal.lCliente
    objNFiscalBD.iFilialCli = objNFiscal.iFilialCli
    objNFiscalBD.lFornecedor = objNFiscal.lFornecedor
    objNFiscalBD.iFilialForn = objNFiscal.iFilialForn

    'Verifica se foi cadastrada uma Nota Fiscal no BD a partir de FilialEmpresa, Serie, Número, Tipo, Data Emissão
    lErro = CF("NFiscal_Le_1", objNFiscalBD)
    If lErro <> SUCESSO And lErro <> 83971 Then gError 89231

    'se for uma nota nova
    If lErro <> SUCESSO Then
    
        'Critica os Valores
        lErro = Critica_Valores(objNFiscal)
        If lErro <> SUCESSO Then gError 66075

    End If

    'Atualiza coleções dos Itens das Notas Fiscais
    Call Atualiza_ItensPC_ItensRC
        
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then

        'Recolhe os dados da Nota Fiscal Original
        lErro = Move_Dados_NFOriginal_Memoria(objNFiscal, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 65891

    End If

    'Verifica se a Quantidade de cada Produto de GridItensNF é igual a soma da Quantidade Recebida desse mesmo Produto de ItensPC
    For iIndice = 1 To objGridItens.iLinhasExistentes
        
        Set objItemNF = objNFiscal.ColItensNF(iIndice)
        
        objProduto.sCodigo = objItemNF.sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 65881

        If lErro = 28030 Then gError 65882

        objItemNF.iClasseUM = objProduto.iClasseUM
        objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
        objItemNF.sUMVenda = objProduto.sSiglaUMVenda

        dQuantidade = 0
        dQuantidadeComMargem = 0
        
        For iLinha = 1 To objGridItensPC.iLinhasExistentes

            'Se a taxa nao estiver preenchida => Erro
            'If Len(Trim(GridItensPC.TextMatrix(iLinha, iGrid_Taxa_Col))) = 0 Then gError 114513
            
            bTeste = True
            Call CF2(Me, "NFEntCom_Produto_Igual_Customizado2", iIndice, iLinha, bTeste)
            
            If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = GridItensPC.TextMatrix(iLinha, iGrid_Prod_Col) And bTeste Then
            
                'Converte a UM de GridItensPC para a UM do GridItens
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItensPC.TextMatrix(iLinha, iGrid_UM_Col), GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), dFator)
                If lErro <> SUCESSO Then gError 89284
            
                Set objItemPCInfo = gcolItemPedCompraInfo.Item(iLinha)
            
                dQuantidade = dQuantidade + GridItensPC.TextMatrix(iLinha, iGrid_Recebido_Col) * dFator
    
                dQuantidadeComMargem = dQuantidadeComMargem + objItemPCInfo.dQuantPedida * dFator * (1 + objItemPCInfo.dPercentMaisReceb)
                dQuantidadeComMargem = dQuantidadeComMargem - (objItemPCInfo.dQuantPedida - CDbl(GridItensPC.TextMatrix(iLinha, iGrid_AReceber_Col)) * dFator)
                 
                If objItemPCInfo.iRecebForaFaixa = ERRO_E_REJEITA_RECEBIMENTO Then
                 
                    If dQuantidade - dQuantidadeComMargem > QTDE_ESTOQUE_DELTA Then gError 106865
                
                End If
                
                If objItemPCInfo.iRecebForaFaixa = AVISA_E_ACEITA_RECEBIMENTO Then
                
                    If dQuantidade - dQuantidadeComMargem > QTDE_ESTOQUE_DELTA Then
                    
                        vbMsg = Rotina_Aviso(vbYesNo, "AVISO_QUANTIDADE_ACIMA_FAIXA", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), dQuantidade, dQuantidadeComMargem)
            
                        If vbMsg = vbNo Then gError 126788
            
                    End If
            
                End If
                
            End If
            
        Next

        sCodProduto = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)
        
        'Se for diferente, Erro
        If Abs(StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)) - dQuantidade) > QTDE_ESTOQUE_DELTA Then gError 65892

    Next
    
    'Guarda os Pedidos e as Requisições que foram marcadas nas listas
    Call Move_PedReqCompra_Memoria(colPedCompras, colReqCompras)

    'Guarda os ItensNF associados aos ItensPC e aos ItensRC
    Call Move_ItensPedReqCompra_Memoria(objNFiscal)

    'Valida os dados da Tributação
    lErro = gobjTribTab.Valida_TributacaoNF_Gravacao(objNFiscal)
    If lErro <> SUCESSO Then gError 65893

    lErro = Calcula_Msg
    If lErro <> SUCESSO Then gError 65893

    'mover a parte do rastreamento
    lErro = gobjRastreamento.Move_Rastro_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 83276
    
    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786

    lErro = gobjTribTab.Valida_Dados
    If lErro <> SUCESSO Then gError 181786
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataEntrada.Text))
    If lErro <> SUCESSO Then gError 92034

    lErro = CF("NFiscal_Valida_Diversos", objNFiscal)
    If lErro <> SUCESSO Then gError 181786
    
    'Grava a Nota Fiscal de Entrada de Compras(inclusive os dados contábeis)
    lErro = CF("NFiscalEntradaCom_Grava", objNFiscal, objContabil, colPedCompras, colReqCompras)
    If lErro <> SUCESSO Then gError 65894

    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_NFISCAL
        gobjAnotacao.sID = CStr(objNFiscal.lNumIntDoc)
        gobjAnotacao.dtDataAlteracao = gdtDataHoje
        
        lErro = CF("Anotacoes_Grava", gobjAnotacao)
        If lErro <> SUCESSO Then gError 130406
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault

    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then

        If Len(Trim(NFiscalInterna.Caption)) = 0 Then vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_NUMERO_NOTA_GRAVADA", objNFiscal.lNumNotaFiscal)

    End If
    
    If ISSerieEletronica(objNFiscal.sSerie) Then
        lErro = CF("NFE_Grava1", objNFiscal.lNumIntDoc, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 203003
    End If
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 65863
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 65864
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 65865
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 65866
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 65867
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 65868
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCOMPRA_NAO_PREENCHIDA", gErr)

        Case 65869, 65880, 65881, 65889, 65890, 65891, 65893, 65894, 66075, 83276, 83281, 89231, 92034, 126788, 140325, 130406, 203003

        Case 65870
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iTipo)

        Case 65871
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)

        Case 65872
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_NAO_PREENCHIDA", gErr)

        Case 65873
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 65876
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_ANTERIOR_DATAEMISSAO", gErr, DataEntrada, DataEmissao)

        Case 65877
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR", gErr, DataEmissao.Text, DataVencimento.Text)

        Case 65878
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)

        Case 65879
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case 65882
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 65883
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 65884
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", gErr, GridItens.TextMatrix(iIndice, iGrid_Produto_Col))

        Case 65885
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 65886
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 65887
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)

        Case 65888
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr, CDbl(PesoLiquido.Text), CDbl(PesoBruto.Text))

        Case 65892
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_DIFERENT_QUANTRECEBIDA", gErr, sCodProduto)
            
        Case 106865
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_MAIOR_SALDO_PC", gErr, sCodProduto)
            
        Case 66062, 181786
            
        Case 75964
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTTOTAL_LOTE_MAIOR_ALMOXARIFADO", gErr, dQuantLote, StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)))
            
        Case 114513
            Call Rotina_Erro(vbOKOnly, "ERRO_TAXA_GRID_IMCOMPLETA", gErr)
                    
        Case 52359
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_BENEF_NAO_PREENCHIDA", gErr)
        
        Case 52360
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORNECEDOR_BENEF_NAO_PREENCHIDA", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156745)

    End Select

    Exit Function

End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objNFiscal As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objFornecedor As New ClassFornecedor
Dim snProt As String
Dim iEnviado As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 89076
    If Len(Trim(Serie.Text)) = 0 Then gError 89077
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 89078

    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    'alterado por cyntia
    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
    
    'le o tipo da origem
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 99345
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 99346
    
    'se for nfiscal interna
    If objTipoDocInfo.iOrigem = DOCINFO_EMPRESA Then
        'Verifica se a existe nota fiscal está cadastrada
        lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
        If lErro <> SUCESSO And lErro <> 43676 Then gError 99352
        
        If lErro = 43676 Then gError 89065
        
'        If gobjCRFAT.iUsaNFe = MARCADO Then
'
'            lErro = CF("NFeFedLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'            If lErro <> SUCESSO Then gError 203042
'
'            If iEnviado = 1 Then gError 203043
'
'        End If
    
    Else
        
        If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 99350
            
        'Guarda em objFornecedor o nome reduzido do Fornecedor
        objFornecedor.sNomeReduzido = Fornecedor.Text
    
        'Lê o código do Fornecedor a partir do nome reduzido obtido na tela
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 99351
        
        'Se não encontrou o fornecedor = > erro
        If lErro = 6681 Then gError 99352
        
        'Guarda no objNFiscal o código do fornecedor
        objNFiscal.lFornecedor = objFornecedor.lCodigo
        objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)
                
        'Verifica se a existe nota fiscal está cadastrada
        lErro = CF("NFiscal_Le_1", objNFiscal)
        If lErro <> SUCESSO And lErro <> 83971 Then gError 99353
        
        If lErro = 83971 Then gError 89065
        
    End If
    'até aqui
    
    'pede confirmacao
    vbMsg = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_NFISCAL", objNFiscal.lNumNotaFiscal)
    If vbMsg = vbYes Then

        'Lê os itens da nota fiscal
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 89081

        'Faz o cancelamento de uma nota fiscal de Saida
        lErro = CF("NotaFiscalEntrada_Excluir", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError 89082
    
        'Limpa a Tela
        Call Limpa_Tela_NFiscalCom1
    
        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 89076
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 89077
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 89078
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 89078, 89081, 89082, 99345, 99351, 99352, 99353, 203042
       
        Case 89065
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
            
        Case 99346
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)
        
        Case 99350
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
        
        Case 99352
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)
       
        Case 203043
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CADASTRADA_NAO_EXCLUI", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156746)
        
    End Select

    Exit Sub

End Sub

Function Atualiza_ItensPC_ItensRC() As Long
'Atualiza as quantidades das coleções globais

Dim iIndice As Integer
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim objItemRCInfo As ClassItemReqComprasInfo

    'Atualiza a quantidade recebida dos itens de pedido de Compras
    For Each objItemPCInfo In gcolItemPedCompraInfo
        iIndice = iIndice + 1
        objItemPCInfo.dQuantRecebida = StrParaDbl(GridItensPC.TextMatrix(iIndice, iGrid_Recebido_Col))
    Next

    iIndice = 0
    'Atualiza a quantidade recebida dos itens de Requisição de Compras
    For Each objItemRCInfo In gcolItemReqCompraInfo
        iIndice = iIndice + 1
        objItemRCInfo.dQuantRecebida = StrParaDbl(GridItensRC.TextMatrix(iIndice, iGrid_RecebidoRC_Col))
        objItemRCInfo.dQuantRecebidoPedido = StrParaDbl(GridItensRC.TextMatrix(iIndice, iGrid_QuantRecebidaPC_Col))
    Next
        
End Function

Function Critica_Valores(objNFiscal As ClassNFiscal) As Long
'Verifica se os valores do Pedidos de Compras e dos ItensPC são
'iguais aos que foram colocados no frame de Itens

Dim lErro As Long
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim dValorTotalFrete As Double
Dim dValorTotalSeguro As Double
Dim dValorTotalDespesas As Double
Dim dValorTotalDescontos As Double
Dim iIndice As Integer
Dim objPedidoCompras As ClassPedidoCompras
Dim objItemPC As ClassItemPedCompra
Dim sProdutoMascarado As String
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice2 As Integer
Dim iItem As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim objItemNF As ClassItemNF
Dim colItemNF As New Collection
Dim iAchou As Integer, dPrecoUnitNF As Double, dPrecoUnitPC As Double
Dim bTeste As Boolean
Dim bArredondamento As Boolean, dQuantidade As Double, dDifTotItem As Double, iLinha As Integer
Dim bAchou As Boolean, bIgnora As Boolean

On Error GoTo Erro_Critica_Valores


    'Lê os valores dos Pedidos de Compras marcados
    lErro = CF("PedidoCompras_Le_Valores", gcolPedidoCompra)
    If lErro <> SUCESSO Then gError 66062
        
    'Para cada Pedido de Compras
    For Each objPedidoCompras In gcolPedidoCompra
        
        For iIndice2 = 0 To PedidosCompra.ListCount - 1
            
            If PedidosCompra.List(iIndice2) = CStr(objPedidoCompras.lCodigo) And PedidosCompra.Selected(iIndice2) = True Then
        
                'Acumula os Valores dos Pedidos marcados
                dValorTotalFrete = dValorTotalFrete + objPedidoCompras.dValorFrete
                dValorTotalSeguro = dValorTotalSeguro + objPedidoCompras.dValorSeguro
                dValorTotalDespesas = dValorTotalDespesas + objPedidoCompras.dOutrasDespesas
                dValorTotalDescontos = dValorTotalDescontos + objPedidoCompras.dValorDesconto
                
                'Lê os Itens do Pedido
                lErro = CF("ItensPC_Le_Codigo", objPedidoCompras)
                If lErro <> SUCESSO Then gError 66071
                
                lErro = CF("PCItens_Le_Customizado", objPedidoCompras)
                If lErro <> SUCESSO Then gError 181263
            
                'Para cada Item do Pedido de Compras
                For Each objItemPC In objPedidoCompras.colItens
                    
                    'Mascara o Produto
                    lErro = Mascara_RetornaProdutoEnxuto(objItemPC.sProduto, sProdutoMascarado)
                    If lErro <> SUCESSO Then gError 66072
                    
                    ProdutoPC.PromptInclude = False
                    ProdutoPC.Text = sProdutoMascarado
                    ProdutoPC.PromptInclude = True
            
                    sProdutoMascarado = ProdutoPC.Text
                    
                    objProduto.sCodigo = objItemPC.sProduto
                
                    'Lê o produto da linha passada por iLinha do GridItens
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 89275
                    
                    If lErro = 28030 Then gError 89276

                    bAchou = False
                    bIgnora = True
                    For Each objItemPCInfo In gcolItemPedCompraInfo
                        If objItemPCInfo.lNumIntDoc = objItemPC.lNumIntDoc Then
                            bAchou = True
                            If objItemPCInfo.dQuantReceber > QTDE_ESTOQUE_DELTA Then bIgnora = False 'Tem algo sendo recebido desse item do pedido
                        End If
                    Next
                    
                    If Not bIgnora Then
                        
                        'Procura no GridItens o Produto igual ao do ItemPC
                        For iIndice = 1 To objGridItens.iLinhasExistentes
                            
                            bTeste = True
                            Call CF2(Me, "NFEntCom_Produto_Igual_Customizado", objItemPC.objInfoUsu, iIndice, bTeste)
                            
                            'Se encontrou
                            If sProdutoMascarado = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) And bTeste Then
                             
                                'Converte a UM de GridItensPC para a UM do GridItens
                                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPC.sUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), dFator)
                                If lErro <> SUCESSO Then gError 89277
                                
                                'preco unitario do item da nf na unidade de medida do item pc
                                dPrecoUnitNF = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col)) * dFator
                                'preco unitario do item de pc em R$
                                'dPrecoUnitPC = objItemPC.dPrecoUnitario * IIf(objItemPC.iMoeda <> MOEDA_REAL, objItemPC.dTaxa, 1)
                                
                                If objItemPC.iMoeda = MOEDA_REAL Then
                                        
                                    'preco unitario do item de pc em R$
                                    dPrecoUnitPC = objItemPC.dPrecoUnitario
                                        
                                Else
                                
                                    For iLinha = 1 To objGridItensPC.iLinhasExistentes
                            
                                        If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = GridItensPC.TextMatrix(iLinha, iGrid_Prod_Col) Then
                                        
                                            dPrecoUnitPC = StrParaDbl(GridItensPC.TextMatrix(iLinha, iGrid_Recebido_RS_Col))
                                            
                                            If dPrecoUnitPC = 0 Then gError 211741
                                            Exit For
                                            
                                        End If
                                        
                                    Next
                                
                                End If
                                
                                'Verifica se possui o mesmo unitário
                                If Abs(dPrecoUnitNF - dPrecoUnitPC) > DELTA_VALORMONETARIO2 Then
                                
                                    bArredondamento = False
                                    
                                    If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then
                                        dQuantidade = CDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
                                    Else
                                        dQuantidade = 1
                                    End If
                                    
                                    dDifTotItem = ((dPrecoUnitNF / dFator * dQuantidade) - (dPrecoUnitPC * dQuantidade / dFator))
                                    If dPrecoUnitPC <> 0 Then
                                        If Abs(dDifTotItem) / (dPrecoUnitPC * dQuantidade / dFator) < 0.011 Then
                                            If (dDifTotItem >= 0 And dDifTotItem < 10) Or (dDifTotItem < 0 And dDifTotItem > -20) Then bArredondamento = True
                                        End If
                                    End If
                                
                                    'Se aceita preços unitários de Itens da Nota Fiscal diferente dos Itens PC
                                    If gobjCOM.iNFDiferentePC = NFISCAL_ACEITA_DIFERENCA_PC Or bArredondamento Then
        
                                        'Exibe mensagem de aviso de o preço unitário for diferente
                                        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_VALORUNITARIO_DIFERENTE_PC", dPrecoUnitNF, iIndice, dPrecoUnitPC)
                                        If vbMsgRes = vbNo Then gError 66073
        
                                    'Se não aceita
                                    Else
        
                                        'Exibe mensagem de erro
                                        gError 65516
        
                                    End If
                                                
                                End If
                                            
                            End If
                        
                        Next
                                                       
                        iAchou = 0
                        
                        For Each objItemNF In colItemNF
                            
                            bTeste = True
                            Call CF2(Me, "NFEntCom_Produto_Igual_Customizado5", objItemPC.objInfoUsu, objItemNF.objInfoUsu, bTeste)
                            
                            If objItemNF.sProduto = sProdutoMascarado And bTeste Then
                                objItemNF.dValorDesconto = objItemNF.dValorDesconto + objItemPC.dValorDesconto
                                iAchou = 1
                                Exit For
                            End If
                        Next
                        
                        If iAchou = 0 Then
                            Set objItemNF = New ClassItemNF
                            objItemNF.sProduto = sProdutoMascarado
                            objItemNF.dValorDesconto = objItemPC.dValorDesconto
                            
                            Call CF2(Me, "NFEntCom_Tranfere_Dados_ItemPC_Para_ItemNF", objItemNF, objItemPC)
                            
                            colItemNF.Add objItemNF
                        End If
                                                       
                        For iItem = 1 To gobjTribTab.objNFTributacao.colItens.Count
                        
                            If gobjTribTab.objNFTributacao.colItens(iItem).sProduto = objItemPC.sProduto Then
    
                                'Se as alíquotas dos ItensNF são diferentes dos Itens PC
                                If Abs(gobjTribTab.objNFTributacao.colItens(iItem).dIPIAliquota - objItemPC.dAliquotaIPI) > DELTA_VALORMONETARIO2 And objItemPC.dAliquotaIPI <> -1 Then
    
                                    'Se é permitido Itens NF que tenha alíquota IPIe alíquota ICMS diferentes dos ItensPC
                                    If gobjCOM.iNFDiferentePC = NFISCAL_ACEITA_DIFERENCA_PC Then
    
                                        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALIQUOTAIPI_DIFERENTE_PC", objItemPC.sProduto, gobjTribTab.objNFTributacao.colItens(iItem).dIPIAliquota, objItemPC.dAliquotaIPI)
                                        If vbMsgRes = vbNo Then gError 70550
                                    Else
                                        gError 67305
                                    End If
    
                                End If
    
                                'Se as alíquotas dos ItensNF são diferentes dos Itens PC
                                If Abs(gobjTribTab.objNFTributacao.colItens(iItem).dICMSAliquota - objItemPC.dAliquotaICMS) > DELTA_VALORMONETARIO2 And objItemPC.dAliquotaICMS <> -1 Then
    
                                    'Se é permitido Itens NF que tenha alíquota IPI e alíquota ICMS diferentes dos ItensPC
                                    If gobjCOM.iNFDiferentePC = NFISCAL_ACEITA_DIFERENCA_PC Then
    
                                        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ALIQUOTAICMS_DIFERENTE_PC", objItemPC.sProduto, gobjTribTab.objNFTributacao.colItens(iItem).dICMSAliquota, objItemPC.dAliquotaICMS)
                                        If vbMsgRes = vbNo Then gError 70551
                                    Else
    
                                        gError 67306
                                    End If
    
                                End If
    
                                Exit For
    
                            End If
                        Next
                        
                    End If
                Next
        
            End If
        
        Next
    
    Next
    
    'Procura no GridItens o Produto igual ao do ItemPC
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        iAchou = 0
    
        For Each objItemNF In colItemNF

            bTeste = True
            Call CF2(Me, "NFEntCom_Produto_Igual_Customizado6", objItemNF.objInfoUsu, iIndice, bTeste)

            If objItemNF.sProduto = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) And bTeste Then
                iAchou = 1
                Exit For
            End If
        Next
    
        If iAchou = 1 Then
            'Verifica se possui o mesmo desconto
            If Abs(StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col)) - objItemNF.dValorDesconto) > DELTA_VALORMONETARIO2 Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DESCONTOITEM_DIFERENTE_PC", StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col)), iIndice, objItemNF.dValorDesconto)
                If vbMsgRes = vbNo Then gError 66074
            End If
        End If
                                                   
    Next
                                                   
    'Se o Valor Frete da Nota Fiscal for diferente do Pedido de Compras
    If Abs(objNFiscal.dValorFrete - dValorTotalFrete) > DELTA_VALORMONETARIO2 Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_VALORFRETE_DIFERENTE_PC", objNFiscal.dValorFrete, dValorTotalFrete)
        If vbMsgRes = vbNo Then gError 66067
    End If
    
    'Se o Valor Seguro da Nota Fiscal for diferente do Pedido de Compras
    If Abs(objNFiscal.dValorSeguro - dValorTotalSeguro) > DELTA_VALORMONETARIO2 Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_VALORSEGURO_DIFERENTE_PC", objNFiscal.dValorSeguro, dValorTotalSeguro)
        If vbMsgRes = vbNo Then gError 66068
    End If
    
    'Se o Valor Despesas da Nota Fiscal for diferente do Pedido de Compras
    If Abs(objNFiscal.dValorOutrasDespesas - dValorTotalDespesas) > DELTA_VALORMONETARIO2 Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_VALORDESPESAS_DIFERENTE_PC", objNFiscal.dValorOutrasDespesas, dValorTotalDespesas)
        If vbMsgRes = vbNo Then gError 66069
    End If
    
    'Se o Valor Desconto da Nota Fiscal for diferente do Pedido de Compras
    If Abs(objNFiscal.dValorDesconto - dValorTotalDescontos) > DELTA_VALORMONETARIO2 Then
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_VALORDESCONTO_DIFERENTE_PC", objNFiscal.dValorDesconto, dValorTotalDescontos)
        If vbMsgRes = vbNo Then gError 66070
    End If
                
    Critica_Valores = SUCESSO
    
    Exit Function
    
Erro_Critica_Valores:

    Critica_Valores = gErr
    
    Select Case gErr
    
        Case 66062, 66067, 66068, 66069, 66070, 66071, 66072, 66073, 66074, 70550, 70551, 89275, 89277
        
        Case 65516
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_PC_PRECOUNITARIO_DIFERENTE", gErr, objItemPC.sProduto)
        
        Case 67305
            Call Rotina_Erro(vbOKOnly, "ERRO_ALIQUOTAIPINF_DIFERENTE_PC", gErr, objItemPC.sProduto, gobjTribTab.objNFTributacao.colItens(iItem).dIPIAliquota, objItemPC.dAliquotaIPI)
        
        Case 67306
            Call Rotina_Erro(vbOKOnly, "ERRO_ALIQUOTAICMSNF_DIFERENTE_PC", gErr, objItemPC.sProduto, gobjTribTab.objNFTributacao.colItens(iItem).dICMSAliquota, objItemPC.dAliquotaICMS)
            
        Case 89276
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case 211741
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_PRECOUNITARIO_NAO_PREENCHIDO_PC", gErr, iLinha)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156747)
    
    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidade de Medida que está deixando de ser a corrente

Dim lErro As Long
Dim sUM As String

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed
    
    
    sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 65895

    'Se a Unidade de Medida foi modificada
    If Trim(sUM) <> Trim(UnidadeMed.Text) Then

        'Atualiza a UM e a quantidade de GridItensPC e GridItensRC
        lErro = Atualiza_UM(GridItens.Row, sUM)
        If lErro <> SUCESSO Then gError 65896
        
        If Not gbCarregandoTela Then
            lErro = gobjTribTab.Pesos_Recalcula()
            If lErro <> SUCESSO Then gError 65896
        End If

    End If

'inicio distribuicao
    If sUM <> UnidadeMed.Text Then
        
        lErro = gobjDistribuicao.Preenche_GridDistribuicaoPC1(gcolItemPedCompraInfo)
        If lErro <> SUCESSO Then gError 89549
        
        lErro = gobjRastreamento.Atualiza_Troca_UM(GridItens.Row)
        If lErro <> SUCESSO Then gError 83282
        
    End If
    
'fim distribuicao
    
    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 65895, 65896, 83282, 89549
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156748)

    End Select

    Exit Function

End Function

Private Function Atualiza_UM(iLinha As Integer, sUM As String) As Long
'Atualiza a UM e a quantidade de GridItensPC e GridItensRC

Dim lErro As Long
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Atualiza_UM

    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 65897

    objProduto.sCodigo = sProdutoFormatado

    'Lê o produto da linha passada por iLinha do GridItens
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 65898
    If lErro = 28030 Then gError 65899

    'Converte a UM de GridItensPC para a UM do GridItens
    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUM, GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), dFator)
    If lErro <> SUCESSO Then gError 65900

    'Atualiza o Grid de Recebimento
    If (StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) * dFator) <> 0 Then GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) * dFator)
    GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col) = Format(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col)) / dFator, FORMATO_PRECO_UNITARIO_EXTERNO)

    Atualiza_UM = SUCESSO
    
    Exit Function

Erro_Atualiza_UM:

    Atualiza_UM = gErr

    Select Case gErr

        Case 65897, 65898, 65900, 65901

        Case 65899
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156749)

    End Select

    Exit Function

End Function

Public Sub UpDownVencimento_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVencimento_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataVencimento, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 65903

    Exit Sub

Erro_UpDownVencimento_DownClick:

    Select Case gErr

        Case 65903

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156750)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencimento_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVencimento_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataVencimento, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 65904

    Exit Sub

Erro_UpDownVencimento_UpClick:

    Select Case gErr

        Case 65904

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156751)

    End Select

    Exit Sub

End Sub

Public Sub VolumeEspecie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VolumeEspecie_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub VolumeEspecie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VolumeEspecie_Validate

    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie, "AVISO_CRIAR_VOLUMEESPECIE")
    If lErro <> SUCESSO Then gError 102454
    
    Exit Sub

Erro_VolumeEspecie_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102454
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156752)

    End Select

End Sub

Public Sub VolumeMarca_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub VolumeMarca_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub VolumeMarca_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VolumeMarca_Validate

    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca, "AVISO_CRIAR_VOLUMEMARCA")
    If lErro <> SUCESSO Then gError 102455
    
    Exit Sub

Erro_VolumeMarca_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102455
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156753)

    End Select

End Sub

Public Sub VolumeNumero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub VolumeQuant_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Move_Dados_NFOriginal_Memoria(objNFiscal As ClassNFiscal, objTipoDocInfo As ClassTipoDocInfo) As Long
'Move os dados da NotaFiscalOriginal para a memória

Dim lErro As Long
Dim objNFOriginal As New ClassNFiscal

On Error GoTo Erro_Move_Dados_NFOriginal_Memoria

    'verifica se a Serie e o Número da NF original estão preenchidos
    If Len(Trim(NFiscalOriginal.ClipText)) = 0 Then gError 65905
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Then gError 65906

    objNFOriginal.lFornecedor = objNFiscal.lFornecedor
    objNFOriginal.iFilialForn = objNFiscal.iFilialForn
    objNFOriginal.sSerie = SerieNFiscalOriginal.Text
    objNFOriginal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
         
    'Lê a Nota Fiscal Original
    lErro = CF("NFiscal_Le_Num_Forn", objNFOriginal, objTipoDocInfo.iTipo)
    If lErro <> SUCESSO And lErro <> 31467 Then gError 65907
    
    If lErro = 31467 Then gError 65908 'Se não encontrar --> Erro
    
    'guarda em objNfiscal o Numero Interno da NF Original
    objNFiscal.lNumIntNotaOriginal = objNFOriginal.lNumIntDoc

    Move_Dados_NFOriginal_Memoria = SUCESSO

    Exit Function

Erro_Move_Dados_NFOriginal_Memoria:

    Move_Dados_NFOriginal_Memoria = gErr

    Select Case gErr

        Case 65905
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NFISCAL_ORIGINAL_NAO_PREENCHIDO", gErr)

        Case 65906
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NFISCAL_ORIGINAL_NAO_PREENCHIDA", gErr)

        Case 65907

        Case 65908
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA1", gErr, objNFOriginal.sSerie, objNFOriginal.lNumNotaFiscal)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156754)

    End Select

    Exit Function

End Function

Private Sub Move_PedReqCompra_Memoria(colPedCompras As Collection, colReqCompras As Collection)
'Move para a memória os Pedidos e as Requisições selecionadas

Dim iIndice As Integer
Dim objPedidoCompra As ClassPedidoCompras

    'Guarda o código dos Pedidos de Compras que estão selecionados
    For iIndice = 0 To PedidosCompra.ListCount - 1
        
        If PedidosCompra.Selected(iIndice) = True Then
            Set objPedidoCompra = New ClassPedidoCompras
            objPedidoCompra.lCodigo = CLng(PedidosCompra.List(iIndice))
            colPedCompras.Add objPedidoCompra
        End If

    Next

    'Guarda o código das Requisições de Compras que estão selecionadas
    For iIndice = 0 To RequisicoesCompra.ListCount - 1
        
        If RequisicoesCompra.Selected(iIndice) = True Then
            colReqCompras.Add CLng(RequisicoesCompra.List(iIndice))
        End If

    Next

End Sub

Function Move_ItensPedReqCompra_Memoria(objNFiscal As ClassNFiscal) As Long
'Move para a memória os ItensNF associados aos ItensPC e aos ItensRC

Dim objItemNF As ClassItemNF
Dim objItemPCInfo As ClassItemPedCompraInfo
Dim objItemNFItemPC As ClassItemNFItemPC
Dim objItemRCInfo As ClassItemReqComprasInfo
Dim objItemNFItemRC As ClassItemNFItemRC

    'Para cada Item da Nota Fiscal
    For Each objItemNF In objNFiscal.ColItensNF
        
        Set objItemNF.colItemNFItemPC = New Collection
        Set objItemNF.colItemNFItemRC = New Collection
        
        For Each objItemPCInfo In gcolItemPedCompraInfo
                
            If objItemNF.sProduto = objItemPCInfo.sProduto Then
            
                Set objItemNFItemPC = New ClassItemNFItemPC
                
                objItemNFItemPC.dQuantidade = objItemPCInfo.dQuantRecebida
                objItemNFItemPC.lItemPedCompra = objItemPCInfo.lNumIntDoc
                objItemNFItemPC.dTaxa = objItemPCInfo.dTaxa
                
                'Adiciona na coleção o objItemNFItemPC
                objItemNF.colItemNFItemPC.Add objItemNFItemPC
            
                For Each objItemRCInfo In gcolItemReqCompraInfo
            
                    If objItemRCInfo.lNumIntDocItemPC = objItemPCInfo.lNumIntDoc Then
                    
                        Set objItemNFItemRC = New ClassItemNFItemRC
                        
                        objItemNFItemRC.dQuantidade = objItemRCInfo.dQuantRecebida
                        objItemNFItemRC.lItemReqCompra = objItemRCInfo.lNumIntDoc
                        objItemNFItemRC.lNumIntDocItemPC = objItemRCInfo.lNumIntDocItemPC
                    
                        'Adiciona na coleção objItemNFItemRC
                        objItemNF.colItemNFItemRC.Add objItemNFItemRC
                            
                    End If
                    
                Next
    
            End If
        
        Next
        
    Next

End Function

Public Function ValorTotal_Calcula() As Long

Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    lErro = gobjTribTab.AtualizarTributacao()
    If lErro <> SUCESSO Then gError 65909
    
    Call ValorDescontoItens_Calcula

    If iEmitenteNF = DOCINFO_EMPRESA Then
    
        'atualizar o total da nota
        Call Total_Calcula(0)
        
    End If
        
    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 65909

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156755)

    End Select

    Exit Function

End Function

Private Function Trata_FilialForn(Optional bSoTrib As Boolean = False) As Long

Dim lErro As Long
Dim objFilialForn As New ClassFilialFornecedor

On Error GoTo Erro_Trata_FilialForn
    
    objFilialForn.iCodFilial = Codigo_Extrai(Filial.Text)
    
    If objFilialForn.iCodFilial <> 0 Then
    
        'Lê a Filial
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilialForn)
        If lErro <> SUCESSO Then gError 65910
        
        Call gobjTribTab.FilialFornecedor_Alterada(objFilialForn.lCodFornecedor, objFilialForn.iCodFilial)
    
        If bSoTrib = False Then
            lErro = CF2(Me, "NFCF_Fornecedor_Filial_Validate", objFilialForn)
            If lErro <> SUCESSO Then gError 131939
        End If
    
    End If
    
    Trata_FilialForn = SUCESSO
     
    Exit Function
    
Erro_Trata_FilialForn:

    Trata_FilialForn = gErr
     
    Select Case gErr
          
        Case 65910, 131939 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156756)
     
    End Select
     
    Exit Function

End Function

Public Function Move_NF_Memoria(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objFornecedor As ClassFornecedor
Dim objNFiscalOriginal As New ClassNFiscal
Dim objProjetoInfo As Object

On Error GoTo Erro_Move_NF_Memoria

    'Se o Fornecedor estiver preenchido lê o código no BD
    If Len(Trim(Fornecedor.Text)) > 0 Then

        Set objFornecedor = New ClassFornecedor

        objFornecedor.sNomeReduzido = Fornecedor.Text

        'Lê o fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 65911

        If lErro = 6681 Then gError 65912 'Se nao encontrar --> erro

        objNFiscal.lFornecedor = objFornecedor.lCodigo

    End If

    'Recolhe os demais campos da Tela
    If Len(Trim(NumRecebimento.Text)) > 0 Then objNFiscal.lNumRecebimento = CLng(NumRecebimento.Text)

    If Len(Trim(Filial.Text)) > 0 Then objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)

    If Len(Trim(TipoNFiscal.Text)) > 0 Then objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    If objNFiscal.iTipoDocInfo = TIPODOCINFO_TIPO_NFEXT Then
        objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
    Else
        objNFiscal.lNumNotaFiscal = StrParaLong(NFiscalInterna.Caption)
    End If

    objNFiscal.iFreteRespons = TipoFrete.ItemData(TipoFrete.ListIndex)

    objNFiscal.dtDataReferencia = DATA_NULA
    objNFiscal.lNumIntDoc = 0
    objNFiscal.sNaturezaOp = NaturezaOp.Text
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.sPlaca = Placa.Text
    objNFiscal.sPlacaUF = PlacaUF.Text
    If Not NFiscalInterna.Visible Then objNFiscal.sChvNFe = ChvNFe.ClipText
    
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objNFiscal.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objNFiscal.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
    
    objNFiscal.sMensagemNota = Mensagem.Text
    objNFiscal.sMensagemCorpoNota = MensagemCorpo.Text
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.dtDataSaida = DATA_NULA
    objNFiscal.iStatus = STATUS_LANCADO
    objNFiscal.sObservacao = Observacao.Text

    objNFiscal.iFilialPedido = Codigo_Extrai(FilialCompra.Text)
    
    If Len(Trim(ValorDesconto.ClipText)) > 0 Then objNFiscal.dValorDesconto = CDbl(ValorDesconto.Text)
    If Len(Trim(ValorDespesas.ClipText)) > 0 Then objNFiscal.dValorOutrasDespesas = CDbl(ValorDespesas.Text)
    If Len(Trim(ValorFrete.ClipText)) > 0 Then objNFiscal.dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorSeguro.ClipText)) > 0 Then objNFiscal.dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(Total.Text)) > 0 Then objNFiscal.dValorTotal = CDbl(Total.Text)
    If Len(Trim(SubTotal.Caption)) > 0 Then objNFiscal.dValorProdutos = CDbl(SubTotal.Caption)

    If Len(Trim(PesoBruto.ClipText)) > 0 Then objNFiscal.dPesoBruto = CDbl(PesoBruto.Text)
    If Len(Trim(PesoLiquido.ClipText)) > 0 Then objNFiscal.dPesoLiq = CDbl(PesoLiquido.Text)

    If Len(Trim(VolumeQuant.ClipText)) > 0 Then objNFiscal.lVolumeQuant = StrParaLong(VolumeQuant.Text)

    If Len(Trim(Transportadora.Text)) > 0 Then objNFiscal.iCodTransportadora = Codigo_Extrai(Transportadora.Text)

    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        objNFiscal.dtDataEmissao = CDate(DataEmissao.Text)
    Else
        objNFiscal.dtDataEmissao = DATA_NULA
    End If
    
    If Len(Trim(DataEntrada.ClipText)) > 0 Then
        objNFiscal.dtDataEntrada = CDate(DataEntrada.Text)
    Else
        objNFiscal.dtDataEntrada = DATA_NULA
    End If
    
'horaentrada
    If objNFiscal.dtDataEntrada <> DATA_NULA Then
        If Len(Trim(HoraEntrada.ClipText)) > 0 Then
            objNFiscal.dtHoraEntrada = CDate(HoraEntrada.Text)
        Else
            objNFiscal.dtHoraEntrada = Time
        End If
    Else
        objNFiscal.dtHoraEntrada = 0
    End If
    
    If Len(Trim(DataVencimento.ClipText)) > 0 Then
        objNFiscal.dtDataVencimento = CDate(DataVencimento.Text)
    Else
        objNFiscal.dtDataVencimento = DATA_NULA
    End If

    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181761
    
    Set objNFiscal.objProjetoInfo = objProjetoInfo
    
    'esta da indica quando está sendo registrada a nota fiscal no sistema. Será utilizado na gravacao do estoque
    objNFiscal.dtDataRegistro = gdtDataAtual
    
    Set objNFiscal.objInfoAdic = gobjInfoAdic
    objNFiscal.dValorItens = StrParaDbl(ValorProdutos2.Caption)
    objNFiscal.dValorDescontoItens = StrParaDbl(ValorDescontoItens.Text)
    
    If Len(Trim(FornecedorBenef.Text)) > 0 Then
    
        Set objFornecedor = New ClassFornecedor
    
        objFornecedor.sNomeReduzido = FornecedorBenef.Text
    
        'Lê o fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 52510
    
        If lErro = 6681 Then gError 52511 'Se nao encontrar --> erro

        objNFiscal.lFornecedorBenef = objFornecedor.lCodigo

        If Len(Trim(FilialFornBenef.List(FilialFornBenef.ListIndex))) > 0 Then objNFiscal.iFilialFornBenef = Codigo_Extrai(FilialFornBenef.List(FilialFornBenef.ListIndex))
    End If

    Move_NF_Memoria = SUCESSO

    Exit Function

Erro_Move_NF_Memoria:

    Move_NF_Memoria = gErr

    Select Case gErr

        Case 52510
        
        Case 52511
          Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, FornecedorBenef.Text)
        
        Case 65911, 181761

        Case 65912
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, Fornecedor.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156757)

    End Select

    Exit Function

End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is NaturezaOp Then
            Call NaturezaLabel_Click
        ElseIf Me.ActiveControl Is Serie Then
            Call SerieLabel_Click
        ElseIf Me.ActiveControl Is NFiscal Then
            Call NFiscalLabel_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
'distribuicao
        ElseIf Me.ActiveControl Is gobjDistribuicao.AlmoxDist Then
            Call gobjDistribuicao.BotaoLocalizacaoDist_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        ElseIf Me.ActiveControl Is NFiscalOriginal Then
            Call NFiscalOriginalLabel_Click
        ElseIf Me.ActiveControl Is SerieNFiscalOriginal Then
            Call NFiscalOriginalLabel_Click
        ElseIf Me.ActiveControl Is NumRecebimento Then
            Call RecebimentoLabel_Click
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        Else
            Call gobjRastreamento.objUserControl_KeyDown(KeyCode, Shift)
        End If
    
    End If

End Sub

Public Sub VolumeQuant_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VolumeQuant, iAlterado)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Nota Fiscal Entrada - c/ Pedidos de Compra"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "NFiscalEntradaCom"

End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho

   RaiseEvent Unload

End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

'inicio contabilidade
Public Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Public Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Public Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

End Sub

Public Sub CTBGridContabil_EnterCell()

    Call objContabil.Contabil_GridContabil_EnterCell

End Sub

Public Sub CTBGridContabil_GotFocus()

    Call objContabil.Contabil_GridContabil_GotFocus

End Sub

Public Sub CTBGridContabil_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)

End Sub

Public Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)

    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)
    
End Sub


Public Sub CTBGridContabil_LeaveCell()

        Call objContabil.Contabil_GridContabil_LeaveCell

End Sub

Public Sub CTBGridContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_GridContabil_Validate(Cancel)

End Sub

Public Sub CTBGridContabil_RowColChange()

    Call objContabil.Contabil_GridContabil_RowColChange

End Sub

Public Sub CTBGridContabil_Scroll()

    Call objContabil.Contabil_GridContabil_Scroll

End Sub

Public Sub CTBConta_Change()

    Call objContabil.Contabil_Conta_Change

End Sub

Public Sub CTBConta_GotFocus()

    Call objContabil.Contabil_Conta_GotFocus

End Sub

Public Sub CTBConta_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)

End Sub

Public Sub CTBConta_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Conta_Validate(Cancel)

End Sub

Public Sub CTBCcl_Change()

    Call objContabil.Contabil_Ccl_Change

End Sub

Public Sub CTBCcl_GotFocus()

    Call objContabil.Contabil_Ccl_GotFocus

End Sub

Public Sub CTBCcl_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)

End Sub

Public Sub CTBCcl_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Ccl_Validate(Cancel)

End Sub

Public Sub CTBCredito_Change()

    Call objContabil.Contabil_Credito_Change

End Sub

Public Sub CTBCredito_GotFocus()

    Call objContabil.Contabil_Credito_GotFocus

End Sub

Public Sub CTBCredito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)

End Sub

Public Sub CTBCredito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Credito_Validate(Cancel)

End Sub

Public Sub CTBDebito_Change()

    Call objContabil.Contabil_Debito_Change

End Sub

Public Sub CTBDebito_GotFocus()

    Call objContabil.Contabil_Debito_GotFocus

End Sub

Public Sub CTBDebito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)

End Sub

Public Sub CTBDebito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Debito_Validate(Cancel)

End Sub

Public Sub CTBSeqContraPartida_Change()

    Call objContabil.Contabil_SeqContraPartida_Change

End Sub

Public Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

Public Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

Public Sub CTBSeqContraPartida_Validate(Cancel As Boolean)

    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)

End Sub

Public Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

End Sub

Public Sub CTBHistorico_Change()

    Call objContabil.Contabil_Historico_Change

End Sub

Public Sub CTBHistorico_GotFocus()

    Call objContabil.Contabil_Historico_GotFocus

End Sub

Public Sub CTBHistorico_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)

End Sub

Public Sub CTBHistorico_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Historico_Validate(Cancel)

End Sub

Public Sub CTBLancAutomatico_Click()

    Call objContabil.Contabil_LancAutomatico_Click

End Sub

Public Sub CTBAglutina_Click()
    
    Call objContabil.Contabil_Aglutina_Click

End Sub

Public Sub CTBAglutina_GotFocus()

    Call objContabil.Contabil_Aglutina_GotFocus

End Sub

Public Sub CTBAglutina_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)

End Sub

Public Sub CTBAglutina_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Aglutina_Validate(Cancel)

End Sub

Public Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_NodeClick(Node)

End Sub

Public Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwCcls_NodeClick(Node)

End Sub

Public Sub CTBListHistoricos_DblClick()

    Call objContabil.Contabil_ListHistoricos_DblClick

End Sub

Public Sub CTBBotaoLimparGrid_Click()

    Call objContabil.Contabil_Limpa_GridContabil

End Sub

Public Sub CTBLote_Change()

    Call objContabil.Contabil_Lote_Change

End Sub

Public Sub CTBLote_GotFocus()

    Call objContabil.Contabil_Lote_GotFocus

End Sub

Public Sub CTBLote_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)

End Sub

Public Sub CTBDataContabil_Change()

    Call objContabil.Contabil_DataContabil_Change

End Sub

Public Sub CTBDataContabil_GotFocus()

    Call objContabil.Contabil_DataContabil_GotFocus

End Sub

Public Sub CTBDataContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)

End Sub

Public Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Public Sub objEventoDoc_evSelecao(obj1 As Object)

    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)

End Sub

Public Sub CTBDocumento_Change()

    Call objContabil.Contabil_Documento_Change

End Sub

Public Sub CTBDocumento_GotFocus()

    Call objContabil.Contabil_Documento_GotFocus

End Sub

Public Sub CTBBotaoImprimir_Click()
    
    Call objContabil.Contabil_BotaoImprimir_Click

End Sub

Public Sub CTBUpDown_DownClick()

    Call objContabil.Contabil_UpDown_DownClick
    
End Sub

Public Sub CTBUpDown_UpClick()

    Call objContabil.Contabil_UpDown_UpClick

End Sub

Public Sub CTBLabelDoc_Click()

    Call objContabil.Contabil_LabelDoc_Click
    
End Sub

Public Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click
    
End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objFilial As New ClassFilialFornecedor
Dim sContaTela As String
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim sContaMascarada As String
Dim objItem As New ClassTribItemNFAux
Dim sProduto As String
Dim sUM As String
Dim dQuantidade As Double
Dim dQuantidadeConvertida As Double
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim iLinha1 As Integer
Dim iAchou As Integer
Dim sCategoriaProduto As String
Dim objProdutoCategoria As New ClassProdutoCategoria
Dim objCatProdItem As ClassCategoriaProdutoItem
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case ICMS_CREDITO_FRETE
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribFrete.dICMSCredito
            
        Case ICMS_CREDITO_SEGURO
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribSeguro.dICMSCredito
            
        Case ICMS_CREDITO_OUTRASDESP
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribOutras.dICMSCredito
            
        Case ICMS_CREDITO_DESCONTO
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribDesconto.dICMSCredito
            
        Case IPI_CREDITO_FRETE
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribFrete.dIPICredito
            
        Case IPI_CREDITO_SEGURO
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribSeguro.dIPICredito
            
        Case IPI_CREDITO_OUTRASDESP
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribOutras.dIPICredito
            
        Case IPI_CREDITO_DESCONTO
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribDesconto.dIPICredito

        Case ICMS_CREDITO_VALOR
            
            For Each objItem In gobjTribTab.objNFTributacao.colItens
                objMnemonicoValor.colValor.Add objItem.dICMSCredito
            Next

        Case IPI_CREDITO_VALOR
            
            For Each objItem In gobjTribTab.objNFTributacao.colItens
                objMnemonicoValor.colValor.Add objItem.dIPICredito
            Next

        Case IPI_PRODUTO
            
            For Each objItem In gobjTribTab.objNFTributacao.colItens
                objMnemonicoValor.colValor.Add objItem.dIPIValor
            Next

        Case DATA_EMISSAO
            If Len(DataEmissao.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataEmissao.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
        
        Case DATA_ENTRADA
            If Len(DataEntrada.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataEntrada.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
            
        Case DESCONTO1
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case DESCRICAO_ITEM
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_DescProduto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_DescProduto_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case FILIAL1
            If Len(Filial.Text) > 0 Then
                objMnemonicoValor.colValor.Add Filial.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case FORNECEDOR1
            If Len(Fornecedor.Text) > 0 Then
                objMnemonicoValor.colValor.Add Fornecedor.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case NATUREZA_OP
            If Len(NaturezaOp.Text) > 0 Then
                objMnemonicoValor.colValor.Add NaturezaOp.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case NFISCAL1
            If Len(NFiscal.Text) > 0 Then
                objMnemonicoValor.colValor.Add CLng(NFiscal.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case SERIE1
            If Len(Serie.Text) > 0 Then
                objMnemonicoValor.colValor.Add Serie.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case TIPO_NOTA_FISCAL
            If Len(TipoNFiscal.Text) > 0 Then
                objMnemonicoValor.colValor.Add TipoNFiscal.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case UNIDADE_MED
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case VALOR_TOTAL
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case VALOR_UNITARIO
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case ICMS
            If Len(ICMSValor1.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ICMSValor1.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case ICMSSUBST
            If Len(ICMSSubstValor1.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ICMSSubstValor1.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case SUBTOTAL1
            If Len(SubTotal.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(SubTotal.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case VALOR_FRETE
            If Len(ValorFrete.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorFrete.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case VALOR_SEGURO
            If Len(ValorSeguro.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorSeguro.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case VALOR_DESPESAS
            If Len(ValorDespesas.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorDespesas.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case IPI
            If Len(IPIValor1.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(IPIValor1.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case VALOR_DESCONTO
            If Len(ValorDesconto.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorDesconto.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case TOTAL1
            If Len(Total.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(Total.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case ALMOXARIFADO1
            
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                
                If Len(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_AlmoxDist_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_AlmoxDist_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            
            Next
        
        Case PRODUTO1
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_Produto_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case QUANTIDADE1
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
                
                    'Guarda os valores que serão passados como parâmetros em UMEstoque_Conversao
                    sProduto = GridItens.TextMatrix(iLinha, iGrid_Produto_Col)
                    sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
                    
                    'Converte a quantidade para UM padrão estoque
                    lErro = CF("UMEstoque_Conversao", sProduto, sUM, dQuantidade, dQuantidadeConvertida)
                    If lErro <> SUCESSO Then gError 79943
                    
                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
                
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        
        Case PRODUTO_ITENS
            
            For iLinha = 1 To objGridItens.iLinhasExistentes
                
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_Produto_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            
            Next
        
        Case PRODUTO_DIST
            
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                
                If Len(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case QUANT_DIST
                
                For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                    
                    If Len(Trim(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col))) > 0 Then
                        If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
                            objMnemonicoValor.colValor.Add CDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col))
                        Else
                            objMnemonicoValor.colValor.Add 0
                        End If
                    Else
                        objMnemonicoValor.colValor.Add 0
                    End If
                Next
        
        'Conta Contabil de Estoque do par Produto x ALmoxarifado
        Case CTACONTABILEST1
                        
            'Para cada linha do grid distribuição
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                
                'Se iLinha do Grid Distribuição tem a quantidade preenchida
                If Len(Trim(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col))) > 0 Then
                    
                    If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
                            
                        objAlmoxarifado.sNomeReduzido = CStr(objMnemonicoValor.vParam(3))

                        'Lê o Nome Reduzido do Almoxarifado
                        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                        If lErro <> SUCESSO And lErro <> 25060 Then gError 64180

                        'Se não encontrou ===> Erro
                        If lErro = SUCESSO Then

                            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
                            
                            lErro = CF("Produto_Formata", objMnemonicoValor.vParam(1), sProdutoFormatado, iPreenchido)
                            If iPreenchido <> PRODUTO_PREENCHIDO Then gError 64179

                            objEstoqueProduto.sProduto = sProdutoFormatado
                            
                            'Lê a conta contábil do Produto no Almoxarifado
                            lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
                            If lErro <> SUCESSO And lErro <> 49991 Then gError 64181

                            If Len(Trim(objEstoqueProduto.sContaContabil)) > 0 Then

                                sContaMascarada = String(STRING_CONTA, 0)

                                lErro = Mascara_MascararConta(objEstoqueProduto.sContaContabil, sContaMascarada)
                                If lErro <> SUCESSO Then gError 64182

                                objMnemonicoValor.colValor.Add sContaMascarada
                            
                            Else
                                
                                objMnemonicoValor.colValor.Add ""
                            
                            End If
                        
                        Else
                            objMnemonicoValor.colValor.Add ""
                        
                        End If
                    
                    Else
                        objMnemonicoValor.colValor.Add ""
                    
                    End If
                
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            
            Next

        Case CTAFILIALFORNECEDOR1
            
            If Len(Trim(Filial.Text)) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 65757
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 65847
                
                Else
                    sContaTela = ""
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
            
        Case ICMS_CREDITO_TOTAL
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dICMSCredito
        
        Case IPI_CREDITO_TOTAL
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dIPICredito
        
        Case PIS_CREDITO_TOTAL
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dPISCredito
        
        Case COFINS_CREDITO_TOTAL
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dCOFINSCredito
        
        Case FATOR_ITEM_EST
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                If Len(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_UMDist_Col)) > 0 Then
                    If Len(GridItens.TextMatrix(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ItemNFDist_Col), iGrid_UnidadeMed_Col)) > 0 Then
                
                        lErro = CF("Produto_Formata", GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iPreenchido)
                        If lErro <> SUCESSO Then gError 183082
                        
                        'Lê o produto
                        objProduto.sCodigo = sProdutoFormatado
                        lErro = CF("Produto_Le", objProduto)
                        If lErro <> SUCESSO And lErro <> 28030 Then gError 183083
            
                        'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ItemNFDist_Col), iGrid_UnidadeMed_Col), GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_UMDist_Col), dFator)
                        If lErro <> SUCESSO Then gError 183084
                
                        objMnemonicoValor.colValor.Add dFator
                    Else
                        objMnemonicoValor.colValor.Add 1
                    End If
                        
                Else
                    objMnemonicoValor.colValor.Add 1
                End If
            Next
        
        Case PRODUTO_TIPO_EST
        
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                If Len(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183085
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183086
        
                    objMnemonicoValor.colValor.Add objProduto.iTipo
                    
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        Case PRODUTO_TIPO
        
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183152
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183153
        
                    objMnemonicoValor.colValor.Add objProduto.iTipo
                    
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        
        'Conta Contabil de Estoque do par Produto x ALmoxarifado
        Case CTACONTABILDIST
                        
            'Para cada linha do grid distribuição
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                
                'Se iLinha do Grid Distribuição tem a quantidade preenchida
                If Len(Trim(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col))) > 0 Then
                    
                    objAlmoxarifado.sNomeReduzido = GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_AlmoxDist_Col)

                    'Lê o Nome Reduzido do Almoxarifado
                    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25060 Then gError 183173

                    'Se não encontrou ===> Erro
                    If lErro = SUCESSO Then

                        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
                        
                        lErro = CF("Produto_Formata", GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iPreenchido)
                        If iPreenchido <> PRODUTO_PREENCHIDO Then gError 183174

                        objEstoqueProduto.sProduto = sProdutoFormatado
                        
                        'Lê a conta contábil do Produto no Almoxarifado
                        lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
                        If lErro <> SUCESSO And lErro <> 49991 Then gError 183175

                        If Len(Trim(objEstoqueProduto.sContaContabil)) > 0 Then

                            sContaMascarada = String(STRING_CONTA, 0)

                            lErro = Mascara_MascararConta(objEstoqueProduto.sContaContabil, sContaMascarada)
                            If lErro <> SUCESSO Then gError 183176

                            objMnemonicoValor.colValor.Add sContaMascarada
                        
                        Else
                            
                            objMnemonicoValor.colValor.Add ""
                        
                        End If
                        
                    Else
                        objMnemonicoValor.colValor.Add ""
                    
                    End If
                
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            
            Next
        
        Case VALOR_UNIT_DIST
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                
                For iLinha1 = 1 To objGridItens.iLinhasExistentes
                
                    If GridItens.TextMatrix(iLinha1, iGrid_Produto_Col) = GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col) Then
                    
                        If Len(GridItens.TextMatrix(iLinha1, iGrid_ValorUnitario_Col)) > 0 Then
                            objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha1, iGrid_ValorUnitario_Col))
                        Else
                            objMnemonicoValor.colValor.Add 0
                        End If
                    
                        Exit For
                    
                    End If
                
                Next
                
            Next
        
        Case ICMS_CRED_VALOR_DIST
            
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
            
                lErro = CF("Produto_Formata", GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iPreenchido)
                If iPreenchido <> PRODUTO_PREENCHIDO Then gError 183181

                iAchou = 0
                
                For Each objItem In gobjTribTab.objNFTributacao.colItens
                
                    If objItem.sProduto = sProdutoFormatado Then
                
                        If StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col)) > 0 And _
                           StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantItemNFDist_Col)) > 0 Then
                
                            objMnemonicoValor.colValor.Add objItem.dICMSCredito * StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col)) / StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantItemNFDist_Col))
                            iAchou = 1
                            Exit For
                        
                        End If
                    
                    End If
                    
                Next
                
                If iAchou = 0 Then
                    objMnemonicoValor.colValor.Add 0
                End If
                
            Next
        
        Case IPI_CRED_VALOR_DIST
            
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
            
                lErro = CF("Produto_Formata", GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iPreenchido)
                If iPreenchido <> PRODUTO_PREENCHIDO Then gError 183182

                iAchou = 0
                
                For Each objItem In gobjTribTab.objNFTributacao.colItens
                
                    If StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col)) > 0 And _
                       StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantItemNFDist_Col)) > 0 Then
            
                        objMnemonicoValor.colValor.Add objItem.dIPICredito * StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col)) / StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantItemNFDist_Col))
                        iAchou = 1
                        Exit For
                    
                    End If
                    
                Next
                    
                If iAchou = 0 Then
                    objMnemonicoValor.colValor.Add 0
                End If
                
            Next
                    

        Case IPI_PRODUTO_DIST
            
            
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
            
                lErro = CF("Produto_Formata", GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iPreenchido)
                If iPreenchido <> PRODUTO_PREENCHIDO Then gError 183183

                iAchou = 0
                
                For Each objItem In gobjTribTab.objNFTributacao.colItens
                
                    If StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col)) > 0 And _
                       StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantItemNFDist_Col)) > 0 Then
            
                        objMnemonicoValor.colValor.Add objItem.dIPIValor * StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col)) / StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantItemNFDist_Col))
                        iAchou = 1
                        Exit For
                    
                    End If
                    
                Next
                    
                If iAchou = 0 Then
                    objMnemonicoValor.colValor.Add 0
                End If
                
            Next
            
        Case QUANT_DIST1
                
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                
                If Len(Trim(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col))) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        Case CONTROLA_EST
        
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183424
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183425
        
                    objMnemonicoValor.colValor.Add objProduto.iControleEstoque
                    
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        Case CONTA_DESP_EST_FORN
            
            If Len(Trim(Filial.Text)) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 183426
                
                If objFilial.sContaDespesa <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaDespesa, sContaTela)
                    If lErro <> SUCESSO Then gError 183427
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
        
        Case CONTACONTABILAPLIC
        
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183530
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183531
        
                    If objProduto.sContaContabil <> "" Then
                    
                        lErro = Mascara_RetornaContaTela(objProduto.sContaContabil, sContaTela)
                        If lErro <> SUCESSO Then gError 183810
                    
                    Else
                    
                        sContaTela = ""
                        
                    End If
        
                    objMnemonicoValor.colValor.Add sContaTela
                    
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case CONTACONTABILCATPROD
                
            For iLinha = 1 To objGridItens.iLinhasExistentes
                
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183552
                
                    objProdutoCategoria.sCategoria = objMnemonicoValor.vParam(1)
                    
                    objProdutoCategoria.sProduto = sProdutoFormatado
                    
                    Set objCatProdItem = New ClassCategoriaProdutoItem
                    
                    lErro = CF("ProdutoCategoria_Le_Item", objProdutoCategoria, objCatProdItem)
                    If lErro <> SUCESSO And lErro <> 183548 Then gError 183553
                    
                    If lErro = SUCESSO Then
                        
                        Set objPlanoConta = New ClassPlanoConta
                    
                        'le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta
                        lErro = CF("PlanoConta_Le_ContaSimples", objCatProdItem.dvalor1, objPlanoConta)
                        If lErro <> SUCESSO And lErro <> 5451 Then gError 183554
                    
                        If lErro = SUCESSO Then
                    
                            If objPlanoConta.sConta <> "" Then
                            
                                lErro = Mascara_RetornaContaTela(objPlanoConta.sConta, sContaTela)
                                If lErro <> SUCESSO Then gError 183811
                            
                            Else
                            
                                sContaTela = ""
                                
                            End If
                    
                            objMnemonicoValor.colValor.Add sContaTela
                            
                        Else
                    
                            objMnemonicoValor.colValor.Add ""
                    
                        End If
                    
                    Else
                        objMnemonicoValor.colValor.Add ""
                    End If
                    
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
                
            Next
        
        Case Else
            lErro = gobjTribTab.Calcula_Mnemonico(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 39649

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39649
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case 65542, 65592, 65702, 65703, 65757, 65847, 183072 To 183076, 183143, 183144, 183424 To 183427, 183530, 183531, 183552, 183553, 183554, 183810, 183811
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156758)

    End Select

    Exit Function

End Function

'Tratamento dos Grids

'GridItens
Public Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGridItens, iAlterado)

End Sub

Public Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGridItens)

End Sub

Public Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_LeaveCell()

    Call Saida_Celula(objGridItens)

End Sub

Public Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)
    
End Sub

Public Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer, Optional ByVal bExibeMsgDel As Boolean = True)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer, lErro As Long
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iLinha As Integer
Dim sProduto As String
Dim bTeste As Boolean
Dim objNFiscal As New ClassNFiscal
Dim objItemNF As New ClassItemNF

On Error GoTo Erro_GridItens_KeyDown

    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iItemAtual = GridItens.Row
    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
    
    Call CF2(Me, "NFEntCom_Move_GridItem_Memoria", objNFiscal, objItemNF, iItemAtual)
    
    If bExibeMsgDel Then
        Call Grid_Trata_Tecla1(KeyCode, objGridItens)
    Else
        Call Grid_Trata_Tecla3(KeyCode, objGridItens)
    End If
    
    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then

        lErro = SubTotal_Calcula()
        If lErro <> SUCESSO Then gError 65998
        
''*** incluidos p/tratamento de tributacao *******************************
        
        lErro = gobjTribTab.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 65902
        
         Call gobjTelaProjetoInfo.Remove_Linha(iItemAtual)
        
'*** fim tributacao
    
        Call ValorTotal_Calcula
        
        If Not gbCarregandoTela Then
            lErro = gobjTribTab.Pesos_Recalcula()
            If lErro <> SUCESSO Then gError 181260
        End If
        
        lErro = CF2(Me, "NFEntCom_Calcula_PesoKg_Total")
        If lErro <> SUCESSO Then gError 181260

        'distribuicao
        lErro = gobjDistribuicao.Exclusao_Item_GridDist(iItemAtual)
        If lErro <> SUCESSO Then gError 89542

        'Remove do Grid de Itens Pedido de Compras os Itens que possuem o mesmo Produto
        For iIndice = objGridItensPC.iLinhasExistentes To 1 Step -1
        
            bTeste = True
            Call CF2(Me, "NFEntCom_Produto_Igual_Customizado7", objItemNF.objInfoUsu, iIndice, bTeste)

            If GridItensPC.TextMatrix(iIndice, iGrid_Prod_Col) = sProduto And bTeste Then
                Call Grid_Exclui_Linha(objGridItensPC, iIndice)
            End If
        Next
        
        'Remove do Grid de Itens de Requisição de Compras os Itens que possuem o mesmo Produto
        For iIndice = objGridItensRC.iLinhasExistentes To 1 Step -1
            If GridItensRC.TextMatrix(iIndice, iGrid_ProdutoRC_Col) = sProduto Then
                Call Grid_Exclui_Linha(objGridItensRC, iIndice)
            End If
        Next
        
        'Formata o Produto
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 65999
        
        'Remove da coleção de Itens Pedidos de Compras
        For iIndice = gcolItemPedCompraInfo.Count To 1 Step -1
            
            bTeste = True
            Call CF2(Me, "NFEntCom_Produto_Igual_Customizado5", gcolItemPedCompraInfo.Item(iIndice).objInfoUsu, objItemNF.objInfoUsu, bTeste)
            
            If gcolItemPedCompraInfo.Item(iIndice).sProduto = sProdutoFormatado And bTeste Then
                gcolItemPedCompraInfo.Remove (iIndice)
            End If
        
        Next
            
        'Remove da coleção de Itens de Requisição de Compras
        For iIndice = gcolItemReqCompraInfo.Count To 1 Step -1
            
            If gcolItemReqCompraInfo.Item(iIndice).sProduto = sProdutoFormatado Then
                gcolItemReqCompraInfo.Remove (iIndice)
            End If
        
        Next
    
        'Verifica se há Pedidos de Compras selecionados que não possuem itens no GridPC
        For iIndice = 0 To PedidosCompra.ListCount - 1
            If PedidosCompra.Selected(iIndice) = True Then
                For iLinha = 1 To objGridItensPC.iLinhasExistentes
                    If GridItensPC.TextMatrix(iLinha, iGrid_PedCompra_Col) = PedidosCompra.List(iIndice) Then
                        Exit For
                    End If
                Next
                If iLinha > objGridItensPC.iLinhasExistentes Then
                    PedidosCompra.Selected(iIndice) = False
                End If
            End If
        Next
        
        'Verifica se há Requisições de Compras selecionadas que não possuem itens no GridRC
        For iIndice = 0 To RequisicoesCompra.ListCount - 1
            If RequisicoesCompra.Selected(iIndice) = True Then
                For iLinha = 1 To objGridItensRC.iLinhasExistentes
                    If GridItensRC.TextMatrix(iLinha, iGrid_ReqCompra_Col) = RequisicoesCompra.List(iIndice) Then
                        Exit For
                    End If
                Next
            
                If iLinha > objGridItensRC.iLinhasExistentes Then
                    RequisicoesCompra.Selected(iIndice) = False
                End If
                
            End If
        
        Next
        
    End If

    Exit Sub
    
Erro_GridItens_KeyDown:

    Select Case gErr

        Case 65902, 65998, 65999, 83275, 89542, 181260

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156759)

    End Select

    Exit Sub
    
End Sub

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Public Sub GridRastro_Click()
'Rastreamento

    Call gobjRastreamento.GridRastro_Click

End Sub

Public Sub GridRastro_EnterCell()
'Rastreamento

    Call gobjRastreamento.GridRastro_EnterCell

End Sub

Public Sub GridRastro_GotFocus()
'Rastreamento

    Call gobjRastreamento.GridRastro_GotFocus

End Sub

Public Sub GridRastro_KeyPress(KeyAscii As Integer)
'Rastreamento

    Call gobjRastreamento.GridRastro_KeyPress(KeyAscii)

End Sub

Public Sub GridRastro_LeaveCell()
'Rastreamento

    Call gobjRastreamento.GridRastro_LeaveCell
    
End Sub

Public Sub GridRastro_Validate(Cancel As Boolean)
'Rastreamento

    Call gobjRastreamento.GridRastro_Validate(Cancel)

End Sub

Public Sub GridRastro_RowColChange()
'Rastreamento

    Call gobjRastreamento.GridRastro_RowColChange

End Sub

Public Sub GridRastro_Scroll()
'Rastreamento

    Call gobjRastreamento.GridRastro_Scroll

End Sub

Public Sub GridRastro_KeyDown(KeyCode As Integer, Shift As Integer)
'Rastreamento

    Call gobjRastreamento.GridRastro_KeyDown(KeyCode, Shift)

End Sub

Public Sub GridDist_Click()
'distribuicao
    
    Call gobjDistribuicao.GridDist_Click

End Sub

Public Sub GridDist_EnterCell()
'distribuicao
    
    Call gobjDistribuicao.GridDist_EnterCell

End Sub

Public Sub GridDist_GotFocus()
'distribuicao
    
    Call gobjDistribuicao.GridDist_GotFocus

End Sub

Public Sub GridDist_KeyPress(KeyAscii As Integer)
'distribuicao
    
    Call gobjDistribuicao.GridDist_KeyPress(KeyAscii)

End Sub

Public Sub GridDist_LeaveCell()
'distribuicao
    
    Call gobjDistribuicao.GridDist_LeaveCell

End Sub

Public Sub GridDist_Validate(Cancel As Boolean)
'distribuicao
    
    Call gobjDistribuicao.GridDist_Validate(Cancel)
    
End Sub

Public Sub GridDist_RowColChange()
'distribuicao
    
    Call gobjDistribuicao.GridDist_RowColChange

End Sub

Public Sub GridDist_KeyDown(KeyCode As Integer, Shift As Integer)
'distribuicao
    
    Call gobjDistribuicao.GridDist_KeyDown(KeyCode, Shift)
    
End Sub

Public Sub GridDist_Scroll()
'distribuicao
    
    Call gobjDistribuicao.GridDist_Scroll

End Sub

Public Sub Produto_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UnidadeMed_Change()
        
        iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMed_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub UnidadeMed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Quantidade_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorUnitario_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorUnitario_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub ValorUnitario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub ValorUnitario_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = ValorUnitario
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PercentDesc_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentDesc_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub PercentDesc_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub PercentDesc_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PercentDesc
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Ccl_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Ccl_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Ccl_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Desconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Desconto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DescricaoItem_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub DescricaoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DescricaoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DescricaoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescricaoItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'GridItensPC
Public Sub GridItensPC_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItensPC, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensPC, iAlterado)
    End If

End Sub

Public Sub GridItensPC_EnterCell()

    Call Grid_Entrada_Celula(objGridItensPC, iAlterado)

End Sub

Public Sub GridItensPC_GotFocus()

    Call Grid_Recebe_Foco(objGridItensPC)

End Sub

Public Sub GridItensPC_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItensPC, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensPC, iAlterado)
    End If

End Sub

Public Sub GridItensPC_LeaveCell()

    Call Saida_Celula(objGridItensPC)

End Sub

Public Sub GridItensPC_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItensPC)
    
End Sub

Public Sub GridItensPC_RowColChange()

    Call Grid_RowColChange(objGridItensPC)

End Sub

Public Sub GridItensPC_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridItensPC_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridItensPC)

    Exit Sub
    
Erro_GridItensPC_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156760)

    End Select

    Exit Sub
    
End Sub

Public Sub GridItensPC_Scroll()

    Call Grid_Scroll(objGridItensPC)

End Sub

Public Sub QuantRecebidaPC_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantRecebidaPC_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensPC)

End Sub

Public Sub QuantRecebidaPC_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensPC)

End Sub

Public Sub QuantRecebidaPC_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItensPC.objControle = QuantRecebidaPC
    lErro = Grid_Campo_Libera_Foco(objGridItensPC)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'GridItensRC
Public Sub GridItensRC_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItensRC, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensRC, iAlterado)
    End If

End Sub

Public Sub GridItensRC_EnterCell()

    Call Grid_Entrada_Celula(objGridItensRC, iAlterado)

End Sub

Public Sub GridItensRC_GotFocus()

    Call Grid_Recebe_Foco(objGridItensRC)

End Sub

Public Sub GridItensRC_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItensRC, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensRC, iAlterado)
    End If

End Sub

Public Sub GridItensRC_LeaveCell()

    Call Saida_Celula(objGridItensRC)

End Sub

Public Sub GridItensRC_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItensRC)
    
End Sub

Public Sub GridItensRC_RowColChange()

    Call Grid_RowColChange(objGridItensRC)

End Sub

Public Sub GridItensRC_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridItensRC_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridItensRC)

    Exit Sub
    
Erro_GridItensRC_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156761)

    End Select

    Exit Sub
    
End Sub

Public Sub GridItensRC_Scroll()

    Call Grid_Scroll(objGridItensRC)

End Sub

Public Sub QuantRecebidaRC_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantRecebidaRC_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensRC)

End Sub

Public Sub QuantRecebidaRC_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensRC)

End Sub

Public Sub QuantRecebidaRC_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItensRC.objControle = QuantRecebidaRC
    lErro = Grid_Campo_Libera_Foco(objGridItensRC)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub BotaoConsultaNFPag_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoConsultaNFPag_Click

    'Verifica se todos os campos necessários para se efetuar a consulta foram preenchidos
    lErro = Critica_CamposNecessarios_ConsultaTitulo()
    If lErro <> SUCESSO Then gError 79767
    
    'Guarda no objNFiscal os dados necessários para consultar o título
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.iTipoDocInfo = Codigo_Extrai(TipoNFiscal.Text)
    objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.dtDataEntrada = StrParaDate(DataEntrada.Text)
    
    'Guarda em objFornecedor o nome reduzido do Fornecedor
    objFornecedor.sNomeReduzido = Fornecedor.Text
    
    'Lê o código do Fornecedor a partir do nome reduzido obtido na tela
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 79768
    
    'Se não encontrou o fornecedor = > erro
    If lErro = 6681 Then gError 79769
    
    'Guarda no objNFiscal o código do fornecedor
    objNFiscal.lFornecedor = objFornecedor.lCodigo
    
    'Lê o NumIntDocCPR da NFiscal e exibe o documento gerado no CPR por essa nota
    lErro = CF("NFiscal_Consulta_DocCPR", objNFiscal)
    If lErro <> SUCESSO And lErro <> 79717 Then gError 79770
    
    'Se não encontrou a nota => erro
    If lErro = 79717 Then gError 79771
    
    Exit Sub
    
Erro_BotaoConsultaNFPag_Click:

    Select Case gErr
    
        Case 79767, 79770, 79768
        
        Case 79771
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA1", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada)
        
        Case 79769
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156762)
    
    End Select
    
    Exit Sub

End Sub

Private Function Critica_CamposNecessarios_ConsultaTitulo() As Long
'Verifica se os campos necessários para encontrar consultar um título gerado por uma NFFatEntrada foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposNecessarios_ConsultaTitulo

    'Verifica se o Fornecedor foi preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 79772
    
    'Verifica se a filial do Fornecedor foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 79773
    
    'Verifica se a Série foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 79774
    
    'Verifica se o Número da Nota foi preenchido
    If Len(Trim(NFiscal.Text)) = 0 Then gError 79775
    
    'Verifica se a data de emissão da nota foi preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 79776
    
    'Verifica se a data de entrada da nota foi preenchida
    If Len(Trim(DataEntrada.ClipText)) = 0 Then gError 79777
    
    'Verifica se o tipo da Nota fiscal foi selecionado
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 79778
    
    Critica_CamposNecessarios_ConsultaTitulo = SUCESSO
    
    Exit Function

Erro_Critica_CamposNecessarios_ConsultaTitulo:

    Critica_CamposNecessarios_ConsultaTitulo = gErr
    
    Select Case gErr
    
        Case 79772
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 79773
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 79774
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 79775
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
        
        Case 79776
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 79777
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_NAO_PREENCHIDA", gErr)
        
        Case 79778
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156763)
    
    End Select
        
    Exit Function
        
End Function

Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

Public Sub ComboPedidoCompras_Click()

Dim lErro As Long

On Error GoTo Erro_ComboPedidoCompras_Click

    Moeda.ListIndex = -1
    
    gbLimpaTaxa = True
    
    'Preenche o grid
    lErro = Preenche_GridItensPC
    If lErro <> SUCESSO Then gError 108992
    
    Exit Sub
    
Erro_ComboPedidoCompras_Click:

    Select Case gErr
        
        Case 108992
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156767)
    
    End Select
    
End Sub

Function Carrega_Moeda() As Long

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection
Dim iPosMoedaReal As Integer
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError 103371
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 103372
    
    For Each objMoeda In colMoedas
    
        Moeda.AddItem objMoeda.sNome
        Moeda.ItemData(iIndice) = objMoeda.iCodigo
        
        iIndice = iIndice + 1
    
    Next
    
    Moeda.ListIndex = -1

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 103371
        
        Case 103372
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156768)
    
    End Select

End Function

Public Sub Taxa_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim dTaxa As Double

On Error GoTo Erro_Taxa_Validate

    'Verifica se foi preenchido
    If Len(Trim(Taxa.Text)) > 0 Then

        'Criica se é Valor não negativo
        lErro = Valor_NaoNegativo_Critica(Taxa.Text)
        If lErro <> SUCESSO Then gError 108993

        Taxa.Text = Format(Taxa.Text, FORMATO_TAXA_CONVERSAO_MOEDA)
        
        gbLimpaTaxa = False
        
        'Preenche o grid
        lErro = Preenche_GridItensPC
        If lErro <> SUCESSO Then gError 108994

    End If

    Exit Sub

Erro_Taxa_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 108993, 108994

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156769)

    End Select

End Sub

Private Function CalculaPrecoUnitario_GridItens() As Long
'Calcula o preco unitário baseado no grid de ItensPC

Dim lErro As Long
Dim dValorItensPC As Double
Dim dQuantidadeItensPC As Double
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim bAchou As Boolean
Dim objItemPedCompraInfo As ClassItemPedCompraInfo
Dim dValorDesconto As Double
Dim dTaxa As Double
Dim dCotacao As Double
Dim dPrecoUnitMaisRecente As Double
Dim objPedCompra As ClassPedidoCompras
Dim dtDataMaisRecente As Date
Dim bTeste As Boolean

On Error GoTo Erro_CalculaPrecoUnitario_GridItens

    'Descobre o Indice do produto alterado no grid de itens
    For iIndice1 = 1 To objGridItens.iLinhasExistentes
        
        sProduto = GridItens.TextMatrix(iIndice1, iGrid_Produto_Col)
        dValorItensPC = 0
        dQuantidadeItensPC = 0
        
        dtDataMaisRecente = DATA_NULA
        
        'Para cada Item da coleção de ItensPC
        For Each objItemPedCompraInfo In gcolItemPedCompraInfo
    
            lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iPreenchido)
            If lErro <> SUCESSO Then gError 114507
            
            bTeste = True
            lErro = CF2(Me, "NFEntCom_Produto_Igual_Customizado", objItemPedCompraInfo.objInfoUsu, iIndice1, bTeste)
            If lErro <> SUCESSO Then gError 181284
            
            If sProdutoFormatado = objItemPedCompraInfo.sProduto And bTeste Then
            
                bAchou = True
            
                objProduto.sCodigo = sProdutoFormatado
        
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 114508
        
                If lErro = 28030 Then gError 114509
        
                'Converte a UM de GridItensPC para a UM do GridItens
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemPedCompraInfo.sUM, GridItens.TextMatrix(iIndice1, iGrid_UnidadeMed_Col), dFator)
                If lErro <> SUCESSO Then gError 114510
        
                If objItemPedCompraInfo.dTaxa > 0 Then
                    dValorItensPC = dValorItensPC + objItemPedCompraInfo.dQuantRecebida * objItemPedCompraInfo.dPrecoUnitario * objItemPedCompraInfo.dTaxa
                    dValorDesconto = objItemPedCompraInfo.dValorDesconto * objItemPedCompraInfo.dTaxa
                Else
                    dValorItensPC = dValorItensPC + objItemPedCompraInfo.dQuantRecebida * objItemPedCompraInfo.dPrecoUnitario
                    dValorDesconto = objItemPedCompraInfo.dValorDesconto
                End If
                
                dQuantidadeItensPC = dQuantidadeItensPC + objItemPedCompraInfo.dQuantRecebida * dFator
            
                If gobjEST.iFormaCalcPrecoEntCom = FORMA_CALC_PESO_ENT_COM_ULT Then
                    For Each objPedCompra In gcolPedidoCompra
                        If objItemPedCompraInfo.lPedCompra = objPedCompra.lCodigo Then
                            If dtDataMaisRecente = DATA_NULA Or dtDataMaisRecente < objPedCompra.dtDataEnvio Then
                                dtDataMaisRecente = objPedCompra.dtDataEnvio
                                If objItemPedCompraInfo.dTaxa > 0 Then
                                    dPrecoUnitMaisRecente = objItemPedCompraInfo.dPrecoUnitario * objItemPedCompraInfo.dTaxa
                                Else
                                    dPrecoUnitMaisRecente = objItemPedCompraInfo.dPrecoUnitario
                                End If
                            End If
                            Exit For
                        End If
                    Next
                End If
                
            End If
    
        Next
        
        If bAchou Then
        
            'Se alterou a quantidade
            If Abs(StrParaDbl(GridItens.TextMatrix(iIndice1, iGrid_Quantidade_Col)) - dQuantidadeItensPC) > QTDE_ESTOQUE_DELTA Then
        
                GridItens.TextMatrix(iIndice1, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidadeItensPC)
                
                lErro = CF2(Me, "NFEntCom_Calcula_PesoKg", iIndice1)
                If lErro <> SUCESSO Then gError 181265
                
            End If
                
            lErro = CF2(Me, "NFEntCom_PrecoTotal_Calcula", iIndice1, dQuantidadeItensPC)
            If lErro <> SUCESSO Then gError 181266
        
            If gobjEST.iFormaCalcPrecoEntCom = FORMA_CALC_PESO_ENT_COM_ULT Then
                dValorItensPC = dQuantidadeItensPC * dPrecoUnitMaisRecente
            End If
            
            If dValorItensPC <> 0 Then
                GridItens.TextMatrix(iIndice1, iGrid_PrecoTotal_Col) = Format(dValorItensPC - dValorDesconto, "STANDARD")
                GridItens.TextMatrix(iIndice1, iGrid_ValorUnitario_Col) = Format(dValorItensPC / dQuantidadeItensPC, IIf(gobjCOM.sFormatoPrecoUnitario <> "", gobjCOM.sFormatoPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO))
                GridItens.TextMatrix(iIndice1, iGrid_Desconto_Col) = Format(dValorDesconto, "STANDARD")
                GridItens.TextMatrix(iIndice1, iGrid_PercDesc_Col) = Format(dValorDesconto / (dValorItensPC), "Percent")
            Else
                GridItens.TextMatrix(iIndice1, iGrid_PrecoTotal_Col) = ""
                GridItens.TextMatrix(iIndice1, iGrid_ValorUnitario_Col) = ""
                GridItens.TextMatrix(iIndice1, iGrid_Desconto_Col) = ""
                GridItens.TextMatrix(iIndice1, iGrid_PercDesc_Col) = ""
            End If
            
            'recalcula os valores de desconto, percentual de desconto e valor total
            lErro = Calcula_Valores(iIndice1)
            If lErro <> SUCESSO Then gError 65752
            
            '*** incluidos p/tratamento de tributacao ****************
            lErro = gobjTribTab.Alteracao_Item_Grid(iIndice1)
            If lErro <> SUCESSO Then gError 65752
            '*** fim tributacao
        
        End If
                
        bAchou = False
        
    Next

    Call SubTotal_Calcula
    
    'Recalcula o total da nota fiscal
    Call ValorTotal_Calcula
    
    CalculaPrecoUnitario_GridItens = SUCESSO
    
    Exit Function
    
Erro_CalculaPrecoUnitario_GridItens:

    CalculaPrecoUnitario_GridItens = gErr
    
    Select Case gErr
    
        Case 114507, 114508, 114510, 181265, 181266, 181284
        
        Case 114509
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156770)
    
    End Select

End Function

Private Sub Fornecedor_Preenche()
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134061

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134061

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156771)

    End Select
    
    Exit Sub

End Sub

'@@@@@ para ser customizada
Function NFEntCom_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Valor Unitário")
    objGridInt.colColuna.Add ("% Desc.")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (objCT.Produto.Name)
    objGridInt.colCampo.Add (objCT.DescricaoItem.Name)
    objGridInt.colCampo.Add (objCT.UnidadeMed.Name)
    objGridInt.colCampo.Add (objCT.Quantidade.Name)
    objGridInt.colCampo.Add (objCT.Ccl.Name)
    objGridInt.colCampo.Add (objCT.ValorUnitario.Name)
    objGridInt.colCampo.Add (objCT.PercentDesc.Name)
    objGridInt.colCampo.Add (objCT.Desconto.Name)
    objGridInt.colCampo.Add (objCT.ValorTotal.Name)
    objGridInt.colCampo.Add (objCT.PrecoTotalB.Name)
    
    'Se é permitido que o valor unitário do ItemNF é diferente do valor unitário do ItemPC
    If gobjCOM.iNFDiferentePC = NFISCAL_NAO_ACEITA_DIFERENCA_PC Then
        objCT.ValorUnitario.Enabled = False
    Else
        objCT.ValorUnitario.Enabled = True
    End If
    
    'Colunas da Grid
    objCT.iGrid_Produto_Col = 1
    objCT.iGrid_DescProduto_Col = 2
    objCT.iGrid_UnidadeMed_Col = 3
    objCT.iGrid_Quantidade_Col = 4
'distribuicao
'    iGrid_Almoxarifado_Col = 5
    objCT.iGrid_Ccl_Col = 5
    objCT.iGrid_ValorUnitario_Col = 6
    objCT.iGrid_PercDesc_Col = 7
    objCT.iGrid_Desconto_Col = 8
    objCT.iGrid_PrecoTotal_Col = 9
    objCT.iGrid_PrecoTotalB_Col = 10

    'Grid do GridInterno
    objGridInt.objGrid = objCT.GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_NF + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    objCT.GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NFEntCom_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

'@@@@@ para ser customizada
Public Function NFEntCom_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long

    NFEntCom_Saida_Celula_GridItens = SUCESSO

End Function

'@@@@@ para ser customizada
Function NFEntCom_Preenche_GridItens(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long
    
    NFEntCom_Preenche_GridItens = SUCESSO
     
End Function

'@@@@@ para ser customizada
Function NFEntCom_Traz_NFiscal_Tela(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long

    NFEntCom_Traz_NFiscal_Tela = SUCESSO
     
End Function

'@@@@@ para ser customizada
Function NFEntCom_Valida(ByVal objCT As Object) As Long

    NFEntCom_Valida = SUCESSO
     
End Function

'@@@@@ para ser customizada
Function NFEntCom_Move_GridItem_Memoria(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal, ByVal objItemNF As ClassItemNF, ByVal iIndice As Integer) As Long
    NFEntCom_Move_GridItem_Memoria = SUCESSO
End Function

'@@@@@ para ser customizada
Function NFEntCom_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iCaminho As Integer, ByVal sProdutoFormatado As String, ByVal iProdutoPreenchido As Integer) As Long
    NFEntCom_Rotina_Grid_Enable = SUCESSO
End Function

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long, sID As String
Dim objNFiscal As New ClassNFiscal
Dim objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_Anotacao_Extrai

    'Lê os dados da Tela NFiscal
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 130404

    If objNFiscal.lNumNotaFiscal <> 0 And objNFiscal.sSerie <> "" And objNFiscal.dtDataEmissao <> DATA_NULA Then
    
        lErro = CF("NFiscal_Le_NumFornCli", objNFiscal)
        If lErro <> SUCESSO And lErro <> 35279 Then gError 130405
        
        If lErro = SUCESSO Then sID = objNFiscal.lNumIntDoc
        
    End If
    
    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_NFISCAL
    If sID <> "" Then
        objAnotacao.sID = sID
    Else
        objAnotacao.sID = ""
        If Not (gobjAnotacao Is Nothing) Then
            objAnotacao.sTextoCompleto = gobjAnotacao.sTextoCompleto
            objAnotacao.sTitulo = gobjAnotacao.sTitulo
        End If
    End If
    
    Exit Sub
     
Erro_Anotacao_Extrai:

    Select Case gErr
          
        Case 130404, 130405
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156772)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Anotacao_Preenche(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Preenche

    'guarda o texto digitado
    Set gobjAnotacao = objAnotacao
        
    Exit Sub
     
Erro_Anotacao_Preenche:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156773)
     
    End Select
     
    Exit Sub

End Sub

Public Sub BotaoSerie_Click()
'Chama a tela de Lote de Rastreamento

Dim lErro As Long

On Error GoTo Erro_BotaoSerie_Click
    
    Call gobjRastreamento.BotaoSerie_Click
                    
    Exit Sub

Erro_BotaoSerie_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141800)
    
    End Select
    
    Exit Sub

End Sub

Public Function NFEntCom_ProdutoLinha_Preenche(ByVal objCT As Object, objProduto As ClassProduto) As Long
    NFEntCom_ProdutoLinha_Preenche = SUCESSO
End Function

Public Function NFEntCom_PrecoTotal_Calcula(ByVal objCT As Object, ByVal iLinha As Integer, dQuantidade As Double) As Long
    NFEntCom_PrecoTotal_Calcula = SUCESSO
End Function

Public Function NFEntCom_Calcula_PesoKg(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    NFEntCom_Calcula_PesoKg = SUCESSO
End Function

Public Function NFEntCom_Calcula_PesoKg_Total(ByVal objCT As Object) As Long
    NFEntCom_Calcula_PesoKg_Total = SUCESSO
End Function

Function NFEntCom_PreencheLinha_GridItens(ByVal objCT As Object, ByVal objItemPC As ClassItemPedCompra, ByVal iLinha As Integer) As Long
    NFEntCom_PreencheLinha_GridItens = SUCESSO
End Function

Function NFEntCom_Produto_Igual_Customizado(ByVal objCT As Object, ByVal objItemPCUsu As Object, ByVal iLinha As Integer, bTeste As Boolean) As Long
    NFEntCom_Produto_Igual_Customizado = SUCESSO
End Function

Function NFEntCom_PreencheLinha_GridPC(ByVal objCT As Object, ByVal objItemPCUsu As Object, ByVal iLinha As Integer) As Long
    NFEntCom_PreencheLinha_GridPC = SUCESSO
End Function

Function NFEntCom_Produto_Igual_Customizado2(ByVal objCT As Object, ByVal iLinhaItem As Integer, ByVal iLinhaPC As Integer, bTeste As Boolean) As Long
    NFEntCom_Produto_Igual_Customizado2 = SUCESSO
End Function

Function NFEntCom_Produto_Igual_Customizado3(ByVal objCT As Object, ByVal objItemPCUsu1 As Object, ByVal objItemPCUsu2 As Object, bTeste As Boolean) As Long
    NFEntCom_Produto_Igual_Customizado3 = SUCESSO
End Function

Function NFEntCom_Produto_Igual_Customizado4(ByVal objCT As Object, ByVal objItemPCFM As Object, ByVal iLinha As Integer, bTeste As Boolean) As Long
    NFEntCom_Produto_Igual_Customizado4 = SUCESSO
End Function

Function NFEntCom_Produto_Igual_Customizado5(ByVal objCT As Object, ByVal objItemPCUsu As Object, ByVal objItemNFUsu As Object, bTeste As Boolean) As Long
    NFEntCom_Produto_Igual_Customizado5 = SUCESSO
End Function

Function NFEntCom_Produto_Igual_Customizado6(ByVal objCT As Object, ByVal objItemNFUsu As Object, ByVal iLinha As Integer, bTeste As Boolean) As Long
    NFEntCom_Produto_Igual_Customizado6 = SUCESSO
End Function

Public Function NFEntCom_Tranfere_Dados_ItemPC_Para_ItemNF(ByVal objCT As Object, ByVal objItemNF As ClassItemNF, ByVal objItemPC As ClassItemPedCompra) As Long
    NFEntCom_Tranfere_Dados_ItemPC_Para_ItemNF = SUCESSO
End Function

Function NFEntCom_Le_PedInfo_Customizado(ByVal objCT As Object, ByVal colItemPedCompraInfo As Collection) As Long
    NFEntCom_Le_PedInfo_Customizado = SUCESSO
End Function

Function NFEntCom_Produto_Igual_Customizado7(ByVal objCT As Object, ByVal objItemNFUsu As Object, ByVal iLinha As Integer, bTeste As Boolean) As Long
    NFEntCom_Produto_Igual_Customizado7 = SUCESSO
End Function

Public Function BotaoProjetos_Click() As Long
        Call gobjTelaProjetoInfo.BotaoProjetos_Click
End Function

Public Function LabelProjeto_Click() As Long
    Call gobjTelaProjetoInfo.LabelProjeto_Click
End Function

Public Function Projeto_GotFocus() As Long
    Call MaskEdBox_TrataGotFocus(Projeto, iAlterado)
End Function

Public Function Projeto_Change() As Long
    iAlterado = REGISTRO_ALTERADO
End Function

Public Function Projeto_Validate(Cancel As Boolean) As Long
    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
End Function

Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensPRJCR As ClassItensPRJCR

On Error GoTo Erro_Obter_ItensPRJCR

    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItensPRJCR = New ClassItensPRJCR
        
        objItensPRJCR.sItem = CStr(iIndice)
        objItensPRJCR.sDescricao = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)
        objItensPRJCR.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) & SEPARADOR & GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)

        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objItensPRJCR.dValorOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))

        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NF
        
        colItensPRJ.Add objItensPRJCR

    Next

    Call gobjTribTab.Preenche_Itens_Projeto_NF(colItensPRJ)

    Obter_ItensPRJCR = SUCESSO

    Exit Function

Erro_Obter_ItensPRJCR:

    Obter_ItensPRJCR = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157006)

    End Select

    Exit Function
    
End Function

Public Function Preenche_Itens_Projeto_Aux(ByVal colItensPRJ As Collection, ByVal sItem As String, ByVal sDescricao As String, ByVal dValor As Double, ByVal iTipo As Integer, Optional ByVal iNegativo As Integer = 0, Optional ByVal sOBS As String = "") As Long

Dim objItensPRJCR As New ClassItensPRJCR

On Error GoTo Erro_Preenche_Itens_Projeto_Aux

    objItensPRJCR.sItem = sItem
    objItensPRJCR.sDescricao = sDescricao
    objItensPRJCR.dValorOriginal = dValor
    objItensPRJCR.iNegativo = iNegativo
    objItensPRJCR.iTipoOrigem = iTipo
    objItensPRJCR.sObservacao = sOBS

    colItensPRJ.Add objItensPRJCR

    Exit Function

Erro_Preenche_Itens_Projeto_Aux:

    Preenche_Itens_Projeto_Aux = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157006)

    End Select

    Exit Function
    
End Function

Public Sub CTBGerencial_Click()
    Call objContabil.Contabil_Gerencial_Click
End Sub

Public Sub CTBGerencial_GotFocus()
    Call objContabil.Contabil_Gerencial_GotFocus
End Sub

Public Sub CTBGerencial_KeyPress(KeyAscii As Integer)
    Call objContabil.Contabil_Gerencial_KeyPress(KeyAscii)
End Sub

Public Sub CTBGerencial_Validate(Cancel As Boolean)
    Call objContabil.Contabil_Gerencial_Validate(Cancel)
End Sub

Function Seleciona_Serie_Padrao() As Long

Dim lErro As Long
Dim sSeriePadrao As String
Dim objSerie As New ClassSerie

On Error GoTo Erro_Seleciona_Serie_Padrao
        
    'Lê série Padrão
    lErro = CF("Serie_Le_Padrao", sSeriePadrao)
    If lErro <> SUCESSO Then gError 500130
    
    'Se possui série padrão
    If Len(Trim(sSeriePadrao)) > 0 Then
    
        'Lê a série
        objSerie.sSerie = sSeriePadrao
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO And lErro <> 22202 Then gError 500131
        
        'Série não cadastrada
        If lErro = 22202 Then gError 500132
               
        'Coloca ela na combo
        Serie.Text = Desconverte_Serie_Eletronica(sSeriePadrao)
        If ISSerieEletronica(sSeriePadrao) Then
            EletronicaFed.Value = vbChecked
        Else
            EletronicaFed.Value = vbUnchecked
        End If
        Call Serie_Validate(bSGECancelDummy)
        
    End If
                    
    Seleciona_Serie_Padrao = SUCESSO
    
    Exit Function

Erro_Seleciona_Serie_Padrao:

    Seleciona_Serie_Padrao = gErr
    
    Select Case gErr
    
        Case 500130, 500131
        
        Case 500132
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157167)
    
    End Select
    
    Exit Function
    
End Function

Public Sub MsgAutomatica_Click()
     Call Calcula_Msg
End Sub

Private Function Calcula_Msg() As Long

Dim lErro As Long

On Error GoTo Erro_Calcula_Msg

    If MsgAutomatica.Value = vbChecked Then
    
        lErro = gobjTribTab.AtualizarMsg
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        MsgAutomatica.Value = vbChecked
    End If
    
    Calcula_Msg = SUCESSO
    
    Exit Function
    
Erro_Calcula_Msg:
    
    Calcula_Msg = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208275)
    
    End Select
    
    Exit Function
        
End Function

Public Sub EletronicaFed_Click()
    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = REGISTRO_ALTERADO
    Call gobjTribTab.Serie_Alterada(Serie.Text)
End Sub

Private Function Carrega_TipoFrete() As Long

Dim lErro As Long
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer
Dim colTipoFrete As Collection

On Error GoTo Erro_Carrega_TipoFrete
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207458

    For Each objTipoFrete In colTipoFrete

        'coloca na combo
        TipoFrete.AddItem objTipoFrete.sDescricao
        TipoFrete.ItemData(TipoFrete.NewIndex) = objTipoFrete.iCodigo

    Next

    TipoFrete.ListIndex = 0

    Carrega_TipoFrete = SUCESSO

    Exit Function

Erro_Carrega_TipoFrete:

    Carrega_TipoFrete = gErr

    Select Case gErr

        Case 207458

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207459)

    End Select

    Exit Function

End Function

Public Sub BotaoInfoAdic_Click()

Dim lErro As Long
Dim sData As String
Dim bCancel As Boolean

On Error GoTo Erro_BotaoInfoAdic_Click

    If gobjInfoAdic Is Nothing Then Set gobjInfoAdic = New ClassInfoAdic

    Call Chama_Tela_Modal("RetiradaEntrega", gobjInfoAdic, Me)

    Exit Sub

Erro_BotaoInfoAdic_Click:

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207733)

    End Select

    Exit Sub

End Sub

Public Sub ValorDescontoItens_Change()
    'Seta iComissoesAlterada
    'iComissoesAlterada = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercDescontoItens_Change()
    'Seta iComissoesAlterada
    'iComissoesAlterada = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorDescontoItens_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorDesconto As Double
Dim vbMsg As VbMsgBoxResult, iIndice As Integer
Dim dDescontoItens As Double, dFator As Double

On Error GoTo Erro_ValorDescontoItens_Validate

    dValorDesconto = 0

    'Verifica se o Valor está preenchido
    If Len(Trim(ValorDescontoItens.Text)) > 0 Then
    
        'Faz a Crítica do Valor digitado
        lErro = Valor_NaoNegativo_Critica(ValorDescontoItens.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dValorDesconto = CDbl(ValorDescontoItens.Text)
        
        'Coloca o Valor formatado na tela
        ValorDescontoItens.Text = Format(dValorDesconto, "Standard")
        
    End If
    
    'Se houve alguma alteração
    If Abs(dValorDescontoItensAnt - dValorDesconto) > DELTA_VALORMONETARIO Then
        
        'Se o desconto foi alterado nos itens pegunta se quer que o sistema recalcule
        If iDescontoAlterado = REGISTRO_ALTERADO Then
        
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_DESCONTO_ITENS_ALTERADO")
            If vbMsg = vbNo Then
                dValorDesconto = dValorDescontoItensAnt
                ValorDescontoItens.Text = Format(dValorDesconto, "Standard")
                gError ERRO_SEM_MENSAGEM
            End If
            iDescontoAlterado = 0
               
        End If
        
        lErro = ValorDescontoItens_Aplica
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If Len(Trim(ValorDescontoItens.Text)) > 0 Then
            dValorDesconto = CDbl(ValorDescontoItens.Text)
        Else
            dValorDesconto = 0
        End If
        
        dValorDescontoItensAnt = dValorDesconto
    
    End If
    
    If StrParaDbl(ValorProdutos2.Caption) > 0 Then
        dPercDescontoItensAnt = Arredonda_Moeda(dValorDesconto / StrParaDbl(ValorProdutos2.Caption), 4)
    Else
        dPercDescontoItensAnt = 0
    End If
    PercDescontoItens.Text = Format(dPercDescontoItensAnt * 100, "FIXED")
        
    Exit Sub

Erro_ValorDescontoItens_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157330)

    End Select

    Exit Sub

End Sub

Public Sub PercDescontoItens_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dPercDesconto As Double
Dim vbMsg As VbMsgBoxResult, iIndice As Integer
Dim dDescontoItens As Double, dFator As Double

On Error GoTo Erro_PercDescontoItens_Validate

    dPercDesconto = 0

    'Verifica se o Valor está preenchido
    If Len(Trim(PercDescontoItens.Text)) > 0 Then
    
        'Faz a Crítica do Valor digitado
        lErro = Porcentagem_Critica(PercDescontoItens.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        dPercDesconto = StrParaDbl(PercDescontoItens.Text) / 100

    End If
    
    'Se houve alguma alteração
    If Abs(dPercDescontoItensAnt - dPercDesconto) > DELTA_VALORMONETARIO2 Then
        
        'Se o desconto foi alterado nos itens pegunta se quer que o sistema recalcule
        If iDescontoAlterado = REGISTRO_ALTERADO Then
        
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_DESCONTO_ITENS_ALTERADO")
            If vbMsg = vbNo Then
                dPercDesconto = dPercDescontoItensAnt
                PercDescontoItens.Text = Format(dPercDesconto * 100, "FIXED")
                gError ERRO_SEM_MENSAGEM
            End If
            iDescontoAlterado = 0
                
        End If
        
        ValorDescontoItens.Text = Format(Arredonda_Moeda(dPercDesconto * StrParaDbl(ValorProdutos2.Caption)), "Standard")
        Call ValorDescontoItens_Validate(bSGECancelDummy)
    
    End If
        
    Exit Sub

Erro_PercDescontoItens_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157330)

    End Select

    Exit Sub

End Sub

Function ValorDescontoItens_Aplica() As Long

Dim lErro As Long, dTotal As Double, dValorItem As Double, dDescontoItem As Double
Dim dDesconto As Double, dFator As Double, dDescontoAplicado As Double, dDiferenca As Double
Dim dPercDesc As Double, iIndice As Integer, dValorTotal As Double
Dim iLinha As Integer

On Error GoTo Erro_ValorDescontoItens_Aplica

    If objGridItens.iLinhasExistentes > 0 Then
    
        dTotal = StrParaDbl(ValorProdutos2.Caption)
        dDesconto = StrParaDbl(ValorDescontoItens.Text)
        If dTotal > 0 Then
            dFator = dDesconto / dTotal
        Else
            dFator = 0
        End If
        
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col)) > 0 And gobjNFiscal.ColItensNF.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
                If iLinha = 0 Then iLinha = iIndice
                dValorItem = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
                dDescontoItem = Arredonda_Moeda(dValorItem * dFator)
                dDescontoAplicado = dDescontoAplicado + dDescontoItem
                GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(dDescontoItem, "STANDARD")
            End If
        Next
        If Abs(dDesconto - dDescontoAplicado) > DELTA_VALORMONETARIO And iLinha > 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col)) + dDescontoAplicado - dDesconto, "STANDARD")
        End If
        
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col)) > 0 And gobjNFiscal.ColItensNF.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
                dPercDesc = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col)) / StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
                GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dPercDesc, "Percent")
                Call Calcula_Valores(iIndice)
                
                lErro = gobjTribTab.Alteracao_Item_Grid(iIndice)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                dValorTotal = dValorTotal + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
            End If
        Next
    
    End If
        
    'Coloca valor total dos produtos na tela
    SubTotal.Caption = Format(dValorTotal, "Standard")

    'Calcula o valor total da nota
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    ValorDescontoItens_Aplica = SUCESSO

    Exit Function

Erro_ValorDescontoItens_Aplica:

    ValorDescontoItens_Aplica = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208382)

    End Select

    Exit Function

End Function

Function ValorDescontoItens_Calcula() As Long

Dim lErro As Long, iIndice As Integer
Dim dDesconto As Double, dPercDesc As Double

On Error GoTo Erro_ValorDescontoItens_Calcula

    dDesconto = 0
    dPercDesc = 0
    If Not (objGridItens Is Nothing) Then
        If objGridItens.iLinhasExistentes > 0 Then
            For iIndice = 1 To objGridItens.iLinhasExistentes
                If gobjNFiscal.ColItensNF.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
                    dDesconto = dDesconto + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
                End If
            Next
            If StrParaDbl(ValorProdutos2.Caption) > 0 Then dPercDesc = Arredonda_Moeda(dDesconto / StrParaDbl(ValorProdutos2.Caption), 4)
            ValorDescontoItens.Text = Format(dDesconto, "Standard")
            PercDescontoItens.Text = Format(dPercDesc * 100, "FIXED")
        Else
            ValorDescontoItens.Text = Format(0, "Standard")
            PercDescontoItens.Text = Format(0, "FIXED")
        End If
        
        dValorDescontoItensAnt = dDesconto
        dPercDescontoItensAnt = dPercDesc
    End If

    ValorDescontoItens_Calcula = SUCESSO

    Exit Function

Erro_ValorDescontoItens_Calcula:

    ValorDescontoItens_Calcula = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208382)

    End Select

    Exit Function

End Function

Public Sub BotaoInfoAdicItem_Click()

Dim lErro  As Long
Dim objItemNF As ClassItemNF
Dim objTela As Object

On Error GoTo Erro_BotaoInfoAdicItem_Click

    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
    
        Set objItemNF = gobjNFiscal.colItens(GridItens.Row)
        
        Set objTela = Me
        
        objItemNF.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMNF
        objItemNF.objInfoAdicDocItem.iItem = GridItens.Row
        objItemNF.objInfoAdicDocItem.sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        objItemNF.objInfoAdicDocItem.sDescProd = GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col)
        
        objItemNF.objInfoAdicDocItem.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
        objItemNF.objInfoAdicDocItem.dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
        objItemNF.objInfoAdicDocItem.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))
        objItemNF.objInfoAdicDocItem.dValorTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))
        
        Call Chama_Tela_Modal("InfoAdicDocItem", objItemNF.objInfoAdicDocItem, objTela, TIPO_ENTRADA)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Call SubTotal_Calcula
        
        Call ValorTotal_Calcula

    End If
    
    Exit Sub

Erro_BotaoInfoAdicItem_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211297)
            
    End Select
    
    Exit Sub

End Sub
Public Sub BotaoTrazerNFe_Click()

Dim lErro  As Long
Dim objTela As Object, objNF As New ClassNFiscal
Dim colPedidoSel As New Collection

On Error GoTo Erro_BotaoTrazerNFe_Click

    If Len(Trim(ChvNFe.ClipText)) = 0 Then gError 211310

    Set objTela = Me
    Set gobjNFiscal = objNF
    
    lErro = CF("NFe_Chv_Monta_NF", objTela, objNF, ChvNFe.ClipText, colPedidoSel)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If objNF.lNumNotaFiscal > 0 Then
        
        lErro = Traz_NFiscal_Tela1(objNF, colPedidoSel)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        Call ValorTotal_Calcula
    
    End If
    
    Exit Sub

Erro_BotaoTrazerNFe_Click:

    Select Case gErr
    
        Case 211310
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CHAVE_NAO_PREENCHIDA", gErr)
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211245)
            
    End Select
    
    Exit Sub

End Sub

Public Sub ChvNFeLabel_Click()

Dim lErro As Long, objNF As New ClassNFiscal
Dim colSelecao As New Collection
Dim iFilialEmpresa As Integer, objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_ChvNFeLabel_Click

    objNF.sChvNFe = ChvNFe.ClipText
    
    iFilialEmpresa = giFilialEmpresa
    If iFilialEmpresa > 50 Then iFilialEmpresa = iFilialEmpresa - 50
    
    objFilialEmpresa.iCodFilial = iFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    colSelecao.Add objFilialEmpresa.sCgc
    
    Call Chama_Tela("NFeXmlLista", colSelecao, objNF, objEventoChvNFe, "CGCEmitente <> ?")

    Exit Sub
    
Erro_ChvNFeLabel_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201452)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoChvNFe_evSelecao(obj1 As Object)

Dim bCancel As Boolean
Dim objNFiscal As New ClassNFiscal
Dim lErro As Long

On Error GoTo Erro_objEventoChvNFe_evSelecao

    Set objNFiscal = obj1

    ChvNFe.PromptInclude = False
    ChvNFe.Text = objNFiscal.sChvNFe
    ChvNFe.PromptInclude = True
    
    Call BotaoTrazerNFe_Click

    Exit Sub
    
Erro_objEventoChvNFe_evSelecao:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188513)

    End Select
    
    Exit Sub
    
End Sub

Public Sub ChvNFe_Validate(Cancel As Boolean)

Dim lErro As Long, objNF As New ClassNFiscal
Dim bExisteNF As Boolean, bExisteXml As Boolean, bTrazer As Boolean, lNumIntNF As Long

On Error GoTo Erro_ChvNFe_Validate

    If gsChvAnt <> ChvNFe.ClipText Then

        lErro = CF("NFe_Chv_Valida", ChvNFe.ClipText, bExisteNF, bExisteXml, bTrazer, lNumIntNF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If bTrazer Then
            If bExisteNF Then
                objNF.lNumIntDoc = lNumIntNF
                lErro = Traz_NFiscal_Tela(objNF)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            ElseIf bExisteXml Then
                Call BotaoTrazerNFe_Click
            End If
        End If
        
        gsChvAnt = ChvNFe.ClipText
        
    End If
       
    Exit Sub

Erro_ChvNFe_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157330)

    End Select

    Exit Sub

End Sub

Public Sub ChvNFe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ChvNFe_GotFocus()
    Call MaskEdBox_TrataGotFocus(ChvNFe, iAlterado)
End Sub

Public Sub Recebimento_Click()
'######FALTA FAZER
End Sub

Private Function Traz_NFiscal_Tela1(objNFiscal As ClassNFiscal, colPedSel As Collection) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim bCancel As Boolean
Dim objItemNF As ClassItemNF
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim colItemPedCompraInfo As New Collection
Dim iLinha As Integer
Dim iIndice As Integer
Dim dPercDesc As Double
Dim objPedidoCompra As ClassPedidoCompras
Dim iKeyCode As Integer, iShift As Integer, bAchou As Boolean
Dim sProdutoFormatado As String, iProdutoPreenchido As Integer
Dim colItensOrd As New ColItensNF, colItensOrdInv As New ColItensNF

On Error GoTo Erro_Traz_NFiscal_Tela1
   
    'Limpa a tela
    Call Limpa_Tela_NFiscalCom1
        
    If objNFiscal.lFornecedor <> 0 Then
        'preenche o Fornecedor
        Fornecedor.Text = objNFiscal.lFornecedor
        Call Fornecedor_Validate(bCancel)
        
        'Preenche a Filial do Fornecedor
        Filial.Text = objNFiscal.iFilialForn
        Call Filial_Validate(bCancel)
    End If

    TipoNFiscal.Text = objNFiscal.iTipoNFiscal
    Call TipoNFiscal_Validate(bCancel)
    
    FilialCompra.Text = objNFiscal.iFilialPedido
    Call FilialCompra_Validate(bCancel)
        
    'Marca os Pedidos de Compras associados ao Recebimento
    For Each objPedidoCompra In colPedSel
        For iIndice = 0 To PedidosCompra.ListCount - 1
            If objPedidoCompra.lCodigo = CLng(PedidosCompra.List(iIndice)) Then
                PedidosCompra.Selected(iIndice) = True
            End If
        Next
    Next
    
    For iIndice = objGridItens.iLinhasExistentes To 1 Step -1
        iKeyCode = 46
        iShift = 0
        'Verifica se o Produto está preenchido
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        bAchou = False
        For Each objItemNF In objNFiscal.colItens
            If sProdutoFormatado = objItemNF.sProduto Then
                bAchou = True
                Exit For
            End If
        Next
        If bAchou Then
        
            colItensOrd.Add1 objItemNF
            
            'Dá uma ajeitada na quantidade e valor
            GridItens.Row = iIndice
            GridItens.Col = iGrid_Quantidade_Col
            Call GridItens_Click
        
            Quantidade.Text = objItemNF.dQuantidade
        
            GridItens.Col = iGrid_ValorUnitario_Col
            Call GridItens_Click
        
            ValorUnitario.Text = objItemNF.dPrecoUnitario
        
            GridItens.Col = iGrid_Desconto_Col
            Call GridItens_Click
        
            Desconto.Text = objItemNF.dValorDesconto
        
            GridItens.Col = iGrid_PrecoTotalB_Col
            Call GridItens_Click
        
        Else
            GridItens.Row = iIndice
            GridItens.Col = iGrid_Produto_Col
            Call GridItens_Click
            'Caso não exista no xml excluí
            Call GridItens_KeyDown(iKeyCode, iShift, False)
        End If
    Next
    
    For iIndice = colItensOrd.Count To 1 Step -1
        Set objItemNF = colItensOrd.Item(iIndice)
        colItensOrdInv.Add1 objItemNF
    Next
    
    'Coloca na ordem que está no grid
    Set objNFiscal.ColItensNF = colItensOrdInv
    
    gbCarregandoTela = True
    
    ChvNFe.PromptInclude = False
    ChvNFe.Text = objNFiscal.sChvNFe
    ChvNFe.PromptInclude = True
    gsChvAnt = objNFiscal.sChvNFe
    
    If objNFiscal.lNumNotaFiscal <> 0 Then NFiscal.Text = objNFiscal.lNumNotaFiscal

    If objNFiscal.sNaturezaOp <> "" Then NaturezaOp.Text = objNFiscal.sNaturezaOp
    Serie.Text = Desconverte_Serie_Eletronica(objNFiscal.sSerie)
    If ISSerieEletronica(objNFiscal.sSerie) Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    Call Serie_Validate(bSGECancelDummy)

    'Se a data não for nula coloca na Tela
    DataEmissao.PromptInclude = False
    If objNFiscal.dtDataEmissao <> DATA_NULA Then DataEmissao.Text = Format(objNFiscal.dtDataEmissao, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    DataEntrada.PromptInclude = False
    If objNFiscal.dtDataEntrada <> DATA_NULA Then DataEntrada.Text = Format(objNFiscal.dtDataEntrada, "dd/mm/yy")
    DataEntrada.PromptInclude = True

    For iIndice = 0 To TipoFrete.ListCount - 1
        If TipoFrete.ItemData(iIndice) = objNFiscal.iFreteRespons Then
            TipoFrete.ListIndex = iIndice
            Exit For
        End If
    Next

    VolumeNumero = objNFiscal.sVolumeNumero
    If objNFiscal.dValorOutrasDespesas > 0 Then ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Standard")
    If objNFiscal.dValorFrete > 0 Then ValorFrete.Text = Format(objNFiscal.dValorFrete, "Standard")
    If objNFiscal.dValorSeguro > 0 Then ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Standard")
    If objNFiscal.dValorTotal > 0 Then Total.Text = Format(objNFiscal.dValorTotal, "Standard")
    If objNFiscal.dValorProdutos > 0 Then SubTotal.Caption = Format(objNFiscal.dValorProdutos, "Standard")

    If objNFiscal.iCodTransportadora <> 0 Then
        Transportadora.Text = objNFiscal.iCodTransportadora
        Call Transportadora_Validate(bCancel)
    End If

    Placa.Text = objNFiscal.sPlaca
    PlacaUF.Text = objNFiscal.sPlacaUF
    
    'Traz a espécie dos volumes do pedido
    If objNFiscal.lVolumeEspecie > 0 Then
        VolumeEspecie.Text = objNFiscal.lVolumeEspecie
        Call VolumeEspecie_Validate(bSGECancelDummy)
    Else
        VolumeEspecie.Text = ""
    End If
    
    'Traz a marca dos volumes do pedido
    If objNFiscal.lVolumeMarca > 0 Then
        VolumeMarca.Text = objNFiscal.lVolumeMarca
        Call VolumeMarca_Validate(bSGECancelDummy)
    Else
        VolumeMarca.Text = ""
    End If
    
    If objNFiscal.lVolumeQuant > 0 Then VolumeQuant.Text = objNFiscal.lVolumeQuant

    Mensagem.Text = objNFiscal.sMensagemNota
    MensagemCorpo.Text = objNFiscal.sMensagemCorpoNota
    Observacao.Text = objNFiscal.sObservacao
    
    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = CStr(objNFiscal.dPesoLiq)
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = CStr(objNFiscal.dPesoBruto)
    
'    'Preenche a Condicao de Pagto
'    If objNFiscal.iCondicaoPagto > 0 Then
'        CondicaoPagamento.Text = objNFiscal.iCondicaoPagto
'        Call CondicaoPagamento_Validate(bCancel)
'    End If
    
    If objNFiscal.dValorItens = 0 Then
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If objNFiscal.ColItensNF.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
                objNFiscal.dValorItens = objNFiscal.dValorItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
                objNFiscal.dValorDescontoItens = objNFiscal.dValorDescontoItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
            End If
        Next
    End If
    
    ValorProdutos2.Caption = Format(objNFiscal.dValorItens, "Standard")
    ValorDescontoItens.Text = Format(objNFiscal.dValorDescontoItens, "Standard")
    If objNFiscal.dValorItens > 0 Then
        dPercDesc = objNFiscal.dValorDescontoItens / objNFiscal.dValorItens
    Else
        dPercDesc = 0
    End If
    PercDescontoItens.Text = Format(dPercDesc * 100, "Fixed")
    
    dValorDescontoItensAnt = objNFiscal.dValorDescontoItens
    dPercDescontoItensAnt = dPercDesc

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 188519
'*** fim tributacao

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        GridItens.Row = iIndice
        
        lErro = gobjDistribuicao.Distribuicao_Processa()
        'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Next

'    lErro = Preenche_GridParcelas(objNFiscal)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    gbCarregandoTela = False
    
    Traz_NFiscal_Tela1 = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela1:

    gbCarregandoTela = False
    
    Traz_NFiscal_Tela1 = gErr

    Select Case gErr

        Case 188515 To 188519, ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188520)

    End Select

    Exit Function

End Function


