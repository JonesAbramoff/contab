VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalFaturaPedido"
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
Public giValorDescontoManual As Integer

Dim gobjEmbalagem As Object

Public gdQuantAnterior As Double

Dim m_objUserControl As Object
   
'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjNFiscal As ClassNFiscal

Dim iFornEntTercAlterado As Integer
Dim giTipoNFiscalPadrao As Integer

'inicia objeto associado a GridComissoes
Public objTabComissoes As New ClassTabComissoes

'Flag que indica se a tela está sendo preenchida.
Public gbCarregandoTela As Boolean
Public giPosCargaOk As Integer
Dim giComplemento As Integer

'Indica que está limpando a tela (é usada para evitar tentar calcular comissões ao limpar a tela)
Public gbLimpandoTela As Boolean

'Associados a contabilidade
Public objContabil As New ClassContabil
Public WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Public WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

Public gobjTelaProjetoInfo As ClassTelaPRJInfo

'declara o objeto da execucao das regras de comissoes e
'o objeto do calculo de comissoes
Public objComissoesRegrasCalc As Object
Public objMnemonicoComissCalc As ClassMnemonicoComissCalc
Public objMnemonicoComissCalcAux As ClassMnemonicoComissCalcAux

'Indica que foi alterado pelo menos um campo na tela que seja utilizado para calcular as comissões
Public iComissoesAlterada As Integer

'Flags e variáveis globais da tela
Public iAlterado As Integer
Public iFrameAtual As Integer
Public gobjTribTab As ClassTribTab
Public glCodPedido As Long
Public giFilialPedido As Integer
Public giFaturaIntegral As Integer
Public iPercAcresFinAlterado As Integer
Public iValorFreteAlterado As Integer
Public iValorSeguroAlterado As Integer
Public iValorDespesasAlterado As Integer
Public iValorDescontoAlterado As Integer
Public gdDesconto As Double
Public gcolItemPedido As Collection
Public gcolItemQuant As Collection
Public gcolItemUM As Collection
Public giFilialEntregaAlterada As Integer
Public giFilialEntregaManual As Integer
Dim giDataEmissaoAlterada As Integer
Public iSerieAlterada As Integer
Public giDataReferenciaAlterada As Integer

'Grids
Public objGridItens As AdmGrid
Public objGridComissoes As AdmGrid
Public objGridAlocacoes As AdmGrid
Public objGridParcelas As AdmGrid
Public objGrid1 As AdmGrid

'Colunas do grid de itens
Public iGrid_Produto_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
Public iGrid_ValorUnitario_Col As Integer
Public iGrid_ValorUnitarioMoeda_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PrecoTotal_Col As Integer
Public iGrid_PrecoTotalB_Col As Integer
Public iGrid_DataEntrega_Col As Integer
Public iGrid_TabPrecoItem_Col As Integer
Public iGrid_ComissaoItem_Col As Integer

'Colunas do grid de Alocações
Public iGrid_Item_Col As Integer
Public iGrid_ProdutoAloc_Col As Integer
Public iGrid_AlmoxAloc_Col As Integer
Public iGrid_UMAloc_Col As Integer
Public iGrid_QuantAloc_Col As Integer
Public iGrid_QuantVend_Col As Integer

'Colunas do grid de parcelas
Public iGrid_Vencimento_Col As Integer
Public iGrid_ValorParcela_Col As Integer
Public iGrid_TipoPagto_Col As Integer
Public iGrid_Desc1Codigo_Col As Integer
Public iGrid_Desc1Ate_Col As Integer
Public iGrid_Desc1Valor_Col As Integer
Public iGrid_Desc1Percentual_Col As Integer
Public iGrid_Desc2Codigo_Col As Integer
Public iGrid_Desc2Ate_Col As Integer
Public iGrid_Desc2Valor_Col As Integer
Public iGrid_Desc2Percentual_Col As Integer
Public iGrid_Desc3Codigo_Col As Integer
Public iGrid_Desc3Ate_Col As Integer
Public iGrid_Desc3Valor_Col As Integer
Public iGrid_Desc3Percentual_Col As Integer

'mnemonicos
Private Const TIPO_NOTA_FISCAL As String = "Tipo_Nota_Fiscal"
Private Const CODIGO1 As String = "Codigo"
Private Const NATUREZA_OP As String = "Natureza_OP"
Private Const CLIENTE1 As String = "Cliente"
Private Const CLIENTE_CODIGO As String = "Cliente_Codigo"
Private Const FILIAL1 As String = "Filial"
Private Const FILIAL_CONTA As String = "FilialCli_Conta_Ctb"
Private Const Serie1 As String = "Serie"
Private Const NFISCAL1 As String = "Nota_Fiscal"
Private Const DATA_EMISSAO As String = "Data_Emissao"
Private Const DATA_SAIDA As String = "Data_Saida"
Private Const PRODUTO1 As String = "Produto_Codigo"
Private Const UNIDADE_MED As String = "Unidade_Medida"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const PRECO_UNITARIO As String = "Preco_Unitario"
Private Const PRECO_TOTAL As String = "Preco_Total"
Private Const DESCONTO1 As String = "Desconto_Produto"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"
'#######Valores####################
Private Const ICMS As String = "ICMS_Valor"

Private Const ICMS_DIFAL_REMETENTE As String = "ICMS_Difal_Remetente"
Private Const ICMS_DIFAL_DESTINO As String = "ICMS_Difal_Destino"
Private Const ICMS_FCP As String = "ICMS_FCP_Valor"

Private Const ICMSSUBST As String = "ICMSSubst_Valor"
Private Const VALOR_PRODUTOS As String = "Valor_Produtos"
Private Const VALOR_FRETE As String = "Valor_Frete"
Private Const VALOR_SEGURO As String = "Valor_Seguro"
Private Const VALOR_DESPESAS As String = "Valor_Despesas"
Private Const IPI As String = "Valor_IPI"
Private Const VALOR_DESCONTO As String = "Valor_Desconto_Glb"
Private Const VALOR_TOTAL As String = "Valor_Total"
'###########Almoxarifado############
Private Const PRODUTO_ALMOX As String = "Produto_Almox"
Private Const ALMOX1 As String = "Almoxarifado"
Private Const QUANT_ALOCADA As String = "Quant_Alocada"
Private Const UNIDADE_MED_EST As String = "Unidade_Med_Est"
'###########Tributação##############
Private Const ISS_VALOR As String = "ISS_Valor"
Private Const ISS_INCLUSO As String = "ISS_Incluso"
Private Const VALOR_IRRF As String = "Valor_IRRF"
Private Const VALOR_INSS As String = "Valor_INSS"
Private Const INSS_RETIDO As String = "INSS_Retido"
Private Const CTACONTABILEST1 As String = "ContaContabilEst"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque"
Private Const PIS_DEBITO_TOTAL As String = "PIS_Debito_Total"
Private Const COFINS_DEBITO_TOTAL As String = "COFINS_Debito_Total"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const ISS_RETIDO As String = "ISS_Retido"
Private Const COFINS_RETIDO As String = "COFINS_Retido"
Private Const CSLL_RETIDO As String = "CSLL_Retido"
Private Const CONTACONTABILAPLIC As String = "ContaContabilAplic"
Private Const CONTACONTABILCATPROD As String = "ContaContabilCatProd"
Private Const CONTACONTABILPROD As String = "ContaContabilProd"
Private Const PIS_VALOR_TOTAL As String = "PIS_Valor_Total"
Private Const COFINS_VALOR_TOTAL As String = "COFINS_Valor_Total"
'fim da contabilidade

'Eventos das telas de browse
Private WithEvents objEventoPedido As AdmEvento
Attribute objEventoPedido.VB_VarHelpID = -1
Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Public WithEvents objEventoNFiscal As AdmEvento
Attribute objEventoNFiscal.VB_VarHelpID = -1
Private WithEvents objEventoNaturezaOp As AdmEvento
Attribute objEventoNaturezaOp.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoEstoqueProd As AdmEvento
Attribute objEventoEstoqueProd.VB_VarHelpID = -1
Public WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoLocalizacao As AdmEvento
Attribute objEventoLocalizacao.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoTranspRedespacho As AdmEvento
Attribute objEventoTranspRedespacho.VB_VarHelpID = -1
Private WithEvents objEventoMensagem As AdmEvento
Attribute objEventoMensagem.VB_VarHelpID = -1
Private WithEvents objEventoFornEntTerc As AdmEvento
Attribute objEventoFornEntTerc.VB_VarHelpID = -1

Public gobjRastreamento As ClassRastreamento

Private Const BOTAO_PEDIDO = 0
Private Const BOTAO_VERPEDIDO = 1
Private Const BOTAO_IMPRIMIR = 2
Private Const BOTAO_CONSULTATITREC = 3
Private Const BOTAO_GRAVAR = 4
Private Const BOTAO_EXCLUIR = 5
Private Const BOTAO_LIMPAR = 6
Private Const BOTAO_FECHAR = 7
Private Const BOTAO_LIMPARNF = 8
Private Const BOTAO_PROJETOS = 9
Private Const BOTAO_GRADE = 10
Private Const BOTAO_KITVENDA = 11
Private Const BOTAO_PRODUTOS = 12
Private Const BOTAO_ESTOQUEPROD = 13
Private Const BOTAO_LOCALIZACAO = 14
Private Const BOTAO_SERIE = 15
Private Const BOTAO_LOTES = 16
Private Const BOTAO_REFAZ_ALOC = 17

'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_Itens = 2
Private Const TAB_Complemento = 3
Private Const TAB_Cobranca = 4
Private Const TAB_Comissoes = 5
Private Const TAB_Almoxarifado = 6
Private Const TAB_Tributacao = 7
Private Const TAB_Contabilizacao = 8
Private Const TAB_RASTRO = 9
Private Const TAB_Embalagens = 10

'--- inicio dos properties get dos controles da tela
Public Property Get Botao() As Object
     Set Botao = objUserControl.Controls("Botao")
End Property

Public Property Get TabPrecoItemNF() As Object
     Set TabPrecoItemNF = objUserControl.Controls("TabPrecoItemNF")
End Property

Public Property Get ComissaoItemNF() As Object
     Set ComissaoItemNF = objUserControl.Controls("ComissaoItemNF")
End Property

Public Property Get FrameFornEntTerc() As Object
     Set FrameFornEntTerc = objUserControl.Controls("FrameFornEntTerc")
End Property

Public Property Get FilialFornEntTerc() As Object
     Set FilialFornEntTerc = objUserControl.Controls("FilialFornEntTerc")
End Property

Public Property Get FornEntTerc() As Object
     Set FornEntTerc = objUserControl.Controls("FornEntTerc")
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

Public Property Get ValorDescontoTit() As Object
     Set ValorDescontoTit = objUserControl.Controls("ValorDescontoTit")
End Property

Public Property Get ValorOriginalTit() As Object
     Set ValorOriginalTit = objUserControl.Controls("ValorOriginalTit")
End Property

Public Property Get ValorTit() As Object
     Set ValorTit = objUserControl.Controls("ValorTit")
End Property

Public Property Get MsgAutomatica() As Object
     Set MsgAutomatica = objUserControl.Controls("MsgAutomatica")
End Property

Public Property Get NaoImpCobranca() As Object
     Set NaoImpCobranca = objUserControl.Controls("NaoImpCobranca")
End Property

'Rastreamento
Public Property Get EscaninhoRastro() As Object
     Set EscaninhoRastro = objUserControl.Controls("EscaninhoRastro")
End Property

Public Property Get TipoNFiscal() As Object
     Set TipoNFiscal = objUserControl.Controls("TipoNFiscal")
End Property

Public Property Get Cubagem() As Object
     Set Cubagem = objUserControl.Controls("Cubagem")
End Property

Public Property Get RedespachoCli() As Object
     Set RedespachoCli = objUserControl.Controls("RedespachoCli")
End Property

Public Property Get TranspRedespacho() As Object
     Set TranspRedespacho = objUserControl.Controls("TranspRedespacho")
End Property

Public Property Get DiretoIndireto() As Object
     Set DiretoIndireto = objUserControl.Controls("DiretoIndireto")
End Property

Public Property Get TotalValorBase() As Object
     Set TotalValorBase = objUserControl.Controls("TotalValorBase")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get TabelaPreco() As Object
     Set TabelaPreco = objUserControl.Controls("TabelaPreco")
End Property

Public Property Get CondicaoPagamento() As Object
     Set CondicaoPagamento = objUserControl.Controls("CondicaoPagamento")
End Property

Public Property Get PercAcrescFin() As Object
     Set PercAcrescFin = objUserControl.Controls("PercAcrescFin")
End Property

Public Property Get CondPagtoLabel() As Object
     Set CondPagtoLabel = objUserControl.Controls("CondPagtoLabel")
End Property

Public Property Get BotaoLimparNF() As Object
     Set BotaoLimparNF = objUserControl.Controls("BotaoLimparNF")
End Property

Public Property Get Serie() As Object
     Set Serie = objUserControl.Controls("Serie")
End Property

Public Property Get NFiscal() As Object
     Set NFiscal = objUserControl.Controls("NFiscal")
End Property

Public Property Get NFiscalLabel() As Object
     Set NFiscalLabel = objUserControl.Controls("NFiscalLabel")
End Property

Public Property Get SerieLabel() As Object
     Set SerieLabel = objUserControl.Controls("SerieLabel")
End Property

Public Property Get Filial() As Object
'    If giTipoVersao = VERSAO_FULL Then
        Set Filial = objUserControl.Controls("Filial")
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'        Set Filial = objUserControl.Controls("FilialL")
'    End If
End Property

Public Property Get Cliente() As Object
'    If giTipoVersao = VERSAO_FULL Then
        Set Cliente = objUserControl.Controls("Cliente")
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'        Set Cliente = objUserControl.Controls("ClienteL")
'    End If
End Property

Public Property Get ClienteLabel() As Object
     Set ClienteLabel = objUserControl.Controls("ClienteLabel")
End Property

Public Property Get UpDownEmissao() As Object
     Set UpDownEmissao = objUserControl.Controls("UpDownEmissao")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get UpDownSaida() As Object
     Set UpDownSaida = objUserControl.Controls("UpDownSaida")
End Property

Public Property Get DataSaida() As Object
     Set DataSaida = objUserControl.Controls("DataSaida")
End Property

'horasaida
Public Property Get HoraSaida() As Object
     Set HoraSaida = objUserControl.Controls("HoraSaida")
End Property

Public Property Get CTBSeqContraPartida() As Object
     Set CTBSeqContraPartida = objUserControl.Controls("CTBSeqContraPartida")
End Property

Public Property Get CTBAglutina() As Object
     Set CTBAglutina = objUserControl.Controls("CTBAglutina")
End Property

Public Property Get CTBHistorico() As Object
     Set CTBHistorico = objUserControl.Controls("CTBHistorico")
End Property

Public Property Get CTBListHistoricos() As Object
     Set CTBListHistoricos = objUserControl.Controls("CTBListHistoricos")
End Property

Public Property Get CTBBotaoModeloPadrao() As Object
     Set CTBBotaoModeloPadrao = objUserControl.Controls("CTBBotaoModeloPadrao")
End Property

Public Property Get CTBBotaoLimparGrid() As Object
     Set CTBBotaoLimparGrid = objUserControl.Controls("CTBBotaoLimparGrid")
End Property

Public Property Get CTBModelo() As Object
     Set CTBModelo = objUserControl.Controls("CTBModelo")
End Property

Public Property Get CTBFrame7() As Object
     Set CTBFrame7 = objUserControl.Controls("CTBFrame7")
End Property

Public Property Get CTBCclLabel() As Object
     Set CTBCclLabel = objUserControl.Controls("CTBCclLabel")
End Property

Public Property Get CTBLabel7() As Object
     Set CTBLabel7 = objUserControl.Controls("CTBLabel7")
End Property

Public Property Get CTBContaDescricao() As Object
     Set CTBContaDescricao = objUserControl.Controls("CTBContaDescricao")
End Property

Public Property Get CTBCclDescricao() As Object
     Set CTBCclDescricao = objUserControl.Controls("CTBCclDescricao")
End Property

Public Property Get CTBBotaoImprimir() As Object
     Set CTBBotaoImprimir = objUserControl.Controls("CTBBotaoImprimir")
End Property

Public Property Get CTBLancAutomatico() As Object
     Set CTBLancAutomatico = objUserControl.Controls("CTBLancAutomatico")
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

Public Property Get CTBTvwCcls() As Object
     Set CTBTvwCcls = objUserControl.Controls("CTBTvwCcls")
End Property

Public Property Get CTBTvwContas() As Object
     Set CTBTvwContas = objUserControl.Controls("CTBTvwContas")
End Property

Public Property Get CTBGridContabil() As Object
     Set CTBGridContabil = objUserControl.Controls("CTBGridContabil")
End Property

Public Property Get CTBLabel21() As Object
     Set CTBLabel21 = objUserControl.Controls("CTBLabel21")
End Property

Public Property Get CTBOrigem() As Object
     Set CTBOrigem = objUserControl.Controls("CTBOrigem")
End Property

Public Property Get CTBLabel14() As Object
     Set CTBLabel14 = objUserControl.Controls("CTBLabel14")
End Property

Public Property Get CTBPeriodo() As Object
     Set CTBPeriodo = objUserControl.Controls("CTBPeriodo")
End Property

Public Property Get CTBExercicio() As Object
     Set CTBExercicio = objUserControl.Controls("CTBExercicio")
End Property

Public Property Get CTBLabel13() As Object
     Set CTBLabel13 = objUserControl.Controls("CTBLabel13")
End Property

Public Property Get CTBLabel5() As Object
     Set CTBLabel5 = objUserControl.Controls("CTBLabel5")
End Property

Public Property Get CTBLabelHistoricos() As Object
     Set CTBLabelHistoricos = objUserControl.Controls("CTBLabelHistoricos")
End Property

Public Property Get CTBLabelContas() As Object
     Set CTBLabelContas = objUserControl.Controls("CTBLabelContas")
End Property

Public Property Get CTBLabelCcl() As Object
     Set CTBLabelCcl = objUserControl.Controls("CTBLabelCcl")
End Property

Public Property Get CTBLabel1() As Object
     Set CTBLabel1 = objUserControl.Controls("CTBLabel1")
End Property

Public Property Get CTBLabelTotais() As Object
     Set CTBLabelTotais = objUserControl.Controls("CTBLabelTotais")
End Property

Public Property Get CTBTotalDebito() As Object
     Set CTBTotalDebito = objUserControl.Controls("CTBTotalDebito")
End Property

Public Property Get CTBTotalCredito() As Object
     Set CTBTotalCredito = objUserControl.Controls("CTBTotalCredito")
End Property

Public Property Get CTBLabel8() As Object
     Set CTBLabel8 = objUserControl.Controls("CTBLabel8")
End Property

Public Property Get CTBLabelDoc() As Object
     Set CTBLabelDoc = objUserControl.Controls("CTBLabelDoc")
End Property

Public Property Get CTBLabelLote() As Object
     Set CTBLabelLote = objUserControl.Controls("CTBLabelLote")
End Property

Public Property Get BotaoLocalizacao() As Object
     Set BotaoLocalizacao = objUserControl.Controls("BotaoLocalizacao")
End Property

Public Property Get ProdutoAlmox() As Object
     Set ProdutoAlmox = objUserControl.Controls("ProdutoAlmox")
End Property

Public Property Get Almox() As Object
     Set Almox = objUserControl.Controls("Almox")
End Property

Public Property Get QuantAlocada() As Object
     Set QuantAlocada = objUserControl.Controls("QuantAlocada")
End Property

Public Property Get ItemNFiscal() As Object
     Set ItemNFiscal = objUserControl.Controls("ItemNFiscal")
End Property

Public Property Get UnidadeMedEst() As Object
     Set UnidadeMedEst = objUserControl.Controls("UnidadeMedEst")
End Property

Public Property Get QuantVendida() As Object
     Set QuantVendida = objUserControl.Controls("QuantVendida")
End Property

Public Property Get GridAlocacao() As Object
     Set GridAlocacao = objUserControl.Controls("GridAlocacao")
End Property

Public Property Get ComissaoAutomatica() As Object
     Set ComissaoAutomatica = objUserControl.Controls("ComissaoAutomatica")
End Property

Public Property Get BotaoVendedores() As Object
     Set BotaoVendedores = objUserControl.Controls("BotaoVendedores")
End Property

Public Property Get ValorBaixa() As Object
     Set ValorBaixa = objUserControl.Controls("ValorBaixa")
End Property

Public Property Get ValorComissao() As Object
     Set ValorComissao = objUserControl.Controls("ValorComissao")
End Property

Public Property Get ValorBase() As Object
     Set ValorBase = objUserControl.Controls("ValorBase")
End Property

Public Property Get PercentualComissao() As Object
     Set PercentualComissao = objUserControl.Controls("PercentualComissao")
End Property

Public Property Get Vendedor() As Object
     Set Vendedor = objUserControl.Controls("Vendedor")
End Property

Public Property Get ValorEmissao() As Object
     Set ValorEmissao = objUserControl.Controls("ValorEmissao")
End Property

Public Property Get PercentualEmissao() As Object
     Set PercentualEmissao = objUserControl.Controls("PercentualEmissao")
End Property

Public Property Get PercentualBaixa() As Object
     Set PercentualBaixa = objUserControl.Controls("PercentualBaixa")
End Property

Public Property Get GridComissoes() As Object
     Set GridComissoes = objUserControl.Controls("GridComissoes")
End Property

Public Property Get TotalPercentualComissao() As Object
     Set TotalPercentualComissao = objUserControl.Controls("TotalPercentualComissao")
End Property

Public Property Get TotalValorComissao() As Object
     Set TotalValorComissao = objUserControl.Controls("TotalValorComissao")
End Property

Public Property Get LabelTotaisComissoes() As Object
     Set LabelTotaisComissoes = objUserControl.Controls("LabelTotaisComissoes")
End Property

Public Property Get CobrancaAutomatica() As Object
     Set CobrancaAutomatica = objUserControl.Controls("CobrancaAutomatica")
End Property

Public Property Get BotaoDataReferenciaDown() As Object
     Set BotaoDataReferenciaDown = objUserControl.Controls("BotaoDataReferenciaDown")
End Property

Public Property Get BotaoDataReferenciaUp() As Object
     Set BotaoDataReferenciaUp = objUserControl.Controls("BotaoDataReferenciaUp")
End Property

Public Property Get Desconto1Codigo() As Object
     Set Desconto1Codigo = objUserControl.Controls("Desconto1Codigo")
End Property

Public Property Get Desconto2Codigo() As Object
     Set Desconto2Codigo = objUserControl.Controls("Desconto2Codigo")
End Property

Public Property Get Desconto3Codigo() As Object
     Set Desconto3Codigo = objUserControl.Controls("Desconto3Codigo")
End Property

Public Property Get Desconto1Percentual() As Object
     Set Desconto1Percentual = objUserControl.Controls("Desconto1Percentual")
End Property

Public Property Get Desconto3Valor() As Object
     Set Desconto3Valor = objUserControl.Controls("Desconto3Valor")
End Property

Public Property Get Desconto3Ate() As Object
     Set Desconto3Ate = objUserControl.Controls("Desconto3Ate")
End Property

Public Property Get Desconto2Valor() As Object
     Set Desconto2Valor = objUserControl.Controls("Desconto2Valor")
End Property

Public Property Get Desconto2Ate() As Object
     Set Desconto2Ate = objUserControl.Controls("Desconto2Ate")
End Property

Public Property Get Desconto1Valor() As Object
     Set Desconto1Valor = objUserControl.Controls("Desconto1Valor")
End Property

Public Property Get Desconto1Ate() As Object
     Set Desconto1Ate = objUserControl.Controls("Desconto1Ate")
End Property

Public Property Get DataVencimento() As Object
     Set DataVencimento = objUserControl.Controls("DataVencimento")
End Property

Public Property Get ValorParcela() As Object
     Set ValorParcela = objUserControl.Controls("ValorParcela")
End Property

Public Property Get Desconto2Percentual() As Object
     Set Desconto2Percentual = objUserControl.Controls("Desconto2Percentual")
End Property

Public Property Get Desconto3Percentual() As Object
     Set Desconto3Percentual = objUserControl.Controls("Desconto3Percentual")
End Property

Public Property Get GridParcelas() As Object
     Set GridParcelas = objUserControl.Controls("GridParcelas")
End Property

Public Property Get DataReferencia() As Object
     Set DataReferencia = objUserControl.Controls("DataReferencia")
End Property

Public Property Get PlacaUF() As Object
     Set PlacaUF = objUserControl.Controls("PlacaUF")
End Property

Public Property Get Placa() As Object
     Set Placa = objUserControl.Controls("Placa")
End Property

Public Property Get Transportadora() As Object
     Set Transportadora = objUserControl.Controls("Transportadora")
End Property

Public Property Get FilialEntrega() As Object
     Set FilialEntrega = objUserControl.Controls("FilialEntrega")
End Property

Public Property Get TipoFrete() As Object
     Set TipoFrete = objUserControl.Controls("TipoFrete")
End Property

Public Property Get TransportadoraLabel() As Object
     Set TransportadoraLabel = objUserControl.Controls("TransportadoraLabel")
End Property

Public Property Get CanalVenda() As Object
     Set CanalVenda = objUserControl.Controls("CanalVenda")
End Property

Public Property Get Mensagem() As Object
     Set Mensagem = objUserControl.Controls("Mensagem")
End Property

Public Property Get NumPedidoTerc() As Object
     Set NumPedidoTerc = objUserControl.Controls("NumPedidoTerc")
End Property

Public Property Get PesoLiquido() As Object
     Set PesoLiquido = objUserControl.Controls("PesoLiquido")
End Property

Public Property Get PesoBruto() As Object
     Set PesoBruto = objUserControl.Controls("PesoBruto")
End Property

Public Property Get MensagemLabel() As Object
     Set MensagemLabel = objUserControl.Controls("MensagemLabel")
End Property

Public Property Get VolumeMarca() As Object
     Set VolumeMarca = objUserControl.Controls("VolumeMarca")
End Property

Public Property Get VolumeEspecie() As Object
     Set VolumeEspecie = objUserControl.Controls("VolumeEspecie")
End Property

Public Property Get VolumeNumero() As Object
     Set VolumeNumero = objUserControl.Controls("VolumeNumero")
End Property

Public Property Get VolumeQuant() As Object
     Set VolumeQuant = objUserControl.Controls("VolumeQuant")
End Property

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
End Property

Public Property Get DescricaoItem() As Object
     Set DescricaoItem = objUserControl.Controls("DescricaoItem")
End Property

Public Property Get UnidadeMed() As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
End Property

Public Property Get Desconto() As Object
     Set Desconto = objUserControl.Controls("Desconto")
End Property

Public Property Get PercentDesc() As Object
     Set PercentDesc = objUserControl.Controls("PercentDesc")
End Property

Public Property Get PrecoUnitario() As Object
     Set PrecoUnitario = objUserControl.Controls("PrecoUnitario")
End Property

Public Property Get Quantidade() As Object
     Set Quantidade = objUserControl.Controls("Quantidade")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

Public Property Get PrecoTotal() As Object
     Set PrecoTotal = objUserControl.Controls("PrecoTotal")
End Property

Public Property Get GridItens() As Object
     Set GridItens = objUserControl.Controls("GridItens")
End Property

Public Property Get BotaoEstoqueProd() As Object
     Set BotaoEstoqueProd = objUserControl.Controls("BotaoEstoqueProd")
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

Public Property Get ICMSSubstValor1() As Object
     Set ICMSSubstValor1 = objUserControl.Controls("ICMSSubstValor1")
End Property

Public Property Get ICMSSubstBase1() As Object
     Set ICMSSubstBase1 = objUserControl.Controls("ICMSSubstBase1")
End Property

Public Property Get ICMSValor1() As Object
     Set ICMSValor1 = objUserControl.Controls("ICMSValor1")
End Property

Public Property Get ICMSBase1() As Object
     Set ICMSBase1 = objUserControl.Controls("ICMSBase1")
End Property

Public Property Get ValorProdutos() As Object
     Set ValorProdutos = objUserControl.Controls("ValorProdutos")
End Property

Public Property Get IPIValor1() As Object
     Set IPIValor1 = objUserControl.Controls("IPIValor1")
End Property

Public Property Get ISSValor1() As Object
     Set ISSValor1 = objUserControl.Controls("ISSValor1")
End Property

Public Property Get LabelTotais() As Object
     Set LabelTotais = objUserControl.Controls("LabelTotais")
End Property

Public Property Get ValorTotal() As Object
     Set ValorTotal = objUserControl.Controls("ValorTotal")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoPedido() As Object
     Set BotaoPedido = objUserControl.Controls("BotaoPedido")
End Property

Public Property Get BotaoVerPedido() As Object
     Set BotaoVerPedido = objUserControl.Controls("BotaoVerPedido")
End Property

Public Property Get FilialPedido() As Object
     Set FilialPedido = objUserControl.Controls("FilialPedido")
End Property

Public Property Get Pedido() As Object
     Set Pedido = objUserControl.Controls("Pedido")
End Property

Public Property Get PedidoLabel() As Object
     Set PedidoLabel = objUserControl.Controls("PedidoLabel")
End Property

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

Public Property Get Status() As Object
    Set Status = objUserControl.Controls("Status")
End Property

Public Property Get ImprimeRomaneio() As Object
     Set ImprimeRomaneio = objUserControl.Controls("ImprimeRomaneio")
End Property

Public Property Get DataEntrega() As Object
     Set DataEntrega = objUserControl.Controls("DataEntrega")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get ImprimeNFGravacao() As Object
     Set ImprimeNFGravacao = objUserControl.Controls("ImprimeNFGravacao")
End Property

Public Property Get BotaoImprimir() As Object
     Set BotaoImprimir = objUserControl.Controls("BotaoImprimir")
End Property

'Rastreamento
Public Property Get GridRastro() As Object
     Set GridRastro = objUserControl.Controls("GridRastro")
End Property

Public Property Get BotaoConsultaTitRec() As Object
     Set BotaoConsultaTitRec = objUserControl.Controls("BotaoConsultaTitRec")
End Property

Public Property Get MensagemCorpo() As Object
     Set MensagemCorpo = objUserControl.Controls("MensagemCorpo")
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

Public Property Get TipoPagto() As Object
     Set TipoPagto = objUserControl.Controls("TipoPagto")
End Property

Public Property Get BotaoTipoPagto() As Object
     Set BotaoTipoPagto = objUserControl.Controls("BotaoTipoPagto")
End Property

Public Property Get Recibo() As Object
     Set Recibo = objUserControl.Controls("Recibo")
End Property

Public Property Get NumNFe() As Object
     Set NumNFe = objUserControl.Controls("NumNFe")
End Property

Public Property Get CodVerificacao() As Object
     Set CodVerificacao = objUserControl.Controls("CodVerificacao")
End Property

Public Property Get EletronicaFed() As Object
     Set EletronicaFed = objUserControl.Controls("EletronicaFed")
End Property

'######################################
'EMBALAGENS
Public Property Get GridEmb() As Object
     Set GridEmb = objUserControl.Controls("GridEmb")
End Property

Public Property Get CalculoAuto() As Object
     Set CalculoAuto = objUserControl.Controls("CalculoAuto")
End Property

Public Property Get Embalagem() As Object
     Set Embalagem = objUserControl.Controls("Embalagem")
End Property
'EMBALAGENS
'######################################

Public Sub ImprimeRomaneio_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub BotaoLimparNF_Click()

    NFiscal.Caption = ""
    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
End Sub

Public Function CargaPosFormLoad(Optional bTrazendoDoc As Boolean = False) As Long

Dim lErro As Long

On Error GoTo Erro_CargaPosFormLoad

    If (giPosCargaOk = 0) Then

        'p/permitir o redesenho da tela
        DoEvents
        
        gbCarregandoTela = True
        
        Set gobjEmbalagem = CreateObject("RotinasMat.ClassMATEmb")
        Set gobjEmbalagem.objTela = Me
        
        'Rastreamento
        Set gobjRastreamento = New ClassRastreamento
        Set gobjRastreamento.objTela = Me
        gobjRastreamento.bSaidaMaterial = True
        gobjRastreamento.bEditavel = False
        
        Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
        Set gobjTelaProjetoInfo.objUserControl = objUserControl
        Set gobjTelaProjetoInfo.objTela = Me
        
        'se a empresa utiliza regras para calculo de comissoes
        If gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
            
            'instancia a classe q executa as regras
            Set objComissoesRegrasCalc = CreateObject("RotinasContab.ClassComissoesRegrasCalc")
            
            'instancia a classe q calcula os mnemonicos
            Set objMnemonicoComissCalc = New ClassMnemonicoComissCalc
            Set objMnemonicoComissCalcAux = New ClassMnemonicoComissCalcAux
        
            'setar o objtela para as classes que calculam mnemonicos
            Set objMnemonicoComissCalc.objTela = Me
            Set objMnemonicoComissCalcAux.objTela = Me
        
        End If

        lErro = gobjEmbalagem.Inicializa(DOCORIGEM_NF)
        If lErro <> SUCESSO Then gError 46777
        
        'Inicializa a parte de tributacao
        Call GridItens_InicializaColunas
        
        Set gobjTribTab = New ClassTribTab
        lErro = gobjTribTab.Ativar(Me)
        If lErro <> SUCESSO Then gError 46777
    
        'Carrega os Tipos de Notas Fiscais
        lErro = Carrega_TiposDocInfo()
        If lErro <> SUCESSO Then gError 130473
        
        If IsDate(DataEmissao.Text) Then Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
        Call gobjTribTab.Serie_Alterada(Serie.Text)
        
        'carregar a combo FilialPedido com as Filiais da empresa
        lErro = Carrega_FilialPedido()
        If lErro <> SUCESSO Then gError 46766
    
        'Carrega a combo de series
        lErro = Carrega_Serie()
        If lErro <> SUCESSO Then gError 46767
        
        'Seleciona série Padrão
        lErro = Seleciona_Serie_Padrao()
        If lErro <> SUCESSO Then gError 500142
        
        'Carrega a combo de tabelas de preços
        lErro = Carrega_TabelaPreco()
        If lErro <> SUCESSO Then gError 46768
    
        'Carrega a combo de transportadoras
        lErro = Carrega_Transportadora()
        If lErro <> SUCESSO Then gError 46769
    
        'Carrega a combo VolumeEspecie
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
        If lErro <> SUCESSO Then gError 102414
    
        'Carrega a combo VolumeMarca
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
        If lErro <> SUCESSO Then gError 102415
    
        'Carrega a combo PlacaUF
        lErro = Carrega_PlacaUF()
        If lErro <> SUCESSO Then gError 46770
    
        'ACrrega a combo CanalVenda
        lErro = Carrega_CanalVenda()
        If lErro <> SUCESSO Then gError 46771
        
        'Carrega na combo as Condições de Pagamento
        lErro = CF("Carrega_CondicaoPagamento", CondicaoPagamento, MODULO_CONTASARECEBER)
        If lErro <> SUCESSO Then gError 46890
        
        'Carrega os Tipos de Desconto
        lErro = Carrega_TipoDesconto()
        If lErro <> SUCESSO Then gError 46965
    
        'Preenche Data Referencia e Data de Emissão coma Data Atual
        DataReferencia.PromptInclude = False
        DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
        DataReferencia.PromptInclude = True
        
        'Inicializa as variáveis de Grid e de Browse
        Set objEventoMensagem = New AdmEvento
        Set objGridItens = New AdmGrid
        Set objGridComissoes = New AdmGrid
        Set objGridAlocacoes = New AdmGrid
        Set objGrid1 = New AdmGrid
        Set objGridParcelas = New AdmGrid
        Set objEventoPedido = New AdmEvento
        Set objEventoSerie = New AdmEvento
        Set objEventoNFiscal = New AdmEvento
        Set objEventoNaturezaOp = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoEstoqueProd = New AdmEvento
        Set objEventoVendedor = New AdmEvento
        Set objEventoLocalizacao = New AdmEvento
        Set objEventoTransportadora = New AdmEvento
        Set objEventoTranspRedespacho = New AdmEvento
        Set objEventoLote = New AdmEvento
        Set objEventoDoc = New AdmEvento
        Set objGridParcelas = New AdmGrid
        Set objEventoFornEntTerc = New AdmEvento
    
        Set gcolItemPedido = New Collection
        Set gcolItemQuant = New Collection
        Set gcolItemUM = New Collection 'Inserido por Wagner
        giDataReferenciaAlterada = 0
        
        'Inicializa a máscara do produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
        If lErro <> SUCESSO Then gError 46776
           
        'Inicializa a máscara do produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAlmox)
        If lErro <> SUCESSO Then gError 46776
           
        'Inicializa o grid de itens
        lErro = CF2(Me, "NFFP_Inicializa_Grid_Itens", objGridItens)
        If lErro <> SUCESSO Then gError 46772
        
        'Inicializa o grid de comissões
        lErro = objTabComissoes.Inicializa_Grid_Comissoes(objGridComissoes)
        If lErro <> SUCESSO Then gError 46773
        
        'Inicializa o grid de Alocações
        lErro = Inicializa_Grid_Alocacoes(objGridAlocacoes)
        If lErro <> SUCESSO Then gError 46774
        
        'Inicializa o grid de Parcelas
        lErro = Inicializa_Grid_Parcelas(objGridParcelas)
        If lErro <> SUCESSO Then gError 46775
        
        lErro = CF2(Me, "NFiscal_Inicializa_Grid_Parcelas", objGridParcelas)
        If lErro <> SUCESSO Then gError 183012
        
        PrecoUnitario.Format = gobjFAT.sFormatoPrecoUnitario
        
        QuantAlocada.Format = FORMATO_ESTOQUE
        Quantidade.Format = FORMATO_ESTOQUE
        QuantVendida.Format = FORMATO_ESTOQUE
    
        'Inicialização da parte de contabilidade
        lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_FATURAMENTO)
        If lErro <> SUCESSO Then gError 46778
    
        lErro = CF2(Me, "NFFP_CargaPosFormLoad", bTrazendoDoc)
        If lErro <> SUCESSO Then gError 83391
        
        gbCarregandoTela = False
        
        iAlterado = 0
        iFornEntTercAlterado = 0
        iComissoesAlterada = 0
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
          
        Case 46766 To 46778, 46890, 46965, 83391, 500142, 102414, 102415, 130473, 183012

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157298)
     
    End Select
     
    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    Set objTabComissoes.objTela = Me
    
    Call CF2(Me, "NFFP_Form_Load2")
    
    If giTipoVersao = VERSAO_LIGHT Then
        
        Botao(BOTAO_GRADE).Visible = False
        
        'sumir com abas
        TabStrip1.Tabs.Remove (TAB_Embalagens)
        TabStrip1.Tabs.Remove (TAB_RASTRO)
        TabStrip1.Tabs.Remove (TAB_Contabilizacao)
        
    End If
        
    MsgAutomatica.Value = vbChecked
    Mensagem.MaxLength = MENSAGEM_NF_MAXLENGTH
    MensagemCorpo.MaxLength = MENSAGEM_NF_MAXLENGTH
    
    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207434
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 207434

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157299)

    End Select
    
    iAlterado = 0
    iComissoesAlterada = 0
    
    Exit Sub

End Sub

Function Seleciona_Serie_Padrao() As Long

Dim lErro As Long
Dim sSeriePadrao As String
Dim objSerie As New ClassSerie

On Error GoTo Erro_Seleciona_Serie_Padrao
        
    'Lê série Padrão
    lErro = CF("Serie_Le_Padrao", sSeriePadrao)
    If lErro <> SUCESSO Then gError 500138
    
    'Se possui série padrão
    If Len(Trim(sSeriePadrao)) > 0 Then
    
        'Lê a série
        objSerie.sSerie = sSeriePadrao
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO And lErro <> 22202 Then gError 500139
        
        'Série não cadastrada
        If lErro = 22202 Then gError 500140
        
        'Se a Série Padrão é para fatura
        If objSerie.iTipoFormulario = TIPO_FORMULARIO_NFISCAL_FATURA Or objSerie.iTipoFormulario = TIPO_FORMULARIO_NFISCAL_FATURA_SERVICO Then
        
            'Coloca ela na combo
            Serie.Text = Desconverte_Serie_Eletronica(sSeriePadrao)
            If ISSerieEletronica(sSeriePadrao) Then
                EletronicaFed.Value = vbChecked
            Else
                EletronicaFed.Value = vbUnchecked
            End If
            Call Serie_Validate(bSGECancelDummy)
        
        End If
        
    End If
                    
    Seleciona_Serie_Padrao = SUCESSO
    
    Exit Function

Erro_Seleciona_Serie_Padrao:

    Seleciona_Serie_Padrao = gErr
    
    Select Case gErr
    
        Case 500138, 500139
        
        Case 500140
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157300)
    
    End Select
    
    Exit Function
    
End Function

Public Function Form_Load_Light()

Dim lErro As Long

On Error GoTo Erro_Form_Load_Light

    'obs.: a carga pos form_load será feita pela classe que trata nffat
    giPosCargaOk = 1
    
    'carregar a combo FilialPedido com as Filiais da empresa
    lErro = Carrega_FilialPedido()
    If lErro <> SUCESSO Then gError 25815
    
    'Carrega variáveis globais exclusivas dessa tela
    Set gcolItemPedido = New Collection
    Set gcolItemQuant = New Collection
    Set gcolItemUM = New Collection 'Inserido por Wagner
    
    'Inicializa as variáveis de Browse
    Set objEventoMensagem = New AdmEvento
    Set objEventoPedido = New AdmEvento
    Set objEventoSerie = New AdmEvento
    Set objEventoNFiscal = New AdmEvento
    Set objEventoNaturezaOp = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoEstoqueProd = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoLocalizacao = New AdmEvento
    Set objEventoTransportadora = New AdmEvento
    Set objEventoLote = New AdmEvento
    Set objEventoDoc = New AdmEvento
    
    Call GridItens_InicializaColunas
    
    'Colunas do Grid Alocacoes
    iGrid_Item_Col = 1
    iGrid_ProdutoAloc_Col = 2
    iGrid_AlmoxAloc_Col = 3
    iGrid_UMAloc_Col = 4
    iGrid_QuantAloc_Col = 5
    iGrid_QuantVend_Col = 6

    'Colunas do Grid Parcelas
    iGrid_Vencimento_Col = 1
    iGrid_ValorParcela_Col = 2
    iGrid_Desc1Codigo_Col = 3
    iGrid_Desc1Ate_Col = 4
    iGrid_Desc1Valor_Col = 5
    iGrid_Desc1Percentual_Col = 6
    iGrid_Desc2Codigo_Col = 7
    iGrid_Desc2Ate_Col = 8
    iGrid_Desc2Valor_Col = 9
    iGrid_Desc2Percentual_Col = 10
    iGrid_Desc3Codigo_Col = 11
    iGrid_Desc3Ate_Col = 12
    iGrid_Desc3Valor_Col = 13
    iGrid_Desc3Percentual_Col = 14

    iAlterado = 0
    iComissoesAlterada = 0
    
    Form_Load_Light = SUCESSO
    
    Exit Function

Erro_Form_Load_Light:

    Form_Load_Light = gErr

    Select Case gErr

        Case 25815

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157301)

    End Select
    
    iAlterado = 0
    iComissoesAlterada = 0
    
    Exit Function

End Function

Private Function Carrega_FilialPedido() As Long
'Carrega FilialPedido com as Filiais Empresas

Dim lErro As Long
Dim objFiliais As AdmFiliais

On Error GoTo Erro_Carrega_FilialPedido

    For Each objFiliais In gcolFiliais

        If objFiliais.iCodFilial <> EMPRESA_TODA Then
            'coloca na combo
            FilialPedido.AddItem CStr(objFiliais.iCodFilial) & SEPARADOR & objFiliais.sNome
            FilialPedido.ItemData(FilialPedido.NewIndex) = objFiliais.iCodFilial
        End If
    Next

    'Seleciona a Filial na qual o usuário entrou no Sistema
    Call Seleciona_FilialEmpresa

    Carrega_FilialPedido = SUCESSO

    Exit Function

Erro_Carrega_FilialPedido:

    Carrega_FilialPedido = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157302)

    End Select

    Exit Function

End Function

Private Function Carrega_Serie() As Long
'Carrega as combos de Série e serie de NF original com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie
Dim sSerieAnt As String

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 46780

    'Carrega na combo
    For Each objSerie In colSerie
        If UCase(sSerieAnt) <> UCase(Desconverte_Serie_Eletronica(objSerie.sSerie)) Then Serie.AddItem Desconverte_Serie_Eletronica(objSerie.sSerie)
        sSerieAnt = Desconverte_Serie_Eletronica(objSerie.sSerie)
    Next

    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case 46780

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157303)

    End Select

    Exit Function

End Function

Private Function Carrega_TabelaPreco() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TabelaPreco

    TabPrecoItemNF.AddItem ""
    TabPrecoItemNF.ItemData(TabPrecoItemNF.NewIndex) = 0
    
    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePrecoVenda", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao, "Ativo = 1")
    If lErro <> SUCESSO Then gError 46781

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        TabelaPreco.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TabelaPreco.ItemData(TabelaPreco.NewIndex) = objCodDescricao.iCodigo

        TabPrecoItemNF.AddItem CStr(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TabPrecoItemNF.ItemData(TabPrecoItemNF.NewIndex) = objCodDescricao.iCodigo
    
    Next

    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = gErr

    Select Case gErr

        Case 46781

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157304)

    End Select

    Exit Function

End Function

Private Function Carrega_Transportadora() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_Transportadora

    'Lê o Código e o NOme de Todas as Transportadoras do BD
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then gError 46782

    'Carrega a combo de Transportadoras
    For Each objCodigoNome In colCodigoNome
        Transportadora.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Transportadora.ItemData(Transportadora.NewIndex) = objCodigoNome.iCodigo
        TranspRedespacho.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        TranspRedespacho.ItemData(TranspRedespacho.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_Transportadora = SUCESSO

    Exit Function

Erro_Carrega_Transportadora:

    Carrega_Transportadora = gErr

    Select Case gErr

        Case 46782

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157305)

    End Select

    Exit Function

End Function

Private Function Carrega_CanalVenda() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_CanalVenda

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "CanalVenda", "Codigo", "NomeReduzido", STRING_CANAL_VENDA_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 46783

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        CanalVenda.AddItem CStr(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        CanalVenda.ItemData(CanalVenda.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_CanalVenda = SUCESSO

    Exit Function

Erro_Carrega_CanalVenda:

    Carrega_CanalVenda = gErr

    Select Case gErr

        Case 46783

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157306)

    End Select

    Exit Function

End Function
'
'Private Function Carrega_CondicaoPagamento() As Long
''Carrega na combo as Condições de Pagamento existentes
'
'Dim lErro As Long
'Dim colCod_DescReduzida As New AdmColCodigoNome
'Dim objCod_DescReduzida As AdmCodigoNome
'
'On Error GoTo Erro_Carrega_CondicaoPagamento
'
'    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
'    lErro = CF("CondicoesPagto_Le_Recebimento", colCod_DescReduzida)
'    If lErro <> SUCESSO Then gError 46891
'
'    For Each objCod_DescReduzida In colCod_DescReduzida
'        'Adiciona novo ítem na List da Combo CondicaoPagamento
'        CondicaoPagamento.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
'        CondicaoPagamento.ItemData(CondicaoPagamento.NewIndex) = objCod_DescReduzida.iCodigo
'    Next
'
'    Carrega_CondicaoPagamento = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_CondicaoPagamento:
'
'    Carrega_CondicaoPagamento = gErr
'
'    Select Case gErr
'
'        Case 46891
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157307)
'
'    End Select
'
'    Exit Function
'
'End Function

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

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157308)

    End Select

End Function

Private Function Carrega_TipoDesconto() As Long
'Carrega na combo os Tipos de Desconto existentes

Dim lErro As Long
Dim objCodDescricao As AdmCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_TipoDesconto

    Set colCodigoDescricao = gobjCRFAT.colTiposDesconto
    
    For Each objCodDescricao In colCodigoDescricao
        'Adiciona o ítem nas List's das Combos de Tipos Desconto
        Desconto1Codigo.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        Desconto1Codigo.ItemData(Desconto1Codigo.NewIndex) = objCodDescricao.iCodigo
        Desconto2Codigo.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        Desconto2Codigo.ItemData(Desconto2Codigo.NewIndex) = objCodDescricao.iCodigo
        Desconto3Codigo.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        Desconto3Codigo.ItemData(Desconto3Codigo.NewIndex) = objCodDescricao.iCodigo
    Next

    Carrega_TipoDesconto = SUCESSO

    Exit Function

Erro_Carrega_TipoDesconto:

    Carrega_TipoDesconto = gErr

    Select Case gErr

        Case 46966

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157309)

    End Select

    Exit Function

End Function

Private Sub Seleciona_FilialEmpresa()

Dim iIndice As Integer

    'Se a filial da empresa não for EMPRESA_TODA
    If giFilialEmpresa <> EMPRESA_TODA Then
        'Seleciona na combo de filial do pedido a filial empresa
        For iIndice = 0 To FilialPedido.ListCount - 1
            If FilialPedido.ItemData(iIndice) = giFilialEmpresa Then
                FilialPedido.ListIndex = iIndice
                Exit For
            End If
        Next

    Else
        FilialPedido.ListIndex = 0
    End If

End Sub

Private Sub GridItens_InicializaColunas()
    Call CF2(Me, "NFFP_GridItens_InicializaColunas")
End Sub

Public Function NFFP_GridItens_InicializaColunas(ByVal objCT As Object) As Long

Dim iColunasEscondidas As Integer
    
    'Colunas da Grid
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        iGrid_TabPrecoItem_Col = 5
    Else
        iColunasEscondidas = iColunasEscondidas + 1
    End If
    
    iGrid_ValorUnitario_Col = 6 - iColunasEscondidas
    iGrid_PercDesc_Col = 7 - iColunasEscondidas
    iGrid_Desconto_Col = 8 - iColunasEscondidas
    iGrid_PrecoTotal_Col = 9 - iColunasEscondidas
    iGrid_PrecoTotalB_Col = 10 - iColunasEscondidas
    iGrid_DataEntrega_Col = 11 - iColunasEscondidas
    
    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        iGrid_ComissaoItem_Col = 12 - iColunasEscondidas
    Else
        iColunasEscondidas = iColunasEscondidas + 1
    End If

    NFFP_GridItens_InicializaColunas = SUCESSO
    
End Function

'@@@@@ customizada
Public Function NFFP_Inicializa_Grid_Itens(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        objGridInt.colColuna.Add ("Tab. Preço")
    End If
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Data Entrega")

    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        objGridInt.colColuna.Add ("Comissão")
    End If

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    
    If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        objGridInt.colCampo.Add (TabPrecoItemNF.Name)
    Else
        TabPrecoItemNF.left = POSICAO_FORA_TELA
    End If
    
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    objGridInt.colCampo.Add (PrecoTotalB.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)

    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        objGridInt.colCampo.Add (ComissaoItemNF.Name)
    Else
        ComissaoItemNF.left = POSICAO_FORA_TELA
    End If
    
    Call GridItens_InicializaColunas

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NFFP_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Alocacoes(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Alocações

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant. Alocada")
    objGridInt.colColuna.Add ("Quant. Vendida")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ItemNFiscal.Name)
    objGridInt.colCampo.Add (ProdutoAlmox.Name)
    objGridInt.colCampo.Add (Almox.Name)
    objGridInt.colCampo.Add (UnidadeMedEst.Name)
    objGridInt.colCampo.Add (QuantAlocada.Name)
    objGridInt.colCampo.Add (QuantVendida.Name)

    'Colunas da Grid
    iGrid_Item_Col = 1
    iGrid_ProdutoAloc_Col = 2
    iGrid_AlmoxAloc_Col = 3
    iGrid_UMAloc_Col = 4
    iGrid_QuantAloc_Col = 5
    iGrid_QuantVend_Col = 6

    'Grid do GridInterno
    objGridInt.objGrid = GridAlocacao

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ALOCACOES + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 13

    'Largura da primeira coluna
    GridAlocacao.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Alocacoes = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Parcelas

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Forma de Pagamento")
    objGridInt.colColuna.Add ("Desconto 1 Tipo")
    objGridInt.colColuna.Add ("Desc. 1 Data")
    objGridInt.colColuna.Add ("Desc. 1 Valor")
    objGridInt.colColuna.Add ("Desc. 1 %")
    objGridInt.colColuna.Add ("Desconto 2 Tipo")
    objGridInt.colColuna.Add ("Desc. 2 Data")
    objGridInt.colColuna.Add ("Desc. 2 Valor")
    objGridInt.colColuna.Add ("Desc. 2 %")
    objGridInt.colColuna.Add ("Desconto 3 Tipo")
    objGridInt.colColuna.Add ("Desc. 3 Data")
    objGridInt.colColuna.Add ("Desc. 3 Valor")
    objGridInt.colColuna.Add ("Desc. 3 %")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (TipoPagto.Name)
    objGridInt.colCampo.Add (Desconto1Codigo.Name)
    objGridInt.colCampo.Add (Desconto1Ate.Name)
    objGridInt.colCampo.Add (Desconto1Valor.Name)
    objGridInt.colCampo.Add (Desconto1Percentual.Name)
    objGridInt.colCampo.Add (Desconto2Codigo.Name)
    objGridInt.colCampo.Add (Desconto2Ate.Name)
    objGridInt.colCampo.Add (Desconto2Valor.Name)
    objGridInt.colCampo.Add (Desconto2Percentual.Name)
    objGridInt.colCampo.Add (Desconto3Codigo.Name)
    objGridInt.colCampo.Add (Desconto3Ate.Name)
    objGridInt.colCampo.Add (Desconto3Valor.Name)
    objGridInt.colCampo.Add (Desconto3Percentual.Name)

    'Colunas do Grid
    iGrid_Vencimento_Col = 1
    iGrid_ValorParcela_Col = 2
    iGrid_TipoPagto_Col = 3
    iGrid_Desc1Codigo_Col = 4
    iGrid_Desc1Ate_Col = 5
    iGrid_Desc1Valor_Col = 6
    iGrid_Desc1Percentual_Col = 7
    iGrid_Desc2Codigo_Col = 8
    iGrid_Desc2Ate_Col = 9
    iGrid_Desc2Valor_Col = 10
    iGrid_Desc2Percentual_Col = 11
    iGrid_Desc3Codigo_Col = 12
    iGrid_Desc3Ate_Col = 13
    iGrid_Desc3Valor_Col = 14
    iGrid_Desc3Percentual_Col = 15
    
    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 700

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)
    
Dim lErro As Long
        
    Call CF2(Me, "NFFP_Form_Unload", Cancel)
    Set gobjInfoUsu = Nothing

    Set gobjTelaProjetoInfo = Nothing
    Set gobjInfoAdic = Nothing

    'Libera as variáveis globais da tela
    Set gobjNFiscal = Nothing
    Set objGridItens = Nothing
    Set objGridComissoes = Nothing
    Set objGridAlocacoes = Nothing
    Set objGridParcelas = Nothing
    Set objGrid1 = Nothing

    Set objEventoMensagem = Nothing
    Set objEventoPedido = Nothing
    Set objEventoSerie = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoNaturezaOp = Nothing
    Set objEventoProduto = Nothing
    Set objEventoEstoqueProd = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoLocalizacao = Nothing
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    Set objEventoFornEntTerc = Nothing

    'Alterado por Tulio em 10/04
    'se foi instanciada a classe de execucao de regras de comissoes => libera
    Set objComissoesRegrasCalc = Nothing
        
    'se foi instanciada a classe que calcula os mnemonicos => libera
    Set objMnemonicoComissCalc = Nothing
    Set objMnemonicoComissCalcAux = Nothing


    Set objEventoTransportadora = Nothing
    Set objEventoTranspRedespacho = Nothing

    Set objContabil = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    'Embalagem
    Set gobjEmbalagem = Nothing

    If Not (gobjTribTab Is Nothing) Then Call gobjTribTab.Desativar
    
    'Fecha o Comando de Setas
    lErro = ComandoSeta_Liberar(Me.Name)

    Set gobjAnotacao = Nothing
    
    Set gobjRastreamento = Nothing

    Set objTabComissoes = Nothing
    Set gobjTribTab = Nothing

    Set gcolItemPedido = Nothing
    Set gcolItemQuant = Nothing
    Set gcolItemUM = Nothing 'Inserido por Wagner

    Set objUserControl = Nothing
        
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157310)
    
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

Public Sub EscaninhoRastro_Change()
'Rastreamento
    Call gobjRastreamento.EscaninhoRastro_Change
End Sub

Public Sub EscaninhoRastro_Click()
'Rastreamento
    Call gobjRastreamento.EscaninhoRastro_Click
End Sub

Public Sub EscaninhoRastro_GotFocus()
'Rastreamento
    Call gobjRastreamento.EscaninhoRastro_GotFocus
End Sub

Public Sub EscaninhoRastro_KeyPress(KeyAscii As Integer)
'Rastreamento
    Call gobjRastreamento.EscaninhoRastro_KeyPress(KeyAscii)
End Sub

Public Sub EscaninhoRastro_Validate(Cancel As Boolean)
'Rastreamento
    Call gobjRastreamento.EscaninhoRastro_Validate(Cancel)
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

Public Sub DataEmissao_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
End Sub

Private Sub DataEmissaoLabel_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataReferencia_Change()
    iAlterado = REGISTRO_ALTERADO
    giDataReferenciaAlterada = REGISTRO_ALTERADO
End Sub

Public Sub DataReferencia_GotFocus()

Dim iDataAux As Integer
    
    iDataAux = giDataReferenciaAlterada
    Call MaskEdBox_TrataGotFocus(DataReferencia, iAlterado)
    giDataReferenciaAlterada = iDataAux
    
End Sub

Public Sub DataSaida_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataSaida, iAlterado)
End Sub

Public Sub TipoFrete_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NFiscal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NFiscal_GotFocus()
    Call MaskEdBox_TrataGotFocus(NFiscal, iAlterado)
End Sub

Public Sub Pedido_GotFocus()
    Call MaskEdBox_TrataGotFocus(Pedido, iAlterado)
End Sub

Public Sub TabStrip1_Click()
    
Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

''*** incluidos p/tratamento de tributacao *******************************
        'se abriu o tab de tributacao
        If TabStrip1.SelectedItem.Index = TAB_Tributacao Then
            lErro = gobjTribTab.TabClick
            If lErro <> SUCESSO Then gError 46785
        End If
'*** fim tributacao

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If UCase(left(TabStrip1.SelectedItem.Caption, 6)) = UCase(left(TITULO_TAB_CONTABILIDADE, 6)) Then Call objContabil.Contabil_Carga_Modelo_Padrao

        Select Case iFrameAtual
        
            Case TAB_DadosPrincipais
                Parent.HelpContextID = IDH_NF_FATURA_PEDIDO_DADOS_PRINCIPAIS
            
            Case TAB_Itens
                Parent.HelpContextID = IDH_NF_FATURA_PEDIDO_ITENS
            
            Case TAB_Complemento
                Call Calcula_Msg
                Parent.HelpContextID = IDH_NF_FATURA_PEDIDO_COMPLEMENTO
            
            Case TAB_Cobranca
                Parent.HelpContextID = IDH_NF_FATURA_PEDIDO_COBRANCA
            
            Case TAB_Comissoes
                'se for tab de comissoes e a empresa
                'utilizar regras para o calculo de comissoes
                lErro = objTabComissoes.Comissoes_Calcula
                If lErro <> SUCESSO Then gError 101536
                
                Parent.HelpContextID = IDH_PEDIDO_VENDA_COMISSOES
            
            Case TAB_Almoxarifado
                Parent.HelpContextID = IDH_NF_FATURA_PEDIDO_ALMOXARIFADO
            
            Case TAB_Tributacao
                Parent.HelpContextID = IDH_NF_FATURA_PEDIDO_TRIBUTACAO
            
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_NF_FATURA_PEDIDO_CONTABILIZACAO
        
        End Select
    
    End If

    Exit Sub
    
Erro_TabStrip1_Click:

    Select Case gErr

        Case 46785, 101536
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157311)

    End Select

    Exit Sub

End Sub

Public Sub Pedido_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialPedido_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialPedido_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialPedido_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_FilialPedido_Validate

    'Se não estiver preenchida, sai.
    If Len(Trim(FilialPedido.Text)) = 0 Then Exit Sub
    'Se não foi selecoinada, sai.
    If FilialPedido.ListIndex <> -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialPedido, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 46786
    If lErro <> SUCESSO Then gError 46787 'Não encontrou.
    
    Exit Sub

Erro_FilialPedido_Validate:

    Cancel = True

    Select Case gErr

        Case 46786

        Case 46787
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialPedido.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157312)

    End Select

    Exit Sub

End Sub

Public Sub Serie_Change()
    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = REGISTRO_ALTERADO
End Sub

Public Sub Serie_Click()
    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = REGISTRO_ALTERADO
End Sub

Public Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Serie_Validate

    If iSerieAlterada <> 0 Then
    
        'Verifica se foi preenchida
        If Len(Trim(Serie.Text)) = 0 Then Exit Sub
        
        'Verifica se foi selecionada
        If Serie.ListIndex = -1 Then
        
            'Tenta selecionar a serie
            lErro = Combo_Item_Igual(Serie)
            If lErro <> SUCESSO And lErro <> 12253 Then gError 46790
            'Se não está na combo
            If lErro <> SUCESSO Then
            
                objSerie.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
                'Busca a série no BD
                lErro = CF("Serie_Le", objSerie)
                If lErro <> SUCESSO And lErro <> 22202 Then gError 51584
                If lErro <> SUCESSO Then gError 46789 'Se não encontrou
            
            End If

        End If
        
''*** incluidos p/tratamento de tributacao *******************************
        Call gobjTribTab.Serie_Alterada(Serie.Text)
'*** fim tributacao
        
        iSerieAlterada = 0
    
    End If
    
    Exit Sub

Erro_Serie_Validate:

    Cancel = True

    Select Case gErr

        Case 46789
            'Pergunta se deseja criar nova série
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_SERIE", Serie.Text)
            'Se a resposta for afirmativa
            If vbMsgRes = vbYes Then
                'Chama a tela de cadastro de séries
                Call Chama_Tela("SerieNFiscal", objSerie)
            End If
            'segura o foco na série

        Case 46790, 51584

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157313)

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

Private Sub Class_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
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

Public Sub NFiscalLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As Collection

    'Chama a tela de browse das notas fiscais geradas a partir de um pedido de venda
    Call Chama_Tela("NFiscalFaturaPedidoLista", colSelecao, objNFiscal, objEventoNFiscal)

End Sub

Private Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objNFiscal = obj1

    'Evita ter que alterar os browses a cada campo inserido na tabela de NF
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 180101

    'Coloca na Tela a Nota Fiscal escolhida
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 46791

    glCodPedido = objNFiscal.lNumPedidoVenda
    giFilialPedido = objNFiscal.iFilialPedido

    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case gErr

        Case 46791, 180101
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157314)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO
    
    iAlterado = REGISTRO_ALTERADO
    giDataEmissaoAlterada = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_DataEmissao_Validate

    If giDataEmissaoAlterada = 0 Then Exit Sub

    'Verifica se a data de emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 46793

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
'*** fim tributacao

    giDataEmissaoAlterada = 0

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 46793
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157315)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 46794

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 46794

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157316)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 46795

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 46795

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157317)

    End Select

    Exit Sub

End Sub

Public Sub DataSaida_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
    
Public Sub DataSaida_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataSaida_Validate

    'Verifica se a data de Saida foi digitada
    If Len(Trim(DataSaida.ClipText)) <> 0 Then

        'Critica a data digitada
        lErro = Data_Critica(DataSaida.Text)
        If lErro <> SUCESSO Then gError 46796

    End If
    
    lErro = Cobranca_Automatica()
    If lErro <> SUCESSO Then gError 56909
    
    Exit Sub

Erro_DataSaida_Validate:

    Cancel = True

    Select Case gErr
        Case 46796, 56909

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157318)

    End Select

    Exit Sub

End Sub

'horasaida
Public Sub HoraSaida_GotFocus()
    Call MaskEdBox_TrataGotFocus(HoraSaida, iAlterado)
End Sub

'horasaida
Public Sub HoraSaida_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

'horasaida
Public Sub HoraSaida_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HoraSaida_Validate

    'Verifica se a hora de saida foi digitada
    If Len(Trim(HoraSaida.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(HoraSaida.Text)
    If lErro <> SUCESSO Then gError 89799

    Exit Sub

Erro_HoraSaida_Validate:

    Cancel = True

    Select Case gErr

        Case 89799

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157319)

    End Select

    Exit Sub

End Sub

Public Sub UpDownSaida_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownSaida_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataSaida, DIMINUI_DATA)
    If lErro Then gError 46797

    Exit Sub

Erro_UpDownSaida_DownClick:

    Select Case gErr

        Case 46797

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157320)

    End Select

    Exit Sub

End Sub

Public Sub UpDownSaida_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownSaida_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataSaida, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 46798

    Exit Sub

Erro_UpDownSaida_UpClick:

    Select Case gErr

        Case 46798

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157321)

    End Select

    Exit Sub

End Sub

Public Sub TabelaPreco_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TabelaPreco_Click()

Dim lErro As Long

On Error GoTo Erro_TabelaPreco_Click

    iAlterado = REGISTRO_ALTERADO

    'Verifica se a Tabela de Preços foi informada
    If TabelaPreco.ListIndex = -1 Then Exit Sub

    'Verifica se existe algum Ítem no Grid
    If objGridItens.iLinhasExistentes = 0 Then Exit Sub

    'Faz o tratamento para a Tabela Preco selecionada
    lErro = Trata_TabelaPreco()
    If lErro <> SUCESSO Then gError 46799

    Exit Sub

Erro_TabelaPreco_Click:

    Select Case gErr

        Case 46799

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157322)

    End Select

    Exit Sub

End Sub

Public Sub TabelaPreco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTabelaPreco As New ClassTabelaPreco
Dim iCodigo As Integer

On Error GoTo Erro_TabelaPreco_Validate

    'Verifica se foi preenchida a ComboBox TabelaPreco
    If Len(Trim(TabelaPreco.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox TabelaPreco
    If TabelaPreco.Text = TabelaPreco.List(TabelaPreco.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TabelaPreco, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 46800

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 46801
        
        If lErro <> SUCESSO Then gError 46803 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 46802

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 46804

    Exit Sub

Erro_TabelaPreco_Validate:

    Cancel = True


    Select Case gErr

        Case 46800, 46801, 46802
    
        Case 46803  'Não encontrou Tabela de Preço no BD
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")
    
            If vbMsgRes = vbYes Then
                'Chama a tela de Tabelas de Preço
                Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
            Else
                'Segura o foco
            End If
    
        Case 46804
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157323)

    End Select

    Exit Sub

End Sub

Public Sub CondicaoPagamento_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CondicaoPagamento_Click()

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult
Dim dPercAcresFin As Double

On Error GoTo Erro_CondicaoPagamento_Click
   
    'Verifica se alguma Condição foi selecionada
    If CondicaoPagamento.ListIndex = -1 Then Exit Sub
    
    'Passa o código da Condição para objCondicaoPagto
    objCondicaoPagto.iCodigo = CondicaoPagamento.ItemData(CondicaoPagamento.ListIndex)

    'Lê Condição de Pagamento à partir do código
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 46805
    
    'Não encontrou a Condição de Pagamento --> erro
    If lErro = 19205 Then gError 46807
    
    If Not gbCarregandoTela Then
    
        dPercAcresFin = StrParaDbl(PercAcrescFin.Text)
        
        If Abs(dPercAcresFin - (objCondicaoPagto.dAcrescimoFinanceiro * 100)) > DELTA_VALORMONETARIO Then
            
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
            If vbMsgRes = vbYes Then
                PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
                Call PercAcrescFin_Validate(bSGECancelDummy)
            End If
        
        End If
    
        'Testa se ValorTotal está preenchido
        If Len(Trim(ValorTotal)) > 0 Then
        
            'Testa se DataEmissao está preenchida e ValorTotal é positivo
            If Len(Trim(DataReferencia.ClipText)) > 0 And (CDbl(ValorTotal.Caption) > 0) Then
                
                'Preenche o GridParcelas
                lErro = Cobranca_Automatica()
                If lErro <> SUCESSO Then gError 46806
                
            End If
        End If
    
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_CondicaoPagamento_Click:

    Select Case gErr

        Case 46805, 46806

        Case 46807
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157324)
           
      End Select
            
    Exit Sub

End Sub

Public Sub CondicaoPagamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Condicaopagamento_Validate

    'Verifica se a Condicaopagamento foi preenchida
    If Len(Trim(CondicaoPagamento.Text)) = 0 Then Exit Sub

    'Verifica se é uma Condicaopagamento selecionada
    If CondicaoPagamento.Text = CondicaoPagamento.List(CondicaoPagamento.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(CondicaoPagamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 46808

    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condição Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 46809
        
        'Não encontrou a Condição de Pagamento
        If lErro = 19205 Then gError 46811

        'Testa se pode ser usada em Contas a Receber
        If objCondicaoPagto.iEmRecebimento = 0 Then gError 46812

        'Coloca na Tela
        CondicaoPagamento.Text = CondPagto_Traz(objCondicaoPagto)
        
        'Se ValorTotal e DataEmissao estiverem preenchidos, preenche GridParcelas
        If Len(Trim(ValorTotal)) > 0 Then
            If Len(Trim(DataReferencia.ClipText)) > 0 And CDbl(ValorTotal.Caption) > 0 Then

                'Preenche o GridParcelas
                lErro = Cobranca_Automatica()
                If lErro <> SUCESSO Then gError 46810

            End If
        End If

    End If

    'Não encontrou o valor que era STRING
    If lErro = 6731 Then gError 46813

    Exit Sub

Erro_Condicaopagamento_Validate:

    Cancel = True


    Select Case gErr

       Case 46808, 46809, 46810

       Case 46811
            'Pergunta se deseja criar a condição de pagamento
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)
            'Se a resposta for positiva
            If vbMsgRes = vbYes Then
                'Chama a tela de Condições de Pagamento
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            Else
            End If

        Case 46812
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", gErr, objCondicaoPagto.iCodigo)

        Case 46813
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157325)

    End Select

    Exit Sub

End Sub

Public Sub PercAcrescFin_Change()
    iAlterado = REGISTRO_ALTERADO
    iPercAcresFinAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercAcrescFin_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_PercAcrescFin_Validate

    'Verific se o percentual foi alterado
    If iPercAcresFinAlterado = 0 Then Exit Sub
    'Se estiver preenchido
    If Len(Trim(PercAcrescFin.Text)) > 0 Then
        lErro = Porcentagem_Critica_Negativa(PercAcrescFin)
        If lErro <> SUCESSO Then gError 26717
    End If

    If Len(Trim(TabelaPreco.Text)) > 0 Then

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_RECALCULAR_PRECOS")
        If vbMsgRes = vbYes Then
            lErro = Trata_TabelaPreco()
            If lErro <> SUCESSO Then gError 46814
        End If
        
    End If
    
    iPercAcresFinAlterado = 0
    
    Exit Sub
    
Erro_PercAcrescFin_Validate:

    Cancel = True


    Select Case gErr
    
        Case 26717

        Case 46814
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157326)
            
    End Select
    
    Exit Sub

End Sub

Public Sub ValorComissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorComissao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComissoes)
End Sub

Public Sub ValorComissao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
End Sub

Public Sub ValorComissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorComissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorFrete_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    iValorFreteAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorFrete_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorFrete As Double

On Error GoTo Erro_ValorFrete_Validate

    If iValorFreteAlterado = 0 Then Exit Sub

    'Verifica se o Valor Frete foi preenchido
    If Len(Trim(ValorFrete.Text)) > 0 Then
    
        'Faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
        If lErro <> SUCESSO Then gError 46815

        dValorFrete = CDbl(ValorFrete.Text)
        
        'Coloca o Valor Formatado na tela
        ValorFrete.Text = Format(dValorFrete, "Fixed")

    End If
    
    iValorFreteAlterado = 0

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorFrete_Validate(Cancel, dValorFrete)
'*** fim tributacao
    
    Call ValorTotal_Calcula

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True

    Select Case gErr

        Case 46815

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157327)

    End Select

    Exit Sub

End Sub

Public Sub ValorProdutos_Change()

Dim dValorDesconto As Double
Dim dValorProdutos As Double
    
    If gbCarregandoTela Then Exit Sub
    
'    If Len(Trim(ValorProdutos.Caption)) <> 0 And IsNumeric(ValorProdutos.Caption) Then
'
'        dValorProdutos = CDbl(ValorProdutos.Caption)
'
'        'Se o cliente possui desconto e o campo desconto não foi alterado pelo usuário
'        If gdDesconto > 0 And iValorDescontoAlterado = 0 Then
'
'            'Calcula o valor do desconto para o cliente e coloca na tela
'            dValorDesconto = gdDesconto * dValorProdutos
'            ValorDesconto.Text = Format(dValorDesconto, "Standard")
'            iValorDescontoAlterado = 0
'
'''*** incluidos p/tratamento de tributacao *******************************
'            Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, dValorDesconto)
''*** fim tributacao
'
'        End If
'
'    End If

End Sub

Public Sub ValorSeguro_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    iValorSeguroAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorSeguro_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorSeguro As Double

On Error GoTo Erro_Valorseguro_Validate

    If iValorSeguroAlterado = 0 Then Exit Sub

    'Se estiver preenchido
    If Len(Trim(ValorSeguro.Text)) > 0 Then
        
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(ValorSeguro.Text)
        If lErro <> SUCESSO Then gError 46816

        dValorSeguro = CDbl(ValorSeguro.Text)
        
        'Coloca o valor formatado na tela
        ValorSeguro.Text = Format(dValorSeguro, "Fixed")

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

        Case 46816

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157328)

    End Select

    Exit Sub

End Sub

Public Sub ValorDespesas_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    iValorDespesasAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorDespesas_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorDespesas As Double

On Error GoTo Erro_ValorDespesas_Validate

    If iValorDespesasAlterado = 0 Then Exit Sub

    'Se  estiver preenchido
    If Len(Trim(ValorDespesas.Text)) > 0 Then

        'Faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorDespesas.Text)
        If lErro <> SUCESSO Then gError 46817

        dValorDespesas = CDbl(ValorDespesas.Text)

        'coloca o valor formatado na tela
        ValorDespesas.Text = Format(dValorDespesas, "Fixed")

    End If

    iValorDespesasAlterado = 0

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDespesas_Validate(Cancel, dValorDespesas)
'*** fim tributacao
    
    Call ValorTotal_Calcula

    Exit Sub

Erro_ValorDespesas_Validate:

    Cancel = True

    Select Case gErr

        Case 46817

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157329)

    End Select

    Exit Sub

End Sub

Public Sub ValorDesconto_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    iValorDescontoAlterado = REGISTRO_ALTERADO
    gdDesconto = 0
    
End Sub

Public Sub ValorDesconto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorDesconto As Double
Dim dValorProdutos As Double

On Error GoTo Erro_ValorDesconto_Validate

    'Verifica se o valor foi alterado
    If iValorDescontoAlterado = 0 Then Exit Sub

    dValorDesconto = 0

    Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)

    'Verifica se o Valor está preenchido
    If Len(Trim(ValorDesconto.Text)) > 0 Then
    
        'Faz a Crítica do Valor digitado
        lErro = Valor_NaoNegativo_Critica(ValorDesconto.Text)
        If lErro <> SUCESSO Then gError 46818

        dValorDesconto = CDbl(ValorDesconto.Text)
        'Coloca o Valor formatado na tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")

        'Verifica se o Valor do Desconto é maior que a soma dos demais valores
        If dValorDesconto > dValorProdutos Then gError 46819

        dValorProdutos = dValorProdutos - dValorDesconto

    End If

    ValorProdutos.Caption = Format(dValorProdutos, "Standard")

    lErro = CF2(Me, "NFFP_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132034

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDesconto_Validate(Cancel, dValorDesconto)
'*** fim tributacao
    
    Call ValorTotal_Calcula

    iValorDescontoAlterado = 0
    
    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True

    Select Case gErr

        Case 46818

        Case 46819
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos)

        Case 132034

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157330)

    End Select

    Exit Sub

End Sub

Public Function Cobranca_Automatica() As Long
'recalcula o tab de cobranca
 
Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Cobranca_Automatica

    'Se automática estiver selecionada e a condicao de pagamento estiver preenchida
    If CobrancaAutomatica.Value = 1 And Len(Trim(CondicaoPagamento.Text)) <> 0 Then
        'Pega a condicao de pagamento da tela
        objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)
        
        'Lê a condição de pagamento
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 46820
        If lErro <> SUCESSO Then gError 46822
        'Preenche o grid parcelas de acordo com a condição de pagamento
        lErro = GridParcelas_Preenche(objCondicaoPagto)
        If lErro <> SUCESSO Then gError 46821
 
    End If
    
    Cobranca_Automatica = SUCESSO
    
    Exit Function
    
Erro_Cobranca_Automatica:
    
    Cobranca_Automatica = gErr
    
    Select Case gErr
    
        Case 46820, 46821
        
        Case 46822
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 157331)
    
    End Select
    
    Exit Function
        
End Function

Public Sub BotaoEstoqueProd_Click()

Dim lErro As Long
Dim objEstoqueProd As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_BotaoEstoqueProd_Click

    lErro = CF2(Me, "NFFP_Botao_Estoque", objEstoqueProd, colSelecao, sSelecao)
    If lErro <> SUCESSO Then gError 131888

    Call Chama_Tela("EstoqueProdutoLista", colSelecao, objEstoqueProd, objEventoEstoqueProd, sSelecao)

    Exit Sub
    
Erro_BotaoEstoqueProd_Click:

    Select Case gErr

        Case 131888
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157332)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoEstoqueProd_evSelecao(obj1 As Object)
    Me.Show
End Sub

Public Sub BotaoProdutos_Click()

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is Produto Then
    
        sProduto1 = Produto.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 43701

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 83086
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr
    
        Case 43701
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 83086
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157333)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)
    Me.Show
End Sub

Public Sub FilialEntrega_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialEntrega_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialEntrega_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialEntrega_Validate

    'Verifica se a FilialEntrega foi preenchida
    If Len(Trim(FilialEntrega.Text)) = 0 Then Exit Sub

    'Verifica se é uma FilialEntrega selecionada
    If FilialEntrega.ListIndex <> -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialEntrega, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 46823

    'Se nao encontra o item com o código informado
    If lErro = 6730 Then

        'Verifica de o Cliente foi digitado
        If Len(Trim(Cliente.Caption)) = 0 Then gError 46826

        sCliente = Cliente.Caption

        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe FilialEntrega com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 46824

        If lErro = 17660 Then gError 46825

        'coloca na tela
        FilialEntrega.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 46827

    Exit Sub

Erro_FilialEntrega_Validate:

    Cancel = True


    Select Case gErr

        Case 46823, 46824

        Case 46825
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Caption)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If
            
        Case 46826
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 46827
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialEntrega.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157334)

    End Select

    Exit Sub

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
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Transportadora_Validate

    'Verifica se foi preenchida a ComboBox Transportadora
    If Len(Trim(Transportadora.Text)) = 0 Then Exit Sub
    'Verifica se está preenchida com o item selecionado na ComboBox Transportadora
    If Transportadora.Text = Transportadora.List(Transportadora.ListIndex) Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Transportadora, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 46828
    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 46829

        If lErro = 19250 Then gError 46830 'Não encontrou Transportadora no BD

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    End If

    'Não existe o item com a String na List da Combo
    If lErro = 6731 Then gError 46831

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case gErr

        Case 46828, 46829

        Case 46830
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then
                Call Chama_Tela("Transportadora", objTransportadora)
            End If


        Case 46831
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157335)

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

Private Sub objEventoTransportadora_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora

    Set objTransportadora = obj1

    'Coloca a Transportdora na Tela
    Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    Me.Show

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

    'Verifica se existe o item na combo
    lErro = Combo_Item_Igual(PlacaUF)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 46832

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 46833

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True

    Select Case gErr

        Case 46832

        Case 46833
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157336)

    End Select

    Exit Sub

End Sub

Public Sub VolumeNumero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub VolumeQuant_Change()
    iAlterado = REGISTRO_ALTERADO
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
    If lErro <> SUCESSO Then gError 102416
    
    Exit Sub

Erro_VolumeEspecie_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102416
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157337)

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
    If lErro <> SUCESSO Then gError 102417
    
    Exit Sub

Erro_VolumeMarca_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102417
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157338)

    End Select

End Sub

Public Sub Mensagem_Change()
    iAlterado = REGISTRO_ALTERADO
    MsgAutomatica.Value = vbUnchecked
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
    If lErro <> SUCESSO Then gError 46834

    'Coloca o valor formatado na Tela
    PesoLiquido.Text = Format(PesoLiquido.Text, PesoLiquido.Format) 'Alterado por Wagner

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True

    Select Case gErr

        Case 46834

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157339)

    End Select

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
    lErro = Valor_Positivo_Critica(PesoBruto.Text)
    If lErro <> SUCESSO Then gError 46835

    'Coloca o valor formatado na Tela
    PesoBruto.Text = Format(PesoBruto.Text, PesoBruto.Format) 'Alterado por Wagner

    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True


    Select Case gErr

        Case 46835

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157340)

    End Select

    Exit Sub

End Sub

Public Sub CanalVenda_Change()
    iAlterado = REGISTRO_ALTERADO
    If gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
        iComissoesAlterada = 1 'preventivamente, pois pode ter alterado algo que afete uma regra
    End If
End Sub

Public Sub CanalVenda_Click()
    iAlterado = REGISTRO_ALTERADO
    If gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
        iComissoesAlterada = 1 'preventivamente, pois pode ter alterado algo que afete uma regra
    End If
End Sub

Public Sub CanalVenda_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCanalVenda As New ClassCanalVenda
Dim iCodigo As Integer

On Error GoTo Erro_CanalVenda_Validate

    'Verifica se foi preenchida a ComboBox CanalVenda
    If Len(Trim(CanalVenda.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox CanalVenda
    If CanalVenda.Text = CanalVenda.List(CanalVenda.ListIndex) Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(CanalVenda, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 46836

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCanalVenda.iCodigo = iCodigo

        'Tenta ler CanalVenda com esse código no BD
        lErro = CF("CanalVenda_Le", objCanalVenda)
        If lErro <> SUCESSO And lErro <> 23597 Then gError 46837

        If lErro = 23597 Then gError 46838 'Não encontrou CanalVenda no BD

        'Encontrou CanalVenda no BD, coloca no Text da Combo
        CanalVenda.Text = CStr(objCanalVenda.iCodigo) & SEPARADOR & objCanalVenda.sNome

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 46838

    Exit Sub

Erro_CanalVenda_Validate:

    Cancel = True

    Select Case gErr

        Case 46836, 46837

        Case 46838  'Não encontrou CanalVenda no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CANALVENDA", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CanalDeVenda", objCanalVenda)
            Else
                'Segura o foco

            End If

        Case 46838
            Call Rotina_Erro(vbOKOnly, "ERRO_CANALVENDA_NAO_ENCONTRADO", gErr, CanalVenda.Text)

        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157341)

    End Select

    Exit Sub

End Sub

Public Sub NumPedidoTerc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CobrancaAutomatica_Click()

    iAlterado = REGISTRO_ALTERADO
    
    'Se cobrança automática estiver selecionado e a condição de pagamento preenchida
    If CobrancaAutomatica.Value = vbChecked And Len(Trim(CondicaoPagamento.Text)) > 0 Then
        'Recalcula a cobrança
        Call Cobranca_Automatica
    End If

    Exit Sub
    
End Sub

Public Sub ComissaoAutomatica_Click()

Dim lErro As Long

On Error GoTo Erro_ComissaoAutomatica_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Se a comissão automática estiver selecionada
    If ComissaoAutomatica.Value = vbChecked Then
        iComissoesAlterada = REGISTRO_ALTERADO
        'Recalcula as comissoes
        lErro = objTabComissoes.Comissoes_Calcula()
        If lErro <> SUCESSO Then gError 46839
        
    End If
    
    Exit Sub

Erro_ComissaoAutomatica_Click:

    Select Case gErr
    
        Case 46839
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157342)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoVendedores_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoVendedores_Click
    
    lErro = objTabComissoes.BotaoVendedores_Click()
    If lErro <> SUCESSO Then gError 43700
    
    Exit Sub

Erro_BotaoVendedores_Click:

    Select Case gErr
    
        Case 43700
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157343)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim lErro As Long

On Error GoTo Erro_objEventoVendedor_evSelecao
    
    Set objVendedor = obj1
    
    If GridComissoes.Row > 0 Then
    
        'Alterado por Tulio 10/04
    
        'se empresa nao utiliza regras p/ o calc de comissoes
        If Not gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
        
            'Preenche a linha de Vendedor com dados default
            lErro = objTabComissoes.VendedorLinha_Preenche(objVendedor)
            If lErro <> SUCESSO Then gError 51582
        
        'senao
         Else
            'verificar se o vendedor esta no grid e tratar o campo 'DiretoIndireto'
            lErro = objTabComissoes.VendedorLinha_Preenche1(objVendedor)
            If lErro <> SUCESSO Then gError 101537
    
        End If
                
    End If

    Me.Show
    
    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr

        Case 51582, 101537 'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157344)

    End Select

    Exit Sub

End Sub

Function Carrega_Dados_NFiscal(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long, objTituloRec As New ClassTituloReceber

On Error GoTo Erro_Carrega_Dados_NFiscal
    
    'Tenta ler a nota Fiscal passada por parametro
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 46842
    
    'Lê as mensagens
    lErro = CF("NFiscal_Le_Msg", objNFiscal)
    If lErro <> SUCESSO Then gError 46842
    
    'Lê os itens da NF
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 46842
    
    'Lê a parte de Tributação
    lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
    If lErro <> SUCESSO And lErro <> 22867 Then gError 46841
    
    'Lê a parte de tributação dos itens da Nota Fiscal
    lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
    If lErro <> SUCESSO Then gError 46843
    
    'Lê as Comissões da Nota Fiscal
    lErro = CF("ComissoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 46844

    'Lê as Alocações dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 46845
    
    'Lê as parcelas da Nota Fiscal Fatura
    lErro = CF("ParcelasRecNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 46846
        
    'carrega dados do titulo a receber associado à nota fiscal
    objTituloRec.lNumIntDoc = objNFiscal.lNumIntDocCPR
    lErro = CF("TituloReceber_Le", objTituloRec, 1)
    If lErro <> SUCESSO And lErro <> 26061 Then gError 56574
    If lErro <> SUCESSO Then
        'Busca nos baixados
        lErro = CF("TituloReceberBaixado_Le", objTituloRec, 1)
        If lErro <> SUCESSO And lErro <> 56570 Then gError 56575
        If lErro <> SUCESSO Then gError 56576
    
    End If
    'Transfere para a nota a condição de pagamento do título
    objNFiscal.iCondicaoPagto = objTituloRec.iCondicaoPagto
    
    Carrega_Dados_NFiscal = SUCESSO
    
    Exit Function
    
Erro_Carrega_Dados_NFiscal:

    Carrega_Dados_NFiscal = gErr
    
    Select Case gErr
    
        Case 46841, 46842, 46843, 46844, 46845, 46846, 56574, 56575
        
        Case 56576
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_REC_INEXISTENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157345)
            
    End Select
    
    Exit Function

End Function

Public Function Traz_NFiscal_Tela(objNFiscal As ClassNFiscal, Optional lNumIntPedido As Long = 0) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim bAlterouCT As Boolean
Dim colEmbProdItem As New Collection

On Error GoTo Erro_Traz_NFiscal_Tela
    
'    'Se for versao LIGHT testa chaveamento
'    If giTipoVersao = VERSAO_LIGHT And lNumIntPedido = 0 Then
'
'        lErro = objUserControl.CT_Chavea_Trazendo(objNFiscal, bAlterouCT)
'        If lErro <> SUCESSO Then gError 25826
'
'        If bAlterouCT Then
'
'            'Chama o Traz_NFiscal_Tela de objCT chaveado
'            lErro = objUserControl.objCT.Traz_NFiscal_Tela(objNFiscal)
'            If lErro <> SUCESSO Then gError 25827
'
'            iAlterado = 0
'            iComissoesAlterada = 0
'            gbCarregandoTela = False
'            Traz_NFiscal_Tela = SUCESSO
'
'            Exit Function
'
'        End If
'
'    End If
    
   gbCarregandoTela = True
    
    'Lê os demais dados da nota fiscal passada se não veio de Pedido
    If lNumIntPedido = 0 Then
        lErro = Carrega_Dados_NFiscal(objNFiscal)
        If lErro <> SUCESSO Then gError 25873
    End If
    
    'Limpa a tela de Nota Fiscal
    lErro = Limpa_Tela_NFiscalFaturaPedido()
    If lErro <> SUCESSO Then gError 500146

    Set gobjNFiscal = objNFiscal
    
    lErro = Traz_NFiscal_Tela2(objNFiscal, lNumIntPedido)
    If lErro <> SUCESSO Then gError 25950
    
    If lNumIntPedido = 0 Then
        lErro = CF("EmbalagensProdItem_Le", objNFiscal.lNumIntDoc, DOCORIGEM_NF, objNFiscal.colItens, colEmbProdItem)
        If lErro <> SUCESSO Then gError 177064
    
        'Preenche o Grid com as Embalagens dos itens do Pedido de Venda
        lErro = gobjEmbalagem.Preenche_GridEmbalagens(colEmbProdItem)
        If lErro <> SUCESSO Then gError 177064
    End If
       
    iAlterado = 0
    iComissoesAlterada = 0
    
    If lNumIntPedido = 0 Then
        MsgAutomatica.Value = vbUnchecked
        gbCarregandoTela = False
    Else
        MsgAutomatica.Value = vbChecked
    End If
    
    Set gobjInfoAdic = objNFiscal.objInfoAdic
    
    Traz_NFiscal_Tela = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela:

    gbCarregandoTela = False

    Traz_NFiscal_Tela = gErr

    Select Case gErr

        Case 25826, 25827, 25873, 25950, 500146

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157346)

    End Select

    Exit Function

End Function

Private Function Traz_NFiscal_Tela2(objNFiscal As ClassNFiscal, lNumIntPedido As Long) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objCliente As New ClassCliente
Dim colCodigoNome As New AdmColCodigoNome
Dim dValorProdutos As Double
Dim bGerarDados As Boolean
Dim dtDataContabilidade As Date
Dim iLote As Integer
Dim bCancel As Boolean, iTipoNFiscal As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim dPercDesc As Double
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Traz_NFiscal_Tela2

    iTipoNFiscal = objNFiscal.iTipoNFiscal
    
    objCliente.lCodigo = objNFiscal.lCliente
    'Lê o cliente
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 19062 Then gError 46847
    If lErro <> SUCESSO Then gError 46853
    'Lê as filiais do cliente
    lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 46848
    'CArrega a combo de filial entrega
    Call CF("Filial_Preenche", FilialEntrega, colCodigoNome)
    
    'coloca o pedido e filialpedido na tela
    Pedido.Text = objNFiscal.lNumPedidoVenda
    FilialPedido.Text = objNFiscal.iFilialPedido
    Call FilialPedido_Validate(bSGECancelDummy)

    glCodPedido = objNFiscal.lNumPedidoVenda
    giFilialPedido = objNFiscal.iFilialPedido
    
    'Preenche o Status da Nota Fiscal
    If objNFiscal.iStatus = STATUS_LANCADO Then
        Status.Caption = STRING_STATUS_LANCADO
    ElseIf objNFiscal.iStatus = STATUS_BAIXADO Then
        Status.Caption = STRING_STATUS_BAIXADO
    ElseIf objNFiscal.iStatus = STATUS_CANCELADO Then
        Status.Caption = STRING_STATUS_CANCELADO
    End If
    
    'Preenche o Cliente
    Cliente.Caption = objCliente.sNomeReduzido
    gdDesconto = objCliente.dDesconto
    
    'Preenche a Filial do Cliente
    For iIndice = 0 To FilialEntrega.ListCount - 1
        If FilialEntrega.ItemData(iIndice) = objNFiscal.iFilialCli Then
            Filial.Caption = FilialEntrega.List(iIndice)
            Exit For
        End If
    Next

    'Coloca a FilialEntrega na tela
    If objNFiscal.iFilialEntrega > 0 Then
        FilialEntrega.Text = objNFiscal.iFilialEntrega
        Call FilialEntrega_Validate(bSGECancelDummy)
    End If

    'Coloca o Tipo na tela
    If iTipoNFiscal = 0 Then
        iTipoNFiscal = giTipoNFiscalPadrao
'        objFilialEmpresa.iCodFilial = giFilialEmpresa
'        objFilialEmpresa.lCodEmpresa = glEmpresa
'        'Lê a FilialEmpresa
'        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
'        If lErro <> SUCESSO And lErro <> 27378 Then Error 46850
'
'        If objFilialEmpresa.iTipoTribAtividade = 3 Then
'            iTipoNFiscal = DOCINFO_NFISFSPV
'        Else
'            iTipoNFiscal = DOCINFO_NFISFVPV
'        End If
    End If
    TipoNFiscal.Text = iTipoNFiscal
    Call TipoNFiscal_Validate(bSGECancelDummy)
    
    'Coloca TabelaPreco na Tela
    If objNFiscal.iTabelaPreco > 0 Then
'        For iIndice = 0 To TabelaPreco.ListCount - 1
'            If objNFiscal.iTabelaPreco = TabelaPreco.ItemData(iIndice) Then
'                TabelaPreco.Text = TabelaPreco.List(iIndice)
'                Exit For
'            End If
'        Next
        TabelaPreco.Text = objNFiscal.iTabelaPreco
        Call TabelaPreco_Validate(bSGECancelDummy)

    End If

    'Preenche a Condicao de Pagto
    If objNFiscal.iCondicaoPagto > 0 Then
        CondicaoPagamento.Text = objNFiscal.iCondicaoPagto
        Call CondicaoPagamento_Validate(bCancel)
    End If
        
    'Preenche os demais dados da tela
    If objNFiscal.lNumNotaFiscal > 0 Then NFiscal.Caption = objNFiscal.lNumNotaFiscal
    
    Serie.Text = Desconverte_Serie_Eletronica(objNFiscal.sSerie)
    If ISSerieEletronica(objNFiscal.sSerie) Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    Call Serie_Validate(bSGECancelDummy)
    
    'Canal Venda
    If objNFiscal.iCanal > 0 Then
        CanalVenda.Text = objNFiscal.iCanal
        Call CanalVenda_Validate(bSGECancelDummy)
    End If
    
    For iIndice = 0 To TipoFrete.ListCount - 1
        If TipoFrete.ItemData(iIndice) = objNFiscal.iFreteRespons Then
            TipoFrete.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Call DateParaMasked(DataReferencia, objNFiscal.dtDataReferencia)
    giDataReferenciaAlterada = 0
    Call DateParaMasked(DataEmissao, objNFiscal.dtDataEmissao)
    giDataEmissaoAlterada = 0
    If objNFiscal.iSemDataSaida = 0 Then
        Call DateParaMasked(DataSaida, objNFiscal.dtDataSaida)
    Else
        Call DateParaMasked(DataSaida, DATA_NULA)
    End If
    
'horasaida
    HoraSaida.PromptInclude = False
    'este teste está correto
    If objNFiscal.dtDataSaida <> DATA_NULA Then HoraSaida.Text = Format(objNFiscal.dtHoraSaida, "hh:mm:ss")
    HoraSaida.PromptInclude = True
    
    'Coloca os Valores na Tela
    If objNFiscal.dValorDesconto > 0 Then ValorDesconto.Text = Format(objNFiscal.dValorDesconto, "Standard")
    If objNFiscal.dValorOutrasDespesas > 0 Then ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Standard")
    If objNFiscal.dValorFrete > 0 Then ValorFrete.Text = Format(objNFiscal.dValorFrete, "Standard")
    If objNFiscal.dValorSeguro > 0 Then ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Standard")
    If objNFiscal.dValorTotal > 0 Then ValorTotal.Caption = Format(objNFiscal.dValorTotal, "Standard")
    
    'Transportadora
    If objNFiscal.iCodTransportadora <> 0 Then
        Transportadora.Text = objNFiscal.iCodTransportadora
        Call Transportadora_Validate(bSGECancelDummy)
    End If

    If objNFiscal.iCodTranspRedesp <> 0 Then
        TranspRedespacho.Text = CStr(objNFiscal.iCodTranspRedesp)
        Call TranspRedespacho_Validate(bCancel)
    End If

    RedespachoCli.Value = objNFiscal.iDetPagFrete
    
    Placa.Text = objNFiscal.sPlaca
    PlacaUF.Text = objNFiscal.sPlacaUF
    
    'Alterado por Luiz Nogueira em 21/08/03
    'Traz a espécie dos volumes do pedido
    If objNFiscal.lVolumeEspecie > 0 Then
        VolumeEspecie.Text = objNFiscal.lVolumeEspecie
        Call VolumeEspecie_Validate(bSGECancelDummy)
    Else
        VolumeEspecie.Text = ""
    End If
    
    'Alterado por Luiz Nogueira em 21/08/03
    'Traz a marca dos volumes do pedido
    If objNFiscal.lVolumeMarca > 0 Then
        VolumeMarca.Text = objNFiscal.lVolumeMarca
        Call VolumeMarca_Validate(bSGECancelDummy)
    Else
        VolumeMarca.Text = ""
    End If
    
    VolumeNumero = objNFiscal.sVolumeNumero

    If objNFiscal.lVolumeQuant > 0 Then VolumeQuant.Text = objNFiscal.lVolumeQuant
    NumPedidoTerc.Text = objNFiscal.sNumPedidoTerc

    Mensagem.Text = objNFiscal.sMensagemNota
    MensagemCorpo.Text = objNFiscal.sMensagemCorpoNota

    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = Format(objNFiscal.dPesoLiq, PesoLiquido.Format) 'Alterado por Wagner
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = Format(objNFiscal.dPesoBruto, PesoBruto.Format) 'Alterado por Wagner
    If objNFiscal.dVolumeTotal > 0 Then Cubagem.Text = Format(objNFiscal.dVolumeTotal, "Fixed")
    
    lErro = CF2(Me, "NFFP_Traz_NFiscal_Tela2", objNFiscal, lNumIntPedido)
    If lErro <> SUCESSO Then gError 46850
    
    'Preenche o Grid com os itens da Nota Fiscal
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 46850

    If objNFiscal.iTipoNFiscal <> 0 Then

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        
        'Le o Tipo para saber se tem Movimento de Estoque
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO Then gError 83392
        
        If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
            'Carrega ItensNF com Rastreamentos
            lErro = gobjRastreamento.Carrega_RastroItensNF(objNFiscal)
            If lErro <> SUCESSO Then gError 83393
        End If

    End If

    'Carrega o Tab Comissões
    lErro = objTabComissoes.Carrega_Tab_Comissoes(objNFiscal)
    If lErro <> SUCESSO Then gError 46851

    If lNumIntPedido > 0 Then bGerarDados = True

    'Preenche o Grid com as Alocações dos itens da Nota Fiscal
    lErro = CF2(Me, "NFFP_Preenche_GridAlocacoes", objNFiscal, bGerarDados)
    If lErro <> SUCESSO Then gError 46884

    'Se a nota fiscal já existe
    If objNFiscal.lNumIntDoc <> 0 Then
    
        'Traz os dados contábeis para a tela (contabilidade)
        lErro = objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
        If lErro <> SUCESSO And lErro <> 36326 Then gError 46852
        
    End If

    dtDataContabilidade = StrParaDate(CTBDataContabil.Text)
    iLote = StrParaInt(CTBLote.Text)
    
    'Se a data contabil está ultrapassada
    If dtDataContabilidade <> DATA_NULA And dtDataContabilidade < gdtDataAtual Then
        'Atualiza a data da contabilidade
        CTBLote.Text = ""
        CTBDataContabil.PromptInclude = False
        CTBDataContabil.Text = Format(gdtDataAtual, "dd/mm/yy")
        CTBDataContabil.PromptInclude = True
        Call CTBDataContabil_Validate(bCancel)
        If iLote > 0 Then CTBLote.Text = iLote
    End If
        
    Call ValorTotal_Calcula
       
    'Preenche o grid de parcelas
    NaoImpCobranca.Value = objNFiscal.iNaoImpCobranca
    lErro = Preenche_Grid_Parcelas(objNFiscal)
    If lErro <> AD_SQL_SUCESSO Then gError 46883

    gobjTribTab.sNatOpInterna = objNFiscal.sNaturezaOp

    lErro = CF2(Me, "NFCF_Cliente_Traz_NF_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 131958

    If objNFiscal.iRecibo = MARCADO Then
        Recibo.Value = vbChecked
    Else
        Recibo.Value = vbUnchecked
    End If
    
    If objNFiscal.lNumNFe = 0 Then
        NumNFe.Caption = objNFiscal.sStatusNFeFed
    Else
        NumNFe.Caption = objNFiscal.lNumNFe
    End If
    
    CodVerificacao.Caption = objNFiscal.sCodVerificacaoNFe
    
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
    
    ValorDescontoTit.Text = Format(objNFiscal.dValorDescontoTit, "Standard")
    Call ValorDescontoTit_Validate(bSGECancelDummy)
    
    If objNFiscal.lFornEntTerc <> 0 Then
        FornEntTerc.Text = objNFiscal.lFornEntTerc
        Call FornEntTerc_Validate(bSGECancelDummy)
        FilialFornEntTerc.Text = objNFiscal.iFilialFornEntTerc
        Call FilialFornEntTerc_Validate(bSGECancelDummy)
    End If
    
    iAlterado = 0
    iComissoesAlterada = 0
    iPercAcresFinAlterado = 0
    
    Traz_NFiscal_Tela2 = SUCESSO
    
    Exit Function

Erro_Traz_NFiscal_Tela2:

    Traz_NFiscal_Tela2 = gErr

    Select Case gErr

        Case 46847, 46848, 46849, 46850, 46851, 46852, 46883, 46884, 51583, 83392, 83393
        
        Case 131958
        
        Case 46853
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157347)

    End Select

    Exit Function

End Function

Private Function Trata_TabelaPreco() As Long

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_TabelaPreco_Click

    For iLinha = 1 To objGridItens.iLinhasExistentes

        lErro = Trata_TabelaPreco_Item(iLinha)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Next

    'Calcula o VAlor Total da Nota
    Call ValorTotal_Calcula

    Trata_TabelaPreco = SUCESSO

    Exit Function

Erro_TabelaPreco_Click:

    Trata_TabelaPreco = gErr

    Select Case gErr

        Case 46660, 46661, 46662, 46663, 56860

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157457)

    End Select

    Exit Function

End Function

Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto) As Long
'Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas

Dim lErro As Long
Dim dValorPagar As Double
Dim iIndice As Integer
Dim dPercAcrescFin As Double
Dim iTamanho As Integer
Dim objParcRec As ClassParcelaReceber
Dim dValorPagarPrimParc As Double
Dim dValorParcST As Double, dtDataParcST As Date
Dim colParcAux As New Collection, bAdicionouST As Boolean
Dim objParcAux1 As ClassParcelaReceber, objParcAux2 As ClassParcelaReceber

On Error GoTo Erro_GridParcelas_Preenche

    Set gobjNFiscal.ColParcelaReceber = New ColParcelaReceber

    'Limpa o GridParcelas
    Call Grid_Limpa(objGridParcelas)
       
    If gobjFAT.iParcRecCobraSTNovaParc = MARCADO Then
        dValorParcST = StrParaDbl(ICMSSubstValor1.Caption)
        If StrParaDate(DataReferencia.Text) = DATA_NULA Then
            dtDataParcST = DATA_NULA
        Else
            dtDataParcST = DateAdd("d", gobjFAT.iParcSTNumDias, StrParaDate(DataReferencia.Text))
        End If
    End If
    
    'Número de Parcelas
    objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas
    
    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then gError 42256
    
    'Se Valor a Pagar for positivo
    If dValorPagar > 0 Then
    
        dValorPagarPrimParc = 0
        
        If objCondicaoPagto.iNumeroParcelas <> 1 Then
        
            lErro = gobjTribTab.ValorPagarPrimParc_Calcula(dValorPagarPrimParc)
            If lErro <> SUCESSO Then gError 42256
    
        End If
    
        If gobjFAT.iParcRecCobraSTNovaParc = MARCADO And dValorParcST > 0 Then
            objCondicaoPagto.dValorTotal = Arredonda_Moeda(dValorPagar - dValorParcST)
        Else
            If dValorPagarPrimParc <> 0 Then
                objCondicaoPagto.dValorTotal = Arredonda_Moeda(dValorPagar - dValorPagarPrimParc)
            Else
                objCondicaoPagto.dValorTotal = dValorPagar
            End If
        End If
               
        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, False, True, True)
        If lErro <> SUCESSO Then gError 42256
    
        If objGridParcelas.iLinhasExistentes <> objCondicaoPagto.colParcelas.Count Then objGridParcelas.iLinhasExistentes = objCondicaoPagto.colParcelas.Count
    
        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            
            If iIndice = 1 And dValorPagarPrimParc <> 0 Then objCondicaoPagto.colParcelas(iIndice).dValor = objCondicaoPagto.colParcelas(iIndice).dValor + dValorPagarPrimParc
            
            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        
        Next
    
    End If

    'Se Data Emissão estiver preenchida
    If Len(Trim(DataReferencia.ClipText)) > 0 Then
    
        objCondicaoPagto.dtDataRef = StrParaDate(DataReferencia.Text)
        objCondicaoPagto.dtDataEntrega = StrParaDate(GridItens.TextMatrix(1, iGrid_DataEntrega_Col))
        If objCondicaoPagto.dtDataEntrega = DATA_NULA Then objCondicaoPagto.dtDataEntrega = StrParaDate(DataSaida.Text)
        If objCondicaoPagto.dtDataEntrega = DATA_NULA Then objCondicaoPagto.dtDataEntrega = StrParaDate(DataEmissao.Text)
        'objCondicaoPagto.dtDataEmissao = StrParaDate(DataEmissao.Text)
        
        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True, True, True)
        If lErro <> SUCESSO Then gError 42257
        
        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.colParcelas.Count  'objCondicaoPagto.iNumeroParcelas
        
            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dtVencimento, "dd/mm/yyyy")
            
'            Set objParcRec = New ClassParcelaReceber
'            gobjNFiscal.ColParcelaReceber.AddObj objParcRec
            
        Next
                        
    End If
    
    If gobjFAT.iParcRecCobraSTNovaParc = MARCADO And dValorParcST > 0 Then
        'Joga as parcelas calculadas para coleção e posiciona a nova parcela no lugar certo
        bAdicionouST = False
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            Set objParcAux1 = New ClassParcelaReceber
            objParcAux1.dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
            objParcAux1.dValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
            If objParcAux1.dtDataVencimento > dtDataParcST And Not bAdicionouST Then
                Set objParcAux2 = New ClassParcelaReceber
                objParcAux2.dtDataVencimento = dtDataParcST
                objParcAux2.dValor = dValorParcST
                colParcAux.Add objParcAux2
                bAdicionouST = True
            End If
            colParcAux.Add objParcAux1
        Next
        If Not bAdicionouST Then
            Set objParcAux2 = New ClassParcelaReceber
            objParcAux2.dtDataVencimento = dtDataParcST
            objParcAux2.dValor = dValorParcST
            colParcAux.Add objParcAux2
            bAdicionouST = True
        End If
        iIndice = 0
        For Each objParcAux1 In colParcAux
            iIndice = iIndice + 1
            If objParcAux1.dtDataVencimento <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objParcAux1.dtDataVencimento, "dd/mm/yyyy")
            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objParcAux1.dValor, "Standard")
        Next
        objGridParcelas.iLinhasExistentes = iIndice
    End If
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        lErro = Preenche_DescontoPadrao(iIndice)
        If lErro <> SUCESSO Then gError 51117
        
        If Len(Trim(DataReferencia.ClipText)) > 0 Then
            Set objParcRec = New ClassParcelaReceber
            gobjNFiscal.ColParcelaReceber.AddObj objParcRec
        End If
    Next
    GridParcelas_Preenche = SUCESSO
    
    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = gErr
    
    Select Case gErr
    
        Case 42256, 42257, 51117
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157349)
        
    End Select
    
    Exit Function

End Function

Function ValorTotal_Calcula() As Long

Dim dValorTotal As Double
Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    lErro = gobjTribTab.ValorTotal_Calcula(dValorTotal)
    If lErro <> SUCESSO Then gError 46860
    
    ValorTotal.Caption = Format(dValorTotal, "Standard")
    
    Call ValorDescontoTit_Validate(bSGECancelDummy)
    
    Call ValorDescontoItens_Calcula

    If Not gbCarregandoTela Then
        'Gera a cobranca em cima do novo valor total
        lErro = Cobranca_Automatica()
        If lErro <> SUCESSO Then gError 56909
        
        'Se a empresa nãu utiliza as regras para o cálculo de comissões
        If Not gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
        
            'Significa que é necessário recalcular as comissões, utilizando o cálculo padrão
            lErro = objTabComissoes.Comissoes_Calcula_Padrao()
            If lErro <> SUCESSO Then gError 46861
        
        Else
            iComissoesAlterada = 1 'preventivamente, pois pode ter alterado algo que afete uma regra
        End If
    
    End If
    
    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 46860, 46861, 56890
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157350)

    End Select

    Exit Function

End Function

Function Pesos_Recalcula() As Long
    Pesos_Recalcula = CF2(Me, "NFFP_Pesos_Recalcula")
End Function

Public Function NFFP_Pesos_Recalcula(ByVal objCT As Object) As Long

Dim lErro As Long
Dim dFator As Double
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dPesoBruto As Double
Dim dPesoLiquido As Double
Dim dQuantidade As Double

On Error GoTo Erro_NFFP_Pesos_Recalcula
    
    lErro = Calcula_VolumeTotal
    If lErro <> SUCESSO Then gError 96134
    
    dPesoBruto = 0
    dPesoLiquido = 0

'    'Para cada linha do Grid de Itens
'    For iIndice = 1 To objGridItens.iLinhasExistentes
'
'        'Se o produto e a UM estiverem preenchidos
'        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) > 0 Then
'
'            'Formata o produto
'            lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'            If lErro <> SUCESSO Then gError 500098
'
'            'Lê o produto
'            objProduto.sCodigo = sProdutoFormatado
'            lErro = CF("Produto_Le", objProduto)
'            If lErro <> SUCESSO And lErro <> 28030 Then gError 500099
'
'            'Se o produto não está cadastrado, erro
'            If lErro = 28030 Then gError 500101
'
'            'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
'            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), objProduto.sSiglaUMVenda, dFator)
'            If lErro <> SUCESSO Then gError 500100
'
'            dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
'
'            'Acumula o peso Bruto e o peso líquido
'            dPesoBruto = dPesoBruto + objProduto.dPesoBruto * dFator * dQuantidade
'            dPesoLiquido = dPesoLiquido + objProduto.dPesoLiq * dFator * dQuantidade
'
'        End If
'
'    Next
'
'    'Coloca os Pesos na tela
'    PesoBruto.Text = Format(dPesoBruto, PesoBruto.Format)
'    PesoLiquido.Text = Format(dPesoLiquido, PesoLiquido.Format)

    dPesoBruto = 0
    dPesoLiquido = 0
    
    'Para cada linha do Grid de Itens
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        'Se o produto e a UM estiverem preenchidos
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) > 0 Then
                       
            'Preenche o Grid de Embalagens
            lErro = gobjEmbalagem.Calcula_Item_GridEmb(StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)), iIndice)
            If lErro <> SUCESSO Then gError 96125
                   
        End If
        
    Next
    
    lErro = gobjEmbalagem.Preenche_Pesos()
    If lErro <> SUCESSO Then gError 96134
    
    NFFP_Pesos_Recalcula = SUCESSO
    
    Exit Function
    
Erro_NFFP_Pesos_Recalcula:

    NFFP_Pesos_Recalcula = gErr
    
    Select Case gErr
    
        Case 500098, 500099, 500100, 96134
        
        Case 500101
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157352)
    
    End Select
    
    Exit Function
    
End Function

Private Function Limpa_Tela_NFiscalFaturaPedido() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_NFiscalFaturaPedido

    'Indica que está limpando a tela
    gbLimpandoTela = True
    
    'Limpa os campos gerais da tela
    Call Limpa_Tela(Me)
     
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    Set gobjAnotacao = Nothing
    Set gobjInfoAdic = Nothing
    
    Set gobjNFiscal = New ClassNFiscal
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next
    
    'Limpa os grids
    Call Grid_Limpa(objGridItens)
    objTabComissoes.Limpa_Grid_Comissoes
    Call Grid_Limpa(objGridAlocacoes)
    Call Grid_Limpa(objGridParcelas)
    
'    If giTipoVersao = VERSAO_FULL Then
        Serie.Text = ""
        Call Serie_Validate(bSGECancelDummy)
        TabelaPreco.Text = ""
'    End If
    
    Status.Caption = ""

    VolumeMarca.Text = ""
    VolumeEspecie.Text = ""
    
    'Recarrega a combo VolumeEspecie e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)

    'Recarrega a combo VolumeMarca e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)

    'Limpa os campos não limpos na rotina geral
    Cliente.Caption = ""
    Filial.Caption = ""
    FilialPedido.Text = giFilialEmpresa
    Call FilialPedido_Validate(bSGECancelDummy)
    ValorProdutos.Caption = ""
    ValorProdutos2.Caption = ""
    ValorOriginalTit.Caption = ""
    ValorTit.Caption = ""
    ValorTotal.Caption = ""
    FilialEntrega.Clear
    Transportadora.Text = ""
    TranspRedespacho.Text = ""
    RedespachoCli.Value = False
    PlacaUF.Text = ""
    CanalVenda.Text = ""
    TotalPercentualComissao.Caption = ""
    TotalValorComissao.Caption = ""
    TotalValorBase.Caption = ""
    CondicaoPagamento.Text = ""
    NFiscal.Caption = ""
    TipoFrete.ListIndex = 0
    EletronicaFed.Value = vbUnchecked
    FilialFornEntTerc.Clear
    
    DataReferencia.PromptInclude = False
    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataReferencia.PromptInclude = True
    
    Call CF2(Me, "NFCF_Default_DataSaida")
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    Call DataEmissao_Validate(bSGECancelDummy)
    
    'Reinicializa od flags da tela
    giFaturaIntegral = 0
    glCodPedido = 0
    giFilialPedido = giFilialEmpresa
    
    TipoNFiscal.Text = ""
    Frame1(TAB_Almoxarifado).Enabled = True
    
    iValorDescontoAlterado = 0
    iPercAcresFinAlterado = 0
    iValorFreteAlterado = 0
    iValorSeguroAlterado = 0
    iValorDespesasAlterado = 0
    iValorDescontoAlterado = 0
    gdDesconto = 0
    giDataReferenciaAlterada = 0
    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0
    
    NaoImpCobranca.Value = vbUnchecked
    
    Recibo.Value = vbUnchecked
    
    NumNFe.Caption = ""
    CodVerificacao.Caption = ""
    
    'Limpa o tab de tributacao
    Call gobjTribTab.Limpa_Tela

    'Limpa o tab de contabilitade
    Call objContabil.Contabil_Limpa_Contabilidade
    
    Call Seleciona_Tipo_Venda

    'Seleciona série padrão
    lErro = Seleciona_Serie_Padrao()
    If lErro <> SUCESSO Then gError 500143
    
    Call CF2(Me, "NFCF_Limpa_Tela")
    Call CF2(Me, "NFFP_Limpa_Tela")
    
    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
    Set gcolItemPedido = New Collection
    Set gcolItemQuant = New Collection
    Set gcolItemUM = New Collection 'Inserido por Wagner
    
    'embalagem
    Call gobjEmbalagem.Limpa_Tela_Embalagem
    
    MsgAutomatica.Value = vbChecked
    
    iAlterado = 0
    iComissoesAlterada = 0

    'Indica que não está mais limpando a tela
    gbLimpandoTela = False
    
    Exit Function

Erro_Limpa_Tela_NFiscalFaturaPedido:
    
    Limpa_Tela_NFiscalFaturaPedido = gErr
    
    Select Case gErr
    
        Case 500143
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157353)
    
    End Select
    
    'Indica que não está mais limpando a tela
    gbLimpandoTela = False
    
    Exit Function
    
End Function

Public Function Preenche_GridItens(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com os itens da Nota Fiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim sProdutoEnxuto As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim dValorProdutos As Double
'Dim dValorDesconto  As Double
'Dim dValorTotal As Double
Dim objProduto As New ClassProduto
Dim iItem As Integer

On Error GoTo Erro_Preenche_GridItens

    iIndice = 0

    'Para cada item da Coleção
    For Each objItemNF In objNFiscal.ColItensNF

        iIndice = iIndice + 1
        objItemNF.iItem = iIndice

        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 46862

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177926

        'Preenche o Grid
'****** IF INCLUÍDO PARA TRATAMENTO DE GRADE ***************
        If objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade.Count > 0 And objProduto.iKitVendaComp <> MARCADO Then GridItens.TextMatrix(iIndice, 0) = "# " & GridItens.TextMatrix(iIndice, 0)
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItemNF.sDescricaoItem
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemNF.sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemNF.dQuantidade)
        
        If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        
            If objItemNF.iTabelaPreco = 0 Then
            
                GridItens.TextMatrix(iIndice, iGrid_TabPrecoItem_Col) = ""
            
            Else
            
                For iItem = 0 To TabPrecoItemNF.ListCount - 1
                    
                    If objItemNF.iTabelaPreco = TabPrecoItemNF.ItemData(iItem) Then
                        
                        GridItens.TextMatrix(iIndice, iGrid_TabPrecoItem_Col) = TabPrecoItemNF.List(iItem)
                        Exit For
                    
                    End If
                
                Next
        
            End If
        
        End If
        
        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(objItemNF.dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(objItemNF.dPercDesc, "Percent")
        
        lErro = CF2(Me, "NFFP_Preenche_GridItens", objItemNF, iIndice)
        If lErro <> SUCESSO Then gError 177048

        If objItemNF.dtDataEntrega <> DATA_NULA Then
            GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col) = Format(objItemNF.dtDataEntrega, "dd/mm/yyyy")
        End If
        
        lErro = CF2(Me, "NFFP_Preenche_GridItem", objNFiscal, objItemNF, iIndice)
        If lErro <> SUCESSO Then gError 124190
        
        lErro = CF2(Me, "NF_Preenche_GridItens", objNFiscal, objItemNF, objProduto, iIndice)
        If lErro <> SUCESSO Then gError 124190
        
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col) = Format(objItemNF.dValorTotal + objItemNF.dValorDesconto, "Standard")
        
        If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        
            If objItemNF.dComissao <> -1 Then
                GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col) = Format(objItemNF.dComissao, "Percent")
            Else
                GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col) = ""
            End If
        
        End If
        
        If iIndice > gcolItemPedido.Count Then gcolItemPedido.Add objItemNF.lNumIntItemPedVenda
        If iIndice > gcolItemQuant.Count Then gcolItemQuant.Add objItemNF.dQuantidade
        If iIndice > gcolItemUM.Count Then gcolItemUM.Add objItemNF.sUnidadeMed
        
    Next
    
    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice

    Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)
    
    ValorProdutos.Caption = Format(dValorProdutos - objNFiscal.dValorDesconto, "Standard")

    lErro = CF2(Me, "NFFP_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132035
    
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objNFiscal.objProjetoInfo, objNFiscal.lNumIntDoc, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181676
    
    Call Grid_Refresh_Checkbox(objGridItens)
    
    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr

        Case 46862, 124190, 132035, 177048, 177926, 181676

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157354)

    End Select

    Exit Function

End Function

Function NFFP_Preenche_GridAlocacoes(ByVal objCT As Object, objNFiscal As ClassNFiscal, Optional bGeraAlocacao As Boolean = False) As Long
'Preenche o Grid com as Alocações da Nota Fiscal

Dim objItemAloc As ClassItemNFAlocacao
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iNumCasasDec As Integer
Dim dAcrescimo As Double
Dim iContador As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_NFFP_Preenche_GridAlocacoes

    'Limpa o grid de alocações
    Call Grid_Limpa(objGridAlocacoes)

    objGridAlocacoes.iLinhasExistentes = 0
    
    'Para cada item da NF
    For iIndice = 1 To objNFiscal.ColItensNF.Count
                
        If objNFiscal.ColItensNF.Item(iIndice).colItensRomaneioGrade.Count = 0 Then
        
            Call AlocacoesNF_Agrupa(objNFiscal.ColItensNF.Item(iIndice).colAlocacoes)
            
            iContador = 0
            'Para cada alocação do Item de NF
            For Each objItemAloc In objNFiscal.ColItensNF.Item(iIndice).colAlocacoes
    
                iContador = iContador + 1
                objGridAlocacoes.iLinhasExistentes = objGridAlocacoes.iLinhasExistentes + 1
                iIndice1 = objGridAlocacoes.iLinhasExistentes
                
                objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 46865
                If lErro <> SUCESSO Then gError 51536
                
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objNFiscal.ColItensNF(iIndice).sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 46866
                'Coloca os dados da alocação na tela
                GridAlocacao.TextMatrix(iIndice1, iGrid_Item_Col) = objNFiscal.ColItensNF(iIndice).iItem
                GridAlocacao.TextMatrix(iIndice1, iGrid_ProdutoAloc_Col) = GridItens.TextMatrix(objNFiscal.ColItensNF(iIndice).iItem, iGrid_Produto_Col)
                GridAlocacao.TextMatrix(iIndice1, iGrid_AlmoxAloc_Col) = objItemAloc.sAlmoxarifado
                GridAlocacao.TextMatrix(iIndice1, iGrid_UMAloc_Col) = objProduto.sSiglaUMEstoque
                GridAlocacao.TextMatrix(iIndice1, iGrid_QuantAloc_Col) = Formata_Estoque(objItemAloc.dQuantidade)
                dQuantidade = StrParaDbl(GridItens.TextMatrix(objNFiscal.ColItensNF(iIndice).iItem, iGrid_Quantidade_Col))
                dQuantidade = Arredonda_Estoque(dQuantidade * dFator)
                
                GridAlocacao.TextMatrix(iIndice1, iGrid_QuantVend_Col) = Formata_Estoque(dQuantidade)
    
            Next
            
            'Quando os dados que estão sendo carregados são originados por um pedido
            'e um item não possui localização tenta gerar automaticamente a alocação
            If iContador = 0 And bGeraAlocacao Then
                GridItens.Row = iIndice
                
                lErro = CF2(Me, "NFFP_Alocacao_Processa")
                If lErro <> SUCESSO And lErro <> 23796 And lErro <> 197701 And lErro <> 197703 Then gError 51602 '??? Jones 09/02/05 passou a tratar 23796
                
            End If
        Else
        
            'Faz a alocação de quem não tem
            For Each objItemRomaneio In objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade
                If objItemRomaneio.colLocalizacao.Count = 0 Then
                    lErro = Alocacao_Processa_ItemGrade(objItemRomaneio, objNFiscal.ColItensNF(iIndice), iIndice)
                    If lErro <> SUCESSO Then gError 46865
                End If
            Next
            
            Call Atualiza_Grid_Alocacao(objNFiscal.ColItensNF.Item(iIndice))
            
            objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 46865
            If lErro <> SUCESSO Then gError 51536

            If objProduto.iKitVendaComp <> MARCADO Then
                Call Atualiza_Grid_Itens(objNFiscal.ColItensNF.Item(iIndice))
            End If
        
        End If

    Next
    
    If gbCarregandoTela Then
    
        lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 46849
    
    End If
    
    NFFP_Preenche_GridAlocacoes = SUCESSO

    Exit Function

Erro_NFFP_Preenche_GridAlocacoes:

    NFFP_Preenche_GridAlocacoes = gErr
    
    Select Case gErr
    
        Case 46865, 46866, 51602
        
        Case 51536
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157355)
            
    End Select

End Function

Private Function PrecoUnitario_Calcula(sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long
    PrecoUnitario_Calcula = CF2(Me, "NFFP_PrecoUnitario_Calcula", sUM, objTabelaPrecoItem, dPrecoUnitario)
End Function

Public Function NFFP_PrecoUnitario_Calcula(ByVal objCT As Object, sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objUM As New ClassUnidadeDeMedida
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim dPercAcresFin As Double
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_NFFP_PrecoUnitario_Calcula

    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto
    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 46868
    If lErro = 28030 Then gError 46869 'Não encontrou

    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 46867
    
    'Converte o preço unitário para UM da tela
    dPrecoUnitario = objTabelaPrecoItem.dPreco / dFator
    
    'Aplica o acréscimo financeiro no preço unitário
    dPercAcresFin = StrParaDbl(PercAcrescFin.Text)
    dPrecoUnitario = dPrecoUnitario * (1 + (dPercAcresFin / 100))
    
    NFFP_PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_NFFP_PrecoUnitario_Calcula:

    NFFP_PrecoUnitario_Calcula = gErr

    Select Case gErr

        Case 46867, 46868

        Case 46869
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objTabelaPrecoItem.sCodProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157356)

    End Select

    Exit Function

End Function

Public Sub PrecoTotal_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrecoTotal_Calcula(iLinha As Integer, Optional iCalculaTotal = True)

Dim dPrecoTotal As Double
Dim dPrecoTotalReal As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim dValorTotal As Double
Dim iIndice As Integer
Dim dValorProdutos As Double
Dim dValorDescontoGlobal As Double, dValorDescontoItens As Double
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_PrecoTotal_Calcula

    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))
    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))

    lErro = CF2(Me, "NFFP_PrecoTotal_Calcula", iLinha, dQuantidade)
    If lErro <> SUCESSO Then gError 177050

    'Recolhe os valores Quantidade, Desconto, PerDesc e Valor Unitário da tela
    If dPrecoUnitario = 0 Or dQuantidade = 0 Then
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = ""
        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(0, "Standard")
        
    Else
        dPrecoTotal = Arredonda_Moeda(dPrecoUnitario * dQuantidade) 'Alterado por Wagner
        dDesconto = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col))
        dPercentDesc = PercentParaDbl(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col))

        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dPrecoUnitario, dPercentDesc, dDesconto, dPrecoTotalReal)

        'Coloca o Desconto calculado na tela
        If dDesconto > 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
        Else
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        End If

        'Coloca o valor Real em Valor Total
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(dPrecoTotalReal, "Standard")

        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dPrecoUnitario, 0, 0, dPrecoTotalReal)

        'Coloca o valor Real em Valor Total
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(dPrecoTotalReal, "Standard")

    End If

    lErro = CF2(Me, "NF_PrecoTotal_Calcula", iLinha, dQuantidade, dPrecoUnitario, dPercentDesc)
    If lErro <> SUCESSO Then gError 184287
    
    Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)

'    'Obtem o valor do Desconto global
'    If Len(Trim(ValorDesconto.Text)) > 0 And IsNumeric(ValorDesconto.Text) Then dValorDescontoGlobal = CDbl(ValorDesconto.Text)
'
'    dValorProdutos = Arredonda_Moeda(dValorProdutos - dValorDescontoGlobal) 'Alterado por Wagner
'
'    'Se com o desconto o valor Produtos ficou negativo
'    If dValorProdutos < 0 And dValorDescontoGlobal > 0 Then
'        'Avisa ao usuário
'        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", dValorDescontoGlobal, dValorProdutos)
'        'Limpa o valor de desconto
'        ValorDesconto.Text = ""
'        dValorDescontoGlobal = 0
'''*** incluidos p/tratamento de tributacao *******************************
'        Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, dValorDescontoGlobal)
''*** fim tributacao
'
'        'Faz a atualização dos valores da tributação
'        lErro = gobjTribTab.AtualizarTributacao()
'        If lErro <> SUCESSO Then gError 56889
'
'        Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)
'
'    End If

    'Coloca o valor dos produtos na tela
    ValorProdutos.Caption = Format(dValorProdutos, "Standard")

    Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotalB_Col, dValorProdutos)
    ValorProdutos2.Caption = Format(dValorProdutos, "Standard")

    Call SubTotal_Calcula(objGridItens, iGrid_Desconto_Col, dValorDescontoItens)
    ValorDescontoItens.Text = Format(dValorDescontoItens, "Standard")
    dValorDescontoItensAnt = dValorDescontoItens
    If dValorProdutos > 0 Then
        PercDescontoItens.Text = Format(dValorDescontoItens / dValorProdutos * 100, "FIXED")
        dPercDescontoItensAnt = Arredonda_Moeda(dValorDescontoItens / dValorProdutos, 4)
    Else
        PercDescontoItens.Text = Format(0, "FIXED")
        dPercDescontoItensAnt = 0
    End If
    
    lErro = CF2(Me, "NFFP_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132036
    
    If iCalculaTotal <> False Then Call ValorTotal_Calcula

    Exit Sub

Erro_PrecoTotal_Calcula:

    Select Case gErr
    
        Case 56889
        
        Case 132036, 177050, 184287
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157357)
            
    End Select
            
    Exit Sub

End Sub

Private Sub SubTotal_Calcula(objGridInt As AdmGrid, iGrid_Coluna As Integer, dSubTotal As Double)
'Faz a soma da Coluna passado no Grid passado e devolve em dValorTotal

Dim iIndice As Integer
Dim objItemNF As ClassItemNF

    dSubTotal = 0

    For iIndice = 1 To objGridInt.iLinhasExistentes
    
        Set objItemNF = gobjNFiscal.colItens.Item(iIndice)
        
        If objItemNF.objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
            'Acumula em dSubTotal
            dSubTotal = Arredonda_Moeda(dSubTotal + CDbl(IIf(Len(Trim(objGridInt.objGrid.TextMatrix(iIndice, iGrid_Coluna))) > 0, objGridInt.objGrid.TextMatrix(iIndice, iGrid_Coluna), 0))) 'Alterado por Wagner
        End If
    Next

    Exit Sub

End Sub

Private Sub ValorReal_Calcula(dQuantidade As Double, dValorUnitario As Double, dPercentDesc As Double, dDesconto As Double, dValorReal As Double)
'Calcula o Valor Real

Dim dValorTotal As Double
Dim dPercDesc1 As Double
Dim dPercDesc2 As Double

    dValorTotal = Arredonda_Moeda(dValorUnitario * dQuantidade) 'Alterado por Wagner

    'Se o Percentual Desconto estiver preenchido
    If dPercentDesc > 0 Then

        'Testa se o desconto está preenchido
        If dDesconto = 0 Then
            dPercDesc2 = 0
        Else
            'Calcula o Percentual em cima dos valores passados
            dPercDesc2 = dDesconto / dValorTotal
            dPercDesc2 = CDbl(Format(dPercDesc2, "0.0000"))
        End If
        'se os percentuais (passado e calculado) forem diferentes calcula-se o desconto
        If dPercentDesc <> dPercDesc2 Then dDesconto = Arredonda_Moeda(dPercentDesc * dValorTotal)

    End If

    dValorReal = Arredonda_Moeda(dValorTotal - dDesconto) 'Alterado por Wagner

End Sub

Public Sub BotaoPedido_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objBloqueioPV As ClassBloqueioPV

On Error GoTo Erro_BotaoPedido_Click

    'Verfifica se o Pedido e a FilialPedido foram informadas
    If Len(Trim(Pedido.Text)) = 0 Then gError 46871
    If Len(Trim(FilialPedido.Text)) = 0 Then gError 46872

    objPedidoVenda.lCodigo = CLng(Pedido.Text)
    objPedidoVenda.iFilialEmpresa = Codigo_Extrai(FilialPedido.Text)
    
    'Verifica se o pedido está baixado
    lErro = CF("PedidoVendaBaixado_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 46135 Then gError 46873
    If lErro = SUCESSO Then gError 46876

    'Busca o pedido na tabela de Pedidos de Venda
    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 46874
    If lErro <> SUCESSO Then gError 46877 'Não encontrou
    
    lErro = CF("BloqueiosPV_Le", objPedidoVenda)
    If lErro <> SUCESSO Then gError 51487
    
    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV
        If objBloqueioPV.dtDataLib = DATA_NULA Then
            If objBloqueioPV.iTipoDeBloqueio <> BLOQUEIO_PARCIAL And objBloqueioPV.iTipoDeBloqueio <> BLOQUEIO_NAO_RESERVA Then gError 51488
        End If
    Next
    
    'Verifica se a Filial de Faturamento do Pedido coincide com a filial de trabalho
    If objPedidoVenda.iFilialEmpresaFaturamento <> giFilialEmpresa Then gError 46878

    lErro = CF("Verifica_TipoPagto_Preenchido", objPedidoVenda)
    If lErro <> SUCESSO Then gError 180003

    'Traz os dados do pedido para a tela
    lErro = Traz_Dados_Pedido_Tela(objPedidoVenda)
    If lErro <> SUCESSO Then gError 46875
    
    Exit Sub

Erro_BotaoPedido_Click:

    Select Case gErr

        Case 46871
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)

        Case 46872
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)

        Case 46873, 46874, 46875, 51487, 180003

        Case 46876
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_BAIXADO", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)

        Case 46877
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)

        Case 46878
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFATURAMENTO_DIFERENTE", gErr, objPedidoVenda.iFilialEmpresaFaturamento, objPedidoVenda.lCodigo, giFilialEmpresa)

        Case 51488
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_BLOQUEIO", gErr, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157358)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim objNFiscal As New ClassNFiscal
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 25958

    'Limpa a Tela de Nota fiscal fatura pedido
    lErro = Limpa_Tela_NFiscalFaturaPedido()
    If lErro <> SUCESSO Then gError 500144
    
    'Limpa as coleções globais
'    Set gcolItemPedido = New Collection
'    Set gcolItemQuant = New Collection
'    Set gcolItemUM = New Collection 'Inserido por Wagner
    
    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        'Prepara para chaveamento para NotaFiscalFatura
'        objNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV
'
'        'Faz chaveamento de classe CT para NotaFiscalFatura
'        lErro = objUserControl.CT_Chavea(objNFiscal)
'        If lErro <> SUCESSO Then gError 25943
'
'    End If
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 25943, 25958, 500144 'tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157359)

    End Select

    Exit Sub
    
End Sub

Public Function Traz_Dados_Pedido_Tela(objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objCliente As New ClassCliente
Dim iIndice As Integer
Dim dValorParcelas As Double, bFaturadoParc As Boolean
Dim objItemNF As ClassItemNF
Dim objDataEntrega As ClassDataEntrega, bRecalcularValorProdutos As Boolean
Dim colEmbProdItem As New Collection, bEmissaoForcaRecalculo As Boolean

On Error GoTo Erro_Traz_Dados_Pedido_Tela

    'Indica que o sistema está carregando dados para a tela,
    'e, portanto, algums cálculos e procedimentos não serão executados
    gbCarregandoTela = True
    
    'Lê todos os dados do pedido de venda
    lErro = CF("PedidoDeVenda_Le_Completo", objPedidoVenda, MARCADO)
    If lErro <> SUCESSO And lErro <> 27633 Then gError 46879
    If lErro <> SUCESSO Then gError 46882
    
    If objPedidoVenda.colParcelas.Count = 0 Then gError 26817
    
    CobrancaAutomatica.Value = objPedidoVenda.iCobrancaAutomatica
    
    'Lê as reservas do pedido de venda
    lErro = CF("ReservasPedido_Le", objPedidoVenda)
    If lErro <> SUCESSO Then gError 46880

    lErro = CF("NFiscal_Le_DatasEntrega", objPedidoVenda)
    If lErro <> SUCESSO Then gError 183544

    'Transfere os dados de objPedidoVenda para objNFiscal
    Call Transfere_Dados_Pedido_NFiscal(objPedidoVenda, objNFiscal, bFaturadoParc, bRecalcularValorProdutos)

    For Each objItemNF In objNFiscal.ColItensNF
        For iIndice = objItemNF.colDataEntrega.Count To 1 Step -1
            
            Set objDataEntrega = objItemNF.colDataEntrega.Item(iIndice)
                
            If objDataEntrega.dQuantidadeEntregue >= objDataEntrega.dQuantidade Then
                objItemNF.colDataEntrega.Remove (iIndice)
            Else
                objDataEntrega.dQuantidade = objDataEntrega.dQuantidade - objDataEntrega.dQuantidadeEntregue
            End If
        
        Next
        
    Next

    Call CF2(Me, "NFCF_DataSaida", objNFiscal)
        
'    If giTipoVersao = VERSAO_LIGHT Then
'        objNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV
'    End If
        
    Set gobjNFiscal = objNFiscal
        
    'Carrega os dados de objNFiscal na tela.
    lErro = Traz_NFiscal_Tela(objNFiscal, objPedidoVenda.lNumIntDoc)
    If lErro <> SUCESSO Then gError 46881
   
    'Coloca na tela o acréscimo financeiro
    PercAcrescFin.Text = Format(objPedidoVenda.dPercAcrescFinanceiro * 100, "Fixed")
    
    Call ValorTotal_Calcula
    
    'se o faturamento é parcial ou mudou valor de produtos por incorporar desconto
    If bFaturadoParc Or bRecalcularValorProdutos Then
    
        'Significa que deve recalcular as comissões, pois o BD está armazenando
        'as comissões referentes ao pedido integralmente
        iComissoesAlterada = REGISTRO_ALTERADO
    
    Else
    
        iComissoesAlterada = 0
    
    End If
   
    'Guarda o código e a filial do Pedido
    glCodPedido = objPedidoVenda.lCodigo
    giFilialPedido = objPedidoVenda.iFilialEmpresa
    
    'armazena o dado de fatura integral do pedido
    giFaturaIntegral = objPedidoVenda.iFaturaIntegral
    
    HoraSaida.PromptInclude = False
    HoraSaida.Text = ""
    HoraSaida.PromptInclude = True
    
    'EMBALAGENS
    If bFaturadoParc Then
            
        lErro = CF2(Me, "NFFP_Calcula_PesoKg_Total")
        If lErro <> SUCESSO Then gError 177064
        
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81512
        
        CalculoAuto.Value = vbChecked
            
        'calcula as embalagens do pedido
        Call CalculoAuto_Click
    
    Else
        
        lErro = CF("EmbalagensProdItem_Le", objPedidoVenda.lNumIntDoc, DOCORIGEM_PV, gcolItemPedido, colEmbProdItem)
        If lErro <> SUCESSO Then gError 177064
    
        'Preenche o Grid com as Embalagens dos itens do Pedido de Venda
        lErro = gobjEmbalagem.Preenche_GridEmbalagens(colEmbProdItem)
        If lErro <> SUCESSO Then gError 177064
        
    End If
    
    'Indica que o sistema não está mais carregando dados para a tela
    gbCarregandoTela = False
    
    If (Month(objPedidoVenda.dtDataEmissao) <> Month(objNFiscal.dtDataEmissao)) Or (Year(objPedidoVenda.dtDataEmissao) <> Year(objNFiscal.dtDataEmissao)) Then
        bEmissaoForcaRecalculo = True
    Else
        bEmissaoForcaRecalculo = False
    End If
    
    If bFaturadoParc Or bRecalcularValorProdutos Or bEmissaoForcaRecalculo Then Call ValorTotal_Calcula
    
    Traz_Dados_Pedido_Tela = SUCESSO

    Exit Function

Erro_Traz_Dados_Pedido_Tela:

    'Indica que o sistema não está mais carregando dados para a tela
    gbCarregandoTela = False
    
    Traz_Dados_Pedido_Tela = gErr

    Select Case gErr

        Case 26817
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_PARCELA_COBRANCA", gErr)
        
        Case 46879, 46880, 46881, 81512, 177064, 183544

        Case 46882
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157360)

    End Select

    Exit Function

End Function

Private Sub Transfere_Dados_Pedido_NFiscal(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objNFiscal As ClassNFiscal, bFaturadoParc As Boolean, bRecalcularValorProdutos As Boolean) 'Tirado Byval por Wagner
'Transfere os dados do objPedidoVenda para objNFiscal

Dim objComissaoPV As ClassComissaoPedVendas
Dim objComissaoNF As ClassComissaoNF
Dim objItemPV As ClassItemPedido
Dim dQuantidade As Double, dValorDesconto As Double
Dim objReserva As ClassReserva
Dim colAlocacoes As ColAlocacoesItemNF
Dim dPercDesc As Double
Dim objItemTributacaoNF As ClassTributacaoDocItem
Dim objParcelaPV As ClassParcelaPedidoVenda
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim objProjetoInfo As Object
Dim dFator As Double, dPrecoTotal As Double, iItem As Integer
Dim objComissaoPorItem As ClassComissaoPorItem, objComissaoPorItemAux As ClassComissaoPorItem
Dim objProduto As New ClassProduto
    
On Error GoTo Erro_Transfere_Dados_Pedido_NFiscal
    
    Call gobjTelaProjetoInfo.Obtem_ProjetoInfo(objProjetoInfo, objPedidoVenda.lNumIntDoc, PRJ_CR_TIPO_PV)
    Set objNFiscal.objProjetoInfo = objProjetoInfo
    
    bFaturadoParc = False
    bRecalcularValorProdutos = False
    
    'Verifica se o pedido já foi faturado parcialmente
    iItem = 0
    For Each objItemPV In objPedidoVenda.colItensPedido
        If objItemPV.dQuantFaturada > 0 Then
            bFaturadoParc = True
            'Exit For
        End If
        
        iItem = iItem + 1
        
        'calcula a quantidade
        dQuantidade = objItemPV.dQuantidade - objItemPV.dQuantFaturada - objItemPV.dQuantCancelada
        
        If gobjFAT.iNFFaturaApenasQuantReservPV = MARCADO And Frame1(TAB_Almoxarifado).Enabled Then
            If objItemPV.dQuantReservada < dQuantidade Then
                dQuantidade = objItemPV.dQuantReservada
                bFaturadoParc = True
            End If
        End If

        If dQuantidade <= DELTA_VALORMONETARIO Then
            Call gobjTelaProjetoInfo.Remove_Linha2(objProjetoInfo, iItem)
            iItem = iItem - 1 'Se excluir a linha 2 a 3 do prj vira 2, logo tem que compatibilizar o item a ser pego
        End If
    Next

    Set objNFiscal.objInfoAdic = objPedidoVenda.objInfoAdic

    With objPedidoVenda
        'Passa os dados gerais da tela de pedido para a Nota Fiscal
        objNFiscal.dPesoBruto = .dPesoBruto
        objNFiscal.dPesoLiq = .dPesoLiq
        objNFiscal.dVolumeTotal = .dVolumeTotal
        objNFiscal.dtDataEmissao = gdtDataAtual
        If Not bFaturadoParc Then
            objNFiscal.dValorDesconto = .dValorDesconto
            objNFiscal.dValorFrete = .dValorFrete
            objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
            objNFiscal.dValorSeguro = .dValorSeguro
            objNFiscal.dValorTotal = .dValorTotal
            objNFiscal.dValorDescontoItens = .dValorDescontoItens
            objNFiscal.dValorDescontoTit = .dValorDescontoTit
            objNFiscal.dValorItens = .dValorItens
        End If
        objNFiscal.iCanal = .iCanalVenda
        objNFiscal.iCodTransportadora = .iCodTransportadora
        objNFiscal.iCodTranspRedesp = .iCodTranspRedesp
        objNFiscal.iDetPagFrete = .iDetPagFrete
        objNFiscal.iFilialCli = .iFilial
        objNFiscal.iFilialEmpresa = .iFilialEmpresaFaturamento
        objNFiscal.iFilialEntrega = .iFilialEntrega
        objNFiscal.iTabelaPreco = .iTabelaPreco
        objNFiscal.lVolumeQuant = .lVolumeQuant
        objNFiscal.lCliente = .lCliente
        objNFiscal.lNumPedidoVenda = .lCodigo
        objNFiscal.sMensagemNota = .sMensagemNota
        objNFiscal.iFilialPedido = .iFilialEmpresa
        objNFiscal.sPlaca = .sPlaca
        objNFiscal.sNaturezaOp = .sNaturezaOp
        objNFiscal.sNumPedidoTerc = .sPedidoCliente
        objNFiscal.sPlacaUF = .sPlacaUF
        objNFiscal.lVolumeEspecie = .lVolumeEspecie
        objNFiscal.lVolumeMarca = .lVolumeMarca
        objNFiscal.sVolumeNumero = .sVolumeNumero
        objNFiscal.iFreteRespons = .iFreteRespons
        objNFiscal.iCondicaoPagto = .iCondicaoPagto
        If Len(Trim(Serie.Text)) > 0 Then objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
        objNFiscal.dtDataReferencia = .dtDataReferencia

    End With
    
    'Transfere os dados das comissões do pedido para a coleção de comissões da Nota Fiscal
    For Each objComissaoPV In objPedidoVenda.colComissoes
        
        With objComissaoPV
            
            'Instancia um novo objComissaoNF
            Set objComissaoNF = New ClassComissaoNF
            
            'Transfere os dados do objComissaoPV para o objComissaoNF
            objComissaoNF.iCodVendedor = .iCodVendedor
            objComissaoNF.dValorBase = .dValorBase
            objComissaoNF.dPercentual = .dPercentual
            objComissaoNF.dValor = .dValor
            objComissaoNF.dPercentualEmissao = .dPercentualEmissao
            objComissaoNF.dValorEmissao = .dValorEmissao
            objComissaoNF.iIndireta = .iIndireta
            objComissaoNF.iSeq = .iSeq
            
        End With
                
        For Each objComissaoPorItem In objComissaoPV.colLinhas
        
            Set objComissaoPorItemAux = New ClassComissaoPorItem
            
            With objComissaoPorItemAux
            
                .dPercentual = objComissaoPorItem.dPercentual
                .dPercentualEmissao = objComissaoPorItem.dPercentualEmissao
                .dValor = objComissaoPorItem.dValor
                .dValorBase = objComissaoPorItem.dValorBase
                .dValorEmissao = objComissaoPorItem.dValorEmissao
                .iLinha = objComissaoPorItem.iLinha
            
            End With
            
            objComissaoNF.colLinhas.Add objComissaoPorItemAux
            
        Next
        
        'Adiciona o objComissaoNF à coleção
        objNFiscal.ColComissoesNF.Add objComissaoNF
        
    Next
    '??? objnfiscal.iComissaoAutomatica=objPedidoVenda.iComissaoAutomatica
    
    If objPedidoVenda.iComissaoAutomatica = MARCADO Then
        ComissaoAutomatica.Value = vbChecked
    Else
        ComissaoAutomatica.Value = vbUnchecked
    End If
    
    lErro = ParcelasPV_Gerando_ParcelasNF(objPedidoVenda, objNFiscal)
    If lErro <> SUCESSO Then gError 51614
    
    'Reinicializa as coleções de itens
    Set gcolItemPedido = New Collection
    Set gcolItemQuant = New Collection
    Set gcolItemUM = New Collection 'Inserido por Wagner
    
    'Para cada item do Pedido
    For Each objItemPV In objPedidoVenda.colItensPedido
        
        lErro = CF("UM_Conversao_Trans", objItemPV.iClasseUM, objItemPV.sUnidadeMed, objItemPV.sUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 51614
               
        'calcula a quantidade
        dQuantidade = objItemPV.dQuantidade - objItemPV.dQuantFaturada - objItemPV.dQuantCancelada
        
        If gobjFAT.iNFFaturaApenasQuantReservPV = MARCADO And Frame1(TAB_Almoxarifado).Enabled Then
            If objItemPV.dQuantReservada < dQuantidade Then
                dQuantidade = objItemPV.dQuantReservada
            End If
        End If

        If dQuantidade > 0 Then
            'Calcula o percentual de desconto
            dPercDesc = objItemPV.dValorDesconto / (objItemPV.dPrecoTotal + objItemPV.dValorDesconto)
                       
            Set colAlocacoes = New ColAlocacoesItemNF
            'Transporta as reservas do PV para a coleção de alocações da NF
            For Each objReserva In objItemPV.colReserva
                
                objAlmoxarifado.iCodigo = objReserva.iAlmoxarifado
                objAlmoxarifado.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
                'Lê o Almoxarifado
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> 25056 And lErro <> SUCESSO Then gError 51560
                If lErro = 25056 Then gError 51561
                              
                If objAlmoxarifado.iFilialEmpresa = giFilialEmpresa Then
                    If bFaturadoParc And objReserva.dQuantidade > dQuantidade * dFator Then
                        colAlocacoes.Add objReserva.iAlmoxarifado, objAlmoxarifado.sNomeReduzido, dQuantidade * dFator
                    Else
                        colAlocacoes.Add objReserva.iAlmoxarifado, objAlmoxarifado.sNomeReduzido, objReserva.dQuantidade
                    End If
                End If
                
            Next
        
            Call gobjTribTab.ItemNFiscal_ObterTrib_ItemPedido(objItemPV, objItemTributacaoNF)
            
            With objItemPV
                                
                'calcula a quantidade do item em funcao da quantidade alocada. Serve para alguns clientes
                lErro = CF("Calcula_Quantidade_Alocada", colAlocacoes, dQuantidade)
                If lErro <> SUCESSO Then gError 105005
                                
                If dQuantidade > 0 Then
                
                    If bFaturadoParc Then
                        dValorDesconto = Arredonda_Moeda(dPercDesc * (dQuantidade * .dPrecoUnitario))
                    Else
                        dValorDesconto = .dValorDesconto
                    End If
                                
                    'Adiciona o item na coleção de itens da NF
                    Set objItemNF = objNFiscal.ColItensNF.Add(0, .iItem, .sProduto, .sUnidadeMed, dQuantidade, .dPrecoUnitario, dPercDesc, dValorDesconto, .dtDataEntrega, .sDescricao, .dValorAbatComissao, .lNumIntDoc, 0, 0, .lNumIntDoc, colAlocacoes, 0, "", "", .iStatus, 0, "", 0, 0, .iControleEstoque, .sUMEstoque, .iClasseUM, 0)
                                       
                    lErro = CF2(Me, "NFFP_Transfere_Dados_Pedido_NFiscal", objItemNF, objItemPV, dQuantidade, objNFiscal)
                    If lErro <> SUCESSO Then gError 177051
                    
                    Set objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).objTributacaoItemNF = objItemTributacaoNF
                    'Guarda a número interno e a quantidade inicial do ítem
                    gcolItemPedido.Add .lNumIntDoc
                    gcolItemQuant.Add dQuantidade
                    gcolItemUM.Add .sUnidadeMed 'Inserido por Wagner
              
                    objItemNF.iRecebForaFaixa = .iRecebForaFaixa
                    objItemNF.dPercentMaisReceb = .dPercentMaisReceb
                    objItemNF.dPercentMenosReceb = .dPercentMenosReceb
                    objItemNF.dComissao = .dComissao
                    objItemNF.iTabelaPreco = .iTabelaPreco
              
                    objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).iPossuiGrade = .iPossuiGrade
                                                   
                    Set objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).colDataEntrega = objItemPV.colItemPVEntrega
                                                   
                    Set objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).colItensRomaneioGrade = New Collection
                               
                    Call Move_ItensGrade_Tela(objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).colItensRomaneioGrade, objItemPV.colItensRomaneioGrade, True)

                    Set objNFiscal.ColItensNF(objNFiscal.ColItensNF.Count).objInfoAdicDocItem = objItemPV.objInfoAdicDocItem

                    dPrecoTotal = (objItemNF.dQuantidade * objItemNF.dPrecoUnitario) - objItemNF.dValorDesconto
                    
                    If gobjFAT.iTrazDescPVParaNF = DESMARCADO And objItemNF.dValorDesconto > DELTA_VALORMONETARIO Then
                        bRecalcularValorProdutos = True
                        objItemNF.dPercDesc = 0
                        objItemNF.dValorDesconto = 0
                        objItemNF.dPrecoUnitario = Arredonda_Moeda(dPrecoTotal / objItemNF.dQuantidade)
                    End If

                    If UCase(objItemTributacaoNF.sUnidadeMed) = UCase(objItemTributacaoNF.sUMTrib) Then
                    
                        objItemTributacaoNF.dQtdTrib = dQuantidade
                        
                    Else
                    
                        objProduto.sCodigo = objItemPV.sProduto
                        
                        lErro = CF("Produto_Le", objProduto)
                        If lErro <> SUCESSO And lErro <> 28030 Then Error 27811
                        If lErro <> SUCESSO Then Error 27811
                        
                        'Faz a conversão da UM da tela para a UM tributaria
                        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemPV.sUnidadeMed, objItemTributacaoNF.sUMTrib, dFator)
                        If lErro <> SUCESSO Then Error 27811
                            
                        objItemTributacaoNF.dQtdTrib = dQuantidade * dFator
                        
                    End If
                    
                    objItemTributacaoNF.dValorUnitTrib = Arredonda_Moeda((objItemNF.dQuantidade * objItemNF.dPrecoUnitario) / objItemTributacaoNF.dQtdTrib, 4)
                
                End If

            End With

        End If

    Next

    'Carrega os dados globais da tributação do PV para a NF
    Call gobjTribTab.NFiscal_ObterTrib_Pedido(objPedidoVenda, objNFiscal)
    
    lErro = CF2(Me, "NFFP_Transfere_Dados_Pedido_NFiscal2", objPedidoVenda, objNFiscal, bFaturadoParc)
    If lErro <> SUCESSO Then gError 124251
    
    If bRecalcularValorProdutos Then
        objNFiscal.dValorItens = 0
    End If
    
    Exit Sub
    
Erro_Transfere_Dados_Pedido_NFiscal:

    Select Case gErr
    
        Case 51560, 105005, 124251, 177051
        
        Case 51561
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157361)
    
    End Select
    
End Sub

'Private Sub Transfere_Dados_Tributacao_Pedido_NFiscal(objPedidoVenda As ClassPedidoDeVenda, objNFiscal As ClassNFiscal)
''Transfere os dados da tributação dos itens de pedido de venda passados para os itens do objNFiscal
'
'    Set objNFiscal.objTributacaoNF = New ClassTributacaoNF
'
'    With objNFiscal.objTributacaoNF
'
'        .dICMSBase = objPedidoVenda.objTributacaoPV.dICMSBase
'        .dICMSSubstBase = objPedidoVenda.objTributacaoPV.dICMSSubstBase
'        .dICMSSubstValor = objPedidoVenda.objTributacaoPV.dICMSSubstValor
'        .dICMSValor = objPedidoVenda.objTributacaoPV.dICMSValor
'        .dIPIBase = objPedidoVenda.objTributacaoPV.dIPIBase
'        .dIPIValor = objPedidoVenda.objTributacaoPV.dIPIValor
'        .dIRRFAliquota = objPedidoVenda.objTributacaoPV.dIRRFAliquota
'        .dIRRFBase = objPedidoVenda.objTributacaoPV.dIRRFBase
'        .dIRRFValor = objPedidoVenda.objTributacaoPV.dIRRFValor
'        .iISSIncluso = objPedidoVenda.objTributacaoPV.iISSIncluso
'        .iTipoTributacao = objPedidoVenda.objTributacaoPV.iTipoTributacao
'        .sNaturezaOpInterna = objPedidoVenda.sNaturezaOp
'
'        Set .objTributacaoDesconto = New ClassTribComplNF
'        Set .objTributacaoFrete = New ClassTribComplNF
'        Set .objTributacaoOutras = New ClassTribComplNF
'        Set .objTributacaoSeguro = New ClassTribComplNF
'
'        .objTributacaoDesconto.dICMSAliquota = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.dICMSAliquota
'        .objTributacaoDesconto.dICMSBase = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.dICMSBase
'        .objTributacaoDesconto.dICMSPercRedBase = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.dICMSPercRedBase
'        .objTributacaoDesconto.dICMSSubstAliquota = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.dICMSSubstAliquota
'        .objTributacaoDesconto.dICMSSubstBase = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.dICMSSubstBase
'        .objTributacaoDesconto.dICMSSubstValor = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.dICMSSubstValor
'        .objTributacaoDesconto.dICMSValor = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.dICMSValor
'        .objTributacaoDesconto.dIPIAliquota = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.dIPIAliquota
'        .objTributacaoDesconto.dIPIBaseCalculo = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.dIPIBaseCalculo
'        .objTributacaoDesconto.dIPIPercRedBase = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.dIPIPercRedBase
'        .objTributacaoDesconto.dIPIValor = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.dIPIValor
'        .objTributacaoDesconto.iICMSTipo = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.iICMSTipo
'        .objTributacaoDesconto.iIPITipo = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.iIPITipo
'        .objTributacaoDesconto.iTipo = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.iTipoTributacaoManual
'        .objTributacaoDesconto.iTipoTributacao = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.iTipoTributacao
'        .objTributacaoDesconto.lNumIntDoc = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.lNumIntDoc
'        .objTributacaoDesconto.sNaturezaOp = objPedidoVenda.objTributacaoPV.objTributacaoDesconto.sNaturezaOp
'
'        .objTributacaoFrete.dICMSAliquota = objPedidoVenda.objTributacaoPV.objTributacaoFrete.dICMSAliquota
'        .objTributacaoFrete.dICMSBase = objPedidoVenda.objTributacaoPV.objTributacaoFrete.dICMSBase
'        .objTributacaoFrete.dICMSPercRedBase = objPedidoVenda.objTributacaoPV.objTributacaoFrete.dICMSPercRedBase
'        .objTributacaoFrete.dICMSSubstAliquota = objPedidoVenda.objTributacaoPV.objTributacaoFrete.dICMSSubstAliquota
'        .objTributacaoFrete.dICMSSubstBase = objPedidoVenda.objTributacaoPV.objTributacaoFrete.dICMSSubstBase
'        .objTributacaoFrete.dICMSSubstValor = objPedidoVenda.objTributacaoPV.objTributacaoFrete.dICMSSubstValor
'        .objTributacaoFrete.dICMSValor = objPedidoVenda.objTributacaoPV.objTributacaoFrete.dICMSValor
'        .objTributacaoFrete.dIPIAliquota = objPedidoVenda.objTributacaoPV.objTributacaoFrete.dIPIAliquota
'        .objTributacaoFrete.dIPIBaseCalculo = objPedidoVenda.objTributacaoPV.objTributacaoFrete.dIPIBaseCalculo
'        .objTributacaoFrete.dIPIPercRedBase = objPedidoVenda.objTributacaoPV.objTributacaoFrete.dIPIPercRedBase
'        .objTributacaoFrete.dIPIValor = objPedidoVenda.objTributacaoPV.objTributacaoFrete.dIPIValor
'        .objTributacaoFrete.iICMSTipo = objPedidoVenda.objTributacaoPV.objTributacaoFrete.iICMSTipo
'        .objTributacaoFrete.iIPITipo = objPedidoVenda.objTributacaoPV.objTributacaoFrete.iIPITipo
'        .objTributacaoFrete.iTipo = objPedidoVenda.objTributacaoPV.objTributacaoFrete.iTipoTributacaoManual
'        .objTributacaoFrete.iTipoTributacao = objPedidoVenda.objTributacaoPV.objTributacaoFrete.iTipoTributacao
'        .objTributacaoFrete.lNumIntDoc = objPedidoVenda.objTributacaoPV.objTributacaoFrete.lNumIntDoc
'        .objTributacaoFrete.sNaturezaOp = objPedidoVenda.objTributacaoPV.objTributacaoFrete.sNaturezaOp
'
'        .objTributacaoOutras.dICMSAliquota = objPedidoVenda.objTributacaoPV.objTributacaoOutras.dICMSAliquota
'        .objTributacaoOutras.dICMSBase = objPedidoVenda.objTributacaoPV.objTributacaoOutras.dICMSBase
'        .objTributacaoOutras.dICMSPercRedBase = objPedidoVenda.objTributacaoPV.objTributacaoOutras.dICMSPercRedBase
'        .objTributacaoOutras.dICMSSubstAliquota = objPedidoVenda.objTributacaoPV.objTributacaoOutras.dICMSSubstAliquota
'        .objTributacaoOutras.dICMSSubstBase = objPedidoVenda.objTributacaoPV.objTributacaoOutras.dICMSSubstBase
'        .objTributacaoOutras.dICMSSubstValor = objPedidoVenda.objTributacaoPV.objTributacaoOutras.dICMSSubstValor
'        .objTributacaoOutras.dICMSValor = objPedidoVenda.objTributacaoPV.objTributacaoOutras.dICMSValor
'        .objTributacaoOutras.dIPIAliquota = objPedidoVenda.objTributacaoPV.objTributacaoOutras.dIPIAliquota
'        .objTributacaoOutras.dIPIBaseCalculo = objPedidoVenda.objTributacaoPV.objTributacaoOutras.dIPIBaseCalculo
'        .objTributacaoOutras.dIPIPercRedBase = objPedidoVenda.objTributacaoPV.objTributacaoOutras.dIPIPercRedBase
'        .objTributacaoOutras.dIPIValor = objPedidoVenda.objTributacaoPV.objTributacaoOutras.dIPIValor
'        .objTributacaoOutras.iICMSTipo = objPedidoVenda.objTributacaoPV.objTributacaoOutras.iICMSTipo
'        .objTributacaoOutras.iIPITipo = objPedidoVenda.objTributacaoPV.objTributacaoOutras.iIPITipo
'        .objTributacaoOutras.iTipo = objPedidoVenda.objTributacaoPV.objTributacaoOutras.iTipoTributacaoManual
'        .objTributacaoOutras.iTipoTributacao = objPedidoVenda.objTributacaoPV.objTributacaoOutras.iTipoTributacao
'        .objTributacaoOutras.lNumIntDoc = objPedidoVenda.objTributacaoPV.objTributacaoOutras.lNumIntDoc
'        .objTributacaoOutras.sNaturezaOp = objPedidoVenda.objTributacaoPV.objTributacaoOutras.sNaturezaOp
'
'        .objTributacaoSeguro.dICMSAliquota = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.dICMSAliquota
'        .objTributacaoSeguro.dICMSBase = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.dICMSBase
'        .objTributacaoSeguro.dICMSPercRedBase = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.dICMSPercRedBase
'        .objTributacaoSeguro.dICMSSubstAliquota = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.dICMSSubstAliquota
'        .objTributacaoSeguro.dICMSSubstBase = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.dICMSSubstBase
'        .objTributacaoSeguro.dICMSSubstValor = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.dICMSSubstValor
'        .objTributacaoSeguro.dICMSValor = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.dICMSValor
'        .objTributacaoSeguro.dIPIAliquota = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.dIPIAliquota
'        .objTributacaoSeguro.dIPIBaseCalculo = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.dIPIBaseCalculo
'        .objTributacaoSeguro.dIPIPercRedBase = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.dIPIPercRedBase
'        .objTributacaoSeguro.dIPIValor = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.dIPIValor
'        .objTributacaoSeguro.iICMSTipo = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.iICMSTipo
'        .objTributacaoSeguro.iIPITipo = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.iIPITipo
'        .objTributacaoSeguro.iTipo = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.iTipoTributacaoManual
'        .objTributacaoSeguro.iTipoTributacao = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.iTipoTributacao
'        .objTributacaoSeguro.lNumIntDoc = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.lNumIntDoc
'        .objTributacaoSeguro.sNaturezaOp = objPedidoVenda.objTributacaoPV.objTributacaoSeguro.sNaturezaOp
'
'    End With
'
'End Sub

Function Preenche_Grid_Parcelas(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com as Parcelas da Nota Fiscal

Dim objParcela As ClassParcelaReceber
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim dValorDesconto As Double

    Call Grid_Limpa(objGridParcelas)
    
    iIndice = 0
            
    'PAra cada parcela da coleção de parcelas
    For Each objParcela In objNFiscal.ColParcelaReceber
        
        iIndice = iIndice + 1
        'Preenche o grid com os dados da parcela
        GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objParcela.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objParcela.dValor, "Standard")
        If objParcela.dtDesconto1Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col) = Format(objParcela.dtDesconto1Ate, "dd/mm/yyyy")
        If objParcela.dtDesconto2Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col) = Format(objParcela.dtDesconto2Ate, "dd/mm/yyyy")
        If objParcela.dtDesconto3Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col) = Format(objParcela.dtDesconto3Ate, "dd/mm/yyyy")
        If objParcela.iDesconto1Codigo = VALOR_FIXO Or objParcela.iDesconto1Codigo = VALOR_ANT_DIA Or objParcela.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(objParcela.dDesconto1Valor, "Standard")
        ElseIf objParcela.iDesconto1Codigo = Percentual Or objParcela.iDesconto1Codigo = PERC_ANT_DIA Or objParcela.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Percentual_Col) = Format(objParcela.dDesconto1Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcela.dValor > 0 Then
                dValorDesconto = objParcela.dDesconto1Valor * objParcela.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        If objParcela.iDesconto2Codigo = VALOR_FIXO Or objParcela.iDesconto2Codigo = VALOR_ANT_DIA Or objParcela.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(objParcela.dDesconto2Valor, "Standard")
        ElseIf objParcela.iDesconto2Codigo = Percentual Or objParcela.iDesconto2Codigo = PERC_ANT_DIA Or objParcela.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col) = Format(objParcela.dDesconto2Valor, "Percent")
            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col) = Format(objParcela.dDesconto2Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcela.dValor > 0 Then
                dValorDesconto = objParcela.dDesconto2Valor * objParcela.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        If objParcela.iDesconto3Codigo = VALOR_FIXO Or objParcela.iDesconto3Codigo = VALOR_ANT_DIA Or objParcela.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(objParcela.dDesconto3Valor, "Standard")
        ElseIf objParcela.iDesconto3Codigo = Percentual Or objParcela.iDesconto3Codigo = PERC_ANT_DIA Or objParcela.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Percentual_Col) = Format(objParcela.dDesconto3Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcela.dValor > 0 Then
                dValorDesconto = objParcela.dDesconto3Valor * objParcela.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        
        For iIndice2 = 0 To Desconto1Codigo.ListCount - 1
            If Desconto1Codigo.ItemData(iIndice2) = objParcela.iDesconto1Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col) = Desconto1Codigo.List(iIndice2)
            If Desconto2Codigo.ItemData(iIndice2) = objParcela.iDesconto2Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col) = Desconto2Codigo.List(iIndice2)
            If Desconto3Codigo.ItemData(iIndice2) = objParcela.iDesconto3Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col) = Desconto3Codigo.List(iIndice2)
        Next
        
        For iIndice2 = 0 To TipoPagto.ListCount - 1
            If TipoPagto.ItemData(iIndice2) = objParcela.iTipoPagto Then
                GridParcelas.TextMatrix(iIndice, iGrid_TipoPagto_Col) = TipoPagto.List(iIndice2)
                Exit For
            End If
        Next
        
    Next
    
    'Inicializa o número de linhas existentes no grid
    objGridParcelas.iLinhasExistentes = iIndice
    
    Preenche_Grid_Parcelas = SUCESSO
    
    Exit Function

End Function

Public Sub BotaoFechar_Click()
    'Fecha a tela
    Unload Me

End Sub

Public Sub BotaoLocalizacao_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim iPreenchido As Integer
Dim sProduto As String
Dim dQuantidade As Double
Dim ColItensNF As New ColItensNF
Dim colAlocacoes As ColAlocacoesItemNF
Dim objProduto As New ClassProduto
Dim objItemNF As New ClassItemNF
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim lbLabel As label
Dim dPrecoUnitario As Double
Dim vbMsgRes As VbMsgBoxResult
Dim objNFiscal As New ClassNFiscal
Dim colItensRomaneio As New Collection
Dim iIndice2 As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objItemRomaneioAux As ClassItemRomaneioGrade
Dim objProdutoAux As New ClassProduto
Dim bAchou  As Boolean

On Error GoTo Erro_BotaoLocalizacao_Click:
    
    'Se não existir itens no grid, sai.
    If objGridItens.iLinhasExistentes = 0 Then gError 43699
    
    'Para cada ítem do grid
    For iIndice = 1 To objGridItens.iLinhasExistentes
        
        Set colItensRomaneio = New Collection
        
        'recolhe a quantidade
        dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        
        'Critica o produto
        lErro = CF("Produto_Critica_Filial2", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), objProduto, iPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 42309
        If lErro = 51381 Then gError 86376
                
        'Se o produto é gerencial sem ser de grade ==> Erro
        If lErro = 86295 And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then
            gError 86377
        ElseIf lErro = 86295 Then
            
             If objProduto.iKitVendaComp <> MARCADO Then
            
                'Lê os filhos analíticos do produto pai de grade passado
                lErro = CF("Produto_Le_Filhos_Grade", objProduto, colItensRomaneio)
                If lErro <> SUCESSO And lErro <> 86304 Then gError 86306
                
                'O produto não tem filhos de grade ou seus filhos são analíticos
                If lErro = 86304 Then gError 86307
            
            Else
            
                lErro = CF("ProdutoKitVenda_Le_Romaneio", objProduto, colItensRomaneio)
                If lErro <> SUCESSO Then gError 177582
            
            End If
            
            'Para cada item lido
            For iIndice2 = colItensRomaneio.Count To 1 Step -1
            
                Set objItemRomaneio = colItensRomaneio(iIndice2)
                
                'Lê o Produto
                objProdutoAux.sCodigo = objItemRomaneio.sProduto
                lErro = CF("Produto_Le", objProdutoAux)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 86378
                If lErro <> SUCESSO Then gError 86379
                
                'Se o produto controlar Reserva
                If objProdutoAux.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    bAchou = False
                    For Each objItemRomaneioAux In gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade
                        If objItemRomaneio.sProduto = objItemRomaneioAux.sProduto Then
                            Call Transfere_Dados_ItensRomaneio(colItensRomaneio, gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade)
                            bAchou = True
                            Exit For
                        End If
                    Next
                End If
                If Not bAchou Or objProdutoAux.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    colItensRomaneio.Remove iIndice2
                End If
            Next
        
        End If
        'Se o produto fizer controle de estoque
        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then

            Set colAlocacoes = New ColAlocacoesItemNF
            'Recolhe as alocações do item
            If gobjNFiscal.ColItensNF(iIndice).iPossuiGrade = DESMARCADO And objProduto.iKitVendaComp = DESMARCADO Then Call Recolhe_Alocacoes(iIndice, colAlocacoes)

            Set objItemNF = ColItensNF.Add(0, iIndice, objProduto.sCodigo, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), dQuantidade, 0, 0, 0, DATA_NULA, GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col), 0, 0, 0, 0, gcolItemPedido(iIndice), colAlocacoes, 0, "", "", 0, 0, "", 0, 0, 0, objProduto.sSiglaUMEstoque, 0, 0)
                        
            lErro = CF2(Me, "NFFP_Move_GridItem_Memoria", gobjNFiscal, objItemNF, iIndice)
            If lErro <> SUCESSO Then gError 197705
                        
            ColItensNF(ColItensNF.Count).iPossuiGrade = gobjNFiscal.ColItensNF(iIndice).iPossuiGrade
            
            If objProduto.iKitVendaComp = MARCADO Then
                ColItensNF(ColItensNF.Count).iPossuiGrade = MARCADO
            End If
            
            If ColItensNF(ColItensNF.Count).iPossuiGrade = MARCADO Then
                Set ColItensNF(ColItensNF.Count).colItensRomaneioGrade = colItensRomaneio
            End If
            
        End If

    Next
    'Chama a tela de localização de produtos
    Call Chama_Tela_Modal("LocalizacaoProduto", ColItensNF, Codigo_Extrai(TipoNFiscal.Text))
       
    'Limpa o grid de Alocações
    Call Grid_Limpa(objGridAlocacoes)
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
    'para cada item
    For Each objItemNF In ColItensNF
        
        objProduto.sCodigo = objItemNF.sProduto
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 46885
        If lErro <> SUCESSO Then gError 46889
    
        'Se as UM de estoque e de venda do itemNF não estiverem preenchidas
        If Len(Trim(objItemNF.sUMEstoque)) = 0 Or Len(Trim(objItemNF.sUnidadeMed)) = 0 Then
            'Preenche as UMs do ItemNF
            objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
            objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
        End If
        
        If objProduto.iKitVendaComp = MARCADO Then
            objItemNF.iPossuiGrade = DESMARCADO
        End If
            
        'Formata o Produto
        lErro = CF("Produto_Formata", GridItens.TextMatrix(objItemNF.iItem, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 46886
        'Se o produto foi substituído
        If sProduto <> objItemNF.sProduto Then
            'Avisa que o produto não pode ser substituido
            lErro = Rotina_Aviso(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_SUBSTITUIDO")
            'Limpa as alocações feitas
            Set objItemNF.colAlocacoes = New ColAlocacoesItemNF
        End If
    Next
    'Passa os itens com suas alocações para o objNFiscal
    Set objNFiscal.ColItensNF = ColItensNF
    
    'Preenche o grid de Alocações
    Call CF2(Me, "NFFP_Preenche_GridAlocacoes", objNFiscal)
    
    For iIndice = 1 To ColItensNF.Count
        Call Transfere_Dados_ItensRomaneio(gobjNFiscal.ColItensNF(ColItensNF(iIndice).iItem).colItensRomaneioGrade, objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade)
    Next
    
    Exit Sub

Erro_BotaoLocalizacao_Click:

    Select Case gErr
    
        Case 26777
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, sProduto)
    
        Case 43699
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)

        Case 46885, 46886, 46888, 86306, 86378, 177582, 197705
        
        Case 46889
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 86307
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, sProduto)
        
        Case 86376
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProduto)

        Case 86379
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objItemRomaneio.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157362)

    End Select

    Exit Sub

End Sub

Public Sub BotaoVerPedido_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoVerPedido_Click

    'Verifica se o Pedido e a filial do pedido foram informados
    If Len(Trim(Pedido.Text)) = 0 Then gError 46892
    If Len(Trim(FilialPedido.Text)) = 0 Then gError 46893
    
    objPedidoVenda.lCodigo = CLng(Pedido.Text)
    objPedidoVenda.iFilialEmpresa = Codigo_Extrai(FilialPedido.Text)
    
    'Verifica se o Pedido está baixado
    lErro = CF("PedidoVendaBaixado_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 46135 Then gError 46894
    If lErro = SUCESSO Then gError 46896 'Se estiver , erro.
    
    'Busca o Pedido na tabela de pedidos de venda
    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 46895
    If lErro <> SUCESSO Then gError 46897 'Não encontrou, erro.
    
    'Chama  a tela de Pedido de Venda passando o pedido de venda da tela
    Call Chama_Tela("PedidoVenda_Consulta", objPedidoVenda)
    
    Exit Sub

Erro_BotaoVerPedido_Click:

    Select Case gErr
    
        Case 46892
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)
            
        Case 46893
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)
        
        Case 46894, 46895
        
        Case 46896
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_BAIXADO", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
        
        Case 46897
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157363)

    End Select

    Exit Sub

End Sub

Public Sub Quantidade_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
    Call CF2(Me, "NFFP_Quantidade_GotFocus")
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

Public Sub PrecoUnitario_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrecoUnitario_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub PrecoUnitario_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub PrecoUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoUnitario
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PercentDesc_Change()
    
    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

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

Public Sub Desconto_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

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

Public Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer
Dim iIndice As Integer
Dim dValorProdutos As Double
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_GridItens_KeyDown

    'Armazena o número de linhas existentes e o item atual selecionado
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iItemAtual = GridItens.Row

    'Se a tecla for a delete e o Pedido só aceitar faturamento integral
    'não permite a exclusão da linha
    If KeyCode = vbKeyDelete And giFaturaIntegral <> 0 Then gError 46899
        
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    'Se algum ítem foi excluído
    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then
        
        '************ grade ************
        GridItens.TextMatrix(GridItens.Row, 0) = GridItens.Row
        gobjNFiscal.ColItensNF.Remove GridItens.Row
        
       Call gobjTelaProjetoInfo.Remove_Linha(iItemAtual)
        
        For iLinhasExistentesAnterior = 1 To objGridItens.iLinhasExistentes
            If gobjNFiscal.ColItensNF(iLinhasExistentesAnterior).iPossuiGrade = MARCADO Then
                GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = "# " & iLinhasExistentesAnterior
            Else
                GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = iLinhasExistentesAnterior
            End If
        Next
        
        GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = iLinhasExistentesAnterior
        '*******************************
        
        'Indica que houve alteração nos dados base para o cálculo de comissões
        iComissoesAlterada = REGISTRO_ALTERADO

        If objGridAlocacoes.iLinhasExistentes > 0 Then
            'Remove as alocações do Pedido
            Call Remove_Alocacoes(iItemAtual)
            'Atualiza o número do item nas outras alocações
            For iIndice = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))) > 0 Then If CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) > iItemAtual Then GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col) = CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) - 1
            Next
        End If
        
        gcolItemPedido.Remove iItemAtual
        gcolItemQuant.Remove iItemAtual
        gcolItemUM.Remove iItemAtual
        
        If objGridItens.iLinhasExistentes <> 0 Then
            Call PrecoTotal_Calcula(objGridItens.iLinhasExistentes, False)
        Else
            If StrParaDbl(ValorDesconto.Text) <> 0 Then
                'Avisa ao usuário
                vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", ValorDesconto.Text, 0)
                
                'Limpa o valor de desconto
                gdDesconto = 0
                ValorDesconto.Text = ""
                iValorDescontoAlterado = 0
                
                Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, 0)

            End If
        End If
        
        'Recalcula o valor dos produtos
        Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)
        'coloca na tela
        ValorProdutos.Caption = Format(dValorProdutos - StrParaDbl(ValorDesconto.Text), "Standard")

        Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotalB_Col, dValorProdutos)

        ValorProdutos2.Caption = Format(dValorProdutos, "Standard")

        lErro = CF2(Me, "NFFP_ValorDesconto_Validate")
        If lErro <> SUCESSO Then gError 132037
        
        If objGridItens.iLinhasExistentes = 0 Then Call Grid_Limpa(objGridParcelas)

''*** incluidos p/tratamento de tributacao *******************************
        'Ataliza a tributação
        lErro = gobjTribTab.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 46898

'*** fim tributacao

        'Recalcula o valor total
        Call ValorTotal_Calcula
        
        'acerta o grid de embalagens
        Call gobjEmbalagem.Exclusao_Item_GridEmb(iItemAtual)
        
        lErro = CF2(Me, "NFFP_Calcula_PesoKg_Total")
        If lErro <> SUCESSO Then gError 177065
    
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81511
    
        'Rastreamento
        lErro = gobjRastreamento.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 83395
    
        For iIndice = 1 To gobjNFiscal.ColItensNF.Count
            gobjNFiscal.ColItensNF(iIndice).iItem = iIndice 'Incluído para ao alterar a grade ou a reserva não fazer do item errado
        Next
        
    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr
    
        Case 46898, 81511, 83395, 132037, 177065
        
        Case 46899
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_FATURA_INTEGRAL", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157364)
            
    End Select

    Exit Sub

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

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'Tratamento de saída de célula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 46900
        
        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridParcelas.Name

                lErro = Saida_Celula_GridParcelas(objGridInt)
                If lErro <> SUCESSO Then gError 46901

            'Se for o GridItens
            Case GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 46902

            'Se for o GridComissoes
            Case GridComissoes.Name

                lErro = objTabComissoes.Saida_Celula_GridComissoes(objGridInt)
                If lErro <> SUCESSO Then gError 46903

            'Rastreamento
            Case GridRastro.Name

                lErro = gobjRastreamento.Saida_Celula()
                If lErro <> SUCESSO Then gError 83396

            Case GridEmb.Name
                
                lErro = gobjEmbalagem.Saida_Celula_GridEmb()
                If lErro <> SUCESSO Then gError 83396
                
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 46904

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 46900, 46901, 46902, 46903, 83396
        
        Case 46904
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157365)

    End Select

    Exit Function

End Function

Private Function Recolhe_Alocacoes(iItem As Integer, colAlocacoes As ColAlocacoesItemNF) As Long
'Recolhe as Alocações

Dim lErro As Long
Dim iIndice As Integer
Dim iItemLinha As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Recolhe_Alocacoes
    'Para cada alocação do grid
    For iIndice = 1 To objGridAlocacoes.iLinhasExistentes
        'Verifica a qual item pertence a alocação
        iItemLinha = StrParaInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))
        'Se a alocação for do item passado
        If iItemLinha = iItem Then
            
            objAlmoxarifado.sNomeReduzido = GridAlocacao.TextMatrix(iIndice, iGrid_AlmoxAloc_Col)
            objAlmoxarifado.iCodigo = 0

            For Each objCodigoNome In colCodigoNome
                If objCodigoNome.sNome = objAlmoxarifado.sNomeReduzido Then
                    objAlmoxarifado.iCodigo = objCodigoNome.iCodigo
                End If
            Next
                
            If objAlmoxarifado.iCodigo = 0 Then

                'Lê o Almoxarifado
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> 25060 And lErro <> SUCESSO Then gError 46979
                
                'Se não encontrou o Almoxarifado --> erro
                If lErro = 25060 Then gError 46980
                            
                colCodigoNome.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido
            
            End If
            
            'Recolhe a alocação
            colAlocacoes.Add objAlmoxarifado.iCodigo, GridAlocacao.TextMatrix(iIndice, iGrid_AlmoxAloc_Col), StrParaDbl(GridAlocacao.TextMatrix(iIndice, iGrid_QuantAloc_Col))

        End If
    Next

    Recolhe_Alocacoes = SUCESSO

    Exit Function

Erro_Recolhe_Alocacoes:

    Recolhe_Alocacoes = gErr

    Select Case gErr

        Case 46979

        Case 46980
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157366)

    End Select

    Exit Function

End Function

Sub Remove_Alocacoes(iItemAtual As Integer)
'Remove do grid de alocações as alocações do item passado

Dim iIndice As Integer
Dim iItemLinha As Integer
Dim bExcluiuLinha As Boolean

    iIndice = 1
    bExcluiuLinha = False
    
    Do While iIndice <= objGridAlocacoes.iLinhasExistentes
        If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))) > 0 Then
            'Recolhe o item a qual pertence a alocação
            iItemLinha = GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)
            'Se for o item passado em iItemAtual
            If iItemLinha = iItemAtual Then
                'Exclui do grid de alocações as alocações daquele ItemNF
                Do While CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) = iItemAtual
                    Call Grid_Exclui_Linha(objGridAlocacoes, iIndice)
                    bExcluiuLinha = True
                    If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))) = 0 Then Exit Do
                Loop
            End If
        End If
        iIndice = iIndice + 1
    Loop

    If bExcluiuLinha Then
        'Rastreamento
        Call gobjRastreamento.Exclusao_Item_Grid(iItemAtual)
    End If

End Sub

Private Function Saida_Celula_GridParcelas(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridParcelas

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col
        'Data de Vencimento
        Case iGrid_Vencimento_Col
            lErro = Saida_Celula_DataVencimento(objGridInt)
            If lErro <> SUCESSO Then gError 46907
        'VAlor da Parcela
        Case iGrid_ValorParcela_Col
            lErro = Saida_Celula_ValorParcela(objGridInt)
            If lErro <> SUCESSO Then gError 46908
        Case iGrid_TipoPagto_Col
            lErro = Saida_Celula_TipoPagto(objGridInt)
            If lErro <> SUCESSO Then gError 183004
        'Descontos da PArcela
        Case iGrid_Desc1Codigo_Col, iGrid_Desc2Codigo_Col, iGrid_Desc3Codigo_Col
            lErro = Saida_Celula_TipoDesconto(objGridInt)
            If lErro <> SUCESSO Then gError 46909
        'Datas de desconto da Parcela
        Case iGrid_Desc1Ate_Col, iGrid_Desc2Ate_Col, iGrid_Desc3Ate_Col
            lErro = Saida_Celula_DescontoData(objGridInt)
            If lErro <> SUCESSO Then gError 46911
        'VAlores dos descontos da parcela
        Case iGrid_Desc1Valor_Col, iGrid_Desc2Valor_Col, iGrid_Desc3Valor_Col
            lErro = Saida_Celula_DescontoValor(objGridInt)
            If lErro <> SUCESSO Then gError 46912
        'Percentuais de desconto da parcela.
        Case iGrid_Desc1Percentual_Col, iGrid_Desc2Percentual_Col, iGrid_Desc3Percentual_Col
            lErro = Saida_Celula_DescontoPerc(objGridInt)
            If lErro <> SUCESSO Then gError 46910
            
    End Select
    
    Saida_Celula_GridParcelas = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_GridParcelas:

    Saida_Celula_GridParcelas = gErr
    
    Select Case gErr
    
        Case 46907, 46908, 46909, 46910, 46911, 46912, 183004
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157367)
            
    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_DataVencimento(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Vencimento que está deixando de ser a corrente

Dim lErro As Long
Dim dtDataReferencia As Date
Dim dtDataVencimento As Date
Dim sDataVencimento As String
Dim iCriouLinha As Integer
Dim objParcRec As ClassParcelaReceber

On Error GoTo Erro_Saida_Celula_DataVencimento

    Set objGridInt.objControle = DataVencimento

    iCriouLinha = False
    
    'Verifica se Data de Vencimento esta preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then gError 46913
        
         dtDataVencimento = CDate(DataVencimento.Text)
         
        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataReferencia.ClipText)) > 0 Then
            dtDataReferencia = StrParaDate(DataReferencia.Text)
            If dtDataVencimento < dtDataReferencia Then gError 46915
        End If
        
        sDataVencimento = Format(dtDataVencimento, "dd/mm/yyyy")
        
        iCriouLinha = False
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            iCriouLinha = True
            
            Set objParcRec = New ClassParcelaReceber
            gobjNFiscal.ColParcelaReceber.AddObj objParcRec
            
        End If
        
    End If
    
    If sDataVencimento <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col) Then CobrancaAutomatica.Value = vbUnchecked
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 46914

    If iCriouLinha Then
        'Coloca desconto padrao (le em CPRConfig)
        lErro = Preenche_DescontoPadrao(GridParcelas.Row)
        If lErro <> SUCESSO Then gError 51499
    End If
    
    Saida_Celula_DataVencimento = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_DataVencimento:

    Saida_Celula_DataVencimento = gErr
    
    Select Case gErr
    
        Case 46913, 46914, 51499
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 46915
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR_REFERENCIA", gErr, dtDataVencimento, GridParcelas.Row, dtDataReferencia)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157368)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
    End Select
    
    Exit Function

End Function

Private Function Saida_Celula_ValorParcela(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor da Parcela que está deixando de sser a corrente

Dim lErro As Long
Dim dColunaSoma As Double
Dim iIndice As Integer
Dim iColDescPerc As Integer
Dim iColTipoDesconto As Integer
Dim lTamanho As Long
Dim dPercentual As Double
Dim dValorParcela As Double
Dim sValorDesconto As String
Dim iTipoDesconto As Integer
Dim objParcRec As ClassParcelaReceber

On Error GoTo Erro_Saida_Celula_ValorParcela

    Set objGridInt.objControle = ValorParcela

    'Verifica se valor está preenchido
    If Len(ValorParcela.ClipText) > 0 Then
        
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(ValorParcela.Text)
        If lErro <> SUCESSO Then gError 46916
        
        ValorParcela.Text = Format(ValorParcela.Text, "Standard")
        If ValorParcela.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) Then
            
            CobrancaAutomatica.Value = vbUnchecked
            
            '***Código para colocar valores de desconto
            dValorParcela = StrParaDbl(ValorParcela.Text)
            If dValorParcela > 0 Then
            
                'Vai varrer todos os 3 descontos para colocar valores
                For iIndice = 1 To 3
                    
                    Select Case iIndice
                        Case 1
                            iColDescPerc = iGrid_Desc1Percentual_Col
                            iColTipoDesconto = iGrid_Desc1Codigo_Col
                        Case 2
                            iColDescPerc = iGrid_Desc2Percentual_Col
                            iColTipoDesconto = iGrid_Desc2Codigo_Col
                        Case 3
                            iColDescPerc = iGrid_Desc3Percentual_Col
                            iColTipoDesconto = iGrid_Desc3Codigo_Col
                    End Select
                    
                    iTipoDesconto = Codigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iColTipoDesconto))
                    lTamanho = Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc)))

                    'Coloca valor de desconto na tela
                    If (iTipoDesconto = Percentual Or iTipoDesconto = PERC_ANT_DIA Or iTipoDesconto = PERC_ANT_DIA_UTIL) And lTamanho > 0 Then
                        dPercentual = PercentParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc))
                        sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
                        GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc - 1) = sValorDesconto
                    End If
                    
                Next
                
            End If
            '***Fim Código para colocar valores de desconto
        End If
        
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            'Coloca DescontoPadrao
            lErro = Preenche_DescontoPadrao(GridParcelas.Row)
            If lErro <> SUCESSO Then gError 51498
            
            Set objParcRec = New ClassParcelaReceber
            gobjNFiscal.ColParcelaReceber.AddObj objParcRec
            
        End If
    Else
        '***Código para colocar valores de desconto
        'Limpa Valores de Desconto
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc1Valor_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc2Valor_Col) = ""
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc3Valor_Col) = ""
        '***Fim Código para colocar valores de desconto
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 46917
                    
    Saida_Celula_ValorParcela = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_ValorParcela:

    Saida_Celula_ValorParcela = gErr
    
    Select Case gErr
    
        Case 46916, 46917, 51498
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157369)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoDesconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo de Desconto que está deixando de sser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim iTipo As Integer
Dim iIndice As Integer
Dim objParcRec As ClassParcelaReceber

On Error GoTo Erro_Saida_Celula_TipoDesconto

    If GridParcelas.Col = iGrid_Desc1Codigo_Col Then
        Set objGridInt.objControle = Desconto1Codigo
    ElseIf GridParcelas.Col = iGrid_Desc2Codigo_Col Then
        Set objGridInt.objControle = Desconto2Codigo
    ElseIf GridParcelas.Col = iGrid_Desc3Codigo_Col Then
        Set objGridInt.objControle = Desconto3Codigo
    End If

    'Verifica se o Tipo foi preenchido
    If Len(Trim(objGridInt.objControle.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If objGridInt.objControle.Text <> objGridInt.objControle.List(objGridInt.objControle.ListIndex) Then

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona_Grid(objGridInt.objControle, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 46918
        
            'Não foi encontrado
            If lErro = 25085 Then gError 46920
            If lErro = 25086 Then gError 46921

        End If

        iTipo = Codigo_Extrai(objGridInt.objControle.Text)
        
        If (iTipo = VALOR_ANT_DIA) Or (iTipo = VALOR_ANT_DIA_UTIL) Or (iTipo = VALOR_FIXO) Then
            GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 3) = ""
        ElseIf iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual Then
            '*** Acrescentado + 1 If para contabilizar com colocação de valores de desconto
            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 3))) = 0 Then
                GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 2) = ""
            End If
        End If
        
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
            Set objParcRec = New ClassParcelaReceber
            gobjNFiscal.ColParcelaReceber.AddObj objParcRec
        
        End If

    Else
        For iIndice = objGridInt.objGrid.Col To iGrid_Desc3Percentual_Col
            GridParcelas.TextMatrix(GridParcelas.Row, iIndice) = ""
        Next
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 46919

    Saida_Celula_TipoDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDesconto:

    Saida_Celula_TipoDesconto = gErr

    Select Case gErr

        Case 46918, 46919
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 46920
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 46921
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", gErr, objGridInt.objControle.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157370)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoData(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Data que está deixando de sser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim dtDataVencimento As Date

On Error GoTo Erro_Saida_Celula_DescontoData

    If GridParcelas.Col = iGrid_Desc1Ate_Col Then
        Set objGridInt.objControle = Desconto1Ate
    ElseIf GridParcelas.Col = iGrid_Desc2Ate_Col Then
        Set objGridInt.objControle = Desconto2Ate
    ElseIf GridParcelas.Col = iGrid_Desc3Ate_Col Then
        Set objGridInt.objControle = Desconto3Ate
    End If

    If Len(Trim(objGridInt.objControle.ClipText)) > 0 Then

        lErro = Data_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 46922
        'Se a data de vencimento estiver preenchida
        If Len(Trim(DataEmissao.ClipText)) = 0 Then
            'critica se DataDesconto ultrapassa DataVencimento
            If StrParaDate(objGridInt.objControle.Text) < StrParaDate(DataEmissao.ClipText) Then gError 46924
        End If

            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))) > 0 Then
            dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))
            If StrParaDate(objGridInt.objControle) > dtDataVencimento Then gError 46925
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 46923

    Saida_Celula_DescontoData = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoData:

    Saida_Celula_DescontoData = gErr

    Select Case gErr

        Case 46922, 46923
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 46924
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_INFERIOR_DATA_EMISSAO", gErr, CDate(objGridInt.objControle.Text))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 46925
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_SUPERIOR_DATA_VENCIMENTO", gErr, CDate(objGridInt.objControle.Text), dtDataVencimento)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157371)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoValor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Valor que está deixando de sser a corrente

Dim lErro As Long
Dim dColunaSoma As Double
Dim objParcRec As ClassParcelaReceber

On Error GoTo Erro_Saida_Celula_DescontoValor

    If GridParcelas.Col = iGrid_Desc1Valor_Col Then
        Set objGridInt.objControle = Desconto1Valor
    ElseIf GridParcelas.Col = iGrid_Desc2Valor_Col Then
        Set objGridInt.objControle = Desconto2Valor
    ElseIf GridParcelas.Col = iGrid_Desc3Valor_Col Then
        Set objGridInt.objControle = Desconto3Valor
    End If

    'Verifica se valor está preenchido
    If Len(objGridInt.objControle.ClipText) > 0 Then
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 46926
        
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
            Set objParcRec = New ClassParcelaReceber
            gobjNFiscal.ColParcelaReceber.AddObj objParcRec
        
        End If
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 46927
                    
    Saida_Celula_DescontoValor = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_DescontoValor:

    Saida_Celula_DescontoValor = gErr
    
    Select Case gErr
    
        Case 46926, 46927
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157372)
        
    End Select

    Exit Function

End Function


Private Function Saida_Celula_DescontoPerc(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Percentual que está deixando de sser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim dPercentual As Double
Dim dValorParcela As Double
Dim sValorDesconto As String

On Error GoTo Erro_Saida_Celula_DescontoPerc

    If GridParcelas.Col = iGrid_Desc1Percentual_Col Then
        Set objGridInt.objControle = Desconto1Percentual
    ElseIf GridParcelas.Col = iGrid_Desc2Percentual_Col Then
        Set objGridInt.objControle = Desconto2Percentual
    ElseIf GridParcelas.Col = iGrid_Desc3Percentual_Col Then
        Set objGridInt.objControle = Desconto3Percentual
    End If

    'Se a Porcentagem estiver preenchida
    If Len(Trim(objGridInt.objControle.Text)) > 0 Then
        'Critica o percentual informado
        lErro = Porcentagem_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 46928
    
        '***Código para colocar valores de desconto
        dPercentual = CDbl(objGridInt.objControle.Text) / 100
        dValorParcela = StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col))
    
        'Coloca Valor do Desconto na tela
        If dValorParcela > 0 Then
            sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
            GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col - 1) = sValorDesconto
        End If
    
    Else
    
        'Limpa Valor de Desconto
        GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col - 1) = ""
        '***Fim Código para colocar valores de desconto

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 46929

    Saida_Celula_DescontoPerc = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoPerc:

    Saida_Celula_DescontoPerc = gErr

    Select Case gErr

        Case 46928, 46929
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157373)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col
        'Quantidade
        Case iGrid_Quantidade_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 46930
        'Preco Unitário
        Case iGrid_ValorUnitario_Col
            lErro = Saida_Celula_PrecoUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 46931
        'Percentual de Desconto
        Case iGrid_PercDesc_Col
            lErro = Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then gError 46932
        'VAlor de desconto
        Case iGrid_Desconto_Col
            lErro = Saida_Celula_Desconto(objGridInt)
            If lErro <> SUCESSO Then gError 46933
        
        Case iGrid_UnidadeMed_Col
            lErro = Saida_Celula_UM(objGridInt)
            If lErro <> SUCESSO Then gError 131916

        Case iGrid_TabPrecoItem_Col
            If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
                lErro = Saida_Celula_TabPrecoItemNF(objGridInt)
                If lErro <> SUCESSO Then gError 131916
            End If

        Case iGrid_ComissaoItem_Col
            If gobjFAT.iExibeComissaoItensPV = MARCADO Then
                lErro = Saida_Celula_ComissaoItemNF(objGridInt)
                If lErro <> SUCESSO Then gError 131916
            End If

        Case Else
            lErro = CF2(Me, "NFFP_Saida_Celula_GridItens", objGridInt)
            If lErro <> SUCESSO Then gError 46933
    
    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 46930, 46931, 46932, 46933, 131916 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157374)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim iSairCelula As Integer
Dim dPrecoUnitario As Double
Dim dQuantidadeAnterior As Double
Dim sUM As String
Dim dFator As Double
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objKitVenda As New ClassKitVenda
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objItemNF As New ClassItemNF
Dim objItemPV As ClassItemPedido
Dim dQtdePV As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    iSairCelula = False
    
    'Verifica o preenchimento de Quantidade
    If Len(Trim(Quantidade.ClipText)) > 0 Then
        'VAlida a quantidae informada
        If giComplemento = DESMARCADO Then
            lErro = Valor_Positivo_Critica(Quantidade.Text)
        Else
            lErro = Valor_NaoNegativo_Critica(Quantidade.Text)
        End If
        If lErro <> SUCESSO Then gError 46936
        'Recolhe a quantidade informada
        dQuantidade = CDbl(Quantidade.Text)
        Quantidade = Formata_Estoque(dQuantidade)
        'Recolhe a quantidade anteriormente digitada
        dQuantidadeAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
        'Verifica se a quantidade foi alterada
        If dQuantidadeAnterior = dQuantidade Then iSairCelula = True
        'Se a quantidade foi alterada verifica se é um valor que pode ser faturado
        
        sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
    
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 132048
    
        objProduto.sCodigo = sProdutoFormatado
    
        'Lê o produto da linha passada por iLinha do GridItens
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 132049
        
        If lErro = 28030 Then gError 132050
    
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, gcolItemUM(GridItens.Row), sUM, dFator)
        If lErro <> SUCESSO Then gError 132051
        
        If Not iSairCelula Then
        
            Set objItemNF = gobjNFiscal.ColItensNF(GridItens.Row)
            
            If objItemNF.iRecebForaFaixa = ERRO_E_REJEITA_RECEBIMENTO Then
                Set objItemPV = New ClassItemPedido
                
                objItemPV.lNumIntDoc = gcolItemPedido(GridItens.Row)
                objItemPV.iFilialEmpresa = Codigo_Extrai(FilialPedido.Text)
                
                lErro = CF("ItensPedido_Le_NumIntDoc1", objItemPV)
                If lErro <> SUCESSO Then gError 196962
                
                dQtdePV = objItemPV.dQuantidade - objItemPV.dQuantCancelada
            
                If dQuantidade > Arredonda_Estoque(gcolItemQuant(GridItens.Row) + (dQtdePV * objItemNF.dPercentMaisReceb) * dFator) Then gError 46934  'Alterado por Wagner
            End If
            
        End If
        
    End If
    'Se a quantidade foi alterada
    If Not iSairCelula Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)
        
        objProduto.sCodigo = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177585
        
        If objProduto.iKitVendaComp <> MARCADO Then
            'Tenta fazer uma alocação p\ oitem de acordo com a nova UM
            lErro = CF2(Me, "NFFP_Alocacao_Processa")
            If lErro <> SUCESSO And lErro <> 23796 And lErro <> 197701 And lErro <> 197703 Then gError 46937 '??? Jones 09/02/05 passou a tratar 23796
            If lErro = 23796 Or lErro = 197701 Or lErro = 197703 Then gError 23796  '??? Jones 09/02/05 passou a tratar 23796
    
        Else
        
            objKitVenda.sProduto = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
            objKitVenda.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
            objKitVenda.dQuantidade = dQuantidade
            
            gobjNFiscal.ColItensNF.Item(GridItens.Row).sUnidadeMed = objKitVenda.sUM
            gobjNFiscal.ColItensNF.Item(GridItens.Row).dQuantidade = dQuantidade
        
            lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade)
            If lErro <> SUCESSO Then gError 177586
                        
            For Each objItemRomaneio In gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade
                lErro = Alocacao_Processa_ItemGrade(objItemRomaneio, gobjNFiscal.ColItensNF(GridItens.Row), GridItens.Row)
                If lErro <> SUCESSO Then gError 177587
            Next
            
            Call Atualiza_Grid_Alocacao(gobjNFiscal.ColItensNF.Item(GridItens.Row))

        End If
    
        lErro = CF2(Me, "NFFP_Calcula_PesoKg", GridItens.Row)
        If lErro <> SUCESSO Then gError 177066
    
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81510
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 46938
    
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))

    If dPrecoUnitario > 0 Then Call PrecoTotal_Calcula(GridItens.Row, False)
    
    '*** incluidos p/tratamento de tributacao ****************
    If GridItens.Row <= objGridItens.iLinhasExistentes Then
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 46935
    End If
    '*** fim tributacao

    'Atualiza o valor total
    Call ValorTotal_Calcula

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 23796 '??? Jones 09/02/05 passou a tratar 23796
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO2", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 46934
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_FATURAR_MENOR", gErr, Arredonda_Estoque(gcolItemQuant(GridItens.Row) + (dQtdePV * objItemNF.dPercentMaisReceb) * dFator))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 46935, 46936, 46937, 46938, 81510
            GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidadeAnterior)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 132048 To 132051, 177066, 177585, 177586, 177587, 196962
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157375)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoUnitario(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double
Dim dQuantCancelada As Double

On Error GoTo Erro_Saida_Celula_PrecoUnitario

    Set objGridInt.objControle = PrecoUnitario

    'Verifica o preenchimento de Preço Unitário
    If Len(Trim(PrecoUnitario.Text)) > 0 Then

        If giComplemento = MARCADO Then
            lErro = Valor_NaoNegativo_Critica(PrecoUnitario.Text)
        Else
            lErro = Valor_Positivo_Critica(PrecoUnitario.Text)
        End If
        If lErro <> SUCESSO Then gError 46931

        PrecoUnitario.Text = Format(PrecoUnitario.Text, gobjFAT.sFormatoPrecoUnitario)
    
        lErro = CF2(Me, "NFFP_Saida_Celula_PrecoUnitario", StrParaDbl(PrecoUnitario.Text), GridItens.Row)
        If lErro <> SUCESSO Then gError 199800
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 46932
    
    Call PrecoTotal_Calcula(GridItens.Row, False)

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 46930
'*** fim tributacao

    'Atualiza o valor total
    Call ValorTotal_Calcula

    Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = gErr

    Select Case gErr

        Case 46930, 46931, 46932, 199800
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157376)

    End Select

    Exit Function

End Function

Function Saida_Celula_PercentDesc(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual de Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim dPercentDescAnterior As Double

On Error GoTo Erro_Saida_Celula_PercentDesc

    Set objGridInt.objControle = PercentDesc
    
    dPercentDesc = StrParaDbl(PercentDesc.Text)
    
    dPercentDescAnterior = PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col)) * 100
    
    'verifica se o percentual está preenchido
    If Len(Trim(PercentDesc.Text)) > 0 Then
    
        'Critica a procentagem
        lErro = Porcentagem_Critica(PercentDesc.Text)
        If lErro <> SUCESSO Then gError 46933
               
        If dPercentDesc <> dPercentDescAnterior Then
        
            'Verifica se o percentual é de 100%
            If dPercentDesc = 100 Then gError 46935
    
            PercentDesc.Text = Format(dPercentDesc, "Fixed")
    
        End If
    Else
        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
    End If

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) > 0 Then dPrecoUnitario = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))

    If Abs(dPercentDesc - dPercentDescAnterior) > DELTA_VALORMONETARIO2 Then
        iDescontoAlterado = REGISTRO_ALTERADO
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 46934

    If dPrecoUnitario > 0 Then
    
        Call PrecoTotal_Calcula(GridItens.Row, False)
    
''*** incluidos p/tratamento de tributacao *******************************
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 39039
'*** fim tributacao
    
        Call ValorTotal_Calcula

    End If

    Saida_Celula_PercentDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDesc:

    Saida_Celula_PercentDesc = gErr

    Select Case gErr

        Case 46933, 46934
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 46935
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157377)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dPrecoTotal As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim iDescontoAlterado As Integer

On Error GoTo Erro_Saida_Celula_Desconto

    Set objGridInt.objControle = Desconto

    iDescontoAlterado = False

    'Verifica o preenchimento de Desconto
    If Len(Trim(Desconto.ClipText)) > 0 Then
        'Faz a critica do Desconot
        lErro = Valor_NaoNegativo_Critica(Desconto.Text)
        If lErro <> SUCESSO Then gError 46937
        'recolhe o valor do desconto informado
        dDesconto = CDbl(Desconto.Text)
    End If
    
    'Verifica se o valor de desconto foi alterado
    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col)) <> dDesconto Then iDescontoAlterado = True
    'Se o desconto foi alterado
    If iDescontoAlterado = True Then
        iDescontoAlterado = REGISTRO_ALTERADO
        'recolhe a quantidade, preço unitário  e calcula o preço total
        dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
        dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))
        dPrecoTotal = dQuantidade * dPrecoUnitario
        'Se o preço total for positivo
        If dPrecoTotal > 0 Then
            'Verifica se o desconto ultrapassa o preço total
            If dDesconto >= dPrecoTotal Then gError 46938
            'Calcula p percentual de desconto em cima do novo desconto
            dPercentDesc = dDesconto / dPrecoTotal
            'Coloca o Percentual de desconto na tela
            GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")

        End If
    End If
    
    If dDesconto <> StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col)) Then iDescontoAlterado = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 35834

    If iDescontoAlterado = True Then

        Call PrecoTotal_Calcula(GridItens.Row, False)
        
        ''*** incluidos p/tratamento de tributacao *******************************
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 46936
        '*** fim tributacao
        
        Call ValorTotal_Calcula
    
    End If
    
    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = gErr

    Select Case gErr

        Case 46936, 46937
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 46938
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCONTO_MAIOR_OU_IGUAL_PRECO_TOTAL", gErr, GridItens.Row, dDesconto, dPrecoTotal)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157378)

    End Select

    Exit Function

End Function

Function NFFP_Alocacao_Processa(ByVal objCT As Object) As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dQuantidade As Double
Dim sUM As String

On Error GoTo Erro_NFFP_Alocacao_Processa

    'Remove as alocações do grid de alocações
    Call Remove_Alocacoes(GridItens.Row)
    
    'Formata o Produto
    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 46949

    objProduto.sCodigo = sProduto
    
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 46950
    If lErro = 28030 Then gError 46952 'Não encontrou
    
    'recolhe a quantidade do grid de itens
    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    
    'Verifica se o produto tem o controle de estoque <> PRODUTO_CONTROLE_SEM_ESTOQUE
    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 And gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA And Codigo_Extrai(TipoNFiscal.Text) <> DOCINFO_NFIFVETPV Then
        'recolhe a UM do ItemNF
        sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
        
        'Tenta Alocar o produto no Almoxarifado padrão
        lErro = CF2(Me, "NFFP_AlocaAlmoxarifadoPradrao", dQuantidade, objProduto, sUM)
        If lErro <> SUCESSO And lErro <> 23796 Then gError 46951 '??? Jones 09/02/05 passou a tratar 23796
        If lErro = 23796 Then gError 23796 '??? Jones 09/02/05 passou a tratar 23796
    
    End If

    NFFP_Alocacao_Processa = SUCESSO

    Exit Function

Erro_NFFP_Alocacao_Processa:

    NFFP_Alocacao_Processa = gErr

    Select Case gErr

        Case 46949, 46950, 46951, 23796 '??? Jones 09/02/05 passou a tratar 23796

        Case 46952
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157379)

    End Select

    Exit Function

End Function

Function NFFP_AlocaAlmoxarifadoPradrao(ByVal objCT As Object, dQuantidade As Double, objProduto As ClassProduto, sUM As String) As Long
'Tenta fazer a alocação do produto no almoxarifado padrão. Caso não consiga chama a tela de Alocação de produto.

Dim lErro As Long
Dim dQuantAlocar As Double
Dim dFator As Double
Dim iAlmoxarifado As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim dQuantDisponivel As Double
Dim objItemNF As ClassItemNF
Dim colOutrosProdutos As New Collection
Dim sProduto As String
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim objItemNFAloc As ClassItemNFAlocacao
Dim sProdutoEnxuto As String
Dim iLinhas As Integer
Dim dAcrescimo As Double
Dim iNumCasasDec As Integer
Dim objItemPV As New ClassItemPedido
Dim colReservaBD As New colReservaItem
Dim dTotal As Double

On Error GoTo Erro_NFFP_AlocaAlmoxarifadoPradrao

    'Faz a conversão da UM da tela para a UM de estoque
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 46954

    'Converte a quantidade para a UM de estoque
    dQuantAlocar = Arredonda_Estoque(dQuantidade * dFator)

    'Busca o Almoxarifado padrão
    lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 46956
    If lErro = 23796 Then gError 23796
    
    'Se encontrou
    If iAlmoxarifado > 0 Then

        objAlmoxarifado.iCodigo = iAlmoxarifado
        'Lê o Aloxarifado
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> 25056 And lErro <> SUCESSO Then gError 46957
        If lErro = 25056 Then gError 46960
        
        objEstoqueProduto.sProduto = objProduto.sCodigo
        objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
        
        'Le os estoques desse produto nesse almoxarifado
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 46955
        If lErro = 21306 Then gError 46961 'Não encontrou

        objItemPV.lNumIntDoc = gcolItemPedido(GridItens.Row)
        objItemPV.sProduto = objProduto.sCodigo

        lErro = CF("ReservasItemPV_Le_NumIntOrigem", objItemPV, colReservaBD)
        If lErro <> SUCESSO And lErro <> 51601 Then gError 62095
        
        For iIndice = 1 To colReservaBD.Count
            If objEstoqueProduto.iAlmoxarifado = colReservaBD(iIndice).iAlmoxarifado Then
                objEstoqueProduto.dQuantDispNossa = objEstoqueProduto.dQuantDispNossa + colReservaBD(iIndice).dQuantidade
                Exit For
            End If
        Next
        
        iCodigo = Codigo_Extrai(TipoNFiscal.Text)
    
        'Seleciona a origem da quantidade disponível
        Select Case iCodigo
            
            'Se o tipo da nota for mat. beneficiado de 3º´s
            Case DOCINFO_NFISFBFPV
            
                'A quantidade disponível deve ser igual a quantidade no escaninho Mat.de 3º´s em Beneficiamento (Benef3)
                dQuantDisponivel = objEstoqueProduto.dQuantBenef3
                
            'Se for outro tipo de nota
            Case Else
                
                'A quantidade disponível deve ser igual a quantidade do escaninho mat. nosso disponível (DispNossa)
                dQuantDisponivel = objEstoqueProduto.dQuantDisponivel
        
        End Select
        
        'dQuantDisponivel = objEstoqueProduto.dQuantDisponivel
        'Remove as alocações do produto do grid de alocações
        Call Remove_Alocacoes(GridItens.Row)

        iLinhas = objGridAlocacoes.iLinhasExistentes
        
        dQuantAlocar = Arredonda_Estoque(dQuantAlocar)

        'Verifica se a Quantidade disponível é maior que a quantidade a alocar
        If (dQuantAlocar - dQuantDisponivel) < QTDE_ESTOQUE_DELTA Or gobjMAT.iAceitaEstoqueNegativo = MARCADO Then

            iLinhas = iLinhas + 1
            'Coloca no Grid alocações a alocação para o almoxarifado padrão
            GridAlocacao.TextMatrix(iLinhas, iGrid_Item_Col) = GridItens.Row
            GridAlocacao.TextMatrix(iLinhas, iGrid_ProdutoAloc_Col) = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
            GridAlocacao.TextMatrix(iLinhas, iGrid_AlmoxAloc_Col) = objAlmoxarifado.sNomeReduzido
            GridAlocacao.TextMatrix(iLinhas, iGrid_QuantAloc_Col) = Formata_Estoque(dQuantAlocar)
            GridAlocacao.TextMatrix(iLinhas, iGrid_QuantVend_Col) = Formata_Estoque(dQuantAlocar)
            GridAlocacao.TextMatrix(iLinhas, iGrid_UMAloc_Col) = objProduto.sSiglaUMEstoque
        
        'Se não for
        Else
            'Se a função não estiver sendo usada pelo carregamento da tela
            'se prepara para chamar a tela de localização
            If Not gbCarregandoTela Then
                
                Set objItemNF = New ClassItemNF
                'Recolhe os dados do item
                objItemNF.iItem = GridItens.Row
                objItemNF.sProduto = objProduto.sCodigo
                objItemNF.sDescricaoItem = objProduto.sDescricao
                objItemNF.dQuantidade = dQuantidade
                objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
                objItemNF.lNumIntItemPedVenda = gcolItemPedido(GridItens.Row)
    
                'Recolhe todos os produtos dos outros itens
                For iIndice = 1 To objGridItens.iLinhasExistentes
                    If iIndice <> GridItens.Row Then
                        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
                        If lErro <> SUCESSO Then gError 46958
                        'Adiciona na coleção de produtos
                        colOutrosProdutos.Add sProduto
                    End If
                Next
    
                'Chama a tela de Localização de Produto
                Call Chama_Tela_Modal("LocalizacaoProduto1", objItemNF, colOutrosProdutos, dQuantAlocar, DOCINFO_NFISFVPV)
                If giRetornoTela = vbCancel Then gError 46963 'Se nada foi feito lá
                If giRetornoTela = vbOK Then
    
                    'Se o produto foi substituido
                    If objProduto.sCodigo <> objItemNF.sProduto Then gError 46962
                                    
                    sProdutoEnxuto = String(STRING_PRODUTO, 0)
    
                    lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
                    If lErro Then gError 46959
    
                    Produto.PromptInclude = False
                    Produto.Text = sProdutoEnxuto
                    Produto.PromptInclude = True
                    
                    'Calcula o Total Alocado
                    For Each objItemNFAloc In objItemNF.colAlocacoes
                        dTotal = dTotal + objItemNFAloc.dQuantidade
                    Next
                    
                    'Para cada alocação feita para o item
                    For Each objItemNFAloc In objItemNF.colAlocacoes
    
                        iLinhas = iLinhas + 1
                        'coloca no grid de alocações os dados das alocações
                        GridAlocacao.TextMatrix(iLinhas, iGrid_Item_Col) = objItemNF.iItem
                        GridAlocacao.TextMatrix(iLinhas, iGrid_ProdutoAloc_Col) = Produto.Text
                        GridAlocacao.TextMatrix(iLinhas, iGrid_AlmoxAloc_Col) = objItemNFAloc.sAlmoxarifado
                        GridAlocacao.TextMatrix(iLinhas, iGrid_QuantAloc_Col) = Formata_Estoque(objItemNFAloc.dQuantidade)
                        GridAlocacao.TextMatrix(iLinhas, iGrid_QuantVend_Col) = Formata_Estoque(dTotal)
                        GridAlocacao.TextMatrix(iLinhas, iGrid_UMAloc_Col) = objItemNF.sUMEstoque
    
                    Next
                
                    'Quantidade.Text = Formata_Estoque(dTotal)
                    If Abs(dQuantAlocar - dTotal) > QTDE_ESTOQUE_DELTA Then
                        Quantidade.Text = Formata_Estoque(dTotal / dFator)
                    Else
                        Quantidade.Text = Formata_Estoque(StrParaDbl(Quantidade.Text))
                    End If
    
                End If
            End If
        End If
        
        'atualiza o numero de linhas existentes no grid de alocações
        objGridAlocacoes.iLinhasExistentes = iLinhas
    End If

    NFFP_AlocaAlmoxarifadoPradrao = SUCESSO

    Exit Function

Erro_NFFP_AlocaAlmoxarifadoPradrao:

    NFFP_AlocaAlmoxarifadoPradrao = gErr

    Select Case gErr

        Case 46954, 46955, 46956, 46957, 46958, 46959, 62095, 23796

        Case 46960
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, iAlmoxarifado)

        Case 46961
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case 46962
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_SUBSTITUIDO", gErr)

        Case 46963
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_LOCALIZACAO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157380)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim iTipo As Integer
Dim sUnidadeMed As String
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 46964
    
    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
    
        'Produto
        Case Produto.Name ', UnidadeMed.Name -- Alterado por wagner
        
        Case UnidadeMed.Name

            UnidadeMed.Clear

            'Guarda a UM que está no Grid
            sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)

            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 131917

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                UnidadeMed.Enabled = False
            Else
                UnidadeMed.Enabled = True

                objProduto.sCodigo = sProdutoFormatado
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 131918
                If lErro = 28030 Then gError 131919 'Não achou

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lâ as Unidades de Medidas da Classe do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 131920
                'Carrega a combo de UM
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next
                'Seleciona na UM que está preenchida
                UnidadeMed.Text = sUM
                If Len(Trim(sUM)) > 0 Then
                    lErro = Combo_Item_Igual(UnidadeMed)
                    If lErro <> SUCESSO And lErro <> 12253 Then gError 131921
                End If
            End If
        
        'Nas demais
        Case TabPrecoItemNF.Name, ComissaoItemNF.Name, PercentDesc.Name, Desconto.Name
            If iProdutoPreenchido = PRODUTO_VAZIO Or gcolItemPedido.Count = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
        Case Quantidade.Name

            If iProdutoPreenchido = PRODUTO_VAZIO Or gcolItemPedido.Count = 0 Or giFaturaIntegral <> 0 Or left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case Desconto1Ate.Name, Desconto1Valor.Name, Desconto1Percentual.Name
            'Habilita os campos de desconto em sequencia
            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc1Codigo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc1Codigo_Col))
                If objControl.Name = Desconto1Ate.Name Then
                    objControl.Enabled = True
                ElseIf objControl.Name = Desconto1Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
                    Desconto1Valor.Enabled = True
                ElseIf objControl.Name = Desconto1Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
                    Desconto1Percentual.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If
        
        Case Desconto2Ate.Name, Desconto2Valor.Name, Desconto2Percentual.Name
            'Habilita os campos de desconto em sequencia
            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc2Codigo_Col))
            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc2Codigo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                If objControl.Name = Desconto2Ate.Name Then
                    objControl.Enabled = True
                ElseIf objControl.Name = Desconto2Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
                    Desconto2Valor.Enabled = True
                ElseIf objControl.Name = Desconto2Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
                    Desconto2Percentual.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If
        
        Case Desconto3Ate.Name, Desconto3Valor.Name, Desconto3Percentual.Name
            'Habilita os campos de desconto em sequencia
            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc3Codigo_Col))
            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc3Codigo_Col))) = 0 Then
                objControl.Enabled = False
            Else
                If objControl.Name = Desconto3Ate.Name Then
                    objControl.Enabled = True
                ElseIf objControl.Name = Desconto3Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
                    Desconto3Valor.Enabled = True
                ElseIf objControl.Name = Desconto3Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
                    Desconto3Percentual.Enabled = True
                Else
                    objControl.Enabled = False
                End If
            End If
        
        
        Case ValorParcela.Name
            'Se o vencimento estiver preenchido, habilita o controle
            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_Col))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case Desconto2Codigo.Name, Desconto3Codigo.Name
            'Habilita os campos de desconto em sequencia
            If Len(Trim(GridParcelas.TextMatrix(iLinha, GridParcelas.Col - 4))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
    
        Case Else
            lErro = CF2(Me, "NFFP_Rotina_Grid_Enable", iLinha, objControl, iCaminho, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 124178
        
    End Select

    'embalagem
    lErro = gobjEmbalagem.Rotina_Grid_Enable_Emb(iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 83394
    
    'Rastreamento
    lErro = gobjRastreamento.Rotina_Grid_Enable(iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 83394

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 46964, 83394
        
        Case 131917, 131918, 131920, 131921 'Inserido por Wagner
        
        Case 131919
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157381)

    End Select

    Exit Sub

End Sub

Public Sub GridComissoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridComissoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
    End If

End Sub

Public Sub GridComissoes_EnterCell()
    Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
End Sub

Public Sub GridComissoes_GotFocus()
    Call Grid_Recebe_Foco(objGridComissoes)
End Sub


Public Sub GridComissoes_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iItemAtual As Integer
Dim iLinhasExistentesAnterior As Integer

    'Guarda linha atual
    iItemAtual = GridComissoes.Row
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridComissoes)
    
    If KeyCode = vbKeyDelete And iItemAtual <= iLinhasExistentesAnterior Then
        Call objTabComissoes.ExcluiLinhaGrid(iItemAtual)
    End If

End Sub

Public Sub GridComissoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridComissoes, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
    End If

End Sub

Public Sub GridComissoes_LeaveCell()
    Call Saida_Celula(objGridComissoes)
End Sub

Public Sub GridComissoes_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridComissoes)
End Sub

Public Sub GridComissoes_RowColChange()
    Call Grid_RowColChange(objGridComissoes)
End Sub

Public Sub GridComissoes_Scroll()
    Call Grid_Scroll(objGridComissoes)
End Sub

Public Sub GridParcelas_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If
    
End Sub

Public Sub GridParcelas_GotFocus()
    Call Grid_Recebe_Foco(objGridParcelas)
End Sub

Public Sub GridParcelas_EnterCell()
    Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
End Sub

Public Sub GridParcelas_LeaveCell()
    Call Saida_Celula(objGridParcelas)
End Sub

Public Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer

On Error GoTo Erro_GridParcelas_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)
    
    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then

        If gobjNFiscal.ColParcelaReceber.Count >= GridParcelas.Row Then gobjNFiscal.ColParcelaPagar.Remove (GridParcelas.Row)
    
    End If

    Exit Sub
    
Erro_GridParcelas_KeyDown:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183015)
            
    End Select

    Exit Sub
    
End Sub

Public Sub GridParcelas_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Public Sub GridParcelas_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridParcelas)
End Sub

Public Sub GridParcelas_RowColChange()
    Call Grid_RowColChange(objGridParcelas)
End Sub

Public Sub GridParcelas_Scroll()
    Call Grid_Scroll(objGridParcelas)
End Sub

Public Sub DataVencimento_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataVencimento_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub DataVencimento_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub DataVencimento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = DataVencimento
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ValorParcela_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorParcela_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub ValorParcela_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub ValorParcela_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorParcela
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto1Ate_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Desconto1Ate_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub Desconto1Ate_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub Desconto1Ate_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto1Ate
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto1Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Desconto1Codigo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub Desconto1Codigo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub Desconto1Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto1Codigo
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto1Percentual_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Desconto1Percentual_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub Desconto1Percentual_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub Desconto1Percentual_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto1Percentual
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto1Valor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Desconto1Valor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub Desconto1Valor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub Desconto1Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto1Valor
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto2Ate_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Desconto2Ate_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub Desconto2Ate_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub Desconto2Ate_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto2Ate
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto2Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Desconto2Codigo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub Desconto2Codigo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub Desconto2Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto2Codigo
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto2Percentual_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Desconto2Percentual_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub Desconto2Percentual_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub Desconto2Percentual_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto2Percentual
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto2Valor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Desconto2Valor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub Desconto2Valor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub Desconto2Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto2Valor
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto3Ate_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Desconto3Ate_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub Desconto3Ate_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub Desconto3Ate_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto3Ate
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto3Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Desconto3Codigo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub Desconto3Codigo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub Desconto3Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto3Codigo
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Desconto3Percentual_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Desconto3Percentual_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub Desconto3Percentual_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub Desconto3Percentual_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto3Percentual
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto3Valor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Desconto3Valor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub Desconto3Valor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub Desconto3Valor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Desconto3Valor
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub BotaoGravar_Click()

Dim objNFiscal As New ClassNFiscal
Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 46967

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscalFaturaPedido()
    If lErro <> SUCESSO Then gError 500145
    
    'Libera as coleções globais
    Set gcolItemPedido = New Collection
    Set gcolItemQuant = New Collection
    Set gcolItemUM = New Collection 'Inserido por Wagner

    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        'Prepara para chaveamento para NotaFiscalFatura
'        objNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV
'
'        'Faz chaveamento de classe CT para NotaFiscalFatura
'        lErro = objUserControl.CT_Chavea(objNFiscal)
'        If lErro <> SUCESSO Then gError 25944
'
'    End If
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 46967, 25944, 500145 'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 157382)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim vbMsgRes As VbMsgBoxResult
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objBloqueioPV As ClassBloqueioPV
Dim dValorTotal As Double
Dim dValorDesconto As Double
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iNFSE As Integer, iIndice As Integer
Dim colEmbalagensItens As New Collection
Dim objEmbalagensItem As ClassEmbalagensItem

On Error GoTo Erro_Gravar_Registro
   
    'Verifica se os dados de algum pedido foi trazido para a tela
    If glCodPedido = 0 Then gError 46968

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 130474
    If Len(Trim(Pedido.Text)) = 0 Then gError 46969
    If Len(Trim(FilialPedido.Text)) = 0 Then gError 46970
    If Len(Trim(Cliente.Caption)) = 0 Then gError 46971
    If Len(Trim(Serie.Text)) = 0 Then gError 46972
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 46974
    If Len(Trim(gobjTribTab.sNatOpInterna)) = 0 Then gError 46976
    If Len(Trim(FilialEntrega.Text)) = 0 Then gError 51493
    If Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_NFIFVETPV Then
        If Len(Trim(FornEntTerc.Text)) = 0 Then gError 209998
        If Len(Trim(FilialFornEntTerc.Text)) = 0 Then gError 209999
    End If

    'Verifica se tem itens no grid
    If objGridItens.iLinhasExistentes = 0 Then gError 46977

    'Verifica se o código de pedido é o mesmo que foi trazido para a tela
    If glCodPedido <> StrParaLong(Pedido.Text) Or giFilialPedido <> Codigo_Extrai(FilialPedido.Text) Then
        'Avisa que os dados da tela são de outro pedido e pergunta se deseja prosseguir.
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PEDIDO_TELA_NAO_UTILIZADO", glCodPedido, giFilialPedido)
        If vbMsgRes = vbNo Then gError 46978 'Sai da rotina.

        Pedido.Text = glCodPedido
        FilialPedido.Text = giFilialPedido
        Call FilialPedido_Validate(bSGECancelDummy)
    End If
    
    objPedidoVenda.lCodigo = StrParaLong(Pedido.Text)
    objPedidoVenda.iFilialEmpresa = Codigo_Extrai(FilialPedido.Text)
    
    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 46990
    
    If objPedidoVenda.dtDataEmissao > StrParaDate(DataEmissao.Text) Then gError 206615

    If Len(Trim(PesoLiquido.ClipText)) > 0 And Len(Trim(PesoBruto.ClipText)) > 0 Then
        If CDbl(PesoLiquido.Text) > CDbl(PesoBruto.Text) Then gError 46989
    End If

    'Verifica se a Data de Saida e menor que a Data de Emissão
    If Len(Trim(DataSaida.ClipText)) <> 0 Then
        If CDate(DataSaida.Text) < CDate(DataEmissao.Text) Then gError 46986
    End If

    dValorTotal = CDbl(IIf(Len(Trim(ValorTotal.Caption)) > 0, ValorTotal.Caption, 0))

    'Se o total for negativo --> Erro
    If dValorTotal < 0 Then gError 46988

    'Verifica se o Valor do Desconto é maior que a soma dos demais valores
    If Len(Trim(ValorDesconto.Text)) > 0 Then dValorDesconto = CDbl(ValorDesconto.Text)
    lErro = DescontoGlobal_Testa(dValorDesconto)
    If lErro <> SUCESSO Then gError 46990
    
    lErro = gobjTribTab.Valida_Dados
    If lErro <> SUCESSO Then gError 46990
    
    objNFiscal.lNumPedidoVenda = glCodPedido
    objNFiscal.iFilialPedido = giFilialPedido
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    '??? objNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV
    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 83451
    
    'Não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 83452

    'Valida os dados do grid de itens
    lErro = Valida_Dados_GridItens()
    If lErro <> SUCESSO Then gError 46987

    'VAlida os dados do grid de comissões
    lErro = objTabComissoes.Valida_Grid_Comissoes()
    If lErro <> SUCESSO Then gError 46991
    
    'Valida os dados do grid de parcelas
    lErro = Valida_Grid_Parcelas()
    If lErro <> SUCESSO Then gError 46995

    lErro = gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
    If lErro <> SUCESSO Then gError 83397

    Set objNFiscal.objRastreamento = gobjRastreamento
    
    If gobjFAT.iPedidoValidaEmbalagem = MARCADO Then
    
        lErro = gobjEmbalagem.Valida_Grid_Embalagens()
        If lErro <> SUCESSO And lErro <> 96133 Then gError 83397
        If lErro = 96133 Then
        
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ACEITA_GRIDEMB_INCOMPLETO")
            If vbMsgRes = vbNo Then gError 83397
            
        End If
    
    End If

    'recolhe os dados da tela
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 46992
    
    'Recolhe os dados dos grids da tela
    lErro = Move_Grids_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 46993
    
    'embalagem
    lErro = gobjEmbalagem.Move_GridEmb_Memoria(colEmbalagensItens, DOCORIGEM_NF)
    If lErro <> SUCESSO Then gError 46993
    
    If colEmbalagensItens.Count > 0 Then
        For iIndice = 1 To objNFiscal.colItens.Count
            For Each objEmbalagensItem In colEmbalagensItens
                If objEmbalagensItem.iItem = iIndice Then
                    Call objNFiscal.colItens(iIndice).colEmbalagens.Add(objEmbalagensItem)
                End If
            Next
        Next
    End If
    
    GL_objMDIForm.MousePointer = vbHourglass
        
    'VAlida os dados de Tributação
    lErro = gobjTribTab.Valida_TributacaoNF_Gravacao(objNFiscal)
    If lErro <> SUCESSO Then gError 46460
    
    lErro = Calcula_Msg
    If lErro <> SUCESSO Then gError 46460

    'tulio160103
    'se for pra verificar limite de credito
    If gobjCRFAT.iCreditoVerificaLimite = VERIFICAR_LIMITE_CREDITO Then
        
        If Len(Trim(NFiscal.Caption)) = 0 Then
            'verifica se o cliente tem crédito
            lErro = CF("NFiscal_Testa_Credito", objNFiscal)
            If lErro <> SUCESSO Then gError 25737
        End If
    
    End If
    
    'mover a parte do rastreamento
    lErro = gobjRastreamento.Move_Rastro_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 83398

    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataEmissao.Text))
    If lErro <> SUCESSO Then gError 92045

    lErro = CF("NFiscal_Valida_MargContr", objNFiscal)
    If lErro <> SUCESSO Then gError 201044
 
    lErro = CF("NFiscal_Valida_Diversos", objNFiscal)
    If lErro <> SUCESSO Then gError 201044
    
    'Grava a Nota Fiscal Fatura Pedido
    lErro = CF("NFiscalFaturaPedido_Grava", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 46994

    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_NFISCAL
        gobjAnotacao.sID = CStr(objNFiscal.lNumIntDoc)
        gobjAnotacao.dtDataAlteracao = gdtDataHoje
        
        lErro = CF("Anotacoes_Grava", gobjAnotacao)
        If lErro <> SUCESSO Then gError 130406
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    If Len(Trim(NFiscal.Caption)) = 0 Then vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_NUMERO_NOTA_GRAVADA", objNFiscal.lNumNotaFiscal)
    
'    If giTipoVersao = VERSAO_FULL Then
        If ImprimeRomaneio.Value = vbChecked Then
            Call CF("NFiscal_Imprime_Romaneio", objNFiscal)
        End If
'    End If
    
    If ImprimeNFGravacao.Value = vbChecked Then
        
        'Dispara função para imprimir nota fiscal
        lErro = NotaFiscal_Imprime(objNFiscal)
        If lErro <> SUCESSO And lErro <> 102262 Then gError 177067
    
    End If
    
    If ISSerieEletronica(objNFiscal.sSerie) Then
        lErro = CF("NFE_Grava1", objNFiscal.lNumIntDoc, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 199997
    End If
    
    lErro = CF("Verifica_NFiscal_Servico_Eletronica", objNFiscal, iNFSE)
    If lErro <> SUCESSO Then gError 207058
    
    If iNFSE = 1 Then
        lErro = CF("NFSE_Grava1", objNFiscal.lNumIntDoc, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 207059
    End If
    
    Gravar_Registro = SUCESSO

    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 130474
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 46968
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PEDIDO_TRAZIDO", gErr)
        
        Case 46969
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)
            
        Case 46970
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)
            
        Case 46971
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 46972
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 46974
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 46975
            Call Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_NAO_PREENCHIDA", gErr)

        Case 46976
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 46977
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)

        Case 46460, 46978, 46983, 46987, 46991, 46992, 46993, 46994, 46995, 83397, 83398, 83451, 92045, 130406, 102263, 177067, 181786, 199997, 201044, 207058, 207059  ' , 46981

        Case 46984
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_BLOQUEIO_CREDITO", gErr, glCodPedido)

        Case 46985
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_BLOQUEIO_TOTAL", gErr)

        Case 46986
            Call Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_ANTERIOR_DATAEMISSAO", gErr, DataSaida, DataEmissao)
        
        Case 46988
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)

        Case 46989
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr, CDbl(PesoLiquido.Text), CDbl(PesoBruto.Text))

        Case 25737, 46990 'Erro tratado pela rotina chamada
        
        Case 51493
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)
        
        Case 56907
            'Call Rotina_Erro(vbOKOnly, "ERRO_IR_FONTE_MAIOR_VALOR_TOTAL", gErr, dValorIRRF, dValorTotal)

        Case 83452
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iTipo)
            
        Case 206615
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_DATA_MENOR_PV", gErr, StrParaDate(DataEmissao.Text), objPedidoVenda.dtDataEmissao)

        Case 209998
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNENTTERC_NAO_PREENCHIDO", gErr)

        Case 209999
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNENTTERC_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157383)

    End Select

    Exit Function

End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objNFiscal As New ClassNFiscal
Dim snProt As String
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iEnviado As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os dados de algum pedido foi trazido para a tela
    If glCodPedido = 0 Then gError 83937

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 130475
    If Len(Trim(Pedido.Text)) = 0 Then gError 83938
    If Len(Trim(FilialPedido.Text)) = 0 Then gError 83939
    If Len(Trim(Cliente.Caption)) = 0 Then gError 83940
    If Len(Trim(Serie.Text)) = 0 Then gError 83941
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 83942

    'Verifica se o código de pedido é o mesmo que foi trazido para a tela
    If glCodPedido <> StrParaLong(Pedido.Text) Or giFilialPedido <> Codigo_Extrai(FilialPedido.Text) Then gError 83943

    objNFiscal.lNumPedidoVenda = glCodPedido
    objNFiscal.iFilialPedido = giFilialPedido
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
'???    objNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV
    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
    
    'le o tipo da origem
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 203022
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 203023
    
    'se origem for de empresa é interna
    If objTipoDocInfo.iOrigem = DOCINFO_EMPRESA Then
        'Verifica se a existe nota fiscal está cadastrada
        lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
        If lErro <> SUCESSO And lErro <> 43676 Then gError 203018
        
        If lErro = 43676 Then gError 203019
        
'        If gobjCRFAT.iUsaNFe = MARCADO And objNFiscal.iRecibo = 0 Then
'
'            lErro = CF("NFeFedLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'            If lErro <> SUCESSO Then gError 203020
'
'            If iEnviado = 1 Then gError 203021
'
'        ElseIf gobjCRFAT.iUsaNFSE = MARCADO And objNFiscal.iRecibo <> 0 Then
'
'            lErro = CF("RPSWEBLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'            If lErro <> SUCESSO Then gError 207077
'
'            If iEnviado = 1 Then gError 207078
'
'        End If

    End If

    'Verifica se a existe nota fiscal está cadastrada
    lErro = CF("NFiscal_Le_Pedido", objNFiscal)
    If lErro <> SUCESSO And lErro <> 46602 Then gError 83944
    
    'se a nota não está cadastrada ==> erro
    If lErro = 46602 Then gError 83945

    'pede confirmacao
    vbMsg = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_NFISCAL", objNFiscal.lNumNotaFiscal)
    If vbMsg = vbYes Then

        'Lê os itens da nota fiscal
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 83948

        'Faz o cancelamento de uma nota fiscal de Saida
        lErro = CF("NotaFiscalSaida_Excluir", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError 83946
    
        'Limpa a Tela
        lErro = Limpa_Tela_NFiscalFaturaPedido()
        If lErro <> SUCESSO Then gError 83947
        
        'Libera as coleções globais
        Set gcolItemPedido = New Collection
        Set gcolItemQuant = New Collection
        Set gcolItemUM = New Collection 'Inserido por Wagner
    
        'fecha o comando de setas
        Call ComandoSeta_Fechar(Me.Name)

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 130475
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 83937
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PEDIDO_TRAZIDO", gErr)
        
        Case 83938
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)
            
        Case 83939
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)
            
        Case 83940
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 83941
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 83942
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 83943
             Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_TELA_DIFERE_BD", gErr, StrParaLong(Pedido.Text), Codigo_Extrai(FilialPedido.Text), glCodPedido, giFilialPedido)
       
        Case 84944, 83946, 83947, 203018, 203020, 203022, 207077
       
        Case 83945
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA2", gErr, objNFiscal.lNumPedidoVenda, objNFiscal.iFilialPedido, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
       
        Case 203019
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
       
        Case 203021
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CADASTRADA_NAO_EXCLUI", gErr)
       
        Case 203023
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)
       
        Case 207078
            Call Rotina_Erro(vbOKOnly, "ERRO_NFSE_CADASTRADA_NAO_EXCLUI", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157384)
        
    End Select

    Exit Sub

End Sub

Private Sub NFiscal_Imprime_Romaneio(objNFiscal As ClassNFiscal)
    Call CF("NFiscal_Imprime_Romaneio", objNFiscal)
End Sub

Function Valida_Dados_GridItens() As Long

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Valida_Dados_GridItens

    'Valida os dados do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 46995
        
        If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col)) = 0 Then gError 46999
        
    Next
    
    Valida_Dados_GridItens = SUCESSO
    
    Exit Function

Erro_Valida_Dados_GridItens:

    Valida_Dados_GridItens = gErr
    
    Select Case gErr
    
        Case 46995
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 46999
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157385)
    
    End Select
            
    Exit Function

End Function

Private Function Valida_Grid_Parcelas() As Long
'Valida os dados do Grid de Parcelas

Dim lErro As Long
Dim iIndice As Integer
Dim dSomaParcelas As Double
Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim iTamanho As Integer
Dim iTipo As Integer
Dim dPercAcrecFin As Double
Dim iDesconto As Integer
Dim dtDataDesconto As Date
Dim iTipoPagto As Integer
Dim objParcPV As ClassParcelaPedidoVenda
Dim objParcRec As ClassParcelaReceber
Dim dValorPagar As Double

On Error GoTo Erro_Valida_Grid_Parcelas

    'Verifica se alguma parcela foi informada
    If objGridParcelas.iLinhasExistentes = 0 Then gError 51006
    
    dSomaParcelas = 0
    
    'Para cada Parcela do grid de parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        dtDataEmissao = StrParaDate(DataEmissao.Text)
        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        
        'verifica se o vencimento e o valor da parcela estão preenchidos
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) = 0 Then gError 51007
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) = 0 Then gError 51010
        
        'verifica se a data de vencimento da parcela é menor que a data de emissão
        'If dtDataVencimento < dtDataEmissao Then gError 51008
        lErro = CF("NF_Valida_Data_Vencimento", dtDataVencimento, dtDataEmissao, iIndice)
        If lErro <> SUCESSO Then gError 140295
        
        'Se o desconto 1 da parcela está preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))) > 0 Then
            iDesconto = 1
            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
            'Verifica se a data do desconto está preenchida
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))) = 0 Then gError 51500
                        
            dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
            'Valida a data do desconto
            If dtDataDesconto > dtDataVencimento Then gError 51513
            'Se o tipo de desconto for por valor
            If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                'Verifica se o valor do desconto está preenchido
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))) = 0 Then gError 51503
            Else
                'Verifica se o percentual de desconto está preenchido
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Percentual_Col))) = 0 Then gError 51504
            End If
            'Se o desconto 2 está preenchido
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))) > 0 Then
                iDesconto = 2
                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
                'Verifica se a data de desconto está preenchida
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))) = 0 Then gError 51501
                'Verifica se a data de desconto está ordenada ou se é igual ao desconto anterior
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 51509
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 51511
                'Se o desconto for do tipo valor
                If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                    'Verifica se o valor está preenchido
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))) = 0 Then gError 51505
                Else
                    'Verifica se o percentual está preenchido
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col))) = 0 Then gError 51506
                End If
                'Valida a data de desconto
                dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
                If dtDataDesconto > dtDataVencimento Then gError 51514
                'Se o desconto 3 está preenchido
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))) > 0 Then
                    iDesconto = 3
                    iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
                    'Verifica se a data de desconto está preenchida
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))) = 0 Then gError 51502
                    'Verifica se a data de desconto está ordenada ou se é igual a do desconto anterior
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 51510
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 51512
                    'Se o desconto for do tipo valor
                    If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                        'Verifica se valor de desconto está preenchido
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))) = 0 Then gError 51507
                    Else
                        'verifica se o percentuial de desconto está preenchido
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Percentual_Col))) = 0 Then gError 51508
                    End If
                    'Valida a data de desconto
                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
                    If dtDataDesconto > dtDataVencimento Then gError 51515
                End If
            End If
        End If
        'Verifica se as data de vencimentos das parcelas estão ordenadas
        If iIndice > 1 Then If CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col)) < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_Col)) Then gError 51009
       'Faz a soma do total das parcelas
        dSomaParcelas = dSomaParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))

        Set objParcPV = New ClassParcelaPedidoVenda
        
        iTipoPagto = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_TipoPagto_Col))
        
        Set objParcRec = gobjNFiscal.ColParcelaReceber(iIndice)
        
        objParcPV.iCodConta = objParcRec.iCodConta
        objParcPV.iBancoCheque = objParcRec.iBancoCheque
        objParcPV.sAgenciaCheque = objParcRec.sAgenciaCheque
        objParcPV.sContaCorrenteCheque = objParcRec.sContaCorrenteCheque
        objParcPV.lNumeroCheque = objParcRec.lNumeroCheque
        objParcPV.dtDataDepositoCheque = objParcRec.dtDataDepositoCheque
        objParcPV.dtDataEmissaoCheque = objParcRec.dtDataEmissaoCheque
        objParcPV.iAdmMeioPagto = objParcRec.iAdmMeioPagto
        objParcPV.iParcelamento = objParcRec.iParcelamento
        objParcPV.sNumeroCartao = objParcRec.sNumeroCartao
        objParcPV.dtValidadeCartao = objParcRec.dtValidadeCartao
        objParcPV.sAprovacaoCartao = objParcRec.sAprovacaoCartao
        objParcPV.dtDataTransacaoCartao = objParcRec.dtDataTransacaoCartao
        
        lErro = CF("Valida_TipoPagto", iIndice, iTipoPagto, objParcPV)
        If lErro <> SUCESSO Then gError 183014
            
    Next
    
    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then gError 183014
    
    If Abs(dValorPagar - dSomaParcelas) > DELTA_VALORMONETARIO Then gError 51011

    Valida_Grid_Parcelas = SUCESSO
    
    Exit Function
    
Erro_Valida_Grid_Parcelas:

    Valida_Grid_Parcelas = gErr
    
    Select Case gErr
    
        Case 51006
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_PARCELA_COBRANCA", gErr)
        
        Case 51007
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_NAO_INFORMADA", gErr, iIndice)
            
        Case 51008
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_MENOR", gErr, iIndice, dtDataVencimento, dtDataEmissao)
        
        Case 51009
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_COBRANCA_NAO_ORDENADA", gErr)
               
        Case 51010
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PARCELA_COBRANCA_NAO_INFORMADO", gErr, iIndice)
        
        Case 51011
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_COBRANCA_INVALIDA", gErr)

        Case 51500, 51501, 51502
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_NAO_PREENCHIDA", gErr, iDesconto, iIndice)
        
        Case 51503, 51504, 51505, 51506, 51507, 51508
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_PARCELA_NAO_PREENCHIDO", gErr, iDesconto, iIndice)

        Case 51509, 51510
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTOS_DESORDENADAS", gErr, iIndice)

        Case 51511, 51512
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTO_IGUAIS", gErr, iDesconto - 1, iDesconto, iIndice)

        Case 51513, 51514, 51515
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESC_PARCELA_SUP_DATA_VENCIMENTO", gErr, dtDataDesconto, iDesconto, iIndice)

        Case 140295

        Case 183013
            Call Rotina_Erro(vbOKOnly, "TIPOPAGTO_NAO_PREENCHIDO", gErr, GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))

        Case 183014

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157386)
            
    End Select
    
    Exit Function

End Function

Function Move_Tela_Memoria(objNFiscal As ClassNFiscal) As Long
'Recolhe os dados da tela

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objProjetoInfo As Object
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria
            
    If Len(Trim(Pedido.ClipText)) > 0 Then objNFiscal.lNumPedidoVenda = CLng(Pedido.Text)
    objNFiscal.iFilialPedido = Codigo_Extrai(FilialPedido.Text)
    
    'Verifica se o Cliente esta preenchido
    If Len(Trim(Cliente.Caption)) > 0 Then
    
        objCliente.sNomeReduzido = Cliente.Caption
        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 51012
        'Se não achou o Cliente --> erro
        If lErro <> SUCESSO Then gError 51013
        'Guarda código do Cliente
        objNFiscal.lCliente = objCliente.lCodigo
        
    End If
    
    'Armazena o responsável pelo frete
    objNFiscal.iFreteRespons = TipoFrete.ItemData(TipoFrete.ListIndex)
    
    
    If Recibo.Value = vbChecked Then
        objNFiscal.iRecibo = MARCADO
    Else
        objNFiscal.iRecibo = DESMARCADO
    End If
    
    'Recolhe os demais dados da tela
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    
    If Len(Trim(NFiscal.Caption)) Then
        objNFiscal.lNumNotaFiscal = CLng(NFiscal.Caption)
    Else
        objNFiscal.lNumNotaFiscal = 0
    End If
    
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)

    lErro = CF2(Me, "NFCF_Move_DataSaida_Memoria", objNFiscal)
    If lErro <> SUCESSO Then gError 132293
    
    objNFiscal.sNaturezaOp = gobjTribTab.sNatOpInterna
    objNFiscal.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)
    objNFiscal.dValorFrete = StrParaDbl(ValorFrete.Text)
    objNFiscal.dValorSeguro = StrParaDbl(ValorSeguro.Text)
    objNFiscal.dValorOutrasDespesas = StrParaDbl(ValorDespesas.Text)
    objNFiscal.dValorDesconto = StrParaDbl(ValorDesconto.Text)
    objNFiscal.dValorProdutos = StrParaDbl(ValorProdutos.Caption)
    objNFiscal.dValorTotal = StrParaDbl(ValorTotal.Caption)
    objNFiscal.iFilialEntrega = Codigo_Extrai(FilialEntrega.Text)
    objNFiscal.iCodTransportadora = Codigo_Extrai(Transportadora.Text)
    objNFiscal.iCodTranspRedesp = Codigo_Extrai(TranspRedespacho.Text)
    objNFiscal.iDetPagFrete = RedespachoCli.Value
    objNFiscal.sPlaca = Placa.Text
    objNFiscal.sPlacaUF = PlacaUF.Text
    objNFiscal.lVolumeQuant = StrParaLong(VolumeQuant.Text)
    objNFiscal.sVolumeNumero = VolumeNumero.Text
    objNFiscal.sMensagemNota = Mensagem.Text
    objNFiscal.sMensagemCorpoNota = MensagemCorpo.Text
    objNFiscal.dPesoLiq = StrParaDbl(PesoLiquido.Text)
    objNFiscal.dPesoBruto = StrParaDbl(PesoBruto.Text)
    objNFiscal.dVolumeTotal = StrParaDbl(Cubagem.Text)
    objNFiscal.iCanal = Codigo_Extrai(CanalVenda.Text)
    objNFiscal.sNumPedidoTerc = NumPedidoTerc.Text
    '???objNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV
    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.dtDataEntrada = DATA_NULA
    objNFiscal.dtDataVencimento = DATA_NULA
    'objNFiscal.iCondicaoPagto = Codigo_Extrai(CondicaoPagamento.Text)
    objNFiscal.iStatus = STATUS_LANCADO
    objNFiscal.dtDataReferencia = StrParaDate(DataReferencia.Text)
    
    objNFiscal.iCondicaoPagto = CondPagto_Extrai(CondicaoPagamento)
            
    'Incluído por Luiz Nogueira em 21/08/03
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objNFiscal.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objNFiscal.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
    
    lErro = CF2(Me, "NFFP_Move_Tela_Memoria", objNFiscal)
    If lErro <> SUCESSO Then gError 124197
    
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181761
    
    Set objNFiscal.objProjetoInfo = objProjetoInfo
    Set objNFiscal.objInfoAdic = gobjInfoAdic
    objNFiscal.dValorItens = StrParaDbl(ValorProdutos2.Caption)
    objNFiscal.dValorDescontoItens = StrParaDbl(ValorDescontoItens.Text)
    
    'Verifica se o Fornecedor está preenchido
    If Len(Trim(FornEntTerc.ClipText)) > 0 Then

        objFornecedor.sNomeReduzido = FornEntTerc.Text

        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError ERRO_SEM_MENSAGEM

        objNFiscal.lFornEntTerc = objFornecedor.lCodigo
        objNFiscal.iFilialFornEntTerc = Codigo_Extrai(FilialFornEntTerc.Text)

    End If
        
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 51012, 124197, 132293, 181761
        
        Case 51013
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157387)
            
    End Select
            
    Exit Function
        
End Function

Private Function Move_Grids_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os dados dos Grids para memória

Dim lErro As Long

On Error GoTo Erro_Move_Grids_Memoria
    
    'Chama Move_GridItens_Memoria
    lErro = Move_GridItens_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 51014

    'Chama Move_GridAlocacao_Memoria
    lErro = Move_GridAlocacao_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 51015

    'Chama Move_GridComissoes_Memoria
    lErro = objTabComissoes.Move_TabComissoes_Memoria(objNFiscal, NOTA_FISCAL)
    If lErro <> SUCESSO Then gError 51016

    'Chama Move_GridParcelas_Memoria
    lErro = Move_GridParcelas_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 51017
    
    'Mover a parte de tributação
    lErro = gobjTribTab.PreencherNF(objNFiscal)
    If lErro <> SUCESSO Then gError 51018
    
    Move_Grids_Memoria = SUCESSO

    Exit Function

Erro_Move_Grids_Memoria:

    Move_Grids_Memoria = gErr

    Select Case gErr

        Case 51014, 51015, 51016, 51017, 51018

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157388)

    End Select

    Exit Function

End Function

Public Function Move_GridItens_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os Itens do Grid para a Memória

Dim iIndice As Integer
Dim lErro As Long
Dim objItemNF As ClassItemNF, objItemNFAux As ClassItemNF, objItemNFAux2 As ClassItemNF
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lTamanho As Long
Dim colAlocacoes As ColAlocacoesItemNF
Dim objDataEntrega As ClassDataEntrega
Dim dQuantidade As Double
Dim vbMsg As VbMsgBoxResult
Dim objItemPV As ClassItemPedido
Dim dQtdePV As Double
Dim sUM As String
Dim objProduto As ClassProduto
Dim dFator As Double
Dim iItem As Integer

On Error GoTo Erro_Move_GridItens_Memoria

    'Para cada linha existente do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemNF = New ClassItemNF

        'Verifica se o Produto está preenchido
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 51019

        'Armazena produto
        objItemNF.sProduto = sProdutoFormatado
      
        'Guarda os demais campos do Grid em objItemNF
        objItemNF.sDescricaoItem = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
        objItemNF.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        objItemNF.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objItemNF.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))
        objItemNF.dtDataEntrega = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col))
        
        lTamanho = Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col)))

        If lTamanho > 0 Then objItemNF.dPercDesc = CDbl(Format(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col), "General Number"))

        objItemNF.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
        
        'Adiciona na coleção de Ítens
        With objItemNF
            Set objItemNFAux = objNFiscal.ColItensNF.Add(0, iIndice, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, 0, 0, 0, 0, gcolItemPedido.Item(iIndice), colAlocacoes, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .sCcl, STATUS_LANCADO, 0, "", 0, 0, .iControleEstoque, .sUMEstoque, .iClasseUM, 0)
        End With
        
        Set objNFiscal.ColItensNF(iIndice).objInfoAdicDocItem = gobjNFiscal.ColItensNF(iIndice).objInfoAdicDocItem

        '********************* TRATAMENTO DE GRADE *****************
        Call Move_ItensGrade_Tela(objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade, gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade)
    
        Set objItemNFAux.colDataEntrega = gobjNFiscal.ColItensNF(iIndice).colDataEntrega
    
        lErro = CF2(Me, "NFFP_Move_GridItem_Memoria", objNFiscal, objItemNFAux, iIndice)
        If lErro <> SUCESSO Then gError 51019
        
        dQuantidade = 0

        For Each objDataEntrega In objItemNFAux.colDataEntrega
            dQuantidade = dQuantidade + objDataEntrega.dQuantidade
        Next
        
        If dQuantidade > 0 Then
            If Abs(dQuantidade - objItemNFAux.dQuantidade) > QTDE_ESTOQUE_DELTA Then gError 183244
        End If
        
        sUM = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
    
        Set objProduto = New ClassProduto
    
        objProduto.sCodigo = sProdutoFormatado
    
        'Lê o produto da linha passada por iLinha do GridItens
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 196966
        
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, gcolItemUM(iIndice), sUM, dFator)
        If lErro <> SUCESSO Then gError 196967
       
        Set objItemNFAux2 = gobjNFiscal.ColItensNF.Item(iIndice)
        
        objItemNFAux.iRecebForaFaixa = objItemNFAux2.iRecebForaFaixa
        objItemNFAux.dPercentMaisReceb = objItemNFAux2.dPercentMaisReceb
        objItemNFAux.dPercentMenosReceb = objItemNFAux2.dPercentMenosReceb
        
        If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        
             For iItem = 0 To TabPrecoItemNF.ListCount - 1
                
                If GridItens.TextMatrix(iIndice, iGrid_TabPrecoItem_Col) = TabPrecoItemNF.List(iItem) Then
        
                    objItemNFAux.iTabelaPreco = TabPrecoItemNF.ItemData(iItem)
                    Exit For
        
                End If
                
            Next
        
        End If
        
        objItemNFAux.dComissao = -1
        If gobjFAT.iExibeComissaoItensPV = MARCADO And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col))) <> 0 Then
            objItemNFAux.dComissao = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col))
        End If
    
        Set objItemPV = New ClassItemPedido
        
        objItemPV.lNumIntDoc = gcolItemPedido(iIndice)
        objItemPV.iFilialEmpresa = Codigo_Extrai(FilialPedido.Text)
        
        lErro = CF("ItensPedido_Le_NumIntDoc1", objItemPV)
        If lErro <> SUCESSO Then gError 196968
        
        dQtdePV = objItemPV.dQuantidade - objItemPV.dQuantCancelada
        
        If objItemNFAux2.iRecebForaFaixa = ERRO_E_REJEITA_RECEBIMENTO Then
        
            If objItemNF.dQuantidade - ((gcolItemQuant(iIndice) + (dQtdePV * objItemNFAux2.dPercentMaisReceb)) * dFator) > QTDE_ESTOQUE_DELTA Then gError 196960
    
        End If
        
        If objItemNFAux2.iRecebForaFaixa = AVISA_E_ACEITA_RECEBIMENTO Then
        
            If objItemNF.dQuantidade - ((gcolItemQuant(iIndice) + (dQtdePV * objItemNFAux2.dPercentMaisReceb)) * dFator) > QTDE_ESTOQUE_DELTA Then
            
                vbMsg = Rotina_Aviso(vbYesNo, "AVISO_QUANTIDADE_ACIMA_FAIXA_PV", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), objItemNF.dQuantidade, ((gcolItemQuant(iIndice) + (dQtdePV * objItemNFAux2.dPercentMaisReceb)) * dFator))
                
                If vbMsg = vbNo Then gError 196961
    
            End If
    
        End If
        
    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 51019, 51020
        
        Case 183244
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTSOLICITADA_DIFERE1", gErr, iIndice, GridItens.TextMatrix(iIndice, iGrid_Produto_Col), objItemNFAux.dQuantidade, dQuantidade)
        
        Case 196960
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_MAIOR_SALDO_PV", gErr, GridItens.TextMatrix(iIndice, iGrid_Produto_Col))
        
        Case 196961, 196966 To 196968
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157389)

    End Select

    Exit Function

End Function

Private Function Move_GridAlocacao_Memoria(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iItem As Integer
Dim objItemAloc As ClassItemNFAlocacao, objItemAloc2 As ClassItemNFAlocacao
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Move_GridAlocacao_Memoria

    'Se o número de linhas existentes e  maior que zero
    If objGridAlocacoes.iLinhasExistentes > 0 Then

        iIndice = 0
        'Percorre o grid de alocações
        For iIndice = 1 To objGridAlocacoes.iLinhasExistentes

            iItem = CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))

            Set objItemAloc = New ClassItemNFAlocacao

            'Preenche o objItemAloc
            objItemAloc.dQuantidade = CDbl(GridAlocacao.TextMatrix(iIndice, iGrid_QuantAloc_Col))
            objItemAloc.sAlmoxarifado = GridAlocacao.TextMatrix(iIndice, iGrid_AlmoxAloc_Col)
            objItemAloc.sUnidadeMed = GridAlocacao.TextMatrix(iIndice, iGrid_UMAloc_Col)
            'Verifica se o Almoxarifado já foi lido
            If colCodigoNome.Count > 0 Then
                For Each objCodigoNome In colCodigoNome
                    If objCodigoNome.sNome = objItemAloc.sAlmoxarifado Then
                        objItemAloc.iAlmoxarifado = objCodigoNome.iCodigo
                        Exit For
                    End If
                Next
            End If
            'Se o almoxarifado ainda não foi lido
            If objItemAloc.iAlmoxarifado = 0 Then
                'Lê o almoxarifado
                objAlmoxarifado.sNomeReduzido = objItemAloc.sAlmoxarifado
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25060 Then gError 51022
                If lErro = 25060 Then gError 51023 'Não encontrou
                
                objItemAloc.iAlmoxarifado = objAlmoxarifado.iCodigo
                'Adiciona o almoxarifado na coleção de almoxarifados lidos
                colCodigoNome.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido
                
            End If
            'Adiciona a alocação do grid na coleção de itens
            With objItemAloc
                Set objItemAloc2 = objNFiscal.ColItensNF(iItem).colAlocacoes.Add(.iAlmoxarifado, .sAlmoxarifado, .dQuantidade)
                objItemAloc2.sUnidadeMed = .sUnidadeMed
            End With
               
        Next
        
    End If

    Move_GridAlocacao_Memoria = SUCESSO

    Exit Function

Erro_Move_GridAlocacao_Memoria:

    Move_GridAlocacao_Memoria = gErr

    Select Case gErr

        Case 51022

        Case 51023
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157390)

    End Select

    Exit Function

End Function

Private Function Move_GridParcelas_Memoria(objNFiscal As ClassNFiscal) As Long
'Move as Parcelas do Grid para a Memória

Dim iIndice As Integer
Dim lTamanho As Long
Dim objParcela As ClassParcelaReceber
Dim dtDataEmissao As Date
Dim dtDataReferencia As Date
Dim lErro As Long
Dim objParcRec As ClassParcelaReceber
Dim objParcRec1 As ClassParcelaReceber

On Error GoTo Erro_Move_GridParcelas_Memoria

    dtDataReferencia = StrParaDate(DataReferencia.Text)
    dtDataEmissao = StrParaDate(DataEmissao.Text)

    If dtDataReferencia <> DATA_NULA Then
        lErro = CF("NF_Valida_Data_Referencia", dtDataReferencia, dtDataEmissao)
        If lErro <> SUCESSO Then gError 140329
    End If

    'Se não há parcelas a recolher, sai da função
    If objGridParcelas.iLinhasExistentes = 0 Then Exit Function
    
    'Para cada parcela do grid
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
    
        Set objParcela = New ClassParcelaReceber
        
        objParcela.iNumParcela = iIndice

        'recolhe os dados da parcela
        objParcela.dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        objParcela.dValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
        objParcela.iTipoPagto = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_TipoPagto_Col))
        objParcela.iDesconto1Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
        objParcela.iDesconto2Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
        objParcela.iDesconto3Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
        objParcela.dtDesconto1Ate = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
        objParcela.dtDesconto2Ate = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
        objParcela.dtDesconto3Ate = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
        
        'Se o tipo de desconto for de Valor recolhe a coluna valor
        'Senão recolhe a coluna percentual
        If objParcela.iDesconto1Codigo = VALOR_FIXO Or objParcela.iDesconto1Codigo = VALOR_ANT_DIA Or objParcela.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
            objParcela.dDesconto1Valor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))
        ElseIf objParcela.iDesconto1Codigo = Percentual Or objParcela.iDesconto1Codigo = PERC_ANT_DIA Or objParcela.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Percentual_Col)))
            If lTamanho > 0 Then objParcela.dDesconto1Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Percentual_Col))
        End If
        
        'Se o tipo de desconto for de Valor recolhe a coluna valor
        'Senão recolhe a coluna percentual
        If objParcela.iDesconto2Codigo = VALOR_FIXO Or objParcela.iDesconto2Codigo = VALOR_ANT_DIA Or objParcela.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
            objParcela.dDesconto2Valor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))
        ElseIf objParcela.iDesconto2Codigo = Percentual Or objParcela.iDesconto2Codigo = PERC_ANT_DIA Or objParcela.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col)))
            If lTamanho > 0 Then objParcela.dDesconto2Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col))
        End If
        
        'Se o tipo de desconto for de Valor recolhe a coluna valor
        'Senão recolhe a coluna percentual
        If objParcela.iDesconto3Codigo = VALOR_FIXO Or objParcela.iDesconto3Codigo = VALOR_ANT_DIA Or objParcela.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
            objParcela.dDesconto3Valor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))
        ElseIf objParcela.iDesconto3Codigo = Percentual Or objParcela.iDesconto3Codigo = PERC_ANT_DIA Or objParcela.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Percentual_Col)))
            If lTamanho > 0 Then objParcela.dDesconto3Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Percentual_Col))
        End If
    
        'Adiciona a parcela na coleção de parcelas da Nota Fiscal
        With objParcela
            Set objParcRec = objNFiscal.ColParcelaReceber.Add(0, 0, iIndice, STATUS_ABERTO, .dtDataVencimento, .dtDataVencimento, .dValor, .dValor, 1, CARTEIRA_CARTEIRA, COBRADOR_PROPRIA_EMPRESA, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0, .iPrevisao, .sObservacao, .dValor)
        
            objParcRec.iTipoPagto = objParcela.iTipoPagto
        
            Set objParcRec1 = gobjNFiscal.ColParcelaReceber.Item(iIndice)
            
            objParcRec.dtDataDepositoCheque = objParcRec1.dtDataDepositoCheque
            objParcRec.dtDataEmissaoCheque = objParcRec1.dtDataEmissaoCheque
            objParcRec.dtValidadeCartao = objParcRec1.dtValidadeCartao
            objParcRec.iAdmMeioPagto = objParcRec1.iAdmMeioPagto
            objParcRec.iBancoCheque = objParcRec1.iBancoCheque
            objParcRec.iCodConta = objParcRec1.iCodConta
            objParcRec.dtDataCredito = objParcRec1.dtDataCredito
            objParcRec.iParcelamento = objParcRec1.iParcelamento
            objParcRec.lNumeroCheque = objParcRec1.lNumeroCheque
            objParcRec.sAgenciaCheque = objParcRec1.sAgenciaCheque
            objParcRec.sAprovacaoCartao = objParcRec1.sAprovacaoCartao
            objParcRec.sContaCorrenteCheque = objParcRec1.sContaCorrenteCheque
            objParcRec.sNumeroCartao = objParcRec1.sNumeroCartao
            objParcRec.dtDataTransacaoCartao = objParcRec1.dtDataTransacaoCartao
            
        End With

    Next

    objNFiscal.iNaoImpCobranca = NaoImpCobranca.Value
    
    objNFiscal.dValorDescontoTit = StrParaDbl(ValorDescontoTit.Text)
    
    Move_GridParcelas_Memoria = SUCESSO
    
    Exit Function

Erro_Move_GridParcelas_Memoria:

    Move_GridParcelas_Memoria = gErr
    
    Select Case gErr
    
        Case 51612
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAREFERENCIA", gErr)
        
        Case 140329
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157391)
            
    End Select

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "NFiscal"

    'Lê os dados da Tela NFiscal
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 51027

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objNFiscal.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "TipoNFiscal", objNFiscal.iTipoNFiscal, 0, "TipoNFiscal"
    colCampoValor.Add "NaturezaOp", objNFiscal.sNaturezaOp, STRING_BUFFER_MAX_TEXTO, "NaturezaOp"
    colCampoValor.Add "Serie", objNFiscal.sSerie, STRING_BUFFER_MAX_TEXTO, "Serie"
    colCampoValor.Add "NumNotaFiscal", objNFiscal.lNumNotaFiscal, 0, "NumNotaFiscal"
    colCampoValor.Add "DataEmissao", objNFiscal.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "DataReferencia", objNFiscal.dtDataReferencia, 0, "DataReferencia"
    colCampoValor.Add "DataSaida", objNFiscal.dtDataSaida, 0, "DataSaida"
'horasaida
    colCampoValor.Add "HoraSaida", CDbl(objNFiscal.dtHoraSaida), 0, "HoraSaida"
    colCampoValor.Add "ValorProdutos", objNFiscal.dValorProdutos, 0, "ValorProdutos"
    colCampoValor.Add "ValorFrete", objNFiscal.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorSeguro", objNFiscal.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "ValorOutrasDespesas", objNFiscal.dValorOutrasDespesas, 0, "ValorOutrasDespesas"
    colCampoValor.Add "ValorDesconto", objNFiscal.dValorDesconto, 0, "ValorDesconto"
    colCampoValor.Add "ValorTotal", objNFiscal.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "Placa", objNFiscal.sPlaca, STRING_BUFFER_MAX_TEXTO, "Placa"
    colCampoValor.Add "PlacaUF", objNFiscal.sPlacaUF, STRING_BUFFER_MAX_TEXTO, "PlacaUF"
    colCampoValor.Add "VolumeQuant", objNFiscal.lVolumeQuant, 0, "VolumeQuant"
    colCampoValor.Add "VolumeEspecie", objNFiscal.lVolumeEspecie, 0, "VolumeEspecie" 'Alterado por Luiz Nogueira em 21/08/03
    colCampoValor.Add "VolumeMarca", objNFiscal.lVolumeMarca, 0, "VolumeMarca" 'Alterado por Luiz Nogueira em 21/08/03
    colCampoValor.Add "MensagemNota", objNFiscal.sMensagemNota, STRING_BUFFER_MAX_TEXTO, "MensagemNota"
    colCampoValor.Add "PesoLiq", objNFiscal.dPesoLiq, 0, "Pesoliq"
    colCampoValor.Add "PesoBruto", objNFiscal.dPesoBruto, 0, "PesoBruto"
    colCampoValor.Add "Canal", objNFiscal.iCanal, 0, "Canal"
    colCampoValor.Add "Cliente", objNFiscal.lCliente, 0, "Cliente"
    colCampoValor.Add "FilialCli", objNFiscal.iFilialCli, 0, "FilialCli"
    colCampoValor.Add "FilialEntrega", objNFiscal.iFilialEntrega, 0, "FilialEntrega"
    colCampoValor.Add "NumPedidoTerc", objNFiscal.sNumPedidoTerc, STRING_BUFFER_MAX_TEXTO, "NumPedidoTerc"
    colCampoValor.Add "NumPedidoVenda", objNFiscal.lNumPedidoVenda, 0, "NumPedidoVenda"
    colCampoValor.Add "CodTransportadora", objNFiscal.iCodTransportadora, 0, "CodTransportadora"
    colCampoValor.Add "CodTranspRedesp", objNFiscal.iCodTranspRedesp, 0, "CodTranspRedesp"
    colCampoValor.Add "DetPagFrete", objNFiscal.iDetPagFrete, 0, "DetPagFrete"
    colCampoValor.Add "TabelaPreco", objNFiscal.iTabelaPreco, 0, "TabelaPreco"
    colCampoValor.Add "FilialPedido", objNFiscal.iFilialPedido, 0, "FilialPedido"
    colCampoValor.Add "NumIntDocCPR", objNFiscal.lNumIntDocCPR, 0, "NumIntDocCPR"
    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_PEDIDOVENDA_VOLUME_NUMERO, "VolumeNumero"
    colCampoValor.Add "FreteRespons", objNFiscal.iFreteRespons, 0, "FreteRespons"
    colCampoValor.Add "Status", objNFiscal.iStatus, 0, "Status"
    colCampoValor.Add "SemDataSaida", objNFiscal.iSemDataSaida, 0, "SemDataSaida"
    colCampoValor.Add "NaoImpCobranca", objNFiscal.iNaoImpCobranca, 0, "NaoImpCobranca"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "FilialPedido", OP_DIFERENTE, 0
    colSelecao.Add "NumPedidoVenda", OP_DIFERENTE, 0
    colSelecao.Add "TipoNFiscal", OP_IGUAL, Codigo_Extrai(TipoNFiscal.Text)
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
'    colSelecao.Add "Status", OP_DIFERENTE, STATUS_CANCELADO
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 51027

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157392)

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

        'Evita ter que alterar essa função a cada campo inserido na tabela de NF
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 180095
        
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 51029

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 51029, 180095

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157393)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

Dim lErro As Long

On Error GoTo Erro_Form_Activate

    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError 59336
        
    Call TelaIndice_Preenche(Me)

    Exit Sub
     
Erro_Form_Activate:

    Select Case gErr
          
        Case 59336
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157394)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Public Sub PedidoLabel_Click()
    
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colSelecao As New Collection

    'Chama a tela de browse dos pedidos que podem gerar Notas Fiscais
    Call Chama_Tela("PedidoNFiscalLista", colSelecao, objPedidoVenda, objEventoPedido)
    
End Sub

Private Sub objEventoPedido_evSelecao(obj1 As Object)

Dim objPedidoVenda As ClassPedidoDeVenda

    Set objPedidoVenda = obj1

    'Coloca o código dopedido na tela
    Pedido.Text = objPedidoVenda.lCodigo
    'Coloca a filial do pedido na tela
    FilialPedido.Text = objPedidoVenda.iFilialEmpresa
    Call FilialPedido_Validate(bSGECancelDummy)
    
    'Traz os dados do pedido selecionado para a tela
    BotaoPedido_Click
    
    Me.Show

End Sub

Public Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objNFiscal Is Nothing) Then

        lErro = CargaPosFormLoad(True)
        If lErro <> SUCESSO Then gError 64006

        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 60189
        
        If lErro <> SUCESSO Then gError 60190

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

        'Lê o Tipo de Documento com o Código de objNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 60191
        
        'Se não encontrar --> Erro
        If lErro = 31415 Then gError 60192
        
'        If giTipoVersao = VERSAO_FULL Then
            If objTipoDocInfo.sNomeTelaNFiscal <> NOME_TELA_NFISCALFATURAPEDIDO Then gError 60193
'        End If
        
        'Traz a nota para a tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 60194

    End If

    iAlterado = 0
    iComissoesAlterada = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 60189, 60191, 60194, 64006

        Case 60190
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case 60192
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 60193
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_DIFERENTE_NF_FATURA_PEDIDO", gErr, objNFiscal.iTipoNFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157395)

    End Select
    
    iAlterado = 0
    iComissoesAlterada = 0

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'tulio261202
    'feito em virtude do redimensionamento da msg da nota p/ inpal
    Call CF("Redimensiona_Tam_MsgNF", Mensagem)

    Parent.HelpContextID = IDH_NF_FATURA_PEDIDO_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Nota Fiscal Fatura - Gerada a Partir de um Pedido"
    Call Form_Load
    
End Function

Public Function Name() As String
    Name = "NFiscalFaturaPedido"
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

Private Function DescontoGlobal_Testa(dValorDesconto As Double) As Long

Dim dValorProdutos As Double
Dim dValorFrete As Double
Dim dValorSeguro As Double
Dim dValorDespesas As Double
    
Dim lErro As Long

On Error GoTo Erro_DescontoGlobal_Testa
    
    'Recolhe os valores da tela
    If Len(Trim(ValorProdutos.Caption)) > 0 Then dValorProdutos = CDbl(ValorProdutos.Caption)
    If Len(Trim(ValorFrete.ClipText)) > 0 Then dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorSeguro.ClipText)) > 0 Then dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(ValorDespesas.ClipText)) > 0 Then dValorDespesas = CDbl(ValorDespesas.Text)

    'Verifica se o Valor do Desconto é maior que a soma dos demais valores
    If dValorProdutos + dValorFrete + dValorSeguro + dValorDespesas <= 0 Then gError 56898 'Alterado por Wagner
    
    DescontoGlobal_Testa = SUCESSO
     
    Exit Function
    
Erro_DescontoGlobal_Testa:

    DescontoGlobal_Testa = gErr
     
    Select Case gErr
                  
        Case 56898
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_ULTRAPASSOU_SOMA_VALORES", gErr, dValorDesconto, dValorProdutos + dValorDesconto + dValorFrete + dValorSeguro + dValorDespesas) 'Alterado por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157396)
     
    End Select
     
    Exit Function

End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Pedido Then
            Call PedidoLabel_Click
        ElseIf Me.ActiveControl Is Serie Then
            Call SerieLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is NFiscal Then
            Call NFiscalLabel_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        ElseIf Me.ActiveControl Is TranspRedespacho Then
            Call TranspRedLabel_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call BotaoVendedores_Click
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        ElseIf Me.ActiveControl Is Embalagem Then
            Call BotaoEmbalagens_Click
        Else
            Call gobjRastreamento.objUserControl_KeyDown(KeyCode, Shift)
        End If
    
    End If

    If KeyCode = KEYCODE_BOTAOCONSULTA Then
        Call BotaoTipoPagto_Click
    End If

End Sub

Public Sub Vendedor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Vendedor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComissoes)
End Sub

Public Sub Vendedor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
End Sub

Public Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = Vendedor
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PercentualComissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercentualComissao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComissoes)
End Sub

Public Sub PercentualComissao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
End Sub

Public Sub PercentualComissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = PercentualComissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorBase_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorBase_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComissoes)
End Sub

Public Sub ValorBase_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
End Sub

Public Sub ValorBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorBase
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PercentualEmissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercentualEmissao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComissoes)
End Sub

Public Sub PercentualEmissao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
End Sub

Public Sub PercentualEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = PercentualEmissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorEmissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorEmissao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComissoes)
End Sub

Public Sub ValorEmissao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
End Sub

Public Sub ValorEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorEmissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PercentualBaixa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercentualBaixa_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComissoes)
End Sub

Public Sub PercentualBaixa_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
End Sub

Public Sub PercentualBaixa_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridComissoes.objControle = PercentualBaixa
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorBaixa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorBaixa_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComissoes)
End Sub

Public Sub ValorBaixa_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
End Sub

Public Sub ValorBaixa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorBaixa
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Preenche_DescontoPadrao(iLinha As Integer) As Long

Dim lErro As Long
Dim colDescontoPadrao As New colDesconto
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim iColuna  As Integer
Dim dtDataVencimento As Date
Dim dPercentual As Double
Dim dValorParcela As Double
Dim sValorDesconto As String

On Error GoTo Erro_Preenche_DescontoPadrao
   
    'Se a data de referencia estiver preenchida
    If Len(Trim(DataReferencia.ClipText)) > 0 Then
    
        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_Col))
        lErro = CF("Parcela_GeraDescontoPadrao", colDescontoPadrao, dtDataVencimento)
        If lErro <> SUCESSO Then gError 26642
        
        If colDescontoPadrao.Count > 0 Then
        
            'Para cada um dos desontos padrão
            For iIndice1 = 1 To colDescontoPadrao.Count
                
                'Seleciona a coluna correspondente ao Desconto
                If iIndice1 = 1 Then iColuna = iGrid_Desc1Codigo_Col
                If iIndice1 = 2 Then iColuna = iGrid_Desc2Codigo_Col
                If iIndice1 = 3 Then iColuna = iGrid_Desc3Codigo_Col
                
                'Seleciona o tipo de desconto
                For iIndice2 = 0 To Desconto1Codigo.ListCount - 1
                    If colDescontoPadrao.Item(iIndice1).iCodigo = Desconto1Codigo.ItemData(iIndice2) Then
                        GridParcelas.TextMatrix(iLinha, iColuna) = Desconto1Codigo.List(iIndice2)
                        GridParcelas.TextMatrix(iLinha, iColuna + 1) = Format(colDescontoPadrao.Item(iIndice1).dtData, "dd/mm/yyyy")
                        GridParcelas.TextMatrix(iLinha, iColuna + 3) = Format(colDescontoPadrao.Item(iIndice1).dValor, "Percent")
                        
                        '*** Inicio colocacao Valor Desconto na tela
                        dPercentual = colDescontoPadrao.Item(iIndice1).dValor
                        dValorParcela = StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col))
                    
                        'Coloca Valor do Desconto na tela
                        If dValorParcela > 0 Then
                            sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
                            GridParcelas.TextMatrix(iLinha, iColuna + 2) = sValorDesconto
                        End If
                        '*** Fim colocacao Valor Desconto na tela

                    End If
                Next
            Next
    
        End If
    
    End If
    
    Preenche_DescontoPadrao = SUCESSO
    
    Exit Function
    
Erro_Preenche_DescontoPadrao:

    Preenche_DescontoPadrao = gErr
    
    Select Case gErr
    
        Case 26642
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157397)
            
    End Select
    
    Exit Function

End Function

'Início contabilidade
Public Sub CTBBotaoModeloPadrao_Click()
    Call objContabil.Contabil_BotaoModeloPadrao_Click
End Sub

Public Sub CTBModelo_Click()
    Call objContabil.Contabil_Modelo_Click
End Sub

Public Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

'    If giTipoVersao = VERSAO_LIGHT Then
'        Call objContabil.Contabil_GridContabil_Consulta_Click
'    End If

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

Public Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)
    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)
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

Private Sub objEventoLote_evSelecao(obj1 As Object)
'Traz o lote selecionado para a tela
    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)
End Sub

Private Sub objEventoDoc_evSelecao(obj1 As Object)
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
Dim dQuantidadeConvertida As Double
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim objCliente As New ClassCliente
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim bEncontrouProduto As Boolean
Dim iLinha2 As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sContaMascarada As String
Dim iAlmoxPadrao As Integer
Dim bEncontrouQuant As Boolean
Dim bEncontrouQuant2 As Boolean
Dim objItemMovEstoque As New ClassItemMovEstoque
Dim sProduto As String
Dim sUM As String
Dim dQuantidade As Double
Dim objFilial As New ClassFilialCliente
Dim objProdutoCategoria As New ClassProdutoCategoria
Dim objCatProdItem As ClassCategoriaProdutoItem
Dim objPlanoConta As New ClassPlanoConta
Dim objProduto As New ClassProduto
Dim objTribItem As Object, dCOFINSCredito As Double, dPISCredito As Double
Dim dValorAux As Double

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
        
        Case ESCANINHO_CUSTO_CONSIG
            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(GridAlocacao.TextMatrix(iLinha, iGrid_ProdutoAloc_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add ESCANINHO_3_EM_CONSIGNACAO
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            
            Next
                    
        Case ESCANINHO_CUSTO
            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(GridAlocacao.TextMatrix(iLinha, iGrid_ProdutoAloc_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add ESCANINHO_NOSSO
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            
            Next
        
        Case QUANT_ALOCADA_CONSIG
            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(GridAlocacao.TextMatrix(iLinha, iGrid_QuantAloc_Col)) > 0 Then
                    
                    'Define o produto que será passado como parâmetro para MovEstoque_Le_QuantVendConsig
                    lErro = CF("Produto_Formata", GridAlocacao.TextMatrix(iLinha, iGrid_ProdutoAloc_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 79012
                    
                    objItemMovEstoque.sProduto = sProdutoFormatado
                    
                    'Define o almoxarifado que será passado como parâmetro para MovEstoque_Le_QuantVendConsig
                    objAlmoxarifado.sNomeReduzido = GridAlocacao.TextMatrix(iLinha, iGrid_AlmoxAloc_Col)
                    
                    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25060 Then gError 79011
                    
                    objItemMovEstoque.iAlmoxarifado = objAlmoxarifado.iCodigo
                    
                    'Define o tipo de movimento, o DocOrigem e o TipoNumIntoDocOrigem que serão passados como parâmetros para MovEstoque_Le_QuantVendConsig
                                        
                    objItemMovEstoque.iTipoMov = MOV_EST_NF_VENDA_MAT_CONSIG
                    objItemMovEstoque.sDocOrigem = "NFISFVPV" & " " & Serie.Text & " " & objContabil.lNumNotaFiscalExt
                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                                                           
                    'Verifica se MovEstoque_Le_QuantVendConsig não encontrou erro
                    lErro = CF("MovEstoque_Le_QuantVendConsig", objItemMovEstoque)
                    If lErro <> SUCESSO And lErro <> 79003 Then gError 79010
                
                    'se não conseguiu encontrar mov. estoque para os parametros em questao ==> quantidade é zerada
                    If lErro = 79003 Then objItemMovEstoque.dQuantidade = 0
                
                    'Passa para o mnemônico o valor encontrado por MovEstoque_Le_QuantVendConsig
                    objMnemonicoValor.colValor.Add objItemMovEstoque.dQuantidade
                    
                Else
                    objMnemonicoValor.colValor.Add 0
                
                End If
            
            Next
        
        Case CTACONTABILEST1
            
            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(Trim(GridAlocacao.TextMatrix(iLinha, iGrid_QuantAloc_Col))) > 0 Then
                    
                    If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
                            
                        objAlmoxarifado.sNomeReduzido = CStr(objMnemonicoValor.vParam(3))

                        'Lê o Nome Reduzido do Almoxarifado
                        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                        If lErro <> SUCESSO And lErro <> 25060 Then gError 64222

                        'Se não encontrou ===> Erro
                        If lErro = SUCESSO Then

                            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
                            
                            lErro = CF("Produto_Formata", objMnemonicoValor.vParam(1), sProdutoFormatado, iPreenchido)
                            If iPreenchido <> PRODUTO_PREENCHIDO Then gError 64223

                            objEstoqueProduto.sProduto = sProdutoFormatado
                            
                            'Lê a conta contábil do Produto no Almoxarifado
                            lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
                            If lErro <> SUCESSO And lErro <> 49991 Then gError 64224

                            If Len(Trim(objEstoqueProduto.sContaContabil)) > 0 Then

                                sContaMascarada = String(STRING_CONTA, 0)

                                lErro = Mascara_MascararConta(objEstoqueProduto.sContaContabil, sContaMascarada)
                                If lErro <> SUCESSO Then gError 64225

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
                    
        Case CODIGO1
            If Len(Pedido.Text) > 0 Then
                objMnemonicoValor.colValor.Add CLng(Pedido.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case QUANT_ESTOQUE
            For iLinha = 1 To objGridItens.iLinhasExistentes
            
                If Len(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
                    
                    lErro = CF("UMEstoque_Conversao", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), CDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)), dQuantidadeConvertida)
                    If lErro <> SUCESSO Then gError 64214

                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        Case ALMOX1
            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(GridAlocacao.TextMatrix(iLinha, iGrid_AlmoxAloc_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridAlocacao.TextMatrix(iLinha, iGrid_AlmoxAloc_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case DATA_EMISSAO
            If Len(DataEmissao.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataEmissao.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
        
        Case DATA_SAIDA
            If Len(DataSaida.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataSaida.FormattedText)
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
            If Len(Filial.Caption) > 0 Then
                objMnemonicoValor.colValor.Add Filial.Caption
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case FILIAL_CONTA
            
            If Len(Filial.Caption) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Caption)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Caption, objFilial)
                If lErro <> SUCESSO Then gError 186138
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError 186139
                
                Else
                
                    sContaMascarada = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaMascarada
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case CLIENTE1
            If Len(Cliente.Caption) > 0 Then
                objMnemonicoValor.colValor.Add Cliente.Caption
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case CLIENTE_CODIGO
            If Len(Cliente.Text) > 0 Then
            
                objCliente.sNomeReduzido = Cliente.Text
                
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then gError 88029
                
                objMnemonicoValor.colValor.Add objCliente.lCodigo
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case NATUREZA_OP
            If Len(gobjTribTab.sNatOpInterna) > 0 Then
                objMnemonicoValor.colValor.Add gobjTribTab.sNatOpInterna
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case NFISCAL1
            If Len(NFiscal.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CLng(NFiscal.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

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
            
        Case Serie1
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
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case VALOR_TOTAL
            If Len(ValorTotal.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorTotal.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case PRECO_UNITARIO
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case PRECO_TOTAL
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col))
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

        Case ICMS_DIFAL_REMETENTE
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dICMSInterestVlrUFRemet

        Case ICMS_DIFAL_DESTINO
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dICMSInterestVlrUFDest

        Case ICMS_FCP
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dICMSInterestVlrFCPUFDest
        
        Case ICMSSUBST
            If Len(ICMSSubstValor1.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ICMSSubstValor1.Caption)
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

        Case VALOR_PRODUTOS
            If Len(ValorProdutos.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorProdutos.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
                    
        Case PRODUTO_ALMOX
            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(GridAlocacao.TextMatrix(iLinha, iGrid_ProdutoAloc_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridAlocacao.TextMatrix(iLinha, iGrid_ProdutoAloc_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case QUANT_ALOCADA
            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(Trim(GridAlocacao.TextMatrix(iLinha, iGrid_QuantAloc_Col))) > 0 Then
                    If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
                        objMnemonicoValor.colValor.Add CDbl(GridAlocacao.TextMatrix(iLinha, iGrid_QuantAloc_Col))
                    Else
                        objMnemonicoValor.colValor.Add 0
                    End If
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case UNIDADE_MED_EST
            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(GridAlocacao.TextMatrix(iLinha, iGrid_UMAloc_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridAlocacao.TextMatrix(iLinha, iGrid_UMAloc_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case PIS_DEBITO_TOTAL
            'objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dPISCredito
            dPISCredito = 0
            For Each objTribItem In gobjTribTab.objNFTributacao.colItens
                dPISCredito = dPISCredito + objTribItem.dPISCredito
            Next
            objMnemonicoValor.colValor.Add Arredonda_Moeda(dPISCredito)
        
        Case COFINS_DEBITO_TOTAL
            'objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dCOFINSCredito
            dCOFINSCredito = 0
            For Each objTribItem In gobjTribTab.objNFTributacao.colItens
                dCOFINSCredito = dCOFINSCredito + objTribItem.dCOFINSCredito
            Next
            objMnemonicoValor.colValor.Add Arredonda_Moeda(dCOFINSCredito)
        
        Case PIS_VALOR_TOTAL
            dValorAux = 0
            For Each objTribItem In gobjTribTab.objNFTributacao.colItens
                dValorAux = dValorAux + objTribItem.dPISValor
            Next
            objMnemonicoValor.colValor.Add Arredonda_Moeda(dValorAux)
        
        Case COFINS_VALOR_TOTAL
            dValorAux = 0
            For Each objTribItem In gobjTribTab.objNFTributacao.colItens
                dValorAux = dValorAux + objTribItem.dCOFINSValor
            Next
            objMnemonicoValor.colValor.Add Arredonda_Moeda(dValorAux)
        
        Case ISS_VALOR
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dISSValor

        Case ISS_INCLUSO
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.iISSIncluso <> 0
        
        Case CONTACONTABILAPLIC
        
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183772
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183773
        
                    If Len(Trim(objProduto.sContaContabil)) <> 0 Then
                    
                        sContaMascarada = String(STRING_CONTA, 0)
    
                        lErro = Mascara_MascararConta(objProduto.sContaContabil, sContaMascarada)
                        If lErro <> SUCESSO Then gError 183798
    
                        objMnemonicoValor.colValor.Add sContaMascarada
                    
                    Else
                    
                        objMnemonicoValor.colValor.Add ""
                    
                    End If
                    
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case CONTACONTABILPROD
        
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183772
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183773
        
                    If Len(Trim(objProduto.sContaContabilProducao)) <> 0 Then
                            
                        sContaMascarada = String(STRING_CONTA, 0)
    
                        lErro = Mascara_MascararConta(objProduto.sContaContabilProducao, sContaMascarada)
                        If lErro <> SUCESSO Then gError 183798
    
                        objMnemonicoValor.colValor.Add sContaMascarada
                        
                    Else
                    
                        objMnemonicoValor.colValor.Add ""
                    
                    End If
                    
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case CONTACONTABILCATPROD
                
            For iLinha = 1 To objGridItens.iLinhasExistentes
                
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183774
                
                    objProdutoCategoria.sCategoria = objMnemonicoValor.vParam(1)
                    
                    objProdutoCategoria.sProduto = sProdutoFormatado
                    
                    Set objCatProdItem = New ClassCategoriaProdutoItem
                    
                    lErro = CF("ProdutoCategoria_Le_Item", objProdutoCategoria, objCatProdItem)
                    If lErro <> SUCESSO And lErro <> 183548 Then gError 183775
                    
                    If lErro = SUCESSO Then
                        
                        Set objPlanoConta = New ClassPlanoConta
                    
                        'le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta
                        lErro = CF("PlanoConta_Le_ContaSimples", objCatProdItem.dvalor1, objPlanoConta)
                        If lErro <> SUCESSO And lErro <> 5451 Then gError 183776
                    
                        If lErro = SUCESSO Then
                    
                            sContaMascarada = String(STRING_CONTA, 0)
        
                            lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
                            If lErro <> SUCESSO Then gError 183799
        
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
        
        Case Else
            lErro = gobjTribTab.Calcula_Mnemonico(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 39663

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39663
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 64214, 64222, 64223, 64224, 64225, 64226, 64227, 64228, 64304, 79010, 79011, 79012, 88029, 39586, 56500, 183772 To 183776, 183798, 183799, 186138, 186139
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157398)

    End Select

    Exit Function

End Function

Public Sub MensagemLabel_Click()

Dim objMensagem As New ClassMensagem
Dim colSelecao As New Collection

    'Coloca a mensagem da tela no objMensagem
    objMensagem.sDescricao = Mensagem.Text
        
    'Chama a tela de browse de Mensagem
    Call Chama_Tela("MensagemLista", colSelecao, objMensagem, objEventoMensagem)

    Exit Sub

End Sub

Private Sub objEventoMensagem_evSelecao(obj1 As Object)

Dim objMensagem As ClassMensagem

    Set objMensagem = obj1
    
    Mensagem.Text = objMensagem.sDescricao
    
    Me.Show

End Sub

Public Sub VolumeQuant_GotFocus()
    Call MaskEdBox_TrataGotFocus(VolumeQuant, iAlterado)
End Sub

Public Sub DataReferencia_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dtDataEmissao As Date
Dim dtDataReferencia As Date
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_DataReferencia_Validate

    If giDataReferenciaAlterada <> REGISTRO_ALTERADO Then Exit Sub
    
    If Len(Trim(DataReferencia.ClipText)) > 0 Then

        'Critica a data digitada
        lErro = Data_Critica(DataReferencia.Text)
        If lErro <> SUCESSO Then gError 26713

        'Compara com data de emissão
        If Len(Trim(DataEmissao.ClipText)) > 0 Then

            dtDataEmissao = CDate(DataEmissao.Text)
            dtDataReferencia = CDate(DataReferencia.Text)

            lErro = CF("NF_Valida_Data_Referencia", dtDataReferencia, dtDataEmissao)
            If lErro <> SUCESSO Then gError 140328

        End If


    End If

    giDataReferenciaAlterada = 0
        
    'Preenche o GridParcelas
    lErro = Cobranca_Automatica()
    If lErro <> SUCESSO Then gError 25436

    Exit Sub

Erro_DataReferencia_Validate:

    Cancel = True
    
    Select Case gErr

        Case 25436 'Tratado na rotina chamada

        Case 26713, 140328

        Case 26714
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAREFERENCIA", gErr, dtDataReferencia, dtDataEmissao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157399)

    End Select

    Exit Sub

End Sub

Function DescontoPadraoParcela_Gera(objParcela As Object, dtDataVencimento As Date) As Long
'Preenche o objParcela com o desconto padrão
'*Obs: O objParcela foi declarado como Object para que a função sirva para Parcelasde Pedido de Venda e Parcelas de Notas Fiscais
                        
Dim colDesconto As colDesconto
Dim lErro As Long
Dim iIndice2 As Integer

On Error GoTo Erro_DescontoPadraoParcela_Gera
                        
    Set colDesconto = New colDesconto
    'Gera os desconto com base nas novas datas de vencimento
    lErro = CF("Parcela_GeraDescontoPadrao", colDesconto, dtDataVencimento)
    If lErro <> SUCESSO Then gError 51613
    With objParcela
        'Limpa no obj os decontos antigos
        .dDesconto1Valor = 0
        .dDesconto2Valor = 0
        .dDesconto3Valor = 0
        .dtDesconto1Ate = DATA_NULA
        .dtDesconto2Ate = DATA_NULA
        .dtDesconto3Ate = DATA_NULA
        .iDesconto1Codigo = 0
        .iDesconto2Codigo = 0
        .iDesconto3Codigo = 0
        
        'Preenche o obj comos novos descontos
        For iIndice2 = 1 To colDesconto.Count
            If iIndice2 = 1 Then
                .iDesconto1Codigo = colDesconto(iIndice2).iCodigo
                .dDesconto1Valor = colDesconto(iIndice2).dValor
                .dtDesconto1Ate = colDesconto(iIndice2).dtData
            ElseIf iIndice2 = 2 Then
                .iDesconto2Codigo = colDesconto(iIndice2).iCodigo
                .dDesconto2Valor = colDesconto(iIndice2).dValor
                .dtDesconto2Ate = colDesconto(iIndice2).dtData
            Else
                .iDesconto3Codigo = colDesconto(iIndice2).iCodigo
                .dDesconto3Valor = colDesconto(iIndice2).dValor
                .dtDesconto3Ate = colDesconto(iIndice2).dtData
            End If
        Next
    End With
    
    DescontoPadraoParcela_Gera = SUCESSO
    
    Exit Function
    
Erro_DescontoPadraoParcela_Gera:

    DescontoPadraoParcela_Gera = gErr
    
    Select Case gErr
        
        Case 51613
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157400)
            
    End Select

    Exit Function
    
End Function

Function ParcelasPV_Gerando_ParcelasNF(objPedidoVenda As ClassPedidoDeVenda, objNFiscal As ClassNFiscal) As Long
'Se a Data de referencia do Pedido estiver ultrapassada em regalação a Data Atual e se
'a cobrança automática do pedido estiver selecionada as  parcelas são atualizadas.
    
Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
'Dim colDataVencimento As New Collection
Dim iIndice As Integer
Dim objParcelaPV As New ClassParcelaPedidoVenda
Dim objParcRec As ClassParcelaReceber
Dim dtDataVencimento As Date
Dim dValorParcST As Double, dtDataParcST As Date
Dim objParcAux1 As ClassCondicaoPagtoParc, objParcAux2 As ClassCondicaoPagtoParc, colParcAux As New Collection
Dim bAdicionouST As Boolean

On Error GoTo Erro_ParcelasPV_Gerando_ParcelasNF
    
    'Se a condição de pagamento estiver preenchida e cobrança automática selecionada e a data de referencia do pedido estiver vazia ou ultrapassada
    If objPedidoVenda.iCondicaoPagto <> 0 And objPedidoVenda.iCobrancaAutomatica = 1 And (objPedidoVenda.dtDataReferencia = DATA_NULA Or objNFiscal.dtDataReferencia < gdtDataAtual) Then
        'A data de referencia da nota fiscal é atualizada
        objNFiscal.dtDataReferencia = gdtDataAtual
        
        objCondicaoPagto.iCodigo = objPedidoVenda.iCondicaoPagto
        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 59095
        If lErro = 19205 Then gError 59096
        
        If objPedidoVenda.objTributacao.dICMSSubstValor > 0 And gobjFAT.iParcRecCobraSTNovaParc = MARCADO Then
            objCondicaoPagto.iNumeroParcelas = objCondicaoPagto.iNumeroParcelas + 1
        End If
        
        If objCondicaoPagto.iNumeroParcelas <> objPedidoVenda.colParcelas.Count Then gError 59097
        
        'Calcula Datas de Vencimento das Parcelas
        'lErro = CF("Parcelas_DatasVencimento", objCondicaoPagto, objNFiscal.dtDataReferencia, colDataVencimento)
        objCondicaoPagto.dtDataRef = objNFiscal.dtDataReferencia
        objCondicaoPagto.dtDataEntrega = StrParaDate(GridItens.TextMatrix(1, iGrid_DataEntrega_Col))
        If objCondicaoPagto.dtDataEntrega = DATA_NULA Then objCondicaoPagto.dtDataEntrega = StrParaDate(DataSaida.Text)
        If objCondicaoPagto.dtDataEntrega = DATA_NULA Then objCondicaoPagto.dtDataEntrega = StrParaDate(DataEmissao.Text)
        'objCondicaoPagto.dtDataEmissao = StrParaDate(DataEmissao.Text)
        
        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True)
        If lErro <> SUCESSO Then gError 59098
        
        If gobjFAT.iParcRecCobraSTNovaParc = MARCADO Then
            dValorParcST = objPedidoVenda.objTributacao.dICMSSubstValor
            dtDataParcST = DateAdd("d", gobjFAT.iParcSTNumDias, objCondicaoPagto.dtDataRef)
            
            If dValorParcST > 0 Then
                bAdicionouST = False
                For Each objParcAux1 In objCondicaoPagto.colParcelas
                    If objParcAux1.dtVencimento > dtDataParcST And Not bAdicionouST Then
                        Set objParcAux2 = New ClassCondicaoPagtoParc
                        objParcAux2.dtVencimento = dtDataParcST
                        colParcAux.Add objParcAux2
                        bAdicionouST = True
                    End If
                    colParcAux.Add objParcAux1
                Next
                If Not bAdicionouST Then
                    Set objParcAux2 = New ClassCondicaoPagtoParc
                    objParcAux2.dtVencimento = dtDataParcST
                    colParcAux.Add objParcAux2
                    bAdicionouST = True
                End If
                Set objCondicaoPagto.colParcelas = colParcAux
            End If
        End If

        'Loop de preenchimento de Parcelas com Datas de Vencimento Recalculadas
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas

            Set objParcelaPV = objPedidoVenda.colParcelas.Item(iIndice)
            
            dtDataVencimento = objCondicaoPagto.colParcelas(iIndice).dtVencimento
            
            'Preenche o desconto padrão
            lErro = DescontoPadraoParcela_Gera(objParcelaPV, dtDataVencimento) 'colDataVencimento(iIndice))
            If lErro <> SUCESSO Then gError 51613
                        
            With objParcelaPV
                'Adiciona a parcela na coleção de parcelas
                Set objParcRec = objNFiscal.ColParcelaReceber.Add(0, 0, .iNumParcela, STATUS_ABERTO, dtDataVencimento, dtDataVencimento, .dValor, .dValor, 1, CARTEIRA_CARTEIRA, COBRADOR_PROPRIA_EMPRESA, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0, 0, "", .dValor)
                
                objParcRec.iTipoPagto = .iTipoPagto
                objParcRec.iCodConta = .iCodConta
                objParcRec.iBancoCheque = .iBancoCheque
                objParcRec.sAgenciaCheque = .sAgenciaCheque
                objParcRec.sContaCorrenteCheque = .sContaCorrenteCheque
                objParcRec.lNumeroCheque = .lNumeroCheque
                objParcRec.dtDataDepositoCheque = .dtDataDepositoCheque
                objParcRec.dtDataEmissaoCheque = .dtDataEmissaoCheque
                objParcRec.iAdmMeioPagto = .iAdmMeioPagto
                objParcRec.iParcelamento = .iParcelamento
                objParcRec.dtValidadeCartao = .dtValidadeCartao
                objParcRec.sNumeroCartao = .sNumeroCartao
                objParcRec.sAprovacaoCartao = .sAprovacaoCartao
            End With

        Next

    Else
    
        'transfere as parcelas do pedido de vendas para a nota fiscal
        For Each objParcelaPV In objPedidoVenda.colParcelas

            With objParcelaPV
                Set objParcRec = objNFiscal.ColParcelaReceber.Add(0, 0, .iNumParcela, STATUS_ABERTO, .dtDataVencimento, .dtDataVencimento, .dValor, .dValor, 1, CARTEIRA_CARTEIRA, COBRADOR_PROPRIA_EMPRESA, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0, 0, "", .dValor)
                
                objParcRec.iTipoPagto = .iTipoPagto
                objParcRec.iCodConta = .iCodConta
                objParcRec.dtDataCredito = .dtDataCredito
                objParcRec.iBancoCheque = .iBancoCheque
                objParcRec.sAgenciaCheque = .sAgenciaCheque
                objParcRec.sContaCorrenteCheque = .sContaCorrenteCheque
                objParcRec.lNumeroCheque = .lNumeroCheque
                objParcRec.dtDataDepositoCheque = .dtDataDepositoCheque
                objParcRec.dtDataEmissaoCheque = .dtDataEmissaoCheque
                objParcRec.iAdmMeioPagto = .iAdmMeioPagto
                objParcRec.iParcelamento = .iParcelamento
                objParcRec.dtValidadeCartao = .dtValidadeCartao
                objParcRec.sNumeroCartao = .sNumeroCartao
                objParcRec.sAprovacaoCartao = .sAprovacaoCartao
                objParcRec.dtDataTransacaoCartao = .dtDataTransacaoCartao
                
            End With

        Next
    
    End If
    
    ParcelasPV_Gerando_ParcelasNF = SUCESSO
    
    Exit Function

Erro_ParcelasPV_Gerando_ParcelasNF:

    ParcelasPV_Gerando_ParcelasNF = gErr
    
    Select Case gErr
        
        Case 59095, 59098, 51613
        
        Case 59096
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
        
        Case 59097
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_ALTERADA_NUM_PARC", gErr, objCondicaoPagto.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157401)
    
    End Select
    
    Exit Function
    
End Function

Public Sub BotaoDataReferenciaDown_Click()

Dim lErro As Long
Dim bCancel As Boolean
Dim sData As String

On Error GoTo Erro_BotaoDataReferenciaDown_Click

    sData = DataReferencia.Text

    'diminui a data em um dia
    lErro = Data_Up_Down_Click(DataReferencia, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 26715

    Call DataReferencia_Validate(bCancel)
    
    If bCancel = True Then DataReferencia.Text = sData
    
    Exit Sub

Erro_BotaoDataReferenciaDown_Click:

    Select Case gErr

        Case 26715
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157402)

    End Select

    Exit Sub

End Sub

Public Sub BotaoDataReferenciaUp_Click()

Dim lErro As Long
Dim sData As String
Dim bCancel As Boolean

On Error GoTo Erro_BotaoDataReferenciaUp_Click

    sData = DataReferencia.Text

    'aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataReferencia, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 26716
    
    Call DataReferencia_Validate(bCancel)

    If bCancel = True Then DataReferencia.Text = sData
    
    Exit Sub

Erro_BotaoDataReferenciaUp_Click:

    Select Case gErr

        Case 26716

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157403)

    End Select

    Exit Sub

End Sub

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

Sub AlocacoesNF_Agrupa(colAlocacoes As ColAlocacoesItemNF)

Dim iIndice As Integer
Dim iIndice1 As Integer

    If colAlocacoes.Count <= 1 Then Exit Sub
    
    'se a alocação se referir a uma transferencia ==> não leva em consideração
    For iIndice = colAlocacoes.Count To 1 Step -1
        If colAlocacoes.Item(iIndice).iTransferencia = TIPOMOV_EST_TRANSFERENCIA Then
            colAlocacoes.Remove iIndice
        End If
    Next
    
    For iIndice = colAlocacoes.Count To 2 Step -1
        
        For iIndice1 = 1 To iIndice - 1
            If (colAlocacoes.Item(iIndice).iAlmoxarifado = colAlocacoes.Item(iIndice1).iAlmoxarifado) Then
                colAlocacoes.Item(iIndice1).dQuantidade = colAlocacoes.Item(iIndice1).dQuantidade + colAlocacoes.Item(iIndice).dQuantidade
                colAlocacoes.Remove iIndice
                Exit For
            End If
        Next
    
    Next

End Sub

Public Sub BotaoConsultaTitRec_Click()
'Abre uma tela para consulta do DocCPR vinculado à nota fiscal

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoConsultaTitRec_Click

    'Verifica se todos os campos necessários para se efetuar a consulta foram preenchidos
    lErro = Critica_CamposNecessarios_ConsultaTitulo()
    If lErro <> SUCESSO Then gError 79806
    
    'Guarda no objNFiscal os dados necessários para consultar o título
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.iTipoDocInfo = Codigo_Extrai(TipoNFiscal.Text)
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Caption)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.dtDataEntrada = DATA_NULA
        
    'Guarda em objFornecedor o nome reduzido do Fornecedor
    objCliente.sNomeReduzido = Cliente.Caption
    
    'Lê o código do Fornecedor a partir do nome reduzido obtido na tela
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 79807
    
    'Se não encontrou o fornecedor = > erro
    If lErro = 12348 Then gError 79808
    
    'Guarda no objNFiscal o código do fornecedor
    objNFiscal.lCliente = objCliente.lCodigo
    
    'Lê o NumIntDocCPR da NFiscal e exibe o documento gerado no CPR por essa nota
    lErro = CF("NFiscal_Consulta_DocCPR", objNFiscal)
    If lErro <> SUCESSO And lErro <> 79717 Then gError 79809
    
    'Se não encontrou a nota => erro
    If lErro = 79717 Then gError 79810
    
    Exit Sub
    
Erro_BotaoConsultaTitRec_Click:

    Select Case gErr
    
        Case 79806, 79809, 79807
        
        Case 79810
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA2", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada)
        
        Case 79808
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157404)
    
    End Select
    
    Exit Sub

End Sub

Private Function Critica_CamposNecessarios_ConsultaTitulo() As Long
'Verifica se os campos necessários para encontrar consultar um título gerado por uma NFFatEntrada foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposNecessarios_ConsultaTitulo

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.Caption)) = 0 Then gError 79811
    
    'Verifica se a filial do Fornecedor foi preenchida
    If Len(Trim(Filial.Caption)) = 0 Then gError 79812
    
    'Verifica se a Série foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 79813
    
    'Verifica se o Número da Nota foi preenchido
    If Len(Trim(NFiscal.Caption)) = 0 Then gError 79814
    
    'Verifica se a data de emissão da nota foi preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 79815
    
    Critica_CamposNecessarios_ConsultaTitulo = SUCESSO
    
    Exit Function

Erro_Critica_CamposNecessarios_ConsultaTitulo:

    Critica_CamposNecessarios_ConsultaTitulo = gErr
    
    Select Case gErr
    
        Case 79811
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 79812
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 79813
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 79814
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
        
        Case 79815
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case Else
            
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157405)
    
    End Select
        
    Exit Function
        
End Function

Private Sub Move_NFiscal_Obj(objNFiscal As ClassNFiscal, tNFiscal As typeNFiscal)

    With tNFiscal

        objNFiscal.lNumIntDoc = .lNumIntDoc
        objNFiscal.iStatus = .iStatus
        objNFiscal.iFilialEmpresa = .iFilialEmpresa
        objNFiscal.iFilialEntrega = .iFilialEntrega
        objNFiscal.dtDataVencimento = .dtDataVencimento
        objNFiscal.dtDataReferencia = .dtDataReferencia
        objNFiscal.iFilialPedido = .iFilialPedido
        objNFiscal.lNumPedidoVenda = .lNumPedidoVenda
        objNFiscal.sNumPedidoTerc = .sNumPedidoTerc
        objNFiscal.iClasseDocCPR = .iClasseDocCPR
        objNFiscal.lNumIntDocCPR = .lNumIntDocCPR
        objNFiscal.dValorTotal = .dValorTotal
        objNFiscal.dValorProdutos = .dValorProdutos
        objNFiscal.dValorFrete = .dValorFrete
        objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscal.dValorDesconto = .dValorDesconto
        objNFiscal.iCodTransportadora = .iCodTransportadora
        objNFiscal.sMensagemNota = .sMensagemNota
        objNFiscal.iTabelaPreco = .iTabelaPreco
        objNFiscal.sNaturezaOp = .sNaturezaOp
        objNFiscal.dPesoLiq = .dPesoLiq
        objNFiscal.dPesoBruto = .dPesoBruto
        objNFiscal.lNumIntTrib = .lNumIntTrib
        objNFiscal.sPlaca = .sPlaca
        objNFiscal.sPlacaUF = .sPlacaUF
        objNFiscal.lVolumeQuant = .lVolumeQuant
        objNFiscal.lVolumeEspecie = .lVolumeEspecie
        objNFiscal.lVolumeMarca = .lVolumeMarca
        objNFiscal.sVolumeNumero = .sVolumeNumero
        objNFiscal.iCanal = .iCanal
        objNFiscal.lNumIntNotaOriginal = .lNumIntNotaOriginal
        objNFiscal.lClienteBenef = .lClienteBenef
        objNFiscal.iFilialCliBenef = .iFilialCliBenef
        objNFiscal.lFornecedorBenef = .lFornecedorBenef
        objNFiscal.iFilialFornBenef = .iFilialFornBenef
        objNFiscal.iFreteRespons = .iFreteRespons
        objNFiscal.lNumRecebimento = .lNumRecebimento
        objNFiscal.sObservacao = .sObservacao
        objNFiscal.sCodUsuarioCancel = .sCodUsuarioCancel
        objNFiscal.sMotivoCancel = .sMotivoCancel

    End With

End Sub

Public Sub DiretoIndireto_Change()

    'Desmarca o cálculo automático de comissões
'    ComissaoAutomatica.Value = vbUnchecked
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DiretoIndireto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComissoes)
End Sub

Public Sub DiretoIndireto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
End Sub

Public Sub DiretoIndireto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = DiretoIndireto
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub BotaoGrade_Click()

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim lErro  As Long
Dim objRomaneioGrade As ClassRomaneioGrade
Dim objItemNF As ClassItemNF

On Error GoTo Erro_BotaoGrade_Click

    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
    
        Set objItemNF = gobjNFiscal.ColItensNF(GridItens.Row)
        
        If objItemNF.iPossuiGrade = MARCADO Then
            
            Set objRomaneioGrade = New ClassRomaneioGrade
            
            objRomaneioGrade.sNomeTela = Me.Name
            Set objRomaneioGrade.objObjetoTela = objItemNF
                        
            Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
            If lErro <> SUCESSO Then gError 86360
        
            Call Atualiza_Grid_Itens(objItemNF)
            
            Call Atualiza_Grid_Alocacao(objItemNF)
        
            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) > 0 Then
                Call PrecoTotal_Calcula(GridItens.Row)
                
                lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
                If lErro <> SUCESSO Then gError 86360
                
                lErro = ValorTotal_Calcula()
                If lErro <> SUCESSO Then gError 51037
            End If
        
        End If
    
    End If
    
    Exit Sub

Erro_BotaoGrade_Click:

    Select Case gErr
    
        Case 51037, 86360
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157407)
            
    End Select
    
    Exit Sub

End Sub

Sub Atualiza_Grid_Itens(objItemNF As ClassItemNF)

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim dQuantidade As Double
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
    
    For Each objItemRomaneioGrade In objItemNF.colItensRomaneioGrade
        dQuantidade = dQuantidade + objItemRomaneioGrade.dQuantidade
    Next

    GridItens.TextMatrix(objItemNF.iItem, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)

    objItemNF.dQuantidade = dQuantidade
    
    Exit Sub

End Sub

Sub Atualiza_Grid_Alocacao(objItemNF As ClassItemNF)

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReserva As ClassReservaItem
Dim sProdutoMascarado As String
Dim lErro As Long
Dim dFator As Double
Dim dFator2 As Double
Dim objProduto As New ClassProduto
Dim dQuantReservada As Double
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Atualiza_Grid_Alocacao

    Call Remove_Alocacoes(objItemNF.iItem)

    For Each objItemRomaneio In objItemNF.colItensRomaneioGrade

        objProduto.sCodigo = objItemRomaneio.sProduto
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
        If lErro = 28030 Then gError 42765 'Não encontrou

        'Faz a conversão da unidade do item para a unidade de estoque
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRomaneio.sUMEstoque, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 42766

        objProduto.sCodigo = objItemNF.sProduto
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
        If lErro = 28030 Then gError 42765 'Não encontrou
        
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator2)
        If lErro <> SUCESSO Then gError 42766
        
        dQuantReservada = 0
        
        For Each objReserva In objItemRomaneio.colLocalizacao
        
            GridAlocacao.TextMatrix(objGridAlocacoes.iLinhasExistentes + 1, iGrid_Item_Col) = objItemNF.iItem
            
            lErro = Mascara_MascararProduto(objItemRomaneio.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 86357
                        
            objAlmoxarifado.iCodigo = objReserva.iAlmoxarifado

            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 46181
            If lErro = 25056 Then gError 46182
                
            objReserva.sAlmoxarifado = objAlmoxarifado.sNomeReduzido
                        
            GridAlocacao.TextMatrix(objGridAlocacoes.iLinhasExistentes + 1, iGrid_ProdutoAloc_Col) = sProdutoMascarado
            GridAlocacao.TextMatrix(objGridAlocacoes.iLinhasExistentes + 1, iGrid_UMAloc_Col) = objItemRomaneio.sUMEstoque
            GridAlocacao.TextMatrix(objGridAlocacoes.iLinhasExistentes + 1, iGrid_AlmoxAloc_Col) = objReserva.sAlmoxarifado
            GridAlocacao.TextMatrix(objGridAlocacoes.iLinhasExistentes + 1, iGrid_QuantVend_Col) = Formata_Estoque((objItemRomaneio.dQuantidade - objItemRomaneio.dQuantCancelada) * dFator2)
            GridAlocacao.TextMatrix(objGridAlocacoes.iLinhasExistentes + 1, iGrid_QuantAloc_Col) = Formata_Estoque(objReserva.dQuantidade * dFator)
            
            objGridAlocacoes.iLinhasExistentes = objGridAlocacoes.iLinhasExistentes + 1
            
        Next
               
    Next

    Exit Sub
    
Erro_Atualiza_Grid_Alocacao:

    Select Case gErr
    
        Case 42764, 42766, 46181, 86357
        
        Case 42765
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 46182
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157408)
            
    End Select
    
    Exit Sub

End Sub

Function Move_ItensGrade_Tela(colItensRomaneio As Collection, colItensRomaneioTela As Collection, Optional bTrazPedido As Boolean = False) As Long

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradeTela As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim objReservaItemTela As ClassReservaItem
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lErro As Long, iFilialEmpresa As Integer

On Error GoTo Erro_Move_ItensGrade_Tela

    'Para cada Item de Romaneio vindo da tela ( Aqueles que já tem quantidade)
    For Each objItemRomaneioGradeTela In colItensRomaneioTela
                    
        Set objItemRomaneioGrade = New ClassItemRomaneioGrade
            
        objItemRomaneioGrade.sProduto = objItemRomaneioGradeTela.sProduto
        objItemRomaneioGrade.dQuantOP = objItemRomaneioGradeTela.dQuantOP
        objItemRomaneioGrade.dQuantSC = objItemRomaneioGradeTela.dQuantSC
        objItemRomaneioGrade.dQuantPV = objItemRomaneioGradeTela.dQuantidade - objItemRomaneioGradeTela.dQuantCancelada - objItemRomaneioGradeTela.dQuantFaturada
        objItemRomaneioGrade.sDescricao = objItemRomaneioGradeTela.sDescricao
        objItemRomaneioGrade.sUMEstoque = objItemRomaneioGradeTela.sUMEstoque
        objItemRomaneioGrade.dQuantidade = objItemRomaneioGradeTela.dQuantidade - objItemRomaneioGradeTela.dQuantCancelada - objItemRomaneioGradeTela.dQuantFaturada
        objItemRomaneioGrade.dQuantAFaturar = objItemRomaneioGradeTela.dQuantAFaturar
        objItemRomaneioGrade.dQuantReservada = objItemRomaneioGradeTela.dQuantReservada
        If bTrazPedido Then
            objItemRomaneioGrade.lNumIntItemPV = objItemRomaneioGradeTela.lNumIntDoc
        Else
            objItemRomaneioGrade.lNumIntDoc = objItemRomaneioGradeTela.lNumIntDoc
            objItemRomaneioGrade.lNumIntItemPV = objItemRomaneioGradeTela.lNumIntItemPV
        End If
                            
        If objItemRomaneioGrade.dQuantidade > 0 Then
                            
            colItensRomaneio.Add objItemRomaneioGrade
        
            'Transfere as informações de Localização
            Set objItemRomaneioGrade.colLocalizacao = New Collection
                
            For Each objReservaItemTela In objItemRomaneioGradeTela.colLocalizacao
                
                objAlmoxarifado.iCodigo = objReservaItemTela.iAlmoxarifado
                            
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> 25056 And lErro <> SUCESSO Then gError 94331
                If lErro = 25056 Then gError 94332
                
                iFilialEmpresa = giFilialEmpresa
                lErro = CF("FilialEmpresaAlmox_Customiza", iFilialEmpresa)
                If lErro <> SUCESSO Then gError 94331
                
                If objAlmoxarifado.iFilialEmpresa = iFilialEmpresa Then
                
                    Set objReservaItem = New ClassReservaItem
                    
                    objReservaItem.dQuantidade = objReservaItemTela.dQuantidade
                    objReservaItem.dtDataValidade = objReservaItemTela.dtDataValidade
                    objReservaItem.iAlmoxarifado = objReservaItemTela.iAlmoxarifado
                    objReservaItem.iFilialEmpresa = objReservaItemTela.iFilialEmpresa
                    objReservaItem.lNumIntDoc = objReservaItemTela.lNumIntDoc
                    objReservaItem.sAlmoxarifado = objReservaItemTela.sAlmoxarifado
                    objReservaItem.sResponsavel = objReservaItemTela.sResponsavel
                    
                    objItemRomaneioGrade.colLocalizacao.Add objReservaItem
                End If
            Next
    
        End If
    
    Next
    
    Move_ItensGrade_Tela = SUCESSO
    
    Exit Function
    
Erro_Move_ItensGrade_Tela:

    Move_ItensGrade_Tela = gErr

    Select Case gErr
            
        Case 94331

        Case 94332
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objReservaItemTela.iAlmoxarifado)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157409)

    End Select
    
    Exit Function

End Function

Function Transfere_Dados_ItensRomaneio(colItensRomaneio As Collection, colItensRomaneioTela As Collection) As Long

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradeTela As ClassItemRomaneioGrade
Dim objReservaItemTela As ClassReservaItem
Dim objReservaItem As ClassReservaItem

    'Para cada Item de Romaneio existente do BD (Produtos Filhos do Produto passado)
    For Each objItemRomaneioGrade In colItensRomaneio
        'Para cada Item de Romaneio vindo da tela ( Aqueles que já tem quantidade)
        For Each objItemRomaneioGradeTela In colItensRomaneioTela
            'Se encontrou o Item
            If objItemRomaneioGrade.sProduto = objItemRomaneioGradeTela.sProduto Then
                'Transfere as informações vindas da tela chamadora para essa tela
                objItemRomaneioGrade.dQuantOP = objItemRomaneioGradeTela.dQuantOP
                objItemRomaneioGrade.dQuantSC = objItemRomaneioGradeTela.dQuantSC
                objItemRomaneioGrade.sDescricao = objItemRomaneioGradeTela.sDescricao
                objItemRomaneioGrade.dQuantAFaturar = objItemRomaneioGradeTela.dQuantAFaturar
                objItemRomaneioGrade.dQuantFaturada = objItemRomaneioGradeTela.dQuantFaturada
                objItemRomaneioGrade.dQuantidade = objItemRomaneioGradeTela.dQuantidade
                objItemRomaneioGrade.dQuantReservada = objItemRomaneioGradeTela.dQuantReservada
                objItemRomaneioGrade.sUMEstoque = objItemRomaneioGradeTela.sUMEstoque
                objItemRomaneioGrade.dQuantCancelada = objItemRomaneioGradeTela.dQuantCancelada
                objItemRomaneioGrade.dQuantPV = objItemRomaneioGradeTela.dQuantPV
                objItemRomaneioGrade.lNumIntItemPV = objItemRomaneioGradeTela.lNumIntItemPV
                objItemRomaneioGrade.lNumIntDoc = objItemRomaneioGradeTela.lNumIntDoc
                
                'Transfere as informações de Localização
                Set objItemRomaneioGrade.colLocalizacao = New Collection
                    
                For Each objReservaItemTela In objItemRomaneioGradeTela.colLocalizacao
                    
                    Set objReservaItem = New ClassReservaItem
                    
                    objReservaItem.dQuantidade = objReservaItemTela.dQuantidade
                    objReservaItem.dtDataValidade = objReservaItemTela.dtDataValidade
                    objReservaItem.iAlmoxarifado = objReservaItemTela.iAlmoxarifado
                    objReservaItem.iFilialEmpresa = objReservaItemTela.iFilialEmpresa
                    objReservaItem.lNumIntDoc = objReservaItemTela.lNumIntDoc
                    objReservaItem.sAlmoxarifado = objReservaItemTela.sAlmoxarifado
                    objReservaItem.sResponsavel = objReservaItemTela.sResponsavel
                    
                    objItemRomaneioGrade.colLocalizacao.Add objReservaItem
                    
                Next
                            
            End If
        
        Next
    Next

    Exit Function

End Function

Public Sub TranspRedespacho_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TranspRedespacho_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TranspRedespacho_Validate(Cancel As Boolean)
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTransportadora As New ClassTransportadora
Dim iCodigo As Integer

On Error GoTo Erro_TranspRedespacho_Validate

    'Verifica se foi preenchida a ComboBox Transportadora
    If Len(Trim(TranspRedespacho.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Transportadora
    If TranspRedespacho.Text = TranspRedespacho.List(TranspRedespacho.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TranspRedespacho, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 19243

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 19244
        
        'Não encontrou Transportadora no BD
        If lErro <> SUCESSO Then gError 19245
        
        'Encontrou Transportadora no BD, coloca no Text da Combo
        TranspRedespacho.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 19246

    Exit Sub

Erro_TranspRedespacho_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 19243, 19244

        Case 19245  'Não encontrou Transportadora no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            End If

        Case 19246
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157410)

    End Select

    Exit Sub

End Sub

Public Sub TranspRedLabel_Click()
'BROWSE TRANSPORTADORA :

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As New Collection
    
    'Se a Transportadora estiver preenchida passa o código para o objTransportadora
    If Len(Trim(TranspRedespacho.Text)) > 0 Then objTransportadora.iCodigo = Codigo_Extrai(TranspRedespacho.Text)
    
    'Chama a tela que lista as transportadoras
    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTranspRedespacho)

End Sub

Private Sub objEventoTranspRedespacho_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora
Dim bCancel As Boolean

    Set objTransportadora = obj1

    'Preenche campo Transportadora
    TranspRedespacho.Text = objTransportadora.iCodigo

    'Chama a rotina de validate
    Call TranspRedespacho_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub RedespachoCli_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

'@@@@@ para ser customizada
Public Function NFFP_Form_Load2(ByVal objCT As Object) As Long
    NFFP_Form_Load2 = SUCESSO
End Function

Function NFFP_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long
    NFFP_Form_Unload = SUCESSO
End Function

'@@@@@ para ser customizada
Function NFFP_CargaPosFormLoad(ByVal objCT As Object, ByVal bTrazendoDoc As Boolean) As Long
    NFFP_CargaPosFormLoad = SUCESSO
End Function

'@@@@@ com customizacao
Function NFFP_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iCaminho As Integer, ByVal sProdutoFormatado As String, ByVal iProdutoPreenchido As Integer) As Long
Dim lErro As Long

On Error GoTo Erro_NFFP_Rotina_Grid_Enable

    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
    
        Case objCT.PrecoUnitario.Name
            If iProdutoPreenchido = PRODUTO_VAZIO Or objCT.gcolItemPedido.Count = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
    End Select
    
    NFFP_Rotina_Grid_Enable = SUCESSO
     
    Exit Function
    
Erro_NFFP_Rotina_Grid_Enable:

    NFFP_Rotina_Grid_Enable = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157411)
     
    End Select
     
    Exit Function

End Function

'@@@@@ para ser customizada
Function NFFP_Traz_NFiscal_Tela2(ByVal objCT As Object, objNFiscal As ClassNFiscal, lNumIntPedido As Long) As Long
    NFFP_Traz_NFiscal_Tela2 = SUCESSO
End Function

'@@@@@ customizada
Function NFFP_Preenche_GridItem(ByVal objCT As Object, objNFiscal As ClassNFiscal, objItemNF As ClassItemNF, iIndice As Integer) As Long
    NFFP_Preenche_GridItem = SUCESSO
End Function

'@@@@@ para ser customizada
Function NFFP_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_NFFP_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col
        
        'Preco Unitário
        Case iGrid_ValorUnitario_Col
            lErro = Saida_Celula_PrecoUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 46931
    
    End Select
    
    NFFP_Saida_Celula_GridItens = SUCESSO
     
    Exit Function
    
Erro_NFFP_Saida_Celula_GridItens:

    NFFP_Saida_Celula_GridItens = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157412)
     
    End Select
     
    Exit Function

End Function

'@@@@@ para ser customizada
Function NFFP_Move_Tela_Memoria(ByVal objCT As Object, objNFiscal As ClassNFiscal) As Long
    NFFP_Move_Tela_Memoria = SUCESSO
End Function

'@@@@@ para ser customizada
Function NFFP_Move_GridItem_Memoria(ByVal objCT As Object, objNFiscal As ClassNFiscal, objItemNF As ClassItemNF, iIndice As Integer) As Long
    NFFP_Move_GridItem_Memoria = SUCESSO
End Function

Public Function Calcula_VolumeTotal() As Long
'Realiza a soma dos volumes dos produtos contido nop pedido de venda

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iIndice As Integer
Dim dFator As Double
Dim sProduto As String
Dim iPreenchido As Integer
Dim dVolume As Double

On Error GoTo Erro_Calcula_VolumeTotal

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        'Retira a formatação do produto
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 125880
        
        If iPreenchido <> PRODUTO_PREENCHIDO Then gError 125881
    
        'preenche o código do produto
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 125876
    
        If lErro = 28030 Then gError 125877
    
        'Realiza a converção para uma mesma UM
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 125878
    
        'Realiza o cálculo do volume
        dVolume = dVolume + ((StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)) * dFator) * (objProduto.dComprimento * objProduto.dEspessura * objProduto.dLargura))
    
    Next

    If dVolume <> 0 Then
        Cubagem.Text = dVolume
    Else
        Cubagem.Text = ""
    End If

    Calcula_VolumeTotal = SUCESSO

    Exit Function

Erro_Calcula_VolumeTotal:

    Calcula_VolumeTotal = gErr

    Select Case gErr

        Case 125876, 125878, 125880

        Case 125877
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case 125881
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157413)

    End Select

    Exit Function
    
End Function

Sub Cubagem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Sub Cubagem_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_Cubagem_Validate

    'Verifica se foi preenchido
    If Len(Trim(Cubagem.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(Cubagem.Text)
    If lErro <> SUCESSO Then gError 46529

    'Coloca o valor formatado na Tela
    Cubagem.Text = Format(Cubagem.Text, "Fixed")

    Exit Sub

Erro_Cubagem_Validate:

    Cancel = True

    Select Case gErr

        Case 46529

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157414)

    End Select

    Exit Sub

End Sub

'@@@@@ para ser customizada
Public Function NFFP_Transfere_Dados_Pedido_NFiscal2(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objNFiscal As ClassNFiscal, ByVal bFaturadoParc As Boolean) As Long
    NFFP_Transfere_Dados_Pedido_NFiscal2 = SUCESSO
End Function

Function NFFP_Botao_Estoque(ByVal objCT As Object, ByVal objEstoqueProd As ClassEstoqueProduto, ByVal colSelecao As Collection, sSelecao As String) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim iAlmoxarifado As Integer
Dim objProduto As New ClassProduto, sProdutoEnxuto As String

On Error GoTo Erro_NFFP_Botao_Estoque

    sSelecao = ""

    'Verifica se a linha do Grid é válida
    If GridItens.Row = 0 Then gError 131889

    'Verifica se o Produto na linha do Grid está preenchido
    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)

    If iProdutoPreenchido <> PRODUTO_VAZIO Then
        objEstoqueProd.sProduto = sProdutoFormatado

        objProduto.sCodigo = sProdutoFormatado
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 101782
        If lErro <> SUCESSO Then gError 101783
        
        If Len(Trim(objProduto.sGrade)) <> 0 Then
        
            Call Mascara_RetornaProdutoEnxuto(sProdutoFormatado, sProdutoEnxuto)
            sSelecao = "Produto LIKE '" & sProdutoEnxuto & "%' AND FilialEmpresa = " & giFilialEmpresa
            Call Chama_Tela("EstProdLista_Consulta", colSelecao, objEstoqueProd, objEventoEstoqueProd, sSelecao)
            gError 58774
            
        End If
        
        colSelecao.Add sProdutoFormatado

        'Lê o Almoxarifado Padrão do Produto
        lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, sProdutoFormatado, iAlmoxarifado)
        If lErro = SUCESSO Then objEstoqueProd.iAlmoxarifado = iAlmoxarifado

    Else
        gError 131890
    End If
 
    NFFP_Botao_Estoque = SUCESSO
    
    Exit Function

Erro_NFFP_Botao_Estoque:

    NFFP_Botao_Estoque = gErr

    Select Case gErr

        Case 101782, 58774
        
        Case 131889
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 131890
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 101783
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157415)

    End Select

    Exit Function
    
End Function

Private Function Saida_Celula_UM(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidade de Medida que está deixando de ser a corrente

Dim lErro As Long
Dim sUmAnterior As String
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim sProduto As String
Dim iPreenchido As Integer
Dim dPrecoUnitario As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objProduto As New ClassProduto
Dim objKitVenda As New ClassKitVenda

On Error GoTo Erro_Saida_Celula_UM

    Set objGridInt.objControle = UnidadeMed
    
    'recolhe a UM anteriormente escolhida
    sUmAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
    
    'coloca no grid a UM atual selecionda
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text
    
    gobjNFiscal.ColItensNF(GridItens.Row).sUnidadeMed = UnidadeMed.Text
    
    'se houve alteração na UM
    If sUmAnterior <> UnidadeMed.Text And Frame1(TAB_Almoxarifado).Enabled = True Then
        
        objProduto.sCodigo = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177583
        
        If gobjNFiscal.ColItensNF(GridItens.Row).iPossuiGrade <> MARCADO And objProduto.iKitVendaComp <> MARCADO Then
            
            lErro = Atualiza_UM(GridItens.Row, sUmAnterior, UnidadeMed.Text)
            If lErro <> SUCESSO Then gError 131927
            
            'Tenta fazer uma alocação p\ o item de acordo com a nova UM
            lErro = CF2(Me, "NFFP_Alocacao_Processa")
            If lErro <> SUCESSO Then gError 131901
    
    
        Else
        
            If objProduto.iKitVendaComp = MARCADO Then
                            
                objKitVenda.sProduto = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
                objKitVenda.sUM = UnidadeMed.Text
                objKitVenda.dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
            
                lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 177584
             
            End If
             
             For Each objItemRomaneio In gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade
                lErro = Alocacao_Processa_ItemGrade(objItemRomaneio, gobjNFiscal.ColItensNF(GridItens.Row), GridItens.Row)
                If lErro <> SUCESSO Then gError 131902
            Next
            
            Call Atualiza_Grid_Alocacao(gobjNFiscal.ColItensNF.Item(GridItens.Row))

        End If
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 131903
    
    'Se a Um selecionada agora é diferente da anterior
    If sUmAnterior <> UnidadeMed.Text Then

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 131904
    
    End If
    
    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case 131901 To 131904, 131927, 177583, 177584
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157416)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Alocacao_Processa_ItemGrade(objItemRomaneio As ClassItemRomaneioGrade, ByVal objItemNF As ClassItemNF, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objProdutoINF As New ClassProduto
Dim dQuantidade As Double
Dim sUM As String

On Error GoTo Erro_Alocacao_Processa_ItemGrade

    objProdutoINF.sCodigo = objItemNF.sProduto

    'Lê o Produto
    lErro = CF("Produto_Le", objProdutoINF)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 131905

    Set objItemRomaneio.colLocalizacao = New Collection

    objProduto.sCodigo = objItemRomaneio.sProduto

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 131905

    'Se não encontrou o Produto --> erro
    If lErro = 28030 Then gError 131906

    dQuantidade = objItemRomaneio.dQuantidade

   If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 And gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA Then

        sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)

        lErro = AlocaAlmoxarifadoPradrao_Grade(objProduto, objItemRomaneio, sUM, dQuantidade, objProdutoINF)
        If lErro <> SUCESSO Then gError 131907
    
    End If

    Alocacao_Processa_ItemGrade = SUCESSO

    Exit Function

Erro_Alocacao_Processa_ItemGrade:

    Alocacao_Processa_ItemGrade = gErr

    Select Case gErr

        Case 131905, 131907

        Case 131906
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157417)

    End Select

    Exit Function

End Function

Function AlocaAlmoxarifadoPradrao_Grade(objProduto As ClassProduto, objItemRomaneio As ClassItemRomaneioGrade, sUM As String, dQuantidade As Double, ByVal objProdutoINF As ClassProduto) As Long
'Tenta fazer a alocação do produto no almoxarifado padrão. Caso não consiga chama a tela de Alocação de produto.

Dim lErro As Long
Dim dQuantAlocar As Double
Dim dFator As Double
Dim iAlmoxarifado As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim dQuantDisponivel As Double
Dim objItemNF As ClassItemNF
Dim colOutrosProdutos As New Collection
Dim sProduto As String
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim objItemNFAloc As ClassItemNFAlocacao
Dim sProdutoEnxuto As String
Dim dAcrescimo As Double
Dim iNumCasasDec As Integer
Dim dTotal As Double
Dim iCodigo As Integer
Dim objReservaItem  As ClassReservaItem

On Error GoTo Erro_AlocaAlmoxarifadoPradrao_Grade

    If objProdutoINF.iKitVendaComp <> MARCADO Then
        'Faz a conversão da UM da tela para a UM de estoque
        lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 131908
    Else
        dFator = 1
    End If

    'Converte a quantidade para a UM de estoque
    dQuantAlocar = Arredonda_Estoque(dQuantidade * dFator) 'Alterado por Wagner

    'Busca o Almoxarifado padrão
    lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 131909

    'Se não encontrou
    If iAlmoxarifado = 0 Then gError 131910

    objAlmoxarifado.iCodigo = iAlmoxarifado
    
    'Lê o Aloxarifado
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> 25056 And lErro <> SUCESSO Then gError 131911
    If lErro = 25056 Then gError 131912

    objEstoqueProduto.sProduto = objProduto.sCodigo
    objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
    
    'Lê os estoques desse produto nesse almoxarifado
    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 21306 Then gError 131913
    
    'Se não achou o Estoque do Produto --> erro
    If lErro = 21306 Then gError 131914

    'A quantidade disponível deve ser igual a quantidade do escaninho mat. nosso disponível (DispNossa)
    dQuantDisponivel = objEstoqueProduto.dQuantDisponivel
    
    Set objItemRomaneio.colLocalizacao = New Collection

    dQuantAlocar = Arredonda_Estoque(dQuantAlocar)
    
    'Verifica se a Quantidade disponível é maior que a quantidade a alocar
    If (dQuantAlocar - dQuantDisponivel) < QTDE_ESTOQUE_DELTA Or gobjMAT.iAceitaEstoqueNegativo = MARCADO Then
        
        Set objReservaItem = New ClassReservaItem
        
        objReservaItem.iAlmoxarifado = objAlmoxarifado.iCodigo
        objReservaItem.sAlmoxarifado = objAlmoxarifado.sNomeReduzido
        objReservaItem.dQuantidade = dQuantAlocar
    
        objItemRomaneio.colLocalizacao.Add objReservaItem
    'Se não for
    Else
        'Se a função não estiver sendo usada pelo carregamento da tela
        'se prepara para chamar a tela de localização
        If Not gbCarregandoTela Then

            Set objItemNF = New ClassItemNF
            'Recolhe os dados do item
            objItemNF.iItem = GridItens.Row
            objItemNF.sProduto = objProduto.sCodigo
            objItemNF.sDescricaoItem = objProduto.sDescricao
            objItemNF.dQuantidade = dQuantidade
            objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
            objItemNF.iPossuiGrade = MARCADO
    
            'Recolhe todos os produtos dos outros itens
            For iIndice = 1 To objGridItens.iLinhasExistentes
                If iIndice <> GridItens.Row Then
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
                    If lErro <> SUCESSO Then gError 131915
                    'Adiciona na coleção de produtos
                    colOutrosProdutos.Add sProduto
                End If
            Next
    
            'Chama a tela de Localização de Produto
            Call Chama_Tela_Modal("LocalizacaoProduto1", objItemNF, colOutrosProdutos, dQuantAlocar, iCodigo)
            If giRetornoTela = vbCancel Then gError 131916
            If giRetornoTela = vbOK Then
                
                'Calcula o Total Alocado
                For Each objItemNFAloc In objItemNF.colAlocacoes
                    dTotal = dTotal + objItemNFAloc.dQuantidade
                Next
                
                'Para cada alocação feita para o item
                For Each objItemNFAloc In objItemNF.colAlocacoes
    
                    Set objReservaItem = New ClassReservaItem
                    
                    objReservaItem.sAlmoxarifado = objItemNFAloc.sAlmoxarifado
                    objReservaItem.dQuantidade = objItemNFAloc.dQuantidade
                    objReservaItem.iAlmoxarifado = objItemNFAloc.iAlmoxarifado
                    
                    objItemRomaneio.colLocalizacao.Add objReservaItem
                Next
                                                
            End If
        End If
    End If
    
    AlocaAlmoxarifadoPradrao_Grade = SUCESSO

    Exit Function

Erro_AlocaAlmoxarifadoPradrao_Grade:

    AlocaAlmoxarifadoPradrao_Grade = gErr

    Select Case gErr

        Case 131908, 131913, 131909, 131911, 131915

        Case 131912
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, iAlmoxarifado)

        Case 131914
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
        
        Case 131916
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_LOCALIZACAO", gErr, objProduto.sCodigo)

        Case 131910
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO1", gErr, objProduto.sCodigo, giFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157418)

    End Select

    Exit Function

End Function

Private Function Atualiza_UM(ByVal iLinha As Integer, ByVal sUmAnterior As String, ByVal sUMNova As String) As Long
'Atualiza quantidades em funcao de troca de UM

Dim lErro As Long
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double, dPrecoUnitario As Double
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Atualiza_UM

    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 131922

    objProduto.sCodigo = sProdutoFormatado

    'Lê o produto da linha passada por iLinha do GridItens
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 131923
    
    If lErro = 28030 Then gError 131924

    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUmAnterior, sUMNova, dFator)
    If lErro <> SUCESSO Then gError 131925

    'Atualiza o Grid
    GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col) = Formata_Estoque(Arredonda_Estoque(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) * dFator))

    'Recalcula preço do ítem e valor total da nota
    lErro = CF2(Me, "NFFP_Atualiza_PrecoUnitarioMoeda", iLinha, dFator)
    If lErro <> SUCESSO Then gError 131928
        
    If StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col)) > 0 Then
        Call PrecoTotal_Calcula(GridItens.Row, False)
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 131926
    End If

    Atualiza_UM = SUCESSO
    
    Exit Function

Erro_Atualiza_UM:

    Atualiza_UM = gErr

    Select Case gErr

        Case 131922, 131923, 131925, 131926, 131928

        Case 131924
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157419)

    End Select

    Exit Function

End Function

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

Public Function NFFP_Atualiza_PrecoUnitarioMoeda(ByVal objCT As Object, ByVal iLinha As Integer, ByVal dFator As Double) As Long
    
Dim dPrecoUnitario As Double
    
     'Preço unitário
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))
      
    GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col) = Format(Arredonda_Moeda(dPrecoUnitario / dFator), gobjFAT.sFormatoPrecoUnitario)
    
    NFFP_Atualiza_PrecoUnitarioMoeda = SUCESSO

End Function

Function NFFP_ValorDesconto_Validate(ByVal objCT As Object) As Long
    NFFP_ValorDesconto_Validate = SUCESSO
End Function

Public Function NFFP_Atualiza_PrecoUnitarioMoeda2(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    NFFP_Atualiza_PrecoUnitarioMoeda2 = SUCESSO
End Function

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long, sID As String
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Anotacao_Extrai

    'Lê os dados da Tela NFiscal
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 130404

    If objNFiscal.lNumNotaFiscal <> 0 And objNFiscal.sSerie <> "" And objNFiscal.dtDataEmissao <> DATA_NULA Then
    
        lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
        If lErro <> SUCESSO And lErro <> 43676 Then gError 130405
        
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157420)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157421)
     
    End Select
     
    Exit Sub

End Sub

Public Sub TipoNFiscal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Carrega_TiposDocInfo() As Long
'Carrega na combo os Tipo de Documentos relacionados com a tela

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim colTipoDocInfo As New colTipoDocInfo
Dim vCodTipoDocInfo As Variant

On Error GoTo Erro_Carrega_TiposDocInfo

    Set colTipoDocInfo = gobjCRFAT.colTiposDocInfo
    
    'Se for versão FULL
'    If giTipoVersao = VERSAO_FULL Then
        'Carrega na combo só os Tipos ligados essa tela
        For Each objTipoDocInfo In colTipoDocInfo
            If objTipoDocInfo.sNomeTelaNFiscal = Me.Name Then
                TipoNFiscal.AddItem CStr(objTipoDocInfo.iCodigo) & SEPARADOR & objTipoDocInfo.sNomeReduzido
                TipoNFiscal.ItemData(TipoNFiscal.NewIndex) = objTipoDocInfo.iCodigo
                'se for o tipo padrao, seleciona-o
                If objTipoDocInfo.iPadrao = Padrao Then
                    TipoNFiscal.ListIndex = TipoNFiscal.NewIndex
                    giTipoNFiscalPadrao = objTipoDocInfo.iCodigo
                End If
            End If
        Next

    'Se for versão LIGHT
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'        'Altera NomeReduzido de NotaFiscal de Venda e NFFatura de Venda
'        colTipoDocInfo(CStr(DOCINFO_NFISFV)).sNomeReduzido = "Venda (NFiscalFatura)"
'        colTipoDocInfo(CStr(DOCINFO_NFISV)).sNomeReduzido = "Venda (NFiscal)"
'
'        'Carrega os tipos da versão LIGHT
'        For Each objTipoDocInfo In colTipoDocInfo
'            For Each vCodTipoDocInfo In objUserControl.colCodTipoDocInfo
'
'                If objTipoDocInfo.iCodigo = vCodTipoDocInfo Then
'
'                    TipoNFiscal.AddItem CStr(objTipoDocInfo.iCodigo) & SEPARADOR & objTipoDocInfo.sNomeReduzido
'                    TipoNFiscal.ItemData(TipoNFiscal.NewIndex) = objTipoDocInfo.iCodigo
'                    'se for NFiscalFatura de Venda seleciona e muda NomeReduzido na tela
'                    If objTipoDocInfo.iCodigo = DOCINFO_NFISFV Then
'                        TipoNFiscal.ListIndex = TipoNFiscal.NewIndex
'                        giTipoNFiscalPadrao = objTipoDocInfo.iCodigo
'                    End If
'
'                    Exit For
'                End If
'
'            Next
'        Next
'
'    End If
    
    Carrega_TiposDocInfo = SUCESSO

    Exit Function

Erro_Carrega_TiposDocInfo:

    Carrega_TiposDocInfo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157162)

    End Select

    Exit Function

End Function

Public Sub TipoNFiscal_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TipoNFiscal_Click

    Frame1(TAB_Almoxarifado).Enabled = True

    If TipoNFiscal.ListIndex = -1 Then Exit Sub

    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    gobjRastreamento.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 42138

    'se não estiver cadastrado ==> erro
    If lErro = 31415 Then gError 42139
    
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then
        giComplemento = MARCADO
    Else
        giComplemento = DESMARCADO
    End If
    
    If objTipoDocInfo.iCodigo = DOCINFO_NFIFVETPV Then
        FrameFornEntTerc.Enabled = True
    Else
        FrameFornEntTerc.Enabled = False
        FornEntTerc.Text = ""
        FilialFornEntTerc.Clear
    End If
    
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Or (objTipoDocInfo.iTipoMovtoEstoque = 0 And objTipoDocInfo.iTipoMovtoEstoque = 0) Then
        Frame1(TAB_Almoxarifado).Enabled = False
    Else
        Frame1(TAB_Almoxarifado).Enabled = True
    End If
    
    If Not objGridAlocacoes Is Nothing Then Call Grid_Limpa(objGridAlocacoes)
    
    'Alterado por cyntia
    If objTipoDocInfo.iEscaninhoRastro = ESCANINHO_DESABILITADO Then
        EscaninhoRastro.Enabled = False
        gobjRastreamento.bEditavel = True
    Else
        EscaninhoRastro.Enabled = True
        gobjRastreamento.bEditavel = False
    End If
    
    'Inicializa o grid de Rastreamento
    lErro = gobjRastreamento.Inicializa_Grid_Rastreamento()
    If lErro <> SUCESSO Then gError 83383

    'se o cliente não estiver preenchido ==> preencher com a natureza padrao o campo natureza da operacao
    If Len(Trim(Cliente.Caption)) = 0 Then Call gobjTribTab.NatOpInterna_AlterarTrocaTipoNF(NaturezaOp_Conv34(objTipoDocInfo.sNaturezaOperacaoPadrao, MaskedParaDate(DataEmissao)))

    ''*** incluidos p/tratamento de tributacao *******************************
    If gbLimpandoTela = False Then
        If objTipoDocInfo.iCodigo = DOCINFO_NFIFVEFPV Then Call gobjTribTab.NatOpInterna_Automatica
        Call gobjTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
        If IsDate(DataEmissao.Text) Then Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
        Call gobjTribTab.Serie_Alterada(Serie.Text)
    End If
    '*** fim tributacao

    'subtipo contabil
    'ReInicializa a parte de contabilidade devido que o tipo implica na mudança dos modelos de contabilização
    lErro = objContabil.Contabil_Processa_Alteracao_Subtipo(objTipoDocInfo.iSubTipoContabil)
    If lErro <> SUCESSO Then gError 71690

    Exit Sub

Erro_TipoNFiscal_Click:

    Select Case gErr

        Case 42138, 71690, 83383

        Case 42139
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157182)

    End Select

    Exit Sub

End Sub

Public Sub TipoNFiscal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TipoNFiscal_Validate

    'Verifica se o tipo está preenchido
    If Len(Trim(TipoNFiscal.Text)) = 0 Then Exit Sub
    
    'Verifica se foi selecionado
    If TipoNFiscal.List(TipoNFiscal.ListIndex) = TipoNFiscal.Text Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(TipoNFiscal, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42140
    
    'Se não conseguir --> Erro
    If lErro <> SUCESSO Then gError 42141
    
    If iCodigo = DOCINFO_NFIFVETPV Then
        FrameFornEntTerc.Enabled = True
    Else
        FrameFornEntTerc.Enabled = False
        FornEntTerc.Text = ""
        FilialFornEntTerc.Clear
    End If

    Exit Sub

Erro_TipoNFiscal_Validate:

    Cancel = True

    Select Case gErr

        Case 42140

        Case 42141
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157183)

    End Select

    Exit Sub

End Sub

Private Sub Seleciona_Tipo_Venda()
'Seleciona na combo de tipos de notas fiscais o tipo "Venda"

Dim lErro As Long
'Dim objFilialEmpresa As New AdmFiliais
Dim iIndice As Integer ', iTipo As Integer

On Error GoTo Erro_Seleciona_Tipo_Venda

'    objFilialEmpresa.iCodFilial = giFilialEmpresa
'    objFilialEmpresa.lCodEmpresa = glEmpresa
'    'Lê a FilialEmpresa
'    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
'    If lErro <> SUCESSO And lErro <> 27378 Then gError ERRO_SEM_MENSAGEM
'
'    If objFilialEmpresa.iTipoTribAtividade = 3 Then
'        iTipo = DOCINFO_NFISFSPV
'    Else
'        iTipo = DOCINFO_NFISFVPV
'    End If

    For iIndice = 0 To TipoNFiscal.ListCount - 1
        If TipoNFiscal.ItemData(iIndice) = giTipoNFiscalPadrao Then 'iTipo Then  'DOCINFO_NFISFVPV Then
            TipoNFiscal.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Exit Sub

Erro_Seleciona_Tipo_Venda:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141800)
    
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

Public Function NFFP_Preenche_GridItens(ByVal objCT As Object, ByVal objItemNF As ClassItemNF, ByVal iLinha As Integer) As Long

Dim dValorDesconto As Double
Dim dValorTotal As Double

On Error GoTo Erro_NFFP_Preenche_GridItens

    dValorDesconto = objItemNF.dPercDesc * objItemNF.dQuantidade * objItemNF.dPrecoUnitario

    dValorTotal = (objItemNF.dPrecoUnitario * objItemNF.dQuantidade) - dValorDesconto

    objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_Desconto_Col) = Format(dValorDesconto, "Standard")
    objCT.GridItens.TextMatrix(iLinha, objCT.iGrid_PrecoTotal_Col) = Format(dValorTotal, "Standard")

    NFFP_Preenche_GridItens = SUCESSO

    Exit Function
    
Erro_NFFP_Preenche_GridItens:

    NFFP_Preenche_GridItens = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177080)

    End Select

    Exit Function

End Function

Public Function NFFP_Calcula_PesoKg(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    NFFP_Calcula_PesoKg = SUCESSO
End Function

Public Function NFFP_Calcula_PesoKg_Total(ByVal objCT As Object) As Long
    NFFP_Calcula_PesoKg_Total = SUCESSO
End Function

Public Function NFFP_PrecoTotal_Calcula(ByVal objCT As Object, ByVal iLinha As Integer, dQuantidade As Double) As Long
    NFFP_PrecoTotal_Calcula = SUCESSO
End Function

Public Function NFFP_Transfere_Dados_Pedido_NFiscal(ByVal objCT As Object, ByVal objItemNF As ClassItemNF, ByVal objItemPV As ClassItemPedido, ByVal dQuantidade As Double, ByVal objNFiscal As ClassNFiscal) As Long
    NFFP_Transfere_Dados_Pedido_NFiscal = SUCESSO
End Function

Private Function NotaFiscal_Imprime(Optional objNFiscal As ClassNFiscal) As Long
'Verifica se a nota em questão existe e, caso exista, dispara a impressão da mesma

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objCliente As ClassCliente
Dim objSerie As New ClassSerie

On Error GoTo Erro_NotaFiscal_Imprime

    'Transforma o ponteiro do mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se objNFiscal não foi passado como parâmetro
    If objNFiscal Is Nothing Then
    
        'Verifica se os campos obrigatórios estão preenchidos
        If Len(Trim(Cliente.Caption)) = 0 Then gError 177070
        If Len(Trim(Serie.Text)) = 0 Then gError 177071
        If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 177072
        If Len(Trim(NFiscal.Caption)) = 0 Then gError 177073
    
        'Instancia o obj
        Set objNFiscal = New ClassNFiscal
    
        'Guarda no obj os dados necessários para identificação da nota a ser impressa
        objNFiscal.iFilialEmpresa = giFilialEmpresa
        objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
        objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
        objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
        objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
        
        'Instancia objCliente, pois será necessário ler o cliente a partir do nome reduzido para obter o código
        Set objCliente = New ClassCliente
        
        'Guarda o nome reduzido do cliente no obj
        objCliente.sNomeReduzido = Trim(Cliente.Caption)
        
        'Lê o cliente a partir do nome reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 177074
        
        'Se não encontrou o cliente => erro
        If lErro = 12348 Then gError 177075
        
        'Guarda o código do cliente no obj
        objNFiscal.lCliente = objCliente.lCodigo
        
        'Guarda o código da filial do cliente
        objNFiscal.iFilialCli = Codigo_Extrai(Filial.Caption)
    
    End If
    
    'Lê os dados da nota fiscal para verificar se a mesma existe no BD
    lErro = CF("NFiscal_Le_1", objNFiscal)
    If lErro <> SUCESSO And lErro <> 83971 Then gError 177076

    'Se não encontrou => erro, pois não é possível imprimir uma nota fiscal inexistente
    If lErro = 83971 Then gError 177077
    
    objSerie.sSerie = objNFiscal.sSerie
    objSerie.iFilialEmpresa = objNFiscal.iFilialEmpresa
    
    'Lock na Tabela Série para a Impreessão
    lErro = CF("Serie_Lock_ImpressaoNFiscal", objSerie)
    If lErro <> SUCESSO And lErro <> 60387 Then gError 188763

    'Se não encontrou a Série --> ERRO
    If lErro = 60387 Then gError 188764
    
    'Dispara a impressão do relatório
    lErro = objRelatorio.ExecutarDireto("Emissão das Notas Fiscais Fatura", "", 1, objSerie.sNomeTsk, "NNFISCALINIC", objNFiscal.lNumNotaFiscal, "NNFISCALFIM", objNFiscal.lNumNotaFiscal, "TSERIE", objNFiscal.sSerie, "NBORDERO", "0")
    If lErro <> SUCESSO Then gError 177078

    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault
    
    NotaFiscal_Imprime = SUCESSO
    
    Exit Function

Erro_NotaFiscal_Imprime:

    GL_objMDIForm.MousePointer = vbDefault

    NotaFiscal_Imprime = gErr
    
    Select Case gErr
           
        Case 177070
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 177071
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
       
        Case 177072
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
            
        Case 177073
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMNOTAFISCAL_NAO_PREENCHIDO", gErr)
            
        Case 177074, 177076
        
        Case 177075
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
        
        Case 177077
            Call Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_NAO_ENCONTRADO", gErr)
            
        Case 177078 'Esse erro deve ser tratado na função chamadora, pois erro na impressão não deve interferir no funcionamento da tela
        
        Case 188764
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177079)
    
    End Select
    
End Function

Public Sub BotaoImprimir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoImprimir_Click

    'Dispara função para imprimir nota fiscal
    lErro = NotaFiscal_Imprime()
    If lErro <> SUCESSO And lErro <> 102262 Then gError 177068
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 177068
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 177069)

    End Select

    Exit Sub

End Sub

Public Sub BotaoKitVenda_Click()

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoKitVenda_Click

    If Me.ActiveControl Is Produto Then
        sProduto1 = Produto.Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 177579

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 177580
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoKitVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub
    
Erro_BotaoKitVenda_Click:

    Select Case gErr
    
        Case 177579
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 177580
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 177581)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub MensagemCorpo_Change()
    iAlterado = REGISTRO_ALTERADO
    MsgAutomatica.Value = vbUnchecked
End Sub

Public Sub Botao_Click(iIndice As Integer)

On Error GoTo Erro_Botao_Click
    
    Select Case iIndice
    
        Case BOTAO_PEDIDO
            Call BotaoPedido_Click
        
        Case BOTAO_VERPEDIDO
            Call BotaoVerPedido_Click
            
        Case BOTAO_IMPRIMIR
            Call BotaoImprimir_Click
        
        Case BOTAO_CONSULTATITREC
            Call BotaoConsultaTitRec_Click
        
        Case BOTAO_GRAVAR
            Call BotaoGravar_Click
        
        Case BOTAO_EXCLUIR
            Call BotaoExcluir_Click
        
        Case BOTAO_LIMPAR
            Call BotaoLimpar_Click
        
        Case BOTAO_FECHAR
            Call BotaoFechar_Click
        
        Case BOTAO_LIMPARNF
            Call BotaoLimparNF_Click
        
        Case BOTAO_PROJETOS
            Call BotaoProjetos_Click
        
        Case BOTAO_GRADE
            Call BotaoGrade_Click
        
        Case BOTAO_KITVENDA
            Call BotaoKitVenda_Click
        
        Case BOTAO_PRODUTOS
            Call BotaoProdutos_Click
        
        Case BOTAO_ESTOQUEPROD
            Call BotaoEstoqueProd_Click
        
        Case BOTAO_LOCALIZACAO
            Call BotaoLocalizacao_Click
        
        Case BOTAO_SERIE
            Call BotaoSerie_Click
        
        Case BOTAO_LOTES
            Call BotaoLotes_Click
            
        Case BOTAO_REFAZ_ALOC
            Call BotaoRefazAloc_Click
        
    End Select

    Exit Sub
    
Erro_Botao_Click:

    Select Case gErr

        Case Else
        
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181838)
    
    End Select
    
    Exit Sub

End Sub

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

        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objItensPRJCR.dValorOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        objItensPRJCR.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) & SEPARADOR & GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)

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

Public Sub TipoPagto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoPagto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub TipoPagto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub TipoPagto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = TipoPagto
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TipoPagto_Detalhe()

Dim objParcRec As ClassParcelaReceber
Dim objForm As Form
Dim objParcPV As New ClassParcelaPedidoVenda

On Error GoTo Erro_TipoPagto_Detalhe

    If TipoPagto.ListIndex <> -1 Then

        If TipoPagto.ItemData(TipoPagto.ListIndex) = TIPOPAGTO_DEPOSITO Then
        
            If gobjNFiscal.ColParcelaReceber.Count >= GridParcelas.Row Then
        
                Set objParcRec = gobjNFiscal.ColParcelaReceber.Item(GridParcelas.Row)
                
                objParcPV.iCodConta = objParcRec.iCodConta
                objParcPV.dtDataCredito = objParcRec.dtDataCredito
                
                Call Chama_Tela_Modal("DepositoConta", objParcPV)
                
                objParcRec.iCodConta = objParcPV.iCodConta
                objParcRec.dtDataCredito = objParcPV.dtDataCredito
                
            End If
    
        ElseIf TipoPagto.ItemData(TipoPagto.ListIndex) = TIPOPAGTO_CHEQUE Then
    
            If gobjNFiscal.ColParcelaReceber.Count >= GridParcelas.Row Then
            
                Set objParcRec = gobjNFiscal.ColParcelaReceber.Item(GridParcelas.Row)
                
                objParcPV.iBancoCheque = objParcRec.iBancoCheque
                objParcPV.sAgenciaCheque = objParcRec.sAgenciaCheque
                objParcPV.sContaCorrenteCheque = objParcRec.sContaCorrenteCheque
                objParcPV.lNumeroCheque = objParcRec.lNumeroCheque
                objParcPV.dtDataDepositoCheque = objParcRec.dtDataDepositoCheque
                objParcPV.dtDataEmissaoCheque = objParcRec.dtDataEmissaoCheque
                
                Call Chama_Tela_Modal("ChequePagto", objParcPV)
                
                objParcRec.iBancoCheque = objParcPV.iBancoCheque
                objParcRec.sAgenciaCheque = objParcPV.sAgenciaCheque
                objParcRec.sContaCorrenteCheque = objParcPV.sContaCorrenteCheque
                objParcRec.lNumeroCheque = objParcPV.lNumeroCheque
                objParcRec.dtDataDepositoCheque = objParcPV.dtDataDepositoCheque
                objParcRec.dtDataEmissaoCheque = objParcPV.dtDataEmissaoCheque
                
            End If
    
        ElseIf TipoPagto.ItemData(TipoPagto.ListIndex) = TIPOPAGTO_CARTAO Then
    
            If gobjNFiscal.ColParcelaReceber.Count >= GridParcelas.Row Then
            
                Set objParcRec = gobjNFiscal.ColParcelaReceber.Item(GridParcelas.Row)
                
                objParcPV.iAdmMeioPagto = objParcRec.iAdmMeioPagto
                objParcPV.iParcelamento = objParcRec.iParcelamento
                objParcPV.sNumeroCartao = objParcRec.sNumeroCartao
                objParcPV.dtValidadeCartao = objParcRec.dtValidadeCartao
                objParcPV.sAprovacaoCartao = objParcRec.sAprovacaoCartao
                objParcPV.dtDataTransacaoCartao = objParcRec.dtDataTransacaoCartao
                
                Call Chama_Tela_Modal("CartaoPagto", objParcPV)
                
                objParcRec.iAdmMeioPagto = objParcPV.iAdmMeioPagto
                objParcRec.iParcelamento = objParcPV.iParcelamento
                objParcRec.sNumeroCartao = objParcPV.sNumeroCartao
                objParcRec.dtValidadeCartao = objParcPV.dtValidadeCartao
                objParcRec.sAprovacaoCartao = objParcPV.sAprovacaoCartao
                objParcRec.dtDataTransacaoCartao = objParcPV.dtDataTransacaoCartao
                
            End If
    
        End If

    End If

    Exit Sub
    
Erro_TipoPagto_Detalhe:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183005)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_TipoPagto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo Pagto que está deixando de serr a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim iTipo As Integer
Dim iIndice As Integer
Dim objParcRec As ClassParcelaReceber

On Error GoTo Erro_Saida_Celula_TipoPagto

    Set objGridInt.objControle = TipoPagto

    'Verifica se o Tipo foi preenchido
    If Len(Trim(TipoPagto.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoPagto.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_TipoPagto_Col) Then

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona_Grid(objGridInt.objControle, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 183006

            'Não foi encontrado
            If lErro = 25085 Then gError 183007
            If lErro = 25086 Then gError 183008

            Call TipoPagto_Detalhe

        End If
        
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            Set objParcRec = New ClassParcelaReceber
            gobjNFiscal.ColParcelaReceber.AddObj objParcRec
            
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 183009

    Saida_Celula_TipoPagto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoPagto:

    Saida_Celula_TipoPagto = gErr

    Select Case gErr

        Case 183006, 183009
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 183007
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMAPAGTO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 183008
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMAPAGTO_NAO_ENCONTRADO", gErr, objGridInt.objControle.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183010)

    End Select

    Exit Function

End Function

Function NFiscal_Inicializa_Grid_Parcelas(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
    
    objGridInt.objGrid.ColWidth(iGrid_TipoPagto_Col) = 0
    objGridInt.objGrid.TextMatrix(0, iGrid_TipoPagto_Col) = ""
    BotaoTipoPagto.Visible = False
    
    NFiscal_Inicializa_Grid_Parcelas = SUCESSO
    
End Function

Public Sub BotaoTipoPagto_Click()

Dim lErro As Long
Dim lCodigo As Long
Dim iTipoPagto As Integer
Dim objParcRec As ClassParcelaReceber
Dim objParcPV As New ClassParcelaPedidoVenda

On Error GoTo Erro_BotaoTipoPagto_Click

    If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_TipoPagto_Col))) > 0 Then

        iTipoPagto = Codigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_TipoPagto_Col))

        If iTipoPagto = TIPOPAGTO_DEPOSITO Then
        
            If gobjNFiscal.ColParcelaReceber.Count >= GridParcelas.Row Then
        
                Set objParcRec = gobjNFiscal.ColParcelaReceber.Item(GridParcelas.Row)
                
                objParcPV.iCodConta = objParcRec.iCodConta
                objParcPV.dtDataCredito = objParcRec.dtDataCredito
                
                Call Chama_Tela_Modal("DepositoConta", objParcPV)
                
                objParcRec.iCodConta = objParcPV.iCodConta
                objParcRec.dtDataCredito = objParcPV.dtDataCredito
        
            End If
    
        ElseIf iTipoPagto = TIPOPAGTO_CHEQUE Then
    
            If gobjNFiscal.ColParcelaReceber.Count >= GridParcelas.Row Then
            
                Set objParcRec = gobjNFiscal.ColParcelaReceber.Item(GridParcelas.Row)
                
                objParcPV.iBancoCheque = objParcRec.iBancoCheque
                objParcPV.sAgenciaCheque = objParcRec.sAgenciaCheque
                objParcPV.sContaCorrenteCheque = objParcRec.sContaCorrenteCheque
                objParcPV.lNumeroCheque = objParcRec.lNumeroCheque
                objParcPV.dtDataDepositoCheque = objParcRec.dtDataDepositoCheque
                objParcPV.dtDataEmissaoCheque = objParcRec.dtDataEmissaoCheque
                
                Call Chama_Tela_Modal("ChequePagto", objParcPV)
                
                objParcRec.iBancoCheque = objParcPV.iBancoCheque
                objParcRec.sAgenciaCheque = objParcPV.sAgenciaCheque
                objParcRec.sContaCorrenteCheque = objParcPV.sContaCorrenteCheque
                objParcRec.lNumeroCheque = objParcPV.lNumeroCheque
                objParcRec.dtDataDepositoCheque = objParcPV.dtDataDepositoCheque
                objParcRec.dtDataEmissaoCheque = objParcPV.dtDataEmissaoCheque
                
            End If
    
        ElseIf iTipoPagto = TIPOPAGTO_CARTAO Then
    
            If gobjNFiscal.ColParcelaReceber.Count >= GridParcelas.Row Then
            
                Set objParcRec = gobjNFiscal.ColParcelaReceber.Item(GridParcelas.Row)
                
                objParcPV.iAdmMeioPagto = objParcRec.iAdmMeioPagto
                objParcPV.iParcelamento = objParcRec.iParcelamento
                objParcPV.sNumeroCartao = objParcRec.sNumeroCartao
                objParcPV.dtValidadeCartao = objParcRec.dtValidadeCartao
                objParcPV.sAprovacaoCartao = objParcRec.sAprovacaoCartao
                objParcPV.dtDataTransacaoCartao = objParcRec.dtDataTransacaoCartao
                
                Call Chama_Tela_Modal("CartaoPagto", objParcPV)
                
                objParcRec.iAdmMeioPagto = objParcPV.iAdmMeioPagto
                objParcRec.iParcelamento = objParcPV.iParcelamento
                objParcRec.sNumeroCartao = objParcPV.sNumeroCartao
                objParcRec.dtValidadeCartao = objParcPV.dtValidadeCartao
                objParcRec.sAprovacaoCartao = objParcPV.sAprovacaoCartao
                objParcRec.dtDataTransacaoCartao = objParcPV.dtDataTransacaoCartao
                
            End If
    
        End If
    
    End If

    Exit Sub

Erro_BotaoTipoPagto_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183011)

    End Select

    Exit Sub

End Sub

Public Function BotaoEntrega_Click() As Long
    Call CF2(Me, "NFFP_BotaoEntrega_Click")
End Function

Public Sub NaoImpCobranca_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

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

Public Function Traz_Pedido_Fluxo() As Long

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Traz_Pedido_Fluxo

    objPedidoVenda.lCodigo = CLng(Pedido.Text)
    objPedidoVenda.iFilialEmpresa = StrParaInt(FilialPedido.Text)
    
    'Traz os dados do pedido para a tela
    lErro = Traz_Dados_Pedido_Tela(objPedidoVenda)
    If lErro <> SUCESSO Then gError 188587
    
    Traz_Pedido_Fluxo = SUCESSO
    
    Exit Function

Erro_Traz_Pedido_Fluxo:

    Traz_Pedido_Fluxo = gErr

    Select Case gErr

        Case 188587

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188588)

    End Select

    Exit Function

End Function

Public Function NFFP_Saida_Celula_PrecoUnitario(ByVal objCT As Object, ByVal dPrecoUnitario As Double, ByVal iLinha As Integer) As Long
    NFFP_Saida_Celula_PrecoUnitario = SUCESSO
End Function

Public Function NFFP_Quantidade_GotFocus(ByVal objCT As Object) As Long
    NFFP_Quantidade_GotFocus = SUCESSO
End Function

Public Function NFFP_BotaoEntrega_Click(ByVal objCT As Object) As Long

Dim lErro As Long
Dim sProdutoTela As String
Dim dQuantidade As Double

On Error GoTo Erro_NFFP_BotaoEntrega_Click

    If GridItens.Row = 0 Then gError 183236

    sProdutoTela = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    If Len(sProdutoTela) = 0 Then gError 183237

    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))

    Call Chama_Tela_Modal("DataEntrega", sProdutoTela, dQuantidade, gobjNFiscal.ColItensNF(GridItens.Row).colDataEntrega)

    NFFP_BotaoEntrega_Click = SUCESSO

    Exit Function

Erro_NFFP_BotaoEntrega_Click:

    NFFP_BotaoEntrega_Click = gErr

    Select Case gErr

        Case 183236
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 183237
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183238)

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
'Carrega versaoNFE

Dim lErro As Long
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer
Dim colTipoFrete As Collection

On Error GoTo Erro_Carrega_TipoFrete
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207435

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

        'Erro já tratado
        Case 207435

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207436)

    End Select

    Exit Function

End Function

Public Sub ValorDescontoTit_Change()
    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorDescontoTit_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorDesconto As Double
Dim dValorPagar As Double

On Error GoTo Erro_ValorDescontoTit_Validate

    dValorDesconto = 0

    'Verifica se o Valor está preenchido
    If Len(Trim(ValorDescontoTit.Text)) > 0 Then
    
        'Faz a Crítica do Valor digitado
        lErro = Valor_NaoNegativo_Critica(ValorDescontoTit.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dValorDesconto = CDbl(ValorDescontoTit.Text)
        
        'Coloca o Valor formatado na tela
        ValorDescontoTit.Text = Format(dValorDesconto, "Standard")

    End If
    
    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ValorTit.Caption = Format(dValorPagar, "Standard")
    ValorOriginalTit.Caption = Format(dValorPagar + dValorDesconto, "Standard")
    
    If Not gbCarregandoTela Then
        lErro = Cobranca_Automatica()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    Exit Sub

Erro_ValorDescontoTit_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157330)

    End Select

    Exit Sub

End Sub

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
    iComissoesAlterada = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercDescontoItens_Change()
    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO
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
        dFator = dDesconto / dTotal
    
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
                Call PrecoTotal_Calcula(iIndice)
                
                lErro = gobjTribTab.Alteracao_Item_Grid(iIndice)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                dValorTotal = dValorTotal + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
            End If
        Next
    
    End If
        
    'Coloca valor total dos produtos na tela
    ValorProdutos.Caption = Format(dValorTotal, "Standard")

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

Public Sub FilialFornEntTerc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialFornEntTerc_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialFornEntTerc_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialFornEntTerc_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialFornEntTerc.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialFornEntTerc.Text = FilialFornEntTerc.List(FilialFornEntTerc.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialFornEntTerc, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(FornEntTerc.Text)) = 0 Then gError 209993

        sFornecedor = FornEntTerc.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError ERRO_SEM_MENSAGEM

        If lErro = 18272 Then gError 209994

        'coloca na tela
        FilialFornEntTerc.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 209995

    Exit Sub

Erro_FilialFornEntTerc_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209993
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 209994
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, FornEntTerc.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 209995
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, FilialFornEntTerc.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209996)

    End Select

    Exit Sub

End Sub

Public Sub FornEntTerc_Change()
    iAlterado = REGISTRO_ALTERADO
    iFornEntTercAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FornEntTerc_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornEntTerc As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_FornEntTerc_Validate
   
    If iFornEntTercAlterado = 1 Then
        If Len(Trim(FornEntTerc.Text)) > 0 Then
    
            'Tenta ler o FornEntTerc (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(FornEntTerc, objFornEntTerc, iCodFilial)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornEntTerc, colCodigoNome)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", FilialFornEntTerc, colCodigoNome)
    
            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", FilialFornEntTerc, iCodFilial)
    
    
        ElseIf Len(Trim(FornEntTerc.Text)) = 0 Then
            FilialFornEntTerc.Clear
        End If
    
        iFornEntTercAlterado = 0
    
    End If
    
    Exit Sub

Erro_FornEntTerc_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209997)

    End Select

    Exit Sub

End Sub

Public Sub FornEntTercLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = FornEntTerc.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornEntTerc)

    Exit Sub

End Sub

Private Sub objEventoFornEntTerc_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    FornEntTerc.Text = objFornecedor.lCodigo
    Call FornEntTerc_Validate(bSGECancelDummy)

    Me.Show

End Sub

Public Function NFFP_Limpa_Tela(ByVal objCT As Object) As Long
    NFFP_Limpa_Tela = SUCESSO
End Function

Public Sub BotaoRefazAloc_Click()
'Chama a tela de Lote de Rastreamento

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim iIndice As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objProduto As ClassProduto

On Error GoTo Erro_BotaoRefazAloc_Click

    'Limpa o grid de alocações
    Call Grid_Limpa(objGridAlocacoes)

    objGridAlocacoes.iLinhasExistentes = 0
    
    iIndice = 0
    For Each objItemNF In gobjNFiscal.ColItensNF
        
        iIndice = iIndice + 1
    
        If objItemNF.colItensRomaneioGrade.Count = 0 Then

            GridItens.Row = iIndice
            
            lErro = CF2(Me, "NFFP_Alocacao_Processa")
            If lErro <> SUCESSO And lErro <> 23796 And lErro <> 197701 And lErro <> 197703 Then gError ERRO_SEM_MENSAGEM

        Else
        
            'Faz a alocação de quem não tem
            For Each objItemRomaneio In objItemNF.colItensRomaneioGrade
                If objItemRomaneio.colLocalizacao.Count = 0 Then
                    lErro = Alocacao_Processa_ItemGrade(objItemRomaneio, objItemNF, iIndice)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                End If
            Next
            
            Call Atualiza_Grid_Alocacao(objItemNF)
            
            Set objProduto = New ClassProduto
            objProduto.sCodigo = objItemNF.sProduto
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

            If objProduto.iKitVendaComp <> MARCADO Then
                Call Atualiza_Grid_Itens(objItemNF)
            End If

        End If
        
    Next

    Exit Sub

Erro_BotaoRefazAloc_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157310)
    
    End Select
    
    Exit Sub

End Sub

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
        
        Call Chama_Tela_Modal("InfoAdicDocItem", objItemNF.objInfoAdicDocItem, objTela, TIPO_SAIDA, gobjInfoAdic)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Call PrecoTotal_Calcula(GridItens.Row, False)
        
        Call ValorTotal_Calcula

    End If
    
    Exit Sub

Erro_BotaoInfoAdicItem_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211245)
            
    End Select
    
    Exit Sub

End Sub

Public Sub ComissaoItemNF_Change()

    iAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO

End Sub

Public Sub ComissaoItemNF_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub ComissaoItemNF_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub ComissaoItemNF_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ComissaoItemNF
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ComissaoItemNF(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dComissaoItemNF As Double
Dim dPercentualCampo As Double
Dim dPercentualGrid As Double

On Error GoTo Erro_Saida_Celula_ComissaoItemNF

    Set objGridInt.objControle = ComissaoItemNF

    dPercentualGrid = PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ComissaoItem_Col))
    dPercentualGrid = dPercentualGrid * 100

    dPercentualCampo = StrParaDbl(ComissaoItemNF.Text)
    
    'Verifica se ComissaoItemNF foi preenchido
    If Len(Trim(ComissaoItemNF.Text)) > 0 Then

        'Critica porcentagem informada
        lErro = Porcentagem_Critica(ComissaoItemNF.Text)
        If lErro <> SUCESSO Then Error 49478

        dComissaoItemNF = StrParaDbl(ComissaoItemNF.Text)

        'Coloca no formato fixed da tela
        ComissaoItemNF.Text = Format(dComissaoItemNF, "Fixed")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49479

    Saida_Celula_ComissaoItemNF = SUCESSO

    Exit Function

Erro_Saida_Celula_ComissaoItemNF:

    Saida_Celula_ComissaoItemNF = Err

    Select Case Err

        Case 49478

        Case 49479
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164612)

    End Select

    Exit Function

End Function

Public Sub TabPrecoItemNF_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TabPrecoItemNF_Click()

    iAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO

End Sub

Public Sub TabPrecoItemNF_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub TabPrecoItemNF_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub TabPrecoItemNF_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = TabPrecoItemNF
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_TabPrecoItemNF(objGridInt As AdmGrid) As Long

Dim lErro As Long, iLinha As Integer, iItem As Integer
Dim iTabelaAnterior As Integer, iTabelaAtual As Integer

On Error GoTo Erro_Saida_Celula_TabPrecoItemNF

    Set objGridInt.objControle = TabPrecoItemNF

    iLinha = GridItens.Row
    
    For iItem = 0 To TabPrecoItemNF.ListCount - 1
        
        If GridItens.TextMatrix(iLinha, iGrid_TabPrecoItem_Col) = TabPrecoItemNF.List(iItem) Then

            iTabelaAnterior = TabPrecoItemNF.ItemData(iItem)
            Exit For

        End If
        
    Next
    
    GridItens.TextMatrix(GridItens.Row, iGrid_TabPrecoItem_Col) = TabPrecoItemNF.Text
    
    For iItem = 0 To TabPrecoItemNF.ListCount - 1
        
        If GridItens.TextMatrix(iLinha, iGrid_TabPrecoItem_Col) = TabPrecoItemNF.List(iItem) Then

            iTabelaAtual = TabPrecoItemNF.ItemData(iItem)
            Exit For

        End If
        
    Next
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49482
    
    If Not gbCarregandoTela And iTabelaAnterior <> iTabelaAtual Then
    
        lErro = Trata_TabelaPreco_Item(iLinha)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Calcula o Valor Total da Nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    Saida_Celula_TabPrecoItemNF = SUCESSO

    Exit Function

Erro_Saida_Celula_TabPrecoItemNF:

    Saida_Celula_TabPrecoItemNF = Err

    Select Case Err

        Case 49482, ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164625)

    End Select

    Exit Function

End Function

Private Function Trata_TabelaPreco_Item(ByVal iLinha As Integer) As Long

Dim lErro As Long, iCodTabela As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim sProduto As String
Dim iPreenchido As Integer, iItem As Integer

On Error GoTo Erro_Trata_TabelaPreco_Item

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 46663

    If iPreenchido <> PRODUTO_VAZIO And Len(Trim(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col))) > 0 Then
        
        If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        
             For iItem = 0 To TabPrecoItemNF.ListCount - 1
                
                If GridItens.TextMatrix(iLinha, iGrid_TabPrecoItem_Col) = TabPrecoItemNF.List(iItem) Then
        
                    iCodTabela = TabPrecoItemNF.ItemData(iItem)
                    Exit For
        
                End If
                
            Next
        
        End If
        
        If iCodTabela = 0 Then iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        
        objTabelaPrecoItem.sCodProduto = sProduto
        objTabelaPrecoItem.iCodTabela = iCodTabela
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
        'Lê a Tabela preço para filialEmpresa
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 46660
        'Se não encontrar
        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
            'Lê a Tabela de Preço a nível de Empresa toda
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 46661

        End If
        'Se  conseguir ler a Tabela de Preços
        If lErro = SUCESSO Then
            'Calcula o Preco Unitário do item
            lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 46662
            'Coloca no Grid
            If dPrecoUnitario > 0 Then
                GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            Else
                GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col) = ""
            End If
            'Calcula o Preco Total do item
            Call PrecoTotal_Calcula(iLinha, False)

            ''*** incluidos p/tratamento de tributacao *******************************
            lErro = gobjTribTab.Alteracao_Item_Grid(iLinha)
            If lErro <> SUCESSO Then gError 56860
            '*** fim tributacao
        
        End If

    End If

    Trata_TabelaPreco_Item = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco_Item:

    Trata_TabelaPreco_Item = gErr

    Select Case gErr

        Case 46660, 46661, 46662, 46663, 56860

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157457)

    End Select

    Exit Function

End Function

'########################################
'EMBALAGEM
Public Sub GridEmb_Click() 'embalagem
    Call gobjEmbalagem.GridEmb_Click
End Sub

Public Sub GridEmb_EnterCell() 'embalagem
    Call gobjEmbalagem.GridEmb_EnterCell
End Sub

Public Sub GridEmb_GotFocus() 'embalagem
    Call gobjEmbalagem.GridEmb_GotFocus
End Sub

Public Sub GridEmb_KeyPress(KeyAscii As Integer) 'embalagem
    Call gobjEmbalagem.GridEmb_KeyPress(KeyAscii)
End Sub

Public Sub GridEmb_LeaveCell() 'embalagem
    Call gobjEmbalagem.GridEmb_LeaveCell
End Sub

Public Sub GridEmb_Validate(Cancel As Boolean) 'embalagem
    Call gobjEmbalagem.GridEmb_Validate(Cancel)
End Sub

Public Sub GridEmb_RowColChange() 'embalagem
    Call gobjEmbalagem.GridEmb_RowColChange
End Sub

Public Sub GridEmb_KeyDown(KeyCode As Integer, Shift As Integer) 'embalagem
    Call gobjEmbalagem.GridEmb_KeyDown(KeyCode, Shift)
End Sub

Public Sub GridEmb_Scroll() 'embalagem
    Call gobjEmbalagem.GridEmb_Scroll
End Sub

Public Sub ItemEmb_KeyPress(KeyAscii As Integer) 'embalagem
    Call gobjEmbalagem.ItemEmb_KeyPress(KeyAscii)
End Sub

Public Sub ItemEmb_Validate(Cancel As Boolean) 'embalagem
    Call gobjEmbalagem.ItemEmb_Validate(Cancel)
End Sub

Public Sub ItemEmb_Change() 'Embalagem
    Call gobjEmbalagem.ItemEmb_Change
End Sub

Public Sub ItemEmb_GotFocus() 'Embalagem
    Call gobjEmbalagem.ItemEmb_GotFocus
End Sub

Public Sub BotaoEmbalagens_Click()
    Call gobjEmbalagem.BotaoEmbalagens_Click
End Sub

Public Sub Embalagem_KeyPress(KeyAscii As Integer) 'embalagem
    Call gobjEmbalagem.Embalagem_KeyPress(KeyAscii)
End Sub

Public Sub Embalagem_Validate(Cancel As Boolean) 'embalagem
    Call gobjEmbalagem.Embalagem_Validate(Cancel)
End Sub

Public Sub Embalagem_Change() 'Embalagem
    Call gobjEmbalagem.Embalagem_Change
End Sub

Public Sub Embalagem_GotFocus() 'Embalagem
    Call gobjEmbalagem.Embalagem_GotFocus
End Sub

Public Sub QuantEmb_KeyPress(KeyAscii As Integer) 'embalagem
    Call gobjEmbalagem.QuantEmb_KeyPress(KeyAscii)
End Sub

Public Sub QuantEmb_Validate(Cancel As Boolean) 'embalagem
    Call gobjEmbalagem.QuantEmb_Validate(Cancel)
End Sub

Public Sub QuantEmb_Change() 'Embalagem
    Call gobjEmbalagem.QuantEmb_Change
End Sub

Public Sub QuantEmb_GotFocus() 'Embalagem
    Call gobjEmbalagem.QuantEmb_GotFocus
End Sub

Public Sub QuantProduto_KeyPress(KeyAscii As Integer) 'embalagem
    Call gobjEmbalagem.QuantProduto_KeyPress(KeyAscii)
End Sub

Public Sub QuantProduto_Validate(Cancel As Boolean) 'embalagem
    Call gobjEmbalagem.QuantProduto_Validate(Cancel)
End Sub

Public Sub QuantProduto_Change() 'Embalagem
    Call gobjEmbalagem.QuantProduto_Change
End Sub

Public Sub QuantProduto_GotFocus() 'Embalagem
    Call gobjEmbalagem.QuantProduto_GotFocus
End Sub

Public Sub CalculoAuto_Click()
Dim lErro As Long
On Error GoTo Erro_CalculoAuto_Click
    If CalculoAuto.Value = vbChecked Then
        Call gobjEmbalagem.Limpa_Tela_Embalagem
        Call Pesos_Recalcula
    End If
    Exit Sub
Erro_CalculoAuto_Click:
    Select Case gErr
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158079)
    End Select
    Exit Sub
End Sub
'EMBALAGEM
'##############################################
