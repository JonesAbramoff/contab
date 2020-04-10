VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalFatGarSRV"
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

Dim m_objUserControl As Object
Dim giIndexTipoNF As Integer
   
'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjNFiscal As ClassNFiscal
Public gcolAPDet As New Collection

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
Public iFiltroAlterado As Integer
Dim giClienteAlterado As Integer
Dim giFilialAlterada As Integer
Public iFrameAtual As Integer
Public iFrameAtualItens As Integer
Public gobjTribTab As ClassTribTab
Public iPercAcresFinAlterado As Integer
Public iValorFreteAlterado As Integer
Public iValorSeguroAlterado As Integer
Public iValorDespesasAlterado As Integer
Public iValorDescontoAlterado As Integer
Public gdDesconto As Double
Public giFilialEntregaAlterada As Integer
Public giFilialEntregaManual As Integer
Public giDataEmissaoAlterada As Integer
Public iSerieAlterada As Integer
Public giDataReferenciaAlterada As Integer
Public lFabricante As Long
Public iFilialFabricante As Integer
Public gcolItensSol As Collection

'Grids
Public objGridItens As AdmGrid
Public objGridItensSRV As AdmGrid
Public objGridComissoes As AdmGrid
Public objGridAlocacoes As AdmGrid
Public objGridParcelas As AdmGrid
Public objGrid1 As AdmGrid
Public objGridAP As AdmGrid

'Colunas do grid de itens
Dim iG_Prod_Col(0 To 1) As Integer
Dim iG_DescProd_Col(0 To 1) As Integer
Dim iG_UM_Col(0 To 1) As Integer
Dim iG_Qtd_Col(0 To 1) As Integer
Dim iG_VlrUnit_Col(0 To 1) As Integer
Dim iG_PercDesc_Col(0 To 1) As Integer
Dim iG_Desc_Col(0 To 1) As Integer
Dim iG_PrecoTot_Col(0 To 1) As Integer
Dim iG_PrecoTotB_Col(0 To 1) As Integer

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

'Colunas do grid de apontamentos
Public iGrid_APSel_Col As Integer
Public iGrid_APDataSol_Col As Integer
Public iGrid_APSolic_Col As Integer
Public iGrid_APProduto_Col As Integer
Public iGrid_APDescProd_Col As Integer
Public iGrid_APGarantia_Col As Integer
Public iGrid_APOS_Col As Integer
Public iGrid_APCodigo_Col As Integer
Public iGrid_APPeca_Col As Integer
Public iGrid_APDescPeca_Col As Integer
Public iGrid_APTipo_Col As Integer
Public iGrid_APDet_Col As Integer
Public iGrid_APQtd_Col As Integer
Public iGrid_APUM_Col As Integer
Public iGrid_APNumIntDoc_Col As Integer

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
Private Const PIS_DEBITO_TOTAL As String = "PIS_Debito_Total"
Private Const COFINS_DEBITO_TOTAL As String = "COFINS_Debito_Total"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const ISS_RETIDO As String = "ISS_Retido"
Private Const COFINS_RETIDO As String = "COFINS_Retido"
Private Const CSLL_RETIDO As String = "CSLL_Retido"
Private Const CONTACONTABILAPLIC As String = "ContaContabilAplic"
Private Const CONTACONTABILCATPROD As String = "ContaContabilCatProd"
'###########Serviço##############
Private Const CONTACONTABILAPLICSRV As String = "ContaContabilAplicSR"
Private Const CONTACONTABILCATSRV As String = "ContaContabilCatSRV"
Private Const PRODUTO1SRV As String = "Servico_CodigoSRV"
Private Const UNIDADE_MEDSRV As String = "Unidade_MedSRV"
Private Const QUANTIDADE1SRV As String = "QuantidadeSRV"
Private Const PRECO_UNITARIOSRV As String = "Preco_UnitarioSRV"
Private Const PRECO_TOTALSRV As String = "Preco_TotalSRV"
Private Const DESCONTO1SRV As String = "Desconto_ServicoSRV"
Private Const DESCRICAO_ITEMSRV As String = "Descricao_ServicoSRV"
'fim da contabilidade

'Eventos das telas de browse
Private WithEvents objEventoPeca As AdmEvento
Attribute objEventoPeca.VB_VarHelpID = -1
Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1
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

Public gobjRastreamento As ClassRastreamento

Private Const BOTAO_PECA = 0
Private Const BOTAO_SERVICO = 1
Private Const BOTAO_IMPRIMIR = 2
Private Const BOTAO_CONSULTATITREC = 3
Private Const BOTAO_GRAVAR = 4
Private Const BOTAO_EXCLUIR = 5
Private Const BOTAO_LIMPAR = 6
Private Const BOTAO_FECHAR = 7
Private Const BOTAO_LIMPARNF = 8
Private Const BOTAO_PROJETOS = 9
Private Const BOTAO_APONTAMENTO = 10
Private Const BOTAO_OS = 11
Private Const BOTAO_SOLICITACAO = 12
Private Const BOTAO_ESTOQUEPROD = 13
Private Const BOTAO_LOCALIZACAO = 14
Private Const BOTAO_SERIE = 15
Private Const BOTAO_LOTES = 16
Private Const BOTAO_GARANTIA = 17
Private Const BOTAO_GERARITENS = 18

'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_Apontamentos = 2
Private Const TAB_Itens = 3
Private Const TAB_Complemento = 4
Private Const TAB_Cobranca = 5
Private Const TAB_Tributacao = 6
Private Const TAB_Contabilizacao = 7

'--- inicio dos properties get dos controles da tela

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

Public Property Get PrecoTotalB(Index As Integer) As Object
     Set PrecoTotalB = objUserControl.Controls("PrecoTotalB")(Index)
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

Public Property Get objGridI(Index As Integer) As Object
    If Index = ITEMNF_TIPO_PECA Then
         Set objGridI = objGridItens
    Else
         Set objGridI = objGridItensSRV
    End If
End Property

Public Property Get iGrid_ValorUnitarioMoedaSRV_Col() As Integer
     iGrid_ValorUnitarioMoedaSRV_Col = 0
End Property

Public Property Get iGrid_ValorUnitarioMoeda_Col() As Integer
     iGrid_ValorUnitarioMoeda_Col = 0
End Property

Public Property Get iGrid_ServicoSRV_Col() As Integer
     iGrid_ServicoSRV_Col = iG_Prod_Col(ITEMNF_TIPO_SERVICO)
End Property

Public Property Get iGrid_DescServicoSRV_Col() As Integer
     iGrid_DescServicoSRV_Col = iG_DescProd_Col(ITEMNF_TIPO_SERVICO)
End Property

Public Property Get iGrid_UnidadeMedSRV_Col() As Integer
     iGrid_UnidadeMedSRV_Col = iG_UM_Col(ITEMNF_TIPO_SERVICO)
End Property

Public Property Get iGrid_QuantidadeSRV_Col() As Integer
     iGrid_QuantidadeSRV_Col = iG_Qtd_Col(ITEMNF_TIPO_SERVICO)
End Property

Public Property Get iGrid_ValorUnitarioSRV_Col() As Integer
     iGrid_ValorUnitarioSRV_Col = iG_VlrUnit_Col(ITEMNF_TIPO_SERVICO)
End Property

Public Property Get iGrid_PercDescSRV_Col() As Integer
     iGrid_PercDescSRV_Col = iG_PercDesc_Col(ITEMNF_TIPO_SERVICO)
End Property

Public Property Get iGrid_DescontoSRV_Col() As Integer
     iGrid_DescontoSRV_Col = iG_Desc_Col(ITEMNF_TIPO_SERVICO)
End Property

Public Property Get iGrid_PrecoTotalSRV_Col() As Integer
     iGrid_PrecoTotalSRV_Col = iG_PrecoTot_Col(ITEMNF_TIPO_SERVICO)
End Property

Public Property Get iGrid_Produto_Col() As Integer
     iGrid_Produto_Col = iG_Prod_Col(ITEMNF_TIPO_PECA)
End Property

Public Property Get iGrid_DescProduto_Col() As Integer
     iGrid_DescProduto_Col = iG_DescProd_Col(ITEMNF_TIPO_PECA)
End Property

Public Property Get iGrid_UnidadeMed_Col() As Integer
     iGrid_UnidadeMed_Col = iG_UM_Col(ITEMNF_TIPO_PECA)
End Property

Public Property Get iGrid_Quantidade_Col() As Integer
     iGrid_Quantidade_Col = iG_Qtd_Col(ITEMNF_TIPO_PECA)
End Property

Public Property Get iGrid_ValorUnitario_Col() As Integer
     iGrid_ValorUnitario_Col = iG_VlrUnit_Col(ITEMNF_TIPO_PECA)
End Property

Public Property Get iGrid_PercDesc_Col() As Integer
     iGrid_PercDesc_Col = iG_PercDesc_Col(ITEMNF_TIPO_PECA)
End Property

Public Property Get iGrid_Desconto_Col() As Integer
     iGrid_Desconto_Col = iG_Desc_Col(ITEMNF_TIPO_PECA)
End Property

Public Property Get iGrid_PrecoTotal_Col() As Integer
     iGrid_PrecoTotal_Col = iG_PrecoTot_Col(ITEMNF_TIPO_PECA)
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

Public Property Get FrameItens() As Object
     Set FrameItens = objUserControl.Controls("FrameItens")
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
    Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Cliente() As Object
    Set Cliente = objUserControl.Controls("Cliente")
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

Public Property Get Emitente() As Object
     Set Emitente = objUserControl.Controls("Emitente")
End Property

Public Property Get Destinatario() As Object
     Set Destinatario = objUserControl.Controls("Destinatario")
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

Public Property Get DescricaoItem(Index As Integer) As Object
     Set DescricaoItem = objUserControl.Controls("DescricaoItem")(Index)
End Property

Public Property Get UnidadeMed(Index As Integer) As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")(Index)
End Property

Public Property Get Desconto(Index As Integer) As Object
     Set Desconto = objUserControl.Controls("Desconto")(Index)
End Property

Public Property Get PercentDesc(Index As Integer) As Object
     Set PercentDesc = objUserControl.Controls("PercentDesc")(Index)
End Property

Public Property Get PrecoUnitario(Index As Integer) As Object
     Set PrecoUnitario = objUserControl.Controls("PrecoUnitario")(Index)
End Property

Public Property Get Quantidade(Index As Integer) As Object
     Set Quantidade = objUserControl.Controls("Quantidade")(Index)
End Property

Public Property Get Produto(Index As Integer) As Object
     Set Produto = objUserControl.Controls("Produto")(Index)
End Property

Public Property Get PrecoTotal(Index As Integer) As Object
     Set PrecoTotal = objUserControl.Controls("PrecoTotal")(Index)
End Property

Public Property Get GridI(Index As Integer) As Object
    If Index = ITEMNF_TIPO_PECA Then
        Set GridI = objUserControl.Controls("GridItens")
    Else
        Set GridI = objUserControl.Controls("GridItensSRV")
    End If
End Property

Public Property Get GridItens() As Object
     Set GridItens = objUserControl.Controls("GridItens")
End Property

Public Property Get GridItensSRV() As Object
     Set GridItensSRV = objUserControl.Controls("GridItensSRV")
End Property

Public Property Get DataEntrega(Index As Integer) As Object
     Set DataEntrega = objUserControl.Controls("DataEntrega")(Index)
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

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

Public Property Get TabStrip2() As Object
     Set TabStrip2 = objUserControl.Controls("TabStrip2")
End Property

Public Property Get Status() As Object
    Set Status = objUserControl.Controls("Status")
End Property

Public Property Get ImprimeRomaneio() As Object
     Set ImprimeRomaneio = objUserControl.Controls("ImprimeRomaneio")
End Property

Public Property Get ImprimeNFGravacao() As Object
     Set ImprimeNFGravacao = objUserControl.Controls("ImprimeNFGravacao")
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

Public Property Get GridAP() As Object
     Set GridAP = objUserControl.Controls("GridAP")
End Property

Public Property Get APSolic() As Object
     Set APSolic = objUserControl.Controls("APSolic")
End Property

Public Property Get APProduto() As Object
     Set APProduto = objUserControl.Controls("APProduto")
End Property

Public Property Get APDescProd() As Object
     Set APDescProd = objUserControl.Controls("APDescProd")
End Property

Public Property Get APGarantia() As Object
     Set APGarantia = objUserControl.Controls("APGarantia")
End Property

Public Property Get APOS() As Object
     Set APOS = objUserControl.Controls("APOS")
End Property

Public Property Get APCodigo() As Object
     Set APCodigo = objUserControl.Controls("APCodigo")
End Property

Public Property Get APPeca() As Object
     Set APPeca = objUserControl.Controls("APPeca")
End Property

Public Property Get APDescPeca() As Object
     Set APDescPeca = objUserControl.Controls("APDescPeca")
End Property

Public Property Get APTipo() As Object
     Set APTipo = objUserControl.Controls("APTipo")
End Property

Public Property Get APDet() As Object
     Set APDet = objUserControl.Controls("APDet")
End Property

Public Property Get APQtd() As Object
     Set APQtd = objUserControl.Controls("APQtd")
End Property

Public Property Get APUM() As Object
     Set APUM = objUserControl.Controls("APUM")
End Property

Public Property Get APDataSol() As Object
     Set APDataSol = objUserControl.Controls("APDataSol")
End Property

Public Property Get APSel() As Object
     Set APSel = objUserControl.Controls("APSel")
End Property

Public Property Get APNumIntDoc() As Object
     Set APNumIntDoc = objUserControl.Controls("APNumIntDoc")
End Property

Public Sub ImprimeRomaneio_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub BotaoLimparNF_Click()

    NFiscal.Caption = ""
    
    'Limpa o Frame de Rastreamento
    'Call gobjRastreamento.Limpa_Tela_Rastreamento
    
    lFabricante = 0
    iFilialFabricante = 0
    Call Trata_Troca_Tab
    
End Sub

Public Function CargaPosFormLoad(Optional bTrazendoDoc As Boolean = False) As Long

Dim lErro As Long
Dim iTipo As Integer

On Error GoTo Erro_CargaPosFormLoad

    If (giPosCargaOk = 0) Then

        'p/permitir o redesenho da tela
        DoEvents
        
        gbCarregandoTela = True
        
        'Rastreamento
        Set gobjRastreamento = New ClassRastreamento
        Set gobjRastreamento.objTela = Me
        gobjRastreamento.bSaidaMaterial = True
        gobjRastreamento.bEditavel = False
        
'        Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
'        Set gobjTelaProjetoInfo.objUserControl = objUserControl
'        Set gobjTelaProjetoInfo.objTela = Me
        
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

        'Inicializa a parte de tributacao
        Call GridItens_InicializaColunas
        
        Set gobjTribTab = New ClassTribTab
        lErro = gobjTribTab.Ativar(Me, , , , , True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Carrega os Tipos de Notas Fiscais
        lErro = Carrega_TiposDocInfo()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If IsDate(DataEmissao.Text) Then Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
        Call gobjTribTab.Serie_Alterada(Serie.Text)
    
        'Carrega a combo de series
        lErro = Carrega_Serie()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Seleciona série Padrão
        lErro = Seleciona_Serie_Padrao()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Carrega a combo de tabelas de preços
        lErro = Carrega_TabelaPreco()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Carrega a combo de transportadoras
        lErro = Carrega_Transportadora()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Carrega a combo VolumeEspecie
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Carrega a combo VolumeMarca
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Carrega a combo PlacaUF
        lErro = Carrega_PlacaUF()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'ACrrega a combo CanalVenda
        lErro = Carrega_CanalVenda()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Carrega na combo as Condições de Pagamento
        lErro = Carrega_CondicaoPagamento()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Carrega os Tipos de Desconto
        lErro = Carrega_TipoDesconto()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Preenche Data Referencia e Data de Emissão coma Data Atual
        DataReferencia.PromptInclude = False
        DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
        DataReferencia.PromptInclude = True
        
        'Inicializa as variáveis de Grid e de Browse
        Set objEventoMensagem = New AdmEvento
        Set objGridComissoes = New AdmGrid
        Set objGridAlocacoes = New AdmGrid
        Set objGrid1 = New AdmGrid
        Set objGridParcelas = New AdmGrid
        Set objEventoSerie = New AdmEvento
        Set objEventoNFiscal = New AdmEvento
        Set objEventoNaturezaOp = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoPeca = New AdmEvento
        Set objEventoServico = New AdmEvento
        Set objEventoEstoqueProd = New AdmEvento
        Set objEventoVendedor = New AdmEvento
        Set objEventoLocalizacao = New AdmEvento
        Set objEventoTransportadora = New AdmEvento
        Set objEventoTranspRedespacho = New AdmEvento
        Set objEventoLote = New AdmEvento
        Set objEventoDoc = New AdmEvento
        Set objGridParcelas = New AdmGrid
        Set objGridAP = New AdmGrid
    
        giDataReferenciaAlterada = 0
        
        Set objGridItens = New AdmGrid
        Set objGridItensSRV = New AdmGrid
        For iTipo = ITEMNF_TIPO_SERVICO To ITEMNF_TIPO_PECA
            Quantidade(iTipo).Format = FORMATO_ESTOQUE
            PrecoUnitario(iTipo).Format = gobjFAT.sFormatoPrecoUnitario
            'Inicializa a máscara do produto
            lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto(iTipo))
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        Next
        
'        lErro = CF("Inicializa_Mascara_Produto_MaskEd", APPeca)
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'        lErro = CF("Inicializa_Mascara_Produto_MaskEd", APProduto)
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Inicializa o grid de itens
        lErro = Inicializa_Grid_Itens(objGridI(ITEMNF_TIPO_PECA))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Inicializa_Grid_ItensSRV(objGridI(ITEMNF_TIPO_SERVICO))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Inicializa o grid de comissões
        lErro = objTabComissoes.Inicializa_Grid_Comissoes(objGridComissoes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Inicializa o grid de Alocações
        lErro = Inicializa_Grid_Alocacoes(objGridAlocacoes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Inicializa o grid de Parcelas
        lErro = Inicializa_Grid_Parcelas(objGridParcelas)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Inicializa_Grid_AP(objGridAP)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objGridParcelas.objGrid.ColWidth(iGrid_TipoPagto_Col) = 0
        objGridParcelas.objGrid.TextMatrix(0, iGrid_TipoPagto_Col) = ""
        BotaoTipoPagto.Visible = False
        
        QuantAlocada.Format = FORMATO_ESTOQUE
        QuantVendida.Format = FORMATO_ESTOQUE
    
        'Inicialização da parte de contabilidade
        lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_SERVICOS, objContabil.iSubTipoTransacaoExt)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        gbCarregandoTela = False
        
        iAlterado = 0
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
          
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205923)
     
    End Select
     
    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set gobjNFiscal = New ClassNFiscal

    iFrameAtual = 1
    iFrameAtualItens = 1
    giIndexTipoNF = -1
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    Set objTabComissoes.objTela = Me
    
    MsgAutomatica.Value = vbChecked
    Mensagem.MaxLength = MENSAGEM_NF_MAXLENGTH
    MensagemCorpo.MaxLength = MENSAGEM_NF_MAXLENGTH
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205924)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Se possui série padrão
    If Len(Trim(sSeriePadrao)) > 0 Then
    
        'Lê a série
        objSerie.sSerie = sSeriePadrao
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO And lErro <> 22202 Then gError ERRO_SEM_MENSAGEM
        
        'Série não cadastrada
        If lErro = 22202 Then gError 205925
        
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
    
        Case ERRO_SEM_MENSAGEM
        
        Case 205925
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205926)
    
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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205928)

    End Select

    Exit Function

End Function

Private Function Carrega_TabelaPreco() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TabelaPreco

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePreco", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        TabelaPreco.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TabelaPreco.ItemData(TabelaPreco.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205929)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205930)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205931)

    End Select

    Exit Function

End Function

Private Function Carrega_CondicaoPagamento() As Long
'Carrega na combo as Condições de Pagamento existentes

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCod_DescReduzida As AdmCodigoNome

On Error GoTo Erro_Carrega_CondicaoPagamento

    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
    lErro = CF("CondicoesPagto_Le_Recebimento", colCod_DescReduzida)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For Each objCod_DescReduzida In colCod_DescReduzida
        'Adiciona novo ítem na List da Combo CondicaoPagamento
        CondicaoPagamento.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
        CondicaoPagamento.ItemData(CondicaoPagamento.NewIndex) = objCod_DescReduzida.iCodigo
    Next

    Carrega_CondicaoPagamento = SUCESSO

    Exit Function

Erro_Carrega_CondicaoPagamento:

    Carrega_CondicaoPagamento = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205932)

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

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205933)

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

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205934)

    End Select

    Exit Function

End Function

Private Sub GridItens_InicializaColunas()

    'Colunas da Grid
    iG_Prod_Col(ITEMNF_TIPO_PECA) = 1
    iG_DescProd_Col(ITEMNF_TIPO_PECA) = 2
    iG_UM_Col(ITEMNF_TIPO_PECA) = 3
    iG_Qtd_Col(ITEMNF_TIPO_PECA) = 4
    iG_VlrUnit_Col(ITEMNF_TIPO_PECA) = 5
    iG_PercDesc_Col(ITEMNF_TIPO_PECA) = 6
    iG_Desc_Col(ITEMNF_TIPO_PECA) = 7
    iG_PrecoTot_Col(ITEMNF_TIPO_PECA) = 8
    iG_PrecoTotB_Col(ITEMNF_TIPO_PECA) = 9
    
    iG_Prod_Col(ITEMNF_TIPO_SERVICO) = 1
    iG_DescProd_Col(ITEMNF_TIPO_SERVICO) = 2
    iG_UM_Col(ITEMNF_TIPO_SERVICO) = 3
    iG_Qtd_Col(ITEMNF_TIPO_SERVICO) = 4
    iG_VlrUnit_Col(ITEMNF_TIPO_SERVICO) = 5
    iG_PercDesc_Col(ITEMNF_TIPO_SERVICO) = 6
    iG_Desc_Col(ITEMNF_TIPO_SERVICO) = 7
    iG_PrecoTot_Col(ITEMNF_TIPO_SERVICO) = 8
    iG_PrecoTotB_Col(ITEMNF_TIPO_SERVICO) = 9

End Sub

Public Function Inicializa_Grid_Itens(ByVal objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

Dim iIndice As Integer

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Peça")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto(ITEMNF_TIPO_PECA).Name)
    objGridInt.colCampo.Add (DescricaoItem(ITEMNF_TIPO_PECA).Name)
    objGridInt.colCampo.Add (UnidadeMed(ITEMNF_TIPO_PECA).Name)
    objGridInt.colCampo.Add (Quantidade(ITEMNF_TIPO_PECA).Name)
    objGridInt.colCampo.Add (PrecoUnitario(ITEMNF_TIPO_PECA).Name)
    objGridInt.colCampo.Add (PercentDesc(ITEMNF_TIPO_PECA).Name)
    objGridInt.colCampo.Add (Desconto(ITEMNF_TIPO_PECA).Name)
    objGridInt.colCampo.Add (PrecoTotal(ITEMNF_TIPO_PECA).Name)
    objGridInt.colCampo.Add (PrecoTotalB(ITEMNF_TIPO_PECA).Name)
    
    For iIndice = 1 To objGridInt.colCampo.Count
        objGridInt.colIndex.Add ITEMNF_TIPO_PECA
    Next

    'Grid do GridInterno
    objGridInt.objGrid = GridI(ITEMNF_TIPO_PECA)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridI(ITEMNF_TIPO_PECA).ColWidth(0) = 500

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Function Inicializa_Grid_ItensSRV(ByVal objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

Dim iIndice As Integer

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Serviço")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto(ITEMNF_TIPO_SERVICO).Name)
    objGridInt.colCampo.Add (DescricaoItem(ITEMNF_TIPO_SERVICO).Name)
    objGridInt.colCampo.Add (UnidadeMed(ITEMNF_TIPO_SERVICO).Name)
    objGridInt.colCampo.Add (Quantidade(ITEMNF_TIPO_SERVICO).Name)
    objGridInt.colCampo.Add (PrecoUnitario(ITEMNF_TIPO_SERVICO).Name)
    objGridInt.colCampo.Add (PercentDesc(ITEMNF_TIPO_SERVICO).Name)
    objGridInt.colCampo.Add (Desconto(ITEMNF_TIPO_SERVICO).Name)
    objGridInt.colCampo.Add (PrecoTotal(ITEMNF_TIPO_SERVICO).Name)
    objGridInt.colCampo.Add (PrecoTotalB(ITEMNF_TIPO_SERVICO).Name)
    
    For iIndice = 1 To objGridInt.colCampo.Count
        objGridInt.colIndex.Add ITEMNF_TIPO_SERVICO
    Next

    'Grid do GridInterno
    objGridInt.objGrid = GridI(ITEMNF_TIPO_SERVICO)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridI(ITEMNF_TIPO_SERVICO).ColWidth(0) = 500

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ItensSRV = SUCESSO

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
    objGridInt.iLinhasVisiveis = 10

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

Public Sub Form_UnLoad(Cancel As Integer)
    
Dim lErro As Long
        
    Set gobjInfoUsu = Nothing

'    Set gobjTelaProjetoInfo = Nothing

    'Libera as variáveis globais da tela
    Set gobjNFiscal = Nothing
    Set gcolAPDet = Nothing
    Set objGridItens = Nothing
    Set objGridItensSRV = Nothing
    Set objGridComissoes = Nothing
    Set objGridAlocacoes = Nothing
    Set objGridParcelas = Nothing
    Set objGrid1 = Nothing
    Set gobjInfoAdic = Nothing

    Set objEventoMensagem = Nothing
    Set objEventoSerie = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoNaturezaOp = Nothing
    Set objEventoProduto = Nothing
    Set objEventoPeca = Nothing
    Set objEventoServico = Nothing
    Set objEventoEstoqueProd = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoLocalizacao = Nothing
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

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

    '?? jones 02/09/12
    Set gcolItensSol = Nothing
    Set objGridAP = Nothing
    
    If Not (gobjTribTab Is Nothing) Then Call gobjTribTab.Desativar
    
    'Fecha o Comando de Setas
    Call ComandoSeta_Liberar(Me.Name)

    Set gobjAnotacao = Nothing
    
    Set gobjRastreamento = Nothing

    Set objTabComissoes = Nothing
    Set gobjTribTab = Nothing

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205936)
    
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

Public Sub Destinatario_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Emitente_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NFiscal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NFiscal_GotFocus()
    Call MaskEdBox_TrataGotFocus(NFiscal, iAlterado)
End Sub

Public Sub TabStrip1_Click()
    
Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'se abriu o tab de tributacao
        If TabStrip1.SelectedItem.Index = TAB_Tributacao Then
            lErro = gobjTribTab.TabClick
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If

        lErro = Trata_Troca_Tab
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If TabStrip1.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao

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
            
'            Case TAB_Comissoes
'                'se for tab de comissoes e a empresa
'                'utilizar regras para o calculo de comissoes
'                lErro = objTabComissoes.Comissoes_Calcula
'                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Parent.HelpContextID = IDH_PEDIDO_VENDA_COMISSOES
            
            Case TAB_Tributacao
                Parent.HelpContextID = IDH_NF_FATURA_PEDIDO_TRIBUTACAO
            
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_NF_FATURA_PEDIDO_CONTABILIZACAO
        
        End Select
    
    End If

    Exit Sub
    
Erro_TabStrip1_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205936)

    End Select

    Exit Sub

End Sub

Public Sub TabStrip2_Click()
    
Dim lErro As Long

On Error GoTo Erro_TabStrip2_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip2.SelectedItem.Index <> iFrameAtualItens Then

        If TabStrip_PodeTrocarTab(iFrameAtualItens, TabStrip2, Me) <> SUCESSO Then Exit Sub

        'se abriu o tab de tributacao
        If TabStrip2.SelectedItem.Index = TAB_Tributacao Then
            lErro = gobjTribTab.TabClick
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If

        'Torna Frame correspondente ao Tab selecionado visivel
        FrameItens(TabStrip2.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        FrameItens(iFrameAtualItens).Visible = False
        'Armazena novo valor de iFrameAtualItens
        iFrameAtualItens = TabStrip2.SelectedItem.Index
            
    End If

    Exit Sub
    
Erro_TabStrip2_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205937)

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
            If lErro <> SUCESSO And lErro <> 12253 Then gError ERRO_SEM_MENSAGEM
            'Se não está na combo
            If lErro <> SUCESSO Then
            
                objSerie.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
                'Busca a série no BD
                lErro = CF("Serie_Le", objSerie)
                If lErro <> SUCESSO And lErro <> 22202 Then gError ERRO_SEM_MENSAGEM
                If lErro <> SUCESSO Then gError 46789 'Se não encontrou
            
            End If

        End If
        
        Call gobjTribTab.Serie_Alterada(Serie.Text)
        
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

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205940)

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
    Call Chama_Tela("NFiscalFatGarSRVLista", colSelecao, objNFiscal, objEventoNFiscal)

End Sub

Private Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objNFiscal = obj1

    'Evita ter que alterar os browses a cada campo inserido na tabela de NF
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM

    'Coloca na Tela a Nota Fiscal escolhida
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205941)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))

    giDataEmissaoAlterada = 0

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205942)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205943)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205944)

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
    If Len(Trim(DataSaida.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataSaida.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_DataSaida_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205945)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_HoraSaida_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205946)

    End Select

    Exit Sub

End Sub

Public Sub UpDownSaida_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownSaida_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataSaida, DIMINUI_DATA)
    If lErro Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownSaida_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205947)

    End Select

    Exit Sub

End Sub

Public Sub UpDownSaida_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownSaida_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataSaida, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownSaida_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205948)

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
    If objGridI(ITEMNF_TIPO_PECA).iLinhasExistentes = 0 Then Exit Sub

    'Faz o tratamento para a Tabela Preco selecionada
    lErro = Trata_TabelaPreco()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_TabelaPreco_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205949)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError ERRO_SEM_MENSAGEM
        
        If lErro <> SUCESSO Then gError 205950 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 205951

    Exit Sub

Erro_TabelaPreco_Validate:

    Cancel = True


    Select Case gErr

        Case ERRO_SEM_MENSAGEM
    
        Case 205950  'Não encontrou Tabela de Preço no BD
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")
            If vbMsgRes = vbYes Then
                'Chama a tela de Tabelas de Preço
                Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
            End If
    
        Case 205951
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205952)

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
    If lErro <> SUCESSO And lErro <> 19205 Then gError ERRO_SEM_MENSAGEM
    
    'Não encontrou a Condição de Pagamento --> erro
    If lErro = 19205 Then gError 205953
    
    dPercAcresFin = StrParaDbl(PercAcrescFin.Text)
    
    If Len(Trim(PercAcrescFin.ClipText)) > 0 And dPercAcresFin <> (objCondicaoPagto.dAcrescimoFinanceiro * 100) Then
        
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
        If vbMsgRes = vbYes Then
            PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
            Call PercAcrescFin_Validate(bSGECancelDummy)
        End If
    Else
        PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
        Call PercAcrescFin_Validate(bSGECancelDummy)
    End If
    
    'Testa se ValorTotal está preenchido
    If Len(Trim(ValorTotal)) > 0 Then
    
        'Testa se DataEmissao está preenchida e ValorTotal é positivo
        If Len(Trim(DataReferencia.ClipText)) > 0 And (CDbl(ValorTotal.Caption) > 0) Then
            
            'Preenche o GridParcelas
            lErro = Cobranca_Automatica()
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_CondicaoPagamento_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 205953
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205954)
           
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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condição Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError ERRO_SEM_MENSAGEM
        
        'Não encontrou a Condição de Pagamento
        If lErro = 19205 Then gError 205955

        'Testa se pode ser usada em Contas a Receber
        If objCondicaoPagto.iEmRecebimento = 0 Then gError 205956

        'Coloca na Tela
        CondicaoPagamento.Text = iCodigo & SEPARADOR & objCondicaoPagto.sDescReduzida

        'Se ValorTotal e DataEmissao estiverem preenchidos, preenche GridParcelas
        If Len(Trim(ValorTotal)) > 0 Then
            If Len(Trim(DataReferencia.ClipText)) > 0 And CDbl(ValorTotal.Caption) > 0 Then

                'Preenche o GridParcelas
                lErro = Cobranca_Automatica()
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            End If
        End If

    End If

    'Não encontrou o valor que era STRING
    If lErro = 6731 Then gError 205957

    Exit Sub

Erro_Condicaopagamento_Validate:

    Cancel = True

    Select Case gErr

       Case ERRO_SEM_MENSAGEM

       Case 205955
            'Pergunta se deseja criar a condição de pagamento
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)
            'Se a resposta for positiva
            If vbMsgRes = vbYes Then
                'Chama a tela de Condições de Pagamento
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            Else
            End If

        Case 205956
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", gErr, objCondicaoPagto.iCodigo)

        Case 205957
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205958)

    End Select

    Exit Sub

End Sub

Public Sub PercAcrescFin_Change()
    iAlterado = REGISTRO_ALTERADO
    iPercAcresFinAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercAcrescFin_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercAcrescFin_Validate

    'Verific se o percentual foi alterado
    If iPercAcresFinAlterado = 0 Then Exit Sub
    'Se estiver preenchido
    If Len(Trim(PercAcrescFin.Text)) > 0 Then
        lErro = Porcentagem_Critica_Negativa(PercAcrescFin)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If

    If Len(Trim(TabelaPreco.Text)) > 0 Then

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    iPercAcresFinAlterado = 0
    
    Exit Sub
    
Erro_PercAcrescFin_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205959)
            
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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dValorFrete = CDbl(ValorFrete.Text)
        
        'Coloca o Valor Formatado na tela
        ValorFrete.Text = Format(dValorFrete, "Fixed")

    End If
    
    iValorFreteAlterado = 0

    Call gobjTribTab.ValorFrete_Validate(Cancel, dValorFrete)
    
    Call ValorTotal_Calcula

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205960)

    End Select

    Exit Sub

End Sub

Public Sub ValorProdutos_Change()

'Dim dValorDesconto As Double
'Dim dValorProdutos As Double
'
'    If gbCarregandoTela Then Exit Sub
'
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
'            Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, dValorDesconto)
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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dValorSeguro = CDbl(ValorSeguro.Text)
        
        'Coloca o valor formatado na tela
        ValorSeguro.Text = Format(dValorSeguro, "Fixed")

    End If

    Call gobjTribTab.ValorSeguro_Validate(Cancel, dValorSeguro)
    
    Call ValorTotal_Calcula

    iValorSeguroAlterado = 0
    
    Exit Sub

Erro_Valorseguro_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205961)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dValorDespesas = CDbl(ValorDespesas.Text)

        'coloca o valor formatado na tela
        ValorDespesas.Text = Format(dValorDespesas, "Fixed")

    End If

    iValorDespesasAlterado = 0

    Call gobjTribTab.ValorDespesas_Validate(Cancel, dValorDespesas)
    
    Call ValorTotal_Calcula

    Exit Sub

Erro_ValorDespesas_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205962)

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
'
'Dim lErro As Long
'Dim dValorDesconto As Double
'Dim dValorProdutos As Double
'
'On Error GoTo Erro_ValorDesconto_Validate
'
'    'Verifica se o valor foi alterado
'    If iValorDescontoAlterado = 0 Then Exit Sub
'
'    dValorDesconto = 0
'
'    Call SubTotal_Calcula(dValorProdutos)
'
'    'Verifica se o Valor está preenchido
'    If Len(Trim(ValorDesconto.Text)) > 0 Then
'
'        'Faz a Crítica do Valor digitado
'        lErro = Valor_NaoNegativo_Critica(ValorDesconto.Text)
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'        dValorDesconto = CDbl(ValorDesconto.Text)
'        'Coloca o Valor formatado na tela
'        ValorDesconto.Text = Format(dValorDesconto, "Standard")
'
'        'Verifica se o Valor do Desconto é maior que a soma dos demais valores
'        If dValorDesconto > dValorProdutos Then gError 205963
'
'        dValorProdutos = dValorProdutos - dValorDesconto
'
'    End If
'
'    ValorProdutos.Caption = Format(dValorProdutos, "Standard")
'
'    Call gobjTribTab.ValorDesconto_Validate(Cancel, dValorDesconto)
'
'    Call ValorTotal_Calcula
'
'    iValorDescontoAlterado = 0
'
'    Exit Sub
'
'Erro_ValorDesconto_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case ERRO_SEM_MENSAGEM
'
'        Case 205963
'            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205964)
'
'    End Select
'
'    Exit Sub

End Sub

Public Function Cobranca_Automatica() As Long
'recalcula o tab de cobranca
 
Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Cobranca_Automatica

    'Se automática estiver selecionada e a condicao de pagamento estiver preenchida
    If CobrancaAutomatica.Value = 1 And Len(Trim(CondicaoPagamento.Text)) <> 0 Then
        'Pega a condicao de pagamento da tela
        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)
        'Lê a condição de pagamento
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError ERRO_SEM_MENSAGEM
        If lErro <> SUCESSO Then gError 46822
        'Preenche o grid parcelas de acordo com a condição de pagamento
        lErro = GridParcelas_Preenche(objCondicaoPagto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
 
    End If
    
    Cobranca_Automatica = SUCESSO
    
    Exit Function
    
Erro_Cobranca_Automatica:
    
    Cobranca_Automatica = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 205965
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 205966)
    
    End Select
    
    Exit Function
        
End Function

Public Sub BotaoEstoqueProd_Click()

Dim lErro As Long
Dim objEstoqueProd As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_BotaoEstoqueProd_Click

    lErro = Botao_Estoque(objEstoqueProd, colSelecao, sSelecao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Chama_Tela("EstoqueProdutoLista", colSelecao, objEstoqueProd, objEventoEstoqueProd, sSelecao)

    Exit Sub
    
Erro_BotaoEstoqueProd_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205967)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoEstoqueProd_evSelecao(obj1 As Object)
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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Se nao encontra o item com o código informado
    If lErro = 6730 Then

        'Verifica de o Cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 205967

        sCliente = Cliente.Text

        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe FilialEntrega com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError ERRO_SEM_MENSAGEM

        If lErro = 17660 Then gError 205968

        'coloca na tela
        FilialEntrega.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 205969

    Exit Sub

Erro_FilialEntrega_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 205967
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 205968
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If
            
        Case 205969
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialEntrega.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205970)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM
    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError ERRO_SEM_MENSAGEM

        If lErro = 19250 Then gError 205971 'Não encontrou Transportadora no BD

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    End If

    'Não existe o item com a String na List da Combo
    If lErro = 6731 Then gError 205972

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 205971
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then
                Call Chama_Tela("Transportadora", objTransportadora)
            End If


        Case 205972
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205973)

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
    If lErro <> SUCESSO And lErro <> 12253 Then gError ERRO_SEM_MENSAGEM

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 205974

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 205974
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205975)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub

Erro_VolumeEspecie_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205976)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub

Erro_VolumeMarca_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205977)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Coloca o valor formatado na Tela
    PesoLiquido.Text = Format(PesoLiquido.Text, PesoLiquido.Format)

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205978)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Coloca o valor formatado na Tela
    PesoBruto.Text = Format(PesoBruto.Text, PesoBruto.Format)

    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205979)

    End Select

    Exit Sub

End Sub

Public Sub CanalVenda_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CanalVenda_Click()
    iAlterado = REGISTRO_ALTERADO
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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCanalVenda.iCodigo = iCodigo

        'Tenta ler CanalVenda com esse código no BD
        lErro = CF("CanalVenda_Le", objCanalVenda)
        If lErro <> SUCESSO And lErro <> 23597 Then gError ERRO_SEM_MENSAGEM

        If lErro = 23597 Then gError 205980 'Não encontrou CanalVenda no BD

        'Encontrou CanalVenda no BD, coloca no Text da Combo
        CanalVenda.Text = CStr(objCanalVenda.iCodigo) & SEPARADOR & objCanalVenda.sNome

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 205981

    Exit Sub

Erro_CanalVenda_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 205980  'Não encontrou CanalVenda no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CANALVENDA", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CanalDeVenda", objCanalVenda)
            End If

        Case 205981
            Call Rotina_Erro(vbOKOnly, "ERRO_CANALVENDA_NAO_ENCONTRADO", gErr, CanalVenda.Text)

        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205982)

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
        'Recalcula as comissoes
        lErro = objTabComissoes.Comissoes_Calcula()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    Exit Sub

Erro_ComissaoAutomatica_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205983)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoVendedores_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoVendedores_Click
    
    lErro = objTabComissoes.BotaoVendedores_Click()
    If lErro <> SUCESSO Then gError 205984
    
    Exit Sub

Erro_BotaoVendedores_Click:

    Select Case gErr
    
        Case 205984
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205985)
    
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
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'senao
         Else
            'verificar se o vendedor esta no grid e tratar o campo 'DiretoIndireto'
            lErro = objTabComissoes.VendedorLinha_Preenche1(objVendedor)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        End If
                
    End If

    Me.Show
    
    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM 'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205986)

    End Select

    Exit Sub

End Sub

Function Carrega_Dados_NFiscal(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long, objTituloRec As New ClassTituloReceber

On Error GoTo Erro_Carrega_Dados_NFiscal
    
    'Tenta ler a nota Fiscal passada por parametro
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Lê os itens da NF
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Lê as mensagens
    lErro = CF("NFiscal_Le_Msg", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Lê a parte de Tributação
    lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
    If lErro <> SUCESSO And lErro <> 22867 Then gError ERRO_SEM_MENSAGEM
    
    'Lê a parte de tributação dos itens da Nota Fiscal
    lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Lê as Comissões da Nota Fiscal
    lErro = CF("ComissoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'    'Lê as Alocações dos itens da Nota Fiscal
'    lErro = CF("AlocacoesNF_Le", objNFiscal)
'    If lErro <> SUCESSO Then gError 46845
    
    'Lê as parcelas da Nota Fiscal Fatura
    lErro = CF("ParcelasRecNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    'carrega dados do titulo a receber associado à nota fiscal
    objTituloRec.lNumIntDoc = objNFiscal.lNumIntDocCPR
    lErro = CF("TituloReceber_Le", objTituloRec, 1)
    If lErro <> SUCESSO And lErro <> 26061 Then gError ERRO_SEM_MENSAGEM
    If lErro <> SUCESSO Then
        'Busca nos baixados
        lErro = CF("TituloReceberBaixado_Le", objTituloRec, 1)
        If lErro <> SUCESSO And lErro <> 56570 Then gError ERRO_SEM_MENSAGEM
        If lErro <> SUCESSO Then gError 205987
    
    End If
    'Transfere para a nota a condição de pagamento do título
    objNFiscal.iCondicaoPagto = objTituloRec.iCondicaoPagto
    
    Carrega_Dados_NFiscal = SUCESSO
    
    Exit Function
    
Erro_Carrega_Dados_NFiscal:

    Carrega_Dados_NFiscal = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 205987
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_REC_INEXISTENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205988)
            
    End Select
    
    Exit Function

End Function

Public Function Traz_NFiscal_Tela(objNFiscal As ClassNFiscal, Optional lNumIntPedido As Long = 0) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim bAlterouCT As Boolean

On Error GoTo Erro_Traz_NFiscal_Tela
        
    gbCarregandoTela = True
    
    'Lê os demais dados da nota fiscal passada se não veio de Pedido
    If lNumIntPedido = 0 Then
        lErro = Carrega_Dados_NFiscal(objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    'Limpa a tela de Nota Fiscal
    lErro = Limpa_Tela_NFiscalFatura()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Traz_NFiscal_Tela2(objNFiscal, lNumIntPedido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Set gobjNFiscal = objNFiscal
    
    If lNumIntPedido <> 0 Then
        MsgAutomatica.Value = vbChecked
    Else
        MsgAutomatica.Value = vbUnchecked
    End If
    
    iAlterado = 0
    iComissoesAlterada = 0
    gbCarregandoTela = False
    
    Traz_NFiscal_Tela = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela:

    gbCarregandoTela = False

    Traz_NFiscal_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205989)

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
Dim objSolicSRV As New ClassSolicSRV
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colSol As New Collection

On Error GoTo Erro_Traz_NFiscal_Tela2

    iTipoNFiscal = objNFiscal.iTipoNFiscal
    
    objCliente.lCodigo = objNFiscal.lCliente
    'Lê o cliente
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 19062 Then gError ERRO_SEM_MENSAGEM
    If lErro <> SUCESSO Then gError 205990
    'Lê as filiais do cliente
    lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    'CArrega a combo de filial entrega
    Call CF("Filial_Preenche", FilialEntrega, colCodigoNome)
        
    'Preenche o Status da Nota Fiscal
    If objNFiscal.iStatus = STATUS_LANCADO Then
        Status.Caption = STRING_STATUS_LANCADO
    ElseIf objNFiscal.iStatus = STATUS_BAIXADO Then
        Status.Caption = STRING_STATUS_BAIXADO
    ElseIf objNFiscal.iStatus = STATUS_CANCELADO Then
        Status.Caption = STRING_STATUS_CANCELADO
    End If
    
    'Preenche o Cliente
    Cliente.Text = objCliente.sNomeReduzido
    Call Cliente_Validate(bSGECancelDummy)
    gdDesconto = objCliente.dDesconto
    
    'Preenche a Filial do Cliente
    For iIndice = 0 To FilialEntrega.ListCount - 1
        If FilialEntrega.ItemData(iIndice) = objNFiscal.iFilialCli Then
            Filial.Text = FilialEntrega.List(iIndice)
            Call Filial_Validate(bSGECancelDummy)
            Exit For
        End If
    Next
    
    'Não vai trazer os apontamentos em aberto
    lFabricante = objNFiscal.lCliente
    iFilialFabricante = objNFiscal.iFilialCli
    
    'LER E TRAZER OS APONTAMENTO DA NF
        
    lErro = CF("NFiscal_Le_Apontamentos", objNFiscal, colSol)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Traz_Apontamento_Tela(colSol)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Coloca a FilialEntrega na tela
    If objNFiscal.iFilialEntrega > 0 Then
        FilialEntrega.Text = objNFiscal.iFilialEntrega
        Call FilialEntrega_Validate(bSGECancelDummy)
    End If

    'Coloca o Tipo na tela
    If iTipoNFiscal = 0 Then iTipoNFiscal = DOCINFO_NFFPSRV
    TipoNFiscal.Text = iTipoNFiscal
    Call TipoNFiscal_Validate(bSGECancelDummy)
    
    'Coloca TabelaPreco na Tela
    If objNFiscal.iTabelaPreco > 0 Then
        For iIndice = 0 To TabelaPreco.ListCount - 1
            If objNFiscal.iTabelaPreco = TabelaPreco.ItemData(iIndice) Then
                TabelaPreco.Text = TabelaPreco.List(iIndice)
                Exit For
            End If
        Next
    End If

    'Preenche a Condicao de Pagto
    If objNFiscal.iCondicaoPagto > 0 Then
        For iIndice = 0 To CondicaoPagamento.ListCount - 1
            If objNFiscal.iCondicaoPagto = CondicaoPagamento.ItemData(iIndice) Then
                CondicaoPagamento.Text = CondicaoPagamento.List(iIndice)
                Exit For
            End If
        Next
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
    'Responsável pelo frete
    If objNFiscal.iFreteRespons = FRETE_EMITENTE Then
        Emitente.Value = True
    Else
        Destinatario.Value = True
    End If
    
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
    
    VolumeNumero = objNFiscal.sVolumeNumero

    If objNFiscal.lVolumeQuant > 0 Then VolumeQuant.Text = objNFiscal.lVolumeQuant
    NumPedidoTerc.Text = objNFiscal.sNumPedidoTerc

    Mensagem.Text = objNFiscal.sMensagemNota
    MensagemCorpo.Text = objNFiscal.sMensagemCorpoNota

    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = Format(objNFiscal.dPesoLiq, PesoLiquido.Format) 'Alterado por Wagner
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = Format(objNFiscal.dPesoBruto, PesoBruto.Format) 'Alterado por Wagner
    If objNFiscal.dVolumeTotal > 0 Then Cubagem.Text = Format(objNFiscal.dVolumeTotal, "Fixed")
    
    'Preenche o Grid com os itens da Nota Fiscal
    lErro = Preenche_GridI(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Preenche o Grid com os itens da Nota Fiscal
    lErro = Preenche_GridISRV(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If objNFiscal.iTipoNFiscal <> 0 Then

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        
        'Le o Tipo para saber se tem Movimento de Estoque
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
            'Carrega ItensNF com Rastreamentos
            lErro = gobjRastreamento.Carrega_RastroItensNF(objNFiscal)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If

    End If

    'Carrega o Tab Comissões
    lErro = objTabComissoes.Carrega_Tab_Comissoes(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If lNumIntPedido > 0 Then bGerarDados = True
    
    If gbCarregandoTela Then
        lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If

'    'Preenche o Grid com as Alocações dos itens da Nota Fiscal
'    lErro = Preenche_GridAlocacoes(Me, objNFiscal, bGerarDados)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Se a nota fiscal já existe
    If objNFiscal.lNumIntDoc <> 0 Then
    
        'Traz os dados contábeis para a tela (contabilidade)
        lErro = objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
        If lErro <> SUCESSO And lErro <> 36326 Then gError ERRO_SEM_MENSAGEM
        
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
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM

    gobjTribTab.sNatOpInterna = objNFiscal.sNaturezaOp

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
    
    Call SubTotal_Calcula
    
    Set gobjInfoAdic = objNFiscal.objInfoAdic
    
    ValorDescontoTit.Text = Format(objNFiscal.dValorDescontoTit, "Standard")
    Call ValorDescontoTit_Validate(bSGECancelDummy)
    
    Traz_NFiscal_Tela2 = SUCESSO
    
    Exit Function

Erro_Traz_NFiscal_Tela2:

    Traz_NFiscal_Tela2 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 205990
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205991)

    End Select

    Exit Function

End Function

Private Function Trata_TabelaPreco() As Long

Dim lErro As Long, iCodTabela As Integer
Dim iLinha As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim sProduto As String
Dim iPreenchido As Integer
Dim iTipo As Integer

On Error GoTo Erro_TabelaPreco_Click

    iCodTabela = Codigo_Extrai(TabelaPreco.Text)
    
    For iTipo = ITEMNF_TIPO_SERVICO To ITEMNF_TIPO_PECA
        
        For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
    
            'Verifica se o Produto está preenchido
            lErro = CF("Produto_Formata", GridI(iTipo).TextMatrix(iLinha, iG_Prod_Col(iTipo)), sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            If iPreenchido <> PRODUTO_VAZIO And Len(Trim(GridI(iTipo).TextMatrix(iLinha, iG_UM_Col(iTipo)))) > 0 Then
                objTabelaPrecoItem.sCodProduto = sProduto
                objTabelaPrecoItem.iCodTabela = iCodTabela
                objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
                'Lê a Tabela preço para filialEmpresa
                lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM
                'Se não encontrar
                If lErro = 28014 Then
                    objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
                    'Lê a Tabela de Preço a nível de Empresa toda
                    lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                    If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM
    
                End If
                'Se  conseguir ler a Tabela de Preços
                If lErro = SUCESSO Then
                    'Calcula o Preco Unitário do item
                    lErro = PrecoUnitario_Calcula(GridI(iTipo).TextMatrix(iLinha, iG_UM_Col(iTipo)), objTabelaPrecoItem, dPrecoUnitario)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    'Coloca no Grid
                    If dPrecoUnitario > 0 Then
                        GridI(iTipo).TextMatrix(iLinha, iG_VlrUnit_Col(iTipo)) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
                    Else
                        GridI(iTipo).TextMatrix(iLinha, iG_VlrUnit_Col(iTipo)) = ""
                    End If
                    'Calcula o Preco Total do item
                    Call PrecoTotal_Calcula(iTipo, iLinha, False)
    
                    ''*** incluidos p/tratamento de tributacao *******************************
                    If iTipo = ITEMNF_TIPO_PECA Then
                        lErro = gobjTribTab.Alteracao_Item_Grid(iLinha)
                    Else
                        lErro = gobjTribTab.Alteracao_Item_Grid(iLinha, True)
                    End If
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    '*** fim tributacao
                
                End If
    
            End If
            
        Next
    Next

    'Calcula o Valor Total da Nota
    Call ValorTotal_Calcula

    Trata_TabelaPreco = SUCESSO

    Exit Function

Erro_TabelaPreco_Click:

    Trata_TabelaPreco = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205992)

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

On Error GoTo Erro_GridParcelas_Preenche

    Set gobjNFiscal.ColParcelaReceber = New ColParcelaReceber

    'Limpa o GridParcelas
    Call Grid_Limpa(objGridParcelas)
    
    'Número de Parcelas
    objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas
    
    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Se Valor a Pagar for positivo
    If dValorPagar > 0 Then
    
        'Calcula os valores das Parcelas
        objCondicaoPagto.dValorTotal = dValorPagar
        
        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, False)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        Next
    
    End If

    'Se Data Emissão estiver preenchida
    If Len(Trim(DataReferencia.ClipText)) > 0 Then
    
        objCondicaoPagto.dtDataRef = StrParaDate(DataReferencia.Text)
        objCondicaoPagto.dtDataEntrega = StrParaDate(DataReferencia.Text)
        
        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas
        
            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dtVencimento, "dd/mm/yyyy")
            
            Set objParcRec = New ClassParcelaReceber
            gobjNFiscal.ColParcelaReceber.AddObj objParcRec
            
        Next
                        
    End If
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        lErro = Preenche_DescontoPadrao(iIndice)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    Next
    GridParcelas_Preenche = SUCESSO
    
    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205993)
        
    End Select
    
    Exit Function

End Function

Function ValorTotal_Calcula() As Long

Dim dValorTotal As Double
Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    lErro = gobjTribTab.ValorTotal_Calcula(dValorTotal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ValorTotal.Caption = Format(dValorTotal, "Standard")

    Call ValorDescontoTit_Validate(bSGECancelDummy)

    Call ValorDescontoItens_Calcula

    If Not gbCarregandoTela Then
        'Gera a cobranca em cima do novo valor total
        lErro = Cobranca_Automatica()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Se a empresa nãu utiliza as regras para o cálculo de comissões
        If Not gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
        
            'Significa que é necessário recalcular as comissões, utilizando o cálculo padrão
            lErro = objTabComissoes.Comissoes_Calcula_Padrao()
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Else
            iComissoesAlterada = 1 'preventivamente, pois pode ter alterado algo que afete uma regra
        End If
    
    End If
    
    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205994)

    End Select

    Exit Function

End Function

Public Function Pesos_Recalcula() As Long

Dim lErro As Long
Dim dFator As Double
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dPesoBruto As Double
Dim dPesoLiquido As Double
Dim dQuantidade As Double

On Error GoTo Erro_Pesos_Recalcula
    
    lErro = Calcula_VolumeTotal
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    dPesoBruto = 0
    dPesoLiquido = 0
    
    'Para cada linha do Grid de Itens
    For iIndice = 1 To objGridI(ITEMNF_TIPO_PECA).iLinhasExistentes
    
        'Se o produto e a UM estiverem preenchidos
        If Len(Trim(GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_Prod_Col(ITEMNF_TIPO_PECA)))) > 0 And Len(Trim(GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_UM_Col(ITEMNF_TIPO_PECA)))) > 0 Then
            
            'Formata o produto
            lErro = CF("Produto_Formata", GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_Prod_Col(ITEMNF_TIPO_PECA)), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            'Lê o produto
            objProduto.sCodigo = sProdutoFormatado
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

            'Se o produto não está cadastrado, erro
            If lErro = 28030 Then gError 205995
                        
            'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_UM_Col(ITEMNF_TIPO_PECA)), objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                                
            dQuantidade = StrParaDbl(GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_Qtd_Col(ITEMNF_TIPO_PECA)))
            
            'Acumula o peso Bruto e o peso líquido
            dPesoBruto = dPesoBruto + objProduto.dPesoBruto * dFator * dQuantidade
            dPesoLiquido = dPesoLiquido + objProduto.dPesoLiq * dFator * dQuantidade
                        
        End If
        
    Next
    
    'Coloca os Pesos na tela
    PesoBruto.Text = Format(dPesoBruto, PesoBruto.Format)
    PesoLiquido.Text = Format(dPesoLiquido, PesoLiquido.Format)
    
    Pesos_Recalcula = SUCESSO
    
    Exit Function
    
Erro_Pesos_Recalcula:

    Pesos_Recalcula = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 205995
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205996)
    
    End Select
    
    Exit Function
    
End Function

Private Function Limpa_Tela_NFiscalFatura() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_NFiscalFatura

    'Indica que está limpando a tela
    gbLimpandoTela = True
    
    'Limpa os campos gerais da tela
    Call Limpa_Tela(Me)
     
'    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
'    Set gobjTelaProjetoInfo.objUserControl = objUserControl
'    Set gobjTelaProjetoInfo.objTela = Me
    
    Set gobjAnotacao = Nothing
    Set gobjInfoAdic = Nothing
    
'    If giTipoVersao = VERSAO_FULL Then
        Serie.Text = ""
        Call Serie_Validate(bSGECancelDummy)
        TabelaPreco.Text = ""
'    End If
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
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
    Filial.Clear
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
    Emitente.Value = True
    EletronicaFed.Value = vbUnchecked
    
    DataReferencia.PromptInclude = False
    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataReferencia.PromptInclude = True
        
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    Call DataEmissao_Validate(bSGECancelDummy)
       
    TipoNFiscal.ListIndex = giIndexTipoNF
    'Frame1(TAB_Almoxarifado).Enabled = True
    
    iValorDescontoAlterado = 0
    iPercAcresFinAlterado = 0
    iValorFreteAlterado = 0
    iValorSeguroAlterado = 0
    iValorDespesasAlterado = 0
    iValorDescontoAlterado = 0
    gdDesconto = 0
    giDataReferenciaAlterada = 0
    
    lFabricante = 0
    iFilialFabricante = 0
    
    For iIndice = 1 To objGridI(ITEMNF_TIPO_PECA).iLinhasExistentes
        GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, 0) = iIndice
    Next
    
    NaoImpCobranca.Value = vbUnchecked
    
    'Limpa os grids
    Call Grid_Limpa(objGridI(ITEMNF_TIPO_PECA))
    Call Grid_Limpa(objGridI(ITEMNF_TIPO_SERVICO))
    Call Grid_Limpa(objGridComissoes)
    Call Grid_Limpa(objGridAlocacoes)
    Call Grid_Limpa(objGridParcelas)
    Call Grid_Limpa(objGridAP)
    
    Set gcolAPDet = New Collection
    
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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento

    MsgAutomatica.Value = vbChecked

    iAlterado = 0
    iComissoesAlterada = 0
    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0

    'Indica que não está mais limpando a tela
    gbLimpandoTela = False
    
    Exit Function

Erro_Limpa_Tela_NFiscalFatura:
    
    Limpa_Tela_NFiscalFatura = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205997)
    
    End Select
    
    'Indica que não está mais limpando a tela
    gbLimpandoTela = False
    
    Exit Function
    
End Function

Public Function Preenche_GridI(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com os itens da Nota Fiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim sProdutoEnxuto As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim dValorProdutos As Double
'Dim dValorDesconto  As Double
Dim dValorTotal As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_Preenche_GridI

    iIndice = 0

    Call Grid_Limpa(objGridI(ITEMNF_TIPO_PECA))

    'Para cada item da Coleção
    For Each objItemNF In objNFiscal.ColItensNF
    
        iIndice = iIndice + 1
        objItemNF.iItem = iIndice

        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Produto(ITEMNF_TIPO_PECA).PromptInclude = False
        Produto(ITEMNF_TIPO_PECA).Text = sProdutoEnxuto
        Produto(ITEMNF_TIPO_PECA).PromptInclude = True

        objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

        If objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade.Count > 0 And objProduto.iKitVendaComp <> MARCADO Then GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, 0) = "# " & GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, 0)
        GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_Prod_Col(ITEMNF_TIPO_PECA)) = Produto(ITEMNF_TIPO_PECA).Text
        GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_DescProd_Col(ITEMNF_TIPO_PECA)) = objItemNF.sDescricaoItem
        GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_UM_Col(ITEMNF_TIPO_PECA)) = objItemNF.sUnidadeMed
        GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_Qtd_Col(ITEMNF_TIPO_PECA)) = Formata_Estoque(objItemNF.dQuantidade)
        GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_VlrUnit_Col(ITEMNF_TIPO_PECA)) = Format(objItemNF.dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_PercDesc_Col(ITEMNF_TIPO_PECA)) = Format(objItemNF.dPercDesc, "Percent")
        
        'dValorDesconto = objItemNF.dPercDesc * objItemNF.dQuantidade * objItemNF.dPrecoUnitario
    
        dValorTotal = (objItemNF.dPrecoUnitario * objItemNF.dQuantidade) - objItemNF.dValorDesconto
    
        GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_Desc_Col(ITEMNF_TIPO_PECA)) = Format(objItemNF.dValorDesconto, "Standard")
        GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_PrecoTot_Col(ITEMNF_TIPO_PECA)) = Format(dValorTotal, "Standard")
        GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_PrecoTotB_Col(ITEMNF_TIPO_PECA)) = Format(dValorTotal + objItemNF.dValorDesconto, "Standard")
        
    Next
    
    'Atualiza o número de linhas existentes
    objGridI(ITEMNF_TIPO_PECA).iLinhasExistentes = iIndice

    Call SubTotal_Calcula
    
    'ValorProdutos.Caption = Format(dValorProdutos - objNFiscal.dValorDesconto, "Standard")
    
'    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objNFiscal.objProjetoInfo, objNFiscal.lNumIntDoc, PRJ_CR_TIPO_NF)
'    If lErro <> SUCESSO Then gError 181676
    
    Preenche_GridI = SUCESSO

    Exit Function

Erro_Preenche_GridI:

    Preenche_GridI = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205998)

    End Select

    Exit Function

End Function

Public Function Preenche_GridISRV(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com os itens da Nota Fiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim sProdutoEnxuto As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim dValorProdutos As Double
'Dim dValorDesconto  As Double
Dim dValorTotal As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_Preenche_GridISRV

    iIndice = 0
    
    Call Grid_Limpa(objGridI(ITEMNF_TIPO_SERVICO))

    'Para cada item da Coleção
    For Each objItemNF In objNFiscal.colItensSRV

        iIndice = iIndice + 1
        objItemNF.iItem = iIndice

        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Produto(ITEMNF_TIPO_SERVICO).PromptInclude = False
        Produto(ITEMNF_TIPO_SERVICO).Text = sProdutoEnxuto
        Produto(ITEMNF_TIPO_SERVICO).PromptInclude = True

        objProduto.sCodigo = objItemNF.sProduto
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

        If objItemNF.colItensRomaneioGrade.Count > 0 And objProduto.iKitVendaComp <> MARCADO Then GridI(ITEMNF_TIPO_SERVICO).TextMatrix(iIndice, 0) = "# " & GridI(ITEMNF_TIPO_SERVICO).TextMatrix(iIndice, 0)
        GridI(ITEMNF_TIPO_SERVICO).TextMatrix(iIndice, iG_Prod_Col(ITEMNF_TIPO_SERVICO)) = Produto(ITEMNF_TIPO_SERVICO).Text
        GridI(ITEMNF_TIPO_SERVICO).TextMatrix(iIndice, iG_DescProd_Col(ITEMNF_TIPO_SERVICO)) = objItemNF.sDescricaoItem
        GridI(ITEMNF_TIPO_SERVICO).TextMatrix(iIndice, iG_UM_Col(ITEMNF_TIPO_SERVICO)) = objItemNF.sUnidadeMed
        GridI(ITEMNF_TIPO_SERVICO).TextMatrix(iIndice, iG_Qtd_Col(ITEMNF_TIPO_SERVICO)) = Formata_Estoque(objItemNF.dQuantidade)
        GridI(ITEMNF_TIPO_SERVICO).TextMatrix(iIndice, iG_VlrUnit_Col(ITEMNF_TIPO_SERVICO)) = Format(objItemNF.dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        GridI(ITEMNF_TIPO_SERVICO).TextMatrix(iIndice, iG_PercDesc_Col(ITEMNF_TIPO_SERVICO)) = Format(objItemNF.dPercDesc, "Percent")
        
        'dValorDesconto = objItemNF.dPercDesc * objItemNF.dQuantidade * objItemNF.dPrecoUnitario
    
        dValorTotal = (objItemNF.dPrecoUnitario * objItemNF.dQuantidade) - objItemNF.dValorDesconto
    
        GridI(ITEMNF_TIPO_SERVICO).TextMatrix(iIndice, iG_Desc_Col(ITEMNF_TIPO_SERVICO)) = Format(objItemNF.dValorDesconto, "Standard")
        GridI(ITEMNF_TIPO_SERVICO).TextMatrix(iIndice, iG_PrecoTot_Col(ITEMNF_TIPO_SERVICO)) = Format(dValorTotal, "Standard")
        GridI(ITEMNF_TIPO_SERVICO).TextMatrix(iIndice, iG_PrecoTotB_Col(ITEMNF_TIPO_SERVICO)) = Format(objItemNF.dValorDesconto, "Standard")
        
    Next
    
    'Atualiza o número de linhas existentes
    objGridI(ITEMNF_TIPO_SERVICO).iLinhasExistentes = iIndice

    Call SubTotal_Calcula
    
    'ValorProdutos.Caption = Format(dValorProdutos - objNFiscal.dValorDesconto, "Standard")
    
    Preenche_GridISRV = SUCESSO

    Exit Function

Erro_Preenche_GridISRV:

    Preenche_GridISRV = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 205999)

    End Select

    Exit Function

End Function

Function Preenche_GridAlocacoes(ByVal objCT As Object, objNFiscal As ClassNFiscal, Optional bGeraAlocacao As Boolean = False) As Long
'
End Function

Public Function PrecoUnitario_Calcula(sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objUM As New ClassUnidadeDeMedida
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim dPercAcresFin As Double
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_PrecoUnitario_Calcula

    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto
    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    If lErro = 28030 Then gError 206000 'Não encontrou

    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Converte o preço unitário para UM da tela
    dPrecoUnitario = objTabelaPrecoItem.dPreco / dFator
    
    'Aplica o acréscimo financeiro no preço unitário
    dPercAcresFin = StrParaDbl(PercAcrescFin.Text)
    dPrecoUnitario = dPrecoUnitario * (1 + (dPercAcresFin / 100))
    
    PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_PrecoUnitario_Calcula:

    PrecoUnitario_Calcula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206000
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objTabelaPrecoItem.sCodProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206001)

    End Select

    Exit Function

End Function

Public Sub PrecoTotal_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrecoTotal_Calcula(ByVal Index As Integer, ByVal iLinha As Integer, Optional iCalculaTotal = True)

Dim dPrecoTotal As Double
Dim dPrecoTotalReal As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim dValorTotal As Double
Dim iIndice As Integer
Dim dValorProdutos As Double
Dim dValorDescontoGlobal As Double, dValorProdutosB As Double, dValorDescontoItens As Double
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_PrecoTotal_Calcula

    dPrecoUnitario = StrParaDbl(GridI(Index).TextMatrix(iLinha, iG_VlrUnit_Col(Index)))
    dQuantidade = StrParaDbl(GridI(Index).TextMatrix(iLinha, iG_Qtd_Col(Index)))

    'Recolhe os valores Quantidade, Desconto, PerDesc e Valor Unitário da tela
    If dPrecoUnitario = 0 Or dQuantidade = 0 Then
        GridI(Index).TextMatrix(iLinha, iG_PrecoTot_Col(Index)) = ""
        GridI(Index).TextMatrix(iLinha, iG_Desc_Col(Index)) = ""
        
    Else
        dPrecoTotal = Arredonda_Moeda(dPrecoUnitario * dQuantidade) 'Alterado por Wagner
        dDesconto = StrParaDbl(GridI((Index)).TextMatrix(iLinha, iG_Desc_Col((Index))))
        dPercentDesc = PercentParaDbl(GridI((Index)).TextMatrix(iLinha, iG_PercDesc_Col((Index))))

        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dPrecoUnitario, dPercentDesc, dDesconto, dPrecoTotalReal)

        'Coloca o Desconto calculado na tela
        If dDesconto > 0 Then
            GridI((Index)).TextMatrix(iLinha, iG_Desc_Col((Index))) = Format(dDesconto, "Standard")
        Else
            GridI((Index)).TextMatrix(iLinha, iG_Desc_Col((Index))) = ""
        End If

        'Coloca o valor Real em Valor Total
        GridI((Index)).TextMatrix(iLinha, iG_PrecoTot_Col((Index))) = Format(dPrecoTotalReal, "Standard")

        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dPrecoUnitario, 0, 0, dPrecoTotalReal)

        'Coloca o valor Real em Valor Total
        GridI((Index)).TextMatrix(iLinha, iG_PrecoTotB_Col(Index)) = Format(dPrecoTotalReal, "Standard")

    End If
    
    Call SubTotal_Calcula
'
'    'Obtem o valor do Desconto global
'    If Len(Trim(ValorDesconto.Text)) > 0 And IsNumeric(ValorDesconto.Text) Then dValorDescontoGlobal = CDbl(ValorDesconto.Text)
'
'    dValorProdutos = Arredonda_Moeda(dValorProdutos - dValorDescontoGlobal)
'
'    'Se com o desconto o valor Produtos ficou negativo
'    If dValorProdutos < 0 And dValorDescontoGlobal > 0 Then
'        'Avisa ao usuário
'        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", dValorDescontoGlobal, dValorProdutos)
'        'Limpa o valor de desconto
'        ValorDesconto.Text = ""
'        dValorDescontoGlobal = 0
'
'        Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, dValorDescontoGlobal)
'
'        'Faz a atualização dos valores da tributação
'        lErro = gobjTribTab.AtualizarTributacao()
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'        Call SubTotal_Calcula(dValorProdutos)
'
'    End If

    
    If iCalculaTotal <> False Then Call ValorTotal_Calcula

    Exit Sub

Erro_PrecoTotal_Calcula:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206002)
            
    End Select
            
    Exit Sub

End Sub

Private Sub SubTotal_Calcula()
'Faz a soma da Coluna passado no Grid passado e devolve em dValorTotal

Dim iIndice As Integer
Dim dSubTotal As Double, dSubTotalBruto As Double, dDesconto As Double

    dSubTotal = 0

    For iIndice = 1 To objGridI(ITEMNF_TIPO_PECA).iLinhasExistentes
        dSubTotal = Arredonda_Moeda(dSubTotal + StrParaDbl(objGridI(ITEMNF_TIPO_PECA).objGrid.TextMatrix(iIndice, iG_PrecoTot_Col(ITEMNF_TIPO_PECA))))
        dSubTotalBruto = dSubTotalBruto + StrParaDbl(objGridI(ITEMNF_TIPO_PECA).objGrid.TextMatrix(iIndice, iG_PrecoTotB_Col(ITEMNF_TIPO_PECA)))
        dDesconto = dDesconto + StrParaDbl(objGridI(ITEMNF_TIPO_PECA).objGrid.TextMatrix(iIndice, iG_Desc_Col(ITEMNF_TIPO_PECA)))
    Next
    
    For iIndice = 1 To objGridI(ITEMNF_TIPO_SERVICO).iLinhasExistentes
        dSubTotal = Arredonda_Moeda(dSubTotal + StrParaDbl(objGridI(ITEMNF_TIPO_SERVICO).objGrid.TextMatrix(iIndice, iG_PrecoTot_Col(ITEMNF_TIPO_SERVICO))))
        dSubTotalBruto = dSubTotalBruto + StrParaDbl(objGridI(ITEMNF_TIPO_SERVICO).objGrid.TextMatrix(iIndice, iG_PrecoTotB_Col(ITEMNF_TIPO_SERVICO)))
        dDesconto = dDesconto + StrParaDbl(objGridI(ITEMNF_TIPO_SERVICO).objGrid.TextMatrix(iIndice, iG_Desc_Col(ITEMNF_TIPO_SERVICO)))
    Next
    
    'Coloca o valor dos produtos na tela
    ValorProdutos.Caption = Format(dSubTotal, "Standard")
    
    ValorProdutos2.Caption = Format(dSubTotalBruto, "Standard")
    ValorDescontoItens.Text = Format(dDesconto, "Standard")
    
    dValorDescontoItensAnt = dDesconto
    If dSubTotalBruto > 0 Then
        PercDescontoItens.Text = Format(dDesconto / dSubTotalBruto * 100, "FIXED")
        dPercDescontoItensAnt = Arredonda_Moeda(dDesconto / dSubTotalBruto, 4)
    Else
        PercDescontoItens.Text = Format(0, "FIXED")
        dPercDescontoItensAnt = 0
    End If

    Exit Sub

End Sub

Private Sub ValorReal_Calcula(dQuantidade As Double, dValorUnitario As Double, dPercentDesc As Double, dDesconto As Double, dValorReal As Double)
'Calcula o Valor Real

Dim dValorTotal As Double
Dim dPercDesc1 As Double
Dim dPercDesc2 As Double

    dValorTotal = Arredonda_Moeda(dValorUnitario * dQuantidade)

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

    dValorReal = Arredonda_Moeda(dValorTotal - dDesconto)

End Sub


Public Sub BotaoLimpar_Click()

Dim objNFiscal As New ClassNFiscal
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa a Tela de Nota fiscal fatura pedido
    lErro = Limpa_Tela_NFiscalFatura()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM 'tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206009)

    End Select

    Exit Sub
    
End Sub

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
'
End Sub

Public Sub Quantidade_Change(Index As Integer)

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_GotFocus(Index As Integer)
    Call Grid_Campo_Recebe_Foco(objGridI(Index))
End Sub

Public Sub Quantidade_KeyPress(Index As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridI(Index))
End Sub

Public Sub Quantidade_Validate(Index As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridI(Index).objControle = Quantidade(Index)
    lErro = Grid_Campo_Libera_Foco(objGridI(Index))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PrecoUnitario_Change(Index As Integer)

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrecoUnitario_GotFocus(Index As Integer)
    Call Grid_Campo_Recebe_Foco(objGridI(Index))
End Sub

Public Sub PrecoUnitario_KeyPress(Index As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridI(Index))
End Sub

Public Sub PrecoUnitario_Validate(Index As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridI(Index).objControle = PrecoUnitario(Index)
    lErro = Grid_Campo_Libera_Foco(objGridI(Index))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PercentDesc_Change(Index As Integer)
    
    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub PercentDesc_GotFocus(Index As Integer)
    Call Grid_Campo_Recebe_Foco(objGridI(Index))
End Sub

Public Sub PercentDesc_KeyPress(Index As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridI(Index))
End Sub

Public Sub PercentDesc_Validate(Index As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridI(Index).objControle = PercentDesc(Index)
    lErro = Grid_Campo_Libera_Foco(objGridI(Index))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Desconto_Change(Index As Integer)

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Desconto_GotFocus(Index As Integer)
    Call Grid_Campo_Recebe_Foco(objGridI(Index))
End Sub

Public Sub Desconto_KeyPress(Index As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridI(Index))
End Sub

Public Sub Desconto_Validate(Index As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridI(Index).objControle = Desconto(Index)
    lErro = Grid_Campo_Libera_Foco(objGridI(Index))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GridItens_Click(Index As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridI(Index), iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridI(Index), iAlterado)
    End If

End Sub

Public Sub GridItens_EnterCell(Index As Integer)
    Call Grid_Entrada_Celula(objGridI(Index), iAlterado)
End Sub

Public Sub GridItens_GotFocus(Index As Integer)
    Call Grid_Recebe_Foco(objGridI(Index))
End Sub

Public Sub GridItens_KeyPress(Index As Integer, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridI(Index), iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridI(Index), iAlterado)
    End If

End Sub

Public Sub GridItens_LeaveCell(Index As Integer)
    Call Saida_Celula(objGridI(Index))
End Sub

Public Sub GridItens_Validate(Index As Integer, Cancel As Boolean)
    Call Grid_Libera_Foco(objGridI(Index))
End Sub

Public Sub GridItens_RowColChange(Index As Integer)
    Call Grid_RowColChange(objGridI(Index))
End Sub

Public Sub GridItens_Scroll(Index As Integer)
    Call Grid_Scroll(objGridI(Index))
End Sub

Public Sub GridItens_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer
Dim iIndice As Integer
Dim dValorProdutos As Double
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_GridItens_KeyDown

    'Armazena o número de linhas existentes e o item atual selecionado
    iLinhasExistentesAnterior = objGridI(Index).iLinhasExistentes
    iItemAtual = GridI(Index).Row
        
    Call Grid_Trata_Tecla1(KeyCode, objGridI(Index))

    'Se algum ítem foi excluído
    If objGridI(Index).iLinhasExistentes < iLinhasExistentesAnterior Then
        
        '************ grade ************
        GridI(Index).TextMatrix(GridI(Index).Row, 0) = GridI(Index).Row
        
        If Index = ITEMNF_TIPO_PECA Then
            gobjNFiscal.ColItensNF.Remove GridI(ITEMNF_TIPO_PECA).Row
        Else
            gobjNFiscal.colItensSRV.Remove GridI(ITEMNF_TIPO_SERVICO).Row
        End If
        
'       Call gobjTelaProjetoInfo.Remove_Linha(iItemAtual)
        
        For iLinhasExistentesAnterior = 1 To objGridI(Index).iLinhasExistentes
            If gobjNFiscal.ColItensNF(iLinhasExistentesAnterior).iPossuiGrade = MARCADO Then
                GridI(Index).TextMatrix(iLinhasExistentesAnterior, 0) = "# " & iLinhasExistentesAnterior
            Else
                GridI(Index).TextMatrix(iLinhasExistentesAnterior, 0) = iLinhasExistentesAnterior
            End If
        Next
        
        GridI(Index).TextMatrix(iLinhasExistentesAnterior, 0) = iLinhasExistentesAnterior
        '*******************************
        
        'Indica que houve alteração nos dados base para o cálculo de comissões
        iComissoesAlterada = REGISTRO_ALTERADO

'        If objGridAlocacoes.iLinhasExistentes > 0 And Index = ITEMNF_TIPO_PECA Then
'            'Remove as alocações do Pedido
'            Call Remove_Alocacoes(iItemAtual)
'            'Atualiza o número do item nas outras alocações
'            For iIndice = 1 To objGridAlocacoes.iLinhasExistentes
'                If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))) > 0 Then If CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) > iItemAtual Then GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col) = CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) - 1
'            Next
'        End If
        
        If objGridI(Index).iLinhasExistentes <> 0 Then
            Call PrecoTotal_Calcula(Index, objGridI(Index).iLinhasExistentes, False)
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
        Call SubTotal_Calcula
        'coloca na tela
        'ValorProdutos.Caption = Format(dValorProdutos - StrParaDbl(ValorDesconto.Text), "Standard")
        
        If objGridI(Index).iLinhasExistentes = 0 Then Call Grid_Limpa(objGridParcelas)

        'Ataliza a tributação
        If Index = ITEMNF_TIPO_PECA Then
            lErro = gobjTribTab.Exclusao_Item_Grid(iItemAtual)
        Else
            lErro = gobjTribTab.Exclusao_Item_Grid(iItemAtual, True)
        End If
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Recalcula o valor total
        Call ValorTotal_Calcula
    
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Rastreamento
        lErro = gobjRastreamento.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    
    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206022
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_FATURA_INTEGRAL", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206023)
            
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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridParcelas.Name

                lErro = Saida_Celula_GridParcelas(objGridInt)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            'Se for o GridI
            Case GridI(ITEMNF_TIPO_PECA).Name

                lErro = Saida_Celula_GridI(ITEMNF_TIPO_PECA, objGridInt)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case GridI(ITEMNF_TIPO_SERVICO).Name

                lErro = Saida_Celula_GridI(ITEMNF_TIPO_SERVICO, objGridInt)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            Case GridAP.Name
            
                'Verifica qual a coluna do Grid em questão
                Select Case objGridInt.objGrid.Col
    
                    Case iGrid_APSel_Col
    
                        lErro = Saida_Celula_Padrao(objGridInt, APSel)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                End Select

'
'            'Se for o GridComissoes
'            Case GridComissoes.Name
'
'                lErro = objTabComissoes.Saida_Celula_GridComissoes(objGridInt)
'                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'            'Rastreamento
'            Case GridRastro.Name
'
'                lErro = gobjRastreamento.Saida_Celula()
'                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 206024

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 206024
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206025)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridParcelas(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridParcelas

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col
    
        'Data de Vencimento
        Case iGrid_Vencimento_Col
            lErro = Saida_Celula_DataVencimento(objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'VAlor da Parcela
        Case iGrid_ValorParcela_Col
            lErro = Saida_Celula_ValorParcela(objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        Case iGrid_TipoPagto_Col
            lErro = Saida_Celula_TipoPagto(objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'Descontos da PArcela
        Case iGrid_Desc1Codigo_Col, iGrid_Desc2Codigo_Col, iGrid_Desc3Codigo_Col
            lErro = Saida_Celula_TipoDesconto(objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'Datas de desconto da Parcela
        Case iGrid_Desc1Ate_Col, iGrid_Desc2Ate_Col, iGrid_Desc3Ate_Col
            lErro = Saida_Celula_DescontoData(objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'VAlores dos descontos da parcela
        Case iGrid_Desc1Valor_Col, iGrid_Desc2Valor_Col, iGrid_Desc3Valor_Col
            lErro = Saida_Celula_DescontoValor(objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'Percentuais de desconto da parcela.
        Case iGrid_Desc1Percentual_Col, iGrid_Desc2Percentual_Col, iGrid_Desc3Percentual_Col
            lErro = Saida_Celula_DescontoPerc(objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
    End Select
    
    Saida_Celula_GridParcelas = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_GridParcelas:

    Saida_Celula_GridParcelas = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206026)
            
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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
         dtDataVencimento = CDate(DataVencimento.Text)
         
        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataReferencia.ClipText)) > 0 Then
            dtDataReferencia = StrParaDate(DataReferencia.Text)
            If dtDataVencimento < dtDataReferencia Then gError 206027
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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iCriouLinha Then
        'Coloca desconto padrao (le em CPRConfig)
        lErro = Preenche_DescontoPadrao(GridParcelas.Row)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    Saida_Celula_DataVencimento = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_DataVencimento:

    Saida_Celula_DataVencimento = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 206027
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR_REFERENCIA", gErr, dtDataVencimento, GridParcelas.Row, dtDataReferencia)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206028)
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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
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
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
    Saida_Celula_ValorParcela = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_ValorParcela:

    Saida_Celula_ValorParcela = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206029)
        
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
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError ERRO_SEM_MENSAGEM
        
            'Não foi encontrado
            If lErro = 25085 Then gError 206030
            If lErro = 25086 Then gError 206031

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_TipoDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDesconto:

    Saida_Celula_TipoDesconto = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 206030
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 206031
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", gErr, objGridInt.objControle.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206032)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        'Se a data de vencimento estiver preenchida
        If Len(Trim(DataEmissao.ClipText)) = 0 Then
            'critica se DataDesconto ultrapassa DataVencimento
            If StrParaDate(objGridInt.objControle.Text) < StrParaDate(DataEmissao.ClipText) Then gError 206033
        End If

            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))) > 0 Then
            dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))
            If StrParaDate(objGridInt.objControle) > dtDataVencimento Then gError 206034
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_DescontoData = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoData:

    Saida_Celula_DescontoData = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 206033
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_INFERIOR_DATA_EMISSAO", gErr, CDate(objGridInt.objControle.Text))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 206034
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_SUPERIOR_DATA_VENCIMENTO", gErr, CDate(objGridInt.objControle.Text), dtDataVencimento)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206035)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
            Set objParcRec = New ClassParcelaReceber
            gobjNFiscal.ColParcelaReceber.AddObj objParcRec
        
        End If
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
    Saida_Celula_DescontoValor = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_DescontoValor:

    Saida_Celula_DescontoValor = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206036)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_DescontoPerc = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoPerc:

    Saida_Celula_DescontoPerc = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206037)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridI(Index As Integer, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridI

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col
        'Quantidade
        Case iG_Qtd_Col(Index)
            lErro = Saida_Celula_Quantidade(Index, objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'Preco Unitário
        Case iG_VlrUnit_Col(Index)
            lErro = Saida_Celula_PrecoUnitario(Index, objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'Percentual de Desconto
        Case iG_PercDesc_Col(Index)
            lErro = Saida_Celula_PercentDesc(Index, objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'VAlor de desconto
        Case iG_Desc_Col(Index)
            lErro = Saida_Celula_Desconto(Index, objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Case iG_UM_Col(Index)
            lErro = Saida_Celula_UM(Index, objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
   
    End Select

    Saida_Celula_GridI = SUCESSO

    Exit Function

Erro_Saida_Celula_GridI:

    Saida_Celula_GridI = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206038)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(Index As Integer, objGridInt As AdmGrid) As Long

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

    Set objGridInt.objControle = Quantidade(Index)

    iSairCelula = False
    
    lErro = CF("Produto_Formata", GridI(Index).TextMatrix(GridI(Index).Row, iG_Prod_Col(Index)), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objProduto.sCodigo = sProdutoFormatado

    'Lê o produto da linha passada por iLinha do GridI
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
    'Verifica o preenchimento de Quantidade
    If Len(Trim(Quantidade(Index).ClipText)) > 0 Then
    
        'VAlida a quantidae informada
        If giComplemento = DESMARCADO Then
            lErro = Valor_Positivo_Critica(Quantidade(Index).Text)
        Else
            lErro = Valor_NaoNegativo_Critica(Quantidade(Index).Text)
        End If
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Recolhe a quantidade informada
        dQuantidade = CDbl(Quantidade(Index).Text)
        Quantidade(Index).Text = Formata_Estoque(dQuantidade)
        
        'Recolhe a quantidade anteriormente digitada
        dQuantidadeAnterior = StrParaDbl(GridI(Index).TextMatrix(GridI(Index).Row, iG_Qtd_Col(Index)))
        
        'Verifica se a quantidade foi alterada
        If dQuantidadeAnterior = dQuantidade Then iSairCelula = True
        'Se a quantidade foi alterada verifica se é um valor que pode ser faturado
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    dPrecoUnitario = StrParaDbl(GridI(Index).TextMatrix(GridI(Index).Row, iG_VlrUnit_Col(Index)))

    Call PrecoTotal_Calcula(Index, GridI(Index).Row, False)
    
    If Index = ITEMNF_TIPO_PECA Then
        lErro = gobjTribTab.Alteracao_Item_Grid(GridI(Index).Row)
    Else
        lErro = gobjTribTab.Alteracao_Item_Grid(GridI(Index).Row, True)
    End If
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Atualiza o valor total
    Call ValorTotal_Calcula

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206039)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoUnitario(Index As Integer, objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double
Dim dQuantCancelada As Double

On Error GoTo Erro_Saida_Celula_PrecoUnitario

    Set objGridInt.objControle = PrecoUnitario(Index)

    'Verifica o preenchimento de Preço Unitário
    If Len(Trim(PrecoUnitario(Index).Text)) > 0 Then

        If giComplemento = MARCADO Then
            lErro = Valor_NaoNegativo_Critica(PrecoUnitario(Index).Text)
        Else
            lErro = Valor_Positivo_Critica(PrecoUnitario(Index).Text)
        End If
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        PrecoUnitario(Index).Text = Format(PrecoUnitario(Index).Text, gobjFAT.sFormatoPrecoUnitario)
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call PrecoTotal_Calcula(Index, GridI(Index).Row, False)

    If Index = ITEMNF_TIPO_PECA Then
        lErro = gobjTribTab.Alteracao_Item_Grid(GridI(Index).Row)
    Else
        lErro = gobjTribTab.Alteracao_Item_Grid(GridI(Index).Row, True)
    End If
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Atualiza o valor total
    Call ValorTotal_Calcula

    Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206040)

    End Select

    Exit Function

End Function

Function Saida_Celula_PercentDesc(Index As Integer, objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual de Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim dPercentDescAnterior As Double

On Error GoTo Erro_Saida_Celula_PercentDesc

    Set objGridInt.objControle = PercentDesc(Index)
    
    dPercentDesc = StrParaDbl(PercentDesc(Index).Text)
    
    dPercentDescAnterior = PercentParaDbl(GridI(Index).TextMatrix(GridI(Index).Row, iG_PercDesc_Col(Index))) * 100
    
    'verifica se o percentual está preenchido
    If Len(Trim(PercentDesc(Index).Text)) > 0 Then
    
        'Critica a procentagem
        lErro = Porcentagem_Critica(PercentDesc(Index).Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
               
        If dPercentDesc <> dPercentDescAnterior Then
        
            'Verifica se o percentual é de 100%
            If dPercentDesc = 100 Then gError 206041
    
            PercentDesc(Index).Text = Format(dPercentDesc, "Fixed")
    
        End If
    Else
        GridI(Index).TextMatrix(GridI(Index).Row, iG_Desc_Col(Index)) = ""
    End If
    
    If Abs(dPercentDesc - dPercentDescAnterior) > DELTA_VALORMONETARIO2 Then
        iDescontoAlterado = REGISTRO_ALTERADO
    End If

    If Len(Trim(GridI(Index).TextMatrix(GridI(Index).Row, iG_VlrUnit_Col(Index)))) > 0 Then dPrecoUnitario = CDbl(GridI(Index).TextMatrix(GridI(Index).Row, iG_VlrUnit_Col(Index)))

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If dPrecoUnitario > 0 Then
    
        Call PrecoTotal_Calcula(Index, GridI(Index).Row, False)
    
        If Index = ITEMNF_TIPO_PECA Then
            lErro = gobjTribTab.Alteracao_Item_Grid(GridI(Index).Row)
        Else
            lErro = gobjTribTab.Alteracao_Item_Grid(GridI(Index).Row, True)
        End If
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        Call ValorTotal_Calcula

    End If

    Saida_Celula_PercentDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDesc:

    Saida_Celula_PercentDesc = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 206041
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206042)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Desconto(Index As Integer, objGridInt As AdmGrid) As Long
'Faz a crítica da célula desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dPrecoTotal As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim iDescontoAlterado As Integer

On Error GoTo Erro_Saida_Celula_Desconto

    Set objGridInt.objControle = Desconto(Index)

    iDescontoAlterado = False

    'Verifica o preenchimento de Desconto
    If Len(Trim(Desconto(Index).ClipText)) > 0 Then
    
        'Faz a critica do Desconot
        lErro = Valor_NaoNegativo_Critica(Desconto(Index).Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'recolhe o valor do desconto informado
        dDesconto = CDbl(Desconto(Index).Text)
    End If
    
    'Verifica se o valor de desconto foi alterado
    If StrParaDbl(GridI(Index).TextMatrix(GridI(Index).Row, iG_Desc_Col(Index))) <> dDesconto Then iDescontoAlterado = True
    
    'Se o desconto foi alterado
    If iDescontoAlterado = True Then
        iDescontoAlterado = REGISTRO_ALTERADO
        'recolhe a quantidade, preço unitário  e calcula o preço total
        dQuantidade = StrParaDbl(GridI(Index).TextMatrix(GridI(Index).Row, iG_Qtd_Col(Index)))
        dPrecoUnitario = StrParaDbl(GridI(Index).TextMatrix(GridI(Index).Row, iG_VlrUnit_Col(Index)))
        dPrecoTotal = dQuantidade * dPrecoUnitario
        
        'Se o preço total for positivo
        If dPrecoTotal > 0 Then
            'Verifica se o desconto ultrapassa o preço total
            If dDesconto >= dPrecoTotal Then gError 206042
            'Calcula p percentual de desconto em cima do novo desconto
            dPercentDesc = dDesconto / dPrecoTotal
            'Coloca o Percentual de desconto na tela
            GridI(Index).TextMatrix(GridI(Index).Row, iG_PercDesc_Col(Index)) = Format(dPercentDesc, "Percent")

        End If
    End If
    
    If dDesconto <> StrParaDbl(GridI(Index).TextMatrix(GridI(Index).Row, iG_Desc_Col(Index))) Then iDescontoAlterado = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iDescontoAlterado = True Then

        Call PrecoTotal_Calcula(Index, GridI(Index).Row, False)
        
        If Index = ITEMNF_TIPO_PECA Then
            lErro = gobjTribTab.Alteracao_Item_Grid(GridI(Index).Row)
        Else
            lErro = gobjTribTab.Alteracao_Item_Grid(GridI(Index).Row, True)
        End If
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Call ValorTotal_Calcula
    
    End If
    
    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 206042
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCONTO_MAIOR_OU_IGUAL_PRECO_TOTAL", gErr, GridI(ITEMNF_TIPO_PECA).Row, dDesconto, dPrecoTotal)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206043)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

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
Dim Index As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    If FrameItens(ITEMNF_TIPO_PECA + 1).Visible Then
        Index = ITEMNF_TIPO_PECA
    Else
        Index = ITEMNF_TIPO_SERVICO
    End If

    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", GridI(Index).TextMatrix(GridI(Index).Row, iG_Prod_Col(Index)), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
       
    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
           
        Case UnidadeMed(Index).Name

            UnidadeMed(Index).Clear

            'Guarda a UM que está no Grid
            sUM = GridI(Index).TextMatrix(GridI(Index).Row, iG_UM_Col(Index))

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                UnidadeMed(Index).Enabled = False
            Else
                UnidadeMed(Index).Enabled = True

                objProduto.sCodigo = sProdutoFormatado
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
                If lErro = 28030 Then gError 206044 'Não achou

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lâ as Unidades de Medidas da Classe do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                'Carrega a combo de UM
                For Each objUM In colSiglas
                    UnidadeMed(Index).AddItem objUM.sSigla
                Next
                'Seleciona na UM que está preenchida
                UnidadeMed(Index).Text = sUM
                If Len(Trim(sUM)) > 0 Then
                    lErro = Combo_Item_Igual(UnidadeMed(Index))
                    If lErro <> SUCESSO And lErro <> 12253 Then gError ERRO_SEM_MENSAGEM
                End If
            End If
                
        'Nas demais
        Case PercentDesc(Index).Name, Desconto(Index).Name
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case Quantidade(Index).Name

            If iProdutoPreenchido = PRODUTO_VAZIO Or left(GridI(ITEMNF_TIPO_PECA).TextMatrix(iLinha, 0), 1) = "#" Then
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
            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))) = 0 Then
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
    
        Case PrecoUnitario(Index).Name
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case APSel.Name
                objControl.Enabled = True
        
        Case APDataSol.Name, APSolic.Name, APOS.Name, APProduto.Name, APDescProd.Name, APPeca.Name, APDescPeca.Name, APDet.Name, APTipo.Name, APUM.Name, APQtd.Name, APGarantia.Name, APCodigo.Name
                objControl.Enabled = True
        
    End Select

'    If Index = ITEMNF_TIPO_PECA Then
'        lErro = gobjRastreamento.Rotina_Grid_Enable(iLinha, objControl, iCaminho)
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'    End If
    
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 206044
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206045)

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

    Call Grid_Trata_Tecla1(KeyCode, objGridComissoes)
    
    If KeyCode = vbKeyDelete Then
        Call objTabComissoes.Soma_Percentual
        Call objTabComissoes.Soma_Valor
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
    iLinhasExistentesAnterior = objGridParcelas.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)
    
    If objGridParcelas.iLinhasExistentes < iLinhasExistentesAnterior Then

        If gobjNFiscal.ColParcelaReceber.Count >= GridParcelas.Row Then gobjNFiscal.ColParcelaPagar.Remove (GridParcelas.Row)
    
    End If

    Exit Sub
    
Erro_GridParcelas_KeyDown:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206046)
            
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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscalFatura()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206047)

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
Dim iNFSE As Integer

On Error GoTo Erro_Gravar_Registro
   
    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 206049
    If Len(Trim(Cliente.Text)) = 0 Then gError 206052
    If Len(Trim(Filial.Text)) = 0 Then gError 206052
    If Len(Trim(Serie.Text)) = 0 Then gError 206053
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 206054
    If Len(Trim(gobjTribTab.sNatOpInterna)) = 0 Then gError 206055
    If Len(Trim(FilialEntrega.Text)) = 0 Then gError 206056

    'Verifica se tem itens no grid
    If objGridI(ITEMNF_TIPO_PECA).iLinhasExistentes + objGridI(ITEMNF_TIPO_SERVICO).iLinhasExistentes = 0 Then gError 206057

    If Len(Trim(PesoLiquido.ClipText)) > 0 And Len(Trim(PesoBruto.ClipText)) > 0 Then
        If CDbl(PesoLiquido.Text) > CDbl(PesoBruto.Text) Then gError 206058
    End If

    'Verifica se a Data de Saida e menor que a Data de Emissão
    If Len(Trim(DataSaida.ClipText)) <> 0 Then
        If CDate(DataSaida.Text) < CDate(DataEmissao.Text) Then gError 206059
    End If

    dValorTotal = CDbl(IIf(Len(Trim(ValorTotal.Caption)) > 0, ValorTotal.Caption, 0))

    'Se o total for negativo --> Erro
    If dValorTotal < 0 Then gError 206060

    'Verifica se o Valor do Desconto é maior que a soma dos demais valores
    If Len(Trim(ValorDesconto.Text)) > 0 Then dValorDesconto = CDbl(ValorDesconto.Text)
    
    lErro = DescontoGlobal_Testa(dValorDesconto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = gobjTribTab.Valida_Dados
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
    
    'Não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 206061

    'Valida os dados do grid de itens
    lErro = Valida_Dados_GridI()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'VAlida os dados do grid de comissões
    lErro = objTabComissoes.Valida_Grid_Comissoes()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Valida os dados do grid de parcelas
    lErro = Valida_Grid_Parcelas()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'    lErro = gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
'    If lErro <> SUCESSO Then gError 83397

    Set objNFiscal.objRastreamento = gobjRastreamento

    'recolhe os dados da tela
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Recolhe os dados dos grids da tela
    lErro = Move_Grids_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    GL_objMDIForm.MousePointer = vbHourglass
        
    'Valida os dados de Tributação
    lErro = gobjTribTab.Valida_TributacaoNF_Gravacao(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'se for pra verificar limite de credito
    If gobjCRFAT.iCreditoVerificaLimite = VERIFICAR_LIMITE_CREDITO Then
        
        If Len(Trim(NFiscal.Caption)) = 0 Then
            'verifica se o cliente tem crédito
            lErro = CF("NFiscal_Testa_Credito", objNFiscal)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
    
    End If
    
'    'mover a parte do rastreamento
'    lErro = gobjRastreamento.Move_Rastro_Memoria(objNFiscal)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'    lErro = gobjTelaProjetoInfo.Valida_Dados
'    If lErro <> SUCESSO Then gError 181786
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataEmissao.Text))
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("NFiscal_Valida_MargContr", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("NFiscal_Valida_Diversos", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
 
    'Grava a Nota Fiscal Fatura Pedido
    lErro = CF("NFiscalFatGarSRV_Grava", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_NFISCAL
        gobjAnotacao.sID = CStr(objNFiscal.lNumIntDoc)
        gobjAnotacao.dtDataAlteracao = gdtDataHoje
        
        lErro = CF("Anotacoes_Grava", gobjAnotacao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    If Len(Trim(NFiscal.Caption)) = 0 Then vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_NUMERO_NOTA_GRAVADA", objNFiscal.lNumNotaFiscal)
    
'    If giTipoVersao = VERSAO_FULL Then
'        If ImprimeRomaneio.Value = vbChecked Then
'            Call CF("NFiscal_Imprime_Romaneio", objNFiscal)
'        End If
'    End If
    
    If ImprimeNFGravacao.Value = vbChecked Then
        
        'Dispara função para imprimir nota fiscal
        lErro = NotaFiscal_Imprime(objNFiscal)
        If lErro <> SUCESSO And lErro <> 102262 Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    If ISSerieEletronica(objNFiscal.sSerie) Then
        lErro = CF("NFE_Grava1", objNFiscal.lNumIntDoc, objTipoDocInfo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    lErro = CF("Verifica_NFiscal_Servico_Eletronica", objNFiscal, iNFSE)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If iNFSE = 1 Then
        lErro = CF("NFSE_Grava1", objNFiscal.lNumIntDoc, objTipoDocInfo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    Gravar_Registro = SUCESSO

    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 206048
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PEDIDOSRV_TRAZIDO", gErr)
    
        Case 206049
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
        
        Case 206050
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)
            
        Case 206051
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)
            
        Case 206052
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 206053
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 206054
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 206055
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 206056
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)

        Case 206057
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)

        Case 206058
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr, CDbl(PesoLiquido.Text), CDbl(PesoBruto.Text))

        Case 206059
            Call Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_ANTERIOR_DATAEMISSAO", gErr, DataSaida, DataEmissao)
        
        Case 206060
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)

        Case 206061
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206062)

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

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 206064
    If Len(Trim(Cliente.Text)) = 0 Then gError 206067
    If Len(Trim(Serie.Text)) = 0 Then gError 206068
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 206069

    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
    
    'le o tipo da origem
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 206071
    
    'se origem for de empresa é interna
    If objTipoDocInfo.iOrigem = DOCINFO_EMPRESA Then
    
        'Verifica se a existe nota fiscal está cadastrada
        lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
        If lErro <> SUCESSO And lErro <> 43676 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = 43676 Then gError 206072
        
'        If gobjCRFAT.iUsaNFe = MARCADO And objNFiscal.iRecibo = 0 Then
'
'            lErro = CF("NFeFedLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'            If iEnviado = 1 Then gError 206073
'
'        ElseIf gobjCRFAT.iUsaNFSE = MARCADO And objNFiscal.iRecibo <> 0 Then
'
'            lErro = CF("RPSWEBLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'            If iEnviado = 1 Then gError 207084
'
'        End If

    End If

    'pede confirmacao
    vbMsg = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_NFISCAL", objNFiscal.lNumNotaFiscal)
    If vbMsg = vbYes Then

        'Lê os itens da nota fiscal
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Faz o cancelamento de uma nota fiscal de Saida
        lErro = CF("NotaFiscalSaida_Excluir", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Limpa a Tela
        lErro = Limpa_Tela_NFiscalFatura()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        'fecha o comando de setas
        Call ComandoSeta_Fechar(Me.Name)

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 206063
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PEDIDOSRV_TRAZIDO", gErr)

        Case 206064
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
       
        Case 206065
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)
            
        Case 206066
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)
            
        Case 206067
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 206068
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 206069
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 206071
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)
             
        Case 206072
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
       
        Case 206073
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CADASTRADA_NAO_EXCLUI", gErr)
       
        Case 206074
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA2", gErr, objNFiscal.lNumPedidoVenda, objNFiscal.iFilialPedido, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
       
        Case 207084
            Call Rotina_Erro(vbOKOnly, "ERRO_NFSE_CADASTRADA_NAO_EXCLUI", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206075)
        
    End Select

    Exit Sub

End Sub

Function Valida_Dados_GridI() As Long

Dim iIndice As Integer
Dim lErro As Long
Dim iTipo As Integer

On Error GoTo Erro_Valida_Dados_GridI

    For iTipo = ITEMNF_TIPO_SERVICO To ITEMNF_TIPO_PECA
    
        'Valida os dados do Grid
        For iIndice = 1 To objGridI(iTipo).iLinhasExistentes
    
            If Len(Trim(GridI(iTipo).TextMatrix(iIndice, iG_Qtd_Col(iTipo)))) = 0 Then gError 206076
            
            If StrParaDbl(GridI(iTipo).TextMatrix(iIndice, iG_VlrUnit_Col(iTipo))) = 0 Then gError 206077
            
        Next
    Next
    
    Valida_Dados_GridI = SUCESSO
    
    Exit Function

Erro_Valida_Dados_GridI:

    Valida_Dados_GridI = gErr
    
    Select Case gErr
    
        Case 206076
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 206077
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206078)
    
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
    If objGridParcelas.iLinhasExistentes = 0 Then gError 206079
    
    dSomaParcelas = 0
    
    'Para cada Parcela do grid de parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        dtDataEmissao = StrParaDate(DataEmissao.Text)
        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        
        'verifica se o vencimento e o valor da parcela estão preenchidos
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) = 0 Then gError 206080
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) = 0 Then gError 206081
        
        'verifica se a data de vencimento da parcela é menor que a data de emissão
        'If dtDataVencimento < dtDataEmissao Then gError 51008
        lErro = CF("NF_Valida_Data_Vencimento", dtDataVencimento, dtDataEmissao, iIndice)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Se o desconto 1 da parcela está preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))) > 0 Then
            iDesconto = 1
            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
            'Verifica se a data do desconto está preenchida
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))) = 0 Then gError 206082
                        
            dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
            'Valida a data do desconto
            If dtDataDesconto > dtDataVencimento Then gError 206083
            'Se o tipo de desconto for por valor
            If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                'Verifica se o valor do desconto está preenchido
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))) = 0 Then gError 206084
            Else
                'Verifica se o percentual de desconto está preenchido
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Percentual_Col))) = 0 Then gError 206085
            End If
            'Se o desconto 2 está preenchido
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))) > 0 Then
                iDesconto = 2
                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
                'Verifica se a data de desconto está preenchida
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))) = 0 Then gError 206086
                'Verifica se a data de desconto está ordenada ou se é igual ao desconto anterior
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 206087
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 206088
                'Se o desconto for do tipo valor
                If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                    'Verifica se o valor está preenchido
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))) = 0 Then gError 206089
                Else
                    'Verifica se o percentual está preenchido
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col))) = 0 Then gError 206090
                End If
                'Valida a data de desconto
                dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
                If dtDataDesconto > dtDataVencimento Then gError 206091
                'Se o desconto 3 está preenchido
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))) > 0 Then
                    iDesconto = 3
                    iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
                    'Verifica se a data de desconto está preenchida
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))) = 0 Then gError 206092
                    'Verifica se a data de desconto está ordenada ou se é igual a do desconto anterior
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 206093
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 206094
                    'Se o desconto for do tipo valor
                    If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                        'Verifica se valor de desconto está preenchido
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))) = 0 Then gError 206095
                    Else
                        'verifica se o percentuial de desconto está preenchido
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Percentual_Col))) = 0 Then gError 206096
                    End If
                    'Valida a data de desconto
                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
                    If dtDataDesconto > dtDataVencimento Then gError 206097
                End If
            End If
        End If
        'Verifica se as data de vencimentos das parcelas estão ordenadas
        If iIndice > 1 Then If CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col)) < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_Col)) Then gError 206098
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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
    Next
    
    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If Abs(dValorPagar - dSomaParcelas) > DELTA_VALORMONETARIO Then gError 206099

    Valida_Grid_Parcelas = SUCESSO
    
    Exit Function
    
Erro_Valida_Grid_Parcelas:

    Valida_Grid_Parcelas = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 206079
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_PARCELA_COBRANCA", gErr)
        
        Case 206080
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_NAO_INFORMADA", gErr, iIndice)
            
        Case 206081
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PARCELA_COBRANCA_NAO_INFORMADO", gErr, iIndice)

        Case 206082, 206086, 206092
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_NAO_PREENCHIDA", gErr, iDesconto, iIndice)

        Case 206083, 206091, 206097
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESC_PARCELA_SUP_DATA_VENCIMENTO", gErr, dtDataDesconto, iDesconto, iIndice)

        Case 206084, 206085, 206089, 206090, 206095, 206096
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_PARCELA_NAO_PREENCHIDO", gErr, iDesconto, iIndice)

        Case 206087, 206093
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTOS_DESORDENADAS", gErr, iIndice)

        Case 206088, 206094
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTO_IGUAIS", gErr, iDesconto - 1, iDesconto, iIndice)

        Case 206098
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_COBRANCA_NAO_ORDENADA", gErr)
        
        Case 206099
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_COBRANCA_INVALIDA", gErr)
  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206100)
            
    End Select
    
    Exit Function

End Function

Function Move_Tela_Memoria(objNFiscal As ClassNFiscal) As Long
'Recolhe os dados da tela

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objProjetoInfo As Object

On Error GoTo Erro_Move_Tela_Memoria

    Set objNFiscal.objInfoUsu = gcolAPDet
            
    'Verifica se o Cliente esta preenchido
    If Len(Trim(Cliente.Text)) > 0 Then
    
        objCliente.sNomeReduzido = Cliente.Text
        
        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
        
        'Se não achou o Cliente --> erro
        If lErro <> SUCESSO Then gError 206101
        'Guarda código do Cliente
        objNFiscal.lCliente = objCliente.lCodigo
        
    End If
    
    'Armazena o responsável pelo frete
    If Emitente.Value Then
        objNFiscal.iFreteRespons = FRETE_EMITENTE
    Else
        objNFiscal.iFreteRespons = FRETE_DESTINATARIO
    End If
    
    If Recibo.Value = vbChecked Then
        objNFiscal.iRecibo = MARCADO
    Else
        objNFiscal.iRecibo = DESMARCADO
    End If
    
    'Recolhe os demais dados da tela
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    
    If Len(Trim(NFiscal.Caption)) Then
        objNFiscal.lNumNotaFiscal = CLng(NFiscal.Caption)
    Else
        objNFiscal.lNumNotaFiscal = 0
    End If
    
    lErro = CF2(Me, "NFCF_Move_DataSaida_Memoria", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
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
    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.dtDataEntrada = DATA_NULA
    objNFiscal.dtDataVencimento = DATA_NULA
    objNFiscal.iCondicaoPagto = Codigo_Extrai(CondicaoPagamento.Text)
    objNFiscal.iStatus = STATUS_LANCADO
    objNFiscal.dtDataReferencia = StrParaDate(DataReferencia.Text)
            
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objNFiscal.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objNFiscal.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
    
'    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_NF)
'    If lErro <> SUCESSO Then gError 181761
    
    Set objNFiscal.objProjetoInfo = objProjetoInfo
    
    Set objNFiscal.objInfoAdic = gobjInfoAdic
    objNFiscal.dValorItens = StrParaDbl(ValorProdutos2.Caption)
    objNFiscal.dValorDescontoItens = StrParaDbl(ValorDescontoItens.Text)
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206101
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206102)
            
    End Select
            
    Exit Function
        
End Function

Private Function Move_Grids_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os dados dos Grids para memória

Dim lErro As Long
Dim iTipo As Integer

On Error GoTo Erro_Move_Grids_Memoria
    
    For iTipo = ITEMNF_TIPO_SERVICO To ITEMNF_TIPO_PECA
        'Chama Move_GridI_Memoria
        lErro = Move_GridI_Memoria(iTipo, objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    Next

'    'Chama Move_GridAlocacao_Memoria
'    lErro = Move_GridAlocacao_Memoria(objNFiscal)
'    If lErro <> SUCESSO Then gError 51015

'    'Chama Move_GridComissoes_Memoria
'    lErro = objTabComissoes.Move_TabComissoes_Memoria(objNFiscal, NOTA_FISCAL)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Chama Move_GridParcelas_Memoria
    lErro = Move_GridParcelas_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Mover a parte de tributação
    lErro = gobjTribTab.PreencherNF(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Move_Grids_Memoria = SUCESSO

    Exit Function

Erro_Move_Grids_Memoria:

    Move_Grids_Memoria = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206103)

    End Select

    Exit Function

End Function

Public Function Move_GridItens_Memoria(objNFiscal As ClassNFiscal) As Long
    Move_GridItens_Memoria = Move_GridI_Memoria(ITEMNF_TIPO_PECA, objNFiscal)
End Function

Public Function Move_GridItensSRV_Memoria(objNFiscal As ClassNFiscal) As Long
    Move_GridItensSRV_Memoria = Move_GridI_Memoria(ITEMNF_TIPO_SERVICO, objNFiscal)
End Function

Public Function Move_GridI_Memoria(Index As Integer, objNFiscal As ClassNFiscal) As Long
'Move os Itens do Grid para a Memória

Dim iIndice As Integer
Dim lErro As Long
Dim objItemNF As ClassItemNF, objItemNFAux As ClassItemNF
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lTamanho As Long
Dim colAlocacoes As ColAlocacoesItemNF
Dim objDataEntrega As ClassDataEntrega
Dim dQuantidade As Double

On Error GoTo Erro_Move_GridI_Memoria

    'Para cada linha existente do Grid
    For iIndice = 1 To objGridI(Index).iLinhasExistentes

        Set objItemNF = New ClassItemNF

        'Verifica se o Produto está preenchido
        lErro = CF("Produto_Formata", GridI(Index).TextMatrix(iIndice, iG_Prod_Col(Index)), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Armazena produto
        objItemNF.sProduto = sProdutoFormatado
      
        'Guarda os demais campos do Grid em objItemNF
        objItemNF.sDescricaoItem = GridI(Index).TextMatrix(iIndice, iG_DescProd_Col(Index))
        objItemNF.sUnidadeMed = GridI(Index).TextMatrix(iIndice, iG_UM_Col(Index))
        objItemNF.dQuantidade = StrParaDbl(GridI(Index).TextMatrix(iIndice, iG_Qtd_Col(Index)))
        objItemNF.dPrecoUnitario = StrParaDbl(GridI(Index).TextMatrix(iIndice, iG_VlrUnit_Col(Index)))
        
        lTamanho = Len(Trim(GridI(Index).TextMatrix(iIndice, iG_PercDesc_Col(Index))))

        If lTamanho > 0 Then objItemNF.dPercDesc = CDbl(Format(GridI(Index).TextMatrix(iIndice, iG_PercDesc_Col(Index)), "General Number"))

        objItemNF.dValorDesconto = StrParaDbl(GridI(Index).TextMatrix(iIndice, iG_Desc_Col(Index)))
        
        'Adiciona na coleção de Ítens
        With objItemNF
            If Index = ITEMNF_TIPO_PECA Then
                Set objItemNFAux = objNFiscal.ColItensNF.Add(0, iIndice, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, 0, 0, 0, 0, 0, colAlocacoes, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .sCcl, STATUS_LANCADO, 0, "", 0, 0, .iControleEstoque, .sUMEstoque, .iClasseUM, 0)
                objItemNFAux.iTipo = Index
                objItemNF.dtDataEntrega = DATA_NULA
            Else
                objItemNF.dtDataEntrega = DATA_NULA 'StrParaDate(GridI(Index).TextMatrix(iIndice, iG_Entrega_Col(Index)))
                Set objItemNFAux = objItemNF
                objItemNFAux.iTipo = Index
                objItemNFAux.iItem = iIndice + objGridI(ITEMNF_TIPO_PECA).iLinhasExistentes
                objNFiscal.colItensSRV.Add objItemNF
            End If
        End With
                
    Next

    Move_GridI_Memoria = SUCESSO

    Exit Function

Erro_Move_GridI_Memoria:

    Move_GridI_Memoria = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 206104
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTSOLICITADA_DIFERE1", gErr, iIndice, GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_Prod_Col), objItemNFAux.dQuantidade, dQuantidade)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206105)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
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
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206106)
            
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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objNFiscal.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "TipoNFiscal", objNFiscal.iTipoNFiscal, 0, "TipoNFiscal"
    colCampoValor.Add "Serie", objNFiscal.sSerie, STRING_BUFFER_MAX_TEXTO, "Serie"
    colCampoValor.Add "NumNotaFiscal", objNFiscal.lNumNotaFiscal, 0, "NumNotaFiscal"
    colCampoValor.Add "DataEmissao", objNFiscal.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "Cliente", objNFiscal.lCliente, 0, "Cliente"
    colCampoValor.Add "FilialCli", objNFiscal.iFilialCli, 0, "FilialCli"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "TipoNFiscal", OP_IGUAL, Codigo_Extrai(TipoNFiscal.Text)
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    'colSelecao.Add "Status", OP_DIFERENTE, STATUS_CANCELADO
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206107)

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
        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
        
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206108)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

Dim lErro As Long

On Error GoTo Erro_Form_Activate

    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Call TelaIndice_Preenche(Me)

    Exit Sub
     
Erro_Form_Activate:

    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206109)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Public Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objNFiscal Is Nothing) Then

        lErro = CargaPosFormLoad(True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
        
        If lErro <> SUCESSO Then gError 206110

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

        'Lê o Tipo de Documento com o Código de objNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
        
        'Se não encontrar --> Erro
        If lErro = 31415 Then gError 206111
        
'        If giTipoVersao = VERSAO_FULL Then
            If objTipoDocInfo.sNomeTelaNFiscal <> NOME_TELA_NFISCALFATURAGARSRV Then gError 206112
'        End If
        
        'Traz a nota para a tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    iAlterado = 0
    iComissoesAlterada = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206110
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case 206111
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 206112
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_DIFERENTE_NF_FATURA_PEDIDO", gErr, objNFiscal.iTipoNFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206113)

    End Select
    
    iAlterado = 0
    iComissoesAlterada = 0

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'feito em virtude do redimensionamento da msg da nota p/ inpal
    Call CF("Redimensiona_Tam_MsgNF", Mensagem)

    Parent.HelpContextID = IDH_NF_FATURA_PEDIDO_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Nota Fiscal Fatura - Itens em Garantia"
    Call Form_Load
    
End Function

Public Function Name() As String
    Name = "NFiscalFatGarSRV"
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
    If dValorProdutos + dValorFrete + dValorSeguro + dValorDespesas <= 0 Then gError 206114
    
    DescontoGlobal_Testa = SUCESSO
     
    Exit Function
    
Erro_DescontoGlobal_Testa:

    DescontoGlobal_Testa = gErr
     
    Select Case gErr
                  
        Case 206114
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_ULTRAPASSOU_SOMA_VALORES", gErr, dValorDesconto, dValorProdutos + dValorDesconto + dValorFrete + dValorSeguro + dValorDespesas) 'Alterado por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206115)
     
    End Select
     
    Exit Function

End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Serie Then
            Call SerieLabel_Click
        ElseIf Me.ActiveControl Is NFiscal Then
            Call NFiscalLabel_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        ElseIf Me.ActiveControl Is TranspRedespacho Then
            Call TranspRedLabel_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call BotaoVendedores_Click
'        ElseIf Me.ActiveControl Is Projeto Then
'            Call LabelProjeto_Click
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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
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
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206116)
            
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
Dim iTipo As Integer

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
                   
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
            iTipo = ITEMNF_TIPO_PECA
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                objMnemonicoValor.colValor.Add StrParaDbl(GridI(iTipo).TextMatrix(iLinha, iG_Desc_Col(iTipo)))
            Next
            
        Case DESCONTO1SRV
            iTipo = ITEMNF_TIPO_SERVICO
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                objMnemonicoValor.colValor.Add StrParaDbl(GridI(iTipo).TextMatrix(iLinha, iG_Desc_Col(iTipo)))
            Next
            
        Case DESCRICAO_ITEM
            iTipo = ITEMNF_TIPO_SERVICO
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                If Len(GridI(iTipo).TextMatrix(iLinha, iG_DescProd_Col(iTipo))) > 0 Then
                    objMnemonicoValor.colValor.Add GridI(iTipo).TextMatrix(iLinha, iG_DescProd_Col(iTipo))
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case DESCRICAO_ITEMSRV
            iTipo = ITEMNF_TIPO_SERVICO
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                If Len(GridI(iTipo).TextMatrix(iLinha, iG_DescProd_Col(iTipo))) > 0 Then
                    objMnemonicoValor.colValor.Add GridI(iTipo).TextMatrix(iLinha, iG_DescProd_Col(iTipo))
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

        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Else
                
                    sContaMascarada = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaMascarada
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case CLIENTE1
            If Len(Cliente.Text) > 0 Then
                objMnemonicoValor.colValor.Add Cliente.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case CLIENTE_CODIGO
            If Len(Cliente.Text) > 0 Then
            
                objCliente.sNomeReduzido = Cliente.Text
                
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
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
            iTipo = ITEMNF_TIPO_PECA
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                objMnemonicoValor.colValor.Add GridI(iTipo).TextMatrix(iLinha, iG_Prod_Col(iTipo))
            Next
                        
        Case PRODUTO1SRV
            iTipo = ITEMNF_TIPO_SERVICO
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                objMnemonicoValor.colValor.Add GridI(iTipo).TextMatrix(iLinha, iG_Prod_Col(iTipo))
            Next
                        
        Case QUANTIDADE1
            iTipo = ITEMNF_TIPO_PECA
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                If Len(GridI(iTipo).TextMatrix(iLinha, iG_Qtd_Col(iTipo))) > 0 Then
                
                    'Guarda os valores que serão passados como parâmetros em UMEstoque_Conversao
                    sProduto = GridI(iTipo).TextMatrix(iLinha, iG_Prod_Col(iTipo))
                    sUM = GridI(iTipo).TextMatrix(iLinha, iG_UM_Col(iTipo))
                    dQuantidade = StrParaDbl(GridI(iTipo).TextMatrix(iLinha, iG_Qtd_Col(iTipo)))
                    
                    'Converte a quantidade para UM padrão estoque
                    lErro = CF("UMEstoque_Conversao", sProduto, sUM, dQuantidade, dQuantidadeConvertida)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
                
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case QUANTIDADE1SRV
            iTipo = ITEMNF_TIPO_SERVICO
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                If Len(GridI(iTipo).TextMatrix(iLinha, iG_Qtd_Col(iTipo))) > 0 Then
                
                    'Guarda os valores que serão passados como parâmetros em UMEstoque_Conversao
                    sProduto = GridI(iTipo).TextMatrix(iLinha, iG_Prod_Col(iTipo))
                    sUM = GridI(iTipo).TextMatrix(iLinha, iG_UM_Col(iTipo))
                    dQuantidade = StrParaDbl(GridI(iTipo).TextMatrix(iLinha, iG_Qtd_Col(iTipo)))
                    
                    'Converte a quantidade para UM padrão estoque
                    lErro = CF("UMEstoque_Conversao", sProduto, sUM, dQuantidade, dQuantidadeConvertida)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
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
            iTipo = ITEMNF_TIPO_PECA
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                objMnemonicoValor.colValor.Add GridI(iTipo).TextMatrix(iLinha, iG_UM_Col(iTipo))
            Next
                        
        Case UNIDADE_MEDSRV
            iTipo = ITEMNF_TIPO_SERVICO
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                objMnemonicoValor.colValor.Add GridI(iTipo).TextMatrix(iLinha, iG_UM_Col(iTipo))
            Next
            
        Case VALOR_TOTAL
            If Len(ValorTotal.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorTotal.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case PRECO_UNITARIO
            iTipo = ITEMNF_TIPO_PECA
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                objMnemonicoValor.colValor.Add StrParaDbl(GridI(iTipo).TextMatrix(iLinha, iG_VlrUnit_Col(iTipo)))
            Next
            
        Case PRECO_UNITARIOSRV
            iTipo = ITEMNF_TIPO_SERVICO
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                objMnemonicoValor.colValor.Add StrParaDbl(GridI(iTipo).TextMatrix(iLinha, iG_VlrUnit_Col(iTipo)))
            Next
            
        Case PRECO_TOTAL
            iTipo = ITEMNF_TIPO_PECA
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                objMnemonicoValor.colValor.Add StrParaDbl(GridI(iTipo).TextMatrix(iLinha, iG_PrecoTot_Col(iTipo)))
            Next
            
        Case PRECO_TOTALSRV
            iTipo = ITEMNF_TIPO_SERVICO
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                objMnemonicoValor.colValor.Add StrParaDbl(GridI(iTipo).TextMatrix(iLinha, iG_PrecoTot_Col(iTipo)))
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
                               
        Case PIS_DEBITO_TOTAL
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dPISCredito
        
        Case COFINS_DEBITO_TOTAL
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dCOFINSCredito
        
        Case ISS_VALOR
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dISSValor

        Case ISS_INCLUSO
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.iISSIncluso <> 0
        
        Case CONTACONTABILAPLIC
        
            iTipo = ITEMNF_TIPO_PECA
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                If Len(GridI(iTipo).TextMatrix(iLinha, iG_Prod_Col(iTipo))) > 0 Then
                
                    lErro = CF("Produto_Formata", GridI(iTipo).TextMatrix(iLinha, iG_Prod_Col(iTipo)), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
                    sContaMascarada = String(STRING_CONTA, 0)

                    lErro = Mascara_MascararConta(objProduto.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                    objMnemonicoValor.colValor.Add sContaMascarada
                    
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case CONTACONTABILAPLICSRV
        
            iTipo = ITEMNF_TIPO_SERVICO
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                If Len(GridI(iTipo).TextMatrix(iLinha, iG_Prod_Col(iTipo))) > 0 Then
                
                    lErro = CF("Produto_Formata", GridI(iTipo).TextMatrix(iLinha, iG_Prod_Col(iTipo)), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
                    sContaMascarada = String(STRING_CONTA, 0)

                    lErro = Mascara_MascararConta(objProduto.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                    objMnemonicoValor.colValor.Add sContaMascarada
                    
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        
        Case CONTACONTABILCATPROD
                
            iTipo = ITEMNF_TIPO_PECA
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                
                If Len(GridI(iTipo).TextMatrix(iLinha, iG_Prod_Col(iTipo))) > 0 Then
                
                    lErro = CF("Produto_Formata", GridI(iTipo).TextMatrix(iLinha, iG_Prod_Col(iTipo)), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                    objProdutoCategoria.sCategoria = objMnemonicoValor.vParam(1)
                    
                    objProdutoCategoria.sProduto = sProdutoFormatado
                    
                    Set objCatProdItem = New ClassCategoriaProdutoItem
                    
                    lErro = CF("ProdutoCategoria_Le_Item", objProdutoCategoria, objCatProdItem)
                    If lErro <> SUCESSO And lErro <> 183548 Then gError ERRO_SEM_MENSAGEM
                    
                    If lErro = SUCESSO Then
                        
                        Set objPlanoConta = New ClassPlanoConta
                    
                        'le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta
                        lErro = CF("PlanoConta_Le_ContaSimples", objCatProdItem.dvalor1, objPlanoConta)
                        If lErro <> SUCESSO And lErro <> 5451 Then gError ERRO_SEM_MENSAGEM
                    
                        If lErro = SUCESSO Then
                    
                            sContaMascarada = String(STRING_CONTA, 0)
        
                            lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
                            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
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
        
        Case CONTACONTABILCATSRV
                
            iTipo = ITEMNF_TIPO_SERVICO
            For iLinha = 1 To objGridI(iTipo).iLinhasExistentes
                
                If Len(GridI(iTipo).TextMatrix(iLinha, iG_Prod_Col(iTipo))) > 0 Then
                
                    lErro = CF("Produto_Formata", GridI(iTipo).TextMatrix(iLinha, iG_Prod_Col(iTipo)), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                    objProdutoCategoria.sCategoria = objMnemonicoValor.vParam(1)
                    
                    objProdutoCategoria.sProduto = sProdutoFormatado
                    
                    Set objCatProdItem = New ClassCategoriaProdutoItem
                    
                    lErro = CF("ProdutoCategoria_Le_Item", objProdutoCategoria, objCatProdItem)
                    If lErro <> SUCESSO And lErro <> 183548 Then gError ERRO_SEM_MENSAGEM
                    
                    If lErro = SUCESSO Then
                        
                        Set objPlanoConta = New ClassPlanoConta
                    
                        'le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta
                        lErro = CF("PlanoConta_Le_ContaSimples", objCatProdItem.dvalor1, objPlanoConta)
                        If lErro <> SUCESSO And lErro <> 5451 Then gError ERRO_SEM_MENSAGEM
                    
                        If lErro = SUCESSO Then
                    
                            sContaMascarada = String(STRING_CONTA, 0)
        
                            lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
                            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
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
            If lErro <> SUCESSO Then gError 206117

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 206117
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206118)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Compara com data de emissão
        If Len(Trim(DataEmissao.ClipText)) > 0 Then

            dtDataEmissao = CDate(DataEmissao.Text)
            dtDataReferencia = CDate(DataReferencia.Text)

            lErro = CF("NF_Valida_Data_Referencia", dtDataReferencia, dtDataEmissao)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206119)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
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
        
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206120)
            
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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call DataReferencia_Validate(bCancel)
    
    If bCancel = True Then DataReferencia.Text = sData
    
    Exit Sub

Erro_BotaoDataReferenciaDown_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206124)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call DataReferencia_Validate(bCancel)

    If bCancel = True Then DataReferencia.Text = sData
    
    Exit Sub

Erro_BotaoDataReferenciaUp_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206125)

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

Public Sub BotaoConsultaTitRec_Click()
'Abre uma tela para consulta do DocCPR vinculado à nota fiscal

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoConsultaTitRec_Click

    'Verifica se todos os campos necessários para se efetuar a consulta foram preenchidos
    lErro = Critica_CamposNecessarios_ConsultaTitulo()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Guarda no objNFiscal os dados necessários para consultar o título
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.iTipoDocInfo = Codigo_Extrai(TipoNFiscal.Text)
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.dtDataEntrada = DATA_NULA
        
    'Guarda em objFornecedor o nome reduzido do Fornecedor
    objCliente.sNomeReduzido = Cliente.Text
    
    'Lê o código do Fornecedor a partir do nome reduzido obtido na tela
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
    
    'Se não encontrou o fornecedor = > erro
    If lErro = 12348 Then gError 206126
    
    'Guarda no objNFiscal o código do fornecedor
    objNFiscal.lCliente = objCliente.lCodigo
    
    'Lê o NumIntDocCPR da NFiscal e exibe o documento gerado no CPR por essa nota
    lErro = CF("NFiscal_Consulta_DocCPR", objNFiscal)
    If lErro <> SUCESSO And lErro <> 79717 Then gError ERRO_SEM_MENSAGEM
    
    'Se não encontrou a nota => erro
    If lErro = 79717 Then gError 206127
    
    Exit Sub
    
Erro_BotaoConsultaTitRec_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 206126
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)
        
        Case 206127
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA2", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206128)
    
    End Select
    
    Exit Sub

End Sub

Private Function Critica_CamposNecessarios_ConsultaTitulo() As Long
'Verifica se os campos necessários para encontrar consultar um título gerado por uma NFFatEntrada foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposNecessarios_ConsultaTitulo

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 206129
    
    'Verifica se a filial do Fornecedor foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 206130
    
    'Verifica se a Série foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 206131
    
    'Verifica se o Número da Nota foi preenchido
    If Len(Trim(NFiscal.Caption)) = 0 Then gError 206132
    
    'Verifica se a data de emissão da nota foi preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 206133
    
    Critica_CamposNecessarios_ConsultaTitulo = SUCESSO
    
    Exit Function

Erro_Critica_CamposNecessarios_ConsultaTitulo:

    Critica_CamposNecessarios_ConsultaTitulo = gErr
    
    Select Case gErr
    
        Case 206129
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 206130
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 206131
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 206132
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
        
        Case 206133
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206134)
    
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
    ComissaoAutomatica.Value = vbUnchecked
    
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
'
End Sub

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError ERRO_SEM_MENSAGEM
        
        'Não encontrou Transportadora no BD
        If lErro <> SUCESSO Then gError 206135
        
        'Encontrou Transportadora no BD, coloca no Text da Combo
        TranspRedespacho.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 206136

    Exit Sub

Erro_TranspRedespacho_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206135  'Não encontrou Transportadora no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            End If

        Case 206136
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206137)

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

Public Function Calcula_VolumeTotal() As Long
''Realiza a soma dos volumes dos produtos contido nop pedido de venda
'
'Dim objProduto As New ClassProduto
'Dim lErro As Long
'Dim iIndice As Integer
'Dim dFator As Double
'Dim sProduto As String
'Dim iPreenchido As Integer
'Dim dVolume As Double
'
'On Error GoTo Erro_Calcula_VolumeTotal
'
'    For iIndice = 1 To objGridI(ITEMNF_TIPO_PECA).iLinhasExistentes
'
'        'Retira a formatação do produto
'        lErro = CF("Produto_Formata", GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_Prod_Col(ITEMNF_TIPO_PECA)), sProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'        If iPreenchido <> PRODUTO_PREENCHIDO Then gError 206138
'
'        'preenche o código do produto
'        objProduto.sCodigo = sProduto
'
'        lErro = CF("Produto_Le", objProduto)
'        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
'
'        If lErro = 28030 Then gError 206139
'
'        'Realiza a converção para uma mesma UM
'        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_UM_Col(ITEMNF_TIPO_PECA)), objProduto.sSiglaUMEstoque, dFator)
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'        'Realiza o cálculo do volume
'        dVolume = dVolume + ((StrParaDbl(GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_Qtd_Col(ITEMNF_TIPO_PECA))) * dFator) * (objProduto.dComprimento * objProduto.dEspessura * objProduto.dLargura))
'
'    Next
'
'    If dVolume <> 0 Then
'        Cubagem.Text = dVolume
'    Else
'        Cubagem.Text = ""
'    End If
'
'    Calcula_VolumeTotal = SUCESSO
'
'    Exit Function
'
'Erro_Calcula_VolumeTotal:
'
'    Calcula_VolumeTotal = gErr
'
'    Select Case gErr
'
'        Case ERRO_SEM_MENSAGEM
'
'        Case 206138
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
'
'        Case 206139
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206140)
'
'    End Select
'
'    Exit Function
    
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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Coloca o valor formatado na Tela
    Cubagem.Text = Format(Cubagem.Text, "Fixed")

    Exit Sub

Erro_Cubagem_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206141)

    End Select

    Exit Sub

End Sub

'@@@@@ para ser customizada
Function Botao_Estoque(ByVal objEstoqueProd As ClassEstoqueProduto, ByVal colSelecao As Collection, sSelecao As String) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim iAlmoxarifado As Integer
Dim objProduto As New ClassProduto, sProdutoEnxuto As String

On Error GoTo Erro_Botao_Estoque

    sSelecao = ""

    'Verifica se a linha do Grid é válida
    If GridI(ITEMNF_TIPO_PECA).Row = 0 Then gError 206142

    'Verifica se o Produto na linha do Grid está preenchido
    lErro = CF("Produto_Formata", GridI(ITEMNF_TIPO_PECA).TextMatrix(GridI(ITEMNF_TIPO_PECA).Row, iG_Prod_Col(ITEMNF_TIPO_PECA)), sProdutoFormatado, iProdutoPreenchido)

    If iProdutoPreenchido <> PRODUTO_VAZIO Then
        objEstoqueProd.sProduto = sProdutoFormatado

        objProduto.sCodigo = sProdutoFormatado
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        If lErro <> SUCESSO Then gError 206144
        
        If Len(Trim(objProduto.sGrade)) <> 0 Then
        
            Call Mascara_RetornaProdutoEnxuto(sProdutoFormatado, sProdutoEnxuto)
            sSelecao = "Produto LIKE '" & sProdutoEnxuto & "%' AND FilialEmpresa = " & giFilialEmpresa
            Call Chama_Tela("EstProdLista_Consulta", colSelecao, objEstoqueProd, objEventoEstoqueProd, sSelecao)
            gError ERRO_SEM_MENSAGEM
            
        End If
        
        colSelecao.Add sProdutoFormatado

        'Lê o Almoxarifado Padrão do Produto
        lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, sProdutoFormatado, iAlmoxarifado)
        If lErro = SUCESSO Then objEstoqueProd.iAlmoxarifado = iAlmoxarifado

    Else
        gError 206143
    End If
 
    Botao_Estoque = SUCESSO
    
    Exit Function

Erro_Botao_Estoque:

    Botao_Estoque = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 206142
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 206143
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 206144
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206145)

    End Select

    Exit Function
    
End Function

Private Function Saida_Celula_UM(Index As Integer, objGridInt As AdmGrid) As Long
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

    Set objGridInt.objControle = UnidadeMed(Index)
    
    'recolhe a UM anteriormente escolhida
    sUmAnterior = GridI(Index).TextMatrix(GridI(Index).Row, iG_UM_Col(Index))
    
    'coloca no grid a UM atual selecionda
    GridI(Index).TextMatrix(GridI(Index).Row, iG_UM_Col(Index)) = UnidadeMed(Index).Text
    
    If Index = ITEMNF_TIPO_PECA Then
        gobjNFiscal.ColItensNF(GridI(Index).Row).sUnidadeMed = UnidadeMed(Index).Text
        objProduto.sCodigo = gobjNFiscal.ColItensNF(GridI(Index).Row).sProduto
    Else
        gobjNFiscal.colItensSRV(GridI(Index).Row).sUnidadeMed = UnidadeMed(Index).Text
        objProduto.sCodigo = gobjNFiscal.colItensSRV(GridI(Index).Row).sProduto
    End If
    
'    'se houve alteração na UM
'    If sUmAnterior <> UnidadeMed(Index).Text And Frame1(TAB_Almoxarifado).Enabled = True Then
'
'        lErro = CF("Produto_Le", objProduto)
'        If lErro <> SUCESSO And lErro <> 28030 Then gError 177583
'
'        If gobjNFiscal.ColItensNF(GridI(Index).Row).iPossuiGrade <> MARCADO And objProduto.iKitVendaComp <> MARCADO Then
'
'            lErro = Atualiza_UM(Index, GridI(Index).Row, sUmAnterior, UnidadeMed(Index).Text)
'            If lErro <> SUCESSO Then gError 131927
'
'            If Index = ITEMNF_TIPO_PECA Then
'                'Tenta fazer uma alocação p\ o item de acordo com a nova UM
'                lErro = Alocacao_Processa
'                If lErro <> SUCESSO Then gError 131901
'            End If
'
'        Else
'
'            If objProduto.iKitVendaComp = MARCADO Then
'
'                objKitVenda.sProduto = objProduto.sCodigo
'                objKitVenda.sUM = UnidadeMed(Index).Text
'                objKitVenda.dQuantidade = StrParaDbl(GridI(Index).TextMatrix(GridI(Index).Row, iG_Qtd_Col(Index)))
'
'                If Index = ITEMNF_TIPO_PECA Then
'                    lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjNFiscal.ColItensNF(GridI(Index).Row).colItensRomaneioGrade)
'                    If lErro <> SUCESSO Then gError 177584
'                End If
'
'            End If
'
'            If Index = ITEMNF_TIPO_PECA Then
'                For Each objItemRomaneio In gobjNFiscal.ColItensNF(GridI(Index).Row).colItensRomaneioGrade
'                    lErro = Alocacao_Processa_ItemGrade(objItemRomaneio, gobjNFiscal.ColItensNF(GridI(Index).Row))
'                    If lErro <> SUCESSO Then gError 131902
'                Next
'
'                Call Atualiza_Grid_Alocacao(gobjNFiscal.ColItensNF.Item(GridI(Index).Row))
'            End If
'
'        End If
'    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Se a Um selecionada agora é diferente da anterior
    If sUmAnterior <> UnidadeMed(Index).Text Then

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206146)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Atualiza_UM(ByVal Index As Integer, ByVal iLinha As Integer, ByVal sUmAnterior As String, ByVal sUMNova As String) As Long
'Atualiza quantidades em funcao de troca de UM

Dim lErro As Long
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double, dPrecoUnitario As Double
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Atualiza_UM

    lErro = CF("Produto_Formata", GridI(Index).TextMatrix(iLinha, iG_Prod_Col(Index)), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objProduto.sCodigo = sProdutoFormatado

    'Lê o produto da linha passada por iLinha do GridI
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
    If lErro = 28030 Then gError 206147

    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUmAnterior, sUMNova, dFator)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Atualiza o Grid
    GridI(ITEMNF_TIPO_PECA).TextMatrix(iLinha, iG_Qtd_Col(Index)) = Formata_Estoque(Arredonda_Estoque(StrParaDbl(GridI(Index).TextMatrix(iLinha, iG_Qtd_Col(Index))) * dFator))

     'Preço unitário
    dPrecoUnitario = StrParaDbl(GridI(Index).TextMatrix(iLinha, iG_VlrUnit_Col(Index)))
      
    GridI(ITEMNF_TIPO_PECA).TextMatrix(iLinha, iG_VlrUnit_Col) = Format(Arredonda_Moeda(dPrecoUnitario / dFator), gobjFAT.sFormatoPrecoUnitario)
        
    If StrParaDbl(GridI(Index).TextMatrix(iLinha, iG_VlrUnit_Col(Index))) > 0 Then
        Call PrecoTotal_Calcula(Index, GridI(Index).Row, False)
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If

    Atualiza_UM = SUCESSO
    
    Exit Function

Erro_Atualiza_UM:

    Atualiza_UM = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206147
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206148)

    End Select

    Exit Function

End Function

Public Sub UnidadeMed_Click(Index As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UnidadeMed_GotFocus(Index As Integer)
    Call Grid_Campo_Recebe_Foco(objGridI(Index))
End Sub

Public Sub UnidadeMed_KeyPress(Index As Integer, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridI(Index))
End Sub

Public Sub UnidadeMed_Validate(Index As Integer, Cancel As Boolean)

Dim lErro As Long

    Set objGridI(Index).objControle = UnidadeMed(Index)
    lErro = Grid_Campo_Libera_Foco(objGridI(Index))
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long, sID As String
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Anotacao_Extrai

    'Lê os dados da Tela NFiscal
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If objNFiscal.lNumNotaFiscal <> 0 And objNFiscal.sSerie <> "" And objNFiscal.dtDataEmissao <> DATA_NULA Then
    
        lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
        If lErro <> SUCESSO And lErro <> 43676 Then gError ERRO_SEM_MENSAGEM
        
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
          
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206149)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206150)
     
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
    
    'Carrega na combo só os Tipos ligados essa tela
    For Each objTipoDocInfo In colTipoDocInfo
        If UCase(objTipoDocInfo.sNomeTelaNFiscal) = UCase(Me.Name) Then
            TipoNFiscal.AddItem CStr(objTipoDocInfo.iCodigo) & SEPARADOR & objTipoDocInfo.sNomeReduzido
            TipoNFiscal.ItemData(TipoNFiscal.NewIndex) = objTipoDocInfo.iCodigo
            'se for o tipo padrao, seleciona-o
            If objTipoDocInfo.iPadrao = Padrao Then TipoNFiscal.ListIndex = TipoNFiscal.NewIndex
        End If
    Next
    giIndexTipoNF = TipoNFiscal.ListIndex
    
    Carrega_TiposDocInfo = SUCESSO

    Exit Function

Erro_Carrega_TiposDocInfo:

    Carrega_TiposDocInfo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206151)

    End Select

    Exit Function

End Function

Public Sub TipoNFiscal_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TipoNFiscal_Click

    'Frame1(TAB_Almoxarifado).Enabled = True

    If TipoNFiscal.ListIndex = -1 Then Exit Sub

    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    gobjRastreamento.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM

    'se não estiver cadastrado ==> erro
    If lErro = 31415 Then gError 206152
    
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then
        giComplemento = MARCADO
    Else
        giComplemento = DESMARCADO
    End If
    
'    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Or (objTipoDocInfo.iTipoMovtoEstoque = 0 And objTipoDocInfo.iTipoMovtoEstoque = 0) Then
'        Frame1(TAB_Almoxarifado).Enabled = False
'    Else
'        Frame1(TAB_Almoxarifado).Enabled = True
'    End If
    
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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'se o cliente não estiver preenchido ==> preencher com a natureza padrao o campo natureza da operacao
    If Len(Trim(Cliente.Text)) = 0 Then gobjTribTab.sNatOpInterna = NaturezaOp_Conv34(objTipoDocInfo.sNaturezaOperacaoPadrao, MaskedParaDate(DataEmissao))

    ''*** incluidos p/tratamento de tributacao *******************************
    If gbLimpandoTela = False Then
        If objTipoDocInfo.iCodigo = DOCINFO_NFFPSRV Then Call gobjTribTab.NatOpInterna_Automatica
        Call gobjTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
        If IsDate(DataEmissao.Text) Then Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
        Call gobjTribTab.Serie_Alterada(Serie.Text)
    End If
    '*** fim tributacao

    'subtipo contabil
    'ReInicializa a parte de contabilidade devido que o tipo implica na mudança dos modelos de contabilização
    lErro = objContabil.Contabil_Processa_Alteracao_Subtipo(objTipoDocInfo.iSubTipoContabil)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_TipoNFiscal_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206152
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206153)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM
    
    'Se não conseguir --> Erro
    If lErro <> SUCESSO Then gError 206154

    Exit Sub

Erro_TipoNFiscal_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 206154
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206155)

    End Select

    Exit Sub

End Sub

Private Sub Seleciona_Tipo_Venda()
'Seleciona na combo de tipos de notas fiscais o tipo "Venda"

Dim iIndice As Integer

    For iIndice = 0 To TipoNFiscal.ListCount - 1
        If TipoNFiscal.ItemData(iIndice) = DOCINFO_NFFPSRV Then
            TipoNFiscal.ListIndex = iIndice
            Exit For
        End If
    Next
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206156)
    
    End Select
    
    Exit Sub

End Sub

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
        If Len(Trim(Cliente.Text)) = 0 Then gError 206157
        If Len(Trim(Serie.Text)) = 0 Then gError 206158
        If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 206159
        If Len(Trim(NFiscal.Caption)) = 0 Then gError 206160
    
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
        objCliente.sNomeReduzido = Trim(Cliente.Text)
        
        'Lê o cliente a partir do nome reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
        
        'Se não encontrou o cliente => erro
        If lErro = 12348 Then gError 206161
        
        'Guarda o código do cliente no obj
        objNFiscal.lCliente = objCliente.lCodigo
        
        'Guarda o código da filial do cliente
        objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
    
    End If
    
    'Lê os dados da nota fiscal para verificar se a mesma existe no BD
    lErro = CF("NFiscal_Le_1", objNFiscal)
    If lErro <> SUCESSO And lErro <> 83971 Then gError ERRO_SEM_MENSAGEM

    'Se não encontrou => erro, pois não é possível imprimir uma nota fiscal inexistente
    If lErro = 83971 Then gError 206162
    
    objSerie.sSerie = objNFiscal.sSerie
    objSerie.iFilialEmpresa = objNFiscal.iFilialEmpresa
    
    'Lock na Tabela Série para a Impreessão
    lErro = CF("Serie_Lock_ImpressaoNFiscal", objSerie)
    If lErro <> SUCESSO And lErro <> 60387 Then gError ERRO_SEM_MENSAGEM

    'Se não encontrou a Série --> ERRO
    If lErro = 60387 Then gError 206163
    
    'Dispara a impressão do relatório
    lErro = objRelatorio.ExecutarDireto("Emissão das Notas Fiscais Fatura", "", 1, objSerie.sNomeTsk, "NNFISCALINIC", objNFiscal.lNumNotaFiscal, "NNFISCALFIM", objNFiscal.lNumNotaFiscal, "TSERIE", objNFiscal.sSerie)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault
    
    NotaFiscal_Imprime = SUCESSO
    
    Exit Function

Erro_NotaFiscal_Imprime:

    GL_objMDIForm.MousePointer = vbDefault

    NotaFiscal_Imprime = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
           
        Case 206157
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 206158
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
       
        Case 206159
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
            
        Case 206160
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMNOTAFISCAL_NAO_PREENCHIDO", gErr)
            
        Case 206161
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
        
        Case 206162
            Call Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_NAO_ENCONTRADO", gErr)
            
        Case 206163
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206164)
    
    End Select
    
End Function

Public Sub BotaoImprimir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoImprimir_Click

    'Dispara função para imprimir nota fiscal
    lErro = NotaFiscal_Imprime()
    If lErro <> SUCESSO And lErro <> 102262 Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206165)

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
        
'        Case BOTAO_PECA
'            Call BotaoPeca_Click
'
'        Case BOTAO_SERVICO
'            Call BotaoServico_Click
'
        Case BOTAO_SOLICITACAO
            Call BotaoSolicitacao_Click

        Case BOTAO_OS
            Call BotaoOS_Click

        Case BOTAO_APONTAMENTO
            Call BotaoApontamento_Click

        Case BOTAO_GARANTIA
            Call BotaoGarantia_Click

        Case BOTAO_ESTOQUEPROD
            Call BotaoEstoqueProd_Click
        
        Case BOTAO_LOCALIZACAO
            Call BotaoLocalizacao_Click
        
        Case BOTAO_SERIE
            Call BotaoSerie_Click
        
        Case BOTAO_LOTES
            Call BotaoLotes_Click
            
        Case BOTAO_GERARITENS
            Call BotaoGerarItens_Click
        
    End Select

    Exit Sub
    
Erro_Botao_Click:

    Select Case gErr

        Case Else
        
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206166)
    
    End Select
    
    Exit Sub

End Sub
'
'Public Function BotaoProjetos_Click() As Long
'        Call gobjTelaProjetoInfo.BotaoProjetos_Click
'End Function
'
'Public Function LabelProjeto_Click() As Long
'    Call gobjTelaProjetoInfo.LabelProjeto_Click
'End Function
'
'Public Function Projeto_GotFocus() As Long
'    Call MaskEdBox_TrataGotFocus(Projeto, iAlterado)
'End Function
'
'Public Function Projeto_Change() As Long
'    iAlterado = REGISTRO_ALTERADO
'End Function
'
'Public Function Projeto_Validate(Cancel As Boolean) As Long
'    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
'End Function
'
'Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim objItensPRJCR As ClassItensPRJCR
'
'On Error GoTo Erro_Obter_ItensPRJCR
'
'    For iIndice = 1 To objGridI(ITEMNF_TIPO_PECA).iLinhasExistentes
'
'        Set objItensPRJCR = New ClassItensPRJCR
'
'        objItensPRJCR.sItem = CStr(iIndice)
'        objItensPRJCR.sDescricao = GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_Prod_Col)
'
'        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_Qtd_Col))
'        objItensPRJCR.dValorOriginal = StrParaDbl(GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_PrecoTot_Col))
'        objItensPRJCR.sObservacao = GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_Prod_Col) & SEPARADOR & GridI(ITEMNF_TIPO_PECA).TextMatrix(iIndice, iG_DescProd_Col)
'
'        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NF
'
'        colItensPRJ.Add objItensPRJCR
'
'    Next
'
'    Call gobjTribTab.Preenche_Itens_Projeto_NF(colItensPRJ)
'
'    Obter_ItensPRJCR = SUCESSO
'
'    Exit Function
'
'Erro_Obter_ItensPRJCR:
'
'    Obter_ItensPRJCR = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157006)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function Preenche_Itens_Projeto_Aux(ByVal colItensPRJ As Collection, ByVal sItem As String, ByVal sDescricao As String, ByVal dValor As Double, ByVal iTipo As Integer, Optional ByVal iNegativo As Integer = 0, Optional ByVal sObs As String = "") As Long
'
'Dim objItensPRJCR As New ClassItensPRJCR
'
'On Error GoTo Erro_Preenche_Itens_Projeto_Aux
'
'    objItensPRJCR.sItem = sItem
'    objItensPRJCR.sDescricao = sDescricao
'    objItensPRJCR.dValorOriginal = dValor
'    objItensPRJCR.iNegativo = iNegativo
'    objItensPRJCR.iTipoOrigem = iTipo
'    objItensPRJCR.sObservacao = sObs
'
'    colItensPRJ.Add objItensPRJCR
'
'    Exit Function
'
'Erro_Preenche_Itens_Projeto_Aux:
'
'    Preenche_Itens_Projeto_Aux = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157006)
'
'    End Select
'
'    Exit Function
'
'End Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206167)

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
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError ERRO_SEM_MENSAGEM

            'Não foi encontrado
            If lErro = 25085 Then gError 206168
            If lErro = 25086 Then gError 206169

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_TipoPagto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoPagto:

    Saida_Celula_TipoPagto = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 206168
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMAPAGTO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 206169
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMAPAGTO_NAO_ENCONTRADO", gErr, objGridInt.objControle.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206170)

    End Select

    Exit Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206171)

    End Select

    Exit Sub

End Sub

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

Public Sub GridAP_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridAP, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAP, iAlterado)
    End If
    
End Sub

Public Sub GridAP_GotFocus()
    Call Grid_Recebe_Foco(objGridAP)
End Sub

Public Sub GridAP_EnterCell()
    Call Grid_Entrada_Celula(objGridAP, iAlterado)
End Sub

Public Sub GridAP_LeaveCell()
    Call Saida_Celula(objGridAP)
End Sub

Public Sub GridAP_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer

On Error GoTo Erro_GridAP_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGridAP.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridAP)
    
    Exit Sub
    
Erro_GridAP_KeyDown:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206046)
            
    End Select

    Exit Sub
    
End Sub

Public Sub GridAP_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAP, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAP, iAlterado)
    End If

End Sub

Public Sub GridAP_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridAP)
End Sub

Public Sub GridAP_RowColChange()
    Call Grid_RowColChange(objGridAP)
End Sub

Public Sub GridAP_Scroll()
    Call Grid_Scroll(objGridAP)
End Sub

Private Function Inicializa_Grid_AP(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Alocações

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Solic.")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição Prod.")
    objGridInt.colColuna.Add ("Garantia")
    objGridInt.colColuna.Add ("OS")
    objGridInt.colColuna.Add ("Apont.")
    objGridInt.colColuna.Add ("Peça\Serv.")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Detalhamento")
    objGridInt.colColuna.Add ("UM")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (APSel.Name)
    objGridInt.colCampo.Add (APDataSol.Name)
    objGridInt.colCampo.Add (APSolic.Name)
    objGridInt.colCampo.Add (APProduto.Name)
    objGridInt.colCampo.Add (APDescProd.Name)
    objGridInt.colCampo.Add (APGarantia.Name)
    objGridInt.colCampo.Add (APOS.Name)
    objGridInt.colCampo.Add (APCodigo.Name)
    objGridInt.colCampo.Add (APPeca.Name)
    objGridInt.colCampo.Add (APDescPeca.Name)
    objGridInt.colCampo.Add (APTipo.Name)
    objGridInt.colCampo.Add (APDet.Name)
    objGridInt.colCampo.Add (APUM.Name)
    objGridInt.colCampo.Add (APQtd.Name)
    objGridInt.colCampo.Add (APNumIntDoc.Name)

    'Colunas da Grid
    iGrid_APSel_Col = 1
    iGrid_APDataSol_Col = 2
    iGrid_APSolic_Col = 3
    iGrid_APProduto_Col = 4
    iGrid_APDescProd_Col = 5
    iGrid_APGarantia_Col = 6
    iGrid_APOS_Col = 7
    iGrid_APCodigo_Col = 8
    iGrid_APPeca_Col = 9
    iGrid_APDescPeca_Col = 10
    iGrid_APTipo_Col = 11
    iGrid_APDet_Col = 12
    iGrid_APUM_Col = 13
    iGrid_APQtd_Col = 14
    iGrid_APNumIntDoc_Col = 15

    'Grid do GridInterno
    objGridInt.objGrid = GridAP

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ALOCACOES + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 15
    
    'Largura da primeira coluna
    GridAP.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    APNumIntDoc.Width = 0

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    GridAP.ColWidth(iGrid_APNumIntDoc_Col) = 0

    Inicializa_Grid_AP = SUCESSO

    Exit Function

End Function

Public Sub APSel_Change()
    iFiltroAlterado = REGISTRO_ALTERADO
End Sub

Public Sub APSel_Click()
    iFiltroAlterado = REGISTRO_ALTERADO
    If GridAP.Row > 0 And GridAP.Row <= objGridAP.iLinhasExistentes Then
        If GridAP.TextMatrix(GridAP.Row, iGrid_APSel_Col) = "1" Then
            gcolAPDet.Item(GridAP.Row).iMarcado = MARCADO
        Else
            gcolAPDet.Item(GridAP.Row).iMarcado = DESMARCADO
        End If
        Call BotaoGerarItens_Click
    End If
End Sub

Public Sub APSel_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAP)
End Sub

Public Sub APSel_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAP)
End Sub

Public Sub APSel_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAP.objControle = APSel
    lErro = Grid_Campo_Libera_Foco(objGridAP)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Trata_Troca_Tab() As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim lCliente As Long, iFilial As Integer
Dim colSol As New Collection

On Error GoTo Erro_Trata_Troca_Tab

    'Verifica preenchimento de Cliente
    If Len(Trim(Cliente.Text)) > 0 Then
    
        objCliente.sNomeReduzido = Cliente.Text
        'Lê o Cliente
        lErro = CF("Cliente_Le_Codigo_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
        
        'Se não encontrou o Cliente --> erro
        If lErro = 12348 Then gError 42412
    
        lCliente = objCliente.lCodigo
        
    End If

    iFilial = Codigo_Extrai(Filial.Text)

    'Se estava no tab principal
    If iFrameAtual = TAB_DadosPrincipais Then
    
        'Se o cliente ou a filial é diferente da última leitura
        If lCliente <> lFabricante Or iFilial <> iFilialFabricante Then
        
            'Se está com os dados prenchidos
            If lCliente <> 0 And iFilial <> 0 Then
            
                lErro = CF("Garantia_Le_Gastos", lCliente, iFilial, colSol)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                lErro = Traz_Apontamento_Tela(colSol)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            End If
        
            lFabricante = lCliente
            iFilialFabricante = iFilial
            iFiltroAlterado = REGISTRO_ALTERADO
        End If


    End If

'    'Se estava no tab apontamentos
'    If iFrameAtual = TAB_Apontamentos Then
'
'        If iFiltroAlterado = REGISTRO_ALTERADO Then
'
'            lErro = Traz_Itens_Tela()
'            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'            iFiltroAlterado = 0
'
'        End If
'
'    End If
    
    Trata_Troca_Tab = SUCESSO

    Exit Function

Erro_Trata_Troca_Tab:

    Trata_Troca_Tab = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206170)

    End Select

    Exit Function
    
End Function

Private Sub BotaoGerarItens_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGerarItens_Click

    lErro = Traz_Itens_Tela()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iFiltroAlterado = 0
    
    Exit Sub

Erro_BotaoGerarItens_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206170)

    End Select

    Exit Sub
    
End Sub

Private Function Traz_Itens_Tela() As Long
'Tem que agrupar e trazer os itens da NF para tela
'LEMBRAR DE: Tributação, Tabela de Preço e objetos globais

Dim lErro As Long
Dim iIndice As Integer
Dim objDet As ClassOSApDet
Dim objNF As New ClassNFiscal
Dim objItemNF As ClassItemNF
Dim bAchou As Boolean, dFator As Double
Dim iItemPeca As Integer, iItemSrv As Integer

On Error GoTo Erro_Traz_Itens_Tela

    For Each objDet In gcolAPDet
    
        If objDet.iMarcado = MARCADO Then
    
            bAchou = False
            iIndice = 0
            For Each objItemNF In objNF.colItens
                iIndice = iIndice + 1
                If objItemNF.sProduto = objDet.sProduto Then
                    lErro = CF("UM_Conversao_Trans", objDet.iClasseUM, objDet.sUM, objItemNF.sUnidadeMed, dFator)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    objItemNF.dQuantidade = objItemNF.dQuantidade + objDet.dQuantidade * dFator
                    objDet.iItemNF = objItemNF.iItemNFOrig
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
                Set objItemNF = New ClassItemNF
                objItemNF.sProduto = objDet.sProduto
                objItemNF.sDescricaoItem = objDet.sProdutoDesc
                objItemNF.sUnidadeMed = objDet.sUM
                objItemNF.dQuantidade = objDet.dQuantidade
                If objDet.iPeca = DESMARCADO Then
                    objItemNF.iTipo = ITEMNF_TIPO_SERVICO
                    iItemSrv = iItemSrv + 1
                    objDet.iItemNF = iItemSrv
                    objItemNF.iItemNFOrig = iItemSrv
                Else
                    objItemNF.iTipo = ITEMNF_TIPO_PECA
                    iItemPeca = iItemPeca + 1
                    objDet.iItemNF = iItemPeca
                    objItemNF.iItemNFOrig = iItemPeca
                End If
                objNF.ColItensNF.Add1 objItemNF
            End If
        End If
    Next
    
'    'Limpa o tab de tributacao
'    Call gobjTribTab.Limpa_Tela

    For iIndice = 1 To gobjNFiscal.ColItensNF.Count
        lErro = gobjTribTab.Exclusao_Item_Grid(iIndice)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    Next
    
    For iIndice = 1 To gobjNFiscal.colItensSRV.Count
        lErro = gobjTribTab.Exclusao_Item_Grid(iIndice, True)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    Next
    
    Set gobjNFiscal.ColItensNF = New ColItensNF
    Set gobjNFiscal.colItensSRV = New Collection
    
    iIndice = 0
    For Each objItemNF In objNF.ColItensNF
        If objItemNF.iTipo = ITEMNF_TIPO_PECA Then
            iIndice = iIndice + 1
            objItemNF.iItemNFOrig = 0
            gobjNFiscal.ColItensNF.Add1 objItemNF
            'permite que a tributacao reflita a inclusao de uma linha no grid
            lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, objItemNF.sProduto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
    Next
    
    iIndice = 0
    For Each objItemNF In objNF.ColItensNF
        If objItemNF.iTipo = ITEMNF_TIPO_SERVICO Then
            iIndice = iIndice + 1
            objItemNF.iItemNFOrig = 0
            gobjNFiscal.colItensSRV.Add objItemNF
            'permite que a tributacao reflita a inclusao de uma linha no grid
            lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, objItemNF.sProduto, True)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
    Next

    'Preenche o Grid com os itens da Nota Fiscal
    lErro = Preenche_GridI(gobjNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Preenche o Grid com os itens da Nota Fiscal
    lErro = Preenche_GridISRV(gobjNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
'    lErro = gobjTribTab.Traz_NFiscal_Tela(gobjNFiscal)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    giFilialAlterada = REGISTRO_ALTERADO
'    Call Cliente_Validate(bSGECancelDummy)
    
    If Len(Trim(TabelaPreco.Text)) > 0 Then
    
        lErro = Trata_TabelaPreco
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    Traz_Itens_Tela = SUCESSO

    Exit Function

Erro_Traz_Itens_Tela:

    Traz_Itens_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206170)

    End Select

    Exit Function
    
End Function

Private Function Traz_Apontamento_Tela(ByVal colItens As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItem As ClassItensSolicSRV
Dim objSol As ClassSolicSRV
Dim objDet As ClassOSApDet
Dim sProdMask As String

On Error GoTo Erro_Traz_Apontamento_Tela

    Call Grid_Limpa(objGridAP)
    Set gcolAPDet = New Collection
    
    For Each objDet In colItens
        Set objSol = objDet.objSol
        Set objItem = objDet.objItemSol
        If objDet.dQuantidade > DELTA_VALORMONETARIO Then
            gcolAPDet.Add objDet
            iIndice = iIndice + 1
            GridAP.TextMatrix(iIndice, iGrid_APSel_Col) = CStr(objDet.iMarcado)
            GridAP.TextMatrix(iIndice, iGrid_APSolic_Col) = CStr(objSol.lCodigo)
            GridAP.TextMatrix(iIndice, iGrid_APCodigo_Col) = CStr(objDet.lCodigoAp)
            GridAP.TextMatrix(iIndice, iGrid_APDataSol_Col) = Format(objSol.dtData, "dd/mm/yyyy")
            GridAP.TextMatrix(iIndice, iGrid_APDescPeca_Col) = objDet.sProdutoDesc
            GridAP.TextMatrix(iIndice, iGrid_APDescProd_Col) = objItem.sProdutoDesc
            GridAP.TextMatrix(iIndice, iGrid_APDet_Col) = objDet.sDet
            GridAP.TextMatrix(iIndice, iGrid_APGarantia_Col) = CStr(objItem.lGarantia)
            GridAP.TextMatrix(iIndice, iGrid_APOS_Col) = objDet.sOS
            
            lErro = Mascara_RetornaProdutoTela(objDet.sProduto, sProdMask)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            GridAP.TextMatrix(iIndice, iGrid_APPeca_Col) = sProdMask
            
            lErro = Mascara_RetornaProdutoTela(objItem.sProduto, sProdMask)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            GridAP.TextMatrix(iIndice, iGrid_APProduto_Col) = sProdMask
            
            GridAP.TextMatrix(iIndice, iGrid_APQtd_Col) = Formata_Estoque(objDet.dQuantidade)
            GridAP.TextMatrix(iIndice, iGrid_APTipo_Col) = objDet.sTipo
            GridAP.TextMatrix(iIndice, iGrid_APUM_Col) = objDet.sUM
            GridAP.TextMatrix(iIndice, iGrid_APNumIntDoc_Col) = CStr(objDet.lNumIntDoc)
        End If
    Next
    objGridAP.iLinhasExistentes = iIndice
    Call Grid_Refresh_Checkbox(objGridAP)
    
    If Not gbCarregandoTela Then Call BotaoGerarItens_Click
    
    Traz_Apontamento_Tela = SUCESSO

    Exit Function

Erro_Traz_Apontamento_Tela:

    Traz_Apontamento_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 206170)

    End Select

    Exit Function

End Function

Public Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO
    giClienteAlterado = 1

    Call Cliente_Preenche

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long, sNatOp As String, iTipoTrib As Integer
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objTipoCliente As New ClassTipoCliente
Dim objMensagem As New ClassMensagem
Dim objContato As New ClassContatos
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_Cliente_Validate

    If giClienteAlterado = 1 Then

        'Verifica se o Cliente está preenchido
        If Len(Trim(Cliente.Text)) > 0 Then

            'Busca o Cliente no BD
            lErro = TP_Cliente_Le3(Cliente, objCliente, iCodFilial, 0)
            If lErro <> SUCESSO And lErro <> 6668 And lErro <> 6676 And lErro <> 6601 And lErro <> 6704 Then gError 188988

            'Se não for um cliente
            If lErro <> SUCESSO Then

                objContato.lCodigo = objCliente.lCodigo
                objContato.sNomeReduzido = objCliente.sNomeReduzido
                objContato.sCgc = objCliente.sCgc

                'Verifica se é um contato
                lErro = TP_Contato_Le(Cliente, objContato, iCodFilial, 0)
                If lErro <> SUCESSO Then

                    'Trata como Cliente novamente
                    lErro = TP_Cliente_Le3(Cliente, objCliente, iCodFilial)
                    If lErro <> SUCESSO Then gError 188989

                Else

                    vbResult = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_COM_CONTATO", Cliente.Text, Cliente.Text)
                    If vbResult = vbNo Then gError 188990

                    'Lê o Contato no BD
                    lErro = CF("Contato_Transforma_Cliente", objContato, objCliente)
                    If lErro <> SUCESSO Then gError 188991

                End If

            End If

            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then gError 188993

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)
            Call CF("Filial_Preenche", FilialEntrega, colCodigoNome)

            If Not gbCarregandoTela Then

                If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then

                    If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ

                    'Seleciona filial na Combo Filial
                    Call CF("Filial_Seleciona", Filial, iCodFilial)
                    Call CF("Filial_Seleciona", FilialEntrega, iCodFilial)

                End If

            End If

            giFilialEntregaManual = 0

            'Se o Tipo estiver preenchido
            If objCliente.iTipo > 0 Then
                objTipoCliente.iCodigo = objCliente.iTipo
                'Lê o Tipo de Cliente
                lErro = CF("TipoCliente_Le", objTipoCliente)
                If lErro <> SUCESSO And lErro <> 19062 Then gError 188995
            End If

            'Tipo de frete
            If objCliente.iTipoFrete = TIPO_FOB Then
                Destinatario.Value = True
            ElseIf objCliente.iTipoFrete = TIPO_CIF Then
                Emitente.Value = True
            End If

            giValorDescontoManual = 0
            'Guarda o valor do desconto do cliente
            If objCliente.dDesconto > 0 Then
                gdDesconto = objCliente.dDesconto
            ElseIf objTipoCliente.dDesconto > 0 Then
                gdDesconto = objTipoCliente.dDesconto
            Else
                gdDesconto = 0
            End If

            If Not gbCarregandoTela Then

                Call DescontoGlobal_Recalcula

                'ATualiza o total com o novo desconto
                lErro = ValorTotal_Calcula()
                If lErro <> SUCESSO Then gError 188996

                'Coloca na tela a tabela do cliente
                If objCliente.iTabelaPreco > 0 Then
                    TabelaPreco.Text = objCliente.iTabelaPreco
                    Call TabelaPreco_Validate(bSGECancelDummy)
                ElseIf objTipoCliente.iTabelaPreco > 0 Then
                    TabelaPreco.Text = objTipoCliente.iTabelaPreco
                    Call TabelaPreco_Validate(bSGECancelDummy)
                End If
                'Se cobrança automática estiver selecionada preenche a CondPagto e dispara o Validate
                If CobrancaAutomatica.Value = 1 Then
                    If objCliente.iCondicaoPagto > 0 Then
                        CondicaoPagamento.Text = objCliente.iCondicaoPagto
                        Call CondicaoPagamento_Validate(bSGECancelDummy)
                    ElseIf objTipoCliente.iCondicaoPagto > 0 Then
                        CondicaoPagamento.Text = objTipoCliente.iCondicaoPagto
                        Call CondicaoPagamento_Validate(bSGECancelDummy)
                    End If

                End If

'                If objCliente.iCodMensagem > 0 Then
'                    'Lê e coloca a mensagem do Cliente na tela
'                    objMensagem.iCodigo = objCliente.iCodMensagem
'
'                    lErro = CF("Mensagem_Le", objMensagem)
'                    If lErro <> SUCESSO And lErro <> 19234 Then gError 188997
'                    If lErro <> SUCESSO Then gError 188998
'
'                    Mensagem.Text = objMensagem.sDescricao
'
'                End If

            End If

        'Se não estiver preenchido
        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            'Limpa a Combo de Filiais
            Filial.Clear
            FilialEntrega.Clear

        End If

        giClienteAlterado = 0

        If Not gbCarregandoTela Then

    ''*** incluidos p/tratamento de tributacao *******************************
            If iCodFilial <> 0 Then Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, iCodFilial) '####
    '*** fim tributacao
        End If
        
    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 188988 To 188998

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188999)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    iAlterado = REGISTRO_ALTERADO

    'Se nenhuma filial foi selecionada, sai.
    If Filial.ListIndex = -1 Then Exit Sub

    'Faz o tratamento para a filial do cliente selecionada
    lErro = Trata_FilialCliente()
    If lErro <> SUCESSO Then gError 191000

    Exit Sub

Erro_Filial_Click:

    Select Case gErr

        Case 191000

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191001)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult
Dim objCliente As New ClassCliente

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida ou alterada
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 191002

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 191003

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 191004

        If lErro = 17660 Then

            'Lê o Cliente
            objCliente.sNomeReduzido = sCliente
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 191005

            'Não encontrou Cliente
            If lErro = 12348 Then gError 191006

            objFilialCliente.lCodCliente = objCliente.lCodigo

            gError 191007

        End If

        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        lErro = Trata_FilialCliente
        If lErro <> SUCESSO Then gError 191008

    End If

    'Não encontrou a STRING
    If lErro = 6731 Then gError 191009

    giFilialAlterada = 0

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 191002, 191004, 191005, 191006, 191008

        Case 191003
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 191007
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 191009
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191010)

    End Select

    Exit Sub

End Sub


Public Sub Cliente_Formata(lCliente As Long)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Formata

    Cliente.Text = lCliente

    'Busca o Cliente no BD
    lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
    If lErro <> SUCESSO Then gError 193221

    lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 193222

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", Filial, colCodigoNome)
    Call CF("Filial_Preenche", FilialEntrega, colCodigoNome)

    'para fazer valer o que veio do bd
    giValorDescontoManual = 1
    giFilialEntregaManual = 1

    giClienteAlterado = 0

    Exit Sub

Erro_Cliente_Formata:

    Select Case gErr

        Case 193221, 193222

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193223)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Formata(objFilial As Object, iFilial As Integer)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Formata

    objFilial.Text = CStr(iFilial)
    sCliente = Cliente.Text
    objFilialCliente.iCodFilial = iFilial

    'Pesquisa se existe Filial com o código extraído
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 193224

    If lErro = 17660 Then gError 193225

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_Filial_Formata:

    Select Case gErr

        Case 193224

        Case 193225
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193226)

    End Select

    Exit Sub

End Sub

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object

On Error GoTo Erro_Cliente_Preenche

    Set objCliente = Cliente

    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 193419

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 193419

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193420)

    End Select

    Exit Sub

End Sub

Private Sub DescontoGlobal_Recalcula()

Dim dValorDesconto As Double
Dim dValorProdutos As Double
Dim lErro As Long

On Error GoTo Erro_DescontoGlobal_Recalcula
    
    If gbCarregandoTela Then Exit Sub

    PercDescontoItens.Text = Format(gdDesconto * 100, "FIXED")
    Call PercDescontoItens_Validate(bSGECancelDummy)

'    If Len(Trim(ValorProdutos.Caption)) <> 0 And IsNumeric(ValorProdutos.Caption) Then
'
'        dValorProdutos = CDbl(ValorProdutos.Caption)
'
'        'Se o cliente possui desconto e o campo desconto não foi alterado pelo usuário
'        If gdDesconto > 0 And iValorDescontoAlterado = 0 Then
'
'            Call SubTotal_Calcula(dValorProdutos)
'
'            'Calcula o valor do desconto para o cliente e coloca na tela
'            dValorDesconto = gdDesconto * dValorProdutos
'            ValorDesconto.Text = Format(dValorDesconto, "Standard")
'            iValorDescontoAlterado = 0
'
'            ValorProdutos.Caption = Format(dValorProdutos - dValorDesconto, "STANDARD")
'
'            lErro = CF2(Me, "PV_ValorDesconto_Validate")
'            If lErro <> SUCESSO Then gError 132033
'
'''*** incluidos p/tratamento de tributacao *******************************
'            Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, dValorDesconto)
''*** fim tributacao
'
'        End If
'
'    End If

    Exit Sub

Erro_DescontoGlobal_Recalcula:

    Select Case gErr

        Case 132033

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157158)

    End Select

    Exit Sub

End Sub

Function Trata_FilialCliente() As Long

Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim objVendedor As New ClassVendedor
Dim objTipoCliente As New ClassTipoCliente
Dim dValorTotal As Double
Dim dValorBase As Double
Dim objTransportadora As New ClassTransportadora
Dim dValorComissao As Double
Dim dValorEmissao As Double
Dim lErro As Long
Dim objMensagem As New ClassMensagem

On Error GoTo Erro_Trata_FilialCliente

    objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
    objCliente.sNomeReduzido = Trim(Cliente.Text)

    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Trim(Cliente.Text), objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 193184
    
    If lErro = 17660 Then gError 193185

    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 193186
    If lErro = 12348 Then gError 193187
    
    Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, objFilialCliente.iCodFilial)

    'Comissões
    If Not (gobjCRFAT.iUsaComissoesRegras = USA_REGRAS) And ComissaoAutomatica.Value = 1 Then

        'calcular as comissoes no modelo antigo...
        lErro = objTabComissoes.Comissao_Automatica_FilialCli_Exibe(objFilialCliente)
        If lErro <> SUCESSO Then gError 193188

    End If

    'Transportadora
    If Len(Trim(Transportadora.Text)) = 0 Then

        'Testa se FilialCliente tem Transportadora
        If objFilialCliente.iCodTransportadora <> 0 Then

            objTransportadora.iCodigo = objFilialCliente.iCodTransportadora
            lErro = CF("Transportadora_Le", objTransportadora)
            If lErro <> SUCESSO And lErro <> 19250 Then gError 193189
            If lErro = SUCESSO Then
                Transportadora.Text = objTransportadora.sNomeReduzido
                Call Transportadora_Validate(False)
            End If

        ElseIf objCliente.iTipo <> 0 Then

            'Testa se TipoCliente já foi lido no BD
            If objTipoCliente.iCodigo = 0 Then

                objTipoCliente.iCodigo = objCliente.iTipo
                lErro = CF("TipoCliente_Le", objTipoCliente)
                If lErro <> SUCESSO And lErro <> 19062 Then gError 193190
                If lErro <> SUCESSO Then objTipoCliente.iCodTransportadora = 0

            End If

            'Testa se TipoCliente tem Transportadora
            If objTipoCliente.iCodTransportadora <> 0 Then
                lErro = CF("Transportadora_Le", objTransportadora)
                If lErro <> SUCESSO And lErro <> 19250 Then gError 193191
                If lErro = SUCESSO Then
                    Transportadora.Text = objTransportadora.sNomeReduzido
                    Call Transportadora_Validate(False)
                End If
            End If

        End If

    End If

    'Verifica se a transportadora de Redespacho está preenchida
    If Len(Trim(TranspRedespacho.Text)) = 0 Then

        If objFilialCliente.iCodTranspRedesp <> 0 Then

            objTransportadora.iCodigo = objFilialCliente.iCodTranspRedesp
            lErro = CF("Transportadora_Le", objTransportadora)
            If lErro <> SUCESSO And lErro <> 19250 Then gError 193192
            
            If lErro <> SUCESSO Then gError 193193

            'Coloca a transportadora na tela
            TranspRedespacho.Text = objTransportadora.sNomeReduzido
            Call TranspRedespacho_Validate(bSGECancelDummy)

            RedespachoCli.Value = objFilialCliente.iDetPagFrete

        End If

    End If

    'Tipo de frete
    If objFilialCliente.iTipoFrete = TIPO_FOB Then
        Destinatario.Value = True
    ElseIf objFilialCliente.iTipoFrete = TIPO_CIF Then
        Emitente.Value = True
    End If

    If giFilialEntregaManual = 0 Then

        FilialEntrega.Text = CStr(objFilialCliente.iCodFilial)
        Call FilialEntrega_Validate(bSGECancelDummy)

        giFilialEntregaManual = 0

    End If

    '#############################################
    'Inserido por Wagner
    lErro = CF2(Me, "NFCF_Cliente_Filial_Validate", objFilialCliente)
    If lErro <> SUCESSO Then gError 193194
    '#############################################

'    '###################################
'    'Inserido por Wagner 27/09/05
'    If objFilialCliente.iCodMensagem > 0 Then
'        objMensagem.iCodigo = objFilialCliente.iCodMensagem
'        'Lê a Mensagem "default" p/ o Cliente
'        lErro = CF("Mensagem_Le", objMensagem)
'        If lErro <> SUCESSO And lErro <> 19234 Then gError 193195
'
'        If lErro = SUCESSO Then Mensagem.Text = objMensagem.sDescricao
'
'    End If
'    '###################################

    Trata_FilialCliente = SUCESSO

    Exit Function

Erro_Trata_FilialCliente:

    Trata_FilialCliente = gErr

    Select Case gErr

        Case 193184, 193186, 193188 To 193192, 193194 To 193195

        Case 193185
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA1", gErr, Cliente.Text, objFilialCliente.iCodFilial)

        Case 193187
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case 193193
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr, objTransportadora.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193196)

    End Select

    Exit Function

End Function

Public Sub BotaoApontamento_Click()

Dim lErro As Long
Dim objOSAp As New ClassOSAp

On Error GoTo Erro_BotaoApontamento_Click
    
    'Verifica se a linha do Grid é válida
    If GridAP.Row = 0 Then gError 206142
    
    objOSAp.lCodigo = StrParaLong(GridAP.TextMatrix(GridAP.Row, iGrid_APCodigo_Col))
    
    Call Chama_Tela("OSApontamento", objOSAp)
                    
    Exit Sub

Erro_BotaoApontamento_Click:

    Select Case gErr
        
        Case 206142
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206156)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoOS_Click()

Dim lErro As Long
Dim objOS As New ClassOS

On Error GoTo Erro_BotaoOS_Click
    
    'Verifica se a linha do Grid é válida
    If GridAP.Row = 0 Then gError 206142
    
    objOS.iFilialEmpresa = giFilialEmpresa
    objOS.sCodigo = GridAP.TextMatrix(GridAP.Row, iGrid_APOS_Col)
    
    Call Chama_Tela("OrdemServico", objOS)
                    
    Exit Sub

Erro_BotaoOS_Click:

    Select Case gErr
        
        Case 206142
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206156)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoGarantia_Click()

Dim lErro As Long
Dim objGarantia As New ClassGarantia

On Error GoTo Erro_BotaoGarantia_Click
    
    'Verifica se a linha do Grid é válida
    If GridAP.Row = 0 Then gError 206142
    
    objGarantia.iFilialEmpresa = giFilialEmpresa
    objGarantia.lCodigo = StrParaLong(GridAP.TextMatrix(GridAP.Row, iGrid_APGarantia_Col))
    
    Call Chama_Tela("Garantia", objGarantia)
                    
    Exit Sub

Erro_BotaoGarantia_Click:

    Select Case gErr
        
        Case 206142
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206156)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoSolicitacao_Click()

Dim lErro As Long
Dim objSolicSRV As New ClassSolicSRV

On Error GoTo Erro_BotaoSolicitacao_Click
    
    'Verifica se a linha do Grid é válida
    If GridAP.Row = 0 Then gError 206142
    
    objSolicSRV.iFilialEmpresa = giFilialEmpresa
    objSolicSRV.lCodigo = StrParaLong(GridAP.TextMatrix(GridAP.Row, iGrid_APSolic_Col))
    
    Call Chama_Tela("SolicitacaoSRV", objSolicSRV)
                    
    Exit Sub

Erro_BotaoSolicitacao_Click:

    Select Case gErr
        
        Case 206142
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206156)
    
    End Select
    
    Exit Sub

End Sub

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

On Error GoTo Erro_ValorDescontoItens_Aplica

    If objGridItens.iLinhasExistentes > 0 Or objGridItensSRV.iLinhasExistentes > 0 Then
    
        dTotal = StrParaDbl(ValorProdutos2.Caption)
        dDesconto = StrParaDbl(ValorDescontoItens.Text)
        dFator = dDesconto / dTotal
    
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If Len(Trim(GridItens.TextMatrix(iIndice, iG_PrecoTotB_Col(ITEMNF_TIPO_PECA)))) > 0 Then
                dValorItem = StrParaDbl(GridItens.TextMatrix(iIndice, iG_PrecoTotB_Col(ITEMNF_TIPO_PECA)))
                dDescontoItem = Arredonda_Moeda(dValorItem * dFator)
                dDescontoAplicado = dDescontoAplicado + dDescontoItem
                GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(dDescontoItem, "STANDARD")
            End If
        Next
        For iIndice = 1 To objGridItensSRV.iLinhasExistentes
            If Len(Trim(GridItensSRV.TextMatrix(iIndice, iG_PrecoTotB_Col(ITEMNF_TIPO_SERVICO)))) > 0 Then
                dValorItem = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iG_PrecoTotB_Col(ITEMNF_TIPO_SERVICO)))
                dDescontoItem = Arredonda_Moeda(dValorItem * dFator)
                dDescontoAplicado = dDescontoAplicado + dDescontoItem
                GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col) = Format(dDescontoItem, "STANDARD")
            End If
        Next
        If Abs(dDesconto - dDescontoAplicado) > DELTA_VALORMONETARIO Then
            If objGridItens.iLinhasExistentes > 0 Then
                GridItens.TextMatrix(1, iGrid_Desconto_Col) = Format(StrParaDbl(GridItens.TextMatrix(1, iGrid_Desconto_Col)) + dDescontoAplicado - dDesconto, "STANDARD")
            Else
                GridItensSRV.TextMatrix(1, iGrid_DescontoSRV_Col) = Format(StrParaDbl(GridItensSRV.TextMatrix(1, iGrid_DescontoSRV_Col)) + dDescontoAplicado - dDesconto, "STANDARD")
            End If
        End If
        
        For iIndice = 1 To objGridItens.iLinhasExistentes
            dPercDesc = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col)) / StrParaDbl(GridItens.TextMatrix(iIndice, iG_PrecoTotB_Col(ITEMNF_TIPO_PECA)))
            GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dPercDesc, "Percent")
            Call PrecoTotal_Calcula(ITEMNF_TIPO_PECA, iIndice)
            
            lErro = gobjTribTab.Alteracao_Item_Grid(iIndice)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            dValorTotal = dValorTotal + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        Next
    
        For iIndice = 1 To objGridItensSRV.iLinhasExistentes
            dPercDesc = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col)) / StrParaDbl(GridItensSRV.TextMatrix(iIndice, iG_PrecoTotB_Col(ITEMNF_TIPO_SERVICO)))
            GridItensSRV.TextMatrix(iIndice, iGrid_PercDescSRV_Col) = Format(dPercDesc, "Percent")
            Call PrecoTotal_Calcula(ITEMNF_TIPO_SERVICO, iIndice)
            
            lErro = gobjTribTab.Alteracao_Item_Grid(iIndice, True)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            dValorTotal = dValorTotal + StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
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
        If objGridItens.iLinhasExistentes > 0 Or objGridItensSRV.iLinhasExistentes > 0 Then
            For iIndice = 1 To objGridItens.iLinhasExistentes
                dDesconto = dDesconto + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
            Next
            For iIndice = 1 To objGridItensSRV.iLinhasExistentes
                dDesconto = dDesconto + StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col))
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

Private Sub Class_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
End Sub
