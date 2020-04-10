VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjAnotacao As ClassAnotacoes
Public gobjInfoAdic As ClassInfoAdic

Public gobjInfoUsu As Object
Public giValorDescontoManual As Integer
Public dValorDescontoItensAnt As Double
Public dPercDescontoItensAnt As Double
Public iDescontoAlterado As Integer

Public gdQuantAnterior As Double

Dim m_objUserControl As Object
Dim FrameL(1 To 7) As Frame
 
'Property Variables:
Dim m_Caption As String
Event Unload()

'inicia objeto associado a GridComissoes
Public objTabComissoes As New ClassTabComissoes

'************ GRADE **********************
Public gobjNFiscal As ClassNFiscal
'*****************************************

Public gobjTelaProjetoInfo As ClassTelaPRJInfo

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

'Flags e variáveis globais da tela
Public iAlterado As Integer
Public iFrameAtual As Integer
Public gobjTribTab As ClassTribTab
Public glCodPedido As Long
Public giFilialPedido As Integer
Public giFaturaIntegral As Integer
Public iValorFreteAlterado As Integer
Public iValorSeguroAlterado As Integer
Public iValorDespesasAlterado As Integer
Public iValorDescontoAlterado As Integer
Public iDataEmissaoAlterada As Integer
Public iSerieAlterada As Integer

Public gdDesconto As Double
Public gcolItemPedido As Collection
Public gcolItemQuant As Collection
Public gcolItemUM As Collection 'Inserido por Wagner
Public giFilialEntregaAlterada As Integer
Public giFilialEntregaManual As Integer

'Grids
Public objGridItens As AdmGrid
Public objGridComissoes As AdmGrid
Public objGridAlocacoes As AdmGrid
Public objGrid1 As AdmGrid

'Colunas do grid de Itens
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

'Colunas do Grid de Alocacoes
Public iGrid_Item_Col As Integer
Public iGrid_ProdutoAloc_Col As Integer
Public iGrid_AlmoxAloc_Col As Integer
Public iGrid_UMAloc_Col As Integer
Public iGrid_QuantAloc_Col As Integer
Public iGrid_QuantVend_Col As Integer

'mnemonicos
Private Const CTACONTABILEST1 As String = "ContaContabilEst"
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
Private Const CODIGO1 As String = "Codigo"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque"
Private Const PIS_DEBITO_TOTAL As String = "PIS_Debito_Total"
Private Const COFINS_DEBITO_TOTAL As String = "COFINS_Debito_Total"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const ISS_RETIDO As String = "ISS_Retido"
Private Const COFINS_RETIDO As String = "COFINS_Retido"
Private Const CSLL_RETIDO As String = "CSLL_Retido"
Private Const CONTACONTABILAPLIC As String = "ContaContabilAplic"
Private Const CONTACONTABILCATPROD As String = "ContaContabilCatProd"
Private Const PIS_VALOR_TOTAL As String = "PIS_Valor_Total"
Private Const COFINS_VALOR_TOTAL As String = "COFINS_Valor_Total"
'fim da contabilidade

'Eventos da tela de browse

Public WithEvents objEventoPedido As AdmEvento
Attribute objEventoPedido.VB_VarHelpID = -1
Public WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Public WithEvents objEventoNFiscal As AdmEvento
Attribute objEventoNFiscal.VB_VarHelpID = -1
Public WithEvents objEventoNaturezaOp As AdmEvento
Attribute objEventoNaturezaOp.VB_VarHelpID = -1
Public WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Public WithEvents objEventoEstoqueProd As AdmEvento
Attribute objEventoEstoqueProd.VB_VarHelpID = -1
Public WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Public WithEvents objEventoLocalizacao As AdmEvento
Attribute objEventoLocalizacao.VB_VarHelpID = -1
Public WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoTranspRedespacho As AdmEvento
Attribute objEventoTranspRedespacho.VB_VarHelpID = -1
Public WithEvents objEventoMensagem As AdmEvento
Attribute objEventoMensagem.VB_VarHelpID = -1


Public gobjRastreamento As ClassRastreamento

'declara o objeto da execucao das regras de comissoes e
'o objeto do calculo de comissoes
Public objComissoesRegrasCalc As Object
Public objMnemonicoComissCalc As ClassMnemonicoComissCalc
Public objMnemonicoComissCalcAux As ClassMnemonicoComissCalcAux

'Indica que foi alterado pelo menos um campo na tela que seja utilizado para calcular as comissões
Public iComissoesAlterada As Integer

'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_Itens = 2
Private Const TAB_Complemento = 3
Private Const TAB_Comissoes = 4
Private Const TAB_Almoxarifado = 5
Private Const TAB_Tributacao = 6
Private Const TAB_Contabilizacao = 7
Private Const TAB_RASTRO = 8

Private mvarsName As String
Private DOCINFO_PADRAO_TELA As Integer
Private DOCINFO_PADRAO_TELA_SIGLA As String

'--- inicio dos properties get dos controles da tela
Public Property Get BotaoGrade() As Object
     Set BotaoGrade = objUserControl.Controls("BotaoGrade")
End Property

Public Property Get TabPrecoItemNF() As Object
     Set TabPrecoItemNF = objUserControl.Controls("TabPrecoItemNF")
End Property

Public Property Get ComissaoItemNF() As Object
     Set ComissaoItemNF = objUserControl.Controls("ComissaoItemNF")
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

Public Property Get RedespachoCli() As Object
     Set RedespachoCli = objUserControl.Controls("RedespachoCli")
End Property

Public Property Get TranspRedespacho() As Object
     Set TranspRedespacho = objUserControl.Controls("TranspRedespacho")
End Property

Public Property Get TotalValorBase() As Object
     Set TotalValorBase = objUserControl.Controls("TotalValorBase")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get BotaoVendedores() As Object
     Set BotaoVendedores = objUserControl.Controls("BotaoVendedores")
End Property

Public Property Get ComissaoAutomatica() As Object
     Set ComissaoAutomatica = objUserControl.Controls("ComissaoAutomatica")
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

Public Property Get ValorBaixa() As Object
     Set ValorBaixa = objUserControl.Controls("ValorBaixa")
End Property

Public Property Get PercentualBaixa() As Object
     Set PercentualBaixa = objUserControl.Controls("PercentualBaixa")
End Property

Public Property Get GridComissoes() As Object
     Set GridComissoes = objUserControl.Controls("GridComissoes")
End Property

Public Property Get LabelTotaisComissoes() As Object
     Set LabelTotaisComissoes = objUserControl.Controls("LabelTotaisComissoes")
End Property

Public Property Get TotalValorComissao() As Object
     Set TotalValorComissao = objUserControl.Controls("TotalValorComissao")
End Property

Public Property Get TotalPercentualComissao() As Object
     Set TotalPercentualComissao = objUserControl.Controls("TotalPercentualComissao")
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

Public Property Get TabelaPreco() As Object
     Set TabelaPreco = objUserControl.Controls("TabelaPreco")
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

Public Property Get CTBSeqContraPartida() As Object
     Set CTBSeqContraPartida = objUserControl.Controls("CTBSeqContraPartida")
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

Public Property Get Frame7() As Object
     Set Frame7 = objUserControl.Controls("Frame7")
End Property

Public Property Get UnidadeMedEst() As Object
     Set UnidadeMedEst = objUserControl.Controls("UnidadeMedEst")
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

Public Property Get QuantVendida() As Object
     Set QuantVendida = objUserControl.Controls("QuantVendida")
End Property

Public Property Get GridAlocacao() As Object
     Set GridAlocacao = objUserControl.Controls("GridAlocacao")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
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

Public Property Get Frame6() As Object
     Set Frame6 = objUserControl.Controls("Frame6")
End Property

Public Property Get TipoFrete() As Object
     Set TipoFrete = objUserControl.Controls("TipoFrete")
End Property

Public Property Get TransportadoraLabel() As Object
     Set TransportadoraLabel = objUserControl.Controls("TransportadoraLabel")
End Property

Public Property Get Frame8() As Object
     Set Frame8 = objUserControl.Controls("Frame8")
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

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get Mensagem() As Object
     Set Mensagem = objUserControl.Controls("Mensagem")
End Property

Public Property Get CanalVenda() As Object
     Set CanalVenda = objUserControl.Controls("CanalVenda")
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

Public Property Get BotaoEstoqueProd() As Object
     Set BotaoEstoqueProd = objUserControl.Controls("BotaoEstoqueProd")
End Property

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
End Property

Public Property Get Frame9() As Object
     Set Frame9 = objUserControl.Controls("Frame9")
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

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
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

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get FilialPedido() As Object
     Set FilialPedido = objUserControl.Controls("FilialPedido")
End Property

Public Property Get BotaoVerPedido() As Object
     Set BotaoVerPedido = objUserControl.Controls("BotaoVerPedido")
End Property

Public Property Get BotaoPedido() As Object
     Set BotaoPedido = objUserControl.Controls("BotaoPedido")
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

'Rastreamento
Public Property Get GridRastro() As Object
     Set GridRastro = objUserControl.Controls("GridRastro")
End Property

Public Property Get BotaoConsultaTitRec() As Object
     Set BotaoConsultaTitRec = objUserControl.Controls("BotaoConsultaTitRec")
End Property

Public Property Get DiretoIndireto() As Object
     Set DiretoIndireto = objUserControl.Controls("DiretoIndireto")
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

Public Sub ImprimeRomaneio_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub BotaoLimparNF_Click()

    NFiscal.Caption = ""
    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
End Sub

Public Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub ComissaoAutomatica_Click()

Dim lErro As Long

On Error GoTo Erro_ComissaoAutomatica_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Alterado por Tulio em 02/04
    'Se a comissão automática estiver selecionada
    If ComissaoAutomatica.Value = vbChecked Then
        
        iComissoesAlterada = REGISTRO_ALTERADO
        'Recalcula as comissoes
        lErro = objTabComissoes.Comissoes_Calcula()
        If lErro <> SUCESSO Then gError 51612
        
    End If
    
    Exit Sub

Erro_ComissaoAutomatica_Click:

    Select Case gErr
    
        Case 51612
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157428)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
Dim lErro As Long
    
    Set gobjInfoUsu = Nothing
    
    Set gobjTelaProjetoInfo = Nothing
    Set gobjInfoAdic = Nothing
    
    'Libera as variáveis globais
    Set gobjNFiscal = Nothing
    Set objGridItens = Nothing
    Set objGridComissoes = Nothing
    Set objGridAlocacoes = Nothing
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

    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    Set gobjRastreamento = Nothing

    If Not (gobjTribTab Is Nothing) Then Call gobjTribTab.Desativar
    
    'Alterado por Tulio em 02/04
    'se foi instanciada a classe de execucao de regras de comissoes => libera
    Set objComissoesRegrasCalc = Nothing
        
    'se foi instanciada a classe que calcula os mnemonicos => libera
    Set objMnemonicoComissCalc = Nothing
    Set objMnemonicoComissCalcAux = Nothing

    
    'Fecha o Comando de Setas
    lErro = ComandoSeta_Liberar(Me.Name)

    Set gobjAnotacao = Nothing
    
    Set objTabComissoes = Nothing
    Set gobjTribTab = Nothing

    Set gcolItemPedido = Nothing
    Set gcolItemQuant = Nothing
    Set gcolItemUM = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoTranspRedespacho = Nothing
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157429)
    
    End Select
    
    Exit Sub

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

Public Sub DataSaida_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataSaida, iAlterado)
End Sub

Public Sub TipoFrete_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NFiscal_GotFocus()
    Call MaskEdBox_TrataGotFocus(NFiscal, iAlterado)
End Sub

Public Sub Pedido_GotFocus()
    Call MaskEdBox_TrataGotFocus(Pedido, iAlterado)
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

Public Sub BotaoLimpar_Click()

Dim objNFiscal As New ClassNFiscal
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 25959

    lErro = Limpa_Tela_NFiscalPedido()
    If lErro <> SUCESSO Then gError 500151
    
    'Limpa as coleções globais
    Set gcolItemPedido = New Collection
    Set gcolItemQuant = New Collection
    Set gcolItemUM = New Collection
    
    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        'Prepara para chaveamento para NotaFiscalFatura
'        objNFiscal.iTipoNFiscal = DOCINFO_NFISFV
'
'        'Faz chaveamento de classe CT para NotaFiscalFatura
'        lErro = objUserControl.CT_Chavea(objNFiscal)
'        If lErro <> SUCESSO Then gError 25957
'
'    End If
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 25957, 25959, 500151 'tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157430)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoVerPedido_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoVerPedido_Click

    'Verifica se o Pedido e a filial do pedido foram informados
    If Len(Trim(Pedido.Text)) = 0 Then gError 46571
    If Len(Trim(FilialPedido.Text)) = 0 Then gError 46572
    
    objPedidoVenda.lCodigo = CLng(Pedido.Text)
    objPedidoVenda.iFilialEmpresa = Codigo_Extrai(FilialPedido.Text)
    
    'Verifica se o Pedido está baixado
    lErro = CF("PedidoVendaBaixado_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 46135 Then gError 46573
    If lErro = SUCESSO Then gError 46574 'Se estiver , erro.
    
    'Busca o Pedido na tabela de pedidos de venda
    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 46575
    If lErro <> SUCESSO Then gError 46576 'Não encontrou, erro.
    
    'Chama  a tela de Pedido de Venda passando o pedido de venda da tela
    Call Chama_Tela("PedidoVenda_Consulta", objPedidoVenda)
    
    Exit Sub

Erro_BotaoVerPedido_Click:

    Select Case gErr
    
        Case 46571
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)
            
        Case 46572
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)
        
        Case 46573, 46575
        
        Case 46574
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_BAIXADO", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
        
        Case 46576
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157431)

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

Public Sub DataEmissao_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO
    
    iAlterado = REGISTRO_ALTERADO
    iDataEmissaoAlterada = REGISTRO_ALTERADO

End Sub

Public Sub DataSaida_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialEntrega_Change()
    iAlterado = REGISTRO_ALTERADO
    giFilialEntregaAlterada = 1
End Sub

Public Sub FilialEntrega_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialPedido_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialPedido_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Form_Activate()

Dim lErro As Long

On Error GoTo Erro_Form_Activate

    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError 59337
        
    Call TelaIndice_Preenche(Me)

    Exit Sub
     
Erro_Form_Activate:

    Select Case gErr
          
        Case 59337
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157432)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Public Function Form_Load_Light()

Dim lErro As Long

On Error GoTo Erro_Form_Load_Light

    'obs.: a carga pos form_load será feita pela classe que trata nffat
    giPosCargaOk = 1
    
    'carregar a combo FilialPedido com as Filiais da empresa
    lErro = Carrega_FilialPedido()
    If lErro <> SUCESSO Then gError 25945
    
    'Carrega variáveis globais exclusivas dessa tela
    Set gcolItemPedido = New Collection
    Set gcolItemQuant = New Collection
    Set gcolItemUM = New Collection
    
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
    Set objEventoTranspRedespacho = New AdmEvento
    Set objEventoLote = New AdmEvento
    Set objEventoDoc = New AdmEvento
    
'    'Inicializa FrameL para versão LIGHT
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        Set FrameL(1) = Frame1(1)
'        Set FrameL(2) = Frame1(2)
'        Set FrameL(3) = Frame1(3)
'        Set FrameL(4) = Frame1(5)
'        Set FrameL(5) = Frame1(6)
'        Set FrameL(6) = Frame1(7)
'        Set FrameL(7) = Frame1(8)
'
'    End If
    
    Call GridItens_InicializaColunas
    
    'Colunas do Grid Alocacoes
    iGrid_Item_Col = 1
    iGrid_ProdutoAloc_Col = 2
    iGrid_AlmoxAloc_Col = 3
    iGrid_UMAloc_Col = 4
    iGrid_QuantAloc_Col = 5
    iGrid_QuantVend_Col = 6

    iAlterado = 0
    iComissoesAlterada = 0
    
    Form_Load_Light = SUCESSO
    
    Exit Function

Erro_Form_Load_Light:

    Form_Load_Light = gErr

    Select Case gErr

        Case 25945

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157434)

    End Select
    
    iAlterado = 0
    iComissoesAlterada = 0
    
    Exit Function

End Function

Public Function CargaPosFormLoad(Optional bTrazendoDoc As Boolean = False) As Long

Dim lErro As Long

On Error GoTo Erro_CargaPosFormLoad

    If (giPosCargaOk = 0) Then

        'p/permitir o redesenho da tela
        DoEvents
        
        gbCarregandoTela = True

        'Inicializa a parte de tributacao
        Call GridItens_InicializaColunas
        Set gobjTribTab = New ClassTribTab
        lErro = gobjTribTab.Ativar(Me)
        If lErro <> SUCESSO Then gError 46553
    
        Call gobjTribTab.TipoNFiscal_Definir(DOCINFO_PADRAO_TELA, DOCINFO_PADRAO_TELA_SIGLA) 'DOCINFO_NFISVPV, "NFISVPV")
        If IsDate(DataEmissao.Text) Then Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
        Call gobjTribTab.Serie_Alterada(Serie.Text)
    
        Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
        Set gobjTelaProjetoInfo.objUserControl = objUserControl
        Set gobjTelaProjetoInfo.objTela = Me
    
        'carregar a combo FilialPedido com as Filiais da empresa
        lErro = Carrega_FilialPedido()
        If lErro <> SUCESSO Then gError 46543
    
        'Carrega a combo de series
        lErro = Carrega_Serie()
        If lErro <> SUCESSO Then gError 46544
            
        'Seleciona série padrão
        lErro = Seleciona_Serie_Padrao()
        If lErro <> SUCESSO Then gError 500148
        
        'Carrega a combo de tabelas de preços
        lErro = Carrega_TabelaPreco()
        If lErro <> SUCESSO Then gError 46545
    
        'Carrega a combo de transportadoras
        lErro = Carrega_Transportadora()
        If lErro <> SUCESSO Then gError 46546
    
        'Incluído por Luiz Nogueira em 21/08/03
        'Carrega a combo VolumeEspecie
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
        If lErro <> SUCESSO Then gError 102418
    
        'Incluído por Luiz Nogueira em 21/08/03
        'Carrega a combo VolumeMarca
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
        If lErro <> SUCESSO Then gError 102419
        
        'Carrega a combo PlacaUF
        lErro = Carrega_PlacaUF()
        If lErro <> SUCESSO Then gError 46547
    
        'ACrrega a combo CanalVenda
        lErro = Carrega_CanalVenda()
        If lErro <> SUCESSO Then gError 46548
    
        'Inicializa as variáveis de Grid e de Browse
        Set objEventoMensagem = New AdmEvento
        Set objGridItens = New AdmGrid
        Set objGridComissoes = New AdmGrid
        Set objGridAlocacoes = New AdmGrid
        Set objGrid1 = New AdmGrid
        Set objEventoPedido = New AdmEvento
        Set objEventoSerie = New AdmEvento
        Set objEventoNFiscal = New AdmEvento
        Set objEventoNaturezaOp = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoEstoqueProd = New AdmEvento
        Set objEventoVendedor = New AdmEvento
        Set objEventoLocalizacao = New AdmEvento
        Set objEventoLote = New AdmEvento
        Set objEventoDoc = New AdmEvento
    
        Set gcolItemPedido = New Collection
        Set gcolItemQuant = New Collection
        Set gcolItemUM = New Collection
        
        'Inicializa a máscara do produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
        If lErro <> SUCESSO Then gError 46552
    
        'Inicializa o grid de itens
        lErro = Inicializa_Grid_Itens(objGridItens)
        If lErro <> SUCESSO Then gError 46549
        
        'Inicializa o grid de comissões
        lErro = objTabComissoes.Inicializa_Grid_Comissoes(objGridComissoes)
        If lErro <> SUCESSO Then gError 46550
        'Inicializa o grid de Alocações
        lErro = Inicializa_Grid_Alocacoes(objGridAlocacoes)
        If lErro <> SUCESSO Then gError 46551
        
        PrecoUnitario.Format = gobjFAT.sFormatoPrecoUnitario
        
        QuantAlocada.Format = FORMATO_ESTOQUE
        Quantidade.Format = FORMATO_ESTOQUE
        QuantVendida.Format = FORMATO_ESTOQUE
       
        'Inicialização da parte de contabilidade
        lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_FATURAMENTO)
        If lErro <> SUCESSO Then gError 46554

        'Rastreamento
        Set gobjRastreamento = New ClassRastreamento
        Set gobjRastreamento.objTela = Me
        gobjRastreamento.bSaidaMaterial = True
        gobjRastreamento.bEditavel = False
        
        'Inicializa o grid de Rastreamento
        lErro = gobjRastreamento.Inicializa_Grid_Rastreamento()
        If lErro <> SUCESSO Then gError 83375
        
        gobjRastreamento.iCodigo = DOCINFO_PADRAO_TELA

        'Alterado por Tulio em 02/04

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
          
        Case 46546 To 46554, 83375, 500148, 102418, 102419

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157435)
     
    End Select
     
    Exit Function

End Function

Public Function Seleciona_Serie_Padrao() As Long

Dim lErro As Long
Dim sSeriePadrao As String

On Error GoTo Erro_Seleciona_Serie_Padrao
        
    'Lê série Padrão
    lErro = CF("Serie_Le_Padrao", sSeriePadrao)
    If lErro <> SUCESSO Then gError 500147
    
    'Se possui série padrão
    If Len(Trim(sSeriePadrao)) > 0 Then
    
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
    
        Case 500147
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157436)
    
    End Select
    
    Exit Function
    
End Function

Public Sub Form_Load()

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    Set objTabComissoes.objTela = Me
    
    Call CF2(Me, "NFP_Form_Load2")
    
    If giTipoVersao = VERSAO_LIGHT Then
        
        BotaoGrade.Visible = False
        
        'sumir com abas de contabilidade e rastro
        TabStrip1.Tabs.Remove (TAB_RASTRO)
        TabStrip1.Tabs.Remove (TAB_Contabilizacao)
        
    End If
    
    MsgAutomatica.Value = vbChecked
    Mensagem.MaxLength = MENSAGEM_NF_MAXLENGTH
    MensagemCorpo.MaxLength = MENSAGEM_NF_MAXLENGTH
    
    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207439
    
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    objFilialEmpresa.lCodEmpresa = glEmpresa
    'Lê a FilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 207439

    If objFilialEmpresa.iTipoTribAtividade = 3 Then
        DOCINFO_PADRAO_TELA = DOCINFO_NFISSPV
        DOCINFO_PADRAO_TELA_SIGLA = "NFISSPV"
    Else
        DOCINFO_PADRAO_TELA = DOCINFO_NFISVPV
        DOCINFO_PADRAO_TELA_SIGLA = "NFISVPV"
    End If
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 207439

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157437)

    End Select
    
    iAlterado = 0
    iComissoesAlterada = 0
    
    Exit Sub

End Sub

Public Function Carrega_FilialPedido() As Long

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157438)

    End Select

    Exit Function

End Function

Public Function Carrega_Serie() As Long
'Carrega as combos de Série e serie de NF original com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie
Dim sSerieAnt As String

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 46640

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

        Case 46640

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157439)

    End Select

    Exit Function

End Function

Public Function Carrega_PlacaUF() As Long
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

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157440)

    End Select

End Function

Public Function Carrega_TabelaPreco() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TabelaPreco

    TabPrecoItemNF.AddItem ""
    TabPrecoItemNF.ItemData(TabPrecoItemNF.NewIndex) = 0
    
    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePrecoVenda", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao, "Ativo = 1")
    If lErro <> SUCESSO Then gError 46642

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

        Case 46642

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157441)

    End Select

    Exit Function

End Function

Public Function Carrega_Transportadora() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_Transportadora

    'Lê o Código e o NOme de Todas as Transportadoras do BD
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then gError 46643

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

        Case 46643

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157442)

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
    If lErro <> SUCESSO Then gError 46644

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

        Case 46644

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157443)

    End Select

    Exit Function

End Function

Private Sub GridItens_InicializaColunas()
    Call CF2(Me, "NFP_GridItens_InicializaColunas")
End Sub

Public Function NFP_GridItens_InicializaColunas(ByVal objCT As Object) As Long

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

    NFP_GridItens_InicializaColunas = SUCESSO
    
End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
    Call CF2(Me, "NFP_Inicializa_Grid_Itens", objGridInt)
End Function

Public Function NFP_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
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

    'Largura automática para o grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NFP_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Function Inicializa_Grid_Alocacoes(objGridInt As AdmGrid) As Long
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
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iProibidoExcluir = 1

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Alocacoes = SUCESSO

    Exit Function

End Function

Public Sub Mensagem_Change()
    iAlterado = REGISTRO_ALTERADO
    MsgAutomatica.Value = vbUnchecked
End Sub

Public Sub NFiscal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumPedidoTerc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Pedido_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PesoBruto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PesoLiquido_Change()
    iAlterado = REGISTRO_ALTERADO
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

Public Sub Serie_Click()
    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = REGISTRO_ALTERADO
End Sub

Public Sub Serie_Change()
    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = REGISTRO_ALTERADO
End Sub

Public Sub TabelaPreco_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Transportadora_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Transportadora_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub


Public Sub ValorDesconto_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    iValorDescontoAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorDespesas_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    iValorDespesasAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorProdutos_Change()

Dim dValorDesconto As Double
Dim dValorProdutos As Double
    
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
'''*** incluidos p/tratamento de tributacao *******************************
'            Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, dValorDesconto)
''*** fim tributacao
'
'        End If
'
'    End If

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
    If lErro <> SUCESSO Then gError 102420
    
    Exit Sub

Erro_VolumeEspecie_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102420
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157444)

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
    If lErro <> SUCESSO Then gError 102421
    
    Exit Sub

Erro_VolumeMarca_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102421
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157445)

    End Select

End Sub

Public Sub VolumeNumero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub VolumeQuant_Change()
    iAlterado = REGISTRO_ALTERADO
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
            If lErro <> SUCESSO And lErro <> 12253 Then gError 46647
            'Se conseguir --> Sai
            If lErro <> SUCESSO Then
            
                objSerie.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
                'Busca a série no BD
                lErro = CF("Serie_Le", objSerie)
                If lErro <> SUCESSO And lErro <> 22202 Then gError 51618
                If lErro <> SUCESSO Then gError 46646 'Se não encontrou
            
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

        Case 46646
            'Pergunta se deseja criar nova série
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_SERIE", Serie.Text)
            'Se a resposta for afirmativa
            If vbMsgRes = vbYes Then
                'Chama a tela de cadastro de séries
                Call Chama_Tela("SerieNFiscal", objSerie)
            End If
            'segura o foco na série

        Case 46647, 51618

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157446)

    End Select

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

Private Sub Class_Initialize()
'setar
    Me.sName = "NFiscalPedido"
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
    Call Chama_Tela("NFiscalPedidoLista", colSelecao, objNFiscal, objEventoNFiscal)

End Sub

Private Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objNFiscal = obj1

    'Evita ter que alterar os browses a cada campo inserido na tabela de NF
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 180104

    'Coloca na Tela a Nota Fiscal escolhida
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 35715

    glCodPedido = objNFiscal.lNumPedidoVenda
    giFilialPedido = objNFiscal.iFilialPedido

    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case gErr

        Case 35715, 180104

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157447)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    If iDataEmissaoAlterada = 1 Then

        'Verifica se a data de emissao foi digitada
        If Len(Trim(DataEmissao.ClipText)) <> 0 Then
    
            'Critica a data digitada
            lErro = Data_Critica(DataEmissao.Text)
            If lErro <> SUCESSO Then gError 46648
        
        End If
        
''*** incluidos p/tratamento de tributacao *******************************
        Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
'*** fim tributacao
        
        iDataEmissaoAlterada = 0

    End If

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case gErr

        Case 46648

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157448)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 46649

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 46649

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157449)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 46650

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 46650

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157450)

    End Select

    Exit Sub

End Sub

Public Sub DataSaida_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataSaida_Validate

    'Verifica se a data de Saida foi digitada
    If Len(Trim(DataSaida.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataSaida.Text)
    If lErro <> SUCESSO Then gError 46651

    Exit Sub

Erro_DataSaida_Validate:

    Cancel = True


    Select Case gErr
        Case 46651

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157451)

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
    If lErro <> SUCESSO Then gError 89800

    Exit Sub

Erro_HoraSaida_Validate:

    Cancel = True

    Select Case gErr

        Case 89800

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157452)

    End Select

    Exit Sub

End Sub

Public Sub UpDownSaida_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownSaida_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataSaida, DIMINUI_DATA)
    If lErro Then gError 46652

    Exit Sub

Erro_UpDownSaida_DownClick:

    Select Case gErr

        Case 46652

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157453)

    End Select

    Exit Sub

End Sub

Public Sub UpDownSaida_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownSaida_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataSaida, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 46653

    Exit Sub

Erro_UpDownSaida_UpClick:

    Select Case gErr

        Case 46653

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157454)

    End Select

    Exit Sub

End Sub

Public Sub TabelaPreco_Click()

Dim lErro As Long

On Error GoTo Erro_TabelaPreco_Click

    iAlterado = REGISTRO_ALTERADO

    'Verifica se a Tabela de Preços foi informada
    If TabelaPreco.ListIndex = -1 Then Exit Sub

    'Verifica se existe algum item no Grid
    If objGridItens.iLinhasExistentes = 0 Then Exit Sub

    'Faz o tratamento para a Tabela Preco selecionada
    lErro = Trata_TabelaPreco()
    If lErro <> SUCESSO Then gError 46654

    Exit Sub

Erro_TabelaPreco_Click:

    Select Case gErr

        Case 46654

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157455)

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

    'Verifica se está preenchida com o item selecionado na ComboBox TabelaPreco
    If TabelaPreco.Text = TabelaPreco.List(TabelaPreco.ListIndex) Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TabelaPreco, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 46655

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 46656

        If lErro <> SUCESSO Then gError 46658 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 46657

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 46659

    Exit Sub

Erro_TabelaPreco_Validate:

    Cancel = True


    Select Case gErr

        Case 46655, 46656, 46657

        Case 46658  'Não encontrou Tabela de Preço no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")

            If vbMsgRes = vbYes Then
                'Chama a tela de Tabelas de Preço
                Call Chama_Tela("TabelaPrecoItem", objTabelaPreco)
            End If

            'Segura o foco

        Case 46659
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157456)

    End Select

    Exit Sub

End Sub

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

Public Sub ValorFrete_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    iValorFreteAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorSeguro_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    iValorSeguroAlterado = REGISTRO_ALTERADO

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
        If lErro <> SUCESSO Then gError 46664

        dValorFrete = CDbl(ValorFrete.Text)
        'Coloca o Valor Formatado na tela
        ValorFrete.Text = Format(dValorFrete, "Fixed")

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

        Case 46664

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157458)

    End Select

    Exit Sub

End Sub

Public Sub ValorSeguro_Validate(Cancel As Boolean)

Dim dValorSeguro As Double
Dim lErro As Long

On Error GoTo Erro_Valorseguro_Validate

    If iValorSeguroAlterado = 0 Then Exit Sub

    'Se estiver preenchido
    If Len(Trim(ValorSeguro.Text)) > 0 Then
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(ValorSeguro.Text)
        If lErro <> SUCESSO Then gError 46665

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

        Case 46665

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157459)

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
        If lErro <> SUCESSO Then gError 46666

        dValorDespesas = CDbl(ValorDespesas.Text)

        'coloca o valor formatado na tela
        ValorDespesas.Text = Format(dValorDespesas, "Fixed")

    End If

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDespesas_Validate(Cancel, dValorDespesas)
'*** fim tributacao

    Call ValorTotal_Calcula

    iValorDespesasAlterado = 0

    Exit Sub

Erro_ValorDespesas_Validate:

    Cancel = True

    Select Case gErr

        Case 46666

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157460)

    End Select

    Exit Sub

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
        If lErro <> SUCESSO Then gError 46667

        dValorDesconto = CDbl(ValorDesconto.Text)
        
        'Coloca o Valor formatado na tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")
        
        lErro = DescontoGlobal_Testa(dValorDesconto)
        If lErro <> SUCESSO Then gError 56893

        dValorProdutos = dValorProdutos - dValorDesconto

    End If

    ValorProdutos.Caption = Format(dValorProdutos, "Standard")

    lErro = CF2(Me, "NFCF_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132038

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDesconto_Validate(Cancel, dValorDesconto)
'*** fim tributacao

    Call ValorTotal_Calcula

    iValorDescontoAlterado = 0

    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True

    Select Case gErr

        Case 46667, 56893, 132038

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157461)

    End Select

    Exit Sub

End Sub

Public Sub FilialEntrega_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialEntrega_Validate

    'Verifica se a FilialEntrega foi preenchida
    If Len(Trim(FilialEntrega.Text)) = 0 Or giFilialEntregaAlterada = 0 Then Exit Sub

    'Verifica se é uma FilialEntrega selecionada
    If FilialEntrega.ListIndex <> -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialEntrega, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 46668

    'Se nao encontra o item com o código informado
    If lErro = 6730 Then

        'Verifica de o Cliente foi digitado
        If Len(Trim(Cliente.Caption)) = 0 Then gError 46671

        sCliente = Cliente.Caption

        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe FilialEntrega com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 46669

        If lErro = 17660 Then gError 46670

        'coloca na tela
        FilialEntrega.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 46672

    giFilialEntregaAlterada = 0
    
    Exit Sub

Erro_FilialEntrega_Validate:

    Cancel = True

    Select Case gErr

        Case 46668, 46669

        Case 46670
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Caption)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If
            
        Case 46671
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 46672
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialEntrega.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157462)

    End Select

    Exit Sub

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 46673
    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 46674

        If lErro = 19250 Then gError 46675 'Não encontrou Transportadora no BD

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    End If

    'Não existe o item com a String na List da Combo
    If lErro = 6731 Then gError 46676

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True


    Select Case gErr

        Case 46673, 46674

        Case 46675
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then
                Call Chama_Tela("Transportadora", objTransportadora)
            End If
            
        Case 46676
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157463)

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

Public Sub PlacaUF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PlacaUF_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(PlacaUF.Text)) = 0 Then Exit Sub

    'Verifica se existe o item na combo
    lErro = Combo_Item_Igual(PlacaUF)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 46677

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 46678

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True

    Select Case gErr

        Case 46677

        Case 46678
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157464)

    End Select

    Exit Sub

End Sub

Public Sub PesoBruto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoBruto_Validate

    'Verifica se está preenchido
    If Len(Trim(PesoBruto.ClipText)) = 0 Then Exit Sub

    'Criticao valor
    lErro = Valor_Positivo_Critica(PesoBruto.Text)
    If lErro <> SUCESSO Then gError 46679

    'Coloca o valor formatado na Tela
    PesoBruto.Text = Format(PesoBruto.Text, PesoBruto.Format)
    
    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True

    Select Case gErr

        Case 46679

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157465)

    End Select

    Exit Sub

End Sub

Public Sub PesoLiquido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoLiquido_Validate

    'Verifica se foi preenchido
    If Len(Trim(PesoLiquido.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(PesoLiquido.Text)
    If lErro <> SUCESSO Then gError 46680

    'Coloca o valor formatado na Tela
    PesoLiquido.Text = Format(PesoLiquido.Text, PesoLiquido.Format) 'Alterado por Wagner

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True


    Select Case gErr

        Case 46680

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157466)

    End Select

    Exit Sub

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 46681

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCanalVenda.iCodigo = iCodigo

        'Tenta ler CanalVenda com esse código no BD
        lErro = CF("CanalVenda_Le", objCanalVenda)
        If lErro <> SUCESSO And lErro <> 23597 Then gError 46682

        If lErro = 23597 Then gError 46683 'Não encontrou CanalVenda no BD

        'Encontrou CanalVenda no BD, coloca no Text da Combo
        CanalVenda.Text = CStr(objCanalVenda.iCodigo) & SEPARADOR & objCanalVenda.sNome

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 46684

    Exit Sub

Erro_CanalVenda_Validate:

    Cancel = True


    Select Case gErr

        Case 46681, 46682

        Case 46683  'Não encontrou CanalVenda no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CANALVENDA", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CanalDeVenda", objCanalVenda)
            End If
            'Segura o foco

        Case 46684
            Call Rotina_Erro(vbOKOnly, "ERRO_CANALVENDA_NAO_ENCONTRADO", gErr, CanalVenda.Text)

        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157467)

    End Select

    Exit Sub

End Sub

Public Sub BotaoVendedores_Click()

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoVendedores_Click

    'Chama a tela de browse de vendedores
    lErro = objTabComissoes.BotaoVendedores_Click()
    If lErro <> SUCESSO Then gError 43703

    Exit Sub
    
Erro_BotaoVendedores_Click:

    Select Case gErr
            
        Case 43703
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157468)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim lErro As Long

On Error GoTo Erro_objEventoVendedor_evSelecao
    
    Set objVendedor = obj1
    
    If GridComissoes.Row > 0 Then
    
        'se empresa nao utiliza regras p/ o calc de comissoes
        If Not gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
        
            'Preenche a linha de Vendedor com dados default
            lErro = objTabComissoes.VendedorLinha_Preenche(objVendedor)
            If lErro <> SUCESSO Then gError 43732
        
        'senao
        Else
            'verificar se o vendedor esta no grid e tratar o campo 'DiretoIndireto'
            lErro = objTabComissoes.VendedorLinha_Preenche1(objVendedor)
            If lErro <> SUCESSO Then gError 101519
    
        End If
    
    End If

    Me.Show
    
    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr

        Case 43732, 101519 'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157469)

    End Select

    Exit Sub

End Sub

Public Sub BotaoEstoqueProd_Click()

Dim lErro As Long
Dim objEstoqueProd As New ClassEstoqueProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim iAlmoxarifado As Integer
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto, sSelecao As String, sProdutoEnxuto As String

On Error GoTo Erro_BotaoEstoqueProd_Click

    'Verifica se a linha do Grid é válida
    If GridItens.Row = 0 Then gError 43730
        
    'Verifica se o Produto na linha do Grid está preenchido
    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    
    If iProdutoPreenchido <> PRODUTO_VAZIO Then
        
        objProduto.sCodigo = sProdutoFormatado
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 101782
        If lErro <> SUCESSO Then gError 101783
        
        objEstoqueProd.sProduto = sProdutoFormatado
    
        If Len(Trim(objProduto.sGrade)) = 0 Then
        
            colSelecao.Add sProdutoFormatado
    
            'Lê o Almoxarifado Padrão do Produto
            lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, sProdutoFormatado, iAlmoxarifado)
            If lErro = SUCESSO Then objEstoqueProd.iAlmoxarifado = iAlmoxarifado
    
            'Chama a tela EstoqueProdutoFilialLista
            Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProd, objEventoEstoqueProd)
        
        Else
        
            Call Mascara_RetornaProdutoEnxuto(sProdutoFormatado, sProdutoEnxuto)
            sSelecao = "Produto LIKE '" & sProdutoEnxuto & "%' AND FilialEmpresa = " & giFilialEmpresa
            Call Chama_Tela("EstProdLista_Consulta", colSelecao, objEstoqueProd, objEventoEstoqueProd, sSelecao)
        
        End If
    
    Else
        gError 43731
    End If

    Exit Sub
    
Erro_BotaoEstoqueProd_Click:

    Select Case gErr
    
        Case 101782
        
        Case 43730
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 43731
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 101783
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157470)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoEstoqueProd_evSelecao(obj1 As Object)
    Me.Show
End Sub

Private Sub Seleciona_FilialEmpresa()

Dim iIndice As Integer

    If giFilialEmpresa <> EMPRESA_TODA Then
        'Seleciona a filial de pedido empresa = giFilialEmpresa
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

Private Function PrecoUnitario_Calcula(sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long

Dim objProduto As New ClassProduto
Dim objUM As New ClassUnidadeDeMedida
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim lErro As Long

On Error GoTo Erro_PrecoUnitario_Calcula

    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto

    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 46686
    If lErro = 28030 Then gError 46687 'Não encontrou

    'Faz a conversão da UM da tela para a UM de Venda
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 46685

    'Calcula o preço unitário
    dPrecoUnitario = objTabelaPrecoItem.dPreco / dFator

    PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_PrecoUnitario_Calcula:

    PrecoUnitario_Calcula = gErr

    Select Case gErr

        Case 46685, 46686

        Case 46687
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objTabelaPrecoItem.sCodProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157471)

    End Select

    Exit Function

End Function

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
        dPrecoTotal = dPrecoUnitario * dQuantidade
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
    If lErro <> SUCESSO Then gError 184288
    
    Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)

'    'Obtem o valor do Desconto global
'    If gdDesconto > 0 And iValorDescontoAlterado = 0 Then
'        dValorDescontoGlobal = gdDesconto * dValorProdutos
'    ElseIf Len(Trim(ValorDesconto.Text)) > 0 And IsNumeric(ValorDesconto.Text) Then
'        dValorDescontoGlobal = CDbl(ValorDesconto.Text)
'    End If
'
'    dValorProdutos = dValorProdutos - dValorDescontoGlobal
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
'        If lErro <> SUCESSO Then gError 56885
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

    lErro = CF2(Me, "NFCF_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132039
    
    If iCalculaTotal <> False Then Call ValorTotal_Calcula

    Exit Sub

Erro_PrecoTotal_Calcula:

    Select Case gErr
    
        Case 56885, 132039, 177050, 184288
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157472)
            
    End Select
            
    Exit Sub

End Sub

Public Function ValorTotal_Calcula() As Long

Dim dValorTotal As Double
Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    lErro = gobjTribTab.ValorTotal_Calcula(dValorTotal)
    If lErro <> SUCESSO Then gError 46689
    
    ValorTotal.Caption = Format(dValorTotal, "Standard")

    Call ValorDescontoItens_Calcula
 
    If Not gbCarregandoTela Then

        'Esse CT também é instanciado por NFiscalRemPedido, mas as comissões só devem
        'ser calculadas quando a tela que o instanciou for NFiscalPedido
        If Me.sName = NOME_TELA_NFISCALPEDIDO Then
            
            'Se a empresa nãu utiliza as regras para o cálculo de comissões
            If Not gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
            
                'Significa que é necessário recalcular as comissões, utilizando o cálculo padrão
                lErro = objTabComissoes.Comissoes_Calcula_Padrao()
                If lErro <> SUCESSO Then gError 46688
            
            Else
                iComissoesAlterada = 1 'preventivamente, pois pode ter alterado algo que afete uma regra
            End If
        
        End If

    End If
    
    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 46688, 46689 ', 56886

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157473)

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
    
    dPesoBruto = 0
    dPesoLiquido = 0
    
    'Para cada linha do Grid de Itens
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        'Se o produto e a UM estiverem preenchidos
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) > 0 Then
            
            'Formata o produto
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 500103
            
            'Lê o produto
            objProduto.sCodigo = sProdutoFormatado
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 500104

            'Se o produto não está cadastrado, erro
            If lErro = 28030 Then gError 500106
                        
            'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then gError 500105
                                
            dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
            
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
    
        Case 500103, 500104, 500105
        
        Case 500106
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157474)
    
    End Select
    
    Exit Function
    
End Function

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

    dValorReal = dValorTotal - dDesconto

End Sub

Public Sub SubTotal_Calcula(objGridInt As AdmGrid, iGrid_Coluna As Integer, dSubTotal As Double) 'alterado p/ public por Leo em 17/01/02
'Faz a soma da Coluna passado no Grid passado e devolve em dValorTotal

Dim iIndice As Integer
Dim objItemNF As ClassItemNF

    dSubTotal = 0

    For iIndice = 1 To objGridInt.iLinhasExistentes
        Set objItemNF = gobjNFiscal.colItens.Item(iIndice)
        
        If objItemNF.objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
            'Acumula em dSubTotal
            dSubTotal = dSubTotal + CDbl(IIf(Len(Trim(objGridInt.objGrid.TextMatrix(iIndice, iGrid_Coluna))) > 0, objGridInt.objGrid.TextMatrix(iIndice, iGrid_Coluna), 0))
        End If
    Next

    Exit Sub

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 46690
    If lErro <> SUCESSO Then gError 46691 'Não encontrou.
    
    Exit Sub

Erro_FilialPedido_Validate:

    Cancel = True


    Select Case gErr

        Case 46690

        Case 46691
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialPedido.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157475)

    End Select

    Exit Sub

End Sub

Public Sub BotaoPedido_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objBloqueioPV As ClassBloqueioPV

On Error GoTo Erro_BotaoPedido_Click

    'Verfifica se o Pedido e a FilialPedido foram informadas
    If Len(Trim(Pedido.Text)) = 0 Then gError 46555
    If Len(Trim(FilialPedido.Text)) = 0 Then gError 46556

    objPedidoVenda.lCodigo = CLng(Pedido.Text)
    objPedidoVenda.iFilialEmpresa = Codigo_Extrai(FilialPedido.Text)
    
    'Verifica se o pedido está baixado
    lErro = CF("PedidoVendaBaixado_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 46135 Then gError 46557
    
    If lErro = SUCESSO Then gError 46558

    'Busca o pedido na tabela de Pedidos de Venda
    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 46559
    
    If lErro <> SUCESSO Then gError 46560 'Não encontrou
    
    lErro = CF("BloqueiosPV_Le", objPedidoVenda)
    If lErro <> SUCESSO Then gError 60773
    
    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV
        If objBloqueioPV.dtDataLib = DATA_NULA Then
            If objBloqueioPV.iTipoDeBloqueio <> BLOQUEIO_PARCIAL And objBloqueioPV.iTipoDeBloqueio <> BLOQUEIO_NAO_RESERVA Then gError 60774
        End If
    Next
    
    'Verifica se a Filial de Faturamento do Pedido coincide com a filial de trabalho
    If objPedidoVenda.iFilialEmpresaFaturamento <> giFilialEmpresa Then gError 46570
    
    lErro = CF("Verifica_TipoPagto_Preenchido", objPedidoVenda)
    If lErro <> SUCESSO Then gError 178999
    
    'Traz os dados do pedido para a tela
    lErro = Traz_Dados_Pedido_Tela(objPedidoVenda)
    If lErro <> SUCESSO Then gError 46561
    
    Exit Sub

Erro_BotaoPedido_Click:

    Select Case gErr

        Case 46555
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)

        Case 46556
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)

        Case 46557, 46559, 46561, 60773, 178999

        Case 46558
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_BAIXADO", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)

        Case 46560
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)

        Case 46570
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFATURAMENTO_DIFERENTE", gErr, objPedidoVenda.iFilialEmpresaFaturamento, objPedidoVenda.lCodigo, giFilialEmpresa)

        Case 60774
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_BLOQUEIO", gErr, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157476)

    End Select

    Exit Sub

End Sub

Public Function Traz_Dados_Pedido_Tela(objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objCliente As New ClassCliente, bFaturadoParc As Boolean
Dim objItemNF As ClassItemNF
Dim objDataEntrega As ClassDataEntrega
Dim iIndice As Integer, bEmissaoForcaRecalculo As Boolean

On Error GoTo Erro_Traz_Dados_Pedido_Tela

    'Indica que o sistema está carregando dados para a tela,
    'e, portanto, algums cálculos e procedimentos não serão executados
    gbCarregandoTela = True
    
    'Lê todos oa dados do pedido de venda
    lErro = CF("PedidoDeVenda_Le_Completo", objPedidoVenda, MARCADO)
    If lErro <> SUCESSO And lErro <> 27633 Then gError 46562
    If lErro <> SUCESSO Then gError 46563
    
    'Lê as reservas do pedido de venda
    lErro = CF("ReservasPedido_Le", objPedidoVenda)
    If lErro <> SUCESSO Then gError 46564

    lErro = CF("NFiscal_Le_DatasEntrega", objPedidoVenda)
    If lErro <> SUCESSO Then gError 183543

    'Transfere os dados de objPedidoVenda para objNFiscal
    Call Transfere_Dados_Pedido_NFiscal(objPedidoVenda, objNFiscal, bFaturadoParc)

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
'        objNFiscal.iTipoNFiscal = DOCINFO_NFISVPV
'    End If
       
    Set gobjNFiscal = objNFiscal
       
    'Carrega os dados de objNFiscal na tela.
    lErro = Traz_NFiscal_Tela(objNFiscal, objPedidoVenda.lNumIntDoc)
    If lErro <> SUCESSO Then gError 46566
   
    'Guarda o código e a filial do Pedido
    glCodPedido = objPedidoVenda.lCodigo
    giFilialPedido = objPedidoVenda.iFilialEmpresa
    
    'Armazena o dado de fatura integral do pedido
    giFaturaIntegral = objPedidoVenda.iFaturaIntegral
    
    'se o faturamento é parcial
    If bFaturadoParc Then
    
        'Significa que deve recalcular as comissões, pois o BD está armazenando
        'as comissões referentes ao pedido integralmente
        iComissoesAlterada = REGISTRO_ALTERADO
        
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81513
    
    Else
    
        iComissoesAlterada = 0
    
    End If
    
    HoraSaida.PromptInclude = False
    HoraSaida.Text = ""
    HoraSaida.PromptInclude = True
    
    'Indica que o sistema não está mais carregando dados para a tela
    gbCarregandoTela = False
    
    Set gobjNFiscal = objNFiscal

    If (Month(objPedidoVenda.dtDataEmissao) <> Month(objNFiscal.dtDataEmissao)) Or (Year(objPedidoVenda.dtDataEmissao) <> Year(objNFiscal.dtDataEmissao)) Then
        bEmissaoForcaRecalculo = True
    Else
        bEmissaoForcaRecalculo = False
    End If
    
    If bFaturadoParc Or bEmissaoForcaRecalculo Then Call ValorTotal_Calcula

    Traz_Dados_Pedido_Tela = SUCESSO

    Exit Function

Erro_Traz_Dados_Pedido_Tela:

    'Indica que o sistema não está mais carregando dados para a tela
    gbCarregandoTela = False
    
    Traz_Dados_Pedido_Tela = gErr

    Select Case gErr

        Case 46562, 46564, 46566, 81513, 183543

        Case 46563
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157477)

    End Select

    Exit Function

End Function

Public Sub Transfere_Dados_Pedido_NFiscal(objPedidoVenda As ClassPedidoDeVenda, objNFiscal As ClassNFiscal, bFaturadoParc As Boolean)
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
Dim objProjetoInfo As Object, bRecalcularValorProdutos As Boolean
Dim dFator As Double, dPrecoTotal As Double, iItem As Integer
Dim objComissaoPorItem As ClassComissaoPorItem, objComissaoPorItemAux As ClassComissaoPorItem

On Error GoTo Erro_Transfere_Dados_Pedido_NFiscal
    
    Call gobjTelaProjetoInfo.Obtem_ProjetoInfo(objProjetoInfo, objPedidoVenda.lNumIntDoc, PRJ_CR_TIPO_PV)
    Set objNFiscal.objProjetoInfo = objProjetoInfo

    Set objNFiscal.objInfoAdic = objPedidoVenda.objInfoAdic

    bFaturadoParc = False
    bRecalcularValorProdutos = False
    
'    'Verifica se o pedido já foi faturado parcialmente
'    For Each objItemPV In objPedidoVenda.colItensPedido
'        If objItemPV.dQuantFaturada > 0 Then
'            bFaturadoParc = True
'            Exit For
'        End If
'    Next

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
    
    With objPedidoVenda
        'Passa os dados gerais da tela de pedido para a Nota Fiscal
        objNFiscal.dPesoBruto = .dPesoBruto
        objNFiscal.dPesoLiq = .dPesoLiq
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
        objNFiscal.dtDataReferencia = .dtDataReferencia
        objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
        objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
        
    End With
    
    'Para cada item do Pedido
   For Each objItemPV In objPedidoVenda.colItensPedido
                 
        lErro = CF("UM_Conversao_Trans", objItemPV.iClasseUM, objItemPV.sUnidadeMed, objItemPV.sUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 51562
                 
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
                
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> 25056 And lErro <> SUCESSO Then gError 51562
                If lErro = 25056 Then gError 51563
                                
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
                If lErro <> SUCESSO Then gError 105154
                                
                If dQuantidade > 0 Then
                                
                    'Adiciona o item na coleção de itens da NF
                    .dPrecoTotal = (dQuantidade * .dPrecoUnitario) * (1 + dPercDesc)
                    
                    If bFaturadoParc Then
                        dValorDesconto = Arredonda_Moeda(dPercDesc * (dQuantidade * .dPrecoUnitario))
                    Else
                        dValorDesconto = .dValorDesconto
                    End If
                     
                    'Adiciona o item na coleção de itens da NF
                    Set objItemNF = objNFiscal.ColItensNF.Add(0, .iItem, .sProduto, .sUnidadeMed, dQuantidade, .dPrecoUnitario, dPercDesc, dValorDesconto, .dtDataEntrega, .sDescricao, .dValorAbatComissao, .lNumIntDoc, 0, 0, .lNumIntDoc, colAlocacoes, 0, "", "", .iStatus, 0, "", 0, 0, .iControleEstoque, .sUMEstoque, .iClasseUM, 0)
                    
                    lErro = CF2(Me, "NFP_Transfere_Dados_Pedido_NFiscal", objItemNF, objItemPV, dQuantidade, objNFiscal)
                    If lErro <> SUCESSO Then gError 177051
                    
                    Set objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).objTributacaoItemNF = objItemTributacaoNF
                    
                    'Guarda a número interno e a quantidade inicial do ítem
                    gcolItemPedido.Add .lNumIntDoc
                    gcolItemQuant.Add dQuantidade
                    gcolItemUM.Add .sUnidadeMed
            
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

                    If gobjFAT.iTrazDescPVParaNF = DESMARCADO And objItemNF.dValorDesconto > DELTA_VALORMONETARIO Then
                        bRecalcularValorProdutos = True
                        dPrecoTotal = (objItemNF.dQuantidade * objItemNF.dPrecoUnitario) - objItemNF.dValorDesconto
                        objItemNF.dPercDesc = 0
                        objItemNF.dValorDesconto = 0
                        objItemNF.dPrecoUnitario = Arredonda_Moeda(dPrecoTotal / objItemNF.dQuantidade)
                        objItemNF.objTributacaoItemNF.dValorUnitTrib = objItemNF.dPrecoUnitario
                    End If
                    
                End If
            
            End With
        End If
    Next
    
    'Transfere os dados das comissões do pedido para a coleção de coissões da Nota Fiscal
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

    For Each objParcelaPV In objPedidoVenda.colParcelas
        With objParcelaPV
            objNFiscal.ColParcelaReceber.Add 0, 0, .iNumParcela, STATUS_LANCADO, .dtDataVencimento, .dtDataVencimento, 0, .dValor, 1, CARTEIRA_CARTEIRA, COBRADOR_PROPRIA_EMPRESA, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0.1, 0, "", .dValor
        End With
    Next
    
    'Reinicializa as coleções de itens
    Set gcolItemPedido = New Collection
    Set gcolItemQuant = New Collection
    Set gcolItemUM = New Collection
    
    
    'Carrega os dados globais da tributação do PV para a NF
    Call gobjTribTab.NFiscal_ObterTrib_Pedido(objPedidoVenda, objNFiscal)
    
    If bRecalcularValorProdutos Then
        objNFiscal.dValorItens = 0
        bFaturadoParc = True
    End If
    
    Exit Sub
    
Erro_Transfere_Dados_Pedido_NFiscal:

    Select Case gErr
    
        Case 51562, 105154, 177051
        
        Case 51563
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157478)
    
    End Select
    
End Sub

Public Function Traz_NFiscal_Tela(objNFiscal As ClassNFiscal, Optional lNumIntPedido As Long = 0) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim bAlterouCT As Boolean

On Error GoTo Erro_Traz_NFiscal_Tela
    
'    'Se for versao LIGHT testa chaveamento
'    If giTipoVersao = VERSAO_LIGHT And lNumIntPedido = 0 Then
'
'        lErro = objUserControl.CT_Chavea_Trazendo(objNFiscal, bAlterouCT)
'        If lErro <> SUCESSO Then gError 25951
'
'        If bAlterouCT Then
'
'            'Chama o Traz_NFiscal_Tela de objCT chaveado
'            lErro = objUserControl.objCT.Traz_NFiscal_Tela(objNFiscal)
'            If lErro <> SUCESSO Then gError 25952
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
        If lErro <> SUCESSO Then gError 25953
    End If
    
    'Limpa a tela de Nota Fiscal
    lErro = Limpa_Tela_NFiscalPedido()
    If lErro <> SUCESSO Then gError 500152
    
    Set gobjNFiscal = objNFiscal

    lErro = Traz_NFiscal_Tela2(objNFiscal, lNumIntPedido)
    If lErro <> SUCESSO Then gError 25954
    
    lErro = CF2(Me, "NFCF_Cliente_Traz_NF_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 131891
       
    iAlterado = 0
    iComissoesAlterada = 0
    
    If lNumIntPedido = 0 Then
        MsgAutomatica.Value = vbUnchecked
        gbCarregandoTela = False
    Else
        MsgAutomatica.Value = vbChecked
    End If

    Traz_NFiscal_Tela = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela:

    gbCarregandoTela = False

    Traz_NFiscal_Tela = gErr

    Select Case gErr

        Case 25951, 25952, 25953, 25954, 500152, 131891 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157479)

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
Dim iLote As Integer
Dim dtDataContabilidade As Date
Dim bCancel As Boolean
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim dPercDesc As Double

On Error GoTo Erro_Traz_NFiscal_Tela2
    
    objCliente.lCodigo = objNFiscal.lCliente
    'Lê o cliente
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 19062 Then gError 46567
    If lErro <> SUCESSO Then gError 46568
    'Lê as filiais do cliente
    lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 46569
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
    
    'Coloca TabelaPreco na Tela
    If objNFiscal.iTabelaPreco > 0 Then
        TabelaPreco.Text = objNFiscal.iTabelaPreco
        Call TabelaPreco_Validate(bSGECancelDummy)
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

    If objNFiscal.iCanal > 0 Then
        CanalVenda.Text = objNFiscal.iCanal
        Call CanalVenda_Validate(bSGECancelDummy)
    End If

    Call DateParaMasked(DataEmissao, objNFiscal.dtDataEmissao)
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

    If objNFiscal.iCodTransportadora <> 0 Then
        Transportadora.Text = objNFiscal.iCodTransportadora
        Call Transportadora_Validate(bSGECancelDummy)
    End If

    If objNFiscal.iCodTranspRedesp <> 0 Then
        TranspRedespacho.Text = CStr(objNFiscal.iCodTranspRedesp)
        Call TranspRedespacho_Validate(bCancel)
    End If

    RedespachoCli.Value = objNFiscal.iDetPagFrete
    
    For iIndice = 0 To TipoFrete.ListCount - 1
        If TipoFrete.ItemData(iIndice) = objNFiscal.iFreteRespons Then
            TipoFrete.ListIndex = iIndice
            Exit For
        End If
    Next


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
    
    lErro = CF2(Me, "NFP_Traz_NFiscal_Tela2", objNFiscal, lNumIntPedido)
    If lErro <> SUCESSO Then gError 46850
    
    gobjTribTab.sNatOpInterna = objNFiscal.sNaturezaOp
    
    'Preenche o Grid com os itens da Nota Fiscal
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 46694

    If objNFiscal.iTipoNFiscal <> 0 Then

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        
        'Le o Tipo para saber se tem Movimento de Estoque
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO Then gError 83376
        
        If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
            'Carrega ItensNF com Rastreamentos
            lErro = gobjRastreamento.Carrega_RastroItensNF(objNFiscal)
            If lErro <> SUCESSO Then gError 83377
        End If

    End If

    'Carrega o Tab Comissões
    lErro = objTabComissoes.Carrega_Tab_Comissoes(objNFiscal)
    If lErro <> SUCESSO Then gError 46693

    If lNumIntPedido > 0 Then bGerarDados = True

    'Preenche o Grid com as Alocações dos itens da Nota Fiscal
    lErro = CF2(Me, "NFP_Preenche_GridAlocacoes", objNFiscal, bGerarDados)
    If lErro <> SUCESSO Then gError 39026

    'Se a nota fiscal já existe
    If objNFiscal.lNumIntDoc <> 0 Then
    
        'Traz os dados contábeis para a tela (contabilidade)
        lErro = objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
        If lErro <> SUCESSO And lErro <> 36326 Then gError 46695
    
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
    
    CodVerificacao.Caption = objNFiscal.sCodVerificacaoNFe
    
    Traz_NFiscal_Tela2 = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela2:

    Traz_NFiscal_Tela2 = gErr

    Select Case gErr

        Case 46567, 46569, 46692, 46694, 46693, 46695, 55861, 83376, 83377
        
        Case 46568
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157480)

    End Select

    Exit Function

End Function

Function Preenche_GridItens(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com os itens da Nota Fiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim sProdutoEnxuto As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim dValorProdutos As Double
Dim dValorDesconto  As Double
Dim dValorTotal As Double
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
        If lErro <> SUCESSO Then gError 46696

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
        
        'dValorDesconto = objItemNF.dPercDesc * objItemNF.dQuantidade * objItemNF.dPrecoUnitario
        GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objItemNF.dValorDesconto, "Standard")

        dValorTotal = (objItemNF.dPrecoUnitario * objItemNF.dQuantidade) - objItemNF.dValorDesconto
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(dValorTotal, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col) = Format(objItemNF.dValorTotal + objItemNF.dValorDesconto, "Standard")

        If objItemNF.dtDataEntrega <> DATA_NULA Then
            GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col) = Format(objItemNF.dtDataEntrega, "dd/mm/yyyy")
        End If
        
        If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        
            If objItemNF.dComissao <> -1 Then
                GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col) = Format(objItemNF.dComissao, "Percent")
            Else
                GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col) = ""
            End If
        
        End If
        
        lErro = CF2(Me, "NFP_Preenche_GridItens", objItemNF, iIndice)
        If lErro <> SUCESSO Then gError 197871
        
        lErro = CF2(Me, "NF_Preenche_GridItens", objNFiscal, objItemNF, objProduto, iIndice)
        If lErro <> SUCESSO Then gError 177926
        
        If iIndice > gcolItemPedido.Count Then gcolItemPedido.Add objItemNF.lNumIntItemPedVenda
        If iIndice > gcolItemQuant.Count Then gcolItemQuant.Add objItemNF.dQuantidade
        If iIndice > gcolItemUM.Count Then gcolItemUM.Add objItemNF.sUnidadeMed
        
    Next
    
    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice

    Set gobjNFiscal = objNFiscal
    Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)
    
    ValorProdutos.Caption = Format(dValorProdutos - objNFiscal.dValorDesconto, "Standard")

    lErro = CF2(Me, "NFCF_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132040
    
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objNFiscal.objProjetoInfo, objNFiscal.lNumIntDoc, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181676
    
    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr

        Case 46696, 181676
        
        Case 132040, 177926, 197871

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157481)

    End Select

    Exit Function

End Function

Public Function NFP_Preenche_GridItens(ByVal objCT As Object, ByVal objItemNF As ClassItemNF, ByVal iLinha As Integer) As Long
    NFP_Preenche_GridItens = SUCESSO
End Function

Public Function NFP_Preenche_GridAlocacoes(ByVal objCT As Object, objNFiscal As ClassNFiscal, Optional bGeraAlocacao As Boolean = False) As Long
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

On Error GoTo Erro_NFP_Preenche_GridAlocacoes

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
                GridAlocacao.TextMatrix(iIndice1, iGrid_ProdutoAloc_Col) = GridItens.TextMatrix(objNFiscal.ColItensNF.Item(iIndice).iItem, iGrid_Produto_Col)
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
                lErro = CF2(Me, "NFP_Alocacao_Processa")
                If lErro <> SUCESSO And lErro <> 23796 Then gError 51602
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
    
    NFP_Preenche_GridAlocacoes = SUCESSO

    Exit Function

Erro_NFP_Preenche_GridAlocacoes:

    NFP_Preenche_GridAlocacoes = gErr
    
    Select Case gErr
    
        Case 46849, 46865, 46866, 51602
        
        Case 51536
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157482)
            
    End Select

    Exit Function
    
End Function

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
            If lErro <> SUCESSO Then gError 46701
        End If
'*** fim tributacao

'        If giTipoVersao = VERSAO_FULL Then
            'Torna Frame correspondente ao Tab selecionado visivel
            Frame1(TabStrip1.SelectedItem.Index).Visible = True
            'Torna Frame atual visivel
            Frame1(iFrameAtual).Visible = False
'        ElseIf giTipoVersao = VERSAO_LIGHT Then
'            'Torna Frame correspondente ao Tab selecionado visivel
'            FrameL(TabStrip1.SelectedItem.Index).Visible = True
'            'Torna Frame atual visivel
'            FrameL(iFrameAtual).Visible = False
'        End If
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If TabStrip1.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao
    
        If TabStrip1.SelectedItem.Caption = TITULO_TAB_COMISSOES Then
        
            'se for tab de comissoes e a empresa
            'utilizar regras para o calculo de comissoes
            lErro = objTabComissoes.Comissoes_Calcula()
            If lErro <> SUCESSO Then gError 101518
    
        End If
    
        Select Case iFrameAtual
        
            Case TAB_DadosPrincipais
                Parent.HelpContextID = IDH_NF_PEDIDO_DADOS_PRINCIPAIS
            
            Case TAB_Itens
                Parent.HelpContextID = IDH_NF_PEDIDO_ITENS
            
            Case TAB_Complemento
                Call Calcula_Msg
                Parent.HelpContextID = IDH_NF_PEDIDO_COMPLEMENTO
            
            Case TAB_Comissoes
                
                Parent.HelpContextID = IDH_PEDIDO_VENDA_COMISSOES
            
            Case TAB_Almoxarifado
                Parent.HelpContextID = IDH_NF_PEDIDO_ALMOXARIFADO
            
            Case TAB_Tributacao
                Parent.HelpContextID = IDH_NF_PEDIDO_TRIBUTACAO
            
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_NF_PEDIDO_CONTABILIZACAO
            
        End Select
    
    End If

    Exit Sub
    
Erro_TabStrip1_Click:

    Select Case gErr

        Case 46701, 101518
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157483)

    End Select

    Exit Sub

End Sub

Private Function Limpa_Tela_NFiscalPedido() As Long
'Limpa a tela de Notas Fiscais

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_NFiscalPedido

    'Indica que está limpando a tela
    gbLimpandoTela = True
    
    'Limpa os campos comuns da tela
    Call Limpa_Tela(Me)
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    Set gobjAnotacao = Nothing
    Set gobjInfoAdic = Nothing
    
'    If giTipoVersao = VERSAO_FULL Then
        Serie.Text = ""
        Call Serie_Validate(bSGECancelDummy)
        TabelaPreco.Text = ""
'    End If
    
    VolumeMarca.Text = ""
    VolumeEspecie.Text = ""
    
    'Recarrega a combo VolumeEspecie e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)

    'Recarrega a combo VolumeMarca e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
    
    'Limpa os campos não limpos pelo Limpa_Tela
    Status.Caption = ""
    NFiscal.Caption = ""
    Cliente.Caption = ""
    TipoFrete.ListIndex = 0
    Filial.Caption = ""
    FilialPedido.Text = giFilialEmpresa
    Call FilialPedido_Validate(bSGECancelDummy)
    ValorProdutos.Caption = ""
    ValorProdutos2.Caption = ""
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
    EletronicaFed.Value = vbUnchecked
    
    Call CF2(Me, "NFCF_Default_DataSaida")

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    Call DataEmissao_Validate(bSGECancelDummy)
    
    giFaturaIntegral = 0
    glCodPedido = 0
    giFilialPedido = giFilialEmpresa
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next
    
    'Limpa os grids
    Call Grid_Limpa(objGridItens)
    objTabComissoes.Limpa_Grid_Comissoes
    Call Grid_Limpa(objGridAlocacoes)
    
    Recibo.Value = vbUnchecked
    
    NumNFe.Caption = ""
    CodVerificacao.Caption = ""
    
    Call gobjTribTab.Limpa_Tela
        
    'Limpa a contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade
        
    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
    'Seleciona Série padrão
    lErro = Seleciona_Serie_Padrao()
    If lErro <> SUCESSO Then gError 500149

    MsgAutomatica.Value = vbChecked

    iValorDescontoAlterado = 0
    iAlterado = 0
    iComissoesAlterada = 0
    iValorDespesasAlterado = 0
    iValorFreteAlterado = 0
    iValorSeguroAlterado = 0
    giFilialEntregaAlterada = 0
    giFilialEntregaManual = 0
    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0

    Call CF2(Me, "NFCF_Limpa_Tela")
 
    'Indica que está limpando a tela
    gbLimpandoTela = False
    
    Exit Function
    
Erro_Limpa_Tela_NFiscalPedido:
    
    Limpa_Tela_NFiscalPedido = gErr
    
    Select Case gErr
    
        Case 500149
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157484)
    
    End Select
    
    'Indica que não está mais limpando a tela
    gbLimpandoTela = False
    
    Exit Function
    
End Function

Public Sub Quantidade_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
    Call CF2(Me, "NFP_Quantidade_GotFocus")
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

    '*** 12/04/02 - Tulio ***
    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO
    '************************************

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

Public Sub PrecoTotal_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrecoTotal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub PrecoTotal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub PrecoTotal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PrecoTotal
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
Dim vbMsgRes As VbMsgBoxResult 'Inserido por Wagner

On Error GoTo Erro_GridItens_KeyDown

    'Armazena o número de linhas existentes e o item atual selecionado
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iItemAtual = GridItens.Row

    'Se a tecla for a delete e o Pedido só aceitar faturamento integral
    'não permite a exclusão da linha
    If KeyCode = vbKeyDelete And giFaturaIntegral <> 0 Then gError 46633
        
    'Se a tecla for a delete e o Grid estiver vazio, Sai da Sub
    If KeyCode = vbKeyDelete And objGridItens.iLinhasExistentes <= 0 Then Exit Sub 'Inserido por Leo em 16/01/02
        
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

        lErro = CF2(Me, "NFCF_ValorDesconto_Validate")
        If lErro <> SUCESSO Then gError 132041

''*** incluidos p/tratamento de tributacao *******************************
        'Ataliza a tributação
        lErro = gobjTribTab.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 46702

'*** fim tributacao

        'Recalcula o valor total
        Call ValorTotal_Calcula
    
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81514
    
        'Rastreamento
        lErro = gobjRastreamento.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 83379
    
        For iIndice = 1 To gobjNFiscal.ColItensNF.Count
            gobjNFiscal.ColItensNF(iIndice).iItem = iIndice 'Incluído para ao alterar a grade ou a reserva não fazer do item errado
        Next

    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr
    
        Case 46702, 81514, 83379
        
        Case 46633
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_FATURA_INTEGRAL", gErr)
            
        Case 132041 'Inserido por Wagner
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157485)
            
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

Public Sub Remove_Alocacoes(iItemAtual As Integer)
'Remove do grid de alocações as alocações do item passado

Dim iIndice As Integer
Dim iItemLinha As Integer
Dim bExcluiuLinha As Boolean

    iIndice = 1
    bExcluiuLinha = False
    
    'Percorre as linhas do grid de alocações
    Do While iIndice <= objGridAlocacoes.iLinhasExistentes
        'Verifica se o item pesquisado é o da linha iIndice
        If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))) > 0 Then
            iItemLinha = GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)
            If iItemLinha = iItemAtual Then
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

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 46706
        
        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridItens
            Case GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 46703

            'se for o GridComissoes
            Case GridComissoes.Name

                lErro = objTabComissoes.Saida_Celula_GridComissoes(objGridInt)
                If lErro <> SUCESSO Then gError 46704

            'Rastreamento
            Case GridRastro.Name

                lErro = gobjRastreamento.Saida_Celula()
                If lErro <> SUCESSO Then gError 83380

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 46705

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 46703, 46704, 46705, 46706, 83380

        Case Else
            
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157486)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Quantidade
        Case iGrid_Quantidade_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 46707

        'PrecoUnitario
        Case iGrid_ValorUnitario_Col
            lErro = Saida_Celula_PrecoUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 46708

        'PercDesc
        Case iGrid_PercDesc_Col
            lErro = Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then gError 46709

        'Desconto
        Case iGrid_Desconto_Col
            lErro = Saida_Celula_Desconto(objGridInt)
            If lErro <> SUCESSO Then gError 46710

        Case iGrid_TabPrecoItem_Col
            If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
                lErro = Saida_Celula_TabPrecoItemNF(objGridInt)
                If lErro <> SUCESSO Then gError 46710
            End If

        Case iGrid_ComissaoItem_Col
            If gobjFAT.iExibeComissaoItensPV = MARCADO Then
                lErro = Saida_Celula_ComissaoItemNF(objGridInt)
                If lErro <> SUCESSO Then gError 46710
            End If

    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 46707, 46708, 46709, 46710

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157487)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
        'Produto
        Case Produto.Name, UnidadeMed.Name
        
        'Nas demais
        Case TabPrecoItemNF.Name, ComissaoItemNF.Name, PrecoUnitario.Name, PercentDesc.Name, Desconto.Name
            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 46713

            If iProdutoPreenchido = PRODUTO_VAZIO Or gcolItemPedido.Count = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
        Case Quantidade.Name
            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 46713

            If iProdutoPreenchido = PRODUTO_VAZIO Or gcolItemPedido.Count = 0 Or giFaturaIntegral <> 0 Or left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

    End Select

    'Rastreamento
    lErro = gobjRastreamento.Rotina_Grid_Enable(iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 83378

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 46713, 83378

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157488)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim iSairCelula As Integer
Dim dPrecoUnitario As Double
Dim dQuantidadeAnterior As Double
Dim objProduto As New ClassProduto
Dim objKitVenda As New ClassKitVenda
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objItemNF As ClassItemNF
Dim objItemPV As ClassItemPedido
Dim dQtdePV As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    iSairCelula = False
    
    'Verifica o preenchimento de Quantidade
    If Len(Trim(Quantidade.ClipText)) > 0 Then

        If giComplemento = DESMARCADO Then
            lErro = Valor_Positivo_Critica(Quantidade.Text)
        Else
            lErro = Valor_NaoNegativo_Critica(Quantidade.Text)
        End If
        If lErro <> SUCESSO Then gError 46716

        dQuantidade = CDbl(Quantidade.Text)
        Quantidade = Formata_Estoque(dQuantidade)
        
        dQuantidadeAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
                   
        If Not iSairCelula Then
        
            Set objItemNF = gobjNFiscal.ColItensNF(GridItens.Row)
            
            If objItemNF.iRecebForaFaixa = ERRO_E_REJEITA_RECEBIMENTO Then
                Set objItemPV = New ClassItemPedido
                
                objItemPV.lNumIntDoc = gcolItemPedido(GridItens.Row)
                objItemPV.iFilialEmpresa = Codigo_Extrai(FilialPedido.Text)
                
                lErro = CF("ItensPedido_Le_NumIntDoc1", objItemPV)
                If lErro <> SUCESSO Then gError 196963
                
                dQtdePV = objItemPV.dQuantidade - objItemPV.dQuantCancelada
                
                If dQuantidade > Arredonda_Estoque(gcolItemQuant(GridItens.Row) + (dQtdePV * objItemNF.dPercentMaisReceb)) Then gError 46714 'Alterado por Wagner
            End If
            
        End If
   
        If dQuantidadeAnterior = dQuantidade Then iSairCelula = True

    End If

    If Not iSairCelula Then
    
        GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)
        
        objProduto.sCodigo = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177585
        
        If objProduto.iKitVendaComp <> MARCADO Then
            'Tenta fazer uma alocação p\ oitem de acordo com a nova UM
            lErro = CF2(Me, "NFP_Alocacao_Processa")
            If lErro <> SUCESSO And lErro <> 23796 Then gError 46717
            If lErro = 23796 Then gError 23796
    
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
            
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81515
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 46718
 
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))

    If dPrecoUnitario > 0 Then Call PrecoTotal_Calcula(GridItens.Row, False)
    
    '*** incluidos p/tratamento de tributacao ****************
    If GridItens.Row <= objGridItens.iLinhasExistentes Then
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 46715
    End If
    '*** fim tributacao

    Call ValorTotal_Calcula

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 23796 '??? Jones 09/02/05 passou a tratar 23796
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO2", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 46714
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_FATURAR_MENOR", gErr, Arredonda_Estoque(gcolItemQuant(GridItens.Row) + (dQtdePV * objItemNF.dPercentMaisReceb)))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 46715, 46716, 46717, 46718, 81515
            GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidadeAnterior)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 177585, 177586, 177587, 196963
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157489)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_PrecoUnitario(objGridInt As AdmGrid) As Long
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
        If lErro <> SUCESSO Then gError 46720

        PrecoUnitario.Text = Format(PrecoUnitario.Text, gobjFAT.sFormatoPrecoUnitario)
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 46721
    
    Call PrecoTotal_Calcula(GridItens.Row, False)

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 46719
'*** fim tributacao

    Call ValorTotal_Calcula

    Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = gErr

    Select Case gErr

        Case 46719, 46720, 46721
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157490)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_PercentDesc(objGridInt As AdmGrid) As Long
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
        If lErro <> SUCESSO Then gError 46723
                
        If dPercentDesc <> dPercentDescAnterior Then
        
            'Verifica se o percentual é de 100%
            If dPercentDesc = 100 Then gError 46725
    
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
    If lErro <> SUCESSO Then gError 46724

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) > 0 Then dPrecoUnitario = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))
    If dPrecoUnitario > 0 Then
    
        Call PrecoTotal_Calcula(GridItens.Row, False)
        
''*** incluidos p/tratamento de tributacao *******************************
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 46722
'*** fim tributacao

        Call ValorTotal_Calcula

    End If

    Saida_Celula_PercentDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDesc:

    Saida_Celula_PercentDesc = gErr

    Select Case gErr

        Case 46722, 46723, 46724
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 46725
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157491)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long
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

    'Veifica o preenchimento de Desconto
    If Len(Trim(Desconto.ClipText)) > 0 Then
        'Faz a critica do Desconot
        lErro = Valor_NaoNegativo_Critica(Desconto.Text)
        If lErro <> SUCESSO Then gError 46727
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
            If dDesconto >= dPrecoTotal Then gError 46728
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
        If lErro <> SUCESSO Then gError 46726
        '*** fim tributacao
        
        Call ValorTotal_Calcula
    
    End If
    
    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = gErr

    Select Case gErr

        Case 46726, 46727
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 46728
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCONTO_MAIOR_OU_IGUAL_PRECO_TOTAL", gErr, GridItens.Row, dDesconto, dPrecoTotal)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157492)

    End Select

    Exit Function

End Function


'Janaina - Alterei para Public
Public Function NFP_Alocacao_Processa(ByVal objCT As Object) As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dQuantidade As Double
Dim sUM As String

On Error GoTo Erro_NFP_Alocacao_Processa

    'Remove as alocações do grid de alocações
    Call Remove_Alocacoes(GridItens.Row)

    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 46738

    objProduto.sCodigo = sProduto
    
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 46739
    If lErro = 28030 Then gError 46741 'Não encontrou

    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    'Verifica se o produto tem o controle de estoque <> PRODUTO_CONTROLE_SEM_ESTOQUE
    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 And gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA And gobjTribTab.mvarobjNFTributacao.iTipoNFiscal <> DOCINFO_NFISRETPV Then

        sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
        
        lErro = CF2(Me, "NFP_AlocaAlmoxarifadoPadrao", dQuantidade, objProduto, sUM)
        If lErro <> SUCESSO And lErro <> 23796 Then gError 46740
        If lErro = 23796 Then gError 23796
    
    End If

    NFP_Alocacao_Processa = SUCESSO

    Exit Function

Erro_NFP_Alocacao_Processa:

    NFP_Alocacao_Processa = gErr

    Select Case gErr

        Case 46738, 46739, 46740, 23796

        Case 46741
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157493)

    End Select

    Exit Function

End Function

Function NFP_AlocaAlmoxarifadoPadrao(ByVal objCT As Object, dQuantidade As Double, objProduto As ClassProduto, sUM As String) As Long
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

On Error GoTo Erro_NFP_AlocaAlmoxarifadoPadrao

    'Faz a conversão da UM da tela para a UM de estoque
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 46743

    'Converte a quantidade para a UM de estoque
    dQuantAlocar = dQuantidade * dFator

    'Busca o Almoxarifado padrão
    lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 46745
    If lErro = 23796 Then gError 23796
    
    'Se encontrou
    If iAlmoxarifado > 0 Then

        objAlmoxarifado.iCodigo = iAlmoxarifado
        'Lê o Aloxarifado
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> 25056 And lErro <> SUCESSO Then gError 46746
        If lErro = 25056 Then gError 46749
        
        objEstoqueProduto.sProduto = objProduto.sCodigo
        objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
        'Le os estoques desse produto nesse almoxarifado
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 46744
        If lErro = 21306 Then gError 46750 'Não encontrou

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
        
        dQuantDisponivel = objEstoqueProduto.dQuantDisponivel
        'Remove as alocações do produto do grid de alocações
        Call Remove_Alocacoes(GridItens.Row)

        iLinhas = objGridAlocacoes.iLinhasExistentes

        dQuantAlocar = Arredonda_Estoque(dQuantAlocar)
        
        'Verifica se a Quantidade disponível é maior que a quantidade a alocar
        If dQuantAlocar <= dQuantDisponivel Or gobjMAT.iAceitaEstoqueNegativo = MARCADO Then

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
                        If lErro <> SUCESSO Then gError 46747
                        'Adiciona na coleção de produtos
                        colOutrosProdutos.Add sProduto
                    End If
                Next
    
                'Chama a tela de Localização de Produto
                Call Chama_Tela_Modal("LocalizacaoProduto1", objItemNF, colOutrosProdutos, dQuantAlocar, DOCINFO_NFISVPV)
                If giRetornoTela = vbCancel Then gError 46752 'Se nada foi feito lá
                If giRetornoTela = vbOK Then
    
                    'Se o produto foi substituido
                    If objProduto.sCodigo <> objItemNF.sProduto Then gError 46751
                                    
                    sProdutoEnxuto = String(STRING_PRODUTO, 0)
    
                    lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
                    If lErro Then gError 46748
    
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

    NFP_AlocaAlmoxarifadoPadrao = SUCESSO

    Exit Function

Erro_NFP_AlocaAlmoxarifadoPadrao:

    NFP_AlocaAlmoxarifadoPadrao = gErr

    Select Case gErr

        Case 46743, 46744, 46745, 46746, 46747, 46748, 23796

        Case 46749
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, iAlmoxarifado)

        Case 46750
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case 46751
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_SUBSTITUIDO", gErr)

        Case 46752
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_LOCALIZACAO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157494)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objNFiscal As ClassNFiscal) As Long
'Recolhe os dados da tela

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objProjetoInfo As Object

On Error GoTo Erro_Move_Tela_Memoria
            
    If Len(Trim(Pedido.ClipText)) > 0 Then objNFiscal.lNumPedidoVenda = CLng(Pedido.Text)
    objNFiscal.iFilialPedido = Codigo_Extrai(FilialPedido.Text)
    
    'Verifica se o Cliente esta preenchido
    If Len(Trim(Cliente.Caption)) > 0 Then
    
        objCliente.sNomeReduzido = Cliente.Caption
        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 46587
        'Se não achou o Cliente --> erro
        If lErro <> SUCESSO Then gError 46588
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
    
    objNFiscal.sVolumeNumero = VolumeNumero.Text
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    
    If Len(Trim(NFiscal.Caption)) > 0 Then
        objNFiscal.lNumNotaFiscal = CLng(NFiscal.Caption)
    Else
        objNFiscal.lNumNotaFiscal = 0
    End If
    
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)

    lErro = CF2(Me, "NFCF_Move_DataSaida_Memoria", objNFiscal)
    If lErro <> SUCESSO Then gError 132295
    
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
    objNFiscal.iCanal = Codigo_Extrai(CanalVenda.Text)
    objNFiscal.sNumPedidoTerc = NumPedidoTerc.Text
    objNFiscal.iTipoNFiscal = DOCINFO_PADRAO_TELA 'DOCINFO_NFISVPV
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.dtDataEntrada = DATA_NULA
    objNFiscal.dtDataVencimento = DATA_NULA
    objNFiscal.iStatus = STATUS_LANCADO
    objNFiscal.dtDataReferencia = DATA_NULA
        
    'Incluído por Luiz Nogueira em 21/08/03
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objNFiscal.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objNFiscal.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
        
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181761
    
    Set objNFiscal.objProjetoInfo = objProjetoInfo
    Set objNFiscal.objInfoAdic = gobjInfoAdic
    objNFiscal.dValorItens = StrParaDbl(ValorProdutos2.Caption)
    objNFiscal.dValorDescontoItens = StrParaDbl(ValorDescontoItens.Text)
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 46587, 132295, 181761 'Inserido por Wagner
        
        Case 46588
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157495)
            
    End Select
            
    Exit Function
        
End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim dValorTotal As Double
Dim dValorDesconto As Double
Dim objNFiscal As New ClassNFiscal
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objBloqueioPV As ClassBloqueioPV
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iNFSE As Integer

On Error GoTo Erro_Gravar_Registro
    
    'Verifica se os dados de algum pedido foi trazido para a tela
    If glCodPedido = 0 Then gError 46591
    
    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(Pedido.Text)) = 0 Then gError 46578
    If Len(Trim(FilialPedido.Text)) = 0 Then gError 46579
    If Len(Trim(Cliente.Caption)) = 0 Then gError 46580
    If Len(Trim(Serie.Text)) = 0 Then gError 46581
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 46583
    If Len(Trim(gobjTribTab.sNatOpInterna)) = 0 Then gError 46585
    If Len(Trim(FilialEntrega.Text)) = 0 Then gError 51533
    
    'Verifica se tem itens no grid
    If objGridItens.iLinhasExistentes = 0 Then gError 46586

    'Verifica se o código de pedido é o mesmo que foi trazido para a tela
    If glCodPedido <> StrParaLong(Pedido.Text) Or giFilialPedido <> Codigo_Extrai(FilialPedido.Text) Then
        'Avisa que os dados da tela são de outro pedido e pergunta se deseja prosseguir.
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PEDIDO_TELA_NAO_UTILIZADO", glCodPedido, giFilialPedido)
        If vbMsgRes = vbNo Then gError 46592 'Sai da rotina.
        
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
        If CDbl(PesoLiquido.Text) > CDbl(PesoBruto.Text) Then gError 46629
    End If
    
    'Verifica se a Data de Saida e menor que a Data de Emissão
    If Len(Trim(DataSaida.ClipText)) <> 0 Then
        If CDate(DataSaida.Text) < CDate(DataEmissao.Text) Then gError 46593
    End If
    
    dValorTotal = CDbl(IIf(Len(Trim(ValorTotal.Caption)) > 0, ValorTotal.Caption, 0))

    'Se o total for negativo --> Erro
    If dValorTotal < 0 Then gError 46628

    'Verifica se o Valor do Desconto é maior que a soma dos demais valores
    If Len(Trim(ValorDesconto.Text)) > 0 Then dValorDesconto = CDbl(ValorDesconto.Text)
    lErro = DescontoGlobal_Testa(dValorDesconto)
    If lErro <> SUCESSO Then gError 45565
    
    lErro = gobjTribTab.Valida_Dados
    If lErro <> SUCESSO Then gError 45565
    
    objNFiscal.lNumPedidoVenda = glCodPedido
    objNFiscal.iFilialPedido = giFilialPedido
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.iTipoNFiscal = DOCINFO_PADRAO_TELA 'DOCINFO_NFISVPV
   
    objPedidoVenda.lCodigo = glCodPedido
    objPedidoVenda.iFilialEmpresa = giFilialPedido
    
    'Lê os bloqueios do Pedido de venda
    lErro = CF("BloqueiosPV_Le", objPedidoVenda)
    If lErro <> SUCESSO Then gError 46634
    
    'Verifica se há algum bloqueio que não permita a geração da nota fiscal
    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV
        If objBloqueioPV.dtDataLib = DATA_NULA Then
            If objBloqueioPV.iTipoDeBloqueio = BLOQUEIO_CREDITO Then gError 46635
            If objBloqueioPV.iTipoDeBloqueio = BLOQUEIO_TOTAL Then gError 46636
        End If
    Next
                
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 83453
    
    'Não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 83454
                
    lErro = Valida_Dados_GridItens()
    If lErro <> SUCESSO Then gError 46637

    'VAlida os dados do grid de comissões
    lErro = objTabComissoes.Valida_Grid_Comissoes()
    If lErro <> SUCESSO Then gError 46630
    
    lErro = gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
    If lErro <> SUCESSO Then gError 83381

    Set objNFiscal.objRastreamento = gobjRastreamento
    
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 46631
    
    lErro = Move_Grids_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 46596
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = gobjTribTab.Valida_TributacaoNF_Gravacao(objNFiscal)
    If lErro <> SUCESSO Then gError 46461
    
    lErro = Calcula_Msg
    If lErro <> SUCESSO Then gError 46461
    
    'se for pra verificar limite de credito
    If gobjCRFAT.iCreditoVerificaLimite = VERIFICAR_LIMITE_CREDITO Then
        If Len(Trim(NFiscal.Caption)) = 0 Then
            'verifica se o cliente tem crédito
            lErro = CF("NFiscal_Testa_Credito", objNFiscal)
            If lErro <> SUCESSO Then gError 25739
        End If
    End If
    
    'mover a parte do rastreamento
    lErro = gobjRastreamento.Move_Rastro_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 83382
    
    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataEmissao.Text))
    If lErro <> SUCESSO Then gError 92046
    
    lErro = CF("NFiscal_Valida_Diversos", objNFiscal)
    If lErro <> SUCESSO Then gError 181786
    
    lErro = CF("NFiscalPedido_Grava", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 46632
    
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
    
    If ISSerieEletronica(objNFiscal.sSerie) Then
        lErro = CF("NFE_Grava1", objNFiscal.lNumIntDoc, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 199998
    End If
    
    lErro = CF("Verifica_NFiscal_Servico_Eletronica", objNFiscal, iNFSE)
    If lErro <> SUCESSO Then gError 207062
    
    If iNFSE = 1 Then
        lErro = CF("NFSE_Grava1", objNFiscal.lNumIntDoc, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 207063
    End If
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 46565, 83381, 83382, 83453, 92046, 130406, 181786, 199998, 207062, 207063
        
        Case 46578
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)
        
        Case 46579
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)
        
        Case 46580
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 46581
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 46583
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 46584
            Call Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_NAO_PREENCHIDA", gErr)
        
        Case 46585
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)
        
        Case 46586
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)
        
        Case 46591
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PEDIDO_TRAZIDO", gErr)
        
        Case 25739, 46461, 46592, 46596, 46630, 46631, 46630, 46632, 46634, 46637   '  , 46618
        
        Case 46593
            Call Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_ANTERIOR_DATAEMISSAO", gErr, DataSaida, DataEmissao)
        
        Case 46628
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)

        Case 46629
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr, CDbl(PesoLiquido.Text), CDbl(PesoBruto.Text))

        Case 46635
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_BLOQUEIO_CREDITO", gErr, glCodPedido)
        
        Case 46636
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_BLOQUEIO_TOTAL", gErr)
                    
        Case 51533
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)
                    
        Case 56906
            'Call Rotina_Erro(vbOKOnly, "ERRO_IR_FONTE_MAIOR_VALOR_TOTAL", gErr, dValorIRRF, dValorTotal)

        Case 83454
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iTipo)
            
        Case 206615
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_DATA_MENOR_PV", gErr, StrParaDate(DataEmissao.Text), objPedidoVenda.dtDataEmissao)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157496)

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
    If glCodPedido = 0 Then gError 83957

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(Pedido.Text)) = 0 Then gError 83958
    If Len(Trim(FilialPedido.Text)) = 0 Then gError 83959
    If Len(Trim(Cliente.Caption)) = 0 Then gError 83960
    If Len(Trim(Serie.Text)) = 0 Then gError 83961
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 83962

    'Verifica se o código de pedido é o mesmo que foi trazido para a tela
    If glCodPedido <> StrParaLong(Pedido.Text) Or giFilialPedido <> Codigo_Extrai(FilialPedido.Text) Then gError 83963

    objNFiscal.lNumPedidoVenda = glCodPedido
    objNFiscal.iFilialPedido = giFilialPedido
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.iTipoNFiscal = DOCINFO_PADRAO_TELA 'DOCINFO_NFISVPV

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'le o tipo da origem
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 203024
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 203025
    
    'se origem for de empresa é interna
    If objTipoDocInfo.iOrigem = DOCINFO_EMPRESA Then
        'Verifica se a existe nota fiscal está cadastrada
        lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
        If lErro <> SUCESSO And lErro <> 43676 Then gError 203026
        
        If lErro = 43676 Then gError 203027
        
'        If gobjCRFAT.iUsaNFe = MARCADO And objNFiscal.iRecibo = 0 Then
'
'            lErro = CF("NFeFedLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'            If lErro <> SUCESSO Then gError 203028
'
'            If iEnviado = 1 Then gError 203029
'
'        ElseIf gobjCRFAT.iUsaNFSE = MARCADO And objNFiscal.iRecibo <> 0 Then
'
'            lErro = CF("RPSWEBLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'            If lErro <> SUCESSO Then gError 207079
'
'            If iEnviado = 1 Then gError 207080
'
'        End If

    End If

    'Verifica se a existe nota fiscal está cadastrada
    lErro = CF("NFiscal_Le_Pedido", objNFiscal)
    If lErro <> SUCESSO And lErro <> 46602 Then gError 83964
    
    'se a nota não está cadastrada ==> erro
    If lErro = 46602 Then gError 83965

    'pede confirmacao
    vbMsg = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_NFISCAL", objNFiscal.lNumNotaFiscal)
    If vbMsg = vbYes Then

        'Lê os itens da nota fiscal
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 83968

        'Faz o cancelamento de uma nota fiscal de Saida
        lErro = CF("NotaFiscalSaida_Excluir", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError 83966
    
        'Limpa a Tela
        lErro = Limpa_Tela_NFiscalPedido()
        If lErro <> SUCESSO Then gError 83967
        
        'Libera as coleções globais
        Set gcolItemPedido = New Collection
        Set gcolItemQuant = New Collection
        Set gcolItemUM = New Collection
    
        'fecha o comando de setas
        Call ComandoSeta_Fechar(Me.Name)

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 83957
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PEDIDO_TRAZIDO", gErr)
        
        Case 83958
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)
            
        Case 83959
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)
            
        Case 83960
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 83961
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 83962
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 83963
             Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_TELA_DIFERE_BD", gErr, StrParaLong(Pedido.Text), Codigo_Extrai(FilialPedido.Text), glCodPedido, giFilialPedido)
       
        Case 84964, 83966, 83967, 207079
       
        Case 83965
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA2", gErr, objNFiscal.lNumPedidoVenda, objNFiscal.iFilialPedido, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
       
        Case 203025
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)
        
        Case 203027
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
       
        Case 203029
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CADASTRADA_NAO_EXCLUI", gErr)
       
        Case 207080
            Call Rotina_Erro(vbOKOnly, "ERRO_NFSE_CADASTRADA_NAO_EXCLUI", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157497)
        
    End Select

    Exit Sub

End Sub

Public Sub NFiscal_Imprime_Romaneio(objNFiscal As ClassNFiscal)
    Call CF("NFiscal_Imprime_Romaneio", objNFiscal)
End Sub

Private Function Move_Grids_Memoria(objNFiscal As ClassNFiscal) As Long
'recolhe os dados dos grids para a memória(objNFiscal)

Dim lErro As Long

On Error GoTo Erro_Move_Grids_Memoria

    'recolhe os dados do grid de itens
    lErro = Move_GridItens_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 46588
    
    'Chama Move_GridComissoes_Memoria
    lErro = objTabComissoes.Move_TabComissoes_Memoria(objNFiscal, NOTA_FISCAL)
    If lErro <> SUCESSO Then gError 46589
    
    'recolhe os dados do grid de alocações
    lErro = Move_GridAlocacao_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 46590
    
    'mover a parte de tributacao
    lErro = gobjTribTab.PreencherNF(objNFiscal)
    If lErro <> SUCESSO Then gError 46597
        
    Move_Grids_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Grids_Memoria:

    Move_Grids_Memoria = gErr
    
    Select Case gErr
    
        Case 46588, 46589, 46590, 46597
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157498)
    
    End Select
    
    Exit Function

End Function

Public Function Move_GridItens_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os Itens do Grid para a Memória

Dim iIndice As Integer
Dim lErro As Long
Dim objItemNF As ClassItemNF, objItemNFAux2 As ClassItemNF
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lTamanho As Long
Dim colAlocacoes As ColAlocacoesItemNF
Dim objItemNF1 As ClassItemNF
Dim objDataEntrega As ClassDataEntrega
Dim dQuantidade As Double
Dim vbMsg As VbMsgBoxResult
Dim objItemPV As ClassItemPedido
Dim dQtdePV As Double
Dim iItem As Integer

On Error GoTo Erro_Move_GridItens_Memoria

    'Para cada linha existente do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemNF = New ClassItemNF

        'Verifica se o Produto está preenchido
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 46753

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
        
        'Adiciona na coleção de itens
        With objItemNF
            Set objItemNF1 = objNFiscal.ColItensNF.Add(0, iIndice, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricaoItem, 0, 0, 0, 0, gcolItemPedido.Item(iIndice), colAlocacoes, 0, "", .sCcl, STATUS_LANCADO, 0, "", 0, 0, .iControleEstoque, .sUMEstoque, .iClasseUM, 0)
        End With

        Set objItemNF1.colDataEntrega = gobjNFiscal.ColItensNF(iIndice).colDataEntrega
        
        lErro = CF2(Me, "NFP_Move_GridItem_Memoria", objNFiscal, objItemNF1, iIndice)
        If lErro <> SUCESSO Then gError 197872
        
        dQuantidade = 0

        For Each objDataEntrega In objItemNF1.colDataEntrega
            dQuantidade = dQuantidade + objDataEntrega.dQuantidade
        Next
        
        If dQuantidade > 0 Then
            If Abs(dQuantidade - objItemNF1.dQuantidade) > QTDE_ESTOQUE_DELTA Then gError 183243
        End If
        
        Set objNFiscal.ColItensNF(iIndice).objInfoAdicDocItem = gobjNFiscal.ColItensNF(iIndice).objInfoAdicDocItem

        '********************* TRATAMENTO DE GRADE *****************
        Call Move_ItensGrade_Tela(objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade, gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade)

        Set objItemNFAux2 = gobjNFiscal.ColItensNF.Item(iIndice)
        
        objItemNF1.iRecebForaFaixa = objItemNFAux2.iRecebForaFaixa
        objItemNF1.dPercentMaisReceb = objItemNFAux2.dPercentMaisReceb
        objItemNF1.dPercentMenosReceb = objItemNFAux2.dPercentMenosReceb
        
        If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        
             For iItem = 0 To TabPrecoItemNF.ListCount - 1
                
                If GridItens.TextMatrix(iIndice, iGrid_TabPrecoItem_Col) = TabPrecoItemNF.List(iItem) Then
        
                    objItemNF1.iTabelaPreco = TabPrecoItemNF.ItemData(iItem)
                    Exit For
        
                End If
                
            Next
        
        End If
    
        objItemNF1.dComissao = -1
        If gobjFAT.iExibeComissaoItensPV = MARCADO And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col))) <> 0 Then
            objItemNF1.dComissao = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col))
        End If
        
        Set objItemPV = New ClassItemPedido
        
        objItemPV.lNumIntDoc = gcolItemPedido(iIndice)
        objItemPV.iFilialEmpresa = Codigo_Extrai(FilialPedido.Text)
        
        lErro = CF("ItensPedido_Le_NumIntDoc1", objItemPV)
        If lErro <> SUCESSO Then gError 196968
        
        dQtdePV = objItemPV.dQuantidade - objItemPV.dQuantCancelada
        
        If objItemNFAux2.iRecebForaFaixa = ERRO_E_REJEITA_RECEBIMENTO Then
        
            If objItemNF.dQuantidade - ((gcolItemQuant(iIndice) + (dQtdePV * objItemNFAux2.dPercentMaisReceb))) > QTDE_ESTOQUE_DELTA Then gError 196960
    
        End If
        
        If objItemNFAux2.iRecebForaFaixa = AVISA_E_ACEITA_RECEBIMENTO Then
        
            If objItemNF.dQuantidade - ((gcolItemQuant(iIndice) + (dQtdePV * objItemNFAux2.dPercentMaisReceb))) > QTDE_ESTOQUE_DELTA Then
            
                vbMsg = Rotina_Aviso(vbYesNo, "AVISO_QUANTIDADE_ACIMA_FAIXA_PV", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), objItemNF.dQuantidade, ((gcolItemQuant(iIndice) + (dQtdePV * objItemNFAux2.dPercentMaisReceb))))
    
                If vbMsg = vbNo Then gError 196961
    
            End If
    
        End If
        
    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 46753, 46754, 197872
      
        Case 183243
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTSOLICITADA_DIFERE1", gErr, iIndice, GridItens.TextMatrix(iIndice, iGrid_Produto_Col), objItemNF1.dQuantidade, dQuantidade)

        Case 196960
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_MAIOR_SALDO_PV", gErr, GridItens.TextMatrix(iIndice, iGrid_Produto_Col))
        
        Case 196961
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157499)

    End Select

    Exit Function

End Function

Public Function Move_GridAlocacao_Memoria(objNFiscal As ClassNFiscal) As Long

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
            
            'Recolhe a quantidade e o almoxarifado
            objItemAloc.dQuantidade = CDbl(GridAlocacao.TextMatrix(iIndice, iGrid_QuantAloc_Col))
            objItemAloc.sAlmoxarifado = GridAlocacao.TextMatrix(iIndice, iGrid_AlmoxAloc_Col)
            objItemAloc.sUnidadeMed = GridAlocacao.TextMatrix(iIndice, iGrid_UMAloc_Col)
            'Verifica se o almoxarifado já foi lido
            If colCodigoNome.Count > 0 Then
                For Each objCodigoNome In colCodigoNome
                    If objCodigoNome.sNome = objItemAloc.sAlmoxarifado Then
                        objItemAloc.iAlmoxarifado = objCodigoNome.iCodigo
                        Exit For
                    End If
                Next
            End If
            'Se o almoxarifado não foi lido
            If objItemAloc.iAlmoxarifado = 0 Then
                'Lê o almoxarifado
                objAlmoxarifado.sNomeReduzido = objItemAloc.sAlmoxarifado
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25060 Then gError 46756
                If lErro = 25060 Then gError 46757 'Não encontrou
                
                objItemAloc.iAlmoxarifado = objAlmoxarifado.iCodigo
                'Adiciona o almoxarifado na coleção de almoxarifados já lidos
                colCodigoNome.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido
                
            End If
            'Guarda os dados da alocação
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

        Case 46756

        Case 46757
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157500)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 46760

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscalPedido()
    If lErro <> SUCESSO Then gError 500150
    
    Set gcolItemPedido = New Collection
    Set gcolItemQuant = New Collection
    Set gcolItemUM = New Collection

    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)

'    If giTipoVersao = VERSAO_LIGHT Then
'
'        'Prepara para chaveamento para NotaFiscalFatura
'        objNFiscal.iTipoNFiscal = DOCINFO_NFISFV
'
'        'Faz chaveamento de classe CT para NotaFiscalFatura
'        lErro = objUserControl.CT_Chavea(objNFiscal)
'        If lErro <> SUCESSO Then gError 25956
'
'    End If

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 46760, 25956, 500150

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 157501)

    End Select

    Exit Sub

End Sub

Function Carrega_Dados_NFiscal(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_Dados_NFiscal

    'Lê as mensagens
    lErro = CF("NFiscal_Le_Msg", objNFiscal)
    If lErro <> SUCESSO Then gError 46608
    
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 46608
    
    'Lê a parte de Tributação
    lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
    If lErro <> SUCESSO And lErro <> 22867 Then gError 46607
    
    'Lê a parte de tributação dos itens da Nota Fiscal
    lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
    If lErro <> SUCESSO Then gError 46609
    
    'Lê as Comissões da Nota Fiscal
    lErro = CF("ComissoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 46610

    'Lê as Alocações dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 46611
    
    Carrega_Dados_NFiscal = SUCESSO
    
    Exit Function
    
Erro_Carrega_Dados_NFiscal:

    Carrega_Dados_NFiscal = gErr
    
    Select Case gErr
    
        Case 46604, 46605, 46607, 46608, 46609, 46610, 46611
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157502)
            
    End Select
    
    Exit Function

End Function

Function Valida_Dados_GridItens() As Long

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Valida_Dados_GridItens

    'Valida os dados do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 46620

        If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col)) = 0 Then gError 46624
        
    Next
    
    Valida_Dados_GridItens = SUCESSO
    
    Exit Function

Erro_Valida_Dados_GridItens:

    Valida_Dados_GridItens = gErr
    
    Select Case gErr
    
        Case 46620
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 46624
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157503)
    
    End Select
            
    Exit Function

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
    If lErro <> SUCESSO Then gError 46764

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objNFiscal.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "TipoNFiscal", objNFiscal.iTipoNFiscal, 0, "TipoNFiscal"
    colCampoValor.Add "NaturezaOp", objNFiscal.sNaturezaOp, STRING_BUFFER_MAX_TEXTO, "NaturezaOp"
    colCampoValor.Add "Serie", objNFiscal.sSerie, STRING_BUFFER_MAX_TEXTO, "Serie"
    colCampoValor.Add "NumNotaFiscal", objNFiscal.lNumNotaFiscal, 0, "NumNotaFiscal"
    colCampoValor.Add "DataEmissao", objNFiscal.dtDataEmissao, 0, "DataEmissao"
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
    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_BUFFER_MAX_TEXTO, "VolumeNumero"
    colCampoValor.Add "FreteRespons", objNFiscal.iFreteRespons, 0, "FreteRespons"
    colCampoValor.Add "Status", objNFiscal.iStatus, 0, "Status"
    colCampoValor.Add "SemDataSaida", objNFiscal.iSemDataSaida, 0, "SemDataSaida"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "FilialPedido", OP_DIFERENTE, 0
    colSelecao.Add "NumPedidoVenda", OP_DIFERENTE, 0
    colSelecao.Add "TipoNFiscal", OP_IGUAL, DOCINFO_PADRAO_TELA
'    colSelecao.Add "Status", OP_DIFERENTE, STATUS_CANCELADO
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 46764

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157504)

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
        If lErro <> SUCESSO And lErro <> 31442 Then gError 180096
        
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 46765

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 46765, 180096

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157505)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPedido_evSelecao(obj1 As Object)

Dim objPedidoVenda As ClassPedidoDeVenda

    Set objPedidoVenda = obj1

    Pedido.Text = objPedidoVenda.lCodigo
    FilialPedido.Text = objPedidoVenda.iFilialEmpresa
    Call FilialPedido_Validate(bSGECancelDummy)
    
    BotaoPedido_Click
    
    Me.Show

End Sub

Public Sub PedidoLabel_Click()
    
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colSelecao As New Collection

    Call Chama_Tela("PedidoNFiscalLista", colSelecao, objPedidoVenda, objEventoPedido)
    
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
        If GridItens.Row = 0 Then gError 43704

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 83087
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub
    
Erro_BotaoProdutos_Click:

    Select Case gErr
    
        Case 43704
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 83087
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157506)
        
    End Select
    
    Exit Sub
    

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)
    Me.Show
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
                        
            lErro = CF2(Me, "NFP_Move_GridItem_Memoria", gobjNFiscal, objItemNF, iIndice)
            If lErro <> SUCESSO Then gError 197873
                        
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
    Call Chama_Tela_Modal("LocalizacaoProduto", ColItensNF, DOCINFO_NFISVPV)
    
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
    Call CF2(Me, "NFP_Preenche_GridAlocacoes", objNFiscal)
    
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

        Case 46885, 46886, 46888, 86306, 86378, 177582, 197873
        
        Case 46889
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 86307
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, sProduto)
        
        Case 86376
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProduto)

        Case 86379
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objItemRomaneio.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157507)

    End Select

    Exit Sub

End Sub

Public Function Recolhe_Alocacoes(iItem As Integer, colAlocacoes As ColAlocacoesItemNF) As Long 'alterado p/ Public por Leo em 12/01/02
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
                If lErro <> 25060 And lErro <> SUCESSO Then gError 46905
                
                'Se não encontrou o Almoxarifado --> erro
                If lErro = 25060 Then gError 46906
                            
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

        Case 46905

        Case 46906
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157508)

    End Select

    Exit Function

End Function

Public Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objNFiscal Is Nothing) Then

        lErro = CargaPosFormLoad(True)
        If lErro <> SUCESSO Then gError 64007

        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 60183
        
        If lErro <> SUCESSO Then gError 60184

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

        'Lê o Tipo de Documento com o Código de objNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 60185
        
        'Se não encontrar --> Erro
        If lErro = 31415 Then gError 60186
        
'        If giTipoVersao = VERSAO_FULL Then
            If objTipoDocInfo.sNomeTelaNFiscal <> NOME_TELA_NFISCALPEDIDO Then gError 60187
'        End If
        
        'Traz a nota para a tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 60188

    End If

    iAlterado = 0
    iComissoesAlterada = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 60183, 60185, 60188, 64007

        Case 60184
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case 60186
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 60187
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_DIFERENTE_NF_PEDIDO", gErr, objNFiscal.iTipoNFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157509)

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

    Parent.HelpContextID = IDH_NF_PEDIDO_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Nota Fiscal Gerada a Partir de um Pedido"
    Call Form_Load
    
End Function

Public Function Name() As String
    Name = "NFiscalPedido"
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

Public Function DescontoGlobal_Testa(dValorDesconto As Double) As Long

Dim dValorProdutos As Double
    
Dim lErro As Long

On Error GoTo Erro_DescontoGlobal_Testa

    Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)

    'Verifica se o Valor do Desconto é maior que a soma dos demais valores
    If dValorDesconto > dValorProdutos Then gError 56895
    
    DescontoGlobal_Testa = SUCESSO
     
    Exit Function
    
Erro_DescontoGlobal_Testa:

    DescontoGlobal_Testa = gErr
     
    Select Case gErr
                  
        Case 56895
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157510)
     
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
        Else
            Call gobjRastreamento.objUserControl_KeyDown(KeyCode, Shift)
        End If
    
    End If

End Sub

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
Dim objProdutoCategoria As New ClassProdutoCategoria
Dim objCatProdItem As ClassCategoriaProdutoItem
Dim objPlanoConta As New ClassPlanoConta
Dim objProduto As New ClassProduto
Dim objFilial As New ClassFilialCliente
Dim sContaTela As String
Dim objTribItemNFAux As ClassTribItemNFAux, dValorAux As Double

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
                    If lErro <> SUCESSO Then gError 79015
                    
                    objItemMovEstoque.sProduto = sProdutoFormatado
                    
                    'Define o almoxarifado que será passado como parâmetro para MovEstoque_Le_QuantVendConsig
                    objAlmoxarifado.sNomeReduzido = GridAlocacao.TextMatrix(iLinha, iGrid_AlmoxAloc_Col)
                    
                    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25060 Then gError 79014
                    
                    objItemMovEstoque.iAlmoxarifado = objAlmoxarifado.iCodigo
                    
                    'Define o tipo de movimento, o DocOrigem e o TipoNumIntoDocOrigem que serão passados como parâmetros para MovEstoque_Le_QuantVendConsig
                                        
                    objItemMovEstoque.iTipoMov = MOV_EST_NF_VENDA_MAT_CONSIG
                    objItemMovEstoque.sDocOrigem = "NFISVPV" & " " & Serie.Text & " " & objContabil.lNumNotaFiscalExt
                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                                                           
                    'Verifica se MovEstoque_Le_QuantVendConsig não encontrou erro
                    lErro = CF("MovEstoque_Le_QuantVendConsig", objItemMovEstoque)
                    If lErro <> SUCESSO And lErro <> 79003 Then gError 79013
                
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
                        If lErro <> SUCESSO And lErro <> 25060 Then gError 64215

                        'Se não encontrou ===> Erro
                        If lErro = SUCESSO Then

                            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
                            
                            lErro = CF("Produto_Formata", objMnemonicoValor.vParam(1), sProdutoFormatado, iPreenchido)
                            If iPreenchido <> PRODUTO_PREENCHIDO Then gError 64216

                            objEstoqueProduto.sProduto = sProdutoFormatado
                            
                            'Lê a conta contábil do Produto no Almoxarifado
                            lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
                            If lErro <> SUCESSO And lErro <> 49991 Then gError 64217

                            If Len(Trim(objEstoqueProduto.sContaContabil)) > 0 Then

                                sContaMascarada = String(STRING_CONTA, 0)

                                lErro = Mascara_MascararConta(objEstoqueProduto.sContaContabil, sContaMascarada)
                                If lErro <> SUCESSO Then gError 64218

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
                    If lErro <> SUCESSO Then gError 64213

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
                If lErro <> SUCESSO Then gError 88030
                
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
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dPISCredito
        
        Case COFINS_DEBITO_TOTAL
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dCOFINSCredito
            
        Case PIS_VALOR_TOTAL
            dValorAux = 0
            For Each objTribItemNFAux In gobjTribTab.objNFTributacao.colItens
                dValorAux = dValorAux + objTribItemNFAux.dPISValor
            Next
            objMnemonicoValor.colValor.Add Arredonda_Moeda(dValorAux)
        
        Case COFINS_VALOR_TOTAL
            dValorAux = 0
            For Each objTribItemNFAux In gobjTribTab.objNFTributacao.colItens
                dValorAux = dValorAux + objTribItemNFAux.dCOFINSValor
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
                    If lErro <> SUCESSO Then gError 183777
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183778
        
                    sContaMascarada = String(STRING_CONTA, 0)

                    lErro = Mascara_MascararConta(objProduto.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError 183800

                    objMnemonicoValor.colValor.Add sContaMascarada
                    
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case CONTACONTABILCATPROD
                
            For iLinha = 1 To objGridItens.iLinhasExistentes
                
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183779
                
                    objProdutoCategoria.sCategoria = objMnemonicoValor.vParam(1)
                    
                    objProdutoCategoria.sProduto = sProdutoFormatado
                    
                    Set objCatProdItem = New ClassCategoriaProdutoItem
                    
                    lErro = CF("ProdutoCategoria_Le_Item", objProdutoCategoria, objCatProdItem)
                    If lErro <> SUCESSO And lErro <> 183548 Then gError 183780
                    
                    If lErro = SUCESSO Then
                        
                        Set objPlanoConta = New ClassPlanoConta
                    
                        'le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta
                        lErro = CF("PlanoConta_Le_ContaSimples", objCatProdItem.dvalor1, objPlanoConta)
                        If lErro <> SUCESSO And lErro <> 5451 Then gError 183781
                    
                        If lErro = SUCESSO Then
                    
                            sContaMascarada = String(STRING_CONTA, 0)
        
                            lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
                            If lErro <> SUCESSO Then gError 183801
        
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
        
        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then gError 186140
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 186141
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
        
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
        
        Case 64213, 64215, 64216, 64217, 64218, 64219, 64220, 64221, 64306, 79013, 79014, 79015, 88030, 183777 To 183781, 183800, 183801, 186140, 186141
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157511)

    End Select

    Exit Function

End Function

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

Public Sub VolumeQuant_GotFocus()
    Call MaskEdBox_TrataGotFocus(VolumeQuant, iAlterado)
End Sub

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
    
    'Traz a mensagem NFiscal para a tela
    Mensagem.Text = objMensagem.sDescricao
    
    Me.Show

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

Public Sub AlocacoesNF_Agrupa(colAlocacoes As ColAlocacoesItemNF)

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
    If lErro <> SUCESSO Then gError 79827
    
    'Guarda no objNFiscal os dados necessários para consultar o título
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.iTipoDocInfo = DOCINFO_PADRAO_TELA 'DOCINFO_NFISVPV
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Caption)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.dtDataEntrada = DATA_NULA
        
    'Guarda em objFornecedor o nome reduzido do Fornecedor
    objCliente.sNomeReduzido = Cliente.Caption
    
    'Lê o código do Fornecedor a partir do nome reduzido obtido na tela
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 79828
    
    'Se não encontrou o fornecedor = > erro
    If lErro = 12348 Then gError 79829
    
    'Guarda no objNFiscal o código do fornecedor
    objNFiscal.lCliente = objCliente.lCodigo
    
    'Lê o NumIntDocCPR da NFiscal e exibe o documento gerado no CPR por essa nota
    lErro = CF("NFiscal_Consulta_DocCPR", objNFiscal)
    If lErro <> SUCESSO And lErro <> 79717 Then gError 79830
    
    'Se não encontrou a nota => erro
    If lErro = 79717 Then gError 79831
    
    Exit Sub
    
Erro_BotaoConsultaTitRec_Click:

    Select Case gErr
    
        Case 79827, 79830, 79828
        
        Case 79831
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA2", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iTipoNFiscal, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEmissao)
        
        Case 79829
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157512)
    
    End Select
    
    Exit Sub

End Sub

Private Function Critica_CamposNecessarios_ConsultaTitulo() As Long
'Verifica se os campos necessários para encontrar consultar um título gerado por uma NFFatEntrada foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposNecessarios_ConsultaTitulo

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.Caption)) = 0 Then gError 79832
    
    'Verifica se a filial do Fornecedor foi preenchida
    If Len(Trim(Filial.Caption)) = 0 Then gError 79833
    
    'Verifica se a Série foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 79834
    
    'Verifica se o Número da Nota foi preenchido
    If Len(Trim(NFiscal.Caption)) = 0 Then gError 79835
    
    'Verifica se a data de emissão da nota foi preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 79836
    
    Critica_CamposNecessarios_ConsultaTitulo = SUCESSO
    
    Exit Function

Erro_Critica_CamposNecessarios_ConsultaTitulo:

    Critica_CamposNecessarios_ConsultaTitulo = gErr
    
    Select Case gErr
    
        Case 79832
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 79833
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 79834
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 79835
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
        
        Case 79836
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157513)
    
    End Select
        
    Exit Function
        
End Function

Public Property Let sName(ByVal vData As String)
    mvarsName = vData
End Property

Public Property Get sName() As String
    sName = mvarsName
End Property

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19243

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 19244
        
        'Não encontrou Transportadora no BD
        If lErro <> SUCESSO Then Error 19245
        
        'Encontrou Transportadora no BD, coloca no Text da Combo
        TranspRedespacho.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 19246

    Exit Sub

Erro_TranspRedespacho_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 19243, 19244

        Case 19245  'Não encontrou Transportadora no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            End If

        Case 19246
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", Err, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157514)

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
                lErro = ValorTotal_Calcula()
                
                lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
                If lErro <> SUCESSO Then gError 86360
                
                If lErro <> SUCESSO Then gError 51037
            End If
        
        End If
    
    End If
    
    Exit Sub

Erro_BotaoGrade_Click:

    Select Case gErr
    
        Case 51037, 86360
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157515)
            
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157516)
            
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157517)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157518)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157519)
     
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

'Public Function BotaoEntrega_Click() As Long
'
'Dim sProdutoTela As String
'Dim dQuantidade As Double
'
'On Error GoTo Erro_BotaoEntrega_Click
'
'    If GridItens.Row = 0 Then gError 183217
'
'    sProdutoTela = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
'
'    If Len(sProdutoTela) = 0 Then gError 183218
'
'    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
'
'    Call Chama_Tela_Modal("DataEntrega", sProdutoTela, dQuantidade, gobjNFiscal.ColItensNF(GridItens.Row).colDataEntrega)
'
'    BotaoEntrega_Click = SUCESSO
'
'    Exit Function
'
'Erro_BotaoEntrega_Click:
'
'    BotaoEntrega_Click = gErr
'
'    Select Case gErr
'
'        Case 183217
'            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
'
'        Case 183218
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183219)
'
'    End Select
'
'    Exit Function
'
'End Function

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

Public Function NFP_Form_Load2(ByVal objCT As Object) As Long
    NFP_Form_Load2 = SUCESSO
End Function

Public Function NFP_Transfere_Dados_Pedido_NFiscal(ByVal objCT As Object, ByVal objItemNF As ClassItemNF, ByVal objItemPV As ClassItemPedido, ByVal dQuantidade As Double, ByVal objNFiscal As ClassNFiscal) As Long
    NFP_Transfere_Dados_Pedido_NFiscal = SUCESSO
End Function

Function NFP_Move_GridItem_Memoria(ByVal objCT As Object, objNFiscal As ClassNFiscal, objItemNF As ClassItemNF, iIndice As Integer) As Long
    NFP_Move_GridItem_Memoria = SUCESSO
End Function

Function NFP_Traz_NFiscal_Tela2(ByVal objCT As Object, objNFiscal As ClassNFiscal, lNumIntPedido As Long) As Long
    NFP_Traz_NFiscal_Tela2 = SUCESSO
End Function

Public Function BotaoEntrega_Click() As Long
    Call CF2(Me, "NFP_BotaoEntrega_Click")
End Function

Public Function NFP_Quantidade_GotFocus(ByVal objCT As Object) As Long
    NFP_Quantidade_GotFocus = SUCESSO
End Function

Public Function NFP_BotaoEntrega_Click(ByVal objCT As Object) As Long

Dim lErro As Long
Dim sProdutoTela As String
Dim dQuantidade As Double

On Error GoTo Erro_NFP_BotaoEntrega_Click

    If GridItens.Row = 0 Then gError 183236

    sProdutoTela = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    If Len(sProdutoTela) = 0 Then gError 183237

    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))

    Call Chama_Tela_Modal("DataEntrega", sProdutoTela, dQuantidade, gobjNFiscal.ColItensNF(GridItens.Row).colDataEntrega)

    NFP_BotaoEntrega_Click = SUCESSO

    Exit Function

Erro_NFP_BotaoEntrega_Click:

    NFP_BotaoEntrega_Click = gErr

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

Public Sub MsgAutomatica_Click()
     Call Calcula_Msg
End Sub

Public Function Calcula_Msg() As Long

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
'Carrega TipoFrete

Dim lErro As Long
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer
Dim colTipoFrete As Collection

On Error GoTo Erro_Carrega_TipoFrete
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207437

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
        Case 207437

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207438)

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

