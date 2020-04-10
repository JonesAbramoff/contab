VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscal"
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

Public gobjInfoUsu As Object 'Inserido por Wagner

Private WithEvents objEventoOrcamento As AdmEvento
Attribute objEventoOrcamento.VB_VarHelpID = -1

'estrutura mantida para auxiliar a manutencao de informacoes p/tributacao
Public gobjNFiscal As ClassNFiscal
Dim giFilialAlterada As Integer
Public giValorDescontoManual As Integer 'Alterado por Wagner
Dim giFilialEntregaManual As Integer
Dim giClienteAlterado As Integer
Dim giComplemento As Integer

Dim bProcessaReserva As Boolean

Dim m_objUserControl As Object
Dim FrameL(1 To 7) As Frame

'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjTelaProjetoInfo As ClassTelaPRJInfo

Private Const NOME_TELA_NFISCAL = "NFiscal"

'Flag que indica se a tela está sendo preenchida.
Public gbCarregandoTela As Boolean

'Indica que está limpando a tela
Public gbLimpandoTela As Boolean

Public giPosCargaOk As Integer

'inicio contabilidade

Public objGrid1 As AdmGrid
Public objContabil As New ClassContabil

'Alterado por Tulio em 27/03
'declara o objeto da execucao das regras de comissoes e
'o objeto do calculo de comissoes
Public objComissoesRegrasCalc As Object
Public objMnemonicoComissCalc As ClassMnemonicoComissCalc
Public objMnemonicoComissCalcAux As ClassMnemonicoComissCalcAux

'*** 12/04/02 - Tulio ***
'Indica que foi alterado pelo menos um campo na tela que seja utilizado para calcular as comissões
Public iComissoesAlterada As Integer

Public WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Public WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

'mnemonicos
Private Const TIPO_NOTA_FISCAL As String = "Tipo_Nota_Fiscal"
Private Const NATUREZA_OP As String = "Natureza_OP"
Private Const CLIENTE1 As String = "Cliente"
Private Const CLIENTE_CODIGO As String = "Cliente_Codigo"
Private Const FILIAL1 As String = "Filial_Cliente"
Private Const FILIAL_CONTA As String = "FilialCli_Conta_Ctb"
Private Const Serie1 As String = "Serie"
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
Private Const IPI As String = "IPI_Valor"
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
Private Const CONTACONTABILEST1 As String = "ContaContabilEst"
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
Private Const PIS_VALOR_TOTAL As String = "PIS_Valor_Total"
Private Const COFINS_VALOR_TOTAL As String = "COFINS_Valor_Total"
'fim da contabilidade

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
Public iGrid_TabPrecoItem_Col As Integer
Public iGrid_ComissaoItem_Col As Integer

Public iGrid_Item_Col As Integer
Public iGrid_ProdutoAloc_Col As Integer
Public iGrid_QuantAloc_Col As Integer
Public iGrid_QuantVend_Col As Integer
Public iGrid_AlmoxAloc_Col As Integer
Public iGrid_UMAloc_Col As Integer

Public iFrameAtual As Integer
Public iAlterado As Integer
Public iClienteAlterado As Integer
Public gdDesconto As Double

Public iValorDescontoAlterado As Integer
Public iValorFreteAlterado As Integer
Public iValorSeguroAlterado As Integer
Public iValorDespesasAlterado As Integer
Public giFilialEntregaAlterada As Integer
Public iDataEmissaoAlterada As Integer
Public iSerieAlterada As Integer

Public objGridItens As AdmGrid
Public objGridComissoes As AdmGrid
Public objGridAlocacoes As AdmGrid

'inicia objeto associado a GridComissoes
Public objTabComissoes As New ClassTabComissoes

Private WithEvents objEventoEstoqueProd As AdmEvento
Attribute objEventoEstoqueProd.VB_VarHelpID = -1
Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1
Public WithEvents objEventoNFiscal As AdmEvento
Attribute objEventoNFiscal.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoTranspRedespacho As AdmEvento
Attribute objEventoTranspRedespacho.VB_VarHelpID = -1
Private WithEvents objEventoSerieOriginal As AdmEvento
Attribute objEventoSerieOriginal.VB_VarHelpID = -1
Private WithEvents objEventoNFiscalOriginal As AdmEvento
Attribute objEventoNFiscalOriginal.VB_VarHelpID = -1
Public WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoMensagem As AdmEvento
Attribute objEventoMensagem.VB_VarHelpID = -1
Private WithEvents objEventoCliIntermediario As AdmEvento
Attribute objEventoCliIntermediario.VB_VarHelpID = -1

''*** incluidos p/tratamento de tributacao *******************************
Public gobjTribTab As New ClassTribTab
Public gobjRastreamento As New ClassRastreamento

'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_Itens = 2
Private Const TAB_Complemento = 3
Private Const TAB_Comissoes = 4
Private Const TAB_Almoxarifado = 5
Private Const TAB_Tributacao = 6
Private Const TAB_Contabilizacao = 7
Private Const TAB_RASTRO = 8

'--- inicio dos properties get dos controles da tela

Public Property Get LabelCliTomador() As Object
     Set LabelCliTomador = objUserControl.Controls("LabelCliTomador")
End Property

Public Property Get LabelCliInt() As Object
     Set LabelCliInt = objUserControl.Controls("LabelCliInt")
End Property

Public Property Get FrameCliInt() As Object
     Set FrameCliInt = objUserControl.Controls("FrameCliInt")
End Property

Public Property Get CliIntermediario() As Object
     Set CliIntermediario = objUserControl.Controls("CliIntermediario")
End Property

Public Property Get FilialCliIntermediario() As Object
     Set FilialCliIntermediario = objUserControl.Controls("FilialCliIntermediario")
End Property

Public Property Get FrameOrcVenda() As Object
     Set FrameOrcVenda = objUserControl.Controls("FrameOrcVenda")
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

Public Property Get DiretoIndireto() As Object
     Set DiretoIndireto = objUserControl.Controls("DiretoIndireto")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get BotaoLocalizacao() As Object
     Set BotaoLocalizacao = objUserControl.Controls("BotaoLocalizacao")
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

Public Property Get ComissaoAutomatica() As Object
     Set ComissaoAutomatica = objUserControl.Controls("ComissaoAutomatica")
End Property

Public Property Get BotaoVendedores() As Object
     Set BotaoVendedores = objUserControl.Controls("BotaoVendedores")
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

Public Property Get TotalPercentualComissao() As Object
     Set TotalPercentualComissao = objUserControl.Controls("TotalPercentualComissao")
End Property

Public Property Get TotalValorBase() As Object
     Set TotalValorBase = objUserControl.Controls("TotalValorBase")
End Property

Public Property Get TotalValorComissao() As Object
     Set TotalValorComissao = objUserControl.Controls("TotalValorComissao")
End Property

Public Property Get LabelTotaisComissoes() As Object
     Set LabelTotaisComissoes = objUserControl.Controls("LabelTotaisComissoes")
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

Public Property Get FrameNFiscalOriginal() As Object
     Set FrameNFiscalOriginal = objUserControl.Controls("FrameNFiscalOriginal")
End Property

Public Property Get SerieNFiscalOriginal() As Object
     Set SerieNFiscalOriginal = objUserControl.Controls("SerieNFiscalOriginal")
End Property

Public Property Get NFiscalOriginal() As Object
     Set NFiscalOriginal = objUserControl.Controls("NFiscalOriginal")
End Property

Public Property Get NFiscalOriginalLabel() As Object
     Set NFiscalOriginalLabel = objUserControl.Controls("NFiscalOriginalLabel")
End Property

Public Property Get SerieNFOriginalLabel() As Object
     Set SerieNFOriginalLabel = objUserControl.Controls("SerieNFOriginalLabel")
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

Public Property Get ValorTotal() As Object
     Set ValorTotal = objUserControl.Controls("ValorTotal")
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

Public Property Get ValorProdutos() As Object
     Set ValorProdutos = objUserControl.Controls("ValorProdutos")
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

Public Property Get UnidadeMed() As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
End Property

Public Property Get DescricaoItem() As Object
     Set DescricaoItem = objUserControl.Controls("DescricaoItem")
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

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
End Property

Public Property Get TabelaPreco() As Object
     Set TabelaPreco = objUserControl.Controls("TabelaPreco")
End Property

Public Property Get BotaoLimparNF() As Object
     Set BotaoLimparNF = objUserControl.Controls("BotaoLimparNF")
End Property

Public Property Get Serie() As Object
     Set Serie = objUserControl.Controls("Serie")
End Property

Public Property Get TipoNFiscal() As Object
     Set TipoNFiscal = objUserControl.Controls("TipoNFiscal")
End Property

Public Property Get NFiscal() As Object
     Set NFiscal = objUserControl.Controls("NFiscal")
End Property

Public Property Get SerieLabel() As Object
     Set SerieLabel = objUserControl.Controls("SerieLabel")
End Property

Public Property Get NFiscalLabel() As Object
     Set NFiscalLabel = objUserControl.Controls("NFiscalLabel")
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

Public Property Get CTBLabelCcl() As Object
     Set CTBLabelCcl = objUserControl.Controls("CTBLabelCcl")
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

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

Public Property Get Status() As Object
    Set Status = objUserControl.Controls("Status")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

'Rastreamento
Public Property Get GridRastro() As Object
     Set GridRastro = objUserControl.Controls("GridRastro")
End Property

'Rastreamento
Public Property Get EscaninhoRastro() As Object
     Set EscaninhoRastro = objUserControl.Controls("EscaninhoRastro")
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

Public Property Get NFiscalAux() As Object
     Set NFiscalAux = objUserControl.Controls("NFiscalAux")
End Property
'--- fim dos properties get dos controles da tela

Public Function CargaPosFormLoad(Optional bTrazendoDoc As Boolean = False) As Long

Dim iSubTipo As Integer
Dim lErro As Long

On Error GoTo Erro_CargaPosFormLoad

    If (giPosCargaOk = 0) Then

        'Rastreamento
        Set gobjRastreamento = New ClassRastreamento
        Set gobjRastreamento.objTela = Me
        gobjRastreamento.bSaidaMaterial = True

        'p/permitir o redesenho da tela
        DoEvents
        
        gbCarregandoTela = True
        
        'Alterado por Tulio em 27/03

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
 
        Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
        Set gobjTelaProjetoInfo.objUserControl = objUserControl
        Set gobjTelaProjetoInfo.objTela = Me
 
        ''*** incluidos p/tratamento de tributacao *******************************
        Call GridItens_InicializaColunas
        Set gobjTribTab = New ClassTribTab
        lErro = gobjTribTab.Ativar(Me)
        If lErro <> SUCESSO Then gError 59736
        '*** fim tributacao
        
        'Carrega os Tipos de Notas Fiscais
        lErro = Carrega_TiposDocInfo()
        If lErro <> SUCESSO Then gError 35685
    
        'Carrega as Tabelas de Preço
        lErro = Carrega_TabelaPreco()
        If lErro <> SUCESSO Then gError 35686
    
        'Carrega as Transportadoras
        lErro = Carrega_Transportadora()
        If lErro <> SUCESSO Then gError 35687
    
        'Carrega a combo VolumeEspecie
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
        If lErro <> SUCESSO Then gError 102414
    
        'Carrega a combo VolumeMarca
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
        If lErro <> SUCESSO Then gError 102415
    
        'Carrega os Canais de Venda
        lErro = Carrega_CanalVenda()
        If lErro <> SUCESSO Then gError 35688
    
        'Carrega as Séries
        lErro = Carrega_Serie()
        If lErro <> SUCESSO Then gError 35697
        
        'Seleciona Série Padrão
        lErro = Seleciona_Serie_Padrao()
        If lErro <> SUCESSO Then gError 500113
        
        'Carrega os Estados
        lErro = Carrega_PlacaUF()
        If lErro <> SUCESSO Then gError 59733
    
        PrecoUnitario.Format = gobjFAT.sFormatoPrecoUnitario
        
        Quantidade.Format = FORMATO_ESTOQUE
        QuantAlocada.Format = FORMATO_ESTOQUE
        QuantVendida.Format = FORMATO_ESTOQUE
    
        'Seta as Variáveis de Grid e das Telas de browse
        Set objEventoMensagem = New AdmEvento
        Set objGridItens = New AdmGrid
        Set objGridComissoes = New AdmGrid
        Set objGridAlocacoes = New AdmGrid
        Set objEventoSerie = New AdmEvento
        Set objEventoCliente = New AdmEvento
        Set objEventoNFiscal = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoVendedor = New AdmEvento
        Set objEventoNatureza = New AdmEvento
        Set objEventoEstoqueProd = New AdmEvento
        Set objEventoSerieOriginal = New AdmEvento
        Set objEventoTransportadora = New AdmEvento
        Set objEventoTranspRedespacho = New AdmEvento
        Set objEventoNFiscalOriginal = New AdmEvento
        Set objEventoCliIntermediario = New AdmEvento
        
        Set objEventoOrcamento = New AdmEvento
        
        'Inicializa a máscara do Produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
        If lErro <> SUCESSO Then gError 35690
        
        'Inicializa o Grid de itens
        lErro = Inicializa_Grid_Itens(objGridItens)
        If lErro <> SUCESSO Then gError 35689
   
        'Inicializa o grid de comissões
        lErro = objTabComissoes.Inicializa_Grid_Comissoes(objGridComissoes)
        If lErro <> SUCESSO Then gError 35691
    
        'Inicializa o Grid de Alocações
        lErro = Inicializa_Grid_Alocacoes(objGridAlocacoes)
        If lErro <> SUCESSO Then gError 35692
        
        'Inicializa a contabilidade carregando apenas os modelos contábeis referentes ao TipoDocInfo default
        If objContabil.objGridInt Is Nothing Then

            'Obtém o código do subtipo contábil referente ao TipoDocInfo default
            lErro = CF("Obtem_Subtipo_Contabil", "TiposDocInfo", DOCINFO_NFISV, iSubTipo)
            If lErro <> SUCESSO And lErro <> 79671 Then gError 79684
            
            'Se não encontrou o subtipo => erro
            If lErro = 79671 Then gError 79685
            
            'Inicialização da parte de contabilidade
            lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_FATURAMENTO, iSubTipo)
            If lErro <> SUCESSO Then gError 79332

        End If
        
        lErro = CF2(Me, "NF_CargaPosFormLoad")
        If lErro <> SUCESSO Then gError 184456
        
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
          
        Case 59736, 35685, 35686, 35687, 35688, 35689, 35690, 35691, 35692, 35697, 39660, 59733, 79332, 500113, 79684, 83323, 102414, 102415, 184456
        
        Case 79685
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPODOCINFO_NAO_ENCONTRADO", gErr, DOCINFO_NFISV)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156172)
     
    End Select
     
    Exit Function

End Function

Public Sub BotaoLimparNF_Click()

    NFiscal.Caption = ""
    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub ComissaoAutomatica_Click()

Dim lErro As Long

On Error GoTo Erro_ComissaoAutomatica_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Se a comissão automática estiver selecionada
    If ComissaoAutomatica.Value = vbChecked Then
        
        'Recalcula as comissoes
        lErro = objTabComissoes.Comissoes_Calcula()
        If lErro <> SUCESSO Then gError 51607
        
    End If
    
    Exit Sub

Erro_ComissaoAutomatica_Click:

    Select Case gErr
    
        Case 51607
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156173)
            
    End Select
    
    Exit Sub
    
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

Public Sub NFiscalOriginal_GotFocus()
    Call MaskEdBox_TrataGotFocus(NFiscalOriginal, iAlterado)
End Sub

Public Sub SerieNFOriginalLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'Recolhe a Série da Nota Fiscal Original da tela
    objSerie.sSerie = SerieNFiscalOriginal.Text

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerieOriginal)

    Exit Sub

End Sub

Public Sub Form_Activate()

Dim lErro As Long

On Error GoTo Erro_Form_Activate

    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError 59333
        
    Call TelaIndice_Preenche(Me)

    Exit Sub
     
Erro_Form_Activate:

    Select Case gErr
          
        Case 59333
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156174)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    giPosCargaOk = 0
    
    iFrameAtual = 1
    
    Set objTabComissoes.objTela = Me
    
    If giTipoVersao = VERSAO_LIGHT Then
        
        FrameOrcVenda.Visible = False
        BotaoGrade.Visible = False
        
        'sumir com abas de contabilidade e rastro
        TabStrip1.Tabs.Remove (TAB_RASTRO)
        TabStrip1.Tabs.Remove (TAB_Contabilizacao)
        
    End If
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    Call CF2(Me, "NFCF_Default_DataSaida")
    
    Call CF2(Me, "NF_Form_Load")
    
    Call Carrega_FilialOrcamento
        
    iAlterado = 0
    iComissoesAlterada = 0
    
    Set gobjNFiscal = New ClassNFiscal
    
    MsgAutomatica.Value = vbChecked
    Mensagem.MaxLength = MENSAGEM_NF_MAXLENGTH
    MensagemCorpo.MaxLength = MENSAGEM_NF_MAXLENGTH
    
    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207417
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 59736, 35685, 35686, 35687, 35688, 35689, 35690, 35691, 35692, 35697, 39660, 59733, 207417

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156175)

    End Select
    
    iAlterado = 0
    iComissoesAlterada = 0
    
    Exit Sub

End Sub

Public Function Form_Load_Light()

Dim lErro As Long

On Error GoTo Erro_Form_Load_Light

    'obs.: a carga pos form_load será feita pela classe que trata nffat
    giPosCargaOk = 1
    
    'Inicializa as Variáveis de browse
    Set objEventoMensagem = New AdmEvento
    Set objEventoSerie = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoNFiscal = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    Set objEventoEstoqueProd = New AdmEvento
    Set objEventoSerieOriginal = New AdmEvento
    Set objEventoTransportadora = New AdmEvento
    Set objEventoNFiscalOriginal = New AdmEvento
    Set objEventoCliIntermediario = New AdmEvento

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
    
    'Colunos do Grid Alocacoes
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

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156176)

    End Select
    
    iAlterado = 0
    iComissoesAlterada = 0
    
    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera as variáveis globais da tela
    Set gobjNFiscal = Nothing
    Set objEventoMensagem = Nothing
    Set objEventoSerie = Nothing
    Set objEventoCliente = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoNatureza = Nothing
    Set objEventoEstoqueProd = Nothing
    Set objEventoSerieOriginal = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoTranspRedespacho = Nothing
    Set objEventoNFiscalOriginal = Nothing
    Set objEventoProduto = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoCliIntermediario = Nothing
    
    Set gobjTelaProjetoInfo = Nothing
    
    Call CF2(Me, "NF_Form_UnLoad")
    
    'se foi instanciada a classe de execucao de regras de comissoes => libera
    Set objComissoesRegrasCalc = Nothing
        
    'se foi instanciada a classe que calcula os mnemonicos => libera
    Set objMnemonicoComissCalc = Nothing
    Set objMnemonicoComissCalcAux = Nothing
    
    Set objGridItens = Nothing
    Set objGridComissoes = Nothing
    Set objGridAlocacoes = Nothing
    Set gobjInfoAdic = Nothing
    
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
    Set gobjRastreamento = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    'Fecha o Comando de Setas
    lErro = ComandoSeta_Liberar(Me.Name)

    Set gobjAnotacao = Nothing
    
''*** incluidos p/tratamento de tributacao *******************************
    If Not (gobjTribTab Is Nothing) Then Call gobjTribTab.Desativar
'*** fim tributacao

    Set objTabComissoes = Nothing
    Set gobjTribTab = Nothing
    
    Set gobjInfoUsu = Nothing 'Inserido por Wagner
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156177)
    
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

Private Function Carrega_TiposDocInfo() As Long
'Carrega os Tipo de Documentos relacionados com a tela corrente

Dim lErro As Long
Dim colTipoDocInfo As New colTipoDocInfo
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iIndice As Integer

On Error GoTo Erro_Carrega_TiposDocInfo

    Set colTipoDocInfo = gobjCRFAT.colTiposDocInfo

    iIndice = 0
    
    'Carrega na combo só os Tipos ligados essa tela
    For Each objTipoDocInfo In colTipoDocInfo
        If objTipoDocInfo.sNomeTelaNFiscal = NOME_TELA_NFISCAL Then
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156178)

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
    If lErro <> SUCESSO Then gError 35694

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

        Case 35694

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156179)

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
    If lErro <> SUCESSO Then gError 35695

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

        Case 35695

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156180)

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
    If lErro <> SUCESSO Then gError 35696

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

        Case 35696

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156181)

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
    If lErro <> SUCESSO Then gError 35698

    'Carrega na combo
    For Each objSerie In colSerie
        If UCase(sSerieAnt) <> UCase(Desconverte_Serie_Eletronica(objSerie.sSerie)) Then Serie.AddItem Desconverte_Serie_Eletronica(objSerie.sSerie)
        SerieNFiscalOriginal.AddItem objSerie.sSerie
        sSerieAnt = Desconverte_Serie_Eletronica(objSerie.sSerie)
    Next
            
    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case 35698

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156182)

    End Select

    Exit Function

End Function

Function Seleciona_Serie_Padrao() As Long

Dim lErro As Long
Dim sSeriePadrao As String

On Error GoTo Erro_Seleciona_Serie_Padrao
        
    'Lê série Padrão
    lErro = CF("Serie_Le_Padrao", sSeriePadrao)
    If lErro <> SUCESSO Then gError 500114
    
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
    
        Case 500114
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156183)
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156184)

    End Select

End Function

Sub GridItens_InicializaColunas()

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

    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        iGrid_ComissaoItem_Col = 11 - iColunasEscondidas
    Else
        iColunasEscondidas = iColunasEscondidas + 1
    End If

End Sub

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens
    Inicializa_Grid_Itens = CF2(Me, "NF_Inicializa_Grid_Itens", objGridInt)
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
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Alocacoes = SUCESSO

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
    sTabela = "NFiscalVenda"

    'Lê os dados da Tela NFiscal
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 35699

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
    colCampoValor.Add "VolumeEspecie", objNFiscal.lVolumeEspecie, 0, "VolumeEspecie"
    colCampoValor.Add "VolumeMarca", objNFiscal.lVolumeMarca, 0, "VolumeMarca"
    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_BUFFER_MAX_TEXTO, "VolumeNumero"
    colCampoValor.Add "FreteRespons", objNFiscal.iFreteRespons, 0, "FreteRespons"
    colCampoValor.Add "MensagemNota", objNFiscal.sMensagemNota, STRING_BUFFER_MAX_TEXTO, "MensagemNota"
    colCampoValor.Add "PesoLiq", objNFiscal.dPesoLiq, 0, "Pesoliq"
    colCampoValor.Add "PesoBruto", objNFiscal.dPesoBruto, 0, "PesoBruto"
    colCampoValor.Add "NumIntNotaOriginal", objNFiscal.lNumIntNotaOriginal, 0, "NumIntNotaOriginal"
    colCampoValor.Add "Canal", objNFiscal.iCanal, 0, "Canal"
    colCampoValor.Add "Cliente", objNFiscal.lCliente, 0, "Cliente"
    colCampoValor.Add "FilialCli", objNFiscal.iFilialCli, 0, "FilialCli"
    colCampoValor.Add "FilialEntrega", objNFiscal.iFilialEntrega, 0, "FilialEntrega"
    colCampoValor.Add "NumPedidoTerc", objNFiscal.sNumPedidoTerc, STRING_BUFFER_MAX_TEXTO, "NumPedidoTerc"
    colCampoValor.Add "CodTransportadora", objNFiscal.iCodTransportadora, 0, "CodTransportadora"
    colCampoValor.Add "CodTranspRedesp", objNFiscal.iCodTranspRedesp, 0, "CodTranspRedesp"
    colCampoValor.Add "DetPagFrete", objNFiscal.iDetPagFrete, 0, "DetPagFrete"
    colCampoValor.Add "TabelaPreco", objNFiscal.iTabelaPreco, 0, "TabelaPreco"
    colCampoValor.Add "Status", objNFiscal.iStatus, 0, "Status"
    colCampoValor.Add "SemDataSaida", objNFiscal.iSemDataSaida, 0, "SemDataSaida"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
'    colSelecao.Add "Status", OP_DIFERENTE, STATUS_CANCELADO
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 35699

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156185)

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

        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 180093
        
        'Coloca os dados da NFiscal na tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 35700

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 35700, 180093

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156186)

    End Select

    Exit Sub

End Sub

Public Function Move_NF_Memoria(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objOrcamentoVenda As ClassOrcamentoVenda
Dim objProjetoInfo As Object

On Error GoTo Erro_Move_NF_Memoria

    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 35701
        
        'Não encontrou p Cliente --> erro
        If lErro = 12348 Then gError 35702

        objNFiscal.lCliente = objCliente.lCodigo
        
    End If
    
    'Verifica se Está Preenchido o Numero da Nota no Label
    If Len(Trim(NFiscal.Caption)) > 0 Then
        objNFiscal.lNumNotaFiscal = CLng(NFiscal.Caption)
    Else
        objNFiscal.lNumNotaFiscal = 0
    End If
    
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
    
    'Se o Código do Orçamento estiver preenchido
    If Len(Trim(Orcamento.Text)) <> 0 Then
        
        'Se a Filial não estiver preenchida ---> Erro.
        If Len(Trim(FilialOrcamento.Text)) = 0 Then gError 94476
        
        Set objOrcamentoVenda = New ClassOrcamentoVenda
        
        objOrcamentoVenda.lCodigo = StrParaLong(Orcamento.Text)
        objOrcamentoVenda.iFilialEmpresa = Codigo_Extrai(FilialOrcamento.Text)
        
        'Lê o Orçamento com o Código e a Filal informados.
        lErro = CF("OrcamentoVenda_Le", objOrcamentoVenda)
        If lErro <> SUCESSO And lErro <> 101232 Then gError 94477
        
        'Se não existe o Orçamento informado ---> Erro.
        If lErro = 101232 Then gError 94478
        
        'Armazena o Codigo do Orcamento
        objNFiscal.lNumOrcamentoVenda = objOrcamentoVenda.lCodigo
        objNFiscal.iFilialOrcamento = objOrcamentoVenda.iFilialEmpresa
         
    End If
    
    'Preenche o objNFiscal com os demais campos da tela
    If Len(Trim(PesoBruto.ClipText)) > 0 Then objNFiscal.dPesoBruto = CDbl(PesoBruto.Text)
    If Len(Trim(PesoLiquido.ClipText)) > 0 Then objNFiscal.dPesoLiq = CDbl(PesoLiquido.Text)

    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        objNFiscal.dtDataEmissao = CDate(DataEmissao.Text)
    Else
        objNFiscal.dtDataEmissao = DATA_NULA
    End If

    lErro = CF2(Me, "NFCF_Move_DataSaida_Memoria", objNFiscal)
    If lErro <> SUCESSO Then gError 132291

    'Verifica se os valores estão preenchidos antes de converter
    If Len(Trim(ValorDesconto.ClipText)) > 0 Then objNFiscal.dValorDesconto = CDbl(ValorDesconto.Text)
    If Len(Trim(ValorFrete.ClipText)) > 0 Then objNFiscal.dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorSeguro.ClipText)) > 0 Then objNFiscal.dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(ValorDespesas.ClipText)) > 0 Then objNFiscal.dValorOutrasDespesas = CDbl(ValorDespesas.Text)
    If Len(Trim(ValorProdutos.Caption)) > 0 Then objNFiscal.dValorProdutos = CDbl(ValorProdutos.Caption)
    If Len(Trim(ValorTotal.Caption)) > 0 Then objNFiscal.dValorTotal = CDbl(ValorTotal.Caption)
    If Len(Trim(VolumeQuant.ClipText)) > 0 Then objNFiscal.lVolumeQuant = StrParaLong(VolumeQuant.Text)

    'Armazena o responsável pelo frete
    objNFiscal.iFreteRespons = TipoFrete.ItemData(TipoFrete.ListIndex)
    
    'Recolhe os demais campos da tela
    objNFiscal.sVolumeNumero = VolumeNumero.Text
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.sNumPedidoTerc = NumPedidoTerc.Text
    objNFiscal.iCodTransportadora = Codigo_Extrai(Transportadora.Text)
    objNFiscal.iCodTranspRedesp = Codigo_Extrai(TranspRedespacho.Text)
    objNFiscal.iDetPagFrete = RedespachoCli.Value
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.iStatus = STATUS_LANCADO
    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
    objNFiscal.sMensagemNota = Mensagem.Text
    objNFiscal.sMensagemCorpoNota = MensagemCorpo.Text
    objNFiscal.sNaturezaOp = gobjTribTab.sNatOpInterna
    objNFiscal.sPlaca = Placa.Text
    objNFiscal.sPlacaUF = PlacaUF.Text
    objNFiscal.dtDataEntrada = DATA_NULA
    objNFiscal.dtDataVencimento = DATA_NULA
    objNFiscal.lNumIntNotaOriginal = 0
    objNFiscal.lNumIntDoc = 0
    objNFiscal.iCanal = Codigo_Extrai(CanalVenda.Text)
    objNFiscal.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)
    objNFiscal.iFilialEntrega = Codigo_Extrai(FilialEntrega.Text)
    objNFiscal.dtDataReferencia = DATA_NULA
    
    If Recibo.Value = vbChecked Then
        objNFiscal.iRecibo = MARCADO
    Else
        objNFiscal.iRecibo = DESMARCADO
    End If
    
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objNFiscal.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objNFiscal.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
    
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181761
    
    Set objNFiscal.objProjetoInfo = objProjetoInfo
    
    lErro = CF2(Me, "NF_Move_NF_Memoria", objNFiscal)
    If lErro <> SUCESSO Then gError 185057
    
    Set objNFiscal.objInfoAdic = gobjInfoAdic
    objNFiscal.dValorItens = StrParaDbl(ValorProdutos2.Caption)
    objNFiscal.dValorDescontoItens = StrParaDbl(ValorDescontoItens.Text)
        
    If Len(Trim(CliIntermediario.ClipText)) > 0 Then
    
        Set objCliente = New ClassCliente

        objCliente.sNomeReduzido = CliIntermediario.Text
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 35701
        
        'Não encontrou p Cliente --> erro
        If lErro = 12348 Then gError 35702

        objNFiscal.lCliIntermediario = objCliente.lCodigo
        
    End If
    objNFiscal.iFilialCliIntermediario = Codigo_Extrai(FilialCliIntermediario.Text)
    
    Move_NF_Memoria = SUCESSO

    Exit Function

Erro_Move_NF_Memoria:

    Move_NF_Memoria = gErr

    Select Case gErr

        Case 35701, 132291, 181761, 185057

        Case 35702
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case 94478
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156187)

    End Select

    Exit Function

End Function

Public Sub Almox_Change()
    iAlterado = REGISTRO_ALTERADO
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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 35849

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCanalVenda.iCodigo = iCodigo

        'Tenta ler CanalVenda com esse código no BD
        lErro = CF("CanalVenda_Le", objCanalVenda)
        If lErro <> SUCESSO And lErro <> 23597 Then gError 35850
        
        If lErro = 23597 Then gError 35851 'Não encontrou CanalVenda no BD

        'Encontrou CanalVenda no BD, coloca no Text da Combo
        CanalVenda.Text = CStr(objCanalVenda.iCodigo) & SEPARADOR & objCanalVenda.sNomeReduzido

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 35852

    Exit Sub

Erro_CanalVenda_Validate:

    Cancel = True

    Select Case gErr

        Case 35849, 35850

        Case 35851  'Não encontrou CanalVenda no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CANALVENDA", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CanalDeVenda", objCanalVenda)
            End If
            'Segura o foco

        Case 35852
            Call Rotina_Erro(vbOKOnly, "ERRO_CANALVENDA_NAO_ENCONTRADO", gErr, CanalVenda.Text)

        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156188)

    End Select

    Exit Sub

End Sub

Public Sub ClienteLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(Cliente.Text)) > 0 Then objCliente.sNomeReduzido = Cliente.Text

    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub Class_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub Cliente_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO

    Call Cliente_Preenche

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objTipoCliente As New ClassTipoCliente
Dim colTipoFrete As Collection
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer

On Error GoTo Erro_Cliente_Validate

    'Verifica se o cliente foi alterado
    If iClienteAlterado = 0 Then Exit Sub
    'Se op cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 35713

        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 35714

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
        Call CF("Filial_Preenche", FilialEntrega, colCodigoNome)

        giFilialEntregaManual = 0
        
        If Not gbCarregandoTela Then
            'Seleciona filial na Combo Filial
            If iCodFilial = FILIAL_MATRIZ Then
                Filial.ListIndex = 0
            Else
                Call CF("Filial_Seleciona", Filial, iCodFilial)
            End If
                    
            lErro = CF("TipoFrete_Le", colTipoFrete)
            If lErro <> SUCESSO Then gError 207420
            
            For Each objTipoFrete In colTipoFrete
                If objTipoFrete.iCodigoNFE = objCliente.iTipoFrete Then
                    For iIndice = 0 To TipoFrete.ListCount - 1
                        If TipoFrete.ItemData(iIndice) = objTipoFrete.iCodigo Then
                            TipoFrete.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next
        
        
        
        
        End If
        
        lErro = CF2(Me, "NFCF_Cliente_Validate", objCliente.lCodigo, iCodFilial)
        If lErro <> SUCESSO Then gError 131876
        
        'Lê o tipo do cliente
        If objCliente.iTipo > 0 Then
            objTipoCliente.iCodigo = objCliente.iTipo
            lErro = CF("TipoCliente_Le", objTipoCliente)
            If lErro <> SUCESSO And lErro <> 19062 Then gError 35739
        End If
        
        lErro = CF2(Me, "NFCF_Cliente_Validate2", objCliente, objTipoCliente)
        If lErro <> SUCESSO Then gError 131876
        
        If Not gbCarregandoTela Then
            
            Call DescontoGlobal_Recalcula
            
            'ATualiza o total com o novo desconto
            lErro = ValorTotal_Calcula()
            
            'verifica se o cliente tem uma tabela default
            If objCliente.iTabelaPreco > 0 Then
                TabelaPreco.Text = objCliente.iTabelaPreco
            ElseIf objTipoCliente.iTabelaPreco > 0 Then
                TabelaPreco.Text = objTipoCliente.iTabelaPreco
            End If
        
            If Len(Trim(TabelaPreco.Text)) > 0 Then TabelaPreco_Validate (bSGECancelDummy)
        
            lErro = CF2(Me, "NFCF_Cliente_Validate3", objCliente)
            If lErro <> SUCESSO Then gError 42758
            
        End If
                
        iClienteAlterado = 0

        If Not gbCarregandoTela Then
''*** incluidos p/tratamento de tributacao *******************************
            Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, iCodFilial) '####
'*** fim tributacao
        End If
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        Filial.Clear
        FilialEntrega.Clear

        Call CF2(Me, "NFCF_Limpa_Tela")

    End If

    iClienteAlterado = 0

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 35713, 35714, 35739, 131876, 42758 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156189)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO
    
    iAlterado = REGISTRO_ALTERADO
    iDataEmissaoAlterada = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    If iDataEmissaoAlterada = 1 Then

        'Verifica se a data de emissao foi digitada
        If Len(Trim(DataEmissao.ClipText)) <> 0 Then
    
            'Critica a data digitada
            lErro = Data_Critica(DataEmissao.Text)
            If lErro <> SUCESSO Then gError 35752
        
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

        Case 35752

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156190)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 35753

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 35753

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156191)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 35754

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 35754

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156192)

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
    If lErro <> SUCESSO Then gError 35755

    Exit Sub

Erro_DataSaida_Validate:

    Cancel = True

    Select Case gErr
    
        Case 35755

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156193)

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
    If lErro <> SUCESSO Then gError 89796

    Exit Sub

Erro_HoraSaida_Validate:

    Cancel = True

    Select Case gErr

        Case 89796

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156194)

    End Select

    Exit Sub

End Sub

Public Sub UpDownSaida_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownSaida_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataSaida, DIMINUI_DATA)
    If lErro Then gError 35757

    Exit Sub

Erro_UpDownSaida_DownClick:

    Select Case gErr

        Case 35757

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156195)

    End Select

    Exit Sub

End Sub

Public Sub UpDownSaida_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownSaida_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataSaida, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 35758

    Exit Sub

Erro_UpDownSaida_UpClick:

    Select Case gErr

        Case 35758

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156196)

    End Select

    Exit Sub

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

Public Sub Filial_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    'Verifica se algo foi selecionada
    If Filial.ListIndex = -1 Then Exit Sub

    'Faz o tratamento da Filial selecionada
    lErro = Trata_FilialCliente()
    If lErro <> SUCESSO Then gError 35741

    Exit Sub

Erro_Filial_Click:

    Select Case gErr

        Case 35741

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156197)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 35724

    'Se nao encontra o item com o código informado
    If lErro = 6730 Then

        'Verifica de o Cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 59734

        sCliente = Cliente.Text

        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 35725

        If lErro = 17660 Then gError 35726

        'Coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        lErro = Trata_FilialCliente()
        If lErro <> SUCESSO Then gError 59735

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 35728

    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case gErr

        Case 35724, 35725

        Case 35726
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 35728
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case 59734
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 59735

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156198)

    End Select

    Exit Sub

End Sub

Public Sub FilialEntrega_Change()
    iAlterado = REGISTRO_ALTERADO
    giFilialEntregaAlterada = 1
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
    If Len(Trim(FilialEntrega.Text)) = 0 Or giFilialEntregaAlterada = 0 Then Exit Sub

    'Verifica se é uma FilialEntrega selecionada
    If FilialEntrega.ListIndex <> -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialEntrega, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 26700

    'Se nao encontra o item com o código informado
    If lErro = 6730 Then

        'Verifica de o Cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 26703

        sCliente = Cliente.Text

        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe FilialEntrega com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 26701

        If lErro = 17660 Then gError 26702

        'coloca na tela
        FilialEntrega.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 26704

    giFilialEntregaAlterada = 0
    
    Exit Sub

Erro_FilialEntrega_Validate:

    Cancel = True


    Select Case gErr

        Case 26700, 26701

        Case 26702
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 26703
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 26704
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialEntrega.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156199)

    End Select

    Exit Sub

End Sub

Public Sub ItemNFiscal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Mensagem_Change()
    iAlterado = REGISTRO_ALTERADO
    MsgAutomatica.Value = vbUnchecked
End Sub

Public Sub NFiscal_Change()
    iAlterado = REGISTRO_ALTERADO
    NFiscalAux.PromptInclude = False
    NFiscalAux.Text = NFiscal.Caption
    NFiscalAux.PromptInclude = True
End Sub

Public Sub NFiscalLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim lErro As Long
Dim colSelecao As Collection

    'Recolhe os dados da Nota Fiscal
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then Exit Sub

    'Chama a Tela NFiscalLista
    Call Chama_Tela("NFiscalLista", colSelecao, objNFiscal, objEventoNFiscal)

    Exit Sub

End Sub

Private Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objNFiscal = obj1

    'Evita ter que alterar os browses a cada campo inserido na tabela de NF
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 180099

    'Coloca na Tela a Nota Fiscal escolhida
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 35715

    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case gErr

        Case 35715, 180099

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156200)

    End Select

    Exit Sub

End Sub

Public Sub NFiscalOriginal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumPedidoTerc_Change()
    iAlterado = REGISTRO_ALTERADO
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

Public Sub PercentualBaixa_Change()
    iAlterado = REGISTRO_ALTERADO
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
    If lErro <> SUCESSO Then gError 35848

    'Coloca o valor formatado na Tela
    PesoBruto.Text = Format(PesoBruto.Text, PesoBruto.Format)

    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True

    Select Case gErr

        Case 35848

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156201)

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
    If lErro <> SUCESSO Then gError 35847

    'Coloca o valor formatado na Tela
    PesoLiquido.Text = Format(PesoLiquido.Text, PesoLiquido.Format) 'Alterado por Wagner

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True

    Select Case gErr

        Case 35847

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156202)

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

    'Verifica se existe o item na combo
    lErro = Combo_Item_Igual(PlacaUF)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 35845

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 35846

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True

    Select Case gErr

        Case 35845

        Case 35846
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156203)

    End Select

    Exit Sub

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

Public Sub Produto_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO
    
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

Public Sub BotaoProdutos_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is Produto Then
        sProduto1 = Produto.Text
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 43693

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 42762
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
    
    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub
    
Erro_BotaoProdutos_Click:

    Select Case gErr
    
        Case 42762
        
        Case 43693
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156204)
            
    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoEnxuto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    'verifica se tem alguma linha do Grid selecionada
    If GridItens.Row = 0 Then gError 35806

    'Verifica se o Produto está preenchido
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then

        Set objProduto = obj1

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 35807

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
        
        If Not (Me.ActiveControl Is Produto) Then
        
            'Coloca o produto, a Descrição e a Unidade de Medida da tela
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text

            'lErro = Saida_Celula_Produto(objGridItens)
            lErro = Produto_Saida_Celula(objProduto)
            If lErro <> SUCESSO Then gError 35808
        
        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    Select Case gErr

        Case 35806, 35808

        Case 35807
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156205)

    End Select

    Exit Sub

End Sub

Public Sub ProdutoAlmox_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub QuantAlocada_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Quantidade_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO
    
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

Public Sub QuantVendida_Change()
    iAlterado = REGISTRO_ALTERADO
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
            If lErro <> SUCESSO And lErro <> 12253 Then gError 35751
            
            'Se não está na combo
            If lErro <> SUCESSO Then
            
                objSerie.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
                'Busca a série no BD
                lErro = CF("Serie_Le", objSerie)
                If lErro <> SUCESSO And lErro <> 22202 Then gError 51606
                If lErro <> SUCESSO Then gError 35750 'Se não encontrou
            
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

        Case 35750
            'Pergunta se deseja criar nova série
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_SERIE", Serie.Text)
            'Se a resposta for afirmativa
            If vbMsgRes = vbYes Then
                'Chama a tela de cadastro de séries
                Call Chama_Tela("SerieNFiscal", objSerie)
            End If
            'segura o foco na série

        Case 35751, 51606

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156206)

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

Public Sub SerieNFiscalOriginal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SerieNFiscalOriginal_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SerieNFiscalOriginal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lTamanho As Long

On Error GoTo Erro_SerieNFiscalOriginal_Validate

    'Verifica se o tamanho da Serie é maior que 3
    If Len(Trim(SerieNFiscalOriginal.Text)) > STRING_SERIE Then gError 35019

    Exit Sub

Erro_SerieNFiscalOriginal_Validate:

    Cancel = True

    Select Case gErr

        Case 35019
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_SERIE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156207)

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

    'Verifica se existe algum item no Grid
    If objGridItens.iLinhasExistentes = 0 Then Exit Sub

    'Faz o tratamento para a Tabela Preco selecionada
    lErro = Trata_TabelaPreco()
    If lErro <> SUCESSO Then gError 35762

    Exit Sub

Erro_TabelaPreco_Click:

    Select Case gErr

        Case 35762

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156208)

    End Select

    Exit Sub

End Sub

Public Sub TabelaPreco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTabelaPreco As New ClassTabelaPreco
Dim iCodigo As Integer
Dim objTabelaPrecoItem As ClassTabelaPrecoItem

On Error GoTo Erro_TabelaPreco_Validate

    'Verifica se foi preenchida a ComboBox TabelaPreco
    If Len(Trim(TabelaPreco.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox TabelaPreco
    If TabelaPreco.Text = TabelaPreco.List(TabelaPreco.ListIndex) Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(TabelaPreco, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 35764

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 35765
        
        If lErro <> SUCESSO Then gError 35766 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 35763

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 35767

    Exit Sub

Erro_TabelaPreco_Validate:

    Cancel = True

    Select Case gErr

        Case 35764, 35765, 35763
    
        Case 35766  'Não encontrou Tabela de Preço no BD
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")
    
            If vbMsgRes = vbYes Then
                    
                'Chama a tela de Criação de Tabelas de Preço
                Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
            End If
            'Segura o foco
    
        Case 35767
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156209)

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
            If lErro <> SUCESSO Then gError 27840
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

        Select Case iFrameAtual
        
            Case TAB_DadosPrincipais
                Parent.HelpContextID = IDH_NF_SAIDA_DADOS_PRINCIPAIS
            
            Case TAB_Itens
                Parent.HelpContextID = IDH_NF_SAIDA_ITENS
            
            Case TAB_Complemento
                Call Calcula_Msg
                Parent.HelpContextID = IDH_NF_SAIDA_COMPLEMENTO
            
            Case TAB_Almoxarifado
                Parent.HelpContextID = IDH_NF_SAIDA_ALMOXARIFADO
            
            'Alterado por Tulio em 27/03
            Case TAB_Comissoes
                'se for tab de comissoes e a empresa
                'utilizar regras para o calculo de comissoes
                lErro = objTabComissoes.Comissoes_Calcula
                If lErro <> SUCESSO Then gError 98991
                
                Parent.HelpContextID = IDH_PEDIDO_VENDA_COMISSOES
            
            Case TAB_Tributacao
                Parent.HelpContextID = IDH_NF_SAIDA_TRIBUTACAO
            
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_NF_SAIDA_CONTABILIZACAO
            
        End Select
    
    End If

    Exit Sub
    
Erro_TabStrip1_Click:

    Select Case gErr

        Case 27840, 98991
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156211)

    End Select

    Exit Sub

End Sub

Public Sub TipoNFiscal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoNFiscal_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TipoNFiscal_Click

'    If giTipoVersao = VERSAO_FULL Then
        Frame1(TAB_Almoxarifado).Enabled = True
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'        FrameL(TAB_Almoxarifado).Enabled = True
'    End If

    If TipoNFiscal.ListIndex = -1 Then Exit Sub

    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    gobjRastreamento.iCodigo = Codigo_Extrai(TipoNFiscal.Text) 'WW

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 35735

    'se não estiver cadastrado ==> erro
    If lErro = 31415 Then gError 35736

    'If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Or objTipoDocInfo.iCodigo = DOCINFO_NFISVNE Or objTipoDocInfo.iCodigo = DOCINFO_NFISS Then
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Or (objTipoDocInfo.iTipoMovtoEstoque = 0 And objTipoDocInfo.iTipoMovtoEstoque = 0) Then
'        If giTipoVersao = VERSAO_FULL Then
            Frame1(TAB_Almoxarifado).Enabled = False
'        ElseIf giTipoVersao = VERSAO_LIGHT Then
'            FrameL(TAB_Almoxarifado).Enabled = False
'        End If
        
    Else
'        If giTipoVersao = VERSAO_FULL Then
            Frame1(TAB_Almoxarifado).Enabled = True
'        ElseIf giTipoVersao = VERSAO_LIGHT Then
'            FrameL(TAB_Almoxarifado).Enabled = True
'        End If
    End If
    
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then
        giComplemento = MARCADO
    Else
        giComplemento = DESMARCADO
    End If
    
    Call Trata_NF_Servico
        
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
    If lErro <> SUCESSO Then gError 83323
        
    'se o cliente não estiver preenchido ==> preencher com a natureza padrao o campo natureza da operacao
    If Len(Trim(Cliente.Text)) = 0 Then Call gobjTribTab.NatOpInterna_AlterarTrocaTipoNF(NaturezaOp_Conv34(objTipoDocInfo.sNaturezaOperacaoPadrao, MaskedParaDate(DataEmissao)))

    ''*** incluidos p/tratamento de tributacao *******************************
    If gbLimpandoTela = False Then
        Call gobjTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
        If IsDate(DataEmissao.Text) Then Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
        Call gobjTribTab.Serie_Alterada(Serie.Text)
    End If
    '*** fim tributacao
    
    If objTipoDocInfo.iCodigo = DOCINFO_NFISVFE And gobjFAT.iVFEUsaTManual = MARCADO Then
        NFiscal.Visible = False
        BotaoLimparNF.Visible = False
        NFiscalAux.Visible = True
        NFiscalAux.PromptInclude = False
        NFiscalAux.Text = ""
        NFiscalAux.PromptInclude = True
    Else
        NFiscal.Visible = True
        BotaoLimparNF.Visible = True
        NFiscalAux.Visible = False
    End If

    'Processa alteração no subtipo contábil para que sejam carregados os modelos referentes ao novo TipoDocInfo selecionado
    lErro = objContabil.Contabil_Processa_Alteracao_Subtipo(objTipoDocInfo.iSubTipoContabil)
    If lErro <> SUCESSO Then gError 79331

    Exit Sub

Erro_TipoNFiscal_Click:

    Select Case gErr

        Case 35735, 79331, 83323

        Case 35736
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156212)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 35733
    'Se não conseguir --> Erro
    If lErro <> SUCESSO Then gError 35734

    Exit Sub

Erro_TipoNFiscal_Validate:

    Cancel = True

    Select Case gErr

        Case 35733

        Case 35734
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156213)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 35841
    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo
        
        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 35842
        
        If lErro = 19250 Then gError 35843 'Não encontrou Transportadora no BD

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    End If

    'Não existe o item com a String na List da Combo
    If lErro = 6731 Then gError 35844

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case gErr

        Case 35841, 35842

        Case 35843
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then
                Call Chama_Tela("Transportadora", objTransportadora)
            End If
            
        Case 35844
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156214)

    End Select

    Exit Sub

End Sub

Public Sub UnidadeMed_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMed_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UnidadeMedEst_Change()
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

Public Sub ValorBaixa_Change()
    iAlterado = REGISTRO_ALTERADO
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

Public Sub ValorDesconto_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    iValorDescontoAlterado = REGISTRO_ALTERADO

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
        If lErro <> SUCESSO Then gError 35837

        dValorDesconto = CDbl(ValorDesconto.Text)
        'Coloca o Valor formatado na tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")
        
        lErro = DescontoGlobal_Testa(dValorDesconto)
        If lErro <> SUCESSO Then gError 56847

        dValorProdutos = dValorProdutos - dValorDesconto

    End If

    ValorProdutos.Caption = Format(dValorProdutos, "Standard")

    lErro = CF2(Me, "NFCF_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132019

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDesconto_Validate(Cancel, dValorDesconto)
'*** fim tributacao

    Call ValorTotal_Calcula

    iValorDescontoAlterado = 0
    
    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True

    Select Case gErr

        Case 35837, 56847, 132019
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156215)

    End Select

    Exit Sub

End Sub

Public Sub ValorDespesas_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    iValorDespesasAlterado = REGISTRO_ALTERADO

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

Public Sub ValorFrete_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    iValorFreteAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorProdutos_Change()
'    If Not gbCarregandoTela Then
'         Call DescontoGlobal_Recalcula
'    End If
End Sub


Public Sub ValorSeguro_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    iValorSeguroAlterado = REGISTRO_ALTERADO

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
        If lErro <> SUCESSO Then gError 35836

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

        Case 35836

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156216)

    End Select

    Exit Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156217)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156218)

    End Select

End Sub

Public Sub VolumeNumero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub VolumeQuant_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Caso a função esteja sendo chamada para carregar um OrcamentoVenda, receberá um objOrcamentoVenda.
Public Function Traz_NFiscal_Tela(objNFiscal As ClassNFiscal, Optional objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim objNFiscalOriginal As New ClassNFiscal
Dim bCancel As Boolean
Dim bAlterouCT As Boolean
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim lCodOrcamento As Long
Dim objTipoDocInfo1 As New ClassTipoDocInfo, iTipoNFiscal As Integer
Dim bOV As Boolean
Dim iIndice As Integer
Dim dPercDesc As Double

On Error GoTo Erro_Traz_NFiscal_Tela
    
    gbCarregandoTela = True
       
    'Limpa a tela
    lErro = Limpa_Tela_NFiscal()
    If lErro <> SUCESSO Then gError 500116
    
    Set gobjNFiscal = objNFiscal

    bOV = True
    'Se a função não for chamada para carregar os dados de um orçamento de venda. Ou seja, a NF já existe no BD.
    If (objOrcamentoVenda Is Nothing) Then
    
        bOV = False
        
        iTipoNFiscal = objNFiscal.iTipoNFiscal

        'Lê as mensagens
        lErro = CF("NFiscal_Le_Msg", objNFiscal)
        If lErro <> SUCESSO Then gError 35705
    
        'Lê os itens da Nota Fiscal
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 35705
        
        'Lê a parte de Tributação
        lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
        If lErro <> SUCESSO And lErro <> 22867 Then gError 39342
    
        'Lê a parte de tributação dos itens da Nota Fiscal
        lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
        If lErro <> SUCESSO Then gError 39343
        
        'Lê as Comissões da Nota Fiscal
        lErro = CF("ComissoesNF_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 35703
    
        'Lê as Alocações dos itens da Nota Fiscal
        lErro = CF("AlocacoesNF_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 35704
        
        Set objOrcamentoVenda = New ClassOrcamentoVenda
        
        objOrcamentoVenda.lNumIntNFiscal = objNFiscal.lNumIntDoc
        
        'Lê os dados do OrcamentoVenda através do seu NumIntNFiscal
        lErro = CF("OrcamentoVenda_Le_NumIntNFiscal", objOrcamentoVenda)
        If lErro <> SUCESSO And lErro <> 101305 Then gError 94485
        
    Else
    
        'Busca nos objs da Colecao o Tipo Padrao => guarda no obj
        For Each objTipoDocInfo1 In gobjCRFAT.colTiposDocInfo
            If objTipoDocInfo1.sNomeTelaNFiscal = NOME_TELA_NFISCAL And objTipoDocInfo1.iPadrao = Padrao Then
                
                iTipoNFiscal = objTipoDocInfo1.iCodigo
                Exit For
            
            End If
        Next
    
    End If
    
    'Se o código do Orcamento estiver carregado
    If objOrcamentoVenda.lCodigo <> 0 Then
        Orcamento.Text = objOrcamentoVenda.lCodigo
    End If
    
    'Se o Código da Filial estiver carregado
    If objOrcamentoVenda.iFilialEmpresa <> 0 Then
        
        FilialOrcamento.Text = objOrcamentoVenda.iFilialEmpresa
        Call FilialOrcamento_Validate(bSGECancelDummy)
    
    End If
    
    'Preenche o Status da Nota Fiscal
    If objNFiscal.iStatus = STATUS_LANCADO Then
        Status.Caption = STRING_STATUS_LANCADO
    ElseIf objNFiscal.iStatus = STATUS_BAIXADO Then
        Status.Caption = STRING_STATUS_BAIXADO
    ElseIf objNFiscal.iStatus = STATUS_CANCELADO Then
        Status.Caption = STRING_STATUS_CANCELADO
    End If
    
    'Se o cliente ainda tiver um código, ou seja, estiver cadastrado.
    If objNFiscal.lCliente <> 0 Then
    
        Call Cliente_Formata(objNFiscal.lCliente)

        Call Filial_Formata(Filial, objNFiscal.iFilialCli)
        
        'Coloca a FilialEntrega na tela
        Call Filial_Formata(FilialEntrega, objNFiscal.iFilialEntrega)
        
        giFilialAlterada = 0
            
    Else
        Cliente.Text = objOrcamentoVenda.sNomeCli
        Filial.Clear
        Filial.Text = objOrcamentoVenda.sNomeFilialCli
    End If
    
    'Coloca TabelaPreco na Tela
    If objNFiscal.iTabelaPreco > 0 Then
        TabelaPreco.Text = objNFiscal.iTabelaPreco
        Call TabelaPreco_Validate(bSGECancelDummy)
    End If

    'Coloca o Tipo na tela
    TipoNFiscal.Text = iTipoNFiscal
    Call TipoNFiscal_Validate(bSGECancelDummy)
    
    If objNFiscal.lNumNotaFiscal <> 0 Then
        'Preenche os demais dados da tela
        NFiscal.Caption = objNFiscal.lNumNotaFiscal
        Serie.Text = Desconverte_Serie_Eletronica(objNFiscal.sSerie)
        If ISSerieEletronica(objNFiscal.sSerie) Then
            EletronicaFed.Value = vbChecked
        Else
            EletronicaFed.Value = vbUnchecked
        End If
        Call Serie_Validate(bSGECancelDummy)
    Else
        'Seleciona série padrão
        lErro = Seleciona_Serie_Padrao()
        If lErro <> SUCESSO Then gError 132032
    End If
    
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
    ValorDesconto.Text = Format(objNFiscal.dValorDesconto, "Standard")
    ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Standard")
    ValorFrete.Text = Format(objNFiscal.dValorFrete, "Standard")
    ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Standard")
    ValorTotal.Caption = Format(objNFiscal.dValorTotal, "Standard")
    ValorProdutos.Caption = Format(objNFiscal.dValorProdutos, "Standard")

    lErro = CF2(Me, "NFCF_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132020

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
    
    For iIndice = 0 To TipoFrete.ListCount - 1
        If TipoFrete.ItemData(iIndice) = objNFiscal.iFreteRespons Then
            TipoFrete.ListIndex = iIndice
            Exit For
        End If
    Next
    
    If objNFiscal.lVolumeQuant > 0 Then VolumeQuant.Text = objNFiscal.lVolumeQuant
    NumPedidoTerc.Text = objNFiscal.sNumPedidoTerc

    Mensagem.Text = objNFiscal.sMensagemNota
    MensagemCorpo.Text = objNFiscal.sMensagemCorpoNota

    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = Format(objNFiscal.dPesoLiq, PesoLiquido.Format) 'Alterado por Wagner
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = Format(objNFiscal.dPesoBruto, PesoBruto.Format) 'Alterado por Wagner

    'Preenche o Grid com os itens da Nota Fiscal
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 39023
    
    objTipoDocInfo.iCodigo = iTipoNFiscal
    
    'Le o Tipo para saber se tem Movimento de Estoque
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO Then gError 83325
    
    'Carrega o Tab Comissões
    lErro = objTabComissoes.Carrega_Tab_Comissoes(objNFiscal)
    If lErro <> SUCESSO Then gError 51609
    
    'Preenche o Grid com as Alocações dos itens da Nota Fiscal
    lErro = Preenche_GridAlocacoes(objNFiscal)
    If lErro <> SUCESSO Then gError 39026

    If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
        'Carrega ItensNF com Rastreamentos
        lErro = gobjRastreamento.Carrega_RastroItensNF(objNFiscal)
        If lErro <> SUCESSO Then gError 83326
    End If
        
    'Se o código da Nota Fiscal Original estiver preenchido
    If objNFiscal.lNumIntNotaOriginal <> 0 Then
        
        objNFiscalOriginal.lNumIntDoc = objNFiscal.lNumIntNotaOriginal
        'Lê a Nota Fiscal Original
        lErro = CF("NFiscal_Le", objNFiscalOriginal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 39024
        'Se não achou a Nota Fiscal Original --> erro
        If lErro = 31442 Then gError 39025
        
        'Coloca os dados relativos a Nota Fiscal Original na Tela
        SerieNFiscalOriginal.Text = objNFiscalOriginal.sSerie
        NFiscalOriginal.Text = objNFiscalOriginal.lNumNotaFiscal
        
    End If
    
    lErro = CF2(Me, "NFCF_Cliente_Traz_NF_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 131877
    
    lErro = CF2(Me, "NF_Traz_NF_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 131877
    
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
    
    Set gobjNFiscal = objNFiscal
    
    ''*** incluídos p/tratamento de tributação ****
    lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 39031
    '*** fim tributação
    
    'Traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 39661

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
    
    If bOV Then
        MsgAutomatica.Value = vbChecked
        
        If objOrcamentoVenda.iMoeda <> MOEDA_REAL Then
            For iIndice = 1 To objGridItens.iLinhasExistentes
                lErro = gobjTribTab.Alteracao_Item_Grid(iIndice)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Call PrecoTotal_Calcula(iIndice)
            Next
            Call ValorTotal_Calcula
        End If
        
    Else
        MsgAutomatica.Value = vbUnchecked
    End If
     
    Set gobjInfoAdic = objNFiscal.objInfoAdic
    
    'Se o cliente ainda tiver um código, ou seja, estiver cadastrado.
    If objNFiscal.lCliIntermediario <> 0 Then
        CliIntermediario.Text = objNFiscal.lCliIntermediario
        Call CliIntermediario_Validate(bSGECancelDummy)

        FilialCliIntermediario.Text = objNFiscal.iFilialCliIntermediario
        Call FilialCliIntermediario_Validate(bSGECancelDummy)
    Else
        CliIntermediario.Text = ""
        FilialCliIntermediario.Clear
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

        Case 25828, 25829, 39031, 35705, 39023, 39024, 39342, 39343, 39661, 51609, 83325, 83326, 500116, 101305, 131877, 132020 'Inserido por Wagner

        Case 39025
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntNotaOriginal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156219)

    End Select

    Exit Function

End Function

Private Sub SubTotal_Calcula(objGridInt As AdmGrid, iGrid_Coluna As Integer, dSubTotal As Double)
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
    If GridItens.Row = 0 Then gError 43694
        
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
            If lErro <> SUCESSO And lErro <> 23796 Then gError 58053
            
            If lErro = SUCESSO Then objEstoqueProd.iAlmoxarifado = iAlmoxarifado
            
            'Chama a tela EstoqueProdutoFilialLista
            Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProd, objEventoEstoqueProd)
    
        Else
        
            Call Mascara_RetornaProdutoEnxuto(sProdutoFormatado, sProdutoEnxuto)
            sSelecao = "Produto LIKE '" & sProdutoEnxuto & "%' AND FilialEmpresa = " & giFilialEmpresa
            Call Chama_Tela("EstProdLista_Consulta", colSelecao, objEstoqueProd, objEventoEstoqueProd, sSelecao)
        
        End If
    
    Else
        gError 43744
    End If
  
    Exit Sub

Erro_BotaoEstoqueProd_Click:

    Select Case gErr
    
        Case 43694
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 43744
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 58053, 101782
        
        Case 101783
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156220)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoEstoqueProd_evSelecao(obj1 As Object)
    Me.Show
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

Private Sub SerieOriginalLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'Recolhe a Série da Nota Fiscal Original da tela
    objSerie.sSerie = SerieNFiscalOriginal.Text

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerieOriginal)

    Exit Sub

End Sub

Private Sub objEventoSerieOriginal_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    'Coloca a Série da Nota Fiscal Original na tela
    SerieNFiscalOriginal.Text = objSerie.sSerie

    Me.Show

End Sub

Public Sub NFiscalOriginalLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection
Dim sCliente As String
Dim lErro As Long

On Error GoTo Erro_NFiscalOriginalLabel_Click

    'Se o cliente não está preenchido => erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 101090
    
    'Recolhe os dados da Nota Fiscal
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 31437

    'Guarda a Serie e o Número da Nota Fiscal Original da Tela
    objNFiscal.sSerie = SerieNFiscalOriginal.Text
    If Len(Trim(NFiscalOriginal.ClipText)) > 0 Then
        objNFiscal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
    Else
        objNFiscal.lNumNotaFiscal = 0
    End If
    
    'Guarda o nome reduzido do Cliente
    sCliente = Cliente.Text
    
    'Faz o filtro por cliente
    colSelecao.Add sCliente
    colSelecao.Add sCliente

    'Chama a Tela NFiscalNormalLista
    Call Chama_Tela("NFiscalNormalLista", colSelecao, objNFiscal, objEventoNFiscalOriginal)

    Exit Sub

Erro_NFiscalOriginalLabel_Click:

    Select Case gErr

        Case 31437
        
        Case 101090
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156221)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156222)
    
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
                If lErro <> SUCESSO Then gError 51608
            
            'senao
             Else
                'verificar se o vendedor esta no grid e tratar o campo 'DiretoIndireto'
                lErro = objTabComissoes.VendedorLinha_Preenche1(objVendedor)
                If lErro <> SUCESSO Then gError 98992
        
            End If
        
    End If

    Me.Show
    
    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr

        Case 51608, 98992 'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156223)

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
        If lErro <> SUCESSO Then gError 64003

        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 35727
        If lErro <> SUCESSO Then gError 35730

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

        'Lê o Tipo de Documento com o Código de objNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 35729
        'Se não encontrar --> Erro
        If lErro = 31415 Then gError 35731
        
'        If giTipoVersao = VERSAO_FULL Then
            If objTipoDocInfo.sNomeTelaNFiscal <> NOME_TELA_NFISCAL Then gError 35732
'        End If
        
        'Traz a nota para a tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 35728

    End If

    iAlterado = 0
    iComissoesAlterada = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 35727, 35728, 35729, 64003

        Case 35730
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case 35731
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 35732
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_DIFERENTE_NF_VENDA", gErr, objNFiscal.iTipoNFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156224)

    End Select
    
    iAlterado = 0
    iComissoesAlterada = 0

    Exit Function

End Function

Private Function Trata_FilialCliente() As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente, objCliente As New ClassCliente
Dim objVendedor As New ClassVendedor
Dim objComissaoNF As New ClassComissaoNF
Dim objTransportadora As New ClassTransportadora
Dim objMensagem As New ClassMensagem
Dim colTipoFrete As Collection
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer

On Error GoTo Erro_Trata_FilialCliente

    objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
    'Lê a FilialCliente
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilialCliente)
    If lErro <> SUCESSO Then gError 35742

    objCliente.sNomeReduzido = Cliente.Text
    
    'Lê o Cliente
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 35743
    If lErro = 12348 Then gError 35744

    Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, objFilialCliente.iCodFilial)
        
    'Comissões
    'se empresa nao utiliza as regras para
    'calculo de comissoes
    If Not gobjCRFAT.iUsaComissoesRegras = USA_REGRAS And ComissaoAutomatica.Value = 1 Then
        
        'calcular as comissoes no modelo antigo...
        lErro = objTabComissoes.Comissao_Automatica_FilialCli_Exibe(objFilialCliente)
        If lErro <> SUCESSO Then gError 59022
    
    End If

    'Verifica se a transportadora está preenchida
    If Len(Trim(Transportadora.Text)) = 0 Then
    
        lErro = CF("Transportadora_Padrao_FilCli", objFilialCliente, objTransportadora)
        If lErro <> SUCESSO Then gError 59023

        If objTransportadora.iCodigo <> 0 Then
                
            'Coloca a transportadora na tela
            Transportadora.Text = objTransportadora.sNomeReduzido
            Call Transportadora_Validate(bSGECancelDummy)
        
        End If
    
    End If
    
    'Verifica se a transportadora de Redespacho está preenchida
    If Len(Trim(TranspRedespacho.Text)) = 0 Then
    
        If objFilialCliente.iCodTranspRedesp <> 0 Then
                
            objTransportadora.iCodigo = objFilialCliente.iCodTranspRedesp
            lErro = CF("Transportadora_Le", objTransportadora)
            If lErro <> SUCESSO And lErro <> 19250 Then gError 124056
            If lErro <> SUCESSO Then gError 124057
            
            'Coloca a transportadora na tela
            TranspRedespacho.Text = objTransportadora.sNomeReduzido
            Call TranspRedespacho_Validate(bSGECancelDummy)
        
            RedespachoCli.Value = objFilialCliente.iDetPagFrete
        
        End If
    
    End If
    
    'Seleciona tipo de frete da Filial do Cliente
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207421
    
    For Each objTipoFrete In colTipoFrete
        If objTipoFrete.iCodigoNFE = objFilialCliente.iTipoFrete Then
            For iIndice = 0 To TipoFrete.ListCount - 1
                If TipoFrete.ItemData(iIndice) = objTipoFrete.iCodigo Then
                    TipoFrete.ListIndex = iIndice
                    Exit For
                End If
            Next
            Exit For
        End If
    Next



    If giFilialEntregaManual = 0 Then
    
        FilialEntrega.Text = CStr(objFilialCliente.iCodFilial)
        Call FilialEntrega_Validate(bSGECancelDummy)
        
        giFilialEntregaManual = 0
        
    End If
        
    lErro = CF2(Me, "NFCF_Cliente_Filial_Validate", objFilialCliente)
    If lErro <> SUCESSO Then gError 131878
        
'    If objFilialCliente.iCodMensagem > 0 Then
'        objMensagem.iCodigo = objFilialCliente.iCodMensagem
'        'Lê a Mensagem "default" p/ o Cliente
'        lErro = CF("Mensagem_Le", objMensagem)
'        If lErro <> SUCESSO And lErro <> 19234 Then gError 140375
'
'        If lErro = SUCESSO Then Mensagem.Text = objMensagem.sDescricao
'
'    End If
    
    Trata_FilialCliente = SUCESSO

    Exit Function

Erro_Trata_FilialCliente:

    Trata_FilialCliente = gErr

    Select Case gErr

        Case 35742, 35743, 35744, 59022, 59023, 124056, 131878, 140375, 207421

        Case 124057
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objTransportadora.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156225)
            
    End Select

    Exit Function

End Function

Public Sub PrecoTotal_Calcula(iLinha As Integer, Optional iCalculaTotal = True)

Dim lErro As Long
Dim iIndice As Integer
Dim lTamanho As Long
Dim vbMsgRes As VbMsgBoxResult
Dim dDesconto As Double
Dim dValorTotal As Double
Dim dPrecoTotal As Double
Dim dQuantidade As Double
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dValorProdutos As Double
Dim dPrecoTotalReal As Double
Dim dValorDescontoGlobal As Double, dValorDescontoItens As Double

On Error GoTo Erro_PrecoTotal_Calcula

    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))
    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))

    If giComplemento <> DESMARCADO Then
        If dQuantidade = 0 Then dQuantidade = 1 'para nao zerar o preco total qdo for complemento de preco
    End If
    
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
    If lErro <> SUCESSO Then gError 184277
    
    Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)
    
'    'Obtem o valor do Desconto global
'    If gdDesconto > 0 And iValorDescontoAlterado = 0 Then
'        dValorDescontoGlobal = gdDesconto * dValorProdutos
'    ElseIf Len(Trim(ValorDesconto.Text)) > 0 And IsNumeric(ValorDesconto.Text) Then
'        dValorDescontoGlobal = CDbl(ValorDesconto.Text)
'    End If
'
'    'Calcula o Valor dos Produtos - Valor do Desconto Global
'    dValorProdutos = dValorProdutos - dValorDescontoGlobal
    
'    'Se com o desconto o valor Produtos ficou negativo
'    If dValorProdutos < 0 And dValorDescontoGlobal > 0 Then
'
'        'Avisa ao usuário
'        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", dValorDescontoGlobal, dValorProdutos)
'
'        'Limpa o valor de desconto
'        gdDesconto = 0
'        Call CF2(Me, "NFCF_Limpa_Desconto")
'        ValorDesconto.Text = ""
'        iValorDescontoAlterado = 0
'        dValorDescontoGlobal = 0
'
'''*** incluidos p/tratamento de tributacao *******************************
'        Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, dValorDescontoGlobal)
''*** fim tributacao
'
'        'Faz a atualização dos valores da tributação
'        lErro = gobjTribTab.AtualizarTributacao()
'        If lErro <> SUCESSO Then gError 56881
'
'        Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)
'
'    End If

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
    If lErro <> SUCESSO Then gError 132021

    If iCalculaTotal <> False Then Call ValorTotal_Calcula

    Exit Sub

Erro_PrecoTotal_Calcula:

    Select Case gErr

        Case 56881, 184277 'Tratado na rotina chamada
        
        Case 132021 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156226)

    End Select
    
    Exit Sub
    
End Sub

Public Function ValorTotal_Calcula() As Long

Dim dValorTotal As Double
Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    lErro = gobjTribTab.ValorTotal_Calcula(dValorTotal)
    If lErro <> SUCESSO Then gError 27836

    'Coloca o valor total na tela
    ValorTotal.Caption = Format(dValorTotal, "Standard")

    Call ValorDescontoItens_Calcula
    
    If Not gbCarregandoTela Then
    
        'Se a empresa nãu utiliza as regras para o cálculo de comissões
        If Not gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
        
            'Significa que é necessário recalcular as comissões, utilizando o cálculo padrão
            lErro = objTabComissoes.Comissoes_Calcula_Padrao()
            If lErro <> SUCESSO Then gError 35809
        
        Else
            iComissoesAlterada = 1 'preventivamente, pois pode ter alterado algo que afete uma regra
        End If
    
    End If

    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 35809, 27836 ', 56882

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156227)

    End Select

    Exit Function

End Function

Function Pesos_Recalcula() As Long

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
            If lErro <> SUCESSO Then gError 500081
            
            'Lê o produto
            objProduto.sCodigo = sProdutoFormatado
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 500082

            'Se o produto não está cadastrado, erro
            If lErro = 28030 Then gError 500084
                        
            'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then gError 500083
                                
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
    
        Case 500081, 500082, 500083
        
        Case 500084
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156228)
    
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
        Case Produto.Name
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 35768

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                Produto.Enabled = True
            Else
                Produto.Enabled = False
            End If
        
        'Unidade de Medida
        Case UnidadeMed.Name

            UnidadeMed.Clear

            'Guarda a UM que está no Grid
            sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 35771

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                UnidadeMed.Enabled = False
            Else
                UnidadeMed.Enabled = True

                objProduto.sCodigo = sProdutoFormatado
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 35769
                If lErro = 28030 Then gError 35775 'Não achou

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lâ as Unidades de Medidas da Classe do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 35770
                'Carrega a combo de UM
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next
                'Seleciona na UM que está preenchida
                UnidadeMed.Text = sUM
                If Len(Trim(sUM)) > 0 Then
                    lErro = Combo_Item_Igual(UnidadeMed)
                    If lErro <> SUCESSO And lErro <> 12253 Then gError 35773
                End If
            End If
                
        'Nas demais
        Case TabPrecoItemNF.Name, ComissaoItemNF.Name, PrecoUnitario.Name, PercentDesc.Name, Desconto.Name, DescricaoItem.Name
            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 35772

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case Quantidade.Name

            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 46713

            If iProdutoPreenchido = PRODUTO_VAZIO Or left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case Else
            Call CF2(Me, "NF_Rotina_Grid_Enable", iLinha, objControl, iCaminho)
    
    End Select

    'Rastreamento
    lErro = gobjRastreamento.Rotina_Grid_Enable(iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 83327

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 35768, 35769, 35770, 35771, 35772, 35773, 46713, 83327

        Case 35775
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156229)

    End Select

    Exit Sub

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

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer, Optional bPergunta As Boolean = True)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer
Dim iIndice As Integer
Dim dValorProdutos As Double
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_GridItens_KeyDown
    
    'Guardo o item atual e o número de linhas existente
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iItemAtual = GridItens.Row

    If bPergunta Then
        Call Grid_Trata_Tecla1(KeyCode, objGridItens)
    Else
        Call Grid_Trata_Tecla3(KeyCode, objGridItens)
    End If
    
    'se alguma linha ofi excluída
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
        
        'Se existem alocacões no grid de alocações
        If objGridAlocacoes.iLinhasExistentes > 0 Then
            'Remove as alocações do item exluído
            Call Remove_Alocacoes(iItemAtual)
            'Atualiza o número do item nas alocações restantes
            For iIndice = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))) > 0 Then If CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) > iItemAtual Then GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col) = CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) - 1
            Next
        End If
        
        If objGridItens.iLinhasExistentes <> 0 Then
            Call PrecoTotal_Calcula(objGridItens.iLinhasExistentes, False)
        Else
            If StrParaDbl(ValorDesconto.Text) <> 0 Then
                'Avisa ao usuário
                vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", ValorDesconto.Text, 0)
                
                giValorDescontoManual = 1
                'Limpa o valor de desconto
                gdDesconto = 0
                Call CF2(Me, "NFCF_Limpa_Desconto")
                ValorDesconto.Text = ""
                iValorDescontoAlterado = 0
                
                Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, 0)

            End If
        End If
        
        'Recalcula o total da nota
        Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)

        ValorProdutos.Caption = Format(dValorProdutos - StrParaDbl(ValorDesconto.Text), "Standard")

        Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotalB_Col, dValorProdutos)

        ValorProdutos2.Caption = Format(dValorProdutos, "Standard")
        
        lErro = CF2(Me, "NFCF_ValorDesconto_Validate")
        If lErro <> SUCESSO Then gError 132022

''*** incluidos p/tratamento de tributacao *******************************
        lErro = gobjTribTab.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 39036
'*** fim tributacao

        Call ValorTotal_Calcula
    
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81522
    
        For iIndice = 1 To gobjNFiscal.ColItensNF.Count
            gobjNFiscal.ColItensNF(iIndice).iItem = iIndice 'Incluído para ao alterar a grade ou a reserva não fazer do item errado
        Next
        
    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr
    
        Case 39036, 81522, 83324
        
        Case 132022 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156230)
            
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

Private Sub Remove_Alocacoes(iItemAtual As Integer)
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

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 39662
        
        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridItens
            Case GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 35776

            'se for o GridComissoes
            Case GridComissoes.Name
                
                lErro = objTabComissoes.Saida_Celula_GridComissoes(objGridInt)
                If lErro <> SUCESSO Then gError 35777

            'Rastreamento
            Case GridRastro.Name

                lErro = gobjRastreamento.Saida_Celula()
                If lErro <> SUCESSO Then gError 83328

        End Select

        lErro = CF2(Me, "NF_Saida_Celula", objGridInt)
        If lErro <> SUCESSO Then gError 83328
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 35779

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 35776, 35777, 35779, 39662, 83328

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156231)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Produto
        Case iGrid_Produto_Col
            lErro = Saida_Celula_Produto(objGridInt)
            If lErro <> SUCESSO Then gError 26593
            
        'DescricaoItem
        Case iGrid_DescProduto_Col
            lErro = Saida_Celula_DescricaoItem(objGridInt)
            If lErro <> SUCESSO Then gError 42327

        'Unidade de Medida
        Case iGrid_UnidadeMed_Col
            lErro = Saida_Celula_UM(objGridInt)
            If lErro <> SUCESSO Then gError 26594

        'Quantidade
        Case iGrid_Quantidade_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 26595

        'PrecoUnitario
        Case iGrid_ValorUnitario_Col
            lErro = Saida_Celula_PrecoUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 26596

        'PercDesc
        Case iGrid_PercDesc_Col
            lErro = Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then gError 26599

        'Desconto
        Case iGrid_Desconto_Col
            lErro = Saida_Celula_Desconto(objGridInt)
            If lErro <> SUCESSO Then gError 42328

        Case iGrid_TabPrecoItem_Col
            If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
                lErro = Saida_Celula_TabPrecoItemNF(objGridInt)
                If lErro <> SUCESSO Then gError 42328
            End If

        Case iGrid_ComissaoItem_Col
            If gobjFAT.iExibeComissaoItensPV = MARCADO Then
                lErro = Saida_Celula_ComissaoItemNF(objGridInt)
                If lErro <> SUCESSO Then gError 42328
            End If

    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 26593, 26594, 26595, 26596, 26599, 42327, 42328

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156232)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto
    
    If Len(Trim(Produto.ClipText)) > 0 Then
        
        sProduto = Produto.Text

        lErro = CF("Trata_Segmento_Produto", sProduto)
        If lErro <> SUCESSO Then gError 199357

        Produto.Text = sProduto
        
        'Executa a saida de célula
        lErro = Produto_Saida_Celula()
        If lErro <> SUCESSO Then gError 35780

    End If

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 35781

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 35780, 35781, 199357
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156233)

    End Select

    Exit Function

End Function

Function Produto_Saida_Celula(Optional objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim vbMsgRes As VbMsgBoxResult
Dim objItemNF As ClassItemNF
Dim iPossuiGrade As Integer
Dim sProdutoPai As String
Dim iAlmoxarifadoPadrao As Integer
Dim colItensRomaneioGrade As New Collection
Dim objItensRomaneio As ClassItemRomaneioGrade
Dim sProduto As String
Dim objRomaneioGrade As New ClassRomaneioGrade
Dim dPercDescontoTab As Double, dPrecoComDescontoTab As Double

On Error GoTo Erro_Produto_Saida_Celula

    'Critica o Produto
    If objProduto Is Nothing Then
        
        Set objProduto = New ClassProduto
        
        'Critica o Produto
        lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 26657
        'Se o produto é gerencial sem ser de grade ==> Erro
        If lErro = 86295 And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then gError 86296
               
        'Se o produto não foi encontrado ==> Pergunta se deseja criar
        If lErro = 51381 Then gError 26658
        
        
    'incluido por tulio para poder ler a UM... 7/8/02
    Else

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 101782
        If lErro <> SUCESSO Then gError 101783
        
        If objProduto.iGerencial = GERENCIAL And objProduto.iKitVendaComp <> MARCADO Then gError 86296
        
    End If
    
    'Sinaliza que o produto possui grade
    If Len(Trim(objProduto.sGrade)) > 0 Then iPossuiGrade = MARCADO

    'Se não for um produto de grade
    If iPossuiGrade = DESMARCADO And objProduto.iKitVendaComp <> MARCADO Then
        
        'Se existir um produto pai de grade no grid
        If Grid_Possui_Grade Then
            
            'Busca, caso exista, o produto pai de grade o prod em questão
            lErro = CF("Produto_Le_PaiGrade", objProduto, sProdutoPai)
            If lErro <> SUCESSO Then gError 86327
            
            'Se o produto tem um pai de grade
            If Len(Trim(sProdutoPai)) > 0 Then
                'Verifica se seu pai aparece no grid
                For iIndice = 1 To gobjNFiscal.ColItensNF.Count
                    'Se aparecer ==> erro
                    If gobjNFiscal.ColItensNF(iIndice).sProduto = sProdutoPai Then gError 86328
                
                Next
            
            End If
            
        End If
        
        'Se o produto controlar estoque
        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
            
            lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifadoPadrao)
            If lErro <> SUCESSO And lErro <> 23796 Then gError 51420
            If lErro <> SUCESSO Or iAlmoxarifadoPadrao = 0 Then gError 51421
    
        End If
    Else
    
        If objProduto.iKitVendaComp <> MARCADO Then

            'Verifica se há filhos válidos com a grade preenchida
            lErro = CF("Produto_Le_Filhos_Grade", objProduto, colItensRomaneioGrade)
            If lErro <> SUCESSO Then gError 86329
        
        Else
        
            lErro = CF("ProdutoKitVenda_Le_Romaneio", objProduto, colItensRomaneioGrade)
            If lErro <> SUCESSO Then gError 177582
        
        End If
        
        'Se nao existir, erro
        If colItensRomaneioGrade.Count = 0 Then gError 86330
        
        'Para cada filho de grade do produto
        For Each objItensRomaneio In colItensRomaneioGrade
            'Verifica se ele já aparece no grid
            For iIndice = 1 To gobjNFiscal.ColItensNF.Count
                'Se aparecer ==> Erro
                If gobjNFiscal.ColItensNF(iIndice).sProduto = objItensRomaneio.sProduto Then gError 86331
            Next
        Next
 
    End If

'    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 26637

    Produto.PromptInclude = False
    Produto.Text = sProduto
    Produto.PromptInclude = True

'    End If

    'Verifica se já está em outra linha do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes
        If iIndice <> GridItens.Row Then
            If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then
                gError 26659
                Exit For
            End If
        End If
    Next
    
    'Verifica se é de Faturamento
    If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 35785

'    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
'    If lErro <> SUCESSO Then gError 39021
'
'    Produto.PromptInclude = False
'    Produto.Text = sProdutoEnxuto
'    Produto.PromptInclude = True
'
'    'Verifica se está no Grid
'    For iIndice = 1 To objGridItens.iLinhasExistentes
'        If iIndice <> GridItens.Row Then If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 35784
'    Next

    Set objItemNF = New ClassItemNF
    
    objItemNF.iPossuiGrade = iPossuiGrade

    If objItemNF.iPossuiGrade = MARCADO Then
        
        objItemNF.sProduto = objProduto.sCodigo
        objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
        objItemNF.iItem = GridItens.Row
        objItemNF.lNumIntDoc = 0
        objItemNF.sDescricaoItem = objProduto.sDescricao
                
        Set objRomaneioGrade = New ClassRomaneioGrade
        
        objRomaneioGrade.sNomeTela = Me.Name
        
        Set objRomaneioGrade.objObjetoTela = objItemNF
                    
        Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
        If giRetornoTela <> vbOK Then gError 86310

        
    End If
    
    'Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda
    'Descricao
    GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objProduto.sDescricao

    GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(gdDesconto, "Percent")

    lErro = CF2(Me, "NF_Produto_Saida_Celula", objProduto)
    If lErro <> SUCESSO Then gError 184279
    
    'Preco Unitário
    
    dPrecoUnitario = 0
    
    lErro = CF2(Me, "NFCF_PrecoUnitario_Calcula", giFilialEmpresa, objProduto.sCodigo, objProduto.sSiglaUMVenda, Cliente.Text, Codigo_Extrai(Filial.Text), Codigo_Extrai(TabelaPreco.Text), 0, MOEDA_REAL, objRomaneioGrade, dPrecoUnitario, False, dPercDescontoTab, dPrecoComDescontoTab)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    If dPrecoUnitario <> 0 Then
        GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
    Else
        GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = ""
    End If
        
    If dPercDescontoTab >= gdDesconto Then
        GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercDescontoTab, "Percent")
    End If
        
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        
        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItens.Row)

''*** incluidos p/tratamento de tributacao *******************************

        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 27832
        
        objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
        
        gobjNFiscal.ColItensNF.Add1 objItemNF
        gobjNFiscal.ColItensNF(GridItens.Row).sUMEstoque = objProduto.sSiglaUMEstoque
        gobjNFiscal.ColItensNF(GridItens.Row).iItem = GridItens.Row
        gobjNFiscal.ColItensNF(GridItens.Row).sProduto = objProduto.sCodigo
        
        '************** GRADE ************
        If iPossuiGrade = MARCADO Or objProduto.iKitVendaComp = MARCADO Then
            
            If objProduto.iKitVendaComp = MARCADO Then
                Set objItemNF.colItensRomaneioGrade = colItensRomaneioGrade
            End If
            
            Set gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade = objItemNF.colItensRomaneioGrade
            
            If objProduto.iKitVendaComp <> MARCADO Then

                gobjNFiscal.ColItensNF(GridItens.Row).iPossuiGrade = MARCADO
                   
                GridItens.TextMatrix(GridItens.Row, 0) = "# " & GridItens.TextMatrix(GridItens.Row, 0)
            
                Call Atualiza_Grid_Itens(objItemNF)
            
            End If
            
            Call Atualiza_Grid_Alocacao(objItemNF)
        End If
        
'*** fim tributacao

    End If

    Call PrecoTotal_Calcula(GridItens.Row, False)
    Call ValorTotal_Calcula
    
    Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_Produto_Saida_Celula:

    Produto_Saida_Celula = gErr

    Select Case gErr

        Case 26637, 26657, 27832, 51420, 86310, 86327, 86329, 101782, 35786, 177582, 184279, ERRO_SEM_MENSAGEM
        
        Case 26658
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)
                
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
            End If
            
        Case 35784
            'Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, Produto.Text, iIndice)

        Case 26659
            If gobjCRFAT.iProdRepetidoNF = PRODUTO_REPETICAO_NAO_PERMITE Then
                Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE3", gErr, Produto.Text, iIndice)
            ElseIf gobjCRFAT.iProdRepetidoNF = PRODUTO_REPETICAO_AVISA Then
                If Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_JA_EXISTENTE", Produto.Text, iIndice) = vbYes Then
                    Resume Next
                End If
            Else
                Resume Next
            End If
            
        Case 35785
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO", gErr)

        Case 39021
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 51421
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO", gErr, objProduto.sCodigo)

        Case 86296
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case 86328
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_GRID", gErr, Trim(sProdutoPai), Trim(Produto.Text))
                    
        Case 86330
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, Produto.Text)
        
        Case 86331
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILHO_GRADE_GRID", gErr, Trim(objProduto.sCodigo), Trim(gobjNFiscal.ColItensNF(iIndice).sProduto))
        
        Case 101783
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156234)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim iSairCelula As Integer
Dim dPrecoUnitario As Double
Dim dQuantidadeAnterior As Double
Dim objProduto As New ClassProduto
Dim objKitVenda As New ClassKitVenda
Dim objItemRomaneio As ClassItemRomaneioGrade

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
        If lErro <> SUCESSO Then gError 35810
        'Recolhe a quantidade informada
        dQuantidade = CDbl(Quantidade.Text)
        Quantidade = Formata_Estoque(dQuantidade)
        'Recolhe a quantidade anteriormente digitada
        dQuantidadeAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
        'Verifica se a quantidade foi alterada
        If dQuantidadeAnterior = dQuantidade Then iSairCelula = True

    End If
    
    'Se a quantidade foi alterada
'    If giTipoVersao = VERSAO_FULL Then
    
        If (Not iSairCelula) And Frame1(TAB_Almoxarifado).Enabled = True Then
            GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)
        
            objProduto.sCodigo = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 177585
        
            If objProduto.iKitVendaComp <> MARCADO Then
                'Tenta fazer uma alocação p\ oitem de acordo com a nova UM
                lErro = Alocacao_Processa()
                If lErro <> SUCESSO Then gError 35811
        
            Else
            
                objKitVenda.sProduto = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
                objKitVenda.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
                objKitVenda.dQuantidade = dQuantidade
                
                gobjNFiscal.ColItensNF.Item(GridItens.Row).sUnidadeMed = objKitVenda.sUM
                gobjNFiscal.ColItensNF.Item(GridItens.Row).dQuantidade = dQuantidade
            
                lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 177586
                            
                For Each objItemRomaneio In gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade
                    
                    lErro = Alocacao_Processa_ItemGrade(objItemRomaneio, gobjNFiscal.ColItensNF(GridItens.Row))
                    If lErro <> SUCESSO Then gError 177587
                
                Next
                
                Call Atualiza_Grid_Alocacao(gobjNFiscal.ColItensNF.Item(GridItens.Row))
    
            End If
            
        End If
    
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'        If (Not iSairCelula) And FrameL(TAB_Almoxarifado).Enabled = True Then
'            GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)
'
'            objProduto.sCodigo = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
'
'            lErro = CF("Produto_Le", objProduto)
'            If lErro <> SUCESSO And lErro <> 28030 Then gError 177585
'
'            If objProduto.iKitVendaComp <> MARCADO Then
'                'Tenta fazer uma alocação p\ oitem de acordo com a nova UM
'                lErro = Alocacao_Processa()
'                If lErro <> SUCESSO Then gError 35811
'
'            Else
'
'                objKitVenda.sProduto = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
'                objKitVenda.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
'                objKitVenda.dQuantidade = dQuantidade
'
'                gobjNFiscal.ColItensNF.Item(GridItens.Row).sUnidadeMed = objKitVenda.sUM
'                gobjNFiscal.ColItensNF.Item(GridItens.Row).dQuantidade = dQuantidade
'
'                lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade)
'                If lErro <> SUCESSO Then gError 177586
'
'                For Each objItemRomaneio In gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade
'
'                    lErro = Alocacao_Processa_ItemGrade(objItemRomaneio, gobjNFiscal.ColItensNF(GridItens.Row))
'                    If lErro <> SUCESSO Then gError 177587
'
'                Next
'
'                Call Atualiza_Grid_Alocacao(gobjNFiscal.ColItensNF.Item(GridItens.Row))
'
'            End If
'
'        End If
'    End If
    
    If (Not iSairCelula) Then
    
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81523
            
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 35812
    
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))

    If dPrecoUnitario > 0 Then Call PrecoTotal_Calcula(GridItens.Row, False)
    
    '*** incluidos p/tratamento de tributacao ****************
    If GridItens.Row <= objGridItens.iLinhasExistentes Then
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 39037
    End If
    '*** fim tributacao

    'Atualiza o valor total
    Call ValorTotal_Calcula

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 35810, 35812, 35811, 81523, 177585, 177586, 177587
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 35811
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156235)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Alocacao_Processa_Todos() As Long
'Faz a alocacao para todas as linhas do grid

Dim iGridRow As Integer
Dim iIndice As Integer
Dim lErro As Long
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_Alocacao_Processa_Todos

    'Guarda a Linha do Grid Corrente
    iGridRow = objGridItens.objGrid.Row
    
    'Para cada linha do grid
    For iIndice = 1 To objGridItens.iLinhasExistentes
        
        'Indica a linha corrente
        objGridItens.objGrid.Row = iIndice
        
        If gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade.Count = 0 Then
            'Processa a alocacao para esta linha
            lErro = Alocacao_Processa()
            If lErro <> SUCESSO Then gError 101320
        Else
        
            For Each objItemRomaneio In gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade
                'Processa a alocacao para esta linha
                lErro = Alocacao_Processa_ItemGrade(objItemRomaneio, gobjNFiscal.ColItensNF(iIndice))
                If lErro <> SUCESSO Then gError 101320
            Next
        
            Call Atualiza_Grid_Alocacao(gobjNFiscal.ColItensNF.Item(iIndice))
        
        End If
    Next
    'Volta com a linha inicial
    objGridItens.objGrid.Row = iGridRow
    
    Alocacao_Processa_Todos = SUCESSO
    
    Exit Function

Erro_Alocacao_Processa_Todos:

    Alocacao_Processa_Todos = gErr
    
    Select Case gErr
    
        Case 101320
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156236)

    End Select

End Function

Private Function Alocacao_Processa() As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dQuantidade As Double
Dim sUM As String

On Error GoTo Erro_Alocacao_Processa

    'Remove as alocações do grid de alocações
    Call Remove_Alocacoes(GridItens.Row)

    'Formata o Produto
    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 35813

    objProduto.sCodigo = sProduto

    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 35814
    If lErro = 28030 Then gError 35815 'Não encontrou

    'recolhe a quantidade do grid de itens
    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))

    'Verifica se o produto tem o controle de estoque <> PRODUTO_CONTROLE_SEM_ESTOQUE
    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 And gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA Then
        'recolhe a UM do ItemNF
        sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

        'Tenta Alocar o produto no Almoxarifado padrão
        lErro = AlocaAlmoxarifadoPradrao(dQuantidade, objProduto, sUM)
        If lErro <> SUCESSO Then gError 35816
    
    End If

    Alocacao_Processa = SUCESSO

    Exit Function

Erro_Alocacao_Processa:

    Alocacao_Processa = gErr

    Select Case gErr

        Case 35813, 35814, 35816

        Case 35815
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156237)

    End Select

    Exit Function

End Function

Function AlocaAlmoxarifadoPradrao(dQuantidade As Double, objProduto As ClassProduto, sUM As String) As Long
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
Dim iLinhas As Integer

On Error GoTo Erro_AlocaAlmoxarifadoPradrao

    'Faz a conversão da UM da tela para a UM de estoque
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 35817

    'Converte a quantidade para a UM de estoque
    dQuantAlocar = dQuantidade * dFator

    lErro = CF2(Me, "NF_AlocaAlmoxarifadoPadrao", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
    If lErro <> SUCESSO Then gError 141067

    objAlmoxarifado.iCodigo = iAlmoxarifado
    
    'Lê o Aloxarifado
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> 25056 And lErro <> SUCESSO Then gError 35823
    If lErro = 25056 Then gError 35824
    
    objEstoqueProduto.sProduto = objProduto.sCodigo
    objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
    
    'Le os estoques desse produto nesse almoxarifado
    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 21306 Then gError 35818
    
    If lErro = 21306 Then gError 35819 'Não encontrou

    iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Seleciona a origem da quantidade disponível
    Select Case iCodigo
    
        'Se o tipo da nota for cobrança de mat. consignado
        Case DOCINFO_NFFISPC
            
            'A quantidade disponível deve ser igual a quantidade no escaninho de mat. em Consignação (Consig)
            dQuantDisponivel = objEstoqueProduto.dQuantConsig
        
        'Se o tipo da nota for mat. beneficiado de 3º´s
        Case DOCINFO_NFISBF
        
            'A quantidade disponível deve ser igual a quantidade no escaninho Mat.de 3º´s em Beneficiamento (Benef3)
            dQuantDisponivel = objEstoqueProduto.dQuantBenef3
            
        'Se for outro tipo de nota
        Case Else
            
            'A quantidade disponível deve ser igual a quantidade do escaninho mat. nosso disponível (DispNossa)
            dQuantDisponivel = objEstoqueProduto.dQuantDisponivel
    
    End Select

    'Remove as alocações do produto do grid de alocações
    Call Remove_Alocacoes(GridItens.Row)

    iLinhas = objGridAlocacoes.iLinhasExistentes

    dQuantAlocar = Arredonda_Estoque(dQuantAlocar)
    
    'Verifica se a Quantidade disponível é maior que a quantidade a alocar
    If (dQuantAlocar - dQuantDisponivel) < QTDE_ESTOQUE_DELTA Then

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

            'Recolhe todos os produtos dos outros itens
            For iIndice = 1 To objGridItens.iLinhasExistentes
                If iIndice <> GridItens.Row Then
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
                    If lErro <> SUCESSO Then gError 35826
                    'Adiciona na coleção de produtos
                    colOutrosProdutos.Add sProduto
                End If
            Next

            'Chama a tela de Localização de Produto
            Call Chama_Tela_Modal("LocalizacaoProduto1", objItemNF, colOutrosProdutos, dQuantAlocar, iCodigo)
            If giRetornoTela = vbCancel Then gError 35825 'Se nada foi feito lá
            If giRetornoTela = vbOK Then

                sProdutoEnxuto = String(STRING_PRODUTO, 0)

                lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
                If lErro <> SUCESSO Then gError 35820

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
                    GridAlocacao.TextMatrix(iLinhas, iGrid_UMAloc_Col) = objItemNF.sUMEstoque
                    GridAlocacao.TextMatrix(iLinhas, iGrid_QuantVend_Col) = Formata_Estoque(dTotal)

                Next
                
                'Quantidade.Text = Formata_Estoque(dTotal)
                If Abs(dQuantAlocar - dTotal) > QTDE_ESTOQUE_DELTA Then
                    Quantidade.Text = Formata_Estoque(dTotal / dFator)
                Else
                    Quantidade.Text = Formata_Estoque(StrParaDbl(Quantidade.Text))
                End If
                
                'Se o produto foi substituido
                If objProduto.sCodigo <> objItemNF.sProduto Then
                    'Coloca o produto no Grid de itens
                    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
                    'Chama a rotina que trata um novo produto
                    lErro = Produto_Saida_Celula()
                    objGridAlocacoes.iLinhasExistentes = iLinhas
                    gError 35821
                End If
            End If
        End If
    End If
    
    'atualiza o numero de linhas existentes no grid de alocações
    objGridAlocacoes.iLinhasExistentes = iLinhas

    AlocaAlmoxarifadoPradrao = SUCESSO

    Exit Function

Erro_AlocaAlmoxarifadoPradrao:

    AlocaAlmoxarifadoPradrao = gErr

    Select Case gErr

        Case 35817, 35818, 35822, 35823, 35826, 35821, 51619, 141067 'Inserido por Wagner

        Case 35819
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case 35820
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)
        
        Case 35824
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, iAlmoxarifado)

        Case 35825
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_LOCALIZACAO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156238)

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

    'Verifica se tem alguma preço digitado
    If Len(Trim(PrecoUnitario.ClipText)) > 0 Then

        'if iCodigo = DOCINFO_NFISICM Or iCodigo = DOCINFO_NFISIPI Then
        If giComplemento = MARCADO Then
            lErro = Valor_NaoNegativo_Critica(PrecoUnitario.Text)
        Else
            lErro = Valor_Positivo_Critica(PrecoUnitario.Text)
        End If
        If lErro <> SUCESSO Then gError 35827
        
        PrecoUnitario.Text = Format(PrecoUnitario.Text, gobjFAT.sFormatoPrecoUnitario)
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 35828

    Call PrecoTotal_Calcula(GridItens.Row, False)
    
''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 39038
'*** fim tributacao

    Call ValorTotal_Calcula

    Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = gErr

    Select Case gErr

        Case 39038, 35827, 35828, 87188
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156239)

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
Dim lTamanho As Long
Dim dPercentDescAnterior As Double

On Error GoTo Erro_Saida_Celula_PercentDesc

    Set objGridInt.objControle = PercentDesc

    dPercentDesc = StrParaDbl(PercentDesc.Text)
    
    lTamanho = Len(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col))
    If lTamanho > 0 Then dPercentDescAnterior = StrParaDbl(left(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col), lTamanho - 1))

    'verifica se o percentual está preenchido
    If Len(Trim(PercentDesc.Text)) > 0 Then
    
        'Critica a procentagem
        lErro = Porcentagem_Critica(PercentDesc.Text)
        If lErro <> SUCESSO Then gError 35829
        
        If dPercentDesc <> dPercentDescAnterior Then
        
            'Verifica se o percentual é de 100%
            If dPercentDesc = 100 Then gError 35831
    
            PercentDesc.Text = Format(dPercentDesc, "Fixed")
    
        End If
    Else
        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 35830

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) > 0 Then dPrecoUnitario = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))

    If Abs(dPercentDesc - dPercentDescAnterior) > DELTA_VALORMONETARIO2 Then
        iDescontoAlterado = REGISTRO_ALTERADO
    End If

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

        Case 39039, 35829, 35830
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 35831
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156240)

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
        If lErro <> SUCESSO Then gError 35832
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
            If dDesconto >= dPrecoTotal Then gError 35833
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
        If lErro <> SUCESSO Then gError 39040
        '*** fim tributacao
        
        Call ValorTotal_Calcula
    
    End If
    
    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = gErr

    Select Case gErr

        Case 39040, 35832
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 35833
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCONTO_MAIOR_OU_IGUAL_PRECO_TOTAL", gErr, GridItens.Row, dDesconto, dPrecoTotal)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156241)

    End Select

    Exit Function

End Function

Public Sub ValorFrete_Validate(Cancel As Boolean)

Dim dValorFrete As Double
Dim lErro As Long

On Error GoTo Erro_ValorFrete_Validate

    If iValorFreteAlterado = 0 Then Exit Sub

    'Verifica se o Valor Frete foi preenchido
    If Len(Trim(ValorFrete.Text)) > 0 Then
        'faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
        If lErro <> SUCESSO Then gError 35835

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

        Case 35835

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156242)

    End Select

    Exit Sub

End Sub

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
            'Tenta fazer uma alocação p\ oitem de acordo com a nova UM
            lErro = Alocacao_Processa()
            If lErro <> SUCESSO Then gError 42195
    
        Else
                        
            If objProduto.iKitVendaComp = MARCADO Then
                
                objKitVenda.sProduto = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
                objKitVenda.sUM = UnidadeMed.Text
                objKitVenda.dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
            
                lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 177584
            
            End If
                            
            For Each objItemRomaneio In gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade
                lErro = Alocacao_Processa_ItemGrade(objItemRomaneio, gobjNFiscal.ColItensNF(GridItens.Row))
                If lErro <> SUCESSO Then gError 42195
            Next
            
            Call Atualiza_Grid_Alocacao(gobjNFiscal.ColItensNF.Item(GridItens.Row))

        End If
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42194
    
    'Se a Um selecionada agora é diferente da anterior
    If sUmAnterior <> UnidadeMed.Text Then

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 500092
    
    End If
    
    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case 42194, 42195, 500092, 177583, 177584
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156243)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescricaoItem(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoItem

    Set objGridInt.objControle = DescricaoItem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42329
    
    Saida_Celula_DescricaoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoItem:

    Saida_Celula_DescricaoItem = gErr

    Select Case gErr

        Case 42329
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156244)

    End Select

    Exit Function

End Function

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

Public Sub ValorDespesas_Validate(Cancel As Boolean)

Dim dValorDespesas As Double
Dim lErro As Long

On Error GoTo Erro_ValorDespesas_Validate

    If iValorDespesasAlterado = 0 Then Exit Sub
    
    'Se  estiver preenchido
    If Len(Trim(ValorDespesas.Text)) > 0 Then

        'Faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorDespesas.Text)
        If lErro <> SUCESSO Then gError 35885

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

        Case 35885

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156245)

    End Select

    Exit Sub

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

        Case 35891

        Case 35892
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156246)

    End Select

    Exit Function

End Function

Public Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Function PrecoUnitario_Calcula(sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long

Dim objProduto As New ClassProduto
Dim objUM As New ClassUnidadeDeMedida
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim lErro As Long
Dim objCotacao As New ClassCotacaoMoeda, objTabelaPreco As New ClassTabelaPreco
Dim objCotacaoAnterior As New ClassCotacaoMoeda

On Error GoTo Erro_PrecoUnitario_Calcula

    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto
    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 35901
    If lErro = 28030 Then gError 35902 'Não encontrou

    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 35900
    
    'Converte o preço unitário para UM da tela
    dPrecoUnitario = objTabelaPrecoItem.dPreco / dFator
    
    If objTabelaPrecoItem.iCodTabela <> 0 Then
    
        objTabelaPreco.iCodigo = objTabelaPrecoItem.iCodTabela
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 35901
        If lErro = SUCESSO And objTabelaPreco.iMoeda <> MOEDA_REAL Then
            
            objCotacao.dtData = gdtDataAtual
            objCotacao.iMoeda = objTabelaPreco.iMoeda
            objCotacaoAnterior.iMoeda = objTabelaPreco.iMoeda
        
            'Chama função de leitura
            lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
            If lErro <> SUCESSO Then gError 35901
            
            'Se nao existe cotacao para a data informada => Mostra a última.
            dPrecoUnitario = dPrecoUnitario * StrParaDbl(IIf(objCotacao.dValor <> 0, Format(objCotacao.dValor, "#.0000"), Format(objCotacaoAnterior.dValor, "#.0000")))
        
        End If
    
    End If
    
    PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_PrecoUnitario_Calcula:

    PrecoUnitario_Calcula = gErr

    Select Case gErr

        Case 35900, 35901

        Case 35902
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objTabelaPrecoItem.sCodProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156247)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 35903

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscal1()
    If lErro <> SUCESSO Then gError 500118
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 35903, 500118

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156248)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iIndice As Integer
Dim dValorTotal As Double
Dim objNFiscal As New ClassNFiscal
Dim dValorDesconto As Double
Dim dFator As Double
Dim iTipoNFiscal As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objProduto As New ClassProduto
Dim iCodigo As Integer
Dim iNFSE As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 35904
    If Len(Trim(gobjTribTab.sNatOpInterna)) = 0 Then gError 35905
    If Len(Trim(Cliente.ClipText)) = 0 Then gError 35906
    If Len(Trim(Filial.Text)) = 0 Then gError 35907
    If Len(Trim(Serie.Text)) = 0 Then gError 35908
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 35911
    If Len(Trim(FilialEntrega.Text)) = 0 Then gError 51532

    If Len(Trim(PesoLiquido.ClipText)) > 0 And Len(Trim(PesoBruto.ClipText)) > 0 Then
        If CDbl(PesoLiquido.Text) > CDbl(PesoBruto.Text) Then gError 35925
    End If
        
    'Verifica se a Data de Saida e menor que a Data de Emissão
    If Len(Trim(DataSaida.ClipText)) <> 0 Then
        If CDate(DataSaida.Text) < CDate(DataEmissao.Text) Then gError 35912
    End If
    
    dValorTotal = StrParaDbl(ValorTotal.Caption)

    'Se o total for negativo --> Erro
    If dValorTotal < 0 Then gError 35924

    lErro = gobjTribTab.Valida_Dados()
    If lErro <> SUCESSO Then gError 56848
    
    'Verifica se o Valor do Desconto é maior que a soma dos demais valores
    If Len(Trim(ValorDesconto.Text)) > 0 Then dValorDesconto = CDbl(ValorDesconto.Text)
    lErro = DescontoGlobal_Testa(dValorDesconto)
    If lErro <> SUCESSO Then gError 56848
    
    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 35913
    
    If lErro = 31415 Then gError 35914 'Não encontrou

    If objTipoDocInfo.iComplementar = DOCINFO_NORMAL Then
        If objGridItens.iLinhasExistentes = 0 Then gError 35915
    End If
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 35916
        
        'If objTipoDocInfo.iCodigo <> DOCINFO_NFISICM And objTipoDocInfo.iCodigo <> DOCINFO_NFISIPI Then
        If objTipoDocInfo.iComplementar = DOCINFO_NORMAL Then
            If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col)) = 0 Then gError 35921
        End If
    Next
            
    'Valida os dados do grid de comissões
    lErro = objTabComissoes.Valida_Grid_Comissoes()
    If lErro <> SUCESSO Then gError 35928

    lErro = gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
    If lErro <> SUCESSO Then gError 83329

    Set objNFiscal.objRastreamento = gobjRastreamento

    'Recolhe os dados da tela
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 35926

    'Se é uma nota de complemento ou se é uma nota de de cobrança de material consignado
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Or objTipoDocInfo.iCodigo = DOCINFO_NFISPC Or objTipoDocInfo.iCodigo = DOCINFO_NFISVFE Then
        
        'Recolhe os dados da Nota Fiscal Original
        lErro = Move_Dados_NFOriginal_Memoria(objNFiscal, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 35927
        
    End If

    lErro = Move_Grids_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 35932
    
    For iIndice = 1 To objNFiscal.ColItensNF.Count
        'Se for uma nota fiscal de serviço
        If objNFiscal.iTipoNFiscal = DOCINFO_NFISS Then
                    
            objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 23080 Then gError 62099
            If lErro <> SUCESSO Then gError 62100
        
            If Not (objProduto.iCompras = PRODUTO_PRODUZIVEL And objProduto.iFaturamento = PRODUTO_VENDAVEL And objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE) Then gError 62101
        
        End If
    Next
    
    lErro = gobjTribTab.Valida_TributacaoNF_Gravacao(objNFiscal)
    If lErro <> SUCESSO Then gError 51102
    
    lErro = Calcula_Msg
    If lErro <> SUCESSO Then gError 51102
    
    'tulio160103
    'se for pra verificar credito
    If gobjCRFAT.iCreditoVerificaLimite = VERIFICAR_LIMITE_CREDITO Then
            
        If Len(Trim(NFiscal.Caption)) = 0 Then
            'verifica se o cliente tem crédito.
            lErro = CF("NFiscal_Testa_Credito", objNFiscal)
            If lErro <> SUCESSO Then gError 44459
        End If
        
    End If
    
    
    
    'mover a parte do rastreamento
    lErro = gobjRastreamento.Move_Rastro_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 83330
    
    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataEmissao.Text))
    If lErro <> SUCESSO Then gError 92042
    
    lErro = CF("NFiscal_Valida_MargContr", objNFiscal)
    If lErro <> SUCESSO Then gError 201042
    
    lErro = CF("NFiscal_Valida_Diversos", objNFiscal)
    If lErro <> SUCESSO Then gError 201042
    
    'Grava a Nota Fiscal de Saída(inclusive os dados contábeis)
    lErro = CF("NFiscal_Grava", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 35936
    
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
        If lErro <> SUCESSO Then gError 199994
    End If
    
    lErro = CF("Verifica_NFiscal_Servico_Eletronica", objNFiscal, iNFSE)
    If lErro <> SUCESSO Then gError 207060
    
    If iNFSE = 1 Then
        lErro = CF("NFSE_Grava1", objNFiscal.lNumIntDoc, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 207061
    End If
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 56848, 62099, 83329, 83330, 92042, 130406, 181786, 199994, 201042, 207060, 207061
                
        Case 35904
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 35905
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 35906
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 35907
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 35908
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 35910
            Call Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_NAO_PREENCHIDA", gErr)

        Case 35911
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 35912
            Call Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_ANTERIOR_DATAEMISSAO", gErr, DataSaida, DataEmissao)

        Case 35913, 35917, 35918, 35926, 35927, 35928, 35932, 35936, 43053, 44459, 51102, 92042

        Case 35914
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iTipo)

        Case 35915
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)

        Case 35916
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case 35921
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 35924
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)

        Case 35925
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr, CDbl(PesoLiquido.Text), CDbl(PesoBruto.Text))

        Case 51532
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)
        
        Case 62100
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                    
        Case 62101
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NOTA_SERVICO", gErr, GridItens.TextMatrix(iIndice, iGrid_Produto_Col))
        
        Case 87189
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156249)

    End Select

    Exit Function

End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objNFiscal As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objCliente As New ClassCliente
Dim snProt As String
Dim iEnviado As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 83981
    If Len(Trim(Serie.Text)) = 0 Then gError 83982
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 83983

    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
    
    'le o tipo da origem
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 99345
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 99346
    
    'se origem for de empresa é interna
    If objTipoDocInfo.iOrigem = DOCINFO_EMPRESA Then
        'Verifica se a existe nota fiscal está cadastrada
        lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
        If lErro <> SUCESSO And lErro <> 43676 Then gError 99352
        
        If lErro = 43676 Then gError 89065
        
'        If gobjCRFAT.iUsaNFe = MARCADO And objNFiscal.iRecibo = 0 Then
'
'            lErro = CF("NFeFedLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'            If lErro <> SUCESSO Then gError 203012
'
'            If iEnviado = 1 Then gError 203013
'
'        ElseIf gobjCRFAT.iUsaNFSE = MARCADO And objNFiscal.iRecibo <> 0 Then
'
'            lErro = CF("RPSWEBLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'            If lErro <> SUCESSO Then gError 207081
'
'            If iEnviado = 1 Then gError 207082
'
'        End If
        
    Else
       
        If Len(Trim(Cliente.ClipText)) = 0 Then gError 99347
        
        objCliente.sNomeReduzido = Cliente.Text
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 99348
        
        If lErro = 12348 Then gError 99349

        objNFiscal.lCliente = objCliente.lCodigo
        objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
                
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
        If lErro <> SUCESSO Then gError 83986

        'Faz o cancelamento de uma nota fiscal de Saida
        lErro = CF("NotaFiscalSaida_Excluir", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError 83987
    
        'Limpa a Tela
        lErro = Limpa_Tela_NFiscal1()
        If lErro <> SUCESSO Then gError 83988
    
        'fecha o comando de setas
        Call ComandoSeta_Fechar(Me.Name)

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 83981
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 83982
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 83983
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 84984, 83986, 83987, 83988, 99345, 99348, 99352, 99353, 203012, 207081
       
        Case 89065
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
            
        Case 99346
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)
        
        Case 99347
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)
        
        Case 99349
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objNFiscal.lCliente)
            
        Case 203013
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CADASTRADA_NAO_EXCLUI", gErr)

        Case 207082
            Call Rotina_Erro(vbOKOnly, "ERRO_NFSE_CADASTRADA_NAO_EXCLUI", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156250)
        
    End Select

    Exit Sub

End Sub

Private Function Move_Grids_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os dados dos Grids para memória

Dim lErro As Long

On Error GoTo Erro_Move_Grids_Memoria

    'Chama Move_GridItens_Memoria
    lErro = Move_GridItens_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 35933

'    If giTipoVersao = VERSAO_FULL Then
        If Frame1(TAB_Almoxarifado).Enabled = True Then
            'Chama Move_GridAlocacao_Memoria
            lErro = Move_GridAlocacao_Memoria(objNFiscal)
            If lErro <> SUCESSO Then gError 35937
        End If
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'        If FrameL(TAB_Almoxarifado).Enabled = True Then
'            'Chama Move_GridAlocacao_Memoria
'            lErro = Move_GridAlocacao_Memoria(objNFiscal)
'            If lErro <> SUCESSO Then gError 35937
'        End If
'    End If

    'Chama Move_GridComissoes_Memoria
    lErro = objTabComissoes.Move_TabComissoes_Memoria(objNFiscal, NOTA_FISCAL)
    If lErro <> SUCESSO Then gError 35938

    'mover a parte de tributacao
    lErro = gobjTribTab.PreencherNF(objNFiscal)
    If lErro <> SUCESSO Then gError 39041
    
    Move_Grids_Memoria = SUCESSO

    Exit Function

Erro_Move_Grids_Memoria:

    Move_Grids_Memoria = gErr

    Select Case gErr

        Case 39041, 35933, 35937, 35938

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156251)

    End Select

    Exit Function

End Function

Public Function Move_GridItens_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os Itens do Grid para a Memória

Dim iIndice As Integer
Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lTamanho As Long
Dim colAlocacoes As ColAlocacoesItemNF
Dim objProduto As New ClassProduto
Dim objItemNF1 As ClassItemNF
Dim iItem As Integer

On Error GoTo Erro_Move_GridItens_Memoria

    'Para cada linha existente do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemNF = New ClassItemNF

        'Verifica se o Produto está preenchido
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 35934

        'Armazena produto
        objItemNF.sProduto = sProdutoFormatado
                
        'Guarda os demais campos do Grid em objItemNF
        objItemNF.sDescricaoItem = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
        objItemNF.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        objItemNF.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objItemNF.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))
        lTamanho = Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col)))

        If lTamanho > 0 Then objItemNF.dPercDesc = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col))

        objItemNF.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))

        'Adiciona na coleção de itens
        With objItemNF
            Set objItemNF1 = objNFiscal.ColItensNF.Add(0, iIndice, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, DATA_NULA, .sDescricaoItem, 0, 0, 0, 0, 0, colAlocacoes, 0, "", "", STATUS_LANCADO, 0, "", 0, 0, 0, "", .iClasseUM, 0)
        End With

        '********************* TRATAMENTO DE GRADE *****************
        Call Move_ItensGrade_Tela(objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade, gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade)

        Set objNFiscal.ColItensNF(iIndice).objInfoAdicDocItem = gobjNFiscal.ColItensNF(iIndice).objInfoAdicDocItem
 
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
   
   Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 35934, 35935, 45339
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156252)

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
            
            If colCodigoNome.Count > 0 Then
                'Verifica se o Almoxarifado já foi lido
                For Each objCodigoNome In colCodigoNome
                    
                    If objCodigoNome.sNome = objItemAloc.sAlmoxarifado Then
                        objItemAloc.iAlmoxarifado = objCodigoNome.iCodigo
                        Exit For
                    End If
                Next
            End If
            'Se o almoxarifado ainda não foi lido
            If objItemAloc.iAlmoxarifado = 0 Then
                objAlmoxarifado.sNomeReduzido = objItemAloc.sAlmoxarifado
                'Lê o Almoxarifado à partir do nome reduzido
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25060 Then gError 35939
                If lErro = 25060 Then gError 35940
                'Adiciona o almoxarifado na coleção de almoxarifados lidos
                objItemAloc.iAlmoxarifado = objAlmoxarifado.iCodigo

                colCodigoNome.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido
                
            End If
            'Adiciona a alocação do grid na coleção de itens
            With objItemAloc
                Set objItemAloc2 = objNFiscal.ColItensNF(iItem).colAlocacoes.Add(.iAlmoxarifado, .sAlmoxarifado, .dQuantidade)
                objItemAloc2.sUnidadeMed = .sUnidadeMed
            End With

            objNFiscal.ColItensNF(iItem).iAlmoxarifado = objItemAloc.iAlmoxarifado
                
        Next
        
    End If

    Move_GridAlocacao_Memoria = SUCESSO

    Exit Function

Erro_Move_GridAlocacao_Memoria:

    Move_GridAlocacao_Memoria = gErr
    
    Select Case gErr

        Case 35939

        Case 35940
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156253)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_NFiscal() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Limpa_Tela_NFiscal

    'Indica que está limpando a tela
    gbLimpandoTela = True
    
    Call Limpa_Tela(Me)
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    Set gobjAnotacao = Nothing
    
    Set gobjNFiscal = New ClassNFiscal
    Set gobjInfoAdic = Nothing
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next
    
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridAlocacoes)
    objTabComissoes.Limpa_Grid_Comissoes
        
    TipoNFiscal.Text = ""
    
    Status.Caption = ""
    
'    If giTipoVersao = VERSAO_FULL Then
        Frame1(TAB_Almoxarifado).Enabled = True
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'        FrameL(TAB_Almoxarifado).Enabled = True
'    End If
    
'    If giTipoVersao = VERSAO_FULL Then
        Serie.Text = ""
        Call Serie_Validate(bSGECancelDummy)
        TabelaPreco.Text = ""
'    End If
    
    NFiscal.Caption = ""
    Filial.Clear
    FilialEntrega.Clear
    FilialCliIntermediario.Clear
    ValorProdutos.Caption = ""
    ValorProdutos2.Caption = ""
    ValorTotal.Caption = ""
    Transportadora.Text = ""
    TranspRedespacho.Text = ""
    RedespachoCli.Value = False
    PlacaUF.Text = ""
    EletronicaFed.Value = vbUnchecked
    
    VolumeMarca.Text = ""
    VolumeEspecie.Text = ""
    
    'Recarrega a combo VolumeEspecie e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)

    'Recarrega a combo VolumeMarca e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
    
    CanalVenda.Text = ""
    SerieNFiscalOriginal = ""
    TotalPercentualComissao.Caption = ""
    TotalValorComissao.Caption = ""
    TotalValorBase.Caption = ""
    TipoFrete.ListIndex = 0

'    For iIndice = 1 To objGridItens.iLinhasExistentes
'        GridItens.TextMatrix(iIndice, 0) = iIndice
'    Next
'
'    Call Grid_Limpa(objGridItens)
'    Call Grid_Limpa(objGridAlocacoes)
'    Call Grid_Limpa(objGridComissoes)

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.Limpa_Tela
'*** fim tributacao
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    Call DataEmissao_Validate(bSGECancelDummy)

    Call CF2(Me, "NFCF_Default_DataSaida")
    
    'Seleciona série padrão
    lErro = Seleciona_Serie_Padrao()
    If lErro <> SUCESSO Then gError 500115
    
    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
    'Daniel 23/05/2002
    FilialOrcamento.Text = ""
    
    'Fecha o Sistema de Setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iValorDescontoAlterado = 0
    iComissoesAlterada = 0
    giFilialEntregaAlterada = 0
    giFilialEntregaManual = 0
    iClienteAlterado = 0
    gdDesconto = 0
    Call CF2(Me, "NFCF_Limpa_Desconto")
    iValorFreteAlterado = 0
    iValorSeguroAlterado = 0
    iValorDespesasAlterado = 0
    
    Recibo.Value = vbUnchecked
    
    NumNFe.Caption = ""
    CodVerificacao.Caption = ""
    
    Call CF2(Me, "NFCF_Limpa_Tela")
    
    Call CF2(Me, "NF_Limpa_Tela_Cust")

    MsgAutomatica.Value = vbChecked
    
    iAlterado = 0
    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0
    
    'Indica que não está mas limpando a tela
    gbLimpandoTela = False
    
    Exit Function

Erro_Limpa_Tela_NFiscal:
    
    Limpa_Tela_NFiscal = gErr
    
    Select Case gErr
        
        Case 500115
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156254)
    
    End Select
    
    'Indica que não está mas limpando a tela
    gbLimpandoTela = False
    
    Exit Function
    
End Function

Private Function Limpa_Tela_NFiscal1() As Long
'Limpa a Tela NFiscalEntrada, mas mantém a natureza e o tipo
    
Dim sNatureza As String
Dim sTipoNFiscal As String
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_NFiscal1

    'Indica que está limpando a tela
    gbLimpandoTela = True
    
    sNatureza = gobjTribTab.sNatOpInterna
    sTipoNFiscal = TipoNFiscal.Text
    
    lErro = Limpa_Tela_NFiscal()
    If lErro <> SUCESSO Then gError 500117
    
    gobjTribTab.sNatOpInterna = sNatureza
    TipoNFiscal.Text = sTipoNFiscal
        
    'Indica que não está mas limpando a tela
    gbLimpandoTela = False
    
    Exit Function

Erro_Limpa_Tela_NFiscal1:
    
    Limpa_Tela_NFiscal1 = gErr
    
    Select Case gErr
    
        Case 500117
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156255)
    
    End Select
    
    'Indica que não está mas limpando a tela
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
Dim objProduto As New ClassProduto
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim iItem As Integer

On Error GoTo Erro_Preenche_GridItens

    iIndice = 0

    'Para cada item da Coleção
    For Each objItemNF In objNFiscal.ColItensNF

        iIndice = iIndice + 1
        objItemNF.iItem = iIndice

        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 35943

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
        GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objItemNF.dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objItemNF.dValorTotal, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col) = Format(objItemNF.dValorTotal + objItemNF.dValorDesconto, "Standard")
        
        If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        
            If objItemNF.dComissao <> -1 Then
                GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col) = Format(objItemNF.dComissao, "Percent")
            Else
                GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col) = ""
            End If
        
        End If
        
        lErro = CF2(Me, "NF_Preenche_GridItens", objNFiscal, objItemNF, objProduto, iIndice)
        If lErro <> SUCESSO Then gError 177926
        
    Next

    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objNFiscal.objProjetoInfo, objNFiscal.lNumIntDoc, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181676
    
    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice
    
    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr

        Case 35943
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)
        
        Case 35944, 177926, 181676

        Case 35945
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156256)

    End Select

    Exit Function

End Function

Private Function Move_Dados_NFOriginal_Memoria(objNFiscal As ClassNFiscal, objTipoDocInfo As ClassTipoDocInfo) As Long
'Move os dados da NotaFiscalOriginal para a memória

Dim lErro As Long
Dim objNFOriginal As New ClassNFiscal
Dim vbResult As VbMsgBoxResult
Dim bCamposPreenchidos As Boolean

On Error GoTo Erro_Move_Dados_NFOriginal_Memoria

    bCamposPreenchidos = True

    'Verifica se a Série e o Número da NF original estão preenchidos
    If Len(Trim(NFiscalOriginal.ClipText)) = 0 And ISSerieEletronica(objNFiscal.sSerie) Then gError 35946
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 And ISSerieEletronica(objNFiscal.sSerie) Then gError 35947

    'Verifica se a Serie e o Número da NF original estão preenchidos
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Then  'gError 35947
        bCamposPreenchidos = False
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_SERIE_NFISCAL_ORIGINAL_NAO_PREENCHIDA")
        If vbResult = vbNo Then gError 141528
    End If
    
    If Len(Trim(NFiscalOriginal.ClipText)) = 0 And bCamposPreenchidos Then 'gError 35946
        bCamposPreenchidos = False
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_NUMERO_NFISCAL_ORIGINAL_NAO_PREENCHIDO")
        If vbResult = vbNo Then gError 141529
    End If

    If bCamposPreenchidos Then
    
        objNFOriginal.lCliente = objNFiscal.lCliente
        objNFOriginal.iFilialCli = objNFiscal.iFilialCli
        objNFOriginal.sSerie = SerieNFiscalOriginal.Text
        objNFOriginal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
    
        'Lê a Nota Fiscal Original
        lErro = CF("NFiscal_Le_Num_Cli", objNFOriginal, objTipoDocInfo.iTipo)
        If lErro <> SUCESSO And lErro <> 35956 Then gError 35950
            If lErro = 35956 Then gError 35952 'Se não encontrar --> Erro
        
        'Guarda em objNfiscal o Numero Interno da NF Original
        objNFiscal.lNumIntNotaOriginal = objNFOriginal.lNumIntDoc

    End If
    
    Move_Dados_NFOriginal_Memoria = SUCESSO

    Exit Function

Erro_Move_Dados_NFOriginal_Memoria:

    Move_Dados_NFOriginal_Memoria = gErr

    Select Case gErr

        Case 35946
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NFISCAL_ORIGINAL_NAO_PREENCHIDO", gErr)

        Case 35947
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NFISCAL_ORIGINAL_NAO_PREENCHIDA", gErr)

        Case 35948, 35950, 35951, 141528, 141529

        Case 35949
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPONFISCALORIGINAL_NAO_ENCONTRADO", gErr)

        Case 35952
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_ASSOCIADA_CLIENTE", gErr, Cliente.Text, SerieNFiscalOriginal.Text, NFiscalOriginal.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156257)

    End Select

    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 39005

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscal()
    If lErro <> SUCESSO Then gError 500119
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 39005, 500119

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156258)

    End Select

    Exit Sub

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
        If lErro = 51381 Then gError 86372
                
        'Se o produto é gerencial sem ser de grade ==> Erro
        If lErro = 86295 And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then
            gError 86373
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

            ColItensNF.Add 0, iIndice, objProduto.sCodigo, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), dQuantidade, 0, 0, 0, DATA_NULA, GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col), 0, 0, 0, 0, 0, colAlocacoes, 0, "", "", 0, 0, "", 0, 0, 0, objProduto.sSiglaUMEstoque, 0, 0
                        
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
    Call Preenche_GridAlocacoes(objNFiscal)
    
    For iIndice = 1 To ColItensNF.Count
    
        Call Transfere_Dados_ItensRomaneio(gobjNFiscal.ColItensNF(ColItensNF(iIndice).iItem).colItensRomaneioGrade, objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade)
    Next
    
    Exit Sub

Erro_BotaoLocalizacao_Click:

    Select Case gErr
    
        Case 43699
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)

        Case 42309, 46885, 46886, 46888, 86306, 86372, 86378, 177582
        
        Case 46889
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 86307
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, sProduto)
        
        Case 86373
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
        
        Case 86379
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objItemRomaneio.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156259)

    End Select
    
    Exit Sub

End Sub

Public Function Preenche_GridAlocacoes(objNFiscal As ClassNFiscal, Optional bGeraAlocacao As Boolean = False) As Long
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

On Error GoTo Erro_Preenche_GridAlocacoes
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
                lErro = Alocacao_Processa()
                If lErro <> SUCESSO Then gError 51602
            End If
        Else
                        
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
    
    Preenche_GridAlocacoes = SUCESSO

    Exit Function

Erro_Preenche_GridAlocacoes:

    Preenche_GridAlocacoes = gErr
    
    Select Case gErr
    
        Case 46865, 46866, 51602
        
        Case 51536
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156260)
            
    End Select

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
Dim iLinha2 As Integer
Dim objCliente As New ClassCliente
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim objFilial As New ClassFilialCliente
Dim sContaTela As String
Dim sContaMascarada As String
Dim iAlmoxPadrao As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim bEncontrouProduto As Boolean
Dim objItemMovEstoque As New ClassItemMovEstoque
Dim sProduto As String
Dim sUM As String
Dim dQuantidade As Double
Dim dQuantidadeConvertida As Double
Dim objProdutoCategoria As New ClassProdutoCategoria
Dim objCatProdItem As ClassCategoriaProdutoItem
Dim objPlanoConta As New ClassPlanoConta
Dim objProduto As New ClassProduto
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
                    If lErro <> SUCESSO Then gError 79006
                    
                    objItemMovEstoque.sProduto = sProdutoFormatado
                    
                    'Define o almoxarifado que será passado como parâmetro para MovEstoque_Le_QuantVendConsig
                    objAlmoxarifado.sNomeReduzido = GridAlocacao.TextMatrix(iLinha, iGrid_AlmoxAloc_Col)
                    
                    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25060 Then gError 79005
                    
                    objItemMovEstoque.iAlmoxarifado = objAlmoxarifado.iCodigo
                    
                    'Define o tipo de movimento, o DocOrigem e o TipoNumIntoDocOrigem que serão passados como parâmetros para MovEstoque_Le_QuantVendConsig
                    objItemMovEstoque.iTipoMov = MOV_EST_NF_VENDA_MAT_CONSIG
                    objItemMovEstoque.sDocOrigem = "NFISV" & " " & Serie.Text & " " & objContabil.lNumNotaFiscalExt
                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                                                           
                    'Verifica se MovEstoque_Le_QuantVendConsig não encontrou erro
                    lErro = CF("MovEstoque_Le_QuantVendConsig", objItemMovEstoque)
                    If lErro <> SUCESSO And lErro <> 79003 Then gError 79004
                
                    'se não conseguiu encontrar mov. estoque para os parametros em questao ==> quantidade é zerada
                    If lErro = 79003 Then objItemMovEstoque.dQuantidade = 0
                
                    'Passa para o mnemônico o valor encontrado por MovEstoque_Le_QuantVendConsig
                    objMnemonicoValor.colValor.Add objItemMovEstoque.dQuantidade
                    
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
            If Len(Filial.Text) > 0 Then
                objMnemonicoValor.colValor.Add Filial.Text
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
            If Len(Trim(Cliente.Text)) > 0 Then
            
                objCliente.sNomeReduzido = Cliente.Text
                
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then gError 88026
                
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
            
        Case CONTACONTABILEST1
            
            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(Trim(GridAlocacao.TextMatrix(iLinha, iGrid_QuantAloc_Col))) > 0 Then
                    
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
                    If lErro <> SUCESSO Then gError 183757
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183758
        
                    sContaMascarada = String(STRING_CONTA, 0)

                    lErro = Mascara_MascararConta(objProduto.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError 183794

                    objMnemonicoValor.colValor.Add sContaMascarada
        
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case CONTACONTABILCATPROD
                
            For iLinha = 1 To objGridItens.iLinhasExistentes
                
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183759
                
                    objProdutoCategoria.sCategoria = objMnemonicoValor.vParam(1)
                    
                    objProdutoCategoria.sProduto = sProdutoFormatado
                    
                    Set objCatProdItem = New ClassCategoriaProdutoItem
                    
                    lErro = CF("ProdutoCategoria_Le_Item", objProdutoCategoria, objCatProdItem)
                    If lErro <> SUCESSO And lErro <> 183548 Then gError 183760
                    
                    If lErro = SUCESSO Then
                        
                        Set objPlanoConta = New ClassPlanoConta
                    
                        'le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta
                        lErro = CF("PlanoConta_Le_ContaSimples", objCatProdItem.dvalor1, objPlanoConta)
                        If lErro <> SUCESSO And lErro <> 5451 Then gError 183761
                    
                        If lErro = SUCESSO Then
                    
                            sContaMascarada = String(STRING_CONTA, 0)
        
                            lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
                            If lErro <> SUCESSO Then gError 183795
        
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
                If lErro <> SUCESSO Then gError 186132
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 186133
                
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
                
        Case 61511, 61512, 61513, 61514, 61516, 64170, 64172, 64173, 79004, 79006, 88026, 183757 To 183761, 183794, 183795, 186132, 186133
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156261)

    End Select

    Exit Function

End Function

'Private Function TabelaPreco_Atualiza_PrecoUnitario(objProduto As ClassProduto) As Long
''Se a Tabela de Preços estiver selecionada tenta calcular o preço unitário para o produto passado
'
'Dim lErro As Long
'Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
'Dim dPrecoUnitario As Double
'
'On Error GoTo Erro_TabelaPreco_Atualiza_PrecoUnitario
'
'    'Verifica se a TabelaPreco está selecionada
'    If Len(Trim(TabelaPreco.Text)) > 0 Then
'
'        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
'        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
'        objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
'
'        'Lê o preço do produto para a filialempresa
'        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
'        If lErro <> SUCESSO And lErro <> 28014 Then gError 35897
'        'Se não encontrar
'        If lErro = 28014 Then
'            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
'
'            'Lê o preço para a empresa toda
'            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
'            If lErro <> SUCESSO And lErro <> 28014 Then gError 35898
'
'        End If
'
'        'Se encontrou o valor na Tabela de preço
'        If lErro <> 28014 Then
'            'Calcula o preço unitário em relação a UM da tela
'            lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
'            If lErro <> SUCESSO Then gError 35899
'
'            'Coloca o preço unitário na tela
'            If dPrecoUnitario > 0 Then
'                GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
'            Else
'                GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = ""
'            End If
'        End If
'    End If
'
'    TabelaPreco_Atualiza_PrecoUnitario = SUCESSO
'
'    Exit Function
'
'Erro_TabelaPreco_Atualiza_PrecoUnitario:
'
'    TabelaPreco_Atualiza_PrecoUnitario = gErr
'
'    Select Case gErr
'
'        Case 35897, 35898, 35899
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156262)
'
'    End Select
'
'    Exit Function
'
'End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'tulio261202
    'feito em virtude do redimensionamento da msg da nota p/ inpal
    Call CF("Redimensiona_Tam_MsgNF", Mensagem)

    Parent.HelpContextID = IDH_NF_SAIDA_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Nota Fiscal de Saída"
    Call Form_Load
    
End Function

Public Function Name() As String
    Name = "NFiscal"
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
    
Dim lErro As Long

On Error GoTo Erro_DescontoGlobal_Testa
    
    Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)
    
    'Verifica se o Valor do Desconto é maior que a soma dos demais valores
    If dValorDesconto > dValorProdutos Then gError 35838
    
    DescontoGlobal_Testa = SUCESSO
     
    Exit Function
    
Erro_DescontoGlobal_Testa:

    DescontoGlobal_Testa = gErr
     
    Select Case gErr
                  
        Case 35838
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156263)
     
    End Select
     
    Exit Function

End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Serie Then
            Call SerieLabel_Click
        ElseIf Me.ActiveControl Is NFiscal Then
            Call NFiscalLabel_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        ElseIf Me.ActiveControl Is TranspRedespacho Then
            Call TranspRedLabel_Click
        ElseIf Me.ActiveControl Is SerieNFiscalOriginal Then
            Call SerieNFOriginalLabel_Click
        ElseIf Me.ActiveControl Is NFiscalOriginal Then
            Call NFiscalOriginalLabel_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call BotaoVendedores_Click
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
    If lErro <> SUCESSO And lErro <> 17660 Then gError 51610

    If lErro = 17660 Then gError 51611

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_Filial_Formata:

    Select Case gErr

        Case 51610
        
        Case 51611
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156264)

    End Select

    Exit Sub

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
    
    Mensagem.Text = objMensagem.sDescricao
    
    Me.Show

End Sub

Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

Public Property Get ImprimeRomaneio() As Object
     Set ImprimeRomaneio = objUserControl.Controls("ImprimeRomaneio")
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

Public Sub ImprimeRomaneio_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Calcula_ValorProdutos(dValorProdutos As Double)

Dim dValorTotal As Double
Dim dValor As Double
Dim iIndice As Integer

    dValor = 0
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        dValorTotal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        dValor = dValor + dValorTotal
    Next
    
    dValorProdutos = dValor
    
End Sub

Private Sub DescontoGlobal_Recalcula()

Dim dValorDesconto As Double
Dim dValorProdutos As Double
Dim lErro As Long

On Error GoTo Erro_DescontoGlobal_Recalcula

    If gbCarregandoTela Then Exit Sub
    
    PercDescontoItens.Text = Format(gdDesconto * 100, "FIXED")
    Call PercDescontoItens_Validate(bSGECancelDummy)

'
'    If Len(Trim(ValorProdutos.Caption)) <> 0 And IsNumeric(ValorProdutos.Caption) Then
'
'        dValorProdutos = CDbl(ValorProdutos.Caption)
'
'        'Se o cliente possui desconto e o campo desconto não foi alterado pelo usuário
'        If gdDesconto > 0 And iValorDescontoAlterado = 0 Then
'
'            Call Calcula_ValorProdutos(dValorProdutos)
'
'            'Calcula o valor do desconto para o cliente e coloca na tela
'            dValorDesconto = gdDesconto * dValorProdutos
'            ValorDesconto.Text = Format(dValorDesconto, "Standard")
'            iValorDescontoAlterado = 0
'
'            ValorProdutos.Caption = Format(dValorProdutos - dValorDesconto, "STANDARD")
'
'            lErro = CF2(Me, "PV_ValorDesconto_Validate")
'            If lErro <> SUCESSO Then gError 132023
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

        Case 132023

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156265)

    End Select

    Exit Sub

End Sub

Private Sub AlocacoesNF_Agrupa(colAlocacoes As ColAlocacoesItemNF)

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

Private Sub NFiscal_Imprime_Romaneio(objNFiscal As ClassNFiscal)
    Call CF("NFiscal_Imprime_Romaneio", objNFiscal)
End Sub

Public Sub BotaoConsultaTitRec_Click()
'Abre uma tela para consulta do DocCPR vinculado à nota fiscal

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoConsultaTitRec_Click

    'Verifica se todos os campos necessários para se efetuar a consulta foram preenchidos
    lErro = Critica_CamposNecessarios_ConsultaTitulo()
    If lErro <> SUCESSO Then gError 79816
    
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
    If lErro <> SUCESSO And lErro <> 12348 Then gError 79817
    
    'Se não encontrou o fornecedor = > erro
    If lErro = 12348 Then gError 79818
    
    'Guarda no objNFiscal o código do fornecedor
    objNFiscal.lCliente = objCliente.lCodigo
    
    'Lê o NumIntDocCPR da NFiscal e exibe o documento gerado no CPR por essa nota
    lErro = CF("NFiscal_Consulta_DocCPR", objNFiscal)
    If lErro <> SUCESSO And lErro <> 79717 Then gError 79819
    
    'Se não encontrou a nota => erro
    If lErro = 79717 Then gError 79820
    
    Exit Sub
    
Erro_BotaoConsultaTitRec_Click:

    Select Case gErr
    
        Case 79816, 79819, 79817
        
        Case 79820
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA2", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iTipoNFiscal, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEmissao)
        
        Case 79818
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156267)
    
    End Select
    
    Exit Sub

End Sub

Private Function Critica_CamposNecessarios_ConsultaTitulo() As Long
'Verifica se os campos necessários para encontrar consultar um título gerado por uma NFFatEntrada foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposNecessarios_ConsultaTitulo

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 79821
    
    'Verifica se a filial do Fornecedor foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 79822
    
    'Verifica se a Série foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 79823
    
    'Verifica se o Número da Nota foi preenchido
    If Len(Trim(NFiscal.Caption)) = 0 Then gError 79824
    
    'Verifica se a data de emissão da nota foi preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 79825
    
    'Verifica se o tipo da Nota fiscal foi selecionado
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 79826
    
    Critica_CamposNecessarios_ConsultaTitulo = SUCESSO
    
    Exit Function

Erro_Critica_CamposNecessarios_ConsultaTitulo:

    Critica_CamposNecessarios_ConsultaTitulo = gErr
    
    Select Case gErr
    
        Case 79821
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 79822
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 79823
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 79824
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
        
        Case 79825
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 79826
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156268)
    
    End Select
        
    Exit Function
        
End Function

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

Public Sub OrcamentoLabel_Click()

Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim colSelecao As New Collection

    Call Chama_Tela("OrcamentoPedidoLista", colSelecao, objOrcamentoVenda, objEventoOrcamento)

End Sub

Public Property Get OrcamentoLabel() As Object
     Set OrcamentoLabel = objUserControl.Controls("OrcamentoLabel")
End Property

Private Sub objEventoOrcamento_evSelecao(obj1 As Object)

Dim objOrcamentoVenda As ClassOrcamentoVenda

    Set objOrcamentoVenda = obj1

    Orcamento.Text = objOrcamentoVenda.lCodigo
    FilialOrcamento.Text = objOrcamentoVenda.iFilialEmpresa
    Call FilialOrcamento_Validate(bSGECancelDummy)

    BotaoOrcamento_Click

    Me.Show
    
End Sub

Public Sub Orcamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Orcamento_GotFocus()

    Call MaskEdBox_TrataGotFocus(Orcamento, iAlterado)

End Sub

Public Property Get Orcamento() As Object
     Set Orcamento = objUserControl.Controls("Orcamento")
End Property

Public Sub FilialOrcamento_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialOrcamento_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Property Get FilialOrcamento() As Object
     Set FilialOrcamento = objUserControl.Controls("FilialOrcamento")
End Property

Private Sub Carrega_FilialOrcamento()

Dim lErro As Long
Dim objFiliais As AdmFiliais

    For Each objFiliais In gcolFiliais

        If objFiliais.iCodFilial <> EMPRESA_TODA Then

            'coloca na combo
            FilialOrcamento.AddItem CStr(objFiliais.iCodFilial) & SEPARADOR & objFiliais.sNome
            FilialOrcamento.ItemData(FilialOrcamento.NewIndex) = objFiliais.iCodFilial

        End If

    Next

    'Seleciona a Filial na qual o usuário entrou no Sistema
    Call Seleciona_FilialEmpresa

    Exit Sub

Erro_Carrega_FilialOrcamento:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156270)

    End Select

    Exit Sub

End Sub

Public Sub FilialOrcamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_FilialOrcamento_Validate

    'Se não estiver preenchida, sai.
    If Len(Trim(FilialOrcamento.Text)) = 0 Then Exit Sub
    'Se não foi selecoinada, sai.
    If FilialOrcamento.ListIndex <> -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialOrcamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 94423
    If lErro <> SUCESSO Then gError 94424 'Não encontrou.

    Exit Sub

Erro_FilialOrcamento_Validate:

    Cancel = True

    Select Case gErr

        Case 94423

        Case 94424
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialOrcamento.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156271)

    End Select

    Exit Sub

End Sub

Public Sub BotaoVerOrcamento_Click()

Dim lErro As Long
Dim objOcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_BotaoVerPedido_Click

    'Verifica se o Orçamento e a filial do pedido foram informados
    If Len(Trim(Orcamento.Text)) = 0 Then gError 94427
    If Len(Trim(FilialOrcamento.Text)) = 0 Then gError 94428
    
    objOcamentoVenda.lCodigo = StrParaLong(Orcamento.Text)
    objOcamentoVenda.iFilialEmpresa = Codigo_Extrai(FilialOrcamento.Text)
        
    'Busca o Pedido na tabela de pedidos de venda
    lErro = CF("OrcamentoVenda_Le", objOcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101232 Then gError 94429
    
    'Não encontrou => erro
    If lErro = 101232 Then gError 94430
    
    'Chama  a tela de Orcamento de Venda
    Call Chama_Tela("OrcamentoVenda", objOcamentoVenda)
    
    Exit Sub

Erro_BotaoVerPedido_Click:

    Select Case gErr
    
        Case 94427
            Call Rotina_Erro(vbOKOnly, "ERRO_CODORCAMENTO_NAO_INFORMADO", gErr)
            
        Case 94428
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALORCAMENTO_NAO_INFORMADA", gErr)
        
        Case 94429
                
        Case 94430
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, objOcamentoVenda.lCodigo, objOcamentoVenda.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156272)

    End Select

    Exit Sub

End Sub

Public Sub BotaoOrcamento_Click()
             
Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_BotaoOrcamento_Click
    
    'Verifica se o Orçamento e a filial do pedido foram informados
    If Len(Trim(Orcamento.Text)) = 0 Then gError 94427
    If Len(Trim(FilialOrcamento.Text)) = 0 Then gError 94428
    
    objOrcamentoVenda.lCodigo = StrParaLong(Orcamento.Text)
    objOrcamentoVenda.iFilialEmpresa = Codigo_Extrai(FilialOrcamento.Text)
        
    'Busca o Pedido na tabela de pedidos de venda
    lErro = CF("OrcamentoVenda_Le_Completo", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101232 Then gError 94429
    
    If lErro = 101232 Then gError 94430 'Não encontrou, erro.
    
    'Se o orçamento estiver vinculado a um Pedido de Venda -> Erro.
    If objOrcamentoVenda.lNumIntPedVenda <> 0 Then
        
        gError 94425
    
    'Se o orçamento estiver vinculado a uma Nota Fiscal -> Erro.
    ElseIf objOrcamentoVenda.lNumIntNFiscal <> 0 Then
        
        gError 94426
    
    End If
        
    'Verifica se a Filial de Faturamento do Orcamento coincide com a filial de trabalho
    If objOrcamentoVenda.iFilialEmpresa <> giFilialEmpresa Then gError 94431
    
    'Traz os dados do pedido para a tela
    lErro = Traz_Dados_Orcamento_Tela(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 94432

    Call ValorTotal_Calcula

    Exit Sub

Erro_BotaoOrcamento_Click:

    Select Case gErr

        Case 94427
            Call Rotina_Erro(vbOKOnly, "ERRO_CODORCAMENTO_NAO_INFORMADO", gErr)

        Case 94428
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALORCAMENTO_NAO_INFORMADA", gErr)

        Case 94429, 94432

        Case 94425
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_VINCULADO_ORCVENDA1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
        
        Case 94426
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_VINCULADO_NFISCAL1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
        
        Case 94430
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case 94431
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFATURAMENTO_DIFERENTE1", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, giFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156273)

    End Select

    Exit Sub

End Sub

Public Property Get BotaoOrcamento() As Object
     Set BotaoOrcamento = objUserControl.Controls("BotaoOrcamento")
End Property

Public Function Traz_Dados_Orcamento_Tela(objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objCliente As New ClassCliente

On Error GoTo Erro_Traz_Dados_Orcamento_Tela


    'Transfere os dados de objOrcamentoVenda para objNFiscal
    Call Transfere_Dados_Orcamento_NFiscal(objOrcamentoVenda, objNFiscal)

    objNFiscal.dtDataEmissao = gdtDataAtual
          
    'Carrega os dados de OrcamentoVenda na tela
    lErro = Traz_NFiscal_Tela(objNFiscal, objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 94434
    
    'Processa as Alocacoes
    lErro = Alocacao_Processa_Todos()
    If lErro <> SUCESSO Then gError 101321
    
    'chamei essa func. para preencher o grid de comissoes tulio241002
    Call Cliente_Validate(bSGECancelDummy)
        
    lErro = CF2(Me, "PV_Traz_Dados_Orcamento_Tela", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 141415
        
    Traz_Dados_Orcamento_Tela = SUCESSO

    Exit Function

Erro_Traz_Dados_Orcamento_Tela:

    Traz_Dados_Orcamento_Tela = gErr

    Select Case gErr

        Case 94432, 94434, 101321, 141415

        Case 94433
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156274)

    End Select

    Exit Function

End Function

Private Sub Transfere_Dados_Orcamento_NFiscal(objOrcamentoVenda As ClassOrcamentoVenda, objNFiscal As ClassNFiscal)
'Transfere os dados do objOrcamentoVenda para objNFiscal

Dim objItemOV As ClassItemOV
Dim objItemNF As ClassItemNF
Dim dQuantidade As Double
Dim dPercDesc As Double
Dim objComissaoNF As New ClassComissaoNF
Dim objParcelaOV As ClassParcelaOV
Dim lErro As Long
Dim iIndice As Integer
Dim objItemTributacaoNF As ClassTributacaoDocItem
Dim objTribTab As New ClassTribTab
Dim objProduto As ClassProduto
Dim colItensRomaneioGrade As New Collection
Dim objKitVenda As New ClassKitVenda
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objProjetoInfo As Object, dtDataBase As Date, dtDataAux As Date
Dim dCotacao As Double
Dim objCotacao As New ClassCotacaoMoeda
Dim objCotacaoAnterior As New ClassCotacaoMoeda
Dim vbMsgResult As VbMsgBoxResult

On Error GoTo Erro_Transfere_Dados_Orcamento_NFiscal

    Call gobjTelaProjetoInfo.Obtem_ProjetoInfo(objProjetoInfo, objOrcamentoVenda.lNumIntDoc, PRJ_CR_TIPO_OV)
    Set objNFiscal.objProjetoInfo = objProjetoInfo

   Set objNFiscal.objInfoAdic = objOrcamentoVenda.objInfoAdic

    If objOrcamentoVenda.iMoeda <> MOEDA_REAL Then
    
        objCotacao.dtData = gdtDataAtual
        objCotacao.iMoeda = objOrcamentoVenda.iMoeda
        objCotacaoAnterior.iMoeda = objOrcamentoVenda.iMoeda
    
        'Chama função de leitura
        lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objCotacao.dValor <> 0 Then
            dCotacao = StrParaDbl(Format(objCotacao.dValor, "#.0000"))
        Else
            vbMsgResult = Rotina_Aviso(vbYesNo, "AVISO_MOEDA_SEM_COTACAO_DATA", objCotacao.iMoeda, Format(gdtDataAtual, "dd/mm/yyyy"), Format(objCotacaoAnterior.dValor, "#.0000"), Format(objCotacaoAnterior.dtData, "dd/mm/yyyy"))
            If vbMsgResult = vbNo Then gError ERRO_SEM_MENSAGEM
            dCotacao = StrParaDbl(Format(objCotacaoAnterior.dValor, "#.0000"))
        End If
        
        If dCotacao = 0 Then gError 211630
        
    Else
        dCotacao = 1
    End If
    
    With objNFiscal
        
        'Passa os dados gerais da tela de Orcamento para o Pedido
        objOrcamentoVenda.dtDataEmissao = gdtDataAtual
            
        .dtDataEmissao = objOrcamentoVenda.dtDataEmissao
        .dtDataReferencia = objOrcamentoVenda.dtDataReferencia
        .dValorDesconto = Arredonda_Moeda(objOrcamentoVenda.dValorDesconto * dCotacao)
        .dValorFrete = Arredonda_Moeda(objOrcamentoVenda.dValorFrete * dCotacao)
        .dValorOutrasDespesas = Arredonda_Moeda(objOrcamentoVenda.dValorOutrasDespesas * dCotacao)
        .dValorProdutos = objOrcamentoVenda.dValorProdutos '* dCotacao
        .dValorSeguro = Arredonda_Moeda(objOrcamentoVenda.dValorSeguro * dCotacao)
        .dValorTotal = objOrcamentoVenda.dValorTotal '* dCotacao
        .iCondicaoPagto = objOrcamentoVenda.iCondicaoPagto
        .iFilialEntrega = objOrcamentoVenda.iFilialEntrega
        .iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa
        .lCliente = objOrcamentoVenda.lCliente
        .iFilialCli = objOrcamentoVenda.iFilial
        .lNumOrcamentoVenda = objOrcamentoVenda.lCodigo
        .iFilialOrcamento = objOrcamentoVenda.iFilial
        .lNumIntDoc = objOrcamentoVenda.lNumIntNFiscal
        .sNaturezaOp = objOrcamentoVenda.sNaturezaOp
        .iTabelaPreco = objOrcamentoVenda.iTabelaPreco
        .dValorDescontoItens = Arredonda_Moeda(objOrcamentoVenda.dValorDescontoItens * dCotacao)
        .dValorDescontoTit = Arredonda_Moeda(objOrcamentoVenda.dValorDescontoTit * dCotacao)
        .dValorItens = objOrcamentoVenda.dValorItens '* dCotacao
        
        .iCodTransportadora = objOrcamentoVenda.iCodTransportadora
        .iCodTranspRedesp = objOrcamentoVenda.iCodTranspRedesp
        .sMensagemNota = objOrcamentoVenda.sMensagemNota
        .sNumPedidoTerc = objOrcamentoVenda.sPedidoCliente
        '.sPedidoRepresentante = objOrcamentoVenda.sPedidoRepresentante
        .iCanal = objOrcamentoVenda.iCanalVenda
        .dPesoBruto = objOrcamentoVenda.dPesoBruto
        .dPesoLiq = objOrcamentoVenda.dPesoLiq
        .sPlaca = objOrcamentoVenda.sPlaca
        .sPlacaUF = objOrcamentoVenda.sPlacaUF
        .lVolumeQuant = objOrcamentoVenda.lVolumeQuant
        .lVolumeEspecie = objOrcamentoVenda.lVolumeEspecie
        .lVolumeMarca = objOrcamentoVenda.lVolumeMarca
        .sVolumeNumero = objOrcamentoVenda.sVolumeNumero
        .dVolumeTotal = objOrcamentoVenda.dVolumeTotal
        .iDetPagFrete = objOrcamentoVenda.iDetPagFrete
        .iFreteRespons = objOrcamentoVenda.iFreteRespons
       
    End With
   
    Call CF2(Me, "NFCF_DataSaida", objNFiscal)

    'Para cada item do Orcamento
    For Each objItemOV In objOrcamentoVenda.colItens
        
        iIndice = iIndice + 1
        
        If objItemOV.dQuantidade > 0 Then
            
            If objItemOV.dPrecoTotal > 0 Then
                
                'Calcula o percentual de desconto
                dPercDesc = objItemOV.dValorDesconto / objItemOV.dPrecoTotal
            End If

            Call objTribTab.ItemNFiscal_ObterTrib_ItemOrcamento(objItemOV, objItemTributacaoNF)
            
            With objItemOV
                                
                'Adiciona o item na coleção de itens da NF
                objNFiscal.ColItensNF.Add 0, iIndice, .sProduto, .sUnidadeMed, .dQuantidade, IIf(dCotacao = 1, .dPrecoUnitario, Arredonda_Moeda(.dPrecoUnitario * dCotacao)), dPercDesc, IIf(dCotacao = 1, .dValorDesconto, Arredonda_Moeda(.dValorDesconto * dCotacao)), .dtDataEntrega, .sDescricao, 0, .lNumIntDoc, 0, 0, .lNumIntDoc, Nothing, 0, "", "", .iStatus, 0, "", 0, 0, 0, "", .iClasseUM, 0
                
                Set objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).objInfoAdicDocItem = objItemOV.objInfoAdicDocItem
                
                If objItemOV.iPrazoEntrega <> 0 Then
                    dtDataBase = StrParaDate(DataEmissao.Text)
                    If dtDataBase = DATA_NULA Then dtDataBase = gdtDataAtual
                    
                    lErro = CF("Calcula_Data", dtDataBase, objItemOV.iPrazoEntrega, 0, CONDPAGTO_TIPOINTERVALO_DIAS_UTEIS, dtDataAux)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).dtDataEntrega = dtDataAux
                End If
                
                Set objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).objTributacaoItemNF = objItemTributacaoNF
                                
                If objItemOV.colItensRomaneioGrade.Count > 0 Then
                    objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).iPossuiGrade = MARCADO
                    Call Move_ItensGrade_Tela(objNFiscal.ColItensNF(objNFiscal.ColItensNF.Count).colItensRomaneioGrade, objItemOV.colItensRomaneioGrade)
                End If
                
            End With
            
            Set objProduto = New ClassProduto
            
            objProduto.sCodigo = objItemOV.sProduto
                        
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 180034
                                    
            If objProduto.iKitVendaComp = MARCADO Then
            
                lErro = CF("ProdutoKitVenda_Le_Romaneio", objProduto, colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 180035
                
                objKitVenda.sProduto = objItemOV.sProduto
                objKitVenda.sUM = objItemOV.sUnidadeMed
                objKitVenda.dQuantidade = objItemOV.dQuantidade
                
                lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 180036
                
                Set objItemOV.colItensRomaneioGrade = colItensRomaneioGrade
            
                Call Move_ItensGrade_Tela(objNFiscal.ColItensNF(objNFiscal.ColItensNF.Count).colItensRomaneioGrade, objItemOV.colItensRomaneioGrade)
                        
            End If

        End If
    
    Next

    'Carrega os dados globais da tributação do OV para a NF
    Call objTribTab.NFiscal_ObterTrib_Orcamento(objOrcamentoVenda, objNFiscal)
    
'    objNFiscal.objTributacaoNF.objTributacaoDesconto.iTipo = ITEM_TIPO_DESCONTO
'    objNFiscal.objTributacaoNF.objTributacaoFrete.iTipo = ITEM_TIPO_FRETE
'    objNFiscal.objTributacaoNF.objTributacaoOutras.iTipo = ITEM_TIPO_OUTRAS_DESP
'    objNFiscal.objTributacaoNF.objTributacaoSeguro.iTipo = ITEM_TIPO_SEGURO
    
    Exit Sub

Erro_Transfere_Dados_Orcamento_NFiscal:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 51562, 180034, 180035, 180036, 180037, 181837

        Case 211630
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_SEM_COTACAO", gErr, objCotacao.iMoeda)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156275)

    End Select

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
    If lErro <> SUCESSO Then gError 56915

    lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 56916

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
    
        Case 56915, 56916
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156276)

    End Select

    Exit Sub

End Sub

'Daniel 23/05/2002
Private Sub Seleciona_FilialEmpresa()

Dim iIndice As Integer
Dim iFilialFaturamento As Integer

    iFilialFaturamento = gobjFAT.iFilialFaturamento
    
    If iFilialFaturamento <> EMPRESA_TODA Then
        
        'seleciona a filial de faturamento na combo
        For iIndice = 0 To FilialOrcamento.ListCount - 1

            If FilialOrcamento.ItemData(iIndice) = iFilialFaturamento Then

                FilialOrcamento.ListIndex = iIndice
                Exit For

            End If
        Next

    Else
    
        FilialOrcamento.ListIndex = 0
        
    End If

    Exit Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156277)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156278)
            
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156279)
            
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156280)

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

Function Grid_Possui_Grade() As Boolean

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim iIndice As Integer

    For iIndice = 1 To gobjNFiscal.ColItensNF.Count
        If gobjNFiscal.ColItensNF(iIndice).iPossuiGrade = MARCADO Then
            Grid_Possui_Grade = True
            Exit Function
        End If
    Next
    
    Grid_Possui_Grade = False
        
    Exit Function
    
End Function

Public Property Get BotaoGrade() As Object
     Set BotaoGrade = objUserControl.Controls("BotaoGrade")
End Property

Private Function Alocacao_Processa_ItemGrade(objItemRomaneio As ClassItemRomaneioGrade, ByVal objItemNF As ClassItemNF) As Long

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
    If lErro <> SUCESSO And lErro <> 28030 Then gError 42201
    
    Set objItemRomaneio.colLocalizacao = New Collection

    objProduto.sCodigo = objItemRomaneio.sProduto

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 42201

    'Se não encontrou o Produto --> erro
    If lErro = 28030 Then gError 42202

    dQuantidade = objItemRomaneio.dQuantidade

   If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 And gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA Then

        sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

        lErro = AlocaAlmoxarifadoPradrao_Grade(objProduto, objItemRomaneio, sUM, dQuantidade, objProdutoINF)
        If lErro <> SUCESSO Then gError 42203
    
    End If

    Alocacao_Processa_ItemGrade = SUCESSO

    Exit Function

Erro_Alocacao_Processa_ItemGrade:

    Alocacao_Processa_ItemGrade = gErr

    Select Case gErr

        Case 42200, 42201, 42203

        Case 42202
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156281)

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
        If lErro <> SUCESSO Then gError 42204
    Else
        dFator = 1
    End If

    'Converte a quantidade para a UM de estoque
    dQuantAlocar = dQuantidade * dFator

    'Busca o Almoxarifado padrão
    lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 42205

    'Se não encontrou
    If iAlmoxarifado = 0 Then gError 51416

    objAlmoxarifado.iCodigo = iAlmoxarifado
    'Lê o Aloxarifado
    lErro = CF("Almoxarifado_Le", objAlmoxarifado)
    If lErro <> 25056 And lErro <> SUCESSO Then gError 42206
    If lErro = 25056 Then gError 42207

    objEstoqueProduto.sProduto = objProduto.sCodigo
    objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
    'Lê os estoques desse produto nesse almoxarifado
    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 21306 Then gError 42208
    'Se não achou o Estoque do Produto --> erro
    If lErro = 21306 Then gError 42209

    iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Seleciona a origem da quantidade disponível
    Select Case iCodigo
    
        'Se o tipo da nota for cobrança de mat. consignado
        Case DOCINFO_NFFISPC
            
            'A quantidade disponível deve ser igual a quantidade no escaninho de mat. em Consignação (Consig)
            dQuantDisponivel = objEstoqueProduto.dQuantConsig
        
        'Se o tipo da nota for mat. beneficiado de 3º´s
        Case DOCINFO_NFISFBF
        
            'A quantidade disponível deve ser igual a quantidade no escaninho Mat.de 3º´s em Beneficiamento (Benef3)
            dQuantDisponivel = objEstoqueProduto.dQuantBenef3
            
        'Se for outro tipo de nota
        Case Else
            
            'A quantidade disponível deve ser igual a quantidade do escaninho mat. nosso disponível (DispNossa)
            dQuantDisponivel = objEstoqueProduto.dQuantDisponivel
    
    End Select
    
    Set objItemRomaneio.colLocalizacao = New Collection

    dQuantAlocar = Arredonda_Estoque(dQuantAlocar)
    
    'Verifica se a Quantidade disponível é maior que a quantidade a alocar
    If (dQuantAlocar - dQuantDisponivel) < QTDE_ESTOQUE_DELTA Then
        
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
                    If lErro <> SUCESSO Then gError 42210
                    'Adiciona na coleção de produtos
                    colOutrosProdutos.Add sProduto
                End If
            Next
    
            'Chama a tela de Localização de Produto
            Call Chama_Tela_Modal("LocalizacaoProduto1", objItemNF, colOutrosProdutos, dQuantAlocar, iCodigo)
            If giRetornoTela = vbCancel Then gError 42211
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

        Case 35820
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)
        
        Case 42204, 42208, 42205, 42206, 42210

        Case 42207
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, iAlmoxarifado)

        Case 42209
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case 42212

        Case 42211
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_LOCALIZACAO", gErr, objProduto.sCodigo)

        Case 51416
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO1", gErr, objProduto.sCodigo, giFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156282)

    End Select

    Exit Function

End Function

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134014

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134014

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156283)

    End Select
    
    Exit Sub

End Sub

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long, sID As String
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Anotacao_Extrai

    'Lê os dados da Tela NFiscal
    lErro = Move_NF_Memoria(objNFiscal)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156284)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156285)
     
    End Select
     
    Exit Sub

End Sub

Public Function NF_AlocaAlmoxarifadoPadrao(ByVal objCT As Object, ByVal iFilialEmpresa As Integer, ByVal sProduto As String, iAlmoxarifadoRetorno As Integer) As Long

Dim lErro  As Long

On Error GoTo Erro_NF_AlocaAlmoxarifadoPadrao

    'Busca o Almoxarifado padrão
    lErro = CF("AlmoxarifadoPadrao_Le", iFilialEmpresa, sProduto, iAlmoxarifadoRetorno)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 51619
    
    'Se encontrou
    If iAlmoxarifadoRetorno = 0 Then gError 35822
    
    Exit Function

Erro_NF_AlocaAlmoxarifadoPadrao:

    NF_AlocaAlmoxarifadoPadrao = gErr

    Select Case gErr
    
        Case 35822, 51619

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156286)

    End Select

    Exit Function
    
End Function

Public Function NF_Form_Load(ByVal objCT As Object) As Long
    NF_Form_Load = SUCESSO
End Function

Public Function NF_Form_UnLoad(ByVal objCT As Object) As Long
    NF_Form_UnLoad = SUCESSO
End Function

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

Public Function NF_Inicializa_Grid_Itens(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

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

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NF_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Function NF_PrecoTotal_Calcula(ByVal objCT As Object, ByVal iLinha As Integer, ByVal dQuantidade As Double, ByVal dPrecoUnitario As Double, ByVal dPercentDesc As Double) As Long
    NF_PrecoTotal_Calcula = SUCESSO
End Function

Public Function NF_Produto_Saida_Celula(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long
    NF_Produto_Saida_Celula = SUCESSO
End Function

Public Function NF_Preenche_GridItens(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal, ByVal objItemNF As ClassItemNF, ByVal objProduto As ClassProduto, ByVal iIndice As Integer) As Long
    NF_Preenche_GridItens = SUCESSO
End Function

Public Function NF_CargaPosFormLoad(ByVal objCT As Object) As Long
    NF_CargaPosFormLoad = SUCESSO
End Function

Public Function NF_Move_NF_Memoria(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long
    NF_Move_NF_Memoria = SUCESSO
End Function

Public Function NF_Traz_NF_Tela(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long
    NF_Traz_NF_Tela = SUCESSO
End Function

Public Function NF_Limpa_GridItens(ByVal objCT As Object) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_NF_Limpa_GridItens

    For iIndice = objCT.objGridItens.iLinhasExistentes To 1 Step -1

        lErro = objCT.gobjTribTab.Exclusao_Item_Grid(iIndice)
        If lErro <> SUCESSO Then gError 99078

        GridItens.TextMatrix(iIndice, 0) = iIndice
    
    Next

    Call Grid_Limpa(objGridItens)

    NF_Limpa_GridItens = SUCESSO
    
    Exit Function
    
Erro_NF_Limpa_GridItens:

    NF_Limpa_GridItens = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184460)

    End Select
    
    Exit Function

End Function

Public Function NF_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iLocalChamada As Integer) As Long
    NF_Rotina_Grid_Enable = SUCESSO
End Function

Function NF_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    NF_Saida_Celula = SUCESSO
End Function

Public Function NF_Limpa_Tela_Cust(ByVal objCT As Object) As Long
    NF_Limpa_Tela_Cust = SUCESSO
End Function

Public Sub NFiscalAux_Change()
    If NFiscalAux.Visible Then NFiscal.Caption = NFiscalAux.Text
End Sub

Public Sub NFiscalAux_GotFocus()
    Call MaskEdBox_TrataGotFocus(NFiscalAux, iAlterado)
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

Private Function Carrega_TipoFrete() As Long
'Carrega versaoNFE

Dim lErro As Long
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer
Dim colTipoFrete As Collection

On Error GoTo Erro_Carrega_TipoFrete
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207418

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
        Case 207418

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207419)

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
Dim dPercDescontoTab As Double, dPrecoComDescontoTab As Double

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
        
        lErro = CF2(Me, "NFCF_PrecoUnitario_Calcula", giFilialEmpresa, sProduto, GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), Cliente.Text, Codigo_Extrai(Filial.Text), iCodTabela, 0, MOEDA_REAL, Nothing, dPrecoUnitario, False, dPercDescontoTab, dPrecoComDescontoTab)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Coloca no Grid
        If dPrecoUnitario > 0 Then
            GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        Else
            GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col) = ""
        End If
        
        If dPercDescontoTab >= gdDesconto Then
            GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col) = Format(dPercDescontoTab, "Percent")
        End If
        
        'Calcula o Preco Total do item
        Call PrecoTotal_Calcula(iLinha, False)

        ''*** incluidos p/tratamento de tributacao *******************************
        lErro = gobjTribTab.Alteracao_Item_Grid(iLinha)
        If lErro <> SUCESSO Then gError 56860
        '*** fim tributacao
        
    End If

    Trata_TabelaPreco_Item = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco_Item:

    Trata_TabelaPreco_Item = gErr

    Select Case gErr

        Case 46660, 46661, 46662, 46663, 56860, ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157457)

    End Select

    Exit Function

End Function

Private Sub Trata_NF_Servico()
    
    If Codigo_Extrai(TipoNFiscal.Text) = 70 Then
        ClienteLabel.Visible = False
        FrameCliInt.Visible = True
        LabelCliTomador.Visible = True
    Else
        ClienteLabel.Visible = True
        FrameCliInt.Visible = False
        LabelCliTomador.Visible = False
        CliIntermediario.Text = ""
        FilialCliIntermediario.Clear
    End If

End Sub

Public Sub CliIntermediario_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CliIntermediario_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_CliIntermediario_Validate
    
    If Len(Trim(CliIntermediario.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Cliente_Le(CliIntermediario, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Lê coleção de códigos, nomes de Filiais do Fornecedor
        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", FilialCliIntermediario, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", FilialCliIntermediario, iCodFilial)

    ElseIf Len(Trim(CliIntermediario.Text)) = 0 Then

        FilialCliIntermediario.Clear

    End If
    
    Exit Sub

Erro_CliIntermediario_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211303)

    End Select

    Exit Sub

End Sub

Public Sub LabelCliTomador_Click()
    Call ClienteLabel_Click
End Sub

Public Sub LabelCliInt_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(CliIntermediario.Text)) > 0 Then objCliente.sNomeReduzido = CliIntermediario.Text

    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliIntermediario)

End Sub

Private Sub objEventoCliIntermediario_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    CliIntermediario.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call CliIntermediario_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub FilialCliIntermediario_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialCliIntermediario_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialCliIntermediario_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialCliIntermediario.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialCliIntermediario.Text = FilialCliIntermediario.List(FilialCliIntermediario.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialCliIntermediario, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(CliIntermediario.Text)) = 0 Then gError 92062

        sCliente = CliIntermediario.Text

        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError ERRO_SEM_MENSAGEM

        If lErro = 17660 Then gError 92064

        'Coloca na tela
        FilialCliIntermediario.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 92065

    Exit Sub

Erro_FilialCliIntermediario_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 92062
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 92064
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, CliIntermediario.Text)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 92065
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialCliIntermediario.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156560)

    End Select

    Exit Sub

End Sub
