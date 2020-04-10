VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public gobjInfoUsu As Object
Dim giFrameAtualItens As Integer
Public gobjInfoAdic As ClassInfoAdic
Dim dValorDescontoItensAnt As Double
Dim dPercDescontoItensAnt As Double
Dim iDescontoAlterado As Integer

Public gobjAnotacao As ClassAnotacoes

Public gobjTelaProjetoInfo As ClassTelaPRJInfo

Const CLIENTE_INEXISTENTE = 1
Const FILIALCLIENTE_INEXISTENTE = 1
Const FILIALENTREGA_INEXISTENTE = 2

'Flag que indica se a tela está sendo preenchida.
Public gbCarregandoTela As Boolean
Public gbLimpandoTela As Boolean

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjPedidoVenda As ClassPedidoDeVenda 'estrutura mantida para auxiliar a manutencao de informacoes p/tributacao
'todos os dados relevantes p/tributacao dentro de gobjPedidoVenda estarao sincronizados com a tela antes da atualizacao da tributacao

'inicia objeto associado a GridComissoes
Dim objTabComissoes As New ClassTabComissoes

'declara o objeto da execucao das regras de comissoes e
'o objeto do calculo de comissoes
Public objComissoesRegrasCalc As Object
Public objMnemonicoComissCalc As ClassMnemonicoComissCalc
Public objMnemonicoComissCalcAux As ClassMnemonicoComissCalcAux

Dim giFrameAtual As Integer
'Dim giFrameAtualTributacao As Integer
Public iAlterado As Integer

Public gcolItemPedido As Collection
Public gcolItemPedidoSRV As Collection
Dim giClienteInexistente As Integer 'Luiz Nogueira em 26/01/04
Dim giFilialCliInexistente As Integer 'Luiz Nogueira em 26/01/04
Dim giClienteAlterado As Integer
Dim giFilialAlterada As Integer
Dim giFilialFatAlterada As Integer
Dim giFilialEntregaAlterada As Integer
Public gdDesconto As Double
Public giValorFreteAlterado As Integer
Public giValorSeguroAlterado As Integer
Public giValorDescontoAlterado As Integer
Public giValorDespesasAlterado  As Integer
Dim giDataReferenciaAlterada As Integer
Dim giNaturezaOpAlterada As Integer
Dim giDataEntregaAlterada As Integer

Public giValorDescontoManual As Integer
Dim giFilialEntregaManual As Integer

Public gobjTribTab As New ClassTribTab
'Dim giTrazendoTribTela As Integer
'Dim giTrazendoTribItemTela As Integer
'
'Dim giISSAliquotaAlterada As Integer
'Dim giISSValorAlterado As Integer
'Dim giValorIRRFAlterado As Integer
'Dim giTipoTributacaoAlterado As Integer
'Dim giAliqIRAlterada As Integer
'Dim iPISRetidoAlterado As Integer
'Dim iCOFINSRetidoAlterado As Integer
'Dim iCSLLRetidoAlterado As Integer
'
'Dim giNatOpItemAlterado As Integer
'Dim giTipoTributacaoItemAlterado As Integer
'Dim giICMSBaseItemAlterado As Integer
'Dim giICMSPercRedBaseItemAlterado As Integer
'Dim giICMSAliquotaItemAlterado As Integer
'Dim giICMSValorItemAlterado As Integer
'Dim giICMSSubstBaseItemAlterado As Integer
'Dim giICMSSubstAliquotaItemAlterado As Integer
'Dim giICMSSubstValorItemAlterado As Integer
'Dim giIPIBaseItemAlterado As Integer
'Dim giIPIPercRedBaseItemAlterado As Integer
'Dim giIPIAliquotaItemAlterado As Integer
'Dim giIPIValorItemAlterado As Integer

Dim giPercAcresFinAlterado As Integer

'Grid_Almoxarifado
Public iGrid_Item_Col As Integer
Public iGrid_ProdutoAlmox_Col As Integer
Public iGrid_UMEstoque_Col As Integer
Public iGrid_Almoxarifado_Col As Integer
Public iGrid_QuantReservar_Col As Integer
Public iGrid_QuantReserv_Col As Integer
Public iGrid_Validade_Col As Integer
Public iGrid_Responsavel_Col As Integer
    
'Grid ItensSRV
Public iGrid_ItemSRV_Col As Integer
Public iGrid_ServicoSRV_Col As Integer
Public iGrid_DescServicoSRV_Col As Integer
Public iGrid_UnidadeMedSRV_Col As Integer
Public iGrid_QuantPedidaSRV_Col As Integer
Public iGrid_QuantCancelSRV_Col As Integer
Public iGrid_ValorUnitarioSRV_Col As Integer
Public iGrid_PercDescSRV_Col As Integer
Public iGrid_DescontoSRV_Col As Integer
Public iGrid_PrecoTotalSRV_Col As Integer
Public iGrid_PrecoTotalSRVB_Col As Integer
Public iGrid_DataEntregaSRV_Col As Integer
Public iGrid_QuantFaturadaSRV_Col As Integer

'Grid Itens
Public iGrid_ItemProduto_Col As Integer
Public iGrid_Produto_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_QuantPedida_Col As Integer
Public iGrid_QuantCancel_Col As Integer
Public iGrid_ValorUnitario_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PrecoTotal_Col As Integer
Public iGrid_PrecoTotalB_Col As Integer
Public iGrid_QuantReservada_Col As Integer
Public iGrid_QuantFaturada_Col As Integer

'Grid Parcelas
Public iGrid_Vencimento_Col  As Integer
Public iGrid_ValorParcela_Col As Integer
Public iGrid_TipoPagto_Col As Integer
Public iGrid_Desc1Codigo_Col As Integer
Public iGrid_Desc1Ate_Col As Integer
Public iGrid_Desc1Valor_Col As Integer
Public iGrid_Desc1Perc_Col As Integer
Public iGrid_Desc2Codigo_Col As Integer
Public iGrid_Desc2Ate_Col As Integer
Public iGrid_Desc2Valor_Col As Integer
Public iGrid_Desc2Perc_Col As Integer
Public iGrid_Desc3Codigo_Col As Integer
Public iGrid_Desc3Ate_Col As Integer
Public iGrid_Desc3Valor_Col As Integer
Public iGrid_Desc3Perc_Col As Integer

Dim gobjTelaTrataBloqGen As ClassTelaTrataBloqGen

Public objGridComissoes As AdmGrid
Public objGridItensSRV As AdmGrid
Public objGridItens As AdmGrid

Dim objGridParcelas As AdmGrid
Dim objGridAlocacao As AdmGrid

Private WithEvents objEventoOrcamento  As AdmEvento 'Incluido por Leo em 28/03/02
Attribute objEventoOrcamento.VB_VarHelpID = -1

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoCondPagto As AdmEvento
Attribute objEventoCondPagto.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1
Private WithEvents objEventoEstoqueProd As AdmEvento
Attribute objEventoEstoqueProd.VB_VarHelpID = -1
Public WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoNaturezaOp As AdmEvento
Attribute objEventoNaturezaOp.VB_VarHelpID = -1
'Private WithEvents objEventoTiposDeTributacao As AdmEvento
Private WithEvents objEventoMensagem As AdmEvento
Attribute objEventoMensagem.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoTranspRedespacho As AdmEvento
Attribute objEventoTranspRedespacho.VB_VarHelpID = -1
Private WithEvents objEventoContato As AdmEvento
Attribute objEventoContato.VB_VarHelpID = -1
Private WithEvents objEventoCodSolicSRV As AdmEvento
Attribute objEventoCodSolicSRV.VB_VarHelpID = -1

Dim giLinhaAnterior As Integer
'Dim giRecalculandoTributacao As Integer
'Dim gcolTiposTribICMS As New Collection
'Dim gcolTiposTribIPI As New Collection
Private giPosCargaOk As Integer

'Dim gsCodigoAnterior As String  'guarda Codigo anterior 'Trecho comentado por Leo em 18/02/02

'Constantes públicas dos tabs
Private Const TAB_Principal = 1
Private Const TAB_Itens = 2
Private Const TAB_Complemento = 3
Private Const TAB_Cobranca = 4
Private Const TAB_Comissoes = 5
Private Const TAB_Bloqueio = 6
'Private Const TAB_Almoxarifado = 7
Private Const TAB_Tributacao = 7

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjNFiscalCTB As ClassNFiscal
Private gobjPedidoVendaCTB As ClassPedidoDeVenda
Private giExercicio As Integer, giPeriodo As Integer
Private gcolAlmoxFilial As New Collection
Private gobjGeracaoNFiscal As ClassGeracaoNFiscal

'mnemonicos
Private Const CODIGO1 As String = "Codigo"
Private Const NATUREZA_OP As String = "Natureza_OP"
Private Const CLIENTE1 As String = "Cliente"
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
Private Const CTACONTABILEST1 As String = "ContaContabilEst"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque" '??? ERRADO: nao está no bd
'fim da contabilidade

'workflow
Private Const CREDITO_ULTRAPASSADO As String = "CreditoUltrapassado"
Private Const CODIGOPEDIDO As String = "CodigoPedido"
Private Const PRODUTO_SEM_RESERVA As String = "ProdutoSemReserva"
Private Const CLIENTE_NOME As String = "Cliente_Nome"

'Incluido por Leo em 24/04/02. Para o  funcionamento do tab de comissoes
Public iComissoesAlterada As Integer

Private gcolParcPV As Collection

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

Public Property Get PrecoTotalB() As Object
     Set PrecoTotalB = objUserControl.Controls("PrecoTotalB")
End Property

Public Property Get PrecoTotalSRVB() As Object
     Set PrecoTotalSRVB = objUserControl.Controls("PrecoTotalSRVB")
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

Public Property Get Cubagem() As Object
     Set Cubagem = objUserControl.Controls("Cubagem")
End Property

Public Property Get UpDownEntregaPV() As Object
     Set UpDownEntregaPV = objUserControl.Controls("UpDownEntregaPV")
End Property

Public Property Get DataEntregaPV() As Object
     Set DataEntregaPV = objUserControl.Controls("DataEntregaPV")
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

Public Property Get FaturaIntegral() As Object
     Set FaturaIntegral = objUserControl.Controls("FaturaIntegral")
End Property

'Public Property Get Frame2() As Object
'     Set Frame2 = objUserControl.Controls("Frame2")
'End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get LabelCliente() As Object
     Set LabelCliente = objUserControl.Controls("LabelCliente")
End Property

Public Property Get CondicaoPagamento() As Object
     Set CondicaoPagamento = objUserControl.Controls("CondicaoPagamento")
End Property

Public Property Get TabelaPreco() As Object
     Set TabelaPreco = objUserControl.Controls("TabelaPreco")
End Property

Public Property Get PercAcrescFin() As Object
     Set PercAcrescFin = objUserControl.Controls("PercAcrescFin")
End Property

Public Property Get CondPagtoLabel() As Object
     Set CondPagtoLabel = objUserControl.Controls("CondPagtoLabel")
End Property

Public Property Get BotaoProxNum() As Object
     Set BotaoProxNum = objUserControl.Controls("BotaoProxNum")
End Property
Public Property Get BotaoNFiscal() As Object
     Set BotaoNFiscal = objUserControl.Controls("BotaoNFiscal")
End Property

Public Property Get FilialFaturamento() As Object
     Set FilialFaturamento = objUserControl.Controls("FilialFaturamento")
End Property

Public Property Get Codigo() As Object
     Set Codigo = objUserControl.Controls("Codigo")
End Property

Public Property Get UpDownEmissao() As Object
     Set UpDownEmissao = objUserControl.Controls("UpDownEmissao")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get NaturezaOp() As Object
     Set NaturezaOp = objUserControl.Controls("NaturezaOp")
End Property

Public Property Get NaturezaLabel() As Object
     Set NaturezaLabel = objUserControl.Controls("NaturezaLabel")
End Property

Public Property Get NumeroLabel() As Object
     Set NumeroLabel = objUserControl.Controls("NumeroLabel")
End Property

'Public Property Get Frame4() As Object
'     Set Frame4 = objUserControl.Controls("Frame4")
'End Property

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

'Public Property Get Frame6() As Object
'     Set Frame6 = objUserControl.Controls("Frame6")
'End Property

Public Property Get Emitente() As Object
     Set Emitente = objUserControl.Controls("Emitente")
End Property

Public Property Get Destinatario() As Object
     Set Destinatario = objUserControl.Controls("Destinatario")
End Property

Public Property Get TransportadoraLabel() As Object
     Set TransportadoraLabel = objUserControl.Controls("TransportadoraLabel")
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

Public Property Get PedidoCliente() As Object
     Set PedidoCliente = objUserControl.Controls("PedidoCliente")
End Property

Public Property Get PesoBruto() As Object
     Set PesoBruto = objUserControl.Controls("PesoBruto")
End Property

Public Property Get PesoLiquido() As Object
     Set PesoLiquido = objUserControl.Controls("PesoLiquido")
End Property

Public Property Get MensagemLabel() As Object
     Set MensagemLabel = objUserControl.Controls("MensagemLabel")
End Property

Public Property Get CanalVendaLabel() As Object
     Set CanalVendaLabel = objUserControl.Controls("CanalVendaLabel")
End Property

'Public Property Get Frame8() As Object
'     Set Frame8 = objUserControl.Controls("Frame8")
'End Property

Public Property Get VolumeNumero() As Object
     Set VolumeNumero = objUserControl.Controls("VolumeNumero")
End Property

Public Property Get VolumeEspecie() As Object
     Set VolumeEspecie = objUserControl.Controls("VolumeEspecie")
End Property

Public Property Get VolumeMarca() As Object
     Set VolumeMarca = objUserControl.Controls("VolumeMarca")
End Property

Public Property Get VolumeQuant() As Object
     Set VolumeQuant = objUserControl.Controls("VolumeQuant")
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

Public Property Get TipoDesconto1() As Object
     Set TipoDesconto1 = objUserControl.Controls("TipoDesconto1")
End Property

Public Property Get TipoDesconto2() As Object
     Set TipoDesconto2 = objUserControl.Controls("TipoDesconto2")
End Property

Public Property Get TipoDesconto3() As Object
     Set TipoDesconto3 = objUserControl.Controls("TipoDesconto3")
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

Public Property Get DataReferencia() As Object
     Set DataReferencia = objUserControl.Controls("DataReferencia")
End Property

Public Property Get GridParcelas() As Object
     Set GridParcelas = objUserControl.Controls("GridParcelas")
End Property

Public Property Get BotaoGerarNFiscal() As Object
     Set BotaoGerarNFiscal = objUserControl.Controls("BotaoGerarNFiscal")
End Property

Public Property Get BotaoBaixarPedido() As Object
     Set BotaoBaixarPedido = objUserControl.Controls("BotaoBaixarPedido")
End Property

Public Property Get BotaoLibera() As Object
     Set BotaoLibera = objUserControl.Controls("BotaoLibera")
End Property

Public Property Get BotaoReserva() As Object
     Set BotaoReserva = objUserControl.Controls("BotaoReserva")
End Property

Public Property Get Frame7() As Object
     Set Frame7 = objUserControl.Controls("Frame7")
End Property

Public Property Get Responsavel() As Object
     Set Responsavel = objUserControl.Controls("Responsavel")
End Property

Public Property Get UnidadeMedEst() As Object
     Set UnidadeMedEst = objUserControl.Controls("UnidadeMedEst")
End Property

Public Property Get DataValidade() As Object
     Set DataValidade = objUserControl.Controls("DataValidade")
End Property

Public Property Get ProdutoAlmox() As Object
     Set ProdutoAlmox = objUserControl.Controls("ProdutoAlmox")
End Property

Public Property Get Almox() As Object
     Set Almox = objUserControl.Controls("Almox")
End Property

Public Property Get QuantReservar() As Object
     Set QuantReservar = objUserControl.Controls("QuantReservar")
End Property

Public Property Get QuantReservada() As Object
     Set QuantReservada = objUserControl.Controls("QuantReservada")
End Property

Public Property Get ItemPedido() As Object
     Set ItemPedido = objUserControl.Controls("ItemPedido")
End Property

Public Property Get GridAlocacao() As Object
     Set GridAlocacao = objUserControl.Controls("GridAlocacao")
End Property

'Public Property Get Frame3() As Object
'     Set Frame3 = objUserControl.Controls("Frame3")
'End Property

Public Property Get ProdutoDescricao() As Object
     Set ProdutoDescricao = objUserControl.Controls("ProdutoDescricao")
End Property

Public Property Get BotaoLiberaBloqueio() As Object
     Set BotaoLiberaBloqueio = objUserControl.Controls("BLGBotaoLibBloq")
End Property

Public Property Get TipoBloqueio() As Object
     Set TipoBloqueio = objUserControl.Controls("BLGTipoBloqueio")
End Property

Public Property Get ResponsavelLib() As Object
     Set ResponsavelLib = objUserControl.Controls("BLGResponsavelLib")
End Property

Public Property Get DataLiberacao() As Object
     Set DataLiberacao = objUserControl.Controls("BLGDataLiberacao")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("BLGObservacao")
End Property

Public Property Get SeqBloqueio() As Object
     Set SeqBloqueio = objUserControl.Controls("BLGSeqBloqueio")
End Property

Public Property Get CodUsuario() As Object
     Set CodUsuario = objUserControl.Controls("BLGUsuario")
End Property

Public Property Get ResponsavelBL() As Object
     Set ResponsavelBL = objUserControl.Controls("BLGResponsavel")
End Property

Public Property Get DataBloqueio() As Object
     Set DataBloqueio = objUserControl.Controls("BLGDataBloqueio")
End Property

Public Property Get GridBloqueio() As Object
     Set GridBloqueio = objUserControl.Controls("BLGGridBloqueio")
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

Public Property Get PercentualBaixa() As Object
     Set PercentualBaixa = objUserControl.Controls("PercentualBaixa")
End Property

Public Property Get GridComissoes() As Object
     Set GridComissoes = objUserControl.Controls("GridComissoes")
End Property

Public Property Get GridEmb() As Object
     Set GridEmb = objUserControl.Controls("GridEmb")
End Property

Public Property Get ValorBaixa() As Object
     Set ValorBaixa = objUserControl.Controls("ValorBaixa")
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

Public Property Get DescricaoProduto() As Object
     Set DescricaoProduto = objUserControl.Controls("DescricaoProduto")
End Property

Public Property Get UnidadeMed() As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
End Property

Public Property Get QuantCancelada() As Object
     Set QuantCancelada = objUserControl.Controls("QuantCancelada")
End Property

Public Property Get QuantFaturada() As Object
     Set QuantFaturada = objUserControl.Controls("QuantFaturada")
End Property

Public Property Get QuantReservadaPV() As Object
     Set QuantReservadaPV = objUserControl.Controls("QuantReservadaPV")
End Property

Public Property Get DataEntrega() As Object
     Set DataEntrega = objUserControl.Controls("DataEntrega")
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

Public Property Get ValorTotal() As Object
     Set ValorTotal = objUserControl.Controls("ValorTotal")
End Property

'Public Property Get FrameTributacao() As Object
'     Set FrameTributacao = objUserControl.Controls("FrameTributacao")
'End Property
'
'Public Property Get Frame9() As Object
'     Set Frame9 = objUserControl.Controls("Frame9")
'End Property
'
'Public Property Get Frame10() As Object
'     Set Frame10 = objUserControl.Controls("Frame10")
'End Property
'
'Public Property Get ICMSSubstValor() As Object
'     Set ICMSSubstValor = objUserControl.Controls("ICMSSubstValor")
'End Property
'
'Public Property Get ICMSSubstBase() As Object
'     Set ICMSSubstBase = objUserControl.Controls("ICMSSubstBase")
'End Property
'
'Public Property Get ICMSValor() As Object
'     Set ICMSValor = objUserControl.Controls("ICMSValor")
'End Property
'
'Public Property Get ICMSBase() As Object
'     Set ICMSBase = objUserControl.Controls("ICMSBase")
'End Property
'
'Public Property Get Frame11() As Object
'     Set Frame11 = objUserControl.Controls("Frame11")
'End Property
'
'Public Property Get IPIValor() As Object
'     Set IPIValor = objUserControl.Controls("IPIValor")
'End Property
'
'Public Property Get IPIBase() As Object
'     Set IPIBase = objUserControl.Controls("IPIBase")
'End Property
'
'Public Property Get Frame12() As Object
'     Set Frame12 = objUserControl.Controls("Frame12")
'End Property
'
'Public Property Get IRAliquota() As Object
'     Set IRAliquota = objUserControl.Controls("IRAliquota")
'End Property
'
'Public Property Get ValorIRRF() As Object
'     Set ValorIRRF = objUserControl.Controls("ValorIRRF")
'End Property
'
'Public Property Get PISRetido() As Object
'     Set PISRetido = objUserControl.Controls("PISRetido")
'End Property
'
'Public Property Get COFINSRetido() As Object
'     Set COFINSRetido = objUserControl.Controls("COFINSRetido")
'End Property
'
'Public Property Get CSLLRetido() As Object
'     Set CSLLRetido = objUserControl.Controls("CSLLRetido")
'End Property
'
'Public Property Get IRBase() As Object
'     Set IRBase = objUserControl.Controls("IRBase")
'End Property
'
'Public Property Get Frame13() As Object
'     Set Frame13 = objUserControl.Controls("Frame13")
'End Property

Public Property Get CalculoAuto() As Object
     Set CalculoAuto = objUserControl.Controls("CalculoAuto")
End Property

'Public Property Get ISSIncluso() As Object
'     Set ISSIncluso = objUserControl.Controls("ISSIncluso")
'End Property
'
'Public Property Get ISSAliquota() As Object
'     Set ISSAliquota = objUserControl.Controls("ISSAliquota")
'End Property
'
'Public Property Get ISSValor() As Object
'     Set ISSValor = objUserControl.Controls("ISSValor")
'End Property
'
'Public Property Get ISSBase() As Object
'     Set ISSBase = objUserControl.Controls("ISSBase")
'End Property
'
'Public Property Get TributacaoRecalcular() As Object
'     Set TributacaoRecalcular = objUserControl.Controls("TributacaoRecalcular")
'End Property
'
'Public Property Get TipoTributacao() As Object
'     Set TipoTributacao = objUserControl.Controls("TipoTributacao")
'End Property
'
'Public Property Get LblNatOpEspelho() As Object
'     Set LblNatOpEspelho = objUserControl.Controls("LblNatOpEspelho")
'End Property
'
'Public Property Get DescNatOp() As Object
'     Set DescNatOp = objUserControl.Controls("DescNatOp")
'End Property
'
'Public Property Get NatOpEspelho() As Object
'     Set NatOpEspelho = objUserControl.Controls("NatOpEspelho")
'End Property
'
'Public Property Get LblTipoTrib() As Object
'     Set LblTipoTrib = objUserControl.Controls("LblTipoTrib")
'End Property
'
'Public Property Get DescTipoTrib() As Object
'     Set DescTipoTrib = objUserControl.Controls("DescTipoTrib")
'End Property
'
'Public Property Get FrameOutrosTrib() As Object
'     Set FrameOutrosTrib = objUserControl.Controls("FrameOutrosTrib")
'End Property
'
'Public Property Get LabelValorOutrasDespesas() As Object
'     Set LabelValorOutrasDespesas = objUserControl.Controls("LabelValorOutrasDespesas")
'End Property
'
'Public Property Get LabelValorSeguro() As Object
'     Set LabelValorSeguro = objUserControl.Controls("LabelValorSeguro")
'End Property
'
'Public Property Get LabelValorDesconto() As Object
'     Set LabelValorDesconto = objUserControl.Controls("LabelValorDesconto")
'End Property
'
'Public Property Get LabelValorFrete() As Object
'     Set LabelValorFrete = objUserControl.Controls("LabelValorFrete")
'End Property
'
'Public Property Get FrameItensTrib() As Object
'     Set FrameItensTrib = objUserControl.Controls("FrameItensTrib")
'End Property
'
'Public Property Get ComboItensTrib() As Object
'     Set ComboItensTrib = objUserControl.Controls("ComboItensTrib")
'End Property
'
'Public Property Get LabelUMItem() As Object
'     Set LabelUMItem = objUserControl.Controls("LabelUMItem")
'End Property
'
'Public Property Get LabelQtdeItem() As Object
'     Set LabelQtdeItem = objUserControl.Controls("LabelQtdeItem")
'End Property
'
'Public Property Get LabelValorItem() As Object
'     Set LabelValorItem = objUserControl.Controls("LabelValorItem")
'End Property
'
'Public Property Get TribSobreOutrasDesp() As Object
'     Set TribSobreOutrasDesp = objUserControl.Controls("TribSobreOutrasDesp")
'End Property
'
'Public Property Get TribSobreSeguro() As Object
'     Set TribSobreSeguro = objUserControl.Controls("TribSobreSeguro")
'End Property
'
'Public Property Get TribSobreDesconto() As Object
'     Set TribSobreDesconto = objUserControl.Controls("TribSobreDesconto")
'End Property
'
'Public Property Get TribSobreFrete() As Object
'     Set TribSobreFrete = objUserControl.Controls("TribSobreFrete")
'End Property
'
'Public Property Get TribSobreItem() As Object
'     Set TribSobreItem = objUserControl.Controls("TribSobreItem")
'End Property
'
'Public Property Get Frame15() As Object
'     Set Frame15 = objUserControl.Controls("Frame15")
'End Property
'
'Public Property Get IPIItemFrame() As Object
'     Set IPIItemFrame = objUserControl.Controls("IPIItemFrame")
'End Property
'
'Public Property Get ComboIPITipo() As Object
'     Set ComboIPITipo = objUserControl.Controls("ComboIPITipo")
'End Property
'
'Public Property Get IPIPercRedBaseItem() As Object
'     Set IPIPercRedBaseItem = objUserControl.Controls("IPIPercRedBaseItem")
'End Property
'
'Public Property Get IPIValorItem() As Object
'     Set IPIValorItem = objUserControl.Controls("IPIValorItem")
'End Property
'
'Public Property Get IPIAliquotaItem() As Object
'     Set IPIAliquotaItem = objUserControl.Controls("IPIAliquotaItem")
'End Property
'
'Public Property Get IPIBaseItem() As Object
'     Set IPIBaseItem = objUserControl.Controls("IPIBaseItem")
'End Property
'
'Public Property Get ICMSSubstValorItem() As Object
'     Set ICMSSubstValorItem = objUserControl.Controls("ICMSSubstValorItem")
'End Property
'
'Public Property Get ICMSSubstAliquotaItem() As Object
'     Set ICMSSubstAliquotaItem = objUserControl.Controls("ICMSSubstAliquotaItem")
'End Property
'
'Public Property Get ICMSSubstBaseItem() As Object
'     Set ICMSSubstBaseItem = objUserControl.Controls("ICMSSubstBaseItem")
'End Property
'
'Public Property Get ComboICMSTipo() As Object
'     Set ComboICMSTipo = objUserControl.Controls("ComboICMSTipo")
'End Property
'
'Public Property Get ICMSValorItem() As Object
'     Set ICMSValorItem = objUserControl.Controls("ICMSValorItem")
'End Property
'
'Public Property Get ICMSAliquotaItem() As Object
'     Set ICMSAliquotaItem = objUserControl.Controls("ICMSAliquotaItem")
'End Property
'
'Public Property Get ICMSPercRedBaseItem() As Object
'     Set ICMSPercRedBaseItem = objUserControl.Controls("ICMSPercRedBaseItem")
'End Property
'
'Public Property Get ICMSBaseItem() As Object
'     Set ICMSBaseItem = objUserControl.Controls("ICMSBaseItem")
'End Property
'
'Public Property Get NaturezaOpItem() As Object
'     Set NaturezaOpItem = objUserControl.Controls("NaturezaOpItem")
'End Property
'
'Public Property Get TipoTributacaoItem() As Object
'     Set TipoTributacaoItem = objUserControl.Controls("TipoTributacaoItem")
'End Property
'
'Public Property Get NaturezaItemLabel() As Object
'     Set NaturezaItemLabel = objUserControl.Controls("NaturezaItemLabel")
'End Property
'
'Public Property Get DescTipoTribItem() As Object
'     Set DescTipoTribItem = objUserControl.Controls("DescTipoTribItem")
'End Property
'
'Public Property Get LabelDescrNatOpItem() As Object
'     Set LabelDescrNatOpItem = objUserControl.Controls("LabelDescrNatOpItem")
'End Property
'
'Public Property Get LblTipoTribItem() As Object
'     Set LblTipoTribItem = objUserControl.Controls("LblTipoTribItem")
'End Property
'
'Public Property Get OpcaoTributacao() As Object
'     Set OpcaoTributacao = objUserControl.Controls("OpcaoTributacao")
'End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoImprimir() As Object
     Set BotaoImprimir = objUserControl.Controls("BotaoImprimir")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get Opcao() As Object
     Set Opcao = objUserControl.Controls("Opcao")
End Property

Public Property Get PedidoRepr() As Object
     Set PedidoRepr = objUserControl.Controls("PedidoRepr")
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

Public Property Get ServicoSRV() As Object
     Set ServicoSRV = objUserControl.Controls("ServicoSRV")
End Property

Public Property Get DescServicoSRV() As Object
     Set DescServicoSRV = objUserControl.Controls("DescServicoSRV")
End Property

Public Property Get UnidadeMedSRV() As Object
     Set UnidadeMedSRV = objUserControl.Controls("UnidadeMedSRV")
End Property

Public Property Get QuantidadeSRV() As Object
     Set QuantidadeSRV = objUserControl.Controls("QuantidadeSRV")
End Property

Public Property Get QuantCanceladaSRV() As Object
     Set QuantCanceladaSRV = objUserControl.Controls("QuantCanceladaSRV")
End Property

Public Property Get QuantFaturadaSRV() As Object
     Set QuantFaturadaSRV = objUserControl.Controls("QuantFaturadaSRV")
End Property

Public Property Get DataEntregaSRV() As Object
     Set DataEntregaSRV = objUserControl.Controls("DataEntregaSRV")
End Property

Public Property Get DescontoSRV() As Object
     Set DescontoSRV = objUserControl.Controls("DescontoSRV")
End Property

Public Property Get PercentDescSRV() As Object
     Set PercentDescSRV = objUserControl.Controls("PercentDescSRV")
End Property

Public Property Get PrecoUnitarioSRV() As Object
     Set PrecoUnitarioSRV = objUserControl.Controls("PrecoUnitarioSRV")
End Property

Public Property Get PrecoTotalSRV() As Object
     Set PrecoTotalSRV = objUserControl.Controls("PrecoTotalSRV")
End Property

Public Property Get GridItensSRV() As Object
     Set GridItensSRV = objUserControl.Controls("GridItensSRV")
End Property

Public Property Get BotaoEstoquePeca() As Object
     Set BotaoEstoquePeca = objUserControl.Controls("BotaoEstoquePeca")
End Property

Public Property Get BotaoServicos() As Object
     Set BotaoServicos = objUserControl.Controls("BotaoServicos")
End Property

Public Property Get BotaoPecas() As Object
     Set BotaoPecas = objUserControl.Controls("BotaoPecas")
End Property

Public Property Get TabStripSRV() As Object
     Set TabStripSRV = objUserControl.Controls("TabStripSRV")
End Property

Public Property Get FrameItens() As Object
     Set FrameItens = objUserControl.Controls("FrameItens")
End Property

Public Property Get CodSolicSRV() As Object
     Set CodSolicSRV = objUserControl.Controls("CodSolicSRV")
End Property

Public Property Get Status() As Object
     Set Status = objUserControl.Controls("Status")
End Property

Public Property Get BotaoDesfazer() As Object
     Set BotaoDesfazer = objUserControl.Controls("BotaoDesfazer")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoDataReferenciaDown_Click()

Dim lErro As Long
Dim bCancel As Boolean
Dim sData As String

On Error GoTo Erro_BotaoDataReferenciaDown_Click

    sData = DataReferencia.Text

    'diminui a data em um dia
    lErro = Data_Up_Down_Click(DataReferencia, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 188766

    Call DataReferencia_Validate(bCancel)

    If bCancel = True Then DataReferencia.Text = sData

    Exit Sub

Erro_BotaoDataReferenciaDown_Click:

    Select Case gErr

        Case 188766

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188767)

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
    If lErro <> SUCESSO Then gError 188768

    Call DataReferencia_Validate(bCancel)

    If bCancel = True Then DataReferencia.Text = sData

    Exit Sub

Erro_BotaoDataReferenciaUp_Click:

    Select Case gErr

        Case 188768

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188769)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long
Dim iIndice As Integer

On Error GoTo Erro_BotaoProxNum_Click

    If Codigo.Enabled = True Then

        lErro = CF("PedidoServico_Automatico", lCodigo)
        If lErro <> SUCESSO Then gError 188772

        Codigo.Text = CStr(lCodigo)

        'limpa a quantidade faturada do grid
        For iIndice = 1 To objGridItens.iLinhasExistentes

            GridItens.TextMatrix(iIndice, iGrid_QuantFaturada_Col) = ""

        Next

        For iIndice = 1 To objGridItensSRV.iLinhasExistentes

            GridItensSRV.TextMatrix(iIndice, iGrid_QuantFaturadaSRV_Col) = ""

        Next

        Call gobjTelaTrataBloqGen.Limpa

    Else

        'Codigo desabilitado não muda
        gError 188773

    End If


    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 188772

        Case 188773
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_DESABILITADO_IMUTAVEL", gErr, Error)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188774)

    End Select

    Exit Sub

End Sub

Public Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub DataEmissao_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Public Sub DataReferencia_GotFocus()
Dim iDataAux As Integer

    iDataAux = giDataReferenciaAlterada
    Call MaskEdBox_TrataGotFocus(DataReferencia, iAlterado)
    giDataReferenciaAlterada = iDataAux

End Sub

Public Sub Destinatario_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Emitente_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialFaturamento_Change()

    iAlterado = REGISTRO_ALTERADO
    giFilialFatAlterada = 1

End Sub

Public Sub FilialFaturamento_Click()

Dim lErro As Long

On Error GoTo Erro_FilialFaturamento_Click

    iAlterado = REGISTRO_ALTERADO

    'Se for FilialFaturamento selecionada faz tratamento de impostos
    If FilialFaturamento.ListIndex <> -1 Then
        If FilialFaturamento_TrataAlteracao() <> SUCESSO Then gError 188775
    End If

    Exit Sub

Erro_FilialFaturamento_Click:

    Select Case gErr

        Case 188775

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188776)

    End Select

    Exit Sub

End Sub

Public Function CargaPosFormLoad(Optional bTrazendoDoc As Boolean = False) As Long

Dim lErro As Long

On Error GoTo Erro_CargaPosFormLoad

    If (giPosCargaOk = 0) Then

        'p/permitir o redesenho da tela
        DoEvents
        
        Set gobjTelaTrataBloqGen = New ClassTelaTrataBloqGen
        
        lErro = gobjTelaTrataBloqGen.Inicializa(MAPBLOQGEN_TIPOTELA_PEDIDOSRV, Me, True)
        If lErro <> SUCESSO Then gError 188776
        
        gbCarregandoTela = True

        'Alterado por Tulio em 27/03

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
        
        ''*** incluidos p/tratamento de tributacao *******************************
        Set gobjTribTab = New ClassTribTab
        lErro = gobjTribTab.Ativar(Me, , , gobjTribTab.TIPOTELA_PV, , True)
        If lErro <> SUCESSO Then gError 188776
        
        Set gobjPedidoVenda = New ClassPedidoDeVenda
        gobjPedidoVenda.iServico = MARCADO
        gobjPedidoVenda.dtDataEmissao = gdtDataHoje
        lErro = gobjTribTab.TributacaoNF_Reset(gobjPedidoVenda)
        If lErro <> SUCESSO Then gError 188776
        '*** fim tributacao
        
''''        lErro = TributacaoPV_Reset()
''''        If lErro <> SUCESSO Then gError 188776

        'Carrega Filiais de Faturamento
        lErro = Carrega_FilialFaturamento()
        If lErro <> SUCESSO Then gError 188777

        'Call BotaoGravarTribCarga

       If bTrazendoDoc = False Then

            'Call CallByName(NaturezaOp, "Refresh", VbMethod)
            Call CallByName(FilialFaturamento, "Refresh", VbMethod)

        End If

'        lErro = CarregaTiposTrib()
'        If lErro <> SUCESSO Then gError 188778

        'Carrega os Estados
        lErro = Carrega_PlacaUF()
        If lErro <> SUCESSO Then gError 188779

        'Carrega a combo combo de Tabela de Preços
        lErro = Carrega_TabelaPreco()
        If lErro <> SUCESSO Then gError 188780

        'Carrega a combo de Transportadora com as Transportadoras
        lErro = Carrega_Transportadoras()
        If lErro <> SUCESSO Then gError 188781

        'Carrega a combo VolumeEspecie
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
        If lErro <> SUCESSO Then gError 188782

        'Carrega a combo VolumeMarca
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
        If lErro <> SUCESSO Then gError 188783

        'Carrega a combo de Canais de Venda
        lErro = Carrega_CanalVenda()
        If lErro <> SUCESSO Then gError 188784

'        'Carrega a combo de Tipos de Bloqueio
'        lErro = Carrega_TipoBloqueio()
'        If lErro <> SUCESSO Then gError 188785

        'Carrega a combo de Condição de Pagamento
        lErro = Carrega_CondicaoPagamento()
        If lErro <> SUCESSO Then gError 188786

        'Carrega a combo de Tipo de desconto
        lErro = Carrega_TipoDesconto()
        If lErro <> SUCESSO Then gError 188787

        PrecoUnitario.Format = gobjFAT.sFormatoPrecoUnitario

        Quantidade.Format = FORMATO_ESTOQUE
        QuantCancelada.Format = FORMATO_ESTOQUE
        
        PrecoUnitarioSRV.Format = gobjFAT.sFormatoPrecoUnitario

        QuantidadeSRV.Format = FORMATO_ESTOQUE
        QuantCanceladaSRV.Format = FORMATO_ESTOQUE

        'Preenche Data Referencia e Data de Emissão coma Data Atual
        DataReferencia.PromptInclude = False
        DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
        DataReferencia.PromptInclude = True
        giDataReferenciaAlterada = 0

        'Iniciliza as variáveis de browse
        Set objEventoMensagem = New AdmEvento
        Set objEventoCliente = New AdmEvento
        Set objEventoNumero = New AdmEvento
        Set objEventoCondPagto = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoServico = New AdmEvento
        Set objEventoEstoqueProd = New AdmEvento
        Set objEventoVendedor = New AdmEvento
        Set objEventoNaturezaOp = New AdmEvento
        'Set objEventoTiposDeTributacao = New AdmEvento
        Set objEventoTransportadora = New AdmEvento
        Set objEventoTranspRedespacho = New AdmEvento
        Set objEventoContato = New AdmEvento
        Set objEventoCodSolicSRV = New AdmEvento

        Set objEventoOrcamento = New AdmEvento

        Set objGridItens = New AdmGrid
        Set objGridItensSRV = New AdmGrid
        
        Set objGridComissoes = New AdmGrid
        Set objGridParcelas = New AdmGrid
'        Set objGridBloqueio = New AdmGrid
        Set objGridAlocacao = New AdmGrid

        Set gcolItemPedido = New Collection
        Set gcolItemPedidoSRV = New Collection

        'Faz as Inicializações dos Grids
        lErro = CF2(Me, "PSRV_Inicializa_Grid_Itens", objGridItens)
        If lErro <> SUCESSO Then gError 188788

        'Faz as Inicializações dos Grids
        lErro = CF2(Me, "PSRV_Inicializa_Grid_ItensSRV", objGridItensSRV)
        If lErro <> SUCESSO Then gError 188821

        lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
        If lErro <> SUCESSO Then gError 188789

        lErro = CF("Inicializa_Mascara_Produto_MaskEd", ServicoSRV)
        If lErro <> SUCESSO Then gError 188790

        Set objTabComissoes.objTela = Me

        lErro = objTabComissoes.Inicializa_Grid_Comissoes(objGridComissoes)
        If lErro <> SUCESSO Then gError 188791

        'alterado por cyntia
        objGridComissoes.iLinhasVisiveis = 5

        'Chama função que inicializa o Grid
        Call Grid_Inicializa(objGridComissoes)

        lErro = Inicializa_Grid_Parcelas(objGridParcelas)
        If lErro <> SUCESSO Then gError 188792

        lErro = CF2(Me, "PSRV_Inicializa_Grid_Parcelas", objGridParcelas)
        If lErro <> SUCESSO Then gError 188793

'        lErro = Inicializa_Grid_Bloqueio(objGridBloqueio)
'        If lErro <> SUCESSO Then gError 188794

        lErro = Inicializa_Grid_Reserva(objGridAlocacao)
        If lErro <> SUCESSO Then gError 188795

        gbCarregandoTela = False

        iComissoesAlterada = 0
        iDescontoAlterado = 0

        iAlterado = 0
        iComissoesAlterada = 0
        gbCarregandoTela = False

        giPosCargaOk = 1

        Call ValorTotal_Calcula

    End If

    CargaPosFormLoad = SUCESSO

    Exit Function

Erro_CargaPosFormLoad:

    gbCarregandoTela = False

    CargaPosFormLoad = gErr

    Select Case gErr

        Case 188776 To 188795, 188821

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188796)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    giPosCargaOk = 0

    giFrameAtual = 1
'    giFrameAtualTributacao = 1
    giFrameAtualItens = 1
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    Call Carrega_FilialOrcamento

    iAlterado = 0
    iComissoesAlterada = 0
    gbCarregandoTela = True

    Set gcolParcPV = New Collection

    Status.Caption = STRING_STATUS_ABERTO

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188797)

    End Select

    iAlterado = 0
    iComissoesAlterada = 0

    Exit Sub

End Sub
'
'Private Sub BotaoGravarTribCarga()
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoGravarTribCarga
'
'    lErro = Tributacao_GravarTela()
'    If lErro <> SUCESSO Then gError 188798
'
'    'Atualiza os valores de tributação
'    lErro = AtualizarTributacao()
'    If lErro <> SUCESSO Then gError 188799
'
'    Exit Sub
'
'Erro_BotaoGravarTribCarga:
'
'    Select Case gErr
'
'        Case 188798, 188799
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188800)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoGravarTrib()
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoGravarTrib
'
'    If giTrazendoTribTela = False Then
'
'        lErro = Tributacao_GravarTela()
'        If lErro <> SUCESSO Then gError 188801
'
'        lErro = ValorTotal_Calcula()
'        If lErro <> SUCESSO Then gError 188802
'
'        lErro = Carrega_Tab_Tributacao(gobjPedidoVenda)
'        If lErro <> SUCESSO Then gError 188803
'
'    End If
'
'    Exit Sub
'
'Erro_BotaoGravarTrib:
'
'    Select Case gErr
'
'        Case 188801 To 188803
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188804)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub BotaoGravarTribItem_Click()
'
'Dim lErro As Long, objTributacaoItemPV As ClassTributacaoItemPV, iIndice As Integer
'
'On Error GoTo Erro_BotaoGravarTribItem_Click
'
'    'atualizar dados da colecao p/o item ou complemento corrente
'
'    'se um item estiver selecionado
'    If TribSobreItem.Value = True Then
'        iIndice = ComboItensTrib.ListIndex
'        If iIndice <> -1 Then
'            If iIndice + 1 > gobjPedidoVenda.colItensPedidoSRV.Count Then
'                Set objTributacaoItemPV = gobjPedidoVenda.colItensPedido.Item(iIndice + 1 - gobjPedidoVenda.colItensPedidoSRV.Count).objTributacaoItemPV
'            Else
'                Set objTributacaoItemPV = gobjPedidoVenda.colItensPedidoSRV.Item(iIndice + 1).objTributacaoItemPV
'            End If
'
'        Else
'            gError 188805
'        End If
'    Else
''        If TribSobreDesconto.Value = True Then
''            Set objTributacaoItemPV = gobjPedidoVenda.objTributacaoPV.objTributacaoDesconto
''        Else
''            If TribSobreFrete.Value = True Then
''                Set objTributacaoItemPV = gobjPedidoVenda.objTributacaoPV.objTributacaoFrete
''            Else
''                If TribSobreSeguro.Value = True Then
''                    Set objTributacaoItemPV = gobjPedidoVenda.objTributacaoPV.objTributacaoSeguro
''                Else
''                    If TribSobreOutrasDesp.Value = True Then
''                        Set objTributacaoItemPV = gobjPedidoVenda.objTributacaoPV.objTributacaoOutras
''                    End If
''                End If
''            End If
''        End If
''XXXX
'    End If
'
'    lErro = TributacaoItem_GravarTela(objTributacaoItemPV)
'    If lErro <> SUCESSO Then gError 188806
'
'    lErro = ValorTotal_Calcula()
'    If lErro <> SUCESSO Then gError 188807
'
'    lErro = TributacaoItem_TrazerTela(objTributacaoItemPV)
'    If lErro <> SUCESSO Then gError 188808
'
'    Exit Sub
'
'Erro_BotaoGravarTribItem_Click:
'
'    Select Case gErr
'
'        Case 188805
'            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ITEM_TRIB_SEL", gErr, Error)
'
'        Case 188806 To 188808
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188809)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ComboICMSTipo_Click()
'
'    If ComboICMSTipo.ListIndex = -1 Then Exit Sub
'
'    If giTrazendoTribItemTela = 0 Then
'        Call BotaoGravarTribItem_Click
'    End If
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ComboIPITipo_Click()
'
'    If ComboIPITipo.ListIndex = -1 Then Exit Sub
'
'    If giTrazendoTribItemTela = 0 Then
'        Call BotaoGravarTribItem_Click
'    End If
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ComboItensTrib_Click()
'
'Dim iIndice As Integer, objItemPedido As ClassItemPedido
'
'    iIndice = ComboItensTrib.ListIndex
'
'    If iIndice <> -1 Then
'
'        'preenche os campos da tela em funcao do item selecionado
'
'        If iIndice + 1 > gobjPedidoVenda.colItensPedidoSRV.Count Then
'
'            Set objItemPedido = gobjPedidoVenda.colItensPedido.Item(iIndice + 1 - gobjPedidoVenda.colItensPedidoSRV.Count)
'
'        Else
'
'            Set objItemPedido = gobjPedidoVenda.colItensPedidoSRV.Item(iIndice + 1)
'
'        End If
'
'
'        LabelValorItem.Caption = Format(objItemPedido.dPrecoTotal, "Standard")
'        LabelQtdeItem.Caption = CStr(objItemPedido.dQuantidade - objItemPedido.dQuantCancelada)
'        LabelUMItem.Caption = objItemPedido.sUnidadeMed
'
'        Call TributacaoItem_TrazerTela(objItemPedido.objTributacaoItemPV)
'
'    End If
'
'End Sub

Private Function Carrega_TabelaPreco() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TabelaPreco

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePreco", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 188810

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

        Case 188810

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188811)

    End Select

    Exit Function

End Function

Private Function Carrega_Transportadoras() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_Transportadoras

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_TRANSPORTADORA_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 188812

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        Transportadora.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        Transportadora.ItemData(Transportadora.NewIndex) = objCodDescricao.iCodigo
        TranspRedespacho.AddItem CStr(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TranspRedespacho.ItemData(TranspRedespacho.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_Transportadoras = SUCESSO

    Exit Function

Erro_Carrega_Transportadoras:

    Carrega_Transportadoras = gErr

    Select Case gErr

        Case 188812

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188813)

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
    If lErro <> SUCESSO Then gError 188814

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        CanalVenda.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        CanalVenda.ItemData(CanalVenda.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_CanalVenda = SUCESSO

    Exit Function

Erro_Carrega_CanalVenda:

    Carrega_CanalVenda = gErr

    Select Case gErr

        Case 188814

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188815)

    End Select

    Exit Function

End Function
'
'Private Function Carrega_TipoBloqueio() As Long
'
'Dim colCodigoDescricao As New AdmColCodigoNome
'Dim objCodDescricao As AdmCodigoNome
'Dim lErro As Long
'
'On Error GoTo Erro_Carrega_TipoBloqueio
'
'    'Lê o código e a descrição de todas as Tabelas de Preços
'    lErro = CF("Cod_Nomes_Le", "TiposDeBloqueio", "Codigo", "NomeReduzido", STRING_TIPO_BLOQUEIO_NOME_REDUZIDO, colCodigoDescricao)
'    If lErro <> SUCESSO Then gError 188816
'
'    For Each objCodDescricao In colCodigoDescricao
'
'        If objCodDescricao.iCodigo <> BLOQUEIO_PARCIAL And objCodDescricao.iCodigo <> BLOQUEIO_NAO_RESERVA And objCodDescricao.iCodigo <> BLOQUEIO_CREDITO Then
'            'Adiciona o item na Lista de Tabela de Preços
'            TipoBloqueio.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
'            TipoBloqueio.ItemData(TipoBloqueio.NewIndex) = objCodDescricao.iCodigo
'        End If
'    Next
'
'    Carrega_TipoBloqueio = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_TipoBloqueio:
'
'    Carrega_TipoBloqueio = gErr
'
'    Select Case gErr
'
'        Case 188816
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188817)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Carrega_CondicaoPagamento() As Long

Dim lErro As Long
Dim colCod_DescReduzida As New AdmColCodigoNome
Dim objCod_DescReduzida As AdmCodigoNome

On Error GoTo Erro_Carrega_CondicaoPagamento

    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
    lErro = CF("CondicoesPagto_Le_Recebimento", colCod_DescReduzida)
    If lErro <> SUCESSO Then gError 188818

    For Each objCod_DescReduzida In colCod_DescReduzida

        'Adiciona novo item na List da Combo CondicaoPagamento
        CondicaoPagamento.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
        CondicaoPagamento.ItemData(CondicaoPagamento.NewIndex) = objCod_DescReduzida.iCodigo

    Next

    Carrega_CondicaoPagamento = SUCESSO

    Exit Function

Erro_Carrega_CondicaoPagamento:

    Carrega_CondicaoPagamento = gErr

    Select Case gErr

        Case 188818

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188819)

    End Select

    Exit Function

End Function

Private Function Carrega_TipoDesconto() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoDesconto

    Set colCodigoDescricao = gobjCRFAT.colTiposDesconto

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item nas List's das Combos de Tipos Desconto
        TipoDesconto1.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        TipoDesconto1.ItemData(TipoDesconto1.NewIndex) = objCodDescricao.iCodigo
        TipoDesconto2.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        TipoDesconto2.ItemData(TipoDesconto2.NewIndex) = objCodDescricao.iCodigo
        TipoDesconto3.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        TipoDesconto3.ItemData(TipoDesconto3.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TipoDesconto = SUCESSO

    Exit Function

Erro_Carrega_TipoDesconto:

    Carrega_TipoDesconto = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188820)

    End Select

    Exit Function

End Function

Public Function PSRV_Inicializa_Grid_ItensSRV(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Serviço")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant. Pedida")
    objGridInt.colColuna.Add ("Quant. Cancelada")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Data Entrega")
    objGridInt.colColuna.Add ("Quant Faturada")


    'Controles que participam do Grid
    objGridInt.colCampo.Add (ServicoSRV.Name)
    objGridInt.colCampo.Add (DescServicoSRV.Name)
    objGridInt.colCampo.Add (UnidadeMedSRV.Name)
    objGridInt.colCampo.Add (QuantidadeSRV.Name)
    objGridInt.colCampo.Add (QuantCanceladaSRV.Name)
    objGridInt.colCampo.Add (PrecoUnitarioSRV.Name)
    objGridInt.colCampo.Add (PercentDescSRV.Name)
    objGridInt.colCampo.Add (DescontoSRV.Name)
    objGridInt.colCampo.Add (PrecoTotalSRV.Name)
    objGridInt.colCampo.Add (PrecoTotalSRVB.Name)
    objGridInt.colCampo.Add (DataEntregaSRV.Name)
    objGridInt.colCampo.Add (QuantFaturadaSRV.Name)

    'Colunas do Grid
    iGrid_ItemSRV_Col = 0
    iGrid_ServicoSRV_Col = 1
    iGrid_DescServicoSRV_Col = 2
    iGrid_UnidadeMedSRV_Col = 3
    iGrid_QuantPedidaSRV_Col = 4
    iGrid_QuantCancelSRV_Col = 5
    iGrid_ValorUnitarioSRV_Col = 6
    iGrid_PercDescSRV_Col = 7
    iGrid_DescontoSRV_Col = 8
    iGrid_PrecoTotalSRV_Col = 9
    iGrid_PrecoTotalSRVB_Col = 10
    iGrid_DataEntregaSRV_Col = 11
    iGrid_QuantFaturadaSRV_Col = 12

    'Grid do GridInterno
    objGridInt.objGrid = GridItensSRV

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    PSRV_Inicializa_Grid_ItensSRV = SUCESSO

End Function

Public Function PSRV_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant. Pedida")
    objGridInt.colColuna.Add ("Quant. Cancelada")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Quant Reservada")
    objGridInt.colColuna.Add ("Quant Faturada")


    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoProduto.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (QuantCancelada.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    objGridInt.colCampo.Add (PrecoTotalB.Name)
    objGridInt.colCampo.Add (QuantReservadaPV.Name)
    objGridInt.colCampo.Add (QuantFaturada.Name)

    'Colunas do Grid
    iGrid_ItemProduto_Col = 0
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_QuantPedida_Col = 4
    iGrid_QuantCancel_Col = 5
    iGrid_ValorUnitario_Col = 6
    iGrid_PercDesc_Col = 7
    iGrid_Desconto_Col = 8
    iGrid_PrecoTotal_Col = 9
    iGrid_PrecoTotalB_Col = 10
    iGrid_QuantReservada_Col = 11
    iGrid_QuantFaturada_Col = 12

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    PSRV_Inicializa_Grid_Itens = SUCESSO

End Function

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Inicializa o Grid

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

    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (TipoPagto.Name)
    objGridInt.colCampo.Add (TipoDesconto1.Name)
    objGridInt.colCampo.Add (Desconto1Ate.Name)
    objGridInt.colCampo.Add (Desconto1Valor.Name)
    objGridInt.colCampo.Add (Desconto1Percentual.Name)
    objGridInt.colCampo.Add (TipoDesconto2.Name)
    objGridInt.colCampo.Add (Desconto2Ate.Name)
    objGridInt.colCampo.Add (Desconto2Valor.Name)
    objGridInt.colCampo.Add (Desconto2Percentual.Name)
    objGridInt.colCampo.Add (TipoDesconto3.Name)
    objGridInt.colCampo.Add (Desconto3Ate.Name)
    objGridInt.colCampo.Add (Desconto3Valor.Name)
    objGridInt.colCampo.Add (Desconto3Percentual.Name)


    'Controles que participam do Grid
    iGrid_Vencimento_Col = 1
    iGrid_ValorParcela_Col = 2
    iGrid_TipoPagto_Col = 3
    iGrid_Desc1Codigo_Col = 4
    iGrid_Desc1Ate_Col = 5
    iGrid_Desc1Valor_Col = 6
    iGrid_Desc1Perc_Col = 7
    iGrid_Desc2Codigo_Col = 8
    iGrid_Desc2Ate_Col = 9
    iGrid_Desc2Valor_Col = 10
    iGrid_Desc2Perc_Col = 11
    iGrid_Desc3Codigo_Col = 12
    iGrid_Desc3Ate_Col = 13
    iGrid_Desc3Valor_Col = 14
    iGrid_Desc3Perc_Col = 15

    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 700

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

'Private Function Inicializa_Grid_Bloqueio(objGridInt As AdmGrid) As Long
''Inicializa o Grid
'
'    'Form do Grid
'    Set objGridInt.objForm = Me
'
'    'Títulos das colunas
'    objGridInt.colColuna.Add (" ")
'    objGridInt.colColuna.Add ("Tipo")
'    objGridInt.colColuna.Add ("Data")
'    objGridInt.colColuna.Add ("Usuário")
'    objGridInt.colColuna.Add ("Responsável")
'    objGridInt.colColuna.Add ("Data Liberação")
'    objGridInt.colColuna.Add ("Resp. Liberação")
'    objGridInt.colColuna.Add ("Observação")
'    objGridInt.colColuna.Add ("Seq.")
'
'    'Controles que participam do Grid
'    objGridInt.colCampo.Add (TipoBloqueio.Name)
'    objGridInt.colCampo.Add (DataBloqueio.Name)
'    objGridInt.colCampo.Add (CodUsuario.Name)
'    objGridInt.colCampo.Add (ResponsavelBL.Name)
'    objGridInt.colCampo.Add (DataLiberacao.Name)
'    objGridInt.colCampo.Add (ResponsavelLib.Name)
'    objGridInt.colCampo.Add (Observacao.Name)
'    objGridInt.colCampo.Add (SeqBloqueio.Name)
'
'    iGrid_TipoBloqueio_Col = 1
'    iGrid_DataBloqueio_Col = 2
'    iGrid_CodUsuario_Col = 3
'    iGrid_ResponsavelBL_Col = 4
'    iGrid_DataLiberacao_Col = 5
'    iGrid_ResponsavelLib_Col = 6
'    iGrid_Observacao_Col = 7
'    iGrid_SeqBloqueio_Col = 8
'
'    'Grid do GridInterno
'    objGridInt.objGrid = GridBloqueio
'
'    'Todas as linhas do grid
'    objGridInt.objGrid.Rows = NUM_MAX_BLOQUEIOS + 1
'
'    'Linhas visíveis do grid
'    objGridInt.iLinhasVisiveis = 8
'
'    'Largura da primeira coluna
'    GridBloqueio.ColWidth(0) = 300
'
'    'Largura automática para as outras colunas
'    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
'
'    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
'
'    'Chama função que inicializa o Grid
'    Call Grid_Inicializa(objGridInt)
'
'    Inicializa_Grid_Bloqueio = SUCESSO
'
'    Exit Function
'
'End Function

Private Function Inicializa_Grid_Reserva(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Reserva

Dim lAlmoxarifado As Long

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Quant.a Reservar")
    objGridInt.colColuna.Add ("Quant. Reservada")
    objGridInt.colColuna.Add ("Validade")
    objGridInt.colColuna.Add ("Responsável")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ItemPedido.Name)
    objGridInt.colCampo.Add (ProdutoAlmox.Name)
    objGridInt.colCampo.Add (UnidadeMedEst.Name)
    objGridInt.colCampo.Add (Almox.Name)
    objGridInt.colCampo.Add (QuantReservar.Name)
    objGridInt.colCampo.Add (QuantReservada.Name)
    objGridInt.colCampo.Add (DataValidade.Name)
    objGridInt.colCampo.Add (Responsavel.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridAlocacao

    iGrid_Item_Col = 1
    iGrid_ProdutoAlmox_Col = 2
    iGrid_UMEstoque_Col = 3
    iGrid_Almoxarifado_Col = 4
    iGrid_QuantReservar_Col = 5
    iGrid_QuantReserv_Col = 6
    iGrid_Validade_Col = 7
    iGrid_Responsavel_Col = 8

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_RESERVAS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridAlocacao.ColWidth(0) = 300

    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
'    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Reserva = SUCESSO

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""
'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colParcelas As ColParcelaPagar

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "PedidoServico"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Pedido_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then gError 188822

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objPedidoVenda.lCodigo, 0, "Codigo"
    colCampoValor.Add "Cliente", objPedidoVenda.lCliente, 0, "Cliente"
    colCampoValor.Add "Filial", objPedidoVenda.iFilial, 0, "Filial"
    colCampoValor.Add "FilialEntrega", objPedidoVenda.iFilialEntrega, 0, "FilialEntrega"
    colCampoValor.Add "CodTransportadora", objPedidoVenda.iCodTransportadora, 0, "CodTransportadora"
    colCampoValor.Add "CodTranspRedesp", objPedidoVenda.iCodTranspRedesp, 0, "CodTranspRedesp"
    colCampoValor.Add "DetPagFrete", objPedidoVenda.iDetPagFrete, 0, "DetPagFrete"
    colCampoValor.Add "CondicaoPagto", objPedidoVenda.iCondicaoPagto, 0, "CondicaoPagto"
    colCampoValor.Add "PercAcrescFinanceiro", objPedidoVenda.dPercAcrescFinanceiro, 0, "PercAcrescFinanceiro"
    colCampoValor.Add "DataEmissao", objPedidoVenda.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "MensagemNota", objPedidoVenda.sMensagemNota, STRING_PEDIDOVENDA_MENSAGEM_NOTA, "MensagemNota"
    colCampoValor.Add "NaturezaOp", objPedidoVenda.sNaturezaOp, STRING_NATUREZAOP_CODIGO, "NaturezaOp"
    colCampoValor.Add "ValorTotal", objPedidoVenda.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "ValorFrete", objPedidoVenda.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorDesconto", objPedidoVenda.dValorDesconto, 0, "ValorDesconto"
    colCampoValor.Add "ValorSeguro", objPedidoVenda.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "PedidoCliente", objPedidoVenda.sPedidoCliente, STRING_PEDIDOVENDA_PEDIDO_CLIENTE, "PedidoCliente"
    colCampoValor.Add "CanalVenda", objPedidoVenda.iCanalVenda, 0, "CanalVenda"
    colCampoValor.Add "TabelaPreco", objPedidoVenda.iTabelaPreco, 0, "TabelaPreco"
    colCampoValor.Add "Placa", objPedidoVenda.sPlaca, STRING_PEDIDOVENDA_PLACA, "Placa"
    colCampoValor.Add "PlacaUF", objPedidoVenda.sPlacaUF, STRING_PEDIDOVENDA_PLACA, "PlacaUF"
    colCampoValor.Add "VolumeQuant", objPedidoVenda.lVolumeQuant, 0, "VolumeQuant"
    colCampoValor.Add "VolumeEspecie", objPedidoVenda.lVolumeEspecie, 0, "VolumeEspecie" 'Alterado por Luiz Nogueira em 21/08/03
    colCampoValor.Add "VolumeMarca", objPedidoVenda.lVolumeMarca, 0, "VolumeMarca" 'Alterado por Luiz Nogueira em 21/08/03
    colCampoValor.Add "VolumeNumero", objPedidoVenda.sVolumeNumero, STRING_PEDIDOVENDA_VOLUME_NUMERO, "VolumeNumero"
    colCampoValor.Add "FreteRespons", objPedidoVenda.iFreteRespons, 0, "FreteRespons"
    colCampoValor.Add "PesoLiq", objPedidoVenda.dPesoLiq, 0, "Pesoliq"
    colCampoValor.Add "PesoBruto", objPedidoVenda.dPesoBruto, 0, "PesoBruto"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 188822

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188823)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Tela_Preenche

    objPedidoVenda.lCodigo = colCampoValor.Item("Codigo").vValor
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa

    If objPedidoVenda.lCodigo <> 0 Then

        'Mostra os dados do Pedido de Venda na tela
        lErro = Traz_Pedido_Tela(objPedidoVenda)
        If lErro <> SUCESSO Then gError 188824

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 188824

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188825)

    End Select

    Exit Sub

End Sub

'Alterado de Private para Public por Wagner
Public Function Traz_Pedido_Tela(objPedidoVenda As ClassPedidoDeVenda, Optional objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Coloca na tela os dados do Pedido de Venda
'Caso a função esteja sendo chamada para carregar um OrcamentoVenda, receberá um objOrcamentoVenda. Por Leo em 25/04/02

Dim lErro As Long
Dim objFilial As AdmFiliais
Dim colEmbProdItem As New Collection
Dim iTipoDocOrigem As Integer
'Dim colItens As New Collection, objItens As ClassItemPedido Removido por Luiz Nogueira em 20/05/03
Dim lCodOrcamento As Long 'por leo em 25/04/02
Dim bProcessarInfoPedido As Boolean 'quando o pedido é carregado a partir de um orçamento, essa variável indica que é preciso completar dados do pedido, como reservas, comissões e embalagens
Dim objCliente As ClassCliente 'Luiz Nogueira 29/05/03
Dim iStatus As Integer
Dim objNFiscal As New ClassNFiscal 'Inserido por Wagner
Dim objParcelas As Object
Dim objParcPV As ClassParcelaPedidoVenda
Dim iIndice As Integer
Dim objSolicSRV As New ClassSolicSRV

On Error GoTo Erro_Traz_Pedido_Tela

    gbCarregandoTela = True

    Call Limpa_Tela_PedidoVenda2

    gbCarregandoTela = True

    'Se a função não for chamada para carregar os dados de um orçamento de venda.
    If (objOrcamentoVenda Is Nothing) Then

        'Lê os dados do Pedido de Venda
        lErro = CF("PedidoServico_Le_Todos_Completo", objPedidoVenda, iStatus)
        If lErro <> SUCESSO Then gError 188876

        Set objOrcamentoVenda = New ClassOrcamentoVenda

        objOrcamentoVenda.lNumIntPedVenda = objPedidoVenda.lNumIntDoc

        'Lê os dados do OrcamentoVenda através do seu NumIntPedVenda
        lErro = CF("OrcamentoSRV_Le_NumIntPedSRV", objOrcamentoVenda)
        If lErro <> SUCESSO And lErro <> 188869 Then gError 188877

    Else 'Incluído por Luiz Nogueira em 20/05/03

        'quando o pedido é carregado a partir de um orçamento, essa variável indica que é
        'preciso completar dados do pedido, como reservas, comissões e embalagens
        bProcessarInfoPedido = True

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

    'lErro = TributacaoPV_Reset(objPedidoVenda)
    lErro = gobjTribTab.Traz_NFiscal_Tela(objPedidoVenda)
    If lErro <> SUCESSO Then gError 188878

'    lErro = CF("ReservasPedidoSRV_Le", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 188879

    ValorTotal.Caption = Format(objPedidoVenda.dValorTotal, "Standard")
    ValorProdutos.Caption = Format(objPedidoVenda.dValorProdutos, "Standard")

    PercAcrescFin.Text = ""
    'Coloca os dados do Pedido na tela

    'Luiz Nogueira 26/01/04 - início
    'Se o código do cliente não está no obj,
    'significa que esse orçamento foi criado sem cadastrar o cliente
    If objPedidoVenda.lCliente = 0 Then

        'Cria uma nova instância da classe cliente
        Set objCliente = New ClassCliente

        'Guarda em objCliente o nome reduzido que foi informado no Orçamento
        objCliente.sNomeReduzido = objOrcamentoVenda.sNomeCli

        'Lê o cliente a partir do do nome reduzido que estava no orçamento
        'para verificar se o mesmo não foi criado após a geração do orçamento
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 188880

        'Se encontrou o cliente
        If lErro = SUCESSO Then

            'guarda o código dele em objPedidoVenda
            objPedidoVenda.lCliente = objCliente.lCodigo

        Else

            giClienteInexistente = CLIENTE_INEXISTENTE 'Incluído por Luiz Nogueira em 29/05/03
            Cliente.Text = objOrcamentoVenda.sNomeCli
            Filial.Clear
            Filial.Text = objOrcamentoVenda.sNomeFilialCli
            FilialEntrega.Clear
            FilialEntrega.Text = objOrcamentoVenda.sNomeFilialCli

        End If

        'Coloca clientes em atraso em vermelho
        If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)
    
    End If

    'Se o código do cliente está carregado em objPedidoVenda
    If objPedidoVenda.lCliente <> 0 Then

        Call Cliente_Formata(objPedidoVenda.lCliente)

        'Se a filial do cliente está carregada em objPedidoVenda
        If objPedidoVenda.iFilial <> 0 Then

            Call Filial_Formata(Filial, objPedidoVenda.iFilial)
            giFilialAlterada = 0

        'Senão, ou seja, se o pedido está sendo carregado a partir de um orçamento
        'onde a filial do cliente não existia qdo o mesmo foi cadastrado
        Else

            'Exibe na combo filialcliente o nome da filial do cliente que foi gravado no orçamento

            Filial.Text = objOrcamentoVenda.sNomeFilialCli

            'Exibe na combo filialentrega o nome da filial do cliente que foi gravado no orçamento
            FilialEntrega.Text = objOrcamentoVenda.sNomeFilialCli

            'Tenta selecionar a filial do cliente na combo
            lErro = Combo_Seleciona(Filial, 0)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 188881

            'Se não conseguiu selecionar
            If lErro = 6730 Or lErro = 6731 Then

                'Indica que a filial do cliente não foi encontrada para que seja exibido tratamento de erro posteriormente
                'O erro não pode ser tratado aqui, pois a tela deve ser carregada normalmente
                giFilialCliInexistente = FILIALCLIENTE_INEXISTENTE

            'Se conseguiu selecionar
            Else

                'Tenta selecionar a filialentrega na combo
                lErro = Combo_Seleciona(FilialEntrega, 0)
                If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 188882

                'Se não conseguiu selecionar => 'Indica que a filial de entrega não foi encontrada para que seja exibido tratamento de erro posteriormente
                If lErro = 6730 Or lErro = 6731 Then giFilialCliInexistente = FILIALENTREGA_INEXISTENTE

            End If


        End If

    End If
    'Luiz Nogueira 26/01/04 - fim

    For Each objFilial In gcolFiliais

        If objFilial.iCodFilial = objPedidoVenda.iFilialEmpresaFaturamento Then

            FilialFaturamento.Text = objFilial.iCodFilial & SEPARADOR & objFilial.sNome
            Exit For
        End If

    Next

    giFilialFatAlterada = 0

    If objPedidoVenda.lCodigo <> 0 Then 'Inserido por Leo em 20/04/02
        Codigo.Text = objPedidoVenda.lCodigo
    End If

'    NaturezaOp.Text = objPedidoVenda.sNaturezaOp
'    Call NaturezaOp_Validate(bSGECancelDummy)
    FaturaIntegral.Value = objPedidoVenda.iFaturaIntegral

    ValorFrete.Text = Format(objPedidoVenda.dValorFrete, "Standard")
    ValorSeguro.Text = Format(objPedidoVenda.dValorSeguro, "Standard")
    ValorDesconto.Text = Format(objPedidoVenda.dValorDesconto, "Standard")
    ValorDespesas.Text = Format(objPedidoVenda.dValorOutrasDespesas, "Standard")

    giValorFreteAlterado = 0
    giValorSeguroAlterado = 0
    giValorDescontoAlterado = 0
    giValorDespesasAlterado = 0

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(objPedidoVenda.dtDataEmissao, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    Call DateParaMasked(DataEntregaPV, objPedidoVenda.dtDataEntrega)
    giDataEntregaAlterada = 0

    'Se a tabela de preços estiver preenchida coloca na tela
    If objPedidoVenda.iTabelaPreco > 0 Then
        TabelaPreco.Text = objPedidoVenda.iTabelaPreco
        Call TabelaPreco_Validate(bSGECancelDummy)
    Else
        TabelaPreco.Text = ""
    End If

    Placa.Text = objPedidoVenda.sPlaca
    PlacaUF.Text = objPedidoVenda.sPlacaUF

    'Traz a espécie dos volumes do pedido
    If objPedidoVenda.lVolumeEspecie > 0 Then
        VolumeEspecie.Text = objPedidoVenda.lVolumeEspecie
        Call VolumeEspecie_Validate(bSGECancelDummy)
    Else
        'Se não está sendo carregado um orçamento, ou seja, está carregando
        'dados de um pedido, deve limpar a tela, pois o pedido foi gravado
        'sem espécie de volume
        If Not bProcessarInfoPedido Then VolumeEspecie.Text = ""
    End If

    'Traz a marca dos volumes do pedido
    If objPedidoVenda.lVolumeMarca > 0 Then
        VolumeMarca.Text = objPedidoVenda.lVolumeMarca
        Call VolumeMarca_Validate(bSGECancelDummy)
    Else
        'Se não está sendo carregado um orçamento, ou seja, está carregando
        'dados de um pedido, deve limpar a tela, pois o pedido foi gravado
        'sem marca de volume
        If Not bProcessarInfoPedido Then VolumeMarca.Text = ""
    End If

    VolumeNumero = objPedidoVenda.sVolumeNumero

    If objPedidoVenda.iFreteRespons = FRETE_EMITENTE Then
        Emitente.Value = True
    Else
        Destinatario.Value = True
    End If

    If objPedidoVenda.dPesoLiq > 0 Then PesoLiquido.Text = Format(objPedidoVenda.dPesoLiq, PesoLiquido.Format) 'Alterado por Wagner
    If objPedidoVenda.dPesoBruto > 0 Then PesoBruto.Text = Format(objPedidoVenda.dPesoBruto, PesoBruto.Format) 'Alterado por Wagner
    If objPedidoVenda.dVolumeTotal > 0 Then Cubagem.Text = Format(objPedidoVenda.dVolumeTotal, "Fixed")

    'Carrega o Tab Complemento
    lErro = Carrega_Tab_Complemento(objPedidoVenda)
    If lErro <> SUCESSO Then gError 188883

    'Carrega o Tab Cobrança
    lErro = Carrega_Tab_Cobranca(objPedidoVenda)
    If lErro <> SUCESSO Then gError 188884

    Set gcolParcPV = New Collection

    For iIndice = 1 To objPedidoVenda.colParcelas.Count

        gcolParcPV.Add objPedidoVenda.colParcelas(iIndice)
    Next

    If objPedidoVenda.lNumIntSolicSRV <> 0 Then
        
        objSolicSRV.lNumIntDoc = objPedidoVenda.lNumIntSolicSRV
        
        lErro = CF("SolicitacaoSRV_Le_NumIntDoc", objSolicSRV)
        If lErro <> SUCESSO And lErro <> 186988 Then gError 193478

        If lErro = SUCESSO Then
            CodSolicSRV.Text = objSolicSRV.lCodigo
        End If
        
    End If

    lErro = Carrega_Grid_ItensSRV(objPedidoVenda, bProcessarInfoPedido)
    If lErro <> SUCESSO Then gError 188893

    'Carrega o Grid de itens, indicando se é para processar reservas ou não
    lErro = Carrega_Grid_Itens(objPedidoVenda, bProcessarInfoPedido)
    If lErro <> SUCESSO Then gError 188885

    'Incluído por Luiz Nogueira em 20/05/03
    'Se não é necessário completar os dados do pedido, ou seja, não está usando um orçamento para
    'carregar a tela, significa que pode carregar todas as informações a partir de objPedidoVenda
    If Not bProcessarInfoPedido Then

        'Carrega o Tab Comissões
        lErro = objTabComissoes.Carrega_Tab_Comissoes(objPedidoVenda)
        If lErro <> SUCESSO Then gError 188886

'        'Carrega o Tab Bloqueio
'        lErro = Carrega_Tab_Bloqueio(objPedidoVenda)
'        If lErro <> SUCESSO Then gError 188887
        
        lErro = gobjTelaTrataBloqGen.Carrega_GridBloqueio(objPedidoVenda)
        If lErro <> SUCESSO Then gError 188887

'        Call Preenche_GridAlocacao(objPedidoVenda.colItensPedido)

    'Senão, ou seja, se está carregando um orçamento para a tela
    Else

        '*** Incluído por Luiz Nogueira em 26/01/04 ***
        'Transferir para versão 2
        'Se o cliente/filial do orçamento estão cadastrados
        If giClienteInexistente <> CLIENTE_INEXISTENTE And giFilialCliInexistente <> FILIALCLIENTE_INEXISTENTE Then
            'Exibe os dados default do cliente que são gravados no orçamento
            'Efetua o cálculo de comissões, se necessário
            lErro = Traz_Orcamento_Cliente(objCliente, objPedidoVenda)
            If lErro <> SUCESSO Then gError 188888
        End If
        '***********************************************

        'calcula as comissões para o pedido
        Call ComissaoAutomatica_Click

'        'calcula as embalagens do pedido
'        CalculoAuto_Click

    End If

    ValorTotal.Caption = Format(objPedidoVenda.dValorTotal, "Standard")

'    'Carrega o Tab de Tributação
'    lErro = Carrega_Tab_Tributacao(objPedidoVenda)
'    If lErro <> SUCESSO Then gError 188889

    '#############################################################
    'Inserido por Wagner
    objNFiscal.lCliente = objPedidoVenda.lCliente
    objNFiscal.iFilialCli = objPedidoVenda.iFilial

    '###################################################
    'Inserido por Wagner 04/08/2006
'    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objPedidoVenda.objProjetoInfo, objPedidoVenda.lNumIntDoc, PRJ_CR_TIPO_PV)
'    If lErro <> SUCESSO Then gError 188891
    '###################################################

    If objPedidoVenda.iStatus = STATUS_BAIXADO Then
        Status.Caption = STRING_STATUS_BAIXADO
        BotaoDesfazer.Enabled = True 'Inserido por Wagner
    Else
        Status.Caption = STRING_STATUS_ABERTO
        BotaoDesfazer.Enabled = False 'Inserido por Wagner
    End If
    
    Call SubTotal_Calcula

    Set gobjInfoAdic = objPedidoVenda.objInfoAdic
    
    ValorDescontoTit.Text = Format(objPedidoVenda.dValorDescontoTit, "Standard")
    Call ValorDescontoTit_Validate(bSGECancelDummy)

    iAlterado = 0

    iComissoesAlterada = 0

    gbCarregandoTela = False

    Traz_Pedido_Tela = SUCESSO

    Exit Function

Erro_Traz_Pedido_Tela:

    gbCarregandoTela = False

    Traz_Pedido_Tela = gErr

    Select Case gErr

        Case 188876 To 188891, 188893, 193478

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188892)

    End Select

    Exit Function

End Function

Private Function Carrega_Grid_ItensSRV(objPedidoVenda As ClassPedidoDeVenda, Optional bProcessaReserva As Boolean = False) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objProduto As New ClassProduto

On Error GoTo Erro_Carrega_Grid_ItensSRV

    Set gcolItemPedidoSRV = New Collection

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItensSRV)

    For iIndice = 1 To objPedidoVenda.colItensPedidoSRV.Count

        lErro = Mascara_RetornaProdutoEnxuto(objPedidoVenda.colItensPedidoSRV(iIndice).sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 188894

        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        objProduto.sCodigo = objPedidoVenda.colItensPedidoSRV(iIndice).sProduto
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 188895

        'Calcula o percentual de desconto
        If objPedidoVenda.colItensPedidoSRV(iIndice).dPrecoTotal > 0 Then
            dPercDesc = objPedidoVenda.colItensPedidoSRV(iIndice).dValorDesconto / (objPedidoVenda.colItensPedidoSRV(iIndice).dPrecoTotal + objPedidoVenda.colItensPedidoSRV(iIndice).dValorDesconto)
        End If

        objPedidoVenda.colItensPedidoSRV(iIndice).iItem = iIndice

        GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col) = Produto.Text
        GridItensSRV.TextMatrix(iIndice, iGrid_DescServicoSRV_Col) = objPedidoVenda.colItensPedidoSRV(iIndice).sProdutoDescricao
        GridItensSRV.TextMatrix(iIndice, iGrid_UnidadeMedSRV_Col) = objPedidoVenda.colItensPedidoSRV(iIndice).sUnidadeMed
        GridItensSRV.TextMatrix(iIndice, iGrid_QuantPedidaSRV_Col) = Formata_Estoque(objPedidoVenda.colItensPedidoSRV(iIndice).dQuantidade)
        GridItensSRV.TextMatrix(iIndice, iGrid_QuantCancelSRV_Col) = Formata_Estoque(objPedidoVenda.colItensPedidoSRV(iIndice).dQuantCancelada)
        GridItensSRV.TextMatrix(iIndice, iGrid_ValorUnitarioSRV_Col) = Format(objPedidoVenda.colItensPedidoSRV(iIndice).dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        GridItensSRV.TextMatrix(iIndice, iGrid_PercDescSRV_Col) = Format(dPercDesc, "Percent")
        GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col) = Format(objPedidoVenda.colItensPedidoSRV(iIndice).dValorDesconto, "Standard")
        GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col) = Format(objPedidoVenda.colItensPedidoSRV(iIndice).dPrecoTotal, "Standard")
        GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRVB_Col) = Format(objPedidoVenda.colItensPedidoSRV(iIndice).dPrecoTotal + objPedidoVenda.colItensPedidoSRV(iIndice).dValorDesconto, "Standard")
        If objPedidoVenda.colItensPedidoSRV(iIndice).dtDataEntrega <> DATA_NULA Then GridItensSRV.TextMatrix(iIndice, iGrid_DataEntregaSRV_Col) = Format(objPedidoVenda.colItensPedidoSRV(iIndice).dtDataEntrega, "dd/mm/yyyy")
        GridItensSRV.TextMatrix(iIndice, iGrid_QuantFaturadaSRV_Col) = Formata_Estoque(objPedidoVenda.colItensPedidoSRV(iIndice).dQuantFaturada)

        'Armazena os números internos dos itens
        gcolItemPedidoSRV.Add objPedidoVenda.colItensPedidoSRV(iIndice).lNumIntDoc

    Next

    'Atualiza o número de linhas existentes
    objGridItensSRV.iLinhasExistentes = gcolItemPedidoSRV.Count

    'Atualiza o grid para mostrar as checkboxes marcadas / desmarcadas
    Call Grid_Refresh_Checkbox(objGridItensSRV)

    Exit Function

Erro_Carrega_Grid_ItensSRV:

    Carrega_Grid_ItensSRV = gErr

    Select Case gErr

        Case 188894
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objPedidoVenda.colItensPedidoSRV(iIndice).sProduto)

        Case 188895

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188896)

    End Select

    Exit Function

End Function

Private Function Carrega_Grid_Itens(objPedidoVenda As ClassPedidoDeVenda, Optional bProcessaReserva As Boolean = False) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objProduto As New ClassProduto

On Error GoTo Erro_Carrega_Grid_Itens

    Set gcolItemPedido = New Collection

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    For iIndice = 1 To objPedidoVenda.colItensPedido.Count

        lErro = Mascara_RetornaProdutoEnxuto(objPedidoVenda.colItensPedido(iIndice).sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 191451

        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        objProduto.sCodigo = objPedidoVenda.colItensPedido(iIndice).sProduto
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 191452

        'Calcula o percentual de desconto
        If objPedidoVenda.colItensPedido(iIndice).dPrecoTotal > 0 Then
            dPercDesc = objPedidoVenda.colItensPedido(iIndice).dValorDesconto / (objPedidoVenda.colItensPedido(iIndice).dPrecoTotal + objPedidoVenda.colItensPedido(iIndice).dValorDesconto)
        End If

        objPedidoVenda.colItensPedido(iIndice).iItem = iIndice

        'Coloca os dados dos itens na tela

        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objPedidoVenda.colItensPedido(iIndice).sProdutoDescricao
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objPedidoVenda.colItensPedido(iIndice).sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col) = Formata_Estoque(objPedidoVenda.colItensPedido(iIndice).dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col) = Formata_Estoque(objPedidoVenda.colItensPedido(iIndice).dQuantCancelada)
        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dPercDesc, "Percent")
        GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dPrecoTotal, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dPrecoTotal + objPedidoVenda.colItensPedido(iIndice).dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col) = Formata_Estoque(objPedidoVenda.colItensPedido(iIndice).dQuantReservada)
        GridItens.TextMatrix(iIndice, iGrid_QuantFaturada_Col) = Formata_Estoque(objPedidoVenda.colItensPedido(iIndice).dQuantFaturada)

        'Armazena os números internos dos itens
        gcolItemPedido.Add objPedidoVenda.colItensPedido(iIndice).lNumIntDoc

        'Incluído por Luiz em 20/05/03 para processar as reservas quando o pedido for gerado por um orçamento
        'Se for para processar reservas
        If bProcessaReserva = True Then

            'Dispara a reserva do produto
            lErro = Reserva_Processa(iIndice, objPedidoVenda.colItensPedido(iIndice).dQuantidade, objPedidoVenda.colItensPedido(iIndice).dQuantCancelada, objPedidoVenda.colItensPedido(iIndice).dQuantFaturada)
            If lErro <> SUCESSO Then gError 191453


        End If

    Next

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = gcolItemPedido.Count

    'Atualiza o grid para mostrar as checkboxes marcadas / desmarcadas
    Call Grid_Refresh_Checkbox(objGridItens)

    Exit Function

Erro_Carrega_Grid_Itens:

    Carrega_Grid_Itens = gErr

    Select Case gErr

        Case 191451
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objPedidoVenda.colItensPedido(iIndice).sProduto)

        Case 191452, 191453

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191454)

    End Select

    Exit Function

End Function

Function Carrega_Tab_Complemento(objPedidoVenda As ClassPedidoDeVenda) As Long
'Coloca os dados do Tab de Complementos na tela

    PedidoCliente.Text = objPedidoVenda.sPedidoCliente
    PedidoRepr.Text = objPedidoVenda.sPedidoRepresentante
    Mensagem.Text = objPedidoVenda.sMensagemNota

    If objPedidoVenda.iCanalVenda > 0 Then
        CanalVenda.Text = objPedidoVenda.iCanalVenda
        Call CanalVenda_Validate(bSGECancelDummy)
    End If

    If objPedidoVenda.iFilialEntrega > 0 Then
        Call Filial_Formata(FilialEntrega, objPedidoVenda.iFilialEntrega)
        giFilialEntregaAlterada = 0
    End If

    If objPedidoVenda.iCodTransportadora > 0 Then
        Transportadora.Text = objPedidoVenda.iCodTransportadora
        Call Transportadora_Validate(bSGECancelDummy)
    End If

    If objPedidoVenda.iCodTranspRedesp > 0 Then
        TranspRedespacho.Text = objPedidoVenda.iCodTranspRedesp
        Call TranspRedespacho_Validate(bSGECancelDummy)
    End If

    RedespachoCli.Value = objPedidoVenda.iDetPagFrete

    If objPedidoVenda.lVolumeQuant > 0 Then VolumeQuant.Text = objPedidoVenda.lVolumeQuant

    Carrega_Tab_Complemento = SUCESSO

End Function

Function Carrega_Tab_Cobranca(objPedidoVenda As ClassPedidoDeVenda) As Long
'Coloca os dados do tab de cobrança na tela

Dim objParcelaPV As ClassParcelaPedidoVenda
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim dValorDesconto As Double

    Call DateParaMasked(DataReferencia, objPedidoVenda.dtDataReferencia)
    giDataReferenciaAlterada = 0

    PercAcrescFin.Text = ""

    If objPedidoVenda.iCondicaoPagto > 0 Then
        CondicaoPagamento.Text = objPedidoVenda.iCondicaoPagto
        Call CondicaoPagamento_Validate(bSGECancelDummy)
    Else
        CondicaoPagamento.Text = ""
    End If

    PercAcrescFin.Text = Format(objPedidoVenda.dPercAcrescFinanceiro * 100, "Fixed")

    CobrancaAutomatica.Value = objPedidoVenda.iCobrancaAutomatica

    'Limpa o Grid de Parcelas antes de preencher com dados da coleção
    Call Grid_Limpa(objGridParcelas)

    iIndice = 0

    For Each objParcelaPV In objPedidoVenda.colParcelas

        iIndice = iIndice + 1
        GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objParcelaPV.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objParcelaPV.dValor, "Standard")
        If objParcelaPV.dtDesconto1Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col) = Format(objParcelaPV.dtDesconto1Ate, "dd/mm/yyyy")
        If objParcelaPV.dtDesconto2Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col) = Format(objParcelaPV.dtDesconto2Ate, "dd/mm/yyyy")
        If objParcelaPV.dtDesconto3Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col) = Format(objParcelaPV.dtDesconto3Ate, "dd/mm/yyyy")
        If objParcelaPV.iDesconto1Codigo = VALOR_FIXO Or objParcelaPV.iDesconto1Codigo = VALOR_ANT_DIA Or objParcelaPV.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(objParcelaPV.dDesconto1Valor, "Standard")
        ElseIf objParcelaPV.iDesconto1Codigo = Percentual Or objParcelaPV.iDesconto1Codigo = PERC_ANT_DIA Or objParcelaPV.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col) = Format(objParcelaPV.dDesconto1Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaPV.dValor > 0 Then
                dValorDesconto = objParcelaPV.dDesconto1Valor * objParcelaPV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        If objParcelaPV.iDesconto2Codigo = VALOR_FIXO Or objParcelaPV.iDesconto2Codigo = VALOR_ANT_DIA Or objParcelaPV.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(objParcelaPV.dDesconto2Valor, "Standard")
        ElseIf objParcelaPV.iDesconto2Codigo = Percentual Or objParcelaPV.iDesconto2Codigo = PERC_ANT_DIA Or objParcelaPV.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col) = Format(objParcelaPV.dDesconto2Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaPV.dValor > 0 Then
                dValorDesconto = objParcelaPV.dDesconto2Valor * objParcelaPV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        If objParcelaPV.iDesconto3Codigo = VALOR_FIXO Or objParcelaPV.iDesconto3Codigo = VALOR_ANT_DIA Or objParcelaPV.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(objParcelaPV.dDesconto3Valor, "Standard")
        ElseIf objParcelaPV.iDesconto3Codigo = Percentual Or objParcelaPV.iDesconto3Codigo = PERC_ANT_DIA Or objParcelaPV.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col) = Format(objParcelaPV.dDesconto3Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaPV.dValor > 0 Then
                dValorDesconto = objParcelaPV.dDesconto3Valor * objParcelaPV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        For iIndice2 = 0 To TipoDesconto1.ListCount - 1
            If TipoDesconto1.ItemData(iIndice2) = objParcelaPV.iDesconto1Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col) = TipoDesconto1.List(iIndice2)
            If TipoDesconto2.ItemData(iIndice2) = objParcelaPV.iDesconto2Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col) = TipoDesconto2.List(iIndice2)
            If TipoDesconto3.ItemData(iIndice2) = objParcelaPV.iDesconto3Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col) = TipoDesconto3.List(iIndice2)
        Next

        For iIndice2 = 0 To TipoPagto.ListCount - 1
            If TipoPagto.ItemData(iIndice2) = objParcelaPV.iTipoPagto Then
                GridParcelas.TextMatrix(iIndice, iGrid_TipoPagto_Col) = TipoPagto.List(iIndice2)
                Exit For
            End If
        Next

    Next

    objGridParcelas.iLinhasExistentes = iIndice

    Carrega_Tab_Cobranca = SUCESSO

    Exit Function

End Function

'
'Function Carrega_Tab_Bloqueio(objPedidoVenda As ClassPedidoDeVenda) As Long
''Coloca os dados do Tab de Bloqueio na tela
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim iIndice2 As Integer
'Dim objBloqueioPV As ClassBloqueioPV
'Dim objTipoBloqueio As New ClassTipoDeBloqueio
'
'On Error GoTo Erro_Carrega_Tab_Bloqueio
'
'    'Limpa o Grid de Bloqueios antes de preencher com os dados da coleção
'    Call Grid_Limpa(objGridBloqueio)
'
'    iIndice = 0
'
'    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV
'
'        iIndice = iIndice + 1
'
'        objTipoBloqueio.iCodigo = objBloqueioPV.iTipoDeBloqueio
'
'        'Lê o Tipo de bloqueio
'        lErro = CF("TipoDeBloqueio_Le", objTipoBloqueio)
'        If lErro <> SUCESSO And lErro <> 23666 Then gError 191455
'
'        If lErro = 23666 Then gError 191456
'
'        'Coloca o bloqieo no Grid de bloqueios
'        GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col) = objTipoBloqueio.iCodigo & SEPARADOR & objTipoBloqueio.sNomeReduzido
'        GridBloqueio.TextMatrix(iIndice, iGrid_DataBloqueio_Col) = Format(objBloqueioPV.dtData, "dd/mm/yy")
'        GridBloqueio.TextMatrix(iIndice, iGrid_CodUsuario_Col) = objBloqueioPV.sCodUsuario
'        GridBloqueio.TextMatrix(iIndice, iGrid_ResponsavelBL_Col) = objBloqueioPV.sResponsavel
'        If objBloqueioPV.dtDataLib <> DATA_NULA Then GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col) = Format(objBloqueioPV.dtDataLib, "dd/mm/yy")
'        GridBloqueio.TextMatrix(iIndice, iGrid_ResponsavelLib_Col) = objBloqueioPV.sResponsavelLib
'        GridBloqueio.TextMatrix(iIndice, iGrid_Observacao_Col) = objBloqueioPV.sObservacao
'        GridBloqueio.TextMatrix(iIndice, iGrid_SeqBloqueio_Col) = CStr(objBloqueioPV.iSequencial)
'
'    Next
'    'Guarda o número de bloqueios trazidos do BD
'    giGridBloqueioLinhasBloqueadas = objPedidoVenda.colBloqueiosPV.Count
'
'    objGridBloqueio.iLinhasExistentes = iIndice
'
'    Carrega_Tab_Bloqueio = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Tab_Bloqueio:
'
'    Carrega_Tab_Bloqueio = gErr
'
'    Select Case gErr
'
'        Case 191455
'
'        Case 191456
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIO_NAO_CADASTRADO", gErr, objTipoBloqueio.iCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191457)
'
'    End Select
'
'End Function

Function Carrega_Grid_Reserva(objPedidoVenda As ClassPedidoDeVenda) As Long
'Carrega na tela as reservas dos ítens do pedido passado

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objItemPedido As ClassItemPedido
Dim objReserva As ClassReserva
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim dFator As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_Carrega_Grid_Reserva

    'Limpa o Grid de Reservas antes de prencher com os dados da coleção
    Call Grid_Limpa(objGridAlocacao)

    iIndice = 0
    iIndice1 = 0

    For Each objItemPedido In objPedidoVenda.colItensPedido

        iIndice1 = iIndice1 + 1

        objProduto.sCodigo = objItemPedido.sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 191458
        
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemPedido.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 191459

        For Each objReserva In objItemPedido.colReserva

            iIndice = iIndice + 1

            Call Remove_Reservas(iIndice)

            'Coloca os dados das reservas na tela
            GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col) = iIndice1
            GridAlocacao.TextMatrix(iIndice, iGrid_ProdutoAlmox_Col) = GridItens.TextMatrix(iIndice1, iGrid_Produto_Col)
            GridAlocacao.TextMatrix(iIndice, iGrid_QuantReservar_Col) = Formata_Estoque(objItemPedido.dQuantidade * dFator)
            GridAlocacao.TextMatrix(iIndice, iGrid_QuantReserv_Col) = Formata_Estoque(objReserva.dQuantidade)
            If objReserva.dtDataValidade <> DATA_NULA Then GridAlocacao.TextMatrix(iIndice, iGrid_Validade_Col) = Format(objReserva.dtDataValidade, "dd/mm/yyyy")
            GridAlocacao.TextMatrix(iIndice, iGrid_Responsavel_Col) = objReserva.sResponsavel
            GridAlocacao.TextMatrix(iIndice, iGrid_UMEstoque_Col) = objItemPedido.sUMEstoque

            objAlmoxarifado.iCodigo = objReserva.iAlmoxarifado

            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 191460
            
            If lErro = 25056 Then gError 191461

            GridAlocacao.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido

        Next

    Next

    objGridAlocacao.iLinhasExistentes = iIndice

    Carrega_Grid_Reserva = SUCESSO

    Exit Function

Erro_Carrega_Grid_Reserva:

    Carrega_Grid_Reserva = gErr

    Select Case gErr

        Case 191458 To 191460

        Case 191461
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191462)

    End Select

    Exit Function

End Function

Private Sub ClienteLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)


End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer, Optional bPergunta As Boolean = True)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer
Dim iIndice As Integer
Dim dValorTotal As Double
Dim lErro As Long
Dim objItemOP As New ClassItemOP
Dim objItemPV As New ClassItemPedido
Dim vbMsgRes As VbMsgBoxResult 'Inserido por Wagner

On Error GoTo Erro_GridItens_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iItemAtual = GridItens.Row

    If KeyCode = vbKeyDelete And iItemAtual <= objGridItens.iLinhasExistentes Then

        'Não deixa deletar se a quantidade faturada for positiva
        If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantFaturada_Col)) > 0 Then gError 188911

        'Se ítem já existe no BD verifica se está ligado a Ordem Producao
        If gcolItemPedido(iItemAtual) <> 0 Then

            objItemPV.lNumIntDoc = gcolItemPedido(iItemAtual)

            lErro = CF("ItemOP_Le_ItemPV", objItemOP, objItemPV)
            If lErro <> SUCESSO And lErro <> 46074 Then gError 188912
            
            If lErro = SUCESSO Then gError 188913

        End If

    End If

    '##################################
    'Alterado por Wagner 09/01/2006
    If bPergunta Then
        Call Grid_Trata_Tecla1(KeyCode, objGridItens)
    Else
        Call Grid_Trata_Tecla3(KeyCode, objGridItens)
    End If
    '##################################

    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then

        '************ grade ************
        GridItens.TextMatrix(GridItens.Row, 0) = GridItens.Row
''''        gobjPedidoVenda.colItensPedido.Remove GridItens.Row

        For iLinhasExistentesAnterior = 1 To objGridItens.iLinhasExistentes
            If gobjPedidoVenda.colItensPedido(iLinhasExistentesAnterior).iPossuiGrade = MARCADO Then
                GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = "# " & iLinhasExistentesAnterior
            Else
                GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = iLinhasExistentesAnterior
            End If
        Next

        GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = iLinhasExistentesAnterior
        '*******************************

        'Indica que houve alteração nos dados base para o cálculo de comissões
        iComissoesAlterada = REGISTRO_ALTERADO

        gcolItemPedido.Remove (iItemAtual)

        Call gobjTribTab.Exclusao_Item_Grid(iItemAtual)
        'Call Tributacao_Remover_Item_Grid(iItemAtual)

        Call gobjTelaProjetoInfo.Remove_Linha(iItemAtual)

        If objGridAlocacao.iLinhasExistentes > 0 Then
            Call Remove_Reservas(iItemAtual)
            For iIndice = 1 To objGridAlocacao.iLinhasExistentes
                If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))) > 0 Then If CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) > iItemAtual Then GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col) = CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) - 1
            Next
        End If

'        'Calcula a soma dos valores de produtos
'        For iIndice = 1 To objGridItens.iLinhasExistentes
'            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
'                If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'            End If
'        Next
'
'        For iIndice = 1 To objGridItensSRV.iLinhasExistentes
'            If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
'                If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
'            End If
'        Next
'
'        'Coloca valor total dos produtos na tela
'        ValorProdutos.Caption = Format(dValorTotal, "Standard")
'
'        '#############################################
'        'Alterado por Wagner
'        If objGridItens.iLinhasExistentes <> 0 Then
'            Call CF2(Me, "PSRV_PrecoTotal_Calcula", objGridItens.iLinhasExistentes)
'        Else
'
'            If StrParaDbl(ValorDesconto.Text) <> 0 Then
'                'Avisa ao usuário
'                vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", ValorDesconto.Text, 0)
'
'                'Limpa o valor de desconto
'                gdDesconto = 0
'                ValorDesconto.Text = ""
'                giValorDescontoAlterado = 0
'
'                Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, 0)
'
'                'Para tributação
'                gobjPedidoVenda.dValorDesconto = 0
'
'            End If
'        End If
'        '#############################################

        Call SubTotal_Calcula

        'Calcula o valor total da nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 188915

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 188916

    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr

        Case 188911
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_PV_QUANT_FAT_POSITIVA", gErr, iItemAtual)

        Case 188912, 188914 To 188917

        Case 188913
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_PV_VINCULADO_ITEM_OP", gErr, iItemAtual, objItemOP.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188918)

    End Select

    Exit Sub

End Sub

Public Sub GridItensSRV_KeyDown(KeyCode As Integer, Shift As Integer, Optional bPergunta As Boolean = True)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer
Dim iIndice As Integer
Dim dValorTotal As Double
Dim lErro As Long
'Dim objItemOS As New ClassItemOS
Dim objItemPV As New ClassItemPedido
Dim vbMsgRes As VbMsgBoxResult 'Inserido por Wagner

On Error GoTo Erro_GridItensSRV_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGridItensSRV.iLinhasExistentes
    iItemAtual = GridItensSRV.Row

    If KeyCode = vbKeyDelete And iItemAtual <= objGridItensSRV.iLinhasExistentes Then

        'Não deixa deletar se a quantidade faturada for positiva
        If StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantFaturadaSRV_Col)) > 0 Then gError 188919

        'Se ítem já existe no BD verifica se está ligado a Ordem Producao
        If gcolItemPedidoSRV(iItemAtual) <> 0 Then

            objItemPV.lNumIntDoc = gcolItemPedidoSRV(iItemAtual)

'            lErro = CF("ItemOS_Le_ItemPV", objItemOS, objItemPV)
'            If lErro <> SUCESSO And lErro <> 46074 Then gError 188920
            
'            If lErro = SUCESSO Then gError 188921

        End If

    End If

    '##################################
    'Alterado por Wagner 09/01/2006
    If bPergunta Then
        Call Grid_Trata_Tecla1(KeyCode, objGridItensSRV)
    Else
        Call Grid_Trata_Tecla3(KeyCode, objGridItensSRV)
    End If
    '##################################

    If objGridItensSRV.iLinhasExistentes < iLinhasExistentesAnterior Then

        '************ grade ************
        GridItensSRV.TextMatrix(GridItensSRV.Row, 0) = GridItensSRV.Row
''''        gobjPedidoVenda.colItensPedido.Remove GridItens.Row

        For iLinhasExistentesAnterior = 1 To objGridItensSRV.iLinhasExistentes
            If gobjPedidoVenda.colItensPedidoSRV(iLinhasExistentesAnterior).iPossuiGrade = MARCADO Then
                GridItensSRV.TextMatrix(iLinhasExistentesAnterior, 0) = "# " & iLinhasExistentesAnterior
            Else
                GridItensSRV.TextMatrix(iLinhasExistentesAnterior, 0) = iLinhasExistentesAnterior
            End If
        Next

        GridItensSRV.TextMatrix(iLinhasExistentesAnterior, 0) = iLinhasExistentesAnterior
        '*******************************

        'Indica que houve alteração nos dados base para o cálculo de comissões
        iComissoesAlterada = REGISTRO_ALTERADO

        gcolItemPedidoSRV.Remove (iItemAtual)

        'Call Tributacao_Remover_Item_GridSRV(iItemAtual)
        Call gobjTribTab.Exclusao_Item_Grid(iItemAtual, True)

        'Call gobjTelaProjetoInfo.Remove_Linha(iItemAtual)

'        If objGridAlocacao.iLinhasExistentes > 0 Then
'            Call Remove_Reservas(iItemAtual)
'            For iIndice = 1 To objGridAlocacao.iLinhasExistentes
'                If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))) > 0 Then If CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) > iItemAtual Then GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col) = CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) - 1
'            Next
'        End If

        'Calcula a soma dos valores de produtos
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
                If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
            End If
        Next

        For iIndice = 1 To objGridItensSRV.iLinhasExistentes
            If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
                If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
            End If
        Next

        'Coloca valor total dos produtos na tela
        ValorProdutos.Caption = Format(dValorTotal, "Standard")

        '#############################################
        'Alterado por Wagner
        If objGridItensSRV.iLinhasExistentes <> 0 Then
            Call CF2(Me, "PSRV_PrecoTotal_CalculaSRV", objGridItensSRV.iLinhasExistentes)
        Else

            If StrParaDbl(ValorDesconto.Text) <> 0 Then
                'Avisa ao usuário
                vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", ValorDesconto.Text, 0)

                'Limpa o valor de desconto
                gdDesconto = 0
                ValorDesconto.Text = ""
                giValorDescontoAlterado = 0

                'Para tributação
                gobjPedidoVenda.dValorDesconto = 0
            End If
        End If
        '#############################################

        'Calcula o valor total da nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 188923

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 188924

    End If

    Exit Sub

Erro_GridItensSRV_KeyDown:

    Select Case gErr

        Case 188919
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_PV_QUANT_FAT_POSITIVA", gErr, iItemAtual)

        Case 188920, 188922 To 188925

        Case 188921
'            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_PV_VINCULADO_ITEM_OS", gErr, iItemAtual, objItemOS.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188926)

    End Select

    Exit Sub

End Sub

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Public Sub GridItensSRV_Scroll()

    Call Grid_Scroll(objGridItensSRV)

End Sub

'Public Sub ICMSAliquotaItem_Change()
'
'    giICMSAliquotaItemAlterado = 1
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ICMSAliquotaItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ICMSAliquotaItem_Validate
'
'    If giICMSAliquotaItemAlterado Then
'
'        If Len(Trim(ICMSAliquotaItem.ClipText)) > 0 Then
'            lErro = Porcentagem_Critica(ICMSAliquotaItem.Text)
'            If lErro <> SUCESSO Then gError 188927
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giICMSAliquotaItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ICMSAliquotaItem_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 188927
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188928)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ICMSBase_Change()
'
'    ICMSBase1.Caption = ICMSBase.Caption
'
'End Sub
'
'Public Sub ICMSBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSBaseItemAlterado = 1
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ICMSBaseItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ICMSBaseItem_Validate
'
'    If giICMSBaseItemAlterado Then
'
'        If Len(Trim(ICMSBaseItem.ClipText)) > 0 Then
'
'            lErro = Valor_NaoNegativo_Critica(ICMSBaseItem.Text)
'            If lErro <> SUCESSO Then gError 188929
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giICMSBaseItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ICMSBaseItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 188929
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188930)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ICMSPercRedBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSPercRedBaseItemAlterado = 1
'
'End Sub
'
'Public Sub ICMSPercRedBaseItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ICMSPercRedBaseItem_Validate
'
'    If giICMSPercRedBaseItemAlterado Then
'
'        If Len(Trim(ICMSPercRedBaseItem.Text)) > 0 Then
'
'            lErro = Porcentagem_Critica(ICMSPercRedBaseItem.Text)
'            If lErro <> SUCESSO Then gError 188931
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giICMSPercRedBaseItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ICMSPercRedBaseItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 188931
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188932)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ICMSSubstAliquotaItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSSubstAliquotaItemAlterado = 1
'
'End Sub
'
'Public Sub ICMSSubstAliquotaItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ICMSSubstAliquotaItem_Validate
'
'    If giICMSSubstAliquotaItemAlterado Then
'
'        If Len(Trim(ICMSSubstAliquotaItem.ClipText)) > 0 Then
'
'            lErro = Porcentagem_Critica(ICMSSubstAliquotaItem.Text)
'            If lErro <> SUCESSO Then gError 188933
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giICMSSubstAliquotaItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ICMSSubstAliquotaItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 188933
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188934)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ICMSSubstBase_Change()
'
'    ICMSSubstBase1.Caption = ICMSSubstBase.Caption
'
'End Sub
'
'Public Sub ICMSSubstBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSSubstBaseItemAlterado = 1
'
'End Sub
'
'Public Sub ICMSSubstBaseItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ICMSSubstBaseItem_Validate
'
'    If giICMSSubstBaseItemAlterado Then
'
'        If Len(Trim(ICMSSubstBaseItem.ClipText)) > 0 Then
'
'            lErro = Valor_NaoNegativo_Critica(ICMSSubstBaseItem.Text)
'            If lErro <> SUCESSO Then gError 188935
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giICMSSubstBaseItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ICMSSubstBaseItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 188935
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188936)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ICMSSubstValor_Change()
'
'    ICMSSubstValor1.Caption = ICMSSubstValor.Caption
'
'End Sub
'
'Public Sub ICMSSubstValorItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSSubstValorItemAlterado = 1
'
'End Sub
'
'Public Sub ICMSSubstValorItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ICMSSubstValorItem_Validate
'
'    If giICMSSubstValorItemAlterado Then
'
'        If Len(Trim(ICMSSubstValorItem.ClipText)) > 0 Then
'
'            lErro = Valor_NaoNegativo_Critica(ICMSSubstValorItem.Text)
'            If lErro <> SUCESSO Then gError 188937
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giICMSSubstValorItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ICMSSubstValorItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 188937
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188938)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ICMSValor_Change()
'
'    ICMSValor1.Caption = ICMSValor.Caption
'
'End Sub
'
'Public Sub ICMSValorItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giICMSValorItemAlterado = 1
'
'End Sub
'
'Public Sub ICMSValorItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ICMSValorItem_Validate
'
'    If giICMSValorItemAlterado Then
'
'        If Len(Trim(ICMSValorItem.ClipText)) > 0 Then
'
'            lErro = Valor_NaoNegativo_Critica(ICMSValorItem.Text)
'            If lErro <> SUCESSO Then gError 188939
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giICMSValorItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_ICMSValorItem_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 188939
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188940)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub IPIAliquotaItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giIPIAliquotaItemAlterado = 1
'
'End Sub
'
'Public Sub IPIAliquotaItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_IPIAliquotaItem_Validate
'
'    If giIPIAliquotaItemAlterado Then
'
'        If Len(Trim(IPIAliquotaItem.ClipText)) > 0 Then
'
'            lErro = Porcentagem_Critica(IPIAliquotaItem.Text)
'            If lErro <> SUCESSO Then gError 188941
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giIPIBaseItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_IPIAliquotaItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 188941
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188942)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub IPIBase_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub IPIBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giIPIBaseItemAlterado = 1
'
'End Sub
'
'Public Sub IPIBaseItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_IPIBaseItem_Validate
'
'    If giIPIBaseItemAlterado Then
'
'        If Len(Trim(IPIBaseItem.ClipText)) > 0 Then
'
'            lErro = Valor_NaoNegativo_Critica(IPIBaseItem.Text)
'            If lErro <> SUCESSO Then gError 188943
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giIPIBaseItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_IPIBaseItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 188943
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188944)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub IPIPercRedBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giIPIPercRedBaseItemAlterado = 1
'
'End Sub
'
'Public Sub IPIPercRedBaseItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_IPIPercRedBaseItem_Validate
'
'    If giIPIPercRedBaseItemAlterado Then
'
'        If Len(Trim(IPIPercRedBaseItem.Text)) > 0 Then
'            lErro = Porcentagem_Critica(IPIPercRedBaseItem.Text)
'            If lErro <> SUCESSO Then gError 188945
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giIPIPercRedBaseItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_IPIPercRedBaseItem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 188945
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188946)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub IPIValor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub IPIValorItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giIPIValorItemAlterado = 1
'
'End Sub
'
'Public Sub IPIValorItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_IPIValorItem_Validate
'
'    If giIPIValorItemAlterado Then
'
'        If Len(Trim(IPIValorItem.ClipText)) > 0 Then
'
'            lErro = Valor_NaoNegativo_Critica(IPIValorItem.Text)
'            If lErro <> SUCESSO Then gError 188947
'
'        End If
'
'        Call BotaoGravarTribItem_Click
'
'        giIPIValorItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_IPIValorItem_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 188947
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188948)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub IRAliquota_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giAliqIRAlterada = 1
'
'End Sub

''Private Function AtualizarTribInfo(sIdentificador As String, vValor As Variant) As Long
''
''Dim lErro As Long
''Dim objTributoDoc As ClassTributoDoc
''
''On Error GoTo Erro_AtualizarTribInfo
''
''    Set objTributoDoc = gobjPedidoVenda
''    lErro = objTributoDoc.GuardarVarTela(sIdentificador, vValor)
''    If lErro <> SUCESSO Then gError 27673
''
''    AtualizarTribInfo = SUCESSO
''
''    Exit Function
''
''Erro_AtualizarTribInfo:
''
''    AtualizarTribInfo = gErr
''
''    Select Case gErr
''
''        Case 27673
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157939)
''
''    End Select
''
''    Exit Function
''
''End Function
'
'Public Sub IRAliquota_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dIRAliquota As Double, dIRValor As Double
'
'On Error GoTo Erro_IRAliquota_Validate
'
'    If giAliqIRAlterada = 0 Then Exit Sub
'
'    If Len(Trim(IRAliquota.ClipText)) > 0 Then
'        lErro = Porcentagem_Critica(IRAliquota.Text)
'        If lErro <> SUCESSO Then gError 188949
'    End If
'
'    Call BotaoGravarTrib
'
'    giAliqIRAlterada = 0
'
'    Exit Sub
'
'Erro_IRAliquota_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 188949
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188950)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub IRBase_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ISSAliquota_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giISSAliquotaAlterada = 1
'
'End Sub
'
'Public Sub ISSAliquota_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ISSAliquota_Validate
'
'    If giISSAliquotaAlterada = 0 Then Exit Sub
'
'    If Len(Trim(ISSAliquota.ClipText)) > 0 Then
'        lErro = Porcentagem_Critica(ISSAliquota.Text)
'        If lErro <> SUCESSO Then gError 188951
'    End If
'
'    Call BotaoGravarTrib
'
'    giISSAliquotaAlterada = 0
'
'    Exit Sub
'
'Erro_ISSAliquota_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 188951
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188952)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ISSBase_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ISSIncluso_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_ISSIncluso_Click
'
'    iAlterado = REGISTRO_ALTERADO
'
'    Call BotaoGravarTrib
'
'    Exit Sub
'
'Erro_ISSIncluso_Click:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188953)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ISSValor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giISSValorAlterado = 1
'
'End Sub
'
'Public Sub ISSValor_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ISSValor_Validate
'
'    If giISSValorAlterado = 0 Then Exit Sub
'
'    If Len(Trim(ISSValor.ClipText)) > 0 Then
'        lErro = Valor_NaoNegativo_Critica(ISSValor.Text)
'        If lErro <> SUCESSO Then gError 188954
'    End If
'
'    Call BotaoGravarTrib
'
'    giISSValorAlterado = 0
'
'    Exit Sub
'
'Erro_ISSValor_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 188954
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188955)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

'Public Sub LblTipoTrib_Click()
'
'Dim colSelecao As New Collection
'Dim objTipoTrib As New ClassTipoDeTributacaoMovto
'
'    'apenas tipos de saida
'    colSelecao.Add "0"
'    colSelecao.Add "0"
'
'    Call Chama_Tela("TiposDeTribMovtoLista", colSelecao, objTipoTrib, objEventoTiposDeTributacao)
'
'End Sub
'
'Public Sub LblTipoTribItem_Click()
'
'    Call LblTipoTrib_Click
'
'End Sub

Public Sub MensagemLabel_Click()

Dim objMensagem As New ClassMensagem
Dim colSelecao As New Collection

    'Coloca a mensagem da tela no objMensagem
    objMensagem.sDescricao = Mensagem.Text

    'Chama a tela de browse de Mensagem
    Call Chama_Tela("MensagemLista", colSelecao, objMensagem, objEventoMensagem)

    Exit Sub

End Sub

'Public Sub NaturezaItemLabel_Click()
'
'Dim objNaturezaOp As New ClassNaturezaOp
'Dim colSelecao As New Collection
'Dim dtDataRef As Date, sSelecao As String
'
'    If Len(Trim(NaturezaOpItem.Text)) > 0 Then objNaturezaOp.sCodigo = NaturezaOpItem.Text
'
'    If Len(Trim(DataEmissao.ClipText)) > 0 Then
'        dtDataRef = MaskedParaDate(DataEmissao)
'    Else
'        dtDataRef = DATA_NULA
'    End If
'
'    sSelecao = "Codigo >= " & NATUREZA_SAIDA_COD_INICIAL & " AND Codigo <= " & NATUREZA_SAIDA_COD_FINAL & " AND {fn LENGTH(Codigo) } = " & IIf(dtDataRef < DATA_INICIO_CFOP4, "3", "4")
'
'    Call Chama_Tela("NaturezaOperacaoLista", colSelecao, objNaturezaOp, objEventoNaturezaOp, sSelecao)
'
'End Sub
'
'Public Sub NaturezaLabel_Click()
'
'Dim objNaturezaOp As New ClassNaturezaOp
'Dim colSelecao As New Collection
'Dim dtDataRef As Date
'
'    'Se NaturezaOP estiver preenchida coloca no Obj
'    objNaturezaOp.sCodigo = NaturezaOp.Text
'
'    If Len(Trim(DataEmissao.ClipText)) > 0 Then
'        dtDataRef = MaskedParaDate(DataEmissao)
'    Else
'        dtDataRef = DATA_NULA
'    End If
'
'    'selecao p/obter apenas as nat de saida
'    colSelecao.Add NATUREZA_SAIDA_COD_INICIAL
'    colSelecao.Add NATUREZA_SAIDA_COD_FINAL
'
'    'Chama a Tela de browse de NaturezaOp
'    Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, objEventoNaturezaOp, "{fn LENGTH(Codigo) } = " & IIf(dtDataRef < DATA_INICIO_CFOP4, "3", "4"))
'
'    Exit Sub
'
'End Sub
'
'Public Sub NaturezaOp_GotFocus()
'
'Dim iNaturezaAux As Integer
'
'    iNaturezaAux = giNaturezaOpAlterada
'    Call MaskEdBox_TrataGotFocus(NaturezaOp, iAlterado)
'    giNaturezaOpAlterada = iNaturezaAux
'
'End Sub
'
'Public Sub NaturezaOpItem_Change()
'
'    giNatOpItemAlterado = 1
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub NaturezaOpItem_GotFocus()
'Dim iNaturezaOpAux As Integer
'
'    iNaturezaOpAux = giNatOpItemAlterado
'    Call MaskEdBox_TrataGotFocus(NaturezaOpItem, iAlterado)
'    giNatOpItemAlterado = iNaturezaOpAux
'
'End Sub
'
'Public Sub NaturezaOpItem_Validate(Cancel As Boolean)
'
'Dim sNatOp As String, lErro As Long, objNaturezaOp As New ClassNaturezaOp
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_NaturezaOpItem_Validate
'
'    If giNatOpItemAlterado = 0 Then Exit Sub
'
'    sNatOp = Trim(NaturezaOpItem.Text)
'
'    If sNatOp <> "" Then
'
'        objNaturezaOp.sCodigo = sNatOp
'
'        If objNaturezaOp.sCodigo < NATUREZA_SAIDA_COD_INICIAL Or objNaturezaOp.sCodigo > NATUREZA_SAIDA_COD_FINAL Then gError 188956
'
'        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
'        If lErro <> SUCESSO And lErro <> 17958 Then gError 188957
'
'        'Se não achou a Natureza de Operação --> erro
'        If lErro <> SUCESSO Then gError 188958
'
'        LabelDescrNatOpItem.Caption = objNaturezaOp.sDescricao
'        Call BotaoGravarTribItem_Click
'    Else
'        'Limpa a descrição
'        LabelDescrNatOpItem.Caption = ""
'    End If
'
'    giNatOpItemAlterado = 0
'
'    Exit Sub
'
'Erro_NaturezaOpItem_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 188956
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", gErr)
'
'        Case 188957
'
'        Case 188958
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_NATUREZA_OPERACAO", NaturezaOpItem.Text)
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("NaturezaOperacao", objNaturezaOp)
'            Else
'            End If
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188959)
'
'    End Select
'
'End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objCliente.sNomeReduzido

    'Dispara o Validate de Cliente
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub NumeroLabel_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colSelecao As Collection
Dim sSelecaoSQL As String

On Error GoTo Erro_NumeroLabel_Click

    Set colSelecao = New Collection

    colSelecao.Add STATUS_BAIXADO

    sSelecaoSQL = "Status<>?"

    lErro = Move_Pedido_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then gError 188960

    Call Chama_Tela("PedSRVAbertosLista", colSelecao, objPedidoVenda, objEventoNumero, sSelecaoSQL)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 188960

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188961)

    End Select

    Exit Sub

End Sub

Private Sub objEventoMensagem_evSelecao(obj1 As Object)

Dim objMensagem As ClassMensagem

    Set objMensagem = obj1

    Mensagem.Text = objMensagem.sDescricao

    Me.Show

End Sub

'Private Sub objEventoNaturezaOp_evSelecao(obj1 As Object)
'
'Dim objNaturezaOp As New ClassNaturezaOp
'
'    Set objNaturezaOp = obj1
'
'    If giFrameAtual = 1 Then
'
'        'Preenche a natureza de Opereração do frame principal
'        NaturezaOp.Text = objNaturezaOp.sCodigo
'        Call NaturezaOp_Validate(bSGECancelDummy)
'
'    Else
'        'Preenche a NatOp do frame de tributação
'        NaturezaOpItem.Text = objNaturezaOp.sCodigo
'        Call NaturezaOpItem_Validate(bSGECancelDummy)
'
'    End If
'
'    Me.Show
'
'End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim objPedidoVenda As ClassPedidoDeVenda
Dim lErro As Long

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objPedidoVenda = obj1

    lErro = Traz_Pedido_Tela(objPedidoVenda)
    If lErro <> SUCESSO Then gError 188962

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 188962

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188963)

    End Select

    Exit Sub

End Sub

Public Sub CondPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection

    'Se Condição de Pagto estiver preenchida, extrai o código
    If Len(Trim(CondicaoPagamento.Text)) > 0 Then
        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)
    End If

    'Chama a Tela CondicoesPagamentoCRLista
    Call Chama_Tela("CondicaoPagtoCRLista", colSelecao, objCondicaoPagto, objEventoCondPagto)

End Sub

Private Sub objEventoCondPagto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCondicaoPagto As ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult
Dim dPercAcresFin As Double

On Error GoTo Erro_objEventoCondPagto_evSelecao

    Set objCondicaoPagto = obj1

    'Preenche campo CondicaoPagamento
    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida

    'Altera PercAcrescFin
    If Len(Trim(PercAcrescFin.ClipText)) > 0 Then

        dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100

        If dPercAcresFin <> objCondicaoPagto.dAcrescimoFinanceiro Then
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
            If vbMsgRes = vbYes Then
                PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
                Call PercAcrescFin_Validate(bSGECancelDummy)
            End If
        End If
    Else
        PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
        Call PercAcrescFin_Validate(bSGECancelDummy)
    End If

    If Len(Trim(ValorTotal.Caption)) > 0 Then
        'Se DataReferencia estiver preenchida e Valor for positivo
        If Len(Trim(DataReferencia.ClipText)) > 0 And CDbl(ValorTotal.Caption) > 0 Then

            'Preenche GridParcelas a partir da Condição de Pagto
            lErro = Cobranca_Automatica()
            If lErro <> SUCESSO Then gError 188964

        End If
    End If

    Me.Show

    Exit Sub

Erro_objEventoCondPagto_evSelecao:

    Select Case gErr

        Case 188964

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188965)

     End Select

     Exit Sub

End Sub

Public Sub BotaoPecas_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String
Dim sDescricao As String
Dim sOrdenacao As String
Dim sSelecaoSQL As String

On Error GoTo Erro_BotaoPecas_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridItens.Row = 0 Then gError 188966

    sOrdenacao = "Código"

    If Me.ActiveControl Is Produto Then
        sProduto1 = Produto.Text
    Else
        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
    End If

    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 188967

    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    If Me.ActiveControl Is DescricaoProduto Then
        sOrdenacao = "Descrição"
        sDescricao = DescricaoProduto.Text
        If iPreenchido = PRODUTO_PREENCHIDO Then gError 188968
    Else
        sDescricao = GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col)
    End If

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto
    objProduto.sDescricao = sDescricao

    Set colSelecao = New Collection

    colSelecao.Add NATUREZA_PROD_SERVICO

    sSelecaoSQL = "Natureza<>?"

    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto, sSelecaoSQL, sOrdenacao)

    Exit Sub

Erro_BotaoPecas_Click:

    Select Case gErr

        Case 188966
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 188967, 188968

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188969)

    End Select

    Exit Sub

End Sub

Public Sub BotaoServicos_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String
Dim sSelecaoSQL As String
Dim sDescricao As String
Dim sOrdenacao As String

On Error GoTo Erro_BotaoServicos_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridItensSRV.Row = 0 Then gError 188970

    sOrdenacao = "Código"

    If Me.ActiveControl Is ServicoSRV Then

        sProduto1 = ServicoSRV.Text

    Else

        sProduto1 = GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ServicoSRV_Col)

    End If

    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 188971

    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    If Me.ActiveControl Is DescricaoProduto Then
        sOrdenacao = "Descrição"
        sDescricao = DescricaoProduto.Text
        If iPreenchido = PRODUTO_PREENCHIDO Then gError 188972
    Else
        sDescricao = GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col)
    End If

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto
    objProduto.sDescricao = sDescricao

    Set colSelecao = New Collection

    colSelecao.Add NATUREZA_PROD_SERVICO

    sSelecaoSQL = "Natureza=?"

    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoServico, sSelecaoSQL, sOrdenacao)

    Exit Sub

Erro_BotaoServicos_Click:

    Select Case gErr

        Case 188970
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 188971, 188972

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188973)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOrcamento_evSelecao(obj1 As Object)

Dim objOrcamentoVenda As ClassOrcamentoVenda

On Error GoTo Erro_objEventoOrcamento_evSelecao

    Set objOrcamentoVenda = obj1

    Orcamento.Text = objOrcamentoVenda.lCodigo
    FilialOrcamento.Text = objOrcamentoVenda.iFilialEmpresa
    Call FilialOrcamento_Validate(bSGECancelDummy)

    Me.Show

    Call BotaoOrcamento_Click

    Exit Sub

Erro_objEventoOrcamento_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridItens.Row < 1 Then Exit Sub
    
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then
    
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 188974
    
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
    
        If Not (Me.ActiveControl Is Produto) Then
    
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
    
            'Faz o Tratamento do produto
            lErro = CF2(Me, "ProdutoSRV_Saida_Celula")
            If lErro <> SUCESSO Then gError 188975
    
        End If
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 188974
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 188975
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188976)

    End Select

    Exit Sub

End Sub

Private Sub objEventoServico_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoServico_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridItensSRV.Row < 1 Then Exit Sub

        If Len(Trim(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ServicoSRV_Col))) = 0 Then
    
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 188977
    
        ServicoSRV.PromptInclude = False
        ServicoSRV.Text = sProduto
        ServicoSRV.PromptInclude = True
    
        If Not (Me.ActiveControl Is ServicoSRV) Then
    
            GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ServicoSRV_Col) = ServicoSRV.Text
    
            'Faz o Tratamento do produto
            lErro = ServicoSRV_Saida_Celula()
            If lErro <> SUCESSO Then gError 188978
    
        End If
    
        'Call ComandoSeta_Fechar(Me.Name)
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr

        Case 188977
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 188978

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188979)

    End Select

    Exit Sub

End Sub

Public Sub BotaoEstoquePeca_Click()

Dim lErro As Long
Dim objEstoqueProd As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_BotaoEstoquePeca_Click

    lErro = CF2(Me, "PSRV_Botao_Estoque", objEstoqueProd, colSelecao, sSelecao)
    If lErro <> SUCESSO Then gError 188980

    Call Chama_Tela("EstoqueProdutoLista", colSelecao, objEstoqueProd, objEventoEstoqueProd, sSelecao)

    Exit Sub

Erro_BotaoEstoquePeca_Click:

    Select Case gErr

        Case 188980

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188981)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim objPedidoVendaAux As New ClassPedidoDeVenda

On Error GoTo Erro_Trata_Parametros

    If Not (objPedidoVenda Is Nothing) Then

        lErro = CargaPosFormLoad(True)
        If lErro <> SUCESSO Then gError 188982

        'Se foi passado o código do orçamento
        If objPedidoVenda.lCodigo > 0 Then  'Incluído por Luiz Nogueira em 13/01/04

            objPedidoVendaAux.lCodigo = objPedidoVenda.lCodigo
            objPedidoVendaAux.iFilialEmpresa = objPedidoVenda.iFilialEmpresa

            'Coloca o Pedido de Venda na tela
            lErro = Traz_Pedido_Tela(objPedidoVendaAux)
            If lErro <> SUCESSO And lErro <> 51162 Then gError 188983
            
            If lErro <> SUCESSO Then  'Não encontrou no BD o código de Pedido

                'Limpa a tela e coloca o código na Tela
                Call Limpa_Tela_PedidoVenda
                Codigo.Text = CStr(objPedidoVenda.lCodigo)

            End If

        '*** Incluído por Luiz Nogueira em 13/01/04 - INÍCIO ***
        'Se foi passado o código do cliente
        ElseIf objPedidoVenda.lCliente > 0 Then

            'Joga o código do cliente na tela
            Cliente.Text = objPedidoVenda.lCliente
            Call Cliente_Validate(bSGECancelDummy)

            'Se foi passada uma filial de cliente
            If objPedidoVenda.iFilial > 0 Then

                'Joga a filial do cliente na tela
                Filial.Text = objPedidoVenda.iFilial
                Call Filial_Validate(bSGECancelDummy)
            End If

            'Cria um número automático para o orçamento
            Call BotaoProxNum_Click
        
        End If

    End If

    iAlterado = 0
    iComissoesAlterada = 0
    gbCarregandoTela = False

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 188982, 188983

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188984)

    End Select

    iAlterado = 0
    iComissoesAlterada = 0
    gbCarregandoTela = False

    Exit Function

End Function

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then gError 188985

    objPedidoVenda.lCodigo = StrParaLong(Codigo.Text)

    objPedidoVenda.iFilialEmpresa = giFilialEmpresa

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 188985, 188986

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188987)

    End Select

    Exit Sub

End Sub

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

            'gobjPedidoVenda.lCliente = objCliente.lCodigo

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

                If objCliente.iCodMensagem > 0 Then
                    'Lê e coloca a mensagem do Cliente na tela
                    objMensagem.iCodigo = objCliente.iCodMensagem

                    lErro = CF("Mensagem_Le", objMensagem)
                    If lErro <> SUCESSO And lErro <> 19234 Then gError 188997
                    If lErro <> SUCESSO Then gError 188998

                    Mensagem.Text = objMensagem.sDescricao

                End If

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
        
        'Coloca clientes em atraso em vermelho
        If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)
    
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
    If Len(Trim(Filial.Text)) = 0 Or giFilialAlterada = 0 Then Exit Sub

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

Public Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 191011

    If gobjPedidoVenda.dtDataEmissao <> StrParaDate(DataEmissao.Text) Then

        gobjPedidoVenda.dtDataEmissao = StrParaDate(DataEmissao.Text)

        Call ValorTotal_Calcula

    End If

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 191011

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191012)

    End Select

    Exit Sub

End Sub
'
'Private Sub objEventoTiposDeTributacao_evSelecao(obj1 As Object)
'
'Dim objTipoTrib As ClassTipoDeTributacaoMovto
'
'    Set objTipoTrib = obj1
'
'    If giFrameAtualTributacao = 1 Then
'
'        TipoTributacao.Text = objTipoTrib.iTipo
'        Call TipoTributacao_Validate(bSGECancelDummy)
'
'    Else
'
'        TipoTributacaoItem.Text = objTipoTrib.iTipo
'        Call TipoTributacaoItem_Validate(bSGECancelDummy)
'
'    End If
'
'    Me.Show
'
'    Exit Sub
'
'End Sub

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

Public Sub PlacaUF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PlacaUF_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(PlacaUF.Text)) = 0 Then Exit Sub

    'Verifica se existe o item na combo
    lErro = Combo_Item_Igual_CI(PlacaUF)
    If lErro <> SUCESSO And lErro <> 58583 Then gError 191013

    'Se não encontrar --> Erro
    If lErro = 58583 Then gError 191014

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True


    Select Case gErr

        Case 191013

        Case 191014
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191015)

    End Select

    Exit Sub

End Sub

'Public Sub TipoTributacao_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giTipoTributacaoAlterado = 1
'
'End Sub
'
'Public Sub TipoTributacao_GotFocus()
'
'Dim iTipoTributacaoAux As Integer
'
'    iTipoTributacaoAux = giTipoTributacaoAlterado
'    Call MaskEdBox_TrataGotFocus(TipoTributacao, iAlterado)
'    giTipoTributacaoAlterado = iTipoTributacaoAux
'
'End Sub
'
'Public Sub TipoTributacao_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objTipoDeTributacao As New ClassTipoDeTributacaoMovto
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_TipoTributacao_Validate
'
'    If Len(Trim(TipoTributacao.Text)) = 0 Then
'        'Limpa o campo da descrição
'        DescTipoTrib.Caption = ""
'        'Correcao por tulio.. esse exit sub nao PODE ESTAR AQUI!
'        'Exit Sub
'    End If
'
'    If (giTipoTributacaoAlterado = 1) Then
'
'        objTipoDeTributacao.iTipo = StrParaInt(TipoTributacao.Text)
'
'        'por tulio... se cod da tributacao nao for 0...
'        If objTipoDeTributacao.iTipo <> 0 Then
'            lErro = CF("TipoTributacao_Le", objTipoDeTributacao)
'            If lErro <> SUCESSO And lErro <> 27259 Then gError 191016
'
'            'Se não encontrou o Tipo da Tributação --> erro
'            If lErro = 27259 Then gError 191017
'        End If
'
'        Call BotaoGravarTrib
'
'        giTipoTributacaoAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_TipoTributacao_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 191016
'
'        Case 191017
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOTRIBUTACAO", TipoTributacao.Text)
'
'            If vbMsgRes = vbYes Then
'
'                Call Chama_Tela("TipoDeTributacao", objTipoDeTributacao)
'
'            Else
'            End If
'
'        Case Else
'            vbMsgRes = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191018)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub TipoTributacaoItem_Change()
'
'    giTipoTributacaoItemAlterado = 1
'    iAlterado = REGISTRO_ALTERADO
'
'
'End Sub
'
'Public Sub TipoTributacaoItem_GotFocus()
'
'Dim iTipoTributacaoItemAux As Integer
'
'    iTipoTributacaoItemAux = giTipoTributacaoItemAlterado
'    Call MaskEdBox_TrataGotFocus(TipoTributacaoItem, iAlterado)
'    giTipoTributacaoItemAlterado = iTipoTributacaoItemAux
'
'End Sub
'
'Public Sub TipoTributacaoItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_TipoTributacaoItem_Validate
'
'    'Se trocou o tipo de tributação
'    If giTipoTributacaoItemAlterado Then
'
'        objTributacaoTipo.iTipo = StrParaInt(TipoTributacaoItem)
'        If objTributacaoTipo.iTipo <> 0 Then
'
'            lErro = CF("TipoTributacao_Le", objTributacaoTipo)
'            If lErro <> SUCESSO And lErro <> 27259 Then gError 191019
'
'            'Se não encontrou o Tipo da Tributação --> erro
'            If lErro = 27259 Then gError 191020
'
'            DescTipoTribItem.Caption = objTributacaoTipo.sDescricao
'
'            Call BotaoGravarTribItem_Click
'        Else
'            'Limpa o campo
'            DescTipoTribItem.Caption = ""
'        End If
'
'        giTipoTributacaoItemAlterado = 0
'
'    End If
'
'    Exit Sub
'
'Erro_TipoTributacaoItem_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 191019
'
'        Case 191020
'
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOTRIBUTACAO", TipoTributacaoItem.Text)
'
'            If vbMsgRes = vbYes Then
'
'                Call Chama_Tela("TipoDeTributacao", objTributacaoTipo)
'
'            Else
'            End If
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191021)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub TribSobreDesconto_Click()
'
'    'se o frame atual for o de itens
'    If FrameItensTrib.Visible = True Then
'        'exibir o de outros
'        FrameOutrosTrib.Visible = True
'        FrameItensTrib.Visible = False
'    End If
'
'    'Call TributacaoItem_TrazerTela(gobjPedidoVenda.objTributacaoPV.objTributacaoDesconto)
'
'End Sub
'
'Public Sub TribSobreFrete_Click()
'
'    'exibir o frame de "outros"
'    FrameOutrosTrib.Visible = True
'    FrameItensTrib.Visible = False
'
'    'Call TributacaoItem_TrazerTela(gobjPedidoVenda.objTributacaoPV.objTributacaoFrete)
'
'End Sub
'
'Public Sub TribSobreItem_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'    'se houver itens na combo
'    If gobjPedidoVenda.colItensPedido.Count <> 0 Or gobjPedidoVenda.colItensPedidoSRV.Count <> 0 Then
'        'mostra o frame de itens e esconde o de outros
'        FrameItensTrib.Visible = True
'        FrameOutrosTrib.Visible = False
'        'selecionar o 1o item
'        ComboItensTrib.ListIndex = 0
'        Call ComboItensTrib_Click
'    Else
'        'senao houver itens na combo selecionar Frete
'        TribSobreFrete.Value = True
'        Call TribSobreFrete_Click
'    End If
'
'End Sub
'
'Public Sub TribSobreOutrasDesp_Click()
'
'    'se o frame atual for o de itens
'    If FrameItensTrib.Visible = True Then
'        'exibir o de outros
'        FrameOutrosTrib.Visible = True
'        FrameItensTrib.Visible = False
'    End If
'
'    'Call TributacaoItem_TrazerTela(gobjPedidoVenda.objTributacaoPV.objTributacaoOutras)
'
'End Sub
'
'Public Sub TribSobreSeguro_Click()
'
'    'se o frame atual for o de itens
'    If FrameItensTrib.Visible = True Then
'        'exibir o de outros
'        FrameOutrosTrib.Visible = True
'        FrameItensTrib.Visible = False
'    End If
'
'    'Call TributacaoItem_TrazerTela(gobjPedidoVenda.objTributacaoPV.objTributacaoSeguro)
'
'End Sub
'
'Public Sub TributacaoRecalcular_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_TributacaoRecalcular_Click
'
'    giRecalculandoTributacao = 1
'
'    lErro = ValorTotal_Calcula()
'    If lErro <> SUCESSO Then gError 191022
'
'    giRecalculandoTributacao = 0
'
'    Exit Sub
'
'Erro_TributacaoRecalcular_Click:
'
'    Select Case gErr
'
'        Case 191022
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191023)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub UnidadeMed_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro Then gError 191024

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 191024

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191025)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro Then gError 191026

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 191026

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191027)

    End Select

    Exit Sub

End Sub

'Public Sub NaturezaOp_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objNaturezaOp As New ClassNaturezaOp
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_NaturezaOp_Validate
'
'    'Se Natureza não está preenchida espelha no frame Tributação
'    If Len(Trim(NaturezaOp.ClipText)) = 0 Then
'
'        NatOpEspelho.Caption = ""
'        DescNatOp.Caption = ""
'
'    End If
'
'    'Verifica se a NaturezaOP foi informada
'    If Len(Trim(NaturezaOp.ClipText)) = 0 Or giNaturezaOpAlterada = 0 Then Exit Sub
'
'    objNaturezaOp.sCodigo = Trim(NaturezaOp.Text)
'
'    If objNaturezaOp.sCodigo < NATUREZA_SAIDA_COD_INICIAL Or objNaturezaOp.sCodigo > NATUREZA_SAIDA_COD_FINAL Then gError 191028
'
'    'Lê a NaturezaOp
'    lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
'    If lErro <> SUCESSO And lErro <> 17958 Then gError 191029
'
'    'Se não existir --> Erro
'    If lErro = 17958 Then gError 191030
'
'    'Espelha Natureza no frame de Tributação
'    NatOpEspelho.Caption = objNaturezaOp.sCodigo
'    DescNatOp.Caption = objNaturezaOp.sDescricao
'
'    If giTrazendoTribTela = 0 And gbCarregandoTela = False Then Call BotaoGravarTrib
'
'    giNaturezaOpAlterada = 0
'
'    Exit Sub
'
'Erro_NaturezaOp_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 191028
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", gErr)
'
'        Case 191029
'
'        Case 191030
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_NATUREZA_OPERACAO", NaturezaOp.Text)
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("NaturezaOperacao", objNaturezaOp)
'            Else
'            End If
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191031)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub TabelaPreco_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double

On Error GoTo Erro_TabelaPreco_Click

    iAlterado = REGISTRO_ALTERADO

    If TabelaPreco.ListIndex = -1 Then Exit Sub

    If objGridItens.iLinhasExistentes = 0 Then Exit Sub

    'Faz o tratamento para a Tabela de Preços escolhida
    lErro = Trata_TabelaPreco()
    If lErro <> SUCESSO Then gError 191032

    Exit Sub

Erro_TabelaPreco_Click:

    Select Case gErr

        Case 191032

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191032)

    End Select

    Exit Sub

End Sub

Public Sub CanalVenda_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CanalVenda_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CobrancaAutomatica_Click()

    iAlterado = REGISTRO_ALTERADO

    If CobrancaAutomatica.Value = vbChecked And Len(Trim(CondicaoPagamento.Text)) > 0 Then
        Call Cobranca_Automatica
    End If

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

    'Lê Condição a partir do código
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 191033
    
    If lErro = 19205 Then gError 191034

    'Altera PercAcrescFin
    If Len(Trim(PercAcrescFin.ClipText)) > 0 Then

        dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100
        If dPercAcresFin <> objCondicaoPagto.dAcrescimoFinanceiro Then
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
            If vbMsgRes = vbYes Then
                PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
                Call PercAcrescFin_Validate(bSGECancelDummy)
            End If
        End If
    Else
        PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
        Call PercAcrescFin_Validate(bSGECancelDummy)
    End If

    'Testa se ValorTotal está preenchido
    If Len(Trim(ValorTotal)) > 0 Then
        'Testa se DataReferencia está preenchida e ValorTotal é positivo
        If Len(Trim(DataReferencia.ClipText)) > 0 And (CDbl(ValorTotal.Caption) > 0) Then

            'Preenche o GridParcelas
            lErro = Cobranca_Automatica()
            If lErro <> SUCESSO Then gError 191035

        End If
    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_CondicaoPagamento_Click:

    Select Case gErr

        Case 191033, 191035

        Case 191034
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191036)

      End Select

    Exit Sub

End Sub

Public Sub DataReferencia_Change()

    iAlterado = REGISTRO_ALTERADO
    giDataReferenciaAlterada = REGISTRO_ALTERADO

End Sub

Public Sub FilialEntrega_Change()

    iAlterado = REGISTRO_ALTERADO
    giFilialEntregaAlterada = 1

End Sub

Public Sub FilialEntrega_Click()

Dim lErro As Long

On Error GoTo Erro_FilialEntrega_Click

    iAlterado = REGISTRO_ALTERADO

    'Se for FilialEntrega selecionada faz tratamento de impostos
    If FilialEntrega.ListIndex <> -1 Then
        If FilialEntrega_TrataAlteracao() <> SUCESSO Then gError 191037
    End If

    Exit Sub

Erro_FilialEntrega_Click:

    Select Case gErr

        Case 191037

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191038)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

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

'#################################################
'Inserido por Wagner
Dim iColuna As Integer
Dim iKey As Integer
Dim iLinha As Integer

    iColuna = GridItens.Col
    iKey = KeyAscii
    iLinha = GridItens.Row
'#################################################

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


Public Sub GridItensSRV_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItensSRV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensSRV, iAlterado)
    End If

End Sub

Public Sub GridItensSRV_EnterCell()

    Call Grid_Entrada_Celula(objGridItensSRV, iAlterado)

End Sub

Public Sub GridItensSRV_GotFocus()

    Call Grid_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub GridItensSRV_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer
Dim iColuna As Integer
Dim iKey As Integer
Dim iLinha As Integer

    iColuna = GridItensSRV.Col
    iKey = KeyAscii
    iLinha = GridItensSRV.Row

     Call Grid_Trata_Tecla(KeyAscii, objGridItensSRV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItensSRV, iAlterado)
    End If

End Sub

Public Sub GridItensSRV_LeaveCell()

    Call Saida_Celula(objGridItensSRV)

End Sub

Public Sub GridItensSRV_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItensSRV)

End Sub

Public Sub GridItensSRV_RowColChange()

    Call Grid_RowColChange(objGridItensSRV)

End Sub

Public Sub Mensagem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'se abriu o tab de detalhamento
        If Opcao.SelectedItem.Index = TAB_Tributacao Then
            lErro = gobjTribTab.TabClick
            If lErro <> SUCESSO Then gError 191039
        End If

        'Esconde o frame atual, mostra o novo
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(giFrameAtual).Visible = False
        'Armazena novo valor de giFrameAtual
        giFrameAtual = Opcao.SelectedItem.Index

        Select Case giFrameAtual

            Case TAB_Principal
                Parent.HelpContextID = IDH_PEDIDO_VENDA_DADOS_PRINCIPAIS

            Case TAB_Itens
                Parent.HelpContextID = IDH_PEDIDO_VENDA_ITENS

            Case TAB_Complemento
                Parent.HelpContextID = IDH_PEDIDO_VENDA_COMPLEMENTO

            Case TAB_Cobranca
                Parent.HelpContextID = IDH_PEDIDO_VENDA_COBRANCA

            Case TAB_Comissoes

                'se for tab de comissoes e a empresa
                'utilizar regras para o calculo de comissoes
                lErro = objTabComissoes.Comissoes_Calcula
                If lErro <> SUCESSO Then gError 191040

                Parent.HelpContextID = IDH_PEDIDO_VENDA_COMISSOES

            Case TAB_Bloqueio
                Parent.HelpContextID = IDH_PEDIDO_VENDA_BLOQUEIO

            Case TAB_Tributacao
                Parent.HelpContextID = IDH_PEDIDO_VENDA_TRIBUTACAO

        End Select


    End If

    Exit Sub

Erro_Opcao_Click:

    Select Case gErr

        Case 191039, 191040

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191041)

    End Select

    Exit Sub

End Sub
'
'Public Sub OpcaoTributacao_Click()
'Dim lErro As Long
'On Error GoTo Erro_OpcaoTributacao_Click
'
'    'Se frame selecionado não for o atual
'    If OpcaoTributacao.SelectedItem.Index <> giFrameAtualTributacao Then
'
'        If TabStrip_PodeTrocarTab(giFrameAtualTributacao, OpcaoTributacao, Me) <> SUCESSO Then Exit Sub
'
'        'Esconde o frame atual, mostra o novo
'        FrameTributacao(OpcaoTributacao.SelectedItem.Index).Visible = True
'        FrameTributacao(giFrameAtualTributacao).Visible = False
'        'Armazena novo valor de giFrameAtualTributacao
'        giFrameAtualTributacao = OpcaoTributacao.SelectedItem.Index
'
'        'se abriu o tab de detalhamento
'        If OpcaoTributacao.SelectedItem.Index = 2 Then
'            lErro = TributacaoItem_InicializaTab
'            If lErro <> SUCESSO Then gError 191042
'        End If
'
'    End If
'
'    Exit Sub
'
'Erro_OpcaoTributacao_Click:
'
'    Select Case gErr
'
'        Case 191042
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191043)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO
    giFilialAlterada = 1

End Sub


Public Sub PedidoCliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercAcrescFin_Change()

    iAlterado = REGISTRO_ALTERADO
    giPercAcresFinAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SerieNFiscalOriginal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub SerieNFiscalOriginal_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComissaoAutomatica_Click()

Dim lErro As Long

On Error GoTo Erro_ComissaoAutomatica_Click

    iAlterado = REGISTRO_ALTERADO

    If ComissaoAutomatica.Value = vbChecked Then

        lErro = objTabComissoes.Comissoes_Calcula()
        If lErro <> SUCESSO Then gError 191044

    End If

    Exit Sub

Erro_ComissaoAutomatica_Click:

    Select Case gErr

        Case 191044

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191045)

    End Select

    Exit Sub

End Sub

Private Sub TipoPedido_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Transportadora_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Transportadora_Click()
    iAlterado = REGISTRO_ALTERADO
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

Public Sub PrecoUnitario_Change()

    iAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO

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

    iAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO

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

Public Sub QuantCancelada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantCancelada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub QuantCancelada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub QuantCancelada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QuantCancelada
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantReservadaPV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantReservadaPV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub QuantReservadaPV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub QuantReservadaPV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QuantReservadaPV
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantFaturada_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantFaturada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub QuantFaturada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub QuantFaturada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QuantFaturada
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

Public Sub DescricaoProduto_Change()

    iAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO

End Sub

Public Sub DescricaoProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DescricaoProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DescricaoProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DescricaoProduto
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ServicoSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ServicoSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub ServicoSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub ServicoSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = ServicoSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UnidadeMedSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UnidadeMedSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub UnidadeMedSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub UnidadeMedSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = UnidadeMedSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantidadeSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantidadeSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub QuantidadeSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub QuantidadeSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = QuantidadeSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PrecoUnitarioSRV_Change()

    iAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO

End Sub

Public Sub PrecoUnitarioSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub PrecoUnitarioSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub PrecoUnitarioSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = PrecoUnitarioSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PrecoTotalSRV_Change()

    iAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO

End Sub

Public Sub PrecoTotalSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub PrecoTotalSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub PrecoTotalSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = PrecoTotalSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantCanceladaSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantCanceladaSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub QuantCanceladaSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub QuantCanceladaSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = QuantCanceladaSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantFaturadaSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantFaturadaSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub QuantFaturadaSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub QuantFaturadaSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = QuantFaturadaSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PercentDescSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentDescSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub PercentDescSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub PercentDescSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = PercentDescSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DescontoSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DescontoSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub DescontoSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub DescontoSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = DescontoSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataEntregaSRV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEntregaSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub DataEntregaSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub DataEntregaSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = DataEntregaSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DescServicoSRV_Change()

    iAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO

End Sub

Public Sub DescServicoSRV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItensSRV)

End Sub

Public Sub DescServicoSRV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItensSRV)

End Sub

Public Sub DescServicoSRV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItensSRV.objControle = DescServicoSRV
    lErro = Grid_Campo_Libera_Foco(objGridItensSRV)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula
    'aquii está devolvendo erro em vez de sucesso
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridParcelas.Name

                lErro = Saida_Celula_GridParcelas(objGridInt)
                If lErro <> SUCESSO Then gError 191047

            'Se for o GridItens
            Case GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 191048

            Case GridItensSRV.Name

                lErro = Saida_Celula_GridItensSRV(objGridInt)
                If lErro <> SUCESSO Then gError 191049

'            'se for o GridComissoes
            Case GridBloqueio.Name

                lErro = gobjTelaTrataBloqGen.Saida_Celula(objGridInt)
                If lErro <> SUCESSO Then gError 191050

            'se for o GridComissoesEmissao
            Case GridComissoes.Name

                lErro = objTabComissoes.Saida_Celula_GridComissoes(objGridInt)
                If lErro <> SUCESSO Then gError 191051


        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 191052

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 191047 To 191052

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191053)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Produto
        Case iGrid_Produto_Col
            lErro = Saida_Celula_Produto(objGridInt)
            If lErro <> SUCESSO Then gError 191054

        'Se for a de Unidade de Medida
        Case iGrid_UnidadeMed_Col
            lErro = Saida_Celula_UM(objGridInt)
            If lErro <> SUCESSO Then gError 191055

        'Se for a de Quantidade Pedida
        Case iGrid_QuantPedida_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 191056

        'Se for a de Quantidade Cancelada
        Case iGrid_QuantCancel_Col
            lErro = Saida_Celula_QuantCancelada(objGridInt)
            If lErro <> SUCESSO Then gError 191057

        'Se for a de Preço Unitário
        Case iGrid_ValorUnitario_Col
            lErro = Saida_Celula_PrecoUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 191058

        'Se for a de desconto produto
        Case iGrid_DescProduto_Col
            lErro = Saida_Celula_DescProduto(objGridInt)
            If lErro <> SUCESSO Then gError 191059

        Case iGrid_PercDesc_Col
            lErro = Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then gError 191060

    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 191054 To 191060

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191061)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridItensSRV(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItensSRV

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Produto
        Case iGrid_ServicoSRV_Col
            lErro = Saida_Celula_ServicoSRV(objGridInt)
            If lErro <> SUCESSO Then gError 191062

        'Se for a de Unidade de Medida
        Case iGrid_UnidadeMedSRV_Col
            lErro = Saida_Celula_UMSRV(objGridInt)
            If lErro <> SUCESSO Then gError 191063

        'Se for a de Quantidade Pedida
        Case iGrid_QuantPedidaSRV_Col
            lErro = Saida_Celula_QuantidadeSRV(objGridInt)
            If lErro <> SUCESSO Then gError 191064

        'Se for a de Quantidade Cancelada
        Case iGrid_QuantCancelSRV_Col
            lErro = Saida_Celula_QuantCanceladaSRV(objGridInt)
            If lErro <> SUCESSO Then gError 191065

        'Se for a de Preço Unitário
        Case iGrid_ValorUnitarioSRV_Col
            lErro = Saida_Celula_PrecoUnitarioSRV(objGridInt)
            If lErro <> SUCESSO Then gError 191066

        'Se for a de Data de Entrega
        Case iGrid_DataEntregaSRV_Col
            lErro = Saida_Celula_DataEntregaSRV(objGridInt)
            If lErro <> SUCESSO Then gError 191067

        'Se for a de desconto produto
        Case iGrid_DescServicoSRV_Col
            lErro = Saida_Celula_DescServicoSRV(objGridInt)
            If lErro <> SUCESSO Then gError 191068

        Case iGrid_PercDescSRV_Col
            lErro = Saida_Celula_PercentDescSRV(objGridInt)
            If lErro <> SUCESSO Then gError 191069

    End Select

    Saida_Celula_GridItensSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItensSRV:

    Saida_Celula_GridItensSRV = gErr

    Select Case gErr

        Case 191062 To 191069

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191070)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto Data que está deixando de ser a corrente

Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) > 0 Then

        lErro = CF2(Me, "ProdutoSRV_Saida_Celula")
        If lErro <> SUCESSO Then gError 191071
        
    End If

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191073

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 191071, 191073
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191074)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ServicoSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula ServicoSRV que está deixando de ser a corrente

Dim lErro As Long
Dim sServico As String

On Error GoTo Erro_Saida_Celula_ServicoSRV

    Set objGridInt.objControle = ServicoSRV

    If Len(Trim(ServicoSRV.ClipText)) > 0 Then

        lErro = ServicoSRV_Saida_Celula()
        If lErro <> SUCESSO And lErro <> 26658 Then gError 191075
        
        If lErro = 26658 Then gError 191076
    End If

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ServicoSRV_Col) = ""

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191077

    Saida_Celula_ServicoSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_ServicoSRV:

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ServicoSRV_Col) = ""

    Saida_Celula_ServicoSRV = gErr

    Select Case gErr

        Case 191075, 191077
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191076

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191078)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UM(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidadede Medida que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double
Dim dQuantCancelada As Double
Dim dQuantFaturada As Double
Dim sUmAnterior As String
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim sProduto As String
Dim iPreenchido As Integer
Dim dPrecoUnitario As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_Saida_Celula_UM

    Set objGridInt.objControle = UnidadeMed

    'Guarda a Unidade de medida anteriormente selecionada
    sUmAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

    'Coloca a Um no grid de itens
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191079

    'Se a Um selecionada agora é diferente da anterior
    If sUmAnterior <> UnidadeMed.Text And objProduto.iKitVendaComp <> MARCADO Then

        objProduto.sCodigo = gobjPedidoVenda.colItensPedido(GridItens.Row).sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 191080


        lErro = Atualiza_UM(GridItens.Row, sUmAnterior, UnidadeMed.Text)
        If lErro <> SUCESSO Then gError 191081

        'Reprocessa a reserva do produto
        dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col))
        dQuantCancelada = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantCancel_Col))
        dQuantFaturada = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantFaturada_Col))

        lErro = Reserva_Processa(GridItens.Row, dQuantidade, dQuantCancelada, dQuantFaturada)
        If lErro <> SUCESSO Then gError 191082

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 191085

    End If

    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case 191079 To 191081, 191083 To 191086
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191082
            GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = sUmAnterior
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191087)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_UMSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidadede Medida que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double
Dim dQuantCancelada As Double
Dim dQuantFaturada As Double
Dim sUmAnterior As String
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim sProduto As String
Dim iPreenchido As Integer
Dim dPrecoUnitario As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_Saida_Celula_UMSRV

    Set objGridInt.objControle = UnidadeMedSRV

    'Guarda a Unidade de medida anteriormente selecionada
    sUmAnterior = GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_UnidadeMedSRV_Col)

    'Coloca a Um no grid de itens
    GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_UnidadeMedSRV_Col) = UnidadeMedSRV.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191089

    'Se a Um selecionada agora é diferente da anterior
    If sUmAnterior <> UnidadeMedSRV.Text And objProduto.iKitVendaComp <> MARCADO Then

        objProduto.sCodigo = gobjPedidoVenda.colItensPedidoSRV(GridItensSRV.Row).sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 191090

        lErro = Atualiza_UMSRV(GridItensSRV.Row, sUmAnterior, UnidadeMedSRV.Text)
        If lErro <> SUCESSO Then gError 191091

        'Reprocessa a reserva do produto
        dQuantidade = StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantPedidaSRV_Col))
        dQuantCancelada = StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantCancelSRV_Col))
        dQuantFaturada = StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantFaturadaSRV_Col))

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 191092

    End If

    Saida_Celula_UMSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_UMSRV:

    Saida_Celula_UMSRV = gErr

    Select Case gErr

        Case 191089 To 191093
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191094)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Function Pesos_Recalcula() As Long
    Pesos_Recalcula = CF2(Me, "PSRV_Pesos_Recalcula")
End Function

Public Function PSRV_Pesos_Recalcula(ByVal objCT As Object) As Long
'calcula pesos e volume total

Dim lErro As Long
Dim dFator As Double
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dPesoBruto As Double, dPesoBrutoItem As Double
Dim lQuantVolumes As Long, lQuantVolumesItens As Long
Dim dPesoLiquido As Double, dPesoLiquidoItem As Double
Dim dQuantPedida As Double, bPesoBrutoInvalido As Boolean
Dim dQuantCancelada As Double

On Error GoTo Erro_PSRV_Pesos_Recalcula

    lErro = Calcula_VolumeTotal
    If lErro <> SUCESSO Then gError 191095

    dPesoBruto = 0
    dPesoLiquido = 0

    lQuantVolumes = 0
    bPesoBrutoInvalido = False

    'Para cada linha do Grid de Itens
    For iIndice = 1 To objGridItens.iLinhasExistentes

        'Se o produto e a UM estiverem preenchidos
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) > 0 Then

            'Formata o produto
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 191096

            'Lê o produto
            objProduto.sCodigo = sProdutoFormatado
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 191097

            'Se o produto não está cadastrado, erro
            If lErro = 28030 Then gError 191098

            'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then gError 191099

            dQuantPedida = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col))
            dQuantCancelada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))

            'Acumula o peso Bruto e o peso líquido
            dPesoBruto = dPesoBruto + objProduto.dPesoBruto * dFator * (dQuantPedida - dQuantCancelada)
            dPesoLiquido = dPesoLiquido + objProduto.dPesoLiq * dFator * (dQuantPedida - dQuantCancelada)

        End If

    Next

    PesoBruto.Text = Format(dPesoBruto, PesoBruto.Format)
    PesoLiquido.Text = Format(dPesoLiquido, PesoBruto.Format)

    PSRV_Pesos_Recalcula = SUCESSO

    Exit Function

Erro_PSRV_Pesos_Recalcula:

    PSRV_Pesos_Recalcula = gErr

    Select Case gErr

        Case 191095 To 191097, 191099

        Case 191098
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191100)

    End Select

    Exit Function

End Function

Private Function Calcula_PesoBruto(objProduto As ClassProduto, dQuantProduto As Double, dPesoBruto As Double, lQuantVolumes As Long) As Long

Dim dCapEmbalagem As Double, objEmbalagem As New ClassEmbalagem
Dim lErro As Long

On Error GoTo Erro_Calcula_PesoBruto

    dPesoBruto = 0
    lQuantVolumes = 0

    'se tem embalagem padrao asssociada
    If objProduto.iEmbalagem <> 0 Then

        objEmbalagem.iCodigo = objProduto.iEmbalagem

        lErro = CF("Embalagem_Le", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 82763 Then gError 191101
        
        If lErro <> SUCESSO Then gError 191102

        'calcula a capacidade de cada embalagem para este produto
        dCapEmbalagem = objProduto.dPesoEspecifico * objEmbalagem.dCapacidade

        If dCapEmbalagem <> 0 Then

            'calcula quantas embalagens serão necessárias
            lQuantVolumes = AsymUp((dQuantProduto * objProduto.dPesoLiq) / dCapEmbalagem)

            dPesoBruto = (objProduto.dPesoLiq * dQuantProduto) + (lQuantVolumes * objEmbalagem.dPeso)

        End If

    End If

    If dPesoBruto = 0 Then

        If objProduto.dPesoBruto <> 0 Then

            dPesoBruto = objProduto.dPesoBruto * dQuantProduto

        End If

    End If

    Calcula_PesoBruto = SUCESSO

    Exit Function

Erro_Calcula_PesoBruto:

    Calcula_PesoBruto = gErr

    Select Case gErr

        Case 191101

        Case 191102
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objEmbalagem.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191103)

    End Select

    Exit Function

End Function

Function AsymUp(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
Dim Temp As Double
     Temp = Int(X * Factor)
     AsymUp = (Temp + IIf(X = Temp, 0, 1)) / Factor
End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidadeque está deixando de ser a corrente

Dim lErro As Long
Dim objItemPedido As New ClassItemPedido
Dim bQuantidadeIgual As Boolean
Dim dQuantidade As Double
Dim dQuantidadeCancelada As Double
Dim dQuantidadeFaturada As Double
Dim iIndice As Integer
Dim iItemAtual As Integer
Dim dPrecoUnitario As Double
Dim dQuantAnterior As Double
Dim objProduto As New ClassProduto
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    bQuantidadeIgual = False

    If Len(Quantidade.Text) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 191104

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    'Comparação com quantidade anterior
    dQuantAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col))
    If dQuantAnterior = StrParaDbl(Quantidade.Text) Then bQuantidadeIgual = True

    GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col) = Formata_Estoque(dQuantidade)

    If Not bQuantidadeIgual Then

        dQuantidade = StrParaDbl(Quantidade.Text)
        dQuantidadeCancelada = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantCancel_Col))
        dQuantidadeFaturada = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantFaturada_Col))
        If dQuantidadeCancelada > 0 And dQuantidade < dQuantidadeCancelada Then gError 191105
        If dQuantidadeFaturada > 0 And dQuantidade - dQuantidadeCancelada < dQuantidadeFaturada Then gError 191106

        '####################################################################
        'Inserido por Wagner 18/05/2006
        objProduto.sCodigo = gobjPedidoVenda.colItensPedido(GridItens.Row).sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 191107
        '####################################################################

        If objProduto.iKitVendaComp <> MARCADO Then
            'Tenta fazer uma alocação p\ oitem de acordo com a nova UM
            lErro = Reserva_Processa(GridItens.Row, dQuantidade, dQuantidadeCancelada, dQuantidadeFaturada)
            If lErro <> SUCESSO Then gError 191108

        Else

            gError 193450

'            '####################################################################
'            'Inserido por Wagner 18/05/2006
'            objKitVenda.sProduto = gobjPedidoVenda.colItensPedido(GridItens.Row).sProduto
'            objKitVenda.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
'            objKitVenda.dQuantidade = dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada
'
'            gobjPedidoVenda.colItensPedido.Item(GridItens.Row).sUnidadeMed = objKitVenda.sUM
'            gobjPedidoVenda.colItensPedido.Item(GridItens.Row).dQuantidade = dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada
'
'            lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjPedidoVenda.colItensPedido(GridItens.Row).colItensRomaneioGrade)
'            If lErro <> SUCESSO Then gError 191109
'
'            For Each objItemRomaneio In gobjPedidoVenda.colItensPedido(GridItens.Row).colItensRomaneioGrade
'
'                lErro = Reserva_Processa_ItemGrade(objItemRomaneio, GridItens.Row)
'                If lErro <> SUCESSO Then gError 191110
'
'            Next
'
'            Call Atualiza_Grid_Alocacao(gobjPedidoVenda.colItensPedido.Item(GridItens.Row))
'            '####################################################################

        End If

    End If

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191111

    'Preço unitário
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))

    'Recalcula preço do ítem e valor total da nota
    If dPrecoUnitario > 0 And Not bQuantidadeIgual Then
        Call CF2(Me, "PSRV_PrecoTotal_Calcula", GridItens.Row)
'        lErro = ValorTotal_Calcula()
'        If lErro <> SUCESSO Then gError 191112
    End If

    If Not bQuantidadeIgual Then

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 191113

        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 191112

    End If
    
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 191112
    
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 191104, 191197, 191109 To 191114
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191105
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_PEDIDA_INFERIOR_CANCELADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191106
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_FATURADA_SUPERIOR", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191108
            GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col) = Formata_Estoque(dQuantAnterior)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 193450
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_KITVENDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191115)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantidadeSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidadeque está deixando de ser a corrente

Dim lErro As Long
Dim objItemPedido As New ClassItemPedido
Dim bQuantidadeIgual As Boolean
Dim dQuantidade As Double
Dim dQuantidadeCancelada As Double
Dim dQuantidadeFaturada As Double
Dim iIndice As Integer
Dim iItemAtual As Integer
Dim dPrecoUnitario As Double
Dim dQuantAnterior As Double
Dim objProduto As New ClassProduto
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_Saida_Celula_QuantidadeSRV

    Set objGridInt.objControle = QuantidadeSRV

    bQuantidadeIgual = False

    If Len(QuantidadeSRV.Text) > 0 Then

        lErro = Valor_Positivo_Critica(QuantidadeSRV.Text)
        If lErro <> SUCESSO Then gError 191116

        QuantidadeSRV.Text = Formata_Estoque(QuantidadeSRV.Text)

    End If

    'Comparação com quantidade anterior
    dQuantAnterior = StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantPedidaSRV_Col))
    If dQuantAnterior = StrParaDbl(QuantidadeSRV.Text) Then bQuantidadeIgual = True

    GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantPedidaSRV_Col) = Formata_Estoque(dQuantidade)

    If Not bQuantidadeIgual Then

        dQuantidade = StrParaDbl(QuantidadeSRV.Text)
        dQuantidadeCancelada = StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantCancelSRV_Col))
        dQuantidadeFaturada = StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantFaturadaSRV_Col))
        If dQuantidadeCancelada > 0 And dQuantidade < dQuantidadeCancelada Then gError 191117
        If dQuantidadeFaturada > 0 And dQuantidade - dQuantidadeCancelada < dQuantidadeFaturada Then gError 191118

    End If

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191119

    'Preço unitário
    dPrecoUnitario = StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ValorUnitarioSRV_Col))

    'Recalcula preço do ítem e valor total da nota
    If dPrecoUnitario > 0 And Not bQuantidadeIgual Then
        Call CF2(Me, "PSRV_PrecoTotal_CalculaSRV", GridItensSRV.Row)
'        lErro = ValorTotal_Calcula()
'        If lErro <> SUCESSO Then gError 191120
    End If

    If Not bQuantidadeIgual Then

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 191121

        lErro = gobjTribTab.Alteracao_Item_Grid(GridItensSRV.Row, True)
        If lErro <> SUCESSO Then gError 191112

    End If

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 191120
        
    Saida_Celula_QuantidadeSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadeSRV:

    Saida_Celula_QuantidadeSRV = gErr

    Select Case gErr

        Case 191116, 191119 To 191121
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191117
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_PEDIDA_INFERIOR_CANCELADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191118
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_FATURADA_SUPERIOR", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191122)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_PedidoVenda()

Dim lErro As Long
Dim sNatureza As String

On Error GoTo Erro_Limpa_Tela_PedidoVenda

    sNatureza = gobjTribTab.sNatOpInterna

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Call Limpa_Tela_PedidoVenda2

    gobjTribTab.sNatOpInterna = sNatureza

    'Call BotaoGravarTrib

    iAlterado = 0
    iComissoesAlterada = 0
    gbCarregandoTela = False

    Exit Sub

Erro_Limpa_Tela_PedidoVenda:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191123)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 191124

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 191125

        If lErro <> SUCESSO Then gError 191126 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 191127

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 191128

    Exit Sub

Erro_TabelaPreco_Validate:

    Cancel = True

    Select Case gErr

    Case 191124, 191125, 191127

    Case 191126
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")

        If vbMsgRes = vbYes Then
            'Preenche o objTabela com o Codigo
            If Len(Trim(TabelaPreco.Text)) > 0 Then objTabelaPreco.iCodigo = CInt(TabelaPreco.Text)
            'Chama a tela de Tabelas de Preço
            Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
        Else
            'Segura o foco

        End If

    Case 191128
        Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191129)

    End Select

    Exit Sub

End Sub

Public Sub CondicaoPagamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult
Dim dPercAcresFin As Double

On Error GoTo Erro_Condicaopagamento_Validate

    'Verifica se a Condicaopagamento foi preenchida
    If Len(Trim(CondicaoPagamento.Text)) = 0 Then Exit Sub

    'Verifica se é uma Condicaopagamento selecionada
    If CondicaoPagamento.Text = CondicaoPagamento.List(CondicaoPagamento.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(CondicaoPagamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 191130

    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 191131
        
        If lErro = 19205 Then gError 191132

        'Testa se pode ser usada em Contas a Receber
        If objCondicaoPagto.iEmRecebimento = 0 Then gError 191133

        'Coloca na Tela
        CondicaoPagamento.Text = iCodigo & SEPARADOR & objCondicaoPagto.sDescReduzida

        'Altera PercAcrescFin
        If Len(Trim(PercAcrescFin.ClipText)) > 0 Then
            dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100
            If dPercAcresFin <> objCondicaoPagto.dAcrescimoFinanceiro Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
                If vbMsgRes = vbYes Then
                    PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
                    Call PercAcrescFin_Validate(bSGECancelDummy)
                End If
            End If
        Else
            PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
            Call PercAcrescFin_Validate(bSGECancelDummy)
        End If

        'Se ValorTotal e DataReferencia estiverem preenchidos, preenche GridParcelas
        If Len(Trim(ValorTotal)) > 0 Then
            If Len(Trim(DataReferencia.ClipText)) > 0 And CLng(ValorTotal.Caption) > 0 Then

                'Preenche o GridParcelas
                lErro = Cobranca_Automatica()
                If lErro <> SUCESSO Then gError 191134

            End If
        End If

    End If

    'Não encontrou o valor que era STRING
    If lErro = 6731 Then gError 191135

    Exit Sub

Erro_Condicaopagamento_Validate:

    Cancel = True

    Select Case gErr

        Case 191130, 191131, 191134

       Case 191132
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If

        Case 191133
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", gErr, objCondicaoPagto.iCodigo)

        Case 191135
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191136)

    End Select

    Exit Sub

End Sub

'Public Sub NaturezaOp_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giNaturezaOpAlterada = 1
'
'End Sub

Public Function PrecoUnitario_Calcula(sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long
    PrecoUnitario_Calcula = CF2(Me, "PSRV_PrecoUnitario_Calcula", sUM, objTabelaPrecoItem, dPrecoUnitario)
End Function

Public Function PSRV_PrecoUnitario_Calcula(ByVal objCT As Object, sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long
'Calcula o Preço unitário do item de acordo com a UM e a tabela de preço

Dim objProduto As New ClassProduto
Dim objUM As New ClassUnidadeDeMedida
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim lErro As Long
Dim dPercAcresFin As Double
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_PSRV_PrecoUnitario_Calcula

    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 191137

    If lErro = 28030 Then gError 191138
    
    'Converte a quantidade para a UM de Venda
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 191139

    dPrecoUnitario = objTabelaPrecoItem.dPreco * dFator

    'Recolhe o percentual de acréscimo financeiro
    dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100

    'Calcula o Preço unitário
    If dPercAcresFin <> 0 Then
        dPrecoUnitario = dPrecoUnitario * (1 + dPercAcresFin)
    End If

    PSRV_PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_PSRV_PrecoUnitario_Calcula:

    PSRV_PrecoUnitario_Calcula = gErr

    Select Case gErr

        Case 191137, 191139

        Case 191138
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objTabelaPrecoItem.sCodProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191140)

    End Select

    Exit Function

End Function

Function ValorTotal_Calcula() As Long
'Calcula o Valor Total do Pedido

'Dim dValorDespesas As Double
'Dim dValorProdutos As Double
Dim dValorTotal As Double
'Dim dValorFrete As Double
'Dim dValorSeguro As Double
'Dim dValorIPI As Double
'Dim dValorICMSSubst As Double
'Dim vbMsgRes As VbMsgBoxResult
'Dim dValorAposIR As Double
'Dim dValorIRRF As Double
Dim lErro As Long ', dValorISS As Double

On Error GoTo Erro_ValorTotal_Calcula

    lErro = gobjTribTab.ValorTotal_Calcula(dValorTotal)
    If lErro <> SUCESSO Then gError 191141

'    If Not gbCarregandoTela Then
'        'Atualiza os valores de tributação
'        lErro = AtualizarTributacao()
'        If lErro <> SUCESSO Then gError 191141
'    End If
'
'    'Recolhe os valores da tela
'    If Len(Trim(ValorProdutos.Caption)) > 0 And IsNumeric(ValorProdutos.Caption) Then dValorProdutos = CDbl(ValorProdutos.Caption)
'    If Len(Trim(ValorFrete.Text)) > 0 And IsNumeric(ValorFrete.Text) Then dValorFrete = CDbl(ValorFrete.Text)
'    If Len(Trim(ValorIRRF.Text)) > 0 And IsNumeric(ValorIRRF.Text) Then dValorIRRF = CDbl(ValorIRRF.Text)
'    If Len(Trim(ValorSeguro.Text)) > 0 And IsNumeric(ValorSeguro.Text) Then dValorSeguro = CDbl(ValorSeguro.Text)
'    If Len(Trim(ValorDespesas.Text)) > 0 And IsNumeric(ValorDespesas.Text) Then dValorDespesas = CDbl(ValorDespesas.Text)
'    If Len(Trim(ICMSSubstValor1.Caption)) > 0 And IsNumeric(ICMSSubstValor1.Caption) Then dValorICMSSubst = CDbl(ICMSSubstValor1.Caption)
'    If Len(Trim(IPIValor1.Caption)) > 0 And IsNumeric(IPIValor1.Caption) Then dValorIPI = CDbl(IPIValor1.Caption)
'    If Len(Trim(ISSValor.Text)) > 0 And IsNumeric(ISSValor.Text) And ISSIncluso.Value = vbUnchecked Then dValorISS = CDbl(ISSValor.Text)
'
'    'Calcula o Valor Total
'    dValorTotal = dValorProdutos + dValorFrete + dValorSeguro + dValorDespesas + dValorIPI + dValorICMSSubst + dValorISS
'
'    dValorAposIR = dValorTotal - dValorIRRF
'
'    If dValorTotal <> 0 And dValorIRRF > 0 And dValorAposIR < 0 Then
'
'        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_IR_FONTE_MAIOR_VALOR_TOTAL", dValorIRRF, dValorTotal)
'        ValorIRRF.Text = ""
'
'        Call ValorIRRF_Validate(bSGECancelDummy)
'
'        'Faz a atualização dos valores da tributação
'        lErro = AtualizarTributacao()
'        If lErro <> SUCESSO Then gError 191142
'
'    End If

    ValorTotal.Caption = Format(dValorTotal, "Standard")
    
    Call ValorDescontoTit_Validate(bSGECancelDummy)

    Call ValorDescontoItens_Calcula

    If Not gbCarregandoTela Then

        'Faz o cálculo automático das comissões
        lErro = Cobranca_Automatica()
        If lErro <> SUCESSO Then gError 191143
        'Alterado por Tulio em 28/03/02
        'Faz o cálculo automático das comissões
        lErro = objTabComissoes.Comissoes_Calcula_Padrao()
        If lErro <> SUCESSO Then gError 191144

    End If

    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 191141 To 191144

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191145)

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
Dim sServicoFormatado As String
Dim iServicoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 191146

    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", GridItensSRV.TextMatrix(iLinha, iGrid_ServicoSRV_Col), sServicoFormatado, iServicoPreenchido)
    If lErro <> SUCESSO Then gError 191147

    Select Case objControl.Name

        Case Produto.Name
            'Se o produto estiver preenchido desabilita
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
            End If

        Case ServicoSRV.Name
            'Se o produto estiver preenchido desabilita
            If iServicoPreenchido <> PRODUTO_VAZIO Then
                ServicoSRV.Enabled = False
            Else
                ServicoSRV.Enabled = True
            End If


        Case UnidadeMed.Name
            'guarda a um go grid nessa coluna
            sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
            'Se já houve um faturamento (parcial ou não ) desse item desabilita a UM
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_QuantFaturada_Col))) > 0 Then
                If CDbl(GridItens.TextMatrix(iLinha, iGrid_QuantFaturada_Col)) > 0 Then
                    UnidadeMed.Enabled = False
                Else
                    UnidadeMed.Enabled = True
                End If
            Else
                UnidadeMed.Enabled = True
            End If

            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = UnidadeMed.Text

            UnidadeMed.Clear

            If iProdutoPreenchido <> PRODUTO_VAZIO Then

                objProduto.sCodigo = sProdutoFormatado
                'Lê o produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 191148

                If lErro = 28030 Then gError 191149

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lê as UMs do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 191150
                'Carrega a combo de UMs
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next

                'Tento selecionar na Combo a Unidade anterior
                If UnidadeMed.ListCount <> 0 Then

                    For iIndice = 0 To UnidadeMed.ListCount - 1

                        If UnidadeMed.List(iIndice) = sUnidadeMed Then
                            UnidadeMed.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                End If

            Else
                UnidadeMed.Enabled = False
            End If

        Case UnidadeMedSRV.Name
            'guarda a um go grid nessa coluna
            sUM = GridItensSRV.TextMatrix(iLinha, iGrid_UnidadeMedSRV_Col)
            'Se já houve um faturamento (parcial ou não ) desse item desabilita a UM
            If Len(Trim(GridItensSRV.TextMatrix(iLinha, iGrid_QuantFaturadaSRV_Col))) > 0 Then
                If CDbl(GridItensSRV.TextMatrix(iLinha, iGrid_QuantFaturadaSRV_Col)) > 0 Then
                    UnidadeMedSRV.Enabled = False
                Else
                    UnidadeMedSRV.Enabled = True
                End If
            Else
                UnidadeMedSRV.Enabled = True
            End If

            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = UnidadeMedSRV.Text

            UnidadeMedSRV.Clear

            If iServicoPreenchido <> PRODUTO_VAZIO Then

                objProduto.sCodigo = sServicoFormatado
                'Lê o produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 191151

                If lErro = 28030 Then gError 191152

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lê as UMs do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 191153
                'Carrega a combo de UMs
                For Each objUM In colSiglas
                    UnidadeMedSRV.AddItem objUM.sSigla
                Next

                'Tento selecionar na Combo a Unidade anterior
                If UnidadeMedSRV.ListCount <> 0 Then

                    For iIndice = 0 To UnidadeMedSRV.ListCount - 1

                        If UnidadeMedSRV.List(iIndice) = sUnidadeMed Then
                            UnidadeMedSRV.ListIndex = iIndice
                            Exit For
                        End If
                    Next
                End If

            Else
                UnidadeMedSRV.Enabled = False
            End If


        Case PrecoUnitario.Name, PercentDesc.Name
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case PrecoUnitarioSRV.Name, PercentDescSRV.Name, DataEntregaSRV.Name
            'Se o produto estiver preenchido, habilita o controle
            If iServicoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If


        Case DescricaoProduto.Name, DescServicoSRV.Name
            objControl.Enabled = True

        '******** O TRATAMENTO DE QUANTIDADE E QUANTCANCELADA FOI DESTACADO PARA TRATAR GRADE
        Case Quantidade.Name, QuantCancelada.Name

            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Or left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case QuantidadeSRV.Name, QuantCanceladaSRV.Name

            'Se o produto estiver preenchido, habilita o controle
            If iServicoPreenchido = PRODUTO_VAZIO Or left(GridItensSRV.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
'
'        Case TipoBloqueio.Name, ResponsavelBL.Name
'            'Se for algum Bloqueio previamente cadastrado, desabilita o controle
'            If iLinha <= giGridBloqueioLinhasBloqueadas Then
'                objControl.Enabled = False
'            Else
'                objControl.Enabled = True
'            End If

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

        Case TipoDesconto2.Name, TipoDesconto3.Name
            'Habilita os campos de desconto em sequencia
            If Len(Trim(GridParcelas.TextMatrix(iLinha, GridParcelas.Col - 4))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

    End Select
    
    lErro = gobjTelaTrataBloqGen.Rotina_Grid_Enabled(iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 198437

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 191146 To 191148, 191149 To 191151, 191153, 198437

        Case 191149, 191152
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191154)

    End Select

    Exit Sub

End Sub

Public Function PSRV_PrecoTotal_Calcula(ByVal objCT As Object, ByVal iLinha As Integer) As Long

Dim dPrecoTotal As Double
Dim dPrecoTotalReal As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dQuantidadeCancelada As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim lTamanho As Long
Dim dValorTotal As Double, dValorTotalB As Double, dValorDesconto As Double
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_PSRV_PrecoTotal_Calcula

    'Quantidades e preço unitário
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))
    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_QuantPedida_Col))
    dQuantidadeCancelada = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_QuantCancel_Col))

    'Cálculo do desconto
    lTamanho = Len(Trim(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col)))
    If lTamanho > 0 Then
        dPercentDesc = CDbl(Format(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col), "General Number"))
    Else
        dPercentDesc = 0
    End If

    dPrecoTotal = dPrecoUnitario * (dQuantidade - dQuantidadeCancelada)

    'Se percentual for >0 tira o desconto
    If dPercentDesc > 0 Then dDesconto = StrParaDbl(Format(dPercentDesc * dPrecoTotal, "Standard"))
    dPrecoTotalReal = dPrecoTotal - dDesconto

    'Coloca valor do desconto no Grid
    If dDesconto > 0 Then
        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
    Else
        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
    End If

    'Coloca preco total do ítem no grid
    GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(dPrecoTotalReal, "Standard")
    GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(dPrecoTotal, "Standard")

'    For iIndice = 1 To objGridItensSRV.iLinhasExistentes
'        If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
'            If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
'        End If
'        dValorTotalB = dValorTotalB + StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRVB_Col))
'        dValorDesconto = dValorDesconto + StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col))
'    Next
'
'    'Calcula a soma dos valores de produtos
'    For iIndice = 1 To objGridItens.iLinhasExistentes
'        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
'            If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'        End If
'        dValorTotalB = dValorTotalB + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
'        dValorDesconto = dValorDesconto + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
'    Next

'    '############################################
'    'Alterado por Wagner
'    'ANTIGO
''    If gdDesconto > 0 Then
''        dValorDesconto = gdDesconto * dValorTotal
''    ElseIf Len(Trim(ValorDesconto.Text)) > 0 And IsNumeric(ValorDesconto.Text) Then
''        dValorDesconto = CDbl(ValorDesconto.Text)
''    End If
''    dValorTotal = dValorTotal - dValorDesconto
'    'NOVO
'    dValorDesconto = gobjPedidoVenda.dValorDesconto
'    dValorTotal = dValorTotal - dValorDesconto
'    '############################################
'
'    'Verifica se o valor de desconto é maior que o valor dos produtos
'    If dValorTotal < 0 And dValorDesconto > 0 Then
'
'        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", dValorDesconto, dValorTotal)
'
'        gdDesconto = 0
'        ValorDesconto.Text = ""
'        giValorDescontoAlterado = 0
'        dValorDesconto = 0
'
'        'Para tributação
'        gobjPedidoVenda.dValorDesconto = dValorDesconto
'
'        Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, dValorDesconto)
'
'        'Faz a atualização dos valores da tributação
'        lErro = gobjTribTab.AtualizarTributacao()
'        If lErro <> SUCESSO Then gError 191155
'
'        'Calcula a soma dos valores de produtos
'        dValorTotal = 0
'        For iIndice = 1 To objGridItens.iLinhasExistentes
'            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
'                If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'            End If
'        Next
'
'        For iIndice = 1 To objGridItensSRV.iLinhasExistentes
'            If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
'                If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
'            End If
'        Next
'
'    End If

    Call SubTotal_Calcula

'    'Coloca valor total dos produtos na tela
'    ValorProdutos.Caption = Format(dValorTotal, "Standard")
    
'    'Altera ítem na coleção de ítens de gobjPedidoVenda que dá suporte p/ Tributação
'    Call Tributacao_Alteracao_Item_Grid(iLinha)

    PSRV_PrecoTotal_Calcula = SUCESSO

    Exit Function

Erro_PSRV_PrecoTotal_Calcula:

    PSRV_PrecoTotal_Calcula = gErr

    Select Case gErr

        Case 191155

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191156)

    End Select

    Exit Function

End Function

Private Sub SubTotal_Calcula()
'Faz a soma da Coluna passado no Grid passado e devolve em dValorTotal

Dim iIndice As Integer
Dim dValorTotal As Double, dValorTotalB As Double, dValorDesconto As Double

    For iIndice = 1 To objGridItensSRV.iLinhasExistentes
        If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
            If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
        End If
        dValorTotalB = dValorTotalB + StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRVB_Col))
        dValorDesconto = dValorDesconto + StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col))
    Next

    'Calcula a soma dos valores de produtos
    For iIndice = 1 To objGridItens.iLinhasExistentes
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
            If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        End If
        dValorTotalB = dValorTotalB + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
        dValorDesconto = dValorDesconto + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
    Next

    'Coloca valor total dos produtos na tela
    ValorProdutos.Caption = Format(dValorTotal, "Standard")
    
    ValorProdutos2.Caption = Format(dValorTotalB, "Standard")
    
    ValorDescontoItens.Text = Format(dValorDesconto, "Standard")
    dValorDescontoItensAnt = dValorDesconto
    If dValorTotalB > 0 Then
        PercDescontoItens.Text = Format(dValorDesconto / dValorTotalB * 100, "FIXED")
        dPercDescontoItensAnt = Arredonda_Moeda(dValorDesconto / dValorTotalB, 4)
    Else
        PercDescontoItens.Text = Format(0, "FIXED")
        dPercDescontoItensAnt = 0
    End If

    Exit Sub

End Sub


Public Function PSRV_PrecoTotal_CalculaSRV(ByVal objCT As Object, ByVal iLinha As Integer) As Long

Dim dPrecoTotal As Double
Dim dPrecoTotalReal As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dQuantidadeCancelada As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim lTamanho As Long
Dim dValorTotal As Double
Dim iIndice As Integer
Dim dValorDesconto As Double
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_PSRV_PrecoTotal_CalculaSRV

    'Quantidades e preço unitário
    dPrecoUnitario = StrParaDbl(GridItensSRV.TextMatrix(iLinha, iGrid_ValorUnitarioSRV_Col))
    dQuantidade = StrParaDbl(GridItensSRV.TextMatrix(iLinha, iGrid_QuantPedidaSRV_Col))
    dQuantidadeCancelada = StrParaDbl(GridItensSRV.TextMatrix(iLinha, iGrid_QuantCancelSRV_Col))

    'Cálculo do desconto
    lTamanho = Len(Trim(GridItensSRV.TextMatrix(iLinha, iGrid_PercDescSRV_Col)))
    If lTamanho > 0 Then
        dPercentDesc = CDbl(Format(GridItensSRV.TextMatrix(iLinha, iGrid_PercDescSRV_Col), "General Number"))
    Else
        dPercentDesc = 0
    End If

    dPrecoTotal = dPrecoUnitario * (dQuantidade - dQuantidadeCancelada)

    'Se percentual for >0 tira o desconto
    If dPercentDesc > 0 Then dDesconto = StrParaDbl(Format(dPercentDesc * dPrecoTotal, "Standard")) 'Alterado por Wagner
    dPrecoTotalReal = dPrecoTotal - dDesconto

    'Coloca valor do desconto no Grid
    If dDesconto > 0 Then
        GridItensSRV.TextMatrix(iLinha, iGrid_DescontoSRV_Col) = Format(dDesconto, "Standard")
    Else
        GridItensSRV.TextMatrix(iLinha, iGrid_DescontoSRV_Col) = ""
    End If

    'Coloca preco total do ítem no grid
    GridItensSRV.TextMatrix(iLinha, iGrid_PrecoTotalSRV_Col) = Format(dPrecoTotalReal, "Standard")
    GridItensSRV.TextMatrix(iLinha, iGrid_PrecoTotalSRVB_Col) = Format(dPrecoTotal, "Standard")
'
'    'Calcula a soma dos valores de produtos
'    For iIndice = 1 To objGridItensSRV.iLinhasExistentes
'        If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
'            If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
'        End If
'    Next
'
'    For iIndice = 1 To objGridItens.iLinhasExistentes
'        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
'            If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'        End If
'    Next
'
'    dValorDesconto = gobjPedidoVenda.dValorDesconto
'    dValorTotal = dValorTotal - dValorDesconto
'    '############################################
'
'    'Verifica se o valor de desconto é maior que o valor dos produtos
'    If dValorTotal < 0 And dValorDesconto > 0 Then
'
'        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", dValorDesconto, dValorTotal)
'
'        gdDesconto = 0
'        ValorDesconto.Text = ""
'        giValorDescontoAlterado = 0
'        dValorDesconto = 0
'
'        'Para tributação
'        gobjPedidoVenda.dValorDesconto = dValorDesconto
'
''        'Faz a atualização dos valores da tributação
''        lErro = AtualizarTributacao()
''        If lErro <> SUCESSO Then gError 191157
'
'        Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, dValorDesconto)
'
'        'Faz a atualização dos valores da tributação
'        lErro = gobjTribTab.AtualizarTributacao()
'        If lErro <> SUCESSO Then gError 191157
'
'        'Calcula a soma dos valores de produtos
'        dValorTotal = 0
'
'        For iIndice = 1 To objGridItens.iLinhasExistentes
'            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
'                If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'            End If
'        Next
'
'        For iIndice = 1 To objGridItensSRV.iLinhasExistentes
'            If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
'                If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
'            End If
'        Next
'
'    End If
'
'    'Coloca valor total dos produtos na tela
'    ValorProdutos.Caption = Format(dValorTotal, "Standard")
    
    Call SubTotal_Calcula

    'Altera ítem na coleção de ítens de gobjPedidoVenda que dá suporte p/ Tributação
'    Call Tributacao_Alteracao_Item_GridSRV(iLinha)

    PSRV_PrecoTotal_CalculaSRV = SUCESSO

    Exit Function

Erro_PSRV_PrecoTotal_CalculaSRV:

    PSRV_PrecoTotal_CalculaSRV = gErr

    Select Case gErr

        Case 191157

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191158)

    End Select

    Exit Function

End Function

Function ProdutoSRV_Saida_Celula(ByVal objCT As Object) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim iIndice As Integer
Dim sProduto As String
Dim vbMsgRes As VbMsgBoxResult
Dim iAlmoxarifadoPadrao As Integer
Dim objItemPV As ClassItemPedido
Dim iPossuiGrade As Integer
Dim objRomaneioGrade As ClassRomaneioGrade
Dim bPossuiFilhos As Boolean
Dim colItensRomaneioGrade As New Collection
Dim objItensRomaneio As ClassItemRomaneioGrade
Dim sProdutoPai As String
Dim objPVAux As New ClassPedidoDeVenda
Dim objKit As New ClassKit 'Inserido por Wagner
Dim sCodigoAntigo As String, dtDataAux As Date
Dim bPreenchePreco As Boolean 'Inserido por Wagner 27/04/2005
Dim lErro1 As Long

On Error GoTo Erro_ProdutoSRV_Saida_Celula

    bPreenchePreco = False 'Inserido por Wagner 27/04/2005

'***********  FUNÇÃO ALTERADA PARA TRATAMENTO DE GRADE  ******************
    'Critica o Produto
    lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 191163
    'Se o produto é gerencial sem ser de grade ==> Erro
    If lErro = 86295 Then
        gError 191164
    End If

    'Se o produto não foi encontrado ==> Pergunta se deseja criar
    If lErro = 51381 Then gError 191165

    If objProduto.iNatureza = NATUREZA_PROD_SERVICO Then gError 191180

    lErro = CF2(Me, "PSRV_Produto_Saida_Celula3", objProduto)
    If lErro <> SUCESSO Then gError 191166

    'Se não for um produto de grade
    If objProduto.iKitVendaComp <> MARCADO Then

        'Se o produto controlar estoque
        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then

            lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifadoPadrao)
            If lErro <> SUCESSO And lErro <> 23796 Then gError 191167
            
            If lErro <> SUCESSO Or iAlmoxarifadoPadrao = 0 Then gError 191168

        End If
    Else

        lErro = CF("ProdutoKitVenda_Le_Romaneio", objProduto, colItensRomaneioGrade)
        If lErro <> SUCESSO Then gError 191169

        'Se nao existir, erro
        If colItensRomaneioGrade.Count = 0 Then gError 191170

        'Para cada filho de grade do produto
        For Each objItensRomaneio In colItensRomaneioGrade
            'Verifica se ele já aparece no grid
            For iIndice = 1 To gobjPedidoVenda.colItensPedido.Count
                'Se aparecer ==> Erro
                If gobjPedidoVenda.colItensPedido(iIndice).sProduto = objItensRomaneio.sProduto Then gError 191171
            Next
        Next

    End If

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 191172

        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True

    End If

    '#####################################################
    'Inserido por Wagner 27/04/2005
    lErro = CF2(Me, "PSRV_Produto_Saida_Celula2")
    If lErro <> SUCESSO Then gError 191173
    '#####################################################

    'Verifica se é de Faturamento
    If objProduto.iFaturamento = 0 Then gError 191174

    Set objItemPV = New ClassItemPedido

    objItemPV.sProduto = objProduto.sCodigo
    objItemPV.sUnidadeMed = objProduto.sSiglaUMVenda
    objItemPV.lCodPedido = StrParaDbl(Codigo.Text)
    objItemPV.iItem = GridItens.Row
    objItemPV.lNumIntDoc = 0
    objItemPV.sDescricao = objProduto.sDescricao

    'Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda

    'Descricao Produto
    If Me.ActiveControl Is DescricaoProduto Then
        DescricaoProduto.Text = objProduto.sDescricao
    Else
        GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objProduto.sDescricao
    End If

'    'Data de Entrega
'    dtDataAux = MaskedParaDate(DataEntregaPV)
'    If dtDataAux <> DATA_NULA Then GridItens.TextMatrix(GridItens.Row, iGrid_DataEntrega_Col) = Format(dtDataAux, "dd/mm/yyyy")

    'Preço Unitário
    If Len(Trim(TabelaPreco.Text)) > 0 Then

        'Coloca Produto no grid (necessario p/usar Trata_TabelaPreco_Item)
        GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text

        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
        objTabelaPrecoItem.sCodProduto = objProduto.sCodigo

        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 191175

        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA

            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 191176

        End If

        If lErro <> 28014 Then

            bPreenchePreco = True 'Inserido por Wagner 27/04/2005

            lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 191177

            If dPrecoUnitario > 0 Then
                GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            Else
                GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = ""
            End If

        End If

    End If

    'Acrescenta uma linha no Grid se for o caso
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1

        gcolItemPedido.Add 0

        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItens.Row)

        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 191178

'        If objProduto.iKitVendaComp = MARCADO Then
'
'            '************** GRADE ************
'            gobjPedidoVenda.colItensPedido(GridItens.Row).iPossuiGrade = MARCADO
'            gobjPedidoVenda.colItensPedido(GridItens.Row).sUMEstoque = objProduto.sSiglaUMEstoque
'            gobjPedidoVenda.colItensPedido(GridItens.Row).iItem = GridItens.Row
'            gobjPedidoVenda.colItensPedido(GridItens.Row).sProduto = objProduto.sCodigo
'
'             If objProduto.iKitVendaComp = MARCADO Then
'                Set objItemPV.colItensRomaneioGrade = colItensRomaneioGrade
'            End If
'
'            Set gobjPedidoVenda.colItensPedido(GridItens.Row).colItensRomaneioGrade = objItemPV.colItensRomaneioGrade
'
'            Call Atualiza_Grid_Alocacao(objItemPV)
'
'        End If

    End If

    Call Grid_Refresh_Checkbox(objGridItens)

    ProdutoSRV_Saida_Celula = SUCESSO

    Exit Function

Erro_ProdutoSRV_Saida_Celula:

    ProdutoSRV_Saida_Celula = gErr

    Select Case gErr

        Case 191163, 191166, 191167, 191169, 191173, 191175, 191176, 191177, 191178

        Case 191164
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case 191165
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
            End If

        Case 191168
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO", gErr, objProduto.sCodigo)

        Case 191170
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, Produto.Text)

        Case 191171
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILHO_GRADE_GRID", gErr, Trim(objProduto.sCodigo), Trim(gobjPedidoVenda.colItensPedido(iIndice).sProduto))

        Case 191172
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, Produto.Text)

        Case 191174
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO", gErr)

        Case 191180
            Call Rotina_Erro(vbOKOnly, "ERRO_PECA_NAO_PODE_TER_NATUREZA_SERVICO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191179)

    End Select

    Exit Function

End Function

Function ServicoSRV_Saida_Celula() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim iIndice As Integer
Dim sProduto As String
Dim vbMsgRes As VbMsgBoxResult
Dim iAlmoxarifadoPadrao As Integer
Dim objItemPV As ClassItemPedido
Dim iPossuiGrade As Integer
Dim objRomaneioGrade As ClassRomaneioGrade
Dim bPossuiFilhos As Boolean
Dim colItensRomaneioGrade As New Collection
Dim objItensRomaneio As ClassItemRomaneioGrade
Dim sProdutoPai As String
Dim objPVAux As New ClassPedidoDeVenda
Dim objKit As New ClassKit 'Inserido por Wagner
Dim sCodigoAntigo As String, dtDataAux As Date
Dim bPreenchePreco As Boolean 'Inserido por Wagner 27/04/2005
Dim lErro1 As Long

On Error GoTo Erro_ServicoSRV_Saida_Celula

    bPreenchePreco = False 'Inserido por Wagner 27/04/2005

'***********  FUNÇÃO ALTERADA PARA TRATAMENTO DE GRADE  ******************
    'Critica o Produto
    lErro = CF("Produto_Critica_Filial2", ServicoSRV.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 191182
    'Se o produto é gerencial sem ser de grade ==> Erro
    If lErro = 86295 Then
        gError 191183
    End If

    'Se o produto não foi encontrado ==> Pergunta se deseja criar
    If lErro = 51381 Then gError 191184

    If objProduto.iNatureza <> NATUREZA_PROD_SERVICO Then gError 191185

    lErro = CF2(Me, "PSRV_Produto_Saida_Celula3SRV", objProduto)
    If lErro <> SUCESSO Then gError 191186

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 191187

        ServicoSRV.PromptInclude = False
        ServicoSRV.Text = sProduto
        ServicoSRV.PromptInclude = True

    End If

    lErro = CF2(Me, "PSRV_Produto_Saida_Celula2SRV")
    If lErro <> SUCESSO Then gError 191188

    'Verifica se é de Faturamento
    If objProduto.iFaturamento = 0 Then gError 191189

    Set objItemPV = New ClassItemPedido

    objItemPV.sProduto = objProduto.sCodigo
    objItemPV.sUnidadeMed = objProduto.sSiglaUMVenda
    objItemPV.lCodPedido = StrParaDbl(Codigo.Text)
    objItemPV.iItem = GridItensSRV.Row
    objItemPV.lNumIntDoc = 0
    objItemPV.sDescricao = objProduto.sDescricao

    'Unidade de Medida
    GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_UnidadeMedSRV_Col) = objProduto.sSiglaUMVenda

    'Descricao Produto
    If Me.ActiveControl Is DescServicoSRV Then
        DescServicoSRV.Text = objProduto.sDescricao
    Else
        GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_DescServicoSRV_Col) = objProduto.sDescricao
    End If

    'Data de Entrega
    dtDataAux = MaskedParaDate(DataEntregaPV)
    If dtDataAux <> DATA_NULA Then GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_DataEntregaSRV_Col) = Format(dtDataAux, "dd/mm/yyyy")

    'Preço Unitário
    If Len(Trim(TabelaPreco.Text)) > 0 Then

        'Coloca Produto no grid (necessario p/usar Trata_TabelaPreco_Item)
        GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ServicoSRV_Col) = ServicoSRV.Text

        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
        objTabelaPrecoItem.sCodProduto = objProduto.sCodigo

        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 191190

        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA

            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 191191

        End If

        If lErro <> 28014 Then

            bPreenchePreco = True 'Inserido por Wagner 27/04/2005

            lErro = PrecoUnitario_Calcula(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_UnidadeMedSRV_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 191192

            If dPrecoUnitario > 0 Then
                GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ValorUnitarioSRV_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            Else
                GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ValorUnitarioSRV_Col) = ""
            End If

        End If

    End If

    'Acrescenta uma linha no Grid se for o caso
    If GridItensSRV.Row - GridItensSRV.FixedRows = objGridItensSRV.iLinhasExistentes Then
        objGridItensSRV.iLinhasExistentes = objGridItensSRV.iLinhasExistentes + 1

        gcolItemPedidoSRV.Add 0

        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItensSRV.Row)

'        'permite que a tributacao reflita a inclusao de uma linha no grid
'        lErro = Tributacao_Inclusao_Item_GridSRV(GridItensSRV.Row, ServicoSRV.Text)
'        If lErro <> SUCESSO Then gError 191193

        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItensSRV.Row, objProduto.sCodigo, True)
        If lErro <> SUCESSO Then gError 191178

    End If

    Call Grid_Refresh_Checkbox(objGridItensSRV)

    ServicoSRV_Saida_Celula = SUCESSO

    Exit Function

Erro_ServicoSRV_Saida_Celula:

    ServicoSRV_Saida_Celula = gErr

    Select Case gErr

        Case 191182, 191186, 191188, 191190, 191191, 191192, 191193

        Case 191183
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case 191184
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", ServicoSRV.Text)
            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItensSRV)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItensSRV)
            End If

        Case 191185
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_NATUREZA_SERVICO", gErr, ServicoSRV.Text)

        Case 191187
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, ServicoSRV.Text)

        Case 191189
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191194)

    End Select

    Exit Function

End Function

Private Function Traz_Produto_Tela() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim iIndice As Integer
Dim sProduto As String, dtDataAux As Date
Dim bPreenchePreco As Boolean 'Inserido por Wagner 27/04/2005

On Error GoTo Erro_Traz_Produto_Tela

    bPreenchePreco = False 'Inserido por Wagner 27/04/2005

    '##################################################################
    'Alterado por Wagner 02/06/2006
    lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 191195
    
    If lErro = 51381 Then gError 191196

    If lErro = 86295 Then gError 191197

    If objProduto.iNatureza = NATUREZA_PROD_SERVICO Then gError 191198

    lErro = CF2(Me, "PSRV_Produto_Saida_Celula2")
    If lErro <> SUCESSO Then gError 191199

    'Verifica se é de Faturamento
    If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 191200

    'Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda

    'Descricao Produto
    GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objProduto.sDescricao

    'Preço Unitário
    If Len(Trim(TabelaPreco.Text)) > 0 Then

        'Coloca Produto no grid (necessario p/usar Trata_TabelaPreco_Item)
        GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text

        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
        objTabelaPrecoItem.sCodProduto = objProduto.sCodigo

        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 191201

        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA

            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 191202

        End If

        If lErro <> 28014 Then

            bPreenchePreco = True 'Inserido por Wagner 27/04/2005

            lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 191203

            If dPrecoUnitario > 0 Then
                GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            Else
                GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = ""
            End If

        End If

    End If


    'Acrescenta uma linha no Grid se for o caso
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1

        gcolItemPedido.Add 0
        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItens.Row)

        'permite que a tributacao reflita a inclusao de uma linha no grid
        'lErro = Tributacao_Inclusao_Item_Grid(GridItens.Row, Produto.Text)
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 191204

    End If

    Call Grid_Refresh_Checkbox(objGridItens)

    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = gErr

    Select Case gErr

        Case 191195, 191199, 191201 To 191204

        Case 191196
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)

        Case 191197
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case 191198
            Call Rotina_Erro(vbOKOnly, "ERRO_PECA_NAO_PODE_TER_NATUREZA_SERVICO", gErr, objProduto.sCodigo)

        Case 191200
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191205)

    End Select

    Exit Function

End Function

'Private Function Traz_Servico_Tela() As Long
'
'Dim lErro As Long
'Dim objProduto As New ClassProduto
'Dim iProdutoPreenchido As Integer
'Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
'Dim dPrecoUnitario As Double
'Dim iIndice As Integer
'Dim sProduto As String, dtDataAux As Date
'Dim bPreenchePreco As Boolean 'Inserido por Wagner 27/04/2005
'
'On Error GoTo Erro_Traz_Servico_Tela
'
'    bPreenchePreco = False 'Inserido por Wagner 27/04/2005
'
'    lErro = CF("Produto_Critica_Filial2", ServicoSRV.Text, objProduto, iProdutoPreenchido)
'    If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 191206
'
'    If lErro = 51381 Then gError 191207
'
'    If lErro = 86295 Then gError 191208
'
'    If objProduto.iNatureza <> NATUREZA_PROD_SERVICO Then gError 191209
'
'    lErro = CF2(Me, "PSRV_Produto_Saida_Celula2SRV")
'    If lErro <> SUCESSO Then gError 191210
'
'    'Verifica se é de Faturamento
'    If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 191211
'
'    'Unidade de Medida
'    GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_UnidadeMedSRV_Col) = objProduto.sSiglaUMVenda
'
'    'Descricao Produto
'    GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_DescServicoSRV_Col) = objProduto.sDescricao
'
'    'Data de Entrega
'    dtDataAux = MaskedParaDate(DataEntregaPV)
'    If dtDataAux <> DATA_NULA Then GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_DataEntregaSRV_Col) = Format(dtDataAux, "dd/mm/yyyy")
'
'    'Preço Unitário
'    If Len(Trim(TabelaPreco.Text)) > 0 Then
'
'        'Coloca Produto no grid (necessario p/usar Trata_TabelaPreco_Item)
'        GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ServicoSRV_Col) = ServicoSRV.Text
'
'        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
'        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
'        objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
'
'        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
'        If lErro <> SUCESSO And lErro <> 28014 Then gError 191212
'
'        If lErro = 28014 Then
'            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
'
'            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
'            If lErro <> SUCESSO And lErro <> 28014 Then gError 191213
'
'        End If
'
'        If lErro <> 28014 Then
'
'            bPreenchePreco = True 'Inserido por Wagner 27/04/2005
'
'            lErro = PrecoUnitario_Calcula(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_UnidadeMedSRV_Col), objTabelaPrecoItem, dPrecoUnitario)
'            If lErro <> SUCESSO Then gError 191214
'
'            If dPrecoUnitario > 0 Then
'                GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ValorUnitarioSRV_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
'            Else
'                GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ValorUnitarioSRV_Col) = ""
'            End If
'
'        End If
'
'    End If
'
'
'    'Acrescenta uma linha no Grid se for o caso
'    If GridItensSRV.Row - GridItensSRV.FixedRows = objGridItensSRV.iLinhasExistentes Then
'        objGridItensSRV.iLinhasExistentes = objGridItensSRV.iLinhasExistentes + 1
'
'        gcolItemPedidoSRV.Add 0
'
'        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItensSRV.Row)
'
'        'permite que a tributacao reflita a inclusao de uma linha no grid
'        lErro = Tributacao_Inclusao_Item_GridSRV(GridItensSRV.Row, ServicoSRV.Text)
'        If lErro <> SUCESSO Then gError 191215
'
'    End If
'
'    Call Grid_Refresh_Checkbox(objGridItensSRV)
'
'    Traz_Servico_Tela = SUCESSO
'
'    Exit Function
'
'Erro_Traz_Servico_Tela:
'
'    Traz_Servico_Tela = gErr
'
'    Select Case gErr
'
'        Case 191206, 191210, 191212 To 191215
'
'        Case 191207
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, ServicoSRV.Text)
'
'        Case 191208
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
'
'        Case 191209
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_NATUREZA_SERVICO", gErr, ServicoSRV.Text)
'
'        Case 191211
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191216)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Sub ValorDesconto_Change()

    iAlterado = REGISTRO_ALTERADO
    giValorDescontoAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorDesconto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorDesconto As Double
Dim dValorProdutos As Double
Dim iIndice As Integer

On Error GoTo Erro_ValorDesconto_Validate

    'Verifica se o valor foi alterado
    If giValorDescontoAlterado = 0 Then Exit Sub

    'Vale o desconto que foi colocado aqui
    giValorDescontoManual = 1

    dValorDesconto = 0

    For iIndice = 1 To objGridItensSRV.iLinhasExistentes
        If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))) > 0 Then
            If CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col)) > 0 Then dValorProdutos = dValorProdutos + CDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
        End If
    Next

    'Calcula a soma dos valores de produtos
    For iIndice = 1 To objGridItens.iLinhasExistentes
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
            If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorProdutos = dValorProdutos + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        End If
    Next

    'Verifica se o Valor está preenchido
    If Len(Trim(ValorDesconto.Text)) > 0 Then

        'Faz a Crítica do Valor digitado
        lErro = Valor_NaoNegativo_Critica(ValorDesconto.Text)
        If lErro <> SUCESSO Then gError 191217

        dValorDesconto = CDbl(ValorDesconto.Text)

        'Coloca o Valor formatado na tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")

        If dValorDesconto > dValorProdutos Then gError 191218

        dValorProdutos = dValorProdutos - dValorDesconto

    End If

    ValorProdutos.Caption = Format(dValorProdutos, "Standard")

    'Para tributação
    gobjPedidoVenda.dValorDesconto = dValorDesconto

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDesconto_Validate(Cancel, dValorDesconto)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 191219

    giValorDescontoAlterado = 0
    
    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True


    Select Case gErr

        Case 191217, 191219

        Case 191218
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191220)

    End Select

    Exit Sub


End Sub

Public Sub ValorDespesas_Change()

     giValorDespesasAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorDespesas_Validate(Cancel As Boolean)

Dim dValorDespesas As Double
Dim lErro As Long

On Error GoTo Erro_ValorDespesas_Validate

    If giValorDespesasAlterado = 0 Then Exit Sub

    'Se  estiver preenchido
    If Len(Trim(ValorDespesas.Text)) > 0 Then

        'Faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorDespesas.Text)
        If lErro <> SUCESSO Then gError 191221

        dValorDespesas = CDbl(ValorDespesas.Text)

        'coloca o valor formatado na tela
        ValorDespesas.Text = Format(dValorDespesas, "Standard")

    End If

'    'Para tributação
'    gobjPedidoVenda.dValorOutrasDespesas = dValorDespesas

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDespesas_Validate(Cancel, dValorDespesas)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 191222

    giValorDespesasAlterado = 0

    Exit Sub

Erro_ValorDespesas_Validate:

    Cancel = True

    Select Case gErr

        Case 191221, 191222

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191223)

    End Select

    Exit Sub


End Sub

Public Sub ValorFrete_Change()

    iAlterado = REGISTRO_ALTERADO
    giValorFreteAlterado = 1

End Sub

Public Sub ValorFrete_Validate(Cancel As Boolean)

Dim dValorFrete As Double
Dim lErro As Long

On Error GoTo Erro_ValorFrete_Validate

    If giValorFreteAlterado = 0 Then Exit Sub

    dValorFrete = 0

    If Len(Trim(ValorFrete.Text)) > 0 Then

        'Critica se valor é não negativo
        lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
        If lErro <> SUCESSO Then gError 191224

        dValorFrete = CDbl(ValorFrete.Text)

        ValorFrete.Text = Format(dValorFrete, "Standard")

    End If

'    'Para tributação
'    gobjPedidoVenda.dValorFrete = dValorFrete

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorFrete_Validate(Cancel, dValorFrete)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 191225

    giValorFreteAlterado = 0

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True

    Select Case gErr

        Case 191224, 191225

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191226)

    End Select

    Exit Sub

End Sub

'Public Sub ValorIRRF_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giValorIRRFAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ValorIRRF_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValorIRRF As Double
'Dim dValorTotal As Double
'
'On Error GoTo Erro_ValorIRRF_Validate
'
'    If giValorIRRFAlterado = 0 Then Exit Sub
'
'    'Verifica se ValorIRRF foi preenchido
'    If Len(Trim(ValorIRRF.Text)) > 0 Then
'
'        'Critica o Valor
'        lErro = Valor_NaoNegativo_Critica(ValorIRRF.Text)
'        If lErro <> SUCESSO Then gError 191227
'
'        dValorIRRF = CDbl(ValorIRRF.Text)
'
'        ValorIRRF.Text = Format(dValorIRRF, "Standard")
'
'        If Len(Trim(ValorTotal.Caption)) > 0 Then dValorTotal = CDbl(ValorTotal.Caption)
'
'        If dValorIRRF > dValorTotal Then gError 191228
'
'    End If
'
'    Call BotaoGravarTrib
'
'    giValorIRRFAlterado = 0
'
'    Exit Sub
'
'Erro_ValorIRRF_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 191227
'
'        Case 191228
'            Call Rotina_Erro(vbOKOnly, "ERRO_IR_FONTE_MAIOR_VALOR_TOTAL", gErr, dValorIRRF, dValorTotal)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191229)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub ValorProdutos_Change()

'    If Not gbCarregandoTela Then
'        Call DescontoGlobal_Recalcula
'    End If

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
'        If gdDesconto > 0 And giValorDescontoManual = 0 Then
'
'            Call Calcula_ValorProdutos(dValorProdutos)
'
'            'Calcula o valor do desconto para o cliente e coloca na tela
'            dValorDesconto = gdDesconto * dValorProdutos
'            ValorDesconto.Text = Format(dValorDesconto, "Standard")
'            giValorDescontoAlterado = 0
'
'            ValorProdutos.Caption = Format(dValorProdutos - dValorDesconto, "STANDARD")
'
'            'Para tributação
'            gobjPedidoVenda.dValorDesconto = dValorDesconto
'
'        End If
'
'    End If

    Exit Sub

Erro_DescontoGlobal_Recalcula:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191230)

    End Select

    Exit Sub

End Sub

Public Sub ValorSeguro_Change()

    iAlterado = REGISTRO_ALTERADO
    giValorSeguroAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorSeguro_Validate(Cancel As Boolean)

Dim dValorSeguro As Double
Dim lErro As Long

On Error GoTo Erro_Valorseguro_Validate

    If giValorSeguroAlterado = 0 Then Exit Sub

    dValorSeguro = 0

    If Len(Trim(ValorSeguro.Text)) > 0 Then

        'Critica se valor é não negativo
        lErro = Valor_NaoNegativo_Critica(ValorSeguro.Text)
        If lErro <> SUCESSO Then gError 191231

        dValorSeguro = CDbl(ValorSeguro.Text)

        ValorSeguro.Text = Format(dValorSeguro, "Standard")

    End If

'    'Para tributação
'    gobjPedidoVenda.dValorSeguro = dValorSeguro
    
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorSeguro_Validate(Cancel, dValorSeguro)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 191232

    giValorSeguroAlterado = 0

    Exit Sub

Erro_Valorseguro_Validate:

    Cancel = True

    Select Case gErr

        Case 191231, 191232

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191233)

    End Select

    Exit Sub

End Sub

Function ReservaAlmoxarifadoPadrao(iIndiceGridItens As Integer, objProduto As ClassProduto, dQuantidade As Double, dQuantCancelada As Double, dQuantFaturada As Double) As Long

Dim lErro As Long
Dim dQuantidadeReservarVenda As Double
Dim dQuantReservadaPedido As Double
Dim dFator As Double
Dim iAlmoxarifadoPadrao As Integer
Dim objItemPV As New ClassItemPedido
Dim objReservaBD As New ClassReserva
Dim objReserva As ClassReserva
Dim sProdutoEnxuto As String
Dim dPrecoUnitario As Double
Dim colItemPedido As New colItemPedido
Dim sProduto As String
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim objAlmoxarifadoPadrao As New ClassAlmoxarifado
Dim objAlmoxarifadoPadrao1 As New ClassAlmoxarifado
Dim dSaldo As Double
Dim colReserva As colReserva
Dim iFilialFaturamento As Integer
Dim iFilialEmpresa As Integer

Dim dQuantidadeReservarEstoque1 As Double 'Reservar Filial 1
Dim dQuantidadeReservarEstoque2 As Double 'Reservar Filial 2
Dim objEstoqueProduto1 As New ClassEstoqueProduto 'Reservar Filial 1
Dim objEstoqueProduto2 As New ClassEstoqueProduto 'Reservar Filial 2


On Error GoTo Erro_ReservaAlmoxarifadoPadrao
    'calula a quantidade reservada
    dQuantidadeReservarVenda = (dQuantidade - dQuantCancelada - dQuantFaturada)

    lErro = CF("UM_Conversao", objProduto.iClasseUM, GridItens.TextMatrix(iIndiceGridItens, iGrid_UnidadeMed_Col), objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 191234
    
    'Converte a quantidade para a unidade de Estoque
    dQuantidadeReservarEstoque1 = dQuantidadeReservarVenda * dFator

    If Len(Trim(FilialFaturamento.Text)) <> 0 Then
        iFilialFaturamento = Codigo_Extrai(FilialFaturamento.Text)
    End If

    If iFilialFaturamento = 0 Then
        'Busca a filial de faturamento
        lErro = CF("FilialFaturamento_Le", giFilialEmpresa, iFilialFaturamento)
        If lErro <> SUCESSO Then gError 191235
        
        If iFilialFaturamento = 0 Then iFilialFaturamento = giFilialEmpresa
    End If

    lErro = CF2(Me, "PSRV_AlocaAlmoxarifadoPadrao", iFilialFaturamento, objProduto.sCodigo, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO And lErro <> 193422 Then gError 191236

    objAlmoxarifadoPadrao.iCodigo = iAlmoxarifadoPadrao

    'Lê o almoxarifado padrão
    lErro = CF("Almoxarifado_Le", objAlmoxarifadoPadrao)
    If lErro = SUCESSO Then

        objEstoqueProduto1.sProduto = objProduto.sCodigo
        objEstoqueProduto1.iAlmoxarifado = iAlmoxarifadoPadrao
        objEstoqueProduto1.sAlmoxarifadoNomeReduzido = objAlmoxarifadoPadrao.sNomeReduzido

        'Lê o estoque do produto
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto1)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 191237
        
        If lErro = 21306 Then gError 191238

        If Len(Trim(Codigo.Text)) > 0 Then objItemPV.lCodPedido = CLng(Codigo.Text)
        objItemPV.sProduto = objProduto.sCodigo
        objItemPV.iFilialEmpresa = giFilialEmpresa
        objItemPV.sProdutoDescricao = objProduto.sDescricao
        objReservaBD.iAlmoxarifado = iAlmoxarifadoPadrao

        'Lê as reservas do produto
        lErro = CF("ReservaItem_Le", objItemPV, objReservaBD)
        If lErro <> SUCESSO And lErro <> 26678 Then gError 191239

        dQuantReservadaPedido = objReservaBD.dQuantidade

        'Saldo enxergado por esse Pedido Venda
        objEstoqueProduto1.dSaldo = objEstoqueProduto1.dQuantDisponivel + dQuantReservadaPedido

        'Se quantidade a reservar for menor que a disponível coloca reserva no GridAlocacao
        If (dQuantidadeReservarEstoque1 - objEstoqueProduto1.dSaldo) < QTDE_ESTOQUE_DELTA Then

            GridItens.TextMatrix(iIndiceGridItens, iGrid_QuantReservada_Col) = Formata_Estoque(dQuantidadeReservarVenda)
            Call Remove_Reservas(iIndiceGridItens)

            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Item_Col) = iIndiceGridItens
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_ProdutoAlmox_Col) = GridItens.TextMatrix(iIndiceGridItens, iGrid_Produto_Col)
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_UMEstoque_Col) = objProduto.sSiglaUMEstoque
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objEstoqueProduto1.sAlmoxarifadoNomeReduzido
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReservar_Col) = Formata_Estoque(dQuantidadeReservarVenda)
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReserv_Col) = Formata_Estoque(dQuantidadeReservarEstoque1)
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Responsavel_Col) = RESERVA_AUTO_RESP
            objGridAlocacao.iLinhasExistentes = objGridAlocacao.iLinhasExistentes + 1

        Else

            lErro = CF("Retorna_Almoxarifado_Alternativo", iAlmoxarifadoPadrao, iFilialEmpresa)
            If lErro <> SUCESSO Then gError 191240

            If iAlmoxarifadoPadrao = 0 Then GoTo FaltaEstoque

            'Verifica quanto falta para reservar
            dQuantidadeReservarEstoque2 = dQuantidadeReservarEstoque1 - objEstoqueProduto1.dSaldo
            dQuantidadeReservarEstoque1 = objEstoqueProduto1.dSaldo

            'Busca na Filial 2
            objAlmoxarifadoPadrao1.iCodigo = iAlmoxarifadoPadrao

            'Lê o almoxarifado
            lErro = CF("Almoxarifado_Le", objAlmoxarifadoPadrao1)
            If lErro = SUCESSO Then

                objEstoqueProduto2.sProduto = objProduto.sCodigo
                objEstoqueProduto2.iAlmoxarifado = iAlmoxarifadoPadrao
                objEstoqueProduto2.sAlmoxarifadoNomeReduzido = objAlmoxarifadoPadrao1.sNomeReduzido

                'Lê o estoque do produto
                lErro = CF("EstoqueProduto_Le", objEstoqueProduto2)
                If lErro <> SUCESSO And lErro <> 21306 Then gError 191241
                
                If lErro = 21306 Then gError 191242

                If Len(Trim(Codigo.Text)) > 0 Then objItemPV.lCodPedido = CLng(Codigo.Text)
                objItemPV.sProduto = objProduto.sCodigo
                objItemPV.iFilialEmpresa = iFilialEmpresa
                objItemPV.sProdutoDescricao = objProduto.sDescricao
                objReservaBD.iAlmoxarifado = iAlmoxarifadoPadrao

                'Lê as reservas do produto
                lErro = CF("ReservaItem_Le", objItemPV, objReservaBD)
                If lErro <> SUCESSO And lErro <> 26678 Then gError 191243

                dQuantReservadaPedido = objReservaBD.dQuantidade

                'Saldo enxergado por esse Pedido Venda
                objEstoqueProduto2.dSaldo = objEstoqueProduto2.dQuantDisponivel + dQuantReservadaPedido

                'Se quantidade a reservar for menor que a disponível coloca reserva no GridAlocacao
                If (dQuantidadeReservarEstoque2 - objEstoqueProduto2.dSaldo) < QTDE_ESTOQUE_DELTA Then

                    GridItens.TextMatrix(iIndiceGridItens, iGrid_QuantReservada_Col) = Formata_Estoque(dQuantidadeReservarVenda)
                    Call Remove_Reservas(iIndiceGridItens)

                    If dQuantidadeReservarEstoque1 > 0 Then
                        'Parte reservada na filial 1
                        GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Item_Col) = iIndiceGridItens
                        GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_ProdutoAlmox_Col) = GridItens.TextMatrix(iIndiceGridItens, iGrid_Produto_Col)
                        GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_UMEstoque_Col) = objProduto.sSiglaUMEstoque
                        GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objEstoqueProduto1.sAlmoxarifadoNomeReduzido
                        GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReservar_Col) = Formata_Estoque(dQuantidadeReservarVenda)
                        GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReserv_Col) = Formata_Estoque(dQuantidadeReservarEstoque1)
                        GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Responsavel_Col) = RESERVA_AUTO_RESP
                        objGridAlocacao.iLinhasExistentes = objGridAlocacao.iLinhasExistentes + 1
                    End If

                    'Parte Reservada na Filial 2
                    GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Item_Col) = iIndiceGridItens
                    GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_ProdutoAlmox_Col) = GridItens.TextMatrix(iIndiceGridItens, iGrid_Produto_Col)
                    GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_UMEstoque_Col) = objProduto.sSiglaUMEstoque
                    GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objEstoqueProduto2.sAlmoxarifadoNomeReduzido
                    GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReservar_Col) = Formata_Estoque(dQuantidadeReservarVenda)
                    GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReserv_Col) = Formata_Estoque(dQuantidadeReservarEstoque2)
                    GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Responsavel_Col) = RESERVA_AUTO_RESP
                    objGridAlocacao.iLinhasExistentes = objGridAlocacao.iLinhasExistentes + 1

                'Caso contrário limpa as reservas desse ítem e chama tela de Falta de Estoque
                Else

FaltaEstoque:
                    Set objItemPV = New ClassItemPedido
                    objItemPV.dQuantidade = dQuantidade
                    objItemPV.dQuantReservada = 0
                    objItemPV.dQuantCancelada = dQuantCancelada
                    objItemPV.dQuantFaturada = dQuantFaturada
                    objItemPV.sProduto = objProduto.sCodigo
                    objItemPV.lCodPedido = StrParaLong(Codigo.Text)
                    objItemPV.iItem = iIndiceGridItens
                    objItemPV.lNumIntDoc = gcolItemPedido(objItemPV.iItem)
                    objItemPV.sUMEstoque = objProduto.sSiglaUMEstoque
                    objItemPV.sProdutoDescricao = objProduto.sDescricao
                    objItemPV.sUnidadeMed = GridItens.TextMatrix(iIndiceGridItens, iGrid_UnidadeMed_Col)
                    objItemPV.iClasseUM = objProduto.iClasseUM


                    'Armazena os nomes dos Produtos do Pedido para não fazer substituição com um que já tem
                    For iIndice = 1 To objGridItens.iLinhasExistentes

                        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
                        If lErro <> SUCESSO Then gError 191244

                        colItemPedido.Add 0, 0, 0, sProduto, 0, 0, 0, 0, DATA_NULA, "", 0, 0, 0, colReserva, "", "", "", "", 0, 0, 0, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0

                    Next

                    'Chama tela de Falta de Estoque
                    lErro = Chama_Tela_Modal("FaltaEstoque", objItemPV, colItemPedido, dQuantidadeReservarEstoque1, objAlmoxarifadoPadrao, objEstoqueProduto1.dSaldo)


                    'Se retornar Cancela erro
                    If giRetornoTela = vbCancel Then gError 191245

                    'Limpa reservas desse ítem no GridAlocacao
                    Call Remove_Reservas(iIndiceGridItens)
                    'Se retornar OK,

                    If giRetornoTela = vbOK Then

                        'Se não substituiu o Produto
                        If objItemPV.sProduto = objProduto.sCodigo Then

                            'Coloca QuantReservada e QuantCancelada no ítem do GridItens
                            If objItemPV.dQuantReservada > 0 Then
                                GridItens.TextMatrix(iIndiceGridItens, iGrid_QuantReservada_Col) = Formata_Estoque(objItemPV.dQuantReservada)
                            Else
                                GridItens.TextMatrix(iIndiceGridItens, iGrid_QuantReservada_Col) = ""
                            End If
                            If objItemPV.dQuantCancelada > 0 Then
                                GridItens.TextMatrix(iIndiceGridItens, iGrid_QuantCancel_Col) = Formata_Estoque(objItemPV.dQuantCancelada)
                            Else
                                GridItens.TextMatrix(iIndiceGridItens, iGrid_QuantCancel_Col) = ""
                            End If

                            'Coloca reseravas desse ítem no GridAlocacao
                            For Each objReserva In objItemPV.colReserva

                                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Item_Col) = iIndiceGridItens
                                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_ProdutoAlmox_Col) = GridItens.TextMatrix(iIndiceGridItens, iGrid_Produto_Col)
                                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_UMEstoque_Col) = objProduto.sSiglaUMEstoque
                                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objReserva.sAlmoxarifado
                                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReservar_Col) = Formata_Estoque((dQuantidadeReservarVenda - objItemPV.dQuantCancelada) * dFator)
                                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReserv_Col) = Formata_Estoque(objReserva.dQuantidade)
                                If objReserva.dtDataValidade <> DATA_NULA Then GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Validade_Col) = Format(objReserva.dtDataValidade, "dd/mm/yyyy")
                                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Responsavel_Col) = objReserva.sResponsavel
                                objGridAlocacao.iLinhasExistentes = objGridAlocacao.iLinhasExistentes + 1
                            Next

                        'Se Produto foi substituído,
                        Else

                            lErro = Mascara_RetornaProdutoEnxuto(objItemPV.sProduto, sProdutoEnxuto)
                            If lErro <> SUCESSO Then gError 191246

                            Produto.PromptInclude = False
                            Produto.Text = sProdutoEnxuto
                            Produto.PromptInclude = True

                            'Substitui Produto no GridItens
                            GridItens.TextMatrix(iIndiceGridItens, iGrid_Produto_Col) = Produto.Text

                            'Chama função que traz os dados defaults do Produto
                            lErro = Traz_Produto_Tela()
                            gError 191247  'Para forçar saída da função e segurar o foco

                        End If
                    End If
                End If
            End If
        End If
    End If

    ReservaAlmoxarifadoPadrao = SUCESSO

    Exit Function

Erro_ReservaAlmoxarifadoPadrao:

    ReservaAlmoxarifadoPadrao = gErr

    Select Case gErr

        Case 191234, 191235, 191236, 191237, 191239, 191240, 191241, 191243, 191244, 191245, 191247

        Case 191238, 191242
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto1.sProduto, objEstoqueProduto1.iAlmoxarifado)

        Case 191245
            Call Rotina_Erro(vbOKOnly, "ERRO_RESERVA_NAO_DECIDIDA", gErr, objProduto.sCodigo)

        Case 191246
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191248)

    End Select

    Exit Function

End Function

Private Sub Remove_Reservas(iItemAtual As Integer)

Dim iIndice As Integer

    iIndice = 1
    'Exclui do grid de reservas as reservas do item passado
    Do While iIndice <= objGridAlocacao.iLinhasExistentes

        If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))) > 0 Then
            If CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) = iItemAtual Then
                Do While CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) = iItemAtual
                    Call Grid_Exclui_Linha(objGridAlocacao, iIndice)
                    If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))) = 0 Then Exit Do
                Loop
            End If
        End If
        iIndice = iIndice + 1
    Loop

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    Set gobjInfoUsu = Nothing

    Set gobjTelaProjetoInfo = Nothing
    Set gobjTelaTrataBloqGen = Nothing

    'variaveis auxiliares à contabilizacao
    Set gobjContabAutomatica = Nothing
    Set gobjNFiscalCTB = Nothing
    Set gobjPedidoVendaCTB = Nothing
    Set gcolAlmoxFilial = Nothing
    Set gobjGeracaoNFiscal = Nothing
    
    Set gobjInfoAdic = Nothing

    'Libera as variáveis globais
    'giGridBloqueioLinhasBloqueadas = 0
    Set objEventoCliente = Nothing
    Set objEventoNumero = Nothing
    Set objEventoCondPagto = Nothing
    Set objEventoProduto = Nothing
    Set objEventoServico = Nothing
    Set objEventoEstoqueProd = Nothing
    Set objEventoNaturezaOp = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoMensagem = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoTranspRedespacho = Nothing
    Set objEventoContato = Nothing
    Set objEventoCodSolicSRV = Nothing

    Set objGridItens = Nothing
    Set objGridComissoes = Nothing
    Set objGridParcelas = Nothing
    'Set objGridBloqueio = Nothing
    Set objGridAlocacao = Nothing

    Set gcolItemPedido = Nothing

    '??? jon 31/08/12
    Set gcolItemPedidoSRV = Nothing
    Set objGridItensSRV = Nothing
    Set objEventoOrcamento = Nothing
    Set gcolParcPV = Nothing

'    Set objEventoTiposDeTributacao = Nothing
'    Set gcolTiposTribICMS = Nothing
'    Set gcolTiposTribIPI = Nothing

    'Libera referencia para essa tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

    Set gobjAnotacao = Nothing

'    'Encerra tributacao
'    Call TributacaoPV_Terminar

    Set objTabComissoes = Nothing
    Set gobjPedidoVenda = Nothing

    'Alterado por Tulio em 27/03
    Set objComissoesRegrasCalc = Nothing
    Set objMnemonicoComissCalc = Nothing
    Set objMnemonicoComissCalcAux = Nothing
    Set gobjTribTab = Nothing

End Sub

Private Function Saida_Celula_PrecoUnitario(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double
Dim dQuantCancelada As Double
Dim bPrecoUnitarioIgual As Boolean

On Error GoTo Erro_Saida_Celula_PrecoUnitario

    bPrecoUnitarioIgual = False

    Set objGridInt.objControle = PrecoUnitario

    If Len(Trim(PrecoUnitario.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(PrecoUnitario.Text)
        If lErro <> SUCESSO Then gError 191249

        PrecoUnitario.Text = Format(PrecoUnitario.Text, gobjFAT.sFormatoPrecoUnitario)

    End If

    'Comparação com Preço Unitário anterior
    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col)) = StrParaDbl(PrecoUnitario.Text) Then bPrecoUnitarioIgual = True

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col))) > 0 Then dQuantidade = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col))
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_QuantCancel_Col))) > 0 Then dQuantCancelada = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantCancel_Col))

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191250

    If Not bPrecoUnitarioIgual Then

        Call CF2(Me, "PSRV_PrecoTotal_Calcula", GridItens.Row)
        
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 191251
        
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 191251

    End If

    Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = gErr


    Select Case gErr

        Case 191249 To 191251
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191252)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoUnitarioSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double
Dim dQuantCancelada As Double
Dim bPrecoUnitarioIgual As Boolean

On Error GoTo Erro_Saida_Celula_PrecoUnitarioSRV

    bPrecoUnitarioIgual = False

    Set objGridInt.objControle = PrecoUnitarioSRV

    If Len(Trim(PrecoUnitarioSRV.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(PrecoUnitarioSRV.Text)
        If lErro <> SUCESSO Then gError 191253

        PrecoUnitarioSRV.Text = Format(PrecoUnitarioSRV.Text, gobjFAT.sFormatoPrecoUnitario)

    End If

    'Comparação com Preço Unitário anterior
    If StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ValorUnitarioSRV_Col)) = StrParaDbl(PrecoUnitarioSRV.Text) Then bPrecoUnitarioIgual = True

    If Len(Trim(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantPedidaSRV_Col))) > 0 Then dQuantidade = CDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantPedidaSRV_Col))
    If Len(Trim(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantCancelSRV_Col))) > 0 Then dQuantCancelada = CDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantCancelSRV_Col))

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191254

    If Not bPrecoUnitarioIgual Then

        Call CF2(Me, "PSRV_PrecoTotal_CalculaSRV", GridItensSRV.Row)
        
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItensSRV.Row, True)
        If lErro <> SUCESSO Then gError 191255
        
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 191255

    End If

    Saida_Celula_PrecoUnitarioSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitarioSRV:

    Saida_Celula_PrecoUnitarioSRV = gErr

    Select Case gErr

        Case 191253 To 191255
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191256)

    End Select

    Exit Function

End Function

Function Saida_Celula_QuantCancelada(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade Cancelada que está deixando de ser a corrente

Dim lErro As Long
Dim objItemPedido As New ClassItemPedido
Dim dQuantidade As Double
Dim dQuantCancelada As Double
Dim dQuantFaturada As Double
Dim dPrecoUnitario As Double
Dim bQuantCanceladaIgual As Boolean
Dim objProduto As New ClassProduto
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_Saida_Celula_QuantCancelada

    bQuantCanceladaIgual = False

    Set objGridInt.objControle = QuantCancelada

    If Len(QuantCancelada.Text) > 0 Then

        'Critica o valor da quantidade
        lErro = Valor_NaoNegativo_Critica(QuantCancelada.Text)
        If lErro <> SUCESSO Then gError 191257

        QuantCancelada.Text = Formata_Estoque(QuantCancelada.Text)

    End If

    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantCancel_Col)) = StrParaDbl(QuantCancelada.Text) Then bQuantCanceladaIgual = True

    If Not bQuantCanceladaIgual Then

        dQuantCancelada = StrParaDbl(QuantCancelada.Text)

        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col))) > 0 Then dQuantidade = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col))
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_QuantFaturada_Col))) > 0 Then dQuantFaturada = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantFaturada_Col))

        If dQuantCancelada > 0 And dQuantidade < dQuantCancelada Then gError 191263
        If dQuantCancelada > 0 And (dQuantidade - dQuantCancelada < dQuantFaturada) Then gError 191264

        '####################################################################
        'Inserido por Wagner 18/05/2006
        objProduto.sCodigo = gobjPedidoVenda.colItensPedido(GridItens.Row).sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 191258
        '####################################################################

        'Tenta fazer uma alocação p\ oitem de acordo com a nova UM
        lErro = Reserva_Processa(GridItens.Row, dQuantidade, dQuantCancelada, dQuantFaturada)
        If lErro <> SUCESSO Then gError 191259

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191260

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) > 0 Then dPrecoUnitario = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))

    If Not bQuantCanceladaIgual And dPrecoUnitario > 0 Then
        'Calcula o preço total
        Call CF2(Me, "PSRV_PrecoTotal_CalculaSRV", GridItens.Row)
    
    End If

    If Not bQuantCanceladaIgual Then

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81728
        
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 81728

    End If

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 191261

    Saida_Celula_QuantCancelada = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantCancelada:

    Saida_Celula_QuantCancelada = gErr

    Select Case gErr

        Case 191257, 191258, 191259, 191260, 191261, 191262
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191263
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_PEDIDA_INFERIOR_CANCELADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191264
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_FATURADA_SUPERIOR", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191265)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Saida_Celula_QuantCanceladaSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade Cancelada que está deixando de ser a corrente

Dim lErro As Long
Dim objItemPedido As New ClassItemPedido
Dim dQuantidade As Double
Dim dQuantCancelada As Double
Dim dQuantFaturada As Double
Dim dPrecoUnitario As Double
Dim bQuantCanceladaIgual As Boolean
Dim objProduto As New ClassProduto
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_Saida_Celula_QuantCanceladaSRV

    bQuantCanceladaIgual = False

    Set objGridInt.objControle = QuantCanceladaSRV

    If Len(QuantCanceladaSRV.Text) > 0 Then

        'Critica o valor da quantidade
        lErro = Valor_NaoNegativo_Critica(QuantCanceladaSRV.Text)
        If lErro <> SUCESSO Then gError 191257

        QuantCanceladaSRV.Text = Formata_Estoque(QuantCanceladaSRV.Text)

    End If

    If StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantCancelSRV_Col)) = StrParaDbl(QuantCanceladaSRV.Text) Then bQuantCanceladaIgual = True

    If Not bQuantCanceladaIgual Then

        dQuantCancelada = StrParaDbl(QuantCanceladaSRV.Text)

        If Len(Trim(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantPedidaSRV_Col))) > 0 Then dQuantidade = CDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantPedidaSRV_Col))
        If Len(Trim(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantFaturadaSRV_Col))) > 0 Then dQuantFaturada = CDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_QuantFaturadaSRV_Col))

        If dQuantCancelada > 0 And dQuantidade < dQuantCancelada Then gError 191263
        If dQuantCancelada > 0 And (dQuantidade - dQuantCancelada < dQuantFaturada) Then gError 191264

        '####################################################################
        'Inserido por Wagner 18/05/2006
        objProduto.sCodigo = gobjPedidoVenda.colItensPedidoSRV(GridItensSRV.Row).sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 191258
        '####################################################################

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191260

    If Len(Trim(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ValorUnitarioSRV_Col))) > 0 Then dPrecoUnitario = CDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_ValorUnitarioSRV_Col))

    If Not bQuantCanceladaIgual And dPrecoUnitario > 0 Then
        'Calcula o preço total
        Call CF2(Me, "PSRV_PrecoTotal_CalculaSRV", GridItens.Row)
'        lErro = ValorTotal_Calcula()
'        If lErro <> SUCESSO Then gError 191261
    End If

    If Not bQuantCanceladaIgual Then

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 191262

        lErro = gobjTribTab.Alteracao_Item_Grid(GridItensSRV.Row, True)
        If lErro <> SUCESSO Then gError 191262
        
    End If
    
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 191262
    
    Saida_Celula_QuantCanceladaSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantCanceladaSRV:

    Saida_Celula_QuantCanceladaSRV = gErr

    Select Case gErr

        Case 191257, 191258, 191259, 191260, 191261, 191262
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191263
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_PEDIDA_INFERIOR_CANCELADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191264
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_FATURADA_SUPERIOR", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191265)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Saida_Celula_PercentDesc(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim dQuantidade As Double
Dim dQuantidadeCancelada As Double
Dim sValorPercAnterior As String
Dim dPercentDescAnterior As Double

On Error GoTo Erro_Saida_Celula_PercentDesc

    Set objGridInt.objControle = PercentDesc

    dPercentDesc = StrParaDbl(PercentDesc.Text)
    
    dPercentDescAnterior = PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col)) * 100

    If Len(PercentDesc.Text) > 0 Then
        'Critica a porcentagem
        lErro = Porcentagem_Critica(PercentDesc.Text)
        If lErro <> SUCESSO Then gError 191266

        If Abs(dPercentDesc - dPercentDescAnterior) > DELTA_VALORMONETARIO2 Then
            'se for igual a 100% -> erro
            If dPercentDesc = 100 Then gError 191267

            PercentDesc.Text = Format(dPercentDesc, "Fixed")

        End If

    Else

        dDesconto = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col))
        dValorTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))

        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(dValorTotal + dDesconto, "Standard")

    End If

    sValorPercAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col)

    If Abs(dPercentDesc - dPercentDescAnterior) > DELTA_VALORMONETARIO2 Then
        iDescontoAlterado = REGISTRO_ALTERADO
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191268
    
    'Se foi alterada
    If Format(dPercentDesc, "#0.#0\%") <> sValorPercAnterior Then
        
        'Recalcula o preço total
        Call CF2(Me, "PSRV_PrecoTotal_Calcula", GridItens.Row)

        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 51044

        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 191269

    End If

    Saida_Celula_PercentDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDesc:

    Saida_Celula_PercentDesc = gErr

    Select Case gErr

        Case 191266, 191268, 191269
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191267
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191270)

    End Select

    Exit Function

End Function

Function Saida_Celula_PercentDescSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim dQuantidade As Double
Dim dQuantidadeCancelada As Double
Dim sValorPercAnterior As String
Dim dPercentDescAnterior As Double

On Error GoTo Erro_Saida_Celula_PercentDescSRV

    Set objGridInt.objControle = PercentDescSRV

    dPercentDesc = StrParaDbl(PercentDescSRV.Text)
    
    dPercentDescAnterior = PercentParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_PercDescSRV_Col)) * 100

    If Len(PercentDescSRV.Text) > 0 Then
        'Critica a porcentagem
        lErro = Porcentagem_Critica(PercentDescSRV.Text)
        If lErro <> SUCESSO Then gError 191271

        dPercentDesc = CDbl(PercentDescSRV.Text)

        If Abs(dPercentDesc - dPercentDescAnterior) > DELTA_VALORMONETARIO2 Then
            'se for igual a 100% -> erro
            If dPercentDesc = 100 Then gError 191272

            PercentDescSRV.Text = Format(dPercentDesc, "Fixed")

        End If

    Else

        dDesconto = StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_DescontoSRV_Col))
        dValorTotal = StrParaDbl(GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_PrecoTotalSRV_Col))

        GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_DescontoSRV_Col) = ""
        GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_PrecoTotalSRV_Col) = Format(dValorTotal + dDesconto, "Standard")

    End If

    If Abs(dPercentDesc - dPercentDescAnterior) > DELTA_VALORMONETARIO2 Then
        iDescontoAlterado = REGISTRO_ALTERADO
    End If
    
    sValorPercAnterior = GridItensSRV.TextMatrix(GridItensSRV.Row, iGrid_PercDescSRV_Col)

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191273
    
    'Se foi alterada
    If Format(dPercentDesc, "#0.#0\%") <> sValorPercAnterior Then
        
        'Recalcula o preço total
        Call CF2(Me, "PSRV_PrecoTotal_CalculaSRV", GridItensSRV.Row)
       
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItensSRV.Row, True)
        If lErro <> SUCESSO Then gError 191274
        
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 191274

    End If

    Saida_Celula_PercentDescSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDescSRV:

    Saida_Celula_PercentDescSRV = gErr

    Select Case gErr

        Case 191271, 191273, 191274
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191272
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191275)

    End Select

    Exit Function

End Function


Function Saida_Celula_DataEntregaSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Entrega que está deixando de ser a corrente

Dim lErro As Long
Dim dtDataEntrega As Date
Dim dtDataEmissao As Date

On Error GoTo Erro_Saida_Celula_DataEntregaSRV

    Set objGridInt.objControle = DataEntregaSRV

    If Len(Trim(DataEntregaSRV.ClipText)) > 0 Then
        'Critica a Data informada
        lErro = Data_Critica(DataEntregaSRV.Text)
        If lErro <> SUCESSO Then gError 191276
        'Se data de emissão estiver preenchida
        If Len(Trim(DataEmissao.ClipText)) > 0 Then

            dtDataEntrega = CDate(DataEntregaSRV.Text)
            dtDataEmissao = CDate(DataEmissao.Text)
            
            'Veerifica se a data de emissão é maior que a data de entrega
            If dtDataEntrega < dtDataEmissao Then gError 191277

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191278

    Saida_Celula_DataEntregaSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_DataEntregaSRV:

    Saida_Celula_DataEntregaSRV = gErr

    Select Case gErr

        Case 191276, 191278
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191277
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAENTREGA", gErr, dtDataEntrega, dtDataEmissao)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191279)

    End Select

    Exit Function

End Function

Function Saida_Celula_DescProduto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Descrição do Produto que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescProduto

    Set objGridInt.objControle = DescricaoProduto

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191280

    Saida_Celula_DescProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_DescProduto:

    Saida_Celula_DescProduto = gErr

    Select Case gErr

        Case 191280
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191281)

    End Select

    Exit Function

End Function

Function Saida_Celula_DescServicoSRV(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Descrição do Produto que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescServicoSRV

    Set objGridInt.objControle = DescServicoSRV

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191282

    Saida_Celula_DescServicoSRV = SUCESSO

    Exit Function

Erro_Saida_Celula_DescServicoSRV:

    Saida_Celula_DescServicoSRV = gErr

    Select Case gErr

        Case 191282
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191282)

    End Select

    Exit Function

End Function

Private Function FilialFaturamento_TrataAlteracao() As Long
'Trata a alteracao da filial de faturamento selecionada para o pedido

Dim lErro As Long
On Error GoTo Erro_FilialFaturamento_TrataAlteracao

    gobjPedidoVenda.iFilialEmpresaFaturamento = Codigo_Extrai(FilialFaturamento)

    'Calula o valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 191283

    FilialFaturamento_TrataAlteracao = SUCESSO

    Exit Function

Erro_FilialFaturamento_TrataAlteracao:

    FilialFaturamento_TrataAlteracao = gErr

    Select Case gErr

        Case 191283

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191284)

    End Select

    Exit Function

End Function

Private Function FilialEntrega_TrataAlteracao() As Long
'Trata a alteracao da filial de entrega selecionada para o pedido

Dim lErro As Long
On Error GoTo Erro_FilialEntrega_TrataAlteracao

    giFilialEntregaManual = 1

    gobjPedidoVenda.iFilialEntrega = Codigo_Extrai(FilialEntrega.Text)

    'Calula o valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 191285

    FilialEntrega_TrataAlteracao = SUCESSO

    Exit Function

Erro_FilialEntrega_TrataAlteracao:

    FilialEntrega_TrataAlteracao = gErr

    Select Case gErr

        Case 191285

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191286)

    End Select

    Exit Function

End Function

Public Sub FilialEntrega_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialEntrega_Validate

    'Verifica se a FilialEntrega foi preenchida
    If Len(Trim(FilialEntrega.Text)) = 0 Or giFilialEntregaAlterada = 0 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialEntrega, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 191287

    'Se nao encontra o item com o código informado
    If lErro = 6730 Then

        'Verifica de o Cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 191288

        sCliente = Cliente.Text

        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe FilialEntrega com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 191289

        'Se não encontrou a Filial do Cliente --> erro
        If lErro = 17660 Then gError 191290

        'Coloca na tela
        FilialEntrega.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        If FilialEntrega_TrataAlteracao() <> SUCESSO Then gError 191291

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 191292

    giFilialEntregaAlterada = 0

    Exit Sub

Erro_FilialEntrega_Validate:

    Cancel = True


    Select Case gErr

        Case 191287, 191289, 191291

        Case 191288
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 191290
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 191292
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialEntrega.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191293)

    End Select

    Exit Sub

End Sub

Public Sub Transportadora_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objTransportadora As New ClassTransportadora
Dim iCodigo As Integer

On Error GoTo Erro_Transportadora_Validate

    'Verifica se foi preenchida a ComboBox Transportadora
    If Len(Trim(Transportadora.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox Transportadora
    If Transportadora.Text = Transportadora.List(Transportadora.ListIndex) Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Transportadora, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 191294

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 191295
        
        If lErro <> SUCESSO Then gError 191296

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 191297

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case gErr

        Case 191294, 191295

        Case 191296

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA", iCodigo)

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            End If
            'Segura o foco

        Case 191297
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191298)

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
    If CanalVenda.ListIndex <> -1 Then Exit Sub

    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(CanalVenda, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 191299

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCanalVenda.iCodigo = iCodigo

        'Tenta ler CanalVenda com esse código no BD
        lErro = CF("CanalVenda_Le", objCanalVenda)
        If lErro <> SUCESSO And lErro <> 23597 Then gError 191300
        
        If lErro <> SUCESSO Then gError 191301 'Não encontrou CanalVenda no BD

        'Encontrou CanalVenda no BD, coloca no Text da Combo
        CanalVenda.Text = CStr(objCanalVenda.iCodigo) & SEPARADOR & objCanalVenda.sNomeReduzido

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 191302

    Exit Sub

Erro_CanalVenda_Validate:

    Cancel = True


    Select Case gErr

        Case 191299, 191300


        Case 191301  'Não encontrou CanalVenda no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CANALVENDA", iCodigo)

            If vbMsgRes = vbYes Then

                Call Chama_Tela("CanalDeVenda", objCanalVenda)

            End If

            'Segura o foco

        Case 191302
            Call Rotina_Erro(vbOKOnly, "ERRO_CANALVENDA_NAO_ENCONTRADO", gErr, CanalVenda.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191303)

    End Select

    Exit Sub

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
        If lErro <> SUCESSO Then gError 191304

        'Compara com data de emissão
        If Len(Trim(DataEmissao.ClipText)) > 0 Then

            dtDataEmissao = CDate(DataEmissao.Text)
            dtDataReferencia = CDate(DataReferencia.Text)

            If dtDataEmissao > dtDataReferencia Then gError 191305

        End If

    End If

    giDataReferenciaAlterada = 0

    'Preenche o GridParcelas
    lErro = Cobranca_Automatica()
    If lErro <> SUCESSO Then gError 191306

    Exit Sub

Erro_DataReferencia_Validate:

    Cancel = True

    Select Case gErr

        Case 191304, 191306

        Case 191305
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAREFERENCIA", gErr, dtDataReferencia, dtDataEmissao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191307)

    End Select

    Exit Sub

End Sub

Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto) As Long
'Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas

Dim lErro As Long
Dim dValorPagar As Double
'Dim dValorIRRF As Double, dPISRetido As Double, dCOFINSRetido As Double, dCSLLRetido As Double
Dim iIndice As Integer
Dim dPercAcrescFin As Double
Dim iTamanho As Integer
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim iColuna As Integer
Dim objParcPV As ClassParcelaPedidoVenda

Dim colDescontoPadrao As New Collection

On Error GoTo Erro_GridParcelas_Preenche

    'Limpa o GridParcelas
    Call Grid_Limpa(objGridParcelas)

    Set gcolParcPV = New Collection

    'Número de Parcelas
    objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas
'
'    If Len(Trim(ValorIRRF.Text)) > 0 Then dValorIRRF = CDbl(ValorIRRF)
'    If Len(Trim(PISRetido.Text)) <> 0 And IsNumeric(PISRetido.Text) Then dPISRetido = CDbl(PISRetido.Text)
'    If Len(Trim(COFINSRetido.Text)) <> 0 And IsNumeric(COFINSRetido.Text) Then dCOFINSRetido = CDbl(COFINSRetido.Text)
'    If Len(Trim(CSLLRetido.Text)) <> 0 And IsNumeric(CSLLRetido.Text) Then dCSLLRetido = CDbl(CSLLRetido.Text)
'
'    'Valor a Pagar
'    dValorPagar = StrParaDbl(ValorTotal) - (dValorIRRF + dPISRetido + dCOFINSRetido + dCSLLRetido)

    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then gError 191308
    
    'Se Valor a Pagar for positivo
    If dValorPagar > 0 Then

        objCondicaoPagto.dValorTotal = dValorPagar

        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, False)
        If lErro <> SUCESSO Then gError 191308

        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        Next

    End If

    'Se Data Referencia estiver preenchida
    If Len(Trim(DataReferencia.ClipText)) > 0 Then

        objCondicaoPagto.dtDataRef = StrParaDate(DataReferencia.Text)
        objCondicaoPagto.dtDataEmissao = StrParaDate(DataEmissao.Text)
        objCondicaoPagto.dtDataEntrega = StrParaDate(DataEntregaPV.Text)

        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True)
        If lErro <> SUCESSO Then gError 191309

        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas

            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dtVencimento, "dd/mm/yyyy")

        Next

    End If

    ' Se dValorPagar>0 coloca desconto padrao (quantos houver, se houver) em todas as parcelas.
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        lErro = Preenche_DescontoPadrao(iIndice)
        If lErro <> SUCESSO Then gError 191310

        Set objParcPV = New ClassParcelaPedidoVenda
        objParcPV.dtDataCredito = DATA_NULA

        gcolParcPV.Add objParcPV

    Next

    GridParcelas_Preenche = SUCESSO

    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = gErr

    Select Case gErr

        Case 191308 To 191310

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191311)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorParcela(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor Parcela que está deixando de ser a corrente

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
Dim objParcPV As ClassParcelaPedidoVenda

On Error GoTo Erro_Saida_Celula_ValorParcela

    Set objGridInt.objControle = ValorParcela

    'Verifica se valor está preenchido
    If Len(ValorParcela.ClipText) > 0 Then

        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(ValorParcela.Text)
        If lErro <> SUCESSO Then gError 191312

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
                            iColDescPerc = iGrid_Desc1Perc_Col
                            iColTipoDesconto = iGrid_Desc1Codigo_Col
                        Case 2
                            iColDescPerc = iGrid_Desc2Perc_Col
                            iColTipoDesconto = iGrid_Desc2Codigo_Col
                        Case 3
                            iColDescPerc = iGrid_Desc3Perc_Col
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
            Set objParcPV = New ClassParcelaPedidoVenda
            gcolParcPV.Add objParcPV

            'Coloca DescontoPadrao
            lErro = Preenche_DescontoPadrao(GridParcelas.Row)
            If lErro <> SUCESSO Then gError 191313


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
    If lErro <> SUCESSO Then gError 191314

    Saida_Celula_ValorParcela = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorParcela:

    Saida_Celula_ValorParcela = gErr

    Select Case gErr

        Case 191312 To 191314
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191315)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataVencimento(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Vencimento que está deixando de serr a corrente

Dim lErro As Long
Dim dtDataReferencia As Date
Dim dtDataVencimento As Date
Dim sDataVencimento As String
Dim iCriouLinha As Boolean
Dim objParcPV As ClassParcelaPedidoVenda

On Error GoTo Erro_Saida_Celula_DataVencimento

    Set objGridInt.objControle = DataVencimento

    'Verifica se Data de Vencimento esta preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then gError 191316

         dtDataVencimento = CDate(DataVencimento.Text)

        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataReferencia.ClipText)) > 0 Then
            dtDataReferencia = CDate(DataReferencia.Text)
            If dtDataVencimento < dtDataReferencia Then gError 191317
        End If

        sDataVencimento = Format(dtDataVencimento, "dd/mm/yyyy")

        iCriouLinha = False
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            Set objParcPV = New ClassParcelaPedidoVenda
            gcolParcPV.Add objParcPV
            iCriouLinha = True
        End If

    End If

    If sDataVencimento <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col) Then CobrancaAutomatica.Value = vbUnchecked

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191318

    If iCriouLinha Then
        'Coloca desconto padrao (le em CPRConfig)
        lErro = Preenche_DescontoPadrao(GridParcelas.Row)
        If lErro <> SUCESSO Then gError 191319
    End If

    Saida_Celula_DataVencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_DataVencimento:

    Saida_Celula_DataVencimento = gErr

    Select Case gErr

        Case 191316, 191318, 191319
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191317
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR_REFERENCIA", gErr, dtDataVencimento, GridParcelas.Row, dtDataReferencia)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191320)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoDesconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo Desconto que está deixando de serr a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim iTipo As Integer
Dim iIndice As Integer
Dim objParcPV As ClassParcelaPedidoVenda

On Error GoTo Erro_Saida_Celula_TipoDesconto

    If GridParcelas.Col = iGrid_Desc1Codigo_Col Then
        Set objGridInt.objControle = TipoDesconto1
    ElseIf GridParcelas.Col = iGrid_Desc2Codigo_Col Then
        Set objGridInt.objControle = TipoDesconto2
    ElseIf GridParcelas.Col = iGrid_Desc3Codigo_Col Then
        Set objGridInt.objControle = TipoDesconto3
    End If

    'Verifica se o Tipo foi preenchido
    If Len(Trim(objGridInt.objControle.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If objGridInt.objControle.Text <> objGridInt.objControle.List(objGridInt.objControle.ListIndex) Then

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona_Grid(objGridInt.objControle, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 191321

            'Não foi encontrado
            If lErro = 25085 Then gError 191322
            If lErro = 25086 Then gError 191323

        End If

        'Extrai o tipo de Desconto
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
            Set objParcPV = New ClassParcelaPedidoVenda
            gcolParcPV.Add objParcPV

        End If

    Else

        For iIndice = objGridInt.objGrid.Col To iGrid_Desc3Perc_Col
            GridParcelas.TextMatrix(GridParcelas.Row, iIndice) = ""
        Next

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191324

    Saida_Celula_TipoDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDesconto:

    Saida_Celula_TipoDesconto = gErr

    Select Case gErr

        Case 191321, 191324
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191322
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191323
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", gErr, objGridInt.objControle.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191325)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoValor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Valor que está deixando de serr a corrente

Dim lErro As Long
Dim dColunaSoma As Double
Dim objParcPV As ClassParcelaPedidoVenda

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
        If lErro <> SUCESSO Then gError 191326

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            Set objParcPV = New ClassParcelaPedidoVenda
            gcolParcPV.Add objParcPV
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191327

    Saida_Celula_DescontoValor = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoValor:

    Saida_Celula_DescontoValor = gErr

    Select Case gErr

        Case 191326, 191327
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191328)

    End Select

    Exit Function

End Function

Public Sub BotaoVendedores_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoVendedores_Click

    lErro = objTabComissoes.BotaoVendedores_Click()
    If lErro <> SUCESSO Then gError 191329

    Exit Sub

Erro_BotaoVendedores_Click:

    Select Case gErr

        Case 191329

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191330)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim lErro As Long

On Error GoTo Erro_objEventoVendedor_evSelecao

    Set objVendedor = obj1

    If GridComissoes.Row > 0 Then

        'Preenche a linha de Vendedor com dados default
        lErro = objTabComissoes.VendedorLinha_Preenche(objVendedor)
        If lErro <> SUCESSO Then gError 191331

    End If

    Me.Show

    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr

        Case 191331

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191332)

    End Select

    Exit Sub

End Sub
'
'Private Function Saida_Celula_TipoBloqueio(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Tipo Bloqueio que está deixando de ser a corrente
'
'Dim lErro As Long
'Dim iCodigo As Integer
'Dim iIndice As Integer
'
'On Error GoTo Erro_Saida_Celula_TipoBloqueio
'
'    Set objGridInt.objControle = TipoBloqueio
'
'    'Verifica se o Tipo foi preenchido
'    If Len(Trim(TipoBloqueio.Text)) > 0 Then
'
'        'Verifica se ele foi selecionado
'        If TipoBloqueio.Text <> TipoBloqueio.List(TipoBloqueio.ListIndex) Then
'
'            'Tenta selecioná-lo na combo
'            lErro = Combo_Seleciona_Grid(TipoBloqueio, iCodigo)
'            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 191333
'
'            'Não foi encontrado
'            If lErro = 25085 Then gError 191334
'            If lErro = 25086 Then gError 191335
'
'        End If
'
'        'Verifica se já existe um bloqueio, não liberado, do mesmo tipo no grid.
'        For iIndice = 1 To objGridInt.iLinhasExistentes
'
'            If iIndice <> GridBloqueio.Row Then
'                iCodigo = Codigo_Extrai(GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))
'
'                If iCodigo = TipoBloqueio.ItemData(TipoBloqueio.ListIndex) Then
'                    If Len(Trim(GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col))) = 0 Then gError 191336
'                End If
'            End If
'
'        Next
'
'        'Acrescenta uma linha no Grid se for o caso
'        If GridBloqueio.Row - GridBloqueio.FixedRows = objGridInt.iLinhasExistentes Then
'            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
'
'            'Coloca data e usuário
'            GridBloqueio.TextMatrix(GridBloqueio.Row, iGrid_DataBloqueio_Col) = Format(gdtDataHoje, "dd/mm/yy")
'            GridBloqueio.TextMatrix(GridBloqueio.Row, iGrid_CodUsuario_Col) = Format(gsUsuario)
'
'        End If
'
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 191337
'
'    Saida_Celula_TipoBloqueio = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_TipoBloqueio:
'
'    Saida_Celula_TipoBloqueio = gErr
'
'    Select Case gErr
'
'        Case 191333, 191337
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 191334
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO", gErr, iCodigo)
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 191335
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO1", gErr, TipoBloqueio.Text)
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 191336
'            Call Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIOPV_REPETIDO", gErr, Codigo_Extrai(TipoBloqueio.Text))
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191338)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_ResponsavelBL(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Responsável BL que está deixando de serr a corrente
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_ResponsavelBL
'
'    Set objGridInt.objControle = ResponsavelBL
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 191339
'
'    Saida_Celula_ResponsavelBL = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_ResponsavelBL:
'
'    Saida_Celula_ResponsavelBL = gErr
'
'    Select Case gErr
'
'        Case 191339
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191340)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Observação que está deixando de ser a corrente
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_Observacao
'
'    Set objGridInt.objControle = Observacao
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 191341
'
'    Saida_Celula_Observacao = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_Observacao:
'
'    Saida_Celula_Observacao = gErr
'
'    Select Case gErr
'
'        Case 191341
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191342)
'
'    End Select
'
'    Exit Function
'
'End Function

'Public Sub BotaoLiberaBloqueio_Click()
'    Call gobjTelaTrataBloqGen.BotaoLibBloq_Click
'End Sub

'
'Public Sub BotaoLiberaBloqueio_Click()
'
'Dim lErro As Long
'Dim objPedidoVenda As New ClassPedidoDeVenda
'
'On Error GoTo Erro_BotaoLiberaBloqueio_Click
'
'    'Verifica se o Código foi preenchido
'    If Len(Trim(Codigo.ClipText)) > 0 Then
'
'        objPedidoVenda.lCodigo = CLng(Codigo.Text)
'        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
'        'Lê o Pedido de Venda
'        lErro = CF("PedidoServico_Le", objPedidoVenda)
'        If lErro <> SUCESSO And lErro <> 26509 Then gError 191343
'
'        'Não encontrou o Pedido de Venda
'        If lErro = 26509 Then objPedidoVenda.lCodigo = 0
'
'    End If
'
'    Call Chama_Tela_Modal("LiberaBloqueio", objPedidoVenda)
'
'    'Le os Bloqueios
'    lErro = CF("BloqueiosPedidoSRV_Le", objPedidoVenda)
'    If lErro <> SUCESSO Then gError 191344
'
'    'Recarrega o tab de Bloqueios
'    lErro = Carrega_Tab_Bloqueio(objPedidoVenda)
'    If lErro <> SUCESSO Then gError 191345
'
'    Exit Sub
'
'Erro_BotaoLiberaBloqueio_Click:
'
'    Select Case gErr
'
'        Case 191343 To 191345
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191346)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub BotaoLibera_Click()

Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoLibera_Click

    If objGridAlocacao.iLinhasExistentes = 0 Then gError 191347

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_LIBERACAO_RESERVAS")
    If vbMsgRes = vbYes Then

        Call Grid_Limpa(objGridAlocacao)

        ProdutoDescricao.Caption = ""

        For iIndice = 1 To objGridItens.iLinhasExistentes
            GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col) = ""
        Next


    End If

    Exit Sub

Erro_BotaoLibera_Click:

    Select Case gErr

        Case 191347
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_HA_RESERVAS_PARA_LIBERAR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191348)

    End Select

    Exit Sub

End Sub

Public Sub BotaoReserva_Click()

Dim colItemPedido As New colItemPedido
Dim iIndice As Integer
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim dQuantidade As Double
Dim colItemReserva As colReserva
Dim objItemPedido As ClassItemPedido
Dim lNumPedido As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim objProduto As New ClassProduto
Dim dQuantFaturar As Double
Dim dQuantCancelada As Double
Dim dQuantReservada As Double
Dim dQuantPedida As Double
Dim dQuantFaturada As Double
Dim sProdutoEnxuto As String
Dim colItensRomaneio As New Collection
Dim objProdutoAux As New ClassProduto
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objItemRomaneioAux As ClassItemRomaneioGrade
Dim bAchou As Boolean
Dim iIndice2 As Integer

On Error GoTo Erro_BotaoReserva_Click

    'Se não existir itens no grid, sai.
    If objGridItens.iLinhasExistentes = 0 Then Exit Sub
    'pega o código da pedido da tela
    If Len(Trim(Codigo.ClipText)) > 0 Then lNumPedido = CLng(Codigo.Text)

    'Para cada item do grid de itens
    For iIndice = 1 To objGridItens.iLinhasExistentes

        dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col))
        dQuantCancelada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))
        dQuantReservada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col))
        dQuantFaturada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantFaturada_Col))
        dQuantFaturar = dQuantidade - dQuantCancelada - dQuantFaturada


        '************* CRIADA A FUNÇÃO Produto_Critica_Filial2 PARA ADAPAR GRADE
        'Faz a critica do produto
        lErro = CF("Produto_Critica_Filial2", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), objProduto, iPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 191349
        
        If lErro = 51381 Then gError 191350

        'Se o produto é gerencial sem ser de grade ==> Erro
        If lErro = 86295 Then
            gError 191351
        End If

        'se o produto faz o controle de reserva
        If objProduto.iControleEstoque = PRODUTO_CONTROLE_RESERVA Then

            Set colItemReserva = New colReserva
            'Recolhe as reservas daquele item
            Call Recolhe_Reserva(iIndice, colItemReserva)
            'adiciona os dados do item a coleção de itens
            colItemPedido.Add giFilialEmpresa, lNumPedido, iIndice, objProduto.sCodigo, dQuantidade, 0, 0, 0, DATA_NULA, objProduto.sDescricao, 0, dQuantCancelada, dQuantReservada, colItemReserva, objProduto.sNomeReduzido, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), "", objProduto.sSiglaUMEstoque, objProduto.iClasseUM, dQuantFaturada, dQuantFaturar, "", 0, objProduto.iControleEstoque, 0, 0, 0, 0, 0, 0, 0, 0, 0
        End If
    Next
    'Chama a tela de alocação de produtos
    Call Chama_Tela_Modal("AlocacaoProduto", colItemPedido)

    'Limpa o grid de reserva
    Call Grid_Limpa(objGridAlocacao)
    'Para cada item da coleção de itens
    For Each objItemPedido In colItemPedido

        'Preenche no grid itens a quantidade reservada e a quantidade cancelada
        If objItemPedido.dQuantReservada > 0 Then
            GridItens.TextMatrix(objItemPedido.iItem, iGrid_QuantReservada_Col) = Formata_Estoque(objItemPedido.dQuantReservada)
        Else
            GridItens.TextMatrix(objItemPedido.iItem, iGrid_QuantReservada_Col) = ""
        End If
        If objItemPedido.dQuantCancelada > 0 Then
            GridItens.TextMatrix(objItemPedido.iItem, iGrid_QuantCancel_Col) = Formata_Estoque(objItemPedido.dQuantCancelada)
        Else
            GridItens.TextMatrix(objItemPedido.iItem, iGrid_QuantCancel_Col) = ""
        End If

        'Formata o produto
        lErro = CF("Produto_Formata", GridItens.TextMatrix(objItemPedido.iItem, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 191352
        'Se o produto foi alterado
        If sProduto <> objItemPedido.sProduto Then

            lErro = Mascara_RetornaProdutoEnxuto(objItemPedido.sProduto, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 191353

            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True
            'Preenche no grid de itens os dados do novo produto
            GridItens.TextMatrix(objItemPedido.iItem, iGrid_Produto_Col) = Produto.Text
            GridItens.TextMatrix(objItemPedido.iItem, iGrid_DescProduto_Col) = objItemPedido.sDescricao
            GridItens.TextMatrix(objItemPedido.iItem, iGrid_QuantPedida_Col) = Formata_Estoque(objItemPedido.dQuantidade)
            GridItens.TextMatrix(objItemPedido.iItem, iGrid_UnidadeMed_Col) = objItemPedido.sUnidadeMed
            'Se a tabla de preços estiver preenchida
            If Len(Trim(TabelaPreco.Text)) > 0 Then

                objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
                objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
                objTabelaPrecoItem.sCodProduto = objItemPedido.sProduto
                'Lê a tabela de preço item para o produto em gifilialempresa
                lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                If lErro <> SUCESSO And lErro <> 28014 Then gError 191354
                If lErro <> SUCESSO Then
                    objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
                    'Lê a tabela de preço item para o produto na empresa toda
                    lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                    If lErro <> SUCESSO And lErro <> 28014 Then gError 191355
                End If
                'Se alguma tabela de preços foi encontrada
                If lErro <> 28014 Then
                    'Calcula o preço unitário do produto
                    lErro = PrecoUnitario_Calcula(objItemPedido.sUnidadeMed, objTabelaPrecoItem, dPrecoUnitario)
                    If lErro <> SUCESSO Then gError 191356
                    'Coloca o valor do preço unitário no grid
                    If dPrecoUnitario > 0 Then
                        GridItens.TextMatrix(objItemPedido.iItem, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
                    Else
                        GridItens.TextMatrix(objItemPedido.iItem, iGrid_ValorUnitario_Col) = ""
                    End If


                End If
            End If
            'Calcula o preço total e o valor total
            Call CF2(Me, "PSRV_PrecoTotal_Calcula", objItemPedido.iItem)
            lErro = ValorTotal_Calcula()
            If lErro <> SUCESSO Then gError 191357

        End If
        
    Next

    'Preenche o grid de reservas
    Call Preenche_GridAlocacao(colItemPedido)

    Exit Sub

Erro_BotaoReserva_Click:

    Select Case gErr

        Case 191349, 191352, 191354, 191355, 191356, 191357

        Case 191350
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProduto)

        Case 191351
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, sProduto)

        Case 191353
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191358)

    End Select

    Exit Sub

End Sub

Private Function Move_Pedido_Memoria(objPedidoVenda As ClassPedidoDeVenda) As Long
'Move os dados da tela para objPedidoVenda

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim dValorTotalParcelas As Double
Dim objParcelaPV As ClassParcelaPedidoVenda
Dim dValorIRRF As Double
Dim iTipoDocOrigem As Integer
Dim colEmbalagensItens As New Collection
Dim objEmbalagensItem As ClassEmbalagensItem
Dim objOrcamentoVenda As ClassOrcamentoVenda 'Por Leo em 25/04/02
Dim objProjetoInfo As Object
Dim objSolicSRV As New ClassSolicSRV

Dim iIndice As Integer

On Error GoTo Erro_Move_Pedido_Memoria

    If Len(Trim(Codigo.Text)) > 0 Then objPedidoVenda.lCodigo = CLng(Codigo.Text)

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 191359

        'Se não achou o Cliente --> erro
        If lErro = 12348 Then gError 191360

        'Guarda código do Cliente em objPedidoVenda
        objPedidoVenda.lCliente = objCliente.lCodigo

    End If

    'Verifica se a Filial está preenchida
    If Len(Trim(Filial.Text)) > 0 Then
        objPedidoVenda.iFilial = Codigo_Extrai(Filial.Text)
        'Se a função código extrai não conseguiu obter o código da filial => erro, pois a filial não está cadastrada
        'isso acontece quando se traz para a tela um orçamento que foi gerado com uma filial de cliente inexistente
        If objPedidoVenda.iFilial = 0 Then gError 191361
    End If

    'Se o Código do Orçamento estiver preenchido
    If Len(Trim(Orcamento.Text)) <> 0 Then

        'Se a Filial não estiver preenchida ---> Erro.
        If Len(Trim(FilialOrcamento.Text)) = 0 Then gError 191362

        Set objOrcamentoVenda = New ClassOrcamentoVenda

        objOrcamentoVenda.lCodigo = StrParaLong(Orcamento.Text)
        objOrcamentoVenda.iFilialEmpresa = Codigo_Extrai(FilialOrcamento.Text)

        'Lê o Orçamento com o Código e a Filal informados.
        lErro = CF("OrcamentoSRV_Le", objOrcamentoVenda)
        If lErro <> SUCESSO And lErro <> 101232 Then gError 191362

        'Se não existe o Orçamento informado ---> Erro.
        If lErro = 101232 Then gError 191363

        'Armazena o NumIntDoc do Orcamento em objPedidoVenda.lNumIntOrcVenda
        objPedidoVenda.lNumIntOrcVenda = objOrcamentoVenda.lNumIntDoc

    End If

    'Preenche objPedidoVenda com dados da tela
    objPedidoVenda.iFilialEmpresaFaturamento = Codigo_Extrai(FilialFaturamento.Text)
    objPedidoVenda.dtDataEmissao = MaskedParaDate(DataEmissao)
    objPedidoVenda.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)
    objPedidoVenda.sNaturezaOp = gobjTribTab.sNatOpInterna
    objPedidoVenda.iFaturaIntegral = FaturaIntegral.Value
    objPedidoVenda.dValorFrete = StrParaDbl(ValorFrete.Text)
    objPedidoVenda.dValorSeguro = StrParaDbl(ValorSeguro.Text)
    objPedidoVenda.dValorDesconto = StrParaDbl(ValorDesconto.Text)
    objPedidoVenda.dValorOutrasDespesas = StrParaDbl(ValorDespesas.Text)
    objPedidoVenda.dValorProdutos = StrParaDbl(ValorProdutos.Caption)
    objPedidoVenda.dValorTotal = StrParaDbl(ValorTotal.Caption)
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa

    If Len(Trim(CodSolicSRV.Text)) > 0 Then
    
        objSolicSRV.iFilialEmpresa = giFilialEmpresa
        objSolicSRV.lCodigo = StrParaLong(CodSolicSRV.Text)
        
        lErro = CF("SolicitacaoSRV_Le", objSolicSRV)
        If lErro <> SUCESSO And lErro <> 183258 Then gError 193476

        If lErro <> SUCESSO Then gError 193477
        
        objPedidoVenda.lNumIntSolicSRV = objSolicSRV.lNumIntDoc
        
    End If


    'Move Grid Itens para memória
    lErro = Move_GridItensSRV_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then gError 191364

    'Move Grid Itens para memória
    lErro = Move_GridItens_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then gError 191365

    'Move Tab Complemento para memória
    Call Move_TabComplemento_Memoria(objPedidoVenda)

    'Move Tab Cobrança para memória
    Call Move_TabCobranca_Memoria(objPedidoVenda)
    'Alterado por Tulio em 28/03/02
    'Move Tab Comissões para memória
    lErro = objTabComissoes.Move_TabComissoes_Memoria(objPedidoVenda, PEDIDO_DE_VENDA)

    'Move Tab Bloqueio para memória
    'Call Move_TabBloqueio_Memoria(objPedidoVenda)
    
    lErro = gobjTelaTrataBloqGen.Move_GridBloqueio_Memoria(objPedidoVenda.colBloqueiosPV)
    If lErro <> SUCESSO Then gError 191365
    
    'Move Grid Reserva para memória
    Call Move_GridAlocacao_Memoria(objPedidoVenda)

    objPedidoVenda.iFilialEmpresa = giFilialEmpresa

    'Move Tributacao para objPedidoVenda
    Set objPedidoVenda.objTributacaoPV = gobjPedidoVenda.objTributacaoPV

'    '##################################################
'    'Inserido por Wagner 09/08/2006
'    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_PV)
'    If lErro <> SUCESSO Then gError 181761
'
'    Set objPedidoVenda.objProjetoInfo = objProjetoInfo
'    '##################################################

    Set objPedidoVenda.objTela = Me
    
    Set objPedidoVenda.objInfoAdic = gobjInfoAdic
    objPedidoVenda.dValorItens = StrParaDbl(ValorProdutos2.Caption)
    objPedidoVenda.dValorDescontoItens = StrParaDbl(ValorDescontoItens.Text)

    Move_Pedido_Memoria = SUCESSO

    Exit Function

Erro_Move_Pedido_Memoria:

    Move_Pedido_Memoria = gErr

    Select Case gErr

        Case 191359, 191362, 191364, 191365, 193476

        Case 191360
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case 191361
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case 191362
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALORCAMENTO_NAO_INFORMADA", gErr)

        Case 191363
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
        
        Case 193477
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRV_NAO_ENCONTRADO", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191366)

    End Select

    Exit Function

End Function

Private Function Move_GridItens_Memoria(objPedidoVenda As ClassPedidoDeVenda) As Long
'Move Grid Itens para memória

Dim lErro As Long, iIndice As Integer
Dim sProduto1 As String, sProduto As String, iPreenchido As Integer

On Error GoTo Erro_Move_GridItens_Memoria

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        sProduto1 = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)
        
        'Verifica se o Produto está preenchido
        If Len(Trim(sProduto1)) > 0 Then
    
            'Formata o produto
            lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 191367
            
        End If

        lErro = Move_GridItem_Memoria(objPedidoVenda, iIndice, sProduto)
        If lErro <> SUCESSO Then gError 191367

    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 191367

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191368)

    End Select

    Exit Function

End Function

Private Function Move_GridItensSRV_Memoria(objPedidoVenda As ClassPedidoDeVenda) As Long
'Move Grid Itens para memória

Dim lErro As Long, iIndice As Integer
Dim sProduto1 As String, sProduto As String, iPreenchido As Integer

On Error GoTo Erro_Move_GridItensSRV_Memoria

    For iIndice = 1 To objGridItensSRV.iLinhasExistentes

        sProduto1 = GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col)
        
        'Verifica se o Produto está preenchido
        If Len(Trim(sProduto1)) > 0 Then
    
            'Formata o produto
            lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 191369
            
        End If
        
        lErro = Move_GridItem_MemoriaSRV(objPedidoVenda, iIndice, sProduto)
        If lErro <> SUCESSO Then gError 191369

    Next

    Move_GridItensSRV_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItensSRV_Memoria:

    Move_GridItensSRV_Memoria = gErr

    Select Case gErr

        Case 191369

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191370)

    End Select

    Exit Function

End Function

Private Sub Move_TabComplemento_Memoria(objPedidoVenda As ClassPedidoDeVenda)

On Error GoTo Erro_Move_TabComplemento_Memoria

    'Recolhe os dados de complemento
    objPedidoVenda.sPlaca = Placa.Text
    objPedidoVenda.sPlacaUF = PlacaUF.Text
    objPedidoVenda.sVolumeNumero = VolumeNumero.Text
    objPedidoVenda.lVolumeQuant = StrParaLong(VolumeQuant.Text)
    objPedidoVenda.dPesoBruto = StrParaDbl(PesoBruto.Text)
    objPedidoVenda.dPesoLiq = StrParaDbl(PesoLiquido.Text)
    objPedidoVenda.dVolumeTotal = StrParaDbl(Cubagem.Text)

    'Armazena o responsável pelo frete
    If Emitente.Value Then
        objPedidoVenda.iFreteRespons = FRETE_EMITENTE
    Else
        objPedidoVenda.iFreteRespons = FRETE_DESTINATARIO
    End If

    If Len(Trim(FilialEntrega.Text)) > 0 Then
        objPedidoVenda.iFilialEntrega = Codigo_Extrai(FilialEntrega.Text)

        'Se a função código extrai não conseguiu obter o código da filial => erro, pois a filial não está cadastrada
        'isso acontece quando se traz para a tela um orçamento que foi gerado com uma filial de cliente inexistente
        If objPedidoVenda.iFilialEntrega = 0 Then gError 191371
    End If

    If Len(Trim(Transportadora.Text)) > 0 Then objPedidoVenda.iCodTransportadora = Codigo_Extrai(Transportadora.Text)
    If Len(Trim(TranspRedespacho.Text)) > 0 Then objPedidoVenda.iCodTranspRedesp = Codigo_Extrai(TranspRedespacho.Text)
    objPedidoVenda.iDetPagFrete = RedespachoCli.Value
    If Len(Trim(PedidoCliente.ClipText)) > 0 Then objPedidoVenda.sPedidoCliente = PedidoCliente.Text
    If Len(Trim(Mensagem.Text)) > 0 Then objPedidoVenda.sMensagemNota = Mensagem.Text
    If Len(Trim(CanalVenda.Text)) > 0 Then objPedidoVenda.iCanalVenda = Codigo_Extrai(CanalVenda.Text)

    If Len(Trim(VolumeEspecie.Text)) > 0 Then objPedidoVenda.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objPedidoVenda.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)

    objPedidoVenda.dtDataEntrega = MaskedParaDate(DataEntregaPV)

    If Len(Trim(PedidoRepr.ClipText)) > 0 Then objPedidoVenda.sPedidoRepresentante = PedidoRepr.Text

    Exit Sub
    
Erro_Move_TabComplemento_Memoria:

    Select Case gErr

        Case 191371

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191372)

    End Select

    Exit Sub

End Sub

Private Sub Move_TabCobranca_Memoria(objPedidoVenda As ClassPedidoDeVenda)
'Recolhe os dados do tab de cobrança

Dim lTamanho As Long
Dim iIndice As Integer
Dim objParcelaPV As ClassParcelaPedidoVenda
Dim objParcPV As ClassParcelaPedidoVenda
Dim objParcPV1 As ClassParcelaPedidoVenda

    'Recolhe os dados da Cobrança
    objPedidoVenda.iCobrancaAutomatica = CobrancaAutomatica.Value
    objPedidoVenda.dtDataReferencia = MaskedParaDate(DataReferencia)
    objPedidoVenda.iCondicaoPagto = Codigo_Extrai(CondicaoPagamento.Text)
    objPedidoVenda.dPercAcrescFinanceiro = StrParaDbl(PercAcrescFin.Text) / 100

    If objGridParcelas.iLinhasExistentes = 0 Then Exit Sub

    'Recolhe os Dados do Grid de Parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        Set objParcelaPV = New ClassParcelaPedidoVenda

        objParcelaPV.iNumParcela = iIndice

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) > 0 Then
            objParcelaPV.dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        Else
            objParcelaPV.dtDataVencimento = DATA_NULA
        End If

        objParcelaPV.dValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
        objParcelaPV.iTipoPagto = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_TipoPagto_Col))
        objParcelaPV.iDesconto1Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
        objParcelaPV.iDesconto2Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
        objParcelaPV.iDesconto3Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))) > 0 Then
            objParcelaPV.dtDesconto1Ate = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
        Else
            objParcelaPV.dtDesconto1Ate = DATA_NULA
        End If
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))) > 0 Then
            objParcelaPV.dtDesconto2Ate = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
        Else
            objParcelaPV.dtDesconto2Ate = DATA_NULA
        End If
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))) > 0 Then
            objParcelaPV.dtDesconto3Ate = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
        Else
            objParcelaPV.dtDesconto3Ate = DATA_NULA
        End If

        If objParcelaPV.iDesconto1Codigo = VALOR_FIXO Or objParcelaPV.iDesconto1Codigo = VALOR_ANT_DIA Or objParcelaPV.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
            objParcelaPV.dDesconto1Valor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))
        ElseIf objParcelaPV.iDesconto1Codigo = Percentual Or objParcelaPV.iDesconto1Codigo = PERC_ANT_DIA Or objParcelaPV.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col)))
            If lTamanho > 0 Then objParcelaPV.dDesconto1Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col))
        End If

        If objParcelaPV.iDesconto2Codigo = VALOR_FIXO Or objParcelaPV.iDesconto2Codigo = VALOR_ANT_DIA Or objParcelaPV.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
            objParcelaPV.dDesconto2Valor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))
        ElseIf objParcelaPV.iDesconto2Codigo = Percentual Or objParcelaPV.iDesconto2Codigo = PERC_ANT_DIA Or objParcelaPV.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col)))
            If lTamanho > 0 Then objParcelaPV.dDesconto2Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col))
        End If

        If objParcelaPV.iDesconto3Codigo = VALOR_FIXO Or objParcelaPV.iDesconto3Codigo = VALOR_ANT_DIA Or objParcelaPV.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
            objParcelaPV.dDesconto3Valor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))
        ElseIf objParcelaPV.iDesconto3Codigo = Percentual Or objParcelaPV.iDesconto3Codigo = PERC_ANT_DIA Or objParcelaPV.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col)))
            If lTamanho > 0 Then objParcelaPV.dDesconto3Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col))
        End If



        With objParcelaPV
            Set objParcPV = objPedidoVenda.colParcelas.Add(.dValor, .dtDataVencimento, .iNumParcela, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .dtDesconto3Ate, .dDesconto3Valor, .iDesconto3Codigo)
            objParcPV.iTipoPagto = objParcelaPV.iTipoPagto

            Set objParcPV1 = gcolParcPV.Item(iIndice)

            objParcPV.dtDataDepositoCheque = objParcPV1.dtDataDepositoCheque
            objParcPV.dtDataEmissaoCheque = objParcPV1.dtDataEmissaoCheque
            objParcPV.dtValidadeCartao = objParcPV1.dtValidadeCartao
            objParcPV.iAdmMeioPagto = objParcPV1.iAdmMeioPagto
            objParcPV.iBancoCheque = objParcPV1.iBancoCheque
            objParcPV.iCodConta = objParcPV1.iCodConta
            objParcPV.dtDataCredito = objParcPV1.dtDataCredito
            objParcPV.iParcelamento = objParcPV1.iParcelamento
            objParcPV.lNumeroCheque = objParcPV1.lNumeroCheque
            objParcPV.sAgenciaCheque = objParcPV1.sAgenciaCheque
            objParcPV.sAprovacaoCartao = objParcPV1.sAprovacaoCartao
            objParcPV.sContaCorrenteCheque = objParcPV1.sContaCorrenteCheque
            objParcPV.sNumeroCartao = objParcPV1.sNumeroCartao
            objParcPV.dtDataTransacaoCartao = objParcPV1.dtDataTransacaoCartao

        End With

    Next
    
    objPedidoVenda.dValorDescontoTit = StrParaDbl(ValorDescontoTit.Text)

End Sub

'Private Sub Move_TabBloqueio_Memoria(objPedidoVenda As ClassPedidoDeVenda)
''Recolhe os dados do tab de Bloqueio
'
'Dim lErro As Long
'Dim objBloqueioPV As ClassBloqueioPV
'Dim iIndice As Integer
'
'    If objGridBloqueio.iLinhasExistentes = 0 Then Exit Sub
'
'    'Recolhe os Bloqueios do Grid de Bloqueios
'    For iIndice = 1 To objGridBloqueio.iLinhasExistentes
'
'        Set objBloqueioPV = New ClassBloqueioPV
'
'        objBloqueioPV.iTipoDeBloqueio = Codigo_Extrai(GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))
'
'        If Len(Trim(GridBloqueio.TextMatrix(iIndice, iGrid_DataBloqueio_Col))) > 0 Then
'            objBloqueioPV.dtData = CDate(GridBloqueio.TextMatrix(iIndice, iGrid_DataBloqueio_Col))
'        Else
'            objBloqueioPV.dtData = DATA_NULA
'        End If
'
'        objBloqueioPV.sCodUsuario = GridBloqueio.TextMatrix(iIndice, iGrid_CodUsuario_Col)
'        objBloqueioPV.sResponsavel = GridBloqueio.TextMatrix(iIndice, iGrid_ResponsavelBL_Col)
'        objBloqueioPV.sObservacao = GridBloqueio.TextMatrix(iIndice, iGrid_Observacao_Col)
'        objBloqueioPV.iSequencial = StrParaInt(GridBloqueio.TextMatrix(iIndice, iGrid_SeqBloqueio_Col))
'
'        If Len(Trim(GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col))) > 0 Then
'            objBloqueioPV.dtDataLib = CDate(GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col))
'        Else
'            objBloqueioPV.dtDataLib = DATA_NULA
'        End If
'
'        objBloqueioPV.sResponsavelLib = GridBloqueio.TextMatrix(iIndice, iGrid_ResponsavelLib_Col)
'
'
'        'Adiciona o bloqueio na coleção de bloqueios
'        With objBloqueioPV
'            objPedidoVenda.colBloqueiosPV.Add giFilialEmpresa, objPedidoVenda.lCodigo, objBloqueioPV.iSequencial, .iTipoDeBloqueio, .sCodUsuario, .sResponsavel, .dtData, "", .sResponsavelLib, .dtDataLib, .sObservacao
'        End With
'
'    Next
'
'    Exit Sub
'
'End Sub

Private Sub Move_GridAlocacao_Memoria(objPedidoVenda As ClassPedidoDeVenda)
'Recolhe os dados do Grid de Reservas

Dim lErro As Long
Dim iIndice As Integer
Dim iItem As Integer
Dim iItemAtual As Integer
Dim objReservaItem As ClassReservaItem
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Move_GridAlocacao_Memoria

    If objGridAlocacao.iLinhasExistentes = 0 Then Exit Sub

    'Armazena o Número do item da primeira reserva no grid
    iItem = CInt(GridAlocacao.TextMatrix(1, iGrid_Item_Col))

    iIndice = 0

    'ESSE CONTROLE É FEITO PARA QUE SE POSSA RECOLHER AS RESERVAS
    'CORRETAMENTE MESMO NÃO ESTANDO ORDENADAS PELO ITEM

    'Enquanto não percorrer todo o grid
    Do While iIndice < objGridAlocacao.iLinhasExistentes

        iItemAtual = iItem
        'Enquanto for o mesmo item
        Do While iItem = iItemAtual

            iIndice = iIndice + 1

            Set objReservaItem = New ClassReservaItem

            objReservaItem.sAlmoxarifado = GridAlocacao.TextMatrix(iIndice, iGrid_Almoxarifado_Col)

            'Lê o almoxarifado
            objAlmoxarifado.sNomeReduzido = objReservaItem.sAlmoxarifado
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25060 Then gError 191373

            'Se não achou o Almoxarifado --> erro
            If lErro <> SUCESSO Then gError 191374

            objReservaItem.iAlmoxarifado = objAlmoxarifado.iCodigo

            If Len(Trim((GridAlocacao.TextMatrix(iIndice, iGrid_QuantReserv_Col)))) <> 0 Then objReservaItem.dQuantidade = CDbl(GridAlocacao.TextMatrix(iIndice, iGrid_QuantReserv_Col))

            If Len(Trim((GridAlocacao.TextMatrix(iIndice, iGrid_Validade_Col)))) <> 0 Then
                objReservaItem.dtDataValidade = CDate(GridAlocacao.TextMatrix(iIndice, iGrid_Validade_Col))
            Else
                objReservaItem.dtDataValidade = DATA_NULA
            End If

            objReservaItem.sResponsavel = GridAlocacao.TextMatrix(iIndice, iGrid_Responsavel_Col)

            'Adiciona a reserva na coleção de reservas do item correspondente
            With objReservaItem
                objPedidoVenda.colItensPedido(iItemAtual).colReserva.Add .iFilialEmpresa, 0, "", .iAlmoxarifado, 0, 0, 0, .dQuantidade, DATA_NULA, .dtDataValidade, "", .sResponsavel, 0, .sAlmoxarifado
            End With

            If Len(Trim(GridAlocacao.TextMatrix(iIndice + 1, iGrid_Item_Col))) > 0 Then
                iItem = GridAlocacao.TextMatrix(iIndice + 1, iGrid_Item_Col)
            Else
                Exit Do
            End If

        Loop

    Loop

    Exit Sub

Erro_Move_GridAlocacao_Memoria:

    Select Case gErr

        Case 191373

        Case 191374
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191375)

    End Select

    Exit Sub

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

Public Sub TipoDesconto1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDesconto1_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub TipoDesconto1_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub TipoDesconto1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = TipoDesconto1
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TipoDesconto2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDesconto2_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub TipoDesconto2_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub TipoDesconto2_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = TipoDesconto2
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TipoDesconto3_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDesconto3_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub TipoDesconto3_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub TipoDesconto3_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = TipoDesconto3
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

        If gcolParcPV.Count >= GridParcelas.Row Then gcolParcPV.Remove (GridParcelas.Row)

    End If

    Exit Sub

Erro_GridParcelas_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191376)

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

Public Sub GridComissoes_GotFocus()

    Call Grid_Recebe_Foco(objGridComissoes)

End Sub

Public Sub GridComissoes_EnterCell()

    Call Grid_Entrada_Celula(objGridComissoes, iAlterado)

End Sub

Public Sub GridComissoes_LeaveCell()

    Call Saida_Celula(objGridComissoes)

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

Public Sub GridComissoes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridComissoes)

End Sub

Public Sub GridComissoes_RowColChange()

    Call Grid_RowColChange(objGridComissoes)

End Sub

Public Sub GridComissoes_Scroll()

    Call Grid_Scroll(objGridComissoes)

End Sub
'
'Public Sub TipoBloqueio_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub TipoBloqueio_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridBloqueio)
'
'End Sub
'
'Public Sub TipoBloqueio_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)
'
'End Sub
'
'Public Sub TipoBloqueio_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridBloqueio.objControle = TipoBloqueio
'    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub DataBloqueio_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub DataBloqueio_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridBloqueio)
'
'End Sub
'
'Public Sub DataBloqueio_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)
'
'End Sub
'
'Public Sub DataBloqueio_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridBloqueio.objControle = DataBloqueio
'    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Observacao_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Observacao_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridBloqueio)
'
'End Sub
'
'Public Sub Observacao_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)
'
'End Sub
'
'Public Sub Observacao_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridBloqueio.objControle = Observacao
'    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub CodUsuario_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub CodUsuario_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridBloqueio)
'
'End Sub
'
'Public Sub CodUsuario_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)
'
'End Sub
'
'Public Sub CodUsuario_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridBloqueio.objControle = CodUsuario
'    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub ResponsavelBL_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ResponsavelBL_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridBloqueio)
'
'End Sub
'
'Public Sub ResponsavelBL_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)
'
'End Sub
'
'Public Sub ResponsavelBL_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridBloqueio.objControle = ResponsavelBL
'    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub GridBloqueio_Click()
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Click(objGridBloqueio, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
'    End If
'
'End Sub
'
'Public Sub GridBloqueio_GotFocus()
'
'    Call Grid_Recebe_Foco(objGridBloqueio)
'
'End Sub
'
'Public Sub GridBloqueio_EnterCell()
'
'    Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
'
'End Sub
'
'Public Sub GridBloqueio_LeaveCell()
'
'    Call Saida_Celula(objGridBloqueio)
'
'End Sub
'
'Public Sub GridBloqueio_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyDelete Then
'
'        If GridBloqueio.Row <= giGridBloqueioLinhasBloqueadas Then Exit Sub
'
'    End If
'
'    Call Grid_Trata_Tecla1(KeyCode, objGridBloqueio)
'
'    Exit Sub
'
'End Sub
'
'Public Sub GridBloqueio_KeyPress(KeyAscii As Integer)
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Trata_Tecla(KeyAscii, objGridBloqueio, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
'    End If
'
'End Sub
'
'Public Sub GridBloqueio_Validate(Cancel As Boolean)
'
'    Call Grid_Libera_Foco(objGridBloqueio)
'
'End Sub
'
'Public Sub GridBloqueio_RowColChange()
'
'    Call Grid_RowColChange(objGridBloqueio)
'
'End Sub
'
'Public Sub GridBloqueio_Scroll()
'
'    Call Grid_Scroll(objGridBloqueio)
'
'End Sub

Public Sub GridAlocacao_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridAlocacao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAlocacao, iAlterado)
    End If

End Sub

Public Sub GridAlocacao_GotFocus()

    Call Grid_Recebe_Foco(objGridAlocacao)

End Sub

Public Sub GridAlocacao_EnterCell()

    Call Grid_Entrada_Celula(objGridAlocacao, iAlterado)

End Sub

Public Sub GridAlocacao_LeaveCell()

    Call Saida_Celula(objGridAlocacao)

End Sub

Public Sub GridAlocacao_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iIndice As Long
Dim iItemAtual As Integer
Dim objProduto As New ClassProduto
Dim sProdutoAnterior As String
Dim dQuantidadeReservadaAnterior As Double
Dim sUmAnterior As String
Dim dQuantidadeReservadaAnteriorConvertida As Double
Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim dFator As Double
Dim objItemRomaneioGrade As ClassItemRomaneioGrade 'Grade
Dim objReservaItem As ClassReservaItem 'Grade
Dim iItem As Integer 'Grade
Dim sAlmoxarifado As String 'Grade
Dim objProdutoItem As ClassProduto

On Error GoTo Erro_GridAlocacao_KeyDown

    If Not (objGridAlocacao.iLinhasExistentes > 0) Then Exit Sub

    iLinhasExistentesAnterior = objGridAlocacao.iLinhasExistentes
    iItemAtual = GridAlocacao.Row
    iItem = GridAlocacao.TextMatrix(iItemAtual, iGrid_Item_Col) 'Grade
    sAlmoxarifado = GridAlocacao.TextMatrix(iItemAtual, iGrid_Almoxarifado_Col) 'Grade

    dQuantidadeReservadaAnterior = CDbl(GridAlocacao.TextMatrix(iItemAtual, iGrid_QuantReserv_Col))
    sProdutoAnterior = GridAlocacao.TextMatrix(iItemAtual, iGrid_ProdutoAlmox_Col)
    sUmAnterior = GridAlocacao.TextMatrix(iItemAtual, iGrid_UMEstoque_Col)

    'Formata o produto
    lErro = CF("Produto_Formata", sProdutoAnterior, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 191377

    objProduto.sCodigo = sProdutoFormatado

    'Lê o produto para pegar a UM_Estoque
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 191378

    If lErro = 28030 Then gError 191379

    Call Grid_Trata_Tecla1(KeyCode, objGridAlocacao)

    If iLinhasExistentesAnterior <> objGridAlocacao.iLinhasExistentes Then

        For iIndice = 1 To objGridItens.iLinhasExistentes


            If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoAnterior Then

                'Converte a Quantidade se Necessário
                lErro = CF("UM_Conversao", objProduto.iClasseUM, sUmAnterior, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), dFator)
                If lErro <> SUCESSO Then gError 191380

                dQuantidadeReservadaAnteriorConvertida = dQuantidadeReservadaAnterior * dFator

                GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col) = Format(CDbl((GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col)) - dQuantidadeReservadaAnteriorConvertida), "Standard")
                Exit For
            End If
        Next
    End If

    Exit Sub

Erro_GridAlocacao_KeyDown:

    Select Case gErr

        Case 191377, 191378, 191380

        Case 191379
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProdutoAnterior)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191381)

    End Select

    Exit Sub

End Sub

Public Sub GridAlocacao_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAlocacao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAlocacao, iAlterado)
    End If

End Sub

Public Sub GridAlocacao_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridAlocacao)

End Sub

Public Sub GridAlocacao_RowColChange()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_GridAlocacao_RowColChange

    Call Grid_RowColChange(objGridAlocacao)

    If objGridAlocacao.iExecutarFuncao = GRID_EXECUTAR_FUNCAO Then

        'se a linha anterior for diferente entao preenche a decricao do produto selecionado
        If (GridAlocacao.Row <> giLinhaAnterior) Then

            If GridAlocacao.Row > 0 Then

                lErro = CF("Produto_Formata", GridAlocacao.TextMatrix(GridAlocacao.Row, iGrid_ProdutoAlmox_Col), sProduto, iPreenchido)
                If lErro <> SUCESSO Then gError 191382

                objProduto.sCodigo = sProduto

                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 191383

                'Se não achou o Produto --> erro
                If lErro = 28030 Then gError 191384

                ProdutoDescricao.Caption = objProduto.sDescricao

            End If

            'Guarda a Linha corrente
            giLinhaAnterior = GridAlocacao.Row

        End If

    End If

    Exit Sub

Erro_GridAlocacao_RowColChange:

    Select Case gErr

        Case 191382, 191383

        Case 191384
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 191385)

    End Select

    Exit Sub

End Sub

Public Sub GridAlocacao_Scroll()

    Call Grid_Scroll(objGridAlocacao)

End Sub

Private Function Saida_Celula_GridParcelas(objGridInt As AdmGrid) As Long
'Faz a crítica do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridParcelas

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Vencimento_Col
            lErro = Saida_Celula_DataVencimento(objGridInt)
            If lErro <> SUCESSO Then gError 191386

        Case iGrid_ValorParcela_Col
            lErro = Saida_Celula_ValorParcela(objGridInt)
            If lErro <> SUCESSO Then gError 191387

        Case iGrid_TipoPagto_Col
            lErro = Saida_Celula_TipoPagto(objGridInt)
            If lErro <> SUCESSO Then gError 191388

        Case iGrid_Desc1Codigo_Col, iGrid_Desc2Codigo_Col, iGrid_Desc3Codigo_Col
            lErro = Saida_Celula_TipoDesconto(objGridInt)
            If lErro <> SUCESSO Then gError 191389

        Case iGrid_Desc1Ate_Col, iGrid_Desc2Ate_Col, iGrid_Desc3Ate_Col
            lErro = Saida_Celula_DescontoData(objGridInt)
            If lErro <> SUCESSO Then gError 191390

        Case iGrid_Desc1Valor_Col, iGrid_Desc2Valor_Col, iGrid_Desc3Valor_Col
            lErro = Saida_Celula_DescontoValor(objGridInt)
            If lErro <> SUCESSO Then gError 191391

        Case iGrid_Desc1Perc_Col, iGrid_Desc2Perc_Col, iGrid_Desc3Perc_Col
            lErro = Saida_Celula_DescontoPerc(objGridInt)
            If lErro <> SUCESSO Then gError 191392

    End Select

    Saida_Celula_GridParcelas = SUCESSO

    Exit Function

Erro_Saida_Celula_GridParcelas:

    Saida_Celula_GridParcelas = gErr

    Select Case gErr

        Case 191386 To 191392

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191393)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoData(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Data que está deixando de ser a corrente

Dim lErro As Long
Dim iCodigo As Integer

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
        If lErro <> SUCESSO Then gError 191394
        'Se a data de vencimento estiver preenchida
        If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))) > 0 Then
            'critica se DataDesconto ultrapassa DataVencimento
            If CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col)) < CDate(objGridInt.objControle.Text) Then gError 191395
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191396

    Saida_Celula_DescontoData = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoData:

    Saida_Celula_DescontoData = gErr

    Select Case gErr

        Case 191394, 191396
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191395
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADESCONTO_MAIOR_DATAVENCIMENTO", gErr, CDate(objGridInt.objControle.Text), CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col)))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191397)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoPerc(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Percentual que está deixando de ser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim dPercentual As Double
Dim dValorParcela As Double
Dim sValorDesconto As String

On Error GoTo Erro_Saida_Celula_DescontoPerc

    If GridParcelas.Col = iGrid_Desc1Perc_Col Then
        Set objGridInt.objControle = Desconto1Percentual
    ElseIf GridParcelas.Col = iGrid_Desc2Perc_Col Then
        Set objGridInt.objControle = Desconto2Percentual
    ElseIf GridParcelas.Col = iGrid_Desc3Perc_Col Then
        Set objGridInt.objControle = Desconto3Percentual
    End If

    If Len(Trim(objGridInt.objControle.Text)) > 0 Then

        'Critica porcentagem
        lErro = Porcentagem_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 191398

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
    If lErro <> SUCESSO Then gError 191399

    Saida_Celula_DescontoPerc = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoPerc:

    Saida_Celula_DescontoPerc = gErr

    Select Case gErr

        Case 191398, 191399
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191400)

    End Select

    Exit Function

End Function
'
'Private Function Saida_Celula_GridBloqueio(objGridInt As AdmGrid) As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_GridBloqueio
'
'    'Verifica qual a coluna atual do Grid
'    Select Case objGridInt.objGrid.Col
'
'        Case iGrid_TipoBloqueio_Col
'            lErro = Saida_Celula_TipoBloqueio(objGridInt)
'            If lErro <> SUCESSO Then gError 191401
'
'        Case iGrid_ResponsavelBL_Col
'            lErro = Saida_Celula_ResponsavelBL(objGridInt)
'            If lErro <> SUCESSO Then gError 191402
'
'        Case iGrid_Observacao_Col
'            lErro = Saida_Celula_Observacao(objGridInt)
'            If lErro <> SUCESSO Then gError 191403
'
'    End Select
'
'    Saida_Celula_GridBloqueio = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_GridBloqueio:
'
'    Saida_Celula_GridBloqueio = gErr
'
'    Select Case gErr
'
'        Case 191401 To 191403
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191404)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Sub Form_Activate()

Dim lErro As Long

On Error GoTo Erro_Form_Activate

    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError 191405

    Call TelaIndice_Preenche(Me)

    Exit Sub

Erro_Form_Activate:

    Select Case gErr

        Case 191405

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191406)

    End Select

    Exit Sub

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 191407

    'Limpa a Tela
    Call Limpa_Tela_PedidoVenda

    iAlterado = 0
    iComissoesAlterada = 0
    gbCarregandoTela = False

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 191407

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191408)

    End Select

    Exit Sub

End Sub

Public Function Reserva_Processa(ByVal iIndiceGridItens As Integer, dQuantidade As Double, dQuantidadeCancelada As Double, dQuantidadeFaturada As Double) As Long

Dim lErro As Long
Dim iItemAtual As Integer
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim iIndice As Integer

On Error GoTo Erro_Reserva_Processa

    'ALTERADO POR WAGNER
    'NÃO VAI RESERVAR
    'A MOVIMENTAÇÂO DE ESTOQUE SE DARÁ PELO APONTAMENTO


'    'Se não houve Quant a Reservar zera Reservas.
'    If dQuantidade = 0 Or dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada = 0 Then
'        GridItens.TextMatrix(iIndiceGridItens, iGrid_QuantReservada_Col) = Format(0, "Standard")
'        iItemAtual = CInt(GridItens.TextMatrix(iIndiceGridItens, iGrid_ItemProduto_Col))
'    End If
'
'    lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndiceGridItens, iGrid_Produto_Col), sProduto, iPreenchido)
'    If lErro <> SUCESSO Then gError 191407
'
'    objProduto.sCodigo = sProduto
'
'    'Lê o Produto
'    lErro = CF("Produto_Le", objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then gError 191408
'
'    'Se não achou o Produto --> erro
'    If lErro = 28030 Then gError 191409
'
'    If objProduto.iControleEstoque = PRODUTO_CONTROLE_RESERVA And dQuantidade > 0 And dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada > 0 And gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA Then
'
'        lErro = ReservaAlmoxarifadoPadrao(iIndiceGridItens, objProduto, dQuantidade, dQuantidadeCancelada, dQuantidadeFaturada)
'        If lErro <> SUCESSO And lErro <> 26680 Then gError 191410
'
'    ElseIf dQuantidade = 0 And dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada = 0 And gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA Then
'
'        For iIndice = 1 To objGridAlocacao.iLinhasExistentes
'
'            If GridAlocacao.TextMatrix(iIndice, iGrid_ProdutoAlmox_Col) = GridItens.TextMatrix(iIndiceGridItens, iGrid_Produto_Col) Then
'                Call Remove_Reservas(CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)))
'            End If
'        Next
'
'    End If

    Reserva_Processa = SUCESSO

    Exit Function

Erro_Reserva_Processa:

    Reserva_Processa = gErr

    Select Case gErr

        Case 191407, 191408, 191410

        Case 191409
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191411)

    End Select

    Exit Function

End Function
'
'Private Function TributacaoPV_Terminar() As Long
'
'Dim lErro As Long, objTributoDoc As ClassTributoDoc
'
'On Error GoTo Erro_TributacaoPV_Terminar
'
'    'Se gobjPedidoVenda já foi inicializado
'    If Not (gobjPedidoVenda Is Nothing) Then
'        Set objTributoDoc = gobjPedidoVenda
'        lErro = objTributoDoc.Desativar
'        If lErro <> SUCESSO Then gError 191412
'        Set gobjPedidoVenda = Nothing
'    End If
'
'    TributacaoPV_Terminar = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoPV_Terminar:
'
'    TributacaoPV_Terminar = gErr
'
'    Select Case gErr
'
'        Case 191412
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191413)
'
'    End Select
'
'End Function
'
'Private Function TributacaoPV_Reset(Optional objPedidoDeVenda As ClassPedidoDeVenda) As Long
''cria ou atualiza gobjPedidoVenda, com dados correspondentes a objPedidoDeVenda (se este for passado) ou com dados "padrao"
'
'Dim lErro As Long, objTributoDoc As ClassTributoDoc
'
'On Error GoTo Erro_TributacaoPV_Reset
'
'    'se gobjPedidoVenda já foi inicializado
'    If Not (gobjPedidoVenda Is Nothing) Then
'        Set objTributoDoc = gobjPedidoVenda
'        lErro = objTributoDoc.Desativar
'        If lErro <> SUCESSO Then gError 191414
'        Set gobjPedidoVenda = Nothing
'    End If
'
'    'se o pedido de venda veio preenchido
'    If Not (objPedidoDeVenda Is Nothing) Then
'
'        Set gobjPedidoVenda = objPedidoDeVenda
'
'    Else
'
'        Set gobjPedidoVenda = New ClassPedidoDeVenda
'        gobjPedidoVenda.dtDataEmissao = gdtDataAtual
'
'    End If
'
'    Set objTributoDoc = gobjPedidoVenda
'    lErro = objTributoDoc.Ativar
'    If lErro <> SUCESSO Then gError 191415
'
'    giNaturezaOpAlterada = 0
'    giISSAliquotaAlterada = 0
'    giISSValorAlterado = 0
'    giValorIRRFAlterado = 0
'    giTipoTributacaoAlterado = 0
'    giAliqIRAlterada = 0
'    iPISRetidoAlterado = 0
'    iCOFINSRetidoAlterado = 0
'    iCSLLRetidoAlterado = 0
'
'    giNatOpItemAlterado = 0
'    giTipoTributacaoItemAlterado = 0
'    giICMSBaseItemAlterado = 0
'    giICMSPercRedBaseItemAlterado = 0
'    giICMSAliquotaItemAlterado = 0
'    giICMSValorItemAlterado = 0
'    giICMSSubstBaseItemAlterado = 0
'    giICMSSubstAliquotaItemAlterado = 0
'    giICMSSubstValorItemAlterado = 0
'    giIPIBaseItemAlterado = 0
'    giIPIPercRedBaseItemAlterado = 0
'    giIPIAliquotaItemAlterado = 0
'    giIPIValorItemAlterado = 0
'
'    TributacaoPV_Reset = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoPV_Reset:
'
'    TributacaoPV_Reset = gErr
'
'    Select Case gErr
'
'        Case 191414, 191415
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191416)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function CarregaTiposTrib() As Long
'
'Dim lErro As Long, sCodigo As String
'Dim objTipoTribICMS As ClassTipoTribICMS
'Dim objTipoTribIPI As ClassTipoTribIPI
'
'On Error GoTo Erro_CarregaTiposTrib
'
'    lErro = CF("TiposTribICMS_Le_Todos", gcolTiposTribICMS)
'    If lErro <> SUCESSO Then gError 191417
'
'    'Preenche ComboICMSTipo
'    For Each objTipoTribICMS In gcolTiposTribICMS
'
'        sCodigo = Space(STRING_TIPO_ICMS_CODIGO - Len(CStr(objTipoTribICMS.iTipo)))
'        sCodigo = sCodigo & CStr(objTipoTribICMS.iTipo) & SEPARADOR & objTipoTribICMS.sDescricao
'        ComboICMSTipo.AddItem (sCodigo)
'        ComboICMSTipo.ItemData(ComboICMSTipo.NewIndex) = objTipoTribICMS.iTipo
'
'    Next
'
'    lErro = CF("TiposTribIPI_Le_Todos", gcolTiposTribIPI)
'    If lErro <> SUCESSO Then gError 191418
'
'    'Preenche ComboIPITipo
'    For Each objTipoTribIPI In gcolTiposTribIPI
'
'        sCodigo = Space(STRING_TIPO_ICMS_CODIGO - Len(CStr(objTipoTribIPI.iTipo)))
'        sCodigo = sCodigo & CStr(objTipoTribIPI.iTipo) & SEPARADOR & objTipoTribIPI.sDescricao
'        ComboIPITipo.AddItem (sCodigo)
'        ComboIPITipo.ItemData(ComboIPITipo.NewIndex) = objTipoTribIPI.iTipo
'
'    Next
'
'    CarregaTiposTrib = SUCESSO
'
'    Exit Function
'
'Erro_CarregaTiposTrib:
'
'    CarregaTiposTrib = gErr
'
'    Select Case gErr
'
'        Case 191417, 191418
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191419)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Carrega_Tab_Tributacao(objPedidoVenda As ClassPedidoDeVenda) As Long
'
'Dim lErro As Long
'Dim objTributacaoPV As New ClassTributacaoPV
'Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto
'
'On Error GoTo Erro_Carrega_Tab_Tributacao
'
'    giTrazendoTribTela = 1
'
'    Set objTributacaoPV = objPedidoVenda.objTributacaoPV
'
'    If NaturezaOp.Text <> objPedidoVenda.sNaturezaOp Then
'
'        NaturezaOp.Text = objPedidoVenda.sNaturezaOp
'        Call NaturezaOp_Validate(bSGECancelDummy)
'
'    End If
'
'    'no frame de "resumo"
'    objTributacaoTipo.iTipo = objTributacaoPV.iTipoTributacao
'    If objTributacaoTipo.iTipo <> 0 Then
'
'        TipoTributacao.Text = CStr(objTributacaoPV.iTipoTributacao)
'
'        lErro = CF("TipoTributacao_Le", objTributacaoTipo)
'        If lErro <> SUCESSO Then gError 191420
'
'        DescTipoTrib.Caption = objTributacaoTipo.sDescricao
'
'        'se nao incide ISS
'        If objTributacaoTipo.iISSIncide = 0 Then
'            ISSValor.Enabled = False
'            ISSAliquota.Enabled = False
'            ISSIncluso.Enabled = False
'        Else
'            ISSValor.Enabled = True
'            ISSAliquota.Enabled = True
'            ISSIncluso.Enabled = True
'        End If
'
'        'se nao incide IR
'        If objTributacaoTipo.iIRIncide = 0 Then
'            ValorIRRF.Enabled = False
'            IRAliquota.Enabled = False
'        Else
'            ValorIRRF.Enabled = True
'            IRAliquota.Enabled = True
'        End If
'
'        'se nao retem PIS
'        If objTributacaoTipo.iPISRetencao = 0 Then
'            PISRetido.Enabled = False
'        Else
'            PISRetido.Enabled = True
'        End If
'
'        'se nao retem COFINS
'        If objTributacaoTipo.iCOFINSRetencao = 0 Then
'            COFINSRetido.Enabled = False
'        Else
'            COFINSRetido.Enabled = True
'        End If
'
'        'se nao retem CSLL
'        If objTributacaoTipo.iCSLLRetencao = 0 Then
'            CSLLRetido.Enabled = False
'        Else
'            CSLLRetido.Enabled = True
'        End If
'
'    Else
'
'        TipoTributacao.Text = ""
'        DescTipoTrib.Caption = ""
'
'    End If
'
'    IPIBase.Caption = Format(objTributacaoPV.dIPIBase, "Standard")
'    IPIValor.Caption = Format(objTributacaoPV.dIPIValor, "Standard")
'    IPIValor1.Caption = IPIValor.Caption
'    ISSBase.Caption = Format(objTributacaoPV.dISSBase, "Standard")
'    ISSAliquota.Text = CStr(objTributacaoPV.dISSAliquota * 100)
'    ISSValor.Text = CStr(objTributacaoPV.dISSValor)
'    ISSValor1.Caption = Format(ISSValor.Text, "Standard")
'    ISSIncluso.Value = objTributacaoPV.iISSIncluso
'    IRBase.Caption = Format(objTributacaoPV.dIRRFBase, "Standard")
'    IRAliquota.Text = CStr(objTributacaoPV.dIRRFAliquota * 100)
'    ValorIRRF.Text = CStr(objTributacaoPV.dIRRFValor)
'    ICMSBase.Caption = Format(objTributacaoPV.dICMSBase, "Standard")
'    ICMSValor.Caption = Format(objTributacaoPV.dICMSValor, "Standard")
'    ICMSSubstBase.Caption = Format(objTributacaoPV.dICMSSubstBase, "Standard")
'    ICMSSubstValor.Caption = Format(objTributacaoPV.dICMSSubstValor, "Standard")
'    PISRetido.Text = CStr(objTributacaoPV.dPISRetido)
'    COFINSRetido.Text = CStr(objTributacaoPV.dCOFINSRetido)
'    CSLLRetido.Text = CStr(objTributacaoPV.dCSLLRetido)
'
'    'o frame de "detalhamento" vou deixar p/carregar qdo o usuario entrar nele
'
'    giISSAliquotaAlterada = 0
'    giISSValorAlterado = 0
'    giValorIRRFAlterado = 0
'    giTipoTributacaoAlterado = 0
'    giAliqIRAlterada = 0
'    iPISRetidoAlterado = 0
'    iCOFINSRetidoAlterado = 0
'    iCSLLRetidoAlterado = 0
'
'    giTrazendoTribTela = 0
'
'    Carrega_Tab_Tributacao = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Tab_Tributacao:
'
'    giTrazendoTribTela = 0
'
'    Carrega_Tab_Tributacao = gErr
'
'    Select Case gErr
'
'        Case 191420
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191421)
'
'    End Select
'
'End Function
'
'Private Function TributacaoItem_InicializaTab() As Long
''deve ser chamada na entrada do tab de detalhamento dentro do tab de tributacao
'Dim lErro As Long, objItemPedido As ClassItemPedido, sItem As String
'Dim objProduto As New ClassProduto
'
'On Error GoTo Erro_TributacaoItem_InicializaTab
'
'    'preencher o valor de frete, seguro, descontos e outras desp no frameOutros
'    LabelValorFrete.Caption = Format(gobjPedidoVenda.dValorFrete, "Standard")
'    LabelValorDesconto.Caption = Format(gobjPedidoVenda.dValorDesconto, "Standard")
'    LabelValorSeguro.Caption = Format(gobjPedidoVenda.dValorSeguro, "Standard")
'    LabelValorOutrasDespesas.Caption = Format(gobjPedidoVenda.dValorOutrasDespesas, "Standard")
'
'    'esvaziar a combo de itens
'    ComboItensTrib.Clear
'
'    'preencher a combo de itens: com "codigo do produto - descricao"
'    For Each objItemPedido In gobjPedidoVenda.colItensPedidoSRV
'
'        lErro = Mascara_MascararProduto(objItemPedido.sProduto, sItem)
'        If lErro <> SUCESSO Then gError 191422
'
'
'        sItem = sItem & " - " & objItemPedido.sDescricao
'        ComboItensTrib.AddItem sItem
'
'    Next
'
'    'preencher a combo de itens: com "codigo do produto - descricao"
'    For Each objItemPedido In gobjPedidoVenda.colItensPedido
'
'        lErro = Mascara_MascararProduto(objItemPedido.sProduto, sItem)
'        If lErro <> SUCESSO Then gError 191423
'
'
'        sItem = sItem & " - " & objItemPedido.sDescricao
'        ComboItensTrib.AddItem sItem
'
'    Next
'
'    TribSobreItem.Value = True
'    Call TribSobreItem_Click
'
'    TributacaoItem_InicializaTab = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoItem_InicializaTab:
'
'    TributacaoItem_InicializaTab = gErr
'
'    Select Case gErr
'
'        Case 191422, 191423
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191424)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function TributacaoItem_TrazerTela(objTributacaoItemPV As ClassTributacaoItemPV) As Long
''Traz para a tela dados de tributacao de um item
'
'Dim iIndice As Integer, objItemPedido As ClassItemPedido
'Dim lErro As Long, objTipoTribIPI As ClassTipoTribIPI, objTipoTribICMS As ClassTipoTribICMS
'Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto
'Dim objNaturezaOp As New ClassNaturezaOp, sNatOp As String
'
'On Error GoTo Erro_TributacaoItem_TrazerTela
'
'    giTrazendoTribItemTela = 1
'
'    NaturezaOpItem.Text = objTributacaoItemPV.sNaturezaOp
'
'    sNatOp = Trim(NaturezaOpItem.Text)
'
'    If sNatOp <> "" Then
'
'        objNaturezaOp.sCodigo = sNatOp
'        'Lê a Natureza de Operação
'        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
'        If lErro <> SUCESSO And lErro <> 17958 Then gError 191425
'
'        'Se não achou a Natureza de Operação --> erro
'        If lErro <> SUCESSO Then gError 191426
'
'        LabelDescrNatOpItem.Caption = objNaturezaOp.sDescricao
'    Else
'        LabelDescrNatOpItem.Caption = ""
'    End If
'
'    objTributacaoTipo.iTipo = objTributacaoItemPV.iTipoTributacao
'    If objTributacaoTipo.iTipo <> 0 Then
'
'        lErro = CF("TipoTributacao_Le", objTributacaoTipo)
'        If lErro <> SUCESSO Then gError 191427
'
'        TipoTributacaoItem.Text = CStr(objTributacaoItemPV.iTipoTributacao)
'        DescTipoTribItem.Caption = objTributacaoTipo.sDescricao
'
'        'Se não incide IPI
'        If objTributacaoTipo.iIPIIncide = 0 Then
'            ComboIPITipo.Enabled = False
'            IPIBaseItem.Enabled = False
'        Else
'            ComboIPITipo.Enabled = True
'            IPIBaseItem.Enabled = True
'        End If
'
'        'Se não incide ICMS
'        If objTributacaoTipo.iICMSIncide = 0 Then
'            ComboICMSTipo.Enabled = False
'            ICMSBaseItem.Enabled = False
'        Else
'            ComboICMSTipo.Enabled = True
'            ICMSBaseItem.Enabled = True
'        End If
'
'    Else
'
'        TipoTributacaoItem.Text = ""
'        DescTipoTribItem.Caption = ""
'
'    End If
'
'    'Setar dados de ICMS
'    Call Combo_Seleciona_ItemData(ComboICMSTipo, objTributacaoItemPV.iICMSTipo)
'    ICMSBaseItem.Text = CStr(objTributacaoItemPV.dICMSBase)
'    ICMSPercRedBaseItem.Text = CStr(objTributacaoItemPV.dICMSPercRedBase * 100)
'    ICMSAliquotaItem.Text = CStr(objTributacaoItemPV.dICMSAliquota * 100)
'    ICMSValorItem.Text = CStr(objTributacaoItemPV.dICMSValor)
'
'    'setar dados ICMS Substituicao
'    ICMSSubstBaseItem.Text = CStr(objTributacaoItemPV.dICMSSubstBase)
'    ICMSSubstAliquotaItem.Text = CStr(objTributacaoItemPV.dICMSSubstAliquota * 100)
'    ICMSSubstValorItem.Text = CStr(objTributacaoItemPV.dICMSSubstValor)
'
'    For Each objTipoTribICMS In gcolTiposTribICMS
'        If objTipoTribICMS.iTipo = objTributacaoItemPV.iICMSTipo Then Exit For
'    Next
'
'    'Se permite redução de base habilitar este campo
'    If objTipoTribICMS.iPermiteReducaoBase Then
'        ICMSPercRedBaseItem.Enabled = True
'    Else
'        'Desabilita-lo e limpa-lo em caso contrário
'        ICMSPercRedBaseItem.Enabled = False
'    End If
'
'    'Se permite aliquota habilitar este campo e valor.
'    If objTipoTribICMS.iPermiteAliquota Then
'        ICMSAliquotaItem.Enabled = True
'        ICMSValorItem.Enabled = True
'    Else
'        'Desabilitar os dois campos e coloca-los com zero
'        ICMSAliquotaItem.Enabled = False
'        ICMSValorItem.Enabled = False
'    End If
'
'    'Se permite margem de lucro habilitar campos do frame de substituicao
'    If objTipoTribICMS.iPermiteMargLucro Then
'        ICMSSubstBaseItem.Enabled = True
'        ICMSSubstAliquotaItem.Enabled = True
'        ICMSSubstValorItem.Enabled = True
'    Else
'        'Limpa-los e desabilita-los
'        ICMSSubstBaseItem.Enabled = False
'        ICMSSubstAliquotaItem.Enabled = False
'        ICMSSubstValorItem.Enabled = False
'    End If
'
'    'Setar dados de IPI
'    Call Combo_Seleciona_ItemData(ComboIPITipo, objTributacaoItemPV.iIPITipo)
'    IPIBaseItem.Text = CStr(objTributacaoItemPV.dIPIBaseCalculo)
'    IPIPercRedBaseItem.Text = CStr(objTributacaoItemPV.dIPIPercRedBase * 100)
'    IPIAliquotaItem.Text = CStr(objTributacaoItemPV.dIPIAliquota * 100)
'    IPIValorItem.Text = CStr(objTributacaoItemPV.dIPIValor)
'
'    For Each objTipoTribIPI In gcolTiposTribIPI
'        If objTipoTribIPI.iTipo = objTributacaoItemPV.iIPITipo Then Exit For
'    Next
'
'    'Se permite redução de base habilitar este campo
'    If objTipoTribIPI.iPermiteReducaoBase Then
'        IPIPercRedBaseItem.Enabled = True
'    Else
'        'desabilita-lo e limpa-lo em caso contrário
'        IPIPercRedBaseItem.Enabled = False
'    End If
'
'    'Se permite alíquota habilitar este campo e valor.
'    If objTipoTribIPI.iPermiteAliquota Then
'        IPIAliquotaItem.Enabled = True
'        IPIValorItem.Enabled = True
'    Else
'        'Desabilitar os dois campos e coloca-los com zero
'        IPIAliquotaItem.Enabled = False
'        IPIValorItem.Enabled = False
'    End If
'
'    giTrazendoTribItemTela = 0
'    giNatOpItemAlterado = 0
'    giTipoTributacaoItemAlterado = 0
'    giICMSBaseItemAlterado = 0
'    giICMSPercRedBaseItemAlterado = 0
'    giICMSAliquotaItemAlterado = 0
'    giICMSValorItemAlterado = 0
'    giICMSSubstBaseItemAlterado = 0
'    giICMSSubstAliquotaItemAlterado = 0
'    giICMSSubstValorItemAlterado = 0
'    giIPIBaseItemAlterado = 0
'    giIPIPercRedBaseItemAlterado = 0
'    giIPIAliquotaItemAlterado = 0
'    giIPIValorItemAlterado = 0
'
'    TributacaoItem_TrazerTela = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoItem_TrazerTela:
'
'    giTrazendoTribItemTela = 0
'
'    TributacaoItem_TrazerTela = gErr
'
'    Select Case gErr
'
'        Case 191425, 191427
'
'        Case 191426
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", objNaturezaOp.sCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191428)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function AtualizarTributacao() As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_AtualizarTributacao
'
'    If Not (gobjPedidoVenda Is Nothing) Then
'
'        'Atualiza os impostos
'        lErro = gobjTributacao.AtualizaImpostos(gobjPedidoVenda, giRecalculandoTributacao)
'        If lErro <> SUCESSO Then gError 191429
'
'        'joga dados do obj atualizado p/a tela
'        lErro = Carrega_Tab_Tributacao(gobjPedidoVenda)
'        If lErro <> SUCESSO Then gError 191430
'
'    End If
'
'    AtualizarTributacao = SUCESSO
'
'    Exit Function
'
'Erro_AtualizarTributacao:
'
'    AtualizarTributacao = gErr
'
'    Select Case gErr
'
'        Case 191429, 191430
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191431)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Tributacao_GravarTela() As Long
''transfere dados de tributacao da tela para gobjPedidoVenda
''os dados que estiverem diferentes devem ser marcados como "manuais"
'
'Dim lErro As Long
'Dim iIndice As Integer, iTemp As Integer, dTemp As Double, objTributacaoPV As ClassTributacaoPV
'
'On Error GoTo Erro_Tributacao_GravarTela
'
'    Set objTributacaoPV = gobjPedidoVenda.objTributacaoPV
'
'    If gobjPedidoVenda.sNaturezaOp <> NaturezaOp.Text Then
'
'        gobjPedidoVenda.sNaturezaOp = NaturezaOp.Text
'        gobjPedidoVenda.iNaturezaOpManual = VAR_PREENCH_MANUAL
'
'    End If
'
'    iTemp = StrParaInt(TipoTributacao.Text)
'    If iTemp <> objTributacaoPV.iTipoTributacao Then
'        objTributacaoPV.iTipoTributacao = iTemp
'        objTributacaoPV.iTipoTributacaoManual = VAR_PREENCH_MANUAL
'    End If
'
'    'setar dados de ISS
'    iTemp = ISSIncluso.Value
'    If iTemp <> objTributacaoPV.iISSIncluso Then
'        objTributacaoPV.iISSIncluso = iTemp
'        objTributacaoPV.iISSInclusoManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ISSAliquota.Text <> CStr(objTributacaoPV.dISSAliquota * 100) Then
'        dTemp = StrParaDbl(ISSAliquota.Text) / 100
'        If objTributacaoPV.dISSAliquota <> dTemp Then
'            objTributacaoPV.dISSAliquota = dTemp
'            objTributacaoPV.iISSAliquotaManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If ISSValor.Text <> CStr(objTributacaoPV.dISSValor) Then
'        dTemp = StrParaDbl(ISSValor.Text)
'        If objTributacaoPV.dISSValor <> dTemp Then
'            objTributacaoPV.dISSValor = dTemp
'            objTributacaoPV.iISSValorManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    'setar dados de IR
'    If IRAliquota.Text <> CStr(objTributacaoPV.dIRRFAliquota * 100) Then
'        dTemp = StrParaDbl(IRAliquota.Text) / 100
'        If objTributacaoPV.dIRRFAliquota <> dTemp Then
'            objTributacaoPV.dIRRFAliquota = dTemp
'            objTributacaoPV.iIRRFAliquotaManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If ValorIRRF.Text <> CStr(objTributacaoPV.dIRRFValor) Then
'        dTemp = StrParaDbl(ValorIRRF.Text)
'        If objTributacaoPV.dIRRFValor <> dTemp Then
'            objTributacaoPV.dIRRFValor = dTemp
'            objTributacaoPV.iIRRFValorManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If PISRetido.Text <> CStr(objTributacaoPV.dPISRetido) Then
'        dTemp = StrParaDbl(PISRetido.Text)
'        If objTributacaoPV.dPISRetido <> dTemp Then
'            objTributacaoPV.dPISRetido = dTemp
'            objTributacaoPV.iPISRetidoManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If COFINSRetido.Text <> CStr(objTributacaoPV.dCOFINSRetido) Then
'        dTemp = StrParaDbl(COFINSRetido.Text)
'        If objTributacaoPV.dCOFINSRetido <> dTemp Then
'            objTributacaoPV.dCOFINSRetido = dTemp
'            objTributacaoPV.iCOFINSRetidoManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    If CSLLRetido.Text <> CStr(objTributacaoPV.dCSLLRetido) Then
'        dTemp = StrParaDbl(CSLLRetido.Text)
'        If objTributacaoPV.dCSLLRetido <> dTemp Then
'            objTributacaoPV.dCSLLRetido = dTemp
'            objTributacaoPV.iCSLLRetidoManual = VAR_PREENCH_MANUAL
'        End If
'    End If
'
'    Tributacao_GravarTela = SUCESSO
'
'    Exit Function
'
'Erro_Tributacao_GravarTela:
'
'    Tributacao_GravarTela = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191432)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function TributacaoItem_GravarTela(objTributacaoItemPV As ClassTributacaoItemPV) As Long
''transfere dados de tributacao de um item da tela para objTributacaoItemPV
''os dados que estiverem diferentes devem ser marcados como "manuais"
'
'Dim lErro As Long
'Dim iIndice As Integer, iTemp As Integer, dTemp As Double, sTemp As String
'
'On Error GoTo Erro_TributacaoItem_GravarTela
'
'    sTemp = Trim(NaturezaOpItem.Text)
'    If Trim(objTributacaoItemPV.sNaturezaOp) <> sTemp Then
'        objTributacaoItemPV.sNaturezaOp = sTemp
'        objTributacaoItemPV.iNaturezaOpManual = VAR_PREENCH_MANUAL
'    End If
'
'    iTemp = StrParaInt(TipoTributacaoItem.Text)
'    If iTemp <> objTributacaoItemPV.iTipoTributacao Then
'        objTributacaoItemPV.iTipoTributacao = iTemp
'        objTributacaoItemPV.iTipoTributacaoManual = VAR_PREENCH_MANUAL
'    End If
'
'    'Setar dados de ICMS
'
'    iTemp = ComboICMSTipo.ItemData(ComboICMSTipo.ListIndex)
'    If iTemp <> objTributacaoItemPV.iICMSTipo Then
'        objTributacaoItemPV.iICMSTipo = iTemp
'        objTributacaoItemPV.iICMSTipoManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSBaseItem.Text <> CStr(objTributacaoItemPV.dICMSBase) Then
'        dTemp = StrParaDbl(ICMSBaseItem.Text)
'        objTributacaoItemPV.dICMSBase = dTemp
'        objTributacaoItemPV.iICMSBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSPercRedBaseItem.Text <> CStr(objTributacaoItemPV.dICMSPercRedBase * 100) Then
'        dTemp = StrParaDbl(ICMSPercRedBaseItem.Text) / 100
'        objTributacaoItemPV.dICMSPercRedBase = dTemp
'        objTributacaoItemPV.iICMSPercRedBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSAliquotaItem.Text <> CStr(objTributacaoItemPV.dICMSAliquota * 100) Then
'        dTemp = StrParaDbl(ICMSAliquotaItem.Text) / 100
'        objTributacaoItemPV.dICMSAliquota = dTemp
'        objTributacaoItemPV.iICMSAliquotaManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSValorItem.Text <> CStr(objTributacaoItemPV.dICMSValor) Then
'        dTemp = StrParaDbl(ICMSValorItem.Text)
'        objTributacaoItemPV.dICMSValor = dTemp
'        objTributacaoItemPV.iICMSValorManual = VAR_PREENCH_MANUAL
'    End If
'
'    'setar dados ICMS Substituicao
'
'    If ICMSSubstBaseItem.Text <> CStr(objTributacaoItemPV.dICMSSubstBase) Then
'        dTemp = StrParaDbl(ICMSSubstBaseItem.Text)
'        objTributacaoItemPV.dICMSSubstBase = dTemp
'        objTributacaoItemPV.iICMSSubstBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSSubstAliquotaItem.Text <> CStr(objTributacaoItemPV.dICMSSubstAliquota * 100) Then
'        dTemp = StrParaDbl(ICMSSubstAliquotaItem.Text) / 100
'        objTributacaoItemPV.dICMSSubstAliquota = dTemp
'        objTributacaoItemPV.iICMSSubstAliquotaManual = VAR_PREENCH_MANUAL
'    End If
'
'    If ICMSSubstValorItem.Text <> CStr(objTributacaoItemPV.dICMSSubstValor) Then
'        dTemp = StrParaDbl(ICMSSubstValorItem.Text)
'        objTributacaoItemPV.dICMSSubstValor = dTemp
'        objTributacaoItemPV.iICMSSubstValorManual = VAR_PREENCH_MANUAL
'    End If
'
'    'setar dados de IPI
'    iTemp = ComboIPITipo.ItemData(ComboIPITipo.ListIndex)
'    If iTemp <> objTributacaoItemPV.iIPITipo Then
'        objTributacaoItemPV.iIPITipo = iTemp
'        objTributacaoItemPV.iIPITipoManual = VAR_PREENCH_MANUAL
'    End If
'
'    If IPIBaseItem.Text <> CStr(objTributacaoItemPV.dIPIBaseCalculo) Then
'        dTemp = StrParaDbl(IPIBaseItem.Text)
'        objTributacaoItemPV.dIPIBaseCalculo = dTemp
'        objTributacaoItemPV.iIPIBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If IPIPercRedBaseItem.Text <> CStr(objTributacaoItemPV.dIPIPercRedBase * 100) Then
'        dTemp = StrParaDbl(IPIPercRedBaseItem.Text) / 100
'        objTributacaoItemPV.dIPIPercRedBase = dTemp
'        objTributacaoItemPV.iIPIPercRedBaseManual = VAR_PREENCH_MANUAL
'    End If
'
'    If IPIAliquotaItem.Text <> CStr(objTributacaoItemPV.dIPIAliquota * 100) Then
'        dTemp = StrParaDbl(IPIAliquotaItem.Text) / 100
'        objTributacaoItemPV.dIPIAliquota = dTemp
'        objTributacaoItemPV.iIPIAliquotaManual = VAR_PREENCH_MANUAL
'    End If
'
'    If IPIValorItem.Text <> CStr(objTributacaoItemPV.dIPIValor) Then
'        dTemp = StrParaDbl(IPIValorItem.Text)
'        objTributacaoItemPV.dIPIValor = dTemp
'        objTributacaoItemPV.iIPIValorManual = VAR_PREENCH_MANUAL
'    End If
'
'    TributacaoItem_GravarTela = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoItem_GravarTela:
'
'    TributacaoItem_GravarTela = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191433)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function Tributacao_Inclusao_Item_Grid(iLinha As Integer, sProduto As String) As Long
''trata a inclusao de uma linha de item no grid
'Dim lErro As Long
'Dim objTributoDocItem As ClassTributoDocItem
'Dim objItemPedido As ClassItemPedido
'On Error GoTo Erro_Tributacao_Inclusao_Item_Grid
'
'    lErro = Move_GridItem_Memoria(gobjPedidoVenda, iLinha, sProduto)
'    If lErro <> SUCESSO Then gError 191434
'
'    Set objItemPedido = gobjPedidoVenda.colItensPedido.Item(iLinha)
'    Set objTributoDocItem = objItemPedido
'    lErro = objTributoDocItem.Ativar(gobjPedidoVenda)
'    If lErro <> SUCESSO Then gError 191435
'
'    Tributacao_Inclusao_Item_Grid = SUCESSO
'
'    Exit Function
'
'Erro_Tributacao_Inclusao_Item_Grid:
'
'    Tributacao_Inclusao_Item_Grid = gErr
'
'    Select Case gErr
'
'        Case 191434, 191435
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191436)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function Tributacao_Inclusao_Item_GridSRV(iLinha As Integer, sProduto As String) As Long
''trata a inclusao de uma linha de item no grid
'Dim lErro As Long
'Dim objTributoDocItem As ClassTributoDocItem
'Dim objItemPedido As ClassItemPedido
'On Error GoTo Erro_Tributacao_Inclusao_Item_GridSRV
'
'    lErro = Move_GridItem_MemoriaSRV(gobjPedidoVenda, iLinha, sProduto)
'    If lErro <> SUCESSO Then gError 191437
'
'    Set objItemPedido = gobjPedidoVenda.colItensPedidoSRV.Item(iLinha)
'    Set objTributoDocItem = objItemPedido
'    lErro = objTributoDocItem.Ativar(gobjPedidoVenda)
'    If lErro <> SUCESSO Then gError 191438
'
'    Tributacao_Inclusao_Item_GridSRV = SUCESSO
'
'    Exit Function
'
'Erro_Tributacao_Inclusao_Item_GridSRV:
'
'    Tributacao_Inclusao_Item_GridSRV = gErr
'
'    Select Case gErr
'
'        Case 191437, 191438
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191439)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function Move_GridItem_Memoria(objPedidoVenda As ClassPedidoDeVenda, iIndice As Integer, sProduto1 As String) As Long
'Recolhe do Grid os dados do item pedido no parametro

Dim lErro As Long
Dim sProduto As String
Dim objItemPedido As ClassItemPedido, objTributacaoItemPV As ClassTributacaoDocItem
Dim iPreenchido As Integer, objItemPedidoNovo As ClassItemPedido
Dim objDataEntrega As ClassDataEntrega
Dim dQuantidade As Double

On Error GoTo Erro_Move_GridItem_Memoria

    Set objItemPedido = New ClassItemPedido

'    'Verifica se o Produto está preenchido
'    If Len(Trim(sProduto1)) > 0 Then
'
'        'Formata o produto
'        lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError 191440

        objItemPedido.sProduto = sProduto1
'    End If

    If Len(Trim(sProduto1)) = 0 Then gError 191441  'Or iPreenchido = PRODUTO_VAZIO Then gError 191441

    'Armazena os dados do item
    objItemPedido.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)

    objItemPedido.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col))
    objItemPedido.dQuantCancelada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))
    objItemPedido.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))
    objItemPedido.dPrecoTotal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
    objItemPedido.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
    objItemPedido.dtDataEntrega = DATA_NULA
    objItemPedido.dQuantReservada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col))
    objItemPedido.dQuantFaturada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantFaturada_Col))

    If gobjPedidoVenda.colItensPedido.Count >= iIndice Then
        Set objTributacaoItemPV = gobjPedidoVenda.colItensPedido.Item(iIndice).objTributacaoItemPV
    Else
        Set objTributacaoItemPV = Nothing
    End If
    objItemPedido.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
    objItemPedido.sDescricao = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
    objItemPedido.iFilialEmpresa = giFilialEmpresa

    'Adiciona o item na colecao de itens do pedido de venda
    With objItemPedido
        Set objItemPedidoNovo = objPedidoVenda.IncluirItem(.sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .colReserva, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, .sUMEstoque, .iClasseUM, .dQuantFaturada, 0, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, gcolItemPedido(iIndice), objTributacaoItemPV)
    End With

    Move_GridItem_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItem_Memoria:

    Move_GridItem_Memoria = gErr

    Select Case gErr

        Case 191440

        Case 191441
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191443)

    End Select

    Exit Function

End Function

Public Function Move_GridItem_MemoriaSRV(objPedidoVenda As ClassPedidoDeVenda, iIndice As Integer, sProduto1 As String) As Long
'Recolhe do Grid os dados do item pedido no parametro

Dim lErro As Long
Dim sProduto As String
Dim objItemPedido As ClassItemPedido, objTributacaoItemPV As ClassTributacaoDocItem
Dim iPreenchido As Integer, objItemPedidoNovo As ClassItemPedido
Dim objDataEntrega As ClassDataEntrega
Dim dQuantidade As Double

On Error GoTo Erro_Move_GridItem_MemoriaSRV

    Set objItemPedido = New ClassItemPedido

'    'Verifica se o Produto está preenchido
'    If Len(Trim(sProduto1)) > 0 Then
'
'        'Formata o produto
'        lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError 191444

        objItemPedido.sProduto = sProduto1
'    End If

    If Len(Trim(sProduto1)) = 0 Then gError 191445  'Or iPreenchido = PRODUTO_VAZIO Then gError 191445

    'Armazena os dados do item
    objItemPedido.sUnidadeMed = GridItensSRV.TextMatrix(iIndice, iGrid_UnidadeMedSRV_Col)

    objItemPedido.dQuantidade = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_QuantPedidaSRV_Col))
    objItemPedido.dQuantCancelada = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_QuantCancelSRV_Col))
    objItemPedido.dPrecoUnitario = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_ValorUnitarioSRV_Col))
    objItemPedido.dPrecoTotal = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
    objItemPedido.dValorDesconto = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col))
    If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_DataEntregaSRV_Col))) > 0 Then
        objItemPedido.dtDataEntrega = CDate(GridItensSRV.TextMatrix(iIndice, iGrid_DataEntregaSRV_Col))
    Else
        objItemPedido.dtDataEntrega = DATA_NULA
    End If
    objItemPedido.dQuantFaturada = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_QuantFaturadaSRV_Col))

    If gobjPedidoVenda.colItensPedidoSRV.Count >= iIndice Then
        Set objTributacaoItemPV = gobjPedidoVenda.colItensPedidoSRV.Item(iIndice).objTributacaoItemPV
    Else
        Set objTributacaoItemPV = Nothing
    End If
    objItemPedido.dValorDesconto = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col))
    objItemPedido.sDescricao = GridItensSRV.TextMatrix(iIndice, iGrid_DescServicoSRV_Col)
    objItemPedido.iFilialEmpresa = giFilialEmpresa

    'Adiciona o item na colecao de itens do pedido de venda
    With objItemPedido
        Set objItemPedidoNovo = objPedidoVenda.IncluirItemSRV(.sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, .dValorAbatComissao, .dQuantCancelada, .dQuantReservada, .colReserva, .sProdutoNomeReduzido, .sUnidadeMed, .sLote, .sUMEstoque, .iClasseUM, .dQuantFaturada, 0, .sDescricao, .iStatus, .iControleEstoque, .dQuantOP, .dQuantSC, 0, 0, 0, 0, 0, 0, gcolItemPedidoSRV(iIndice), objTributacaoItemPV)
    End With

    Set objItemPedidoNovo.colItemPVEntrega = gobjPedidoVenda.colItensPedidoSRV(iIndice).colItemPVEntrega

    dQuantidade = 0

    For Each objDataEntrega In objItemPedidoNovo.colItemPVEntrega
        dQuantidade = dQuantidade + objDataEntrega.dQuantidade
    Next

    If dQuantidade > 0 Then

        If dQuantidade <> objItemPedidoNovo.dQuantidade - objItemPedidoNovo.dQuantCancelada Then gError 191446

    End If

    Move_GridItem_MemoriaSRV = SUCESSO

    Exit Function

Erro_Move_GridItem_MemoriaSRV:

    Move_GridItem_MemoriaSRV = gErr

    Select Case gErr

        Case 191444

        Case 191445
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 191446
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTSOLICITADA_DIFERE1", gErr, iIndice, sProduto1, objItemPedidoNovo.dQuantidade, dQuantidade)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191447)

    End Select

    Exit Function

End Function
'
'Function Tributacao_Remover_Item_Grid(iLinha As Integer) As Long
''trata a exclusao de uma linha de item no grid
'Dim objItemPedido As ClassItemPedido, objTributoDocItem As ClassTributoDocItem
'
'        Set objItemPedido = gobjPedidoVenda.colItensPedido(iLinha)
'        Set objTributoDocItem = objItemPedido
'        Call objTributoDocItem.Desativar
'        Call gobjPedidoVenda.RemoverItem(iLinha)
'
'End Function
'
'Function Tributacao_Remover_Item_GridSRV(iLinha As Integer) As Long
''trata a exclusao de uma linha de item no grid
'Dim objItemPedido As ClassItemPedido, objTributoDocItem As ClassTributoDocItem
'
'        Set objItemPedido = gobjPedidoVenda.colItensPedidoSRV(iLinha)
'        Set objTributoDocItem = objItemPedido
'        Call objTributoDocItem.Desativar
'        Call gobjPedidoVenda.RemoverItemSRV(iLinha)
'
'End Function
'
'
'Function Tributacao_Alteracao_Item_Grid(iIndice As Integer) As Long
''trata a alteracao de uma linha de item no grid
'
'Dim lErro As Long, sProduto As String, iPreenchido As Integer
'Dim objItemPedido As ClassItemPedido
'
'On Error GoTo Erro_Tributacao_Alteracao_Item_Grid
'
'    Set objItemPedido = gobjPedidoVenda.colItensPedido.Item(iIndice)
'
'    If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 Then
'
'        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError 191161
'
'        objItemPedido.sProduto = sProduto
'    End If
'
'    objItemPedido.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
'
'    objItemPedido.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col))
'    objItemPedido.dQuantCancelada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))
'    objItemPedido.dPrecoTotal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'    objItemPedido.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
'
'    Tributacao_Alteracao_Item_Grid = SUCESSO
'
'    Exit Function
'
'Erro_Tributacao_Alteracao_Item_Grid:
'
'    Tributacao_Alteracao_Item_Grid = gErr
'
'    Select Case gErr
'
'        Case 191161
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191162)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Tributacao_Alteracao_Item_GridSRV(iIndice As Integer) As Long
''trata a alteracao de uma linha de item no grid
'
'Dim lErro As Long, sProduto As String, iPreenchido As Integer
'Dim objItemPedido As ClassItemPedido
'
'On Error GoTo Erro_Tributacao_Alteracao_Item_GridSRV
'
'    Set objItemPedido = gobjPedidoVenda.colItensPedidoSRV.Item(iIndice)
'
'    If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col))) > 0 Then
'
'        lErro = CF("Produto_Formata", GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col), sProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError 191159
'
'        objItemPedido.sProduto = sProduto
'    End If
'
'    objItemPedido.sUnidadeMed = GridItensSRV.TextMatrix(iIndice, iGrid_UnidadeMedSRV_Col)
'
'    objItemPedido.dQuantidade = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_QuantPedidaSRV_Col))
'    objItemPedido.dQuantCancelada = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_QuantCancelSRV_Col))
'    objItemPedido.dPrecoTotal = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))
'    objItemPedido.dValorDesconto = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col))
'
'    Tributacao_Alteracao_Item_GridSRV = SUCESSO
'
'    Exit Function
'
'Erro_Tributacao_Alteracao_Item_GridSRV:
'
'    Tributacao_Alteracao_Item_GridSRV = gErr
'
'    Select Case gErr
'
'        Case 191159
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191160)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Recolhe_Reserva(iItem As Integer, colReserva As colReserva) As Long
'Recolhe em colReserva as reservas do item passado em iItem como parâmetro

Dim lErro As Long
Dim iIndice As Integer
Dim iItemLinha As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim dtData As Date

On Error GoTo Erro_Recolhe_Reserva

    'Percorre o grid de reservas
    For iIndice = 1 To objGridAlocacao.iLinhasExistentes
        'Guarda o item
        iItemLinha = StrParaInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))
        'Se o item da reserva for o mesmo passado por parâmetro
        If iItemLinha = iItem Then

            objAlmoxarifado.sNomeReduzido = GridAlocacao.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
            'Lê o almoxarifado
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> 25060 And lErro <> SUCESSO Then gError 191448
            
            If lErro = 25060 Then gError 191449 'Não encontrou

            If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Validade_Col))) > 0 Then
                dtData = CDate(GridAlocacao.TextMatrix(iIndice, iGrid_Validade_Col))
            Else
                dtData = DATA_NULA

            End If
            'Adiciona os dados da reserva na coleção de reservas
            colReserva.Add giFilialEmpresa, 0, "", objAlmoxarifado.iCodigo, 0, 0, 0, StrParaDbl(GridAlocacao.TextMatrix(iIndice, iGrid_QuantReserv_Col)), DATA_NULA, dtData, gsUsuario, GridAlocacao.TextMatrix(iIndice, iGrid_Responsavel_Col), 0, objAlmoxarifado.sNomeReduzido

        End If
    Next

    Recolhe_Reserva = SUCESSO

    Exit Function

Erro_Recolhe_Reserva:

    Recolhe_Reserva = gErr

    Select Case gErr

        Case 191448

        Case 191449
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191450)

    End Select

    Exit Function

End Function

Private Function Preenche_GridAlocacao(colItensPV As colItemPedido) As Long
'Preenche o grid de reserva com as reservas dos itens passados em colItensPV

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim dFator As Double
Dim objReserva As ClassReserva
Dim objProduto As New ClassProduto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objParcPV As ClassParcelaPedidoVenda

On Error GoTo Erro_Preenche_GridAlocacao

    iIndice1 = 0

    'Limpa o grid de reservas
    Call Grid_Limpa(objGridAlocacao)

    'Percorre a coleção de itens
    For iIndice = 1 To colItensPV.Count

        '******** COMPARAÇÃO INCLUÍDA PARA TRATAMENTO DA GRADE ********
        If colItensPV.Item(iIndice).iPossuiGrade = MARCADO Then
            Call Atualiza_Grid_Alocacao(colItensPV.Item(iIndice))

            objProduto.sCodigo = colItensPV.Item(iIndice).sProduto
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 191463
            
            If lErro <> SUCESSO Then gError 191464

            If objProduto.iKitVendaComp <> MARCADO Then
                Call Atualiza_Grid_Itens(colItensPV(iIndice))
            End If
        Else

            objProduto.sCodigo = colItensPV.Item(iIndice).sProduto
            
            'Lê o produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 191465
            
            If lErro = 28030 Then gError 191466 'Não encontrou
            
            'Faz a conversão da unidade do item para a unidade de estoque
            lErro = CF("UM_Conversao", objProduto.iClasseUM, colItensPV.Item(iIndice).sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 191467
            
            'Para cada reserva do item
            For Each objReserva In colItensPV.Item(iIndice).colReserva

                'Coloca os dados da reserva no grid
                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Item_Col) = colItensPV.Item(iIndice).iItem
                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_ProdutoAlmox_Col) = GridItens.TextMatrix(colItensPV.Item(iIndice).iItem, iGrid_Produto_Col)
                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_UMEstoque_Col) = colItensPV.Item(iIndice).sUMEstoque
                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objReserva.sAlmoxarifado
                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReserv_Col) = Formata_Estoque(objReserva.dQuantidade)
                If objReserva.dtDataValidade <> DATA_NULA Then GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Validade_Col) = Format(objReserva.dtDataValidade, "dd/mm/yyyy")
                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Responsavel_Col) = objReserva.sResponsavel
                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReservar_Col) = Formata_Estoque((colItensPV.Item(iIndice).dQuantidade - colItensPV.Item(iIndice).dQuantCancelada) * dFator)

                objAlmoxarifado.iCodigo = objReserva.iAlmoxarifado

                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError 191468
                
                If lErro = 25056 Then gError 191469

                GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido

                objGridAlocacao.iLinhasExistentes = objGridAlocacao.iLinhasExistentes + 1

            Next
        End If
    Next

    Preenche_GridAlocacao = SUCESSO

    Exit Function

Erro_Preenche_GridAlocacao:

    Preenche_GridAlocacao = gErr

    Select Case gErr

        Case 191463, 191465, 191467, 191468

        Case 191464, 191466
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 191469
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191470)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 191471

    'Limpa a Tela
    Call Limpa_Tela_PedidoVenda

    iAlterado = 0
    iComissoesAlterada = 0
    gbCarregandoTela = False

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 191471

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 191472)

    End Select

    Exit Sub

End Sub

Public Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim dValor As Double

On Error GoTo Erro_BotaoImprimir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 191473
    If Len(Trim(Cliente.Text)) = 0 Then gError 191474
    If Len(Trim(Filial.Text)) = 0 Then gError 191475
    If Len(Trim(FilialEntrega.Text)) = 0 Then gError 191476
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 191477
    'If Len(Trim(NaturezaOp.Text)) = 0 Then gError 191478

    lErro = Move_Pedido_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then gError 191479

    lErro = CF("PedidoServico_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 188828 Then gError 191480

    If lErro = 188828 Then gError 191481

    lErro = objRelatorio.ExecutarDireto("Pedido de Serviço", "PedidoServico = @NPSDE ", 1, "PedServ", "NPSDE", objPedidoVenda.lCodigo, "NPSATE", objPedidoVenda.lCodigo)
    If lErro <> SUCESSO Then gError 191482

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 191473
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)

        Case 191474
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 191475
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)

        Case 191476
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)

        Case 191477
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 191478
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 191479, 191480, 191482

        Case 191481
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_SERVICO_NAO_CADASTRADO", gErr, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 191483)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim dValorTotal As Double
Dim dValor As Double, bExibirBloqsPV As Boolean
Dim vbMsgRes As VbMsgBoxResult, colBloqPVAux As New Collection, objBloqPV As ClassBloqueioGen, objBloqPVAux As ClassBloqueioGen
Dim objPedidoServico As New ClassPedidoServico
Dim objPedidoVenda_BD As New ClassPedidoDeVenda

On Error GoTo Erro_Gravar_Registro

    If Len(Trim(Codigo.Text)) = 0 Then gError 193007
    If Len(Trim(Cliente.Text)) = 0 Then gError 193008
    If Len(Trim(Filial.Text)) = 0 Then gError 193009
    If Len(Trim(FilialEntrega.Text)) = 0 Then gError 193010
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 193011
    If Len(Trim(gobjTribTab.sNatOpInterna)) = 0 Then gError 193012

'    If giTipoVersao = VERSAO_FULL Then
        If Len(Trim(FilialFaturamento.Text)) = 0 Then gError 193013
'    End If

    dValor = CDbl(ValorTotal.Caption)

    If dValor < 0 Then gError 193014

    If Len(Trim(PesoBruto.Text)) > 0 And Len(Trim(PesoLiquido.Text)) > 0 Then
        If StrParaDbl(PesoLiquido.Text) > StrParaDbl(PesoBruto.Text) Then gError 193015
    End If

    lErro = CF("PSRV_PedidoRepr_Valida", StrParaDbl(Codigo.Text), giFilialEmpresa, PedidoRepr.Text)
    If lErro <> SUCESSO Then gError 193461

    lErro = Valida_Grid_ItensSRV()
    If lErro <> SUCESSO Then gError 193016

    lErro = Valida_Grid_Itens()
    If lErro <> SUCESSO Then gError 193017

    lErro = Valida_Grid_Parcelas()
    If lErro <> SUCESSO Then gError 193018

    lErro = objTabComissoes.Valida_Grid_Comissoes()
    If lErro <> SUCESSO Then gError 193019

    lErro = gobjTelaTrataBloqGen.Valida_Grid_Bloqueio()
    If lErro <> SUCESSO Then gError 193020
'
'    lErro = Valida_Tributacao_Gravacao()
'    If lErro <> SUCESSO Then gError 193021

    lErro = Move_Pedido_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then gError 193022

    objPedidoVenda_BD.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
    objPedidoVenda_BD.lCodigo = objPedidoVenda.lCodigo

    lErro = CF("PedidoServico_Le", objPedidoVenda_BD)
    If lErro <> SUCESSO And lErro <> 188828 Then gError 195757
    
    If objPedidoVenda_BD.iStatus = STATUS_BAIXADO Then gError 195758

    lErro = CF("PedidoVenda_Verifica_Preenchimento", objPedidoVenda)
    If lErro <> SUCESSO Then gError 193023

    'guarda em colecao os bloqueios automaticos nao liberados que haviam antes da gravacao
    For Each objBloqPV In objPedidoVenda.colBloqueiosPV

        If objBloqPV.dtDataLib = DATA_NULA And objBloqPV.sResponsavel = BLOQUEIO_AUTO_RESP Then

            Set objBloqPVAux = New ClassBloqueioPV
            With objBloqPVAux
                objBloqPVAux.iTipoDeBloqueio = objBloqPV.iTipoDeBloqueio
                objBloqPVAux.iSequencial = objBloqPV.iSequencial
            End With

            Call colBloqPVAux.Add(objBloqPVAux)

        End If

    Next

    lErro = Trata_Alteracao(objPedidoServico, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> SUCESSO Then gError 193024

'    lErro = gobjTelaProjetoInfo.Valida_Dados
'    If lErro <> SUCESSO Then gError 193025
    
    lErro = gobjTribTab.Valida_Dados()
    If lErro <> SUCESSO Then gError 181786

    'Cyntia
    GL_objMDIForm.MousePointer = vbHourglass

    'Grava no BD
    lErro = CF("PedidoServico_Grava", objPedidoVenda)
    If lErro <> SUCESSO Then gError 193026

    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then

        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then

            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_PEDIDOSERVICO
            gobjAnotacao.sID = CStr(objPedidoVenda.iFilialEmpresa) & "," & CStr(objPedidoVenda.lCodigo)
            gobjAnotacao.dtDataAlteracao = gdtDataHoje

            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 193027

        End If

    End If
    
    lErro = gobjTelaTrataBloqGen.Trata_Bloqueios_Gravacao(objPedidoVenda)
    If lErro <> SUCESSO Then gError 193028

'    Call PedidoVenda_ExibirBloqs(objPedidoVenda, colBloqPVAux, bExibirBloqsPV)
'    If bExibirBloqsPV Then
'
'        lErro = Chama_Tela("BloqPSRV_ConsultaTodos", objPedidoVenda)
'        If lErro <> SUCESSO Then gError 193028
'
'    End If

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 193007
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)

        Case 193008
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 193009
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)

        Case 193010
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)

        Case 193011
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 193012
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 193013
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FATURAMENTO_NAO_PREENCHIDA", gErr)

        Case 193014
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_PSRV_NEGATIVO", gErr)

        Case 193015
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr)

        Case 193016 To 193028, 193461, 195757
        
        Case 195758
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_SERVICO_BAIXADO", gErr, objPedidoVenda_BD.iFilialEmpresa, objPedidoVenda_BD.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193029)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Itens() As Long

Dim iIndice As Integer
Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Valida_Grid_Itens

    'para cada item do grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col))) = 0 Then gError 193030

        lErro = Valor_Positivo_Critica(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col))
        If lErro <> SUCESSO Then gError 193031

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) = 0 Then gError 193032

        lErro = Valor_Positivo_Critica(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))
        If lErro <> SUCESSO Then gError 193033

    Next

    Valida_Grid_Itens = SUCESSO

    Exit Function

Erro_Valida_Grid_Itens:

    Valida_Grid_Itens = gErr

    Select Case gErr

        Case 193030
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA_SRV2", gErr, iIndice)

        Case 193031, 193033

        Case 193032
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO_SRV2", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193034)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_ItensSRV() As Long

Dim iIndice As Integer
Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Valida_Grid_ItensSRV

    'Verifica se há itens no grid
    If objGridItensSRV.iLinhasExistentes = 0 Then gError 193035

    'para cada item do grid
    For iIndice = 1 To objGridItensSRV.iLinhasExistentes

        If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_QuantPedidaSRV_Col))) = 0 Then gError 193036

        lErro = Valor_Positivo_Critica(GridItensSRV.TextMatrix(iIndice, iGrid_QuantPedidaSRV_Col))
        If lErro <> SUCESSO Then gError 193037

        If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_ValorUnitarioSRV_Col))) = 0 Then gError 193038

        lErro = Valor_Positivo_Critica(GridItensSRV.TextMatrix(iIndice, iGrid_ValorUnitarioSRV_Col))
        If lErro <> SUCESSO Then gError 193039

    Next

    Valida_Grid_ItensSRV = SUCESSO

    Exit Function

Erro_Valida_Grid_ItensSRV:

    Valida_Grid_ItensSRV = gErr

    Select Case gErr

        Case 193035
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_PSRV", gErr)

        Case 193036
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA_SRV1", gErr, iIndice)

        Case 193037, 193039
        
        Case 193038
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO_SRV1", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193040)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Parcelas() As Long
'Valida os dados do Grid de Parcelas

Dim lErro As Long
Dim iIndice As Integer
Dim dSomaParcelas As Double
'Dim dValorIRRF As Double, dPISRetido As Double, dCOFINSRetido As Double, dCSLLRetido As Double
'Dim dValorTotal As Double
Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim iTamanho As Integer
Dim iTipo As Integer
Dim dPercAcrecFin As Double
Dim iDesconto As Integer
Dim dtDataDesconto As Date
Dim iTipoPagto As Integer
Dim dValorPagar As Double

On Error GoTo Erro_Valida_Grid_Parcelas

    'Verifica se alguma parcela foi informada
    If objGridParcelas.iLinhasExistentes = 0 Then gError 193041

    dSomaParcelas = 0

    'Para cada Parcela do grid de parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        dtDataEmissao = StrParaDate(DataEmissao.Text)
        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) = 0 Then gError 193042
        
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) = 0 Then gError 193043

        If dtDataVencimento < dtDataEmissao Then gError 193044

        'Se o tipo de desconto 1 estiver preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))) > 0 Then
            iDesconto = 1
            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
            'Verifica se a data de desconto está preenchdida
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))) = 0 Then gError 193045
            'Recolhe o Valor ou Percentual de desconto
            If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))) = 0 Then gError 193046
            Else
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col))) = 0 Then gError 193047
            End If
            'Se o tipo de desconto 2 estiver preenchido
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))) > 0 Then
                iDesconto = 2
                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
                'Verifica se a data de desconto está preenchdida
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))) = 0 Then gError 193048
                'Faz a crítica da ordem das datas de desconto
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 193049
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 193050
                'Recolhe o Valor ou Percentual de desconto
                If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))) = 0 Then gError 193051
                Else
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col))) = 0 Then gError 193052
                End If
                'Se o tipo de desconto 3 estiver preenchido
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))) > 0 Then
                    iDesconto = 3
                    iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
                    'Verifica se a data de desconto está preenchdida
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))) = 0 Then gError 193053
                    'Faz a crítica da ordem das datas de desconto
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 193054
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 193055
                    'Recolhe o Valor ou Percentual de desconto
                    If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))) = 0 Then gError 193056
                    Else
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col))) = 0 Then gError 193057
                    End If
                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
                    If dtDataDesconto > dtDataVencimento Then gError 193058
                Else
                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
                    If dtDataDesconto > dtDataVencimento Then gError 193059
                End If
            Else
                dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
                If dtDataDesconto > dtDataVencimento Then gError 193060
            End If

        End If

        If iIndice > 1 Then If CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col)) < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_Col)) Then gError 193061

        dSomaParcelas = dSomaParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_TipoPagto_Col))) > 0 Then

            iTipoPagto = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_TipoPagto_Col))

            lErro = CF("Valida_TipoPagto", iIndice, iTipoPagto, gcolParcPV(iIndice))
            If lErro <> SUCESSO Then gError 193062

        End If

    Next

'    dValorTotal = StrParaDbl(ValorTotal.Caption)
'    dValorIRRF = StrParaDbl(ValorIRRF.Text)
'    If Len(Trim(PISRetido.Text)) <> 0 And IsNumeric(PISRetido.Text) Then dPISRetido = CDbl(PISRetido.Text)
'    If Len(Trim(COFINSRetido.Text)) <> 0 And IsNumeric(COFINSRetido.Text) Then dCOFINSRetido = CDbl(COFINSRetido.Text)
'    If Len(Trim(CSLLRetido.Text)) <> 0 And IsNumeric(CSLLRetido.Text) Then dCSLLRetido = CDbl(CSLLRetido.Text)
'
'    If Format((dValorTotal - (dValorIRRF + dPISRetido + dCOFINSRetido + dCSLLRetido)), "Standard") <> Format(dSomaParcelas, "Standard") Then gError 193063

    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then gError 193062
    
    If Abs(dValorPagar - dSomaParcelas) > DELTA_VALORMONETARIO Then gError 193063

    Valida_Grid_Parcelas = SUCESSO

    Exit Function

Erro_Valida_Grid_Parcelas:

    Valida_Grid_Parcelas = gErr

    Select Case gErr

        Case 193041
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_PARCELA_COBRANCA", gErr)

        Case 193042
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_NAO_INFORMADA", gErr, iIndice)

        Case 193043
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PARCELA_COBRANCA_NAO_INFORMADO", gErr, iIndice)

        Case 193044
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_MENOR", gErr, iIndice, dtDataVencimento, dtDataEmissao)

        Case 193045, 193048, 193053
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_NAO_PREENCHIDA", gErr, iDesconto, iIndice)

        Case 193046, 193047, 193051, 193052, 193056, 193057
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_PARCELA_NAO_PREENCHIDO", gErr, iDesconto, iIndice)
        
        Case 193049, 193054
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTOS_DESORDENADAS", gErr, iIndice)

        Case 193050, 193055
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTO_IGUAIS", gErr, iDesconto - 1, iDesconto, iIndice)

        Case 193058, 193059, 193060
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESC_PARCELA_SUP_DATA_VENCIMENTO", gErr, dtDataDesconto, iDesconto, iIndice)

        Case 193061
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_COBRANCA_NAO_ORDENADA", gErr)

        Case 193062

        Case 193063
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_COBRANCA_INVALIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193064)

    End Select

    Exit Function

End Function
'
'Private Function Valida_Grid_Bloqueio() As Long
''Validação de GridBloqueio
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim bBloqueioTotal As Boolean
'Dim iTipoBloqueio As Integer
'Dim vbMsgRes As VbMsgBoxResult
'Dim dtDataLiberacao As Date
'
'On Error GoTo Erro_Valida_Grid_Bloqueio
'
'    bBloqueioTotal = False
'
'    If objGridBloqueio.iLinhasExistentes > giGridBloqueioLinhasBloqueadas Then
'
'        For iIndice = giGridBloqueioLinhasBloqueadas + 1 To objGridBloqueio.iLinhasExistentes
'
'            'Testa se preencheu tipo de bloqueio
'            If Len(Trim(GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))) = 0 Then gError 193065
'
'        Next
'
'    End If
'
'    If objGridBloqueio.iLinhasExistentes > 0 Then
'
'        'Testa se existe bloqueio total não liberado
'        For iIndice = giGridBloqueioLinhasBloqueadas + 1 To objGridBloqueio.iLinhasExistentes
'
'            iTipoBloqueio = Codigo_Extrai(GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))
'            dtDataLiberacao = StrParaDate(GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col))
'
'            If iTipoBloqueio = BLOQUEIO_TOTAL And dtDataLiberacao = DATA_NULA Then bBloqueioTotal = True
'
'        Next
'
'        If bBloqueioTotal Then
'
'            'Pergunta se deseja continuar com bloqueio total existente
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_BLOQUEIO_TOTAL")
'            If vbMsgRes = vbNo Then gError 193066
'
'            'Se continuar com bloqueio total limpa reservas
'            Call Grid_Limpa(objGridAlocacao)
'
'            'Zera QuantReservadas no GridItens
'            For iIndice = 1 To objGridItens.iLinhasExistentes
'                GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col) = ""
'            Next
'
'        End If
'
'    End If
'
'    Valida_Grid_Bloqueio = SUCESSO
'
'    Exit Function
'
'Erro_Valida_Grid_Bloqueio:
'
'    Valida_Grid_Bloqueio = gErr
'
'    Select Case gErr
'
'        Case 193065
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_BLOQUEIO_NAO_INFORMADO", gErr, iIndice)
'
'        Case 193066
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193067)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Sub BotaoBaixarPedido_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoBaixarPedido_Click

    GL_objMDIForm.MousePointer = vbHourglass

    lErro = PedidoVenda_Critica_Le(objPedidoVenda)
    If lErro <> SUCESSO Then gError 193073

    lErro = CF("PedidoServico_Baixar_Pedido", objPedidoVenda)
    If lErro <> SUCESSO Then gError 193074

    Call Limpa_Tela_PedidoVenda

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoBaixarPedido_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 193073, 193074

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193075)

    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim vbMsgRes As VbMsgBoxResult
Dim lCodigo As Long

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o código foi informado
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 193076

    objPedidoVenda.lCodigo = CLng(Codigo.Text)
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa
    Set objPedidoVenda.objTela = Me

    'Verifica se é um pedido baixado
    lErro = CF("PedidoServico_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 188832 Then gError 193077

    If lErro <> SUCESSO Then gError 193078
    
    'Se for um Pedido Baixado --> erro
    If objPedidoVenda.iStatus = STATUS_BAIXADO Then gError 193079

    'Pede a confirmação da exclusão do Pedido de Venda
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PEDIDO_SRV", objPedidoVenda.lCodigo)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If


    'Faz a exclusão do Pedido de Venda
    lErro = CF("PedidoServico_Exclui", objPedidoVenda)
    If lErro <> SUCESSO Then gError 193080

    'Limpa a Tela de Pedido de Venda
    Call Limpa_Tela_PedidoVenda

    iAlterado = 0
    iComissoesAlterada = 0
    gbCarregandoTela = False

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 193076
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)

        Case 193077, 193080

        Case 193078
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_SERVICO_NAO_CADASTRADO", gErr, objPedidoVenda.lCodigo)

        Case 193079
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOSERVICO_BAIXADO", gErr, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193081)

    End Select

    Exit Sub

End Sub

Private Function Trata_TabelaPreco() As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Trata_TabelaPreco

    If Not gbCarregandoTela Then

        For iLinha = 1 To objGridItens.iLinhasExistentes

            lErro = Trata_TabelaPreco_Item(iLinha)
            If lErro <> SUCESSO Then gError 193184

        Next

        For iLinha = 1 To objGridItensSRV.iLinhasExistentes

            lErro = Trata_TabelaPreco_ItemSRV(iLinha)
            If lErro <> SUCESSO Then gError 193185

        Next

        'Calcula o Valor Total da Nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 193186

    End If

    Trata_TabelaPreco = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco:

    Trata_TabelaPreco = gErr

    Select Case gErr

        Case 193184 To 193186

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193187)

    End Select

    Exit Function

End Function

Private Function Trata_TabelaPreco_Item(iLinha As Integer) As Long
'faz tratamento de tabela de preço para um ítem (produto)

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Trata_TabelaPreco_Item

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 193179

    If iPreenchido <> PRODUTO_VAZIO And Len(Trim(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col))) > 0 Then

        objTabelaPrecoItem.sCodProduto = sProduto
        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa

        'Lê a Tabela preço para filialEmpresa
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 193180

        'Se não encontrar
        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
            'Lê a Tabela de Preço a nível de Empresa toda
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 193181

        End If

        'Se  conseguir ler a Tabela de Preços
        If lErro = SUCESSO Then
            
            'Calcula o Preco Unitário do Ítem
            lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 193182
            'Coloca no Grid
            If dPrecoUnitario > 0 Then
                GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            Else
                GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col) = ""
            End If

            'Calcula o Preco Total do Ítem
            Call CF2(Me, "PSRV_PrecoTotal_Calcula", iLinha)

            ''*** incluidos p/tratamento de tributacao *******************************
            lErro = gobjTribTab.Alteracao_Item_Grid(iLinha)
            If lErro <> SUCESSO Then gError 193179
            '*** fim tributacao
         End If

    End If

    Trata_TabelaPreco_Item = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco_Item:

    Trata_TabelaPreco_Item = gErr

    Select Case gErr

        Case 193179 To 193182

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193183)

    End Select

    Exit Function

End Function

Private Function Trata_TabelaPreco_ItemSRV(iLinha As Integer) As Long
'faz tratamento de tabela de preço para um ítem (produto)

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Trata_TabelaPreco_ItemSRV

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", GridItensSRV.TextMatrix(iLinha, iGrid_ServicoSRV_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 193174

    If iPreenchido <> PRODUTO_VAZIO And Len(Trim(GridItensSRV.TextMatrix(iLinha, iGrid_UnidadeMedSRV_Col))) > 0 Then

        objTabelaPrecoItem.sCodProduto = sProduto
        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa

        'Lê a Tabela preço para filialEmpresa
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 193175

        'Se não encontrar
        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
            'Lê a Tabela de Preço a nível de Empresa toda
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 193176

        End If

        'Se  conseguir ler a Tabela de Preços
        If lErro = SUCESSO Then
            
            'Calcula o Preco Unitário do Ítem
            lErro = PrecoUnitario_Calcula(GridItensSRV.TextMatrix(iLinha, iGrid_UnidadeMedSRV_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 193177
            'Coloca no Grid
            If dPrecoUnitario > 0 Then
                GridItensSRV.TextMatrix(iLinha, iGrid_ValorUnitarioSRV_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            Else
                GridItens.TextMatrix(iLinha, iGrid_ValorUnitarioSRV_Col) = ""
            End If

            'Calcula o Preco Total do Ítem
            Call CF2(Me, "PSRV_PrecoTotal_CalculaSRV", iLinha)

            ''*** incluidos p/tratamento de tributacao *******************************
            lErro = gobjTribTab.Alteracao_Item_Grid(iLinha, True)
            If lErro <> SUCESSO Then gError 193179
            '*** fim tributacao

         End If

    End If

    Trata_TabelaPreco_ItemSRV = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco_ItemSRV:

    Trata_TabelaPreco_ItemSRV = gErr

    Select Case gErr

        Case 193174 To 193177

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193178)

    End Select

    Exit Function

End Function

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

    '###################################
    'Inserido por Wagner 27/09/05
    If objFilialCliente.iCodMensagem > 0 Then
        objMensagem.iCodigo = objFilialCliente.iCodMensagem
        'Lê a Mensagem "default" p/ o Cliente
        lErro = CF("Mensagem_Le", objMensagem)
        If lErro <> SUCESSO And lErro <> 19234 Then gError 193195

        If lErro = SUCESSO Then Mensagem.Text = objMensagem.sDescricao

    End If
    '###################################

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

Private Function Carrega_FilialFaturamento() As Long
'Carrega FilialFaturamento com as Filiais Empresas

Dim lErro As Long
Dim objFiliais As AdmFiliais

On Error GoTo Erro_Carrega_FilialFaturamento

    For Each objFiliais In gcolFiliais

        If objFiliais.iCodFilial <> EMPRESA_TODA Then

            FilialFaturamento.AddItem CStr(objFiliais.iCodFilial) & SEPARADOR & objFiliais.sNome
            FilialFaturamento.ItemData(FilialFaturamento.NewIndex) = objFiliais.iCodFilial

        End If

    Next

    'Seleciona a Filial na qual o usuário entrou no Sistema
    Call Seleciona_FilialEmpresa

    Carrega_FilialFaturamento = SUCESSO

    Exit Function

Erro_Carrega_FilialFaturamento:

    Carrega_FilialFaturamento = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193197)

    End Select

    Exit Function

End Function

Public Sub FilialFaturamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_FilialFaturamento_Validate

    'Se não estiver preenchida ou alterada pula a crítica
    If Len(Trim(FilialFaturamento.Text)) = 0 Or giFilialFatAlterada = 0 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialFaturamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 193198

    'Nao encontrou o item com o código informado
    If lErro = 6730 Then gError 193199

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 193200

    giFilialFatAlterada = 0

    Exit Sub

Erro_FilialFaturamento_Validate:

    Cancel = True


    Select Case gErr

        Case 193198

        Case 193199
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, iCodigo)

        Case 193200
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialFaturamento.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193201)

    End Select

    Exit Sub

End Sub

Private Sub Seleciona_FilialEmpresa()

Dim lErro As Long
Dim iIndice As Integer
Dim iFilialFaturamento As Integer

On Error GoTo Erro_Seleciona_FilialEmpresa

'    If giTipoVersao = VERSAO_FULL Then

        iFilialFaturamento = gobjFAT.iFilialFaturamento

'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'        iFilialFaturamento = giFilialEmpresa
'    End If

''    If giTipoVersao = VERSAO_FULL Then
''
''        'Lê a filial de faturamento
''        lErro = CF("FilialFaturamento_Le",giFilialEmpresa, iFilialFaturamento)
''        If lErro <> SUCESSO Then gError 51139
''
''    ElseIf giTipoVersao = VERSAO_LIGHT Then
''        iFilialFaturamento = giFilialEmpresa
''    End If

    If iFilialFaturamento <> EMPRESA_TODA Then
        'seleciona a filial de faturamento na combo
        For iIndice = 0 To FilialFaturamento.ListCount - 1

            If FilialFaturamento.ItemData(iIndice) = iFilialFaturamento Then

                FilialFaturamento.ListIndex = iIndice
                Exit For

            End If
        Next

    Else
        FilialFaturamento.ListIndex = 0
    End If

    Exit Sub

Erro_Seleciona_FilialEmpresa:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193202)

    End Select

    Exit Sub

End Sub

Public Sub PercAcrescFin_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercAcrescFin_Validate

    If giPercAcresFinAlterado = 0 Then Exit Sub

    If Len(Trim(PercAcrescFin.Text)) > 0 Then
        lErro = Porcentagem_Critica_Negativa(PercAcrescFin)
        If lErro <> SUCESSO Then gError 193203
    End If

    If Len(Trim(TabelaPreco.Text)) > 0 Then

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 193204

    End If

    giPercAcresFinAlterado = 0

    Exit Sub

Erro_PercAcrescFin_Validate:

    Cancel = True

    Select Case gErr

        Case 193203, 193204

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193205)

    End Select

    Exit Sub

End Sub

Private Function Cobranca_Automatica() As Long
'recalcula o tab de cobranca

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Cobranca_Automatica

    If CobrancaAutomatica.Value = 1 And Len(Trim(CondicaoPagamento.Text)) <> 0 Then

        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)

        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 193206
        
        If lErro <> SUCESSO Then gError 193207

        lErro = GridParcelas_Preenche(objCondicaoPagto)
        If lErro <> SUCESSO Then gError 193208

    End If

    Cobranca_Automatica = SUCESSO

    Exit Function

Erro_Cobranca_Automatica:

    Cobranca_Automatica = gErr

    Select Case gErr

        Case 193206, 193208

        Case 193207
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 193209)

    End Select

    Exit Function

End Function

Public Sub VolumeEspecie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeEspecie_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeEspecie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VolumeEspecie_Validate

    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie, "AVISO_CRIAR_VOLUMEESPECIE")
    If lErro <> SUCESSO Then gError 193210

    Exit Sub

Erro_VolumeEspecie_Validate:

    Cancel = True

    Select Case gErr

        Case 193210

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193211)

    End Select

End Sub

Public Sub VolumeMarca_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeMarca_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeMarca_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VolumeMarca_Validate

    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca, "AVISO_CRIAR_VOLUMEMARCA")
    If lErro <> SUCESSO Then gError 193212

    Exit Sub

Erro_VolumeMarca_Validate:

    Cancel = True

    Select Case gErr

        Case 193212

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193213)

    End Select

End Sub

Public Sub VolumeNumero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VolumeQuant_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PesoLiquido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoLiquido_Validate

    'Verifica se foi preenchido
    If Len(Trim(PesoLiquido.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(PesoLiquido.Text)
    If lErro <> SUCESSO Then gError 193214

    'Coloca o valor formatado na Tela
    PesoLiquido.Text = Format(PesoLiquido.Text, PesoLiquido.Format) 'Alterado por Wagner

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True

    Select Case gErr

        Case 193214

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193215)

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
    If lErro <> SUCESSO Then gError 193216

    'Coloca o valor formatado na Tela
    PesoBruto.Text = Format(PesoBruto.Text, PesoBruto.Format) 'Alterado por Wagner

    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True

    Select Case gErr

        Case 193216

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193217)

    End Select

    Exit Sub

End Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193218)

    End Select

End Function

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
        If lErro <> SUCESSO Then gError 193219

        If colDescontoPadrao.Count > 0 Then

            'Para cada um dos desontos padrão
            For iIndice1 = 1 To colDescontoPadrao.Count

                'Seleciona a coluna correspondente ao Desconto
                If iIndice1 = 1 Then iColuna = iGrid_Desc1Codigo_Col
                If iIndice1 = 2 Then iColuna = iGrid_Desc2Codigo_Col
                If iIndice1 = 3 Then iColuna = iGrid_Desc3Codigo_Col

                'Seleciona o tipo de desconto
                For iIndice2 = 0 To TipoDesconto1.ListCount - 1
                    If colDescontoPadrao.Item(iIndice1).iCodigo = TipoDesconto1.ItemData(iIndice2) Then
                        GridParcelas.TextMatrix(iLinha, iColuna) = TipoDesconto1.List(iIndice2)
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

        Case 193219

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193220)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_PEDIDO_SERVICO_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Pedido de Serviço"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "PedidoServico"

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

Private Sub Limpa_Tela_PedidoVenda2()
'Limpa os campos da tela sem fechar o sistema de setas

Dim iIndice As Integer

    Call Limpa_Tela(Me)

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    Set gobjInfoAdic = Nothing

    Codigo.Text = ""
    Filial.Clear
    ComissaoAutomatica.Value = vbChecked
    CobrancaAutomatica.Value = vbChecked
    TotalPercentualComissao.Caption = ""
    TotalValorComissao.Caption = ""

    'Incluído por Luiz Nogueira em 21/08/03
    VolumeMarca.Text = ""
    VolumeEspecie.Text = ""

    'Incluído por Luiz Nogueira em 21/08/03
    'Recarrega a combo VolumeEspecie e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)

    'Incluído por Luiz Nogueira em 21/08/03
    'Recarrega a combo VolumeMarca e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)

    ValorTotal.Caption = ""
    ValorProdutos.Caption = ""
    ValorProdutos2.Caption = ""
    ValorOriginalTit.Caption = ""
    ValorTit.Caption = ""
    FilialEntrega.Clear
    Transportadora.Text = ""
    TranspRedespacho.Text = ""
    RedespachoCli.Value = False
    CanalVenda.Text = ""
    CondicaoPagamento.Text = ""
    PlacaUF.Text = ""
    DataReferencia.PromptInclude = False
    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataReferencia.PromptInclude = True
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    Call DataEmissao_Validate(bSGECancelDummy)

    Emitente.Value = True
    ProdutoDescricao.Caption = ""

    FilialOrcamento.Text = "" 'Incluido por leo em 25/04/02

'    If giTipoVersao = VERSAO_FULL Then
        Seleciona_FilialEmpresa
        TabelaPreco.Text = ""
'    End If

    'tab de tributacao resumo
'    ISSIncluso.Value = 0
'    IPIBase.Caption = ""
'    IPIValor.Caption = ""
'    ISSBase.Caption = ""
'    DescTipoTrib.Caption = ""
'    IRBase.Caption = ""
'    ICMSBase.Caption = ""
'    ICMSValor.Caption = ""
'    ICMSSubstBase.Caption = ""
'    ICMSSubstValor.Caption = ""
'
'    'tab de tributacao itens
'    LabelValorFrete.Caption = ""
'    LabelValorDesconto.Caption = ""
'    LabelValorSeguro.Caption = ""
'    LabelValorOutrasDespesas.Caption = ""
'    ComboItensTrib.Clear
'    LabelValorItem.Caption = ""
'    LabelQtdeItem.Caption = ""
'    LabelUMItem.Caption = ""
'    LabelDescrNatOpItem.Caption = ""
'    DescTipoTribItem.Caption = ""

    '************** TRATAMENTO DE GRADE **************
    For iIndice = 1 To objGridItens.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next
    '*************************************************

    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridItensSRV)
    'Call Grid_Limpa(objGridBloqueio)
    Call Grid_Limpa(objGridAlocacao)
    Call Grid_Limpa(objGridComissoes)
    Call Grid_Limpa(objGridParcelas)

    Set gcolParcPV = New Collection

    Set gcolItemPedido = New Collection

'    'Resseta tributação
'    Call TributacaoPV_Reset

    iAlterado = 0
    iComissoesAlterada = 0
    gbCarregandoTela = False

    giValorDescontoAlterado = 0
    giClienteAlterado = 0
    giFilialAlterada = 0
    giFilialFatAlterada = 0
    giFilialEntregaAlterada = 0
    giDataReferenciaAlterada = 0
    giDataEntregaAlterada = 0

    giValorDescontoManual = 0
    giFilialEntregaManual = 0
    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0

    'giGridBloqueioLinhasBloqueadas = 0

    Set gobjAnotacao = Nothing
    
    Call gobjTelaTrataBloqGen.Limpa

    Call gobjTribTab.Limpa_Tela

    'Coloca clientes em atraso em vermelho
    If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)
    
    BotaoDesfazer.Enabled = False 'Inserido por Wagner
    
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

    'Coloca clientes em atraso em vermelho
    If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)
    
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
'
'Public Function Valida_Tributacao_Gravacao() As Long
'
'Dim lErro As Long
'Dim objItemPV As ClassItemPedido
'Dim iIndice As Integer, dtDataRef As Date
'
'On Error GoTo Erro_Valida_Tributacao_Gravacao
'
'    If gobjPedidoVenda.objTributacaoPV.iTipoTributacao = 0 Then gError 193227
'
'    dtDataRef = gobjPedidoVenda.dtDataEmissao
'
'    iIndice = 0
'
'    For Each objItemPV In gobjPedidoVenda.colItensPedidoSRV
'
'        iIndice = iIndice + 1
'        If Len(Trim(objItemPV.objTributacaoItemPV.sNaturezaOp)) = 0 Then gError 193228
'        If objItemPV.objTributacaoItemPV.iTipoTributacao = 0 Then gError 56922
'        If Natop_ErroTamanho(dtDataRef, objItemPV.objTributacaoItemPV.sNaturezaOp) Then gError 193230
'
'    Next
'
'    For Each objItemPV In gobjPedidoVenda.colItensPedido
'
'        iIndice = iIndice + 1
'        If Len(Trim(objItemPV.objTributacaoItemPV.sNaturezaOp)) = 0 Then gError 193231
'        If objItemPV.objTributacaoItemPV.iTipoTributacao = 0 Then gError 193232
'        If Natop_ErroTamanho(dtDataRef, objItemPV.objTributacaoItemPV.sNaturezaOp) Then gError 193233
'
'    Next
'
'
''    If Len(Trim(gobjPedidoVenda.objTributacaoPV.objTributacaoDesconto.sNaturezaOp)) = 0 Then gError 193234
''    If gobjPedidoVenda.objTributacaoPV.objTributacaoDesconto.iTipoTributacao = 0 Then gError 193235
''
''    If Len(Trim(gobjPedidoVenda.objTributacaoPV.objTributacaoFrete.sNaturezaOp)) = 0 Then gError 193236
''    If gobjPedidoVenda.objTributacaoPV.objTributacaoFrete.iTipoTributacao = 0 Then gError 193237
''
''    If Len(Trim(gobjPedidoVenda.objTributacaoPV.objTributacaoOutras.sNaturezaOp)) = 0 Then gError 193238
''    If gobjPedidoVenda.objTributacaoPV.objTributacaoOutras.iTipoTributacao = 0 Then gError 193239
''
''    If Len(Trim(gobjPedidoVenda.objTributacaoPV.objTributacaoSeguro.sNaturezaOp)) = 0 Then gError 193240
''    If gobjPedidoVenda.objTributacaoPV.objTributacaoSeguro.iTipoTributacao = 0 Then gError 193241
''
''    If Natop_ErroTamanho(dtDataRef, gobjPedidoVenda.sNaturezaOp) Or _
''        Natop_ErroTamanho(dtDataRef, gobjPedidoVenda.objTributacaoPV.objTributacaoDesconto.sNaturezaOp) Or _
''        Natop_ErroTamanho(dtDataRef, gobjPedidoVenda.objTributacaoPV.objTributacaoFrete.sNaturezaOp) Or _
''        Natop_ErroTamanho(dtDataRef, gobjPedidoVenda.objTributacaoPV.objTributacaoOutras.sNaturezaOp) Or _
''        Natop_ErroTamanho(dtDataRef, gobjPedidoVenda.objTributacaoPV.objTributacaoSeguro.sNaturezaOp) Then gError 193242
'
'    Valida_Tributacao_Gravacao = SUCESSO
'
'    Exit Function
'
'Erro_Valida_Tributacao_Gravacao:
'
'    Valida_Tributacao_Gravacao = gErr
'
'    Select Case gErr
'
'        Case 193227
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_PREENCHIDO", gErr)
'
'        Case 193228, 193231
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_ITEM_TRIBUTACAO_NAO_PREENCHIDA", iIndice)
'
'        Case 193229, 193232
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
'
'        Case 193230, 193233, 193242
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_TAMANHO_INCORRETO", gErr)
'
'        Case 193234
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_DESCONTO_NAO_PRENCHIDA", gErr)
'
'        Case 193235
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_DESCONTO_NAO_PREENCHIDO", gErr)
'
'        Case 193236
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_FRETE_NAO_PRENCHIDA", gErr)
'
'        Case 193237
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_FRETE_NAO_PREENCHIDO", gErr)
'
'        Case 193238
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_DESPESAS_NAO_PRENCHIDA", gErr)
'
'        Case 193239
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_DESPESAS_NAO_PREENCHIDO", gErr)
'
'        Case 193240
'            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SEGURO_NAO_PRENCHIDA", gErr)
'
'        Case 193241
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_SEGURO_NAO_PREENCHIDO", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193243)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function FilialFaturamento_Formata(objFilialFaturamento As Object, iFilialFaturamento As Integer) As Long

Dim iIndice As Integer

    For iIndice = 0 To objFilialFaturamento.ListCount - 1
        If objFilialFaturamento.ItemData(iIndice) = iFilialFaturamento Then
            objFilialFaturamento.Text = objFilialFaturamento.List(iIndice)
        End If
    Next

    FilialFaturamento_Formata = SUCESSO

    Exit Function

End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Codigo Then
            Call NumeroLabel_Click
'        ElseIf Me.ActiveControl Is NaturezaOp Then
'            Call NaturezaLabel_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is CondicaoPagamento Then
            Call CondPagtoLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoPecas_Click
        ElseIf Me.ActiveControl Is ServicoSRV Then
            Call BotaoServicos_Click
        ElseIf Me.ActiveControl Is Mensagem Then
            Call MensagemLabel_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call BotaoVendedores_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        ElseIf Me.ActiveControl Is TranspRedespacho Then
            Call TranspRedLabel_Click
        ElseIf Me.ActiveControl Is DescricaoProduto Then
            Call BotaoPecas_Click
        ElseIf Me.ActiveControl Is DescServicoSRV Then
            Call BotaoServicos_Click
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        End If
    End If

    If KeyCode = KEYCODE_BOTAOCONSULTA Then
        Call BotaoTipoPagto_Click
    End If

End Sub

Public Sub TransportadoraLabel_Click()

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As New Collection

    'Preenche o código da Transportadora
    If Len(Trim(Transportadora.Text)) <> 0 Then objTransportadora.iCodigo = Codigo_Extrai(Transportadora.Text)

    objTransportadora.sNomeReduzido = Transportadora.Text

    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTransportadora)

End Sub

Public Sub TranspRedLabel_Click()

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As New Collection

    'Preenche o código da Transportadora
    If Len(Trim(Transportadora.Text)) <> 0 Then objTransportadora.iCodigo = Codigo_Extrai(TranspRedespacho.Text)

    objTransportadora.sNomeReduzido = TranspRedespacho.Text

    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTranspRedespacho)

End Sub

Private Sub objEventoTransportadora_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora

    Set objTransportadora = obj1

    'Preenche o Text com Código e NomeReduzido
    Transportadora.Text = objTransportadora.iCodigo & "-" & objTransportadora.sNomeReduzido

    Me.Show

End Sub

Private Sub objEventoTranspRedespacho_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora

    Set objTransportadora = obj1

    'Preenche o Text com Código e NomeReduzido
    TranspRedespacho.Text = objTransportadora.iCodigo & "-" & objTransportadora.sNomeReduzido

    Me.Show

End Sub

Public Sub VolumeQuant_GotFocus()

    Call MaskEdBox_TrataGotFocus(VolumeQuant, iAlterado)

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

'Public Sub BotaoGerarNFiscal_Click()
'
'Dim lErro As Long
'Dim objPedidoVenda As New ClassPedidoDeVenda
'Dim objItemPV As ClassItemPedido
'Dim objNFiscal As New ClassNFiscal
'Dim objNFiscalInfo As New ClassNFiscalInfo
'Dim sSeriePadrao As String
'Dim objSerie As New ClassSerie
'
'On Error GoTo Erro_BotaoGerarNFiscal_Click
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    lErro = PedidoVenda_Critica_Le(objPedidoVenda)
'    If lErro <> SUCESSO Then gError 62126
'
'    For Each objItemPV In objPedidoVenda.colItensPedido
'        If objItemPV.dQuantFaturada > 0 Then
'            gError 62127
'            Exit For
'        End If
'    Next
'
'    Call Transfere_Dados_Pedido_NFiscal(objPedidoVenda, objNFiscal)
'
'    'Lê a Serie Padrao para Geracao de NF
'    lErro = CF("Serie_Le_Padrao", sSeriePadrao)
'    If lErro <> SUCESSO Then gError 64121
'
'    If Trim(sSeriePadrao) = "" Then gError 64124
'
'    objSerie.sSerie = sSeriePadrao
'    objSerie.iFilialEmpresa = giFilialEmpresa
'
'    'Lê a Filial Empresa no BD
'    lErro = CF("Serie_Le", objSerie)
'    If lErro <> SUCESSO And lErro <> 22202 Then gError 64122
'
'    If lErro = 22202 Then gError 64123
'
'    Set gobjGeracaoNFiscal = New ClassGeracaoNFiscal
'
'    If objSerie.iTipoFormulario = TIPO_FORMULARIO_NFISCAL Or objSerie.iTipoFormulario = TIPO_FORMULARIO_NFISCAL_SERVICO Then
'        gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISVPV
'    ElseIf objSerie.iTipoFormulario = TIPO_FORMULARIO_NFISCAL_FATURA Or objSerie.iTipoFormulario = TIPO_FORMULARIO_NFISCAL_FATURA_SERVICO Then
'        gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV
'    End If
'
'    gobjGeracaoNFiscal.sSerie = sSeriePadrao
'    Call CF2(Me, "PV_NF_DataSaida", gobjGeracaoNFiscal)  'Inserido por Wagner
'
'    objNFiscalInfo.dValorTotal = objNFiscal.dValorTotal
'    objNFiscalInfo.iCodTransp = objNFiscal.iCodTransportadora
'    objNFiscalInfo.iFilialCliente = objNFiscal.iFilialCli
'    objNFiscalInfo.lCliente = objNFiscal.lCliente
'    objNFiscalInfo.iFilialEmpresa = giFilialEmpresa
'    objNFiscalInfo.iMarcada = MARCADO
'    objNFiscalInfo.lPedido = objNFiscal.lNumPedidoVenda
'
'    gobjGeracaoNFiscal.colNFiscalInfo.Add objNFiscalInfo
'
'    lErro = GeracaoNF_Prepara_CTB()
'    If lErro <> SUCESSO Then gError 59458
'
'    lErro = CF("GeracaoNFiscal_GerarNFs", gobjGeracaoNFiscal)
'    If lErro <> SUCESSO Then gError 62130
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Call Limpa_Tela_PedidoVenda
'
'    Exit Sub
'
'Erro_BotaoGerarNFiscal_Click:
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case gErr
'
'        Case 59458, 62126, 62129, 62130, 62131, 64121, 64122
'
'        Case 62127
'            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_JA_FATURADO_PARCIALMENTE", gErr, objPedidoVenda.lCodigo)
'
'        Case 64123
'            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
'
'        Case 64124
'            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CONFIGURADA", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158069)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Function PedidoVenda_Critica_Le(objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long

On Error GoTo Erro_PedidoVenda_Critica_Le

    'Critica PedidoVenda
    lErro = PedidoServico_Critica(objPedidoVenda)
    If lErro <> SUCESSO Then gError 193261

    lErro = CF("PedidoServico_Le_Completo", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 193245 Then gError 193262
    
    If lErro <> SUCESSO Then gError 193263

    'Lê as reservas do pedido de venda
    lErro = CF("ReservasPedidoSRV_Le", objPedidoVenda)
    If lErro <> SUCESSO Then gError 193264

    PedidoVenda_Critica_Le = SUCESSO

    Exit Function

Erro_PedidoVenda_Critica_Le:

    PedidoVenda_Critica_Le = gErr

    Select Case gErr

        Case 193261, 193262, 193264

        Case 193263
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_SERVICO_NAO_CADASTRADO", gErr, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193265)

    End Select

    Exit Function

End Function

Function PedidoServico_Critica(objPedidoVenda As ClassPedidoDeVenda) As Long
'Faz crítica do PedidoVenda

Dim lErro As Long

On Error GoTo Erro_PedidoServico_Critica

    If Len(Trim(Codigo.Text)) = 0 Then gError 193266

    objPedidoVenda.lCodigo = StrParaLong(Codigo.Text)
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa

    lErro = CF("PedidoServico_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 188828 Then gError 193267
    
    If lErro = SUCESSO And objPedidoVenda.iStatus = STATUS_BAIXADO Then gError 193268

    PedidoServico_Critica = SUCESSO

    Exit Function

Erro_PedidoServico_Critica:

    PedidoServico_Critica = gErr

    Select Case gErr

        Case 193266
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)

        Case 193267

        Case 193268
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOSERVICO_BAIXADO", gErr, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193269)

    End Select

    Exit Function

End Function

'Private Sub Transfere_Dados_Pedido_NFiscal(objPedidoVenda As ClassPedidoDeVenda, objNFiscal As ClassNFiscal)
''Transfere os dados do objPedidoVenda para objNFiscal
'
'Dim objComissaoPV As ClassComissaoPedVendas
'Dim objComissaoNF As ClassComissaoNF
'Dim objItemPV As ClassItemPedido
'Dim dQuantidade As Double
'Dim objReserva As ClassReserva
'Dim colAlocacoes As ColAlocacoesItemNF
'Dim dPercDesc As Double
'Dim objItemTributacaoNF As ClassTribItemNF
'Dim objParcelaPV As ClassParcelaPedidoVenda
'Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim lErro As Long
'Dim iIndice As Integer
'Dim objTribTab As New ClassTribTab
'
'On Error GoTo Erro_Transfere_Dados_Pedido_NFiscal
'
'    With objPedidoVenda
'        'Passa os dados gerais da tela de pedido para a Nota Fiscal
'        objNFiscal.dPesoBruto = .dPesoBruto
'        objNFiscal.dPesoLiq = .dPesoLiq
'        objNFiscal.dVolumeTotal = .dVolumeTotal
'        objNFiscal.dtDataEmissao = gdtDataAtual
'        objNFiscal.dValorDesconto = .dValorDesconto
'        objNFiscal.dValorFrete = .dValorFrete
'        objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
'        objNFiscal.dValorSeguro = .dValorSeguro
'        objNFiscal.dValorTotal = .dValorTotal
'        objNFiscal.iCanal = .iCanalVenda
'        objNFiscal.iCodTransportadora = .iCodTransportadora
'        objNFiscal.iCodTranspRedesp = .iCodTranspRedesp
'        objNFiscal.iDetPagFrete = .iDetPagFrete
'        objNFiscal.iFilialCli = .iFilial
'        objNFiscal.iFilialEmpresa = .iFilialEmpresaFaturamento
'        objNFiscal.iFilialEntrega = .iFilialEntrega
'        objNFiscal.iTabelaPreco = .iTabelaPreco
'        objNFiscal.lVolumeQuant = .lVolumeQuant
'        objNFiscal.lCliente = .lCliente
'        objNFiscal.lNumPedidoVenda = .lCodigo
'        objNFiscal.sMensagemNota = .sMensagemNota
'        objNFiscal.iFilialPedido = .iFilialEmpresa
'        objNFiscal.sPlaca = .sPedidoCliente
'        objNFiscal.sNaturezaOp = .sNaturezaOp
'        objNFiscal.sNumPedidoTerc = .sPedidoCliente
'        objNFiscal.sPlacaUF = .sPlacaUF
'        objNFiscal.lVolumeEspecie = .lVolumeEspecie
'        objNFiscal.lVolumeMarca = .lVolumeMarca
'        objNFiscal.sVolumeNumero = .sVolumeNumero
'        objNFiscal.iFreteRespons = .iFreteRespons
'        objNFiscal.iCondicaoPagto = .iCondicaoPagto
'        objNFiscal.dtDataReferencia = .dtDataReferencia
'        objNFiscal.sSerie = SERIE_VERSAO_LIGHT
'        objNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV
'
'        Set objNFiscal.objProjetoInfo = .objProjetoInfo
'
'    End With
'
'    'Transfere os dados das comissões do pedido para a coleção de coissões da Nota Fiscal
'    For Each objComissaoPV In objPedidoVenda.colComissoes
'
'        With objComissaoPV
'
'            'Instancia um novo objComissaoNF
'            Set objComissaoNF = New ClassComissaoNF
'
'            'Transfere os dados do objComissaoPV para o objComissaoNF
'            objComissaoNF.iCodVendedor = .iCodVendedor
'            objComissaoNF.dValorBase = .dValorBase
'            objComissaoNF.dPercentual = .dPercentual
'            objComissaoNF.dValor = .dValor
'            objComissaoNF.dPercentualEmissao = .dPercentualEmissao
'            objComissaoNF.dValorEmissao = .dValorEmissao
'
'        End With
'
'        'Adiciona o objComissaoNF à coleção
'        objNFiscal.ColComissoesNF.Add objComissaoNF
'    Next
'
'    lErro = ParcelasPV_Gerando_ParcelasNF(objPedidoVenda, objNFiscal)
'    If lErro <> SUCESSO Then gError 51614
'
'    'Para cada item do Pedido
'    For Each objItemPV In objPedidoVenda.colItensPedido
'
'        iIndice = iIndice + 1
'
'        'calcula a quantidade
'        dQuantidade = objItemPV.dQuantidade - objItemPV.dQuantFaturada - objItemPV.dQuantCancelada
'
'        If dQuantidade > 0 Then
'
'            If objItemPV.dPrecoTotal > 0 Then
'                'Calcula o percentual de desconto
'                dPercDesc = objItemPV.dValorDesconto / objItemPV.dPrecoTotal
'            End If
'
'            Set colAlocacoes = New ColAlocacoesItemNF
'            'Transporta as reservas do PV para a coleção de alocações da NF
'            For Each objReserva In objItemPV.colReserva
'
'                objAlmoxarifado.iCodigo = objReserva.iAlmoxarifado
'                objAlmoxarifado.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
'                'Lê o ALmoxarifado
'                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
'                If lErro <> 25056 And lErro <> SUCESSO Then gError 51560
'                If lErro = 25056 Then gError 51561
'
'                If objAlmoxarifado.iFilialEmpresa = giFilialEmpresa Then colAlocacoes.Add objReserva.iAlmoxarifado, objAlmoxarifado.sNomeReduzido, dQuantidade
'
'            Next
'
'            Call objTribTab.ItemNFiscal_ObterTrib_ItemPedido(objItemPV, objItemTributacaoNF)
'
'            With objItemPV
'
'                'Adiciona o item na coleção de itens da NF
'                objNFiscal.ColItensNF.Add 0, iIndice, .sProduto, .sUnidadeMed, dQuantidade, .dPrecoUnitario, dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricao, .dValorAbatComissao, .lNumIntDoc, 0, 0, .lNumIntDoc, colAlocacoes, 0, "", "", .iStatus, 0, "", 0, 0, .iControleEstoque, .sUMEstoque, .iClasseUM, 0
'                Set objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).objTributacaoItemNF = objItemTributacaoNF
'            End With
'        End If
'    Next
'
'    'Carrega os dados globais da tributação do PV para a NF
'    Call objTribTab.NFiscal_ObterTrib_Pedido(objPedidoVenda, objNFiscal)
'
'    objNFiscal.objTributacaoNF.objTributacaoDesconto.iTipo = ITEM_TIPO_DESCONTO
'    objNFiscal.objTributacaoNF.objTributacaoFrete.iTipo = ITEM_TIPO_FRETE
'    objNFiscal.objTributacaoNF.objTributacaoOutras.iTipo = ITEM_TIPO_OUTRAS_DESP
'    objNFiscal.objTributacaoNF.objTributacaoSeguro.iTipo = ITEM_TIPO_SEGURO
'
'    Exit Sub
'
'Erro_Transfere_Dados_Pedido_NFiscal:
'
'    Select Case gErr
'
'        Case 51560
'
'        Case 51561
'            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158072)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Function ParcelasPV_Gerando_ParcelasNF(objPedidoVenda As ClassPedidoDeVenda, objNFiscal As ClassNFiscal) As Long
''Se a Data de referencia do Pedido estiver ultrapassada em regalação a Data Atual e se
''a cobrança automática do pedido estiver selecionada as  parcelas são atualizadas.
'
'Dim lErro As Long
'Dim objCondicaoPagto As New ClassCondicaoPagto
'Dim colDataVencimento As New Collection
'Dim iIndice As Integer
'Dim objParcelaPV As New ClassParcelaPedidoVenda
'
'On Error GoTo Erro_ParcelasPV_Gerando_ParcelasNF
'
'    'Se a condição de pagamento estiver preenchida e cobrança automática selecionada e a data de referencia do pedido estiver vazia ou ultrapassada
'    If objPedidoVenda.iCondicaoPagto <> 0 And objPedidoVenda.iCobrancaAutomatica = 1 And (objPedidoVenda.dtDataReferencia = DATA_NULA Or objNFiscal.dtDataReferencia < gdtDataAtual) Then
'        'A data de referencia da nota fiscal é atualizada
'        objNFiscal.dtDataReferencia = gdtDataAtual
'
'        objCondicaoPagto.iCodigo = objPedidoVenda.iCondicaoPagto
'        'Lê Condicao Pagamento no BD
'        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
'        If lErro <> SUCESSO And lErro <> 19205 Then gError 59095
'        If lErro = 19205 Then gError 59096
'
'        If objCondicaoPagto.iNumeroParcelas <> objPedidoVenda.colParcelas.Count Then gError 59097
'
'        'Calcula Datas de Vencimento das Parcelas
'        lErro = CF("Parcelas_DatasVencimento", objCondicaoPagto, objNFiscal.dtDataReferencia, colDataVencimento)
'        If lErro <> SUCESSO Then gError 59098
'
'        'Loop de preenchimento de Parcelas com Datas de Vencimento Recalculadas
'        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas
'
'            Set objParcelaPV = objPedidoVenda.colParcelas.Item(iIndice)
'
'            'Preenche o desconto padrão
'            lErro = DescontoPadraoParcela_Gera(objParcelaPV, colDataVencimento(iIndice))
'            If lErro <> SUCESSO Then gError 51613
'
'            With objParcelaPV
'                'Adiciona a parcela na coleçào de parcelas
'            '##############################################
'            'ALTERADO POR WAGNER - CLASSPARCELAPV SEM PREVISÃO E OBSERVAÇÃO
'                objNFiscal.ColParcelaReceber.Add 0, 0, .iNumParcela, STATUS_ABERTO, colDataVencimento(iIndice), colDataVencimento(iIndice), .dValor, .dValor, 1, CARTEIRA_CARTEIRA, COBRADOR_PROPRIA_EMPRESA, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0, 0, "", .dValor
'            '##############################################
'            End With
'
'        Next
'
'    Else
'
'        'transfere as parcelas do pedido de vendas para a nota fiscal
'        For Each objParcelaPV In objPedidoVenda.colParcelas
'
'            With objParcelaPV
'            '##############################################
'            'ALTERADO POR WAGNER - CLASSPARCELAPV SEM PREVISÃO E OBSERVAÇÃO
'                objNFiscal.ColParcelaReceber.Add 0, 0, .iNumParcela, STATUS_ABERTO, .dtDataVencimento, .dtDataVencimento, .dValor, .dValor, 1, CARTEIRA_CARTEIRA, COBRADOR_PROPRIA_EMPRESA, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0, 0, "", .dValor
'            '##############################################
'            End With
'
'        Next
'
'    End If
'
'    ParcelasPV_Gerando_ParcelasNF = SUCESSO
'
'    Exit Function
'
'Erro_ParcelasPV_Gerando_ParcelasNF:
'
'    ParcelasPV_Gerando_ParcelasNF = gErr
'
'    Select Case gErr
'
'        Case 59095, 59098, 51613
'
'        Case 59096
'            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
'
'        Case 59097
'            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_ALTERADA_NUM_PARC", gErr, objCondicaoPagto.iCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158073)
'
'    End Select
'
'    Exit Function
'
'End Function

'Function DescontoPadraoParcela_Gera(objParcela As Object, dtDataVencimento As Date) As Long
''Preenche o objParcela com o desconto padrão
''*Obs: O objParcela foi declarado como Object para que a função sirva para Parcelasde Pedido de Venda e Parcelas de Notas Fiscais
'
'Dim colDesconto As colDesconto
'Dim lErro As Long
'Dim iIndice2 As Integer
'
'On Error GoTo Erro_DescontoPadraoParcela_Gera
'
'    Set colDesconto = New colDesconto
'    'Gera os desconto com base nas novas datas de vencimento
'    lErro = CF("Parcela_GeraDescontoPadrao", colDesconto, dtDataVencimento)
'    If lErro <> SUCESSO Then gError 51613
'    With objParcela
'        'Limpa no obj os decontos antigos
'        .dDesconto1Valor = 0
'        .dDesconto2Valor = 0
'        .dDesconto3Valor = 0
'        .dtDesconto1Ate = DATA_NULA
'        .dtDesconto2Ate = DATA_NULA
'        .dtDesconto3Ate = DATA_NULA
'        .iDesconto1Codigo = 0
'        .iDesconto2Codigo = 0
'        .iDesconto3Codigo = 0
'
'        'Preenche o obj comos novos descontos
'        For iIndice2 = 1 To colDesconto.Count
'            If iIndice2 = 1 Then
'                .iDesconto1Codigo = colDesconto(iIndice2).iCodigo
'                .dDesconto1Valor = colDesconto(iIndice2).dValor
'                .dtDesconto1Ate = colDesconto(iIndice2).dtData
'            ElseIf iIndice2 = 2 Then
'                .iDesconto2Codigo = colDesconto(iIndice2).iCodigo
'                .dDesconto2Valor = colDesconto(iIndice2).dValor
'                .dtDesconto2Ate = colDesconto(iIndice2).dtData
'            Else
'                .iDesconto3Codigo = colDesconto(iIndice2).iCodigo
'                .dDesconto3Valor = colDesconto(iIndice2).dValor
'                .dtDesconto3Ate = colDesconto(iIndice2).dtData
'            End If
'        Next
'    End With
'
'    DescontoPadraoParcela_Gera = SUCESSO
'
'    Exit Function
'
'Erro_DescontoPadraoParcela_Gera:
'
'    DescontoPadraoParcela_Gera = gErr
'
'    Select Case gErr
'
'        Case 51613
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158074)
'
'    End Select
'
'    Exit Function
'
'End Function

'Public Sub BotaoNFiscal_Click()
'
'Dim lErro As Long
'Dim objPedidoVenda As New ClassPedidoDeVenda
'Dim objNFiscal As New ClassNFiscal
'Dim sSeriePadrao As String
'
'On Error GoTo Erro_BotaoNFiscal_Click
'
'    lErro = PedidoVenda_Critica(objPedidoVenda)
'    If lErro <> SUCESSO Then gError 62132
'
'    'Lê a Serie Padrao para gerar NF
'    lErro = CF("Serie_Le_Padrao", sSeriePadrao)
'    If lErro <> SUCESSO Then gError 25955
'
'    'Testa se a Serie Padrão está cadastrada
'    If Trim(sSeriePadrao) = "" Then gError 25956
'
'    Call Chama_Tela("NFiscalFatura", objNFiscal, objPedidoVenda)
'
'    Exit Sub
'
'Erro_BotaoNFiscal_Click:
'
'    Select Case gErr
'
'        Case 25955, 62132
'
'        Case 25956
'            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CONFIGURADA", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158075)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub Calcula_ValorProdutos(dValorProdutos As Double)

Dim dValorTotal As Double
Dim dValor As Double
Dim iIndice As Integer

    dValor = 0

    For iIndice = 1 To objGridItens.iLinhasExistentes

        dValorTotal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))

        dValor = dValor + dValorTotal

    Next
    
    For iIndice = 1 To objGridItensSRV.iLinhasExistentes

        dValorTotal = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRV_Col))

        dValor = dValor + dValorTotal

    Next

    dValorProdutos = dValor

End Sub

'Private Function GeracaoNF_Prepara_CTB() As Long
''prepara informacoes necessarias para a contabilizacao
'
'Dim lErro As Long, objPeriodo As New ClassPeriodo
'
'On Error GoTo Erro_GeracaoNF_Prepara_CTB
'
'    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
'
'        gobjGeracaoNFiscal.objTelaAtualizacao = Me
'        gobjGeracaoNFiscal.dtContabil = gdtDataAtual
'        gobjGeracaoNFiscal.iLoteContabil = 0
'
'        'Coloca o periodo relativo a data na tela
'        lErro = CF("Periodo_Le", gobjGeracaoNFiscal.dtContabil, objPeriodo)
'        If lErro <> SUCESSO Then gError 59457
'
'        giPeriodo = objPeriodo.iPeriodo
'        giExercicio = objPeriodo.iExercicio
'
'    End If
'
'    GeracaoNF_Prepara_CTB = SUCESSO
'
'    Exit Function
'
'Erro_GeracaoNF_Prepara_CTB:
'
'    GeracaoNF_Prepara_CTB = gErr
'
'    Select Case gErr
'
'        Case 59457
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158076)
'
'    End Select
'
'    Exit Function
'
'End Function

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor, Optional objContexto As Object) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim dQuantidadeConvertida As Double
Dim sProdutoFormatado As String, sProdutoTela As String
Dim iPreenchido As Integer
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim bEncontrouProduto As Boolean
Dim iLinha2 As Integer, objCodigoNome As New AdmlCodigoNome
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sContaMascarada As String, sAlmoxNomeRed As String
Dim iAlmoxPadrao As Integer, objCliente As New ClassCliente
Dim bEncontrouQuant As Boolean, objFilialCliente As New ClassFilialCliente
Dim bEncontrouQuant2 As Boolean, objItem As ClassItemNF, objAlocacao As ClassItemNFAlocacao
Dim objBloqueioPV As ClassBloqueioPV
Dim iAchou As Integer

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case CTACONTABILEST1 'parametros: produto no formato da tela do grid de itens, produto no formato da tela do grid de alocacoes e nome reduzido do almoxarifado

            For Each objItem In gobjNFiscalCTB.ColItensNF

                For Each objAlocacao In objItem.colAlocacoes

                    If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then

                        objAlmoxarifado.sNomeReduzido = CStr(objMnemonicoValor.vParam(3))

                        'Lê o Nome Reduzido do Almoxarifado
                        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                        If lErro <> SUCESSO And lErro <> 25060 Then gError 193274

                        'Se não encontrou ===> Erro
                        If lErro = SUCESSO Then

                            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo

                            lErro = CF("Produto_Formata", objMnemonicoValor.vParam(1), sProdutoFormatado, iPreenchido)
                            If iPreenchido <> PRODUTO_PREENCHIDO Then gError 193275

                            objEstoqueProduto.sProduto = sProdutoFormatado

                            'Lê a conta contábil do Produto no Almoxarifado
                            lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
                            If lErro <> SUCESSO And lErro <> 49991 Then gError 193276

                            If Len(Trim(objEstoqueProduto.sContaContabil)) > 0 Then

                                sContaMascarada = String(STRING_CONTA, 0)

                                lErro = Mascara_RetornaContaTela(objEstoqueProduto.sContaContabil, sContaMascarada)
                                If lErro <> SUCESSO Then gError 193277

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

            Next

        Case CODIGO1
            objMnemonicoValor.colValor.Add gobjPedidoVendaCTB.lCodigo

        Case QUANT_ESTOQUE

            For Each objItem In gobjNFiscalCTB.ColItensNF

                lErro = CF("UMEstoque_Conversao", objItem.sProduto, objItem.sUnidadeMed, objItem.dQuantidade, dQuantidadeConvertida)
                If lErro <> SUCESSO Then gError 193278

                objMnemonicoValor.colValor.Add dQuantidadeConvertida

            Next

        Case ALMOX1

            If gcolAlmoxFilial.Count = 0 Then

                lErro = CF("Almoxarifados_Le_FilialEmpresa", gobjNFiscalCTB.iFilialEmpresa, gcolAlmoxFilial)
                If lErro <> SUCESSO Then gError 193279

            End If

            For Each objItem In gobjNFiscalCTB.ColItensNF

                For Each objAlocacao In objItem.colAlocacoes

                    sAlmoxNomeRed = ""

                    For Each objAlmoxarifado In gcolAlmoxFilial

                        If objAlocacao.iAlmoxarifado = objAlmoxarifado.iCodigo Then

                            sAlmoxNomeRed = objAlmoxarifado.sNomeReduzido
                            Exit For

                        End If

                    Next

                    objMnemonicoValor.colValor.Add sAlmoxNomeRed

                Next

            Next

        Case DATA_EMISSAO
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.dtDataEmissao

        Case DATA_SAIDA
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.dtDataSaida

        Case DESCONTO1
            For Each objItem In gobjNFiscalCTB.ColItensNF

                objMnemonicoValor.colValor.Add objItem.dValorDesconto

            Next

        Case DESCRICAO_ITEM
            For Each objItem In gobjNFiscalCTB.ColItensNF

                objMnemonicoValor.colValor.Add objItem.sDescricaoItem

            Next

        Case FILIAL_CONTA

            If Len(Filial.Text) > 0 Then

                objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilialCliente)
                If lErro <> SUCESSO Then gError 193280

                If objFilialCliente.sContaContabil <> "" Then

                    lErro = Mascara_RetornaContaTela(objFilialCliente.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError 193281

                Else

                    sContaMascarada = ""

                End If

                objMnemonicoValor.colValor.Add sContaMascarada

            Else

                objMnemonicoValor.colValor.Add ""

            End If

        Case FILIAL1

            objFilialCliente.lCodCliente = gobjNFiscalCTB.lCliente
            objFilialCliente.iCodFilial = gobjNFiscalCTB.iFilialCli

            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 193282
            
            If lErro <> SUCESSO Then gError 193271

            objMnemonicoValor.colValor.Add objFilialCliente.sNome

        Case CLIENTE1

            objCodigoNome.lCodigo = gobjNFiscalCTB.lCliente

            lErro = CF("Cliente_Le_NomeRed", objCodigoNome)
            If lErro <> SUCESSO And lErro <> 12553 Then gError 193283
            
            If lErro <> SUCESSO Then gError 193272

            objMnemonicoValor.colValor.Add objCodigoNome.sNome

        Case NATUREZA_OP
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.sNaturezaOp

        Case NFISCAL1
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.lNumNotaFiscal

        Case PRODUTO1
            For Each objItem In gobjNFiscalCTB.ColItensNF

                lErro = Mascara_RetornaProdutoTela(objItem.sProduto, sProdutoTela)
                If lErro <> SUCESSO Then gError 193284

                objMnemonicoValor.colValor.Add sProdutoTela

            Next

        Case QUANTIDADE1
            For Each objItem In gobjNFiscalCTB.ColItensNF

                objMnemonicoValor.colValor.Add objItem.dQuantidade

            Next

        Case Serie1
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.sSerie

        Case UNIDADE_MED
            For Each objItem In gobjNFiscalCTB.ColItensNF

                objMnemonicoValor.colValor.Add objItem.sUnidadeMed

            Next

        Case VALOR_TOTAL
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.dValorTotal

        Case PRECO_UNITARIO
            For Each objItem In gobjNFiscalCTB.ColItensNF

                objMnemonicoValor.colValor.Add objItem.dPrecoUnitario

            Next

        Case PRECO_TOTAL
            For Each objItem In gobjNFiscalCTB.ColItensNF

                objMnemonicoValor.colValor.Add objItem.dValorTotal

            Next

        Case ICMS
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.objTributacaoNF.dICMSValor

        Case ICMSSUBST
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.objTributacaoNF.dICMSSubstValor

        Case VALOR_FRETE
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.dValorFrete

        Case VALOR_SEGURO
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.dValorSeguro

        Case VALOR_DESPESAS
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.dValorOutrasDespesas

        Case IPI
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.objTributacaoNF.dIPIValor

        Case VALOR_DESCONTO
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.dValorDesconto

        Case ISS_VALOR
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.objTributacaoNF.dISSValor

        Case ISS_INCLUSO

            objMnemonicoValor.colValor.Add gobjNFiscalCTB.objTributacaoNF.iISSIncluso <> 0

        Case VALOR_IRRF
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.objTributacaoNF.dIRRFValor

        Case VALOR_PRODUTOS
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.dValorProdutos

        Case PRODUTO_ALMOX
            For Each objItem In gobjNFiscalCTB.ColItensNF

                For Each objAlocacao In objItem.colAlocacoes

                    lErro = Mascara_RetornaProdutoTela(objItem.sProduto, sProdutoTela)
                    If lErro <> SUCESSO Then gError 193285

                    objMnemonicoValor.colValor.Add sProdutoTela

                Next

            Next

        Case QUANT_ALOCADA

            For Each objItem In gobjNFiscalCTB.ColItensNF

                If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then

                    For Each objAlocacao In objItem.colAlocacoes

                        objMnemonicoValor.colValor.Add objAlocacao.dQuantidade

                    Next

                Else

                    For Each objAlocacao In objItem.colAlocacoes

                        objMnemonicoValor.colValor.Add 0

                    Next

                End If

            Next

        Case UNIDADE_MED_EST
            For Each objItem In gobjNFiscalCTB.ColItensNF

                For Each objAlocacao In objItem.colAlocacoes

                    If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
                        objMnemonicoValor.colValor.Add objItem.sUMEstoque
                    Else
                        objMnemonicoValor.colValor.Add ""
                    End If

                Next

            Next


        Case CREDITO_ULTRAPASSADO

            For Each objBloqueioPV In objContexto.colBloqueiosPV

                If objBloqueioPV.iTipoDeBloqueio = BLOQUEIO_CREDITO Then
                    iAchou = 1
                    Exit For
                End If
            Next

            objMnemonicoValor.colValor.Add iAchou

        Case CODIGOPEDIDO
            objMnemonicoValor.colValor.Add objContexto.lCodigo


        Case PRODUTO_SEM_RESERVA

            For Each objBloqueioPV In objContexto.colBloqueiosPV

                If objBloqueioPV.iTipoDeBloqueio = BLOQUEIO_PARCIAL Or objBloqueioPV.iTipoDeBloqueio = BLOQUEIO_NAO_RESERVA Then
                    iAchou = 1
                    Exit For
                End If
            Next

            objMnemonicoValor.colValor.Add iAchou


        Case CLIENTE_NOME

            objCodigoNome.lCodigo = objContexto.lCliente

            lErro = CF("Cliente_Le_NomeRed", objCodigoNome)
            If lErro <> SUCESSO And lErro <> 12553 Then gError 193286

            If lErro <> SUCESSO Then gError 193273

            objMnemonicoValor.colValor.Add objCodigoNome.sNome

        Case Else
            gError 193270

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 193270
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case 193274 To 193286

        Case 193271
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objFilialCliente.iCodFilial, objFilialCliente.lCodCliente)

        Case 193272
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, gobjNFiscalCTB.lCliente)

        Case 193273
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objContexto.lCliente)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193287)

    End Select

    Exit Function

End Function

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é chamada a cada atualizacao de nota fiscal e é responsavel por gerar a contabilizacao correspondente

Dim lErro As Long, lDoc As Long, objItem As ClassItemNF, iNumAlocacoes As Integer

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjNFiscalCTB = vParams(0)
    Set gobjPedidoVendaCTB = vParams(1)

    'percorre itens otendo qtde de alocacoes da nf como um todo
    For Each objItem In gobjNFiscalCTB.ColItensNF
        iNumAlocacoes = iNumAlocacoes + objItem.colAlocacoes.Count
    Next

    GridAlocacao.Tag = iNumAlocacoes
    GridItens.Tag = gobjNFiscalCTB.ColItensNF.Count

    'obter numero de Doc
    lErro = CF("Voucher_Automatico1", gobjNFiscalCTB.iFilialEmpresa, giExercicio, giPeriodo, MODULO_FATURAMENTO, lDoc)
    If lErro <> SUCESSO Then gError 193288

    'grava a contabilizacao
    lErro = objContabAutomatica.Gravar_Registro(Me, IIf(gobjNFiscalCTB.iTipoNFiscal = DOCINFO_NFISFVPV, "NFiscalFaturaPedido", "NFiscalPedido"), gobjNFiscalCTB.lNumIntDoc, gobjNFiscalCTB.lCliente, gobjNFiscalCTB.iFilialCli, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjNFiscalCTB.iFilialEmpresa, gobjGeracaoNFiscal.iLoteContabil, gobjNFiscalCTB.lNumNotaFiscal)
    If lErro <> SUCESSO Then gError 193289

    GeraContabilizacao = SUCESSO

    Exit Function

Erro_GeraContabilizacao:

    GeraContabilizacao = gErr

    Select Case gErr

        Case 193288, 193289

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193290)

    End Select

    Exit Function

End Function

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
    Call Seleciona_FilialEmpresaOrcamento

    Exit Sub

Erro_Carrega_FilialOrcamento:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193291)

    End Select

    Exit Sub

End Sub

Private Sub Seleciona_FilialEmpresaOrcamento()

Dim lErro As Long
Dim iIndice As Integer
Dim iFilialOrcamento As Integer

On Error GoTo Erro_Seleciona_FilialEmpresaOrcamento

'    If giTipoVersao = VERSAO_FULL Then

        iFilialOrcamento = gobjFAT.iFilialFaturamento

'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'        iFilialOrcamento = giFilialEmpresa
'    End If

''    If giTipoVersao = VERSAO_FULL Then
''
''        'Lê a filial de faturamento
''        lErro = CF("FilialFaturamento_Le",giFilialEmpresa, iFilialOrcamento)
''        If lErro <> SUCESSO Then gError 51139
''
''    ElseIf giTipoVersao = VERSAO_LIGHT Then
''        iFilialOrcamento = giFilialEmpresa
''    End If

    If iFilialOrcamento <> EMPRESA_TODA Then
        'seleciona a filial de faturamento na combo
        For iIndice = 0 To FilialOrcamento.ListCount - 1

            If FilialOrcamento.ItemData(iIndice) = iFilialOrcamento Then

                FilialOrcamento.ListIndex = iIndice
                Exit For

            End If
        Next

    Else
        FilialOrcamento.ListIndex = 0
    End If

    Exit Sub

Erro_Seleciona_FilialEmpresaOrcamento:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193292)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 192292
    If lErro <> SUCESSO Then gError 193293

    Exit Sub

Erro_FilialOrcamento_Validate:

    Cancel = True

    Select Case gErr

        Case 193292

        Case 193293
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialOrcamento.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193294)

    End Select

    Exit Sub

End Sub

Public Sub OrcamentoLabel_Click()

Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim colSelecao As New Collection

    '??? Isso está certo ???
'    colSelecao.Add gdtDataHoje

    objOrcamentoVenda.lCodigo = StrParaLong(Orcamento.Text)

    Call Chama_Tela("OrcamentoPedidoSRVLista", colSelecao, objOrcamentoVenda, objEventoOrcamento)

End Sub

Public Property Get OrcamentoLabel() As Object
     Set OrcamentoLabel = objUserControl.Controls("OrcamentoLabel")
End Property

Public Property Get DiretoIndireto() As Object
     Set DiretoIndireto = objUserControl.Controls("DiretoIndireto")
End Property


Public Sub DiretoIndireto_Change()

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

Public Sub BotaoVerOrcamento_Click()

Dim lErro As Long
Dim objOcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_BotaoVerPedido_Click

    'Verifica se o Orçamento e a filial do pedido foram informados
    If Len(Trim(Orcamento.Text)) = 0 Then gError 193295
    If Len(Trim(FilialOrcamento.Text)) = 0 Then gError 193296

    objOcamentoVenda.lCodigo = StrParaLong(Orcamento.Text)
    objOcamentoVenda.iFilialEmpresa = Codigo_Extrai(FilialOrcamento.Text)

    'Busca o Pedido na tabela de pedidos de venda
    lErro = CF("OrcamentoSRV_Le", objOcamentoVenda)
    If lErro <> SUCESSO And lErro <> 186534 Then gError 193297

    If lErro <> SUCESSO Then gError 193298 'Não encontrou, erro.

    'Chama  a tela de Pedido de Venda passando o pedido de venda da tela
    Call Chama_Tela("OrcamentoSRV", objOcamentoVenda)

    Exit Sub

Erro_BotaoVerPedido_Click:

    Select Case gErr

        Case 193295
            Call Rotina_Erro(vbOKOnly, "ERRO_CODORCAMENTO_NAO_INFORMADO", gErr)

        Case 193296
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALORCAMENTO_NAO_INFORMADA", gErr)

        Case 193297

        Case 193298
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOSRV_NAO_CADASTRADO1", gErr, objOcamentoVenda.lCodigo, objOcamentoVenda.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193299)

    End Select

    Exit Sub

End Sub

Public Sub BotaoOrcamento_Click()

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim objMapBloqGen As New ClassMapeamentoBloqGen
Dim objBloqueioGen As ClassBloqueioGen
Dim colTipoBloq As New Collection
Dim objTipoBloq As ClassTiposDeBloqueioGen

On Error GoTo Erro_BotaoOrcamento_Click

    'Verifica se o Orçamento e a filial do pedido foram informados
    If Len(Trim(Orcamento.Text)) = 0 Then gError 193341
    If Len(Trim(FilialOrcamento.Text)) = 0 Then gError 193342

    objOrcamentoVenda.lCodigo = StrParaLong(Orcamento.Text)
    objOrcamentoVenda.iFilialEmpresa = Codigo_Extrai(FilialOrcamento.Text)

    'Busca o Pedido na tabela de pedidos de venda
    lErro = CF("OrcamentoSRV_Le_Completo", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 193343 Then gError 193352

    If lErro <> SUCESSO Then gError 193353

    'Se o orçamento estiver vinculado a um Pedido de Venda -> Erro.
    If objOrcamentoVenda.lNumIntPedVenda <> 0 Then

        gError 193354

    'Se o orçamento estiver vinculado a uma Nota Fiscal -> Erro.
    ElseIf objOrcamentoVenda.lNumIntNFiscal <> 0 Then

        gError 193355

    End If
    
    objMapBloqGen.iTipoTelaBloqueio = MAPBLOQGEN_TIPOTELA_ORCAMENTOSRV
    
     lErro = CF("MapeamentoBloqGen_Le", objMapBloqGen)
    If lErro <> SUCESSO Then gError 198549
    
    lErro = CF("BloqueiosGen_Le", objMapBloqGen, objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 198550
    
    'Le todos os Tipos de Bloqueio
    lErro = CF("TiposDeBloqueioGen_Le_TipoTela", objMapBloqGen.iTipoTelaBloqueio, colTipoBloq)
    If lErro <> SUCESSO And lErro <> 29168 Then gError 198551
    
    For Each objTipoBloq In colTipoBloq
        If objTipoBloq.iNaoApareceTelaLib = DESMARCADO Then
            For Each objBloqueioGen In objOrcamentoVenda.colBloqueiosPV
                If objTipoBloq.iCodigo = objBloqueioGen.iTipoDeBloqueio Then
                    If objBloqueioGen.dtDataLib = DATA_NULA Then gError 198552
                End If
            Next
        End If
    Next

    'Verifica se a Filial de Faturamento do Orcamento coincide com a filial de trabalho
    If objOrcamentoVenda.iFilialEmpresa <> giFilialEmpresa Then gError 193356

    'Traz os dados do pedido para a tela
    lErro = Traz_Dados_Orcamento_Tela(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 193357

    Call ValorTotal_Calcula

    'Trata clientes não cadastrados no orçamento
    Call Cadastra_Cliente_Inexistente(objOrcamentoVenda)

    Exit Sub

Erro_BotaoOrcamento_Click:

    Select Case gErr

        Case 193341
            Call Rotina_Erro(vbOKOnly, "ERRO_CODORCAMENTO_NAO_INFORMADO", gErr)

        Case 193342
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALORCAMENTO_NAO_INFORMADA", gErr)

        Case 193352, 193357, 198549 To 198551

        Case 193353
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOSRV_NAO_CADASTRADO1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case 193354
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOSRV_VINCULADO_PEDIDOSRV1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case 193355
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOSRV_VINCULADO_NFISCAL1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case 193356
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFATURAMENTO_DIFERENTE1", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, giFilialEmpresa)

        Case 198552
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCSRV_BLOQUEADO", gErr, objOrcamentoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193358)

    End Select

    Exit Sub

End Sub

Public Property Get BotaoOrcamento() As Object
     Set BotaoOrcamento = objUserControl.Controls("BotaoOrcamento")
End Property

Public Function Traz_Dados_Orcamento_Tela(objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objCliente As New ClassCliente

On Error GoTo Erro_Traz_Dados_Orcamento_Tela


    'Transfere os dados de objPedidoVenda para objNFiscal
    Call Transfere_Dados_Orcamento_Pedido(objOrcamentoVenda, objPedidoVenda)

    objPedidoVenda.dtDataEmissao = gdtDataAtual
    objPedidoVenda.dtDataEntrega = DATA_NULA

    'Carrega os dados de PedidoVenda na tela passando a constante que identifica como oriundo de Orçamento Venda.
    lErro = Traz_Pedido_Tela(objPedidoVenda, objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 193359

    Traz_Dados_Orcamento_Tela = SUCESSO

    Exit Function

Erro_Traz_Dados_Orcamento_Tela:

    Traz_Dados_Orcamento_Tela = gErr

    Select Case gErr

        Case 193359

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193360)

    End Select

    Exit Function

End Function

Public Sub Transfere_Dados_Orcamento_Pedido(objOrcamentoVenda As ClassOrcamentoVenda, objPedidoVenda As ClassPedidoDeVenda)
'Transfere os dados do objOrcamentoVenda para objPedidoDeVenda

Dim objItemOV As ClassItemOV
Dim objItemPV As ClassItemPedido
Dim dQuantidade As Double
Dim dPercDesc As Double
Dim objItemTributacaoPedido As ClassTributacaoDocItem
Dim objComissaoPV As New ClassComissaoPedVendas
Dim objParcelaOV As ClassParcelaOV
Dim lErro As Long
Dim objProduto As ClassProduto
Dim colItensRomaneioGrade As New Collection
Dim objProjetoInfo As Object

On Error GoTo Erro_Transfere_Dados_Pedido_NFiscal

'    Call gobjTelaProjetoInfo.Obtem_ProjetoInfo(objProjetoInfo, objOrcamentoVenda.lNumIntDoc, PRJ_CR_TIPO_OV)
'    Set objPedidoVenda.objProjetoInfo = objProjetoInfo

    Set objPedidoVenda.objInfoAdic = objOrcamentoVenda.objInfoAdic

    With objPedidoVenda

        'AKI!!! Transferir o resto dos campos

        'Passa os dados gerais da tela de Orcamento para o Pedido
        objOrcamentoVenda.dtDataEmissao = gdtDataAtual

        .dtDataEmissao = objOrcamentoVenda.dtDataEmissao
        .dtDataReferencia = objOrcamentoVenda.dtDataReferencia
        .dValorDesconto = objOrcamentoVenda.dValorDesconto
        .dValorFrete = objOrcamentoVenda.dValorFrete
        .dValorOutrasDespesas = objOrcamentoVenda.dValorOutrasDespesas
        .dValorSeguro = objOrcamentoVenda.dValorSeguro
        .dValorTotal = objOrcamentoVenda.dValorTotal
        .lNumIntDoc = objOrcamentoVenda.lNumIntPedVenda
        .iNaturezaOpManual = VAR_PREENCH_MANUAL
        .iCobrancaAutomatica = objOrcamentoVenda.iCobrancaAutomatica
        .dValorProdutos = objOrcamentoVenda.dValorProdutos
        .dPercAcrescFinanceiro = objOrcamentoVenda.dPercAcrescFinanceiro
        .iFilial = objOrcamentoVenda.iFilial
        .iFilialEntrega = objOrcamentoVenda.iFilial
        .iFilialEmpresaFaturamento = objOrcamentoVenda.iFilialEmpresa
        .iTabelaPreco = objOrcamentoVenda.iTabelaPreco
        .lCliente = objOrcamentoVenda.lCliente
        .sNaturezaOp = objOrcamentoVenda.sNaturezaOp
        .iCondicaoPagto = objOrcamentoVenda.iCondicaoPagto
        .lNumIntSolicSRV = objOrcamentoVenda.lNumIntSolicSRV
        .dValorDescontoItens = objOrcamentoVenda.dValorDescontoItens
        .dValorDescontoTit = objOrcamentoVenda.dValorDescontoTit
        .dValorItens = objOrcamentoVenda.dValorItens

    End With


    'Para cada item do Pedido
    For Each objItemOV In objOrcamentoVenda.colItensSRV

'        Set objItemPV = New ClassItemPedido
'
'        Call ItemPedidoVenda_ObterTrib_ItemOrcamento(objItemPV, objItemOV)

        With objItemOV

            Set objItemPV = objPedidoVenda.IncluirItemSRV(.sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, 0, 0, 0, Nothing, "", .sUnidadeMed, "", "", 0, 0, 0, "", STATUS_ABERTO, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, .objTributacaoItemOV)

        End With

    Next



    'Para cada item do Pedido
    For Each objItemOV In objOrcamentoVenda.colItens

'        Set objItemPV = New ClassItemPedido
'
'        Call ItemPedidoVenda_ObterTrib_ItemOrcamento(objItemPV, objItemOV)

        With objItemOV

            Set objItemPV = objPedidoVenda.IncluirItem(.sProduto, .dQuantidade, .dPrecoUnitario, .dPrecoTotal, .dValorDesconto, .dtDataEntrega, .sProdutoDescricao, 0, 0, 0, Nothing, "", .sUnidadeMed, "", "", 0, 0, 0, "", STATUS_ABERTO, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, .objTributacaoItemOV)

        End With

    Next

    Call gobjTribTab.PedidoVenda_ObterTrib_Orcamento(objPedidoVenda, objOrcamentoVenda)

    If objOrcamentoVenda.iVendedor <> 0 Then

        'adiciona na coleção de comissões os dados o código do vendedor
        objComissaoPV.iCodVendedor = objOrcamentoVenda.iVendedor
        objPedidoVenda.colComissoes.Add objComissaoPV

    End If

    For Each objParcelaOV In objOrcamentoVenda.colParcela
        With objParcelaOV
            objPedidoVenda.colParcelas.Add .dValor, .dtDataVencimento, .iNumParcela, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .dtDesconto3Ate, .dDesconto3Valor, .iDesconto3Codigo
        End With
    Next

    'Reinicializa as coleções de itens
    Set gcolItemPedido = New Collection

    Exit Sub

Erro_Transfere_Dados_Pedido_NFiscal:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193361)

    End Select

End Sub
'
'Private Sub ComplPV_ObterTrib_ComplOV(objTribComplPV As ClassTributacaoItemPV, objTribComplOV As ClassTributacaoItemPV)
'
'    With objTribComplPV
'
'        .dICMSAliquota = objTribComplOV.dICMSAliquota
'        .dICMSBase = objTribComplOV.dICMSBase
'        .dICMSPercRedBase = objTribComplOV.dICMSPercRedBase
'        .dICMSSubstAliquota = objTribComplOV.dICMSSubstAliquota
'        .dICMSSubstBase = objTribComplOV.dICMSSubstBase
'        .dICMSSubstValor = objTribComplOV.dICMSSubstValor
'        .dICMSValor = objTribComplOV.dICMSValor
'        .dIPIAliquota = objTribComplOV.dIPIAliquota
'        .dIPIBaseCalculo = objTribComplOV.dIPIBaseCalculo
'        .dIPIPercRedBase = objTribComplOV.dIPIPercRedBase
'        .dIPIValor = objTribComplOV.dIPIValor
'        .iICMSTipo = objTribComplOV.iICMSTipo
'        .iIPITipo = objTribComplOV.iIPITipo
'        .iTipoTributacao = objTribComplOV.iTipoTributacao
'        .lNumIntDoc = objTribComplOV.lNumIntDoc
'        .sNaturezaOp = objTribComplOV.sNaturezaOp
'
'        .iNaturezaOpManual = objTribComplOV.iNaturezaOpManual
'        .iTipoTributacaoManual = objTribComplOV.iTipoTributacaoManual
'        .iIPITipoManual = objTribComplOV.iIPITipoManual
'        .iIPIBaseManual = objTribComplOV.iIPIBaseManual
'        .iIPIPercRedBaseManual = objTribComplOV.iIPIPercRedBaseManual
'        .iIPIAliquotaManual = objTribComplOV.iIPIAliquotaManual
'        .iIPIValorManual = objTribComplOV.iIPIValorManual
'        .iICMSTipoManual = objTribComplOV.iICMSTipoManual
'        .iICMSBaseManual = objTribComplOV.iICMSBaseManual
'        .iICMSPercRedBaseManual = objTribComplOV.iICMSPercRedBaseManual
'        .iICMSAliquotaManual = objTribComplOV.iICMSAliquotaManual
'        .iICMSValorManual = objTribComplOV.iICMSValorManual
'        .iICMSSubstBaseManual = objTribComplOV.iICMSSubstBaseManual
'        .iICMSSubstAliquotaManual = objTribComplOV.iICMSSubstAliquotaManual
'        .iICMSSubstValorManual = objTribComplOV.iICMSSubstValorManual
'
'    End With
'
'End Sub
'
'Sub PedidoVenda_ObterTrib_Orcamento(objPedidoVenda As ClassPedidoDeVenda, objOrcamentoVenda As ClassOrcamentoVenda)
'
'    Set objPedidoVenda.objTributacaoPV = New ClassTributacaoPV
'
'    With objPedidoVenda.objTributacaoPV
'
'        .dICMSBase = objOrcamentoVenda.objTributacaoOV.dICMSBase
'        .dICMSSubstBase = objOrcamentoVenda.objTributacaoOV.dICMSSubstBase
'        .dICMSSubstValor = objOrcamentoVenda.objTributacaoOV.dICMSSubstValor
'        .dICMSValor = objOrcamentoVenda.objTributacaoOV.dICMSValor
'        .dIPIBase = objOrcamentoVenda.objTributacaoOV.dIPIBase
'        .dIPIValor = objOrcamentoVenda.objTributacaoOV.dIPIValor
'        .dIRRFAliquota = objOrcamentoVenda.objTributacaoOV.dIRRFAliquota
'        .dIRRFBase = objOrcamentoVenda.objTributacaoOV.dIRRFBase
'        .dIRRFValor = objOrcamentoVenda.objTributacaoOV.dIRRFValor
'        .dISSBase = objOrcamentoVenda.objTributacaoOV.dISSBase
'        .dISSValor = objOrcamentoVenda.objTributacaoOV.dISSValor
'        .iISSIncluso = objOrcamentoVenda.objTributacaoOV.iISSIncluso
'        .iTipoTributacao = objOrcamentoVenda.objTributacaoOV.iTipoTributacao
'        .dPISRetido = objOrcamentoVenda.objTributacaoOV.dPISRetido
'        .dCOFINSRetido = objOrcamentoVenda.objTributacaoOV.dCOFINSRetido
'        .dCSLLRetido = objOrcamentoVenda.objTributacaoOV.dCSLLRetido
'
'        .iTipoTributacaoManual = objOrcamentoVenda.objTributacaoOV.iTipoTributacaoManual
'        .iICMSBaseManual = objOrcamentoVenda.objTributacaoOV.iICMSBaseManual
'        .iICMSValorManual = objOrcamentoVenda.objTributacaoOV.iICMSValorManual
'        .iICMSSubstBaseManual = objOrcamentoVenda.objTributacaoOV.iICMSSubstBaseManual
'        .iICMSSubstValorManual = objOrcamentoVenda.objTributacaoOV.iICMSSubstValorManual
'        .iIPIBaseManual = objOrcamentoVenda.objTributacaoOV.iIPIBaseManual
'        .iIPIValorManual = objOrcamentoVenda.objTributacaoOV.iIPIValorManual
'        .iIRRFAliquotaManual = objOrcamentoVenda.objTributacaoOV.iIRRFAliquotaManual
'        .iIRRFValorManual = objOrcamentoVenda.objTributacaoOV.iIRRFValorManual
'        .iISSAliquotaManual = objOrcamentoVenda.objTributacaoOV.iISSAliquotaManual
'        '.iISSInclusoManual = objOrcamentoVenda.objTributacaoOV.iISSInclusoManual
'        .iISSValorManual = objOrcamentoVenda.objTributacaoOV.iISSValorManual
'        .iPISRetidoManual = objOrcamentoVenda.objTributacaoOV.iPISRetidoManual
'        .iCOFINSRetidoManual = objOrcamentoVenda.objTributacaoOV.iCOFINSRetidoManual
'        .iCSLLRetidoManual = objOrcamentoVenda.objTributacaoOV.iCSLLRetidoManual
'
''        Set .objTributacaoDesconto = New ClassTributacaoItemPV
''        Set .objTributacaoFrete = New ClassTributacaoItemPV
''        Set .objTributacaoOutras = New ClassTributacaoItemPV
''        Set .objTributacaoSeguro = New ClassTributacaoItemPV
''
''        Call ComplPV_ObterTrib_ComplOV(.objTributacaoDesconto, objOrcamentoVenda.objTributacaoOV.objTributacaoDesconto)
''        Call ComplPV_ObterTrib_ComplOV(.objTributacaoFrete, objOrcamentoVenda.objTributacaoOV.objTributacaoFrete)
''        Call ComplPV_ObterTrib_ComplOV(.objTributacaoOutras, objOrcamentoVenda.objTributacaoOV.objTributacaoOutras)
''        Call ComplPV_ObterTrib_ComplOV(.objTributacaoSeguro, objOrcamentoVenda.objTributacaoOV.objTributacaoSeguro)
'
'    End With
'
'End Sub
'
'Sub ItemPedidoVenda_ObterTrib_ItemOrcamento(objItemPV As ClassItemPedido, objItemOV As ClassItemOV)
'
'    Set objItemPV = New ClassItemPedido
'
'    With objItemOV
'
'        objItemPV.objTributacaoItemPV.dICMSAliquota = .objTributacaoItemOV.dICMSAliquota
'        objItemPV.objTributacaoItemPV.dICMSBase = .objTributacaoItemOV.dICMSBase
'        objItemPV.objTributacaoItemPV.dICMSPercRedBase = .objTributacaoItemOV.dICMSPercRedBase
'        objItemPV.objTributacaoItemPV.dICMSSubstAliquota = .objTributacaoItemOV.dICMSSubstAliquota
'        objItemPV.objTributacaoItemPV.dICMSSubstBase = .objTributacaoItemOV.dICMSSubstBase
'        objItemPV.objTributacaoItemPV.dICMSSubstValor = .objTributacaoItemOV.dICMSSubstValor
'        objItemPV.objTributacaoItemPV.dICMSValor = .objTributacaoItemOV.dICMSValor
'        objItemPV.objTributacaoItemPV.dIPIAliquota = .objTributacaoItemOV.dIPIAliquota
'        objItemPV.objTributacaoItemPV.dIPIBaseCalculo = .objTributacaoItemOV.dIPIBaseCalculo
'        objItemPV.objTributacaoItemPV.dIPIPercRedBase = .objTributacaoItemOV.dIPIPercRedBase
'        objItemPV.objTributacaoItemPV.dIPIValor = .objTributacaoItemOV.dIPIValor
'        objItemPV.objTributacaoItemPV.iICMSTipo = .objTributacaoItemOV.iICMSTipo
'        objItemPV.objTributacaoItemPV.iIPITipo = .objTributacaoItemOV.iIPITipo
'        objItemPV.objTributacaoItemPV.iTipoTributacao = .objTributacaoItemOV.iTipoTributacao
'        objItemPV.objTributacaoItemPV.sNaturezaOp = .objTributacaoItemOV.sNaturezaOp
'
'        objItemPV.objTributacaoItemPV.iNaturezaOpManual = .objTributacaoItemOV.iNaturezaOpManual
'        objItemPV.objTributacaoItemPV.iTipoTributacaoManual = .objTributacaoItemOV.iTipoTributacaoManual
'        objItemPV.objTributacaoItemPV.iIPITipoManual = .objTributacaoItemOV.iIPITipoManual
'        objItemPV.objTributacaoItemPV.iIPIBaseManual = .objTributacaoItemOV.iIPIBaseManual
'        objItemPV.objTributacaoItemPV.iIPIPercRedBaseManual = .objTributacaoItemOV.iIPIPercRedBaseManual
'        objItemPV.objTributacaoItemPV.iIPIAliquotaManual = .objTributacaoItemOV.iIPIAliquotaManual
'        objItemPV.objTributacaoItemPV.iIPIValorManual = .objTributacaoItemOV.iIPIValorManual
'        objItemPV.objTributacaoItemPV.iICMSTipoManual = .objTributacaoItemOV.iICMSTipoManual
'        objItemPV.objTributacaoItemPV.iICMSBaseManual = .objTributacaoItemOV.iICMSBaseManual
'        objItemPV.objTributacaoItemPV.iICMSPercRedBaseManual = .objTributacaoItemOV.iICMSPercRedBaseManual
'        objItemPV.objTributacaoItemPV.iICMSAliquotaManual = .objTributacaoItemOV.iICMSAliquotaManual
'        objItemPV.objTributacaoItemPV.iICMSValorManual = .objTributacaoItemOV.iICMSValorManual
'        objItemPV.objTributacaoItemPV.iICMSSubstBaseManual = .objTributacaoItemOV.iICMSSubstBaseManual
'        objItemPV.objTributacaoItemPV.iICMSSubstAliquotaManual = .objTributacaoItemOV.iICMSSubstAliquotaManual
'        objItemPV.objTributacaoItemPV.iICMSSubstValorManual = .objTributacaoItemOV.iICMSSubstValorManual
'
'    End With
'
'End Sub

Sub Atualiza_Grid_Alocacao(objItemPV As ClassItemPedido)

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

'    Call Remove_Reservas(objItemPV.iItem)
'
'    For Each objItemRomaneio In objItemPV.colItensRomaneioGrade
'
'        objProduto.sCodigo = objItemPV.sProduto
'        'Lê o produto
'        lErro = CF("Produto_Le", objProduto)
'        If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
'        If lErro = 28030 Then gError 42765 'Não encontrou
'        'Faz a conversão da unidade do item para a unidade de estoque
'        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRomaneio.sUMEstoque, objProduto.sSiglaUMEstoque, dFator)
'        If lErro <> SUCESSO Then gError 42766
'
'        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemPV.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator2)
'        If lErro <> SUCESSO Then gError 42766
'
'        dQuantReservada = 0
'
'        For Each objReserva In objItemRomaneio.colLocalizacao
'
'            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Item_Col) = objItemPV.iItem
'
'            lErro = Mascara_MascararProduto(objItemRomaneio.sProduto, sProdutoMascarado)
'            If lErro <> SUCESSO Then gError 86357
'
'            objAlmoxarifado.iCodigo = objReserva.iAlmoxarifado
'
'            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
'            If lErro <> SUCESSO And lErro <> 25056 Then gError 46181
'            If lErro = 25056 Then gError 46182
'
'            objReserva.sAlmoxarifado = objAlmoxarifado.sNomeReduzido
'
'            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
'            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_ProdutoAlmox_Col) = sProdutoMascarado
'            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_UMEstoque_Col) = objItemRomaneio.sUMEstoque
'            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objReserva.sAlmoxarifado
'            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReservar_Col) = Formata_Estoque((objItemRomaneio.dQuantidade - objItemRomaneio.dQuantCancelada) * dFator2)
'            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReserv_Col) = Formata_Estoque(objReserva.dQuantidade * dFator)
'            If objReserva.dtDataValidade <> DATA_NULA Then GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Validade_Col) = Format(objReserva.dtDataValidade, "dd/mm/yyyy")
'            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Responsavel_Col) = objReserva.sResponsavel
'
'            objGridAlocacao.iLinhasExistentes = objGridAlocacao.iLinhasExistentes + 1
'
'        Next
'
'    Next

    Exit Sub

Erro_Atualiza_Grid_Alocacao:

    Select Case gErr

        Case 42764, 42766, 46181, 86357

        Case 42765
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", objProduto.sCodigo, gErr)

        Case 46182
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158087)

    End Select

    Exit Sub

End Sub

Sub Atualiza_Grid_Itens(objItemPV As ClassItemPedido)

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim dQuantidade As Double
Dim dQuantCancelada As Double
Dim dQuantReservada As Double
Dim dQuantFaturada As Double
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem

'    For Each objItemRomaneioGrade In objItemPV.colItensRomaneioGrade
'
'        dQuantidade = dQuantidade + objItemRomaneioGrade.dQuantidade
'        dQuantCancelada = dQuantCancelada + objItemRomaneioGrade.dQuantCancelada
'        dQuantReservada = dQuantReservada + objItemRomaneioGrade.dQuantReservada
'        dQuantFaturada = dQuantFaturada + objItemRomaneioGrade.dQuantFaturada
'
'    Next
'
'    GridItens.TextMatrix(objItemPV.iItem, iGrid_QuantPedida_Col) = Formata_Estoque(dQuantidade)
'    GridItens.TextMatrix(objItemPV.iItem, iGrid_QuantCancel_Col) = Formata_Estoque(dQuantCancelada)
'    GridItens.TextMatrix(objItemPV.iItem, iGrid_QuantReservada_Col) = Formata_Estoque(dQuantReservada)
'    GridItens.TextMatrix(objItemPV.iItem, iGrid_QuantFaturada_Col) = Formata_Estoque(dQuantFaturada)
'
'    objItemPV.dQuantidade = dQuantidade
'    objItemPV.dQuantCancelada = dQuantCancelada

    Exit Sub

End Sub

Private Function Reserva_Processa_ItemGrade(objItemRomaneioGrade As ClassItemRomaneioGrade, ByVal iItem As Integer) As Long

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************


Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dQuantidade As Double, dQuantidadeCancelada As Double, dQuantidadeFaturada As Double

On Error GoTo Erro_Reserva_Processa_ItemGrade

'    dQuantidade = objItemRomaneioGrade.dQuantidade
'    dQuantidadeCancelada = objItemRomaneioGrade.dQuantCancelada
'    dQuantidadeFaturada = objItemRomaneioGrade.dQuantFaturada
'
'    'Se não houve Quant a Reservar zera Reservas.
'    If dQuantidade = 0 Or dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada = 0 Then
'        objItemRomaneioGrade.dQuantReservada = 0
'    End If
'
'    objProduto.sCodigo = objItemRomaneioGrade.sProduto
'
'    'Lê o Produto
'    lErro = CF("Produto_Le", objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then gError 26669
'
'    'Se não achou o Produto --> erro
'    If lErro = 28030 Then gError 26668
'
'    objItemRomaneioGrade.sUMEstoque = objProduto.sSiglaUMEstoque
'
'    'Se controla Reserva e há quantidade a reservar gera reserva automático
'    If objProduto.iControleEstoque = PRODUTO_CONTROLE_RESERVA And dQuantidade > 0 And dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada > 0 And gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA Then
'
'        'tenta reservar no Almoxarifado padrão
'        lErro = ReservaAlmoxarifadoPadrao_Grade(objProduto, objItemRomaneioGrade, dQuantidade, dQuantidadeCancelada, dQuantidadeFaturada, iItem)
'        If lErro <> SUCESSO And lErro <> 26680 Then gError 26679
'
'    'Se não há quantidade a reservar
'    ElseIf dQuantidade = 0 And dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada = 0 And gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA Then
'
'        'Elimina as reservas feitas anteriormente
'        Set objItemRomaneioGrade.colLocalizacao = New Collection
'        objItemRomaneioGrade.dQuantReservada = 0
'
'    End If

    Reserva_Processa_ItemGrade = SUCESSO

    Exit Function

Erro_Reserva_Processa_ItemGrade:

    Reserva_Processa_ItemGrade = gErr

    Select Case gErr

        Case 26669, 26679

        Case 26668
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158089)

    End Select

    Exit Function

End Function

Function ReservaAlmoxarifadoPadrao_Grade(objProduto As ClassProduto, objItemRomaneioGrade As ClassItemRomaneioGrade, dQuantidade As Double, dQuantCancelada As Double, dQuantFaturada As Double, ByVal iItem As Integer) As Long

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim lErro As Long
Dim dQuantidadeReservarVenda As Double
Dim dQuantReservadaPedido As Double
Dim dFator As Double
Dim iAlmoxarifadoPadrao As Integer
Dim objItemPV As New ClassItemPedido
Dim objReservaBD As New ClassReserva
Dim objReservaItem As ClassReservaItem
Dim objReserva As ClassReserva
Dim colItemPedido As New colItemPedido
Dim objAlmoxarifadoPadrao As New ClassAlmoxarifado
Dim objAlmoxarifadoPadrao1 As New ClassAlmoxarifado
Dim iFilialEmpresa As Integer
Dim iFilialFaturamento As Integer
Dim dQuantidadeReservarEstoque1 As Double 'Reservar Filial 1
Dim dQuantidadeReservarEstoque2 As Double 'Reservar Filial 2
Dim objEstoqueProduto1 As New ClassEstoqueProduto 'Reservar Filial 1
Dim objEstoqueProduto2 As New ClassEstoqueProduto 'Reservar Filial 2

On Error GoTo Erro_ReservaAlmoxarifadoPadrao_Grade

'    'calula a quantidade reservada
'    dQuantidadeReservarVenda = (dQuantidade - dQuantCancelada - dQuantFaturada)
'
'    lErro = CF("UM_Conversao", objProduto.iClasseUM, gobjPedidoVenda.colItensPedido(iItem).sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
'    If lErro <> SUCESSO Then gError 26670
'
'    'Converte a quantidade para a unidade de Estoque
'    dQuantidadeReservarEstoque1 = dQuantidadeReservarVenda * dFator
'
'    iFilialFaturamento = Codigo_Extrai(FilialFaturamento.Text)
'
'    If iFilialFaturamento = 0 Then
'        'Busca a filial de faturamento
'        lErro = CF("FilialFaturamento_Le", giFilialEmpresa, iFilialFaturamento)
'        If lErro <> SUCESSO Then gError 30414
'        If iFilialFaturamento = 0 Then iFilialFaturamento = giFilialEmpresa
'    End If
'
'    'Busca o almoxarifado padrão
'    lErro = CF("AlmoxarifadoPadrao_Le", iFilialFaturamento, objProduto.sCodigo, iAlmoxarifadoPadrao)
'    If lErro <> SUCESSO And lErro <> 23796 Then gError 51417
'    If iAlmoxarifadoPadrao = 0 Then gError 51418
'
'    objAlmoxarifadoPadrao.iCodigo = iAlmoxarifadoPadrao
'    'Lê o almoxarifado padrão
'    lErro = CF("Almoxarifado_Le", objAlmoxarifadoPadrao)
'    If lErro = SUCESSO Then
'
'        objEstoqueProduto1.sProduto = objProduto.sCodigo
'        objEstoqueProduto1.iAlmoxarifado = iAlmoxarifadoPadrao
'        objEstoqueProduto1.sAlmoxarifadoNomeReduzido = objAlmoxarifadoPadrao.sNomeReduzido
'
'        'Lê o estoque do produto
'        lErro = CF("EstoqueProduto_Le", objEstoqueProduto1)
'        If lErro <> SUCESSO And lErro <> 21306 Then gError 26672
'        If lErro = 21306 Then gError 26673
'
'        objItemPV.lCodPedido = StrParaLong(Codigo.Text)
'        objItemPV.sProduto = objProduto.sCodigo
'        objItemPV.iFilialEmpresa = giFilialEmpresa
'        objItemPV.sProdutoDescricao = objProduto.sDescricao
'        objReservaBD.iAlmoxarifado = iAlmoxarifadoPadrao
'
'        'Lê as reservas do produto
'        lErro = CF("ReservaItem_Le", objItemPV, objReservaBD)
'        If lErro <> SUCESSO And lErro <> 26678 Then gError 26674
'
'        dQuantReservadaPedido = objReservaBD.dQuantidade
'
'        'Saldo enxergado por esse Pedido Venda
'        objEstoqueProduto1.dSaldo = objEstoqueProduto1.dQuantDisponivel + dQuantReservadaPedido
'
'        'Se quantidade a reservar for menor que a disponível coloca reserva no GridAlocacao
'        If (dQuantidadeReservarEstoque1 - objEstoqueProduto1.dSaldo) < QTDE_ESTOQUE_DELTA Then
'
'            objItemRomaneioGrade.dQuantReservada = dQuantidadeReservarVenda
'
'            Set objItemRomaneioGrade.colLocalizacao = New Collection
'
'            Set objReservaItem = New ClassReservaItem
'
'            objReservaItem.iAlmoxarifado = objEstoqueProduto1.iAlmoxarifado
'            objReservaItem.sAlmoxarifado = objEstoqueProduto1.sAlmoxarifadoNomeReduzido
'            objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarVenda)
'            objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarEstoque1)
'            objReservaItem.sResponsavel = RESERVA_AUTO_RESP
'            objReservaItem.dtDataValidade = DATA_NULA
'
'            objItemRomaneioGrade.colLocalizacao.Add objReservaItem
'
'        Else
'
'            lErro = CF("Retorna_Almoxarifado_Alternativo", iAlmoxarifadoPadrao, iFilialEmpresa)
'            If lErro <> SUCESSO Then gError 105004
'
'            If iAlmoxarifadoPadrao = 0 Then GoTo FaltaEstoque
'
'            'Verifica quanto falta para reservar
'            dQuantidadeReservarEstoque2 = dQuantidadeReservarEstoque1 - objEstoqueProduto1.dSaldo
'            dQuantidadeReservarEstoque1 = objEstoqueProduto1.dSaldo
'
'            'Busca na Filial 2
'            objAlmoxarifadoPadrao1.iCodigo = iAlmoxarifadoPadrao
'
'            'Lê o almoxarifado
'            lErro = CF("Almoxarifado_Le", objAlmoxarifadoPadrao1)
'            If lErro = SUCESSO Then
'
'                objEstoqueProduto2.sProduto = objProduto.sCodigo
'                objEstoqueProduto2.iAlmoxarifado = iAlmoxarifadoPadrao
'                objEstoqueProduto2.sAlmoxarifadoNomeReduzido = objAlmoxarifadoPadrao1.sNomeReduzido
'
'                'Lê o estoque do produto
'                lErro = CF("EstoqueProduto_Le", objEstoqueProduto2)
'                If lErro <> SUCESSO And lErro <> 21306 Then gError 26672
'                If lErro = 21306 Then gError 26673
'
'                objItemPV.lCodPedido = StrParaLong(Codigo)
'                objItemPV.sProduto = objProduto.sCodigo
'                objItemPV.iFilialEmpresa = iFilialEmpresa
'                objItemPV.sProdutoDescricao = objProduto.sDescricao
'                objReservaBD.iAlmoxarifado = iAlmoxarifadoPadrao
'
'                'Lê as reservas do produto
'                lErro = CF("ReservaItem_Le", objItemPV, objReservaBD)
'                If lErro <> SUCESSO And lErro <> 26678 Then gError 26674
'
'                dQuantReservadaPedido = objReservaBD.dQuantidade
'
'                'Saldo enxergado por esse Pedido Venda
'                objEstoqueProduto2.dSaldo = objEstoqueProduto2.dQuantDisponivel + dQuantReservadaPedido
'
'                'Se quantidade a reservar for menor que a disponível coloca reserva no GridAlocacao
'                If (dQuantidadeReservarEstoque2 - objEstoqueProduto2.dSaldo) < QTDE_ESTOQUE_DELTA Then
'
'                    If dQuantidadeReservarEstoque1 > 0 Then
'
'                        Set objReservaItem = New ClassReservaItem
'
'                        'Parte reservada na filial 1
'                        objReservaItem.iAlmoxarifado = objEstoqueProduto1.iAlmoxarifado
'                        objReservaItem.sAlmoxarifado = objEstoqueProduto1.sAlmoxarifadoNomeReduzido
'                        objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarVenda)
'                        objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarEstoque1)
'                        objReservaItem.sResponsavel = RESERVA_AUTO_RESP
'                        objReservaItem.dtDataValidade = DATA_NULA
'
'                        objItemRomaneioGrade.colLocalizacao.Add objReservaItem
'
'                    End If
'
'                    Set objReservaItem = New ClassReservaItem
'
'                    'Parte Reservada na Filial 2
'                        objReservaItem.iAlmoxarifado = objEstoqueProduto2.iAlmoxarifado
'                        objReservaItem.sAlmoxarifado = objEstoqueProduto2.sAlmoxarifadoNomeReduzido
'                        objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarVenda)
'                        objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarEstoque2)
'                        objReservaItem.dtDataValidade = DATA_NULA
'                        objReservaItem.sResponsavel = RESERVA_AUTO_RESP
'
'                        objItemRomaneioGrade.colLocalizacao.Add objReservaItem
'
'                'Caso contrário limpa as reservas desse ítem e chama tela de Falta de Estoque
'                Else
'
'FaltaEstoque:
'                    Set objItemPV = New ClassItemPedido
'                    objItemPV.dQuantidade = dQuantidade
'                    objItemPV.dQuantReservada = 0
'                    objItemPV.dQuantCancelada = dQuantCancelada
'                    objItemPV.dQuantFaturada = dQuantFaturada
'                    objItemPV.sProduto = objProduto.sCodigo
'                    objItemPV.lCodPedido = StrParaLong(Codigo.Text)
'                    objItemPV.iItem = iItem
'                    objItemPV.lNumIntDoc = gcolItemPedido(iItem)
'                    objItemPV.sUMEstoque = objProduto.sSiglaUMEstoque
'                    objItemPV.sProdutoDescricao = objProduto.sDescricao
'                    objItemPV.sUnidadeMed = gobjPedidoVenda.colItensPedido(iItem).sUnidadeMed
'                    objItemPV.iClasseUM = objProduto.iClasseUM
'                    objItemPV.iPossuiGrade = MARCADO
'
'                    'Chama tela de Falta de Estoque
'                    lErro = Chama_Tela_Modal("FaltaEstoque", objItemPV, colItemPedido, dQuantidadeReservarEstoque1, objAlmoxarifadoPadrao, objEstoqueProduto1.dSaldo)
'
'                    'Se retornar Cancela erro
'                    If giRetornoTela = vbCancel Then gError 26680
'
'                    'Limpa reservas desse ítem no GridAlocacao
'                    Set objItemRomaneioGrade.colLocalizacao = New Collection
'
'                    If giRetornoTela = vbOK Then
'
'                       'Se não substituiu o Produto
'                        'Coloca QuantReservada e QuantCancelada no ítem do GridItens
'                        objItemRomaneioGrade.dQuantReservada = objItemPV.dQuantReservada
'
'                        objItemRomaneioGrade.dQuantCancelada = objItemPV.dQuantCancelada
'
'                        'Coloca reseravas desse ítem no GridAlocacao
'                        For Each objReserva In objItemPV.colReserva
'                            Set objReservaItem = New ClassReservaItem
'
'                            objReservaItem.iAlmoxarifado = objReserva.iAlmoxarifado
'                            objReservaItem.dQuantidade = objReserva.dQuantidade
'                            objReservaItem.dtDataValidade = objReserva.dtDataValidade
'                            objReservaItem.sAlmoxarifado = objReserva.sAlmoxarifado
'                            objReservaItem.sResponsavel = objReserva.sResponsavel
'
'                            objItemRomaneioGrade.colLocalizacao.Add objReservaItem
'
'                        Next
'                    End If
'                End If
'            End If
'        End If
'    End If

    ReservaAlmoxarifadoPadrao_Grade = SUCESSO

    Exit Function

Erro_ReservaAlmoxarifadoPadrao_Grade:

    ReservaAlmoxarifadoPadrao_Grade = gErr

    Select Case gErr

        Case 26670, 26672, 26674, 26682, 26772, 30414, 51417, 105004

        Case 26681
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 26673
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto1.sProduto, objEstoqueProduto1.iAlmoxarifado)

        Case 26680
            Call Rotina_Erro(vbOKOnly, "ERRO_RESERVA_NAO_DECIDIDA", gErr, objProduto.sCodigo)

        Case 51418
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO", gErr, objProduto.sCodigo, giFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158090)

    End Select

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 193362

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 193363

        'Não encontrou Transportadora no BD
        If lErro <> SUCESSO Then gError 193364

        'Encontrou Transportadora no BD, coloca no Text da Combo
        TranspRedespacho.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 193365

    Exit Sub

Erro_TranspRedespacho_Validate:

    Cancel = True

    Select Case gErr

        Case 193362, 193363

        Case 193364  'Não encontrou Transportadora no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            End If

        Case 193365
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, TranspRedespacho.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193366)

    End Select

    Exit Sub

End Sub

Public Sub RedespachoCli_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Extrai

    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_PEDIDOVENDA
    If Len(Trim(Codigo.Text)) > 0 Then
        objAnotacao.sID = CStr(giFilialEmpresa) & "," & Codigo.Text
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

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193367)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193368)

    End Select

    Exit Sub

End Sub

Public Sub DataEntregaPV_Change()
    iAlterado = REGISTRO_ALTERADO
    giDataEntregaAlterada = REGISTRO_ALTERADO
End Sub

Public Sub DataEntregaPV_GotFocus()

Dim iDataAux As Integer

    iDataAux = giDataEntregaAlterada
    Call MaskEdBox_TrataGotFocus(DataEntregaPV, iAlterado)
    giDataEntregaAlterada = iDataAux

End Sub

Public Sub DataEntregaPV_Validate(Cancel As Boolean)

Dim lErro As Long, iIndice As Integer, dtData As Date

On Error GoTo Erro_DataEntregaPV_Validate

    If giDataEntregaAlterada <> REGISTRO_ALTERADO Then Exit Sub

    'Verifica se a Data foi digitada
    If Len(Trim(DataEntregaPV.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEntregaPV.Text)
    If lErro <> SUCESSO Then gError 193369

    dtData = MaskedParaDate(DataEntregaPV)

    For iIndice = 1 To objGridItensSRV.iLinhasExistentes
        If dtData <> DATA_NULA Then
            GridItensSRV.TextMatrix(iIndice, iGrid_DataEntregaSRV_Col) = Format(dtData, "dd/mm/yyyy")
        Else
            GridItensSRV.TextMatrix(iIndice, iGrid_DataEntregaSRV_Col) = ""
        End If
    Next

    If gobjPedidoVenda.dtDataEntrega <> StrParaDate(DataEntregaPV.Text) Then

        gobjPedidoVenda.dtDataEntrega = StrParaDate(DataEntregaPV.Text)

        Call ValorTotal_Calcula

    End If

    giDataEntregaAlterada = 0

    Exit Sub

Erro_DataEntregaPV_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 193369

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193370)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntregaPV_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntregaPV_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEntregaPV, DIMINUI_DATA)
    If lErro Then gError 193371

    Exit Sub

Erro_UpDownEntregaPV_DownClick:

    Select Case gErr

        Case 193371

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193372)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntregaPV_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntregaPV_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEntregaPV, AUMENTA_DATA)
    If lErro Then gError 193373

    Exit Sub

Erro_UpDownEntregaPV_UpClick:

    Select Case gErr

        Case 193373

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193374)

    End Select

    Exit Sub

End Sub

'Incluída por Luiz Nogueira em 26/01/04
Private Function Cadastra_Cliente_Inexistente(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Pergunta ao usuário se deseja cadastrar um cliente inexiste e chama a tela de cliente
'Serve para quando o pedido é criado a partir de um orçamento onde o cliente não foi cadastrado

Dim objCliente As ClassCliente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Cadastra_Cliente_Inexistente

    'Se o cliente que está na tela não está cadastrado no BD
    If giClienteInexistente = CLIENTE_INEXISTENTE Then

        'Pergunta ao usuário se deseja cadastrar o cliente
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_ORCAMENTO", Trim(Cliente.Text), Orcamento.Text)

        'Se o usuário respondeu que sim
        If vbMsgRes = vbYes Then

            'Cria uma nova instância da classe cliente
            Set objCliente = New ClassCliente

            'Guarda o nome reduzido do cliente no obj
            objCliente.sNomeReduzido = Trim(Cliente.Text)

            'Chama a tela de clientes
            Call Chama_Tela("Clientes", objCliente)
        End If

    End If

    'Se a filial do cliente não está cadastrada => erro
    If giFilialCliInexistente = FILIALCLIENTE_INEXISTENTE Then gError 193375

    'Se a filial de entrega não está cadastrada => erro
    If giFilialCliInexistente = FILIALENTREGA_INEXISTENTE Then gError 193376

    giFilialCliInexistente = 0
    giClienteInexistente = 0

    Cadastra_Cliente_Inexistente = SUCESSO

    Exit Function

Erro_Cadastra_Cliente_Inexistente:

    Select Case gErr

        Case 193375
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case 193376
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_INEXISTENTE", gErr, FilialEntrega.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193377)

    End Select

End Function

Private Function Traz_Orcamento_Cliente(objCliente As ClassCliente, ByVal objPedidoVenda As ClassPedidoDeVenda)
'Traz para tela os dados defaults do cliente, quando a tela é preenchida a partir de orçamento
'Dispara o cálculo de comissões caso a empresa não utilize o sistema de regras de comissões

Dim lErro As Long
Dim objMensagem As ClassMensagem
Dim objTipoCliente As ClassTipoCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objTransportadora As New ClassTransportadora

On Error GoTo Erro_Traz_Orcamento_Cliente

    '**** LEITURA DO CLIENTE ****
    'Se ainda não leu os dados do cliente durante a carga do pedido para a tela
    If objCliente Is Nothing Then

        'Instancia objCliente
        Set objCliente = New ClassCliente

        'Guarda o código do cliente no obj
        objCliente.lCodigo = objPedidoVenda.lCliente

        'Lê os dados do cliente
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 193378

        'Se não encontrou o cliente => erro
        If lErro = 12348 Then gError 193379

    End If
    '*****************************

    '**** LEITURA DA FILIAL DO CLIENTE *****
    'Guarda em objFilialCliente os dados necessários para ler a filial
    objFilialCliente.lCodCliente = objCliente.lCodigo
    objFilialCliente.iCodFilial = objPedidoVenda.iFilial

    'Lê a Filial cliente para obter os dados default que serão exibidos na tela
    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 12567 Then gError 193380

    'Se não encontrou a filial => erro
    If lErro = 12567 Then gError 193381
    '*****************************************

    '**** TRATAMENTO PARA TRANSPORTADORA DEFAULT ****
    'Guarda o código da transportadora em objtransportadora
    objTransportadora.iCodigo = objFilialCliente.iCodTransportadora

    'Se ainda não encontrou Transportadora default
    If objTransportadora.iCodigo = 0 Then

        'Instancia objTipoCliente
        Set objTipoCliente = New ClassTipoCliente

        'Guarda o tipo de cliente no obj
        objTipoCliente.iCodigo = objCliente.iTipo

        'Lê o tipo de cliente para checar se o tipo tem transportadora default
        lErro = CF("TipoCliente_Le", objTipoCliente)
        If lErro <> SUCESSO And lErro <> 19062 Then gError 193382

        'Se não encontrou => erro
        If lErro = 19062 Then gError 193383

        'Guarda o código da transportadora em objtransportadora
        objTransportadora.iCodigo = objTipoCliente.iCodTransportadora

    End If

    'Se encontrou uma transportadora default
    If objTransportadora.iCodigo <> 0 Then

'        'Lê os dados da transportadora
'        lErro = CF("Transportadora_Le", objTransportadora)
'        If lErro <> SUCESSO And lErro <> 19250 Then gError 51063
'        If lErro = SUCESSO Then
            Transportadora.Text = objTransportadora.iCodigo
            Call Transportadora_Validate(False)
'        End If
    End If
    '*****************************************

    '**** TRATAMENTO PARA TIPO DE FRETE ****
    'Tipo de frete
    If objFilialCliente.iTipoFrete = TIPO_FOB Then
        Destinatario.Value = True
    ElseIf objFilialCliente.iTipoFrete = TIPO_CIF Then
        Emitente.Value = True
    End If
    '*****************************************

    '**** TRATAMENTO PARA MENSAGEM DE NOTA FISCAL ****
    'Se existe mensagem de nota fiscal para o cliente
    If objCliente.iCodMensagem > 0 Then

        'Instancia objMensagem
        Set objMensagem = New ClassMensagem

        'Guarda o código da mensagem no obj
        objMensagem.iCodigo = objCliente.iCodMensagem

        'Lê a mensagem do Cliente
        lErro = CF("Mensagem_Le", objMensagem)
        If lErro <> SUCESSO And lErro <> 19234 Then gError 193384

        'Se não encontrou a mensagem => erro
        If lErro = 19234 Then gError 193385

        'Exibe a mensagem na tela
        Mensagem.Text = objMensagem.sDescricao

    End If
    '**************************************************

    '*** TRATAMENTO PARA COMISSÕES ***
    'Se a empresa não utiliza as regras para cálculo de comissões, calcula as comissões no modelo antigo
    'Caso a empresa utilize as regras, o cáclulo será efetuado ao clicar no tab comissões ou ao gravar o pedido
    If Not (gobjCRFAT.iUsaComissoesRegras = USA_REGRAS) Then

        'calcular as comissoes no modelo antigo...
        lErro = objTabComissoes.Comissao_Automatica_FilialCli_Exibe(objFilialCliente)
        If lErro <> SUCESSO Then gError 193386

    End If
    '*********************************

    Traz_Orcamento_Cliente = SUCESSO

    Exit Function

Erro_Traz_Orcamento_Cliente:

    Traz_Orcamento_Cliente = gErr

    Select Case gErr

        Case 193378, 193380, 193382, 193384, 193386

        Case 193379
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)

        Case 193381
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_INEXISTENTE", gErr, objCliente.lCodigo, objFilialCliente.iCodFilial)

        Case 193383
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo, objCliente.iTipo)

        Case 193385
            Call Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_CADASTRADA", gErr, objCliente.lCodigo, objCliente.iCodMensagem)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193387)

    End Select

End Function
'********************************

'Public Sub PISRetido_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iPISRetidoAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PISRetido_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValor As Double
'Dim dValorTotal As Double
'
'On Error GoTo Erro_PISRetido_Validate
'
'    If iPISRetidoAlterado = 0 Then Exit Sub
'
'    'Verifica se foi preenchido
'    If Len(Trim(PISRetido.Text)) > 0 Then
'
'        'Critica o Valor
'        lErro = Valor_NaoNegativo_Critica(PISRetido.Text)
'        If lErro <> SUCESSO Then gError 193388
'
'        dValor = CDbl(PISRetido.Text)
'
'        PISRetido.Text = Format(dValor, "Standard")
'
'        If Len(Trim(ValorTotal.Caption)) > 0 Then dValorTotal = CDbl(ValorTotal.Caption)
'
'        If dValor > dValorTotal Then gError 193389
'
'    End If
'
'    Call BotaoGravarTrib
'
'    iPISRetidoAlterado = 0
'
'    Exit Sub
'
'Erro_PISRetido_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 193388
'
'        Case 193389
'            Call Rotina_Erro(vbOKOnly, "ERRO_PIS_RETIDO_MAIOR_VALOR_TOTAL", gErr, dValor, dValorTotal)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193390)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub COFINSRetido_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iCOFINSRetidoAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub COFINSRetido_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValor As Double
'Dim dValorTotal As Double
'
'On Error GoTo Erro_COFINSRetido_Validate
'
'    If iCOFINSRetidoAlterado = 0 Then Exit Sub
'
'    'Verifica se foi preenchido
'    If Len(Trim(COFINSRetido.Text)) > 0 Then
'
'        'Critica o Valor
'        lErro = Valor_NaoNegativo_Critica(COFINSRetido.Text)
'        If lErro <> SUCESSO Then gError 193391
'
'        dValor = CDbl(COFINSRetido.Text)
'
'        COFINSRetido.Text = Format(dValor, "Standard")
'
'        If Len(Trim(ValorTotal.Caption)) > 0 Then dValorTotal = CDbl(ValorTotal.Caption)
'
'        If dValor > dValorTotal Then gError 193392
'
'    End If
'
'    Call BotaoGravarTrib
'
'    iCOFINSRetidoAlterado = 0
'
'    Exit Sub
'
'Erro_COFINSRetido_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 193391
'
'        Case 193392
'            Call Rotina_Erro(vbOKOnly, "ERRO_COFINS_RETIDO_MAIOR_VALOR_TOTAL", gErr, dValor, dValorTotal)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193393)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub CSLLRetido_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iCSLLRetidoAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub CSLLRetido_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValor As Double
'Dim dValorTotal As Double
'
'On Error GoTo Erro_CSLLRetido_Validate
'
'    If iCSLLRetidoAlterado = 0 Then Exit Sub
'
'    'Verifica se foi preenchido
'    If Len(Trim(CSLLRetido.Text)) > 0 Then
'
'        'Critica o Valor
'        lErro = Valor_NaoNegativo_Critica(CSLLRetido.Text)
'        If lErro <> SUCESSO Then gError 193394
'
'        dValor = CDbl(CSLLRetido.Text)
'
'        CSLLRetido.Text = Format(dValor, "Standard")
'
'        If Len(Trim(ValorTotal.Caption)) > 0 Then dValorTotal = CDbl(ValorTotal.Caption)
'
'        If dValor > dValorTotal Then gError 193395
'
'    End If
'
'    Call BotaoGravarTrib
'
'    iCSLLRetidoAlterado = 0
'
'    Exit Sub
'
'Erro_CSLLRetido_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 193394
'
'        Case 193395
'            Call Rotina_Erro(vbOKOnly, "ERRO_CSLL_RETIDO_MAIOR_VALOR_TOTAL", gErr, dValor, dValorTotal)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193396)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Private Sub PedidoVenda_ExibirBloqs(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal colBloqPVAux As Collection, bExibirBloqs As Boolean)
''retorna em bExibirBloqsPV se deve abrir a tela que mostra os bloqueios de pv
'
'Dim lErro As Long, objBloqPV As ClassBloqueioPV, objBloqPVAux As ClassBloqueioPV
'Dim bCriado As Boolean
'
'On Error GoTo Erro_PedidoVenda_ExibirBloqs
'
'    'verifica se algum bloqueio automatico foi criado durante a gravacao
'    bCriado = False
'
'    For Each objBloqPV In objPedidoVenda.colBloqueiosPV
'
'        If objBloqPV.dtDataLib = DATA_NULA Then
'
'            bCriado = True
'
'            For Each objBloqPVAux In colBloqPVAux
'
'                If objBloqPVAux.iTipoDeBloqueio = objBloqPV.iTipoDeBloqueio Then
'
'                    bCriado = False
'                    Exit For
'
'                End If
'
'            Next
'
'            If bCriado Then Exit For
'
'        End If
'
'    Next
'
'    bExibirBloqs = bCriado
'
'    Exit Sub
'
'Erro_PedidoVenda_ExibirBloqs:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193397)
'
'    End Select
'
'    Exit Sub
'
'End Sub

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
        If lErro <> SUCESSO Then gError 193398

        If iPreenchido <> PRODUTO_PREENCHIDO Then gError 193399

        'preenche o código do produto
        objProduto.sCodigo = sProduto

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 193400

        If lErro = 28030 Then gError 193401

        'Realiza a converção para uma mesma UM
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 193402

        'Realiza o cálculo do volume
        dVolume = dVolume + (((StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col)) - StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))) * dFator) * (objProduto.dComprimento * objProduto.dEspessura * objProduto.dLargura))

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

        Case 193398, 193400, 193402

        Case 193399
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 193401
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193403)

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
    If lErro <> SUCESSO Then gError 193404

    'Coloca o valor formatado na Tela
    Cubagem.Text = Format(Cubagem.Text, "Fixed")

    Exit Sub

Erro_Cubagem_Validate:

    Cancel = True

    Select Case gErr

        Case 193404

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193405)

    End Select

    Exit Sub

End Sub

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
    If lErro <> SUCESSO Then gError 193406

    objProduto.sCodigo = sProdutoFormatado

    'Lê o produto da linha passada por iLinha do GridItens
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 193407
    If lErro = 28030 Then gError 193408

    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUmAnterior, sUMNova, dFator)
    If lErro <> SUCESSO Then gError 193409

    'Atualiza o Grid
    If StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_QuantPedida_Col)) * dFator > QTDE_ESTOQUE_DELTA Then
        GridItens.TextMatrix(iLinha, iGrid_QuantPedida_Col) = Formata_Estoque(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_QuantPedida_Col)) * dFator)
    Else
        GridItens.TextMatrix(iLinha, iGrid_QuantPedida_Col) = ""
    End If

    If StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_QuantCancel_Col)) * dFator > QTDE_ESTOQUE_DELTA Then
        GridItens.TextMatrix(iLinha, iGrid_QuantCancel_Col) = Formata_Estoque(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_QuantCancel_Col)) * dFator)
    Else
        GridItens.TextMatrix(iLinha, iGrid_QuantCancel_Col) = ""
    End If

    'Preço unitário
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))

    'Recalcula preço do ítem e valor total da nota
    If dPrecoUnitario > 0 Then
        GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario / dFator, gobjFAT.sFormatoPrecoUnitario)
        Call CF2(Me, "PSRV_PrecoTotal_Calcula", iLinha)
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 193410
    End If

    Atualiza_UM = SUCESSO

    Exit Function

Erro_Atualiza_UM:

    Atualiza_UM = gErr

    Select Case gErr

        Case 193406, 193407, 193409, 193410

        Case 193408
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193411)

    End Select

    Exit Function

End Function

Private Function Atualiza_UMSRV(ByVal iLinha As Integer, ByVal sUmAnterior As String, ByVal sUMNova As String) As Long
'Atualiza quantidades em funcao de troca de UM

Dim lErro As Long
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double, dPrecoUnitario As Double
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Atualiza_UMSRV

    lErro = CF("Produto_Formata", GridItensSRV.TextMatrix(iLinha, iGrid_ServicoSRV_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 191088

    objProduto.sCodigo = sProdutoFormatado

    'Lê o produto da linha passada por iLinha do GridItens
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 191089
    
    If lErro = 28030 Then gError 191090

    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUmAnterior, sUMNova, dFator)
    If lErro <> SUCESSO Then gError 191091

    'Atualiza o Grid
    If StrParaDbl(GridItensSRV.TextMatrix(iLinha, iGrid_QuantPedidaSRV_Col)) * dFator > QTDE_ESTOQUE_DELTA Then
        GridItensSRV.TextMatrix(iLinha, iGrid_QuantPedidaSRV_Col) = Formata_Estoque(StrParaDbl(GridItensSRV.TextMatrix(iLinha, iGrid_QuantPedidaSRV_Col)) * dFator)
    Else
        GridItensSRV.TextMatrix(iLinha, iGrid_QuantPedidaSRV_Col) = ""
    End If

    If StrParaDbl(GridItensSRV.TextMatrix(iLinha, iGrid_QuantCancelSRV_Col)) * dFator > QTDE_ESTOQUE_DELTA Then
        GridItensSRV.TextMatrix(iLinha, iGrid_QuantCancelSRV_Col) = Formata_Estoque(StrParaDbl(GridItensSRV.TextMatrix(iLinha, iGrid_QuantCancelSRV_Col)) * dFator)
    Else
        GridItensSRV.TextMatrix(iLinha, iGrid_QuantCancelSRV_Col) = ""
    End If

    'Preço unitário
    dPrecoUnitario = StrParaDbl(GridItensSRV.TextMatrix(iLinha, iGrid_ValorUnitarioSRV_Col))

    'Recalcula preço do ítem e valor total da nota
    If dPrecoUnitario > 0 Then
        GridItensSRV.TextMatrix(iLinha, iGrid_ValorUnitarioSRV_Col) = Format(dPrecoUnitario / dFator, gobjFAT.sFormatoPrecoUnitario)
        Call CF2(Me, "PSRV_PrecoTotal_CalculaSRV", iLinha)
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 191092
    End If

    Atualiza_UMSRV = SUCESSO

    Exit Function

Erro_Atualiza_UMSRV:

    Atualiza_UMSRV = gErr

    Select Case gErr

        Case 191088, 191089, 191091, 191092

        Case 191090
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191093)

    End Select

    Exit Function

End Function


'###########################################################################
'Inserido por Wagner
Function PSRV_Botao_Estoque(ByVal objCT As Object, ByVal objEstoqueProd As ClassEstoqueProduto, ByVal colSelecao As Collection, sSelecao As String) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim iAlmoxarifado As Integer

On Error GoTo Erro_PSRV_Botao_Estoque

    sSelecao = ""

    'VerIfica se tem alguma linha selecionada
    If objCT.GridItens.Row = 0 Then gError 193412

    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 193413

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        '*********** INCLUIDO PARA O TRATAMENTO DE ITEM DE GRADE ****************
        If gobjPedidoVenda.colItensPedido(GridItens.Row).iPossuiGrade = MARCADO Then
            Call BotaoReserva_Click
            Exit Function
        End If

        objEstoqueProd.sProduto = sProdutoFormatado

        colSelecao.Add sProdutoFormatado

        lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, sProdutoFormatado, iAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 23796 Then gError 193414

        If lErro = SUCESSO Then objEstoqueProd.iAlmoxarifado = iAlmoxarifado

    Else
        gError 193415


    End If

    PSRV_Botao_Estoque = SUCESSO

    Exit Function


Erro_PSRV_Botao_Estoque:

    PSRV_Botao_Estoque = gErr

    Select Case gErr

        Case 193412
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 193413, 193414

        Case 193415
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193416)

    End Select

    Exit Function

End Function

Public Sub BotaoTodosPedidos_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colSelecao As Collection

On Error GoTo Erro_NumeroLabel_Click

    lErro = Move_Pedido_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then gError 193417

    Call Chama_Tela("PedidoServico_Lista", colSelecao, objPedidoVenda, objEventoNumero)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 193417

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193418)

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

Public Function PSRV_AlocaAlmoxarifadoPadrao(ByVal objCT As Object, ByVal iFilialEmpresa As Integer, ByVal sProduto As String, iAlmoxarifadoRetorno As Integer) As Long

Dim lErro  As Long

On Error GoTo Erro_PSRV_AlocaAlmoxarifadoPadrao

    'Busca o Almoxarifado padrão
    lErro = CF("AlmoxarifadoPadrao_Le", iFilialEmpresa, sProduto, iAlmoxarifadoRetorno)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 193421

    'Se encontrou
    If iAlmoxarifadoRetorno = 0 Then gError 193422

    Exit Function

Erro_PSRV_AlocaAlmoxarifadoPadrao:

    PSRV_AlocaAlmoxarifadoPadrao = gErr

    Select Case gErr

        Case 193421, 193422

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193423)

    End Select

    Exit Function

End Function

'########################################################################
'Inserido por Wagner 23/03/2006
Public Function PSRV_Produto_Saida_Celula2(ByVal objCT As Object) As Long

Dim iIndice As Integer

On Error GoTo Erro_PSRV_Produto_Saida_Celula2

    'Verifica se já está em outra linha do Grid
    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
        If iIndice <> objCT.GridItens.Row Then
            If objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Produto_Col) = objCT.Produto.Text Then gError 193424
        End If
    Next

    PSRV_Produto_Saida_Celula2 = SUCESSO
    
    Exit Function

Erro_PSRV_Produto_Saida_Celula2:

    PSRV_Produto_Saida_Celula2 = gErr

    Select Case gErr

        Case 193424
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, objCT.Produto.Text, objCT.Produto.Text, iIndice)

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193425)

    End Select

    Exit Function

End Function

Public Function PSRV_Produto_Saida_Celula2SRV(ByVal objCT As Object) As Long

Dim iIndice As Integer

On Error GoTo Erro_PSRV_Produto_Saida_Celula2SRV

    'Verifica se já está em outra linha do Grid
    For iIndice = 1 To objCT.objGridItensSRV.iLinhasExistentes
        If iIndice <> objCT.GridItensSRV.Row Then
            If objCT.GridItensSRV.TextMatrix(iIndice, objCT.iGrid_ServicoSRV_Col) = objCT.ServicoSRV.Text Then gError 177164
        End If
    Next

    PSRV_Produto_Saida_Celula2SRV = SUCESSO
    
    Exit Function

Erro_PSRV_Produto_Saida_Celula2SRV:

    PSRV_Produto_Saida_Celula2SRV = gErr

    Select Case gErr

        Case 177164
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, objCT.ServicoSRV.Text, objCT.ServicoSRV.Text, iIndice)

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177165)

    End Select

    Exit Function

End Function

Public Function PSRV_Produto_Saida_Celula3(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim sCodigoAntigo As String

On Error GoTo Erro_PSRV_Produto_Saida_Celula3

    sCodigoAntigo = objProduto.sCodigo

    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
        If iIndice <> objCT.GridItens.Row Then
            If objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Produto_Col) = objCT.Produto.Text Then
                If vbYes = Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_EQUIVALENTE", objCT.Produto.Text, iIndice) Then

                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 193426

                    Call Chama_Tela_Modal("ProdutoEqui", objProduto)

                    If giRetornoTela <> OK Then gError 193427

                    lErro = CF("Produto_Grava_Equivalente", objProduto, sCodigoAntigo)
                    If lErro <> SUCESSO Then gError 193428

                Else
                    gError 193429
                End If
            End If
        End If
    Next

    PSRV_Produto_Saida_Celula3 = SUCESSO

    Exit Function

Erro_PSRV_Produto_Saida_Celula3:

    PSRV_Produto_Saida_Celula3 = gErr

    Select Case gErr

        Case 193426 To 193429

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193430)

    End Select

    Exit Function

End Function

Public Function PSRV_Produto_Saida_Celula3SRV(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim sCodigoAntigo As String

On Error GoTo Erro_PSRV_Produto_Saida_Celula3SRV

    sCodigoAntigo = objProduto.sCodigo

    For iIndice = 1 To objCT.objGridItensSRV.iLinhasExistentes
        If iIndice <> objCT.GridItensSRV.Row Then
            If objCT.GridItensSRV.TextMatrix(iIndice, objCT.iGrid_ServicoSRV_Col) = objCT.ServicoSRV.Text Then
                If vbYes = Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_EQUIVALENTE", objCT.ServicoSRV.Text, iIndice) Then

                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 193431

                    Call Chama_Tela_Modal("ProdutoEqui", objProduto)

                    If giRetornoTela <> OK Then gError 193432

                    lErro = CF("Produto_Grava_Equivalente", objProduto, sCodigoAntigo)
                    If lErro <> SUCESSO Then gError 193433

                Else
                    gError 193434
                End If
            End If
        End If
    Next

    PSRV_Produto_Saida_Celula3SRV = SUCESSO

    Exit Function

Erro_PSRV_Produto_Saida_Celula3SRV:

    PSRV_Produto_Saida_Celula3SRV = gErr

    Select Case gErr

        Case 193431 To 193434

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193435)

    End Select

    Exit Function

End Function

Public Function BotaoRefazAlocacao_Click() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dQuantidade As Double
Dim dQuantidadeCancelada As Double
Dim dQuantidadeFaturada As Double
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_BotaoRefazAlocacao_Click

    For iIndice = 1 To objGridItens.iLinhasExistentes

        If gobjPedidoVenda.colItensPedido(iIndice).colItensRomaneioGrade.Count = 0 Then

            dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col))
            dQuantidadeCancelada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))
            dQuantidadeFaturada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantFaturada_Col))

            lErro = Reserva_Processa(iIndice, dQuantidade, dQuantidadeCancelada, dQuantidadeFaturada)
            If lErro <> SUCESSO Then gError 193436

        Else

            For Each objItemRomaneio In gobjPedidoVenda.colItensPedido(iIndice).colItensRomaneioGrade

                lErro = Reserva_Processa_ItemGrade(objItemRomaneio, iIndice)
                If lErro <> SUCESSO Then gError 193437

            Next

            Call Atualiza_Grid_Alocacao(gobjPedidoVenda.colItensPedido.Item(iIndice))

        End If

    Next

    BotaoRefazAlocacao_Click = SUCESSO

    Exit Function

Erro_BotaoRefazAlocacao_Click:

    BotaoRefazAlocacao_Click = gErr

    Select Case gErr

        Case 193436, 193437

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193438)

    End Select

    Exit Function

End Function

'##############################################################

Public Sub PedidoRepr_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Function BotaoContato_Click() As Long

Dim objContato As New ClassContatos
Dim colSelecao As Collection
Dim lErro As Long

On Error GoTo Erro_BotaoContato_Click

    objContato.sNomeReduzido = Cliente.Text

    'Chama Tela ContatoLista
    Call Chama_Tela("ContatosLista", colSelecao, objContato, objEventoContato)

    BotaoContato_Click = SUCESSO

    Exit Function

Erro_BotaoContato_Click:

    BotaoContato_Click = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193439)

    End Select

    Exit Function

End Function

Private Sub objEventoContato_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objContato As ClassContatos
Dim objCliente As New ClassCliente

    Set objContato = obj1

    'Lê o Contato no BD
    lErro = CF("Contato_Transforma_Cliente", objContato, objCliente)
    If lErro <> SUCESSO Then gError 193440

    Cliente.Text = objCliente.sNomeReduzido

    Call Cliente_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_BotaoContato_Click:

    Select Case gErr

        Case 193440

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193441)

    End Select

    Exit Sub

End Sub

'Inserido por Wagner 03/08/2006
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

        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col)) - StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))
        objItensPRJCR.dValorOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))

'        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_PV

        colItensPRJ.Add objItensPRJCR

    Next
    
    Call gobjTribTab.Preenche_Itens_Projeto_NF(colItensPRJ)


'    Set objItensPRJCR = New ClassItensPRJCR
'
'    objItensPRJCR.sItem = "FR"
'    objItensPRJCR.sDescricao = "Frete"
'
'    objItensPRJCR.dValorOriginal = StrParaDbl(ValorFrete.Text)
'
''    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_PV
'
'    colItensPRJ.Add objItensPRJCR
'
'    Set objItensPRJCR = New ClassItensPRJCR
'
'    objItensPRJCR.sItem = "SG"
'    objItensPRJCR.sDescricao = "Seguro"
'
'    objItensPRJCR.dValorOriginal = StrParaDbl(ValorSeguro.Text)
'
''    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_PV
'
'    colItensPRJ.Add objItensPRJCR
'
'    Set objItensPRJCR = New ClassItensPRJCR
'
'    objItensPRJCR.sItem = "OD"
'    objItensPRJCR.sDescricao = "Outras Despesas"
'
'    objItensPRJCR.dValorOriginal = StrParaDbl(ValorDespesas.Text)
'
''    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_PV
'
'    colItensPRJ.Add objItensPRJCR
'
'    Set objItensPRJCR = New ClassItensPRJCR
'
'    objItensPRJCR.sItem = "ICMS"
'    objItensPRJCR.sDescricao = "ICMS Substituto"
'
'    objItensPRJCR.dValorOriginal = StrParaDbl(ICMSSubstValor1.Caption)
'
''    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_PV
'
'    colItensPRJ.Add objItensPRJCR
'
'    Set objItensPRJCR = New ClassItensPRJCR
'
'    objItensPRJCR.sItem = "IPI"
'    objItensPRJCR.sDescricao = "IPI"
'
'    objItensPRJCR.dValorOriginal = StrParaDbl(IPIValor1.Caption)
'
''    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_PV
'
'    colItensPRJ.Add objItensPRJCR
'
'    Set objItensPRJCR = New ClassItensPRJCR
'
'    objItensPRJCR.sItem = "ISS"
'    objItensPRJCR.sDescricao = "ISS"
'
'    objItensPRJCR.dValorOriginal = StrParaDbl(ISSValor.Text)
'
''    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_PV
'
'    colItensPRJ.Add objItensPRJCR
'
'    Set objItensPRJCR = New ClassItensPRJCR
'
'    objItensPRJCR.sItem = "DESC"
'    objItensPRJCR.sDescricao = "Desconto"
'
'    objItensPRJCR.dValorOriginal = StrParaDbl(ValorDesconto.Text)
'
'    objItensPRJCR.iNegativo = MARCADO
'
''    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_OV
'
'    colItensPRJ.Add objItensPRJCR

    Obter_ItensPRJCR = SUCESSO

    Exit Function

Erro_Obter_ItensPRJCR:

    Obter_ItensPRJCR = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193442)

    End Select

    Exit Function
    
End Function

Public Function Preenche_Itens_Projeto_Aux(ByVal colItensPRJ As Collection, ByVal sItem As String, ByVal sDescricao As String, ByVal dValor As Double, ByVal iTipo As Integer, Optional ByVal iNegativo As Integer = 0, Optional ByVal sObs As String = "") As Long

Dim objItensPRJCR As New ClassItensPRJCR

On Error GoTo Erro_Preenche_Itens_Projeto_Aux

    objItensPRJCR.sItem = sItem
    objItensPRJCR.sDescricao = sDescricao
    objItensPRJCR.dValorOriginal = dValor
    objItensPRJCR.iNegativo = iNegativo
    objItensPRJCR.iTipoOrigem = iTipo
    objItensPRJCR.sObservacao = sObs

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
'#####################################


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

Dim objParcPV As ClassParcelaPedidoVenda
Dim objForm As Form

On Error GoTo Erro_TipoPagto_Detalhe

    If TipoPagto.ListIndex <> -1 Then

        If TipoPagto.ItemData(TipoPagto.ListIndex) = TIPOPAGTO_DEPOSITO Then

            If gcolParcPV.Count >= GridParcelas.Row Then

                Set objParcPV = gcolParcPV.Item(GridParcelas.Row)

                Call Chama_Tela_Modal("DepositoConta", objParcPV)

            End If

        ElseIf TipoPagto.ItemData(TipoPagto.ListIndex) = TIPOPAGTO_CHEQUE Then

            If gcolParcPV.Count >= GridParcelas.Row Then

                Set objParcPV = gcolParcPV.Item(GridParcelas.Row)

                Call Chama_Tela_Modal("ChequePagto", objParcPV)

            End If

        ElseIf TipoPagto.ItemData(TipoPagto.ListIndex) = TIPOPAGTO_CARTAO Then

            If gcolParcPV.Count >= GridParcelas.Row Then

                Set objParcPV = gcolParcPV.Item(GridParcelas.Row)

                Call Chama_Tela_Modal("CartaoPagto", objParcPV)

            End If

        End If

    End If

    Exit Sub

Erro_TipoPagto_Detalhe:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193443)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_TipoPagto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo Pagto que está deixando de serr a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim iTipo As Integer
Dim iIndice As Integer
Dim objParcPV As ClassParcelaPedidoVenda

On Error GoTo Erro_Saida_Celula_TipoPagto

    Set objGridInt.objControle = TipoPagto

    'Verifica se o Tipo foi preenchido
    If Len(Trim(TipoPagto.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoPagto.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_TipoPagto_Col) Then

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona_Grid(objGridInt.objControle, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 193444

            'Não foi encontrado
            If lErro = 25085 Then gError 193445
            If lErro = 25086 Then gError 193446

            Call TipoPagto_Detalhe

            CobrancaAutomatica.Value = vbUnchecked

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            Set objParcPV = New ClassParcelaPedidoVenda
            gcolParcPV.Add objParcPV

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 193447

    Saida_Celula_TipoPagto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoPagto:

    Saida_Celula_TipoPagto = gErr

    Select Case gErr

        Case 193444, 193447
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 193445
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMAPAGTO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 193446
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMAPAGTO_NAO_ENCONTRADO", gErr, objGridInt.objControle.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193448)

    End Select

    Exit Function

End Function

Public Sub BotaoTipoPagto_Click()

Dim lErro As Long
Dim lCodigo As Long
Dim iTipoPagto As Integer
Dim objParcPV As ClassParcelaPedidoVenda

On Error GoTo Erro_BotaoTipoPagto_Click

    If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_TipoPagto_Col))) > 0 Then

        iTipoPagto = Codigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_TipoPagto_Col))

        If iTipoPagto = TIPOPAGTO_DEPOSITO Then

            If gcolParcPV.Count >= GridParcelas.Row Then

                Set objParcPV = gcolParcPV.Item(GridParcelas.Row)

                Call Chama_Tela_Modal("DepositoConta", objParcPV)

            End If

        ElseIf iTipoPagto = TIPOPAGTO_CHEQUE Then

            If gcolParcPV.Count >= GridParcelas.Row Then

                Set objParcPV = gcolParcPV.Item(GridParcelas.Row)

                Call Chama_Tela_Modal("ChequePagto", objParcPV)

            End If

        ElseIf iTipoPagto = TIPOPAGTO_CARTAO Then

            If gcolParcPV.Count >= GridParcelas.Row Then

                Set objParcPV = gcolParcPV.Item(GridParcelas.Row)

                Call Chama_Tela_Modal("CartaoPagto", objParcPV)

            End If

        End If

    End If

    Exit Sub

Erro_BotaoTipoPagto_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193449)

    End Select

    Exit Sub

End Sub

Function PSRV_Inicializa_Grid_Parcelas(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
    
    objGridInt.objGrid.ColWidth(iGrid_TipoPagto_Col) = 0
    objGridInt.objGrid.TextMatrix(0, iGrid_TipoPagto_Col) = ""
    
    PSRV_Inicializa_Grid_Parcelas = SUCESSO
    
End Function

Public Sub PedidoRepr_Validate(Cancel As Boolean)

Dim lErro As Long

    lErro = CF("PSRV_PedidoRepr_Valida", StrParaDbl(Codigo.Text), giFilialEmpresa, PedidoRepr.Text)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TabStripSRV_Click()

Dim lErro As Long

On Error GoTo Erro_TabStripSRV_Click

    'Se frame selecionado não for o atual
    If TabStripSRV.SelectedItem.Index <> giFrameAtualItens Then

        If TabStrip_PodeTrocarTab(giFrameAtualItens, TabStripSRV, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        FrameItens(TabStripSRV.SelectedItem.Index).Visible = True
        FrameItens(giFrameAtualItens).Visible = False

        'Armazena novo valor de giFrameAtual
        giFrameAtualItens = TabStripSRV.SelectedItem.Index

    End If

    Exit Sub

Erro_TabStripSRV_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193463)

    End Select

    Exit Sub

End Sub

Public Sub LabelSolicitaSRV_Click()

Dim objSolicSRV As New ClassSolicSRV
Dim colSelecao As New Collection

    objSolicSRV.lCodigo = StrParaLong(CodSolicSRV.Text)
    
    Call Chama_Tela("SolicitacaoSRVLista", colSelecao, objSolicSRV, objEventoCodSolicSRV)
    
End Sub

Private Sub objEventoCodSolicSRV_evSelecao(obj1 As Object)

Dim objSolicSRV As ClassSolicSRV
Dim bCancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoCodSolicSRV_evSelecao

    Set objSolicSRV = obj1
    
    CodSolicSRV.Text = objSolicSRV.lCodigo
    Call BotaoSolicitacaoSRV_Click
    
    Me.Show

    Exit Sub

Erro_objEventoCodSolicSRV_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193464)
    
    End Select

End Sub

Private Sub CodSolicSRV_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodSolicSRV_GotFocus()
    Call MaskEdBox_TrataGotFocus(CodSolicSRV, iAlterado)
End Sub

Private Sub CodSolicSRV_Validate(Cancel As Boolean)

Dim objSolicSRV As New ClassSolicSRV
Dim lErro As Long

On Error GoTo Erro_CodSolicSRV_Validate

    If Len(Trim(CodSolicSRV.Text)) > 0 Then

        objSolicSRV.lCodigo = StrParaLong(CodSolicSRV.Text)
        objSolicSRV.iFilialEmpresa = giFilialEmpresa

        'Lê no BD os dados da solicitacao em questao
        lErro = CF("SolicitacaoSRV_Le", objSolicSRV)
        If lErro <> SUCESSO And lErro <> 183258 Then gError 193468
    
        'Se não encontrou a solicitacao => erro
        If lErro <> SUCESSO Then gError 193469

    End If
    
    Exit Sub
    
Erro_CodSolicSRV_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 193468
    
        Case 193469
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRV_NAO_ENCONTRADO", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193470)
    
    End Select
    
    Exit Sub

End Sub


'#################################################
'Bloqueio

Public Sub TipoBloqueio_Change()
    Call gobjTelaTrataBloqGen.TipoBloqueio_Change
End Sub

Public Sub TipoBloqueio_GotFocus()
    Call gobjTelaTrataBloqGen.TipoBloqueio_GotFocus
End Sub

Public Sub TipoBloqueio_KeyPress(KeyAscii As Integer)
    Call gobjTelaTrataBloqGen.TipoBloqueio_KeyPress(KeyAscii)
End Sub

Public Sub TipoBloqueio_Validate(Cancel As Boolean)
    Call gobjTelaTrataBloqGen.TipoBloqueio_Validate(Cancel)
End Sub

Public Sub DataBloqueio_Change()
    Call gobjTelaTrataBloqGen.DataBloqueio_Change
End Sub

Public Sub DataBloqueio_GotFocus()
    Call gobjTelaTrataBloqGen.DataBloqueio_GotFocus
End Sub

Public Sub DataBloqueio_KeyPress(KeyAscii As Integer)
    Call gobjTelaTrataBloqGen.DataBloqueio_KeyPress(KeyAscii)
End Sub

Public Sub DataBloqueio_Validate(Cancel As Boolean)
    Call gobjTelaTrataBloqGen.DataBloqueio_Validate(Cancel)
End Sub

Public Sub Observacao_Change()
    Call gobjTelaTrataBloqGen.Observacao_Change
End Sub

Public Sub Observacao_GotFocus()
    Call gobjTelaTrataBloqGen.Observacao_GotFocus
End Sub

Public Sub Observacao_KeyPress(KeyAscii As Integer)
    Call gobjTelaTrataBloqGen.Observacao_KeyPress(KeyAscii)
End Sub

Public Sub Observacao_Validate(Cancel As Boolean)
    Call gobjTelaTrataBloqGen.Observacao_Validate(Cancel)
End Sub

Public Sub CodUsuario_Change()
    Call gobjTelaTrataBloqGen.Usuario_Change
End Sub

Public Sub CodUsuario_GotFocus()
    Call gobjTelaTrataBloqGen.Usuario_GotFocus
End Sub

Public Sub CodUsuario_KeyPress(KeyAscii As Integer)
    Call gobjTelaTrataBloqGen.Usuario_KeyPress(KeyAscii)
End Sub

Public Sub CodUsuario_Validate(Cancel As Boolean)
    Call gobjTelaTrataBloqGen.Usuario_Validate(Cancel)
End Sub

Public Sub ResponsavelBL_Change()
    Call gobjTelaTrataBloqGen.Responsavel_Change
End Sub

Public Sub ResponsavelBL_GotFocus()
    Call gobjTelaTrataBloqGen.Responsavel_GotFocus
End Sub

Public Sub ResponsavelBL_KeyPress(KeyAscii As Integer)
    Call gobjTelaTrataBloqGen.Responsavel_KeyPress(KeyAscii)
End Sub

Public Sub ResponsavelBL_Validate(Cancel As Boolean)
    Call gobjTelaTrataBloqGen.Responsavel_Validate(Cancel)
End Sub

Public Sub GridBloqueio_Click()
    Call gobjTelaTrataBloqGen.GridBloqueio_Click
End Sub

Public Sub GridBloqueio_GotFocus()
    Call gobjTelaTrataBloqGen.GridBloqueio_GotFocus
End Sub

Public Sub GridBloqueio_EnterCell()
    Call gobjTelaTrataBloqGen.GridBloqueio_EnterCell
End Sub

Public Sub GridBloqueio_LeaveCell()
    Call gobjTelaTrataBloqGen.GridBloqueio_LeaveCell
End Sub

Public Sub GridBloqueio_KeyDown(KeyCode As Integer, Shift As Integer)
    Call gobjTelaTrataBloqGen.GridBloqueio_KeyDown(KeyCode, Shift)
End Sub

Public Sub GridBloqueio_KeyPress(KeyAscii As Integer)
    Call gobjTelaTrataBloqGen.GridBloqueio_KeyPress(KeyAscii)
End Sub

Public Sub GridBloqueio_Validate(Cancel As Boolean)
    Call gobjTelaTrataBloqGen.GridBloqueio_Validate(Cancel)
End Sub

Public Sub GridBloqueio_RowColChange()
    Call gobjTelaTrataBloqGen.GridBloqueio_RowColChange
End Sub

Public Sub GridBloqueio_Scroll()
    Call gobjTelaTrataBloqGen.GridBloqueio_Scroll
End Sub

Public Sub BotaoLiberaBloqueio_Click()
    Call gobjTelaTrataBloqGen.BotaoLibBloq_Click(Codigo)
End Sub
'Bloqueio
'#################################################

Public Sub BotaoSolicitacaoSRV_Click()
    Call Trazer_SolicitacaoSRV_Click
End Sub

Public Sub BotaoSolicitacaoSRVAP_Click()
    Call Trazer_SolicitacaoSRV_Click(True)
End Sub

Private Sub Trazer_SolicitacaoSRV_Click(Optional ByVal bAP As Boolean = False)

Dim lErro As Long
Dim objSolicSRV As New ClassSolicSRV

On Error GoTo Erro_Trazer_SolicitacaoSRV_Click

    'Verfifica se o Pedido e a FilialPedido foram informadas
    If Len(Trim(CodSolicSRV.Text)) = 0 Then gError 186998

    objSolicSRV.iFilialEmpresa = giFilialEmpresa
    objSolicSRV.lCodigo = StrParaLong(CodSolicSRV.Text)

    lErro = CF("SolicitacaoSRV_Le", objSolicSRV)
    If lErro <> SUCESSO And lErro <> 183258 Then gError 186999
    
    If lErro <> SUCESSO Then gError 188000

    'Traz os dados do pedido para a tela
    lErro = Traz_Dados_SolicSRV_Tela(objSolicSRV, bAP)
    If lErro <> SUCESSO Then gError 188001
    
    Exit Sub

Erro_Trazer_SolicitacaoSRV_Click:

    Select Case gErr

        Case 186998
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICITACAO_NAO_PREENCHIDA", gErr)

        Case 186999, 188001

        Case 188000
            Call Rotina_Erro(vbOKOnly, "ERRO_SOLICSRV_NAO_ENCONTRADO", gErr, objSolicSRV.iFilialEmpresa, objSolicSRV.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188002)

    End Select

    Exit Sub

End Sub

Public Function Traz_Dados_SolicSRV_Tela(objSolicSRV As ClassSolicSRV, ByVal bAP As Boolean) As Long

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim objItensSolicSRV As ClassItensSolicSRV
Dim sServicoEnxuto As String
Dim iIndice As Integer
Dim objProdSolicSRV As ClassProdSolicSRV
Dim colItens As New Collection
Dim dPrecoUnitario As Double

On Error GoTo Erro_Traz_Dados_SolicSRV_Tela

    'Limpa a tela
        Call Limpa_Tela_Apontamento
    

    
        If objSolicSRV.iVendedor <> 0 Then
            
            objVendedor.iCodigo = objSolicSRV.iVendedor
                
            lErro = CF("Vendedor_Le", objVendedor)
            If lErro <> SUCESSO And lErro <> 12582 Then gError 188003
            
            If lErro <> SUCESSO Then gError 188004
        
            Vendedor.Text = objVendedor.sNomeReduzido
            
        End If
            
            
        CodSolicSRV.Text = objSolicSRV.lCodigo
    
    
        'Se existe um código para o Cliente
        If objSolicSRV.lCliente <> 0 Then
    
            giClienteAlterado = 1
            giFilialAlterada = 1
            giFilialEntregaAlterada = 1
            

    
            Call Cliente_Formata(objSolicSRV.lCliente)
            Call Filial_Formata(Filial, objSolicSRV.iFilial)
            Call CF("Filial_Seleciona", FilialEntrega, objSolicSRV.iFilial)
            
            Call Cliente_Validate(bSGECancelDummy)
            Call Filial_Validate(bSGECancelDummy)
            Call FilialEntrega_Validate(bSGECancelDummy)
    
        End If
    
    
    lErro = CF("Solicitacao_Le_Gastos", objSolicSRV, colItens, bAP)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iIndice = 0
    For Each objItensSolicSRV In colItens
    
        If objItensSolicSRV.iTipo = ITEMNF_TIPO_SERVICO Then
    
            lErro = Mascara_RetornaProdutoEnxuto(objItensSolicSRV.sProduto, sServicoEnxuto)
            If lErro <> SUCESSO Then gError 188005
    
            'Mascara o produto enxuto
            ServicoSRV.PromptInclude = False
            ServicoSRV.Text = sServicoEnxuto
            ServicoSRV.PromptInclude = True
    
            iIndice = iIndice + 1
    
            'Coloca os dados dos itens na tela
            GridItensSRV.TextMatrix(iIndice, iGrid_ServicoSRV_Col) = ServicoSRV.Text
            GridItensSRV.TextMatrix(iIndice, iGrid_DescServicoSRV_Col) = objItensSolicSRV.sProdutoDesc
            GridItensSRV.TextMatrix(iIndice, iGrid_UnidadeMedSRV_Col) = objItensSolicSRV.sUM
            GridItensSRV.TextMatrix(iIndice, iGrid_QuantPedidaSRV_Col) = Formata_Estoque(objItensSolicSRV.dQuantidade)
    
            objGridItensSRV.iLinhasExistentes = iIndice
    
    
'            Set objProdSolicSRV = New ClassProdSolicSRV
'
'            objProdSolicSRV.sServicoOrcSRV = objItensSolicSRV.sServico
'            objProdSolicSRV.dQuantidade = objItensSolicSRV.dQuantidade
'            objProdSolicSRV.sProduto = objItensSolicSRV.sProduto
'            objProdSolicSRV.sContrato = objItensSolicSRV.sContrato
'            objProdSolicSRV.lGarantia = objItensSolicSRV.lGarantia
'            objProdSolicSRV.sLote = objItensSolicSRV.sLote
'            objProdSolicSRV.iFilialOP = objItensSolicSRV.iFilialOP
'
'            gcolProdSolicSRV.Add objProdSolicSRV

            gcolItemPedidoSRV.Add 0
    
            Call gobjTelaProjetoInfo.Adiciona_Linha(iIndice)
        
            'permite que a tributacao reflita a inclusao de uma linha no grid
            lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, objItensSolicSRV.sProduto, True)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            
            
            'Preço Unitário
            If Len(Trim(TabelaPreco.Text)) > 0 Then
            
                lErro = Trata_TabelaPreco_ItemSRV(iIndice)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            End If
            
            
        End If
    
    
    
    Next
    
    
    iIndice = 0
    For Each objItensSolicSRV In colItens
    
        If objItensSolicSRV.iTipo = ITEMNF_TIPO_PECA Then
    
            lErro = Mascara_RetornaProdutoEnxuto(objItensSolicSRV.sProduto, sServicoEnxuto)
            If lErro <> SUCESSO Then gError 188005
    
            'Mascara o produto enxuto
            Produto.PromptInclude = False
            Produto.Text = sServicoEnxuto
            Produto.PromptInclude = True
    
            iIndice = iIndice + 1
    
            'Coloca os dados dos itens na tela
            GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
            GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItensSolicSRV.sProdutoDesc
            GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItensSolicSRV.sUM
            GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col) = Formata_Estoque(objItensSolicSRV.dQuantidade)

            objGridItens.iLinhasExistentes = iIndice


'            Set objProdSolicSRV = New ClassProdSolicSRV
'
'            objProdSolicSRV.sServicoOrcSRV = objItensSolicSRV.sServico
'            objProdSolicSRV.dQuantidade = objItensSolicSRV.dQuantidade
'            objProdSolicSRV.sProduto = objItensSolicSRV.sProduto
'            objProdSolicSRV.sContrato = objItensSolicSRV.sContrato
'            objProdSolicSRV.lGarantia = objItensSolicSRV.lGarantia
'            objProdSolicSRV.sLote = objItensSolicSRV.sLote
'            objProdSolicSRV.iFilialOP = objItensSolicSRV.iFilialOP
'
'            gcolProdSolicSRV.Add objProdSolicSRV

            gcolItemPedido.Add 0
    
            Call gobjTelaProjetoInfo.Adiciona_Linha(iIndice)

            'permite que a tributacao reflita a inclusao de uma linha no grid
            lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, objItensSolicSRV.sProduto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                'Preço Unitário
            If Len(Trim(TabelaPreco.Text)) > 0 Then
            
                lErro = Trata_TabelaPreco_Item(iIndice)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            End If
            
            
        End If
    
    Next
        
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        
    Traz_Dados_SolicSRV_Tela = SUCESSO

    Exit Function

Erro_Traz_Dados_SolicSRV_Tela:

    'Indica que o sistema não está mais carregando dados para a tela
    gbCarregandoTela = False
    
    Traz_Dados_SolicSRV_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 188003, 195756

        Case 188004
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, objVendedor.iCodigo)

        Case 188005
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItensSolicSRV.sServico)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188006)

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
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))) > 0 Then
                dValorItem = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
                dDescontoItem = Arredonda_Moeda(dValorItem * dFator)
                dDescontoAplicado = dDescontoAplicado + dDescontoItem
                GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(dDescontoItem, "STANDARD")
            End If
        Next
        For iIndice = 1 To objGridItensSRV.iLinhasExistentes
            If Len(Trim(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRVB_Col))) > 0 Then
                dValorItem = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRVB_Col))
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
            dPercDesc = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col)) / StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
            GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dPercDesc, "Percent")
            Call CF2(Me, "PSRV_PrecoTotal_Calcula", iIndice)
            
            lErro = gobjTribTab.Alteracao_Item_Grid(iIndice)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            dValorTotal = dValorTotal + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        Next
    
        For iIndice = 1 To objGridItensSRV.iLinhasExistentes
            dPercDesc = StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_DescontoSRV_Col)) / StrParaDbl(GridItensSRV.TextMatrix(iIndice, iGrid_PrecoTotalSRVB_Col))
            GridItensSRV.TextMatrix(iIndice, iGrid_PercDescSRV_Col) = Format(dPercDesc, "Percent")
            Call CF2(Me, "PSRV_PrecoTotal_CalculaSRV", iIndice)
            
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

Public Sub BotaoDesfazer_Click()

Dim objPedidoVenda As New ClassPedidoDeVenda
Dim lErro As Long

On Error GoTo Erro_BotaoDesfazer_Click

    If Len(Trim(Codigo.Text)) = 0 Then gError 210647

    objPedidoVenda.lCodigo = StrParaLong(Codigo.Text)
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa

    lErro = CF("PedidoSRV_Desfaz_Baixa", objPedidoVenda)
    If lErro <> SUCESSO Then gError 210648

    Call Limpa_Tela_PedidoVenda
    
    Exit Sub

Erro_BotaoDesfazer_Click:

    Select Case gErr
    
        Case 210647
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)
        
        Case 210648
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210649)
            
    End Select
    
    Exit Sub

End Sub

Private Sub Class_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
End Sub

Private Sub Limpa_Tela_Apontamento()

Dim iIndice As Integer

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me

    '************** TRATAMENTO DE GRADE **************
    For iIndice = 1 To objGridItens.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next
    '*************************************************

    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridItensSRV)
    Call Grid_Limpa(objGridComissoes)
    Call Grid_Limpa(objGridParcelas)

    Set gcolParcPV = New Collection

    Set gcolItemPedido = New Collection
    Set gcolItemPedidoSRV = New Collection
   
       Call gobjTribTab.Limpa_Tela

End Sub
