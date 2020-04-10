VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPedidoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjInfoUsu As Object

Public gobjAnotacao As ClassAnotacoes
Public gobjInfoAdic As ClassInfoAdic
Dim dValorDescontoItensAnt As Double
Dim dPercDescontoItensAnt As Double
Dim iDescontoAlterado As Integer

Public gobjTelaProjetoInfo As ClassTelaPRJInfo
Public gdQuantAnterior As Double

Public glVolumeQuantCalc As Long 'qtde de volumes calculadas para a Nena

Public iImportacao As Integer

Const CLIENTE_INEXISTENTE = 1 'Luiz Nogueira em 26/01/04
Const FILIALCLIENTE_INEXISTENTE = 1 'Luiz Nogueira em 26/01/04
Const FILIALENTREGA_INEXISTENTE = 2 'Luiz Nogueira em 26/01/04

'Flag que indica se a tela está sendo preenchida.
Public gbCarregandoTela As Boolean
Public gbLimpandoTela As Boolean

Dim m_objUserControl As Object

Dim bPulaValorTotalCalcula As Boolean
Dim iTratamentoPadraoFaltaEstoque As Integer

Dim gbSaiuGrade As Boolean

'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjPedidoVenda As ClassPedidoDeVenda 'estrutura mantida para auxiliar a manutencao de informacoes p/tributacao
'todos os dados relevantes p/tributacao dentro de gobjPedidoVenda estarao sincronizados com a tela antes da atualizacao da tributacao

'inicia objeto associado a GridComissoes
Public objTabComissoes As New ClassTabComissoes

'declara o objeto da execucao das regras de comissoes e
'o objeto do calculo de comissoes
Public objComissoesRegrasCalc As Object
Public objMnemonicoComissCalc As ClassMnemonicoComissCalc
Public objMnemonicoComissCalcAux As ClassMnemonicoComissCalcAux

Public giFrameAtual As Integer
Public iAlterado As Integer

Public gcolItemPedido As Collection
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
Dim giDataRefFluxoAlterada As Integer
Public giNaturezaOpAlterada As Integer
Dim giDataEntregaAlterada As Integer

Public giValorDescontoManual As Integer
Dim giFilialEntregaManual As Integer

Public gobjTribTab As New ClassTribTab

Public giPercAcresFinAlterado As Integer

'Embalagem
Public gobjEmbalagem As Object

'variaveis relacionadas ao grid
Dim giGridBloqueioLinhasBloqueadas As Integer

'Grid Itens
Public iGrid_Item_Col As Integer
Public iGrid_ProdutoAlmox_Col As Integer
Public iGrid_UMEstoque_Col As Integer
Public iGrid_Almoxarifado_Col As Integer
Public iGrid_QuantReservar_Col As Integer
Public iGrid_QuantReserv_Col As Integer
Public iGrid_Validade_Col As Integer
Public iGrid_Responsavel_Col As Integer
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
Public iGrid_DataEntrega_Col As Integer
Public iGrid_QuantReservada_Col As Integer
Public iGrid_QuantFaturada_Col As Integer
Public iGrid_QuantFaturadaAMais_Col As Integer
Public iGrid_PercentMenosReceb_Col As Integer
Public iGrid_PercentMaisReceb_Col As Integer
Public iGrid_RecebForaFaixa_Col As Integer
Public iGrid_QuantFatAMais_Col As Integer
Public iGrid_Prioridade_Col As Integer
Public iGrid_TabPrecoItem_Col As Integer
Public iGrid_ComissaoItem_Col As Integer

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

'Grid Bloqueio
Public iGrid_TipoBloqueio_Col As Integer
Public iGrid_DataBloqueio_Col As Integer
Public iGrid_CodUsuario_Col As Integer
Public iGrid_ResponsavelBL_Col As Integer
Public iGrid_DataLiberacao_Col As Integer
Public iGrid_ResponsavelLib_Col As Integer
Public iGrid_Observacao_Col As Integer
Public iGrid_SeqBloqueio_Col As Integer

Public objGridComissoes As AdmGrid
'Janaina
'Dim objGridItens As AdmGrid
Public objGridItens As AdmGrid
'Janaina
Public objGridParcelas As AdmGrid
Public objGridBloqueio As AdmGrid
Public objGridAlocacao As AdmGrid

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
Private WithEvents objEventoEstoqueProd As AdmEvento
Attribute objEventoEstoqueProd.VB_VarHelpID = -1
Public WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoNaturezaOp As AdmEvento
Attribute objEventoNaturezaOp.VB_VarHelpID = -1
Private WithEvents objEventoMensagem As AdmEvento
Attribute objEventoMensagem.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoTranspRedespacho As AdmEvento
Attribute objEventoTranspRedespacho.VB_VarHelpID = -1
Private WithEvents objEventoContato As AdmEvento
Attribute objEventoContato.VB_VarHelpID = -1

Dim giLinhaAnterior As Integer
Private giPosCargaOk As Integer

'Constantes públicas dos tabs
Private Const TAB_Principal = 1
Private Const TAB_Itens = 2
Private Const TAB_Complemento = 3
Private Const TAB_Cobranca = 4
Private Const TAB_Comissoes = 5
Private Const TAB_Bloqueio = 6
Private Const TAB_Almoxarifado = 7
Private Const TAB_Tributacao = 8
Private Const TAB_Embalagens = 9

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
Private Const CTACONTABILEST1 As String = "ContaContabilEst"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque" '??? ERRADO: nao está no bd
'fim da contabilidade

'workflow
Private Const CREDITO_ULTRAPASSADO As String = "CreditoUltrapassado"
Private Const CODIGOPEDIDO As String = "CodigoPedido"
Private Const PRODUTO_SEM_RESERVA As String = "ProdutoSemReserva"
Private Const CLIENTE_NOME As String = "Cliente_Nome"
Private Const CPF_CNPJ As String = "CPF_CNPJ"
Private Const UF_PRINCIPAL As String = "UF_Principal"
Private Const CONDPAGTO As String = "CondPagto"
Private Const EXISTE_PRECOMENORTABELA As String = "Existe_PrecoMenorTab"
Private Const PRECOMENORTABELA As String = "PrecoMenorTabela"
Private Const EMAIL_VEND_CLI As String = "Email_Vend_Cli"
Private Const TIPO_FRETE_CLIENTE As String = "Tipo_Frete_Cliente"
Private Const TIPO_FRETE_PEDIDO As String = "Tipo_Frete_Pedido"
Private Const TIPO_FRETE_CLIENTE_DESC As String = "Tipo_Frete_Cli_Desc"
Private Const TIPO_FRETE_PEDIDO_DESC As String = "Tipo_Frete_PV_Desc"

'Incluido por Leo em 24/04/02. Para o  funcionamento do tab de comissoes
Public iComissoesAlterada As Integer

Private gcolParcPV As Collection

'--- inicio dos properties get dos controles da tela
Public Property Get Label1() As Object
     Set Label1 = objUserControl.Controls("Label1")
End Property

Public Property Get BotaoAnaliseVenda() As Object
     Set BotaoAnaliseVenda = objUserControl.Controls("BotaoAnaliseVenda")
End Property

Public Property Get FrameOrcVenda() As Object
     Set FrameOrcVenda = objUserControl.Controls("FrameOrcVenda")
End Property

Public Property Get TabPrecoItemPV() As Object
     Set TabPrecoItemPV = objUserControl.Controls("TabPrecoItemPV")
End Property

Public Property Get ComissaoItemPV() As Object
     Set ComissaoItemPV = objUserControl.Controls("ComissaoItemPV")
End Property

Public Property Get FrameCodBase() As Object
     Set FrameCodBase = objUserControl.Controls("FrameCodBase")
End Property

Public Property Get CodBase() As Object
     Set CodBase = objUserControl.Controls("CodBase")
End Property

Public Property Get Parc() As Object
     Set Parc = objUserControl.Controls("Parc")
End Property

Public Property Get FrameParc() As Object
     Set FrameParc = objUserControl.Controls("FrameParc")
End Property

Public Property Get ImprimirConfGravacao() As Object
     Set ImprimirConfGravacao = objUserControl.Controls("ImprimirConfGravacao")
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

Public Property Get EmailGravacao() As Object
     Set EmailGravacao = objUserControl.Controls("EmailGravacao")
End Property

Public Property Get ImprimeGravacao() As Object
     Set ImprimeGravacao = objUserControl.Controls("ImprimeGravacao")
End Property

Public Property Get BotaoEmail() As Object
     Set BotaoEmail = objUserControl.Controls("BotaoEmail")
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

Public Property Get BotaoGrade() As Object
     Set BotaoGrade = objUserControl.Controls("BotaoGrade")
End Property

Public Property Get BotaoInfoAdicDocItem() As Object
     Set BotaoInfoAdicDocItem = objUserControl.Controls("BotaoInfoAdicDocItem")
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

Public Property Get TransportadoraLabel() As Object
     Set TransportadoraLabel = objUserControl.Controls("TransportadoraLabel")
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

Public Property Get ProdutoDescricao() As Object
     Set ProdutoDescricao = objUserControl.Controls("ProdutoDescricao")
End Property

Public Property Get BotaoLiberaBloqueio() As Object
     Set BotaoLiberaBloqueio = objUserControl.Controls("BotaoLiberaBloqueio")
End Property

Public Property Get TipoBloqueio() As Object
     Set TipoBloqueio = objUserControl.Controls("TipoBloqueio")
End Property

Public Property Get ResponsavelLib() As Object
     Set ResponsavelLib = objUserControl.Controls("ResponsavelLib")
End Property

Public Property Get DataLiberacao() As Object
     Set DataLiberacao = objUserControl.Controls("DataLiberacao")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get SeqBloqueio() As Object
     Set SeqBloqueio = objUserControl.Controls("SeqBloqueio")
End Property

Public Property Get CodUsuario() As Object
     Set CodUsuario = objUserControl.Controls("CodUsuario")
End Property

Public Property Get ResponsavelBL() As Object
     Set ResponsavelBL = objUserControl.Controls("ResponsavelBL")
End Property

Public Property Get DataBloqueio() As Object
     Set DataBloqueio = objUserControl.Controls("DataBloqueio")
End Property

Public Property Get GridBloqueio() As Object
     Set GridBloqueio = objUserControl.Controls("GridBloqueio")
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

Public Property Get BotaoEstoqueProd() As Object
     Set BotaoEstoqueProd = objUserControl.Controls("BotaoEstoqueProd")
End Property

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
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

Public Property Get CalculoAuto() As Object
     Set CalculoAuto = objUserControl.Controls("CalculoAuto")
End Property

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

Public Property Get BotaoDataRefFluxoDown() As Object
     Set BotaoDataRefFluxoDown = objUserControl.Controls("BotaoDataRefFluxoDown")
End Property

Public Property Get BotaoDataRefFluxoUp() As Object
     Set BotaoDataRefFluxoUp = objUserControl.Controls("BotaoDataRefFluxoUp")
End Property

Public Property Get DataRefFluxo() As Object
     Set DataRefFluxo = objUserControl.Controls("DataRefFluxo")
End Property

Public Property Get RecebForaFaixa() As Object
     Set RecebForaFaixa = objUserControl.Controls("RecebForaFaixa")
End Property

Public Property Get PercentMenosReceb() As Object
     Set PercentMenosReceb = objUserControl.Controls("PercentMenosReceb")
End Property

Public Property Get PercentMaisReceb() As Object
     Set PercentMaisReceb = objUserControl.Controls("PercentMaisReceb")
End Property

Public Property Get QuantFatAMais() As Object
     Set QuantFatAMais = objUserControl.Controls("QuantFatAMais")
End Property

Public Property Get Prioridade() As Object
     Set Prioridade = objUserControl.Controls("Prioridade")
End Property

Public Property Get PrioridadePadrao() As Object
     Set PrioridadePadrao = objUserControl.Controls("PrioridadePadrao")
End Property

Public Property Get BotaoExportar() As Object
     Set BotaoExportar = objUserControl.Controls("BotaoExportar")
End Property

Public Property Get TipoFrete() As Object
     Set TipoFrete = objUserControl.Controls("TipoFrete")
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
    If lErro <> SUCESSO Then gError 26715

    Call DataReferencia_Validate(bCancel)
    
    If bCancel = True Then DataReferencia.Text = sData
    
    Exit Sub

Erro_BotaoDataReferenciaDown_Click:

    Select Case gErr

        Case 26715
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157906)

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
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157907)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long
Dim iIndice As Integer

On Error GoTo Erro_BotaoProxNum_Click

    If Codigo.Enabled = True Then
    
        lErro = CF("PedidoDeVenda_Automatico", lCodigo)
        If lErro <> SUCESSO Then gError 57537

        Codigo.Text = CStr(lCodigo)
        
        'limpa a quantidade faturada do grid
        For iIndice = 1 To objGridItens.iLinhasExistentes
    
            GridItens.TextMatrix(iIndice, iGrid_QuantFaturada_Col) = ""
            
        Next
        
        Call Grid_Limpa(objGridBloqueio)
        
    Else
    
        'Codigo desabilitado não muda
        gError 25662
    
    End If
        

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 57537
        
        Case 25662
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_DESABILITADO_IMUTAVEL", gErr, Error)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157908)
    
    End Select

    Exit Sub

End Sub

Public Sub Codigo_GotFocus()

    'gsCodigoAnterior = Codigo.Text 'Trecho comentado por Leo em 18/02/02
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

Public Sub TipoFrete_Click()

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
        If FilialFaturamento_TrataAlteracao() <> SUCESSO Then gError 25442
    End If

    Exit Sub

Erro_FilialFaturamento_Click:

    Select Case gErr

        Case 25442

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157909)

    End Select

    Exit Sub

End Sub

'@@@@@ para customizacao
Public Function PV_CargaPosFormLoad(ByVal objCT As Object, bTrazendoDoc As Boolean) As Long
    PV_CargaPosFormLoad = SUCESSO
End Function

Public Function CargaPosFormLoad(Optional bTrazendoDoc As Boolean = False) As Long

Dim lErro As Long

On Error GoTo Erro_CargaPosFormLoad

    If (giPosCargaOk = 0) Then
    
        'p/permitir o redesenho da tela
        DoEvents
        
        gbCarregandoTela = True
             
        Set gobjEmbalagem = CreateObject("RotinasMat.ClassMATEmb")
        Set gobjEmbalagem.objTela = Me
        
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
        lErro = gobjTribTab.Ativar(Me, , , gobjTribTab.TIPOTELA_PV)
        If lErro <> SUCESSO Then gError 27643
        
        Set gobjPedidoVenda = New ClassPedidoDeVenda
        gobjPedidoVenda.dtDataEmissao = StrParaDate(DataEmissao.Text)
        lErro = gobjTribTab.TributacaoNF_Reset(gobjPedidoVenda)
        If lErro <> SUCESSO Then gError 27643
        '*** fim tributacao
        
''''        lErro = TributacaoPV_Reset()
''''        If lErro <> SUCESSO Then gError 27643
    
        'Carrega Filiais de Faturamento
        lErro = Carrega_FilialFaturamento()
        If lErro <> SUCESSO Then gError 46177
    
''''        Call AtualizarTabTrib
    
       If bTrazendoDoc = False Then

            'Call CallByName(NaturezaOp, "Refresh", VbMethod)
            Call CallByName(FilialFaturamento, "Refresh", VbMethod)

        End If
        
''''        lErro = CarregaTiposTrib()
''''        If lErro <> SUCESSO Then gError 27638
    
        'Carrega os Estados
        lErro = Carrega_PlacaUF()
        If lErro <> SUCESSO Then gError 46531
    
        'Carrega a combo combo de Tabela de Preços
        lErro = Carrega_TabelaPreco()
        If lErro <> SUCESSO Then gError 26481
    
        'Carrega a combo de Transportadora com as Transportadoras
        lErro = Carrega_Transportadoras()
        If lErro <> SUCESSO Then gError 26483
    
        'Incluído por Luiz Nogueira em 21/08/03
        'Carrega a combo VolumeEspecie
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
        If lErro <> SUCESSO Then gError 102389
    
        'Incluído por Luiz Nogueira em 21/08/03
        'Carrega a combo VolumeMarca
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
        If lErro <> SUCESSO Then gError 102390
    
        'Carrega a combo de Canais de Venda
        lErro = Carrega_CanalVenda()
        If lErro <> SUCESSO Then gError 26485
    
        'Carrega a combo de Tipos de Bloqueio
        lErro = Carrega_TipoBloqueio()
        If lErro <> SUCESSO Then gError 26487
    
        'Carrega a combo de Condição de Pagamento
        lErro = CF("Carrega_CondicaoPagamento", CondicaoPagamento, MODULO_CONTASARECEBER)
        If lErro <> SUCESSO Then gError 26490
    
        'Carrega a combo de Tipo de desconto
        lErro = Carrega_TipoDesconto()
        If lErro <> SUCESSO Then gError 26491
    
        PrecoUnitario.Format = gobjFAT.sFormatoPrecoUnitario
        
        Quantidade.Format = FORMATO_ESTOQUE
        QuantCancelada.Format = FORMATO_ESTOQUE
    
        'Preenche Data Referencia e Data de Emissão coma Data Atual
        DataReferencia.PromptInclude = False
        DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
        DataReferencia.PromptInclude = True
        giDataReferenciaAlterada = 0
    
        DataRefFluxo.PromptInclude = False
        DataRefFluxo.Text = Format(gdtDataAtual, "dd/mm/yy")
        DataRefFluxo.PromptInclude = True
        giDataRefFluxoAlterada = 0
    
    
        'Iniciliza as variáveis de browse
        Set objEventoMensagem = New AdmEvento
        Set objEventoCliente = New AdmEvento
        Set objEventoNumero = New AdmEvento
        Set objEventoCondPagto = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoEstoqueProd = New AdmEvento
        Set objEventoVendedor = New AdmEvento
        Set objEventoNaturezaOp = New AdmEvento
        Set objEventoTransportadora = New AdmEvento
        Set objEventoTranspRedespacho = New AdmEvento
        Set objEventoContato = New AdmEvento
        
        Set objEventoOrcamento = New AdmEvento 'Por Leo em 28/03/02
        
        Set objGridItens = New AdmGrid
        Set objGridComissoes = New AdmGrid
        Set objGridParcelas = New AdmGrid
        Set objGridBloqueio = New AdmGrid
        Set objGridAlocacao = New AdmGrid
    
        Set gcolItemPedido = New Collection
    
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
        If lErro <> SUCESSO Then gError 26636
        
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAlmox)
        If lErro <> SUCESSO Then gError 26636
    
        'Faz as Inicializações dos Grids
        lErro = CF2(Me, "PV_Inicializa_Grid_Itens", objGridItens)
        If lErro <> SUCESSO Then gError 26493
    
        'Janaina
        lErro = gobjEmbalagem.Inicializa()
        If lErro <> SUCESSO Then gError 96126
        'Janaina
        
        Set objTabComissoes.objTela = Me
    
        lErro = objTabComissoes.Inicializa_Grid_Comissoes(objGridComissoes)
        If lErro <> SUCESSO Then gError 26495
        
        'alterado por cyntia
        
        If m_objUserControl.Height > 9000 Then
            objGridComissoes.iLinhasVisiveis = 16
            objGridComissoes.objGrid.Rows = objGridComissoes.iLinhasVisiveis + 1
        Else
            objGridComissoes.iLinhasVisiveis = 5
        End If
        
        'Chama função que inicializa o Grid
        Call Grid_Inicializa(objGridComissoes)
    
        lErro = Inicializa_Grid_Parcelas(objGridParcelas)
        If lErro <> SUCESSO Then gError 26496
    
        lErro = CF2(Me, "PV_Inicializa_Grid_Parcelas", objGridParcelas)
        If lErro <> SUCESSO Then gError 124180
    
        lErro = Inicializa_Grid_Bloqueio(objGridBloqueio)
        If lErro <> SUCESSO Then gError 26497
    
        lErro = Inicializa_Grid_Reserva(objGridAlocacao)
        If lErro <> SUCESSO Then gError 26635
        
        'Para versão LIGHT Pedido Fatura Integral
'        If giTipoVersao = VERSAO_LIGHT Then
'            FaturaIntegral.Value = vbChecked
'            TabelaPreco.ListIndex = 0
'        End If

        lErro = CF2(Me, "PV_CargaPosFormLoad", bTrazendoDoc)
        If lErro <> SUCESSO Then gError 124162
        
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
          
        Case 46531, 26481, 26483, 26485, 26487, 26490, 26491
        Case 26493, 26636, 26495, 26496, 26497, 26635
        Case 27638, 46177, 27643, 96126, 102389, 102390, 124162
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157910)
     
    End Select
     
    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    giPosCargaOk = 0
    
    giFrameAtual = 1
    
    bPulaValorTotalCalcula = False
    iTratamentoPadraoFaltaEstoque = 0
    
    gbSaiuGrade = False

    If giTipoVersao = VERSAO_LIGHT Then
        
        FrameOrcVenda.Visible = False
        BotaoAnaliseVenda.Visible = False
        Label1(17).Visible = False 'label de filial de faturamento
        FilialFaturamento.Visible = False
        BotaoGrade.Visible = False
        
        'sumir com aba de embalagem
        Opcao.Tabs.Remove (TAB_Embalagens)
        
    End If
        
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    If gobjFAT.iImprimePVGravacao <> 0 Then ImprimeGravacao.Value = vbChecked
    
    If gobjFAT.iUsaParcDePV = MARCADO Then
        FrameParc.Visible = True
        FrameCodBase.Visible = True
    End If
    
    Call Carrega_FilialOrcamento
    
    'Carrega a combo de RecebForaFaixa
    lErro = Carrega_RecebForaFaixa()
    If lErro <> SUCESSO Then gError 196959

    iAlterado = 0
    iComissoesAlterada = 0
    gbCarregandoTela = True

    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207407

    Call CF2(Me, "PV_Form_Load")

    Set gcolParcPV = New Collection

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 196959, 207407

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157911)

    End Select

    iAlterado = 0
    iComissoesAlterada = 0
    
    Exit Sub

End Sub

Private Function Carrega_TabelaPreco() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TabelaPreco

    TabPrecoItemPV.AddItem ""
    TabPrecoItemPV.ItemData(TabPrecoItemPV.NewIndex) = 0
    
    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePrecoVenda", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao, "Ativo = 1")
    If lErro <> SUCESSO Then gError 26482

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        TabelaPreco.AddItem CStr(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TabelaPreco.ItemData(TabelaPreco.NewIndex) = objCodDescricao.iCodigo

        TabPrecoItemPV.AddItem CStr(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TabPrecoItemPV.ItemData(TabPrecoItemPV.NewIndex) = objCodDescricao.iCodigo
    
    Next
    
    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = gErr

    Select Case gErr

        Case 26482

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157915)

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
    If lErro <> SUCESSO Then gError 26484

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

        Case 26484

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157916)

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
    If lErro <> SUCESSO Then gError 26486

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

        Case 26486

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157917)

    End Select

    Exit Function

End Function

Private Function Carrega_TipoBloqueio() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TipoBloqueio

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TiposDeBloqueio", "Codigo", "NomeReduzido", STRING_TIPO_BLOQUEIO_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 26488

    For Each objCodDescricao In colCodigoDescricao

        If objCodDescricao.iCodigo <> BLOQUEIO_PARCIAL And objCodDescricao.iCodigo <> BLOQUEIO_NAO_RESERVA And objCodDescricao.iCodigo <> BLOQUEIO_CREDITO Then
            'Adiciona o item na Lista de Tabela de Preços
            TipoBloqueio.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
            TipoBloqueio.ItemData(TipoBloqueio.NewIndex) = objCodDescricao.iCodigo
        End If
    Next

    Carrega_TipoBloqueio = SUCESSO

    Exit Function

Erro_Carrega_TipoBloqueio:

    Carrega_TipoBloqueio = gErr

    Select Case gErr

        Case 26488

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157918)

    End Select

    Exit Function

End Function

'Private Function Carrega_CondicaoPagamento() As Long
'
'Dim lErro As Long
'Dim colCod_DescReduzida As New AdmColCodigoNome
'Dim objCod_DescReduzida As AdmCodigoNome
'
'On Error GoTo Erro_Carrega_CondicaoPagamento
'
'    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
'    lErro = CF("CondicoesPagto_Le_Recebimento", colCod_DescReduzida)
'    If lErro <> SUCESSO Then gError 26489
'
'    For Each objCod_DescReduzida In colCod_DescReduzida
'
'        'Adiciona novo item na List da Combo CondicaoPagamento
'        If gobjCRFAT.iCondPagtoSemCodigo = 0 Then
'            CondicaoPagamento.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
'        Else
'            CondicaoPagamento.AddItem objCod_DescReduzida.sNome
'        End If
'        CondicaoPagamento.ItemData(CondicaoPagamento.NewIndex) = objCod_DescReduzida.iCodigo
'
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
'        Case 26489
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157919)
'
'    End Select
'
'    Exit Function
'
'End Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157920)

    End Select

    Exit Function

End Function

'@@@@@ customizada
Public Function PV_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

Dim iColunasEscondidas As Integer

    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant. Pedida")
    If gobjFAT.iExibeQtdeCanceladaItensPV = MARCADO Then
        objGridInt.colColuna.Add ("Quant. Cancelada")
    End If
    If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        objGridInt.colColuna.Add ("Tab. Preço")
    End If
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Data Entrega")
    objGridInt.colColuna.Add ("Quant Reservada")
    objGridInt.colColuna.Add ("Quant Faturada")
    objGridInt.colColuna.Add ("Prioridade")

    If gobjFAT.iExibePercFatPV = MARCADO Then
        objGridInt.colColuna.Add ("% a Menos Fat")
        objGridInt.colColuna.Add ("% a Mais Fat")
        objGridInt.colColuna.Add ("Ação Receb Fora Faixa")
        objGridInt.colColuna.Add ("Quant Fat. A Mais")
    End If
    
    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        objGridInt.colColuna.Add ("Comissão")
    End If

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoProduto.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    
    If gobjFAT.iExibeQtdeCanceladaItensPV = MARCADO Then
        objGridInt.colCampo.Add (QuantCancelada.Name)
    Else
        QuantCancelada.left = POSICAO_FORA_TELA
    End If
    
    If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        objGridInt.colCampo.Add (TabPrecoItemPV.Name)
    Else
        TabPrecoItemPV.left = POSICAO_FORA_TELA
    End If
    
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    objGridInt.colCampo.Add (PrecoTotalB.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)
    objGridInt.colCampo.Add (QuantReservadaPV.Name)
    objGridInt.colCampo.Add (QuantFaturada.Name)
    objGridInt.colCampo.Add (Prioridade.Name)

    If gobjFAT.iExibePercFatPV = MARCADO Then
        objGridInt.colCampo.Add (PercentMenosReceb.Name)
        objGridInt.colCampo.Add (PercentMaisReceb.Name)
        objGridInt.colCampo.Add (RecebForaFaixa.Name)
        objGridInt.colCampo.Add (QuantFatAMais.Name)
    Else
        PercentMenosReceb.left = POSICAO_FORA_TELA
        PercentMaisReceb.left = POSICAO_FORA_TELA
        RecebForaFaixa.left = POSICAO_FORA_TELA
        QuantFatAMais.left = POSICAO_FORA_TELA
    End If
    
    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        objGridInt.colCampo.Add (ComissaoItemPV.Name)
    Else
        ComissaoItemPV.left = POSICAO_FORA_TELA
    End If
    
    'Colunas do Grid
    iGrid_ItemProduto_Col = 0
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_QuantPedida_Col = 4
        
    If gobjFAT.iExibeQtdeCanceladaItensPV = MARCADO Then
        iGrid_QuantCancel_Col = 5
    Else
        iColunasEscondidas = iColunasEscondidas + 1
    End If
    
    If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        iGrid_TabPrecoItem_Col = 6 - iColunasEscondidas
    Else
        iColunasEscondidas = iColunasEscondidas + 1
    End If
    
    iGrid_ValorUnitario_Col = 7 - iColunasEscondidas
    iGrid_PercDesc_Col = 8 - iColunasEscondidas
    iGrid_Desconto_Col = 9 - iColunasEscondidas
    iGrid_PrecoTotal_Col = 10 - iColunasEscondidas
    iGrid_PrecoTotalB_Col = 11 - iColunasEscondidas
    iGrid_DataEntrega_Col = 12 - iColunasEscondidas
    iGrid_QuantReservada_Col = 13 - iColunasEscondidas
    iGrid_QuantFaturada_Col = 14 - iColunasEscondidas
    iGrid_Prioridade_Col = 15 - iColunasEscondidas
    
    If gobjFAT.iExibePercFatPV = MARCADO Then

        iGrid_PercentMenosReceb_Col = 16 - iColunasEscondidas
        iGrid_PercentMaisReceb_Col = 17 - iColunasEscondidas
        iGrid_RecebForaFaixa_Col = 18 - iColunasEscondidas
        iGrid_QuantFatAMais_Col = 19 - iColunasEscondidas
    Else
        iColunasEscondidas = iColunasEscondidas + 4
    End If
    
    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        iGrid_ComissaoItem_Col = 20 - iColunasEscondidas
    Else
        iColunasEscondidas = iColunasEscondidas + 1
    End If

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    PV_Inicializa_Grid_Itens = SUCESSO

    Exit Function

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
    If m_objUserControl.Height > 9000 Then
        objGridInt.iLinhasVisiveis = 18
    Else
        objGridInt.iLinhasVisiveis = 7
    End If

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 700

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Bloqueio(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Usuário")
    objGridInt.colColuna.Add ("Responsável")
    objGridInt.colColuna.Add ("Data Liberação")
    objGridInt.colColuna.Add ("Resp. Liberação")
    objGridInt.colColuna.Add ("Observação")
    objGridInt.colColuna.Add ("Seq.")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (TipoBloqueio.Name)
    objGridInt.colCampo.Add (DataBloqueio.Name)
    objGridInt.colCampo.Add (CodUsuario.Name)
    objGridInt.colCampo.Add (ResponsavelBL.Name)
    objGridInt.colCampo.Add (DataLiberacao.Name)
    objGridInt.colCampo.Add (ResponsavelLib.Name)
    objGridInt.colCampo.Add (Observacao.Name)
    objGridInt.colCampo.Add (SeqBloqueio.Name)

    iGrid_TipoBloqueio_Col = 1
    iGrid_DataBloqueio_Col = 2
    iGrid_CodUsuario_Col = 3
    iGrid_ResponsavelBL_Col = 4
    iGrid_DataLiberacao_Col = 5
    iGrid_ResponsavelLib_Col = 6
    iGrid_Observacao_Col = 7
    iGrid_SeqBloqueio_Col = 8

    'Grid do GridInterno
    objGridInt.objGrid = GridBloqueio

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_BLOQUEIOS + 1

    'Linhas visíveis do grid
    If m_objUserControl.Height > 9000 Then
        objGridInt.iLinhasVisiveis = 19
    Else
        objGridInt.iLinhasVisiveis = 8
    End If

    'Largura da primeira coluna
    GridBloqueio.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Bloqueio = SUCESSO

    Exit Function

End Function

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
    If m_objUserControl.Height > 9000 Then
        objGridInt.iLinhasVisiveis = 24
    Else
        objGridInt.iLinhasVisiveis = 10
    End If

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
    sTabela = "PedidosDeVenda"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Pedido_Memoria(objPedidoVenda, False)
    If lErro <> SUCESSO Then gError 51159

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

        Case 51159

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157921)

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
        If lErro <> SUCESSO Then gError 51160

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 51160

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157922)

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
Dim objCliente As ClassCliente
Dim iStatus As Integer
Dim objNFiscal As New ClassNFiscal
Dim objParcelas As Object
Dim objParcPV As ClassParcelaPedidoVenda
Dim iIndice As Integer
Dim dPercDesc As Double

On Error GoTo Erro_Traz_Pedido_Tela

    gbCarregandoTela = True
        
    Call Limpa_Tela_PedidoVenda2
    
    gbCarregandoTela = True
        
    'Se a função não for chamada para carregar os dados de um orçamento de venda.
    If (objOrcamentoVenda Is Nothing) Then
    
         Call CF2(Me, "PV_Codigo_Enabled", False)
    
        'Lê os dados do Pedido de Venda
        lErro = CF("PedidoDeVenda_Le_Todos_Completo", objPedidoVenda, iStatus)
        If lErro <> SUCESSO And lErro <> 60304 Then gError 51161
        
        Set objOrcamentoVenda = New ClassOrcamentoVenda
        
        objOrcamentoVenda.lNumIntPedVenda = objPedidoVenda.lNumIntDoc
        
        'Lê os dados do OrcamentoVenda através do seu NumIntPedVenda
        lErro = CF("OrcamentoVenda_Le_NumIntPedVenda", objOrcamentoVenda)
        If lErro <> SUCESSO And lErro <> 94456 Then gError 94485
                
    Else
    
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
    
    lErro = gobjTribTab.Traz_NFiscal_Tela(objPedidoVenda)
    If lErro <> SUCESSO Then gError 181676
    
    If Not (objOrcamentoVenda Is Nothing) Then
        Call PedidoVenda_Atualiza_Orcamento(objOrcamentoVenda)
    End If

    lErro = CF("ReservasPedido_Le", objPedidoVenda)
    If lErro <> SUCESSO Then gError 46044
       
    ValorTotal.Caption = Format(objPedidoVenda.dValorTotal, "Standard")
    ValorProdutos.Caption = Format(objPedidoVenda.dValorProdutos, "Standard")
    
    PercAcrescFin.Text = Format(objPedidoVenda.dPercAcrescFinanceiro * 100, "Fixed")
    'Coloca os dados do Pedido na tela
    
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
        If lErro <> SUCESSO And lErro <> 12348 Then gError 102196
        
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
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 102197
            
            'Se não conseguiu selecionar
            If lErro = 6730 Or lErro = 6731 Then
            
                'Indica que a filial do cliente não foi encontrada para que seja exibido tratamento de erro posteriormente
                'O erro não pode ser tratado aqui, pois a tela deve ser carregada normalmente
                giFilialCliInexistente = FILIALCLIENTE_INEXISTENTE
            
            'Se conseguiu selecionar
            Else
            
                'Tenta selecionar a filialentrega na combo
                lErro = Combo_Seleciona(FilialEntrega, 0)
                If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 102198
            
                'Se não conseguiu selecionar => 'Indica que a filial de entrega não foi encontrada para que seja exibido tratamento de erro posteriormente
                If lErro = 6730 Or lErro = 6731 Then giFilialCliInexistente = FILIALENTREGA_INEXISTENTE
                
            End If


        End If
            
    End If
    
    For Each objFilial In gcolFiliais
        
        If objFilial.iCodFilial = objPedidoVenda.iFilialEmpresaFaturamento Then
            
            FilialFaturamento.Text = objFilial.iCodFilial & SEPARADOR & objFilial.sNome
            Exit For
        End If
    
    Next
    
    giFilialFatAlterada = 0
    
    If objPedidoVenda.lCodigo <> 0 Then
        Codigo.Text = objPedidoVenda.lCodigo
    End If
    
''''    NaturezaOp.Text = objPedidoVenda.sNaturezaOp
''''    Call NaturezaOp_Validate(bSGECancelDummy)
    FaturaIntegral.Value = objPedidoVenda.iFaturaIntegral
    
    ValorFrete.Text = Format(objPedidoVenda.dValorFrete, "Standard")
    ValorSeguro.Text = Format(objPedidoVenda.dValorSeguro, "Standard")
    ValorDesconto.Text = Format(objPedidoVenda.dValorDesconto, "Standard")
    ValorDespesas.Text = Format(objPedidoVenda.dValorOutrasDespesas, "Standard")
    
    lErro = CF2(Me, "PV_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132004
   
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
    
    'Alterado por Luiz Nogueira em 21/08/03
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

    For iIndice = 0 To TipoFrete.ListCount - 1
        If TipoFrete.ItemData(iIndice) = objPedidoVenda.iFreteRespons Then
            TipoFrete.ListIndex = iIndice
            Exit For
        End If
    Next
            
    If objPedidoVenda.dPesoLiq > 0 Then PesoLiquido.Text = Format(objPedidoVenda.dPesoLiq, PesoLiquido.Format) 'Alterado por Wagner
    If objPedidoVenda.dPesoBruto > 0 Then PesoBruto.Text = Format(objPedidoVenda.dPesoBruto, PesoBruto.Format) 'Alterado por Wagner
    If objPedidoVenda.dVolumeTotal > 0 Then Cubagem.Text = Format(objPedidoVenda.dVolumeTotal, "Fixed")

    lErro = CF2(Me, "PV_Traz_Pedido_Tela", objPedidoVenda, objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 124181
    
    'Carrega o Tab Complemento
    lErro = Carrega_Tab_Complemento(objPedidoVenda)
    If lErro <> SUCESSO Then gError 51163

    'Carrega o Tab Cobrança
    lErro = Carrega_Tab_Cobranca(objPedidoVenda)
    If lErro <> SUCESSO Then gError 51164

    Set gcolParcPV = New Collection
   
    For iIndice = 1 To objPedidoVenda.colParcelas.Count
   
        gcolParcPV.Add objPedidoVenda.colParcelas(iIndice)
    Next

    'Alterado por Luiz Nogueira em 20/05/03
    'Carrega o Grid de itens, indicando se é para processar reservas ou não
    lErro = Carrega_Grid_Itens(objPedidoVenda, bProcessarInfoPedido)
    If lErro <> SUCESSO Then gError 26569
    
    lErro = CF2(Me, "PV_Traz_Pedido_Tela2", objPedidoVenda, objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 124181
    
    'Incluído por Luiz Nogueira em 20/05/03
    'Se não é necessário completar os dados do pedido, ou seja, não está usando um orçamento para
    'carregar a tela, significa que pode carregar todas as informações a partir de objPedidoVenda
    If Not bProcessarInfoPedido Then
    
        'Carrega o Tab Comissões
        lErro = objTabComissoes.Carrega_Tab_Comissoes(objPedidoVenda)
        If lErro <> SUCESSO Then gError 51165
    
        'Carrega o Tab Bloqueio
        lErro = Carrega_Tab_Bloqueio(objPedidoVenda)
        If lErro <> SUCESSO Then gError 26633
    
        Call Preenche_GridAlocacao(objPedidoVenda.colItensPedido)
        
        iTipoDocOrigem = DOCORIGEM_PV
        
        lErro = CF("EmbalagensProdItem_Le", objPedidoVenda.lNumIntDoc, iTipoDocOrigem, gcolItemPedido, colEmbProdItem)
        If lErro <> SUCESSO Then gError 96069
    
        'Preenche o Grid com as Embalagens dos itens do Pedido de Venda
        lErro = gobjEmbalagem.Preenche_GridEmbalagens(colEmbProdItem)
        If lErro <> SUCESSO Then gError 89500
            
    'Senão, ou seja, se está carregando um orçamento para a tela
    Else
    
        '*** Incluído por Luiz Nogueira em 26/01/04 ***
        'Transferir para versão 2
        'Se o cliente/filial do orçamento estão cadastrados
        If giClienteInexistente <> CLIENTE_INEXISTENTE And giFilialCliInexistente <> FILIALCLIENTE_INEXISTENTE Then
            'Exibe os dados default do cliente que não gravados no orçamento
            'Efetua o cálculo de comissões, se necessário
            lErro = Traz_Orcamento_Cliente(objCliente, objPedidoVenda)
            If lErro <> SUCESSO Then gError 102253
        End If
        '***********************************************
        
        'calcula as comissões para o pedido
        Call ComissaoAutomatica_Click
        
        'calcula as embalagens do pedido
        CalculoAuto_Click
    
        lErro = CF2(Me, "PV_Calcula_PesoKg_Total")
        If lErro <> SUCESSO Then gError 177159
    
        If objOrcamentoVenda.lVolumeQuant > 0 Then VolumeQuant.Text = objOrcamentoVenda.lVolumeQuant
        If objOrcamentoVenda.dPesoLiq > 0 Then PesoLiquido.Text = Format(objOrcamentoVenda.dPesoLiq, PesoLiquido.Format) 'Alterado por Wagner
        If objOrcamentoVenda.dPesoBruto > 0 Then PesoBruto.Text = Format(objOrcamentoVenda.dPesoBruto, PesoBruto.Format) 'Alterado por Wagner
        If objOrcamentoVenda.dVolumeTotal > 0 Then Cubagem.Text = Format(objOrcamentoVenda.dVolumeTotal, "Fixed")
        
    End If
    
    ValorTotal.Caption = Format(objPedidoVenda.dValorTotal, "Standard")

    objNFiscal.lCliente = objPedidoVenda.lCliente
    objNFiscal.iFilialCli = objPedidoVenda.iFilial
    
    lErro = CF2(Me, "NFCF_Cliente_Traz_NF_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 131961
   
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objPedidoVenda.objProjetoInfo, objPedidoVenda.lNumIntDoc, PRJ_CR_TIPO_PV)
    If lErro <> SUCESSO Then gError 181676
    
    If objPedidoVenda.dValorItens = 0 Then
        For iIndice = 1 To objGridItens.iLinhasExistentes
            objPedidoVenda.dValorItens = objPedidoVenda.dValorItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
            objPedidoVenda.dValorDescontoItens = objPedidoVenda.dValorDescontoItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
        Next
    End If
    
    ValorProdutos2.Caption = Format(objPedidoVenda.dValorItens, "Standard")
    ValorDescontoItens.Text = Format(objPedidoVenda.dValorDescontoItens, "Standard")
    If objPedidoVenda.dValorItens > 0 Then
        dPercDesc = objPedidoVenda.dValorDescontoItens / objPedidoVenda.dValorItens
    Else
        dPercDesc = 0
    End If
    PercDescontoItens.Text = Format(dPercDesc * 100, "Fixed")
    
    dValorDescontoItensAnt = objPedidoVenda.dValorDescontoItens
    dPercDescontoItensAnt = dPercDesc
    
    Set gobjInfoAdic = objPedidoVenda.objInfoAdic
    
    ValorDescontoTit.Text = Format(objPedidoVenda.dValorDescontoTit, "Standard")
    Call ValorDescontoTit_Validate(bSGECancelDummy)
    
    If Not (objOrcamentoVenda Is Nothing) Then
        If objOrcamentoVenda.iMoeda <> MOEDA_REAL Then
            For iIndice = 1 To objGridItens.iLinhasExistentes
                lErro = gobjTribTab.Alteracao_Item_Grid(iIndice)
                If lErro <> SUCESSO Then gError 177159
                
                Call PrecoTotal_Calcula(iIndice)
            Next
            Call ValorTotal_Calcula
        End If
    End If
    
    If gobjFAT.iUsaParcDePV = MARCADO Then
    
        If objPedidoVenda.iParc <> 0 Then
            Parc.PromptInclude = False
            Parc.Text = CStr(objPedidoVenda.iParc)
            Parc.PromptInclude = True
        End If
        If objPedidoVenda.lCodigoBase <> 0 Then
            CodBase.Caption = CStr(objPedidoVenda.lCodigoBase)
        End If
    End If
    
    iAlterado = 0
    iComissoesAlterada = 0
    giPercAcresFinAlterado = 0
    
    iAlterado = 0
    
    iComissoesAlterada = 0
    
    gbCarregandoTela = False
     
    Traz_Pedido_Tela = SUCESSO

    Exit Function

Erro_Traz_Pedido_Tela:

    gbCarregandoTela = False
     
    Traz_Pedido_Tela = gErr

    Select Case gErr

        Case 51161, 26569, 51163, 51164, 51165, 26633, 27640, 27642, 46044, 51162, 96069, 89500, 94463, 94485, 102196, 102197, 102198, 102253, 124181, 131961, 132004, 177159, 181676

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157923)

    End Select

    Exit Function

End Function

'Alterado por Luiz Nogueira em 20/05/03
Private Function Carrega_Grid_Itens(objPedidoVenda As ClassPedidoDeVenda, Optional bProcessaReserva As Boolean = False) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objProduto As New ClassProduto
Dim iItem As Integer

On Error GoTo Erro_Carrega_Grid_Itens

    Set gcolItemPedido = New Collection

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    For iIndice = 1 To objPedidoVenda.colItensPedido.Count
       
        lErro = Mascara_RetornaProdutoEnxuto(objPedidoVenda.colItensPedido(iIndice).sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 26649

        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
        
        objProduto.sCodigo = objPedidoVenda.colItensPedido(iIndice).sProduto
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177926
        
        'Calcula o percentual de desconto
        If objPedidoVenda.colItensPedido(iIndice).dPrecoTotal > 0 Then
            dPercDesc = objPedidoVenda.colItensPedido(iIndice).dValorDesconto / (objPedidoVenda.colItensPedido(iIndice).dPrecoTotal + objPedidoVenda.colItensPedido(iIndice).dValorDesconto)
        Else
            dPercDesc = 0
        End If
        
        objPedidoVenda.colItensPedido(iIndice).iItem = iIndice
        
        'Coloca os dados dos itens na tela
        
        '****** IF INCLUÍDO PARA TRATAMENTO DE GRADE ***************
        If objPedidoVenda.colItensPedido(iIndice).iPossuiGrade = MARCADO And objProduto.iKitVendaComp <> MARCADO Then GridItens.TextMatrix(iIndice, 0) = "# " & GridItens.TextMatrix(iIndice, 0)
        
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objPedidoVenda.colItensPedido(iIndice).sDescricao 'sProdutoDescricao
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objPedidoVenda.colItensPedido(iIndice).sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col) = Formata_Estoque(objPedidoVenda.colItensPedido(iIndice).dQuantidade)
        If gobjFAT.iExibeQtdeCanceladaItensPV = MARCADO Then GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col) = Formata_Estoque(objPedidoVenda.colItensPedido(iIndice).dQuantCancelada)
        
        If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        
            If objPedidoVenda.colItensPedido(iIndice).iTabelaPreco = 0 Then
            
                GridItens.TextMatrix(iIndice, iGrid_TabPrecoItem_Col) = ""
            
            Else
            
                For iItem = 0 To TabPrecoItemPV.ListCount - 1
                    If objPedidoVenda.colItensPedido(iIndice).iTabelaPreco = TabPrecoItemPV.ItemData(iItem) Then
                        'coloca no Grid Itens RecebForaFaixa
                        GridItens.TextMatrix(iIndice, iGrid_TabPrecoItem_Col) = TabPrecoItemPV.List(iItem)
                        Exit For
                    End If
                Next
        
            End If
        End If
        
        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dPercDesc, "Percent")
        GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dPrecoTotal, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dPrecoTotal + objPedidoVenda.colItensPedido(iIndice).dValorDesconto, "Standard")
        If objPedidoVenda.colItensPedido(iIndice).dtDataEntrega <> DATA_NULA Then GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dtDataEntrega, "dd/mm/yyyy")
        GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col) = Formata_Estoque(objPedidoVenda.colItensPedido(iIndice).dQuantReservada)
        GridItens.TextMatrix(iIndice, iGrid_QuantFaturada_Col) = Formata_Estoque(objPedidoVenda.colItensPedido(iIndice).dQuantFaturada)

        'preenche prioridade
        GridItens.TextMatrix(iIndice, iGrid_Prioridade_Col) = CStr(objPedidoVenda.colItensPedido(iIndice).iPrioridade)

        lErro = CF2(Me, "PV_Carrega_Grid_Itens_Linha", objPedidoVenda, iIndice)
        If lErro <> SUCESSO Then gError 132823

        lErro = CF2(Me, "PV_Preenche_GridItens", objPedidoVenda, objPedidoVenda.colItensPedido(iIndice), objProduto, iIndice, dPercDesc)
        If lErro <> SUCESSO Then gError 132823
        
        'Armazena os números internos dos itens
        gcolItemPedido.Add objPedidoVenda.colItensPedido(iIndice).lNumIntDoc

        'Incluído por Luiz em 20/05/03 para processar as reservas quando o pedido for gerado por um orçamento
        'Se for para processar reservas
        If bProcessaReserva = True Then
        
            If objPedidoVenda.colItensPedido(iIndice).iPossuiGrade = DESMARCADO And objProduto.iKitVendaComp <> MARCADO Then
            
                'Dispara a reserva do produto
                lErro = Reserva_Processa(iIndice, objPedidoVenda.colItensPedido(iIndice).dQuantidade, objPedidoVenda.colItensPedido(iIndice).dQuantCancelada, objPedidoVenda.colItensPedido(iIndice).dQuantFaturada)
                If lErro <> SUCESSO Then gError 102174
            
            Else
                
                For Each objItemRomaneio In objPedidoVenda.colItensPedido(iIndice).colItensRomaneioGrade
                    lErro = Reserva_Processa_ItemGrade(objItemRomaneio, iIndice)
                    If lErro <> SUCESSO Then gError 102174
                    
                Next
                
                Call Atualiza_Grid_Alocacao(objPedidoVenda.colItensPedido(iIndice))
                
                If objProduto.iKitVendaComp <> MARCADO Then
                    Call Atualiza_Grid_Itens(objPedidoVenda.colItensPedido(iIndice))
                End If
            
            End If
            
        End If
        
        objPedidoVenda.colItensPedido(iIndice).dQuantFaturadaAMais = objPedidoVenda.colItensPedido(iIndice).dQuantFaturada - objPedidoVenda.colItensPedido(iIndice).dQuantidade
        If objPedidoVenda.colItensPedido(iIndice).dQuantFaturadaAMais < 0 Then objPedidoVenda.colItensPedido(iIndice).dQuantFaturadaAMais = 0
        
        If gobjFAT.iExibePercFatPV = MARCADO Then

            GridItens.TextMatrix(iIndice, iGrid_PercentMaisReceb_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dPercentMaisReceb, "Percent")
            GridItens.TextMatrix(iIndice, iGrid_PercentMenosReceb_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dPercentMenosReceb, "Percent")
            GridItens.TextMatrix(iIndice, iGrid_QuantFatAMais_Col) = Formata_Estoque(objPedidoVenda.colItensPedido(iIndice).dQuantFaturadaAMais)
    
            For iItem = 0 To RecebForaFaixa.ListCount - 1
                If objPedidoVenda.colItensPedido(iIndice).iRecebForaFaixa = RecebForaFaixa.ItemData(iItem) Then
                    'coloca no Grid Itens RecebForaFaixa
                    GridItens.TextMatrix(iIndice, iGrid_RecebForaFaixa_Col) = RecebForaFaixa.List(iItem)
                End If
            Next
        
        End If
        
        If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        
            If objPedidoVenda.colItensPedido(iIndice).dComissao <> -1 Then
                GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dComissao, "Percent")
            Else
                GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col) = ""
            End If
        
        End If
    
    Next

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = gcolItemPedido.Count

    lErro = CF2(Me, "PV_Carrega_Grid_Itens", objPedidoVenda, bProcessaReserva)
    If lErro <> SUCESSO Then gError 124180
        
    'Atualiza o grid para mostrar as checkboxes marcadas / desmarcadas
    Call Grid_Refresh_Checkbox(objGridItens)
        
    Exit Function

Erro_Carrega_Grid_Itens:

    Carrega_Grid_Itens = gErr

    Select Case gErr

        Case 26649
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objPedidoVenda.colItensPedido(iIndice).sProduto)

        Case 102174, 132823, 177926 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157924)

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

    Call DateParaMasked(DataRefFluxo, objPedidoVenda.dtDataRefFluxo)
    giDataRefFluxoAlterada = 0

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


Function Carrega_Tab_Bloqueio(objPedidoVenda As ClassPedidoDeVenda) As Long
'Coloca os dados do Tab de Bloqueio na tela

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objBloqueioPV As ClassBloqueioPV
Dim objTipoBloqueio As New ClassTipoDeBloqueio

On Error GoTo Erro_Carrega_Tab_Bloqueio

    'Limpa o Grid de Bloqueios antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridBloqueio)

    iIndice = 0

    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV

        iIndice = iIndice + 1

        objTipoBloqueio.iCodigo = objBloqueioPV.iTipoDeBloqueio

        'Lê o Tipo de bloqueio
        lErro = CF("TipoDeBloqueio_Le", objTipoBloqueio)
        If lErro <> SUCESSO And lErro <> 23666 Then gError 42308
        If lErro = 23666 Then gError 42309
        'Coloca o bloqieo no Grid de bloqueios
        GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col) = objTipoBloqueio.iCodigo & SEPARADOR & objTipoBloqueio.sNomeReduzido
        GridBloqueio.TextMatrix(iIndice, iGrid_DataBloqueio_Col) = Format(objBloqueioPV.dtData, "dd/mm/yy")
        GridBloqueio.TextMatrix(iIndice, iGrid_CodUsuario_Col) = objBloqueioPV.sCodUsuario
        GridBloqueio.TextMatrix(iIndice, iGrid_ResponsavelBL_Col) = objBloqueioPV.sResponsavel
        If objBloqueioPV.dtDataLib <> DATA_NULA Then GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col) = Format(objBloqueioPV.dtDataLib, "dd/mm/yy")
        GridBloqueio.TextMatrix(iIndice, iGrid_ResponsavelLib_Col) = objBloqueioPV.sResponsavelLib
        GridBloqueio.TextMatrix(iIndice, iGrid_Observacao_Col) = objBloqueioPV.sObservacao
        GridBloqueio.TextMatrix(iIndice, iGrid_SeqBloqueio_Col) = CStr(objBloqueioPV.iSequencial)

    Next
    'Guarda o número de bloqueios trazidos do BD
    giGridBloqueioLinhasBloqueadas = objPedidoVenda.colBloqueiosPV.Count

    objGridBloqueio.iLinhasExistentes = iIndice

    Carrega_Tab_Bloqueio = SUCESSO

    Exit Function

Erro_Carrega_Tab_Bloqueio:

    Carrega_Tab_Bloqueio = gErr

    Select Case gErr

        Case 42308
        
        Case 42309
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIO_NAO_CADASTRADO", gErr, objTipoBloqueio.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157925)

    End Select

End Function

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
        If lErro <> SUCESSO And lErro <> 28030 Then gError 59730
        If lErro <> SUCESSO Then gError 59731
        
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemPedido.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 59732
        
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
            If lErro <> SUCESSO And lErro <> 25056 Then gError 46181
            If lErro = 25056 Then gError 46182

            GridAlocacao.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido

        Next

    Next

    objGridAlocacao.iLinhasExistentes = iIndice

    Carrega_Grid_Reserva = SUCESSO

    Exit Function

Erro_Carrega_Grid_Reserva:

    Carrega_Grid_Reserva = gErr

    Select Case gErr

        Case 46181, 59730, 59731, 59732

        Case 46182
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157926)

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
Dim dValorTotal As Double, dValorTotalB As Double
Dim lErro As Long
Dim objItemOp As New ClassItemOP
Dim objItemPV As New ClassItemPedido
Dim vbMsgRes As VbMsgBoxResult 'Inserido por Wagner

On Error GoTo Erro_GridItens_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iItemAtual = GridItens.Row

    If KeyCode = vbKeyDelete And iItemAtual <= objGridItens.iLinhasExistentes Then
        
        'Não deixa deletar se a quantidade faturada for positiva
        If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantFaturada_Col)) > 0 Then gError 26647
        
        'Se ítem já existe no BD verifica se está ligado a Ordem Producao
        If gcolItemPedido(iItemAtual) <> 0 Then
        
            objItemPV.lNumIntDoc = gcolItemPedido(iItemAtual)
            
            lErro = CF("ItemOP_Le_ItemPV", objItemOp, objItemPV)
            If lErro <> SUCESSO And lErro <> 46074 Then gError 24458
            If lErro = SUCESSO Then gError 24457
        
        End If
    
        Call CF2(Me, "PV_Exclui_ItemPV", iItemAtual)
    
    End If

    If bPergunta Then
        Call Grid_Trata_Tecla1(KeyCode, objGridItens)
    Else
        Call Grid_Trata_Tecla3(KeyCode, objGridItens)
    End If

    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then

        '************ grade ************
        GridItens.TextMatrix(GridItens.Row, 0) = GridItens.Row
        
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

        Call gobjTelaProjetoInfo.Remove_Linha(iItemAtual)

        If objGridAlocacao.iLinhasExistentes > 0 Then
            Call Remove_Reservas(iItemAtual)
            For iIndice = 1 To objGridAlocacao.iLinhasExistentes
                If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))) > 0 Then If CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) > iItemAtual Then GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col) = CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) - 1
            Next
        End If
        
        'Calcula a soma dos valores de produtos
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
                If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
            End If
            dValorTotalB = dValorTotalB + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
        Next

        'Coloca valor total dos produtos na tela
        ValorProdutos.Caption = Format(dValorTotal, "Standard")
        ValorProdutos2.Caption = Format(dValorTotalB, "Standard")

        If objGridItens.iLinhasExistentes <> 0 Then
            Call CF2(Me, "PV_PrecoTotal_Calcula", objGridItens.iLinhasExistentes)
        Else
        
            If StrParaDbl(ValorDesconto.Text) <> 0 Then
                'Avisa ao usuário
                vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", ValorDesconto.Text, 0)
                
                'Limpa o valor de desconto
                gdDesconto = 0
                Call CF2(Me, "NFCF_Limpa_Desconto")
                ValorDesconto.Text = ""
                giValorDescontoAlterado = 0
    
                Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, 0)
    
                'Para tributação
                gobjPedidoVenda.dValorDesconto = 0
            End If
        End If
        
        lErro = CF2(Me, "PV_ValorDesconto_Validate")
        If lErro <> SUCESSO Then gError 132007

        'Calcula o valor total da nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 51033

        'acerta o grid de embalagens
        Call gobjEmbalagem.Exclusao_Item_GridEmb(iItemAtual)
        
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 500073
        
        lErro = CF2(Me, "PV_Calcula_PesoKg_Total")
        If lErro <> SUCESSO Then gError 177160
        
        For iIndice = 1 To gobjPedidoVenda.colItens.Count
            If gobjPedidoVenda.colItens(iIndice).iPossuiGrade = MARCADO Then
                GridItens.TextMatrix(iIndice, 0) = "# " & CStr(iIndice)
            Else
                GridItens.TextMatrix(iIndice, 0) = CStr(iIndice)
            End If
            gobjPedidoVenda.colItens(iIndice).iItem = iIndice 'Incluído para ao alterar a grade ou a reserva não fazer do item errado
        Next
        GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, 0) = CStr(objGridItens.iLinhasExistentes + 1)

    End If

    Exit Sub
    
Erro_GridItens_KeyDown:

    Select Case gErr
        
        Case 24457
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_PV_VINCULADO_ITEM_OP", gErr, iItemAtual, objItemOp.lNumIntDoc)
        
        Case 24458, 51033, 500073, 177160
        
        Case 26647
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_PV_QUANT_FAT_POSITIVA", gErr, iItemAtual)
        
        Case 132007 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157927)
            
    End Select

    Exit Sub

End Sub

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Public Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

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

Private Sub Class_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
End Sub

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

On Error GoTo Erro_NumeroLabel_Click

    lErro = Move_Pedido_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then gError 26498

    Call Chama_Tela("PedidoVendaLista", colSelecao, objPedidoVenda, objEventoNumero)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 26498

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157945)

    End Select

    Exit Sub

End Sub

Private Sub objEventoMensagem_evSelecao(obj1 As Object)

Dim objMensagem As ClassMensagem

    Set objMensagem = obj1
    
    Mensagem.Text = objMensagem.sDescricao
    
    Me.Show

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim objPedidoVenda As ClassPedidoDeVenda
Dim lErro As Long

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objPedidoVenda = obj1

    lErro = Traz_Pedido_Tela(objPedidoVenda)
    If lErro <> SUCESSO Then gError 26499

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 26499

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157946)

    End Select

    Exit Sub

End Sub

Public Sub CondPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection

    objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)

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
    CondicaoPagamento.Text = CondPagto_Traz(objCondicaoPagto)
    Call CondicaoPagamento_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoCondPagto_evSelecao:

    Select Case gErr

        Case 26500

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157947)

     End Select

     Exit Sub

End Sub

Public Sub BotaoProdutos_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String
Dim sDescricao As String
Dim sOrdenacao As String
Dim sNomeBrowse As String

On Error GoTo Erro_BotaoProdutos_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridItens.Row = 0 Then gError 58771

    sOrdenacao = "Código"

    If Me.ActiveControl Is Produto Then
        sProduto1 = Produto.Text
    Else
        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
    End If
       
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 58772
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    If Me.ActiveControl Is DescricaoProduto Then
        sOrdenacao = "Descrição"
        sDescricao = DescricaoProduto.Text
        If iPreenchido = PRODUTO_PREENCHIDO Then gError 180430
    Else
        sDescricao = GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col)
    End If
    
    'preenche o codigo do produto
    objProduto.sCodigo = sProduto
    objProduto.sDescricao = sDescricao
    
    Call CF2(Me, "PV_BotaoProdutos_Click", objProduto, sOrdenacao)
    Call CF2(Me, "PV_BotaoProdutos_Click2", sNomeBrowse)
    
    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela(sNomeBrowse, colSelecao, objProduto, objEventoProduto, "", sOrdenacao)

    Exit Sub
        
Erro_BotaoProdutos_Click:
    
    Select Case gErr
        
        Case 58771
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 58772, 180430 'Tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157948)

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
Dim lErro As Long, iExecutaEntradaCelula As Integer

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridItens.Row < 1 Then Exit Sub

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 55133
    
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
    
        'Incluído por Luiz Nogueira em 29/04/03
        If Not (Me.ActiveControl Is Produto) Then
        
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
        
            'Faz o Tratamento do produto
            'lErro = Traz_Produto_Tela()
            lErro = CF2(Me, "Produto_Saida_Celula")
            If lErro <> SUCESSO Then gError 31391
            
            If gbSaiuGrade Then
            
                gbSaiuGrade = False
                
                If gobjFAT.iPulaLinhaGrade = MARCADO Then
                
                    objGridItens.objGrid.Row = objGridItens.objGrid.Row + 1
                    objGridItens.objGrid.Col = 1
                   
                    objGridItens.objForm.Show
                    objGridItens.objGrid.SetFocus
                    
                    Call Grid_Trata_Tecla(vbKeyReturn, objGridItens, iExecutaEntradaCelula)
                
                    If iExecutaEntradaCelula = 1 Then
                        Call Grid_Entrada_Celula(objGridItens, iAlterado)
                    End If
                    
                End If
            
            End If
    
        End If
    
    End If
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    gbSaiuGrade = False

    Select Case gErr

        Case 31391
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""
            
        Case 55133
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157949)

    End Select

    Exit Sub

End Sub

Public Sub BotaoEstoqueProd_Click()

Dim lErro As Long
Dim objEstoqueProd As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_BotaoEstoqueProd_Click

    lErro = CF2(Me, "PV_Botao_Estoque", objEstoqueProd, colSelecao, sSelecao)
    If lErro <> SUCESSO Then gError 131881

    Call Chama_Tela("EstoqueProdutoLista", colSelecao, objEstoqueProd, objEventoEstoqueProd, sSelecao)
    
    Exit Sub

Erro_BotaoEstoqueProd_Click:
    
    Select Case gErr
    
        Case 131881 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157950)
    
    End Select
    
    Exit Sub
    
End Sub

Function Trata_Parametros(Optional objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim objPedidoVendaAux As New ClassPedidoDeVenda

On Error GoTo Erro_Trata_Parametros

    If Not (objPedidoVenda Is Nothing) Then
        
        lErro = CargaPosFormLoad(True)
        If lErro <> SUCESSO Then gError 59288
        
        'Se foi passado o código do orçamento
        If objPedidoVenda.lCodigo > 0 Then  'Incluído por Luiz Nogueira em 13/01/04
        
            objPedidoVendaAux.lCodigo = objPedidoVenda.lCodigo
            objPedidoVendaAux.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
            
            'Coloca o Pedido de Venda na tela
            lErro = Traz_Pedido_Tela(objPedidoVendaAux)
            If lErro <> SUCESSO And lErro <> 51162 Then gError 26512
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
        '*** Incluído por Luiz Nogueira em 13/01/04 - FIM ***
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

        Case 26512, 59288

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157951)

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
    If lErro <> SUCESSO Then gError 26513

    objPedidoVenda.lCodigo = StrParaLong(Codigo.Text)
    
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa 'Incluido por Leo em 18/02/02
    
    lErro = CF2(Me, "PV_Codigo_Validate", objPedidoVenda)
    If lErro <> SUCESSO Then gError 132824
    
    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 26513, 26683, 94329, 132824 'Trecho alterado por Leo em 18/02/02
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157952)

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
Dim iIndice As Integer
Dim colTipoFrete As Collection
Dim objTipoFrete As ClassTipoFrete

On Error GoTo Erro_Cliente_Validate
    
    If giClienteAlterado = 1 Then

        'Verifica se o Cliente está preenchido
        If Len(Trim(Cliente.Text)) > 0 Then
            'Busca o Cliente no BD
            lErro = TP_Cliente_Le3(Cliente, objCliente, iCodFilial, 0)
            If lErro <> SUCESSO And lErro <> 6668 And lErro <> 6676 And lErro <> 6601 And lErro <> 6704 Then gError 26135

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
                    If lErro <> SUCESSO Then gError 181160
                    
                Else
                    vbResult = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTE_COM_CONTATO", Cliente.Text, Cliente.Text)
                    If vbResult = vbNo Then gError 181161
                    
                    'Lê o Contato no BD
                    lErro = CF("Contato_Transforma_Cliente", objContato, objCliente)
                    If lErro <> SUCESSO Then gError 181162
                
                End If
            
            End If
            
            lErro = CF2(Me, "PV_Cliente_Validate", objCliente.lCodigo)
            If lErro <> SUCESSO Then gError 26135
            
            'gobjPedidoVenda.lCliente = objCliente.lCodigo
            
            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then gError 26136

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
            
            lErro = CF2(Me, "NFCF_Cliente_Validate", objCliente.lCodigo, iCodFilial)
            If lErro <> SUCESSO Then gError 131959
            
            giFilialEntregaManual = 0
            
            'Se o Tipo estiver preenchido
            If objCliente.iTipo > 0 Then
                objTipoCliente.iCodigo = objCliente.iTipo
                'Lê o Tipo de Cliente
                
                lErro = CF("TipoCliente_Le", objTipoCliente)
                If lErro <> SUCESSO And lErro <> 19062 Then gError 26514
                
            End If
            
            lErro = CF("TipoFrete_Le", colTipoFrete)
            If lErro <> SUCESSO Then gError 207413
            
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
            
            giValorDescontoManual = 0

            lErro = CF2(Me, "NFCF_Cliente_Validate2", objCliente, objTipoCliente)
            If lErro <> SUCESSO Then gError 51034
        
            If Not gbCarregandoTela Then
                
                Call DescontoGlobal_Recalcula

                'ATualiza o total com o novo desconto
                lErro = ValorTotal_Calcula()

                If lErro <> SUCESSO Then gError 51034
                
                'Coloca na tela a tabela do cliente
                If objCliente.iTabelaPreco > 0 Then
                    TabelaPreco.Text = objCliente.iTabelaPreco
                    Call TabelaPreco_Validate(bSGECancelDummy)
                ElseIf objTipoCliente.iTabelaPreco > 0 Then
                    TabelaPreco.Text = objTipoCliente.iTabelaPreco
                    Call TabelaPreco_Validate(bSGECancelDummy)
                End If
                
                lErro = CF2(Me, "PV_Cliente_Validate2", objCliente.lCodigo)
                If lErro <> SUCESSO Then gError 26135
                
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

                lErro = CF2(Me, "NFCF_Cliente_Validate3", objCliente)
                If lErro <> SUCESSO Then gError 51034

            End If
            
        'Se não estiver preenchido
        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            'Limpa a Combo de Filiais
            Filial.Clear

            FilialEntrega.Clear

            Call CF2(Me, "NFCF_Limpa_Tela")

        End If

        giClienteAlterado = 0
        
        If Not gbCarregandoTela Then

    ''*** incluidos p/tratamento de tributacao *******************************
            If iCodFilial <> 0 Then
            
                Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, iCodFilial) '####
                
            End If
    '*** fim tributacao
        End If

        'Coloca clientes em atraso em vermelho
        If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)
        
    End If

    Exit Sub

Erro_Cliente_Validate:
   
    Cancel = True
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 26135, 26136, 26514, 51034, 51377, 51378, 131959, 181160, 181161, 181162, 207413 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157953)

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
    If lErro <> SUCESSO Then gError 23581

    Exit Sub

Erro_Filial_Click:

    Select Case gErr

        Case 23581

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157954)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 26622

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 26623

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 26624
        
        If lErro = 17660 Then
            
            'Lê o Cliente
            objCliente.sNomeReduzido = sCliente
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 25664
            
            'Não encontrou Cliente
            If lErro = 12348 Then gError 25665
            
            objFilialCliente.lCodCliente = objCliente.lCodigo
        
            gError 26625
            
        End If

        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        lErro = Trata_FilialCliente
        If lErro <> SUCESSO Then gError 25435

    End If

    'Não encontrou a STRING
    If lErro = 6731 Then gError 26626

    giFilialAlterada = 0
    
    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case gErr

        Case 25435, 25664, 25665 'tratado na rotina chamada

        Case 26622, 26624

        Case 26623
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 26625
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)
            
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 26626
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157955)

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
    If lErro <> SUCESSO Then gError 26582

'    If gobjPedidoVenda.dtDataEmissao <> StrParaDate(DataEmissao.Text) Then
'
'        gobjPedidoVenda.dtDataEmissao = StrParaDate(DataEmissao.Text)
'
'        Call ValorTotal_Calcula
'
'    End If
    
''*** incluidos p/tratamento de tributacao *******************************
        Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
'*** fim tributacao
    
    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 26582

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157956)

    End Select

    Exit Sub

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

Public Sub PlacaUF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PlacaUF_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(PlacaUF.Text)) = 0 Then Exit Sub

    'Verifica se existe o item na combo
    lErro = Combo_Item_Igual_CI(PlacaUF)
    If lErro <> SUCESSO And lErro <> 58583 Then gError 46527

    'Se não encontrar --> Erro
    If lErro = 58583 Then gError 46528

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True


    Select Case gErr

        Case 46527

        Case 46528
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157957)

    End Select

    Exit Sub

End Sub

Public Sub UnidadeMed_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro Then gError 26583

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 26583

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157961)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro Then gError 26584

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 26584

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157962)

    End Select

    Exit Sub

End Sub

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
    If lErro <> SUCESSO Then gError 30525

    Exit Sub

Erro_TabelaPreco_Click:

    Select Case gErr

        Case 30525

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157964)

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
    If lErro <> SUCESSO And lErro <> 19205 Then gError 26718
    If lErro = 19205 Then gError 26720
    
    'Altera PercAcrescFin
    If Not gbCarregandoTela Then
'        If Len(Trim(PercAcrescFin.ClipText)) > 0 Then
    
        dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100
        If Abs(dPercAcresFin - objCondicaoPagto.dAcrescimoFinanceiro) > DELTA_VALORMONETARIO2 Then
            vbMsgRes = vbYes
            'Se já tiverem itens cadastrados pergunta se deseja substituir o acrescimo financeiro porque
            'ao alterá-lo o preço unitário dos itens será recalculado de acordo com
            'a tabela de preço\acrescimo financeiro
            If objGridItens.iLinhasExistentes > 0 Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
            End If
            If vbMsgRes = vbYes Then
                PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
                Call PercAcrescFin_Validate(bSGECancelDummy)
            End If
        End If
'        Else
'            If objCondicaoPagto.dAcrescimoFinanceiro <> 0 Then
'                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
'                If vbMsgRes = vbYes Then
'                    PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
'                    Call PercAcrescFin_Validate(bSGECancelDummy)
'                End If
'            End If
'        End If
        
        lErro = CF2(Me, "PV_CondicaoPagamento_Click", objCondicaoPagto)
        If lErro <> SUCESSO Then gError 26718
    
        'Testa se ValorTotal está preenchido
        If Len(Trim(ValorTotal)) > 0 Then
            'Testa se DataReferencia está preenchida e ValorTotal é positivo
            If Len(Trim(DataReferencia.ClipText)) > 0 And (CDbl(ValorTotal.Caption) > 0) Then
    
                'Preenche o GridParcelas
                lErro = Cobranca_Automatica()
                If lErro <> SUCESSO Then gError 26719
    
            End If
        End If
    
    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_CondicaoPagamento_Click:

    Select Case gErr

        Case 26718, 26719

        Case 26720
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157965)

      End Select

    Exit Sub

End Sub

Public Function PV_CondicaoPagamento_Click(ByVal objCT As Object, ByVal objCondicaoPagto As ClassCondicaoPagto) As Long
    PV_CondicaoPagamento_Click = SUCESSO
End Function

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
        If FilialEntrega_TrataAlteracao() <> SUCESSO Then gError 27670
    End If

    Exit Sub

Erro_FilialEntrega_Click:

    Select Case gErr

        Case 27670

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157966)

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
Dim iColuna As Integer
Dim iKey As Integer
Dim iLinha As Integer

    iColuna = GridItens.Col
    iKey = KeyAscii
    iLinha = GridItens.Row

     Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

     Call CF2(Me, "PV_Muda_Tabulacao", iKey, iLinha, iColuna)

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
            If lErro <> SUCESSO Then gError 27641
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
                If lErro <> SUCESSO Then gError 101521
                
                Parent.HelpContextID = IDH_PEDIDO_VENDA_COMISSOES

                
            Case TAB_Bloqueio
                Parent.HelpContextID = IDH_PEDIDO_VENDA_BLOQUEIO
                
            Case TAB_Almoxarifado
                Parent.HelpContextID = IDH_PEDIDO_VENDA_ALMOXARIFADO
                
            Case TAB_Tributacao
                Parent.HelpContextID = IDH_PEDIDO_VENDA_TRIBUTACAO
                
        End Select
    
    
    End If

    Exit Sub

Erro_Opcao_Click:

    Select Case gErr

        Case 27641, 101521

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157967)

    End Select

    Exit Sub

End Sub

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
    
    'alterado por Tulio em 28/03/02
        lErro = objTabComissoes.Comissoes_Calcula()
        If lErro <> SUCESSO Then gError 25727
        
    End If
    
    Exit Sub

Erro_ComissaoAutomatica_Click:

    Select Case gErr

        Case 25727 'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157969)

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
    iComissoesAlterada = REGISTRO_ALTERADO
End Sub

Public Sub Produto_GotFocus()

Dim lErro As Long

    Call Grid_Campo_Recebe_Foco(objGridItens)

    If gobjFAT.iPedVendaCodBarra = 1 Then

        If objGridItens.lErroSaidaCelula = 0 Then

            lErro = CF2(Me, "PV_Trata_CodigoBarras1")

            Call Grid_Entrada_Celula(objGridItens, iAlterado)

            If lErro <> SUCESSO Then
    
                objGridItens.lErroSaidaCelula = 1
            End If

        Else
    
            objGridItens.lErroSaidaCelula = 0
    
        End If
        
    End If

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
    iComissoesAlterada = REGISTRO_ALTERADO
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
    iComissoesAlterada = REGISTRO_ALTERADO
End Sub

Public Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

    Call CF2(Me, "PV_Quantidade_GotFocus")

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
    iComissoesAlterada = REGISTRO_ALTERADO
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
    iComissoesAlterada = REGISTRO_ALTERADO
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

Public Sub DataEntrega_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataEntrega_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub DataEntrega_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub DataEntrega_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataEntrega
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

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long, iExecutaEntradaCelula As Integer

On Error GoTo Erro_Saida_Celula
    'aquii está devolvendo erro em vez de sucesso
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridParcelas.Name

                lErro = Saida_Celula_GridParcelas(objGridInt)
                If lErro <> SUCESSO Then gError 26064

            'Se for o GridItens
            Case GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 26065

            'se for o GridComissoes
            Case GridBloqueio.Name

                lErro = Saida_Celula_GridBloqueio(objGridInt)
                If lErro <> SUCESSO Then gError 26066

            'se for o GridComissoesEmissao
            Case GridComissoes.Name

                lErro = objTabComissoes.Saida_Celula_GridComissoes(objGridInt)
                If lErro <> SUCESSO Then gError 26067
            
            'Janaina
            'se for o Grid de Embalagens
            Case GridEmb.Name
                
                lErro = gobjEmbalagem.Saida_Celula_GridEmb()
                If lErro <> SUCESSO Then gError 96122
            'Janaina
            
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 26068
    
    End If
    
    If gbSaiuGrade Then

        If objGridInt.objGrid.Name = "GridItens" Then

            gbSaiuGrade = False

            If gobjFAT.iPulaLinhaGrade = MARCADO Then

                objGridInt.objForm.Show
                objGridInt.objGrid.SetFocus

                objGridInt.objGrid.Row = objGridInt.objGrid.Row + 1
'                objGridInt.objGrid.Col = objGridInt.objGrid.FixedCols


                SendKeys "{ENTER}", True
                SendKeys "{ENTER}", True


'                Call Grid_Trata_Tecla(vbKeyReturn, objGridInt, iExecutaEntradaCelula)
'
'                If iExecutaEntradaCelula = 1 Then
'                    Call Grid_Entrada_Celula(objGridInt, iAlterado)
'                End If

            End If

        End If

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    gbSaiuGrade = False

    Saida_Celula = gErr

    Select Case gErr

        Case 26064, 26065, 26066, 26067, 26068, 96122

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157970)

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
            If lErro <> SUCESSO Then gError 26593


        'Se for a de Unidade de Medida
        Case iGrid_UnidadeMed_Col
            lErro = Saida_Celula_UM(objGridInt)
            If lErro <> SUCESSO Then gError 26594

        'Se for a de Quantidade Pedida
        Case iGrid_QuantPedida_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 26595
    
        'Se for a de Quantidade Cancelada
        Case iGrid_QuantCancel_Col
            If gobjFAT.iExibeQtdeCanceladaItensPV = MARCADO Then
                lErro = Saida_Celula_QuantCancelada(objGridInt)
                If lErro <> SUCESSO Then gError 26598
            End If
            
        'Se for a de Preço Unitário
        Case iGrid_ValorUnitario_Col
            lErro = Saida_Celula_PrecoUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 26596
            
        'Se for a de Percentual de Desconto
        Case iGrid_PercDesc_Col
            lErro = Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then gError 26599

        Case iGrid_Desconto_Col
            lErro = Saida_Celula_Desconto(objGridInt)
            If lErro <> SUCESSO Then gError 26599

        'Se for a de Data de Entrega
        Case iGrid_DataEntrega_Col
            lErro = Saida_Celula_DataEntrega(objGridInt)
            If lErro <> SUCESSO Then gError 26601

        'Se for a de desconto produto
        Case iGrid_DescProduto_Col
            lErro = Saida_Celula_DescProduto(objGridInt)
            If lErro <> SUCESSO Then gError 87524

        'Se for a de desconto produto
        Case iGrid_PercentMaisReceb_Col
            If gobjFAT.iExibePercFatPV = MARCADO Then
                lErro = Saida_Celula_PercentMaisReceb(objGridInt)
                If lErro <> SUCESSO Then gError 87524
            End If

        'Se for a de desconto produto
        Case iGrid_PercentMenosReceb_Col
            If gobjFAT.iExibePercFatPV = MARCADO Then
                lErro = Saida_Celula_PercentMenosReceb(objGridInt)
                If lErro <> SUCESSO Then gError 87524
            End If

        'Se for a de desconto produto
        Case iGrid_RecebForaFaixa_Col
            If gobjFAT.iExibePercFatPV = MARCADO Then
                lErro = Saida_Celula_RecebForaFaixa(objGridInt)
                If lErro <> SUCESSO Then gError 87524
            End If

        Case iGrid_Prioridade_Col
            lErro = Saida_Celula_Prioridade(objGridInt)
            If lErro <> SUCESSO Then gError 87524
            
        Case iGrid_TabPrecoItem_Col
            If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
                lErro = Saida_Celula_TabPrecoItemPV(objGridInt)
                If lErro <> SUCESSO Then gError 87524
            End If

        Case iGrid_ComissaoItem_Col
            If gobjFAT.iExibeComissaoItensPV = MARCADO Then
                lErro = Saida_Celula_ComissaoItemPV(objGridInt)
                If lErro <> SUCESSO Then gError 87524
            End If

        Case Else
            lErro = CF2(Me, "PV_Saida_Celula_GridItens", objGridInt)
            If lErro <> SUCESSO Then gError 87524
    
    End Select
    
    
    
    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 26593, 26594, 26595, 26596, 26598, 26599, 26601, 87524

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157971)

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
    
        sProduto = Produto.Text

        lErro = CF("Trata_Segmento_Produto", sProduto)
        If lErro <> SUCESSO Then gError 199361

        Produto.Text = sProduto
    
        lErro = CF2(Me, "Produto_Saida_Celula")
        If lErro <> SUCESSO And lErro <> 26658 Then gError 26606
        If lErro = 26658 Then gError 51366
    End If
       
    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 26616
    
    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 26606, 26616, 199361
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 51366

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157972)

    End Select

    Exit Function

End Function

Public Function PV_Saida_Celula_UM(ByVal objCT As Object, ByVal sUmAnterior As String, bIgnora As Boolean) As Long
    PV_Saida_Celula_UM = SUCESSO
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
Dim objKitVenda As New ClassKitVenda
Dim bIgnora As Boolean

On Error GoTo Erro_Saida_Celula_UM

    Set objGridInt.objControle = UnidadeMed

    'Guarda a Unidade de medida anteriormente selecionada
    sUmAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

    'Coloca a Um no grid de itens
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 26627

    'Se a Um selecionada agora é diferente da anterior
    If sUmAnterior <> UnidadeMed.Text And objProduto.iKitVendaComp <> MARCADO Then
    
        objProduto.sCodigo = gobjPedidoVenda.colItensPedido(GridItens.Row).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177583
    
        '*********** IF INCLUÍDO PARA TRATAMENTO DE GRADE *********************
        If gobjPedidoVenda.colItensPedido(GridItens.Row).iPossuiGrade <> MARCADO And objProduto.iKitVendaComp <> MARCADO Then
            
            lErro = Atualiza_UM(GridItens.Row, sUmAnterior, UnidadeMed.Text)
            If lErro <> SUCESSO Then gError 65896
            
            bIgnora = False
            Call CF2(Me, "PV_Saida_Celula_UM", sUmAnterior, bIgnora)
            
            If Not bIgnora Then
                'Reprocessa a reserva do produto
                dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col))
                If iGrid_QuantCancel_Col <> 0 Then dQuantCancelada = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantCancel_Col))
                dQuantFaturada = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantFaturada_Col))
        
                lErro = Reserva_Processa(GridItens.Row, dQuantidade, dQuantCancelada, dQuantFaturada)
                If lErro <> SUCESSO Then gError 26852
            End If
        
        Else
            
            gobjPedidoVenda.colItensPedido(GridItens.Row).sUnidadeMed = UnidadeMed.Text
        
            lErro = Reserva_Trata_Grade_TrocaUM(gobjPedidoVenda.colItensPedido(GridItens.Row))
            If lErro <> SUCESSO Then gError 86332
            
            If objProduto.iKitVendaComp = MARCADO Then
            
                objKitVenda.sProduto = gobjPedidoVenda.colItensPedido(GridItens.Row).sProduto
                objKitVenda.sUM = UnidadeMed.Text
                objKitVenda.dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col)) - IIf(iGrid_QuantCancel_Col <> 0, StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantCancel_Col)), 0)
            
                lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjPedidoVenda.colItensPedido(GridItens.Row).colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 177584
                                
            End If
                                
        End If
    
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 500077
        
        lErro = CF2(Me, "PV_Calcula_PesoKg", GridItens.Row)
        If lErro <> SUCESSO Then gError 177161
    
    End If

    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case 26627, 86332, 500077, 65896, 177161, 177583, 177584
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26852
            GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = sUmAnterior
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157973)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Function Pesos_Recalcula() As Long
    Pesos_Recalcula = CF2(Me, "PV_Pesos_Recalcula")
End Function

Public Function PV_Pesos_Recalcula(ByVal objCT As Object) As Long
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

On Error GoTo Erro_PV_Pesos_Recalcula
    
    lErro = Calcula_VolumeTotal
    If lErro <> SUCESSO Then gError 96134
    
    dPesoBruto = 0
    dPesoLiquido = 0
    
    lQuantVolumes = 0
    bPesoBrutoInvalido = False
    
    'Para cada linha do Grid de Itens
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        'Se o produto e a UM estiverem preenchidos
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) > 0 Then
           
            dQuantPedida = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col))
            If iGrid_QuantCancel_Col <> 0 Then dQuantCancelada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))
            
            'Preenche o Grid de Embalagens
            lErro = gobjEmbalagem.Calcula_Item_GridEmb(dQuantPedida - dQuantCancelada, iIndice)
            If lErro <> SUCESSO Then gError 96125
                   
        End If
        
    Next
    
    lErro = gobjEmbalagem.Preenche_Pesos()
    If lErro <> SUCESSO Then gError 96134
    
    PV_Pesos_Recalcula = SUCESSO
    
    Exit Function

Erro_PV_Pesos_Recalcula:

    PV_Pesos_Recalcula = gErr
    
    Select Case gErr
    
        Case 500074, 500075, 500077, 85153, 96125, 96134
        
        Case 500076
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157974)
    
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
        If lErro <> SUCESSO And lErro <> 82763 Then gError 85153
        If lErro <> SUCESSO Then gError 81540
            
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
          
        Case 85153
        
        Case 81540
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objEmbalagem.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157975)
     
    End Select
     
    Exit Function
    
End Function

Function AsymUp(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
Dim Temp As Double
     Temp = Int(X * Factor)
     AsymUp = (Temp + IIf(X = Temp, 0, 1)) / Factor
End Function

Function Embalagem_Le(objEmbalagem As ClassEmbalagem) As Long
'Lê a embalagem a partir do código passado como parametro

Dim lErro As Long
Dim lComando As Long
Dim tEmbalagem As typeEmbalagem

On Error GoTo Erro_Embalagem_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 82760
    
    tEmbalagem.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)
    tEmbalagem.sSigla = String(STRING_DESCRICAO, 0)
    
    'Busca no BD uma embalagem com mesmo código
    lErro = Comando_Executar(lComando, "SELECT Codigo,Descricao, Sigla, Capacidade, Peso FROM Embalagens WHERE Codigo=?", tEmbalagem.iCodigo, tEmbalagem.sDescricao, tEmbalagem.sSigla, tEmbalagem.dCapacidade, tEmbalagem.dPeso, objEmbalagem.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 82761
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 82762
    
    'Se encontrou ==> erro
    If lErro <> AD_SQL_SUCESSO Then gError 82763
    
    objEmbalagem.dCapacidade = tEmbalagem.dCapacidade
    objEmbalagem.dPeso = tEmbalagem.dPeso
    objEmbalagem.iCodigo = tEmbalagem.iCodigo
    objEmbalagem.sDescricao = tEmbalagem.sDescricao
    objEmbalagem.sSigla = tEmbalagem.sSigla
    
    'fecha o comando
    Call Comando_Fechar(lComando)
    
    Embalagem_Le = SUCESSO
    
    Exit Function
    
Erro_Embalagem_Le:

    Embalagem_Le = gErr
    
    Select Case gErr
    
        Case 82760
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 82761, 82762
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMBALAGEM", gErr, objEmbalagem.iCodigo)
            
        Case 82763
            'Erro tratado na rotina chamadora
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157976)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function PV_Saida_Celula_Quantidade(ByVal objCT As Object, ByVal bQuantidadeIgual As Boolean, bIgnora As Boolean) As Long
    PV_Saida_Celula_Quantidade = SUCESSO
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
Dim objKitVenda As New ClassKitVenda
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim bIgnora As Boolean, bTrataKitVenda As Boolean

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    bQuantidadeIgual = False

    bTrataKitVenda = False

    If Len(Quantidade.Text) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 26665

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    'Comparação com quantidade anterior
    dQuantAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col))
    If dQuantAnterior = StrParaDbl(Quantidade.Text) Then bQuantidadeIgual = True
    
    GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col) = Formata_Estoque(dQuantidade)
    
    bIgnora = False
    Call CF2(Me, "PV_Saida_Celula_Quantidade", bQuantidadeIgual, bIgnora)

    If Not bQuantidadeIgual And Not bIgnora Then

        dQuantidade = StrParaDbl(Quantidade.Text)
        If iGrid_QuantCancel_Col <> 0 Then dQuantidadeCancelada = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantCancel_Col))
        dQuantidadeFaturada = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantFaturada_Col))
        If dQuantidadeCancelada > 0 And dQuantidade < dQuantidadeCancelada Then gError 26666
        If dQuantidadeFaturada > 0 And dQuantidade - dQuantidadeCancelada < dQuantidadeFaturada Then gError 26667

        objProduto.sCodigo = gobjPedidoVenda.colItensPedido(GridItens.Row).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177585
        
        If objProduto.iKitVendaComp <> MARCADO Then
            'Tenta fazer uma alocação p\ oitem de acordo com a nova UM
            lErro = Reserva_Processa(GridItens.Row, dQuantidade, dQuantidadeCancelada, dQuantidadeFaturada)
            If lErro <> SUCESSO Then gError 26831
    
        Else
        
            objKitVenda.sProduto = gobjPedidoVenda.colItensPedido(GridItens.Row).sProduto
            objKitVenda.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
            objKitVenda.dQuantidade = dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada
            
            gobjPedidoVenda.colItensPedido.Item(GridItens.Row).sUnidadeMed = objKitVenda.sUM
            gobjPedidoVenda.colItensPedido.Item(GridItens.Row).dQuantidade = dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada
        
            lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjPedidoVenda.colItensPedido(GridItens.Row).colItensRomaneioGrade)
            If lErro <> SUCESSO Then gError 177586
                        
            For Each objItemRomaneio In gobjPedidoVenda.colItensPedido(GridItens.Row).colItensRomaneioGrade
                
                lErro = Reserva_Processa_ItemGrade(objItemRomaneio, GridItens.Row)
                If lErro <> SUCESSO Then gError 177587
            
            Next
            
            Call Atualiza_Grid_Alocacao(gobjPedidoVenda.colItensPedido.Item(GridItens.Row))

            If gobjFAT.iKitVendaExibeComp = MARCADO Then bTrataKitVenda = True

        End If

    End If
    
    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 59727

    'Preço unitário
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))

    'Recalcula preço do ítem e valor total da nota
    If dPrecoUnitario > 0 And Not bQuantidadeIgual Then
        Call CF2(Me, "PV_PrecoTotal_Calcula", GridItens.Row)
    End If

    If Not bQuantidadeIgual Then
        
        If Not bPulaValorTotalCalcula Then
            'Recalcula o peso bruto e o peso líquido
            lErro = Pesos_Recalcula()
            If lErro <> SUCESSO Then gError 81525
        End If

        lErro = CF2(Me, "PV_Calcula_PesoKg", GridItens.Row)
        If lErro <> SUCESSO Then gError 177162

        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 51037

    End If
    
    If Not bPulaValorTotalCalcula Then
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 51037
    End If
    
    If bTrataKitVenda Then
        
        lErro = Trata_Kit_de_Venda(objProduto, objKitVenda, GridItens.Row)
        If lErro <> SUCESSO Then gError 51037
    
    End If
    
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 59727, 26665, 51037, 81525, 177162, 177585, 177586, 177587
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26831
            GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col) = Formata_Estoque(dQuantAnterior)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 26666
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_PEDIDA_INFERIOR_CANCELADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26667
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_FATURADA_SUPERIOR", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157977)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_PedidoVenda()

Dim lErro As Long
'Dim sNatureza As String

On Error GoTo Erro_Limpa_Tela_PedidoVenda

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

'    sNatureza = gobjTribTab.sNatOpInterna

    Call Limpa_Tela_PedidoVenda2

'    gobjTribTab.sNatOpInterna = sNatureza
    Call gobjTribTab.TributacaoRecalcular_Click
    
    iAlterado = 0
    iComissoesAlterada = 0
    gbCarregandoTela = False
    
    Exit Sub

Erro_Limpa_Tela_PedidoVenda:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157978)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 26538

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 26539

        If lErro <> SUCESSO Then gError 26540 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 30527

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 26541

    Exit Sub

Erro_TabelaPreco_Validate:

    Cancel = True

    Select Case gErr

    Case 26538, 26539, 30527

    Case 26540  'Não encontrou Tabela de Preço no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")

        If vbMsgRes = vbYes Then
            'Preenche o objTabela com o Codigo
            If Len(Trim(TabelaPreco.Text)) > 0 Then objTabelaPreco.iCodigo = CInt(TabelaPreco.Text)
            'Chama a tela de Tabelas de Preço
            Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
        Else
            'Segura o foco

        End If

    Case 26541

        Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)

    Case Else

        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157979)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 26542

    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 26543
        If lErro = 19205 Then gError 26545

        'Testa se pode ser usada em Contas a Receber
        If objCondicaoPagto.iEmRecebimento = 0 Then gError 26547

        'Coloca na Tela
        CondicaoPagamento.Text = CondPagto_Traz(objCondicaoPagto)
        
'        'Altera PercAcrescFin
'        If Len(Trim(PercAcrescFin.ClipText)) > 0 Then
'            dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100
'            If dPercAcresFin <> objCondicaoPagto.dAcrescimoFinanceiro Then
'                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
'                If vbMsgRes = vbYes Then
'                    PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
'                    Call PercAcrescFin_Validate(bSGECancelDummy)
'                End If
'            End If
'        Else
'            PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
'            Call PercAcrescFin_Validate(bSGECancelDummy)
'        End If

        dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100
        If Abs(dPercAcresFin - objCondicaoPagto.dAcrescimoFinanceiro) > DELTA_VALORMONETARIO2 Then
            vbMsgRes = vbYes
            'Se já tiverem itens cadastrados pergunta se deseja substituir o acrescimo financeiro porque
            'ao alterá-lo o preço unitário dos itens será recalculado de acordo com
            'a tabela de preço\acrescimo financeiro
            If objGridItens.iLinhasExistentes > 0 Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
            End If
            If vbMsgRes = vbYes Then
                PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
                Call PercAcrescFin_Validate(bSGECancelDummy)
            End If
        End If
        
        lErro = CF2(Me, "PV_CondicaoPagamento_Click", objCondicaoPagto)
        If lErro <> SUCESSO Then gError 26544

        'Se ValorTotal e DataReferencia estiverem preenchidos, preenche GridParcelas
        If Len(Trim(ValorTotal)) > 0 Then
            If Len(Trim(DataReferencia.ClipText)) > 0 And CLng(ValorTotal.Caption) > 0 Then

                'Preenche o GridParcelas
                lErro = Cobranca_Automatica()
                If lErro <> SUCESSO Then gError 26544

            End If
        End If

    End If

    'Não encontrou o valor que era STRING
    If lErro = 6731 Then gError 26546

    Exit Sub

Erro_Condicaopagamento_Validate:

    Cancel = True

    Select Case gErr

       Case 26542, 26543, 26544

       Case 26545
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If
            
        Case 26546
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)

        Case 26547
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", gErr, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157980)

    End Select

    Exit Sub

End Sub

Public Function PrecoUnitario_Calcula(sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long
    PrecoUnitario_Calcula = CF2(Me, "PV_PrecoUnitario_Calcula", 1, sUM, objTabelaPrecoItem, dPrecoUnitario)
End Function

Public Function PV_PrecoUnitario_Calcula(ByVal objCT As Object, ByVal iLinha As Integer, ByVal sUM As String, ByVal objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long
'Calcula o Preço unitário do item de acordo com a UM e a tabela de preço

Dim objProduto As New ClassProduto
Dim objUM As New ClassUnidadeDeMedida
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim lErro As Long, objTabelaPreco As New ClassTabelaPreco
Dim dPercAcresFin As Double
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objCotacao As New ClassCotacaoMoeda
Dim objCotacaoAnterior As New ClassCotacaoMoeda

On Error GoTo Erro_PV_PrecoUnitario_Calcula

    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 26638

    If lErro = 28030 Then gError 26639
    'Converte a quantidade para a UM de Venda
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 26640

    dPrecoUnitario = objTabelaPrecoItem.dPreco / dFator

    'Recolhe o percentual de acréscimo financeiro
    dPercAcresFin = StrParaDbl(PercAcrescFin.Text) / 100

    'Calcula o Preço unitário
    If dPercAcresFin <> 0 Then
        dPrecoUnitario = dPrecoUnitario * (1 + dPercAcresFin)
    End If

    If objTabelaPrecoItem.iCodTabela <> 0 Then
    
        objTabelaPreco.iCodigo = objTabelaPrecoItem.iCodTabela
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 26638
        If lErro = SUCESSO And objTabelaPreco.iMoeda <> MOEDA_REAL Then
            
            objCotacao.dtData = gdtDataAtual
            objCotacao.iMoeda = objTabelaPreco.iMoeda
            objCotacaoAnterior.iMoeda = objTabelaPreco.iMoeda
        
            'Chama função de leitura
            lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
            If lErro <> SUCESSO Then gError 26638
            
            'Se nao existe cotacao para a data informada => Mostra a última.
            dPrecoUnitario = dPrecoUnitario * StrParaDbl(IIf(objCotacao.dValor <> 0, Format(objCotacao.dValor, "#.0000"), Format(objCotacaoAnterior.dValor, "#.0000")))
        
        End If
    
    End If
    
    PV_PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_PV_PrecoUnitario_Calcula:

    PV_PrecoUnitario_Calcula = gErr

    Select Case gErr

        Case 26638, 26640

        Case 26639
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objTabelaPrecoItem.sCodProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157981)

    End Select

    Exit Function

End Function

Function ValorTotal_Calcula() As Long
'Calcula o Valor Total do Pedido

Dim dValorTotal As Double
Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    lErro = gobjTribTab.ValorTotal_Calcula(dValorTotal)
    If lErro <> SUCESSO Then gError 27648

    ValorTotal.Caption = Format(dValorTotal, "Standard")
    
    Call ValorDescontoTit_Validate(bSGECancelDummy)

    Call ValorDescontoItens_Calcula

    If Not gbCarregandoTela Then
    
        'Faz o cálculo automático das comissões
        lErro = Cobranca_Automatica()
        If lErro <> SUCESSO Then gError 56912
        
        'Se a empresa nãu utiliza as regras para o cálculo de comissões
        If Not gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
        
            'Significa que é necessário recalcular as comissões, utilizando o cálculo padrão
            lErro = objTabComissoes.Comissoes_Calcula_Padrao()
            If lErro <> SUCESSO Then gError 26643
        
        Else
            iComissoesAlterada = 1 'preventivamente, pois pode ter alterado algo que afete uma regra
        End If

    End If
    
    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 26643, 27648, 56912, 56914 'tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157982)

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


    lErro = CF2(Me, "PV_Rotina_Grid_Enable1", iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 215033

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 215033

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 215034)

    End Select

    Exit Sub

End Sub

Function PV_Rotina_Grid_Enable1(ByVal objCT As Object, iLinha As Integer, objControl As Object, iCaminho As Integer) As Long
    
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

On Error GoTo Erro_PV_Rotina_Grid_Enable1
    
    'Formata o produto do grid de itens
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 31389

    Select Case objControl.Name
        
        Case Produto.Name
            'Se o produto estiver preenchido desabilita
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
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
                If lErro <> SUCESSO And lErro <> 28030 Then gError 26644

                If lErro = 28030 Then gError 26645

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lê as UMs do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 26646
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

        Case TabPrecoItemPV.Name, ComissaoItemPV.Name, PercentDesc.Name, Desconto.Name, DataEntrega.Name, PercentMaisReceb.Name, PercentMenosReceb.Name, RecebForaFaixa.Name, Prioridade.Name
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case PrecoUnitario.Name
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Or gobjFAT.iBloqEditPrecoPV = MARCADO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
        Case DescricaoProduto.Name
                objControl.Enabled = True
                
        '******** O TRATAMENTO DE QUANTIDADE E QUANTCANCELADA FOI DESTACADO PARA TRATAR GRADE
        Case Quantidade.Name, QuantCancelada.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Or left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case TipoBloqueio.Name, ResponsavelBL.Name
            'Se for algum Bloqueio previamente cadastrado, desabilita o controle
            If iLinha <= giGridBloqueioLinhasBloqueadas Then
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

        Case TipoDesconto2.Name, TipoDesconto3.Name
            'Habilita os campos de desconto em sequencia
            If Len(Trim(GridParcelas.TextMatrix(iLinha, GridParcelas.Col - 4))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
        Case QuantFatAMais.Name
            objControl.Enabled = False
        
        Case Else
            lErro = CF2(Me, "PV_Rotina_Grid_Enable", iLinha, objControl, iCaminho, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 124178
    
    End Select

    'embalagem
    lErro = gobjEmbalagem.Rotina_Grid_Enable_Emb(iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 96127
    
    lErro = CF2(Me, "PV_Rotina_Grid_Enable2", iLinha, objControl, iCaminho, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 124178
       
    PV_Rotina_Grid_Enable1 = SUCESSO
    
    Exit Function

Erro_PV_Rotina_Grid_Enable1:

    PV_Rotina_Grid_Enable1 = gErr

    Select Case gErr

        Case 26644, 26646, 31389, 96127, 124178

        Case 26645
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157983)

    End Select

    Exit Function

End Function

Public Function PrecoTotal_Calcula(ByVal iLinha As Integer) As Long
    PrecoTotal_Calcula = CF2(Me, "PV_PrecoTotal_Calcula", iLinha)
End Function

Public Function PV_PrecoTotal_Calcula(ByVal objCT As Object, ByVal iLinha As Integer) As Long

'Dim dPrecoTotal As Double
'Dim dPrecoTotalReal As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dQuantidadeCancelada As Double
'Dim dDesconto As Double
Dim dPercentDesc As Double
Dim lTamanho As Long
Dim dValorTotal As Double, dValorTotalB As Double, dValorDescontoItens As Double
Dim iIndice As Integer
Dim dValorDesconto As Double
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long, objItemPV As ClassItemPedido

On Error GoTo Erro_PV_PrecoTotal_Calcula

    'Quantidades e preço unitário
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))
    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_QuantPedida_Col))
    If iGrid_QuantCancel_Col <> 0 Then dQuantidadeCancelada = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_QuantCancel_Col))

    lErro = CF2(Me, "PV_PrecoTotal_Calcula2", iLinha, dQuantidade)
    If lErro <> SUCESSO Then gError 177050
    
    'Cálculo do desconto
    lTamanho = Len(Trim(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col)))
    If lTamanho > 0 Then
        dPercentDesc = CDbl(Format(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col), "General Number"))
    Else
        dPercentDesc = 0
    End If

    lErro = CF2(Me, "PV_PrecoTotal_Calcula3", iLinha, (dQuantidade - dQuantidadeCancelada), dPrecoUnitario, dPercentDesc)
    If lErro <> SUCESSO Then gError 177050

    'Calcula a soma dos valores de produtos
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        If gobjPedidoVenda.colItensPedido.Count >= iIndice Then
        
            Set objItemPV = gobjPedidoVenda.colItensPedido(iIndice)
            
            If objItemPV.objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
        
                If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
                    If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
                End If
                dValorTotalB = dValorTotalB + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
                dValorDescontoItens = dValorDescontoItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
        
            End If
        
        End If
        
    Next

'    dValorDesconto = gobjPedidoVenda.dValorDesconto
'    dValorTotal = dValorTotal - dValorDesconto
'
'    'Verifica se o valor de desconto é maior que o valor dos produtos
'    If dValorTotal < 0 And dValorDesconto > 0 Then
'
'        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", dValorDesconto, dValorTotal)
'
'        gdDesconto = 0
'        Call CF2(Me, "NFCF_Limpa_Desconto")
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
'        If lErro <> SUCESSO Then gError 56913
'
'        'Calcula a soma dos valores de produtos
'        dValorTotal = 0
'        For iIndice = 1 To objGridItens.iLinhasExistentes
'            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
'                If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
'            End If
'        Next
'
'    End If

    'Coloca valor total dos produtos na tela
    ValorProdutos.Caption = Format(dValorTotal, "Standard")
    ValorProdutos2.Caption = Format(dValorTotalB, "Standard")
    ValorDesconto.Text = Format(dValorDesconto, "Standard")
    ValorDescontoItens.Text = Format(dValorDescontoItens, "Standard")
    dValorDescontoItensAnt = dValorDescontoItens
    If dValorTotalB > 0 Then
        PercDescontoItens.Text = Format(dValorDescontoItens / dValorTotalB * 100, "FIXED")
        dPercDescontoItensAnt = Arredonda_Moeda(dValorDescontoItens / dValorTotalB, 4)
    Else
        PercDescontoItens.Text = Format(0, "FIXED")
        dPercDescontoItensAnt = 0
    End If
    
    lErro = CF2(Me, "PV_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132005
    
    If iLinha <= gobjPedidoVenda.colItensPedido.Count Then
        lErro = gobjTribTab.Alteracao_Item_Grid(iLinha)
        If lErro <> SUCESSO Then gError 56913
    End If

    PV_PrecoTotal_Calcula = SUCESSO
    
    Exit Function

Erro_PV_PrecoTotal_Calcula:

    PV_PrecoTotal_Calcula = gErr
    
    Select Case gErr

        Case 56913, 132005 'tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157984)

    End Select

    Exit Function

End Function

Function Produto_Saida_Celula(ByVal objCT As Object) As Long

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
Dim objKit As New ClassKit
Dim sCodigoAntigo As String, dtDataAux As Date
Dim bPreenchePreco As Boolean
Dim lErro1 As Long
Dim objCliente As New ClassCliente
Dim iRecebForaFaixa As Integer
Dim iTemFaixaReceb As Integer
Dim dPercentMaisReceb As Double
Dim dPercentMenosReceb As Double
Dim lCliente As Long
Dim dPercDescontoTab As Double, dPrecoComDescontoTab As Double

On Error GoTo Erro_Produto_Saida_Celula

    bPreenchePreco = False

'***********  FUNÇÃO ALTERADA PARA TRATAMENTO DE GRADE  ******************
    'Critica o Produto
    lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 26657
    'Se o produto é gerencial sem ser de grade ==> Erro
    If lErro = 86295 And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then
        gError 86296
    ElseIf Len(Trim(objProduto.sGrade)) > 0 Then
        'Sinaliza que o produto possui grade
        iPossuiGrade = MARCADO
    End If
    
    If lErro = 51381 Then
        lErro1 = CF2(Me, "PV_Produto_Saida_Celula4", objProduto, lErro)
        If lErro1 <> SUCESSO Then gError 178535
    End If
    
    'Se o produto não foi encontrado ==> Pergunta se deseja criar
    If lErro = 51381 Then gError 26658
    
    lErro = CF2(Me, "PV_Produto_Saida_Celula3", objProduto)
    If lErro <> SUCESSO Then gError 177395

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
                For iIndice = 1 To gobjPedidoVenda.colItensPedido.Count
                    'Se aparecer ==> erro
                    If gobjPedidoVenda.colItensPedido(iIndice).sProduto = sProdutoPai Then gError 86328
                
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
            For iIndice = 1 To gobjPedidoVenda.colItensPedido.Count
                'Se aparecer ==> Erro
                If gobjPedidoVenda.colItensPedido(iIndice).sProduto = objItensRomaneio.sProduto Then gError 86331
            Next
        Next
 
    End If

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 26637

        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True

    End If

    If Len(Trim(PrioridadePadrao.ClipText)) > 0 Then
        Prioridade.PromptInclude = False
        Prioridade.Text = PrioridadePadrao.Text
        GridItens.TextMatrix(GridItens.Row, iGrid_Prioridade_Col) = PrioridadePadrao.Text
        Prioridade.PromptInclude = True
    End If

    If objProduto.iGerencial = DESMARCADO Then
        lErro = CF2(Me, "PV_Produto_Saida_Celula2")
        If lErro <> SUCESSO Then gError 132825
    End If
    
    'Verifica se é de Faturamento
    If objProduto.iFaturamento = 0 Then gError 26660
       
    Set objItemPV = New ClassItemPedido
    
    objItemPV.sProduto = objProduto.sCodigo
    objItemPV.sUnidadeMed = objProduto.sSiglaUMVenda
    objItemPV.lCodPedido = StrParaDbl(Codigo.Text)
    objItemPV.iItem = GridItens.Row
    objItemPV.lNumIntDoc = 0
    objItemPV.sDescricao = objProduto.sDescricao
    
    objItemPV.iPossuiGrade = iPossuiGrade

    If objItemPV.iPossuiGrade = MARCADO Then
                        
        Set objRomaneioGrade = New ClassRomaneioGrade
        
        objRomaneioGrade.iFilialFaturamento = Codigo_Extrai(FilialFaturamento.Text)
        objRomaneioGrade.sNomeTela = Me.Name
        
        Set objRomaneioGrade.objObjetoTela = objItemPV
                    
        lErro = GrupoGrade_Valida_Itens(objRomaneioGrade)
        If lErro <> SUCESSO Then gError 86310
                    
        Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
        If giRetornoTela <> vbOK Then gError 86310
        
        lErro = GrupoGrade_Trata_Itens(objRomaneioGrade, objProduto)
        If lErro <> SUCESSO Then gError 86310

        gbSaiuGrade = True
        
    End If

    'Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda

    'Descricao Produto
    If Me.ActiveControl Is DescricaoProduto Then
        DescricaoProduto.Text = objProduto.sDescricao
    Else
        GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objProduto.sDescricao
    End If
    
    GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(gdDesconto, "Percent")

    lErro = CF2(Me, "PV_Produto_Saida_Celula5", objProduto)
    If lErro <> SUCESSO Then gError 132825

    'Preço Unitário
    
    'Coloca Produto no grid (necessario p/usar Trata_TabelaPreco_Item)
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
            
    dPrecoUnitario = 0
    
    lErro = CF2(Me, "NFCF_PrecoUnitario_Calcula", giFilialEmpresa, objProduto.sCodigo, GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col), Cliente.Text, Codigo_Extrai(Filial.Text), Codigo_Extrai(TabelaPreco.Text), StrParaDbl(PercAcrescFin.Text) / 100, MOEDA_REAL, objRomaneioGrade, dPrecoUnitario, False, dPercDescontoTab, dPrecoComDescontoTab)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    If dPrecoUnitario > 0 Then
        GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
    Else
        GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = ""
    End If
    
    If dPercDescontoTab >= gdDesconto Then
        GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercDescontoTab, "Percent")
    End If
    
    lErro = CF2(Me, "PV_Produto_Saida_Celula", True, dPrecoUnitario)
    If lErro <> SUCESSO Then gError 132825

    'Acrescenta uma linha no Grid se for o caso
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1

        lErro = CF2(Me, "PV_Produto_Saida_Celula6")
        If lErro <> SUCESSO Then gError 197728

        gcolItemPedido.Add 0

        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItens.Row)

        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 27681
        
        If iPossuiGrade = MARCADO Or objProduto.iKitVendaComp = MARCADO Then
        
            '************** GRADE ************
            gobjPedidoVenda.colItensPedido(GridItens.Row).iPossuiGrade = MARCADO
            gobjPedidoVenda.colItensPedido(GridItens.Row).sUMEstoque = objProduto.sSiglaUMEstoque
            gobjPedidoVenda.colItensPedido(GridItens.Row).iItem = GridItens.Row
            gobjPedidoVenda.colItensPedido(GridItens.Row).sProduto = objProduto.sCodigo
                       
             If objProduto.iKitVendaComp = MARCADO Then
                Set objItemPV.colItensRomaneioGrade = colItensRomaneioGrade
            End If
            
            Set gobjPedidoVenda.colItensPedido(GridItens.Row).colItensRomaneioGrade = objItemPV.colItensRomaneioGrade
            
            If objProduto.iKitVendaComp <> MARCADO Then
            
                GridItens.TextMatrix(GridItens.Row, 0) = "# " & GridItens.TextMatrix(GridItens.Row, 0)
                       
                Call Atualiza_Grid_Itens(objItemPV)
                
            End If
            
            Call Atualiza_Grid_Alocacao(objItemPV)
            
        End If

    End If
   
    If gobjFAT.iExibePercFatPV = MARCADO Then
    
        'Verifica se o Cliente foi preenchido
        iTemFaixaReceb = gobjFAT.iTemFaixaReceb
        dPercentMaisReceb = gobjFAT.dPercentMaisReceb
        dPercentMenosReceb = gobjFAT.dPercentMenosReceb
        iRecebForaFaixa = gobjFAT.iRecebForaFaixa
        If Len(Trim(Cliente.ClipText)) > 0 Then
    
            objCliente.sNomeReduzido = Cliente.Text
    
            'Lê o Cliente através do Nome Reduzido
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 26779
    
            If objCliente.iIgnoraRecebPadrao Then
                iTemFaixaReceb = objCliente.iTemFaixaReceb
                dPercentMaisReceb = objCliente.dPercentMaisReceb
                dPercentMenosReceb = objCliente.dPercentMenosReceb
                iRecebForaFaixa = objCliente.iRecebForaFaixa
            End If
    
        End If
        
        If iTemFaixaReceb = DESMARCADO Then
            GridItens.TextMatrix(GridItens.Row, iGrid_PercentMaisReceb_Col) = Format(dPercentMaisReceb, "PERCENT")
            GridItens.TextMatrix(GridItens.Row, iGrid_PercentMenosReceb_Col) = Format(dPercentMenosReceb, "PERCENT")
    
            If iRecebForaFaixa = REJEITA_RECEBIMENTO Then
                'coloca no Grid Itens RecebForaFaixa
                GridItens.TextMatrix(GridItens.Row, iGrid_RecebForaFaixa_Col) = RecebForaFaixa.List(ERRO_E_REJEITA_RECEBIMENTO)
            Else
                'coloca no Grid Itens RecebForaFaixa
                GridItens.TextMatrix(GridItens.Row, iGrid_RecebForaFaixa_Col) = RecebForaFaixa.List(AVISA_E_ACEITA_RECEBIMENTO)
            End If
        Else
            'coloca no Grid Itens RecebForaFaixa
            GridItens.TextMatrix(GridItens.Row, iGrid_RecebForaFaixa_Col) = RecebForaFaixa.List(NAO_AVISA_E_ACEITA_RECEBIMENTO)
        End If
        
    End If
   
    Call Grid_Refresh_Checkbox(objGridItens)
   
    lErro = CF2(Me, "PV_Calcula_PesoKg", GridItens.Row)
    If lErro <> SUCESSO Then gError 177163
    
    If objItemPV.iPossuiGrade = MARCADO Then
        Call CF2(Me, "PV_PrecoTotal_Calcula", GridItens.Row)
        Call ValorTotal_Calcula
    End If
    
    lErro = GrupoGrade_Separa_Itens(objRomaneioGrade)
    If lErro <> SUCESSO Then gError 177163
      
    Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_Produto_Saida_Celula:

    Produto_Saida_Celula = gErr

    Select Case gErr

        Case 26637, 131013
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, Produto.Text)
                            
        Case 26657, 26661, 26662, 26663, 27681, 51420, 86310, 86327, 86329, 131295, 131296, 131297, 131298, 132825, 177163, 177395, 177582, 178535, 197728, ERRO_SEM_MENSAGEM

        Case 26658
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)
                
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
            End If
        
        Case 26660
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO", gErr)

        Case 51421
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO", gErr, objProduto.sCodigo)
        
        Case 86296
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case 86328
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_GRID", gErr, Trim(sProdutoPai), Trim(Produto.Text))
                    
        Case 86330
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, Produto.Text)
        
        Case 86331
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILHO_GRADE_GRID", gErr, Trim(objProduto.sCodigo), Trim(gobjPedidoVenda.colItensPedido(iIndice).sProduto))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157985)

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
Dim bPreenchePreco As Boolean
Dim objCliente As New ClassCliente, lCliente As Long
Dim dPercDescontoTab As Double, dPrecoComDescontoTab As Double

On Error GoTo Erro_Traz_Produto_Tela

    bPreenchePreco = False
    
    lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 55134
    If lErro = 51381 Then gError 55135
        
    lErro = CF2(Me, "PV_Produto_Saida_Celula2")
    If lErro <> SUCESSO Then gError 132825

    'Verifica se é de Faturamento
    If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 55137

    'Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda

    'Descricao Produto
    GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objProduto.sDescricao

    lErro = CF2(Me, "PV_Traz_Produto_Tela", objProduto)
    If lErro <> SUCESSO Then gError 55134
    
    'Preço Unitário
            
    'Coloca Produto no grid (necessario p/usar Trata_TabelaPreco_Item)
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
    
    dPrecoUnitario = 0
    
    lErro = CF2(Me, "NFCF_PrecoUnitario_Calcula", giFilialEmpresa, objProduto.sCodigo, GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col), Cliente.Text, Codigo_Extrai(Filial.Text), Codigo_Extrai(TabelaPreco.Text), StrParaDbl(PercAcrescFin.Text) / 100, MOEDA_REAL, Nothing, dPrecoUnitario, False, dPercDescontoTab, dPrecoComDescontoTab)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    If dPrecoUnitario > 0 Then
        GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
    Else
        GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = ""
    End If
    
    If dPercDescontoTab >= gdDesconto Then
        GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercDescontoTab, "Percent")
    End If
    
    lErro = CF2(Me, "PV_Produto_Saida_Celula", True, dPrecoUnitario)
    If lErro <> SUCESSO Then gError 132826

    'Acrescenta uma linha no Grid se for o caso
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1

        gcolItemPedido.Add 0
        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItens.Row)

        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 55141

    End If

    Call Grid_Refresh_Checkbox(objGridItens)
    
    Call CF2(Me, "PV_Traz_Produto_Tela2")

    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = gErr

    Select Case gErr

        Case 55134, 55138, 55139, 55140, 55141, 132826, ERRO_SEM_MENSAGEM

        Case 55135
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)

        Case 55136
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, Produto.Text, iIndice)

        Case 55137
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157986)

    End Select

    Exit Function

End Function

Public Function PV_Traz_Produto_Tela2(ByVal objCT As Object) As Long
    PV_Traz_Produto_Tela2 = SUCESSO
End Function

Public Sub ValorDesconto_Change()

    iAlterado = REGISTRO_ALTERADO
    giValorDescontoAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO

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
        If lErro <> SUCESSO Then gError 26652

        dValorDesconto = CDbl(ValorDesconto.Text)

        'Coloca o Valor formatado na tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")

        If dValorDesconto > dValorProdutos Then gError 26653
            
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
    If lErro <> SUCESSO Then gError 51038

    lErro = CF2(Me, "PV_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132002

    giValorDescontoAlterado = 0

    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True


    Select Case gErr

        Case 26652, 51038

        Case 26653
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos)

        Case 132002 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157987)

    End Select

    Exit Sub


End Sub

Public Sub ValorDespesas_Change()

     giValorDespesasAlterado = REGISTRO_ALTERADO
     iComissoesAlterada = REGISTRO_ALTERADO

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
        If lErro <> SUCESSO Then gError 35885

        dValorDespesas = CDbl(ValorDespesas.Text)

        'coloca o valor formatado na tela
        ValorDespesas.Text = Format(dValorDespesas, "Standard")

    End If

    'Para tributação
    gobjPedidoVenda.dValorOutrasDespesas = dValorDespesas

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDespesas_Validate(Cancel, dValorDespesas)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 51039

    giValorDespesasAlterado = 0

    Exit Sub

Erro_ValorDespesas_Validate:

    Cancel = True


    Select Case gErr

        Case 35885

        Case 51039

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157988)

    End Select

    Exit Sub


End Sub

Public Sub ValorFrete_Change()

    iAlterado = REGISTRO_ALTERADO
    giValorFreteAlterado = 1
    iComissoesAlterada = REGISTRO_ALTERADO

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
        If lErro <> SUCESSO Then gError 26650

        dValorFrete = CDbl(ValorFrete.Text)

        ValorFrete.Text = Format(dValorFrete, "Standard")

    End If

    'Para tributação
    gobjPedidoVenda.dValorFrete = dValorFrete
    
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorFrete_Validate(Cancel, dValorFrete)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 51040

    giValorFreteAlterado = 0

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True


    Select Case gErr

        Case 26650

        Case 51040

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157989)

    End Select

    Exit Sub

End Sub

Public Sub ValorProdutos_Change()

'    If Not gbCarregandoTela Then
'        Call DescontoGlobal_Recalcula
'    End If
'
End Sub

Private Sub DescontoGlobal_Recalcula()

Dim dValorDesconto As Double
Dim dValorProdutos As Double
Dim lErro As Long

On Error GoTo Erro_DescontoGlobal_Recalcula

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
'            lErro = CF2(Me, "PV_ValorDesconto_Validate")
'            If lErro <> SUCESSO Then gError 132003
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

        Case 132003

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157991)

    End Select

    Exit Sub
    
End Sub

Public Sub ValorSeguro_Change()

    iAlterado = REGISTRO_ALTERADO
    giValorSeguroAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO

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
        If lErro <> SUCESSO Then gError 26651

        dValorSeguro = CDbl(ValorSeguro.Text)

        ValorSeguro.Text = Format(dValorSeguro, "Standard")

    End If

    'Para tributação
    gobjPedidoVenda.dValorSeguro = dValorSeguro
    
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorSeguro_Validate(Cancel, dValorSeguro)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 51041

    giValorSeguroAlterado = 0

    Exit Sub

Erro_Valorseguro_Validate:

    Cancel = True


    Select Case gErr

        Case 26651

        Case 51041

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157992)

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

Dim dQuantidadeReservarEstoque As Double 'Reservar Filial 1

Dim dQuantidadeReservarEstoque1 As Double 'Reservar Filial 1
Dim dQuantidadeReservarEstoque2 As Double 'Reservar Filial 2
Dim objEstoqueProduto1 As New ClassEstoqueProduto 'Reservar Filial 1
Dim objEstoqueProduto2 As New ClassEstoqueProduto 'Reservar Filial 2


On Error GoTo Erro_ReservaAlmoxarifadoPadrao
    'calula a quantidade reservada
    dQuantidadeReservarVenda = (dQuantidade - dQuantCancelada - dQuantFaturada)
    
    lErro = CF("UM_Conversao", objProduto.iClasseUM, GridItens.TextMatrix(iIndiceGridItens, iGrid_UnidadeMed_Col), objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 26670
    'Converte a quantidade para a unidade de Estoque
    dQuantidadeReservarEstoque = dQuantidadeReservarVenda * dFator
    
    dQuantidadeReservarEstoque1 = dQuantidadeReservarEstoque
    
    If Len(Trim(FilialFaturamento.Text)) <> 0 Then
        iFilialFaturamento = Codigo_Extrai(FilialFaturamento.Text)
    End If
        
    If iFilialFaturamento = 0 Then
        'Busca a filial de faturamento
        lErro = CF("FilialFaturamento_Le", giFilialEmpresa, iFilialFaturamento)
        If lErro <> SUCESSO Then gError 30414
        If iFilialFaturamento = 0 Then iFilialFaturamento = giFilialEmpresa
    End If

    lErro = CF2(Me, "PV_AlocaAlmoxarifadoPadrao", iFilialFaturamento, objProduto.sCodigo, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO And lErro <> 35822 Then gError 141081
    
    objAlmoxarifadoPadrao.iCodigo = iAlmoxarifadoPadrao
    
    'Lê o almoxarifado padrão
    lErro = CF("Almoxarifado_Le", objAlmoxarifadoPadrao)
    If lErro = SUCESSO Then

        objEstoqueProduto1.sProduto = objProduto.sCodigo
        objEstoqueProduto1.iAlmoxarifado = iAlmoxarifadoPadrao
        objEstoqueProduto1.sAlmoxarifadoNomeReduzido = objAlmoxarifadoPadrao.sNomeReduzido
        
        'Lê o estoque do produto
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto1)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 26672
        If lErro = 21306 Then gError 26673

        If Len(Trim(Codigo.Text)) > 0 Then objItemPV.lCodPedido = CLng(Codigo.Text)
        objItemPV.sProduto = objProduto.sCodigo
        objItemPV.iFilialEmpresa = giFilialEmpresa
        objItemPV.sProdutoDescricao = objProduto.sDescricao
        objReservaBD.iAlmoxarifado = iAlmoxarifadoPadrao
        
        'Lê as reservas do produto
        lErro = CF("ReservaItem_Le", objItemPV, objReservaBD)
        If lErro <> SUCESSO And lErro <> 26678 Then gError 26674

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
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReservar_Col) = Formata_Estoque(dQuantidadeReservarEstoque) 'Formata_Estoque(dQuantidadeReservarVenda)
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReserv_Col) = Formata_Estoque(dQuantidadeReservarEstoque1)
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Responsavel_Col) = RESERVA_AUTO_RESP
            objGridAlocacao.iLinhasExistentes = objGridAlocacao.iLinhasExistentes + 1

        Else
        
            lErro = CF("Retorna_Almoxarifado_Alternativo", iAlmoxarifadoPadrao, iFilialEmpresa)
            If lErro <> SUCESSO Then gError 105004
        
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
                If lErro <> SUCESSO And lErro <> 21306 Then gError 26672
                If lErro = 21306 Then gError 26673
        
                If Len(Trim(Codigo.Text)) > 0 Then objItemPV.lCodPedido = CLng(Codigo.Text)
                objItemPV.sProduto = objProduto.sCodigo
                objItemPV.iFilialEmpresa = iFilialEmpresa
                objItemPV.sProdutoDescricao = objProduto.sDescricao
                objReservaBD.iAlmoxarifado = iAlmoxarifadoPadrao
                
                'Lê as reservas do produto
                lErro = CF("ReservaItem_Le", objItemPV, objReservaBD)
                If lErro <> SUCESSO And lErro <> 26678 Then gError 26674
        
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
                        GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReservar_Col) = Formata_Estoque(dQuantidadeReservarEstoque) 'Formata_Estoque(dQuantidadeReservarVenda)
                        GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReserv_Col) = Formata_Estoque(dQuantidadeReservarEstoque1)
                        GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Responsavel_Col) = RESERVA_AUTO_RESP
                        objGridAlocacao.iLinhasExistentes = objGridAlocacao.iLinhasExistentes + 1
                    End If
                    
                    'Parte Reservada na Filial 2
                    GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Item_Col) = iIndiceGridItens
                    GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_ProdutoAlmox_Col) = GridItens.TextMatrix(iIndiceGridItens, iGrid_Produto_Col)
                    GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_UMEstoque_Col) = objProduto.sSiglaUMEstoque
                    GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objEstoqueProduto2.sAlmoxarifadoNomeReduzido
                    GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReservar_Col) = Formata_Estoque(dQuantidadeReservarEstoque) 'Formata_Estoque(dQuantidadeReservarVenda)
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
                    
                    If gobjPedidoVenda.colItensPedido(objItemPV.iItem).iPossuiGrade = MARCADO Then objItemPV.iPossuiGrade = MARCADO
        
                    
                    'Armazena os nomes dos Produtos do Pedido para não fazer substituição com um que já tem
                    For iIndice = 1 To objGridItens.iLinhasExistentes
        
                        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
                        If lErro <> SUCESSO Then gError 26772
        
                        colItemPedido.Add 0, 0, 0, sProduto, 0, 0, 0, 0, DATA_NULA, "", 0, 0, 0, colReserva, "", "", "", "", 0, 0, 0, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
        
                    Next
                    
                    If iImportacao = MARCADO Then
                        ReservaAlmoxarifadoPadrao = SUCESSO
                        Exit Function
                    End If
                    
                    objItemPV.iTratamentoFaltaEstoque = iTratamentoPadraoFaltaEstoque
                    objItemPV.bTelaFaltaEstExibeBtnAplicarEmTodos = bPulaValorTotalCalcula
                    
                    'Chama tela de Falta de Estoque
                    If objItemPV.iTratamentoFaltaEstoque <> 0 Then
                        lErro = Chama_Tela_Modal_Invisivel("FaltaEstoque", objItemPV, colItemPedido, dQuantidadeReservarEstoque1, objAlmoxarifadoPadrao, objEstoqueProduto1.dSaldo)
                    Else
                        lErro = Chama_Tela_Modal("FaltaEstoque", objItemPV, colItemPedido, dQuantidadeReservarEstoque1, objAlmoxarifadoPadrao, objEstoqueProduto1.dSaldo)
                    End If
                    If objItemPV.iTratamentoFaltaEstoque <> 0 Then iTratamentoPadraoFaltaEstoque = objItemPV.iTratamentoFaltaEstoque
                  
                    'Se retornar Cancela erro
                    If giRetornoTela = vbCancel Then gError 26680
                            
                    If objItemPV.iPossuiGrade <> MARCADO Then
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
                                If iGrid_QuantCancel_Col <> 0 Then
                                    If objItemPV.dQuantCancelada > 0 Then
                                        GridItens.TextMatrix(iIndiceGridItens, iGrid_QuantCancel_Col) = Formata_Estoque(objItemPV.dQuantCancelada)
                                    Else
                                        GridItens.TextMatrix(iIndiceGridItens, iGrid_QuantCancel_Col) = ""
                                    End If
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
                                If lErro <> SUCESSO Then gError 26681
            
                                Produto.PromptInclude = False
                                Produto.Text = sProdutoEnxuto
                                Produto.PromptInclude = True
            
                                'Substitui Produto no GridItens
                                GridItens.TextMatrix(iIndiceGridItens, iGrid_Produto_Col) = Produto.Text
            
                                'Chama função que traz os dados defaults do Produto
                                lErro = Traz_Produto_Tela()
                                gError 26682  'Para forçar saída da função e segurar o foco
            
                            End If
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

        Case 26670, 26672, 26674, 26682, 26772, 30414, 105004, 141081 'Inserido por Wagner

        Case 26681
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 26673
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto1.sProduto, objEstoqueProduto1.iAlmoxarifado)

        Case 26680
            Call Rotina_Erro(vbOKOnly, "ERRO_RESERVA_NAO_DECIDIDA", gErr, objProduto.sCodigo)

        Case 51418
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO", gErr, objProduto.sCodigo, giFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157993)

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

'@@@@@ para ser customizada
Public Function PV_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long
    
    PV_Form_Unload = SUCESSO
    
End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Call CF2(Me, "PV_Form_Unload", Cancel)
    Set gobjInfoUsu = Nothing

    Set gobjTelaProjetoInfo = Nothing
    Set gobjInfoAdic = Nothing

    'variaveis auxiliares à contabilizacao
    Set gobjContabAutomatica = Nothing
    Set gobjNFiscalCTB = Nothing
    Set gobjPedidoVendaCTB = Nothing
    Set gcolAlmoxFilial = Nothing
    Set gobjGeracaoNFiscal = Nothing

    'Libera as variáveis globais
    giGridBloqueioLinhasBloqueadas = 0
    Set objEventoCliente = Nothing
    Set objEventoNumero = Nothing
    Set objEventoCondPagto = Nothing
    Set objEventoProduto = Nothing
    Set objEventoEstoqueProd = Nothing
    Set objEventoNaturezaOp = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoMensagem = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoTranspRedespacho = Nothing
    Set objEventoContato = Nothing

    Set objEventoOrcamento = Nothing

    Set objGridItens = Nothing
    Set objGridComissoes = Nothing
    Set objGridParcelas = Nothing
    Set objGridBloqueio = Nothing
    Set objGridAlocacao = Nothing

    Set gcolItemPedido = Nothing
        
    'Embalagem
    Set gobjEmbalagem = Nothing

    'Libera referencia para essa tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

    Set gobjAnotacao = Nothing
    
    Set objTabComissoes = Nothing
    Set gobjPedidoVenda = Nothing
    
    'Alterado por Tulio em 27/03
    Set objComissoesRegrasCalc = Nothing
    Set objMnemonicoComissCalc = Nothing
    Set objMnemonicoComissCalcAux = Nothing
    Set gobjTribTab = Nothing
    
    Set objUserControl = Nothing
    Set gcolParcPV = Nothing

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
        If lErro <> SUCESSO Then gError 26684

        PrecoUnitario.Text = Format(PrecoUnitario.Text, gobjFAT.sFormatoPrecoUnitario)
    
        lErro = CF2(Me, "PV_Saida_Celula_PrecoUnitario2", StrParaDbl(PrecoUnitario.Text), GridItens.Row)
        If lErro <> SUCESSO Then gError 197750
    
    End If

    'Comparação com Preço Unitário anterior
    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col)) = StrParaDbl(PrecoUnitario.Text) Then bPrecoUnitarioIgual = True

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col))) > 0 Then dQuantidade = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col))
    If iGrid_QuantCancel_Col <> 0 Then If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_QuantCancel_Col))) > 0 Then dQuantCancelada = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantCancel_Col))

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 26685

    If Not bPrecoUnitarioIgual Then
    
        lErro = CF2(Me, "PV_Saida_Celula_PrecoUnitario", StrParaDbl(PrecoUnitario.Text), GridItens.Row)
        If lErro <> SUCESSO Then gError 132830

        Call CF2(Me, "PV_PrecoTotal_Calcula", GridItens.Row)
        
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 51037

        If Not bPulaValorTotalCalcula Then
            lErro = ValorTotal_Calcula()
            If lErro <> SUCESSO Then gError 51042
        End If
        
    End If

    lErro = CF2(Me, "PV_Saida_Celula_PrecoUnitario1", StrParaDbl(PrecoUnitario.Text), GridItens.Row)
    If lErro <> SUCESSO Then gError 197750

    lErro = CF2(Me, "PV_Saida_Celula_PrecoUnitario3", bPrecoUnitarioIgual)
    If lErro <> SUCESSO Then gError 197750

    Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = gErr


    Select Case gErr

        Case 26684, 26685, 51042, 132830, 197750
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157994)

    End Select

    Exit Function

End Function

Public Function PV_Saida_Celula_PrecoUnitario3(ByVal objCT As Object, ByVal bPrecoUnitarioIgual As Boolean) As Long
    PV_Saida_Celula_PrecoUnitario3 = SUCESSO
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
Dim objKitVenda As New ClassKitVenda
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim bIgnora As Boolean

On Error GoTo Erro_Saida_Celula_QuantCancelada

    bQuantCanceladaIgual = False

    Set objGridInt.objControle = QuantCancelada

    If Len(QuantCancelada.Text) > 0 Then
    
        'Critica o valor da quantidade
        lErro = Valor_NaoNegativo_Critica(QuantCancelada.Text)
        If lErro <> SUCESSO Then gError 26686

        QuantCancelada.Text = Formata_Estoque(QuantCancelada.Text)

    End If

    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantCancel_Col)) = StrParaDbl(QuantCancelada.Text) Then bQuantCanceladaIgual = True
    
    bIgnora = False
    Call CF2(Me, "PV_Saida_Celula_Quantidade", bQuantCanceladaIgual, bIgnora)

    If Not bQuantCanceladaIgual And Not bIgnora Then

        dQuantCancelada = StrParaDbl(QuantCancelada.Text)

        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col))) > 0 Then dQuantidade = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col))
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_QuantFaturada_Col))) > 0 Then dQuantFaturada = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantFaturada_Col))

        If dQuantCancelada > 0 And dQuantidade < dQuantCancelada Then gError 26687
        If dQuantCancelada > 0 And (dQuantidade - dQuantCancelada < dQuantFaturada) Then gError 26688

        objProduto.sCodigo = gobjPedidoVenda.colItensPedido(GridItens.Row).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177585
        
        If objProduto.iKitVendaComp <> MARCADO Then
            'Tenta fazer uma alocação p\ oitem de acordo com a nova UM
            lErro = Reserva_Processa(GridItens.Row, dQuantidade, dQuantCancelada, dQuantFaturada)
            If lErro <> SUCESSO Then gError 26832
    
        Else
        
            objKitVenda.sProduto = gobjPedidoVenda.colItensPedido(GridItens.Row).sProduto
            objKitVenda.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
            objKitVenda.dQuantidade = dQuantidade - dQuantCancelada - dQuantFaturada
            
            gobjPedidoVenda.colItensPedido.Item(GridItens.Row).sUnidadeMed = objKitVenda.sUM
            gobjPedidoVenda.colItensPedido.Item(GridItens.Row).dQuantidade = dQuantidade - dQuantCancelada - dQuantFaturada
        
            lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjPedidoVenda.colItensPedido(GridItens.Row).colItensRomaneioGrade)
            If lErro <> SUCESSO Then gError 177586
                        
            For Each objItemRomaneio In gobjPedidoVenda.colItensPedido(GridItens.Row).colItensRomaneioGrade
                
                lErro = Reserva_Processa_ItemGrade(objItemRomaneio, GridItens.Row)
                If lErro <> SUCESSO Then gError 177587
            
            Next
            
            Call Atualiza_Grid_Alocacao(gobjPedidoVenda.colItensPedido.Item(GridItens.Row))
            
        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 26693
    
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) > 0 Then dPrecoUnitario = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))

    If Not bQuantCanceladaIgual And dPrecoUnitario > 0 Then
        'Calcula o preço total
        Call CF2(Me, "PV_PrecoTotal_Calcula", GridItens.Row)
    End If

    If Not bQuantCanceladaIgual Then
    
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81728
        
        lErro = CF2(Me, "PV_Calcula_PesoKg", GridItens.Row)
        If lErro <> SUCESSO Then gError 177162
        
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 51037
        
    End If
    
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 51043
    
    Saida_Celula_QuantCancelada = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantCancelada:

    Saida_Celula_QuantCancelada = gErr

    Select Case gErr

        Case 26686, 26693, 26832, 51043, 81728, 177585, 177586, 177587, 177162
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26687
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_PEDIDA_INFERIOR_CANCELADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26688
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANT_FATURADA_SUPERIOR", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157995)
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

On Error GoTo Erro_Saida_Celula_PercentDesc

    Set objGridInt.objControle = PercentDesc

    If Len(PercentDesc.Text) > 0 Then
        'Critica a porcentagem
        lErro = Porcentagem_Critica(PercentDesc.Text)
        If lErro <> SUCESSO Then gError 26694

        dPercentDesc = CDbl(PercentDesc.Text)
        
        If Format(dPercentDesc, "#0.#0\%") <> GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) Then
            'se for igual a 100% -> erro
            If dPercentDesc = 100 Then gError 26695

            PercentDesc.Text = Format(dPercentDesc, "Fixed")
        
        End If

    Else
        
        dDesconto = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col))
        dValorTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))

        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(dValorTotal + dDesconto, "Standard")

    End If

    sValorPercAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col)
    If sValorPercAnterior = "" Then sValorPercAnterior = "0,00%"
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 26696
    
    'Se foi alterada
    If Format(dPercentDesc, "#0.#0\%") <> sValorPercAnterior Then
    
        iDescontoAlterado = REGISTRO_ALTERADO
    
        'Recalcula o preço total
        Call CF2(Me, "PV_PrecoTotal_Calcula", GridItens.Row)
        
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 51044

        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 51044

    End If

    Saida_Celula_PercentDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDesc:

    Saida_Celula_PercentDesc = gErr

    Select Case gErr

        Case 26694, 26696, 51044
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26695
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157996)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dPrecoTotal As Double
Dim dDesconto As Double
Dim dPercentDesc As Double

On Error GoTo Erro_Saida_Celula_Desconto

    Set objGridInt.objControle = Desconto
    'Verifica se o o desconto foi digitado
    If Len(Trim(Desconto.ClipText)) > 0 Then
        
        'Critica o valor digitado
        lErro = Valor_NaoNegativo_Critica(Desconto.Text)
        If lErro <> SUCESSO Then gError 42219

        dDesconto = CDbl(Desconto.Text)
        
    End If
        
    If dDesconto <> StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col)) Then
        iDescontoAlterado = REGISTRO_ALTERADO
        dPrecoTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotalB_Col))
        'Se o Preço total e positivo
        If dPrecoTotal > 0 Then
            'Verifica se o Valor do desconto é superior ao Preço Total
            If dDesconto >= dPrecoTotal Then gError 42220
            
            'Recalcula o percentual de desconto
            dPercentDesc = dDesconto / dPrecoTotal

            GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42221

    'Recalcula o preço total
    Call CF2(Me, "PV_PrecoTotal_Calcula", GridItens.Row)

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 42222
'*** fim tributacao
    
    Call ValorTotal_Calcula
    
    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = gErr

    Select Case gErr

        Case 42219, 42221
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 42220
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCONTO_MAIOR_OU_IGUAL_PRECO_TOTAL", gErr, GridItens.Row, dDesconto, dPrecoTotal)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 42222

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157211)

    End Select

    Exit Function

End Function

Function Saida_Celula_DataEntrega(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Entrega que está deixando de ser a corrente

Dim lErro As Long
Dim dtDataEntrega As Date
Dim dtDataEmissao As Date

On Error GoTo Erro_Saida_Celula_DataEntrega

    Set objGridInt.objControle = DataEntrega

    If Len(Trim(DataEntrega.ClipText)) > 0 Then
        'Critica a Data informada
        lErro = Data_Critica(DataEntrega.Text)
        If lErro <> SUCESSO Then gError 26697
        'Se data de emissão estiver preenchida
        If Len(Trim(DataEmissao.ClipText)) > 0 Then

            dtDataEntrega = CDate(DataEntrega.Text)
            dtDataEmissao = CDate(DataEmissao.Text)
            'Veerifica se a data de emissão é maior que a data de entrega
            If dtDataEntrega < dtDataEmissao Then gError 26698

        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 26699

    Saida_Celula_DataEntrega = SUCESSO

    Exit Function

Erro_Saida_Celula_DataEntrega:

    Saida_Celula_DataEntrega = gErr

    Select Case gErr

        Case 26697, 26699
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26698
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAENTREGA", gErr, dtDataEntrega, dtDataEmissao)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157997)

    End Select

    Exit Function

End Function

Function Saida_Celula_DescProduto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Descrição do Produto que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescProduto

    Set objGridInt.objControle = DescricaoProduto
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 75488

    Saida_Celula_DescProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_DescProduto:

    Saida_Celula_DescProduto = gErr

    Select Case gErr

        Case 75488
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157998)

    End Select

    Exit Function

End Function

Private Function FilialFaturamento_TrataAlteracao() As Long
'Trata a alteracao da filial de faturamento selecionada para o pedido

Dim lErro As Long, objItemPV As ClassItemPedido, iFilFatAnterior As Integer
Dim objFilialEmpPedido As New AdmFiliais, objFilialEmpFat As New AdmFiliais

On Error GoTo Erro_FilialFaturamento_TrataAlteracao

    iFilFatAnterior = gobjPedidoVenda.iFilialEmpresaFaturamento
    If iFilFatAnterior = 0 Then iFilFatAnterior = giFilialEmpresa
    gobjPedidoVenda.iFilialEmpresaFaturamento = Codigo_Extrai(FilialFaturamento)
    
    If Not gbCarregandoTela And gobjPedidoVenda.iFilialEmpresaFaturamento <> iFilFatAnterior Then
    
        objFilialEmpPedido.iCodFilial = iFilFatAnterior
        lErro = CF("FilialEmpresa_Le", objFilialEmpPedido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objFilialEmpFat.iCodFilial = gobjPedidoVenda.iFilialEmpresaFaturamento
        lErro = CF("FilialEmpresa_Le", objFilialEmpFat)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objFilialEmpPedido.iSuperSimples <> objFilialEmpFat.iSuperSimples Then
        
            For Each objItemPV In gobjPedidoVenda.colItens
            
                objItemPV.objTributacao.iRegimeTributarioManual = VAR_PREENCH_AUTOMATICO
            
            Next
        
            Call gobjTribTab.TributacaoRecalcular_Click
        
        End If
            
    End If
    
    'Calula o valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 25441

    FilialFaturamento_TrataAlteracao = SUCESSO

    Exit Function

Erro_FilialFaturamento_TrataAlteracao:

    FilialFaturamento_TrataAlteracao = gErr

    Select Case gErr

        Case 25441

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157999)

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
    If lErro <> SUCESSO Then gError 27672

    FilialEntrega_TrataAlteracao = SUCESSO

    Exit Function

Erro_FilialEntrega_TrataAlteracao:

    FilialEntrega_TrataAlteracao = gErr

    Select Case gErr

        Case 27672

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158000)

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

        'Se não encontrou a Filial do Cliente --> erro
        If lErro = 17660 Then gError 26702

        'Coloca na tela
        FilialEntrega.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        If FilialEntrega_TrataAlteracao() <> SUCESSO Then gError 27671

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 26704

    giFilialEntregaAlterada = 0
    
    Exit Sub

Erro_FilialEntrega_Validate:

    Cancel = True


    Select Case gErr

        Case 26700, 26701, 27671

        Case 26702
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 26703
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 26704
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialEntrega.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158001)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 26705

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 26706
        If lErro <> SUCESSO Then gError 26707 'Não encontrou Transportadora no BD

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 26708

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case gErr

        Case 26705, 26706


        Case 26707  'Não encontrou Transportadora no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA", iCodigo)

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            End If
            'Segura o foco

        Case 26708
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158002)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 26709

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCanalVenda.iCodigo = iCodigo

        'Tenta ler CanalVenda com esse código no BD
        lErro = CF("CanalVenda_Le", objCanalVenda)
        If lErro <> SUCESSO And lErro <> 23597 Then gError 26710
        If lErro <> SUCESSO Then gError 26711 'Não encontrou CanalVenda no BD

        'Encontrou CanalVenda no BD, coloca no Text da Combo
        CanalVenda.Text = CStr(objCanalVenda.iCodigo) & SEPARADOR & objCanalVenda.sNomeReduzido

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 26712

    Exit Sub

Erro_CanalVenda_Validate:

    Cancel = True


    Select Case gErr

        Case 26709, 26710


        Case 26711  'Não encontrou CanalVenda no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CANALVENDA", iCodigo)

            If vbMsgRes = vbYes Then

                Call Chama_Tela("CanalDeVenda", objCanalVenda)

            End If
            
            'Segura o foco

        Case 26712
            Call Rotina_Erro(vbOKOnly, "ERRO_CANALVENDA_NAO_ENCONTRADO", gErr, CanalVenda.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158003)

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
        If lErro <> SUCESSO Then gError 26713

        'Compara com data de emissão
        If Len(Trim(DataEmissao.ClipText)) > 0 Then

            dtDataEmissao = CDate(DataEmissao.Text)
            dtDataReferencia = CDate(DataReferencia.Text)

            If dtDataEmissao > dtDataReferencia Then gError 26714

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

        Case 25436, 26713 'Tratado na rotina chamada

        Case 26714
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAREFERENCIA", gErr, dtDataReferencia, dtDataEmissao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158004)

    End Select

    Exit Sub

End Sub

Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto) As Long
'Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas

Dim lErro As Long
Dim dValorPagar As Double
Dim iIndice As Integer
Dim dPercAcrescFin As Double
Dim iTamanho As Integer
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim iColuna As Integer
Dim objParcPV As ClassParcelaPedidoVenda
Dim dValorPagarPrimParc As Double
Dim colDescontoPadrao As New Collection
Dim dValorParcST As Double, dtDataParcST As Date
Dim colParcAux As New Collection, bAdicionouST As Boolean
Dim objParcPVAux1 As ClassParcelaPedidoVenda, objParcPVAux2 As ClassParcelaPedidoVenda

On Error GoTo Erro_GridParcelas_Preenche

    'Limpa o GridParcelas
    Call Grid_Limpa(objGridParcelas)

    Set gcolParcPV = New Collection
    
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
    If lErro <> SUCESSO Then gError 26721

    'Se Valor a Pagar for positivo
    If dValorPagar > 0 Then

        dValorPagarPrimParc = 0
        
        If objCondicaoPagto.iNumeroParcelas <> 1 Then
        
            lErro = gobjTribTab.ValorPagarPrimParc_Calcula(dValorPagarPrimParc)
            If lErro <> SUCESSO Then gError 26721
    
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
        If lErro <> SUCESSO Then gError 26721

        If objGridParcelas.iLinhasExistentes <> objCondicaoPagto.colParcelas.Count Then objGridParcelas.iLinhasExistentes = objCondicaoPagto.colParcelas.Count

        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            
            If iIndice = 1 And dValorPagarPrimParc <> 0 Then objCondicaoPagto.colParcelas(iIndice).dValor = objCondicaoPagto.colParcelas(iIndice).dValor + dValorPagarPrimParc
            
            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
            
        Next

    End If

    'Se Data Referencia estiver preenchida
    If Len(Trim(DataReferencia.ClipText)) > 0 Then

        objCondicaoPagto.dtDataRef = StrParaDate(DataReferencia.Text)
        'objCondicaoPagto.dtDataEmissao = StrParaDate(DataEmissao.Text)
        objCondicaoPagto.dtDataEntrega = StrParaDate(DataEntregaPV.Text)

        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True, True, True)
        If lErro <> SUCESSO Then gError 26722

        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.colParcelas.Count 'objCondicaoPagto.iNumeroParcelas

            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dtVencimento, "dd/mm/yyyy")

        Next

    End If
    
    If gobjFAT.iParcRecCobraSTNovaParc = MARCADO And dValorParcST > 0 Then
        'Joga as parcelas calculadas para coleção e posiciona a nova parcela no lugar certo
        bAdicionouST = False
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            Set objParcPVAux1 = New ClassParcelaPedidoVenda
            objParcPVAux1.dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
            objParcPVAux1.dValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
            If objParcPVAux1.dtDataVencimento > dtDataParcST And Not bAdicionouST Then
                Set objParcPVAux2 = New ClassParcelaPedidoVenda
                objParcPVAux2.dtDataVencimento = dtDataParcST
                objParcPVAux2.dValor = dValorParcST
                colParcAux.Add objParcPVAux2
                bAdicionouST = True
            End If
            colParcAux.Add objParcPVAux1
        Next
        If Not bAdicionouST Then
            Set objParcPVAux2 = New ClassParcelaPedidoVenda
            objParcPVAux2.dtDataVencimento = dtDataParcST
            objParcPVAux2.dValor = dValorParcST
            colParcAux.Add objParcPVAux2
            bAdicionouST = True
        End If
        iIndice = 0
        For Each objParcPVAux1 In colParcAux
            iIndice = iIndice + 1
            If objParcPVAux1.dtDataVencimento <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objParcPVAux1.dtDataVencimento, "dd/mm/yyyy")
            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objParcPVAux1.dValor, "Standard")
        Next
        objGridParcelas.iLinhasExistentes = iIndice
    End If

    ' Se dValorPagar>0 coloca desconto padrao (quantos houver, se houver) em todas as parcelas.
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        lErro = Preenche_DescontoPadrao(iIndice)
        If lErro <> SUCESSO Then gError 51031
        
        Set objParcPV = New ClassParcelaPedidoVenda
        objParcPV.dtDataCredito = DATA_NULA
        
        gcolParcPV.Add objParcPV
        
    Next
    
    GridParcelas_Preenche = SUCESSO

    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = gErr

    Select Case gErr

        Case 26721, 26722, 51031

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158005)

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
        If lErro <> SUCESSO Then gError 26724

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
            If lErro <> SUCESSO Then gError 51061
        
        
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
    If lErro <> SUCESSO Then gError 26725

    Saida_Celula_ValorParcela = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorParcela:

    Saida_Celula_ValorParcela = gErr

    Select Case gErr

        Case 26724, 26725, 51061
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158006)

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
        If lErro <> SUCESSO Then gError 26726

         dtDataVencimento = CDate(DataVencimento.Text)

        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataReferencia.ClipText)) > 0 Then
            dtDataReferencia = CDate(DataReferencia.Text)
            If dtDataVencimento < dtDataReferencia Then gError 26728
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
    If lErro <> SUCESSO Then gError 26727

    If iCriouLinha Then
        'Coloca desconto padrao (le em CPRConfig)
        lErro = Preenche_DescontoPadrao(GridParcelas.Row)
        If lErro <> SUCESSO Then gError 51032
    End If
        
    Saida_Celula_DataVencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_DataVencimento:

    Saida_Celula_DataVencimento = gErr

    Select Case gErr

        Case 26726, 26727, 51032
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26728
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR_REFERENCIA", gErr, dtDataVencimento, GridParcelas.Row, dtDataReferencia)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158007)

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
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 26729

            'Não foi encontrado
            If lErro = 25085 Then gError 26730
            If lErro = 25086 Then gError 26731

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
    If lErro <> SUCESSO Then gError 26732

    Saida_Celula_TipoDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDesconto:

    Saida_Celula_TipoDesconto = gErr

    Select Case gErr

        Case 26729, 26732
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26730
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26731
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", gErr, objGridInt.objControle.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158008)

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
        If lErro <> SUCESSO Then gError 26733

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            Set objParcPV = New ClassParcelaPedidoVenda
            gcolParcPV.Add objParcPV
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 26734

    Saida_Celula_DescontoValor = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoValor:

    Saida_Celula_DescontoValor = gErr

    Select Case gErr

        Case 26733, 26734
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158009)

    End Select

    Exit Function

End Function


Public Sub BotaoVendedores_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoVendedores_Click
    
    lErro = objTabComissoes.BotaoVendedores_Click()
    If lErro <> SUCESSO Then gError 25734

    Exit Sub
    
Erro_BotaoVendedores_Click:

    Select Case gErr
    
        Case 25734  'Tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158010)
        
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
            If lErro <> SUCESSO Then gError 25728
        
        'senao
        Else
            'verificar se o vendedor esta no grid e tratar o campo 'DiretoIndireto'
            lErro = objTabComissoes.VendedorLinha_Preenche2(objVendedor)
            If lErro <> SUCESSO Then gError 101521
        
        End If
    
    End If

    Me.Show
    
    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr

        Case 25728, 101521 'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158011)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_TipoBloqueio(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo Bloqueio que está deixando de ser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_TipoBloqueio

    Set objGridInt.objControle = TipoBloqueio

    'Verifica se o Tipo foi preenchido
    If Len(Trim(TipoBloqueio.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoBloqueio.Text <> TipoBloqueio.List(TipoBloqueio.ListIndex) Then

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona_Grid(TipoBloqueio, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 26764

            'Não foi encontrado
            If lErro = 25085 Then gError 26766
            If lErro = 25086 Then gError 26767

        End If

        'Verifica se já existe um bloqueio, não liberado, do mesmo tipo no grid.
        For iIndice = 1 To objGridInt.iLinhasExistentes

            If iIndice <> GridBloqueio.Row Then
                iCodigo = Codigo_Extrai(GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))

                If iCodigo = TipoBloqueio.ItemData(TipoBloqueio.ListIndex) Then
                    If Len(Trim(GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col))) = 0 Then gError 51379
                End If
            End If

        Next
        
        'Acrescenta uma linha no Grid se for o caso
        If GridBloqueio.Row - GridBloqueio.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

            'Coloca data e usuário
            GridBloqueio.TextMatrix(GridBloqueio.Row, iGrid_DataBloqueio_Col) = Format(gdtDataHoje, "dd/mm/yy")
            GridBloqueio.TextMatrix(GridBloqueio.Row, iGrid_CodUsuario_Col) = Format(gsUsuario)
        
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 26765

    Saida_Celula_TipoBloqueio = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoBloqueio:

    Saida_Celula_TipoBloqueio = gErr

    Select Case gErr

        Case 26764, 26765
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26766
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26767
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO1", gErr, TipoBloqueio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 51379
            Call Rotina_Erro(vbOKOnly, "ERRO_BLOQUEIOPV_REPETIDO", gErr, Codigo_Extrai(TipoBloqueio.Text))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158012)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ResponsavelBL(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Responsável BL que está deixando de serr a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ResponsavelBL

    Set objGridInt.objControle = ResponsavelBL

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 26768

    Saida_Celula_ResponsavelBL = SUCESSO

    Exit Function

Erro_Saida_Celula_ResponsavelBL:

    Saida_Celula_ResponsavelBL = gErr

    Select Case gErr

        Case 26768
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158013)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Observação que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = Observacao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 125293
    
    Saida_Celula_Observacao = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = gErr
    
    Select Case gErr
    
        Case 125293
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158014)

    End Select

    Exit Function

End Function

Public Sub BotaoLiberaBloqueio_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoLiberaBloqueio_Click

    'Verifica se o Código foi preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then

        objPedidoVenda.lCodigo = CLng(Codigo.Text)
        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
        'Lê o Pedido de Venda
        lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
        If lErro <> SUCESSO And lErro <> 26509 Then gError 26769

        'Não encontrou o Pedido de Venda
        If lErro = 26509 Then objPedidoVenda.lCodigo = 0
        
    End If

    Call Chama_Tela_Modal("LiberaBloqueio", objPedidoVenda)
    
    'Le os Bloqueios
    lErro = CF("BloqueiosPV_Le", objPedidoVenda)
    If lErro <> SUCESSO Then gError 61485
    
    'Recarrega o tab de Bloqueios
    lErro = Carrega_Tab_Bloqueio(objPedidoVenda)
    If lErro <> SUCESSO Then gError 61486
    
    Exit Sub

Erro_BotaoLiberaBloqueio_Click:

    Select Case gErr

        Case 26769

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158015)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLibera_Click()

Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objItemGrade As ClassItemRomaneioGrade
Dim objItemPedido As ClassItemPedido

On Error GoTo Erro_BotaoLibera_Click

    If objGridAlocacao.iLinhasExistentes = 0 Then gError 60772

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_LIBERACAO_RESERVAS")
    If vbMsgRes = vbYes Then

        Call Grid_Limpa(objGridAlocacao)
        
        ProdutoDescricao.Caption = ""
        
        For iIndice = 1 To objGridItens.iLinhasExistentes
            GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col) = ""
        Next
    
    
    End If
    
    For Each objItemPedido In gobjPedidoVenda.colItensPedido
        If Not (objItemPedido.colItensRomaneioGrade Is Nothing) Then
            For Each objItemGrade In objItemPedido.colItensRomaneioGrade
                objItemGrade.dQuantReservada = 0
                Set objItemGrade.colLocalizacao = New Collection
            Next
        End If
        objItemPedido.dQuantReservada = 0
        Set objItemPedido.colReserva = New colReserva
    Next
    
    Exit Sub
    
Erro_BotaoLibera_Click:

    Select Case gErr

        Case 60772
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_HA_RESERVAS_PARA_LIBERAR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158016)

    End Select

    Exit Sub
    
End Sub

Public Function PV_BotaoReserva_Click(ByVal objCT As Object, ByVal objItemPedido As ClassItemPedido) As Long

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double

On Error GoTo Erro_PV_BotaoReserva_Click

    'Se a tabla de preços estiver preenchida
    If Len(Trim(TabelaPreco.Text)) > 0 Then

        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        objTabelaPrecoItem.sCodProduto = objItemPedido.sProduto
        'Lê a tabela de preço item para o produto em gifilialempresa
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 26778
        If lErro <> SUCESSO Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
            'Lê a tabela de preço item para o produto na empresa toda
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 28206
        End If
        'Se alguma tabela de preços foi encontrada
        If lErro <> 28014 Then
            'Calcula o preço unitário do produto
            lErro = CF2(Me, "PV_PrecoUnitario_Calcula", objItemPedido.iItem, objItemPedido.sUnidadeMed, objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 28207
            'Coloca o valor do preço unitário no grid
            If dPrecoUnitario > 0 Then
                GridItens.TextMatrix(objItemPedido.iItem, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            Else
                GridItens.TextMatrix(objItemPedido.iItem, iGrid_ValorUnitario_Col) = ""
            End If
            
            lErro = CF2(Me, "PV_Saida_Celula_PrecoUnitario", dPrecoUnitario, objItemPedido.iItem)
            If lErro <> SUCESSO Then gError 132827

        End If
    End If
    
    PV_BotaoReserva_Click = SUCESSO
    
    Exit Function
    
Erro_PV_BotaoReserva_Click:

    PV_BotaoReserva_Click = SUCESSO

    Select Case gErr
    
        Case 26778, 28206, 28207, 132827

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158016)

    End Select

    Exit Function
    
End Function

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
    
        Set colItensRomaneio = New Collection
        
        dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col))
        If iGrid_QuantCancel_Col <> 0 Then dQuantCancelada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))
        dQuantReservada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col))
        dQuantFaturada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantFaturada_Col))
        dQuantFaturar = dQuantidade - dQuantCancelada - dQuantFaturada
        
        
        '************* CRIADA A FUNÇÃO Produto_Critica_Filial2 PARA ADAPAR GRADE
        'Faz a critica do produto
        lErro = CF("Produto_Critica_Filial2", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), objProduto, iPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 28208
        If lErro = 51381 Then gError 86376
                
        'Se o produto é gerencial sem ser de grade ==> Erro
        If lErro = 86295 And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then
            gError 86377
        ElseIf lErro = 86295 Then
'************ TRATAMENTO DE GRADE ***********************
            
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
                If objProdutoAux.iControleEstoque = PRODUTO_CONTROLE_RESERVA Then
                    bAchou = False
                    For Each objItemRomaneioAux In gobjPedidoVenda.colItensPedido(iIndice).colItensRomaneioGrade
                        If objItemRomaneio.sProduto = objItemRomaneioAux.sProduto Then
                            Call Transfere_Dados_ItensRomaneio(colItensRomaneio, gobjPedidoVenda.colItensPedido(iIndice).colItensRomaneioGrade)
                            bAchou = True
                            Exit For
                        End If
                    Next
                End If
                If Not bAchou Or objProdutoAux.iControleEstoque <> PRODUTO_CONTROLE_RESERVA Then
                    colItensRomaneio.Remove iIndice2
                End If
            Next
'***********************************
        End If
    
        'se o produto faz o controle de reserva
        If objProduto.iControleEstoque = PRODUTO_CONTROLE_RESERVA Or gobjPedidoVenda.colItensPedido(iIndice).iPossuiGrade = MARCADO Or objProduto.iKitVendaComp = DESMARCADO Then

            Set colItemReserva = New colReserva
            'Recolhe as reservas daquele item
            Call Recolhe_Reserva(iIndice, colItemReserva)
            'adiciona os dados do item a coleção de itens
            colItemPedido.Add giFilialEmpresa, lNumPedido, iIndice, objProduto.sCodigo, dQuantidade, 0, 0, 0, DATA_NULA, objProduto.sDescricao, 0, dQuantCancelada, dQuantReservada, colItemReserva, objProduto.sNomeReduzido, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), "", objProduto.sSiglaUMEstoque, objProduto.iClasseUM, dQuantFaturada, dQuantFaturar, "", 0, objProduto.iControleEstoque, 0, 0, 0, 0, 0, 0, 0, 0, 0
            
            '**************** GRADE **********************
            If gobjPedidoVenda.colItensPedido(iIndice).iPossuiGrade = MARCADO Or objProduto.iKitVendaComp = MARCADO Then
                
                Set colItemPedido.Item(colItemPedido.Count).colItensRomaneioGrade = colItensRomaneio
                colItemPedido.Item(colItemPedido.Count).iPossuiGrade = MARCADO
                                
            End If

        End If
    Next
    'Chama a tela de alocação de produtos
    Call Chama_Tela_Modal("AlocacaoProduto", colItemPedido)
       
    'Limpa o grid de reserva
    Call Grid_Limpa(objGridAlocacao)
    'Para cada item da coleção de itens
    For Each objItemPedido In colItemPedido
        
        '************************* GRADE ************************
        If objItemPedido.iPossuiGrade = MARCADO Then
        
            For iIndice2 = 1 To gobjPedidoVenda.colItensPedido.Count
                If objItemPedido.sProduto = gobjPedidoVenda.colItensPedido.Item(iIndice2).sProduto Then
                    Set gobjPedidoVenda.colItensPedido.Item(iIndice2).colItensRomaneioGrade = objItemPedido.colItensRomaneioGrade
                    gobjPedidoVenda.colItensPedido.Item(iIndice2).dQuantCancelada = objItemPedido.dQuantCancelada
                    gobjPedidoVenda.colItensPedido.Item(iIndice2).dQuantidade = objItemPedido.dQuantidade
                    gobjPedidoVenda.colItensPedido.Item(iIndice2).dQuantReservada = objItemPedido.dQuantReservada
                End If
            Next
            
        Else
                    
            'Preenche no grid itens a quantidade reservada e a quantidade cancelada
            If objItemPedido.dQuantReservada > 0 Then
                GridItens.TextMatrix(objItemPedido.iItem, iGrid_QuantReservada_Col) = Formata_Estoque(objItemPedido.dQuantReservada)
            Else
                GridItens.TextMatrix(objItemPedido.iItem, iGrid_QuantReservada_Col) = ""
            End If
            If iGrid_QuantCancel_Col <> 0 Then
                If objItemPedido.dQuantCancelada > 0 Then
                    GridItens.TextMatrix(objItemPedido.iItem, iGrid_QuantCancel_Col) = Formata_Estoque(objItemPedido.dQuantCancelada)
                Else
                    GridItens.TextMatrix(objItemPedido.iItem, iGrid_QuantCancel_Col) = ""
                End If
            End If
            
            'Formata o produto
            lErro = CF("Produto_Formata", GridItens.TextMatrix(objItemPedido.iItem, iGrid_Produto_Col), sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 26776
            'Se o produto foi alterado
            If sProduto <> objItemPedido.sProduto Then
    
                lErro = Mascara_RetornaProdutoEnxuto(objItemPedido.sProduto, sProdutoEnxuto)
                If lErro <> SUCESSO Then gError 26777
    
                Produto.PromptInclude = False
                Produto.Text = sProdutoEnxuto
                Produto.PromptInclude = True
                'Preenche no grid de itens os dados do novo produto
                GridItens.TextMatrix(objItemPedido.iItem, iGrid_Produto_Col) = Produto.Text
                GridItens.TextMatrix(objItemPedido.iItem, iGrid_DescProduto_Col) = objItemPedido.sDescricao
                GridItens.TextMatrix(objItemPedido.iItem, iGrid_QuantPedida_Col) = Formata_Estoque(objItemPedido.dQuantidade)
                GridItens.TextMatrix(objItemPedido.iItem, iGrid_UnidadeMed_Col) = objItemPedido.sUnidadeMed

                lErro = CF2(Me, "PV_BotaoReserva_Click", objItemPedido)
                If lErro <> SUCESSO Then gError 26777
                
                'Calcula o preço total e o valor total
                Call CF2(Me, "PV_PrecoTotal_Calcula", objItemPedido.iItem)
    
            End If
        End If
    Next
        
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 51045
        
    'Preenche o grid de reservas
    Call Preenche_GridAlocacao(colItemPedido)

    Exit Sub

Erro_BotaoReserva_Click:

    Select Case gErr

        Case 26776, 26778, 28206, 28207, 28208, 51045, 86306, 86378, 132827, 177582 'Alterado por Wagner

        Case 26777
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, sProduto)
    
        Case 86376
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProduto)

        Case 86379
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objItemRomaneio.sProduto)

        Case 86377
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, sProduto)

        Case 86307
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158017)

    End Select

    Exit Sub

End Sub

Public Function Move_Pedido_Memoria(objPedidoVenda As ClassPedidoDeVenda, Optional ByVal bMoveTudo As Boolean = True) As Long
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
Dim objTabelaPreco As New ClassTabelaPreco

Dim iIndice As Integer

On Error GoTo Erro_Move_Pedido_Memoria

    If Len(Trim(Codigo.Text)) > 0 Then objPedidoVenda.lCodigo = CLng(Codigo.Text)

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 26779

        'Se não achou o Cliente --> erro
        If lErro = 12348 Then gError 26780

        'Guarda código do Cliente em objPedidoVenda
        objPedidoVenda.lCliente = objCliente.lCodigo

    End If

    'Verifica se a Filial está preenchida
    If Len(Trim(Filial.Text)) > 0 Then
        objPedidoVenda.iFilial = Codigo_Extrai(Filial.Text)
        'Se a função código extrai não conseguiu obter o código da filial => erro, pois a filial não está cadastrada
        'isso acontece quando se traz para a tela um orçamento que foi gerado com uma filial de cliente inexistente
        If objPedidoVenda.iFilial = 0 Then gError 102201 'Incluído por Luiz Nogueira em 29/05/03
    End If
   
   'Trecho Incluido por Leo em 25/04/02
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
        
        'Armazena o NumIntDoc do Orcamento em objPedidoVenda.lNumIntOrcVenda
        objPedidoVenda.lNumIntOrcVenda = objOrcamentoVenda.lNumIntDoc
         
    End If
    'Fim Leo
   
    'Preenche objPedidoVenda com dados da tela
    objPedidoVenda.iFilialEmpresaFaturamento = Codigo_Extrai(FilialFaturamento.Text)
    objPedidoVenda.dtDataEmissao = MaskedParaDate(DataEmissao)
    objPedidoVenda.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)
    
    objTabelaPreco.iCodigo = objPedidoVenda.iTabelaPreco
    
    If objTabelaPreco.iCodigo <> 0 Then

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError ERRO_SEM_MENSAGEM
        
        If objTabelaPreco.iAtivo = DESMARCADO Then gError 213834
        
    End If
    
    objPedidoVenda.sNaturezaOp = gobjTribTab.sNatOpInterna
    objPedidoVenda.iFaturaIntegral = FaturaIntegral.Value
    objPedidoVenda.dValorFrete = StrParaDbl(ValorFrete.Text)
    objPedidoVenda.dValorSeguro = StrParaDbl(ValorSeguro.Text)
    objPedidoVenda.dValorDesconto = StrParaDbl(ValorDesconto.Text)
    objPedidoVenda.dValorOutrasDespesas = StrParaDbl(ValorDespesas.Text)
    objPedidoVenda.dValorProdutos = StrParaDbl(ValorProdutos.Caption)
    objPedidoVenda.dValorTotal = StrParaDbl(ValorTotal.Caption)
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa
    
    If gobjFAT.iUsaParcDePV Then
        objPedidoVenda.iParc = StrParaInt(Parc.Text)
    End If
    
    If bMoveTudo Then

        'Move Grid Itens para memória
        lErro = Move_GridItens_Memoria(objPedidoVenda)
        If lErro <> SUCESSO Then gError 26781
    
        'Move Tab Complemento para memória
        Call Move_TabComplemento_Memoria(objPedidoVenda)
    
        'Move Tab Cobrança para memória
        Call Move_TabCobranca_Memoria(objPedidoVenda)
        'Alterado por Tulio em 28/03/02
        
        'Move Tab Comissões para memória
        lErro = objTabComissoes.Move_TabComissoes_Memoria(objPedidoVenda, PEDIDO_DE_VENDA)
    
        'Move Tab Bloqueio para memória
        Call Move_TabBloqueio_Memoria(objPedidoVenda)
    
        'Move Grid Reserva para memória
        Call Move_GridAlocacao_Memoria(objPedidoVenda)
    
        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
    
        'Move Tributacao para objPedidoVenda
        Set objPedidoVenda.objTributacaoPV = gobjPedidoVenda.objTributacaoPV
    
        'Janaina
        iTipoDocOrigem = DOCORIGEM_PV
        
        'embalagem
        lErro = gobjEmbalagem.Move_GridEmb_Memoria(colEmbalagensItens, iTipoDocOrigem)
        If lErro <> SUCESSO Then gError 89466
        
        If colEmbalagensItens.Count > 0 Then
        
            For iIndice = 1 To objPedidoVenda.colItensPedido.Count
                
                For Each objEmbalagensItem In colEmbalagensItens
    
                    If objEmbalagensItem.iItem = iIndice Then
                        Call objPedidoVenda.colItensPedido(iIndice).colEmbalagens.Add(objEmbalagensItem)
                    End If
                
                Next
            
            Next
            
        End If
        'Janaina
        
        lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_PV)
        If lErro <> SUCESSO Then gError 181761
        
        Set objPedidoVenda.objProjetoInfo = objProjetoInfo
        
        Set objPedidoVenda.objTela = Me
        
        Set objPedidoVenda.objInfoAdic = gobjInfoAdic
        objPedidoVenda.dValorItens = StrParaDbl(ValorProdutos2.Caption)
        objPedidoVenda.dValorDescontoItens = StrParaDbl(ValorDescontoItens.Text)
    
    End If
    
    Move_Pedido_Memoria = SUCESSO

    Exit Function

Erro_Move_Pedido_Memoria:

    Move_Pedido_Memoria = gErr

    Select Case gErr

        Case 26779, 26781, 94477, 181761

        Case 26780
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case 94476
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALORCAMENTO_NAO_INFORMADA", gErr)
            
        Case 94478
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
            
        Case 102201 'Incluído por Luiz Nogueira em 29/05/03
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)
            
        Case 102202 'Incluído por Luiz Nogueira em 29/05/03
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_INEXISTENTE", gErr, FilialEntrega.Text)
            
        Case ERRO_SEM_MENSAGEM
            
        Case 213834
            Call Rotina_Erro(vbOKOnly, "ERRO_TABPRECO_INATIVA", gErr, objTabelaPreco.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158018)

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
            If lErro <> SUCESSO Then gError 27682
            
        End If

        lErro = Move_GridItem_Memoria(objPedidoVenda, iIndice, sProduto)
        If lErro <> SUCESSO Then gError 26782

        '********************* TRATAMENTO DE GRADE *****************
        Call Move_ItensGrade_Tela(objPedidoVenda.colItensPedido(iIndice).colItensRomaneioGrade, gobjPedidoVenda.colItensPedido(iIndice).colItensRomaneioGrade)
            
    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 26782

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158019)

    End Select

    Exit Function

End Function

Private Sub Move_TabComplemento_Memoria(objPedidoVenda As ClassPedidoDeVenda)

    'Recolhe os dados de complemento
    objPedidoVenda.sPlaca = Placa.Text
    objPedidoVenda.sPlacaUF = PlacaUF.Text
    objPedidoVenda.sVolumeNumero = VolumeNumero.Text
    objPedidoVenda.lVolumeQuant = StrParaLong(VolumeQuant.Text)
    objPedidoVenda.dPesoBruto = StrParaDbl(PesoBruto.Text)
    objPedidoVenda.dPesoLiq = StrParaDbl(PesoLiquido.Text)
    objPedidoVenda.dVolumeTotal = StrParaDbl(Cubagem.Text)
    
    'Armazena o responsável pelo frete
    objPedidoVenda.iFreteRespons = TipoFrete.ItemData(TipoFrete.ListIndex)
    
    If Len(Trim(FilialEntrega.Text)) > 0 Then
        objPedidoVenda.iFilialEntrega = Codigo_Extrai(FilialEntrega.Text)
        
        'Se a função código extrai não conseguiu obter o código da filial => erro, pois a filial não está cadastrada
        'isso acontece quando se traz para a tela um orçamento que foi gerado com uma filial de cliente inexistente
        If objPedidoVenda.iFilialEntrega = 0 Then gError 102202 'Incluído por Luiz Nogueira em 29/05/03
    End If
    
    If Len(Trim(Transportadora.Text)) > 0 Then objPedidoVenda.iCodTransportadora = Codigo_Extrai(Transportadora.Text)
    If Len(Trim(TranspRedespacho.Text)) > 0 Then objPedidoVenda.iCodTranspRedesp = Codigo_Extrai(TranspRedespacho.Text)
    objPedidoVenda.iDetPagFrete = RedespachoCli.Value
    If Len(Trim(PedidoCliente.ClipText)) > 0 Then objPedidoVenda.sPedidoCliente = PedidoCliente.Text
    If Len(Trim(Mensagem.Text)) > 0 Then objPedidoVenda.sMensagemNota = Mensagem.Text
    If Len(Trim(CanalVenda.Text)) > 0 Then objPedidoVenda.iCanalVenda = Codigo_Extrai(CanalVenda.Text)

    'Incluído por Luiz Nogueira em 21/08/03
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objPedidoVenda.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objPedidoVenda.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
    
    objPedidoVenda.dtDataEntrega = MaskedParaDate(DataEntregaPV)

    If Len(Trim(PedidoRepr.ClipText)) > 0 Then objPedidoVenda.sPedidoRepresentante = PedidoRepr.Text

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
    objPedidoVenda.dtDataRefFluxo = MaskedParaDate(DataRefFluxo)
    objPedidoVenda.iCondicaoPagto = CondPagto_Extrai(CondicaoPagamento)
    
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

Private Sub Move_TabBloqueio_Memoria(objPedidoVenda As ClassPedidoDeVenda)
'Recolhe os dados do tab de Bloqueio

Dim lErro As Long
Dim objBloqueioPV As ClassBloqueioPV
Dim iIndice As Integer

    If objGridBloqueio.iLinhasExistentes = 0 Then Exit Sub

    'Recolhe os Bloqueios do Grid de Bloqueios
    For iIndice = 1 To objGridBloqueio.iLinhasExistentes

        Set objBloqueioPV = New ClassBloqueioPV

        objBloqueioPV.iTipoDeBloqueio = Codigo_Extrai(GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))

        If Len(Trim(GridBloqueio.TextMatrix(iIndice, iGrid_DataBloqueio_Col))) > 0 Then
            objBloqueioPV.dtData = CDate(GridBloqueio.TextMatrix(iIndice, iGrid_DataBloqueio_Col))
        Else
            objBloqueioPV.dtData = DATA_NULA
        End If

        objBloqueioPV.sCodUsuario = GridBloqueio.TextMatrix(iIndice, iGrid_CodUsuario_Col)
        objBloqueioPV.sResponsavel = GridBloqueio.TextMatrix(iIndice, iGrid_ResponsavelBL_Col)
        objBloqueioPV.sObservacao = GridBloqueio.TextMatrix(iIndice, iGrid_Observacao_Col)
        objBloqueioPV.iSequencial = StrParaInt(GridBloqueio.TextMatrix(iIndice, iGrid_SeqBloqueio_Col))

        If Len(Trim(GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col))) > 0 Then
            objBloqueioPV.dtDataLib = CDate(GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col))
        Else
            objBloqueioPV.dtDataLib = DATA_NULA
        End If

        objBloqueioPV.sResponsavelLib = GridBloqueio.TextMatrix(iIndice, iGrid_ResponsavelLib_Col)


        'Adiciona o bloqueio na coleção de bloqueios
        With objBloqueioPV
            objPedidoVenda.colBloqueiosPV.Add giFilialEmpresa, objPedidoVenda.lCodigo, objBloqueioPV.iSequencial, .iTipoDeBloqueio, .sCodUsuario, .sResponsavel, .dtData, "", .sResponsavelLib, .dtDataLib, .sObservacao
        End With

    Next

    Exit Sub

End Sub

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
            
            '******************** TRATAMENTO DE GRADE **********************
            'Só recolhe do grid as reservas dos itens que não são de grade
            If gobjPedidoVenda.colItensPedido(iItem).iPossuiGrade <> MARCADO Then
    
                Set objReservaItem = New ClassReservaItem
    
                objReservaItem.sAlmoxarifado = GridAlocacao.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
    
                'Lê o almoxarifado
                objAlmoxarifado.sNomeReduzido = objReservaItem.sAlmoxarifado
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25060 Then gError 42997
    
                'Se não achou o Almoxarifado --> erro
                If lErro <> SUCESSO Then gError 42998
    
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
            End If
            
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

        Case 42997

        Case 42998
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158020)

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
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)

    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then

        If gcolParcPV.Count >= GridParcelas.Row Then gcolParcPV.Remove (GridParcelas.Row)
    
    End If

    Exit Sub
    
Erro_GridParcelas_KeyDown:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178943)
            
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
    'nao pode colocar o codigo abaixo pois grid limpa limpa o campo faz com q seja acionado o change e desligue o comissaoautomatica
'    ComissaoAutomatica.Value = vbUnchecked

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
    'nao pode colocar o codigo abaixo pois grid limpa limpa o campo faz com q seja acionado o change e desligue o comissaoautomatica
'    ComissaoAutomatica.Value = vbUnchecked

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
    'nao pode colocar o codigo abaixo pois grid limpa limpa o campo faz com q seja acionado o change e desligue o comissaoautomatica
'    ComissaoAutomatica.Value = vbUnchecked

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
    'nao pode colocar o codigo abaixo pois grid limpa limpa o campo faz com q seja acionado o change e desligue o comissaoautomatica
'    ComissaoAutomatica.Value = vbUnchecked

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
    'nao pode colocar o codigo abaixo pois grid limpa limpa o campo faz com q seja acionado o change e desligue o comissaoautomatica
    
'    ComissaoAutomatica.Value = vbUnchecked

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
    'nao pode colocar o codigo abaixo pois grid limpa limpa o campo faz com q seja acionado o change e desligue o comissaoautomatica
'    ComissaoAutomatica.Value = vbUnchecked

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
    'nao pode colocar o codigo abaixo pois grid limpa limpa o campo faz com q seja acionado o change e desligue o comissaoautomatica
'    ComissaoAutomatica.Value = vbUnchecked

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
    'nao pode colocar o codigo abaixo pois grid limpa limpa o campo faz com q seja acionado o change e desligue o comissaoautomatica
'    ComissaoAutomatica.Value = vbUnchecked

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

Public Sub GridComissoes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridComissoes)
    
End Sub

Public Sub GridComissoes_RowColChange()

    Call Grid_RowColChange(objGridComissoes)

End Sub

Public Sub GridComissoes_Scroll()

    Call Grid_Scroll(objGridComissoes)

End Sub

Public Sub TipoBloqueio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoBloqueio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Public Sub TipoBloqueio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Public Sub TipoBloqueio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = TipoBloqueio
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataBloqueio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataBloqueio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Public Sub DataBloqueio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Public Sub DataBloqueio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = DataBloqueio
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Observacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Public Sub Observacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Public Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = Observacao
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub CodUsuario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CodUsuario_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Public Sub CodUsuario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Public Sub CodUsuario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = CodUsuario
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ResponsavelBL_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ResponsavelBL_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Public Sub ResponsavelBL_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Public Sub ResponsavelBL_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = ResponsavelBL
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GridBloqueio_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridBloqueio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
    End If

End Sub

Public Sub GridBloqueio_GotFocus()

    Call Grid_Recebe_Foco(objGridBloqueio)

End Sub

Public Sub GridBloqueio_EnterCell()

    Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)

End Sub

Public Sub GridBloqueio_LeaveCell()

    Call Saida_Celula(objGridBloqueio)

End Sub

Public Sub GridBloqueio_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then

        If GridBloqueio.Row <= giGridBloqueioLinhasBloqueadas Then Exit Sub

    End If

    Call Grid_Trata_Tecla1(KeyCode, objGridBloqueio)

    Exit Sub

End Sub

Public Sub GridBloqueio_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridBloqueio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
    End If

End Sub

Public Sub GridBloqueio_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridBloqueio)

End Sub

Public Sub GridBloqueio_RowColChange()

    Call Grid_RowColChange(objGridBloqueio)

End Sub

Public Sub GridBloqueio_Scroll()

    Call Grid_Scroll(objGridBloqueio)

End Sub

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
    If lErro <> SUCESSO Then gError 64199

    objProduto.sCodigo = sProdutoFormatado
                    
    'Lê o produto para pegar a UM_Estoque
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 64200
    
    If lErro = 28030 Then gError 64201

    Call Grid_Trata_Tecla1(KeyCode, objGridAlocacao)

    If iLinhasExistentesAnterior <> objGridAlocacao.iLinhasExistentes Then
    
        For iIndice = 1 To objGridItens.iLinhasExistentes
        
            '********* TRATAMENTO DE GRADE *************
            If gobjPedidoVenda.colItensPedido(iItem).iPossuiGrade <> MARCADO Then
        
                If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoAnterior Then
                    
                    'Converte a Quantidade se Necessário
                    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUmAnterior, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), dFator)
                    If lErro <> SUCESSO Then gError 64198
                    
                    dQuantidadeReservadaAnteriorConvertida = dQuantidadeReservadaAnterior * dFator
                    
                    GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col) = Format(StrParaDbl((GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col)) - dQuantidadeReservadaAnteriorConvertida), "Standard")
                    Exit For
                End If
            Else
            
                For Each objItemRomaneioGrade In gobjPedidoVenda.colItensPedido(iItem).colItensRomaneioGrade
                
                    If objProduto.sCodigo = objItemRomaneioGrade.sProduto Then

                        Set objProdutoItem = New ClassProduto
                        
                        objProdutoItem.sCodigo = gobjPedidoVenda.colItensPedido(iItem).sProduto
                        
                        'Lê o produto para pegar a UM_Estoque
                        lErro = CF("Produto_Le", objProdutoItem)
                        If lErro <> SUCESSO And lErro <> 28030 Then gError 64200
                    
                        If objProdutoItem.iKitVendaComp = DESMARCADO Then
                            'Converte a Quantidade se Necessário
                            lErro = CF("UM_Conversao", objProduto.iClasseUM, sUmAnterior, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), dFator)
                            If lErro <> SUCESSO Then gError 64198
                        Else
                            dFator = 1
                        End If
                        
                        dQuantidadeReservadaAnteriorConvertida = dQuantidadeReservadaAnterior * dFator
    
                        If objProdutoItem.iKitVendaComp = DESMARCADO Then
                            GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col) = Format((StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col)) - dQuantidadeReservadaAnteriorConvertida), "Standard")
                        End If
                        
                        'gobjPedidoVenda.colItensPedido(iItemAtual).dQuantReservada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col))
                    
                        objItemRomaneioGrade.dQuantReservada = objItemRomaneioGrade.dQuantReservada - dQuantidadeReservadaAnteriorConvertida
                        iIndice = 0
                        For Each objReservaItem In objItemRomaneioGrade.colLocalizacao
                            iIndice = iIndice + 1
                            If objReservaItem.sAlmoxarifado = sAlmoxarifado Then
                                objItemRomaneioGrade.colLocalizacao.Remove iIndice
                                Exit For
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    Exit Sub
    
Erro_GridAlocacao_KeyDown:

    Select Case gErr
        
        Case 64198, 64199, 64200
        
        Case 64201
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProdutoAnterior)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158021)

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
                If lErro <> SUCESSO Then gError 52115
    
                objProduto.sCodigo = sProduto
    
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 52116
    
                'Se não achou o Produto --> erro
                If lErro = 28030 Then gError 52117
                    
                ProdutoDescricao.Caption = objProduto.sDescricao
            
            End If
            
            'Guarda a Linha corrente
            giLinhaAnterior = GridAlocacao.Row
                
        End If
    
    End If

    Exit Sub

Erro_GridAlocacao_RowColChange:

    Select Case gErr
        
        Case 52115, 52116
        
        Case 52117
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158022)

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
            If lErro <> SUCESSO Then gError 26784

        Case iGrid_ValorParcela_Col
            lErro = Saida_Celula_ValorParcela(objGridInt)
            If lErro <> SUCESSO Then gError 26785

        Case iGrid_TipoPagto_Col
            lErro = Saida_Celula_TipoPagto(objGridInt)
            If lErro <> SUCESSO Then gError 178945

        Case iGrid_Desc1Codigo_Col, iGrid_Desc2Codigo_Col, iGrid_Desc3Codigo_Col
            lErro = Saida_Celula_TipoDesconto(objGridInt)
            If lErro <> SUCESSO Then gError 26786

        Case iGrid_Desc1Ate_Col, iGrid_Desc2Ate_Col, iGrid_Desc3Ate_Col
            lErro = Saida_Celula_DescontoData(objGridInt)
            If lErro <> SUCESSO Then gError 26787

        Case iGrid_Desc1Valor_Col, iGrid_Desc2Valor_Col, iGrid_Desc3Valor_Col
            lErro = Saida_Celula_DescontoValor(objGridInt)
            If lErro <> SUCESSO Then gError 26830

        Case iGrid_Desc1Perc_Col, iGrid_Desc2Perc_Col, iGrid_Desc3Perc_Col
            lErro = Saida_Celula_DescontoPerc(objGridInt)
            If lErro <> SUCESSO Then gError 26788



    End Select

    Saida_Celula_GridParcelas = SUCESSO

    Exit Function

Erro_Saida_Celula_GridParcelas:

    Saida_Celula_GridParcelas = gErr

    Select Case gErr

        Case 26784, 26785, 26786, 26788, 26787, 26830, 178945

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158023)

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
        If lErro <> SUCESSO Then gError 26792
        'Se a data de vencimento estiver preenchida
        If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))) > 0 Then
            'critica se DataDesconto ultrapassa DataVencimento
            If CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col)) < CDate(objGridInt.objControle.Text) Then gError 26592
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 26793

    Saida_Celula_DescontoData = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoData:

    Saida_Celula_DescontoData = gErr

    Select Case gErr

        Case 26792, 26793
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 26592
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADESCONTO_MAIOR_DATAVENCIMENTO", gErr, CDate(objGridInt.objControle.Text), CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col)))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158024)

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
        If lErro <> SUCESSO Then gError 26794
    
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
    If lErro <> SUCESSO Then gError 26795
    
    Saida_Celula_DescontoPerc = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoPerc:

    Saida_Celula_DescontoPerc = gErr

    Select Case gErr

        Case 26794, 26795
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158025)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridBloqueio(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridBloqueio

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_TipoBloqueio_Col
            lErro = Saida_Celula_TipoBloqueio(objGridInt)
            If lErro <> SUCESSO Then gError 59728

        Case iGrid_ResponsavelBL_Col
            lErro = Saida_Celula_ResponsavelBL(objGridInt)
            If lErro <> SUCESSO Then gError 59729

        Case iGrid_Observacao_Col
            lErro = Saida_Celula_Observacao(objGridInt)
            If lErro <> SUCESSO Then gError 125292

    End Select

    Saida_Celula_GridBloqueio = SUCESSO

    Exit Function

Erro_Saida_Celula_GridBloqueio:

    Saida_Celula_GridBloqueio = gErr

    Select Case gErr

        Case 59728, 59729, 125292

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158026)

    End Select

    Exit Function

End Function

Public Sub Form_Activate()

Dim lErro As Long

On Error GoTo Erro_Form_Activate

    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError 59332
        
    Call TelaIndice_Preenche(Me)

    Exit Sub
     
Erro_Form_Activate:

    Select Case gErr
          
        Case 59332
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158027)
     
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
    If lErro <> SUCESSO Then gError 26805

    'Limpa a Tela
    Call Limpa_Tela_PedidoVenda

    iAlterado = 0
    iComissoesAlterada = 0
    gbCarregandoTela = False

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 26805

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158028)

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

    'Se não houve Quant a Reservar zera Reservas.
    If dQuantidade = 0 Or dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada = 0 Then
        GridItens.TextMatrix(iIndiceGridItens, iGrid_QuantReservada_Col) = Format(0, "Standard")
        iItemAtual = CInt(GridItens.TextMatrix(iIndiceGridItens, iGrid_ItemProduto_Col))
    End If

    lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndiceGridItens, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 26796

    objProduto.sCodigo = sProduto

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 26669

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 26668

    If objProduto.iControleEstoque = PRODUTO_CONTROLE_RESERVA And dQuantidade > 0 And dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada > 0 And (gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA) Then

        lErro = ReservaAlmoxarifadoPadrao(iIndiceGridItens, objProduto, dQuantidade, dQuantidadeCancelada, dQuantidadeFaturada)
        If lErro <> SUCESSO And lErro <> 26680 Then gError 26679

        If gobjFAT.iFaltaEstObrigaDecisao = MARCADO And lErro = 26680 Then gError 26679
    
    ElseIf dQuantidade = 0 And dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada = 0 And gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA Then
    
        For iIndice = 1 To objGridAlocacao.iLinhasExistentes
                
            If GridAlocacao.TextMatrix(iIndice, iGrid_ProdutoAlmox_Col) = GridItens.TextMatrix(iIndiceGridItens, iGrid_Produto_Col) Then
                Call Remove_Reservas(CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)))
            End If
        Next
        
    End If

    Reserva_Processa = SUCESSO

    Exit Function

Erro_Reserva_Processa:

    Reserva_Processa = gErr

    Select Case gErr

        Case 26796, 26669, 26679

        Case 26668
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158029)

    End Select

    Exit Function

End Function

Public Function Move_GridItem_Memoria(objPedidoVenda As ClassPedidoDeVenda, iIndice As Integer, sProduto1 As String) As Long
'Recolhe do Grid os dados do item pedido no parametro

Dim lErro As Long
Dim sProduto As String
Dim objItemPedido As ClassItemPedido, objTributacaoItemPV As ClassTributacaoDocItem
Dim iPreenchido As Integer, objItemPedidoNovo As ClassItemPedido
Dim objDataEntrega As ClassDataEntrega
Dim dQuantidade As Double
Dim iItem As Integer

On Error GoTo Erro_Move_GridItem_Memoria

    Set objItemPedido = New ClassItemPedido

'    'Verifica se o Produto está preenchido
'    If Len(Trim(sProduto1)) > 0 Then
'
'        'Formata o produto
'        lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError 27682
'
    objItemPedido.sProduto = sProduto1
'    End If

    If Len(Trim(sProduto1)) = 0 Then gError 20767 'Or iPreenchido = PRODUTO_VAZIO Then gError 20767

    'Armazena os dados do item
    objItemPedido.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
    
    objItemPedido.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col))
    If gobjFAT.iExibeQtdeCanceladaItensPV = MARCADO Then objItemPedido.dQuantCancelada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))
    objItemPedido.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))
    objItemPedido.dPrecoTotal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
    objItemPedido.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
    If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col))) > 0 Then
        objItemPedido.dtDataEntrega = CDate(GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col))
    Else
        objItemPedido.dtDataEntrega = DATA_NULA
    End If
    objItemPedido.dQuantReservada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col))
    objItemPedido.dQuantFaturada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantFaturada_Col))
    objItemPedido.iPrioridade = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_Prioridade_Col))

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
    objItemPedidoNovo.iPrioridade = objItemPedido.iPrioridade

    lErro = CF2(Me, "PV_Move_GridItem_Memoria", objPedidoVenda, objItemPedidoNovo, iIndice, sProduto1)
    If lErro <> SUCESSO Then gError 124179
    
    Set objItemPedidoNovo.colItemPVEntrega = gobjPedidoVenda.colItensPedido(iIndice).colItemPVEntrega
    Set objItemPedidoNovo.objInfoAdicDocItem = gobjPedidoVenda.colItensPedido(iIndice).objInfoAdicDocItem
    
    dQuantidade = 0

    For Each objDataEntrega In objItemPedidoNovo.colItemPVEntrega
        dQuantidade = dQuantidade + objDataEntrega.dQuantidade
    Next
    
    If dQuantidade > 0 Then
    
        If Abs(dQuantidade - (objItemPedidoNovo.dQuantidade - objItemPedidoNovo.dQuantCancelada)) > QTDE_ESTOQUE_DELTA Then gError 183245
    
    End If
    
    If gobjFAT.iExibePercFatPV = MARCADO Then

        objItemPedidoNovo.dQuantFaturadaAMais = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantFatAMais_Col))
        objItemPedidoNovo.dPercentMaisReceb = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_PercentMaisReceb_Col))
        objItemPedidoNovo.dPercentMenosReceb = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_PercentMenosReceb_Col))
    
         For iItem = 0 To RecebForaFaixa.ListCount - 1
            If GridItens.TextMatrix(iIndice, iGrid_RecebForaFaixa_Col) = RecebForaFaixa.List(iItem) Then
    
                'coloca no Grid Itens RecebForaFaixa
                objItemPedidoNovo.iRecebForaFaixa = RecebForaFaixa.ItemData(iItem)
    
            End If
        Next
        
    Else
        objItemPedidoNovo.iRecebForaFaixa = ERRO_E_REJEITA_RECEBIMENTO
    End If
        
    If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
    
         For iItem = 0 To TabPrecoItemPV.ListCount - 1
            
            If GridItens.TextMatrix(iIndice, iGrid_TabPrecoItem_Col) = TabPrecoItemPV.List(iItem) Then
    
                objItemPedidoNovo.iTabelaPreco = TabPrecoItemPV.ItemData(iItem)
                Exit For
    
            End If
            
        Next
    
    End If
    
    objItemPedidoNovo.dComissao = -1
    If gobjFAT.iExibeComissaoItensPV = MARCADO And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col))) <> 0 Then
        
        objItemPedidoNovo.dComissao = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col))
    
    End If
    
    Move_GridItem_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItem_Memoria:

    Move_GridItem_Memoria = gErr

    Select Case gErr

        Case 20767
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 27682

        Case 183245
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTSOLICITADA_DIFERE1", gErr, iIndice, sProduto1, objItemPedidoNovo.dQuantidade, dQuantidade)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158040)

    End Select
    
    Exit Function

End Function

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
            If lErro <> 25060 And lErro <> SUCESSO Then gError 35891
            If lErro = 25060 Then gError 35892 'Não encontrou
            
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

        Case 35891

        Case 35892
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158042)

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
            If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
            If lErro <> SUCESSO Then gError 42765

            If objProduto.iKitVendaComp <> MARCADO Then
                Call Atualiza_Grid_Itens(colItensPV(iIndice))
            End If
        Else
    
            objProduto.sCodigo = colItensPV.Item(iIndice).sProduto
            'Lê o produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
            If lErro = 28030 Then gError 42765 'Não encontrou
            'Faz a conversão da unidade do item para a unidade de estoque
            lErro = CF("UM_Conversao", objProduto.iClasseUM, colItensPV.Item(iIndice).sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 42766
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
                If lErro <> SUCESSO And lErro <> 25056 Then gError 46181
                If lErro = 25056 Then gError 46182
    
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

        Case 42764, 42766, 46181

        Case 42765
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 46182
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158043)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 26806

    'Limpa a Tela
    Call Limpa_Tela_PedidoVenda

    iAlterado = 0
    iComissoesAlterada = 0
    gbCarregandoTela = False

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 26806

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158044)

    End Select

    Exit Sub

End Sub

Public Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objPedidoVendaDB As New ClassPedidoDeVenda
Dim dValor As Double
Dim bBloqueio As Boolean
Dim objBloqueioPV As ClassBloqueioPV

On Error GoTo Erro_BotaoImprimir_Click

'    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Codigo.Text)) = 0 Then gError 93215
    If Len(Trim(Cliente.Text)) = 0 Then gError 93216
    If Len(Trim(Filial.Text)) = 0 Then gError 93217
    If Len(Trim(FilialEntrega.Text)) = 0 Then gError 93218
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 93219
'    If Len(Trim(NaturezaOp.Text)) = 0 Then gError 93230
        
    lErro = Move_Pedido_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then gError 93228

    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 93214

    If lErro = 26509 Then gError 93231
    
    If gobjFAT.iImpedeImpPVComBloq = MARCADO Then
    
        objPedidoVendaDB.lCodigo = objPedidoVenda.lCodigo
        objPedidoVendaDB.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
    
        bBloqueio = False
        lErro = CF("BloqueiosPV_Le", objPedidoVendaDB)
        If lErro <> SUCESSO Then gError 198692
        
        For Each objBloqueioPV In objPedidoVendaDB.colBloqueiosPV
            If objBloqueioPV.dtDataLib = DATA_NULA Then
                bBloqueio = True
                Exit For
            End If
        Next
        
        If bBloqueio Then gError 198693
    
    End If
    
'    GL_objMDIForm.MousePointer = vbDefault

    lErro = objRelatorio.ExecutarDireto("Pedido de Venda", "PedidoVenda >= @NPEDVENDINIC E PedidoVenda <= @NPEDVENDFIM", 1, "PedVenda", "NPEDVENDINIC", objPedidoVenda.lCodigo, "NPEDVENDFIM", objPedidoVenda.lCodigo)
    If lErro <> SUCESSO Then gError 93229
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 93214, 93229, 198692

        Case 93215
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)

        Case 93216
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 93217
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)

        Case 93218
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)
        
        Case 93219
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 93230
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 93231
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoVenda.lCodigo)

        Case 198693
            Call Rotina_Erro(vbOKOnly, "ERRO_PV_COM_BLOQ_IMP_IMPRESSAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158045)

    End Select

'    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Public Function Gravar_Registro(Optional bNaoGravar As Boolean = False) As Long

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim dValorTotal As Double
Dim dValor As Double, bExibirBloqsPV As Boolean
Dim vbMsgRes As VbMsgBoxResult, colBloqPVAux As New Collection, objBloqPV As ClassBloqueioPV, objBloqPVAux As ClassBloqueioPV
Dim objAnaliseVendas As New ClassAnaliseVendaInfo

On Error GoTo Erro_Gravar_Registro

    If Len(Trim(Codigo.Text)) = 0 Then gError 26807
    If Len(Trim(Cliente.Text)) = 0 Then gError 26808
    If Len(Trim(Filial.Text)) = 0 Then gError 26809
    If Len(Trim(FilialEntrega.Text)) = 0 Then gError 56957
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 26810
    If Len(Trim(gobjTribTab.sNatOpInterna)) = 0 Then gError 26811
    
    If TipoFrete.ItemData(TipoFrete.ListIndex) = FRETE_DESTINATARIO And Len(Trim(Transportadora.Text)) = 0 Then gError 75489
    If TipoFrete.ItemData(TipoFrete.ListIndex) = FRETE_SEM And StrParaDbl(ValorFrete.Text) <> 0 Then gError 207416
     
'    If giTipoVersao = VERSAO_FULL Then
        If Len(Trim(FilialFaturamento.Text)) = 0 Then gError 42378
'    End If

    dValor = CDbl(ValorTotal.Caption)

    If dValor < 0 Then gError 30594

    If Len(Trim(PesoBruto.Text)) > 0 And Len(Trim(PesoLiquido.Text)) > 0 Then
        If StrParaDbl(PesoLiquido.Text) > StrParaDbl(PesoBruto.Text) Then gError 51376
    End If
    
    lErro = CF("PV_PedidoRepr_Valida", StrParaDbl(Codigo.Text), giFilialEmpresa, PedidoRepr.Text)
    If lErro <> SUCESSO Then gError 180380

    lErro = Valida_Grid_Itens()
    If lErro <> SUCESSO Then gError 26812

    lErro = Valida_Grid_Parcelas()
    If lErro <> SUCESSO Then gError 26816

    lErro = objTabComissoes.Valida_Grid_Comissoes()
    If lErro <> SUCESSO Then gError 26823

    lErro = Valida_Grid_Bloqueio()
    If lErro <> SUCESSO Then gError 26827
    
    If gobjFAT.iPedidoValidaEmbalagem = MARCADO Then
    
        lErro = gobjEmbalagem.Valida_Grid_Embalagens()
        If lErro <> SUCESSO And lErro <> 96133 Then gError 96132
        If lErro = 96133 Then
        
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ACEITA_GRIDEMB_INCOMPLETO")
            If vbMsgRes = vbNo Then gError 81746
            
        End If
    
    End If
    
    lErro = Move_Pedido_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then gError 26829
    
    lErro = CF2(Me, "PV_Move_Pedido_Memoria", objPedidoVenda)
    If lErro <> SUCESSO Then gError 124169
    
    lErro = CF("PedidoVenda_Verifica_Preenchimento", objPedidoVenda)
    If lErro <> SUCESSO Then gError 124070
    
    lErro = gobjTribTab.Valida_TributacaoNF_Gravacao(objPedidoVenda)
    If lErro <> SUCESSO Then gError 56931
    
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
    
    lErro = Trata_Alteracao(objPedidoVenda, objPedidoVenda.iFilialEmpresa, objPedidoVenda.lCodigo)
    If lErro <> SUCESSO Then gError 130496
    
    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786
    
    lErro = gobjTribTab.Valida_Dados
    If lErro <> SUCESSO Then gError 181786
    
    lErro = CF("PedidoVenda_Valida_MargContr", objPedidoVenda)
    If lErro <> SUCESSO Then gError 201040
    
    If gobjCRFAT.iValidaDescMaxTabPreco <> 0 And gobjCRFAT.iBloquearPVPrecoBaixo = 0 And objPedidoVenda.colParcelas.Count <> 0 Then
    
        Call objPedidoVenda.AnaliseVenda_Preenche(objAnaliseVendas)
        lErro = CF("AnaliseVenda_ValidaDescMax", objAnaliseVendas)
        If lErro <> SUCESSO Then gError 181786
    
    End If
    
    If Not bNaoGravar Then
    
        'Cyntia
        GL_objMDIForm.MousePointer = vbHourglass
           
        'Grava no BD
        lErro = CF("PedidoVenda_Grava", objPedidoVenda)
        If lErro <> SUCESSO Then gError 46183
        
        GL_objMDIForm.MousePointer = vbDefault
        
        'gravar anotacao, se houver
        If Not (gobjAnotacao Is Nothing) Then
        
            If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
            
                gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_PEDIDOVENDA
                gobjAnotacao.sID = CStr(objPedidoVenda.iFilialEmpresa) & "," & CStr(objPedidoVenda.lCodigo)
                gobjAnotacao.dtDataAlteracao = gdtDataHoje
                
                lErro = CF("Anotacoes_Grava", gobjAnotacao)
                If lErro <> SUCESSO Then gError 124058
                
            End If
            
        End If
        
        Call PedidoVenda_ExibirBloqsPV(objPedidoVenda, colBloqPVAux, bExibirBloqsPV)
        If bExibirBloqsPV Then
            
            lErro = Chama_Tela("BloqPV_ConsultaTodos", objPedidoVenda)
            If lErro <> SUCESSO Then gError 124128
        
        End If
        
        'Se for para enviar por email depois da gravação
        If EmailGravacao.Value = vbChecked Then
            Call BotaoEmail_Click
        End If
        
        If ImprimeGravacao.Value = vbChecked Then
            Call BotaoImprimir_Click
        End If
        
        If ImprimirConfGravacao.Value = vbChecked Then
            Call BotaoImprimirConf_Click
        End If
   
    End If
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    If Not bNaoGravar Then
        GL_objMDIForm.MousePointer = vbDefault
    End If
    
    Select Case gErr

        Case 81746, 124058, 124070, 124128, 124169, 130496, 181786, 201040
        
        Case 26807
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)

        Case 26808
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 26809
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)

        Case 56957
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)
        
        Case 26810
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 26811
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 26812, 26816, 26823, 26827, 26829, 46183, 56931, 96132, 180380

        Case 30594
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_PV_NEGATIVO", gErr)

        Case 42378
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FATURAMENTO_NAO_PREENCHIDA", gErr)

        Case 51376
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr)

        Case 75489
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_TRANSPORTADORA_NAO_PREENCHIDA", gErr)
            If vbMsgRes = vbYes Then Resume Next
            
        Case 207416
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_FRETE_COM_VALOR", gErr)
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158046)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Itens() As Long

Dim iIndice As Integer
Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Valida_Grid_Itens

    'Verifica se há itens no grid
    If objGridItens.iLinhasExistentes = 0 Then gError 26813
    
    'para cada item do grid
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        lErro = CF2(Me, "PV_Valida_Grid_Itens", iIndice)
        If lErro <> SUCESSO Then gError 177169
        
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col))) = 0 Then gError 51455
        
        lErro = Valor_Positivo_Critica(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col))
        If lErro <> SUCESSO Then gError 26814

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) = 0 Then gError 51456

        lErro = Valor_Positivo_Critica(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))
        If lErro <> SUCESSO Then gError 26815

    Next

    Valida_Grid_Itens = SUCESSO

    Exit Function

Erro_Valida_Grid_Itens:

    Valida_Grid_Itens = gErr

    Select Case gErr

        Case 26813
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_PV", gErr)

        Case 51455
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 51456
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 26814, 26815, 177169

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158047)

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
Dim dValorPagar As Double
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Valida_Grid_Parcelas

    'Verifica se alguma parcela foi informada
    If objGridParcelas.iLinhasExistentes = 0 Then gError 26817
    
    dSomaParcelas = 0
    
    'Para cada Parcela do grid de parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        dtDataEmissao = StrParaDate(DataEmissao.Text)
        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) = 0 Then gError 26818
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) = 0 Then gError 26821
                        
        If dtDataVencimento < dtDataEmissao Then gError 26819
        
        'Se o tipo de desconto 1 estiver preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))) > 0 Then
            iDesconto = 1
            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
            'Verifica se a data de desconto está preenchdida
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))) = 0 Then gError 51066
            'Recolhe o Valor ou Percentual de desconto
            If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))) = 0 Then gError 51069
            Else
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col))) = 0 Then gError 51070
            End If
            'Se o tipo de desconto 2 estiver preenchido
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))) > 0 Then
                iDesconto = 2
                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
                'Verifica se a data de desconto está preenchdida
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))) = 0 Then gError 51067
                'Faz a crítica da ordem das datas de desconto
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 51075
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 51077
                'Recolhe o Valor ou Percentual de desconto
                If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))) = 0 Then gError 51071
                Else
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col))) = 0 Then gError 51072
                End If
                'Se o tipo de desconto 3 estiver preenchido
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))) > 0 Then
                    iDesconto = 3
                    iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
                    'Verifica se a data de desconto está preenchdida
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))) = 0 Then gError 51068
                    'Faz a crítica da ordem das datas de desconto
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 51076
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 51078
                    'Recolhe o Valor ou Percentual de desconto
                    If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))) = 0 Then gError 51073
                    Else
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col))) = 0 Then gError 51074
                    End If
                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
                    If dtDataDesconto > dtDataVencimento Then gError 51364
                Else
                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
                    If dtDataDesconto > dtDataVencimento Then gError 51363
                End If
            Else
                dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
                If dtDataDesconto > dtDataVencimento Then gError 51362
            End If
           
        End If
        
        If iIndice > 1 Then If CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col)) < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_Col)) Then gError 26820

        dSomaParcelas = dSomaParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_TipoPagto_Col))) > 0 Then
            
            iTipoPagto = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_TipoPagto_Col))
            
            lErro = CF("Valida_TipoPagto", iIndice, iTipoPagto, gcolParcPV(iIndice))
            If lErro <> SUCESSO Then gError 178975
            
        End If
            
    Next
    
    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then gError 140294
    
    If Abs(dValorPagar - dSomaParcelas) > DELTA_VALORMONETARIO Then gError 26822

    Valida_Grid_Parcelas = SUCESSO
    
    Exit Function
    
Erro_Valida_Grid_Parcelas:

    Valida_Grid_Parcelas = gErr
    
    Select Case gErr
    
        Case 26817
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PV_NF_REMESSA")
            If vbMsgRes = vbYes Then Valida_Grid_Parcelas = SUCESSO
            
        Case 26818
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_NAO_INFORMADA", gErr, iIndice)
            
        Case 26819
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_MENOR", gErr, iIndice, dtDataVencimento, dtDataEmissao)
        
        Case 26820
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_COBRANCA_NAO_ORDENADA", gErr)
        
        Case 26821
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PARCELA_COBRANCA_NAO_INFORMADO", gErr, iIndice)
        
        Case 26822
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_COBRANCA_INVALIDA", gErr)

        Case 51066, 51067, 51068
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_NAO_PREENCHIDA", gErr, iDesconto, iIndice)
        
        Case 51069, 51070, 51071, 51072, 51073, 51074
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_PARCELA_NAO_PREENCHIDO", gErr, iDesconto, iIndice)

        Case 51075, 51076
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTOS_DESORDENADAS", gErr, iIndice)

        Case 51077, 51078
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTO_IGUAIS", gErr, iDesconto - 1, iDesconto, iIndice)

        Case 51362, 51363, 51364
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESC_PARCELA_SUP_DATA_VENCIMENTO", gErr, dtDataDesconto, iDesconto, iIndice)

        Case 178975

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158048)
            
    End Select
    
    Exit Function

End Function

Private Function Valida_Grid_Bloqueio() As Long
'Validação de GridBloqueio

Dim lErro As Long
Dim iIndice As Integer
Dim bBloqueioTotal As Boolean
Dim iTipoBloqueio As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim dtDataLiberacao As Date

On Error GoTo Erro_Valida_Grid_Bloqueio

    bBloqueioTotal = False

    If objGridBloqueio.iLinhasExistentes > giGridBloqueioLinhasBloqueadas Then

        For iIndice = giGridBloqueioLinhasBloqueadas + 1 To objGridBloqueio.iLinhasExistentes

            'Testa se preencheu tipo de bloqueio
            If Len(Trim(GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))) = 0 Then gError 26828

        Next

    End If

    If objGridBloqueio.iLinhasExistentes > 0 Then

        'Testa se existe bloqueio total não liberado
        For iIndice = giGridBloqueioLinhasBloqueadas + 1 To objGridBloqueio.iLinhasExistentes

            iTipoBloqueio = Codigo_Extrai(GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col))
            dtDataLiberacao = StrParaDate(GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col))

            If iTipoBloqueio = BLOQUEIO_TOTAL And dtDataLiberacao = DATA_NULA Then bBloqueioTotal = True

        Next

        If bBloqueioTotal Then

            'Pergunta se deseja continuar com bloqueio total existente
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_BLOQUEIO_TOTAL")
            If vbMsgRes = vbNo Then gError 30057

            'Se continuar com bloqueio total limpa reservas
            Call Grid_Limpa(objGridAlocacao)
            
            'Zera QuantReservadas no GridItens
            For iIndice = 1 To objGridItens.iLinhasExistentes
                GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col) = ""
            Next

        End If

    End If

    Valida_Grid_Bloqueio = SUCESSO

    Exit Function

Erro_Valida_Grid_Bloqueio:

    Valida_Grid_Bloqueio = gErr

    Select Case gErr

        Case 26828
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_BLOQUEIO_NAO_INFORMADO", gErr, iIndice)

        Case 30057

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158049)

    End Select

    Exit Function

End Function

Public Sub BotaoBaixarPedido_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoBaixarPedido_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = PedidoVenda_Critica_Le(objPedidoVenda)
    If lErro <> SUCESSO Then gError 62125
    
    lErro = CF("PedidoDeVenda_Baixar_Pedido", objPedidoVenda)
    If lErro <> SUCESSO Then gError 62124
    
    Call Limpa_Tela_PedidoVenda
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoBaixarPedido_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 62124, 62125
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158050)
    
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
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 46130

    objPedidoVenda.lCodigo = CLng(Codigo.Text)
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa
    Set objPedidoVenda.objTela = Me
    
    'Verifica se é um pedido baixado
    lErro = CF("PedidoVendaBaixado_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 46135 Then gError 46131

    'Se for um Pedido Baixado --> erro
    If lErro = SUCESSO Then gError 46136

    'Lê o Pedido
    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 46137

    'Se não encontrou o Pedido de Venda --> erro
    If lErro = 26509 Then gError 46138

    'Pede a confirmação da exclusão do Pedido de Venda
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PEDIDO_VENDA", objPedidoVenda.lCodigo)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Faz a exclusão do Pedido de Venda
    lErro = CF("PedidoVenda_Exclui", objPedidoVenda)
    If lErro <> SUCESSO Then gError 46139

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

        Case 46130
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)

        Case 46131, 46137, 46139

        Case 46136
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDVENDA_BAIXADO_EXCLUSAO", gErr, objPedidoVenda.lCodigo)

        Case 46138
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", gErr, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158051)

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
            If lErro <> SUCESSO Then gError 25437
    
        Next
    
        'Calcula o Valor Total da Nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 51046

    End If
    
    Trata_TabelaPreco = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco:

    Trata_TabelaPreco = gErr

    Select Case gErr

        Case 25437 'tratado na rotina chamada

        Case 51046

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158052)

    End Select

    Exit Function

End Function

Public Function Trata_TabelaPreco_Item(iLinha As Integer) As Long
'faz tratamento de tabela de preço para um ítem (produto)

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim sProduto As String
Dim iPreenchido As Integer, iTabela As Integer
Dim iItem As Integer
Dim dPercDescontoTab As Double, dPrecoComDescontoTab As Double

On Error GoTo Erro_Trata_TabelaPreco_Item

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 39147

    If iPreenchido <> PRODUTO_VAZIO And Len(Trim(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col))) > 0 Then

        objTabelaPrecoItem.sCodProduto = sProduto
        
        If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        
             For iItem = 0 To TabPrecoItemPV.ListCount - 1
                
                If GridItens.TextMatrix(iLinha, iGrid_TabPrecoItem_Col) = TabPrecoItemPV.List(iItem) Then
        
                    iTabela = TabPrecoItemPV.ItemData(iItem)
                    Exit For
        
                End If
                
            Next
        
        End If
        
        If iTabela = 0 Then iTabela = Codigo_Extrai(TabelaPreco.Text)
        
        lErro = CF2(Me, "NFCF_PrecoUnitario_Calcula", giFilialEmpresa, sProduto, GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), Cliente.Text, Codigo_Extrai(Filial.Text), iTabela, StrParaDbl(PercAcrescFin.Text) / 100, MOEDA_REAL, Nothing, dPrecoUnitario, False, dPercDescontoTab, dPrecoComDescontoTab)
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
    
        lErro = CF2(Me, "PV_Saida_Celula_PrecoUnitario", dPrecoUnitario, iLinha)
        If lErro <> SUCESSO Then gError 132828

        'Calcula o Preco Total do Ítem
        Call CF2(Me, "PV_PrecoTotal_Calcula", iLinha)

        ''*** incluidos p/tratamento de tributacao *******************************
        lErro = gobjTribTab.Alteracao_Item_Grid(iLinha)
        If lErro <> SUCESSO Then gError 132828
        '*** fim tributacao
    
    End If

    Trata_TabelaPreco_Item = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco_Item:

    Trata_TabelaPreco_Item = gErr

    Select Case gErr

        Case 39147, 39148, 39149, 39150, 132828, ERRO_SEM_MENSAGEM 'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158053)

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
Dim colTipoFrete As New Collection
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer

On Error GoTo Erro_Trata_FilialCliente

    objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
    objCliente.sNomeReduzido = Trim(Cliente.Text)

    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Trim(Cliente.Text), objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 26515
    If lErro = 17660 Then gError 25434

    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 26516
    If lErro = 12348 Then gError 26517
    
    Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, objFilialCliente.iCodFilial)

    'Comissões
    'Alterado por Tulio em 27/03
    'se empresa nao utiliza as regras para
    'calculo de comissoes
    'Início trecho Corrigido por Rafael Menezes em 20/09/2002
    If Not (gobjCRFAT.iUsaComissoesRegras = USA_REGRAS) And ComissaoAutomatica.Value = 1 Then
    'Fim trecho Corrigido por Rafael Menezes em 20/09/2002
        
        objTabComissoes.Limpa_Grid_Comissoes
        
        'calcular as comissoes no modelo antigo...
        lErro = objTabComissoes.Comissao_Automatica_FilialCli_Exibe(objFilialCliente)
        If lErro <> SUCESSO Then gError 26724
    
    End If
    
    'Transportadora
    If Len(Trim(Transportadora.Text)) = 0 Then

        'Testa se FilialCliente tem Transportadora
        If objFilialCliente.iCodTransportadora <> 0 Then

            objTransportadora.iCodigo = objFilialCliente.iCodTransportadora
            lErro = CF("Transportadora_Le", objTransportadora)
            If lErro <> SUCESSO And lErro <> 19250 Then gError 51064
            If lErro = SUCESSO Then
                Transportadora.Text = objTransportadora.sNomeReduzido
                Call Transportadora_Validate(False)
            End If

        ElseIf objCliente.iTipo <> 0 Then

            'Testa se TipoCliente já foi lido no BD
            If objTipoCliente.iCodigo = 0 Then

                objTipoCliente.iCodigo = objCliente.iTipo
                lErro = CF("TipoCliente_Le", objTipoCliente)
                If lErro <> SUCESSO And lErro <> 19062 Then gError 51065
                If lErro <> SUCESSO Then objTipoCliente.iCodTransportadora = 0

            End If

            'Testa se TipoCliente tem Transportadora
            If objTipoCliente.iCodTransportadora <> 0 Then
                lErro = CF("Transportadora_Le", objTransportadora)
                If lErro <> SUCESSO And lErro <> 19250 Then gError 51063
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
            If lErro <> SUCESSO And lErro <> 19250 Then gError 124056
            If lErro <> SUCESSO Then gError 124057
            
            'Coloca a transportadora na tela
            TranspRedespacho.Text = objTransportadora.sNomeReduzido
            Call TranspRedespacho_Validate(bSGECancelDummy)
        
            RedespachoCli.Value = objFilialCliente.iDetPagFrete
        
        End If
    
    End If
    
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207414
    
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
    
    If objFilialCliente.iCodMensagem = 0 Then objFilialCliente.iCodMensagem = objCliente.iCodMensagem
    
    lErro = CF2(Me, "NFCF_Cliente_Filial_Validate", objFilialCliente)
    If lErro <> SUCESSO Then gError 131894

    If objFilialCliente.iCodMensagem > 0 Then
        objMensagem.iCodigo = objFilialCliente.iCodMensagem
        'Lê a Mensagem "default" p/ o Cliente
        lErro = CF("Mensagem_Le", objMensagem)
        If lErro <> SUCESSO And lErro <> 19234 Then gError 140379

        If lErro = SUCESSO Then Mensagem.Text = objMensagem.sDescricao
    End If

    If gobjCRFAT.iPrecoUnitarioFormPreco <> 0 Then Call Trata_TabelaPreco
    
    Trata_FilialCliente = SUCESSO

    Exit Function

Erro_Trata_FilialCliente:

    Trata_FilialCliente = gErr

    Select Case gErr

        Case 25434
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA1", gErr, Cliente.Text, objFilialCliente.iCodFilial)

        Case 26515, 26516, 51063, 51064, 51065, 25724, 124056, 131894, 140379, 207414

        Case 26517
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case 124057
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr, objTransportadora.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158054)

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

        'Erro já tratado
        Case 22318

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158055)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 25443

    'Nao encontrou o item com o código informado
    If lErro = 6730 Then gError 25444

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 25445

    giFilialFatAlterada = 0
    
    Exit Sub

Erro_FilialFaturamento_Validate:

    Cancel = True


    Select Case gErr

        Case 25443

        Case 25444
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, iCodigo)

        Case 25445
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialFaturamento.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158056)

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

        Case 51139

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158057)

    End Select

    Exit Sub

End Sub

Public Sub PercAcrescFin_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercAcrescFin_Validate

    If giPercAcresFinAlterado = 0 Then Exit Sub

    If Len(Trim(PercAcrescFin.Text)) > 0 Then
        lErro = Porcentagem_Critica_Negativa(PercAcrescFin)
        If lErro <> SUCESSO Then gError 26717
    End If

    If Len(Trim(TabelaPreco.Text)) > 0 Then

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 46190

    End If

    giPercAcresFinAlterado = 0

    Exit Sub

Erro_PercAcrescFin_Validate:

    Cancel = True


    Select Case gErr

        Case 26717

        Case 46190

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158058)

    End Select

    Exit Sub

End Sub

Public Function Cobranca_Automatica() As Long
'recalcula o tab de cobranca

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Cobranca_Automatica

    If CobrancaAutomatica.Value = 1 And Len(Trim(CondicaoPagamento.Text)) <> 0 Then
    
        objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)
        
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 46187
        If lErro <> SUCESSO Then gError 46188
        
        lErro = GridParcelas_Preenche(objCondicaoPagto)
        If lErro <> SUCESSO Then gError 46189
 
    End If
    
    Cobranca_Automatica = SUCESSO
    
    Exit Function
    
Erro_Cobranca_Automatica:
    
    Cobranca_Automatica = gErr
    
    Select Case gErr
    
        Case 46187, 46189
        
        Case 46188
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158059)
    
    End Select
    
    Exit Function
        
End Function

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
    If lErro <> SUCESSO Then gError 102401
    
    Exit Sub

Erro_VolumeEspecie_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102401
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158060)

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
    If lErro <> SUCESSO Then gError 102400
    
    Exit Sub

Erro_VolumeMarca_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102400
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158061)

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
    If lErro <> SUCESSO Then gError 46529

    'Coloca o valor formatado na Tela
    PesoLiquido.Text = Format(PesoLiquido.Text, PesoLiquido.Format)

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True


    Select Case gErr

        Case 46529

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158062)

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
    If lErro <> SUCESSO Then gError 46530

    'Coloca o valor formatado na Tela
    PesoBruto.Text = Format(PesoBruto.Text, PesoBruto.Format)

    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True

    Select Case gErr

        Case 46530

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158063)

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

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158064)

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
        If lErro <> SUCESSO Then gError 26642
        
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
    
        Case 26642
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158065)
            
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PEDIDO_VENDA_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Pedido de Venda"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PedidoVenda"
    
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

    Call CF2(Me, "PV_Codigo_Enabled", True)

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    Set gobjInfoAdic = Nothing

    Codigo.Text = ""
    CodBase.Caption = ""
    Filial.Clear
    ComissaoAutomatica.Value = vbChecked
    CobrancaAutomatica.Value = vbChecked
    TotalPercentualComissao.Caption = ""
    TotalValorComissao.Caption = ""
    TotalValorBase.Caption = ""
    
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
    
    DataRefFluxo.PromptInclude = False
    DataRefFluxo.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataRefFluxo.PromptInclude = True
    
    TipoFrete.ListIndex = 0
    ProdutoDescricao.Caption = ""
        
    FilialOrcamento.Text = "" 'Incluido por leo em 25/04/02
        
'    If giTipoVersao = VERSAO_FULL Then
        Seleciona_FilialEmpresa
        TabelaPreco.Text = ""
'    End If

    
    '************** TRATAMENTO DE GRADE **************
    For iIndice = 1 To objGridItens.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next
    '*************************************************
    
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridBloqueio)
    Call Grid_Limpa(objGridAlocacao)
    
    objTabComissoes.Limpa_Grid_Comissoes
    
    Call Grid_Limpa(objGridParcelas)

    Set gcolParcPV = New Collection

    Set gcolItemPedido = New Collection

    Call gobjTribTab.Limpa_Tela
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    Call DataEmissao_Validate(bSGECancelDummy)
    
    If gobjFAT.iImprimePVGravacao <> 0 Then
        ImprimeGravacao.Value = vbChecked
    Else
        ImprimeGravacao.Value = vbUnchecked
    End If
    ImprimirConfGravacao.Value = vbUnchecked
    EmailGravacao.Value = vbUnchecked
    
    FaturaIntegral.Value = vbUnchecked
    
    iAlterado = 0
    iComissoesAlterada = 0
    gbCarregandoTela = False
    
    giValorDescontoAlterado = 0
    giClienteAlterado = 0
    giFilialAlterada = 0
    giFilialFatAlterada = 0
    giFilialEntregaAlterada = 0
    giDataReferenciaAlterada = 0
    giDataRefFluxoAlterada = 0
    giDataEntregaAlterada = 0

    giValorDescontoManual = 0
    giFilialEntregaManual = 0
    
    giGridBloqueioLinhasBloqueadas = 0
    
    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0
    giPercAcresFinAlterado = 0
    
    Call CF2(Me, "NFCF_Limpa_Tela")
    
    Call CF2(Me, "PV_Limpa_Tela")
    
    Call gobjTribTab.Limpa_Tela
    
    'embalagem
    Call gobjEmbalagem.Limpa_Tela_Embalagem

    Set gobjAnotacao = Nothing
    
    'Coloca clientes em atraso em vermelho
    If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)
        
    Exit Sub

End Sub

Public Function PV_Limpa_Tela(ByVal objCT As Object) As Long
    PV_Limpa_Tela = SUCESSO
End Function

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

    lErro = CF2(Me, "PV_Cliente_Formata", objCliente)
    If lErro <> SUCESSO Then gError 56916
        
    'Coloca clientes em atraso em vermelho
    If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)
                
    'para fazer valer o que veio do bd
    giValorDescontoManual = 1
    giFilialEntregaManual = 1
    
    giClienteAlterado = 0
    
    Exit Sub

Erro_Cliente_Formata:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 56915, 56916
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158066)

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
    If lErro <> SUCESSO And lErro <> 17660 Then gError 56918

    If lErro = 17660 Then gError 56919

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_Filial_Formata:

    Select Case gErr

        Case 56918
        
        Case 56919
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158067)

    End Select

    Exit Sub

End Sub

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
    
    Dim lErro As Long
    
    If KeyCode = vbKeyF4 Then
        Call BotaoUltVendasCli_Click
    End If
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is CondicaoPagamento Then
            Call CondPagtoLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Mensagem Then
            Call MensagemLabel_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call BotaoVendedores_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        ElseIf Me.ActiveControl Is TranspRedespacho Then
            Call TranspRedLabel_Click
        ElseIf Me.ActiveControl Is DescricaoProduto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        End If
    ElseIf KeyCode = KEYCODE_CODBARRAS Then
        lErro = CF2(Me, "PV_Trata_CodigoBarras1")
    End If
    
    If KeyCode = KEYCODE_BOTAOCONSULTA Then
        Call BotaoTipoPagto_Click
    End If
    
    
    Call CF2(Me, "PV_UserControl_KeyDown", KeyCode, Shift)

End Sub

Public Function PV_UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer) As Long
    PV_UserControl_KeyDown = SUCESSO
End Function

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

Public Sub BotaoGerarNFiscal_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objItemPV As ClassItemPedido
Dim objNFiscal As New ClassNFiscal
Dim objNFiscalInfo As New ClassNFiscalInfo
Dim sSeriePadrao As String
Dim objSerie As New ClassSerie

On Error GoTo Erro_BotaoGerarNFiscal_Click

    GL_objMDIForm.MousePointer = vbHourglass

    lErro = PedidoVenda_Critica_Le(objPedidoVenda)
    If lErro <> SUCESSO Then gError 62126

    For Each objItemPV In objPedidoVenda.colItensPedido
        If objItemPV.dQuantFaturada > 0 Then
            gError 62127
            Exit For
        End If
    Next
    
    Call Transfere_Dados_Pedido_NFiscal(objPedidoVenda, objNFiscal)
    
    'Lê a Serie Padrao para Geracao de NF
    lErro = CF("Serie_Le_Padrao", sSeriePadrao)
    If lErro <> SUCESSO Then gError 64121
    
    If Trim(sSeriePadrao) = "" Then gError 64124
    
    objSerie.sSerie = sSeriePadrao
    objSerie.iFilialEmpresa = giFilialEmpresa
    
    'Lê a Filial Empresa no BD
    lErro = CF("Serie_Le", objSerie)
    If lErro <> SUCESSO And lErro <> 22202 Then gError 64122
    
    If lErro = 22202 Then gError 64123
    
    Set gobjGeracaoNFiscal = New ClassGeracaoNFiscal
    
    If objSerie.iTipoFormulario = TIPO_FORMULARIO_NFISCAL Or objSerie.iTipoFormulario = TIPO_FORMULARIO_NFISCAL_SERVICO Then
        gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISVPV
    ElseIf objSerie.iTipoFormulario = TIPO_FORMULARIO_NFISCAL_FATURA Or objSerie.iTipoFormulario = TIPO_FORMULARIO_NFISCAL_FATURA_SERVICO Then
        gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV
    End If
    
    gobjGeracaoNFiscal.sSerie = sSeriePadrao
    Call CF2(Me, "PV_NF_DataSaida", gobjGeracaoNFiscal)
    
    objNFiscalInfo.dValorTotal = objNFiscal.dValorTotal
    objNFiscalInfo.iCodTransp = objNFiscal.iCodTransportadora
    objNFiscalInfo.iFilialCliente = objNFiscal.iFilialCli
    objNFiscalInfo.lCliente = objNFiscal.lCliente
    objNFiscalInfo.iFilialEmpresa = giFilialEmpresa
    objNFiscalInfo.iMarcada = MARCADO
    objNFiscalInfo.lPedido = objNFiscal.lNumPedidoVenda
    
    gobjGeracaoNFiscal.colNFiscalInfo.Add objNFiscalInfo

    lErro = GeracaoNF_Prepara_CTB()
    If lErro <> SUCESSO Then gError 59458
        
    lErro = CF("GeracaoNFiscal_GerarNFs", gobjGeracaoNFiscal)
    If lErro <> SUCESSO Then gError 62130
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Call Limpa_Tela_PedidoVenda

    Exit Sub

Erro_BotaoGerarNFiscal_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 59458, 62126, 62129, 62130, 62131, 64121, 64122

        Case 62127
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_JA_FATURADO_PARCIALMENTE", gErr, objPedidoVenda.lCodigo)
        
        Case 64123
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
            
        Case 64124
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CONFIGURADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158069)

    End Select

    Exit Sub

End Sub

Function PedidoVenda_Critica_Le(objPedidoVenda As ClassPedidoDeVenda) As Long
    
Dim lErro As Long

On Error GoTo Erro_PedidoVenda_Critica_Le
    
    'Critica PedidoVenda
    lErro = PedidoVenda_Critica(objPedidoVenda)
    If lErro <> SUCESSO Then gError 25871
    
    lErro = CF("PedidoDeVenda_Le_Completo", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 27633 Then gError 62114
    If lErro <> SUCESSO Then gError 62115
        
    'Lê as reservas do pedido de venda
    lErro = CF("ReservasPedido_Le", objPedidoVenda)
    If lErro <> SUCESSO Then gError 62118
    
    PedidoVenda_Critica_Le = SUCESSO

    Exit Function
    
Erro_PedidoVenda_Critica_Le:

    PedidoVenda_Critica_Le = gErr
    
    Select Case gErr
    
        Case 25871, 62114, 62118
        
        Case 62115
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoVenda.lCodigo)
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158070)
    
    End Select
    
    Exit Function
    
End Function

Function PedidoVenda_Critica(objPedidoVenda As ClassPedidoDeVenda) As Long
'Faz crítica do PedidoVenda
    
Dim lErro As Long

On Error GoTo Erro_PedidoVenda_Critica
    
    If Len(Trim(Codigo.Text)) = 0 Then gError 62113
    
    objPedidoVenda.lCodigo = StrParaLong(Codigo.Text)
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("PedidoVendaBaixado_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 46135 Then gError 62116
    If lErro = SUCESSO Then gError 62117
    
    PedidoVenda_Critica = SUCESSO

    Exit Function
    
Erro_PedidoVenda_Critica:

    PedidoVenda_Critica = gErr
    
    Select Case gErr
    
        Case 62113
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)
        
        Case 62116
        
        Case 62117
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_BAIXADO", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158071)
    
    End Select
    
    Exit Function
    
End Function

Private Sub Transfere_Dados_Pedido_NFiscal(objPedidoVenda As ClassPedidoDeVenda, objNFiscal As ClassNFiscal)
'Transfere os dados do objPedidoVenda para objNFiscal

Dim objComissaoPV As ClassComissaoPedVendas
Dim objComissaoNF As ClassComissaoNF
Dim objItemPV As ClassItemPedido
Dim dQuantidade As Double
Dim objReserva As ClassReserva
Dim colAlocacoes As ColAlocacoesItemNF
Dim dPercDesc As Double
Dim objItemTributacaoNF As ClassTributacaoDocItem
Dim objParcelaPV As ClassParcelaPedidoVenda
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lErro As Long
Dim iIndice As Integer
Dim objComissaoPorItem As ClassComissaoPorItem, objComissaoPorItemAux As ClassComissaoPorItem
    
On Error GoTo Erro_Transfere_Dados_Pedido_NFiscal
    
    With objPedidoVenda
        'Passa os dados gerais da tela de pedido para a Nota Fiscal
        objNFiscal.dPesoBruto = .dPesoBruto
        objNFiscal.dPesoLiq = .dPesoLiq
        objNFiscal.dVolumeTotal = .dVolumeTotal
        objNFiscal.dtDataEmissao = gdtDataAtual
        objNFiscal.dValorDesconto = .dValorDesconto
        objNFiscal.dValorFrete = .dValorFrete
        objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
        objNFiscal.dValorSeguro = .dValorSeguro
        objNFiscal.dValorTotal = .dValorTotal
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
        objNFiscal.sPlaca = .sPedidoCliente
        objNFiscal.sNaturezaOp = .sNaturezaOp
        objNFiscal.sNumPedidoTerc = .sPedidoCliente
        objNFiscal.sPlacaUF = .sPlacaUF
        objNFiscal.lVolumeEspecie = .lVolumeEspecie
        objNFiscal.lVolumeMarca = .lVolumeMarca
        objNFiscal.sVolumeNumero = .sVolumeNumero
        objNFiscal.iFreteRespons = .iFreteRespons
        objNFiscal.iCondicaoPagto = .iCondicaoPagto
        objNFiscal.dtDataReferencia = .dtDataReferencia
        objNFiscal.sSerie = SERIE_VERSAO_LIGHT
        objNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV
        
        Set objNFiscal.objProjetoInfo = .objProjetoInfo

    End With
    
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
    
    lErro = ParcelasPV_Gerando_ParcelasNF(objPedidoVenda, objNFiscal)
    If lErro <> SUCESSO Then gError 51614

    'Para cada item do Pedido
    For Each objItemPV In objPedidoVenda.colItensPedido
        
        iIndice = iIndice + 1
        
        'calcula a quantidade
        dQuantidade = objItemPV.dQuantidade - objItemPV.dQuantFaturada - objItemPV.dQuantCancelada

        If dQuantidade > 0 Then
            
            If objItemPV.dPrecoTotal > 0 Then
                'Calcula o percentual de desconto
                dPercDesc = objItemPV.dValorDesconto / objItemPV.dPrecoTotal
            End If
            
            Set colAlocacoes = New ColAlocacoesItemNF
            'Transporta as reservas do PV para a coleção de alocações da NF
            For Each objReserva In objItemPV.colReserva
                
                objAlmoxarifado.iCodigo = objReserva.iAlmoxarifado
                objAlmoxarifado.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
                'Lê o ALmoxarifado
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> 25056 And lErro <> SUCESSO Then gError 51560
                If lErro = 25056 Then gError 51561
                              
                If objAlmoxarifado.iFilialEmpresa = giFilialEmpresa Then colAlocacoes.Add objReserva.iAlmoxarifado, objAlmoxarifado.sNomeReduzido, dQuantidade
                
            Next
        
            Call gobjTribTab.ItemNFiscal_ObterTrib_ItemPedido(objItemPV, objItemTributacaoNF)
            
            With objItemPV
                                
                'Adiciona o item na coleção de itens da NF
                objNFiscal.ColItensNF.Add 0, iIndice, .sProduto, .sUnidadeMed, dQuantidade, .dPrecoUnitario, dPercDesc, .dValorDesconto, .dtDataEntrega, .sDescricao, .dValorAbatComissao, .lNumIntDoc, 0, 0, .lNumIntDoc, colAlocacoes, 0, "", "", .iStatus, 0, "", 0, 0, .iControleEstoque, .sUMEstoque, .iClasseUM, 0
                Set objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).objTributacaoItemNF = objItemTributacaoNF
            End With
        End If
    Next

    'Carrega os dados globais da tributação do PV para a NF
    Call gobjTribTab.NFiscal_ObterTrib_Pedido(objPedidoVenda, objNFiscal)
    
'    objNFiscal.objTributacaoNF.objTributacaoDesconto.iTipo = ITEM_TIPO_DESCONTO
'    objNFiscal.objTributacaoNF.objTributacaoFrete.iTipo = ITEM_TIPO_FRETE
'    objNFiscal.objTributacaoNF.objTributacaoOutras.iTipo = ITEM_TIPO_OUTRAS_DESP
'    objNFiscal.objTributacaoNF.objTributacaoSeguro.iTipo = ITEM_TIPO_SEGURO

    Exit Sub

Erro_Transfere_Dados_Pedido_NFiscal:

    Select Case gErr

        Case 51560

        Case 51561
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158072)

    End Select

    Exit Sub

End Sub

Function ParcelasPV_Gerando_ParcelasNF(objPedidoVenda As ClassPedidoDeVenda, objNFiscal As ClassNFiscal) As Long
'Se a Data de referencia do Pedido estiver ultrapassada em regalação a Data Atual e se
'a cobrança automática do pedido estiver selecionada as  parcelas são atualizadas.
    
Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
'Dim colDataVencimento As New Collection
Dim iIndice As Integer
Dim objParcelaPV As New ClassParcelaPedidoVenda
Dim dtDataVencimento As Date

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
        
        If objCondicaoPagto.iNumeroParcelas <> objPedidoVenda.colParcelas.Count Then gError 59097
        
        'Calcula Datas de Vencimento das Parcelas
        'lErro = CF("Parcelas_DatasVencimento", objCondicaoPagto, objNFiscal.dtDataReferencia, colDataVencimento)
        objCondicaoPagto.dtDataRef = objNFiscal.dtDataReferencia
        objCondicaoPagto.dtDataEntrega = StrParaDate(DataEntregaPV.Text)
        
        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True)
        If lErro <> SUCESSO Then gError 59098

        'Loop de preenchimento de Parcelas com Datas de Vencimento Recalculadas
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas

            Set objParcelaPV = objPedidoVenda.colParcelas.Item(iIndice)
            
            dtDataVencimento = objCondicaoPagto.colParcelas(iIndice).dtVencimento
            
            'Preenche o desconto padrão
            lErro = DescontoPadraoParcela_Gera(objParcelaPV, dtDataVencimento) 'colDataVencimento(iIndice))
            If lErro <> SUCESSO Then gError 51613
                        
            With objParcelaPV
                'Adiciona a parcela na coleçào de parcelas
                objNFiscal.ColParcelaReceber.Add 0, 0, .iNumParcela, STATUS_ABERTO, dtDataVencimento, dtDataVencimento, .dValor, .dValor, 1, CARTEIRA_CARTEIRA, COBRADOR_PROPRIA_EMPRESA, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0, 0, "", .dValor
            End With

        Next

    Else
    
        'transfere as parcelas do pedido de vendas para a nota fiscal
        For Each objParcelaPV In objPedidoVenda.colParcelas
            With objParcelaPV
                objNFiscal.ColParcelaReceber.Add 0, 0, .iNumParcela, STATUS_ABERTO, .dtDataVencimento, .dtDataVencimento, .dValor, .dValor, 1, CARTEIRA_CARTEIRA, COBRADOR_PROPRIA_EMPRESA, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0, 0, "", .dValor
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158073)
    
    End Select
    
    Exit Function
    
End Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158074)
            
    End Select

    Exit Function
    
End Function

Public Sub BotaoNFiscal_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objNFiscal As New ClassNFiscal
Dim sSeriePadrao As String

On Error GoTo Erro_BotaoNFiscal_Click

    lErro = PedidoVenda_Critica(objPedidoVenda)
    If lErro <> SUCESSO Then gError 62132
    
    'Lê a Serie Padrao para gerar NF
    lErro = CF("Serie_Le_Padrao", sSeriePadrao)
    If lErro <> SUCESSO Then gError 25955
    
    'Testa se a Serie Padrão está cadastrada
    If Trim(sSeriePadrao) = "" Then gError 25956
    
    Call Chama_Tela("NFiscalFatura", objNFiscal, objPedidoVenda)

    Exit Sub
    
Erro_BotaoNFiscal_Click:

    Select Case gErr
    
        Case 25955, 62132
        
        Case 25956
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CONFIGURADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158075)
            
    End Select
    
    Exit Sub

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

Private Function GeracaoNF_Prepara_CTB() As Long
'prepara informacoes necessarias para a contabilizacao

Dim lErro As Long, objPeriodo As New ClassPeriodo

On Error GoTo Erro_GeracaoNF_Prepara_CTB

    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
    
        gobjGeracaoNFiscal.objTelaAtualizacao = Me
        gobjGeracaoNFiscal.dtContabil = gdtDataAtual
        gobjGeracaoNFiscal.iLoteContabil = 0
    
        'Coloca o periodo relativo a data na tela
        lErro = CF("Periodo_Le", gobjGeracaoNFiscal.dtContabil, objPeriodo)
        If lErro <> SUCESSO Then gError 59457
        
        giPeriodo = objPeriodo.iPeriodo
        giExercicio = objPeriodo.iExercicio

    End If
    
    GeracaoNF_Prepara_CTB = SUCESSO
     
    Exit Function
    
Erro_GeracaoNF_Prepara_CTB:

    GeracaoNF_Prepara_CTB = gErr
     
    Select Case gErr
          
        Case 59457
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158076)
     
    End Select
     
    Exit Function

End Function

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
Dim objEndereco As New ClassEndereco
Dim colPVMotivo As New Collection
Dim sTexto As String, iFrete As Integer, sFrete As String
Dim objVendedor As New ClassVendedor
Dim objTipoFrete As ClassTipoFrete
Dim colTipoFrete As Collection

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
        
        Case CTACONTABILEST1 'parametros: produto no formato da tela do grid de itens, produto no formato da tela do grid de alocacoes e nome reduzido do almoxarifado
            
            For Each objItem In gobjNFiscalCTB.ColItensNF
                
                For Each objAlocacao In objItem.colAlocacoes
                
                    If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
                            
                        objAlmoxarifado.sNomeReduzido = CStr(objMnemonicoValor.vParam(3))

                        'Lê o Nome Reduzido do Almoxarifado
                        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                        If lErro <> SUCESSO And lErro <> 25060 Then gError 59388

                        'Se não encontrou ===> Erro
                        If lErro = SUCESSO Then

                            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
                            
                            lErro = CF("Produto_Formata", objMnemonicoValor.vParam(1), sProdutoFormatado, iPreenchido)
                            If iPreenchido <> PRODUTO_PREENCHIDO Then gError 59390

                            objEstoqueProduto.sProduto = sProdutoFormatado
                            
                            'Lê a conta contábil do Produto no Almoxarifado
                            lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
                            If lErro <> SUCESSO And lErro <> 49991 Then gError 59391

                            If Len(Trim(objEstoqueProduto.sContaContabil)) > 0 Then

                                sContaMascarada = String(STRING_CONTA, 0)

                                lErro = Mascara_RetornaContaTela(objEstoqueProduto.sContaContabil, sContaMascarada)
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
                
                Next
                
            Next
                    
        Case CODIGO1
            objMnemonicoValor.colValor.Add gobjPedidoVendaCTB.lCodigo
        
        Case QUANT_ESTOQUE
        
            For Each objItem In gobjNFiscalCTB.ColItensNF
            
                lErro = CF("UMEstoque_Conversao", objItem.sProduto, objItem.sUnidadeMed, objItem.dQuantidade, dQuantidadeConvertida)
                If lErro <> SUCESSO Then gError 64214
                
                objMnemonicoValor.colValor.Add dQuantidadeConvertida
                
            Next
            
        Case ALMOX1
        
            If gcolAlmoxFilial.Count = 0 Then
            
                lErro = CF("Almoxarifados_Le_FilialEmpresa", gobjNFiscalCTB.iFilialEmpresa, gcolAlmoxFilial)
                If lErro <> SUCESSO Then gError 59423
                
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
                If lErro <> SUCESSO Then gError 39586
                
                If objFilialCliente.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilialCliente.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError 56500
                
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
            If lErro <> SUCESSO And lErro <> 12567 Then gError 59392
            If lErro <> SUCESSO Then gError 59393
            
            objMnemonicoValor.colValor.Add objFilialCliente.sNome
            
        Case CLIENTE1
        
            objCodigoNome.lCodigo = gobjNFiscalCTB.lCliente
            
            lErro = CF("Cliente_Le_NomeRed", objCodigoNome)
            If lErro <> SUCESSO And lErro <> 12553 Then gError 59394
            If lErro <> SUCESSO Then gError 59395
                        
            objMnemonicoValor.colValor.Add objCodigoNome.sNome
            
        Case NATUREZA_OP
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.sNaturezaOp

        Case NFISCAL1
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.lNumNotaFiscal

        Case PRODUTO1
            For Each objItem In gobjNFiscalCTB.ColItensNF
            
                lErro = Mascara_RetornaProdutoTela(objItem.sProduto, sProdutoTela)
                If lErro <> SUCESSO Then gError 59397
                
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

        Case ICMS_DIFAL_REMETENTE
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.objTributacaoNF.dICMSInterestVlrUFRemet

        Case ICMS_DIFAL_DESTINO
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.objTributacaoNF.dICMSInterestVlrUFDest

        Case ICMS_FCP
            objMnemonicoValor.colValor.Add gobjNFiscalCTB.objTributacaoNF.dICMSInterestVlrFCPUFDest
        
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
                    If lErro <> SUCESSO Then gError 59409
                    
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
            If lErro <> SUCESSO And lErro <> 12553 Then gError 178248
            
            If lErro <> SUCESSO Then gError 178249
                        
            objMnemonicoValor.colValor.Add objCodigoNome.sNome
            
        Case CPF_CNPJ
        
            objFilialCliente.lCodCliente = objContexto.lCliente
            objFilialCliente.iCodFilial = objContexto.iFilial
            
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 186569
            
            If lErro <> SUCESSO Then gError 186570
            
            objMnemonicoValor.colValor.Add objFilialCliente.sCgc
            
        Case EMAIL_VEND_CLI
            objFilialCliente.lCodCliente = objContexto.lCliente
            objFilialCliente.iCodFilial = objContexto.iFilial
            
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 186569
            
            If lErro <> SUCESSO Then gError 186570
        
            If objFilialCliente.iVendedor = 0 Then
                objMnemonicoValor.colValor.Add ""
            Else
                objVendedor.iCodigo = objFilialCliente.iVendedor
                
                lErro = CF("Vendedor_Le", objVendedor)
                If lErro <> SUCESSO And lErro <> 12582 Then gError 186569
                
                If lErro <> SUCESSO Then gError 201230
                
                objEndereco.lCodigo = objVendedor.lEndereco
                
                lErro = CF("Endereco_Le", objEndereco)
                If lErro <> SUCESSO And lErro <> 12309 Then gError 186573
                
                objMnemonicoValor.colValor.Add objEndereco.sEmail & IIf(objEndereco.sEmail2 <> "", ";" & objEndereco.sEmail2, "")
                
            End If
            
        Case TIPO_FRETE_CLIENTE, TIPO_FRETE_CLIENTE_DESC
        
            objFilialCliente.lCodCliente = objContexto.lCliente
            objFilialCliente.iCodFilial = objContexto.iFilial
            
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 186569
            
            If lErro <> SUCESSO Then gError 186570
            
            lErro = CF("TipoFrete_Le", colTipoFrete)
            If lErro <> SUCESSO Then gError 186569
            
            iFrete = -1
            sFrete = ""
            For Each objTipoFrete In colTipoFrete
                If objTipoFrete.iCodigoNFE = objFilialCliente.iTipoFrete Then
                    iFrete = objTipoFrete.iCodigoNFE
                    sFrete = objTipoFrete.sDescricao
                    Exit For
                End If
            Next
            
            If objMnemonicoValor.sMnemonico = TIPO_FRETE_CLIENTE_DESC Then
                objMnemonicoValor.colValor.Add sFrete
            Else
                objMnemonicoValor.colValor.Add iFrete
            End If
            
        Case TIPO_FRETE_PEDIDO, TIPO_FRETE_PEDIDO_DESC
        
            lErro = CF("TipoFrete_Le", colTipoFrete)
            If lErro <> SUCESSO Then gError 186569
        
            iFrete = -1
            sFrete = ""
            For Each objTipoFrete In colTipoFrete
                If objTipoFrete.iCodigo = objContexto.iFreteRespons Then
                    iFrete = objTipoFrete.iCodigoNFE
                    sFrete = objTipoFrete.sDescricao
                    Exit For
                End If
            Next
        
            If objMnemonicoValor.sMnemonico = TIPO_FRETE_PEDIDO_DESC Then
                objMnemonicoValor.colValor.Add sFrete
            Else
                objMnemonicoValor.colValor.Add iFrete
            End If
        
        Case UF_PRINCIPAL
        
            objFilialCliente.lCodCliente = objContexto.lCliente
            objFilialCliente.iCodFilial = objContexto.iFilial
            
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 186571
            
            If lErro <> SUCESSO Then gError 186572
            
            objEndereco.lCodigo = objFilialCliente.lEndereco
            
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO And lErro <> 12309 Then gError 186573
            
            objMnemonicoValor.colValor.Add objEndereco.sSiglaEstado
            
        Case CONDPAGTO
            objMnemonicoValor.colValor.Add objContexto.iCondicaoPagto
            
        Case EXISTE_PRECOMENORTABELA
        
            lErro = CF("PV_Preco_Menor_Tabela", objContexto, colPVMotivo)
            If lErro <> SUCESSO Then gError 197763

            If colPVMotivo.Count > 0 Then
                iAchou = 1
            End If

            objMnemonicoValor.colValor.Add iAchou

        
        
        Case PRECOMENORTABELA
        
            lErro = CF("PV_Preco_Menor_Tabela_Texto", objContexto, sTexto)
            If lErro <> SUCESSO Then gError 197764
        
            objMnemonicoValor.colValor.Add sTexto
        

        Case Else
            gError 59389

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 59389
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 59388, 59390, 59391, 59392, 59394, 59397, 59409, 59410, 59423, 178248, 186569, 186571, 186573, 197763, 197764
        
        Case 59393, 186570, 186572
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objFilialCliente.iCodFilial, objFilialCliente.lCodCliente)
        
        Case 59395
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, gobjNFiscalCTB.lCliente)
        
        Case 201230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, objVendedor.iCodigo)
        
        Case 178249
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objContexto.lCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158077)

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
    If lErro <> SUCESSO Then gError 59398
    
    'grava a contabilizacao
    lErro = objContabAutomatica.Gravar_Registro(Me, IIf(gobjNFiscalCTB.iTipoNFiscal = DOCINFO_NFISFVPV, "NFiscalFaturaPedido", "NFiscalPedido"), gobjNFiscalCTB.lNumIntDoc, gobjNFiscalCTB.lCliente, gobjNFiscalCTB.iFilialCli, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjNFiscalCTB.iFilialEmpresa, gobjGeracaoNFiscal.iLoteContabil, gobjNFiscalCTB.lNumNotaFiscal)
    If lErro <> SUCESSO Then gError 59399

    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = gErr
     
    Select Case gErr
          
        Case 59398, 59399
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158078)
     
    End Select
     
    Exit Function

End Function

Public Sub GridEmb_Click()
'embalagem
    
    Call gobjEmbalagem.GridEmb_Click

End Sub

Public Sub GridEmb_EnterCell()
'embalagem
    
    Call gobjEmbalagem.GridEmb_EnterCell

End Sub

Public Sub GridEmb_GotFocus()
'embalagem
    
    Call gobjEmbalagem.GridEmb_GotFocus

End Sub

Public Sub GridEmb_KeyPress(KeyAscii As Integer)
'embalagem
    
    Call gobjEmbalagem.GridEmb_KeyPress(KeyAscii)

End Sub

Public Sub GridEmb_LeaveCell()
'embalagem
    
    Call gobjEmbalagem.GridEmb_LeaveCell

End Sub

Public Sub GridEmb_Validate(Cancel As Boolean)
'embalagem
    
    Call gobjEmbalagem.GridEmb_Validate(Cancel)
    
End Sub

Public Sub GridEmb_RowColChange()
'embalagem
    
    Call gobjEmbalagem.GridEmb_RowColChange

End Sub

Public Sub GridEmb_KeyDown(KeyCode As Integer, Shift As Integer)
'embalagem
    
    Call gobjEmbalagem.GridEmb_KeyDown(KeyCode, Shift)
    
End Sub

Public Sub GridEmb_Scroll()
'embalagem
    
    Call gobjEmbalagem.GridEmb_Scroll

End Sub

Public Sub ItemEmb_KeyPress(KeyAscii As Integer)
'embalagem

    Call gobjEmbalagem.ItemEmb_KeyPress(KeyAscii)

End Sub

Public Sub ItemEmb_Validate(Cancel As Boolean)
'embalagem

    Call gobjEmbalagem.ItemEmb_Validate(Cancel)

End Sub

Public Sub ItemEmb_Change()
'Embalagem

    Call gobjEmbalagem.ItemEmb_Change

End Sub

Public Sub ItemEmb_GotFocus()
'Embalagem

    Call gobjEmbalagem.ItemEmb_GotFocus

End Sub

'Janaina
Public Sub BotaoEmbalagens_Click()
    Call gobjEmbalagem.BotaoEmbalagens_Click
End Sub

Public Sub Embalagem_KeyPress(KeyAscii As Integer)
'embalagem

    Call gobjEmbalagem.Embalagem_KeyPress(KeyAscii)

End Sub

Public Sub Embalagem_Validate(Cancel As Boolean)
'embalagem

    Call gobjEmbalagem.Embalagem_Validate(Cancel)

End Sub

Public Sub Embalagem_Change()
'Embalagem

    Call gobjEmbalagem.Embalagem_Change

End Sub

Public Sub Embalagem_GotFocus()
'Embalagem

    Call gobjEmbalagem.Embalagem_GotFocus

End Sub

Public Sub QuantEmb_KeyPress(KeyAscii As Integer)
'embalagem

    Call gobjEmbalagem.QuantEmb_KeyPress(KeyAscii)

End Sub

Public Sub QuantEmb_Validate(Cancel As Boolean)
'embalagem

    Call gobjEmbalagem.QuantEmb_Validate(Cancel)

End Sub

Public Sub QuantEmb_Change()
'Embalagem

    Call gobjEmbalagem.QuantEmb_Change

End Sub

Public Sub QuantEmb_GotFocus()
'Embalagem

    Call gobjEmbalagem.QuantEmb_GotFocus

End Sub

Public Sub QuantProduto_KeyPress(KeyAscii As Integer)
'embalagem

    Call gobjEmbalagem.QuantProduto_KeyPress(KeyAscii)

End Sub

Public Sub QuantProduto_Validate(Cancel As Boolean)
'embalagem

    Call gobjEmbalagem.QuantProduto_Validate(Cancel)

End Sub

Public Sub QuantProduto_Change()
'Embalagem

    Call gobjEmbalagem.QuantProduto_Change

End Sub

Public Sub QuantProduto_GotFocus()
'Embalagem

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

'Leo daqui p/ baixo em 26/03/02

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158080)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158081)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158082)

    End Select

    Exit Sub

End Sub

Public Sub OrcamentoLabel_Click()

Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim colSelecao As New Collection

    Call Chama_Tela("OrcamentoPedidoLista", colSelecao, objOrcamentoVenda, objEventoOrcamento)

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
    If Len(Trim(Orcamento.Text)) = 0 Then gError 94427
    If Len(Trim(FilialOrcamento.Text)) = 0 Then gError 94428
    
    objOcamentoVenda.lCodigo = StrParaLong(Orcamento.Text)
    objOcamentoVenda.iFilialEmpresa = Codigo_Extrai(FilialOrcamento.Text)
        
    'Busca o Pedido na tabela de pedidos de venda
    lErro = CF("OrcamentoVenda_Le", objOcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101232 Then gError 94429
    
    If lErro = 101232 Then gError 94430 'Não encontrou, erro.
    
    'Chama  a tela de Pedido de Venda passando o pedido de venda da tela
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158083)

    End Select

    Exit Sub

End Sub

Public Sub BotaoOrcamento_Click()
             
Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim objAnotacoes As New ClassAnotacoes
Dim sTextoMontado As String

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

    'Trata clientes não cadastrados no orçamento
    Call Cadastra_Cliente_Inexistente(objOrcamentoVenda) 'Luiz 26/01/04
    
    objAnotacoes.iTipoDocOrigem = ANOTACAO_ORIGEM_ORCVENDA
    
    objAnotacoes.sID = CStr(giFilialEmpresa) & "," & objOrcamentoVenda.lCodigo
    
    'Verifica no bd se ja existe alguma anotação referente ao Id/ Origem
    lErro = CF("Anotacoes_Le_OrigemID", objAnotacoes)
    If lErro <> SUCESSO And lErro <> 119844 Then gError 210445
    
    If lErro = SUCESSO Then
    
        'busca o texto da anotação que foi "montado" a partir do numintdoc
        lErro = CF("AnotacoesLinha_Le_Texto1", sTextoMontado, objAnotacoes.lNumIntDoc)
        If lErro <> SUCESSO And lErro <> 119875 Then gError 210446
        
        If lErro = SUCESSO Then
            Set gobjAnotacao = New ClassAnotacoes
        
            gobjAnotacao.sTitulo = objAnotacoes.sTitulo
            gobjAnotacao.sTextoCompleto = sTextoMontado
        End If
    
    End If
    
    Exit Sub

Erro_BotaoOrcamento_Click:

    Select Case gErr

        Case 94427
            Call Rotina_Erro(vbOKOnly, "ERRO_CODORCAMENTO_NAO_INFORMADO", gErr)

        Case 94428
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALORCAMENTO_NAO_INFORMADA", gErr)

        Case 94429, 94432, 210445, 210446

        Case 94425
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_VINCULADO_PEDVENDA1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
        
        Case 94426
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_VINCULADO_NFISCAL1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
        
        Case 94430
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case 94431
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFATURAMENTO_DIFERENTE1", gErr, objOrcamentoVenda.iFilialEmpresa, objOrcamentoVenda.lCodigo, giFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158084)

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
    'objPedidoVenda.dtDataEntrega = DATA_NULA
    
    'Carrega os dados de PedidoVenda na tela passando a constante que identifica como oriundo de Orçamento Venda.
    lErro = Traz_Pedido_Tela(objPedidoVenda, objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 94434
        
    lErro = CF2(Me, "PV_Traz_Dados_Orcamento_Tela", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 141415
        
    Traz_Dados_Orcamento_Tela = SUCESSO

    Exit Function

Erro_Traz_Dados_Orcamento_Tela:

    Traz_Dados_Orcamento_Tela = gErr

    Select Case gErr

        Case 94432, 94434, 141415

        Case 94433
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158085)

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
Dim objKitVenda As New ClassKitVenda
Dim objProjetoInfo As Object, dtDataBase As Date, dtDataAux As Date
Dim dCotacao As Double
Dim objCotacao As New ClassCotacaoMoeda
Dim objCotacaoAnterior As New ClassCotacaoMoeda
Dim vbMsgResult As VbMsgBoxResult

On Error GoTo Erro_Transfere_Dados_Pedido_NFiscal

    Call gobjTelaProjetoInfo.Obtem_ProjetoInfo(objProjetoInfo, objOrcamentoVenda.lNumIntDoc, PRJ_CR_TIPO_OV)
    Set objPedidoVenda.objProjetoInfo = objProjetoInfo

    Set objPedidoVenda.objInfoAdic = objOrcamentoVenda.objInfoAdic
    
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

    With objPedidoVenda
        
        'AKI!!! Transferir o resto dos campos
        
        'Passa os dados gerais da tela de Orcamento para o Pedido
        objOrcamentoVenda.dtDataEmissao = gdtDataAtual
        
        .dtDataEmissao = objOrcamentoVenda.dtDataEmissao
        .dtDataReferencia = objOrcamentoVenda.dtDataReferencia
        .dtDataRefFluxo = objOrcamentoVenda.dtDataReferencia
        .dValorDesconto = Arredonda_Moeda(objOrcamentoVenda.dValorDesconto * dCotacao)
        .dValorFrete = Arredonda_Moeda(objOrcamentoVenda.dValorFrete * dCotacao)
        .dValorOutrasDespesas = Arredonda_Moeda(objOrcamentoVenda.dValorOutrasDespesas * dCotacao)
        .dValorSeguro = Arredonda_Moeda(objOrcamentoVenda.dValorSeguro * dCotacao)
        .dValorTotal = objOrcamentoVenda.dValorTotal '* dCotacao
        .lNumIntDoc = objOrcamentoVenda.lNumIntPedVenda
        .iNaturezaOpManual = objOrcamentoVenda.iNaturezaOpManual
        .iCobrancaAutomatica = objOrcamentoVenda.iCobrancaAutomatica
        .dValorProdutos = objOrcamentoVenda.dValorProdutos '* dCotacao
        .dPercAcrescFinanceiro = objOrcamentoVenda.dPercAcrescFinanceiro
        .iFilial = objOrcamentoVenda.iFilial
        .iFilialEntrega = objOrcamentoVenda.iFilialEntrega
        .iFilialEmpresaFaturamento = objOrcamentoVenda.iFilialEmpresa
        .iTabelaPreco = objOrcamentoVenda.iTabelaPreco
        .lCliente = objOrcamentoVenda.lCliente
        .sNaturezaOp = objOrcamentoVenda.sNaturezaOp
        .iCondicaoPagto = objOrcamentoVenda.iCondicaoPagto
        .dValorDescontoItens = Arredonda_Moeda(objOrcamentoVenda.dValorDescontoItens * dCotacao)
        .dValorDescontoTit = Arredonda_Moeda(objOrcamentoVenda.dValorDescontoTit * dCotacao)
        .dValorItens = objOrcamentoVenda.dValorItens '* dCotacao
        
        .dtDataEntrega = objOrcamentoVenda.dtDataEntrega
        .iCodTransportadora = objOrcamentoVenda.iCodTransportadora
        .iCodTranspRedesp = objOrcamentoVenda.iCodTranspRedesp
        .sMensagemNota = objOrcamentoVenda.sMensagemNota
        .sPedidoCliente = objOrcamentoVenda.sPedidoCliente
        .sPedidoRepresentante = objOrcamentoVenda.sPedidoRepresentante
        .iCanalVenda = objOrcamentoVenda.iCanalVenda
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
    
    dtDataBase = StrParaDate(DataEmissao.Text)
    If dtDataBase = DATA_NULA Then dtDataBase = gdtDataAtual
    
    If objPedidoVenda.dtDataEntrega = DATA_NULA Then
        If objOrcamentoVenda.iDataEnt = OV_DATA_ENTREGA_PRAZO_DIAS_UTEIS Then
            lErro = CF("Calcula_Data", dtDataBase, objOrcamentoVenda.iPrazoEntrega, 0, CONDPAGTO_TIPOINTERVALO_DIAS_UTEIS, dtDataAux)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            objPedidoVenda.dtDataEntrega = dtDataAux
        ElseIf objOrcamentoVenda.iDataEnt = OV_DATA_ENTREGA_PRAZO_DIAS_CORRIDOS Then
            dtDataAux = DateAdd("d", objOrcamentoVenda.iPrazoEntrega, dtDataBase)
            objPedidoVenda.dtDataEntrega = dtDataAux
        ElseIf objOrcamentoVenda.iDataEnt = OV_DATA_ENTREGA_PRAZO_SEMANAS Then
            dtDataAux = DateAdd("d", objOrcamentoVenda.iPrazoEntrega * 7, dtDataBase)
            objPedidoVenda.dtDataEntrega = dtDataAux
        ElseIf objOrcamentoVenda.iDataEnt = OV_DATA_ENTREGA_PRAZO_MESES Then
            dtDataAux = DateAdd("m", objOrcamentoVenda.iPrazoEntrega, dtDataBase)
            objPedidoVenda.dtDataEntrega = dtDataAux
        End If
    End If

    'Para cada item do Pedido
    For Each objItemOV In objOrcamentoVenda.colItens
        
        Set objItemPV = New ClassItemPedido
        
        Call gobjTribTab.ItemPedidoVenda_ObterTrib_ItemOrcamento(objItemPV, objItemOV)
        
        With objItemOV
        
            Set objItemPV = objPedidoVenda.IncluirItem(.sProduto, .dQuantidade, IIf(dCotacao = 1, .dPrecoUnitario, Arredonda_Moeda(.dPrecoUnitario * dCotacao)), IIf(dCotacao = 1, .dPrecoTotal, Arredonda_Moeda(.dPrecoTotal * dCotacao)), IIf(dCotacao = 1, .dValorDesconto, Arredonda_Moeda(.dValorDesconto * dCotacao)), .dtDataEntrega, .sProdutoDescricao, 0, 0, 0, Nothing, "", .sUnidadeMed, "", "", 0, 0, 0, .sDescricao, STATUS_ABERTO, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, .objTributacaoItemOV)
            
            If objItemOV.colItensRomaneioGrade.Count > 0 Then objItemPV.iPossuiGrade = MARCADO
            
            Set objItemPV.objInfoAdicDocItem = objItemOV.objInfoAdicDocItem
            
            If objItemPV.dtDataEntrega = DATA_NULA Then
                If objOrcamentoVenda.iDataEnt = OV_DATA_ENTREGA_PRAZO_DIAS_UTEIS Then
                    lErro = CF("Calcula_Data", dtDataBase, objItemOV.iPrazoEntrega, 0, CONDPAGTO_TIPOINTERVALO_DIAS_UTEIS, dtDataAux)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    objItemPV.dtDataEntrega = dtDataAux
                ElseIf objOrcamentoVenda.iDataEnt = OV_DATA_ENTREGA_PRAZO_DIAS_CORRIDOS Then
                    dtDataAux = DateAdd("d", objItemOV.iPrazoEntrega, dtDataBase)
                    objItemPV.dtDataEntrega = dtDataAux
                ElseIf objOrcamentoVenda.iDataEnt = OV_DATA_ENTREGA_PRAZO_SEMANAS Then
                    dtDataAux = DateAdd("d", objItemOV.iPrazoEntrega * 7, dtDataBase)
                    objItemPV.dtDataEntrega = dtDataAux
                ElseIf objOrcamentoVenda.iDataEnt = OV_DATA_ENTREGA_PRAZO_MESES Then
                    dtDataAux = DateAdd("m", objItemOV.iPrazoEntrega, dtDataBase)
                    objItemPV.dtDataEntrega = dtDataAux
                End If
            End If
                                  
            Set objProduto = New ClassProduto
            
            objProduto.sCodigo = objItemOV.sProduto
                        
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 180030
                                    
            If objProduto.iKitVendaComp = MARCADO Then
            
                Set colItensRomaneioGrade = New Collection
            
                lErro = CF("ProdutoKitVenda_Le_Romaneio", objProduto, colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 180031
                
                objKitVenda.sProduto = objItemOV.sProduto
                objKitVenda.sUM = objItemOV.sUnidadeMed
                objKitVenda.dQuantidade = objItemOV.dQuantidade
                
                lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 180032
                
                Set objItemOV.colItensRomaneioGrade = colItensRomaneioGrade
            
            End If
                        
            lErro = CF2(Me, "PV_Transfere_Dados_Orcamento_Pedido", objItemOV, objItemPV)
            If lErro <> SUCESSO Then gError 177225
            
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
            objPedidoVenda.colParcelas.Add .dValor * dCotacao, .dtDataVencimento, .iNumParcela, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .dtDesconto3Ate, .dDesconto3Valor, .iDesconto3Codigo
        End With
    Next

    'Reinicializa as coleções de itens
    Set gcolItemPedido = New Collection

    Exit Sub

Erro_Transfere_Dados_Pedido_NFiscal:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 51562, 177225, 180030, 180031
        
        Case 211630
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_SEM_COTACAO", gErr, objCotacao.iMoeda)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158086)

    End Select

End Sub

Sub Atualiza_Grid_Alocacao(objItemPV As ClassItemPedido)

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReserva As ClassReservaItem
Dim sProdutoMascarado As String
Dim lErro As Long
Dim dFator As Double
Dim dFator2 As Double
Dim objProduto As New ClassProduto
Dim dQuantReservada As Double, dQuantReservadaItem As Double
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Atualiza_Grid_Alocacao

    Call Remove_Reservas(objItemPV.iItem)

    dQuantReservadaItem = 0
    For Each objItemRomaneio In objItemPV.colItensRomaneioGrade

        objProduto.sCodigo = objItemRomaneio.sProduto
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
        If lErro = 28030 Then gError 42765 'Não encontrou
        'Faz a conversão da unidade do item para a unidade de estoque
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRomaneio.sUMEstoque, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 42766
        
        objProduto.sCodigo = objItemPV.sProduto
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
        If lErro = 28030 Then gError 42765 'Não encontrou
        
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemPV.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator2)
        If lErro <> SUCESSO Then gError 42766
        
        dQuantReservada = 0
        
        For Each objReserva In objItemRomaneio.colLocalizacao
        
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Item_Col) = objItemPV.iItem
            
            lErro = Mascara_MascararProduto(objItemRomaneio.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 86357
                        
            objAlmoxarifado.iCodigo = objReserva.iAlmoxarifado

            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 46181
            If lErro = 25056 Then gError 46182
                
            objReserva.sAlmoxarifado = objAlmoxarifado.sNomeReduzido
                        
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_ProdutoAlmox_Col) = sProdutoMascarado
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_UMEstoque_Col) = objItemRomaneio.sUMEstoque
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objReserva.sAlmoxarifado
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReservar_Col) = Formata_Estoque((objItemRomaneio.dQuantidade - objItemRomaneio.dQuantCancelada) * dFator2)
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_QuantReserv_Col) = Formata_Estoque(objReserva.dQuantidade * dFator)
            If objReserva.dtDataValidade <> DATA_NULA Then GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Validade_Col) = Format(objReserva.dtDataValidade, "dd/mm/yyyy")
            GridAlocacao.TextMatrix(objGridAlocacao.iLinhasExistentes + 1, iGrid_Responsavel_Col) = objReserva.sResponsavel
            
            dQuantReservadaItem = dQuantReservadaItem + StrParaDbl(Formata_Estoque(objReserva.dQuantidade * dFator))
            
            objGridAlocacao.iLinhasExistentes = objGridAlocacao.iLinhasExistentes + 1
            
        Next
               
    Next
    
    If dQuantReservadaItem > 0 Then
        GridItens.TextMatrix(objItemPV.iItem, iGrid_QuantReservada_Col) = Formata_Estoque(dQuantReservadaItem)
    Else
        GridItens.TextMatrix(objItemPV.iItem, iGrid_QuantReservada_Col) = ""
    End If

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
    
    For Each objItemRomaneioGrade In objItemPV.colItensRomaneioGrade
            
        dQuantidade = dQuantidade + objItemRomaneioGrade.dQuantidade
        dQuantCancelada = dQuantCancelada + objItemRomaneioGrade.dQuantCancelada
        dQuantReservada = dQuantReservada + objItemRomaneioGrade.dQuantReservada
        dQuantFaturada = dQuantFaturada + objItemRomaneioGrade.dQuantFaturada
        
    Next

    GridItens.TextMatrix(objItemPV.iItem, iGrid_QuantPedida_Col) = Formata_Estoque(dQuantidade)
    If iGrid_QuantCancel_Col <> 0 Then GridItens.TextMatrix(objItemPV.iItem, iGrid_QuantCancel_Col) = Formata_Estoque(dQuantCancelada)
    GridItens.TextMatrix(objItemPV.iItem, iGrid_QuantReservada_Col) = Formata_Estoque(dQuantReservada)
    GridItens.TextMatrix(objItemPV.iItem, iGrid_QuantFaturada_Col) = Formata_Estoque(dQuantFaturada)

    objItemPV.dQuantidade = dQuantidade
    objItemPV.dQuantCancelada = dQuantCancelada

    Exit Sub

End Sub

Public Sub BotaoInfoAdicItem_Click()

Dim lErro  As Long
Dim objItemPV As ClassItemPedido
Dim objTela As Object

On Error GoTo Erro_BotaoInfoAdicItem_Click

    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
    
        Set objItemPV = gobjPedidoVenda.colItensPedido(GridItens.Row)
        
        Set objTela = Me
        
        objItemPV.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMPV
        objItemPV.objInfoAdicDocItem.iItem = GridItens.Row
        objItemPV.objInfoAdicDocItem.sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        objItemPV.objInfoAdicDocItem.sDescProd = GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col)
        
        objItemPV.objInfoAdicDocItem.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
        objItemPV.objInfoAdicDocItem.dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col)) - StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantCancel_Col))
        objItemPV.objInfoAdicDocItem.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))
        objItemPV.objInfoAdicDocItem.dValorTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))
        
        Call Chama_Tela_Modal("InfoAdicDocItem", objItemPV.objInfoAdicDocItem, objTela, TIPO_SAIDA, gobjInfoAdic)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Call CF2(Me, "PV_PrecoTotal_Calcula", GridItens.Row)
        
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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

Public Sub BotaoGrade_Click()

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim lErro  As Long
Dim objRomaneioGrade As ClassRomaneioGrade
Dim objItemPV As ClassItemPedido

On Error GoTo Erro_BotaoGrade_Click

    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
    
        Set objItemPV = gobjPedidoVenda.colItensPedido(GridItens.Row)
        
        If objItemPV.iPossuiGrade = MARCADO Then
            
            Set objRomaneioGrade = New ClassRomaneioGrade
            
            objRomaneioGrade.sNomeTela = Me.Name
            objRomaneioGrade.iFilialFaturamento = Codigo_Extrai(FilialFaturamento.Text)
            Set objRomaneioGrade.objObjetoTela = objItemPV
                        
            Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
            If lErro <> SUCESSO Then gError 86360
        
            Call Atualiza_Grid_Itens(objItemPV)
            
            Call Atualiza_Grid_Alocacao(objItemPV)
        
            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) > 0 Then
                Call CF2(Me, "PV_PrecoTotal_Calcula", GridItens.Row)
                
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158088)
            
    End Select
    
    Exit Sub

End Sub

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

    For iIndice = 1 To gobjPedidoVenda.colItensPedido.Count
        If gobjPedidoVenda.colItensPedido(iIndice).iPossuiGrade = MARCADO Then
            Grid_Possui_Grade = True
            Exit Function
        End If
    Next
    
    Grid_Possui_Grade = False
        
    Exit Function
    
End Function

Private Function Reserva_Processa_ItemGrade(objItemRomaneioGrade As ClassItemRomaneioGrade, ByVal iItem As Integer, Optional ByVal bIgnoraConfig As Boolean = False) As Long

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************


Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dQuantidade As Double, dQuantidadeCancelada As Double, dQuantidadeFaturada As Double

On Error GoTo Erro_Reserva_Processa_ItemGrade

    dQuantidade = objItemRomaneioGrade.dQuantidade
    dQuantidadeCancelada = objItemRomaneioGrade.dQuantCancelada
    dQuantidadeFaturada = objItemRomaneioGrade.dQuantFaturada

    'Se não houve Quant a Reservar zera Reservas.
    If dQuantidade = 0 Or dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada = 0 Then
        objItemRomaneioGrade.dQuantReservada = 0
    End If

    objProduto.sCodigo = objItemRomaneioGrade.sProduto

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 26669

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 26668
    
    objItemRomaneioGrade.sUMEstoque = objProduto.sSiglaUMEstoque

    'Se controla Reserva e há quantidade a reservar gera reserva automático
    If objProduto.iControleEstoque = PRODUTO_CONTROLE_RESERVA And dQuantidade > 0 And dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada > 0 And (gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA Or bIgnoraConfig) Then
        
        'tenta reservar no Almoxarifado padrão
        lErro = ReservaAlmoxarifadoPadrao_Grade(objProduto, objItemRomaneioGrade, dQuantidade, dQuantidadeCancelada, dQuantidadeFaturada, iItem)
        If lErro <> SUCESSO And lErro <> 26680 Then gError 26679

    'Se não há quantidade a reservar
    ElseIf dQuantidade = 0 And dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada = 0 And gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA Then
        
        'Elimina as reservas feitas anteriormente
        Set objItemRomaneioGrade.colLocalizacao = New Collection
        objItemRomaneioGrade.dQuantReservada = 0
                        
    End If
    
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
Dim objProdutoIPV As New ClassProduto

On Error GoTo Erro_ReservaAlmoxarifadoPadrao_Grade
   
    'calula a quantidade reservada
    dQuantidadeReservarVenda = (dQuantidade - dQuantCancelada - dQuantFaturada)

    objProdutoIPV.sCodigo = gobjPedidoVenda.colItensPedido(iItem).sProduto

    'Lê o Produto
    lErro = CF("Produto_Le", objProdutoIPV)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 30414
    
    If objProdutoIPV.iKitVendaComp <> MARCADO Then
        lErro = CF("UM_Conversao", objProduto.iClasseUM, gobjPedidoVenda.colItensPedido(iItem).sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 26670
    Else
        dFator = 1 'O Romaneio grade já está na UM de estoque e não necessáriamente a classe UM
        'do item de venda é igual aos produtos do romaneio
    End If

    'Converte a quantidade para a unidade de Estoque
    dQuantidadeReservarEstoque1 = dQuantidadeReservarVenda * dFator

    iFilialFaturamento = Codigo_Extrai(FilialFaturamento.Text)

    If iFilialFaturamento = 0 Then
        'Busca a filial de faturamento
        lErro = CF("FilialFaturamento_Le", giFilialEmpresa, iFilialFaturamento)
        If lErro <> SUCESSO Then gError 30414
        If iFilialFaturamento = 0 Then iFilialFaturamento = giFilialEmpresa
    End If

    'Busca o almoxarifado padrão
    lErro = CF("AlmoxarifadoPadrao_Le", iFilialFaturamento, objProduto.sCodigo, iAlmoxarifadoPadrao)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 51417
    If iAlmoxarifadoPadrao = 0 Then gError 51418

    objAlmoxarifadoPadrao.iCodigo = iAlmoxarifadoPadrao
    'Lê o almoxarifado padrão
    lErro = CF("Almoxarifado_Le", objAlmoxarifadoPadrao)
    If lErro = SUCESSO Then

        objEstoqueProduto1.sProduto = objProduto.sCodigo
        objEstoqueProduto1.iAlmoxarifado = iAlmoxarifadoPadrao
        objEstoqueProduto1.sAlmoxarifadoNomeReduzido = objAlmoxarifadoPadrao.sNomeReduzido

        'Lê o estoque do produto
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto1)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 26672
        If lErro = 21306 Then gError 26673

        objItemPV.lCodPedido = StrParaLong(Codigo.Text)
        objItemPV.sProduto = objProduto.sCodigo
        objItemPV.iFilialEmpresa = giFilialEmpresa
        objItemPV.sProdutoDescricao = objProduto.sDescricao
        objReservaBD.iAlmoxarifado = iAlmoxarifadoPadrao

        'Lê as reservas do produto
        lErro = CF("ReservaItem_Le", objItemPV, objReservaBD)
        If lErro <> SUCESSO And lErro <> 26678 Then gError 26674

        dQuantReservadaPedido = objReservaBD.dQuantidade

        'Saldo enxergado por esse Pedido Venda
        objEstoqueProduto1.dSaldo = objEstoqueProduto1.dQuantDisponivel + dQuantReservadaPedido

        'Se quantidade a reservar for menor que a disponível coloca reserva no GridAlocacao
        If (dQuantidadeReservarEstoque1 - objEstoqueProduto1.dSaldo) < QTDE_ESTOQUE_DELTA Then
        
            objItemRomaneioGrade.dQuantReservada = dQuantidadeReservarVenda
            
            Set objItemRomaneioGrade.colLocalizacao = New Collection
            
            Set objReservaItem = New ClassReservaItem
            
            objReservaItem.iAlmoxarifado = objEstoqueProduto1.iAlmoxarifado
            objReservaItem.sAlmoxarifado = objEstoqueProduto1.sAlmoxarifadoNomeReduzido
            objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarVenda)
            objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarEstoque1)
            objReservaItem.sResponsavel = RESERVA_AUTO_RESP
            objReservaItem.dtDataValidade = DATA_NULA

            objItemRomaneioGrade.colLocalizacao.Add objReservaItem

        Else

            lErro = CF("Retorna_Almoxarifado_Alternativo", iAlmoxarifadoPadrao, iFilialEmpresa)
            If lErro <> SUCESSO Then gError 105004

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
                If lErro <> SUCESSO And lErro <> 21306 Then gError 26672
                If lErro = 21306 Then gError 26673

                objItemPV.lCodPedido = StrParaLong(Codigo)
                objItemPV.sProduto = objProduto.sCodigo
                objItemPV.iFilialEmpresa = iFilialEmpresa
                objItemPV.sProdutoDescricao = objProduto.sDescricao
                objReservaBD.iAlmoxarifado = iAlmoxarifadoPadrao

                'Lê as reservas do produto
                lErro = CF("ReservaItem_Le", objItemPV, objReservaBD)
                If lErro <> SUCESSO And lErro <> 26678 Then gError 26674

                dQuantReservadaPedido = objReservaBD.dQuantidade

                'Saldo enxergado por esse Pedido Venda
                objEstoqueProduto2.dSaldo = objEstoqueProduto2.dQuantDisponivel + dQuantReservadaPedido

                'Se quantidade a reservar for menor que a disponível coloca reserva no GridAlocacao
                If (dQuantidadeReservarEstoque2 - objEstoqueProduto2.dSaldo) < QTDE_ESTOQUE_DELTA Then

                    If dQuantidadeReservarEstoque1 > 0 Then
                        
                        Set objReservaItem = New ClassReservaItem
                        
                        'Parte reservada na filial 1
                        objReservaItem.iAlmoxarifado = objEstoqueProduto1.iAlmoxarifado
                        objReservaItem.sAlmoxarifado = objEstoqueProduto1.sAlmoxarifadoNomeReduzido
                        objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarVenda)
                        objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarEstoque1)
                        objReservaItem.sResponsavel = RESERVA_AUTO_RESP
                        objReservaItem.dtDataValidade = DATA_NULA
                        
                        objItemRomaneioGrade.colLocalizacao.Add objReservaItem

                    End If

                    Set objReservaItem = New ClassReservaItem
                    
                    'Parte Reservada na Filial 2
                        objReservaItem.iAlmoxarifado = objEstoqueProduto2.iAlmoxarifado
                        objReservaItem.sAlmoxarifado = objEstoqueProduto2.sAlmoxarifadoNomeReduzido
                        objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarVenda)
                        objReservaItem.dQuantidade = Formata_Estoque(dQuantidadeReservarEstoque2)
                        objReservaItem.dtDataValidade = DATA_NULA
                        objReservaItem.sResponsavel = RESERVA_AUTO_RESP
                        
                        objItemRomaneioGrade.colLocalizacao.Add objReservaItem

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
                    objItemPV.iItem = iItem
                    objItemPV.lNumIntDoc = gcolItemPedido(iItem)
                    objItemPV.sUMEstoque = objProduto.sSiglaUMEstoque
                    objItemPV.sProdutoDescricao = objProduto.sDescricao
                    If objProdutoIPV.iKitVendaComp = MARCADO Then
                        objItemPV.sUnidadeMed = objItemRomaneioGrade.sUMEstoque
                    Else
                        objItemPV.sUnidadeMed = gobjPedidoVenda.colItensPedido(iItem).sUnidadeMed
                    End If
                    objItemPV.iClasseUM = objProduto.iClasseUM
                    objItemPV.iPossuiGrade = MARCADO
                    
                    If iImportacao = MARCADO Then
                        ReservaAlmoxarifadoPadrao_Grade = SUCESSO
                        Exit Function
                    End If
                    
                    'Chama tela de Falta de Estoque
                    lErro = Chama_Tela_Modal("FaltaEstoque", objItemPV, colItemPedido, dQuantidadeReservarEstoque1, objAlmoxarifadoPadrao, objEstoqueProduto1.dSaldo)

                    'Se retornar Cancela erro
                    If giRetornoTela = vbCancel Then gError 26680

                    'Limpa reservas desse ítem no GridAlocacao
                    Set objItemRomaneioGrade.colLocalizacao = New Collection
                    
                    If giRetornoTela = vbOK Then

                       'Se não substituiu o Produto
                        'Coloca QuantReservada e QuantCancelada no ítem do GridItens
                        objItemRomaneioGrade.dQuantReservada = objItemPV.dQuantReservada
                       
                        objItemRomaneioGrade.dQuantCancelada = objItemPV.dQuantCancelada

                        'Coloca reseravas desse ítem no GridAlocacao
                        For Each objReserva In objItemPV.colReserva
                            Set objReservaItem = New ClassReservaItem
                            
                            objReservaItem.iAlmoxarifado = objReserva.iAlmoxarifado
                            objReservaItem.dQuantidade = objReserva.dQuantidade
                            objReservaItem.dtDataValidade = objReserva.dtDataValidade
                            objReservaItem.sAlmoxarifado = objReserva.sAlmoxarifado
                            objReservaItem.sResponsavel = objReserva.sResponsavel
                            
                            objItemRomaneioGrade.colLocalizacao.Add objReservaItem
                            
                        Next
                    End If
                End If
            End If
        End If
    End If

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
Function Reserva_Trata_Grade_TrocaUM(objItemPV As ClassItemPedido) As Long

Dim lErro As Long
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim colItensRomaneio As New Collection
Dim objItemRomaneioAux As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim objReservaItemTela As ClassReservaItem

On Error GoTo Erro_Reserva_Trata_Grade_TrocaUM
    
    For Each objItemRomaneio In objItemPV.colItensRomaneioGrade

        Set objItemRomaneioAux = New ClassItemRomaneioGrade
        
        objItemRomaneioAux.dQuantAFaturar = objItemRomaneio.dQuantAFaturar
        objItemRomaneioAux.dQuantCancelada = objItemRomaneio.dQuantCancelada
        objItemRomaneioAux.dQuantFaturada = objItemRomaneio.dQuantFaturada
        objItemRomaneioAux.dQuantidade = objItemRomaneio.dQuantidade
        objItemRomaneioAux.dQuantOP = objItemRomaneio.dQuantOP
        objItemRomaneioAux.dQuantReservada = objItemRomaneio.dQuantReservada
        objItemRomaneioAux.dQuantSC = objItemRomaneio.dQuantSC
        objItemRomaneioAux.lNumIntDoc = objItemRomaneio.lNumIntDoc
        objItemRomaneioAux.sDescricao = objItemRomaneio.sDescricao
        objItemRomaneioAux.sProduto = objItemRomaneio.sProduto
        objItemRomaneioAux.sUMEstoque = objItemRomaneio.sUMEstoque
        
        For Each objReservaItemTela In objItemRomaneio.colLocalizacao
            
            Set objReservaItem = New ClassReservaItem
            
            objReservaItem.dQuantidade = objReservaItemTela.dQuantidade
            objReservaItem.dtDataValidade = objReservaItemTela.dtDataValidade
            objReservaItem.iAlmoxarifado = objReservaItemTela.iAlmoxarifado
            objReservaItem.iFilialEmpresa = objReservaItemTela.iFilialEmpresa
            objReservaItem.lNumIntDoc = objReservaItemTela.lNumIntDoc
            objReservaItem.sAlmoxarifado = objReservaItemTela.sAlmoxarifado
            objReservaItem.sResponsavel = objReservaItemTela.sResponsavel
            
            objItemRomaneioAux.colLocalizacao.Add objReservaItem
            
        Next
        
        lErro = Reserva_Processa_ItemGrade(objItemRomaneioAux, GridItens.Row)
        If lErro <> SUCESSO Then gError 26852
        
        colItensRomaneio.Add objItemRomaneioAux
        
    Next
    
    Set objItemPV.colItensRomaneioGrade = colItensRomaneio
    
    Call Atualiza_Grid_Itens(objItemPV)
    Call Atualiza_Grid_Alocacao(objItemPV)
    
    Reserva_Trata_Grade_TrocaUM = SUCESSO
    
    Exit Function

Erro_Reserva_Trata_Grade_TrocaUM:

    Reserva_Trata_Grade_TrocaUM = gErr
    
    Select Case gErr
    
        Case 26852
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158091)
            
    End Select
    
    Exit Function
    
End Function

Function Move_ItensGrade_Tela(colItensRomaneio As Collection, colItensRomaneioTela As Collection) As Long

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradeTela As ClassItemRomaneioGrade
Dim objReservaItemTela As ClassReservaItem
Dim objReservaItem As ClassReservaItem

    'Para cada Item de Romaneio vindo da tela ( Aqueles que já tem quantidade)
    For Each objItemRomaneioGradeTela In colItensRomaneioTela
                    
        Set objItemRomaneioGrade = New ClassItemRomaneioGrade
            
        objItemRomaneioGrade.sProduto = objItemRomaneioGradeTela.sProduto
        objItemRomaneioGrade.dQuantOP = objItemRomaneioGradeTela.dQuantOP
        objItemRomaneioGrade.dQuantSC = objItemRomaneioGradeTela.dQuantSC
        objItemRomaneioGrade.sDescricao = objItemRomaneioGradeTela.sDescricao
        objItemRomaneioGrade.dQuantAFaturar = objItemRomaneioGradeTela.dQuantAFaturar
        objItemRomaneioGrade.dQuantFaturada = objItemRomaneioGradeTela.dQuantFaturada
        objItemRomaneioGrade.dQuantidade = objItemRomaneioGradeTela.dQuantidade
        objItemRomaneioGrade.dQuantReservada = objItemRomaneioGradeTela.dQuantReservada
        objItemRomaneioGrade.sUMEstoque = objItemRomaneioGradeTela.sUMEstoque
        objItemRomaneioGrade.dQuantCancelada = objItemRomaneioGradeTela.dQuantCancelada
        
        'Transfere as informações de Localização
        Set objItemRomaneioGrade.colLocalizacao = New Collection
            
        For Each objReservaItemTela In objItemRomaneioGradeTela.colLocalizacao
            
            Set objReservaItem = New ClassReservaItem
            
            objReservaItem.dQuantidade = objReservaItemTela.dQuantidade
            objReservaItem.dtDataValidade = objReservaItemTela.dtDataValidade
            objReservaItem.iAlmoxarifado = objReservaItemTela.iAlmoxarifado
            objReservaItem.iFilialEmpresa = giFilialEmpresa
            objReservaItem.lNumIntDoc = objReservaItemTela.lNumIntDoc
            objReservaItem.sAlmoxarifado = objReservaItemTela.sAlmoxarifado
            objReservaItem.sResponsavel = objReservaItemTela.sResponsavel
                        
            objItemRomaneioGrade.colLocalizacao.Add objReservaItem
            
        Next
                    
        colItensRomaneio.Add objItemRomaneioGrade
    Next

    Exit Function

End Function

Function Transfere_Dados_ItensRomaneioOrcamento(colItensRomaneioPV As Collection, colItensRomaneioOV As Collection) As Long

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradePV As ClassItemRomaneioGrade
Dim objReservaItemTela As ClassReservaItem
Dim objReservaItem As ClassReservaItem

    'Para cada Item de Romaneio vindo da tela ( Aqueles que já tem quantidade)
    For Each objItemRomaneioGradePV In colItensRomaneioOV

        Set objItemRomaneioGrade = New ClassItemRomaneioGrade

        'Transfere as informações vindas da tela chamadora para essa tela
        objItemRomaneioGrade.sDescricao = objItemRomaneioGradePV.sDescricao
        objItemRomaneioGrade.dQuantidade = objItemRomaneioGradePV.dQuantidade
        objItemRomaneioGrade.sUMEstoque = objItemRomaneioGradePV.sUMEstoque
        objItemRomaneioGrade.sProduto = objItemRomaneioGradePV.sProduto

        colItensRomaneioPV.Add objItemRomaneioGrade

    Next

    Exit Function

End Function

Sub PedidoVenda_Atualiza_Orcamento(ByVal objOrcamento As ClassOrcamentoVenda)

Dim iIndice As Integer

    For iIndice = 1 To objOrcamento.colItens.Count
        Call Transfere_Dados_ItensRomaneioOrcamento(gobjPedidoVenda.colItensPedido(iIndice).colItensRomaneioGrade, objOrcamento.colItens(iIndice).colItensRomaneioGrade)
    Next

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
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, TranspRedespacho.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158092)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158093)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158094)
     
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
    
    lErro = CF2(Me, "PV_DataEntregaPV_Validate", Cancel)
    If lErro = SUCESSO Then 'gError 26582

        'If Cancel Then Exit Sub
    
        'Verifica se a Data foi digitada
        If Len(Trim(DataEntregaPV.ClipText)) <> 0 Then
        
            'Critica a data digitada
            lErro = Data_Critica(DataEntregaPV.Text)
            If lErro <> SUCESSO Then gError 26582
            
        End If
        
        dtData = MaskedParaDate(DataEntregaPV)
        
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If dtData <> DATA_NULA Then
                GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col) = Format(dtData, "dd/mm/yyyy")
            Else
                GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col) = ""
            End If
        Next
        
        If gobjPedidoVenda.dtDataEntrega <> StrParaDate(DataEntregaPV.Text) Then
            
            gobjPedidoVenda.dtDataEntrega = StrParaDate(DataEntregaPV.Text)
            
            Call ValorTotal_Calcula
            
        End If
        
        giDataEntregaAlterada = 0
        
    End If
    
    Exit Sub

Erro_DataEntregaPV_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 26582

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158095)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntregaPV_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntregaPV_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEntregaPV, DIMINUI_DATA)
    If lErro Then gError 26583

    Exit Sub

Erro_UpDownEntregaPV_DownClick:

    Select Case gErr

        Case 26583

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158096)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntregaPV_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntregaPV_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEntregaPV, AUMENTA_DATA)
    If lErro Then gError 26584

    Exit Sub

Erro_UpDownEntregaPV_UpClick:

    Select Case gErr

        Case 26584

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158097)

    End Select

    Exit Sub

End Sub

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
    If giFilialCliInexistente = FILIALCLIENTE_INEXISTENTE Then gError 102199
    
    'Se a filial de entrega não está cadastrada => erro
    If giFilialCliInexistente = FILIALENTREGA_INEXISTENTE Then gError 102200
    
    giFilialCliInexistente = 0
    giClienteInexistente = 0
    
    Cadastra_Cliente_Inexistente = SUCESSO
    
    Exit Function
    
Erro_Cadastra_Cliente_Inexistente:

    Select Case gErr
    
        Case 102199
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)
            
        Case 102200
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_INEXISTENTE", gErr, FilialEntrega.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158098)
    
    End Select

End Function

'Transferir para versão 2
Private Function Traz_Orcamento_Cliente(objCliente As ClassCliente, ByVal objPedidoVenda As ClassPedidoDeVenda)
'Traz para tela os dados defaults do cliente, quando a tela é preenchida a partir de orçamento
'Dispara o cálculo de comissões caso a empresa não utilize o sistema de regras de comissões

Dim lErro As Long
Dim objMensagem As ClassMensagem
Dim objTipoCliente As ClassTipoCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objTransportadora As New ClassTransportadora
Dim objTipoFrete As ClassTipoFrete
Dim colTipoFrete As New Collection
Dim iIndice As Integer

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
        If lErro <> SUCESSO And lErro <> 12348 Then gError 102245
        
        'Se não encontrou o cliente => erro
        If lErro = 12348 Then gError 102246
        
    End If
    '*****************************

    '**** LEITURA DA FILIAL DO CLIENTE *****
    'Guarda em objFilialCliente os dados necessários para ler a filial
    objFilialCliente.lCodCliente = objCliente.lCodigo
    objFilialCliente.iCodFilial = objPedidoVenda.iFilial
    
    'Lê a Filial cliente para obter os dados default que serão exibidos na tela
    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 12567 Then gError 102250

    'Se não encontrou a filial => erro
    If lErro = 12567 Then gError 102251
    '*****************************************

''    '**** TRATAMENTO PARA TRANSPORTADORA DEFAULT ****
''    'Guarda o código da transportadora em objtransportadora
''    objTransportadora.iCodigo = objFilialCliente.iCodTransportadora
''
''    'Se ainda não encontrou Transportadora default
''    If objTransportadora.iCodigo = 0 Then
''
''        'Instancia objTipoCliente
''        Set objTipoCliente = New ClassTipoCliente
''
''        'Guarda o tipo de cliente no obj
''        objTipoCliente.iCodigo = objCliente.iTipo
''
''        'Lê o tipo de cliente para checar se o tipo tem transportadora default
''        lErro = CF("TipoCliente_Le", objTipoCliente)
''        If lErro <> SUCESSO And lErro <> 19062 Then gError 102252
''
''        'Se não encontrou => erro
''        If lErro = 19062 Then gError 102249
''
''        'Guarda o código da transportadora em objtransportadora
''        objTransportadora.iCodigo = objTipoCliente.iCodTransportadora
''
''    End If
''
''    'Se encontrou uma transportadora default
''    If objTransportadora.iCodigo <> 0 Then
''        Transportadora.Text = objTransportadora.iCodigo
''        Call Transportadora_Validate(False)
''    End If
''    '*****************************************
    
''    '**** TRATAMENTO PARA TIPO DE FRETE ****
''    'Tipo de frete
''
''    lErro = CF("TipoFrete_Le", colTipoFrete)
''    If lErro <> SUCESSO Then gError 207417
''
''    For Each objTipoFrete In colTipoFrete
''        If objTipoFrete.iCodigoNFE = objFilialCliente.iTipoFrete Then
''            For iIndice = 0 To TipoFrete.ListCount - 1
''                If TipoFrete.ItemData(iIndice) = objTipoFrete.iCodigo Then
''                    TipoFrete.ListIndex = iIndice
''                    Exit For
''                End If
''            Next
''            Exit For
''        End If
''    Next
''
''    '*****************************************
''
''
''
''
''
''    '**** TRATAMENTO PARA MENSAGEM DE NOTA FISCAL ****
''    'Se existe mensagem de nota fiscal para o cliente
''    If objCliente.iCodMensagem > 0 Then
''
''        'Instancia objMensagem
''        Set objMensagem = New ClassMensagem
''
''        'Guarda o código da mensagem no obj
''        objMensagem.iCodigo = objCliente.iCodMensagem
''
''        'Lê a mensagem do Cliente
''        lErro = CF("Mensagem_Le", objMensagem)
''        If lErro <> SUCESSO And lErro <> 19234 Then gError 102247
''
''        'Se não encontrou a mensagem => erro
''        If lErro = 19234 Then gError 102248
''
''        'Exibe a mensagem na tela
''        Mensagem.Text = objMensagem.sDescricao
''
''    End If
''    '**************************************************
    
    '*** TRATAMENTO PARA COMISSÕES ***
    'Se a empresa não utiliza as regras para cálculo de comissões, calcula as comissões no modelo antigo
    'Caso a empresa utilize as regras, o cáclulo será efetuado ao clicar no tab comissões ou ao gravar o pedido
    If Not (gobjCRFAT.iUsaComissoesRegras = USA_REGRAS) Then
        
        'calcular as comissoes no modelo antigo...
        lErro = objTabComissoes.Comissao_Automatica_FilialCli_Exibe(objFilialCliente)
        If lErro <> SUCESSO Then gError 102254
    
    End If
    '*********************************
    
    Traz_Orcamento_Cliente = SUCESSO

    Exit Function
    
Erro_Traz_Orcamento_Cliente:

    Traz_Orcamento_Cliente = gErr
    
    Select Case gErr
    
        Case 102245, 102247, 102250, 102152, 207415, 207417
        
        Case 102246
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
        
        Case 102251
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_INEXISTENTE", gErr, objCliente.lCodigo, objFilialCliente.iCodFilial)
        
        Case 102249
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo, objCliente.iTipo)
        
        Case 102248
            Call Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_CADASTRADA", gErr, objCliente.lCodigo, objCliente.iCodMensagem)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158099)
    
    End Select
    
End Function
'********************************

Sub PedidoVenda_ExibirBloqsPV(ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal colBloqPVAux As Collection, bExibirBloqsPV As Boolean)
'retorna em bExibirBloqsPV se deve abrir a tela que mostra os bloqueios de pv

Dim lErro As Long, objBloqPV As ClassBloqueioPV, objBloqPVAux As ClassBloqueioPV
Dim bCriado As Boolean

On Error GoTo Erro_PedidoVenda_ExibirBloqsPV

    'verifica se algum bloqueio automatico foi criado durante a gravacao
    bCriado = False
    
    For Each objBloqPV In objPedidoVenda.colBloqueiosPV
    
        If objBloqPV.dtDataLib = DATA_NULA Then
            
            bCriado = True
    
            For Each objBloqPVAux In colBloqPVAux
            
                If objBloqPVAux.iTipoDeBloqueio = objBloqPV.iTipoDeBloqueio Then
                
                    bCriado = False
                    Exit For
                    
                End If
                
            Next
            
            If bCriado Then Exit For
        
        End If
        
    Next
    
    bExibirBloqsPV = bCriado
    
    Exit Sub
     
Erro_PedidoVenda_ExibirBloqsPV:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158103)
     
    End Select
     
    Exit Sub

End Sub

'@@@@@ para ser customizada
Function PV_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
        
Dim lErro As Long

On Error GoTo Erro_PV_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col
    
        'Se for a de Percentual de Desconto
        Case iGrid_PercDesc_Col
            lErro = Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then gError 26599

    End Select
    
    PV_Saida_Celula_GridItens = SUCESSO
     
    Exit Function
    
Erro_PV_Saida_Celula_GridItens:

    PV_Saida_Celula_GridItens = gErr
     
    Select Case gErr
          
        Case 26599
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158104)
     
    End Select
     
    Exit Function

End Function

'@@@@@ para ser customizada
Function PV_Move_Pedido_Memoria(ByVal objCT As Object, objPedidoVenda As ClassPedidoDeVenda) As Long
    PV_Move_Pedido_Memoria = SUCESSO
End Function

'@@@@@ para ser customizada
Function PV_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iCaminho As Integer, ByVal sProdutoFormatado As String, ByVal iProdutoPreenchido As Integer) As Long
    PV_Rotina_Grid_Enable = SUCESSO
End Function

Function PV_Rotina_Grid_Enable2(ByVal objCT As Object, iLinha As Integer, objControl As Object, iCaminho As Integer, ByVal sProdutoFormatado As String, ByVal iProdutoPreenchido As Integer) As Long
    PV_Rotina_Grid_Enable2 = SUCESSO
End Function

Function PV_PrecoTotal_Calcula3(ByVal objCT As Object, ByVal iLinha As Integer, ByVal dQuantidade As Double, ByVal dPrecoUnitario As Double, ByVal dPercentDesc As Double) As Long

Dim dPrecoTotal As Double, dPrecoTotalReal As Double, dDesconto As Double
        
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
        
'    dPrecoTotal = dPrecoUnitario * dQuantidade
'
'    'Se percentual for >0 tira o desconto
'    If dPercentDesc > 0 Then dDesconto = StrParaDbl(Format(dPercentDesc * dPrecoTotal, "Standard")) 'Alterado por Wagner
'    dPrecoTotalReal = dPrecoTotal - dDesconto
'
'    'Coloca valor do desconto no Grid
'    If dDesconto > 0 Then
'        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
'    Else
'        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
'    End If
'
'    'Coloca preco total do ítem no grid
'    GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(dPrecoTotalReal, "Standard")
'    GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(dPrecoTotal, "Standard")


    
    PV_PrecoTotal_Calcula3 = SUCESSO
    
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
        'se os percentuais (passado e calulado) forem diferentes calcula-se o desconto
        If dPercentDesc <> dPercDesc2 Then dDesconto = Arredonda_Moeda(dPercentDesc * dValorTotal)

    End If

    dValorReal = dValorTotal - dDesconto

End Sub

'@@@@@ para ser customizada
Function PV_Move_GridItem_Memoria(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objItemPedido As ClassItemPedido, ByVal iIndice As Integer, ByVal sProduto1 As String) As Long
    PV_Move_GridItem_Memoria = SUCESSO
End Function

'@@@@@ para ser customizada
Function PV_Carrega_Grid_Itens(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal bProcessaReserva As Boolean) As Long
    PV_Carrega_Grid_Itens = SUCESSO
End Function

'@@@@@ para ser customizada
Function PV_Traz_Pedido_Tela(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
    PV_Traz_Pedido_Tela = SUCESSO
End Function

Function PV_Traz_Pedido_Tela2(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
    PV_Traz_Pedido_Tela2 = SUCESSO
End Function

Function PV_Preenche_GridItens(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda, ByVal objItemPedido As ClassItemPedido, ByVal objProduto As ClassProduto, ByVal iIndice As Integer, ByVal dPercDesc As Double) As Long
    PV_Preenche_GridItens = SUCESSO
End Function

Public Function Calcula_VolumeTotal() As Long
'Realiza a soma dos volumes dos produtos contido nop pedido de venda

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iIndice As Integer
Dim dFator As Double
Dim sProduto As String
Dim iPreenchido As Integer
Dim dVolume As Double, dQuant As Double

On Error GoTo Erro_Calcula_VolumeTotal

    glVolumeQuantCalc = 0
    
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
    
        dQuant = (StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col)) - IIf(iGrid_QuantCancel_Col <> 0, StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col)), 0)) * dFator
    
        If objProduto.dQtdeEmbBase = 0 Then
            glVolumeQuantCalc = glVolumeQuantCalc + dQuant
        Else
            glVolumeQuantCalc = glVolumeQuantCalc + Arredonda_ParaCima(dQuant / objProduto.dQtdeEmbBase)
        End If
        
        'Realiza o cálculo do volume
        dVolume = dVolume + (dQuant * (objProduto.dComprimento * objProduto.dEspessura * objProduto.dLargura))
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158105)

    End Select

    Exit Function
    
End Function

'@@@@@ para ser customizada
Function PV_Cliente_Validate(ByVal objCT As Object, ByVal lCliente As Long) As Long
    PV_Cliente_Validate = SUCESSO
End Function

Function PV_Cliente_Validate2(ByVal objCT As Object, ByVal lCliente As Long) As Long
    PV_Cliente_Validate2 = SUCESSO
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158106)

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
    If lErro <> SUCESSO Then gError 65897

    objProduto.sCodigo = sProdutoFormatado

    'Lê o produto da linha passada por iLinha do GridItens
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 65898
    If lErro = 28030 Then gError 65899

    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUmAnterior, sUMNova, dFator)
    If lErro <> SUCESSO Then gError 65900

    'Atualiza o Grid
    If StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_QuantPedida_Col)) * dFator > QTDE_ESTOQUE_DELTA Then
        GridItens.TextMatrix(iLinha, iGrid_QuantPedida_Col) = Formata_Estoque(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_QuantPedida_Col)) * dFator)
    Else
        GridItens.TextMatrix(iLinha, iGrid_QuantPedida_Col) = ""
    End If
    
    If iGrid_QuantCancel_Col <> 0 Then
        If StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_QuantCancel_Col)) * dFator > QTDE_ESTOQUE_DELTA Then
            GridItens.TextMatrix(iLinha, iGrid_QuantCancel_Col) = Formata_Estoque(StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_QuantCancel_Col)) * dFator)
        Else
            GridItens.TextMatrix(iLinha, iGrid_QuantCancel_Col) = ""
        End If
    End If

    'Preço unitário
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))

    'Recalcula preço do ítem e valor total da nota
    If dPrecoUnitario > 0 Then
        GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario / dFator, gobjFAT.sFormatoPrecoUnitario)
        Call CF2(Me, "PV_PrecoTotal_Calcula", iLinha)
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 51037
    End If

    Atualiza_UM = SUCESSO
    
    Exit Function

Erro_Atualiza_UM:

    Atualiza_UM = gErr

    Select Case gErr

        Case 65897, 65898, 65900, 65901, 51037

        Case 65899
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158107)

    End Select

    Exit Function

End Function

Function PV_Botao_Estoque(ByVal objCT As Object, ByVal objEstoqueProd As ClassEstoqueProduto, ByVal colSelecao As Collection, sSelecao As String) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String, sProdutoEnxuto As String
Dim iAlmoxarifado As Integer

On Error GoTo Erro_PV_Botao_Estoque

    sSelecao = ""

    'VerIfica se tem alguma linha selecionada
    If objCT.GridItens.Row = 0 Then gError 58773

    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 58774

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        objEstoqueProd.sProduto = sProdutoFormatado
        
        '*********** INCLUIDO PARA O TRATAMENTO DE ITEM DE GRADE ****************
        If gobjPedidoVenda.colItensPedido(GridItens.Row).iPossuiGrade = MARCADO Then
            Call Mascara_RetornaProdutoEnxuto(sProdutoFormatado, sProdutoEnxuto)
            sSelecao = "Produto LIKE '" & sProdutoEnxuto & "%'"
            Call Chama_Tela("EstProdLista_Consulta", colSelecao, objEstoqueProd, objEventoEstoqueProd, sSelecao)
            gError 58774
        End If

        colSelecao.Add sProdutoFormatado

        lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, sProdutoFormatado, iAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 23796 Then gError 58775

        If lErro = SUCESSO Then objEstoqueProd.iAlmoxarifado = iAlmoxarifado

    Else
        gError 58776


    End If

    PV_Botao_Estoque = SUCESSO
    
    Exit Function

Erro_PV_Botao_Estoque:

    PV_Botao_Estoque = gErr

    Select Case gErr

        Case 58773
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 58774, 58775 'Tratado na rotina chamada

        Case 58776
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158108)

    End Select

    Exit Function
    
End Function

Function PV_Muda_Tabulacao(ByVal objCT As Object, ByVal iKey As Integer, ByVal iLinha As Integer, ByVal iColuna As Integer) As Long
    PV_Muda_Tabulacao = SUCESSO
End Function

Function PV_ValorDesconto_Validate(ByVal objCT As Object) As Long
    PV_ValorDesconto_Validate = SUCESSO
End Function

Function PV_NF_DataSaida(ByVal objCT As Object, ByVal gobjGeracaoNFiscal As ClassGeracaoNFiscal) As Long
    gobjGeracaoNFiscal.dtDataSaida = gdtDataHoje
End Function

Public Sub BotaoTodosPedidos_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colSelecao As Collection

On Error GoTo Erro_NumeroLabel_Click

    lErro = Move_Pedido_Memoria(objPedidoVenda, False)
    If lErro <> SUCESSO Then gError 26498

    Call Chama_Tela("PedidoVendaTodosLista", colSelecao, objPedidoVenda, objEventoNumero)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 26498

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158109)

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
    If lErro <> SUCESSO Then gError 134013

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134013

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158110)

    End Select
    
    Exit Sub

End Sub

Public Function PV_Carrega_Grid_Itens_Linha(ByVal objCT As Object, objPedidoVenda As ClassPedidoDeVenda, ByVal iIndice As Integer) As Long
    PV_Carrega_Grid_Itens_Linha = SUCESSO
End Function

Public Function PV_Codigo_Validate(ByVal objCT As Object, objPedidoVenda As ClassPedidoDeVenda) As Long
    PV_Codigo_Validate = SUCESSO
End Function

Public Function PV_Codigo_Enabled(ByVal objCT As Object, ByVal bFlag As Boolean) As Long
    PV_Codigo_Enabled = SUCESSO
End Function

Public Function PV_Produto_Saida_Celula(ByVal objCT As Object, ByVal bPreenchePreco As Boolean, ByVal dPrecoUnitario As Double) As Long
    PV_Produto_Saida_Celula = SUCESSO
End Function

Public Function PV_Saida_Celula_PrecoUnitario(ByVal objCT As Object, ByVal dPrecoUnitario As Double, ByVal iLinha As Integer) As Long
    PV_Saida_Celula_PrecoUnitario = SUCESSO
End Function

Public Function PV_AlocaAlmoxarifadoPadrao(ByVal objCT As Object, ByVal iFilialEmpresa As Integer, ByVal sProduto As String, iAlmoxarifadoRetorno As Integer) As Long

Dim lErro  As Long

On Error GoTo Erro_PV_AlocaAlmoxarifadoPadrao

    'Busca o Almoxarifado padrão
    lErro = CF("AlmoxarifadoPadrao_Le", iFilialEmpresa, sProduto, iAlmoxarifadoRetorno)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 51619
    
    'Se encontrou
    If iAlmoxarifadoRetorno = 0 Then gError 35822
    
    Exit Function

Erro_PV_AlocaAlmoxarifadoPadrao:

    PV_AlocaAlmoxarifadoPadrao = gErr

    Select Case gErr
    
        Case 35822, 51619

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158111)

    End Select

    Exit Function
    
End Function

Public Function PV_Form_Load(ByVal objCT As Object) As Long
    PV_Form_Load = SUCESSO
End Function

Public Function PV_Traz_Dados_Orcamento_Tela(ByVal objCT As Object, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
    PV_Traz_Dados_Orcamento_Tela = SUCESSO
End Function

Public Function PV_Produto_Saida_Celula2(ByVal objCT As Object) As Long

Dim iIndice As Integer

On Error GoTo Erro_PV_Produto_Saida_Celula2

    If gobjCRFAT.iProdRepetidoPV = PRODUTO_REPETICAO_NAO_PERMITE Then
    
        'Verifica se já está em outra linha do Grid
        For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
            If iIndice <> objCT.GridItens.Row Then
                If objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Produto_Col) = objCT.Produto.Text Then gError 177164
            End If
        Next

    End If
    
    PV_Produto_Saida_Celula2 = SUCESSO
    
    Exit Function

Erro_PV_Produto_Saida_Celula2:

    PV_Produto_Saida_Celula2 = gErr

    Select Case gErr
    
        Case 177164
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, objCT.Produto.Text, objCT.Produto.Text, iIndice)

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177165)

    End Select

    Exit Function
    
End Function

Public Function PV_Produto_Saida_Celula3(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim sCodigoAntigo As String

On Error GoTo Erro_PV_Produto_Saida_Celula3

    If gobjCRFAT.iProdRepetidoPV = PRODUTO_REPETICAO_NAO_PERMITE Then
            
        sCodigoAntigo = objProduto.sCodigo
    
        If objProduto.iGerencial = DESMARCADO Then
    
            For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
                If iIndice <> objCT.GridItens.Row Then
                    If objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Produto_Col) = objCT.Produto.Text Then
                        If vbYes = Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_EQUIVALENTE", objCT.Produto.Text, iIndice) Then
            
                            lErro = CF("Produto_Le", objProduto)
                            If lErro <> SUCESSO And lErro <> 28030 Then gError 131295
            
                            Call Chama_Tela_Modal("ProdutoEqui", objProduto)
            
                            If giRetornoTela <> OK Then gError 131296
            
                            lErro = CF("Produto_Grava_Equivalente", objProduto, sCodigoAntigo)
                            If lErro <> SUCESSO Then gError 131297
            
                        Else
                            gError 131298
                        End If
                    End If
                End If
            Next
            
        End If
    
    End If
    
    PV_Produto_Saida_Celula3 = SUCESSO
    
    Exit Function

Erro_PV_Produto_Saida_Celula3:

    PV_Produto_Saida_Celula3 = gErr

    Select Case gErr
    
        Case 131295, 131296, 131297, 131298
         
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177165)

    End Select

    Exit Function
    
End Function

Public Function PV_Calcula_PesoKg(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    PV_Calcula_PesoKg = SUCESSO
End Function

Public Function PV_Calcula_PesoKg_Total(ByVal objCT As Object) As Long
    PV_Calcula_PesoKg_Total = SUCESSO
End Function

Public Function PV_PrecoTotal_Calcula2(ByVal objCT As Object, ByVal iLinha As Integer, dQuantidade As Double) As Long
    PV_PrecoTotal_Calcula2 = SUCESSO
End Function

Public Function PV_Valida_Grid_Itens(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    PV_Valida_Grid_Itens = SUCESSO
End Function

Public Function PV_Transfere_Dados_Orcamento_Pedido(ByVal objCT As Object, ByVal objItemOV As ClassItemOV, ByVal objItemPV As ClassItemPedido) As Long
    PV_Transfere_Dados_Orcamento_Pedido = SUCESSO
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
            If iGrid_QuantCancel_Col <> 0 Then dQuantidadeCancelada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))
            dQuantidadeFaturada = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantFaturada_Col))
        
            lErro = Reserva_Processa2(iIndice, dQuantidade, dQuantidadeCancelada, dQuantidadeFaturada, True)
            If lErro <> SUCESSO Then gError 177335
            
        Else

            For Each objItemRomaneio In gobjPedidoVenda.colItensPedido(iIndice).colItensRomaneioGrade
                
                lErro = Reserva_Processa_ItemGrade(objItemRomaneio, iIndice, True)
                If lErro <> SUCESSO Then gError 177587
            
            Next
            
            Call Atualiza_Grid_Alocacao(gobjPedidoVenda.colItensPedido.Item(iIndice))
            
        End If
            
    Next

    BotaoRefazAlocacao_Click = SUCESSO

    Exit Function

Erro_BotaoRefazAlocacao_Click:

    BotaoRefazAlocacao_Click = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177336)

    End Select

    Exit Function

End Function

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

Public Sub PedidoRepr_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Function PV_Exibe_QuantEstoque(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    PV_Exibe_QuantEstoque = SUCESSO
End Function

Public Sub PedidoRepr_Validate(Cancel As Boolean)

Dim lErro As Long

    lErro = CF("PV_PedidoRepr_Valida", StrParaDbl(Codigo.Text), giFilialEmpresa, PedidoRepr.Text)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function PV_BotaoProdutos_Click(ByVal objCT As Object, ByVal objProduto As ClassProduto, sOrdenacao As String) As Long
    PV_BotaoProdutos_Click = SUCESSO
End Function

Public Function PV_BotaoProdutos_Click2(ByVal objCT As Object, sNomeBrowse As String) As Long
    sNomeBrowse = "ProdutoVendaLista"
    PV_BotaoProdutos_Click2 = SUCESSO
End Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180810)

    End Select

    Exit Function
    
End Function

Private Sub objEventoContato_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objContato As ClassContatos
Dim objCliente As New ClassCliente

On Error GoTo Erro_objEventoContato_evSelecao

    Set objContato = obj1

    'Lê o Contato no BD
    lErro = CF("Contato_Transforma_Cliente", objContato, objCliente)
    If lErro <> SUCESSO Then gError 181158
    
    Cliente.Text = objCliente.sNomeReduzido

    Call Cliente_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoContato_evSelecao:

    Select Case gErr
    
        Case 181158

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181159)

    End Select

    Exit Sub

End Sub

Public Function PV_Produto_Saida_Celula4(ByVal objCT As Object, ByVal objProduto As ClassProduto, lErro As Long) As Long
    PV_Produto_Saida_Celula4 = SUCESSO
End Function

Public Function PV_Produto_Saida_Celula5(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long

Dim dtDataAux As Date

    'Data de Entrega
    dtDataAux = MaskedParaDate(DataEntregaPV)
    If dtDataAux <> DATA_NULA Then GridItens.TextMatrix(GridItens.Row, iGrid_DataEntrega_Col) = Format(dtDataAux, "dd/mm/yyyy")
    
    PV_Produto_Saida_Celula5 = SUCESSO
    
End Function

Public Function PV_Produto_Saida_Celula6(ByVal objCT As Object) As Long
    PV_Produto_Saida_Celula6 = SUCESSO
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

        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col)) - IIf(iGrid_QuantCancel_Col <> 0, StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col)), 0)
        objItensPRJCR.dValorOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        objItensPRJCR.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) & SEPARADOR & GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)

        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_PV
        
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178944)

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
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 178946

            'Não foi encontrado
            If lErro = 25085 Then gError 178947
            If lErro = 25086 Then gError 178948

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
    If lErro <> SUCESSO Then gError 178949

    Saida_Celula_TipoPagto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoPagto:

    Saida_Celula_TipoPagto = gErr

    Select Case gErr

        Case 178946, 178949
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 178947
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMAPAGTO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 178948
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMAPAGTO_NAO_ENCONTRADO", gErr, objGridInt.objControle.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178950)

    End Select

    Exit Function

End Function

Function PV_Inicializa_Grid_Parcelas(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
    
    objGridInt.objGrid.ColWidth(iGrid_TipoPagto_Col) = 0
    objGridInt.objGrid.TextMatrix(0, iGrid_TipoPagto_Col) = ""
    
    PV_Inicializa_Grid_Parcelas = SUCESSO
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178974)

    End Select

    Exit Sub

End Sub

Public Function BotaoEntrega_Click() As Long

Dim sProdutoTela As String
Dim objDataEntrega As ClassDataEntrega
Dim dQuantidade As Double

On Error GoTo Erro_BotaoEntrega_Click

    If GridItens.Row = 0 Then gError 183202

    sProdutoTela = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantPedida_Col)) - IIf(iGrid_QuantCancel_Col <> 0, StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_QuantCancel_Col)), 0)

    If Len(sProdutoTela) = 0 Then gError 183203

    Call Chama_Tela_Modal("DataEntrega", sProdutoTela, dQuantidade, gobjPedidoVenda.colItensPedido(GridItens.Row).colItemPVEntrega, 1, 1)
    
    Call CF2(Me, "PV_Trata_Quantidade_ItemPV")
    
    BotaoEntrega_Click = SUCESSO

    Exit Function

Erro_BotaoEntrega_Click:

    BotaoEntrega_Click = gErr

    Select Case gErr

        Case 183202
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 183203
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183196)

    End Select

    Exit Function
    
End Function

Public Sub BotaoDataRefFluxoDown_Click()

Dim lErro As Long
Dim bCancel As Boolean
Dim sData As String

On Error GoTo Erro_BotaoDataRefFluxoDown_Click

    sData = DataRefFluxo.Text

    'diminui a data em um dia
    lErro = Data_Up_Down_Click(DataRefFluxo, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 188576

    Call DataRefFluxo_Validate(bCancel)
    
    If bCancel = True Then DataRefFluxo.Text = sData
    
    Exit Sub

Erro_BotaoDataRefFluxoDown_Click:

    Select Case gErr

        Case 188576
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188577)

    End Select

    Exit Sub

End Sub

Public Sub BotaoDataRefFluxoUp_Click()

Dim lErro As Long
Dim sData As String
Dim bCancel As Boolean

On Error GoTo Erro_BotaoDataRefFluxoUp_Click

    sData = DataRefFluxo.Text

    'aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataRefFluxo, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 188578
    
    Call DataRefFluxo_Validate(bCancel)

    If bCancel = True Then DataRefFluxo.Text = sData
    
    Exit Sub

Erro_BotaoDataRefFluxoUp_Click:

    Select Case gErr

        Case 188578

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188579)

    End Select

    Exit Sub

End Sub

Public Sub DataRefFluxo_GotFocus()
Dim iDataAux As Integer
    
    iDataAux = giDataRefFluxoAlterada
    Call MaskEdBox_TrataGotFocus(DataRefFluxo, iAlterado)
    giDataRefFluxoAlterada = iDataAux
    
End Sub

Public Sub DataRefFluxo_Change()
    
    iAlterado = REGISTRO_ALTERADO
    giDataRefFluxoAlterada = REGISTRO_ALTERADO

End Sub

Public Sub DataRefFluxo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataRefFluxo_Validate

    If giDataRefFluxoAlterada <> REGISTRO_ALTERADO Then Exit Sub
    
    If Len(Trim(DataRefFluxo.ClipText)) > 0 Then

        'Critica a data digitada
        lErro = Data_Critica(DataRefFluxo.Text)
        If lErro <> SUCESSO Then gError 188560

    End If

    giDataRefFluxoAlterada = 0
        
    Exit Sub

Erro_DataRefFluxo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 188560

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188561)

    End Select

    Exit Sub

End Sub

Public Function PV_Traz_Produto_Tela(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long

Dim dtDataAux As Date

    'Data de Entrega
    dtDataAux = MaskedParaDate(DataEntregaPV)
    If dtDataAux <> DATA_NULA Then GridItens.TextMatrix(GridItens.Row, iGrid_DataEntrega_Col) = Format(dtDataAux, "dd/mm/yyyy")
    
    PV_Traz_Produto_Tela = SUCESSO
    
End Function

Function PV_Cliente_Formata(ByVal objCT As Object, ByVal objCliente As ClassCliente) As Long
    PV_Cliente_Formata = SUCESSO
End Function

Function PV_DataEntregaPV_Validate(ByVal objCT As Object, Cancel As Boolean) As Long
    
On Error GoTo Erro_PV_DataEntregaPV_Validate
    
    If Len(Trim(DataEntregaPV.ClipText)) = 0 Then gError 192970
    
    PV_DataEntregaPV_Validate = SUCESSO
    
    Exit Function

Erro_PV_DataEntregaPV_Validate:

    PV_DataEntregaPV_Validate = gErr
    
    Select Case gErr
    
        Case 192970

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192971)

    End Select

    Exit Function
    
End Function

Public Sub RecebForaFaixa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub RecebForaFaixa_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub RecebForaFaixa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub RecebForaFaixa_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub RecebForaFaixa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = RecebForaFaixa
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PercentMaisReceb_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentMaisReceb_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub PercentMaisReceb_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub PercentMaisReceb_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PercentMaisReceb
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PercentMenosReceb_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentMenosReceb_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub PercentMenosReceb_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub PercentMenosReceb_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = PercentMenosReceb
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_PercentMaisReceb(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercentMaisReceb As Double

On Error GoTo Erro_Saida_Celula_PercentMaisReceb

    Set objGridInt.objControle = PercentMaisReceb

    'Verifica se PercentMaisReceb foi preenchido
    If Len(Trim(PercentMaisReceb.Text)) > 0 Then

        'Critica porcentagem informada
        lErro = Porcentagem_Critica(PercentMaisReceb.Text)
        If lErro <> SUCESSO Then Error 49478

        dPercentMaisReceb = StrParaDbl(PercentMaisReceb.Text)

        'Coloca no formato fixed da tela
        PercentMaisReceb.Text = Format(dPercentMaisReceb, "Fixed")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49479

    Saida_Celula_PercentMaisReceb = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentMaisReceb:

    Saida_Celula_PercentMaisReceb = Err

    Select Case Err

        Case 49478

        Case 49479
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164612)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercentMenosReceb(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercentMenosReceb As Double

On Error GoTo Erro_Saida_Celula_PercentMenosReceb

    Set objGridInt.objControle = PercentMenosReceb

    'Verifica se PercentMenosReceb foi preenchido
    If Len(Trim(PercentMenosReceb.Text)) > 0 Then

        'Critica porcentagem informada
        lErro = Porcentagem_Critica(PercentMenosReceb.Text)
        If lErro <> SUCESSO Then Error 49478

        dPercentMenosReceb = StrParaDbl(PercentMenosReceb.Text)

        'Coloca no formato fixed da tela
        PercentMenosReceb.Text = Format(dPercentMenosReceb, "Fixed")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49479

    Saida_Celula_PercentMenosReceb = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentMenosReceb:

    Saida_Celula_PercentMenosReceb = Err

    Select Case Err

        Case 49478

        Case 49479
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164612)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_RecebForaFaixa(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_RecebForaFaixa

    Set objGridInt.objControle = RecebForaFaixa

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49482

    Saida_Celula_RecebForaFaixa = SUCESSO

    Exit Function

Erro_Saida_Celula_RecebForaFaixa:

    Saida_Celula_RecebForaFaixa = Err

    Select Case Err

        Case 49482
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164625)

    End Select

    Exit Function


End Function

Public Sub QuantFatAMais_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantFatAMais_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub QuantFatAMais_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub QuantFatAMais_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = QuantFatAMais
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Carrega_RecebForaFaixa() As Long

    'Limpa a combo
    RecebForaFaixa.Clear

    RecebForaFaixa.AddItem MENSAGEM_NAO_AVISA_ACEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = NAO_AVISA_E_ACEITA_RECEBIMENTO

    RecebForaFaixa.AddItem MENSAGEM_REJEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = ERRO_E_REJEITA_RECEBIMENTO

    RecebForaFaixa.AddItem MENSAGEM_ACEITA_RECEBIMENTO
    RecebForaFaixa.ItemData(RecebForaFaixa.NewIndex) = AVISA_E_ACEITA_RECEBIMENTO

    Exit Function

End Function

Public Function PV_Exclui_ItemPV(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    PV_Exclui_ItemPV = SUCESSO
End Function

Public Function PV_Saida_Celula_PrecoUnitario1(ByVal objCT As Object, ByVal dPrecoUnitario As Double, ByVal iLinha As Integer) As Long
    PV_Saida_Celula_PrecoUnitario1 = SUCESSO
End Function

Function PV_Preco_Menor_Tabela(ByVal objPedidoVenda As ClassPedidoDeVenda, colPVMotivo As Collection) As Long
    PV_Preco_Menor_Tabela = SUCESSO
End Function

Function PV_Preco_Menor_Tabela_Texto(ByVal objPedidoVenda As ClassPedidoDeVenda, sTexto As String) As Long
    PV_Preco_Menor_Tabela_Texto = SUCESSO
End Function
    
Private Sub BotaoUltVendasCli_Click()

Dim lErro As Long, objPedidoVenda As New ClassPedidoDeVenda
Dim objCliente As New ClassCliente
Dim colSelecao As New Collection, sSelecao As String

On Error GoTo Erro_BotaoUltVendasCli_Click

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 26779

        'Se não achou o Cliente --> erro
        If lErro = 12348 Then gError 26780

        'Guarda código do Cliente em objPedidoVenda
        objPedidoVenda.lCliente = objCliente.lCodigo

    Else
    
        gError 184764
        
    End If

    'Verifica se a Filial está preenchida
    If Len(Trim(Filial.Text)) > 0 Then
        objPedidoVenda.iFilial = Codigo_Extrai(Filial.Text)
    Else
        gError 184765
    End If
    
    colSelecao.Add giFilialEmpresa
    colSelecao.Add objPedidoVenda.lCliente
    colSelecao.Add objPedidoVenda.iFilial
        
    sSelecao = "FilialEmpresa = ? AND Cliente = ? AND Filial = ?"
    
    If Len(Trim(Codigo.Text)) > 0 Then
        objPedidoVenda.lCodigo = CLng(Codigo.Text)
        colSelecao.Add objPedidoVenda.lCodigo
        sSelecao = sSelecao & " AND PVCodigo <> ? AND Posicao < 3"
        Call Chama_Tela("UltPVPorProdLista", colSelecao, Nothing, Nothing, sSelecao)
    Else
        sSelecao = sSelecao & " AND Posicao = 1"
        Call Chama_Tela("UltPVPorProdLista", colSelecao, Nothing, Nothing, sSelecao)
    End If
    
    Exit Sub
    
Erro_BotaoUltVendasCli_Click:

    Select Case gErr

        Case 26779
        
        Case 26780
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case 184764
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 184765
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184763)

    End Select
    
    Exit Sub

End Sub

Public Function PV_Saida_Celula_PrecoUnitario2(ByVal objCT As Object, ByVal dPrecoUnitario As Double, ByVal iLinha As Integer) As Long
    PV_Saida_Celula_PrecoUnitario2 = SUCESSO
End Function

Public Sub PV_Quantidade_GotFocus(ByVal objCT As Object)
'
End Sub

Public Sub PV_Trata_Quantidade_ItemPV(ByVal objCT As Object)
'
End Sub

Public Sub Prioridade_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Prioridade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Prioridade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Prioridade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Prioridade
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PrioridadePadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(PrioridadePadrao, iAlterado)

End Sub

Private Function Saida_Celula_Prioridade(objGridInt As AdmGrid) As Long
'faz a critica da celula de Prioridade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Prioridade

    Set objGridInt.objControle = Prioridade

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 41323

    Saida_Celula_Prioridade = SUCESSO

    Exit Function

Erro_Saida_Celula_Prioridade:

    Saida_Celula_Prioridade = gErr

    Select Case gErr

        Case 41323
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163836)

    End Select

    Exit Function

End Function

Public Sub BotaoExportar_Click()
    Call CF2(Me, "BotaoExportar1_Click")
End Sub

Public Sub BotaoExportar1_Click(ByVal objCT As Object)
End Sub

Private Function Reserva_Processa2(ByVal iIndiceGridItens As Integer, dQuantidade As Double, dQuantidadeCancelada As Double, dQuantidadeFaturada As Double, Optional ByVal bIgnoraConfig As Boolean = False) As Long

Dim lErro As Long
Dim iItemAtual As Integer
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim iIndice As Integer

On Error GoTo Erro_Reserva_Processa2

    'Se não houve Quant a Reservar zera Reservas.
    If dQuantidade = 0 Or dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada = 0 Then
        GridItens.TextMatrix(iIndiceGridItens, iGrid_QuantReservada_Col) = Format(0, "Standard")
        iItemAtual = CInt(GridItens.TextMatrix(iIndiceGridItens, iGrid_ItemProduto_Col))
    End If

    lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndiceGridItens, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 26796

    objProduto.sCodigo = sProduto

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 26669

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 26668

    If objProduto.iControleEstoque = PRODUTO_CONTROLE_RESERVA And dQuantidade > 0 And dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada > 0 And (gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA Or bIgnoraConfig) Then

        lErro = ReservaAlmoxarifadoPadrao(iIndiceGridItens, objProduto, dQuantidade, dQuantidadeCancelada, dQuantidadeFaturada)
        If lErro <> SUCESSO And lErro <> 26680 Then gError 26679

        If gobjFAT.iFaltaEstObrigaDecisao = MARCADO And lErro = 26680 Then gError 26679
    
    ElseIf dQuantidade = 0 And dQuantidade - dQuantidadeCancelada - dQuantidadeFaturada = 0 And gobjFAT.iPedidoReservaAutomatica = PEDVENDA_RESERVA_AUTOMATICA Then
    
        For iIndice = 1 To objGridAlocacao.iLinhasExistentes
                
            If GridAlocacao.TextMatrix(iIndice, iGrid_ProdutoAlmox_Col) = GridItens.TextMatrix(iIndiceGridItens, iGrid_Produto_Col) Then
                Call Remove_Reservas(CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)))
            End If
        Next
        
    End If

    Reserva_Processa2 = SUCESSO

    Exit Function

Erro_Reserva_Processa2:

    Reserva_Processa2 = gErr

    Select Case gErr

        Case 26796, 26669, 26679

        Case 26668
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158029)

    End Select

    Exit Function

End Function


'*************************************
'Funções para tratamento de quebra de itens de grade
Function GrupoGrade_Trata_Itens(ByVal objRomaneioGrade As ClassRomaneioGrade, ByVal objProduto As ClassProduto) As Long

Dim lErro As Long, iIndice As Integer, bAchou As Boolean
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objTabelaPrecoItem As ClassTabelaPrecoItem
Dim objRomaneioGradeAux As ClassRomaneioGrade

On Error GoTo Erro_GrupoGrade_Trata_Itens

    'Se é para separar por preço
    If gobjCRFAT.iSeparaItensGradePrecoDif = MARCADO Then
    
        'Limpa coleção com preços diferentes
        Set objRomaneioGrade.colRomaneioGradeAux = New Collection
    
        'Se tem itens no romaneio grade
        If objRomaneioGrade.colItensRomaneioGrade.Count > 0 Then
        
            iIndice = 0
            'Para cada item
            For Each objItemRomaneioGrade In objRomaneioGrade.colItensRomaneioGrade
            
                If objItemRomaneioGrade.dQuantidade > DELTA_VALORMONETARIO Then
            
                    iIndice = iIndice + 1
                
                    'Obtem o preço
                    Set objTabelaPrecoItem = New ClassTabelaPrecoItem
            
                    objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
                    objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
                    objTabelaPrecoItem.sCodProduto = objItemRomaneioGrade.sProduto
            
                    lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                    If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM
            
                    If lErro = 28014 Then
                        objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
            
                        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                        If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM
            
                    End If
                    
                    objItemRomaneioGrade.dPreco = objTabelaPrecoItem.dPreco
                    objItemRomaneioGrade.sTextoGrade = objTabelaPrecoItem.sTextoGrade
                    
                    'Se o Romaneio ainda está sem preço pega o preço do primeiro item
                    If objRomaneioGrade.dPreco = 0 And iIndice = 1 Then
                        objRomaneioGrade.dPreco = objItemRomaneioGrade.dPreco
                        objRomaneioGrade.sTextoGrade = objItemRomaneioGrade.sTextoGrade
                        Set objRomaneioGrade.objTabelaPrecoItem = objTabelaPrecoItem
                    End If
                
                    'Se o preço do Romaneio divergir do Item gera um novo Romaneio e coloca na coleção auxiliar
                    If Abs(objItemRomaneioGrade.dPreco - objRomaneioGrade.dPreco) > DELTA_VALORMONETARIO2 Then
                        
                        'Procura na coleção auxiliar um Romaneio com mesmo preço
                        bAchou = False
                        For Each objRomaneioGradeAux In objRomaneioGrade.colRomaneioGradeAux
                            If Abs(objItemRomaneioGrade.dPreco - objRomaneioGradeAux.dPreco) < DELTA_VALORMONETARIO2 Then
                                bAchou = True
                                Exit For
                            End If
                        Next
                        'Se não encontrou adiciona
                        If Not bAchou Then
                            Set objRomaneioGradeAux = New ClassRomaneioGrade
                        
                            objRomaneioGradeAux.dPreco = objItemRomaneioGrade.dPreco
                            objRomaneioGradeAux.sTextoGrade = objItemRomaneioGrade.sTextoGrade
                            Set objRomaneioGradeAux.objTabelaPrecoItem = objTabelaPrecoItem
                        
                            objRomaneioGrade.colRomaneioGradeAux.Add objRomaneioGradeAux
                        End If
                        objRomaneioGradeAux.colItensRomaneioGrade.Add objItemRomaneioGrade
                    
                    End If
                    
                End If
        
            Next
        
            'Remove todos os itens com preço diferentes
            For iIndice = objRomaneioGrade.colItensRomaneioGrade.Count To 1 Step -1
                Set objItemRomaneioGrade = objRomaneioGrade.colItensRomaneioGrade.Item(iIndice)
                If Abs(objItemRomaneioGrade.dPreco - objRomaneioGrade.dPreco) > DELTA_VALORMONETARIO2 Then
                    objRomaneioGrade.colItensRomaneioGrade.Remove iIndice
                End If
            Next
            Set objRomaneioGrade.objObjetoTela.colItensRomaneioGrade = objRomaneioGrade.colItensRomaneioGrade
            
            'Adiciona o TextoGrade a Descrição do produto para diferenciar e "explicar" o preço diferente
            If Not (objProduto Is Nothing) Then
                Set objRomaneioGrade.objProduto = objProduto
                If Len(Trim(objRomaneioGrade.sTextoGrade)) > 0 Then objProduto.sDescricao = objProduto.sDescricao & SEPARADOR & objRomaneioGrade.sTextoGrade
            End If
        
        End If
    
    End If
    
    GrupoGrade_Trata_Itens = SUCESSO

    Exit Function

Erro_GrupoGrade_Trata_Itens:

    GrupoGrade_Trata_Itens = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208379)

    End Select

    Exit Function

End Function

Function GrupoGrade_Valida_Itens(ByVal objRomaneioGrade As ClassRomaneioGrade) As Long

Dim lErro As Long

On Error GoTo Erro_GrupoGrade_Valida_Itens

    If gobjCRFAT.iSeparaItensGradePrecoDif = MARCADO Then
    
        If Codigo_Extrai(TabelaPreco.Text) = 0 Then gError 208380 'Sem a tabela de preço não tem como separar os itens
    
    End If
    
    GrupoGrade_Valida_Itens = SUCESSO

    Exit Function

Erro_GrupoGrade_Valida_Itens:

    GrupoGrade_Valida_Itens = gErr

    Select Case gErr
    
        Case 208380 'ERRO_TABELAPRECO_NAO_PREENCHIDA
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208381)

    End Select

    Exit Function

End Function

Function GrupoGrade_Separa_Itens(ByVal objRomaneioGrade As ClassRomaneioGrade) As Long

Dim lErro As Long, iItem As Integer
Dim objRomaneioGradeAux As ClassRomaneioGrade
Dim objProduto As New ClassProduto, objItemPedido As ClassItemPedido
Dim dPrecoUnitario As Double

On Error GoTo Erro_GrupoGrade_Separa_Itens

     If gobjCRFAT.iSeparaItensGradePrecoDif = MARCADO Then
    
        'Se abriu a tela de grade
        If Not (objRomaneioGrade Is Nothing) Then
        
            'Pega o Pai de Grade
            Set objProduto = objRomaneioGrade.objProduto
        
            'Para cada romaneio da coleção auxiliar (preço diferente) insere um item no grid
            For Each objRomaneioGradeAux In objRomaneioGrade.colRomaneioGradeAux
            
                objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
                
                iItem = objGridItens.iLinhasExistentes
            
                'Unidade de Medida
                GridItens.TextMatrix(iItem, iGrid_Produto_Col) = Produto.Text
                GridItens.TextMatrix(iItem, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda
            
                'Descricao Produto
                GridItens.TextMatrix(iItem, iGrid_DescProduto_Col) = left(objProduto.sDescricao, Len(objProduto.sDescricao) - Len(objRomaneioGrade.sTextoGrade)) & IIf(Len(objRomaneioGrade.sTextoGrade) = 0, SEPARADOR, "") & objRomaneioGradeAux.sTextoGrade
            
                If objRomaneioGradeAux.objTabelaPrecoItem.dPreco > 0 Then
                    
                    lErro = CF2(Me, "PV_PrecoUnitario_Calcula", iItem, GridItens.TextMatrix(iItem, iGrid_UnidadeMed_Col), objRomaneioGradeAux.objTabelaPrecoItem, dPrecoUnitario)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                    If dPrecoUnitario > 0 Then
                        GridItens.TextMatrix(iItem, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
                    Else
                        GridItens.TextMatrix(iItem, iGrid_ValorUnitario_Col) = ""
                    End If
                    
                End If
                
                gcolItemPedido.Add 0
               
                Call gobjTelaProjetoInfo.Adiciona_Linha(iItem)
                
                'permite que a tributacao reflita a inclusao de uma linha no grid
                lErro = gobjTribTab.Inclusao_Item_Grid(iItem, objProduto.sCodigo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                gobjPedidoVenda.colItensPedido(iItem).iPossuiGrade = MARCADO
                gobjPedidoVenda.colItensPedido(iItem).sUMEstoque = objProduto.sSiglaUMEstoque
                gobjPedidoVenda.colItensPedido(iItem).iItem = iItem
                gobjPedidoVenda.colItensPedido(iItem).sProduto = objProduto.sCodigo
                           
                Set gobjPedidoVenda.colItens(iItem).colItensRomaneioGrade = objRomaneioGradeAux.colItensRomaneioGrade
                   
                GridItens.TextMatrix(iItem, 0) = "# " & GridItens.TextMatrix(iItem, 0)
                
                Set objItemPedido = gobjPedidoVenda.colItens(iItem)
                       
                Call Atualiza_Grid_Itens(objItemPedido)
                
                lErro = CF2(Me, "PV_Calcula_PesoKg", iItem)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Call CF2(Me, "PV_PrecoTotal_Calcula", iItem)
            
                Call Atualiza_Grid_Alocacao(objItemPedido)
            
            Next
        
            Call ValorTotal_Calcula
        
        End If
    
    End If
   
    GrupoGrade_Separa_Itens = SUCESSO

    Exit Function

Erro_GrupoGrade_Separa_Itens:

    GrupoGrade_Separa_Itens = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208382)

    End Select

    Exit Function

End Function
'Funções para tratamento de quebra de itens de grade
'*************************************

Public Sub BotaoEmail_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objPedidoVendaDB As New ClassPedidoDeVenda
Dim dValor As Double
Dim bBloqueio As Boolean
Dim objBloqueioPV As ClassBloqueioPV
Dim sMailTo As String, objFilialCliente As New ClassFilialCliente, objEndereco As New ClassEndereco

On Error GoTo Erro_BotaoEmail_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Codigo.Text)) = 0 Then gError 93215
    If Len(Trim(Cliente.Text)) = 0 Then gError 93216
    If Len(Trim(Filial.Text)) = 0 Then gError 93217
    If Len(Trim(FilialEntrega.Text)) = 0 Then gError 93218
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 93219
'    If Len(Trim(NaturezaOp.Text)) = 0 Then gError 93230
        
    lErro = Move_Pedido_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then gError 93228

    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 93214

    If lErro = 26509 Then gError 93231
    
    If gobjFAT.iImpedeImpPVComBloq = MARCADO Then
    
        objPedidoVendaDB.lCodigo = objPedidoVenda.lCodigo
        objPedidoVendaDB.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
    
        bBloqueio = False
        lErro = CF("BloqueiosPV_Le", objPedidoVendaDB)
        If lErro <> SUCESSO Then gError 198692
        
        For Each objBloqueioPV In objPedidoVendaDB.colBloqueiosPV
            If objBloqueioPV.dtDataLib = DATA_NULA Then
                bBloqueio = True
                Exit For
            End If
        Next
        
        If bBloqueio Then gError 198693
    
    End If
    
    objFilialCliente.lCodCliente = objPedidoVenda.lCliente
    objFilialCliente.iCodFilial = objPedidoVenda.iFilial
    
    If objFilialCliente.lCodCliente <> 0 And objFilialCliente.iCodFilial Then

        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then gError 124017
         
        If lErro = SUCESSO Then
        
            objEndereco.lCodigo = objFilialCliente.lEndereco
            
            lErro = CF("Endereco_Le", objEndereco)
            If lErro <> SUCESSO Then gError 124018
        
            sMailTo = objEndereco.sEmail
            
        End If
        
    End If
    
    lErro = objRelatorio.ExecutarDiretoEmail("Pedido de Venda", "PedidoVenda >= @NPEDVENDINIC E PedidoVenda <= @NPEDVENDFIM", 0, "PedVenda", "NPEDVENDINIC", objPedidoVenda.lCodigo, "NPEDVENDFIM", objPedidoVenda.lCodigo, "TTO_EMAIL", sMailTo, "TSUBJECT", "Pedido de Venda " & CStr(objPedidoVenda.lCodigo), "TALIASATTACH", "PV" & CStr(objPedidoVenda.lCodigo))
    If lErro <> SUCESSO Then gError 93229
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoEmail_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 93214, 93229, 198692

        Case 93215
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)

        Case 93216
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 93217
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)

        Case 93218
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)
        
        Case 93219
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 93230
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 93231
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoVenda.lCodigo)

        Case 198693
            Call Rotina_Erro(vbOKOnly, "ERRO_PV_COM_BLOQ_IMP_IMPRESSAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158045)

    End Select

    Exit Sub

End Sub

Public Function PV_Trata_CodigoBarras1(ByVal objCT As Object) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoEnxuto As String
Dim sCodBarras As String
Dim sCodBarrasOriginal As String
Dim dCusto As Double

On Error GoTo Erro_PV_Trata_CodigoBarras1

    If objCT.objGridItens.iLinhasExistentes + 1 = objCT.GridItens.Row Then
    
        'Verifica se o Produto está preenchido
        If Len(Trim(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Produto_Col))) = 0 Then
            
            If objCT.ActiveControl Is objCT.Produto Then
                    
                    Set objCT.objGridItens.objControle = objCT.Produto
            
                    lErro = Grid_Abandona_Celula(objCT.objGridItens)
                    If lErro <> SUCESSO Then gError 199540
                    
            End If
            
            objProduto.lErro = 1
    
            Call Chama_Tela_Modal("CodigoBarras", objProduto)
    
            
            If objProduto.sCodigoBarras <> "Cancel" Then
                If objProduto.lErro = SUCESSO Then
    
                    lErro = CF("INV_Trata_CodigoBarras", objProduto)
                    If lErro <> SUCESSO Then gError 199366
    
                End If
    
                'Lê os demais atributos do Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 199305

                'Se não encontrou o Produto --> Erro
                If lErro = 28030 Then gError 199306

                lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
                If lErro <> SUCESSO Then gError 199307
        
                objCT.Show
        
                objCT.Produto.PromptInclude = False
                objCT.Produto.Text = sProdutoEnxuto
                objCT.Produto.PromptInclude = True
                
                If Not objCT.ActiveControl Is objCT.Produto Then
                    objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Produto_Col) = objCT.Produto.Text
                
                    lErro = CF2(Me, "Produto_Saida_Celula")
                    If lErro <> SUCESSO And lErro <> 26658 Then gError 26606
                    If lErro = 26658 Then gError 51366
                
                
'                    'Preenche a Linha do Grid
'                    lErro = objCT.ProdutoLinha_Preenche(objProduto)
'                    If lErro <> SUCESSO Then gError 199305
'
'                    'calcula o custo atual do produto passado como parametro e devolve-o de dCusto
'                    lErro = CF("Calcula_Custo_Atual", objProduto, dCusto, objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Tipo_Col))
'                    If lErro <> SUCESSO Then gError 199305
'
'                    objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_CustoUnitario_Col) = Formata_Custo(dCusto)
'
'                    If Len(Trim(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_LoteProduto_Col))) = 0 Then
'
'                        lErro = objCT.QuantEstoque_Calcula(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Produto_Col), objCT.GridItens.Row)
'                        If lErro <> SUCESSO Then gError 199305
'
'                    Else
'
'                        lErro = objCT.QuantLote_Calcula(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Produto_Col), objCT.GridItens.Row, objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_LoteProduto_Col), Codigo_Extrai(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_FilialOP_Col)))
'                        If lErro <> SUCESSO Then gError 199305
'
'                    End If
'
'                    lErro = objCT.Preenche_ContaContabil()
'                    If lErro <> SUCESSO Then gError 52240
                
                End If
    
            Else
            
                gError 199834
    
    
            End If
    
        End If
    
    End If

    PV_Trata_CodigoBarras1 = SUCESSO

    Exit Function

Erro_PV_Trata_CodigoBarras1:

    PV_Trata_CodigoBarras1 = gErr

    objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Produto_Col) = ""

    Select Case gErr

        Case 199305, 199308, 199363, 199366, 199540, 199834

        Case 199306, 199364
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 199307
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199309)

    End Select

    Exit Function

End Function

Private Function Carrega_TipoFrete() As Long
'Carrega versaoNFE

Dim lErro As Long
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer
Dim colTipoFrete As Collection

On Error GoTo Erro_Carrega_TipoFrete
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207405

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
        Case 207405

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207406)

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
            If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col)) > 0 And gobjPedidoVenda.colItens.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
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
            If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col)) > 0 And gobjPedidoVenda.colItens.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
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
                If gobjPedidoVenda.colItens.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
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

Public Sub BotaoImprimirConf_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoImprimirConf_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Codigo.Text)) = 0 Then gError 211460
    
    objPedidoVenda.lCodigo = StrParaLong(Codigo.Text)
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa
    
    lErro = Move_Pedido_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError ERRO_SEM_MENSAGEM

    If lErro = 26509 Then gError 211461
    
    lErro = objRelatorio.ExecutarDireto("Pedido de Venda - Conferência do Estoque", "PedidoVenda >= @NPEDVENDINIC E PedidoVenda <= @NPEDVENDFIM", 1, "PVSEP", "NPEDVENDINIC", objPedidoVenda.lCodigo, "NPEDVENDFIM", objPedidoVenda.lCodigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoImprimirConf_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 211460
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)

        Case 211461
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoVenda.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 211462)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Public Sub Parc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Parc_GotFocus()
    Call MaskEdBox_TrataGotFocus(Parc, iAlterado)
End Sub

Public Sub ComissaoItemPV_Change()

    iAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO

End Sub

Public Sub ComissaoItemPV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub ComissaoItemPV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub ComissaoItemPV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ComissaoItemPV
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ComissaoItemPV(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dComissaoItemPV As Double
Dim dPercentualCampo As Double
Dim dPercentualGrid As Double

On Error GoTo Erro_Saida_Celula_ComissaoItemPV

    Set objGridInt.objControle = ComissaoItemPV

    dPercentualGrid = PercentParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ComissaoItem_Col))
    dPercentualGrid = dPercentualGrid * 100

    dPercentualCampo = StrParaDbl(ComissaoItemPV.Text)
    
    'Verifica se ComissaoItemPV foi preenchido
    If Len(Trim(ComissaoItemPV.Text)) > 0 Then

        'Critica porcentagem informada
        lErro = Porcentagem_Critica(ComissaoItemPV.Text)
        If lErro <> SUCESSO Then Error 49478

        dComissaoItemPV = StrParaDbl(ComissaoItemPV.Text)

        'Coloca no formato fixed da tela
        ComissaoItemPV.Text = Format(dComissaoItemPV, "Fixed")

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49479

    Saida_Celula_ComissaoItemPV = SUCESSO

    Exit Function

Erro_Saida_Celula_ComissaoItemPV:

    Saida_Celula_ComissaoItemPV = Err

    Select Case Err

        Case 49478

        Case 49479
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164612)

    End Select

    Exit Function

End Function

Public Sub TabPrecoItemPV_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TabPrecoItemPV_Click()

    iAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO

End Sub

Public Sub TabPrecoItemPV_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub TabPrecoItemPV_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub TabPrecoItemPV_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = TabPrecoItemPV
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_TabPrecoItemPV(objGridInt As AdmGrid) As Long

Dim lErro As Long, iLinha As Integer, iItem As Integer
Dim iTabelaAnterior As Integer, iTabelaAtual As Integer

On Error GoTo Erro_Saida_Celula_TabPrecoItemPV

    Set objGridInt.objControle = TabPrecoItemPV

    iLinha = GridItens.Row
    
    For iItem = 0 To TabPrecoItemPV.ListCount - 1
        
        If GridItens.TextMatrix(iLinha, iGrid_TabPrecoItem_Col) = TabPrecoItemPV.List(iItem) Then

            iTabelaAnterior = TabPrecoItemPV.ItemData(iItem)
            Exit For

        End If
        
    Next
    
    GridItens.TextMatrix(GridItens.Row, iGrid_TabPrecoItem_Col) = TabPrecoItemPV.Text
    
    For iItem = 0 To TabPrecoItemPV.ListCount - 1
        
        If GridItens.TextMatrix(iLinha, iGrid_TabPrecoItem_Col) = TabPrecoItemPV.List(iItem) Then

            iTabelaAtual = TabPrecoItemPV.ItemData(iItem)
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

    Saida_Celula_TabPrecoItemPV = SUCESSO

    Exit Function

Erro_Saida_Celula_TabPrecoItemPV:

    Saida_Celula_TabPrecoItemPV = Err

    Select Case Err

        Case 49482, ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164625)

    End Select

    Exit Function

End Function

Private Function Trata_Kit_de_Venda(ByVal objProd As ClassProduto, ByVal objKitVenda As ClassKitVenda, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim vbMsgBox As VbMsgBoxResult
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim iItem As Integer, sProdutoEnxuto As String
Dim objProduto As ClassProduto
Dim objItemPedido As ClassItemPedido
Dim objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double
Dim colItensRomaneioGrade As Collection

On Error GoTo Erro_Trata_Kit_de_Venda

    vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_SUBSTITUIR_KIT_VENDA", objKitVenda.sProduto, Formata_Estoque(objKitVenda.dQuantidade), objKitVenda.sUM)
    If vbMsgBox = vbYes Then
    
        Set colItensRomaneioGrade = gobjPedidoVenda.colItensPedido(iLinha).colItensRomaneioGrade
    
        Call GridItens_KeyDown(vbKeyDelete, 0, False)
    
        For Each objItemRomaneio In colItensRomaneioGrade
        
            Set objProduto = New ClassProduto
    
            objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
            
            iItem = objGridItens.iLinhasExistentes
            
            objProduto.sCodigo = objItemRomaneio.sProduto
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
            lErro = Mascara_RetornaProdutoEnxuto(objItemRomaneio.sProduto, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            'Mascara o produto enxuto
            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True
        
            'Unidade de Medida
            GridItens.TextMatrix(iItem, iGrid_Produto_Col) = Produto.Text
            GridItens.TextMatrix(iItem, iGrid_UnidadeMed_Col) = objItemRomaneio.sUMEstoque
        
            'Descricao Produto
            GridItens.TextMatrix(iItem, iGrid_DescProduto_Col) = objProduto.sDescricao
            GridItens.TextMatrix(iItem, iGrid_QuantPedida_Col) = Formata_Estoque(objItemRomaneio.dQuantidade)
                       
            'Preço Unitário
            If Len(Trim(TabelaPreco.Text)) > 0 Then
            
                Set objTabelaPrecoItem = New ClassTabelaPrecoItem
                    
                objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
                objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
                objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
        
                lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM
        
                If lErro = 28014 Then
                    objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
        
                    lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                    If lErro <> SUCESSO And lErro <> 28014 Then gError ERRO_SEM_MENSAGEM
        
                End If
                If lErro <> 28014 Then
        
                    lErro = CF2(Me, "PV_PrecoUnitario_Calcula", iItem, GridItens.TextMatrix(iItem, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
                    If lErro <> SUCESSO Then gError 26663
        
                    If dPrecoUnitario > 0 Then
                        GridItens.TextMatrix(iItem, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
                    Else
                        GridItens.TextMatrix(iItem, iGrid_ValorUnitario_Col) = ""
                    End If
        
                End If
        
            End If
            
            gcolItemPedido.Add 0
           
            Call gobjTelaProjetoInfo.Adiciona_Linha(iItem)
            
            'permite que a tributacao reflita a inclusao de uma linha no grid
            lErro = gobjTribTab.Inclusao_Item_Grid(iItem, objProduto.sCodigo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            gobjPedidoVenda.colItensPedido(iItem).iPossuiGrade = DESMARCADO
            gobjPedidoVenda.colItensPedido(iItem).sUMEstoque = objProduto.sSiglaUMEstoque
            gobjPedidoVenda.colItensPedido(iItem).iItem = iItem
            gobjPedidoVenda.colItensPedido(iItem).sProduto = objProduto.sCodigo
                                   
            Set objItemPedido = gobjPedidoVenda.colItens(iItem)
            
            objItemPedido.dQuantidade = objItemRomaneio.dQuantidade

            Call CF2(Me, "PV_PrecoTotal_Calcula", iItem)
        
            Call Atualiza_Grid_Alocacao(objItemPedido)
            
            lErro = CF2(Me, "PV_Calcula_PesoKg", iItem)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            lErro = gobjTribTab.Alteracao_Item_Grid(iItem)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Next
    
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        Call ValorTotal_Calcula
    
    End If
    
    Trata_Kit_de_Venda = SUCESSO

    Exit Function

Erro_Trata_Kit_de_Venda:

    Trata_Kit_de_Venda = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 164625)

    End Select

    Exit Function

End Function

Public Sub BotaoAnaliseVenda_Click()

Dim objAnaliseVendas As New ClassAnaliseVendaInfo
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim lErro As Long

On Error GoTo Erro_BotaoAnaliseVenda_Click

    Call Move_Pedido_Memoria(objPedidoVenda)

    If objPedidoVenda.colItens.Count = 0 Then gError 201541

    Call objPedidoVenda.AnaliseVenda_Preenche(objAnaliseVendas)
    
    Call Chama_Tela_Modal("AnaliseVenda", objAnaliseVendas)

    Exit Sub
    
Erro_BotaoAnaliseVenda_Click:

    Select Case gErr

        Case 201541
            Call Rotina_Erro(vbOKOnly, "ERRO_ANALISE_VENDA_SEM_INFO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201542)

    End Select
    
    Exit Sub

End Sub

Public Sub BotaoImportarItens_Click()

Dim lErro As Long, sProdutoEnxuto As String
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objItem As ClassItemPedido, iIndice As Integer

On Error GoTo Erro_BotaoImportarItens_Click

    Call Chama_Tela_Modal("PVImportItens", objPedidoVenda)
    
    bPulaValorTotalCalcula = True

    iIndice = 0
    For Each objItem In objPedidoVenda.colItens
        iIndice = iIndice + 1
        GridItens.Row = objGridItens.iLinhasExistentes + 1
        GridItens.Col = iGrid_Produto_Col
    
        lErro = Mascara_RetornaProdutoEnxuto(objItem.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
    
        lErro = CF2(Me, "Produto_Saida_Celula")
        If lErro <> SUCESSO And lErro <> 26658 Then gError ERRO_SEM_MENSAGEM
        
        GridItens.Col = iGrid_QuantPedida_Col
        
        Quantidade.Text = Formata_Estoque(objItem.dQuantidade)
        
        Call Saida_Celula_Quantidade(objGridItens)
        
        If objItem.dPrecoUnitario > 0 Then
        
            GridItens.Col = iGrid_ValorUnitario_Col
            
            PrecoUnitario.Text = Format(objItem.dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            
            Call Saida_Celula_PrecoUnitario(objGridItens)
        
        End If
        
    Next
    
    bPulaValorTotalCalcula = False
    iTratamentoPadraoFaltaEstoque = 0

    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Recalcula o peso bruto e o peso líquido
    lErro = Pesos_Recalcula()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF2(Me, "PV_Calcula_PesoKg_Total")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub
    
Erro_BotaoImportarItens_Click:

    bPulaValorTotalCalcula = False
    iTratamentoPadraoFaltaEstoque = 0

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201542)

    End Select
    
    Exit Sub

End Sub
