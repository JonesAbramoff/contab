VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTOrcamentoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

Public gobjAnotacao As ClassAnotacoes
Public gobjInfoUsu As Object
Public gobjInfoAdic As ClassInfoAdic
Dim dValorDescontoItensAnt As Double
Dim dPercDescontoItensAnt As Double
Dim iDescontoAlterado As Integer

Public gobjTelaProjetoInfo As ClassTelaPRJInfo
Public gobjIdiomaTela As ClassIdiomaTela

'para a aba de formaçao de preços
Public gcolItensOVPrecoCalc As Collection
Dim gcolPCSituacoes As Collection
Dim gcolFPSituacoes As Collection
Public gobjTribTab As New ClassTribTab

Private iCodigoAlterado As Integer
Dim lMotivoPerdaAnt As Long
Dim iMoedaAnt As Integer
Dim iPrazoAnt As Integer
Dim iIdiomaAnt As Integer

'Grid Itens
Public iGrid_Item_Col As Integer
Public iGrid_UMEstoque_Col As Integer
Public iGrid_Responsavel_Col As Integer
Public iGrid_ItemProduto_Col As Integer
Public iGrid_Produto_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
Public iGrid_PrecoUnitario_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PrecoTotal_Col As Integer
Public iGrid_PrecoTotalB_Col As Integer
Public iGrid_DataEntrega_Col As Integer
Public iGrid_PrazoEntrega_Col As Integer
Public iGrid_Status_Col As Integer
Public iGrid_MotivoPerda_Col As Integer
Public iGrid_Observacao_Col As Integer
Public iGrid_SomaItem_Col As Integer

Const VALOR_CODIGO = 1
Const VALOR_CODIGOBASE = 0

Public iGrid_VersaoKit_Col As Integer
Public iGrid_VersaoKitBase_Col As Integer

Private WithEvents objEventoNumeroBase As AdmEvento
Attribute objEventoNumeroBase.VB_VarHelpID = -1
Private WithEvents objEventoVersao As AdmEvento
Attribute objEventoVersao.VB_VarHelpID = -1

'Grid Parcelas
Public iGrid_Vencimento_Col  As Integer
Public iGrid_ValorParcela_Col As Integer
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

'Grid PrecosCalculados
Dim iGrid_PCSelecionado_Col  As Integer
Dim iGrid_PCProduto_Col  As Integer
Dim iGrid_PCDescProduto_Col  As Integer
Dim iGrid_PCUnidMed_Col  As Integer
Dim iGrid_PCQtde_Col  As Integer
Dim iGrid_PCPrecoUnitCalc_Col  As Integer
Dim iGrid_PCPrecoUnit_Col  As Integer
Dim iGrid_PCPrecoTotal_Col  As Integer
Dim iGrid_PCSituacao_Col  As Integer

'Grid FormacaoPreco
Dim iGrid_FPProduto_Col  As Integer
Dim iGrid_FPDescricaoProduto_Col  As Integer
Dim iGrid_FPUnidMed_Col  As Integer
Dim iGrid_FPQtde_Col  As Integer
Dim iGrid_FPCustoUnit_Col  As Integer
Dim iGrid_FPPercentMargem_Col  As Integer
Dim iGrid_FPPrecoUnit_Col  As Integer
Dim iGrid_FPPrecoTotal_Col  As Integer
Dim iGrid_FPSituacao_Col  As Integer

Private WithEvents objEventoBotaoPedAtualizados As AdmEvento
Attribute objEventoBotaoPedAtualizados.VB_VarHelpID = -1

Dim giFrameAtual As Integer
Dim gsCodigoAnterior  As String
Dim giFilialEntregaManual As Integer

Public iAlterado As Integer
Public iVendedorAlterado As Integer
Public iVendedor2Alterado As Integer

Public giClienteAlterado As Integer
Public giFilialAlterada As Integer
Public gdDesconto As Double
Public giValorFreteAlterado As Integer
Public giValorSeguroAlterado As Integer
Public giValorDescontoAlterado As Integer
Public giValorDespesasAlterado  As Integer
Public giDataReferenciaAlterada As Integer
Dim giDataEntregaAlterada As Integer
Dim giFilialEntregaAlterada As Integer

Dim giValorDescontoManual As Integer

Dim giPercAcresFinAlterado As Integer

Public gobjOrcamentoVenda As New ClassOrcamentoVenda 'estrutura mantida para auxiliar a manutencao de informacoes p/tributacao
    'todos os dados relevantes p/tributacao dentro de gobjOrcamentoVenda estarao sincronizados com a tela antes da atualizacao da tributacao

Public objGridItens As AdmGrid
Dim objGridParcelas As AdmGrid
Dim objGridPrecosCalculados As AdmGrid
Dim objGridFormacaoPreco As AdmGrid

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoCondPagto As AdmEvento
Attribute objEventoCondPagto.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Public WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Public WithEvents objEventoVendedor2 As AdmEvento
Attribute objEventoVendedor2.VB_VarHelpID = -1
Private WithEvents objEventoNaturezaOp As AdmEvento
Attribute objEventoNaturezaOp.VB_VarHelpID = -1
Private WithEvents objEventoTiposDeTributacao As AdmEvento
Attribute objEventoTiposDeTributacao.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoTranspRedespacho As AdmEvento
Attribute objEventoTranspRedespacho.VB_VarHelpID = -1
Private WithEvents objEventoMensagem As AdmEvento
Attribute objEventoMensagem.VB_VarHelpID = -1
Private WithEvents objEventoEstoqueProd As AdmEvento
Attribute objEventoEstoqueProd.VB_VarHelpID = -1


Dim giLinhaAnterior As Integer

'variaveis auxiliares para criacao da contabilizacao
'Private gobjContabAutomatica As ClassContabAutomatica
'Private gobjNFiscalCTB As ClassNFiscal
'Private gobjOrcamentoVendaCTB As ClassOrcamentoVenda
'Private giExercicio As Integer, giPeriodo As Integer
'Private gcolAlmoxFilial As New Collection
'Private gobjGeracaoNFiscal As ClassGeracaoNFiscal

'Constantes públicas dos tabs
Private Const TAB_Principal = 1
Private Const TAB_Itens = 2
Private Const TAB_Complemento = 3
Private Const TAB_Cobranca = 4
Private Const TAB_Tributacao = 5

'Property Variables:
Dim m_Caption As String
Public gbCarregandoTela As Boolean
Public iFrameAtual As Integer
Dim bTrazendoDoc As Boolean
Dim giPosCargaOk As Integer

Event Unload()
    
'--- inicio dos properties get dos controles da tela
Public Property Get Idioma() As Object
     Set Idioma = objUserControl.Controls("Idioma")
End Property

Public Property Get BotaoSobe() As Object
     Set BotaoSobe = objUserControl.Controls("BotaoSobe")
End Property
Public Property Get BotaoDesce() As Object
     Set BotaoDesce = objUserControl.Controls("BotaoDesce")
End Property
Public Property Get FrameTS() As Object
     Set FrameTS = objUserControl.Controls("FrameTS")
End Property
Public Property Get TS() As Object
     Set TS = objUserControl.Controls("TS")
End Property
Public Property Get SomaItem() As Object
     Set SomaItem = objUserControl.Controls("SomaItem")
End Property

Public Property Get Moeda() As Object
     Set Moeda = objUserControl.Controls("Moeda")
End Property

Public Property Get DataPerda() As Object
     Set DataPerda = objUserControl.Controls("DataPerda")
End Property

Public Property Get PrazoEntVar() As Object
     Set PrazoEntVar = objUserControl.Controls("PrazoEntVar")
End Property

Public Property Get PrazoTexto() As Object
     Set PrazoTexto = objUserControl.Controls("PrazoTexto")
End Property

Public Property Get ImprimirOVComCodProd() As Object
     Set ImprimirOVComCodProd = objUserControl.Controls("ImprimirOVComCodProd")
End Property

Public Property Get CalcularST() As Object
     Set CalcularST = objUserControl.Controls("CalcularST")
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

Public Property Get ImprimirOVComPreco() As Object
     Set ImprimirOVComPreco = objUserControl.Controls("ImprimirOVComPreco")
End Property

Public Property Get ISSValor1() As Object
     Set ICMSValor1 = objUserControl.Controls("ICMSValor1")
End Property

Public Property Get GridPrecosCalculados() As Object
     Set GridPrecosCalculados = objUserControl.Controls("GridPrecosCalculados")
End Property

Public Property Get GridFormacaoPreco() As Object
     Set GridFormacaoPreco = objUserControl.Controls("GridFormacaoPreco")
End Property

Public Property Get PCSelecionado() As Object
     Set PCSelecionado = objUserControl.Controls("PCSelecionado")
End Property

Public Property Get PCProduto() As Object
     Set PCProduto = objUserControl.Controls("PCProduto")
End Property

Public Property Get PCDescProduto() As Object
     Set PCDescProduto = objUserControl.Controls("PCDescProduto")
End Property

Public Property Get PCUnidMed() As Object
     Set PCUnidMed = objUserControl.Controls("PCUnidMed")
End Property

Public Property Get PCQtde() As Object
     Set PCQtde = objUserControl.Controls("PCQtde")
End Property

Public Property Get PCPrecoUnitCalc() As Object
     Set PCPrecoUnitCalc = objUserControl.Controls("PCPrecoUnitCalc")
End Property

Public Property Get PCPrecoUnit() As Object
     Set PCPrecoUnit = objUserControl.Controls("PCPrecoUnit")
End Property

Public Property Get PCPrecoTotal() As Object
     Set PCPrecoTotal = objUserControl.Controls("PCPrecoTotal")
End Property

Public Property Get PCSituacao() As Object
     Set PCSituacao = objUserControl.Controls("PCSituacao")
End Property

Public Property Get FPProduto() As Object
     Set FPProduto = objUserControl.Controls("FPProduto")
End Property

Public Property Get FPDescricaoProduto() As Object
     Set FPDescricaoProduto = objUserControl.Controls("FPDescricaoProduto")
End Property

Public Property Get FPUnidMed() As Object
     Set FPUnidMed = objUserControl.Controls("FPUnidMed")
End Property

Public Property Get FPQtde() As Object
     Set FPQtde = objUserControl.Controls("FPQtde")
End Property

Public Property Get FPCustoUnit() As Object
     Set FPCustoUnit = objUserControl.Controls("FPCustoUnit")
End Property

Public Property Get FPPercentMargem() As Object
     Set FPPercentMargem = objUserControl.Controls("FPPercentMargem")
End Property

Public Property Get FPPrecoUnit() As Object
     Set FPPrecoUnit = objUserControl.Controls("FPPrecoUnit")
End Property

Public Property Get FPPrecoTotal() As Object
     Set FPPrecoTotal = objUserControl.Controls("FPPrecoTotal")
End Property

Public Property Get FPSituacao() As Object
     Set FPSituacao = objUserControl.Controls("FPSituacao")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get CobrancaAutomatica() As Object
     Set CobrancaAutomatica = objUserControl.Controls("CobrancaAutomatica")
End Property

Public Property Get SSFrame3() As Object
     Set SSFrame3 = objUserControl.Controls("SSFrame3")
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

Public Property Get BotaoVersaoKitBase() As Object
     Set BotaoVersaoKitBase = objUserControl.Controls("BotaoVersaoKitBase")
End Property

Public Property Get BotaoGrade() As Object
     Set BotaoGrade = objUserControl.Controls("BotaoGrade")
End Property

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
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

Public Property Get ValorTotal() As Object
     Set ValorTotal = objUserControl.Controls("ValorTotal")
End Property

Public Property Get VersaoKit() As Object
     Set VersaoKit = objUserControl.Controls("VersaoKit")
End Property

Public Property Get VersaoKitBase() As Object
     Set VersaoKitBase = objUserControl.Controls("VersaoKitBase")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

Public Property Get DescricaoProduto() As Object
     Set DescricaoProduto = objUserControl.Controls("DescricaoProduto")
End Property

Public Property Get UnidadeMed() As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
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

Public Property Get PrecoTotal() As Object
     Set PrecoTotal = objUserControl.Controls("PrecoTotal")
End Property

Public Property Get GridItens() As Object
     Set GridItens = objUserControl.Controls("GridItens")
End Property

Public Property Get EmailOrcamentoGravacao() As Object
     Set EmailOrcamentoGravacao = objUserControl.Controls("EmailOrcamentoGravacao")
End Property

Public Property Get PrazoValidade() As Object
     Set PrazoValidade = objUserControl.Controls("PrazoValidade")
End Property

Public Property Get Vendedor() As Object
     Set Vendedor = objUserControl.Controls("Vendedor")
End Property

Public Property Get Vendedor2() As Object
     Set Vendedor2 = objUserControl.Controls("Vendedor2")
End Property

Public Property Get Contato() As Object
     Set Contato = objUserControl.Controls("Contato")
End Property

Public Property Get Email() As Object
     Set Email = objUserControl.Controls("Email")
End Property

Public Property Get Usuario() As Object
     Set Usuario = objUserControl.Controls("Usuario")
End Property

Public Property Get UsuarioLabel() As Object
     Set UsuarioLabel = objUserControl.Controls("UsuarioLabel")
End Property

Public Property Get VendedorLabel() As Object
     Set VendedorLabel = objUserControl.Controls("VendedorLabel")
End Property

Public Property Get Vendedor2Label() As Object
     Set Vendedor2Label = objUserControl.Controls("Vendedor2Label")
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

Public Property Get UpDownEmissao() As Object
     Set UpDownEmissao = objUserControl.Controls("UpDownEmissao")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get Codigo() As Object
     Set Codigo = objUserControl.Controls("Codigo")
End Property

Public Property Get CodigoBase() As Object
     Set CodigoBase = objUserControl.Controls("CodigoBase")
End Property

Public Property Get NumeroBaseLabel() As Object
     Set NumeroBaseLabel = objUserControl.Controls("NumeroBaseLabel")
End Property

Public Property Get NaturezaLabel() As Object
     Set NaturezaLabel = objUserControl.Controls("NaturezaLabel")
End Property

Public Property Get NumeroLabel() As Object
     Set NumeroLabel = objUserControl.Controls("NumeroLabel")
End Property

Public Property Get BotaoEmail() As Object
     Set BotaoEmail = objUserControl.Controls("BotaoEmail")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoImprimir() As Object
     Set BotaoImprimir = objUserControl.Controls("BotaoImprimir")
End Property

Public Property Get ImprimeOrcamentoGravacao() As Object
     Set ImprimeOrcamentoGravacao = objUserControl.Controls("ImprimeOrcamentoGravacao")
End Property

Public Property Get Opcao() As Object
     Set Opcao = objUserControl.Controls("Opcao")
End Property

Public Property Get StatusItem() As Object
     Set StatusItem = objUserControl.Controls("StatusItem")
End Property

Public Property Get Status() As Object
     Set Status = objUserControl.Controls("Status")
End Property

Public Property Get StatusComercial() As Object
     Set StatusComercial = objUserControl.Controls("StatusComercial")
End Property

Public Property Get MotivoPerdaItem() As Object
     Set MotivoPerdaItem = objUserControl.Controls("MotivoPerdaItem")
End Property

Public Property Get MotivoPerda() As Object
     Set MotivoPerda = objUserControl.Controls("MotivoPerda")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

Public Property Get OVVersao() As Object
     Set OVVersao = objUserControl.Controls("OVVersao")
End Property

Public Property Get TrocarVersao() As Object
     Set TrocarVersao = objUserControl.Controls("TrocarVersao")
End Property

Public Property Get DataUltAlt() As Object
     Set DataUltAlt = objUserControl.Controls("DataUltAlt")
End Property

Public Property Get HoraUltAlt() As Object
     Set HoraUltAlt = objUserControl.Controls("HoraUltAlt")
End Property

Public Property Get FrameVersao() As Object
     Set FrameVersao = objUserControl.Controls("FrameVersao")
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

Public Property Get BotaoInfoAdicDocItem() As Object
     Set BotaoInfoAdicDocItem = objUserControl.Controls("BotaoInfoAdicDocItem")
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

Public Property Get DataEnvio() As Object
     Set DataEnvio = objUserControl.Controls("DataEnvio")
End Property

Public Property Get DataProxContato() As Object
     Set DataProxContato = objUserControl.Controls("DataProxContato")
End Property

Public Property Get DataFechamento() As Object
     Set DataFechamento = objUserControl.Controls("DataFechamento")
End Property

Public Property Get PrazoEntrega() As Object
     Set PrazoEntrega = objUserControl.Controls("PrazoEntrega")
End Property

Public Property Get PrazoEntregaItem() As Object
     Set PrazoEntregaItem = objUserControl.Controls("PrazoEntregaItem")
End Property

Public Property Get PedidoRepr() As Object
     Set PedidoRepr = objUserControl.Controls("PedidoRepr")
End Property

Public Property Get TipoFrete() As Object
     Set TipoFrete = objUserControl.Controls("TipoFrete")
End Property

Public Property Get OptDataPrazoEnt() As Object
     Set OptDataPrazoEnt = objUserControl.Controls("OptDataPrazoEnt")
End Property

Public Property Get FrameDataPrazoEnt() As Object
     Set FrameDataPrazoEnt = objUserControl.Controls("FrameDataPrazoEnt")
End Property

'--- fim dos properties get dos controles da tela

Public Sub Form_Load()

Dim lErro As Long
Dim objContainer As Object
Dim objUserControl As Object

On Error GoTo Erro_Form_Load

    'precodesc
    Set objContainer = Frame2(3)
    Set objUserControl = Me

    Email.MaxLength = STRING_EMAIL
    Contato.MaxLength = STRING_CONTATO

    'funcao usada para alguns clientes customizarem o  Form_Load
    lErro = CF("Orcamento_Form_Load", objUserControl, objContainer)
    If lErro <> SUCESSO Then gError 126500
    
    If gobjFAT.iOVExibeOrdItens = MARCADO Then
        BotaoSobe.Visible = True
        BotaoDesce.Visible = True
    End If
    If gobjFAT.iOVExibeTS = MARCADO Then
        FrameTS.Visible = True
    End If
    
    Set objEventoNumeroBase = New AdmEvento
    Set objEventoVersao = New AdmEvento
    Usuario.Caption = gsUsuario

    Set gobjIdiomaTela = New ClassIdiomaTela
    
    lErro = gobjIdiomaTela.Iniciar(Me.Name)
    If lErro <> SUCESSO Then gError 126500
    
    lErro = gobjIdiomaTela.Carrega_Combo_Idioma(Idioma)
    If lErro <> SUCESSO Then gError 126500
    
    iIdiomaAnt = Codigo_Extrai(Idioma.Text)
    
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        BotaoVersaoKitBase.Visible = True
    Else
        BotaoVersaoKitBase.Visible = False
    End If
    
    If gobjCRFAT.iOVGravacaoVersao = OV_VERSAO_NAO_GRAVA Then
        TrocarVersao.Value = vbUnchecked
        FrameVersao.Enabled = False
    Else
        FrameVersao.Enabled = True
        TrocarVersao.Value = vbChecked
    End If
    
    If gobjFAT.iExibeOPImpPreco = MARCADO Then
        ImprimirOVComPreco.Visible = True
    End If

    giPosCargaOk = 0
    
    giFrameAtual = 1
    
    'Preenche Data Referencia e Data de Emissão coma Data Atual
    DataReferencia.PromptInclude = False
    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataReferencia.PromptInclude = True
    giDataReferenciaAlterada = 0
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    lErro = Carrega_TipoDesconto
    If lErro <> SUCESSO Then gError 103050

    Call Carrega_Status(Status)
    Call Carrega_Status(StatusItem)
    Call Carrega_Motivo(MotivoPerda)
    Call Carrega_Motivo(MotivoPerdaItem)
    
    lErro = CF2(Me, "OV_Form_Load")
    If lErro <> SUCESSO Then gError 141406
    
    iCodigoAlterado = 0
    iAlterado = 0
    iDescontoAlterado = 0
    
    iVendedorAlterado = 0
    iVendedor2Alterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 84082, 84083, 84084, 84085, 84086, 84087, 84094, 84307, 101105, 103050, 126500, 141406

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157719)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157720)

    End Select

    Exit Function

End Function
'
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
'    If lErro <> SUCESSO Then gError 84088 '26489
'
'   For Each objCod_DescReduzida In colCod_DescReduzida
'
'        'Adiciona novo item na List da Combo CondicaoPagamento
'        CondicaoPagamento.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
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
'        Case 84088
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157721)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Call CF2(Me, "OV_Form_Unload", Cancel)
    
    Set gobjOrcamentoVenda = Nothing
    Set gcolItensOVPrecoCalc = Nothing
    Set gcolPCSituacoes = Nothing
    Set gcolFPSituacoes = Nothing
    Set gobjTelaProjetoInfo = Nothing
    Set gobjInfoAdic = Nothing
    Set gobjIdiomaTela = Nothing

    Set objGridItens = Nothing
    Set objGridParcelas = Nothing
    Set objGridPrecosCalculados = Nothing
    Set objGridFormacaoPreco = Nothing

    'trecho incluido por Leo em 20/03/02 - Filtrar
    Set objEventoCliente = Nothing
    Set objEventoNumero = Nothing
    Set objEventoCondPagto = Nothing
    Set objEventoProduto = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoVendedor2 = Nothing
    Set objEventoNaturezaOp = Nothing
    Set objEventoTiposDeTributacao = Nothing
    
    Set objEventoTransportadora = Nothing
    Set objEventoTranspRedespacho = Nothing
    Set objEventoMensagem = Nothing
    Set objEventoEstoqueProd = Nothing

    Set objEventoNumeroBase = Nothing
    Set objEventoVersao = Nothing

    Set objEventoBotaoPedAtualizados = Nothing
    
    Set gobjAnotacao = Nothing 'Inserido por Wagner
    Set gobjInfoUsu = Nothing 'Inserido por Wagner
    Set gobjTribTab = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Orçamento"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "OrcamentoVenda"

End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Function Carrega_TabelaPreco() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TabelaPreco

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePrecoVenda", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao, "Ativo = 1")
    If lErro <> SUCESSO Then gError 84012

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

        Case 84012

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157722)

    End Select

    Exit Function

End Function

Function Carrega_FilialEmpresa() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_FilialEmpresa

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "FiliaisEmpresa", "FilialEmpresa", "Nome", STRING_FILIAISEMPRESA_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 84093

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Preços
        Filial.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        Filial.ItemData(Filial.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr

    Select Case gErr

        Case 84093
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_FILIALEMPRESA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157723)

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
    If lErro <> SUCESSO Then gError 84070 '26715

    Call DataReferencia_Validate(bCancel)

    If bCancel = True Then DataReferencia.Text = sData

    Exit Sub

Erro_BotaoDataReferenciaDown_Click:

    Select Case gErr

        Case 84070

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157724)

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
    If lErro <> SUCESSO Then gError 84072

    Call DataReferencia_Validate(bCancel)

    If bCancel = True Then DataReferencia.Text = sData

    Exit Sub

Erro_BotaoDataReferenciaUp_Click:

    Select Case gErr

        Case 84072

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157725)

    End Select

    Exit Sub


End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o código foi informado
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 84307 '46130

    objOrcamentoVenda.lCodigo = StrParaLong(Codigo.Text)
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
    objOrcamentoVenda.iVersao = StrParaInt(OVVersao.Caption)
    Set objOrcamentoVenda.objTela = Me

    'Lê o Orcamento
    lErro = CF("OrcamentoVenda_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101232 Then gError 84308
    If lErro = 101232 Then gError 101274
    
    If objOrcamentoVenda.iHistorico = MARCADO Then gError 206830
    
    'por Leo em 29/04/02 **********
    'Se o orçamento estiver vinculado a um Pedido de Venda, não poderá ser excluido -> Erro.
    If objOrcamentoVenda.lNumIntPedVenda <> 0 Then
        
        gError 94486
    
    'Se o orçamento estiver vinculado a uma Nota Fiscal, não poderá ser excluido -> Erro.
    ElseIf objOrcamentoVenda.lNumIntNFiscal <> 0 Then
        
        gError 94487
    
    End If
    'leo *********
    
    'Pede a confirmação da exclusão do Orcamento de Venda
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_ORCAMENTO_VENDA", objOrcamentoVenda.lCodigo)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Faz a exclusão do Orcamento de Venda
    lErro = CF("OrcamentoVenda_Exclui", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84310 '46139

    'Limpa a Tela de Orcamento de Venda
    Call Limpa_OrcamentoVenda
    
    'fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 84307
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", gErr)

        Case 84308, 84310

        Case 101274
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO", gErr, objOrcamentoVenda.lCodigo)

        Case 94486
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORCAMENTO_VINCULADO_PEDIDO", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
            
        Case 94487
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ORCAMENTO_VINCULADO_NFISCAL", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)
              
        Case 206830
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_OV_HISTORICO", gErr)
              
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157726)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub


Public Sub BotaoGravar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 84243 '26806

    'Limpa a Tela
    Call Limpa_OrcamentoVenda

    iCodigoAlterado = 0
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 84243

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 157727)

    End Select

    Exit Sub

End Sub

Public Sub BotaoVersaoKitBase_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoVersaoKitBase_Click

    If Not (GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes) Then Exit Sub

    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 129966

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        objKit.sProdutoRaiz = sProdutoFormatado
            
        colSelecao.Add sProdutoFormatado
        
        Call Chama_Tela("KitVersaoLista", colSelecao, objKit, objEventoVersao)
    
    Else
         gError 129967
         
    End If
    
    Exit Sub

Erro_BotaoVersaoKitBase_Click:

    Select Case gErr

        Case 129966
        
        Case 129967
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTORAIZKIT_NAO_PREENCHIDO2", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157728)

    End Select

    Exit Sub


End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 84065 '26805

    'Limpa a Tela
    Call Limpa_OrcamentoVenda
    iCodigoAlterado = 0
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 84065

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157729)

    End Select

    Exit Sub


End Sub

Private Sub Limpa_OrcamentoVenda()

Dim lErro As Long
Dim sNatureza As String

On Error GoTo Erro_OrcamentoVenda
    
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    sNatureza = gobjTribTab.sNatOpInterna
    
    Call Limpa_OrcamentoVenda2

    gobjTribTab.sNatOpInterna = sNatureza

    Set gobjAnotacao = Nothing
    
    'Coloca clientes em atraso em vermelho
    If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)
        
    iCodigoAlterado = 0
    iAlterado = 0
          
    iVendedorAlterado = 0
    iVendedor2Alterado = 0
     
    Exit Sub

Erro_OrcamentoVenda:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157730)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_OrcamentoVenda2()
'Limpa os campos da tela sem fechar o sistema de setas

Dim iIndice As Integer

    Call Limpa_Tela(Me)
    
    DataUltAlt.Caption = ""
    HoraUltAlt.Caption = ""
    Usuario.Caption = ""
    OVVersao.Caption = ""
       
    If gobjCRFAT.iOVGravacaoVersao = OV_VERSAO_NAO_GRAVA Then
        TrocarVersao.Value = vbUnchecked
    Else
        TrocarVersao.Value = vbChecked
    End If

    Codigo.Enabled = True
    Codigo.Text = ""
    Filial.Clear
    FilialEntrega.Clear
    
    OptDataPrazoEnt(0).Value = True
    Call OptDataPrazoEnt_Click(0)
    
    PrazoTexto.Text = ""

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me

    Set gcolItensOVPrecoCalc = New Collection
    Set gobjInfoAdic = Nothing
    
    CobrancaAutomatica.Value = vbChecked
    
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
    Transportadora.Text = ""
    TranspRedespacho.Text = ""
    RedespachoCli.Value = False
    CanalVenda.Text = ""
    CondicaoPagamento.Text = ""
    PlacaUF.Text = ""
    DataReferencia.PromptInclude = False
    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataReferencia.PromptInclude = True
    'Trecho Incluido por Leo em 22/03/02
    Vendedor.PromptInclude = False
    Vendedor.Text = ""
    Vendedor.PromptInclude = True
    
    Vendedor2.PromptInclude = False
    Vendedor2.Text = ""
    Vendedor2.PromptInclude = True
    
    TipoFrete.ListIndex = 0
    
    If gobjFAT.iExibeOPCalcST = MARCADO Then
        If gobjFAT.iCalcSTPadrao = MARCADO Then
            CalcularST.Value = vbChecked
        Else
            CalcularST.Value = vbUnchecked
        End If
    Else
        CalcularST.Value = vbChecked
    End If
    
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
    Call Grid_Limpa(objGridParcelas)
    Call Grid_Limpa(objGridPrecosCalculados)
    Call Grid_Limpa(objGridFormacaoPreco)
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    Call DataEmissao_Validate(bSGECancelDummy)
    
    ImprimirOVComPreco.Value = vbChecked
    ImprimirOVComCodProd.Value = vbUnchecked
    
    lMotivoPerdaAnt = 0
    Call Combo_Seleciona_ItemData(Moeda, MOEDA_REAL)
    iMoedaAnt = Codigo_Extrai(Moeda.Text)
    
    Call gobjIdiomaTela.Reiniciar
    
    iPrazoAnt = 0
    
    iCodigoAlterado = 0
    iAlterado = 0
    giValorDescontoAlterado = 0
    giClienteAlterado = 0
    giFilialAlterada = 0
    giFilialEntregaAlterada = 0
    giDataReferenciaAlterada = 0
    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0
    giDataEntregaAlterada = 0

    giValorDescontoManual = 0
    giFilialEntregaManual = 0
    giPercAcresFinAlterado = 0

    Status.ListIndex = -1
    StatusComercial.ListIndex = -1
    MotivoPerda.ListIndex = -1
    
    Call gobjTribTab.Limpa_Tela
    
    If CalcularST.Value = vbChecked Then
        gobjOrcamentoVenda.iNaoCalcICMSST = DESMARCADO
    Else
        gobjOrcamentoVenda.iNaoCalcICMSST = MARCADO
    End If
    
    Call CF2(Me, "OV_Limpa_OrcamentoVenda2")
    
    Exit Sub

End Sub

Sub Seleciona_FilialEmpresa()

Dim lErro As Long
Dim iIndice As Integer
Dim iFilialFaturamento As Integer

On Error GoTo Erro_Seleciona_FilialEmpresa

    iFilialFaturamento = gobjFAT.iFilialFaturamento

    Exit Sub

Erro_Seleciona_FilialEmpresa:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157731)

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

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is Produto Then

        sProduto1 = Produto.Text

    Else

        'Verifica se tem alguma linha selecionada no Grid
        If GridItens.Row = 0 Then gError 84074 '58771

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    End If

    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 84075 '58772

    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto

    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 84074 '58771
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 84075 '58772 Tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157732)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lOrcamentoVenda As Long

On Error GoTo Erro_PedidoDeVenda_Automatico

    lErro = CF("Config_ObterAutomatico", "FatConfig", "NUM_PROX_CODIGO_ORCAMENTOVENDA", "OrcamentoVenda", "Codigo", lOrcamentoVenda)
    If lErro <> SUCESSO Then gError 94422
    
    Codigo.Text = lOrcamentoVenda
    Call Codigo_Validate(bSGECancelDummy)

    Exit Sub

Erro_PedidoDeVenda_Automatico:

    Select Case gErr

        Case 94422

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157733)

    End Select

    Exit Sub
    
End Sub

Public Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
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

On Error GoTo Erro_Cliente_Validate

    If giClienteAlterado = 1 Then

        'Verifica se o Cliente está preenchido
        If Len(Trim(Cliente.Text)) > 0 Then

            'Busca o Cliente no BD
            lErro = TP_Cliente_Le_Orcamento(Cliente, objCliente, iCodFilial)
            If lErro <> SUCESSO And lErro <> 94449 And lErro <> 94450 And lErro <> 94451 And lErro <> 94452 Then gError 84037
                                        
            If lErro = SUCESSO Then
            
                'gobjOrcamentoVenda.lCliente = objCliente.lCodigo
            
                lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
                If lErro <> SUCESSO Then gError 84038 '26136
    
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
                    If lErro <> SUCESSO And lErro <> 19062 Then gError 126956
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
                    If lErro <> SUCESSO Then gError 84039
    
                    'Incluído por Luiz Nogueira em 26/01/04
                    'Coloca na tela a tabela do cliente
                    If objCliente.iVendedor > 0 Then
                        
                        Vendedor.Text = objCliente.iVendedor
                        Call Vendedor_Validate(bSGECancelDummy)
                    
                    ElseIf objTipoCliente.iVendedor > 0 Then
                        
                        Vendedor.Text = objTipoCliente.iVendedor
                        Call Vendedor_Validate(bSGECancelDummy)
                    
                    End If
                    'Fim Luiz Nogueira - 26/01/04
    
                    'Coloca na tela a tabela do cliente
                    If objCliente.iTabelaPreco > 0 Then
                        
                        TabelaPreco.Text = objCliente.iTabelaPreco
                        Call TabelaPreco_Validate(bSGECancelDummy)
                    
                    ElseIf objTipoCliente.iTabelaPreco > 0 Then
                        
                        TabelaPreco.Text = objTipoCliente.iTabelaPreco
                        Call TabelaPreco_Validate(bSGECancelDummy)
                    
                    End If
                    
                    'Se cobrança automática estiver selecionada preenche a CondPagto e dispara o Validate
                    If CobrancaAutomatica.Value = MARCADO Then
                        
                        If objCliente.iCondicaoPagto > 0 Then
                            
                            CondicaoPagamento.Text = objCliente.iCondicaoPagto
                            Call CondicaoPagamento_Validate(bSGECancelDummy)
                        
                        ElseIf objTipoCliente.iCondicaoPagto > 0 Then
                            
                            CondicaoPagamento.Text = objTipoCliente.iCondicaoPagto
                            Call CondicaoPagamento_Validate(bSGECancelDummy)
                        
                        End If
    
                    End If
    
                End If
                
                giClienteAlterado = 0
    
            Else
            
                gobjOrcamentoVenda.lCliente = 0
                giValorDescontoManual = 0
                gdDesconto = 0
                                
                If Not gbCarregandoTela Then
                
                    Call DescontoGlobal_Recalcula
    
                    'ATualiza o total com o novo desconto
                    lErro = ValorTotal_Calcula()
                    If lErro <> SUCESSO Then gError 84039 '51034
                          
                    TabelaPreco.ListIndex = -1
                    
                    objCliente.lCodigo = 0
                    
                    Filial.Clear
                            
                End If
                
            End If
        
        'Se não estiver preenchido
        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            'Limpa a Combo de Filiais
            Filial.Clear
            
            FilialEntrega.Clear

        End If
        
        PercDescontoItens.Text = Format(gdDesconto * 100, "FIXED")
        Call PercDescontoItens_Validate(bSGECancelDummy)

        giClienteAlterado = 0
    
        If Not gbCarregandoTela Then

    ''*** incluidos p/tratamento de tributacao *******************************
            If iCodFilial <> 0 Then Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, iCodFilial) '####
    '*** fim tributacao
        End If
        
        'Coloca clientes em atraso em vermelho
        If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)

        Call CF2(Me, "OV_Cliente_Validate", objCliente)
        
    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 84037, 84038, 84039, 84040, 84041, 126956

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157734)

    End Select

    Exit Sub

End Sub

Private Sub DescontoGlobal_Recalcula()

'Dim dValorDesconto As Double
'Dim dValorProdutos As Double
'
'    If Len(Trim(ValorProdutos.Caption)) <> 0 And IsNumeric(ValorProdutos.Caption) Then
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
'            'Para tributação
'            gobjOrcamentoVenda.dValorDesconto = dValorDesconto
'
'        End If
'
'    End If

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

Public Sub CobrancaAutomatica_Click()

  iAlterado = REGISTRO_ALTERADO
  
  Call Cobranca_Automatica

End Sub

Public Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    iCodigoAlterado = REGISTRO_ALTERADO

    Call Atualiza_VersaoKit_Grid

End Sub

Public Sub Codigo_GotFocus()

    gsCodigoAnterior = Codigo.Text
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub CodigoBase_Validate(Cancel As Boolean)

Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim lErro As Long

On Error GoTo Erro_CodigoBase_Validate

    If StrParaLong(CodigoBase.Text) <> 0 Then

        objOrcamentoVenda.lCodigo = StrParaLong(CodigoBase.Text)
        objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
        
        'Lê os dados do Orcamento de Venda
        lErro = CF("OrcamentoVenda_Le", objOrcamentoVenda)
        If lErro <> SUCESSO And lErro <> 101232 Then gError 131274
        If lErro = 101232 Then gError 131275

        lErro = Traz_OrcamentoVenda_Tela(objOrcamentoVenda, VALOR_CODIGOBASE)
        If lErro <> SUCESSO Then gError 131276
    
    End If
    
    Exit Sub

Erro_CodigoBase_Validate:

    Cancel = True

    Select Case gErr
    
        Case 131274, 131276
        
        Case 131275
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157735)

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

    'Lê Condição a partir do código
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 84043 '26718
    If lErro = 19205 Then gError 84044 '26720

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
            If lErro <> SUCESSO Then gError 84045 '26719

        End If
    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_CondicaoPagamento_Click:

    Select Case gErr

        Case 84045, 84043

        Case 84044
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157736)

      End Select

    Exit Sub

End Sub

Private Function Cobranca_Automatica() As Long
'recalcula o tab de cobranca

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Cobranca_Automatica

    If CobrancaAutomatica.Value = vbChecked And Len(Trim(CondicaoPagamento.Text)) <> 0 Then

        objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)
        
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 84046
        If lErro <> SUCESSO Then gError 84048

        lErro = GridParcelas_Preenche(objCondicaoPagto)
        If lErro <> SUCESSO Then gError 84047

    End If

    Cobranca_Automatica = SUCESSO

    Exit Function

Erro_Cobranca_Automatica:

    Cobranca_Automatica = gErr

    Select Case gErr

        Case 84046, 84047

        Case 84048
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 157737)

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
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim iColuna As Integer
Dim dValorPagarPrimParc As Double

Dim colDescontoPadrao As New Collection

On Error GoTo Erro_GridParcelas_Preenche

    'Limpa o GridParcelas
    Call Grid_Limpa(objGridParcelas)

    'Número de Parcelas
    objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas

    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then gError 84076

    'Se Valor a Pagar for positivo
    If dValorPagar > 0 Then

        dValorPagarPrimParc = 0
        
        If objCondicaoPagto.iNumeroParcelas <> 1 Then
        
            lErro = gobjTribTab.ValorPagarPrimParc_Calcula(dValorPagarPrimParc)
            If lErro <> SUCESSO Then gError 84076 '26721
    
        End If
    
        If dValorPagarPrimParc <> 0 Then
            objCondicaoPagto.dValorTotal = Arredonda_Moeda(dValorPagar - dValorPagarPrimParc)
        Else
            objCondicaoPagto.dValorTotal = dValorPagar
        End If
                
        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, False, True, True)
        If lErro <> SUCESSO Then gError 84076 '26721

        If objGridParcelas.iLinhasExistentes <> objCondicaoPagto.colParcelas.Count Then objGridParcelas.iLinhasExistentes = objCondicaoPagto.colParcelas.Count

        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
            If iIndice = 1 And dValorPagarPrimParc <> 0 Then objCondicaoPagto.colParcelas(iIndice).dValor = objCondicaoPagto.colParcelas(iIndice).dValor + dValorPagarPrimParc
            
            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        
        Next

    End If

    'Se Data Referencia estiver preenchida
    If Len(Trim(DataReferencia.ClipText)) > 0 Then

        objCondicaoPagto.dtDataRef = CDate(DataReferencia.Text)

        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True, True, True)
        If lErro <> SUCESSO Then gError 84077

        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.colParcelas.Count  'objCondicaoPagto.iNumeroParcelas

            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dtVencimento, "dd/mm/yyyy")

        Next

    End If

    ' Se dValorPagar>0 coloca desconto padrao (quantos houver, se houver) em todas as parcelas.
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        lErro = Preenche_DescontoPadrao(iIndice)
        If lErro <> SUCESSO Then gError 84078
    Next

    GridParcelas_Preenche = SUCESSO

    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = gErr

    Select Case gErr

        Case 84076 To 84078

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157738)

    End Select

    Exit Function

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
        If lErro <> SUCESSO Then gError 84098

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

        Case 84098

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157739)

    End Select

    Exit Function

End Function

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 84049 '26542

    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 84050 '26543
        If lErro = 19205 Then gError 84054 '26545

        'Testa se pode ser usada em Contas a Receber
        If objCondicaoPagto.iEmRecebimento = 0 Then gError 84051 '26547

        'Coloca na Tela
        CondicaoPagamento.Text = CondPagto_Traz(objCondicaoPagto)

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
                If lErro <> SUCESSO Then gError 84052 '26544

            End If
        End If

    End If

    'Não encontrou o valor que era STRING
    If lErro = 6731 Then gError 84053 '26546

    Exit Sub

Erro_Condicaopagamento_Validate:

    Cancel = True

    Select Case gErr

       Case 84049, 84050, 84052

        Case 84051
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", gErr, objCondicaoPagto.iCodigo)

        Case 84053
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)

       Case 84054
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157740)

    End Select

    Exit Sub

End Sub

Public Sub CondPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection

    'Se Condição de Pagto estiver preenchida, extrai o código
    If Len(Trim(CondicaoPagamento.Text)) > 0 Then
        objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)
    End If

    'Chama a Tela CondicoesPagamentoCRLista
    Call Chama_Tela("CondicaoPagtoCRLista", colSelecao, objCondicaoPagto, objEventoCondPagto)

End Sub

Public Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissao_GotFocus()

     Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 84037

'    If gobjOrcamentoVenda.dtDataEmissao <> StrParaDate(DataEmissao.Text) Then
'
'        gobjOrcamentoVenda.dtDataEmissao = StrParaDate(DataEmissao.Text)
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
        Case 84037

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157741)

    End Select

    Exit Sub

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
        If lErro <> SUCESSO Then gError 84062 '26713

        'Compara com data de emissão
        If Len(Trim(DataEmissao.ClipText)) > 0 Then

            dtDataEmissao = CDate(DataEmissao.Text)
            dtDataReferencia = CDate(DataReferencia.Text)

            If dtDataEmissao > dtDataReferencia Then gError 84063

        End If


    End If

    giDataReferenciaAlterada = 0

    'Preenche o GridParcelas
    lErro = Cobranca_Automatica()
    If lErro <> SUCESSO Then gError 84064

    Exit Sub

Erro_DataReferencia_Validate:

    Cancel = True

    Select Case gErr

        Case 84062, 84064 'Tratado na rotina chamada

        Case 84063
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAREFERENCIA", gErr, dtDataReferencia, dtDataEmissao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157742)

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

    iAlterado = REGISTRO_ALTERADO

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

Public Sub Desconto3Percentual_Change()

    iAlterado = REGISTRO_ALTERADO

 End Sub

Private Sub Desconto3Percentual_GotFocus()

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

Public Sub DescricaoProduto_Change()

    iAlterado = REGISTRO_ALTERADO

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

Public Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO
    giFilialAlterada = 1

End Sub

Public Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    iAlterado = REGISTRO_ALTERADO

    'Se nenhuma filial foi selecionada, sai.
    If Filial.ListIndex = -1 Then Exit Sub

    'Faz o tratamento para a filial do cliente selecionada
    lErro = Trata_FilialCliente()
    If lErro <> SUCESSO Then gError 84042 '23581

    Exit Sub

Erro_Filial_Click:

    Select Case gErr

        Case 84042

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157743)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 84000 '26622

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 84001 '26623

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 84002 '26624

        If lErro = 17660 Then

            'Lê o Cliente
            objCliente.sNomeReduzido = sCliente
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 84003 '25664

            'Se encontrou o Cliente
            If lErro = SUCESSO Then
                
                objFilialCliente.lCodCliente = objCliente.lCodigo

                gError 84005
            
            End If
            
        End If
        
        If iCodigo <> 0 Then
        
            'Coloca na tela a Filial lida
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
        
            lErro = Trata_FilialCliente
            If lErro <> SUCESSO Then gError 84006 '25435
        
        Else
            
            objCliente.lCodigo = 0
            objFilialCliente.iCodFilial = 0
            
        End If
        
    'Não encontrou a STRING
    ElseIf lErro = 6731 Then
        
        'trecho incluido por Leo em 17/04/02
        objCliente.sNomeReduzido = Cliente.Text
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 94448 '25664
        
        If lErro = SUCESSO Then gError 84007
        
    End If
    
    giFilialAlterada = 0

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 84000, 84002

        Case 84001
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 84003, 84006, 94448 'tratado na rotina chamada

        Case 84005
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 84007
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157744)

    End Select

    Exit Sub

End Sub

'Function Trata_FilialCliente() As Long
'
'Dim objFilialCliente As New ClassFilialCliente
'Dim objCliente As New ClassCliente
'Dim objVendedor As New ClassVendedor
'Dim objTipoCliente As New ClassTipoCliente
'Dim dValorTotal As Double
'Dim dValorBase As Double
'Dim objTransportadora As New ClassTransportadora
'Dim dValorComissao As Double
'Dim dValorEmissao As Double
'Dim lErro As Long
'
'On Error GoTo Erro_Trata_FilialCliente
'
'    objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
'    objCliente.sNomeReduzido = Trim(Cliente.Text)
'
'    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Trim(Cliente.Text), objFilialCliente)
'    If lErro <> SUCESSO And lErro <> 17660 Then gError 84019
'    If lErro = 17660 Then gError 84020
'
'    Call gobjTribTab.FilialCliente_Alterada(objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
'
'    gobjOrcamentoVenda.iFilial = objFilialCliente.iCodFilial
'
'    'Calula o valor total
'    lErro = ValorTotal_Calcula()
'    If lErro <> SUCESSO Then gError 124223
'
'    Trata_FilialCliente = SUCESSO
'
'    Exit Function
'
'Erro_Trata_FilialCliente:
'
'    Trata_FilialCliente = gErr
'
'    Select Case gErr
'
'        Case 84019, 124223
'
'        Case 84020
'            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA1", gErr, Cliente.Text, objFilialCliente.iCodFilial)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157745)
'
'    End Select
'
'    Exit Function
'
'End Function

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
    If lErro <> SUCESSO And lErro <> 17660 Then gError ERRO_SEM_MENSAGEM
    If lErro = 17660 Then gError 211527

    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
    If lErro = 12348 Then gError 211528
    
    Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, objFilialCliente.iCodFilial)
    
    'Transportadora
    If Len(Trim(Transportadora.Text)) = 0 Then

        'Testa se FilialCliente tem Transportadora
        If objFilialCliente.iCodTransportadora <> 0 Then

            objTransportadora.iCodigo = objFilialCliente.iCodTransportadora
            lErro = CF("Transportadora_Le", objTransportadora)
            If lErro <> SUCESSO And lErro <> 19250 Then gError ERRO_SEM_MENSAGEM
            If lErro = SUCESSO Then
                Transportadora.Text = objTransportadora.sNomeReduzido
                Call Transportadora_Validate(False)
            End If

        ElseIf objCliente.iTipo <> 0 Then

            'Testa se TipoCliente já foi lido no BD
            If objTipoCliente.iCodigo = 0 Then

                objTipoCliente.iCodigo = objCliente.iTipo
                lErro = CF("TipoCliente_Le", objTipoCliente)
                If lErro <> SUCESSO And lErro <> 19062 Then gError ERRO_SEM_MENSAGEM
                If lErro <> SUCESSO Then objTipoCliente.iCodTransportadora = 0

            End If

            'Testa se TipoCliente tem Transportadora
            If objTipoCliente.iCodTransportadora <> 0 Then
                lErro = CF("Transportadora_Le", objTransportadora)
                If lErro <> SUCESSO And lErro <> 19250 Then gError ERRO_SEM_MENSAGEM
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
            If lErro <> SUCESSO And lErro <> 19250 Then gError ERRO_SEM_MENSAGEM
            If lErro <> SUCESSO Then gError 211529
            
            'Coloca a transportadora na tela
            TranspRedespacho.Text = objTransportadora.sNomeReduzido
            Call TranspRedespacho_Validate(bSGECancelDummy)
        
            RedespachoCli.Value = objFilialCliente.iDetPagFrete
        
        End If
    
    End If
    
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If objFilialCliente.iCodMensagem > 0 Then
        objMensagem.iCodigo = objFilialCliente.iCodMensagem
        'Lê a Mensagem "default" p/ o Cliente
        lErro = CF("Mensagem_Le", objMensagem)
        If lErro <> SUCESSO And lErro <> 19234 Then gError ERRO_SEM_MENSAGEM

        If lErro = SUCESSO Then Mensagem.Text = objMensagem.sDescricao

    End If

    If gobjCRFAT.iPrecoUnitarioFormPreco <> 0 Then Call Trata_TabelaPreco
    
    Trata_FilialCliente = SUCESSO

    Exit Function

Erro_Trata_FilialCliente:

    Trata_FilialCliente = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 211527
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA1", gErr, Cliente.Text, objFilialCliente.iCodFilial)

        Case 211528
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case 211529
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr, objTransportadora.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211530)

    End Select

    Exit Function

End Function

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

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer
Dim iIndice As Integer
Dim dValorTotal As Double, dValorTotalB As Double
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_GridItens_KeyDown

    'Guarda o número de linhas existentes e a linha atual
    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iItemAtual = GridItens.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then

        '************ grade ************
        'Retira a "#" caso o item excluído tenha sido um de grade
        GridItens.TextMatrix(GridItens.Row, 0) = GridItens.Row
        '*******************************

        Call gobjTelaProjetoInfo.Remove_Linha(iItemAtual)
              
        Call gobjTribTab.Exclusao_Item_Grid(iItemAtual)
        
        Call FormPreco_Remove_Linha(iItemAtual)
        
        'Calcula a soma dos valores de produtos
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
                If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
            End If
            dValorTotalB = dValorTotalB + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
        Next
        
        If objGridItens.iLinhasExistentes <> 0 Then
            Call PrecoTotal_Calcula(objGridItens.iLinhasExistentes)
        Else
            If StrParaDbl(ValorDesconto.Text) <> 0 Then
                'Avisa ao usuário
                vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", ValorDesconto.Text, 0)
                
                'Limpa o valor de desconto
                gdDesconto = 0
                ValorDesconto.Text = ""
                giValorDescontoAlterado = 0
                
                Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, 0)
               
                'Para tributação
                gobjOrcamentoVenda.dValorDesconto = 0

            End If
        End If
        
        For iIndice = 1 To gobjOrcamentoVenda.colItens.Count
            If gobjOrcamentoVenda.colItens(iIndice).iPossuiGrade = MARCADO Then
                GridItens.TextMatrix(iIndice, 0) = "# " & CStr(iIndice)
            Else
                GridItens.TextMatrix(iIndice, 0) = CStr(iIndice)
            End If
            gobjOrcamentoVenda.colItens(iIndice).iItem = iIndice 'Incluído para ao alterar a grade ou a reserva não fazer do item errado
        Next
        GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, 0) = CStr(objGridItens.iLinhasExistentes + 1)
        
        'Coloca valor total dos produtos na tela
        ValorProdutos.Caption = Format(dValorTotal - StrParaDbl(ValorDesconto.Text), "Standard")
        ValorProdutos2.Caption = Format(dValorTotalB, "Standard")

        'Calcula o valor total da nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 84145

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 84145

        lErro = FormPreco_Calcula()
        If lErro <> SUCESSO Then gError 84145
    
    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr

        Case 84145

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157746)

    End Select

    Exit Sub

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

Public Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGridItens)

End Sub

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Public Sub GridItens_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridItens)

End Sub

Public Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Public Sub GridParcelas_EnterCell()

     Call Grid_Entrada_Celula(objGridParcelas, iAlterado)

End Sub

Public Sub GridParcelas_GotFocus()

    Call Grid_Recebe_Foco(objGridParcelas)

End Sub

Public Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)

End Sub

Public Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Public Sub GridParcelas_LeaveCell()

    Call Saida_Celula(objGridParcelas)

End Sub

Public Sub GridParcelas_RowColChange()

    Call Grid_RowColChange(objGridParcelas)

End Sub

Public Sub GridParcelas_Scroll()

    Call Grid_Scroll(objGridParcelas)

End Sub

Public Sub GridParcelas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridParcelas)

End Sub

Public Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)


End Sub

Public Sub NumeroLabel_Click()

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim colSelecao As Collection

On Error GoTo Erro_NumeroLabel_Click
    
    lErro = Move_OrcamentoVenda_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84106

    Call Chama_Tela("OrcamentoVendaCGLista", colSelecao, objOrcamentoVenda, objEventoNumero)
    
    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 84106

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157762)

    End Select

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

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOrcamentoVenda As ClassOrcamentoVenda

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objOrcamentoVenda = obj1

    lErro = Traz_OrcamentoVenda_Tela(objOrcamentoVenda, VALOR_CODIGO)
    If lErro <> SUCESSO And lErro <> 84363 Then gError 84479

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 84479
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157763)

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
        If lErro <> SUCESSO Then gError 94414
    
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
    
        If Not (Me.ActiveControl Is Produto) Then
    
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
        
            'Faz o Tratamento do produto
            lErro = Produto_Saida_Celula()
            If lErro <> SUCESSO Then
            
                If Not (Me.ActiveControl Is Produto) Then
                
                    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""
                
                End If
            
                gError 94415
            End If
            
        End If
        
    End If
    
    'Call ComandoSeta_Fechar(Me.Name)
    
    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr
            
        Case 94414
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 94415

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157764)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1

    'Preenche campo Vendedor
    Vendedor.Text = objVendedor.sNomeReduzido

    Me.Show

End Sub

Private Sub objEventoVendedor2_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1

    'Preenche campo Vendedor
    Vendedor2.Text = objVendedor.sNomeReduzido

    Me.Show

End Sub

Public Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'se abriu o tab de tributacao
        If Opcao.SelectedItem.Index = TAB_Tributacao Then
            
            lErro = gobjTribTab.TabClick
            If lErro <> SUCESSO Then gError 84308 '27840
            
        '??? Alteração Daniel em 29/10/2002
        ElseIf Opcao.SelectedItem.Index = TAB_Cobranca Then
        
            'Recalcula as parcelas
            Call CobrancaAutomatica_Click
        
        End If

        'Esconde o frame atual, mostra o novo
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(giFrameAtual).Visible = False

        'Armazena novo valor de giFrameAtual
        giFrameAtual = Opcao.SelectedItem.Index
       
    End If

    Exit Sub

Erro_Opcao_Click:

    Select Case gErr

        Case 84308

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157765)

    End Select

    Exit Sub

End Sub

Public Sub PercAcrescFin_Change()

    iAlterado = REGISTRO_ALTERADO
    giPercAcresFinAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercAcrescFin_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercAcrescFin_Validate

    If giPercAcresFinAlterado = 0 Then Exit Sub

    If Len(Trim(PercAcrescFin.Text)) > 0 Then
        lErro = Porcentagem_Critica_Negativa(PercAcrescFin)
        If lErro <> SUCESSO Then gError 84091 '26717
    End If

    If Len(Trim(TabelaPreco.Text)) > 0 Then

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 84092 '46190

    End If

    giPercAcresFinAlterado = 0

    Exit Sub

Erro_PercAcrescFin_Validate:

    Cancel = True


    Select Case gErr

        Case 84091, 84092

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157767)

    End Select

    Exit Sub

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

Public Sub PrazoValidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrecoTotal_Change()

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

Public Sub TabelaPreco_Click()

Dim lErro As Long

On Error GoTo Erro_TabelaPreco_Click

    iAlterado = REGISTRO_ALTERADO

    If TabelaPreco.ListIndex = -1 Then Exit Sub

    If objGridItens.iLinhasExistentes = 0 Then Exit Sub

    'Faz o tratamento para a Tabela de Preços escolhida
    lErro = Trata_TabelaPreco()
    If lErro <> SUCESSO Then gError 84013

    Exit Sub

Erro_TabelaPreco_Click:

    Select Case gErr

        Case 84013

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157768)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 84014

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 84015 '26539

        If lErro <> SUCESSO Then gError 84016 '26540 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 84017 '30527

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 84018 '26541

    Exit Sub

Erro_TabelaPreco_Validate:

    Cancel = True

    Select Case gErr

    Case 84014, 84015, 84017

    Case 84016  'Não encontrou Tabela de Preço no BD

        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")

        If vbMsgRes = vbYes Then
            'Preenche o objTabela com o Codigo
            If Len(Trim(TabelaPreco.Text)) > 0 Then objTabelaPreco.iCodigo = CInt(TabelaPreco.Text)
            'Chama a tela de Tabelas de Preço
            Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
        End If

    Case 84018
        Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157769)

    End Select

    Exit Sub

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

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 84066 '26583

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 84066 '26583

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157772)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 84067

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 84067

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157773)

    End Select

    Exit Sub

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

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Valor")
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
    iGrid_Desc1Codigo_Col = 3
    iGrid_Desc1Ate_Col = 4
    iGrid_Desc1Valor_Col = 5
    iGrid_Desc1Perc_Col = 6
    iGrid_Desc2Codigo_Col = 7
    iGrid_Desc2Ate_Col = 8
    iGrid_Desc2Valor_Col = 9
    iGrid_Desc2Perc_Col = 10
    iGrid_Desc3Codigo_Col = 11
    iGrid_Desc3Ate_Col = 12
    iGrid_Desc3Valor_Col = 13
    iGrid_Desc3Perc_Col = 14

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

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Inicializa_Grid_Itens = CF2(Me, "OV_Inicializa_Grid_Itens", objGridInt) 'Inserido por Wagner

End Function

Private Function Trata_TabelaPreco() As Long

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_Trata_TabelaPreco

    If Not gbCarregandoTela Then

        For iLinha = 1 To objGridItens.iLinhasExistentes

            lErro = Trata_TabelaPreco_Item(iLinha)
            If lErro <> SUCESSO Then gError 84019

        Next

        'Calcula o Valor Total da Nota
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 84020

    End If

    Trata_TabelaPreco = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco:

    Trata_TabelaPreco = gErr

    Select Case gErr

        Case 84019 'tratado na rotina chamada

        Case 84020

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157774)

    End Select

    Exit Function

End Function

Public Function Trata_TabelaPreco_Item(iLinha As Integer) As Long
'faz tratamento de tabela de preço para um ítem (produto)

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim sProduto As String
Dim iPreenchido As Integer, objCliente As New ClassCliente, lCliente As Long
Dim dPercDescontoTab As Double, dPrecoComDescontoTab As Double

On Error GoTo Erro_Trata_TabelaPreco_Item

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 84021 '39147

    If iPreenchido <> PRODUTO_VAZIO And Len(Trim(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col))) > 0 Then

        lErro = CF2(Me, "NFCF_PrecoUnitario_Calcula", giFilialEmpresa, sProduto, GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), Cliente.Text, Codigo_Extrai(Filial.Text), Codigo_Extrai(TabelaPreco.Text), StrParaDbl(PercAcrescFin.Text) / 100, Codigo_Extrai(Moeda.Text), Nothing, dPrecoUnitario, True, dPercDescontoTab, dPrecoComDescontoTab)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If dPercDescontoTab >= gdDesconto Then
            GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col) = Format(dPercDescontoTab, "Percent")
        End If
        
        lErro = CF2(Me, "OV_Move_PrecoUnitario_Tela", dPrecoUnitario, iLinha)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Calcula o Preco Total do Ítem
        Call PrecoTotal_Calcula(iLinha)
        
        ''*** incluidos p/tratamento de tributacao *******************************
        lErro = gobjTribTab.Alteracao_Item_Grid(iLinha)
        If lErro <> SUCESSO Then gError 84023
        '*** fim tributacao

    End If

    Trata_TabelaPreco_Item = SUCESSO

    Exit Function

Erro_Trata_TabelaPreco_Item:

    Trata_TabelaPreco_Item = gErr

    Select Case gErr

        Case 84021, 84022, 84023, 84024, ERRO_SEM_MENSAGEM 'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157775)

    End Select

    Exit Function

End Function

Function ValorTotal_Calcula() As Long
'Calcula o Valor Total do Pedido

Dim dValorTotal As Double
Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    lErro = gobjTribTab.ValorTotal_Calcula(dValorTotal)
    If lErro <> SUCESSO Then gError 101102
    
    ValorTotal.Caption = Format(dValorTotal, "Standard")
    
    Call ValorDescontoTit_Validate(bSGECancelDummy)
    
    Call ValorDescontoItens_Calcula

    Call CF2(Me, "OV_ValorTotal_Calcula")

    If Not gbCarregandoTela Then
    
        'Faz o cálculo automático das comissões
        lErro = Cobranca_Automatica()
        If lErro <> SUCESSO Then gError 101103
        
    End If

    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 101102, 101103

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157776)

    End Select

    Exit Function

End Function

Public Function PrecoUnitario_Calcula(sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long
    PrecoUnitario_Calcula = CF2(Me, "OV_PrecoUnitario_Calcula", sUM, objTabelaPrecoItem, dPrecoUnitario)
End Function

Public Function OV_PrecoUnitario_Calcula(ByVal objCT As Object, sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long
'Calcula o Preço unitário do item de acordo com a UM e a tabela de preço
'apenas para grade

Dim objProduto As New ClassProduto
Dim objUM As New ClassUnidadeDeMedida
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim lErro As Long
Dim dPercAcresFin As Double, objTabelaPreco As New ClassTabelaPreco
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objCotacao1 As New ClassCotacaoMoeda
Dim objCotacaoAnterior1 As New ClassCotacaoMoeda
Dim objCotacao2 As New ClassCotacaoMoeda
Dim objCotacaoAnterior2 As New ClassCotacaoMoeda
Dim dCotacao1 As Double, dCotacao2 As Double
Dim vbMsgResult As VbMsgBoxResult

On Error GoTo Erro_OV_PrecoUnitario_Calcula

    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto
    'Lê o produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 84029 '26638
    If lErro = 28030 Then gError 84030 '26639

    'Converte a quantidade para a UM de Venda
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 84031 '26640

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
        If lErro <> SUCESSO And lErro <> 28004 Then gError 84029
        If lErro = SUCESSO And (objTabelaPreco.iMoeda <> Codigo_Extrai(Moeda.Text) Or iMoedaAnt <> Codigo_Extrai(Moeda.Text)) Then
        'Se possuem moedas diferentes ou se trocou a moeda precisa calcular/recalcular a cotação
            If objTabelaPreco.iMoeda <> MOEDA_REAL Then

                objCotacao1.dtData = gdtDataAtual
                objCotacao1.iMoeda = objTabelaPreco.iMoeda
                objCotacaoAnterior1.iMoeda = objTabelaPreco.iMoeda

                'Chama função de leitura
                lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao1, objCotacaoAnterior1)
                If lErro <> SUCESSO Then gError 84029

            Else
                objCotacao1.dValor = 1
            End If

            If Codigo_Extrai(Moeda.Text) <> MOEDA_REAL Then

                objCotacao2.dtData = gdtDataAtual
                objCotacao2.iMoeda = Codigo_Extrai(Moeda.Text)
                objCotacaoAnterior2.iMoeda = Codigo_Extrai(Moeda.Text)

                'Chama função de leitura
                lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao2, objCotacaoAnterior2)
                If lErro <> SUCESSO Then gError 84029

            Else
                objCotacao2.dValor = 1
            End If

            If objCotacao1.dValor <> 0 Then
                dCotacao1 = StrParaDbl(Format(objCotacao1.dValor, "#.0000"))
            Else
                vbMsgResult = Rotina_Aviso(vbYesNo, "AVISO_MOEDA_SEM_COTACAO_DATA", objCotacao1.iMoeda, Format(gdtDataAtual, "dd/mm/yyyy"), Format(objCotacaoAnterior1.dValor, "#.0000"), Format(objCotacaoAnterior1.dtData, "dd/mm/yyyy"))
                If vbMsgResult = vbNo Then gError ERRO_SEM_MENSAGEM
                dCotacao1 = StrParaDbl(Format(objCotacaoAnterior1.dValor, "#.0000"))
            End If

            If objCotacao2.dValor <> 0 Then
                dCotacao2 = StrParaDbl(Format(objCotacao2.dValor, "#.0000"))
            Else
                vbMsgResult = Rotina_Aviso(vbYesNo, "AVISO_MOEDA_SEM_COTACAO_DATA", objCotacao2.iMoeda, Format(gdtDataAtual, "dd/mm/yyyy"), Format(objCotacaoAnterior2.dValor, "#.0000"), Format(objCotacaoAnterior2.dtData, "dd/mm/yyyy"))
                If vbMsgResult = vbNo Then gError ERRO_SEM_MENSAGEM
                dCotacao2 = StrParaDbl(Format(objCotacaoAnterior2.dValor, "#.0000"))
            End If

            If dCotacao1 = 0 Then gError 211631
            If dCotacao2 = 0 Then gError 211632

            'Se nao existe cotacao para a data informada => Mostra a última.
            dPrecoUnitario = dPrecoUnitario * dCotacao1 / dCotacao2

        End If

    End If
    
    OV_PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_OV_PrecoUnitario_Calcula:

    OV_PrecoUnitario_Calcula = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 84029, 84031

        Case 84030
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objTabelaPrecoItem.sCodProduto)

        Case 211631
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_SEM_COTACAO", gErr, objCotacao1.iMoeda)

        Case 211632
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_SEM_COTACAO", gErr, objCotacao2.iMoeda)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157777)

    End Select

    Exit Function

End Function

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
        If lErro <> SUCESSO Then gError 84055 '26652

        dValorDesconto = CDbl(ValorDesconto.Text)

        'Coloca o Valor formatado na tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")

        If dValorDesconto > dValorProdutos Then gError 84056 '26653

        dValorProdutos = dValorProdutos - dValorDesconto

    End If

    ValorProdutos.Caption = Format(dValorProdutos, "Standard")

    'Para tributação
    gobjOrcamentoVenda.dValorDesconto = dValorDesconto

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDesconto_Validate(Cancel, dValorDesconto)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 84057 '51038

    giValorDescontoAlterado = 0

    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True

    Select Case gErr

        Case 84055, 84057

        Case 84056
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157778)

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
        If lErro <> SUCESSO Then gError 84108 '35885

        dValorDespesas = CDbl(ValorDespesas.Text)

        'coloca o valor formatado na tela
        ValorDespesas.Text = Format(dValorDespesas, "Standard")

    End If

    'Para tributação
    gobjOrcamentoVenda.dValorOutrasDespesas = dValorDespesas

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDespesas_Validate(Cancel, dValorDespesas)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 84109 '51039

    giValorDespesasAlterado = 0

    Exit Sub

Erro_ValorDespesas_Validate:

    Cancel = True

    Select Case gErr

        Case 84108

        Case 84109

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157779)

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
        If lErro <> SUCESSO Then gError 84058 '26650

        dValorFrete = CDbl(ValorFrete.Text)

        ValorFrete.Text = Format(dValorFrete, "Standard")

    End If

    'Para tributação
    gobjOrcamentoVenda.dValorFrete = dValorFrete

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorFrete_Validate(Cancel, dValorFrete)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 84059 '51040

    giValorFreteAlterado = 0

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True

    Select Case gErr

        Case 84058, 84059

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157780)

    End Select

    Exit Sub

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
        If lErro <> SUCESSO Then gError 84060 '26651

        dValorSeguro = CDbl(ValorSeguro.Text)

        ValorSeguro.Text = Format(dValorSeguro, "Standard")

    End If

    'Para tributação
    gobjOrcamentoVenda.dValorSeguro = dValorSeguro

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorSeguro_Validate(Cancel, dValorSeguro)
'*** fim tributacao

    'Recalcula valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError 84061 '51041

    giValorSeguroAlterado = 0

    Exit Sub

Erro_Valorseguro_Validate:

    Cancel = True

    Select Case gErr

        Case 84060, 84061

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157782)

    End Select

    Exit Sub

End Sub

'mario
Private Function Move_GridItens_Memoria(objOrcamentoVenda As ClassOrcamentoVenda) As Long
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
            If lErro <> SUCESSO Then gError 84102
            
        End If
        
        lErro = Move_GridItem_Memoria(objOrcamentoVenda, iIndice, sProduto)
        If lErro <> SUCESSO Then gError 84102

        '********************* TRATAMENTO DE GRADE *****************
        Call Move_ItensGrade_Tela(objOrcamentoVenda.colItens(iIndice).colItensRomaneioGrade, gobjOrcamentoVenda.colItens(iIndice).colItensRomaneioGrade)

    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 84102

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157783)

    End Select

    Exit Function

End Function

Public Function Move_GridItem_Memoria(objOrcamentoVenda As ClassOrcamentoVenda, iIndice As Integer, sProduto1 As String) As Long
'Recolhe do Grid os dados do item orçamento no parametro

Dim lErro As Long
Dim sProduto As String
Dim objItemOrcamento As New ClassItemOV, objTributacaoOV As New ClassTributacaoDoc
Dim iPreenchido As Integer

On Error GoTo Erro_Move_GridItem_Memoria

    Set objItemOrcamento = New ClassItemOV

'    'Verifica se o Produto está preenchido
'    If Len(Trim(sProduto1)) > 0 Then
'
'        'Formata o produto
'        lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError 84103 '27682

    objItemOrcamento.sProduto = sProduto1
'    End If

    If Len(Trim(sProduto1)) = 0 Then gError 84104 'Or iPreenchido = PRODUTO_VAZIO Then gError 84104 '20767

    'Armazena os dados do item
    objItemOrcamento.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
    objItemOrcamento.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
    objItemOrcamento.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))
    objItemOrcamento.dPrecoTotal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
    objItemOrcamento.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
    objItemOrcamento.dtDataEntrega = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col))
    objItemOrcamento.iPrazoEntrega = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_PrazoEntrega_Col))
    'objItemOrcamento.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
    objItemOrcamento.sDescricao = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
    objItemOrcamento.iFilialEmpresa = giFilialEmpresa
    
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objItemOrcamento.sVersaoKit = GridItens.TextMatrix(iIndice, iGrid_VersaoKit_Col)
        objItemOrcamento.sVersaoKitBase = GridItens.TextMatrix(iIndice, iGrid_VersaoKitBase_Col)
    End If
    
    If gobjOrcamentoVenda.colItens.Count >= iIndice Then
        Set objItemOrcamento.objTributacaoItemOV = gobjOrcamentoVenda.colItens.Item(iIndice).objTributacaoItemOV
        Set objItemOrcamento.objInfoAdicDocItem = gobjOrcamentoVenda.colItens(iIndice).objInfoAdicDocItem
    Else
        Set objItemOrcamento.objTributacaoItemOV = Nothing
        Set objItemOrcamento.objInfoAdicDocItem = New ClassInfoAdicDocItem
    End If
    
    objItemOrcamento.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Observacao_Col)
    objItemOrcamento.lStatus = LCodigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_Status_Col))
    objItemOrcamento.lMotivoPerda = LCodigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_MotivoPerda_Col))
 
    lErro = CF2(Me, "OV_Move_GridItens_Memoria", objItemOrcamento, iIndice)
    If lErro <> SUCESSO Then gError 141405
    
    If iGrid_SomaItem_Col <> 0 Then
        objItemOrcamento.iMarcado = StrParaInt(GridItens.TextMatrix(iIndice, iGrid_SomaItem_Col))
    End If
    
    'Adiciona o item na colecao de itens do orçamento de venda
     objOrcamentoVenda.colItens.Add objItemOrcamento

    Move_GridItem_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItem_Memoria:

    Move_GridItem_Memoria = gErr

    Select Case gErr

        Case 84103, 141405

        Case 84104
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157784)

    End Select

    Exit Function

End Function

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
                If lErro <> SUCESSO Then gError 84132 '26064

            'Se for o GridItens
            Case GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 84133 '26065

            Case GridPrecosCalculados.Name

                lErro = Saida_Celula_GridPrecosCalculados(objGridInt)
                If lErro <> SUCESSO Then gError 84133

            Case GridFormacaoPreco.Name

                lErro = Saida_Celula_GridFormacaoPreco(objGridInt)
                If lErro <> SUCESSO Then gError 84133

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 84134 '26068

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 84132, 84133, 84134

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157785)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridParcelas(objGridInt As AdmGrid) As Long
'Faz a crítica do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridParcelas

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Vencimento_Col
            lErro = Saida_Celula_DataVencimento(objGridInt)
            If lErro <> SUCESSO Then gError 84126 '26784

        Case iGrid_ValorParcela_Col
            lErro = Saida_Celula_ValorParcela(objGridInt)
            If lErro <> SUCESSO Then gError 84127 '26785

        Case iGrid_Desc1Codigo_Col, iGrid_Desc2Codigo_Col, iGrid_Desc3Codigo_Col
            lErro = Saida_Celula_TipoDesconto(objGridInt)
            If lErro <> SUCESSO Then gError 84128 '26786

        Case iGrid_Desc1Ate_Col, iGrid_Desc2Ate_Col, iGrid_Desc3Ate_Col
            lErro = Saida_Celula_DescontoData(objGridInt)
            If lErro <> SUCESSO Then gError 84129 '26787

        Case iGrid_Desc1Valor_Col, iGrid_Desc2Valor_Col, iGrid_Desc3Valor_Col
            lErro = Saida_Celula_DescontoValor(objGridInt)
            If lErro <> SUCESSO Then gError 84130 '26830

        Case iGrid_Desc1Perc_Col, iGrid_Desc2Perc_Col, iGrid_Desc3Perc_Col
            lErro = Saida_Celula_DescontoPerc(objGridInt)
            If lErro <> SUCESSO Then gError 84131 '26788

    End Select

    Saida_Celula_GridParcelas = SUCESSO

    Exit Function

Erro_Saida_Celula_GridParcelas:

    Saida_Celula_GridParcelas = gErr

    Select Case gErr

        Case 84126, 84127, 84128, 84129, 84130, 84131

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157786)

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
            If lErro <> SUCESSO Then gError 84135 '26593
            
        Case iGrid_DescProduto_Col
            lErro = Saida_Celula_DescProduto(objGridInt)
            If lErro <> SUCESSO Then gError 129216

        'Se for a de Unidade de Medida
        Case iGrid_UnidadeMed_Col
            lErro = Saida_Celula_UM(objGridInt)
            If lErro <> SUCESSO Then gError 84136 '26594

        'Se for a de Quantidade Pedida
        Case iGrid_Quantidade_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 84137 '26595

        'Se for a de Preço Unitário
        Case iGrid_PrecoUnitario_Col
            lErro = Saida_Celula_PrecoUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 84139 '26596

        'Se for a de Percentual de Desconto
        Case iGrid_PercDesc_Col
            lErro = Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then gError 84140 '26599
            
        Case iGrid_Desconto_Col
            lErro = Saida_Celula_Desconto(objGridInt)
            If lErro <> SUCESSO Then gError 84140

        'Se for a de Data de Entrega
        Case iGrid_DataEntrega_Col
            lErro = Saida_Celula_DataEntrega(objGridInt)
            If lErro <> SUCESSO Then gError 84141 '26601

        Case iGrid_PrazoEntrega_Col
            lErro = Saida_Celula_PrazoEntrega(objGridInt)
            If lErro <> SUCESSO Then gError 84141 '26601

        Case iGrid_VersaoKit_Col
            If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
                lErro = Saida_Celula_VersaoKit(objGridInt)
                If lErro <> SUCESSO Then gError 129965
            End If

        Case iGrid_VersaoKitBase_Col
            If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
                lErro = Saida_Celula_VersaoKitBase(objGridInt)
                If lErro <> SUCESSO Then gError 129966
            End If
    
        Case iGrid_Observacao_Col
            lErro = Saida_Celula_Observacao(objGridInt)
            If lErro <> SUCESSO Then gError 141385
        
        Case iGrid_MotivoPerda_Col
            lErro = Saida_Celula_MotivoPerdaItem(objGridInt)
            If lErro <> SUCESSO Then gError 141389
        
        Case iGrid_Status_Col
            lErro = Saida_Celula_StatusItem(objGridInt)
            If lErro <> SUCESSO Then gError 141390
        
        Case iGrid_PrecoTotal_Col
            lErro = Saida_Celula_PrecoTotal(objGridInt)
            If lErro <> SUCESSO Then gError 141390
        
    End Select

    lErro = CF2(Me, "OV_Saida_Celula_GridItens", objGridInt)
    If lErro <> SUCESSO Then gError 141410
    
    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 84135, 84136, 84137, 84138, 84139, 84140, 84141, 129216, 129965, 129966, 141410, 141385, 141389, 141390

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157787)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim objOrcamentoVendaBD As New ClassOrcamentoVenda
Dim dValorTotal As Double
Dim dValor As Double
Dim vbResult As VbMsgBoxResult
Dim objAnaliseVendas As New ClassAnaliseVendaInfo

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 84308 '26807
    If Len(Trim(Cliente.Text)) = 0 Then gError 84309 '26808
    If Len(Trim(Filial.Text)) = 0 Then gError 84310 '26809
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 84311 '26810
    If Len(Trim(gobjTribTab.sNatOpInterna)) = 0 Then gError 94496
    If Len(Trim(FilialEntrega.Text)) = 0 Then gError 211526
    
    dValor = CDbl(ValorTotal.Caption)

    If dValor < 0 Then gError 84312 ''30594

    lErro = Valida_Grid_Itens()
    If lErro <> SUCESSO Then gError 84313 '26812

    lErro = Valida_Grid_Parcelas()
    If lErro <> SUCESSO Then gError 84314 '26816

    lErro = Move_OrcamentoVenda_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84315 '26829

    lErro = CF2(Me, "OV_Valida_Itens", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 141385
    
    lErro = gobjTribTab.Valida_TributacaoNF_Gravacao(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 56931

    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786

    lErro = gobjTribTab.Valida_Dados
    If lErro <> SUCESSO Then gError 181786

    lErro = CF2(Me, "OV_FormPreco_ValidaDados", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 181786
    
    objOrcamentoVendaBD.lCodigo = objOrcamentoVenda.lCodigo
    objOrcamentoVendaBD.iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa
    
    lErro = CF("OrcamentoVenda_Le", objOrcamentoVendaBD)
    If lErro <> SUCESSO And lErro <> 101232 Then gError 181786
    
    'Se for uma alteração confirma a decisão tomada
    If lErro = SUCESSO Then
        vbResult = vbYes
        If TrocarVersao.Value = vbChecked Then
            If gobjCRFAT.iOVGravacaoVersao = OV_VERSAO_PERGUNTA Then vbResult = Rotina_Aviso(vbYesNo, "AVISO_OV_TROCA_VERSAO")
        Else
            If objOrcamentoVendaBD.iVersao <> objOrcamentoVenda.iVersao Then gError 206849 'Não é a versão ativa
            If gobjCRFAT.iOVGravacaoVersao = OV_VERSAO_PERGUNTA Then vbResult = Rotina_Aviso(vbYesNo, "AVISO_OV_ALTERA_VERSAO")
        End If
        If vbResult = vbNo Then gError 141373
    End If
    
    If CalcularST.Value = vbUnchecked Then
        'Se calculou diferente das demais telas muda para não trocar
        Call gobjTribTab.ForcaICMSTipoManual
    End If
    
    Call objOrcamentoVenda.AnaliseVenda_Preenche(objAnaliseVendas)
    lErro = CF("AnaliseVenda_ValidaDescMax", objAnaliseVendas)
    If lErro <> SUCESSO Then gError 181786
    
    'Grava no BD
    lErro = CF("OrcamentoVenda_Grava", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84317 '46183

    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
        
            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_ORCVENDA
            gobjAnotacao.sID = CStr(objOrcamentoVenda.iFilialEmpresa) & "," & CStr(objOrcamentoVenda.lCodigo)
            gobjAnotacao.dtDataAlteracao = gdtDataHoje
            
            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 141373
            
        End If
        
    End If
    
    'Se for para imprimir o orçamento depois da gravação
    If ImprimeOrcamentoGravacao.Value = vbChecked Then
        
        'Dispara função para imprimir orçamento
        lErro = Orcamento_Imprime(Trim(objOrcamentoVenda.lCodigo), True)
        If lErro <> SUCESSO Then gError 102239
    
    End If

    'Se for para enviar por email o orçamento depois da gravação
    If EmailOrcamentoGravacao.Value = vbChecked Then
        
        'Dispara função para enviar por email orçamento
        lErro = Orcamento_Email(objOrcamentoVenda.lCodigo)
        If lErro <> SUCESSO And lErro <> 102237 Then gError 131029
    
    End If
   
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 56931, 141385, 141373, 181786
        
        Case 84308
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)

        Case 84309
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 84310
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)

        Case 84311
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 84312
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_OV_NEGATIVO", gErr)

        Case 84313, 84314, 84315, 84316, 84317, 131029

        Case 94496
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)
            
        Case 206849
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_OV_HISTORICO", gErr)
        
        Case 211526
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157788)

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
    If lErro <> SUCESSO Then gError 84147 '31389

    Select Case objControl.Name

        Case Produto.Name
            'Se o produto estiver preenchido desabilita
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
            End If

        Case VersaoKit.Name
            'Se o produto estiver preenchido desabilita
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                Produto.Enabled = True
            Else
                Produto.Enabled = False
            End If

        Case VersaoKitBase.Name
            'Se o produto estiver preenchido desabilita
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                Produto.Enabled = True
            Else
                Produto.Enabled = False
            End If
            
        Case DescricaoProduto.Name
            If gobjFAT.iOVBloqEditDescProd = MARCADO Then
                DescricaoProduto.Enabled = False
            Else
                DescricaoProduto.Enabled = True
            End If

        Case UnidadeMed.Name
            'guarda a um go grid nessa coluna
            sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)

            UnidadeMed.Enabled = True

            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = UnidadeMed.Text

            UnidadeMed.Clear

            If iProdutoPreenchido <> PRODUTO_VAZIO Then

                objProduto.sCodigo = sProdutoFormatado
                'Lê o produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 84148 '26644

                If lErro = 28030 Then gError 84149 '26645

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lê as UMs do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 84150 '26646
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

        Case PrecoUnitario.Name, PercentDesc.Name, Desconto.Name, DataEntrega.Name, MotivoPerdaItem.Name, StatusItem.Name, Observacao.Name 'Alterado por Wagner
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        '******** O TRATAMENTO DE QUANTIDADE FOI DESTACADO PARA TRATAR GRADE
        Case Quantidade.Name
                
            'Se o produto estiver preenchido, habilita o controle
            If iProdutoPreenchido = PRODUTO_VAZIO Or left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
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
            
        Case PrecoTotal.Name
            'Habilita os campos de desconto em sequencia
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

    End Select

    lErro = CF2(Me, "OV_Rotina_Grid_Enable", objControl.Name, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 141412
    
    If gobjCRFAT.iUsaBloqAcessoPorTelaControle = MARCADO Then
        lErro = CF("Rotina_Grid_Enable_BloqueiaAcesso", Me.Name, objControl)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 84148, 84150, 84147, 141412
        
        Case ERRO_SEM_MENSAGEM

        Case 84149
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157789)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto Data que está deixando de ser a corrente

Dim lErro As Long
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) > 0 Then

        sProduto = Produto.Text

        lErro = CF("Trata_Segmento_Produto", sProduto)
        If lErro <> SUCESSO Then gError 199362

        Produto.Text = sProduto

        lErro = Produto_Saida_Celula()
        If lErro <> SUCESSO And lErro <> 26658 Then gError 84152
        If lErro = 26658 Then gError 84151
    End If

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84153

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 84151 To 84153, 199362
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157790)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UM(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidadede Medida que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UM

    Set objGridInt.objControle = UnidadeMed

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84154 '26627

    lErro = CF2(Me, "OV_Calcula_PesoKg", GridItens.Row)
    If lErro <> SUCESSO Then gError 177170

    lErro = FormPreco_Calcula()
    If lErro <> SUCESSO Then gError 177170
    
    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case 84154, 177170
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157791)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

End Function

Private Function Saida_Celula_DescProduto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Descrição do Produto que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescProduto

    Set objGridInt.objControle = DescricaoProduto

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_DescProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_DescProduto:

    Saida_Celula_DescProduto = gErr

    Select Case gErr

        Case 129217
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157792)

    End Select

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade que está deixando de ser a corrente

Dim lErro As Long
Dim bQuantidadeIgual As Boolean
Dim iIndice As Integer
Dim dPrecoUnitario As Double
Dim dQuantAnterior As Double
Dim objProduto As New ClassProduto
Dim objKitVenda As New ClassKitVenda
Dim objItemRomaneio As ClassItemRomaneioGrade

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    bQuantidadeIgual = False

    If Len(Quantidade.Text) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 84157 '26665

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    'Comparação com quantidade anterior
    dQuantAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    If dQuantAnterior = StrParaDbl(Quantidade.Text) Then bQuantidadeIgual = True

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84161 '59727

    'Preço unitário
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))

    lErro = CF2(Me, "OV_Calcula_PesoKg", GridItens.Row)
    If lErro <> SUCESSO Then gError 177171

    'Recalcula preço do ítem e valor total da nota
    If Not bQuantidadeIgual Then
    
        lErro = CF2(Me, "OV_Saida_Celula_Quantidade", objGridInt)
        If lErro <> SUCESSO Then gError 84162
    
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 84162

        Call PrecoTotal_Calcula(GridItens.Row)
        
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 84162

        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 84162 '51037
    End If

    If Not bQuantidadeIgual Then
        lErro = FormPreco_Calcula(GridItens.Row)
        If lErro <> SUCESSO Then gError 84162 '51037
    End If
    
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 84161, 84157, 84162, 177171
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157793)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoUnitario(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim bPrecoUnitarioIgual As Boolean

On Error GoTo Erro_Saida_Celula_PrecoUnitario

    bPrecoUnitarioIgual = False

    Set objGridInt.objControle = PrecoUnitario

    If Len(Trim(PrecoUnitario.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(PrecoUnitario.Text)
        If lErro <> SUCESSO Then gError 84170  '26684

        PrecoUnitario.Text = Format(PrecoUnitario.Text, gobjFAT.sFormatoPrecoUnitario)
    
    End If

    'Comparação com Preço Unitário anterior
    If StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col)) = StrParaDbl(PrecoUnitario.Text) Then bPrecoUnitarioIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84171 '26685

    If Not bPrecoUnitarioIgual Then
    
        lErro = CF2(Me, "OV_Saida_Celula_PrecoUnitario", objGridInt)
        If lErro <> SUCESSO Then gError 84172
        
        Call PrecoTotal_Calcula(GridItens.Row)
        
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 84172
        
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 84172 '51042

        lErro = FormPreco_Calcula(GridItens.Row)
        If lErro <> SUCESSO Then gError 84172
    
    End If

   Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = gErr


    Select Case gErr

        Case 84170, 84171, 84172
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157794)

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
Dim sValorPercAnterior As String

On Error GoTo Erro_Saida_Celula_PercentDesc

    Set objGridInt.objControle = PercentDesc

    If Len(PercentDesc.Text) > 0 Then
        'Critica a porcentagem
        lErro = Porcentagem_Critica_Negativa(PercentDesc.Text) 'Alterado por Wagner
        If lErro <> SUCESSO Then gError 84329 '26694

        dPercentDesc = CDbl(PercentDesc.Text)

        If Format(dPercentDesc, "#0.#0\%") <> GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) Then
            'se for igual a 100% -> erro
            If dPercentDesc = 100 Then gError 84330 '26695

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
    If lErro <> SUCESSO Then gError 84331 '26696
    
    'Se foi alterada
    If Format(dPercentDesc, "#0.#0\%") <> sValorPercAnterior Then

        iDescontoAlterado = REGISTRO_ALTERADO
        
        'Recalcula o preço total
        Call PrecoTotal_Calcula(GridItens.Row)
        
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 84333

        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 84333 '51044
        
        'Preenche GridParcelas a partir da Condição de Pagto
        lErro = Cobranca_Automatica()
        If lErro <> SUCESSO Then gError 84332

    End If

    Saida_Celula_PercentDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDesc:

    Saida_Celula_PercentDesc = gErr

    Select Case gErr

        Case 84329, 84331, 84333
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84330
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84332

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157795)

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

    Call PrecoTotal_Calcula(GridItens.Row)

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 42222
'*** fim tributacao
    
    Call ValorTotal_Calcula
    
    'Preenche GridParcelas a partir da Condição de Pagto
    lErro = Cobranca_Automatica()
    If lErro <> SUCESSO Then gError 42222
    
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
        If lErro <> SUCESSO Then gError 84173 ' 26697
        'Se data de emissão estiver preenchida
        If Len(Trim(DataEmissao.ClipText)) > 0 Then

            dtDataEntrega = CDate(DataEntrega.Text)
            dtDataEmissao = CDate(DataEmissao.Text)
            'Veerifica se a data de emissão é maior que a data de entrega
            If dtDataEntrega < dtDataEmissao Then gError 84174 '26698

        End If
        
        GridItens.TextMatrix(GridItens.Row, iGrid_PrazoEntrega_Col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84175 '26699

    Saida_Celula_DataEntrega = SUCESSO

    Exit Function

Erro_Saida_Celula_DataEntrega:

    Saida_Celula_DataEntrega = gErr

    Select Case gErr

        Case 84173, 84175
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84174
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAENTREGA", gErr, dtDataEntrega, dtDataEmissao)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157796)

    End Select

    Exit Function

End Function

Function Produto_Saida_Celula() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim iIndice As Integer
Dim sProduto As String
Dim vbMsgRes As VbMsgBoxResult
Dim objItemOV As ClassItemOV
Dim iPossuiGrade As Integer
Dim objRomaneioGrade As ClassRomaneioGrade
Dim colItensRomaneioGrade As New Collection
Dim objItensRomaneio As ClassItemRomaneioGrade
Dim sProdutoPai As String
Dim objGridItens1 As Object
Dim objKit As New ClassKit 'Inserido por Wagner
Dim sCodigoAntigo As String
Dim sCliente As String
Dim iFilialCli As Integer
Dim sRetorno As String
Dim lCliente As Long, sUM As String, iTabelaPreco As Integer
Dim objCliente As New ClassCliente
Dim dPercDescontoTab As Double, dPrecoComDescontoTab As Double

On Error GoTo Erro_Produto_Saida_Celula

'***********  FUNÇÃO ALTERADA PARA TRATAMENTO DE GRADE  ******************
    
    'Critica o Produto
    lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 84176
       
    If lErro = 86295 And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then
        gError 86296
    ElseIf Len(Trim(objProduto.sGrade)) > 0 Then
        'Sinaliza que o produto possui grade
        iPossuiGrade = MARCADO
    End If
    
    'Se o produto não foi encontrado ==> Pergunta se deseja criar
    If lErro = 51381 Then gError 84177
       
    lErro = CF2(Me, "OV_Produto_Saida_Celula3", objProduto)
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
                For iIndice = 1 To gobjOrcamentoVenda.colItens.Count
                    'Se aparecer ==> erro
                    If gobjOrcamentoVenda.colItens(iIndice).sProduto = sProdutoPai Then gError 86328
                
                Next
            
            End If
            
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
            For iIndice = 1 To gobjOrcamentoVenda.colItens.Count
                'Se aparecer ==> Erro
                If gobjOrcamentoVenda.colItens(iIndice).sProduto = objItensRomaneio.sProduto Then gError 86331
            Next
        Next
 
    End If
    
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 84178

        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True

    End If

    If objProduto.iGerencial = DESMARCADO Then
        lErro = CF2(Me, "OV_Produto_Valida")
        If lErro <> SUCESSO Then gError 84179
    End If
    
    lErro = CF2(Me, "OV_Produto_Saida_Celula", objProduto)
    If lErro <> SUCESSO Then gError 177172

    Set objItemOV = New ClassItemOV
    
    objItemOV.iPossuiGrade = iPossuiGrade

    If objItemOV.iPossuiGrade = MARCADO Then
        
        objItemOV.sProduto = objProduto.sCodigo
        objItemOV.sUnidadeMed = objProduto.sSiglaUMVenda
        objItemOV.lCodOrcamento = StrParaDbl(Codigo.Text)
        objItemOV.iItem = GridItens.Row
        objItemOV.lNumIntDoc = 0
        objItemOV.sDescricao = objProduto.sDescricao
                
        Set objRomaneioGrade = New ClassRomaneioGrade
        
        objRomaneioGrade.sNomeTela = Me.Name
        
        Set objRomaneioGrade.objObjetoTela = objItemOV
                    
        lErro = GrupoGrade_Valida_Itens(objRomaneioGrade)
        If lErro <> SUCESSO Then gError 86310
                    
        Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
        If giRetornoTela <> vbOK Then gError 86310
        
        lErro = GrupoGrade_Trata_Itens(objRomaneioGrade, objProduto)
        If lErro <> SUCESSO Then gError 86310

    End If

    'Verifica se é um kit
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then

        If objProduto.iCompras = PRODUTO_PRODUZIVEL Then
                        
            objKit.sProdutoRaiz = objProduto.sCodigo
            lErro = CF("Kit_Le_Padrao", objKit)
            If lErro <> SUCESSO And lErro <> 106304 Then gError 131270
            
            'Se encontrou => É UM KIT
            If lErro <> 106304 Then
                GridItens.TextMatrix(GridItens.Row, iGrid_VersaoKit_Col) = Codigo.Text & SEPARADOR & CStr(GridItens.Row)
                GridItens.TextMatrix(GridItens.Row, iGrid_VersaoKitBase_Col) = objKit.sVersao
            End If
        
        End If
        
    End If

    'Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda
    GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(gdDesconto, "Percent")

    lErro = gobjIdiomaTela.Traduzir("Produtos", "Descricao", objProduto.sDescricao, objProduto.sCodigo, 0, 0, Codigo_Extrai(Idioma.Text), sRetorno)
    If lErro <> SUCESSO Then gError 131270

    'Descricao Produto
    GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = sRetorno 'objProduto.sDescricao

    GridItens.TextMatrix(GridItens.Row, iGrid_Status_Col) = Status.Text
    GridItens.TextMatrix(GridItens.Row, iGrid_MotivoPerda_Col) = MotivoPerda.Text

    If StrParaDate(DataEntregaPV.Text) <> DATA_NULA Then
        GridItens.TextMatrix(GridItens.Row, iGrid_DataEntrega_Col) = Format(MaskedParaDate(DataEntregaPV), "dd/mm/yyyy")
    End If
    If StrParaInt(PrazoEntrega.Text) <> 0 Then
        GridItens.TextMatrix(GridItens.Row, iGrid_PrazoEntrega_Col) = PrazoEntrega.Text
    End If

    'Preço Unitário
    
    'Coloca Produto no grid (necessario p/usar Trata_TabelaPreco_Item)
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
            
    iFilialCli = Codigo_Extrai(Filial.Text)
    dPrecoUnitario = 0
    
    lErro = CF2(Me, "NFCF_PrecoUnitario_Calcula", giFilialEmpresa, objProduto.sCodigo, objProduto.sSiglaUMVenda, Cliente.Text, iFilialCli, Codigo_Extrai(TabelaPreco.Text), StrParaDbl(PercAcrescFin.Text) / 100, Codigo_Extrai(Moeda.Text), objRomaneioGrade, dPrecoUnitario, True, dPercDescontoTab, dPrecoComDescontoTab)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If dPercDescontoTab >= gdDesconto Then
        GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercDescontoTab, "Percent")
    End If
    
    lErro = CF2(Me, "OV_Move_PrecoUnitario_Tela", dPrecoUnitario, GridItens.Row)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'precodesc
    Set objGridItens1 = GridItens
    sCliente = Cliente.Text
    Call CF("Produto_Saida_Celula_PrecoDesc", objGridItens1, GridItens.Row, iGrid_PrecoUnitario_Col + 1, dPrecoUnitario, objProduto.sCodigo, iGrid_DescProduto_Col, sCliente, iFilialCli, gdDesconto, iGrid_PrecoUnitario_Col + 2)

    'Acrescenta uma linha no Grid se for o caso
    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        
        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItens.Row)

        Call FormPreco_Adiciona_Linha(GridItens.Row, objProduto)
        
        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 101105
    
        If iPossuiGrade = MARCADO Or objProduto.iKitVendaComp = MARCADO Then
        
            '************** GRADE ************
            gobjOrcamentoVenda.colItens(GridItens.Row).iPossuiGrade = MARCADO

            gobjOrcamentoVenda.colItens(GridItens.Row).iItem = GridItens.Row
            gobjOrcamentoVenda.colItens(GridItens.Row).sProduto = objProduto.sCodigo
                       
            Set gobjOrcamentoVenda.colItens(GridItens.Row).colItensRomaneioGrade = objItemOV.colItensRomaneioGrade
            
            If objProduto.iKitVendaComp <> MARCADO Then

                GridItens.TextMatrix(GridItens.Row, 0) = "# " & GridItens.TextMatrix(GridItens.Row, 0)
                       
                Call Atualiza_Grid_Itens(objItemOV)
                
            End If
                        
        End If
    
    End If
    
    If objItemOV.iPossuiGrade = MARCADO Then
        Call PrecoTotal_Calcula(GridItens.Row)
        Call ValorTotal_Calcula
    End If
    
    lErro = GrupoGrade_Separa_Itens(objRomaneioGrade)
    If lErro <> SUCESSO Then gError 101105

    'Atualiza a checkbox do grid para exibir a figura marcada/desmarcada
    Call Grid_Refresh_Checkbox(objGridItens)
    

    Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_Produto_Saida_Celula:

    Produto_Saida_Celula = gErr

    Select Case gErr
    
        Case 86328 ' Inserido por Wagner
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_GRID", gErr, sProdutoPai, Produto.Text)

        Case 84178, 131014 ' Inserido por Wagner
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, Produto.Text)

        Case 84176, 84181, 84182, 84183, 84184, 84334, 101105, 86310, 129999, 131000, 131270, 131291, 131292, 177172, 177395, 177582 'Alterado por Wagner

        Case 84177
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
            End If

        Case 84179, ERRO_SEM_MENSAGEM
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, Produto.Text, Produto.Text, iIndice)

        Case 86296
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case 86330
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, Produto.Text)

        Case 86331
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILHO_GRADE_GRID", gErr, Trim(objProduto.sCodigo), Trim(gobjOrcamentoVenda.colItens(iIndice).sProduto))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157797)

    End Select

    Exit Function

End Function

Public Sub PrecoTotal_Calcula(iLinha As Integer)

Dim dPrecoTotal As Double
Dim dPrecoTotalReal As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim lTamanho As Long
Dim dValorTotal As Double, dValorTotalB As Double, dValorDescontoItens As Double
Dim iIndice As Integer
Dim dValorDesconto As Double
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long
Dim objGridItens1 As Object
Dim objItemOV As ClassItemOV

On Error GoTo Erro_PrecoTotal_Calcula

    'Quantidades e preço unitário
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col))
    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))

    lErro = CF2(Me, "OV_PrecoTotal_Calcula", iLinha, dQuantidade)
    If lErro <> SUCESSO Then gError 177173
    'Cálculo do desconto
    lTamanho = Len(Trim(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col)))
    If lTamanho > 0 Then
        dPercentDesc = CDbl(Format(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col), "General Number"))
    Else
        dPercentDesc = 0
    End If

    dPrecoTotal = dPrecoUnitario * (dQuantidade)

'    'Se percentual for >0 tira o desconto
'    'If dPercentDesc > 0 Then dDesconto = dPercentDesc * dPrecoTotal
'    dDesconto = dPercentDesc * dPrecoTotal
'    dPrecoTotalReal = dPrecoTotal - dDesconto 'Inserido por Wagner
'
'    'precodesc
'    Set objGridItens1 = GridItens
'    Call CF("PrecoTotal_Calcula_PrecoDesc", objGridItens1, iLinha, iGrid_PrecoUnitario_Col + 1, Format(dPrecoUnitario * (1 - dPercentDesc), "Standard"))
'
'    'Coloca valor do desconto no Grid
'    GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
'
'    'Coloca preco total do ítem no grid
'    GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(dPrecoTotalReal, "Standard")
'    GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(dPrecoTotal, "Standard")

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

    'Calcula a soma dos valores de produtos
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        Set objItemOV = gobjOrcamentoVenda.colItens(iIndice)
        
        If objItemOV.objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
    
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) > 0 Then
                If CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)) > 0 Then dValorTotal = dValorTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
            End If
            dValorTotalB = dValorTotalB + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
            dValorDescontoItens = dValorDescontoItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
    
        End If
    
    Next

'    If gdDesconto > 0 Then
'        dValorDesconto = gdDesconto * dValorTotal
'    ElseIf Len(Trim(ValorDesconto.Text)) > 0 And IsNumeric(ValorDesconto.Text) Then
'        dValorDesconto = CDbl(ValorDesconto.Text)
'    End If
    dValorDesconto = 0
'    dValorTotal = dValorTotal - dValorDesconto

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
'        gobjOrcamentoVenda.dValorDesconto = dValorDesconto
'
'        Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, dValorDesconto)
'
'        'Faz a atualização dos valores da tributação
'        lErro = gobjTribTab.AtualizarTributacao()
'        If lErro <> SUCESSO Then gError 101104
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

    Exit Sub

Erro_PrecoTotal_Calcula:

    Select Case gErr

        Case 84195, 84247, 101104, 177173

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157798)

    End Select

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
        'se os percentuais (passado e calulado) forem diferentes calcula-se o desconto
        If dPercentDesc <> dPercDesc2 Then dDesconto = Arredonda_Moeda(dPercentDesc * dValorTotal)

    End If

    dValorReal = dValorTotal - dDesconto

End Sub

Private Function Saida_Celula_DataVencimento(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Vencimento que está deixando de serr a corrente

Dim lErro As Long
Dim dtDataReferencia As Date
Dim dtDataVencimento As Date
Dim sDataVencimento As String
Dim iCriouLinha As Boolean

On Error GoTo Erro_Saida_Celula_DataVencimento

    Set objGridInt.objControle = DataVencimento

    'Verifica se Data de Vencimento esta preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then gError 84196 '26726

         dtDataVencimento = CDate(DataVencimento.Text)

        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataReferencia.ClipText)) > 0 Then
            dtDataReferencia = CDate(DataReferencia.Text)
            If dtDataVencimento < dtDataReferencia Then gError 84197 '26728
        End If

        sDataVencimento = Format(dtDataVencimento, "dd/mm/yyyy")

        iCriouLinha = False
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            iCriouLinha = True
        End If

    End If

    If sDataVencimento <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col) Then CobrancaAutomatica.Value = vbUnchecked

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84198 '26727

    If iCriouLinha Then
        'Coloca desconto padrao (le em CPRConfig)
        lErro = Preenche_DescontoPadrao(GridParcelas.Row)
        If lErro <> SUCESSO Then gError 84199 '51032
    End If

    Saida_Celula_DataVencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_DataVencimento:

    Saida_Celula_DataVencimento = gErr

    Select Case gErr

        Case 84196, 84198, 84199
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84197
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR_REFERENCIA", gErr, dtDataVencimento, GridParcelas.Row, dtDataReferencia)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157799)

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

On Error GoTo Erro_Saida_Celula_ValorParcela

    Set objGridInt.objControle = ValorParcela

    'Verifica se valor está preenchido
    If Len(ValorParcela.ClipText) > 0 Then

        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(ValorParcela.Text)
        If lErro <> SUCESSO Then gError 84200 '26724

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
            'Coloca DescontoPadrao
            lErro = Preenche_DescontoPadrao(GridParcelas.Row)
            If lErro <> SUCESSO Then gError 84201  '51061

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
    If lErro <> SUCESSO Then gError 84202 '26725

    Saida_Celula_ValorParcela = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorParcela:

    Saida_Celula_ValorParcela = gErr

    Select Case gErr

        Case 84200, 84201, 84202
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157800)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoDesconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo Desconto que está deixando de serr a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim iTipo As Integer
Dim iIndice As Integer

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
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 84203 '26729

            'Não foi encontrado
            If lErro = 25085 Then gError 84204 '26730
            If lErro = 25086 Then gError 84205 '26731

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
        End If

    Else

        For iIndice = objGridInt.objGrid.Col To iGrid_Desc3Perc_Col
            GridParcelas.TextMatrix(GridParcelas.Row, iIndice) = ""
        Next

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84206 '26732

    Saida_Celula_TipoDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDesconto:

    Saida_Celula_TipoDesconto = gErr

    Select Case gErr

        Case 84203, 84206
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84204
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84205
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", gErr, objGridInt.objControle.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157801)

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
        If lErro <> SUCESSO Then gError 84207 '26792
        'Se a data de vencimento estiver preenchida
        If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))) > 0 Then
            'critica se DataDesconto ultrapassa DataVencimento
            If CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col)) < CDate(objGridInt.objControle.Text) Then gError 84208 '26592
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84209 '26793

    Saida_Celula_DescontoData = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoData:

    Saida_Celula_DescontoData = gErr

    Select Case gErr

        Case 84207, 84209
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 84208
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADESCONTO_MAIOR_DATAVENCIMENTO", gErr, CDate(objGridInt.objControle.Text), CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col)))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157802)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoValor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Valor que está deixando de serr a corrente

Dim lErro As Long
Dim dColunaSoma As Double

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
        If lErro <> SUCESSO Then gError 84210 '26733

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84211 '26734

    Saida_Celula_DescontoValor = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoValor:

    Saida_Celula_DescontoValor = gErr

    Select Case gErr

        Case 84210, 84211
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157803)

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
        lErro = Porcentagem_Critica_Negativa(objGridInt.objControle.Text) ' Alterado por Wagner
        If lErro <> SUCESSO Then gError 84212 '26794

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
    If lErro <> SUCESSO Then gError 84213 '26795

    Saida_Celula_DescontoPerc = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoPerc:

    Saida_Celula_DescontoPerc = gErr

    Select Case gErr

        Case 84212, 84213
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157804)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Itens() As Long

Dim iIndice As Integer
Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Valida_Grid_Itens

    'Verifica se há itens no grid
    If objGridItens.iLinhasExistentes = 0 Then gError 84214 '26813

    'para cada item do grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        lErro = CF2(Me, "OV_Valida_Grid_Itens", iIndice)
        If lErro <> SUCESSO Then gError 177174

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 84215 '51455

        lErro = Valor_Positivo_Critica(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        If lErro <> SUCESSO Then gError 84216 '26814

        If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_VersaoKit_Col))) = 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_VersaoKitBase_Col))) <> 0 Then gError 129977
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_VersaoKitBase_Col))) = 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_VersaoKit_Col))) <> 0 Then gError 129978
        End If
    
    
    Next

    Valida_Grid_Itens = SUCESSO

    Exit Function

Erro_Valida_Grid_Itens:

    Valida_Grid_Itens = gErr

    Select Case gErr

        Case 84214
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_OV", gErr)

        Case 84215
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case 84216, 84218, 177174

        Case 84217
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 129977
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAOKIT_ITEM_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 129978
            Call Rotina_Erro(vbOKOnly, "ERRO_VERSAOKITBASE_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157805)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Parcelas() As Long
'Valida os dados do Grid de Parcelas

Dim lErro As Long
Dim iIndice As Integer
Dim dSomaParcelas As Double
Dim dValorTotal As Double
Dim dValorPagar As Double
Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim iTamanho As Integer
Dim iTipo As Integer
Dim dPercAcrecFin As Double
Dim iDesconto As Integer
Dim dtDataDesconto As Date

On Error GoTo Erro_Valida_Grid_Parcelas

    'Verifica se alguma parcela foi informada
''''''    If objGridParcelas.iLinhasExistentes = 0 Then gError 84219 '26817

    dSomaParcelas = 0

    'Para cada Parcela do grid de parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        dtDataEmissao = StrParaDate(DataEmissao.Text)
        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) = 0 Then gError 84220 '26818
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) = 0 Then gError 84221 '26821

        'Se o tipo de desconto 1 estiver preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))) > 0 Then
            iDesconto = 1
            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
            'Verifica se a data de desconto está preenchdida
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))) = 0 Then gError 84223 '51066
            'Recolhe o Valor ou Percentual de desconto
            If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))) = 0 Then gError 84224 '51069
            Else
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col))) = 0 Then gError 84225 '51070
            End If
            'Se o tipo de desconto 2 estiver preenchido
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))) > 0 Then
                iDesconto = 2
                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
                'Verifica se a data de desconto está preenchdida
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))) = 0 Then gError 84226 '51067
                'Faz a crítica da ordem das datas de desconto
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 84227 '51075
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 84228 '51077
                'Recolhe o Valor ou Percentual de desconto
                If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))) = 0 Then gError 84229 '51071
                Else
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col))) = 0 Then gError 84230 '51072
                End If
                'Se o tipo de desconto 3 estiver preenchido
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))) > 0 Then
                    iDesconto = 3
                    iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
                    'Verifica se a data de desconto está preenchdida
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))) = 0 Then gError 84231 '51068
                    'Faz a crítica da ordem das datas de desconto
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 84232 '51076
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 84233 '51078
                    'Recolhe o Valor ou Percentual de desconto
                    If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))) = 0 Then gError 84234 '51073
                    Else
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col))) = 0 Then gError 84235 '51074
                    End If
                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
                    If dtDataDesconto > dtDataVencimento Then gError 84236 '51364
                Else
                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
                    If dtDataDesconto > dtDataVencimento Then gError 84237 '51363
                End If
            Else
                dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
                If dtDataDesconto > dtDataVencimento Then gError 84238 '51362
            End If
        End If



        If iIndice > 1 Then If CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col)) < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_Col)) Then gError 84239 '26820

        dSomaParcelas = dSomaParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))

    Next

    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then gError 84238
    
    If Abs(dValorPagar - dSomaParcelas) > DELTA_VALORMONETARIO Then gError 26822

    Valida_Grid_Parcelas = SUCESSO

    Exit Function

Erro_Valida_Grid_Parcelas:

    Valida_Grid_Parcelas = gErr

    Select Case gErr

        Case 84219
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_PARCELA_COBRANCA", gErr)

        Case 84220
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_NAO_INFORMADA", gErr, iIndice)

        Case 84222
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_MENOR", gErr, iIndice, dtDataVencimento, dtDataEmissao)
        Case 84221
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PARCELA_COBRANCA_NAO_INFORMADO", gErr, iIndice)

        Case 84223, 84226, 84231
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_NAO_PREENCHIDA", gErr, iDesconto, iIndice)

        Case 84224, 84225, 84229, 84230, 84234, 84235
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_PARCELA_NAO_PREENCHIDO", gErr, iDesconto, iIndice)

        Case 84227, 84232
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTOS_DESORDENADAS", gErr, iIndice)

        Case 84228, 84233
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTO_IGUAIS", gErr, iDesconto - 1, iDesconto, iIndice)

        Case 84238, 84237, 84236
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESC_PARCELA_SUP_DATA_VENCIMENTO", gErr, dtDataDesconto, iDesconto, iIndice)

        Case 26822
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_COBRANCA_INVALIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157806)

    End Select

    Exit Function

End Function

Private Function Move_OrcamentoVenda_Memoria(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Move os dados da tela para objOrcamentoVenda

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim dValorTotalParcelas As Double
Dim dValorIRRF As Double
Dim objVendedor As New ClassVendedor
Dim objProjetoInfo As Object
Dim objTabelaPreco As New ClassTabelaPreco

On Error GoTo Erro_OrcamentoVenda_Memoria

    If Len(Trim(Codigo.Text)) > 0 Then objOrcamentoVenda.lCodigo = CLng(Codigo.Text)

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 84239 '26779

        If lErro = SUCESSO Then
            'Guarda código do Cliente em objOrcamentoVenda
            objOrcamentoVenda.lCliente = objCliente.lCodigo
        End If
            
        objOrcamentoVenda.sNomeCli = objCliente.sNomeReduzido
            
    End If
    
    'Verifica se vendedor existe
    If Len(Trim(Vendedor.Text)) > 0 Then
        
        objVendedor.sNomeReduzido = Trim(Vendedor.Text)

        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 94418

        'Não encontrou o vendedor ==> erro
        If lErro = 25008 Then gError 94419

        objOrcamentoVenda.iVendedor = objVendedor.iCodigo

    End If
    
    'Verifica se vendedor2 existe
    If Len(Trim(Vendedor2.Text)) > 0 Then
        
        objVendedor.sNomeReduzido = Trim(Vendedor2.Text)

        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError 94418

        'Não encontrou o vendedor ==> erro
        If lErro = 25008 Then gError 94419

        objOrcamentoVenda.iVendedor2 = objVendedor.iCodigo

    End If
    
    objOrcamentoVenda.sContato = Contato.Text
    objOrcamentoVenda.sEmail = Email.Text
    
    'Verifica se a Filial está preenchida
    If Len(Trim(Filial.Text)) > 0 Then
        
        'Se o Cliente estiver cadastrado
        If objOrcamentoVenda.lCliente <> 0 Then
        
            'a filial tb deverá estar cadastrada e por isso teremos o código da filial na tela
            objOrcamentoVenda.iFilial = Codigo_Extrai(Filial.Text)
            objOrcamentoVenda.sNomeFilialCli = Nome_Extrai(Filial.Text)
                                                
        Else
            'se não, guardaremos o Texto digitado pelo usuário
            objOrcamentoVenda.sNomeFilialCli = Trim(Filial.Text)
            
        End If
            
    End If

    'Preenche objOrcamentoVenda com dados da tela
    objOrcamentoVenda.dtDataEmissao = MaskedParaDate(DataEmissao)
    objOrcamentoVenda.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)
    
    objTabelaPreco.iCodigo = objOrcamentoVenda.iTabelaPreco
    
    If objTabelaPreco.iCodigo <> 0 Then

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError ERRO_SEM_MENSAGEM
        
        If objTabelaPreco.iAtivo = DESMARCADO Then gError 213836
        
    End If
        
    objOrcamentoVenda.sNaturezaOp = Trim(gobjTribTab.sNatOpInterna)
    objOrcamentoVenda.dValorFrete = StrParaDbl(ValorFrete.Text)
    objOrcamentoVenda.dValorSeguro = StrParaDbl(ValorSeguro.Text)
    objOrcamentoVenda.dValorDesconto = StrParaDbl(ValorDesconto.Text)
    objOrcamentoVenda.dValorOutrasDespesas = StrParaDbl(ValorDespesas.Text)
    objOrcamentoVenda.dValorProdutos = StrParaDbl(ValorProdutos.Caption)
    objOrcamentoVenda.dValorTotal = StrParaDbl(ValorTotal.Caption)
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
    objOrcamentoVenda.iPrazoValidade = StrParaInt(PrazoValidade.Text)
    objOrcamentoVenda.dPercAcrescFinanceiro = StrParaDbl(PercAcrescFin.ClipText)
    objOrcamentoVenda.dtDataReferencia = MaskedParaDate(DataReferencia)
    objOrcamentoVenda.dValorOutrasDespesas = StrParaDbl(ValorDespesas.ClipText)
    objOrcamentoVenda.iCobrancaAutomatica = StrParaInt(CobrancaAutomatica.Value)
    objOrcamentoVenda.iCondicaoPagto = CondPagto_Extrai(CondicaoPagamento)
    objOrcamentoVenda.iIdioma = CondPagto_Extrai(Idioma)
    objOrcamentoVenda.sUsuario = gsUsuario
    objOrcamentoVenda.lCodigoBase = StrParaLong(CodigoBase.Text)
    
    'Move Grid Itens para memória
    lErro = Move_GridItens_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84241 '26781

    'Move Tab Cobrança para memória
    Call Move_TabCobranca_Memoria(objOrcamentoVenda)

    lErro = Move_TabComplemento_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa

    'Move Tributacao para objOrcamentoVenda
    Set objOrcamentoVenda.objTributacaoOV = gobjOrcamentoVenda.objTributacaoOV

    If Status.ListIndex <> -1 Then
        objOrcamentoVenda.lStatus = Status.ItemData(Status.ListIndex)
    End If
    If MotivoPerda.ListIndex <> -1 Then
        objOrcamentoVenda.lMotivoPerda = MotivoPerda.ItemData(MotivoPerda.ListIndex)
    End If
    
    If StatusComercial.ListIndex <> -1 Then
        objOrcamentoVenda.iStatusComercial = StatusComercial.ItemData(StatusComercial.ListIndex)
    End If
    
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_OV)
    If lErro <> SUCESSO Then gError 181761
    
    Set objOrcamentoVenda.objProjetoInfo = objProjetoInfo
   
    lErro = FormPreco_Move_Tela_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 181761
    
    If TrocarVersao.Value = vbChecked Then
        objOrcamentoVenda.iTrocarVersao = MARCADO
    Else
        objOrcamentoVenda.iTrocarVersao = DESMARCADO
    End If
    
    objOrcamentoVenda.iVersao = StrParaInt(OVVersao.Caption)
    
    Set objOrcamentoVenda.objTela = Me
    
    Set objOrcamentoVenda.objInfoAdic = gobjInfoAdic
    objOrcamentoVenda.dValorItens = StrParaDbl(ValorProdutos2.Caption)
    objOrcamentoVenda.dValorDescontoItens = StrParaDbl(ValorDescontoItens.Text)
    
    objOrcamentoVenda.dtDataEntrega = MaskedParaDate(DataEntregaPV)
    objOrcamentoVenda.dtDataEnvio = MaskedParaDate(DataEnvio)
    objOrcamentoVenda.dtDataPerda = MaskedParaDate(DataPerda)
    objOrcamentoVenda.dtDataProxCobr = MaskedParaDate(DataProxContato)
    objOrcamentoVenda.dtDataPrevReceb = MaskedParaDate(DataFechamento)
    objOrcamentoVenda.iPrazoEntrega = StrParaInt(PrazoEntrega.Text)
    objOrcamentoVenda.iMoeda = Codigo_Extrai(Moeda.Text)
    objOrcamentoVenda.sPrazoTexto = PrazoTexto.Text
    
    If OptDataPrazoEnt(1).Value Then
        objOrcamentoVenda.iDataEnt = PrazoEntVar.ItemData(PrazoEntVar.ListIndex)
    ElseIf OptDataPrazoEnt(0).Value Then
        objOrcamentoVenda.iDataEnt = OV_DATA_ENTREGA_DATA
    ElseIf OptDataPrazoEnt(2).Value Then
        objOrcamentoVenda.iDataEnt = OV_DATA_ENTREGA_TEXTO
    End If
    
    Set objOrcamentoVenda.objIdiomaTela = gobjIdiomaTela
    
    lErro = CF2(Me, "OV_Move_OrcamentoVenda_Memoria", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 141405
    
    Move_OrcamentoVenda_Memoria = SUCESSO

    Exit Function

Erro_OrcamentoVenda_Memoria:

    Move_OrcamentoVenda_Memoria = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 84239, 84241, 181761

        Case 84240
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)
        'Por Leo em 21/03/02
        Case 94418
        
        Case 94419
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objVendedor.sNomeReduzido)
        'Leo até aqui
        
        Case ERRO_SEM_MENSAGEM
        
        Case 213836
            Call Rotina_Erro(vbOKOnly, "ERRO_TABPRECO_INATIVA", gErr, objTabelaPreco.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157807)

    End Select

    Exit Function

End Function

Public Function OV_Move_OrcamentoVenda_Memoria(ByVal objCT As Object, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
    OV_Move_OrcamentoVenda_Memoria = SUCESSO
End Function


Private Sub Move_TabCobranca_Memoria(objOrcamentoVenda As ClassOrcamentoVenda)
'Recolhe os dados do tab de cobrança

Dim lTamanho As Long
Dim iIndice As Integer
Dim objParcelaOV As ClassParcelaOV

    'Recolhe os dados da Cobrança
    objOrcamentoVenda.iCobrancaAutomatica = CobrancaAutomatica.Value
    objOrcamentoVenda.dtDataReferencia = MaskedParaDate(DataReferencia)
    objOrcamentoVenda.iCondicaoPagto = CondPagto_Extrai(CondicaoPagamento)
    objOrcamentoVenda.dPercAcrescFinanceiro = StrParaDbl(PercAcrescFin.Text) / 100

    If objGridParcelas.iLinhasExistentes = 0 Then Exit Sub

    'Recolhe os Dados do Grid de Parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        Set objParcelaOV = New ClassParcelaOV

        objParcelaOV.iNumParcela = iIndice

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) > 0 Then
            objParcelaOV.dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        Else
            objParcelaOV.dtDataVencimento = DATA_NULA
        End If

        objParcelaOV.dValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
        objParcelaOV.iDesconto1Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
        objParcelaOV.iDesconto2Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
        objParcelaOV.iDesconto3Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))) > 0 Then
            objParcelaOV.dtDesconto1Ate = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
        Else
            objParcelaOV.dtDesconto1Ate = DATA_NULA
        End If
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))) > 0 Then
            objParcelaOV.dtDesconto2Ate = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
        Else
            objParcelaOV.dtDesconto2Ate = DATA_NULA
        End If
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))) > 0 Then
            objParcelaOV.dtDesconto3Ate = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
        Else
            objParcelaOV.dtDesconto3Ate = DATA_NULA
        End If

        If objParcelaOV.iDesconto1Codigo = VALOR_FIXO Or objParcelaOV.iDesconto1Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
            objParcelaOV.dDesconto1Valor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))
        ElseIf objParcelaOV.iDesconto1Codigo = Percentual Or objParcelaOV.iDesconto1Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col)))
            If lTamanho > 0 Then objParcelaOV.dDesconto1Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col))
        End If

        If objParcelaOV.iDesconto2Codigo = VALOR_FIXO Or objParcelaOV.iDesconto2Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
            objParcelaOV.dDesconto2Valor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))
        ElseIf objParcelaOV.iDesconto2Codigo = Percentual Or objParcelaOV.iDesconto2Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col)))
            If lTamanho > 0 Then objParcelaOV.dDesconto2Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col))
        End If

        If objParcelaOV.iDesconto3Codigo = VALOR_FIXO Or objParcelaOV.iDesconto3Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
            objParcelaOV.dDesconto3Valor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))
        ElseIf objParcelaOV.iDesconto3Codigo = Percentual Or objParcelaOV.iDesconto3Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col)))
            If lTamanho > 0 Then objParcelaOV.dDesconto3Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col))
        End If

        objOrcamentoVenda.colParcela.Add objParcelaOV

    Next
    
    objOrcamentoVenda.dValorDescontoTit = StrParaDbl(ValorDescontoTit.Text)

End Sub

Function Traz_OrcamentoVenda_Tela(objOrcamentoVenda As ClassOrcamentoVenda, ByVal iBrowse As Integer) As Long
'Coloca na tela os dados do Orcamento de Venda

Dim lErro As Long
Dim objFilial As AdmFiliais
Dim bCancel As Boolean
Dim lCodigo As Long, iIndice As Integer
Dim objItemOV As ClassItemOV
Dim dPercDesc As Double

On Error GoTo Erro_Traz_OrcamentoVenda_Tela

    gbCarregandoTela = True
    
    lCodigo = StrParaLong(Codigo.Text)
    
    Call Limpa_OrcamentoVenda2

    'Lê os dados do Orcamento de Venda
    lErro = CF("OrcamentoVenda_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101232 Then gError 84359
    If lErro = 101232 Then gError 84363

    'Lê a parte de Tributação
    lErro = CF("OrcamentoVenda_Le_Tributacao", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101162 Then gError 84360
    If lErro = 101162 Then gError 101282

    lErro = CF("OrcamentoVenda_Le_Itens_ComTributacao", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101278 Then gError 84469
    If lErro = 101278 Then gError 101280
        
    If iBrowse <> VALOR_CODIGOBASE Then
    
        lErro = CF("OrcamentoVenda_Le_ItensFormPreco", objOrcamentoVenda)
        If lErro <> SUCESSO Then gError 101100
        
        Set gcolItensOVPrecoCalc = New Collection
        For Each objItemOV In objOrcamentoVenda.colItens
            gcolItensOVPrecoCalc.Add objItemOV
        Next
        'Comentado por Wagner, com as duas coleções apontando para o mesmo endereço de memória
        'ao adicionar um item ou remover nas coleções em separado há a duplicidade
        'Set gcolItensOVPrecoCalc = objOrcamentoVenda.colItens
    
    End If
    
    lErro = CF("ParcelasOV_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101284 Then gError 84470
    
    If objOrcamentoVenda.dtDataUltAlt <> DATA_NULA Then DataUltAlt.Caption = Format(objOrcamentoVenda.dtDataUltAlt, "dd/mm/yyyy")
    OVVersao.Caption = CStr(objOrcamentoVenda.iVersao)
    HoraUltAlt.Caption = Format(objOrcamentoVenda.dHoraUltAlt, "HH:MM:SS")
  
    ValorTotal.Caption = Format(objOrcamentoVenda.dValorTotal, "Standard")
    ValorProdutos.Caption = Format(objOrcamentoVenda.dValorProdutos, "Standard")
    
    PercAcrescFin.Text = ""
    'Coloca os dados do Orcamento na tela
    
    Usuario.Caption = objOrcamentoVenda.sUsuario
    
    If iBrowse = VALOR_CODIGO Then
        Codigo.Text = objOrcamentoVenda.lCodigo
        If objOrcamentoVenda.lCodigoBase <> 0 Then CodigoBase.Text = objOrcamentoVenda.lCodigoBase
    Else
        CodigoBase.Text = objOrcamentoVenda.lCodigo
        If lCodigo <> 0 Then Codigo.Text = lCodigo
    End If
    
    'Se existe um código para o Cliente
    If objOrcamentoVenda.lCliente <> 0 Then
    
        Call Cliente_Formata(objOrcamentoVenda.lCliente)
        Call Filial_Formata(Filial, objOrcamentoVenda.iFilial)
                
    Else 'Trecho por Leo em 18/04/02 ***
        
        'Preenche o Cliente e a Filial com os Nomes Informados.
        Cliente.Text = objOrcamentoVenda.sNomeCli
        Filial.Text = objOrcamentoVenda.sNomeFilialCli
                
    End If 'Leo até aqui ***
    
    giFilialAlterada = 0
    
    ValorFrete.Text = Format(objOrcamentoVenda.dValorFrete, "Standard")
    ValorSeguro.Text = Format(objOrcamentoVenda.dValorSeguro, "Standard")
    ValorDesconto.Text = Format(objOrcamentoVenda.dValorDesconto, "Standard")
    ValorDespesas.Text = Format(objOrcamentoVenda.dValorOutrasDespesas, "Standard")

    giValorFreteAlterado = 0
    giValorSeguroAlterado = 0
    giValorDescontoAlterado = 0
    giValorDespesasAlterado = 0

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(objOrcamentoVenda.dtDataEmissao, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    If objOrcamentoVenda.iPrazoValidade <> 0 Then 'Incluido por Leo em 29/04/02
        PrazoValidade.Text = objOrcamentoVenda.iPrazoValidade
    End If
    
    'Se a tabela de preços estiver preenchida coloca na tela
    If objOrcamentoVenda.iTabelaPreco > 0 Then
        TabelaPreco.Text = objOrcamentoVenda.iTabelaPreco
        Call TabelaPreco_Validate(bSGECancelDummy)
    Else
        TabelaPreco.Text = ""
    End If
    
    'Preenche o campo de vendedores
    If objOrcamentoVenda.iVendedor <> 0 Then
        
        Vendedor.Text = objOrcamentoVenda.iVendedor
        Call Vendedor_Validate(bCancel)
    
    End If
    
    'Preenche o campo de vendedor 2
    If objOrcamentoVenda.iVendedor2 <> 0 Then
        
        Vendedor2.Text = objOrcamentoVenda.iVendedor2
        Call Vendedor2_Validate(bCancel)
    
    End If
    
    Contato.Text = objOrcamentoVenda.sContato
    Email.Text = objOrcamentoVenda.sEmail
    
    'Carrega o Tab Cobrança
    lErro = Carrega_Tab_Cobranca(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84361 '51164

    'Carrega o Tab Complemento
    lErro = Carrega_Tab_Complemento(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84361 '51164

    'Carrega o Grid de itens
    lErro = Carrega_Grid_Itens(objOrcamentoVenda, iBrowse)
    If lErro <> SUCESSO Then gError 84362 '26569
    
    ValorTotal.Caption = Format(objOrcamentoVenda.dValorTotal, "Standard")
        
    If objOrcamentoVenda.lStatus <> 0 Then Call Combo_Seleciona_ItemData(Status, objOrcamentoVenda.lStatus)
    If objOrcamentoVenda.lMotivoPerda <> 0 Then Call Combo_Seleciona_ItemData(MotivoPerda, objOrcamentoVenda.lMotivoPerda)
    
    If objOrcamentoVenda.iStatusComercial <> 0 Then Call Combo_Seleciona_ItemData(StatusComercial, objOrcamentoVenda.iStatusComercial)
    
    If objOrcamentoVenda.iIdioma <> 0 Then Call Combo_Seleciona_ItemData(Idioma, objOrcamentoVenda.iIdioma)
    
    If objOrcamentoVenda.iHistorico <> MARCADO Then
        objOrcamentoVenda.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_OV
        lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objOrcamentoVenda.objProjetoInfo, objOrcamentoVenda.lNumIntDoc, PRJ_CR_TIPO_OV)
    Else
        objOrcamentoVenda.objInfoAdic.iTipoDoc = TIPODOC_INFOADIC_OVHIST
        lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objOrcamentoVenda.objProjetoInfo, objOrcamentoVenda.lNumIntDoc, PRJ_CR_TIPO_OVHIST)
    End If
    If lErro <> SUCESSO Then gError 181676
    
    objOrcamentoVenda.objInfoAdic.lNumIntDoc = objOrcamentoVenda.lNumIntDoc
    lErro = CF("InfoAdicionais_Le", objOrcamentoVenda.objInfoAdic)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 181676
    
    lErro = PrecosCalculados_Traz_Dados_Tela()
    If lErro <> SUCESSO Then gError 181676
    
    lErro = gobjTribTab.Traz_NFiscal_Tela(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 181676
    
    'Carrega o Tab Complemento
    lErro = Carrega_Tab_Complemento(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 181676
    
    If objOrcamentoVenda.dValorItens = 0 Then
        For iIndice = 1 To objGridItens.iLinhasExistentes
            objOrcamentoVenda.dValorItens = objOrcamentoVenda.dValorItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
            objOrcamentoVenda.dValorDescontoItens = objOrcamentoVenda.dValorDescontoItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
        Next
    End If
    
    ValorProdutos2.Caption = Format(objOrcamentoVenda.dValorItens, "Standard")
    ValorDescontoItens.Text = Format(objOrcamentoVenda.dValorDescontoItens, "Standard")
    If objOrcamentoVenda.dValorItens > 0 Then
        dPercDesc = objOrcamentoVenda.dValorDescontoItens / objOrcamentoVenda.dValorItens
    Else
        dPercDesc = 0
    End If
    PercDescontoItens.Text = Format(dPercDesc * 100, "Fixed")
    
    dValorDescontoItensAnt = objOrcamentoVenda.dValorDescontoItens
    dPercDescontoItensAnt = dPercDesc
    
    Set gobjInfoAdic = objOrcamentoVenda.objInfoAdic
    
    ValorDescontoTit.Text = Format(objOrcamentoVenda.dValorDescontoTit, "Standard")
    Call ValorDescontoTit_Validate(bSGECancelDummy)
    
    If CalcularST.Value = vbChecked Then
        gobjOrcamentoVenda.iNaoCalcICMSST = DESMARCADO
    Else
        gobjOrcamentoVenda.iNaoCalcICMSST = MARCADO
    End If
    
    Call DateParaMasked(DataEnvio, gobjOrcamentoVenda.dtDataEnvio)
    Call DateParaMasked(DataPerda, gobjOrcamentoVenda.dtDataPerda)
    Call DateParaMasked(DataFechamento, gobjOrcamentoVenda.dtDataPrevReceb)
    Call DateParaMasked(DataProxContato, gobjOrcamentoVenda.dtDataProxCobr)
    Call Combo_Seleciona_ItemData(Moeda, gobjOrcamentoVenda.iMoeda)
    If gobjOrcamentoVenda.iDataEnt = OV_DATA_ENTREGA_DATA Then
        Call DateParaMasked(DataEntregaPV, gobjOrcamentoVenda.dtDataEntrega)
        OptDataPrazoEnt(0).Value = True
    ElseIf gobjOrcamentoVenda.iDataEnt = OV_DATA_ENTREGA_TEXTO Then
        OptDataPrazoEnt(2).Value = True
        PrazoTexto.Text = gobjOrcamentoVenda.sPrazoTexto
    Else
        OptDataPrazoEnt(1).Value = True
        If gobjOrcamentoVenda.iPrazoEntrega <> 0 Then
            PrazoEntrega.PromptInclude = False
            PrazoEntrega.Text = CStr(gobjOrcamentoVenda.iPrazoEntrega)
            PrazoEntrega.PromptInclude = True
        End If
        Call Combo_Seleciona_ItemData(PrazoEntVar, gobjOrcamentoVenda.iDataEnt)
    End If
    Call OptDataPrazoEnt_Click(0)
    Call PrazoEntrega_Validate(bSGECancelDummy)
    
    lErro = CF2(Me, "OV_Traz_OrcamentoVenda_Tela", objOrcamentoVenda, iBrowse)
    If lErro <> SUCESSO Then gError 181676
    
    Call gobjIdiomaTela.Reiniciar
    
    lErro = gobjIdiomaTela.Associar("OrcamentoVenda", 0, "", objOrcamentoVenda.lNumIntDoc, 0)
    If lErro <> SUCESSO Then gError 181676
    
    lErro = CF("IdiomaTextos_Le", gobjIdiomaTela)
    If lErro <> SUCESSO Then gError 181676
    
    giDataEntregaAlterada = 0
    
    iCodigoAlterado = 0
    iAlterado = 0
    giPercAcresFinAlterado = 0

    gbCarregandoTela = False

    Traz_OrcamentoVenda_Tela = SUCESSO

    Exit Function

Erro_Traz_OrcamentoVenda_Tela:

    gbCarregandoTela = False

    Traz_OrcamentoVenda_Tela = gErr

    Select Case gErr

        Case 84359 To 84363, 84469, 84470, 101100, 101101, 27640, 101282, 101286, 181676
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157811)

    End Select

    Exit Function

End Function

Public Sub Form_Activate()

    Dim lErro As Long

On Error GoTo Erro_Form_Activate

'???? Trocar numeracao de erro
    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError 59332
        
    Call TelaIndice_Preenche(Me)

    Exit Sub
     
Erro_Form_Activate:

    Select Case gErr
          
        Case 59332
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157812)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""
'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "OrcamentoVenda"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_OrcamentoVenda_Memoria(objOrcamentoVenda)
    If lErro <> SUCESSO Then gError 84404 '51159

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objOrcamentoVenda.lCodigo, 0, "Codigo"
    colCampoValor.Add "Cliente", objOrcamentoVenda.lCliente, 0, "Cliente"
    colCampoValor.Add "Filial", objOrcamentoVenda.iFilial, 0, "Filial"
    colCampoValor.Add "CondicaoPagto", objOrcamentoVenda.iCondicaoPagto, 0, "CondicaoPagto"
    colCampoValor.Add "PercAcrescFinanceiro", objOrcamentoVenda.dPercAcrescFinanceiro, 0, "PercAcrescFinanceiro"
    colCampoValor.Add "DataEmissao", objOrcamentoVenda.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "ValorTotal", objOrcamentoVenda.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "ValorFrete", objOrcamentoVenda.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorDesconto", objOrcamentoVenda.dValorDesconto, 0, "ValorDesconto"
    colCampoValor.Add "ValorSeguro", objOrcamentoVenda.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "TabelaPreco", objOrcamentoVenda.iTabelaPreco, 0, "TabelaPreco"
    colCampoValor.Add "Vendedor", objOrcamentoVenda.iVendedor, 0, "Vendedor"
    colCampoValor.Add "Vendedor2", objOrcamentoVenda.iVendedor2, 0, "Vendedor2"
    colCampoValor.Add "NomeCli", objOrcamentoVenda.sNomeCli, STRING_CLIENTE_RAZAO_SOCIAL, "NomeCli"
    colCampoValor.Add "NomeFilialCli", objOrcamentoVenda.sNomeFilialCli, STRING_FILIAL_CLIENTE_NOME, "NomeFilialCli"
    colCampoValor.Add "NaturezaOp", objOrcamentoVenda.sNaturezaOp, STRING_NATUREZAOP_CODIGO, "NaturezaOp" 'por Leo em 02/05/02
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 84404

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157813)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda

On Error GoTo Erro_Tela_Preenche

    objOrcamentoVenda.lCodigo = colCampoValor.Item("Codigo").vValor
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa

    If objOrcamentoVenda.lCodigo <> 0 Then

        'Mostra os dados do Pedido de Venda na tela
        lErro = Traz_OrcamentoVenda_Tela(objOrcamentoVenda, VALOR_CODIGO)
        If lErro <> SUCESSO Then gError 84405

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 84405

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157814)

    End Select

    Exit Sub

End Sub

Function Carrega_Tab_Cobranca(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Coloca os dados do tab de cobrança na tela

Dim objParcelaOV As ClassParcelaOV
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim dValorDesconto As Double

    Call DateParaMasked(DataReferencia, objOrcamentoVenda.dtDataReferencia)
    giDataReferenciaAlterada = 0

    PercAcrescFin.Text = ""

    If objOrcamentoVenda.iCondicaoPagto > 0 Then
        CondicaoPagamento.Text = objOrcamentoVenda.iCondicaoPagto
        Call CondicaoPagamento_Validate(bSGECancelDummy)
    Else
        CondicaoPagamento.Text = ""
    End If

    PercAcrescFin.Text = Format(objOrcamentoVenda.dPercAcrescFinanceiro * 100, "Fixed")

    CobrancaAutomatica.Value = objOrcamentoVenda.iCobrancaAutomatica

    'Limpa o Grid de Parcelas antes de preencher com dados da coleção
    Call Grid_Limpa(objGridParcelas)

    iIndice = 0

    For Each objParcelaOV In objOrcamentoVenda.colParcela

        iIndice = iIndice + 1
        GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objParcelaOV.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objParcelaOV.dValor, "Standard")
        If objParcelaOV.dtDesconto1Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col) = Format(objParcelaOV.dtDesconto1Ate, "dd/mm/yyyy")
        If objParcelaOV.dtDesconto2Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col) = Format(objParcelaOV.dtDesconto2Ate, "dd/mm/yyyy")
        If objParcelaOV.dtDesconto3Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col) = Format(objParcelaOV.dtDesconto3Ate, "dd/mm/yyyy")
        If objParcelaOV.iDesconto1Codigo = VALOR_FIXO Or objParcelaOV.iDesconto1Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(objParcelaOV.dDesconto1Valor, "Standard")
        ElseIf objParcelaOV.iDesconto1Codigo = Percentual Or objParcelaOV.iDesconto1Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col) = Format(objParcelaOV.dDesconto1Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaOV.dValor > 0 Then
                dValorDesconto = objParcelaOV.dDesconto1Valor * objParcelaOV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        If objParcelaOV.iDesconto2Codigo = VALOR_FIXO Or objParcelaOV.iDesconto2Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(objParcelaOV.dDesconto2Valor, "Standard")
        ElseIf objParcelaOV.iDesconto2Codigo = Percentual Or objParcelaOV.iDesconto2Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col) = Format(objParcelaOV.dDesconto2Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaOV.dValor > 0 Then
                dValorDesconto = objParcelaOV.dDesconto2Valor * objParcelaOV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        If objParcelaOV.iDesconto3Codigo = VALOR_FIXO Or objParcelaOV.iDesconto3Codigo = VALOR_ANT_DIA Or objParcelaOV.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(objParcelaOV.dDesconto3Valor, "Standard")
        ElseIf objParcelaOV.iDesconto3Codigo = Percentual Or objParcelaOV.iDesconto3Codigo = PERC_ANT_DIA Or objParcelaOV.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col) = Format(objParcelaOV.dDesconto3Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaOV.dValor > 0 Then
                dValorDesconto = objParcelaOV.dDesconto3Valor * objParcelaOV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        For iIndice2 = 0 To TipoDesconto1.ListCount - 1
            If TipoDesconto1.ItemData(iIndice2) = objParcelaOV.iDesconto1Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col) = TipoDesconto1.List(iIndice2)
            If TipoDesconto2.ItemData(iIndice2) = objParcelaOV.iDesconto2Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col) = TipoDesconto2.List(iIndice2)
            If TipoDesconto3.ItemData(iIndice2) = objParcelaOV.iDesconto3Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col) = TipoDesconto3.List(iIndice2)
        Next

    Next

    objGridParcelas.iLinhasExistentes = iIndice

    Carrega_Tab_Cobranca = SUCESSO

    Exit Function

End Function

Private Function Carrega_Grid_Itens(objOrcamentoVenda As ClassOrcamentoVenda, ByVal iBrowse As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim objGridItens1 As Object

On Error GoTo Erro_Carrega_Grid_Itens

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    For iIndice = 1 To objOrcamentoVenda.colItens.Count

        objOrcamentoVenda.colItens(iIndice).iItem = iIndice

        lErro = Mascara_RetornaProdutoEnxuto(objOrcamentoVenda.colItens(iIndice).sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 84406

        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        'Calcula o percentual de desconto
        If objOrcamentoVenda.colItens(iIndice).dPrecoTotal + objOrcamentoVenda.colItens(iIndice).dValorDesconto > 0 Then
            dPercDesc = objOrcamentoVenda.colItens(iIndice).dValorDesconto / (objOrcamentoVenda.colItens(iIndice).dPrecoTotal + objOrcamentoVenda.colItens(iIndice).dValorDesconto)
        Else
            dPercDesc = 0
        End If

        '****** IF INCLUÍDO PARA TRATAMENTO DE GRADE ***************
        If objOrcamentoVenda.colItens(iIndice).iPossuiGrade = MARCADO Then GridItens.TextMatrix(iIndice, 0) = "# " & GridItens.TextMatrix(iIndice, 0)
        
        'Coloca os dados dos itens na tela
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objOrcamentoVenda.colItens(iIndice).sDescricao
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objOrcamentoVenda.colItens(iIndice).sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objOrcamentoVenda.colItens(iIndice).dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col) = Format(objOrcamentoVenda.colItens(iIndice).dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        
        If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
            If iBrowse = VALOR_CODIGO Then
                GridItens.TextMatrix(iIndice, iGrid_VersaoKit_Col) = objOrcamentoVenda.colItens(iIndice).sVersaoKit
                GridItens.TextMatrix(iIndice, iGrid_VersaoKitBase_Col) = objOrcamentoVenda.colItens(iIndice).sVersaoKitBase
            Else
                GridItens.TextMatrix(iIndice, iGrid_VersaoKit_Col) = Codigo.Text & SEPARADOR & iIndice
                GridItens.TextMatrix(iIndice, iGrid_VersaoKitBase_Col) = objOrcamentoVenda.colItens(iIndice).sVersaoKit
            End If
        End If
        
        'precoDesc
        Set objGridItens1 = GridItens
        Call CF("Carrega_Grid_Itens_PrecoDesc", objGridItens1, iIndice, iGrid_PrecoUnitario_Col + 1, Format(objOrcamentoVenda.colItens(iIndice).dPrecoUnitario * (1 - dPercDesc), gobjFAT.sFormatoPrecoUnitario), objOrcamentoVenda.colItens(iIndice).sProduto, iGrid_DescProduto_Col)
        
        GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dPercDesc, "Percent")
        GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objOrcamentoVenda.colItens(iIndice).dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objOrcamentoVenda.colItens(iIndice).dPrecoTotal, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col) = Format(objOrcamentoVenda.colItens(iIndice).dPrecoTotal + objOrcamentoVenda.colItens(iIndice).dValorDesconto, "Standard")
        If objOrcamentoVenda.colItens(iIndice).dtDataEntrega <> DATA_NULA Then GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col) = Format(objOrcamentoVenda.colItens(iIndice).dtDataEntrega, "dd/mm/yyyy")
        If objOrcamentoVenda.colItens(iIndice).iPrazoEntrega <> 0 Then GridItens.TextMatrix(iIndice, iGrid_PrazoEntrega_Col) = CStr(objOrcamentoVenda.colItens(iIndice).iPrazoEntrega)

        GridItens.TextMatrix(iIndice, iGrid_Observacao_Col) = objOrcamentoVenda.colItens(iIndice).sObservacao
        If objOrcamentoVenda.colItens(iIndice).lStatus <> 0 Then
            Call Combo_Seleciona_ItemData(StatusItem, objOrcamentoVenda.colItens(iIndice).lStatus)
            GridItens.TextMatrix(iIndice, iGrid_Status_Col) = StatusItem.Text
        End If
        If objOrcamentoVenda.colItens(iIndice).lMotivoPerda <> 0 Then
            Call Combo_Seleciona_ItemData(MotivoPerdaItem, objOrcamentoVenda.colItens(iIndice).lMotivoPerda)
            GridItens.TextMatrix(iIndice, iGrid_MotivoPerda_Col) = MotivoPerdaItem.Text
        End If
        
        lErro = CF2(Me, "OV_Traz_GridItens_Tela", objOrcamentoVenda.colItens(iIndice), iIndice)
        If lErro <> SUCESSO Then gError 141404
        
        If iGrid_SomaItem_Col <> 0 Then
             GridItens.TextMatrix(iIndice, iGrid_SomaItem_Col) = CStr(objOrcamentoVenda.colItens(iIndice).iMarcado)
        End If
        
    Next

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = objOrcamentoVenda.colItens.Count

    'Atualiza o grid para mostrar as checkboxes marcadas / desmarcadas
    Call Grid_Refresh_Checkbox(objGridItens)

    If iGrid_SomaItem_Col <> 0 Then
        Call SomaItem_Click
    End If

    Carrega_Grid_Itens = SUCESSO
    
    Exit Function

Erro_Carrega_Grid_Itens:

    Carrega_Grid_Itens = gErr

    Select Case gErr

        Case 84406
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objOrcamentoVenda.colItens(iIndice).sProduto)

        Case 141404 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157815)

    End Select

    Exit Function

End Function

Public Sub Cliente_Formata(lCliente As Long)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objTipoCliente As New ClassTipoCliente

On Error GoTo Erro_Cliente_Formata

    Cliente.Text = lCliente

    'Busca o Cliente no BD
    lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
    If lErro <> SUCESSO Then gError 84411 '56915

    lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 84412 '56916

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", Filial, colCodigoNome)

    'Se o Tipo estiver preenchido
    If objCliente.iTipo > 0 Then
        objTipoCliente.iCodigo = objCliente.iTipo
        'Lê o Tipo de Cliente
        lErro = CF("TipoCliente_Le", objTipoCliente)
        If lErro <> SUCESSO And lErro <> 19062 Then gError 126955
    End If

    'Guarda o valor do desconto do cliente
    If objCliente.dDesconto > 0 Then
        gdDesconto = objCliente.dDesconto
    ElseIf objTipoCliente.dDesconto > 0 Then
        gdDesconto = objTipoCliente.dDesconto
    Else
        gdDesconto = 0
    End If
    
    PercDescontoItens.Text = Format(gdDesconto * 100, "FIXED")
    Call PercDescontoItens_Validate(bSGECancelDummy)

    'para fazer valer o que veio do bd
    giValorDescontoManual = 1

    'Coloca clientes em atraso em vermelho
    If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)
        
    giClienteAlterado = 0
    
    Exit Sub

Erro_Cliente_Formata:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 84411, 84412, 126955

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157816)

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
    If lErro <> SUCESSO And lErro <> 17660 Then gError 84414 '56918

    If lErro = 17660 Then gError 84415 '56919

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_Filial_Formata:

    Select Case gErr

        Case 84414

        Case 84415
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157817)

    End Select

    Exit Sub

End Sub

Public Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iVendedorAlterado = 1


End Sub

Public Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    'Se Vendedor foi alterado,
    If iVendedorAlterado = 1 Then

        If Len(Trim(Vendedor.Text)) > 0 Then
            
            'Tenta ler o Vendedor (NomeReduzido ou Código)
            lErro = TP_Vendedor_Le(Vendedor, objVendedor)
            If lErro <> SUCESSO Then gError 94417

        End If

        iVendedorAlterado = 0

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 94417   'Tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157818)

    End Select
    
End Sub

Public Sub VendedorLabel_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Public Sub Vendedor2_Change()

    iAlterado = REGISTRO_ALTERADO
    iVendedor2Alterado = 1


End Sub

Public Sub Vendedor2_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor2_Validate

    'Se Vendedor foi alterado,
    If iVendedor2Alterado = 1 Then

        If Len(Trim(Vendedor2.Text)) > 0 Then
            
            'Tenta ler o Vendedor (NomeReduzido ou Código)
            lErro = TP_Vendedor_Le(Vendedor2, objVendedor)
            If lErro <> SUCESSO Then gError 94417

        End If

        iVendedor2Alterado = 0

    End If

    Exit Sub

Erro_Vendedor2_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 94417   'Tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157818)

    End Select
    
End Sub

Public Sub Vendedor2Label_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Vendedor2.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor2.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor2)

End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Vendedor Then
            Call VendedorLabel_Click
        ElseIf Me.ActiveControl Is Vendedor2 Then
            Call Vendedor2Label_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is CodigoBase Then
            Call NumeroBaseLabel_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        End If
          
    End If
    
    If KeyCode = KEYCODE_IDIOMAS Then
        Call gobjIdiomaTela.AbreTela(Me.ActiveControl)
    End If

End Sub

Private Function Tipo_Cliente(ByVal sCliente As String) As enumTipo

    If Len(Trim(sCliente)) = 0 Then
        Tipo_Cliente = TIPO_VAZIO
    ElseIf Not IsNumeric(sCliente) Then
        Tipo_Cliente = TIPO_STRING
    ElseIf Int(CDbl(sCliente)) <> CDbl(sCliente) Then
        Tipo_Cliente = TIPO_DECIMAL
    ElseIf CDbl(sCliente) <= 0 Then
        Tipo_Cliente = TIPO_NAO_POSITIVO
    ElseIf Len(Trim(sCliente)) > STRING_CGC Then
        Tipo_Cliente = TIPO_OVERFLOW
    ElseIf Len(Trim(sCliente)) > STRING_CPF Then
        Tipo_Cliente = TIPO_CGC
    ElseIf CDbl(sCliente) > NUM_MAX_CLIENTES Then
        Tipo_Cliente = TIPO_CPF
    Else
        Tipo_Cliente = TIPO_CODIGO
    End If

End Function

Public Function Nome_Extrai(sTexto As String) As String
'Função que retira de um texto no formato "Codigo - Nome" apenas o nome.

Dim iPosicao As Integer
Dim sString As String

    iPosicao = InStr(1, sTexto, "-")
    sString = Mid(sTexto, iPosicao + 1)
    
    Nome_Extrai = sString
    
    Exit Function

End Function

Public Function CargaPosFormLoad(Optional bTrazendoDoc As Boolean = False) As Long

Dim lErro As Long

On Error GoTo Erro_CargaPosFormLoad

    If (giPosCargaOk = 0) Then
    
        'p/permitir o redesenho da tela
        DoEvents
        
        gbCarregandoTela = True
               
        'Carrega a combo combo de Tabela de Preços
        lErro = Carrega_TabelaPreco()
        If lErro <> SUCESSO Then gError 26481
    
        'Carrega a combo de Condição de Pagamento
        lErro = CF("Carrega_CondicaoPagamento", CondicaoPagamento, MODULO_CONTASARECEBER)
        If lErro <> SUCESSO Then gError 26490
        
        'Carrega os Estados
        lErro = Carrega_PlacaUF()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Carrega a combo de Transportadora com as Transportadoras
        lErro = Carrega_Transportadoras()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Incluído por Luiz Nogueira em 21/08/03
        'Carrega a combo VolumeEspecie
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Incluído por Luiz Nogueira em 21/08/03
        'Carrega a combo VolumeMarca
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        'Carrega a Combo TipoFrete
        lErro = Carrega_TipoFrete()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Carrega a combo de Canais de Venda
        lErro = Carrega_CanalVenda()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Carrega_Moeda()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = Carrega_PrazoTexto()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        PrecoUnitario.Format = gobjFAT.sFormatoPrecoUnitario
        
        Quantidade.Format = FORMATO_ESTOQUE
    
        'Preenche Data Referencia e Data de Emissão coma Data Atual
        DataReferencia.PromptInclude = False
        DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
        DataReferencia.PromptInclude = True
        giDataReferenciaAlterada = 0
    
        Set objGridItens = New AdmGrid
        Set objGridParcelas = New AdmGrid
        Set objGridPrecosCalculados = New AdmGrid
        Set objGridFormacaoPreco = New AdmGrid
        
        Set objEventoCliente = New AdmEvento
        Set objEventoNumero = New AdmEvento
        Set objEventoCondPagto = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoVendedor = New AdmEvento
        Set objEventoVendedor2 = New AdmEvento
        Set objEventoNaturezaOp = New AdmEvento
        Set objEventoTiposDeTributacao = New AdmEvento
        Set objEventoTransportadora = New AdmEvento
        Set objEventoTranspRedespacho = New AdmEvento
        Set objEventoMensagem = New AdmEvento
        Set objEventoEstoqueProd = New AdmEvento

        Set objEventoBotaoPedAtualizados = New AdmEvento
    
        Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
        Set gobjTelaProjetoInfo.objUserControl = objUserControl
        Set gobjTelaProjetoInfo.objTela = Me

        Set gcolItensOVPrecoCalc = New Collection
        
        Call PCSituacoes_Carrega
        Call FPSituacoes_Carrega
           
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
        If lErro <> SUCESSO Then gError 26636
    
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", FPProduto)
        If lErro <> SUCESSO Then gError 26636
        
        'Faz as Inicializações dos Grids
        lErro = Inicializa_Grid_Itens(objGridItens)
        If lErro <> SUCESSO Then gError 26493
    
        lErro = Inicializa_Grid_Parcelas(objGridParcelas)
        If lErro <> SUCESSO Then gError 26496
    
        lErro = Inicializa_Grid_PrecosCalculados(objGridPrecosCalculados)
        If lErro <> SUCESSO Then gError 26493
    
        lErro = Inicializa_Grid_FormacaoPreco(objGridFormacaoPreco)
        If lErro <> SUCESSO Then gError 26493
        
        ''*** incluidos p/tratamento de tributacao *******************************
        Set gobjTribTab = New ClassTribTab
        lErro = gobjTribTab.Ativar(Me, , , gobjTribTab.TIPOTELA_OV)
        If lErro <> SUCESSO Then gError 26493
        
        Set gobjOrcamentoVenda = New ClassOrcamentoVenda
        gobjOrcamentoVenda.dtDataEmissao = StrParaDate(DataEmissao.Text)
        lErro = gobjTribTab.TributacaoNF_Reset(gobjOrcamentoVenda)
        If lErro <> SUCESSO Then gError 26493
        '*** fim tributacao
        
        Call OptDataPrazoEnt_Click(0)
    
        gbCarregandoTela = False
     
        iAlterado = 0
        iCodigoAlterado = 0

        giPosCargaOk = 1
        
        If gobjFAT.iExibeOPImpCodProd = MARCADO Then
            ImprimirOVComCodProd.Visible = True
        End If
        If gobjFAT.iExibeOPCalcST = MARCADO Then
            CalcularST.Visible = True
            If gobjFAT.iCalcSTPadrao = MARCADO Then
                CalcularST.Value = vbChecked
            Else
                CalcularST.Value = vbUnchecked
            End If
        Else
            CalcularST.Value = vbChecked
        End If
        If CalcularST.Value = vbChecked Then
            gobjOrcamentoVenda.iNaoCalcICMSST = DESMARCADO
        Else
            gobjOrcamentoVenda.iNaoCalcICMSST = MARCADO
        End If

        Call ValorTotal_Calcula
    
    End If
   
    CargaPosFormLoad = SUCESSO
    
    Exit Function
     
Erro_CargaPosFormLoad:

    gbCarregandoTela = False
    
    CargaPosFormLoad = gErr
    
    Select Case gErr
          '????? Trocar a numeracao de erro
        Case 46531, 26481, 26483, 26485, 26487, 26490, 26491
        Case 26493, 26636, 26495, 26496, 26497, 26635
        Case 27638, 46177, 27643, 96126
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157834)
     
    End Select
     
    Exit Function

End Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157835)

     End Select

     Exit Sub

End Sub

Public Sub BotaoImprimir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoImprimir_Click

    'Se o código do orçamento não foi informado => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 102238
    
    'Dispara função para imprimir orçamento
    lErro = Orcamento_Imprime(Trim(Codigo.Text))
    If lErro <> SUCESSO Then gError 102239
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 102239
        
        Case 102238
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 157836)

    End Select

    Exit Sub

End Sub

Private Function Orcamento_Imprime(ByVal lOrcamento As Long, Optional ByVal bGravando As Boolean = False) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim sNomeTsk As String, sConteudo1 As String, sConteudo2 As String, sConteudo3 As String, sCodRel As String

On Error GoTo Erro_Orcamento_Imprime

    'Transforma o ponteiro do mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Guarda no obj o código do orçamento passado como parâmetro
    objOrcamentoVenda.lCodigo = lOrcamento
    
    'Guarda a FilialEmpresa ativa como filial do orçamento
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
    objOrcamentoVenda.iVersao = StrParaInt(OVVersao.Caption)
    
    'Lê os dados do orçamento para verificar se o mesmo existe no BD
    lErro = CF("OrcamentoVenda_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101232 Then gError 102235

    'Se não encontrou => erro, pois não é possível imprimir um orçamento inexistente
    If lErro = 101232 Then gError 102236
    
    sNomeTsk = "OrcVenda"
    sCodRel = "Orçamento de Vendas"
    
    lErro = CF("Config_Le", "FATConfig", "TROCA_TSK_OV_POR_STATUS", 0, sConteudo1)
    If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(sConteudo1) = MARCADO Then
    
        lErro = CF("Config_Le", "FATConfig", "NOME_TSK_OV_POR_STATUS" & CStr(objOrcamentoVenda.lStatus), 0, sConteudo2)
        If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then
            sNomeTsk = sConteudo2
        End If
        
        lErro = CF("Config_Le", "FATConfig", "NOME_REL_OV_POR_STATUS" & CStr(objOrcamentoVenda.lStatus), 0, sConteudo3)
        If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then
            sCodRel = sConteudo3
        End If
    
    End If
    
    If Codigo_Extrai(Idioma.Text) <> 1 And sNomeTsk = "OrcVenda" Then
        sNomeTsk = "OV_ID" & Codigo_Extrai(Idioma.Text)
    End If
    
    'Dispara a impressão do relatório
    If objOrcamentoVenda.iHistorico <> MARCADO Or bGravando Then 'Ao gravar imprime sempre a versão atual
        lErro = objRelatorio.ExecutarDireto(sCodRel, "OrcamentoVenda >= @NORCVENDINIC E OrcamentoVenda <= @NORCVENDFIM", 1, sNomeTsk, "NORCVENDINIC", Trim(Codigo.Text), "NORCVENDFIM", Trim(Codigo.Text), "NEXIBIRORC", IIf(ImprimirOVComPreco.Value = vbChecked, 1, 0), "NEXIBIRCODPROD", IIf(ImprimirOVComCodProd.Value = vbChecked, 1, 0))
    Else
        lErro = objRelatorio.ExecutarDireto("Orçamento de Vendas Histórico", "OrcamentoVenda >= @NORCVENDINIC E OrcamentoVenda <= @NORCVENDFIM E Versao = @NVERSAO", 1, "ORCVENHI", "NORCVENDINIC", Trim(Codigo.Text), "NORCVENDFIM", Trim(Codigo.Text), "NEXIBIRORC", IIf(ImprimirOVComPreco.Value = vbChecked, 1, 0), "NEXIBIRCODPROD", IIf(ImprimirOVComCodProd.Value = vbChecked, 1, 0), "NVERSAO", OVVersao.Caption)
    End If
    If lErro <> SUCESSO Then gError 102237

    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault
    
    Orcamento_Imprime = SUCESSO
    
    Exit Function

Erro_Orcamento_Imprime:

    Orcamento_Imprime = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 102235, 102237
        
        Case 102236
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO", gErr, objOrcamentoVenda.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157837)
    
    End Select
    
    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault

End Function

Public Sub BotaoGrade_Click()

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim lErro  As Long
Dim objRomaneioGrade As ClassRomaneioGrade
Dim objItemOV As ClassItemOV

On Error GoTo Erro_BotaoGrade_Click

    'Se a linha selecionada for uma existente
    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
    
        'Recolhe o item  corr
        Set objItemOV = gobjOrcamentoVenda.colItens(GridItens.Row)
        
        If objItemOV.iPossuiGrade = MARCADO Then
            
            Set objRomaneioGrade = New ClassRomaneioGrade
            
            objRomaneioGrade.sNomeTela = Me.Name
            Set objRomaneioGrade.objObjetoTela = objItemOV
                        
            Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
            If lErro <> SUCESSO Then gError 86360
        
            Call Atualiza_Grid_Itens(objItemOV)
                    
            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))) > 0 Then
                
                Call PrecoTotal_Calcula(GridItens.Row)
                
                lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
                If lErro <> SUCESSO Then gError 86360
                
                lErro = ValorTotal_Calcula()
                If lErro <> SUCESSO Then gError 84162
            End If
        
        End If
    
    End If
    
    Exit Sub

Erro_BotaoGrade_Click:

    Select Case gErr
    
        Case 84162, 86360
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157838)
            
    End Select
    
    Exit Sub

End Sub

Sub Atualiza_Grid_Itens(objItemOV As ClassItemOV)

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim dQuantidade As Double
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
    
    For Each objItemRomaneioGrade In objItemOV.colItensRomaneioGrade
            
        dQuantidade = dQuantidade + objItemRomaneioGrade.dQuantidade
        
    Next

    GridItens.TextMatrix(objItemOV.iItem, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)

    objItemOV.dQuantidade = dQuantidade
    
    Exit Sub

End Sub

Function Grid_Possui_Grade() As Boolean

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim iIndice As Integer

    For iIndice = 1 To gobjOrcamentoVenda.colItens.Count
        If gobjOrcamentoVenda.colItens(iIndice).iPossuiGrade = MARCADO Then
            Grid_Possui_Grade = True
            Exit Function
        End If
    Next
    
    Grid_Possui_Grade = False
        
    Exit Function
    
End Function

Function Move_ItensGrade_Tela(colItensRomaneio As Collection, colItensRomaneioTela As Collection) As Long

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradeTela As ClassItemRomaneioGrade

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
                            
        colItensRomaneio.Add objItemRomaneioGrade
    Next

    Exit Function

End Function

Public Sub VersaoKit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VersaoKit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub VersaoKit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub VersaoKit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = VersaoKit
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub VersaoKitBase_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VersaoKitBase_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub VersaoKitBase_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub VersaoKitBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = VersaoKitBase
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_VersaoKitBase(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit base que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_VersaoKitBase

    Set objGridInt.objControle = VersaoKitBase

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_VersaoKitBase = SUCESSO

    Exit Function

Erro_Saida_Celula_VersaoKitBase:

    Saida_Celula_VersaoKitBase = gErr

    Select Case gErr

        Case 129217
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157842)

    End Select

End Function

Private Function Saida_Celula_VersaoKit(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Versão do Kit que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_VersaoKit

    Set objGridInt.objControle = VersaoKit

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129217

    Saida_Celula_VersaoKit = SUCESSO

    Exit Function

Erro_Saida_Celula_VersaoKit:

    Saida_Celula_VersaoKit = gErr

    Select Case gErr

        Case 129217
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157843)

    End Select

End Function

Private Sub objEventoVersao_evSelecao(obj1 As Object)

Dim objKit As ClassKit
Dim lErro As Long

On Error GoTo Erro_objEventoVersao_evSelecao

    Set objKit = obj1

    VersaoKitBase.Text = objKit.sVersao
    
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        GridItens.TextMatrix(GridItens.Row, iGrid_VersaoKitBase_Col) = VersaoKitBase.Text
    End If
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoVersao_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157844)

    End Select

    Exit Sub
    
End Sub

Private Function Atualiza_VersaoKit_Grid() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Atualiza_VersaoKit_Grid
    
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
    
        For iIndice = 1 To objGridItens.iLinhasExistentes
        
            GridItens.TextMatrix(iIndice, iGrid_VersaoKit_Col) = Codigo.Text & SEPARADOR & iIndice
        
        Next
    
    End If

    Atualiza_VersaoKit_Grid = SUCESSO

    Exit Function

Erro_Atualiza_VersaoKit_Grid:

    Atualiza_VersaoKit_Grid = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157845)

    End Select

    Exit Function
    
End Function

Public Sub NumeroBaseLabel_Click()

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim colSelecao As Collection

On Error GoTo Erro_NumeroLabel_Click
    
    objOrcamentoVenda.lCodigo = StrParaLong(CodigoBase.Text)
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa

    Call Chama_Tela("OrcamentoVendaLista", colSelecao, objOrcamentoVenda, objEventoNumeroBase)
    
    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157846)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumeroBase_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOrcamentoVenda As ClassOrcamentoVenda

On Error GoTo Erro_objEventoNumeroBase_evSelecao

    Set objOrcamentoVenda = obj1

    lErro = Traz_OrcamentoVenda_Tela(objOrcamentoVenda, VALOR_CODIGOBASE)
    If lErro <> SUCESSO And lErro <> 84363 Then gError 129972

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNumeroBase_evSelecao:

    Select Case gErr

        Case 129972
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157847)

    End Select

    Exit Sub

End Sub

Public Sub CodigoBase_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub CodigoBase_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodigoBase, iAlterado)

End Sub

Private Function Orcamento_Email(ByVal lOrcamento As Long) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim sMailTo As String, objFilialCliente As New ClassFilialCliente, objEndereco As New ClassEndereco
Dim sNomeTsk As String, sConteudo1 As String, sConteudo2 As String, sConteudo3 As String, sCodRel As String

On Error GoTo Erro_Orcamento_Email

    'Transforma o ponteiro do mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Guarda no obj o código do orçamento passado como parâmetro
    objOrcamentoVenda.lCodigo = lOrcamento
    
    'Guarda a FilialEmpresa ativa como filial do orçamento
    objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
    objOrcamentoVenda.iVersao = StrParaInt(OVVersao.Caption)
    
    'Lê os dados do orçamento para verificar se o mesmo existe no BD
    lErro = CF("OrcamentoVenda_Le", objOrcamentoVenda)
    If lErro <> SUCESSO And lErro <> 101232 Then gError 102235

    'Se não encontrou => erro, pois não é possível imprimir um orçamento inexistente
    If lErro = 101232 Then gError 102236
    
    If Len(Trim(objOrcamentoVenda.sEmail)) <> 0 Then
    
        sMailTo = Trim(objOrcamentoVenda.sEmail)
    
    Else
    
        objFilialCliente.lCodCliente = objOrcamentoVenda.lCliente
        objFilialCliente.iCodFilial = objOrcamentoVenda.iFilial
        
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
    
    End If
    
    sNomeTsk = "OrcVenda"
    sCodRel = "Orçamento de Vendas"
    
    lErro = CF("Config_Le", "FATConfig", "TROCA_TSK_OV_POR_STATUS", 0, sConteudo1)
    If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
    
    If StrParaInt(sConteudo1) = MARCADO Then
    
        lErro = CF("Config_Le", "FATConfig", "NOME_TSK_OV_POR_STATUS" & CStr(objOrcamentoVenda.lStatus), 0, sConteudo2)
        If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then
            sNomeTsk = sConteudo2
        End If
        
        lErro = CF("Config_Le", "FATConfig", "NOME_REL_OV_POR_STATUS" & CStr(objOrcamentoVenda.lStatus), 0, sConteudo3)
        If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO Then
            sCodRel = sConteudo3
        End If
    
    End If
    
    'Dispara a impressão do relatório
    If objOrcamentoVenda.iHistorico <> MARCADO Then
        lErro = objRelatorio.ExecutarDiretoEmail(sCodRel, "OrcamentoVenda >= @NORCVENDINIC E OrcamentoVenda <= @NORCVENDFIM", 0, sNomeTsk, "NORCVENDINIC", Trim(Codigo.Text), "NORCVENDFIM", Trim(Codigo.Text), "NEXIBIRORC", IIf(ImprimirOVComPreco.Value = vbChecked, 1, 0), "TTO_EMAIL", sMailTo, "TSUBJECT", "Orçamento " & CStr(lOrcamento), "TALIASATTACH", "Orc" & CStr(lOrcamento))
    Else
        lErro = objRelatorio.ExecutarDiretoEmail("Orçamento de Vendas Histórico", "OrcamentoVenda >= @NORCVENDINIC E OrcamentoVenda <= @NORCVENDFIM E Versao = @NVERSAO", 0, "ORCVENHI", "NORCVENDINIC", Trim(Codigo.Text), "NORCVENDFIM", Trim(Codigo.Text), "NEXIBIRORC", IIf(ImprimirOVComPreco.Value = vbChecked, 1, 0), "TTO_EMAIL", sMailTo, "TSUBJECT", "Orçamento " & CStr(lOrcamento), "TALIASATTACH", "Orc" & CStr(lOrcamento) & "V" & Format(OVVersao.Caption, "000"), "NVERSAO", OVVersao.Caption)
    End If
    If lErro <> SUCESSO Then gError 102237
    
    objOrcamentoVenda.dtDataEnvio = gdtDataAtual
    
    lErro = CF("OrcamentoVenda_Atualiza_Envio", objOrcamentoVenda)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call DateParaMasked(DataEnvio, objOrcamentoVenda.dtDataEnvio)

    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault
    
    Orcamento_Email = SUCESSO
    
    Exit Function

Erro_Orcamento_Email:

    Orcamento_Email = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 102235, 124017, 124018
        
        Case 102237 'Esse erro deve ser tratado na função chamadora, pois erro na impressão não deve interferir no funcionamento da tela
        
        Case 102236
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO", gErr, objOrcamentoVenda.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157848)
    
    End Select
    
    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault

End Function

Public Sub BotaoEmail_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoEmail_Click

    'Se o código do orçamento não foi informado => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 102238
    
    'Dispara função para imprimir orçamento
    lErro = Orcamento_Email(Trim(Codigo.Text))
    If lErro <> SUCESSO And lErro <> 102237 Then gError 102239
    
    Exit Sub

Erro_BotaoEmail_Click:

    Select Case gErr

        Case 102239
        
        Case 102238
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 157849)

    End Select

    Exit Sub

End Sub

Public Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134009

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134009

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157850)

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

Public Function OV_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

Dim iIncremento As Integer
Dim objUserControl As Object

    Set objGridInt.objForm = Me
    Set objUserControl = Me

    'Títulos das colunas
    If gobjFAT.iOVExibeOrdItens = MARCADO Then
        objGridInt.colColuna.Add (" ")
    Else
        objGridInt.colColuna.Add ("Item")
    End If
    
    If gobjFAT.iOVExibeTS = MARCADO Then
        objGridInt.colColuna.Add ("S")
    End If
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objGridInt.colColuna.Add ("Versão do Kit")
        objGridInt.colColuna.Add ("Versao do Kit Base")
    End If
    
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens4", objGridInt)
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens1", objGridInt)
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Prazo Entrega")
    objGridInt.colColuna.Add ("Data Entrega")
    
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Motivo Perda")
    objGridInt.colColuna.Add ("Observação")

    'Controles que participam do Grid
    If gobjFAT.iOVExibeTS = MARCADO Then
        objGridInt.colCampo.Add (SomaItem.Name)
    End If
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoProduto.Name)
    
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        objGridInt.colCampo.Add (VersaoKit.Name)
        objGridInt.colCampo.Add (VersaoKitBase.Name)
        
    Else
    
        VersaoKit.left = POSICAO_FORA_TELA
        VersaoKit.TabStop = False
    
        VersaoKitBase.left = POSICAO_FORA_TELA
        VersaoKitBase.TabStop = False

    End If

    Call CF("Orcamento_Inicializa_Grid_Itens5", objGridInt, objUserControl)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    'precodesc
    Call CF("Orcamento_Inicializa_Grid_Itens2", objGridInt, objUserControl)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    objGridInt.colCampo.Add (PrecoTotalB.Name)
    objGridInt.colCampo.Add (PrazoEntregaItem.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)
    
    objGridInt.colCampo.Add (StatusItem.Name)
    objGridInt.colCampo.Add (MotivoPerdaItem.Name)
    objGridInt.colCampo.Add (Observacao.Name)
    
    'Colunas do Grid
    iIncremento = 0
    iGrid_ItemProduto_Col = 0
    If gobjFAT.iOVExibeTS = MARCADO Then
        iGrid_SomaItem_Col = 1
        iIncremento = iIncremento + 1
    End If
    iGrid_Produto_Col = 1 + iIncremento
    iGrid_DescProduto_Col = 2 + iIncremento
    
    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
        iGrid_VersaoKit_Col = 3 + iIncremento
        iGrid_VersaoKitBase_Col = 4 + iIncremento
        iIncremento = iIncremento + 2
        
    End If
    
    Call CF("Orcamento_Inicializa_Grid_Itens6", iIncremento)
    iGrid_UnidadeMed_Col = 3 + iIncremento
    iGrid_Quantidade_Col = 4 + iIncremento
    iGrid_PrecoUnitario_Col = 5 + iIncremento
    Call CF("Orcamento_Inicializa_Grid_Itens3", iIncremento)
    iGrid_PercDesc_Col = 6 + iIncremento
    iGrid_Desconto_Col = 7 + iIncremento
    iGrid_PrecoTotal_Col = 8 + iIncremento
    iGrid_PrecoTotalB_Col = 9 + iIncremento
    iGrid_PrazoEntrega_Col = 10 + iIncremento
    iGrid_DataEntrega_Col = 11 + iIncremento
    
    iGrid_Status_Col = 12 + iIncremento
    iGrid_MotivoPerda_Col = 13 + iIncremento
    iGrid_Observacao_Col = 14 + iIncremento
    
    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    If m_objUserControl.Height > 9000 Then
        If gobjFAT.iOVExibeTS = MARCADO Then
            objGridInt.iLinhasVisiveis = 15
        Else
            objGridInt.iLinhasVisiveis = 16
        End If
    Else
        If gobjFAT.iOVExibeTS = MARCADO Then
            objGridInt.iLinhasVisiveis = 7
        Else
            objGridInt.iLinhasVisiveis = 8
        End If
    End If
    
    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    'objGridInt.iProibidoIncluirNoMeioGrid = GRID_PERMITIDO_INCLUIR_NO_MEIO

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    OV_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Sub StatusComercial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub StatusComercial_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Status_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MotivoPerda_Change()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Perda
End Sub

Public Sub Status_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MotivoPerda_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Perda
End Sub

Private Function Carrega_Status(ByVal objComboBox As ComboBox) As Long
'Carrega a combo de Tipo

Dim lErro As Long

On Error GoTo Erro_Carrega_Status

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_STATUSOV, objComboBox)
    If lErro <> SUCESSO Then gError 141371

    objComboBox.AddItem ""
    objComboBox.ItemData(objComboBox.NewIndex) = 0

    Carrega_Status = SUCESSO

    Exit Function

Erro_Carrega_Status:

    Carrega_Status = gErr

    Select Case gErr
    
        Case 141371

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157851)

    End Select

    Exit Function

End Function

Private Function Carrega_Motivo(ByVal objComboBox As ComboBox) As Long
'Carrega a combo de Tipo

Dim lErro As Long

On Error GoTo Erro_Carrega_Motivo

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_MOTIVOSOV, objComboBox)
    If lErro <> SUCESSO Then gError 141372
    
    objComboBox.AddItem ""
    objComboBox.ItemData(objComboBox.NewIndex) = 0
    
    lMotivoPerdaAnt = objComboBox.ListIndex

    Carrega_Motivo = SUCESSO

    Exit Function

Erro_Carrega_Motivo:

    Carrega_Motivo = gErr

    Select Case gErr
    
        Case 141372

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157852)

    End Select

    Exit Function

End Function

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Extrai

    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_ORCVENDA
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157853)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157854)
     
    End Select
     
    Exit Sub

End Sub

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Observacao que está deixando de ser a corrente

Dim lErro As Long
Dim sObservacao As String

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = Observacao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 141386

    Saida_Celula_Observacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = gErr

    Select Case gErr

        Case 141386
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157855)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MotivoPerdaItem(objGridInt As AdmGrid) As Long
'Faz a crítica da célula MotivoPerdaItem que está deixando de ser a corrente

Dim lErro As Long
Dim sMotivoPerdaItem As String

On Error GoTo Erro_Saida_Celula_MotivoPerdaItem

    Set objGridInt.objControle = MotivoPerdaItem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 141387

    Saida_Celula_MotivoPerdaItem = SUCESSO

    Exit Function

Erro_Saida_Celula_MotivoPerdaItem:

    Saida_Celula_MotivoPerdaItem = gErr

    Select Case gErr

        Case 141387
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157856)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_StatusItem(objGridInt As AdmGrid) As Long
'Faz a crítica da célula StatusItem que está deixando de ser a corrente

Dim lErro As Long
Dim sStatusItem As String

On Error GoTo Erro_Saida_Celula_StatusItem

    Set objGridInt.objControle = StatusItem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 141388

    Saida_Celula_StatusItem = SUCESSO

    Exit Function

Erro_Saida_Celula_StatusItem:

    Saida_Celula_StatusItem = gErr

    Select Case gErr

        Case 141388
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157857)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Observacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Observacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Observacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Observacao
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub StatusItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub StatusItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub StatusItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub StatusItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = StatusItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub MotivoPerdaItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub MotivoPerdaItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub MotivoPerdaItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub MotivoPerdaItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = MotivoPerdaItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function OV_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    OV_Saida_Celula_GridItens = SUCESSO
End Function

Public Function OV_Move_GridItens_Memoria(ByVal objCT As Object, ByVal objItemOV As ClassItemOV, ByVal iLinha As Integer) As Long
    OV_Move_GridItens_Memoria = SUCESSO
End Function

Public Function OV_Traz_GridItens_Tela(ByVal objCT As Object, ByVal objItemOV As ClassItemOV, ByVal iLinha As Integer) As Long
    OV_Traz_GridItens_Tela = SUCESSO
End Function

Public Function OV_Form_Load(ByVal objCT As Object) As Long
    OV_Form_Load = SUCESSO
End Function

Public Function OV_Valida_Itens(ByVal objCT As Object, ByVal objOV As ClassOrcamentoVenda) As Long
    OV_Valida_Itens = SUCESSO
End Function

Public Function OV_Rotina_Grid_Enable(ByVal objCT As Object, ByVal sNomeControle As String, ByVal iProdutoPreenchido As Integer) As Long
    OV_Rotina_Grid_Enable = SUCESSO
End Function

Public Function OV_Produto_Valida(ByVal objCT As Object) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_OV_Produto_Valida

    If gobjCRFAT.iProdRepetidoOV = PRODUTO_REPETICAO_NAO_PERMITE Then
    
        'Verifica se já está em outra linha do Grid
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If iIndice <> GridItens.Row Then
                If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 141411
            End If
        Next

    End If
    
    Exit Function

Erro_OV_Produto_Valida:

    OV_Produto_Valida = gErr

    Select Case gErr
    
        Case 141411
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, Produto.Text, Produto.Text, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157858)

    End Select

    Exit Function
    
End Function

Public Function OV_Calcula_PesoKg(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    OV_Calcula_PesoKg = SUCESSO
End Function

Public Function OV_Calcula_PesoKg_Total(ByVal objCT As Object) As Long
    OV_Calcula_PesoKg_Total = SUCESSO
End Function

Public Function OV_PrecoTotal_Calcula(ByVal objCT As Object, ByVal iLinha As Integer, dQuantidade As Double) As Long
    OV_PrecoTotal_Calcula = SUCESSO
End Function

Public Function OV_Produto_Saida_Celula(ByVal objCT As Object, objProduto As ClassProduto) As Long
    OV_Produto_Saida_Celula = SUCESSO
End Function

Public Function OV_Valida_Grid_Itens(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    OV_Valida_Grid_Itens = SUCESSO
End Function

Public Function OV_Produto_Saida_Celula3(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim sCodigoAntigo As String

On Error GoTo Erro_OV_Produto_Saida_Celula3

    If gobjCRFAT.iProdRepetidoOV = PRODUTO_REPETICAO_NAO_PERMITE Then
            
        sCodigoAntigo = objProduto.sCodigo
    
        If objProduto.iGerencial = DESMARCADO Then
            For iIndice = 1 To objGridItens.iLinhasExistentes
                If iIndice <> GridItens.Row Then
                    If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then
                        If vbYes = Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_EQUIVALENTE", Produto.Text, iIndice) Then
        
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
    
    OV_Produto_Saida_Celula3 = SUCESSO
    
    Exit Function

Erro_OV_Produto_Saida_Celula3:

    OV_Produto_Saida_Celula3 = gErr

    Select Case gErr
    
        Case 131295, 131296, 131297, 131298
         
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177165)

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

        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_OV
        
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

Private Function Inicializa_Grid_PrecosCalculados(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("R$ Calculado")
    objGridInt.colColuna.Add ("R$ Escolhido")
    objGridInt.colColuna.Add ("Situação")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Qtde")
    objGridInt.colColuna.Add ("Total")

    objGridInt.colCampo.Add (PCSelecionado.Name)
    objGridInt.colCampo.Add (PCProduto.Name)
    objGridInt.colCampo.Add (PCDescProduto.Name)
    objGridInt.colCampo.Add (PCPrecoUnitCalc.Name)
    objGridInt.colCampo.Add (PCPrecoUnit.Name)
    objGridInt.colCampo.Add (PCSituacao.Name)
    objGridInt.colCampo.Add (PCUnidMed.Name)
    objGridInt.colCampo.Add (PCQtde.Name)
    objGridInt.colCampo.Add (PCPrecoTotal.Name)

    'Controles que participam do Grid
    iGrid_PCSelecionado_Col = 1
    iGrid_PCProduto_Col = 2
    iGrid_PCDescProduto_Col = 3
    iGrid_PCPrecoUnitCalc_Col = 4
    iGrid_PCPrecoUnit_Col = 5
    iGrid_PCSituacao_Col = 6
    iGrid_PCUnidMed_Col = 7
    iGrid_PCQtde_Col = 8
    iGrid_PCPrecoTotal_Col = 9

    'Grid do GridInterno
    objGridInt.objGrid = GridPrecosCalculados

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    If m_objUserControl.Height > 9000 Then
        objGridInt.iLinhasVisiveis = 8
    Else
        objGridInt.iLinhasVisiveis = 3
    End If

    'Largura da primeira coluna
    GridPrecosCalculados.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_PrecosCalculados = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_FormacaoPreco(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Qtde")
    objGridInt.colColuna.Add ("Custo Unit.")
    objGridInt.colColuna.Add ("% Margem")
    objGridInt.colColuna.Add ("Preço Unit.")
    objGridInt.colColuna.Add ("Total")
    objGridInt.colColuna.Add ("Situação")

    objGridInt.colCampo.Add (FPProduto.Name)
    objGridInt.colCampo.Add (FPDescricaoProduto.Name)
    objGridInt.colCampo.Add (FPUnidMed.Name)
    objGridInt.colCampo.Add (FPQtde.Name)
    objGridInt.colCampo.Add (FPCustoUnit.Name)
    objGridInt.colCampo.Add (FPPercentMargem.Name)
    objGridInt.colCampo.Add (FPPrecoUnit.Name)
    objGridInt.colCampo.Add (FPPrecoTotal.Name)
    objGridInt.colCampo.Add (FPSituacao.Name)

    'Controles que participam do Grid
    iGrid_FPProduto_Col = 1
    iGrid_FPDescricaoProduto_Col = 2
    iGrid_FPUnidMed_Col = 3
    iGrid_FPQtde_Col = 4
    iGrid_FPCustoUnit_Col = 5
    iGrid_FPPercentMargem_Col = 6
    iGrid_FPPrecoUnit_Col = 7
    iGrid_FPPrecoTotal_Col = 8
    iGrid_FPSituacao_Col = 9

    'Grid do GridInterno
    objGridInt.objGrid = GridFormacaoPreco

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = (NUM_MAXIMO_ITENS * 10) + 1

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    
    If m_objUserControl.Height > 9000 Then
        objGridInt.iLinhasVisiveis = 9
    Else
        objGridInt.iLinhasVisiveis = 4
    End If
    
    'Largura da primeira coluna
    GridFormacaoPreco.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_FormacaoPreco = SUCESSO

    Exit Function

End Function

Public Sub GridPrecosCalculados_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridPrecosCalculados, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPrecosCalculados, iAlterado)
    End If

End Sub

Public Sub GridPrecosCalculados_EnterCell()

     Call Grid_Entrada_Celula(objGridPrecosCalculados, iAlterado)

End Sub

Public Sub GridPrecosCalculados_GotFocus()

    Call Grid_Recebe_Foco(objGridPrecosCalculados)

End Sub

Public Sub GridPrecosCalculados_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridPrecosCalculados)

End Sub

Public Sub GridPrecosCalculados_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridPrecosCalculados, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridPrecosCalculados, iAlterado)
    End If

End Sub

Public Sub GridPrecosCalculados_LeaveCell()

    Call Saida_Celula(objGridPrecosCalculados)

End Sub

Public Sub GridPrecosCalculados_RowColChange()

    Call Grid_RowColChange(objGridPrecosCalculados)

End Sub

Public Sub GridPrecosCalculados_Scroll()

    Call Grid_Scroll(objGridPrecosCalculados)

End Sub

Public Sub GridPrecosCalculados_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridPrecosCalculados)

End Sub

Public Sub GridFormacaoPreco_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridFormacaoPreco, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFormacaoPreco, iAlterado)
    End If

End Sub

Public Sub GridFormacaoPreco_EnterCell()

     Call Grid_Entrada_Celula(objGridFormacaoPreco, iAlterado)

End Sub

Public Sub GridFormacaoPreco_GotFocus()

    Call Grid_Recebe_Foco(objGridFormacaoPreco)

End Sub

Public Sub GridFormacaoPreco_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFormacaoPreco)

End Sub

Public Sub GridFormacaoPreco_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFormacaoPreco, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFormacaoPreco, iAlterado)
    End If

End Sub

Public Sub GridFormacaoPreco_LeaveCell()

    Call Saida_Celula(objGridFormacaoPreco)

End Sub

Public Sub GridFormacaoPreco_RowColChange()

    Call Grid_RowColChange(objGridFormacaoPreco)

End Sub

Public Sub GridFormacaoPreco_Scroll()

    Call Grid_Scroll(objGridFormacaoPreco)

End Sub

Public Sub GridFormacaoPreco_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridFormacaoPreco)

End Sub

Public Sub PCPrecoUnit_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PCPrecoUnit_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPrecosCalculados)

End Sub

Public Sub PCPrecoUnit_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPrecosCalculados)

End Sub

Public Sub PCPrecoUnit_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPrecosCalculados.objControle = PCPrecoUnit
    lErro = Grid_Campo_Libera_Foco(objGridPrecosCalculados)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PCSituacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PCSituacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPrecosCalculados)

End Sub

Public Sub PCSituacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPrecosCalculados)

End Sub

Public Sub PCSituacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPrecosCalculados.objControle = PCSituacao
    lErro = Grid_Campo_Libera_Foco(objGridPrecosCalculados)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub FPPercentMargem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FPPercentMargem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFormacaoPreco)

End Sub

Public Sub FPPercentMargem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFormacaoPreco)

End Sub

Public Sub FPPercentMargem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFormacaoPreco.objControle = FPPercentMargem
    lErro = Grid_Campo_Libera_Foco(objGridFormacaoPreco)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub FPSituacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FPSituacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridFormacaoPreco)

End Sub

Public Sub FPSituacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFormacaoPreco)

End Sub

Public Sub FPSituacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridFormacaoPreco.objControle = FPSituacao
    lErro = Grid_Campo_Libera_Foco(objGridFormacaoPreco)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PCSelecionado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PCSelecionado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridPrecosCalculados)

End Sub

Public Sub PCSelecionado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridPrecosCalculados)

End Sub

Public Sub PCSelecionado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridPrecosCalculados.objControle = PCSelecionado
    lErro = Grid_Campo_Libera_Foco(objGridPrecosCalculados)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FormPreco_Remove_Linha(ByVal iLinha As Integer)
    
    If gobjFAT.iOVSemFormPreco = 0 Then
    
        If gcolItensOVPrecoCalc.Count >= iLinha Then
        
            gcolItensOVPrecoCalc.Remove (iLinha)
            
        End If
    
        Call FormPreco_AtualizarSituacao
    
    End If
    
End Sub

Public Function OV_FormPreco_ValidaDados(ByVal objCT As Object, ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
    OV_FormPreco_ValidaDados = SUCESSO
End Function

Private Sub FormPreco_Adiciona_Linha(ByVal iLinha As Integer, ByVal objProduto As ClassProduto)

Dim objItemOV As New ClassItemOV

    If gobjFAT.iOVSemFormPreco = 0 Then
    
        'adiciona item em gcolItensOVPrecoCalc
        objItemOV.sProduto = objProduto.sCodigo
        
        If gobjFAT.iCotarItensOrcVenda <> 0 Then
            objItemOV.iPCSituacao = ITEMOV_PCSITUACAO_COTAR
        End If
        
        gcolItensOVPrecoCalc.Add objItemOV
        
        Call FormPreco_AtualizarSituacao
    
    End If
    
End Sub

Private Function FormPreco_Move_Tela_Memoria(ByVal objOrcamentoVenda As ClassOrcamentoVenda) As Long
'transfere dados

Dim lErro As Long, iLinhaFP As Integer
Dim objItemOV As ClassItemOV, iIndice As Integer, objItemOVAux As ClassItemOV
Dim objItemFP As ClassItemFormPreco, objItemFPAux As ClassItemFormPreco, iSeq As Integer

On Error GoTo Erro_FormPreco_Move_Tela_Memoria

    If gobjFAT.iOVSemFormPreco = 0 Then
    
        For Each objItemOV In objOrcamentoVenda.colItens
        
            iIndice = iIndice + 1
            objItemOV.dPCPrecoUnitCalc = StrParaDbl(GridPrecosCalculados.TextMatrix(iIndice, iGrid_PCPrecoUnitCalc_Col))
            
            Select Case GridPrecosCalculados.TextMatrix(iIndice, iGrid_PCSituacao_Col)
            
                Case ITEMOV_PCSITUACAO_STRING_COTAR
                    objItemOV.iPCSituacao = ITEMOV_PCSITUACAO_COTAR
                Case ITEMOV_PCSITUACAO_STRING_EM_COTACAO
                    objItemOV.iPCSituacao = ITEMOV_PCSITUACAO_EM_COTACAO
                Case ITEMOV_PCSITUACAO_STRING_COTADO
                    objItemOV.iPCSituacao = ITEMOV_PCSITUACAO_COTADO
                Case Else
                    objItemOV.iPCSituacao = ITEMOV_PCSITUACAO_NAO_COTAR
                
            End Select
            
            Set objItemOV.colItensFormPreco = New Collection
            
            If gcolItensOVPrecoCalc.Count >= iIndice Then
            
                Set objItemOVAux = gcolItensOVPrecoCalc.Item(iIndice)
                iSeq = 0
                For Each objItemFPAux In objItemOVAux.colItensFormPreco
                    Set objItemFP = New ClassItemFormPreco
                    Call objItemFP.Copiar(objItemFPAux)
                    
                    iSeq = iSeq + 1
                    objItemFP.iSequencial = iSeq
                    objItemOV.colItensFormPreco.Add objItemFP
                Next
                
            End If
            
        Next
    
    End If
    
    FormPreco_Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_FormPreco_Move_Tela_Memoria:

    FormPreco_Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184311)

    End Select
    
    Exit Function

End Function

Private Function FormPreco_PreencheGrid(ByVal iIndice As Integer) As Long

Dim lErro As Long, sSituacao As String
Dim iLinhaFP As Integer, objItemFP As ClassItemFormPreco, objItemOV As ClassItemOV
Dim sProdutoEnxuto As String, objProduto As New ClassProduto

On Error GoTo Erro_FormPreco_PreencheGrid
    
    If gobjFAT.iOVSemFormPreco = 0 Then
    
        Call Grid_Limpa(objGridFormacaoPreco)
            
        If gcolItensOVPrecoCalc.Count >= iIndice And iIndice <> 0 Then
        
            Set objItemOV = gcolItensOVPrecoCalc.Item(iIndice)
    
            For Each objItemFP In objItemOV.colItensFormPreco
            
                iLinhaFP = iLinhaFP + 1
                
                lErro = Mascara_RetornaProdutoEnxuto(objItemFP.sProduto, sProdutoEnxuto)
                If lErro <> SUCESSO Then gError 184323
        
                'Mascara o produto enxuto
                FPProduto.PromptInclude = False
                FPProduto.Text = sProdutoEnxuto
                FPProduto.PromptInclude = True
                
                objProduto.sCodigo = objItemFP.sProduto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 184324
                If lErro <> SUCESSO Then gError 184325
                
                GridFormacaoPreco.TextMatrix(iLinhaFP, iGrid_FPProduto_Col) = FPProduto.Text
                GridFormacaoPreco.TextMatrix(iLinhaFP, iGrid_FPDescricaoProduto_Col) = objProduto.sDescricao
                GridFormacaoPreco.TextMatrix(iLinhaFP, iGrid_FPUnidMed_Col) = objItemFP.sUnidMed
                GridFormacaoPreco.TextMatrix(iLinhaFP, iGrid_FPQtde_Col) = Formata_Estoque(objItemFP.dQtde)
                GridFormacaoPreco.TextMatrix(iLinhaFP, iGrid_FPCustoUnit_Col) = Format(objItemFP.dCustoUnit, gobjFAT.sFormatoPrecoUnitario)
                GridFormacaoPreco.TextMatrix(iLinhaFP, iGrid_FPPercentMargem_Col) = Format(objItemFP.dPercentMargem, "Percent")
                GridFormacaoPreco.TextMatrix(iLinhaFP, iGrid_FPPrecoUnit_Col) = Format(objItemFP.dPrecoUnit, gobjFAT.sFormatoPrecoUnitario)
                GridFormacaoPreco.TextMatrix(iLinhaFP, iGrid_FPPrecoTotal_Col) = Format(objItemFP.dPrecoTotal, "Standard")
                
                Select Case objItemFP.iSituacao
                    Case FPSITUACAO_NAO_COTAR
                        sSituacao = FPSITUACAO_STRING_NAO_COTAR
                    Case FPSITUACAO_COTAR
                        sSituacao = FPSITUACAO_STRING_COTAR
                    Case FPSITUACAO_EM_COTACAO
                        sSituacao = FPSITUACAO_STRING_EM_COTACAO
                    Case FPSITUACAO_COTADO
                        sSituacao = FPSITUACAO_STRING_COTADO
                End Select
                GridFormacaoPreco.TextMatrix(iLinhaFP, iGrid_FPSituacao_Col) = sSituacao
                
            Next
            
            objGridFormacaoPreco.iLinhasExistentes = iLinhaFP
        
        End If
    
    End If
    
    FormPreco_PreencheGrid = SUCESSO
    
    Exit Function
    
Erro_FormPreco_PreencheGrid:

    FormPreco_PreencheGrid = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184310)

    End Select
    
    Exit Function

End Function

Private Function PrecosCalculados_Traz_Dados_Tela() As Long

Dim lErro As Long
Dim objItemOV As ClassItemOV, iIndice As Integer
Dim objProduto As ClassProduto

On Error GoTo Erro_PrecosCalculados_Traz_Dados_Tela

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridPrecosCalculados)
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        If gcolItensOVPrecoCalc.Count >= iIndice Then
        
            Set objItemOV = gcolItensOVPrecoCalc.Item(iIndice)
            
            GridPrecosCalculados.TextMatrix(iIndice, iGrid_PCProduto_Col) = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)
            GridPrecosCalculados.TextMatrix(iIndice, iGrid_PCDescProduto_Col) = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
            GridPrecosCalculados.TextMatrix(iIndice, iGrid_PCUnidMed_Col) = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
            GridPrecosCalculados.TextMatrix(iIndice, iGrid_PCQtde_Col) = GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)
            GridPrecosCalculados.TextMatrix(iIndice, iGrid_PCPrecoUnit_Col) = GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col)
            GridPrecosCalculados.TextMatrix(iIndice, iGrid_PCPrecoTotal_Col) = GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col)
            
            GridPrecosCalculados.TextMatrix(iIndice, iGrid_PCPrecoUnitCalc_Col) = Format(objItemOV.dPCPrecoUnitCalc, gobjFAT.sFormatoPrecoUnitario)
            GridPrecosCalculados.TextMatrix(iIndice, iGrid_PCSituacao_Col) = gcolPCSituacoes.Item(objItemOV.iPCSituacao + 1)
            
        End If
        
    Next
    
    objGridPrecosCalculados.iLinhasExistentes = objGridItens.iLinhasExistentes
    
    Call Grid_Refresh_Checkbox(objGridPrecosCalculados)
    
    Call Grid_Limpa(objGridFormacaoPreco)

    PrecosCalculados_Traz_Dados_Tela = SUCESSO
    
    Exit Function
    
Erro_PrecosCalculados_Traz_Dados_Tela:

    PrecosCalculados_Traz_Dados_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184309)

    End Select
    
    Exit Function

End Function

Function FormPreco_Calcula(Optional ByVal iLinhaCalc As Integer = 0) As Long
    
Dim lErro As Long, bAtualizar As Boolean, dMargem As Double
Dim objItemOV As ClassItemOV, iIndice As Integer, objItemFP As ClassItemFormPreco
Dim sUnidadeMed As String, dQuantidade As Double, sVersaoKit As String, sVersaoKitBase As String
Dim colItensFormPreco As Collection, colMP As Collection, colEmb As Collection
Dim objProdutoFilial As ClassProdutoFilial, objKit As ClassKit, dQtdeBaseKit As Double, dFator As Double
Dim objCustoDirFabrPlanMP As ClassCustoDirFabrPlanMP, objCustoDirFabrPlanEmb As ClassCustoDirFabrPlanEmb
Dim objProduto As ClassProduto, objProdutoAux As ClassProduto, dQtdeAux As Double
Dim objItemFPAux As ClassItemFormPreco, bRecarregarGrids As Boolean, dCusto As Double

On Error GoTo Erro_FormPreco_Calcula

    If gobjFAT.iOVSemFormPreco = 0 Then
    
        bRecarregarGrids = False
        
        For iIndice = 1 To objGridItens.iLinhasExistentes
        
            If iLinhaCalc = 0 Or iLinhaCalc = iIndice Then
            
                If gcolItensOVPrecoCalc.Count >= iIndice Then
                
                    Set objItemOV = gcolItensOVPrecoCalc.Item(iIndice)
                    
                    Set objProduto = New ClassProduto
                    objProduto.sCodigo = objItemOV.sProduto
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 184317
                    If lErro <> SUCESSO Then gError 184318
        
                    If objProduto.iKitVendaComp = MARCADO Then
                        FormPreco_Calcula = SUCESSO
                        Exit Function
                    End If
                               
                    'Armazena os dados do item
                    sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
                    dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
                    
                    If gobjFAT.iTemVersaoOV = TEM_VERSAO_OV Then
                        sVersaoKit = GridItens.TextMatrix(iIndice, iGrid_VersaoKit_Col)
                        sVersaoKitBase = GridItens.TextMatrix(iIndice, iGrid_VersaoKitBase_Col)
                    End If
            
        '            If objItemOV.sUnidadeMed <> sUnidadeMed Or Abs(objItemOV.dQuantidade - dQuantidade) > QTDE_ESTOQUE_DELTA Or _
                       objItemOV.sVersaoKit <> sVersaoKit Or objItemOV.sVersaoKitBase <> sVersaoKitBase Then
                        
                        Set objProdutoFilial = New ClassProdutoFilial
                        objProdutoFilial.iFilialEmpresa = giFilialEmpresa
                        objProdutoFilial.sProduto = objItemOV.sProduto
            
                        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
                        If lErro <> SUCESSO And lErro <> 28261 Then gError 184312
                        'Se não encontrou
                        If lErro = 28261 Then gError 184313
                        
                        'criar colecao temporaria
                        Set colItensFormPreco = New Collection
                                                        
                        bAtualizar = True
                            
                        'se o produto nao é produzido na filial entao pegar apenas o custo dele mesmo como materia prima
                        If objProdutoFilial.iProdNaFilial = 0 Then
        
                            If dQuantidade <> 0 And Len(Trim(sUnidadeMed)) <> 0 Then
                            
                                Set objItemFP = New ClassItemFormPreco
                                objItemFP.sProduto = objItemOV.sProduto
                                
                                Set objProdutoAux = New ClassProduto
                                objProdutoAux.sCodigo = objItemFP.sProduto
                                lErro = CF("Produto_Le", objProdutoAux)
                                If lErro <> SUCESSO And lErro <> 28030 Then gError 184317
                                If lErro <> SUCESSO Then gError 184318
                                
                                dQtdeAux = dQuantidade
                                
                                If sUnidadeMed <> objProdutoAux.sSiglaUMCompra Then
                                    
                                    'converter qtde para UM de compra
                                    lErro = CF("UM_Conversao", objProdutoAux.iClasseUM, sUnidadeMed, objProdutoAux.sSiglaUMCompra, dFator)
                                    If lErro <> SUCESSO Then gError 184376
                                    
                                    dQtdeAux = dQtdeAux * dFator
                                
                                End If
                                
                                objItemFP.sUnidMed = objProdutoAux.sSiglaUMCompra
                                objItemFP.dQtde = dQtdeAux
                                
                                If gobjFAT.iCotarItensOrcVenda <> 0 Then
                                    objItemFP.iSituacao = FPSITUACAO_COTAR
                                End If
                                
                                lErro = CF("Produto_ObtemMargemPadrao", objItemFP.sProduto, dMargem)
                                If lErro <> SUCESSO Then gError 184376
                                objItemFP.dPercentMargem = dMargem
                                
                                '###Obter custo da matéria prima
                                dCusto = 0
                                lErro = CF("CustoMedioAtual_Le", objProduto.sCodigo, dCusto, giFilialEmpresa)
                                If lErro <> SUCESSO Then gError 184376
                                objItemFP.dCustoUnit = dCusto
                                
                                objItemFP.dPrecoUnit = Arredonda_Moeda(objItemFP.dCustoUnit * (1 + objItemFP.dPercentMargem))
                                objItemFP.dPrecoTotal = Arredonda_Moeda(objItemFP.dPrecoUnit * objItemFP.dQtde)
                                objItemOV.dPCPrecoUnitCalc = objItemFP.dPrecoUnit
                                    
                                colItensFormPreco.Add objItemFP
                                '###
                        
                            End If
                            
                        Else
                        
                            If dQuantidade <> 0 And Len(Trim(sUnidadeMed)) <> 0 Then
                            
                                'obter versao a ser utilizada
                                Set objKit = New ClassKit
                                objKit.sProdutoRaiz = objItemOV.sProduto
                                objKit.sVersao = sVersaoKit
                                
                                lErro = CF("Kit_Le", objKit)
                                If lErro <> SUCESSO And lErro <> 21826 Then gError 184314
                                If lErro <> SUCESSO Then
                                
                                    objKit.sVersao = sVersaoKitBase
                                    lErro = CF("Kit_Le", objKit)
                                    If lErro <> SUCESSO And lErro <> 21826 Then gError 184315
                                    If lErro <> SUCESSO Then
                                    
                                        lErro = CF("Kit_Le_FormPreco", objKit)
                                        If lErro <> SUCESSO And lErro <> 106304 Then gError 184316
                            
                                    End If
                                    
                                End If
                                
                                If lErro = SUCESSO Then
                                    
                                    '??? converter qtde da tela para a do kit
                                    dQtdeBaseKit = dQuantidade
                                    
                                    'ler do bd a colecao de insumos e embalagens
                                    Set colMP = New Collection
                                    Set colEmb = New Collection
                                    lErro = CF("Produto_ObterEmbMP", objKit.sProdutoRaiz, giFilialEmpresa, colMP, colEmb)
                                    If lErro <> SUCESSO Then gError 184319
                                    
                                    For Each objCustoDirFabrPlanMP In colMP
                                    
                                        Set objItemFP = New ClassItemFormPreco
                                        objItemFP.sProduto = objCustoDirFabrPlanMP.sProdutoMP
                                        
                                        Set objProdutoAux = New ClassProduto
                                        objProdutoAux.sCodigo = objItemFP.sProduto
                                        lErro = CF("Produto_Le", objProdutoAux)
                                        If lErro <> SUCESSO And lErro <> 28030 Then gError 184317
                                        If lErro <> SUCESSO Then gError 184318
                                        
                                        dQtdeAux = objCustoDirFabrPlanMP.dParticipacao * dQtdeBaseKit
                                        
                                        If objProdutoAux.sSiglaUMEstoque <> objProdutoAux.sSiglaUMCompra Then
                                            'converter qtde para UM de compra
                                        End If
                                        
                                        objItemFP.sUnidMed = objProdutoAux.sSiglaUMCompra
                                        objItemFP.dQtde = dQtdeAux
                                        
                                        If gobjFAT.iCotarItensOrcVenda <> 0 Then
                                            objItemFP.iSituacao = FPSITUACAO_COTAR
                                        End If
                                        
                                        lErro = CF("Produto_ObtemMargemPadrao", objItemFP.sProduto, dMargem)
                                        If lErro <> SUCESSO Then gError 184376
                                        objItemFP.dPercentMargem = dMargem
                                        
                                        colItensFormPreco.Add objItemFP
                                        
                                    Next
                                    
                                    For Each objCustoDirFabrPlanEmb In colEmb
                                    
                                        Set objItemFP = New ClassItemFormPreco
                                        objItemFP.sProduto = objCustoDirFabrPlanEmb.sProdutoEmb
                                        
                                        Set objProdutoAux = New ClassProduto
                                        objProdutoAux.sCodigo = objItemFP.sProduto
                                        lErro = CF("Produto_Le", objProdutoAux)
                                        If lErro <> SUCESSO And lErro <> 28030 Then gError 184320
                                        If lErro <> SUCESSO Then gError 184321
                                        
                                        dQtdeAux = objCustoDirFabrPlanEmb.dQtde * dQtdeBaseKit
                                        
                                        If objProdutoAux.sSiglaUMEstoque <> objProdutoAux.sSiglaUMCompra Then
                                            'converter qtde para UM de compra
                                        End If
                                        
                                        objItemFP.sUnidMed = objProdutoAux.sSiglaUMCompra
                                        objItemFP.dQtde = dQtdeAux
                                        
                                        If gobjFAT.iCotarItensOrcVenda <> 0 Then
                                            objItemFP.iSituacao = FPSITUACAO_COTAR
                                        End If
                                        
                                        lErro = CF("Produto_ObtemMargemPadrao", objItemFP.sProduto, dMargem)
                                        If lErro <> SUCESSO Then gError 184376
                                        objItemFP.dPercentMargem = dMargem
                                        
                                        colItensFormPreco.Add objItemFP
                                        
                                    Next
                                    
                                End If
                            
                            End If
                        
                        End If
                        
                        'procurar na colecao já existente os itens que já existiam.
                        For Each objItemFPAux In objItemOV.colItensFormPreco
                        
                            'para os itens que já existiam,
                            For Each objItemFP In colItensFormPreco
                            
                                If objItemFP.sProduto = objItemFPAux.sProduto Then
                                
                                    'se estiver cotado ou aguardando cotacao nao poderá atualizar
                                    If objItemFPAux.iSituacao = FPSITUACAO_COTADO Or objItemFPAux.iSituacao = FPSITUACAO_EM_COTACAO Then
                                    
                                        If Abs(objItemFPAux.dPercentMargem - objItemFP.dPercentMargem) > 0.01 Then
                                            objItemFPAux.dPercentMargem = objItemFP.dPercentMargem
                                            objItemFPAux.dPrecoUnit = Arredonda_Moeda(objItemFPAux.dCustoUnit * (1 + objItemFPAux.dPercentMargem))
                                            objItemFPAux.dPrecoTotal = Arredonda_Moeda(objItemFPAux.dPrecoUnit * objItemFPAux.dQtde)
                                            bRecarregarGrids = True
                                        End If
                                        
                                        bAtualizar = False
                                        Exit For
                                        
                                    Else
                                        'sobrepor As informacoes que já haviam
                                        objItemFP.dCustoUnit = objItemFPAux.dCustoUnit
                                        objItemFP.dPercentMargem = objItemFPAux.dPercentMargem
                                        objItemFP.dPrecoUnit = objItemFPAux.dPrecoUnit
                                        objItemFP.iSituacao = objItemFPAux.iSituacao
                                        
                                        objItemFP.dPrecoTotal = Arredonda_Moeda(objItemFP.dPrecoUnit * objItemFP.dQtde)
                                        
                                    End If
                        
                                    Exit For
                                    
                                End If
                                
        '                        If bAtualizar = False Then Exit For
                                
                            Next
                            
                        Next
                        
                        If bAtualizar Then
                        
                            'substituir a colecao antiga pela nova
                            Set objItemOV.colItensFormPreco = colItensFormPreco
                            bRecarregarGrids = True
                            
                        End If
                        
        ''            End If
                    
                End If
                
            End If
            
        Next
        
        If bRecarregarGrids Then
            
            Call PrecosCalculados_Traz_Dados_Tela
            
        End If
    
    End If
    
    FormPreco_Calcula = SUCESSO
    
    Exit Function
    
Erro_FormPreco_Calcula:

    FormPreco_Calcula = gErr

    Select Case gErr
    
        Case 184312, 184314 To 184317, 184319, 184320, 184376
    
        Case 184313
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILIAL_NAO_CADASTRADO", gErr, objProdutoFilial.sProduto, objProdutoFilial.iFilialEmpresa)
        
        Case 184318, 184321
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184308)

    End Select
    
    Exit Function

End Function

Private Sub PCSituacoes_Carrega()

Dim iIndice As Integer
    
    Set gcolPCSituacoes = New Collection

    For iIndice = 1 To PCSituacao.ListCount
        gcolPCSituacoes.Add PCSituacao.List(iIndice - 1)
    Next
    
End Sub

Private Sub FPSituacoes_Carrega()

Dim iIndice As Integer
    
    Set gcolFPSituacoes = New Collection

    For iIndice = 1 To FPSituacao.ListCount
        gcolFPSituacoes.Add FPSituacao.List(iIndice - 1)
    Next
    
End Sub

Public Sub PCSelecionado_Click()
    'carrega o grid de formacao de precos correspondente a linha selecionada no grid de precos calculados
    
    Call FormPreco_PreencheGrid(GridPrecosCalculados.Row)
    
End Sub

Private Function Saida_Celula_GridPrecosCalculados(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridPrecosCalculados

    Select Case objGridInt.objGrid.Col

        Case iGrid_PCSituacao_Col
            lErro = Saida_Celula_PCSituacao(objGridInt)
            If lErro <> SUCESSO Then gError 184358
        
        Case iGrid_PCPrecoUnit_Col
            lErro = Saida_Celula_PCPrecoUnit(objGridInt)
            If lErro <> SUCESSO Then gError 184358
        
    End Select

    Saida_Celula_GridPrecosCalculados = SUCESSO

    Exit Function

Erro_Saida_Celula_GridPrecosCalculados:

    Saida_Celula_GridPrecosCalculados = gErr

    Select Case gErr

        Case 184358

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184356)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridFormacaoPreco(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridFormacaoPreco

    Select Case objGridInt.objGrid.Col

        Case iGrid_FPSituacao_Col
            lErro = Saida_Celula_FPSituacao(objGridInt)
            If lErro <> SUCESSO Then gError 184359
        
        Case iGrid_FPPercentMargem_Col
            lErro = Saida_Celula_FPPercentMargem(objGridInt)
            If lErro <> SUCESSO Then gError 184359
        
    End Select

    Saida_Celula_GridFormacaoPreco = SUCESSO

    Exit Function

Erro_Saida_Celula_GridFormacaoPreco:

    Saida_Celula_GridFormacaoPreco = gErr

    Select Case gErr

        Case 184359

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184357)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PCSituacao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula PCSituacao que está deixando de ser a corrente

Dim lErro As Long
Dim sPCSituacao As String

On Error GoTo Erro_Saida_Celula_PCSituacao

    Set objGridInt.objControle = PCSituacao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 141388

    '??? poderia testar se trocou
    Call FormPreco_Atualiza_Linha(GridFormacaoPreco.Row)
    
    Call FormPreco_AtualizarSituacao

    Saida_Celula_PCSituacao = SUCESSO

    Exit Function

Erro_Saida_Celula_PCSituacao:

    Saida_Celula_PCSituacao = gErr

    Select Case gErr

        Case 141388
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157857)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FPSituacao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula FPSituacao que está deixando de ser a corrente

Dim lErro As Long
Dim sFPSituacao As String

On Error GoTo Erro_Saida_Celula_FPSituacao

    Set objGridInt.objControle = FPSituacao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 141388

    '??? poderia testar se trocou
    Call FormPreco_Atualiza_Linha(GridFormacaoPreco.Row)

    Call FormPreco_AtualizarSituacao
    
    Saida_Celula_FPSituacao = SUCESSO

    Exit Function

Erro_Saida_Celula_FPSituacao:

    Saida_Celula_FPSituacao = gErr

    Select Case gErr

        Case 141388
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157857)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PCPrecoUnit(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim bPrecoUnitarioIgual As Boolean

On Error GoTo Erro_Saida_Celula_PCPrecoUnit

    bPrecoUnitarioIgual = False

    Set objGridInt.objControle = PCPrecoUnit

    If Len(Trim(PCPrecoUnit.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(PCPrecoUnit.Text)
        If lErro <> SUCESSO Then gError 84170  '26684

        PrecoUnitario.Text = Format(PCPrecoUnit.Text, gobjFAT.sFormatoPrecoUnitario)
    
    End If

    'Comparação com Preço Unitário anterior
    If StrParaDbl(GridPrecosCalculados.TextMatrix(GridPrecosCalculados.Row, iGrid_PCPrecoUnit_Col)) = StrParaDbl(PCPrecoUnit.Text) Then bPrecoUnitarioIgual = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84171 '26685

    If Not bPrecoUnitarioIgual Then

        GridItens.TextMatrix(GridPrecosCalculados.Row, iGrid_PrecoUnitario_Col) = Format(StrParaDbl(PCPrecoUnit.Text), gobjFAT.sFormatoPrecoUnitario)
        
        '???? furo 12/07/2013 ???? testar
        lErro = gobjTribTab.Alteracao_Item_Grid(GridPrecosCalculados.Row)
        If lErro <> SUCESSO Then gError 84172
        
        Call PrecoTotal_Calcula(GridPrecosCalculados.Row)
        
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError 84172 '51042

    End If

   Saida_Celula_PCPrecoUnit = SUCESSO

    Exit Function

Erro_Saida_Celula_PCPrecoUnit:

    Saida_Celula_PCPrecoUnit = gErr

    Select Case gErr

        Case 84170, 84171, 84172
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157794)

    End Select

    Exit Function

End Function

Function Saida_Celula_FPPercentMargem(objGridInt As AdmGrid) As Long
'

Dim lErro As Long
Dim dFPPercentMargem As Double
Dim sValorPercAnterior As String
Dim dPrecoUnitario As Double, dPrecoTotal As Double

On Error GoTo Erro_Saida_Celula_FPPercentMargem

    Set objGridInt.objControle = FPPercentMargem

    If Len(FPPercentMargem.Text) > 0 Then
        'Critica a porcentagem
        lErro = Porcentagem_Critica_Negativa(FPPercentMargem.Text) 'Alterado por Wagner
        If lErro <> SUCESSO Then gError 84329 '26694

        dFPPercentMargem = CDbl(FPPercentMargem.Text)

        If Format(dFPPercentMargem, "#0.#0\%") <> GridFormacaoPreco.TextMatrix(GridFormacaoPreco.Row, iGrid_FPPercentMargem_Col) Then

            FPPercentMargem.Text = Format(dFPPercentMargem, "Fixed")

        End If

    End If

    sValorPercAnterior = GridFormacaoPreco.TextMatrix(GridFormacaoPreco.Row, iGrid_FPPercentMargem_Col)

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 84331 '26696
    
    'Se foi alterada
    If Format(dFPPercentMargem, "#0.#0\%") <> sValorPercAnterior Then
        
        'recalcular preco do item e atualizar a coluna de preco calculado
        Call FormPreco_Atualiza_Linha(GridFormacaoPreco.Row)
        
    End If

    Saida_Celula_FPPercentMargem = SUCESSO

    Exit Function

Erro_Saida_Celula_FPPercentMargem:

    Saida_Celula_FPPercentMargem = gErr

    Select Case gErr

        Case 84329, 84331
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157795)

    End Select

    Exit Function

End Function

Private Sub FormPreco_Atualiza_Linha(ByVal iLinhaFP As Integer)

Dim objItemOV As ClassItemOV, iLinhaPC As Integer, objItemFP As ClassItemFormPreco
Dim dPrecoTotalNovo As Double, iLinhaAux As Integer, iSituacao As Integer
Dim dQtdeItemOV As Double, dPrecoTotalAcumulado As Double, dPrecoUnitNovo As Double
Dim bTrocouSituacao As Boolean

    If gobjFAT.iOVSemFormPreco = 0 Then
    
        bTrocouSituacao = False
        
        iLinhaPC = GridPrecosCalculados.Row
        
        If gcolItensOVPrecoCalc.Count >= iLinhaPC And iLinhaPC <> 0 Then
        
            Set objItemOV = gcolItensOVPrecoCalc.Item(iLinhaPC)
            
            If objItemOV.colItensFormPreco.Count >= iLinhaFP And iLinhaFP <> 0 Then
            
                Set objItemFP = objItemOV.colItensFormPreco.Item(iLinhaFP)
                
                objItemFP.dCustoUnit = StrParaDbl(GridFormacaoPreco.TextMatrix(iLinhaFP, iGrid_FPCustoUnit_Col))
                objItemFP.dPercentMargem = PercentParaDbl(GridFormacaoPreco.TextMatrix(iLinhaFP, iGrid_FPPercentMargem_Col))
                dPrecoUnitNovo = Arredonda_Moeda(objItemFP.dCustoUnit * (1 + objItemFP.dPercentMargem))
                If Abs(objItemFP.dPrecoUnit - dPrecoUnitNovo) > DELTA_VALORMONETARIO Then
                    objItemFP.dPrecoUnit = dPrecoUnitNovo
                    GridFormacaoPreco.TextMatrix(iLinhaFP, iGrid_FPPrecoUnit_Col) = Format(objItemFP.dPrecoUnit, gobjFAT.sFormatoPrecoUnitario)
                End If
                
                Select Case GridFormacaoPreco.TextMatrix(iLinhaFP, iGrid_FPSituacao_Col)
                    Case FPSITUACAO_STRING_NAO_COTAR
                        iSituacao = FPSITUACAO_NAO_COTAR
                    Case FPSITUACAO_STRING_COTAR
                        iSituacao = FPSITUACAO_COTAR
                    Case FPSITUACAO_STRING_EM_COTACAO
                        iSituacao = FPSITUACAO_EM_COTACAO
                    Case FPSITUACAO_STRING_COTADO
                        iSituacao = FPSITUACAO_COTADO
                End Select
                        
                If objItemFP.iSituacao <> iSituacao Then
                    objItemFP.iSituacao = iSituacao
                    bTrocouSituacao = True
                End If
                
                dPrecoTotalNovo = Arredonda_Moeda(objItemFP.dPrecoUnit * objItemFP.dQtde)
                If Abs(objItemFP.dPrecoTotal - dPrecoTotalNovo) > DELTA_VALORMONETARIO Then
                    objItemFP.dPrecoTotal = dPrecoTotalNovo
                    GridFormacaoPreco.TextMatrix(iLinhaFP, iGrid_FPPrecoTotal_Col) = Format(objItemFP.dPrecoTotal, "Standard")
                    
                    'recalcular preço total da linha de precos calculados
                    dQtdeItemOV = StrParaDbl(GridItens.TextMatrix(iLinhaPC, iGrid_Quantidade_Col))
                    If dQtdeItemOV <> 0 Then
                        dPrecoTotalAcumulado = 0
                        For iLinhaAux = 1 To objItemOV.colItensFormPreco.Count
                            dPrecoTotalAcumulado = dPrecoTotalAcumulado + objItemOV.colItensFormPreco.Item(iLinhaAux).dPrecoTotal
                        Next
                        GridPrecosCalculados.TextMatrix(iLinhaPC, iGrid_PCPrecoUnitCalc_Col) = Format(dPrecoTotalAcumulado / dQtdeItemOV, gobjFAT.sFormatoPrecoUnitario)
                    Else
                        GridPrecosCalculados.TextMatrix(iLinhaPC, iGrid_PCPrecoUnitCalc_Col) = ""
                    End If
                End If
                
            End If
            
        End If
        
        If bTrocouSituacao Then Call FormPreco_AtualizarSituacao
        
    End If
    
End Sub

Public Sub BotaoCotacoesPendentes_Click()
    
Dim objPedidoCotacao As New ClassPedidoCotacao, objItemFP As ClassItemFormPreco
Dim colSelecao As New Collection, objItemOV As ClassItemOV, lErro As Long
Dim iLinhaPC As Integer, iLinhaFP As Integer, lNumIntCotProd As Long, sFiltro As String

On Error GoTo Erro_BotaoCotacoesPendentes_Click

    iLinhaPC = GridPrecosCalculados.Row
    
    If iLinhaPC <> 0 And gcolItensOVPrecoCalc.Count >= iLinhaPC Then
    
        Set objItemOV = gcolItensOVPrecoCalc.Item(iLinhaPC)
        
        iLinhaFP = GridFormacaoPreco.Row
        If iLinhaFP <> 0 And iLinhaFP <= objItemOV.colItensFormPreco.Count Then
    
            sFiltro = "ItemCotPrecoUnitario IS NULL"
            
            Set objItemFP = objItemOV.colItensFormPreco.Item(iLinhaFP)
        
            If objItemFP.lNumIntDoc <> 0 Then
            
                lErro = CF("ItemFormPreco_Obter_CotProd", objItemFP.lNumIntDoc, lNumIntCotProd)
                If lErro <> SUCESSO Then gError 184370
                
                If lNumIntCotProd <> 0 Then
                    sFiltro = sFiltro & " AND NumIntDocCotProd=?"
                    colSelecao.Add lNumIntCotProd
                End If
                
            End If
            
            'Chama a tela
            Call Chama_Tela("PedCotCompletoLista", colSelecao, objPedidoCotacao, Nothing, sFiltro)
            
        End If
    Else
        '??? dar msg de erro
    End If
    
    Exit Sub
    
Erro_BotaoCotacoesPendentes_Click:

    Select Case gErr

        Case 184370
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184371)

    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoCotacoesRecebidas_Click()
    
Dim objPedidoCotacao As New ClassPedidoCotacao, objItemFP As ClassItemFormPreco
Dim colSelecao As New Collection, objItemOV As ClassItemOV, lErro As Long
Dim iLinhaPC As Integer, iLinhaFP As Integer, lNumIntCotProd As Long, sFiltro As String

On Error GoTo Erro_BotaoCotacoesRecebidas_Click

    iLinhaPC = GridPrecosCalculados.Row
    
    If iLinhaPC <> 0 And gcolItensOVPrecoCalc.Count >= iLinhaPC Then
    
        Set objItemOV = gcolItensOVPrecoCalc.Item(iLinhaPC)
        
        iLinhaFP = GridFormacaoPreco.Row
        If iLinhaFP <> 0 And iLinhaFP <= objItemOV.colItensFormPreco.Count Then
    
            sFiltro = "ItemCotPrecoUnitario <> 0"
            
            Set objItemFP = objItemOV.colItensFormPreco.Item(iLinhaFP)
        
            If objItemFP.lNumIntDoc <> 0 Then
            
                lErro = CF("ItemFormPreco_Obter_CotProd", objItemFP.lNumIntDoc, lNumIntCotProd)
                If lErro <> SUCESSO Then gError 184370
                
                If lNumIntCotProd <> 0 Then
                    sFiltro = sFiltro & " AND NumIntDocCotProd=?"
                    colSelecao.Add lNumIntCotProd
                End If
                
            End If
            
            'Chama a tela
            Call Chama_Tela("PedCotCompletoLista", colSelecao, objPedidoCotacao, objEventoBotaoPedAtualizados, sFiltro)
            
        End If
    Else
        '??? dar msg de erro
    End If
    
    Exit Sub
    
Erro_BotaoCotacoesRecebidas_Click:

    Select Case gErr

        Case 184370
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184371)

    End Select
    
    Exit Sub
    
End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    If (iCodigoAlterado = REGISTRO_ALTERADO) Then
    
        If StrParaLong(Codigo.Text) <> 0 Then
    
            objOrcamentoVenda.lCodigo = StrParaLong(Codigo.Text)
            objOrcamentoVenda.iFilialEmpresa = giFilialEmpresa
            
            'Lê os dados do Orcamento de Venda
            lErro = CF("OrcamentoVenda_Le", objOrcamentoVenda)
            If lErro <> SUCESSO And lErro <> 101232 Then gError 131274
            If lErro = SUCESSO Then
    
                lErro = Traz_OrcamentoVenda_Tela(objOrcamentoVenda, VALOR_CODIGO)
                If lErro <> SUCESSO Then gError 131276
                
            Else
            
                Call PrecosCalculados_Recriar
            
            End If
            
        Else
        
            Call PrecosCalculados_Recriar
            
        End If
    
        iCodigoAlterado = 0
    
    End If
    
    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr
    
        Case 131274, 131276
        
        Case 131275
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157735)

      End Select

    Exit Sub

End Sub

Private Sub PrecosCalculados_Recriar()
    
Dim lErro As Long, objItemOV As New ClassItemOV, iIndice As Integer
Dim sProduto As String, sProduto1 As String
Dim iPreenchido As Integer

On Error GoTo Erro_PrecosCalculados_Recriar

    Set gcolItensOVPrecoCalc = New Collection
    Call Grid_Limpa(objGridPrecosCalculados)
    Call Grid_Limpa(objGridFormacaoPreco)

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        Set objItemOV = New ClassItemOV
        sProduto1 = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)
        
        'Verifica se o Produto está preenchido
        If Len(Trim(sProduto1)) > 0 Then
    
            'Formata o produto
            lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 184411
    
            objItemOV.sProduto = sProduto
            
        End If
    
        If Len(Trim(sProduto1)) = 0 Or iPreenchido = PRODUTO_VAZIO Then Exit For
        
        If gobjFAT.iCotarItensOrcVenda <> 0 Then
            objItemOV.iPCSituacao = ITEMOV_PCSITUACAO_COTAR
        End If
    
        gcolItensOVPrecoCalc.Add objItemOV
        
    Next
    
    lErro = FormPreco_Calcula()
    If lErro <> SUCESSO Then gError 184412
    
    Call FormPreco_AtualizarSituacao
    
    Call PrecosCalculados_Traz_Dados_Tela
    
    Exit Sub
    
Erro_PrecosCalculados_Recriar:

    Select Case gErr

        Case 184411, 184412

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 184413)

    End Select
    
End Sub

Private Sub FormPreco_AtualizarSituacao()

Dim lErro As Long, objItemOV As ClassItemOV, iIndice As Integer, iStatusComercialAnterior As Integer
Dim objItemFP As ClassItemFormPreco, iPCSituacao As Integer, bSituacaoItemOVAlterada As Boolean
Dim bTemItemFPCotar As Boolean, bTemItemFPEmCotacao As Boolean, bTemItemFPCotado As Boolean
Dim bTemItemOVEmCotacao As Boolean, bTemItemOVCotado As Boolean, iStatusComercial As Integer

    If gobjFAT.iOVSemFormPreco = 0 Then
    
        bSituacaoItemOVAlterada = False
        For iIndice = 1 To objGridItens.iLinhasExistentes
        
            If gcolItensOVPrecoCalc.Count >= iIndice Then
            
                Set objItemOV = gcolItensOVPrecoCalc.Item(iIndice)
                
                bTemItemFPCotar = False
                bTemItemFPEmCotacao = False
                bTemItemFPCotado = False
                For Each objItemFP In objItemOV.colItensFormPreco
                    
                    Select Case objItemFP.iSituacao
                    
                        Case FPSITUACAO_COTAR
                            bTemItemFPCotar = True
                            
                        Case FPSITUACAO_EM_COTACAO
                            bTemItemFPEmCotacao = True
                            
                        Case FPSITUACAO_COTADO
                            bTemItemFPCotado = True
                            
                    End Select
                    
                Next
                
                If bTemItemFPCotar Then
                    iPCSituacao = ITEMOV_PCSITUACAO_COTAR
                Else
                    If bTemItemFPEmCotacao Then
                        iPCSituacao = ITEMOV_PCSITUACAO_EM_COTACAO
                    Else
                        If bTemItemFPCotado Then
                            iPCSituacao = ITEMOV_PCSITUACAO_COTADO
                        Else
                            iPCSituacao = ITEMOV_PCSITUACAO_NAO_COTAR
                        End If
                    End If
                End If
                
                If iPCSituacao <> objItemOV.iPCSituacao Then
                    bSituacaoItemOVAlterada = True
                    objItemOV.iPCSituacao = iPCSituacao
                    GridPrecosCalculados.TextMatrix(iIndice, iGrid_PCSituacao_Col) = gcolPCSituacoes.Item(objItemOV.iPCSituacao + 1)
                End If
                
            End If
            
        Next
        
        bTemItemOVEmCotacao = False
        bTemItemOVCotado = False
        
        For iIndice = 1 To objGridItens.iLinhasExistentes
        
            If gcolItensOVPrecoCalc.Count >= iIndice Then
            
                Set objItemOV = gcolItensOVPrecoCalc.Item(iIndice)
                
                Select Case GridPrecosCalculados.TextMatrix(iIndice, iGrid_PCSituacao_Col)
                
                    Case ITEMOV_PCSITUACAO_STRING_COTAR
                        objItemOV.iPCSituacao = ITEMOV_PCSITUACAO_COTAR
                    Case ITEMOV_PCSITUACAO_STRING_EM_COTACAO
                        objItemOV.iPCSituacao = ITEMOV_PCSITUACAO_EM_COTACAO
                    Case ITEMOV_PCSITUACAO_STRING_COTADO
                        objItemOV.iPCSituacao = ITEMOV_PCSITUACAO_COTADO
                    Case Else
                        objItemOV.iPCSituacao = ITEMOV_PCSITUACAO_NAO_COTAR
                    
                End Select
                
                Select Case objItemOV.iPCSituacao
                
                    Case ITEMOV_PCSITUACAO_COTAR, ITEMOV_PCSITUACAO_EM_COTACAO
                        bTemItemOVEmCotacao = True
                        
                    Case ITEMOV_PCSITUACAO_COTADO
                        bTemItemOVCotado = True
                        
                End Select
                
            End If
            
        Next
        
        iStatusComercial = OV_STATUS_COMERCIAL_NAO_COTAR
        If bTemItemOVEmCotacao Then
            iStatusComercial = OV_STATUS_COMERCIAL_EM_COTACAO
        Else
            If bTemItemOVCotado Then
                iStatusComercial = OV_STATUS_COMERCIAL_COTADO
            End If
        End If
        
        iStatusComercialAnterior = OV_STATUS_COMERCIAL_NAO_COTAR
        If StatusComercial.ListIndex <> -1 Then
            iStatusComercialAnterior = StatusComercial.ItemData(StatusComercial.ListIndex)
        End If
        
        If iStatusComercialAnterior <> OV_STATUS_COMERCIAL_LIBERADO Then
            If iStatusComercialAnterior <> iStatusComercial Then
                Call Combo_Seleciona_ItemData(StatusComercial, iStatusComercial)
            End If
        End If
    
    End If
    
End Sub

Public Function OV_Form_UnLoad(ByVal objCT As Object, Cancel As Integer) As Long
    OV_Form_UnLoad = SUCESSO
End Function

Function Trata_Parametros(Optional objOrcamentoVenda As ClassOrcamentoVenda) As Long

Dim lErro As Long
Dim objOrcamentoVendaAux As New ClassOrcamentoVenda

On Error GoTo Erro_Trata_Parametros

    lErro = CargaPosFormLoad(True)
    If lErro <> SUCESSO Then gError 59288
    
    If Not (objOrcamentoVenda Is Nothing) Then
        
        'Se foi passado o código do orçamento
        If objOrcamentoVenda.lCodigo > 0 Then 'Incluído por Luiz Nogueira em 13/01/04
            
            objOrcamentoVendaAux.lCodigo = objOrcamentoVenda.lCodigo
            objOrcamentoVendaAux.iFilialEmpresa = objOrcamentoVenda.iFilialEmpresa
            objOrcamentoVendaAux.iVersao = objOrcamentoVenda.iVersao
            
            'Coloca o Pedido de Venda na tela
            
            lErro = Traz_OrcamentoVenda_Tela(objOrcamentoVendaAux, VALOR_CODIGO)
            If lErro <> SUCESSO And lErro <> 84363 Then gError 84479
    
            If lErro <> SUCESSO Then  'Não encontrou no BD o código de Pedido
    
                'Limpa a tela e coloca o código na Tela
                Call Limpa_OrcamentoVenda
                Codigo.Text = CStr(objOrcamentoVenda.lCodigo)
    
            End If

        '*** Incluído por Luiz Nogueira em 13/01/04 - INÍCIO ***
        'Se foi passado o código do cliente
        ElseIf objOrcamentoVenda.lCliente > 0 Then
        
            'Joga o código do cliente na tela
            Cliente.Text = objOrcamentoVenda.lCliente
            Call Cliente_Validate(bSGECancelDummy)
            
            'Se foi passada uma filial de cliente
            If objOrcamentoVenda.iFilial > 0 Then
            
                'Joga a filial do cliente na tela
                Filial.Text = objOrcamentoVenda.iFilial
                Call Filial_Validate(bSGECancelDummy)
            End If
            
            'Cria um número automático para o orçamento
            Call BotaoProxNum_Click
        '*** Incluído por Luiz Nogueira em 13/01/04 - FIM ***
        End If

    End If

    iCodigoAlterado = 0
    iAlterado = 0
    iVendedorAlterado = 0
    iVendedor2Alterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 84479, 59288

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157718)

    End Select

    iAlterado = 0

    Exit Function

End Function


Public Sub BotaoVersoesOVs_Click()

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoVersoesOVs_Click
    

    Call Chama_Tela("OrcamentoVendaCGHistLista", colSelecao, objOrcamentoVenda, objEventoNumero)
    
    Exit Sub

Erro_BotaoVersoesOVs_Click:

    Select Case gErr

        Case 84106

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157762)

    End Select

    Exit Sub

End Sub

Public Sub BotaoOVVersoes_Click()

Dim lErro As Long
Dim objOrcamentoVenda As New ClassOrcamentoVenda
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoOVVersoes_Click
    
    colSelecao.Add StrParaLong(Codigo.Text)

    Call Chama_Tela("OrcamentoVendaCGHistLista", colSelecao, objOrcamentoVenda, objEventoNumero, "Codigo = ?")
    
    Exit Sub

Erro_BotaoOVVersoes_Click:

    Select Case gErr

        Case 84106

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157762)

    End Select

    Exit Sub

End Sub

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
Dim objProduto As New ClassProduto, objItemOV As ClassItemOV
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
            
                GridItens.TextMatrix(iItem, iGrid_Status_Col) = Status.Text
                GridItens.TextMatrix(iItem, iGrid_MotivoPerda_Col) = MotivoPerda.Text

                If objRomaneioGradeAux.objTabelaPrecoItem.dPreco > 0 Then
        
                    lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(iItem, iGrid_UnidadeMed_Col), objRomaneioGradeAux.objTabelaPrecoItem, dPrecoUnitario)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                    lErro = CF2(Me, "OV_Move_PrecoUnitario_Tela", dPrecoUnitario, iItem)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
'                    If dPrecoUnitario > 0 Then
'                        GridItens.TextMatrix(iItem, iGrid_PrecoUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
'                    Else
'                        GridItens.TextMatrix(iItem, iGrid_PrecoUnitario_Col) = ""
'                    End If
                    
                End If
                
                Call gobjTelaProjetoInfo.Adiciona_Linha(iItem)
        
                Call FormPreco_Adiciona_Linha(iItem, objProduto)
                
                'permite que a tributacao reflita a inclusao de uma linha no grid
                lErro = gobjTribTab.Inclusao_Item_Grid(iItem, objProduto.sCodigo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                gobjOrcamentoVenda.colItens(iItem).iPossuiGrade = MARCADO
                gobjOrcamentoVenda.colItens(iItem).iItem = iItem
                gobjOrcamentoVenda.colItens(iItem).sProduto = objProduto.sCodigo
                           
                Set gobjOrcamentoVenda.colItens(iItem).colItensRomaneioGrade = objRomaneioGradeAux.colItensRomaneioGrade
                   
                GridItens.TextMatrix(iItem, 0) = "# " & GridItens.TextMatrix(iItem, 0)
                
                Set objItemOV = gobjOrcamentoVenda.colItens(iItem)
                       
                Call Atualiza_Grid_Itens(objItemOV)
                
                Call PrecoTotal_Calcula(iItem)
            
            Next
        
            Call ValorTotal_Calcula
        
        End If
    
    End If
   
    GrupoGrade_Separa_Itens = SUCESSO

    Exit Function

Erro_GrupoGrade_Separa_Itens:

    GrupoGrade_Separa_Itens = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208382)

    End Select

    Exit Function

End Function
'Funções para tratamento de quebra de itens de grade
'*************************************

Public Sub ValorDescontoTit_Change()
    'Seta iComissoesAlterada
    'iComissoesAlterada = REGISTRO_ALTERADO
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
Dim iLinha As Integer

On Error GoTo Erro_ValorDescontoItens_Aplica

    If objGridItens.iLinhasExistentes > 0 Then
    
        dTotal = StrParaDbl(ValorProdutos2.Caption)
        dDesconto = StrParaDbl(ValorDescontoItens.Text)
        dFator = dDesconto / dTotal
    
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col)) > 0 And gobjOrcamentoVenda.colItens.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
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
            If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col)) > 0 And gobjOrcamentoVenda.colItens.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
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
                If gobjOrcamentoVenda.colItens.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
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

Public Sub Contato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Email_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CalcularST_Click()
    
    'Flag usado para trocar a forma como o CST é calculado
    If CalcularST.Value = vbChecked Then
        gobjOrcamentoVenda.iNaoCalcICMSST = DESMARCADO
    Else
        gobjOrcamentoVenda.iNaoCalcICMSST = MARCADO
    End If

    'Força o cálculo com o novo parâmetro
    Call gobjTribTab.TributacaoRecalcular_Click
    
End Sub

'==================================================
'Incluido em 06/06/2012

Public Sub BotaoInfoAdicItem_Click()

Dim lErro  As Long
Dim objItemOV As ClassItemOV
Dim objTela As Object

On Error GoTo Erro_BotaoInfoAdicItem_Click

    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
    
        Set objItemOV = gobjOrcamentoVenda.colItens(GridItens.Row)
        
        Set objTela = Me
        
        objItemOV.objInfoAdicDocItem.iTipoDoc = TIPODOC_INFOADIC_ITEMOV
        objItemOV.objInfoAdicDocItem.iItem = GridItens.Row
        objItemOV.objInfoAdicDocItem.sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        objItemOV.objInfoAdicDocItem.sDescProd = GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col)
        
        objItemOV.objInfoAdicDocItem.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
        objItemOV.objInfoAdicDocItem.dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
        objItemOV.objInfoAdicDocItem.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))
        objItemOV.objInfoAdicDocItem.dValorTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))
        
        Call Chama_Tela_Modal("InfoAdicDocItem", objItemOV.objInfoAdicDocItem, objTela, TIPO_SAIDA, gobjInfoAdic)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Call PrecoTotal_Calcula(GridItens.Row)
        
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Exit Sub

Erro_BotaoInfoAdicItem_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211485)
            
    End Select
    
    Exit Sub

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
        If FilialEntrega_TrataAlteracao() <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If

    Exit Sub

Erro_FilialEntrega_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211486)

    End Select

    Exit Sub

End Sub

Public Sub Transportadora_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Transportadora_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoFrete_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Function FilialEntrega_TrataAlteracao() As Long
'Trata a alteracao da filial de entrega selecionada para o pedido

Dim lErro As Long
On Error GoTo Erro_FilialEntrega_TrataAlteracao

    giFilialEntregaManual = 1
    
    gobjOrcamentoVenda.iFilialEntrega = Codigo_Extrai(FilialEntrega.Text)
    
    'Calula o valor total
    lErro = ValorTotal_Calcula()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    FilialEntrega_TrataAlteracao = SUCESSO

    Exit Function

Erro_FilialEntrega_TrataAlteracao:

    FilialEntrega_TrataAlteracao = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211487)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Se nao encontra o item com o código informado
    If lErro = 6730 Then

        'Verifica de o Cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 211488

        sCliente = Cliente.Text

        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe FilialEntrega com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError ERRO_SEM_MENSAGEM

        'Se não encontrou a Filial do Cliente --> erro
        If lErro = 17660 Then gError 211489

        'Coloca na tela
        FilialEntrega.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        If FilialEntrega_TrataAlteracao() <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 211490

    giFilialEntregaAlterada = 0
    
    Exit Sub

Erro_FilialEntrega_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 211488
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 211489
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 211490
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialEntrega.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211491)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError ERRO_SEM_MENSAGEM
        If lErro <> SUCESSO Then gError 211492 'Não encontrou Transportadora no BD

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 211493

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 211492  'Não encontrou Transportadora no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA", iCodigo)

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            End If
            'Segura o foco

        Case 211493
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211494)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Nao existe o item com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCanalVenda.iCodigo = iCodigo

        'Tenta ler CanalVenda com esse código no BD
        lErro = CF("CanalVenda_Le", objCanalVenda)
        If lErro <> SUCESSO And lErro <> 23597 Then gError ERRO_SEM_MENSAGEM
        If lErro <> SUCESSO Then gError 211495 'Não encontrou CanalVenda no BD

        'Encontrou CanalVenda no BD, coloca no Text da Combo
        CanalVenda.Text = CStr(objCanalVenda.iCodigo) & SEPARADOR & objCanalVenda.sNomeReduzido

    End If

    'Não existe o item com a STRING na List da ComboBox
    If lErro = 6731 Then gError 211496

    Exit Sub

Erro_CanalVenda_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 211495  'Não encontrou CanalVenda no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CANALVENDA", iCodigo)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("CanalDeVenda", objCanalVenda)
            End If

        Case 211496
            Call Rotina_Erro(vbOKOnly, "ERRO_CANALVENDA_NAO_ENCONTRADO", gErr, CanalVenda.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211497)

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
    If lErro <> SUCESSO And lErro <> 58583 Then gError ERRO_SEM_MENSAGEM

    'Se não encontrar --> Erro
    If lErro = 58583 Then gError 211498

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 211498
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211499)

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
Dim vbMsgResult As VbMsgBoxResult
Dim iLinha As Integer

On Error GoTo Erro_DataEntregaPV_Validate

    If giDataEntregaAlterada <> REGISTRO_ALTERADO Then Exit Sub
    
    If StrParaDate(DataEntregaPV.Text) <> DATA_NULA Then

        'If Cancel Then Exit Sub
    
        'Verifica se a Data foi digitada
        If Len(Trim(DataEntregaPV.ClipText)) <> 0 Then
        
            'Critica a data digitada
            lErro = Data_Critica(DataEntregaPV.Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        dtData = MaskedParaDate(DataEntregaPV)
                
        If gobjOrcamentoVenda.dtDataEntrega <> StrParaDate(DataEntregaPV.Text) Then
            
            gobjOrcamentoVenda.dtDataEntrega = StrParaDate(DataEntregaPV.Text)
            
            Call ValorTotal_Calcula
            
            If Not gbCarregandoTela Then
                If objGridItens.iLinhasExistentes > 0 Then
                    vbMsgResult = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_TROCAR_DATAENTREGA_ITENS_CADASTRADOS")
                    If vbMsgResult = vbYes Then
                        For iLinha = 1 To objGridItens.iLinhasExistentes
                            GridItens.TextMatrix(iLinha, iGrid_DataEntrega_Col) = Format(dtData, "dd/mm/yyyy")
                            GridItens.TextMatrix(iLinha, iGrid_PrazoEntrega_Col) = ""
                        Next
                    End If
                End If
            End If
            
        End If
        
        giDataEntregaAlterada = 0
        
    End If
    
    
    Exit Sub

Erro_DataEntregaPV_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211500)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntregaPV_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntregaPV_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEntregaPV, DIMINUI_DATA)
    If lErro Then gError ERRO_SEM_MENSAGEM
    
    Call DataEntregaPV_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDownEntregaPV_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211501)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntregaPV_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntregaPV_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEntregaPV, AUMENTA_DATA)
    If lErro Then gError ERRO_SEM_MENSAGEM
    
    Call DataEntregaPV_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDownEntregaPV_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211502)

    End Select

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo

        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError ERRO_SEM_MENSAGEM
        
        'Não encontrou Transportadora no BD
        If lErro <> SUCESSO Then gError 211503
        
        'Encontrou Transportadora no BD, coloca no Text da Combo
        TranspRedespacho.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 211504

    Exit Sub

Erro_TranspRedespacho_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 211503  'Não encontrou Transportadora no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)

            End If

        Case 211504
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, TranspRedespacho.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211505)

    End Select

    Exit Sub

End Sub

Public Sub RedespachoCli_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211506)

    End Select

    Exit Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211507)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211508)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Coloca o valor formatado na Tela
    PesoLiquido.Text = Format(PesoLiquido.Text, PesoLiquido.Format)

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211509)

    End Select

    Exit Sub

End Sub

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If iPreenchido <> PRODUTO_PREENCHIDO Then gError 211510
    
        'preenche o código do produto
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
        If lErro = 28030 Then gError 211511
    
        'Realiza a converção para uma mesma UM
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Realiza o cálculo do volume
        dVolume = dVolume + (((StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) * dFator) * (objProduto.dComprimento * objProduto.dEspessura * objProduto.dLargura))
    
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

        Case ERRO_SEM_MENSAGEM

        Case 211510
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 211511
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211512)

    End Select

    Exit Function
    
End Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211513)

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

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211514)

    End Select

End Function

Public Sub Mensagem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CanalVenda_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CanalVenda_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PedidoRepr_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PedidoRepr_Validate(Cancel As Boolean)

Dim lErro As Long

    lErro = CF("PV_PedidoRepr_Valida", StrParaDbl(Codigo.Text), giFilialEmpresa, PedidoRepr.Text)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Carrega_Transportadoras() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_Transportadoras

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_TRANSPORTADORA_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211515)

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
        CanalVenda.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        CanalVenda.ItemData(CanalVenda.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_CanalVenda = SUCESSO

    Exit Function

Erro_Carrega_CanalVenda:

    Carrega_CanalVenda = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211516)

    End Select

    Exit Function

End Function

Function Carrega_Tab_Complemento(objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Coloca os dados do Tab de Complementos na tela

Dim iIndice As Integer

On Error GoTo Erro_Carrega_Tab_Complemento

    PedidoCliente.Text = objOrcamentoVenda.sPedidoCliente
    PedidoRepr.Text = objOrcamentoVenda.sPedidoRepresentante
    Mensagem.Text = objOrcamentoVenda.sMensagemNota

    If objOrcamentoVenda.iCanalVenda > 0 Then
        CanalVenda.Text = objOrcamentoVenda.iCanalVenda
        Call CanalVenda_Validate(bSGECancelDummy)
    End If

    If objOrcamentoVenda.iFilialEntrega > 0 Then
        Call Filial_Formata(FilialEntrega, objOrcamentoVenda.iFilialEntrega)
        giFilialEntregaAlterada = 0
    End If

    If objOrcamentoVenda.iCodTransportadora > 0 Then
        Transportadora.Text = objOrcamentoVenda.iCodTransportadora
        Call Transportadora_Validate(bSGECancelDummy)
    End If

    If objOrcamentoVenda.iCodTranspRedesp > 0 Then
        TranspRedespacho.Text = objOrcamentoVenda.iCodTranspRedesp
        Call TranspRedespacho_Validate(bSGECancelDummy)
    End If

    RedespachoCli.Value = objOrcamentoVenda.iDetPagFrete

    If objOrcamentoVenda.lVolumeQuant > 0 Then VolumeQuant.Text = objOrcamentoVenda.lVolumeQuant

    Placa.Text = objOrcamentoVenda.sPlaca
    PlacaUF.Text = objOrcamentoVenda.sPlacaUF
    
    'Traz a espécie dos volumes do pedido
    If objOrcamentoVenda.lVolumeEspecie > 0 Then
        VolumeEspecie.Text = objOrcamentoVenda.lVolumeEspecie
        Call VolumeEspecie_Validate(bSGECancelDummy)
    End If
    
    'Traz a marca dos volumes do pedido
    If objOrcamentoVenda.lVolumeMarca > 0 Then
        VolumeMarca.Text = objOrcamentoVenda.lVolumeMarca
        Call VolumeMarca_Validate(bSGECancelDummy)
    End If
    
    VolumeNumero = objOrcamentoVenda.sVolumeNumero

    For iIndice = 0 To TipoFrete.ListCount - 1
        If TipoFrete.ItemData(iIndice) = objOrcamentoVenda.iFreteRespons Then
            TipoFrete.ListIndex = iIndice
            Exit For
        End If
    Next
            
    If objOrcamentoVenda.dPesoLiq > 0 Then PesoLiquido.Text = Format(objOrcamentoVenda.dPesoLiq, PesoLiquido.Format)
    If objOrcamentoVenda.dPesoBruto > 0 Then PesoBruto.Text = Format(objOrcamentoVenda.dPesoBruto, PesoBruto.Format)
    If objOrcamentoVenda.dVolumeTotal > 0 Then Cubagem.Text = Format(objOrcamentoVenda.dVolumeTotal, "Fixed")

    Carrega_Tab_Complemento = SUCESSO

    Exit Function

Erro_Carrega_Tab_Complemento:

    Carrega_Tab_Complemento = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211517)

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

Public Sub PrazoEntregaItem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PrazoEntregaItem_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub PrazoEntregaItem_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub PrazoEntregaItem_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridItens.objControle = PrazoEntregaItem
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub DataEnvio_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataEnvio_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataEnvio, iAlterado)
End Sub

Public Sub DataEnvio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEnvio_Validate

    'Verifica se a Data de Envio foi digitada
    If Len(Trim(DataEnvio.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEnvio.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_DataEnvio_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211518)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEnvio_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEnvio_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEnvio, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownEnvio_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211519)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEnvio_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEnvio_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEnvio, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownEnvio_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211520)

    End Select

    Exit Sub

End Sub

Public Sub PedidoCliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PrazoEntrega_GotFocus()
    Call MaskEdBox_TrataGotFocus(PrazoEntrega, iAlterado)
End Sub

Public Sub PrazoEntrega_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PrazoEntrega_Validate(Cancel As Boolean)
Dim vbMsgResult As VbMsgBoxResult
Dim iLinha As Integer
    If Not gbCarregandoTela Then
        If iPrazoAnt <> StrParaInt(PrazoEntrega.Text) Then
            If objGridItens.iLinhasExistentes > 0 Then
                vbMsgResult = Rotina_Aviso(vbYesNo, "AVISO_DESEJA_TROCAR_PRAZOENTREGA_ITENS_CADASTRADOS")
                If vbMsgResult = vbYes Then
                    For iLinha = 1 To objGridItens.iLinhasExistentes
                        GridItens.TextMatrix(iLinha, iGrid_PrazoEntrega_Col) = CStr(StrParaInt(PrazoEntrega.Text))
                        GridItens.TextMatrix(iLinha, iGrid_DataEntrega_Col) = ""
                    Next
                End If
            End If
        End If
    End If
    iPrazoAnt = StrParaInt(PrazoEntrega.Text)
End Sub

Private Function Move_TabComplemento_Memoria(objOV As ClassOrcamentoVenda) As Long

On Error GoTo Erro_Move_TabComplemento_Memoria

    'Recolhe os dados de complemento
    objOV.sPlaca = Placa.Text
    objOV.sPlacaUF = PlacaUF.Text
    objOV.sVolumeNumero = VolumeNumero.Text
    objOV.lVolumeQuant = StrParaLong(VolumeQuant.Text)
    objOV.dPesoBruto = StrParaDbl(PesoBruto.Text)
    objOV.dPesoLiq = StrParaDbl(PesoLiquido.Text)
    objOV.dVolumeTotal = StrParaDbl(Cubagem.Text)
    
    'Armazena o responsável pelo frete
    objOV.iFreteRespons = TipoFrete.ItemData(TipoFrete.ListIndex)
    
    objOV.iFilialEntrega = Codigo_Extrai(FilialEntrega.Text)
    
    If Len(Trim(Transportadora.Text)) > 0 Then objOV.iCodTransportadora = Codigo_Extrai(Transportadora.Text)
    If Len(Trim(TranspRedespacho.Text)) > 0 Then objOV.iCodTranspRedesp = Codigo_Extrai(TranspRedespacho.Text)
    objOV.iDetPagFrete = RedespachoCli.Value
    If Len(Trim(PedidoCliente.ClipText)) > 0 Then objOV.sPedidoCliente = PedidoCliente.Text
    If Len(Trim(Mensagem.Text)) > 0 Then objOV.sMensagemNota = Mensagem.Text
    If Len(Trim(CanalVenda.Text)) > 0 Then objOV.iCanalVenda = Codigo_Extrai(CanalVenda.Text)

    'Incluído por Luiz Nogueira em 21/08/03
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objOV.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objOV.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
    
    If Len(Trim(PedidoRepr.ClipText)) > 0 Then objOV.sPedidoRepresentante = PedidoRepr.Text
    
    Move_TabComplemento_Memoria = SUCESSO

    Exit Function

Erro_Move_TabComplemento_Memoria:

    Move_TabComplemento_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211521)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211522)

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
Dim dVolume As Double

On Error GoTo Erro_Pesos_Recalcula

    If gobjFAT.iPesosRecalcula <> 0 Then

    '    lErro = Calcula_VolumeTotal
    '    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        dPesoBruto = 0
        dPesoLiquido = 0
        dVolume = 0
        
        'Para cada linha do Grid de Itens
        For iIndice = 1 To objGridItens.iLinhasExistentes
        
            'Se o produto e a UM estiverem preenchidos
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) > 0 Then
                
                'Formata o produto
                lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                'Lê o produto
                objProduto.sCodigo = sProdutoFormatado
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
                'Se o produto não está cadastrado, erro
                If lErro = 28030 Then gError 211523
                            
                'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), objProduto.sSiglaUMVenda, dFator)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                                    
                dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
                
                'Acumula o peso Bruto e o peso líquido
                dPesoBruto = dPesoBruto + objProduto.dPesoBruto * dFator * dQuantidade
                dPesoLiquido = dPesoLiquido + objProduto.dPesoLiq * dFator * dQuantidade
                            
                'Realiza o cálculo do volume
                dVolume = dVolume + (((StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) * dFator) * (objProduto.dComprimento * objProduto.dEspessura * objProduto.dLargura))
            
            End If
            
        Next
        
        'Coloca os Pesos na tela
        PesoBruto.Text = Format(dPesoBruto, PesoBruto.Format)
        PesoLiquido.Text = Format(dPesoLiquido, PesoLiquido.Format)
        
        If dVolume <> 0 Then
            Cubagem.Text = dVolume
        Else
            Cubagem.Text = ""
        End If
    
    End If
    
    Pesos_Recalcula = SUCESSO
    
    Exit Function
    
Erro_Pesos_Recalcula:

    Pesos_Recalcula = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 211523
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211524)
    
    End Select
    
    Exit Function
    
End Function

Public Sub OptDataPrazoEnt_Click(Index As Integer)

    If OptDataPrazoEnt(1).Value Then
        FrameDataPrazoEnt(1).Enabled = True
        FrameDataPrazoEnt(0).Enabled = False
        FrameDataPrazoEnt(2).Enabled = False
        DataEntregaPV.PromptInclude = False
        DataEntregaPV.Text = ""
        DataEntregaPV.PromptInclude = True
        gobjOrcamentoVenda.dtDataEntrega = DATA_NULA
        PrazoTexto.Text = ""
        If PrazoEntVar.ListIndex = -1 Then PrazoEntVar.ListIndex = 0
    ElseIf OptDataPrazoEnt(0).Value Then
        FrameDataPrazoEnt(0).Enabled = True
        FrameDataPrazoEnt(1).Enabled = False
        FrameDataPrazoEnt(2).Enabled = False
        PrazoEntrega.PromptInclude = False
        PrazoEntrega.Text = ""
        PrazoEntrega.PromptInclude = True
        iPrazoAnt = 0
        PrazoTexto.Text = ""
        PrazoEntVar.ListIndex = -1
    ElseIf OptDataPrazoEnt(2).Value Then
        FrameDataPrazoEnt(1).Enabled = False
        FrameDataPrazoEnt(0).Enabled = False
        FrameDataPrazoEnt(2).Enabled = True
        DataEntregaPV.PromptInclude = False
        DataEntregaPV.Text = ""
        DataEntregaPV.PromptInclude = True
        PrazoEntrega.PromptInclude = False
        PrazoEntrega.Text = ""
        PrazoEntrega.PromptInclude = True
        iPrazoAnt = 0
        PrazoEntVar.ListIndex = -1
    End If

End Sub

Function Saida_Celula_PrazoEntrega(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PrazoEntrega

    Set objGridInt.objControle = PrazoEntregaItem

    If Len(Trim(PrazoEntregaItem.ClipText)) > 0 Then

        GridItens.TextMatrix(GridItens.Row, iGrid_DataEntrega_Col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 211525

    Saida_Celula_PrazoEntrega = SUCESSO

    Exit Function

Erro_Saida_Celula_PrazoEntrega:

    Saida_Celula_PrazoEntrega = gErr

    Select Case gErr

        Case 211525
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211526)

    End Select

    Exit Function

End Function

Public Sub BotaoEstoqueProd_Click()

Dim lErro As Long
Dim objEstoqueProd As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_BotaoEstoqueProd_Click

    lErro = OV_Botao_Estoque(objEstoqueProd, colSelecao, sSelecao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Chama_Tela("EstoqueProdutoLista", colSelecao, objEstoqueProd, objEventoEstoqueProd, sSelecao)
    
    Exit Sub

Erro_BotaoEstoqueProd_Click:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211531)
    
    End Select
    
    Exit Sub
    
End Sub

Function OV_Botao_Estoque(ByVal objEstoqueProd As ClassEstoqueProduto, ByVal colSelecao As Collection, sSelecao As String) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String, sProdutoEnxuto As String
Dim iAlmoxarifado As Integer

On Error GoTo Erro_OV_Botao_Estoque

    sSelecao = ""

    'VerIfica se tem alguma linha selecionada
    If GridItens.Row = 0 Then gError 211531

    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        objEstoqueProd.sProduto = sProdutoFormatado
        
        '*********** INCLUIDO PARA O TRATAMENTO DE ITEM DE GRADE ****************
        If gobjOrcamentoVenda.colItens(GridItens.Row).iPossuiGrade = MARCADO Then
            Call Mascara_RetornaProdutoEnxuto(sProdutoFormatado, sProdutoEnxuto)
            sSelecao = "Produto LIKE '" & sProdutoEnxuto & "%'"
            Call Chama_Tela("EstProdLista_Consulta", colSelecao, objEstoqueProd, objEventoEstoqueProd, sSelecao)
            gError ERRO_SEM_MENSAGEM
        End If

        colSelecao.Add sProdutoFormatado

        lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, sProdutoFormatado, iAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 23796 Then gError ERRO_SEM_MENSAGEM

        If lErro = SUCESSO Then objEstoqueProd.iAlmoxarifado = iAlmoxarifado

    Else
        gError 211532
    End If

    OV_Botao_Estoque = SUCESSO
    
    Exit Function

Erro_OV_Botao_Estoque:

    OV_Botao_Estoque = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 211531
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 211532
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211533)

    End Select

    Exit Function
    
End Function

Private Function Carrega_Moeda() As Long

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 103372
    
    For Each objMoeda In colMoedas
    
        Moeda.AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
        Moeda.ItemData(Moeda.NewIndex) = objMoeda.iCodigo
    
    Next
    
    Call Combo_Seleciona_ItemData(Moeda, MOEDA_REAL)
    iMoedaAnt = Codigo_Extrai(Moeda.Text)

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 103372
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164661)
    
    End Select

End Function

Private Function Carrega_PrazoTexto() As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_PrazoTexto
    
    lErro = CF("Carrega_Combo_Texto", PrazoTexto, "OrcamentoVenda", STRING_MAXIMO, "PrazoTexto")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Carrega_PrazoTexto = SUCESSO
    
    Exit Function
    
Erro_Carrega_PrazoTexto:

    Carrega_PrazoTexto = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164661)
    
    End Select

End Function

Public Sub Moeda_Change()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Moeda
End Sub

Public Sub Moeda_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Moeda
End Sub

Public Sub PrazoTexto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PrazoTexto_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PrazoEntVar_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataPerda_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataPerda_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataPerda, iAlterado)
End Sub

Public Sub DataPerda_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataPerda_Validate

    'Verifica se a Data de Perda foi digitada
    If Len(Trim(DataPerda.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataPerda.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_DataPerda_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211518)

    End Select

    Exit Sub

End Sub

Public Sub UpDownPerda_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownPerda_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataPerda, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownPerda_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211519)

    End Select

    Exit Sub

End Sub

Public Sub UpDownPerda_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPerda_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataPerda, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownPerda_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211520)

    End Select

    Exit Sub

End Sub

Private Sub Trata_Perda()

    If Not gbCarregandoTela Then
        If MotivoPerda.ListIndex <> -1 And lMotivoPerdaAnt <> MotivoPerda.ListIndex Then
            Call DateParaMasked(DataPerda, gdtDataAtual)
        End If
    End If
    lMotivoPerdaAnt = MotivoPerda.ListIndex

End Sub

Private Sub Trata_Moeda()

Dim lErro As Long

On Error GoTo Erro_Trata_Moeda

    If Not gbCarregandoTela Then
        If iMoedaAnt <> Codigo_Extrai(Moeda.Text) Then
    
            lErro = Trata_TabelaPreco()
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
    End If
    iMoedaAnt = Codigo_Extrai(Moeda.Text)

    Exit Sub

Erro_Trata_Moeda:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211520)

    End Select

    Exit Sub

End Sub

Public Function OV_Move_PrecoUnitario_Tela(ByVal objCT As Object, ByVal dPrecoUnitario As Double, ByVal iLinha As Integer) As Long

        If dPrecoUnitario > 0 Then
            GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        Else
            GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = ""
        End If
        
        OV_Move_PrecoUnitario_Tela = SUCESSO

End Function

Public Sub BotaoDesce_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoDesce_Click

    If GridItens.Row < objGridItens.iLinhasExistentes Then
    
        lErro = Trata_Troca_Linha(GridItens.Row, GridItens.Row + 1)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    Exit Sub

Erro_BotaoDesce_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211520)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoSobe_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoSobe_Click

    If GridItens.Row > 1 Then
    
        lErro = Trata_Troca_Linha(GridItens.Row, GridItens.Row - 1)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    Exit Sub

Erro_BotaoSobe_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211520)

    End Select

    Exit Sub
    
End Sub

Private Function Trata_Troca_Linha(ByVal iLinhaAtual As Integer, ByVal iLinhaNova As Integer) As Long

Dim lErro As Long
Dim objOVTela As New ClassOrcamentoVenda
Dim objItemOV As ClassItemOV

On Error GoTo Erro_Trata_Troca_Linha

    lErro = Move_GridItens_Memoria(objOVTela)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Colecao_Altera_Ordem_Itens(objOVTela.colItens, iLinhaAtual, iLinhaNova)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Colecao_Altera_Ordem_Itens(gobjOrcamentoVenda.colItens, iLinhaAtual, iLinhaNova)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Colecao_Altera_Ordem_Itens(gcolItensOVPrecoCalc, iLinhaAtual, iLinhaNova)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = gobjTelaProjetoInfo.Troca_Linha(iLinhaAtual, iLinhaNova)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Carrega_Grid_Itens(objOVTela, VALOR_CODIGO)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    GridItens.Row = iLinhaNova
    GridItens.Col = iGrid_Produto_Col
    Call GridItens_Click

    Trata_Troca_Linha = SUCESSO
    
    Exit Function
    
Erro_Trata_Troca_Linha:

    Trata_Troca_Linha = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 164661)
    
    End Select
    
End Function

Public Sub SomaItem_Click()
    If iGrid_SomaItem_Col <> 0 Then
        Call Soma_Coluna_Grid(objGridItens, iGrid_PrecoTotal_Col, TS, False, iGrid_SomaItem_Col)
    End If
End Sub

Private Function Saida_Celula_PrecoTotal(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Total que está deixando de ser a corrente

Dim lErro As Long
Dim dPrecoAnt As Double, dPrecoAtual As Double
Dim dQuantidade As Double, dPrecoUnit As Double
Dim dPrecoBrutoCalc As Double, dDesconto As Double
Dim bPrecoTotalIgual As Boolean, dPrecoTotal As Double, dPercentDesc As Double

On Error GoTo Erro_Saida_Celula_PrecoTotal

    bPrecoTotalIgual = False

    Set objGridInt.objControle = PrecoTotal

    If Len(Trim(PrecoTotal.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(PrecoTotal.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    dPrecoAnt = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))
    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    dPrecoUnit = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))
    dDesconto = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col))
    dPrecoAtual = StrParaDbl(PrecoTotal.Text)
    
    dPrecoBrutoCalc = Arredonda_Moeda(dQuantidade * dPrecoUnit)

    'Comparação com Preço Unitário anterior
    If Abs(dPrecoAnt - dPrecoAtual) < DELTA_VALORMONETARIO Then
        bPrecoTotalIgual = True
    Else
        bPrecoTotalIgual = False
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If Not bPrecoTotalIgual Then
        
        If dPrecoBrutoCalc < dPrecoAtual Then
            GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col) = Format(dPrecoAtual / dQuantidade, gobjFAT.sFormatoPrecoUnitario)
            GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
            GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = ""
            
            lErro = CF2(Me, "OV_Saida_Celula_PrecoUnitario", objGridInt)
            If lErro <> SUCESSO Then gError 84172
            
        Else
            GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = Format(dPrecoBrutoCalc - dPrecoAtual, "STANDARD")
            
            dPrecoTotal = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotalB_Col))
            dDesconto = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col))
            
            'Se o Preço total é positivo
            If dPrecoTotal > 0 Then
                'Recalcula o percentual de desconto
                dPercentDesc = dDesconto / dPrecoTotal
                GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")
            End If
        End If
        
        Call PrecoTotal_Calcula(GridItens.Row)
        
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        lErro = FormPreco_Calcula()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

   Saida_Celula_PrecoTotal = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoTotal:

    Saida_Celula_PrecoTotal = gErr


    Select Case gErr

        Case 84170, 84171, 84172
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157794)

    End Select

    Exit Function

End Function

Public Sub DataFechamento_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataFechamento_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataFechamento, iAlterado)
End Sub

Public Sub DataFechamento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFechamento_Validate

    'Verifica se a Data de Fechamento foi digitada
    If Len(Trim(DataFechamento.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataFechamento.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_DataFechamento_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211518)

    End Select

    Exit Sub

End Sub

Public Sub UpDownFechamento_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownFechamento_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataFechamento, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownFechamento_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211519)

    End Select

    Exit Sub

End Sub

Public Sub UpDownFechamento_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownFechamento_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataFechamento, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownFechamento_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211520)

    End Select

    Exit Sub

End Sub

Public Sub DataProxContato_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataProxContato_GotFocus()
     Call MaskEdBox_TrataGotFocus(DataProxContato, iAlterado)
End Sub

Public Sub DataProxContato_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataProxContato_Validate

    'Verifica se a Data de ProxContato foi digitada
    If Len(Trim(DataProxContato.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataProxContato.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_DataProxContato_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211518)

    End Select

    Exit Sub

End Sub

Public Sub UpDownProxContato_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownProxContato_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataProxContato, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownProxContato_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211519)

    End Select

    Exit Sub

End Sub

Public Sub UpDownProxContato_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownProxContato_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataProxContato, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownProxContato_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211520)

    End Select

    Exit Sub

End Sub

Public Function OV_ValorTotal_Calcula(ByVal objCT As Object) As Long
    OV_ValorTotal_Calcula = SUCESSO
End Function

Public Function OV_Limpa_OrcamentoVenda2(ByVal objCT As Object) As Long
    OV_Limpa_OrcamentoVenda2 = SUCESSO
End Function

Public Function OV_Saida_Celula_Quantidade(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    OV_Saida_Celula_Quantidade = SUCESSO
End Function

Public Function OV_Saida_Celula_PrecoUnitario(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    OV_Saida_Celula_PrecoUnitario = SUCESSO
End Function

Function OV_Traz_OrcamentoVenda_Tela(ByVal objCT As Object, objOrcamentoVenda As ClassOrcamentoVenda, ByVal iBrowse As Integer) As Long
    OV_Traz_OrcamentoVenda_Tela = SUCESSO
End Function

Public Sub BotaoAnaliseVenda_Click()

Dim objAnaliseVendas As New ClassAnaliseVendaInfo
Dim objOrcamentoVendas As New ClassOrcamentoVenda
Dim lErro As Long

On Error GoTo Erro_BotaoAnaliseVenda_Click

    Call Move_OrcamentoVenda_Memoria(objOrcamentoVendas)

    If objOrcamentoVendas.colItens.Count = 0 Then gError 201541

    Call objOrcamentoVendas.AnaliseVenda_Preenche(objAnaliseVendas)
    
    Call CF2(Me, "OV_AnaliseVenda_Preenche", objOrcamentoVendas, objAnaliseVendas)

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

Public Function OV_AnaliseVenda_Preenche(ByVal objCT As Object, ByVal objOrcamentoVenda As ClassOrcamentoVenda, ByVal objAnaliseVendas As ClassAnaliseVendaInfo) As Long
    OV_AnaliseVenda_Preenche = SUCESSO
End Function

Public Sub Idioma_Change()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Idioma
End Sub

Public Sub Idioma_Click()
    iAlterado = REGISTRO_ALTERADO
    Call Trata_Idioma
End Sub

Private Sub Trata_Idioma()

Dim lErro As Long
Dim iLinha As Integer
Dim sProduto As String, iPreenchido As Integer
Dim sRetorno As String

On Error GoTo Erro_Trata_Idioma

    If Not gbCarregandoTela Then
    
        If iIdiomaAnt <> Codigo_Extrai(Idioma.Text) Then
    
            If Not (objGridItens Is Nothing) Then
            
                For iLinha = 1 To objGridItens.iLinhasExistentes
                    
                    'Verifica se o Produto está preenchido
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                    If iPreenchido <> PRODUTO_VAZIO Then
                    
                        lErro = gobjIdiomaTela.Traduzir("Produtos", "Descricao", GridItens.TextMatrix(iLinha, iGrid_DescProduto_Col), sProduto, 0, 0, Codigo_Extrai(Idioma.Text), sRetorno)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                        GridItens.TextMatrix(iLinha, iGrid_DescProduto_Col) = sRetorno
                    
                    End If
                    
                Next
        
            End If
        
        End If
    
    End If
    
    iIdiomaAnt = Codigo_Extrai(Idioma.Text)

    Exit Sub

Erro_Trata_Idioma:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211520)

    End Select

    Exit Sub

End Sub

Public Sub BtnExportaDoc_Click()

Dim lErro  As Long
Dim objOV As New ClassOrcamentoVenda

On Error GoTo Erro_BtnExportaDoc_Click

    'Guarda no obj o código do orçamento passado como parâmetro
    objOV.lCodigo = StrParaLong(Codigo.Text)
    
    'Guarda a FilialEmpresa ativa como filial do orçamento
    objOV.iFilialEmpresa = giFilialEmpresa
    objOV.iVersao = StrParaInt(OVVersao.Caption)
    
    'Lê os dados do orçamento para verificar se o mesmo existe no BD
    lErro = CF("OrcamentoVenda_Le", objOV)
    If lErro <> SUCESSO And lErro <> 101232 Then gError ERRO_SEM_MENSAGEM

    'Se não encontrou => erro, pois não é possível imprimir um orçamento inexistente
    If lErro = 101232 Then gError 102236
    
    'Lê a parte de Tributação
    lErro = CF("OrcamentoVenda_Le_Tributacao", objOV)
    If lErro <> SUCESSO And lErro <> 101162 Then gError ERRO_SEM_MENSAGEM

    lErro = CF("OrcamentoVenda_Le_Itens_ComTributacao", objOV)
    If lErro <> SUCESSO And lErro <> 101278 Then gError ERRO_SEM_MENSAGEM
    
    lErro = CF("ParcelasOV_Le", objOV)
    If lErro <> SUCESSO And lErro <> 101284 Then gError ERRO_SEM_MENSAGEM
               
    Call Chama_Tela_Modal("ExportaParaDoc", MNEMONICO_MALADIRETA_TIPO_OV, objOV)
    
    Exit Sub

Erro_BtnExportaDoc_Click:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case 102236
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO", gErr, objOV.lCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157838)
            
    End Select
    
    Exit Sub

End Sub

Public Function OV_Cliente_Validate(ByVal objCT As Object, ByVal objCli As ClassCliente) As Long
    OV_Cliente_Validate = SUCESSO
End Function
