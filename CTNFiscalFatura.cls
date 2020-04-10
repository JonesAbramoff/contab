VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalFatura"
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

Dim giNaoRepeteEvento As String

Private WithEvents objEventoOrcamento As AdmEvento
Attribute objEventoOrcamento.VB_VarHelpID = -1

Public gobjTelaProjetoInfo As ClassTelaPRJInfo

'estrutura mantida para auxiliar a manutencao de informacoes p/tributacao
Public gobjNFiscal As ClassNFiscal
Dim giFilialAlterada As Integer
Public giValorDescontoManual As Integer 'Alterado por Wagner
Dim giFilialEntregaManual As Integer
Dim giClienteAlterado As Integer
Dim giComplemento As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

'Flag que indica se a tela está sendo preenchida.
Public gbCarregandoTela As Boolean
Public gbLimpandoTela As Boolean
Public giPosCargaOk As Integer

'inicia objeto associado a GridComissoes
Public objTabComissoes As New ClassTabComissoes

'inicio contabilidade
Public objGrid1 As AdmGrid
Public objContabil As New ClassContabil
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
Private Const UNIDADE_MED As String = "Unidade_Med"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const PRECO_UNITARIO As String = "Preco_Unitario"
Private Const PRECO_TOTAL As String = "Preco_Total"
Private Const DESCONTO1 As String = "Desconto_Produto"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"
Private Const CTACONTABILEST1 As String = "ContaContabilEst"
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
'###########Pagamento###############
Private Const PERC_ACRESC_FIN As String = "Perc_Acresc_Fin"
Private Const VALOR_PARCELA As String = "Valor_Parcela"
'###########Tributação##############
Private Const ISS_VALOR As String = "ISS_Valor"
Private Const ISS_INCLUSO As String = "ISS_Incluso"
Private Const VALOR_IRRF As String = "Valor_IRRF"
Private Const QUANT_DISPONIVEL As String = "Quant_Disponivel"
Private Const QUANT_CONSIGTERC As String = "Quant_ConsigTerc"
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

Public iAlterado As Integer
Public iPercAcresFinAlterado As Integer
Public iFrameAtual As Integer
Public gdDesconto As Double
Public iClienteAlterado As Integer
Public iValorDescontoAlterado As Integer  'É usado em ValorTotal_Change, CondicaoPagamento_Validate, GridParcelas_Preenche
Public iValorFreteAlterado As Integer
Public iValorSeguroAlterado As Integer
Public iValorDespesasAlterado As Integer
Public iDataEmissaoAlterada As Integer
Public iSerieAlterada As Integer
Public giFilialEntregaAlterada As Integer
Public giDataReferenciaAlterada As Integer

'declara o objeto da execucao das regras de comissoes e
'o objeto do calculo de comissoes
Public objComissoesRegrasCalc As Object
Public objMnemonicoComissCalc As ClassMnemonicoComissCalc
Public objMnemonicoComissCalcAux As ClassMnemonicoComissCalcAux

'Indica que foi alterado pelo menos um campo na tela que seja utilizado para calcular as comissões
Public iComissoesAlterada As Integer

Public objGridParcelas As AdmGrid
Dim iGrid_Vencimento_Col As Integer
Dim iGrid_ValorParcela_Col As Integer
Dim iGrid_Desc1Codigo_Col As Integer
Dim iGrid_Desc1Ate_Col As Integer
Dim iGrid_Desc1Valor_Col As Integer
Dim iGrid_Desc1Percentual_Col As Integer
Dim iGrid_Desc2Codigo_Col As Integer
Dim iGrid_Desc2Ate_Col As Integer
Dim iGrid_Desc2Valor_Col As Integer
Dim iGrid_Desc2Percentual_Col As Integer
Dim iGrid_Desc3Codigo_Col As Integer
Dim iGrid_Desc3Ate_Col As Integer
Dim iGrid_Desc3Valor_Col As Integer
Dim iGrid_Desc3Percentual_Col As Integer

Public objGridItens As AdmGrid
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
Public iGrid_Contrato_Col As Integer
Public iGrid_ItemContrato_Col As Integer
Public iGrid_DataCobranca_Col As Integer
Public iGrid_SerieNFOrig_Col As Integer
Public iGrid_NumNFOrig_Col As Integer
Public iGrid_ItemNFOrig_Col As Integer
Public iGrid_TabPrecoItem_Col As Integer
Public iGrid_ComissaoItem_Col As Integer

Public objGridComissoes As AdmGrid

Public objGridAlocacoes As AdmGrid
Public iGrid_Item_Col As Integer
Public iGrid_ProdutoAloc_Col As Integer
Public iGrid_AlmoxAloc_Col As Integer
Public iGrid_UMAloc_Col As Integer
Public iGrid_QuantAloc_Col As Integer
Public iGrid_QuantVend_Col As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoCondPagto As AdmEvento
Attribute objEventoCondPagto.VB_VarHelpID = -1
Private WithEvents objEventoEstoqueProd As AdmEvento
Attribute objEventoEstoqueProd.VB_VarHelpID = -1
Public WithEvents objEventoNFiscal As AdmEvento
Attribute objEventoNFiscal.VB_VarHelpID = -1
Private WithEvents objEventoNFiscalOriginal As AdmEvento
Attribute objEventoNFiscalOriginal.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Private WithEvents objEventoSerieOriginal As AdmEvento
Attribute objEventoSerieOriginal.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoTranspRedespacho As AdmEvento
Attribute objEventoTranspRedespacho.VB_VarHelpID = -1
Public WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoMensagem As AdmEvento
Attribute objEventoMensagem.VB_VarHelpID = -1
Private WithEvents objEventoContrato As AdmEvento
Attribute objEventoContrato.VB_VarHelpID = -1

''*** incluidos p/tratamento de tributacao *******************************
Public gobjTribTab As New ClassTribTab

Public gobjRastreamento As New ClassRastreamento

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

Private WithEvents objEventoItemContrato As AdmEvento
Attribute objEventoItemContrato.VB_VarHelpID = -1
Private WithEvents objEventoMedicao As AdmEvento
Attribute objEventoMedicao.VB_VarHelpID = -1

Dim gcolcolMedicoes As Collection

Const TIPODOCINFO_COMPLEMENTO_IPI = 53

Const BOTAO_ORCAMENTO = 0
Const BOTAO_VERORCAMENTO = 1
Const BOTAO_CONSULTATITREC = 2
Const BOTAO_GRAVAR = 3
Const BOTAO_EXCLUIR = 4
Const BOTAO_LIMPAR = 5
Const BOTAO_FECHAR = 6
Const BOTAO_EXIBIRDADOS = 7
Const BOTAO_GRADE = 8
Const BOTAO_KITVENDA = 9
Const BOTAO_PRODUTOS = 10
Const BOTAO_ITEMCONTRATO = 11
Const BOTAO_MEDICAO = 12
Const BOTAO_DOCCONTRATO = 13
Const BOTAO_ESTOQUEPROD = 14
'Const BOTAO_VENDEDORES = 15
Const BOTAO_LOCALIZACAO = 16
Const BOTAO_SERIE = 17
Const BOTAO_LOTES = 18
Const BOTAO_PROJETOS = 19
Const BOTAO_LIMPARNF = 20
Const BOTAO_CONTRATO = 21
Const BOTAO_VERCONTRATO = 22
Const BOTAO_CONTRATOORCAMENTO = 999

'--- inicio dos properties get dos controles da tela
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

Public Property Get FrameCliente() As Object
     Set FrameCliente = objUserControl.Controls("FrameCliente")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Fornecedor() As Object
     Set Fornecedor = objUserControl.Controls("Fornecedor")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get ClienteLabel() As Object
     Set ClienteLabel = objUserControl.Controls("ClienteLabel")
End Property

Public Property Get FornecedorLabel() As Object
     Set FornecedorLabel = objUserControl.Controls("FornecedorLabel")
End Property

Public Property Get FramePrecos() As Object
     Set FramePrecos = objUserControl.Controls("FramePrecos")
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

Public Property Get FrameDatas() As Object
     Set FrameDatas = objUserControl.Controls("FrameDatas")
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

Public Property Get FrameIdentificacao() As Object
     Set FrameIdentificacao = objUserControl.Controls("FrameIdentificacao")
End Property

Public Property Get BotaoLimparNF() As Object
     Set BotaoLimparNF = objUserControl.Controls("Botao").Item(BOTAO_LIMPARNF)
End Property

Public Property Get TipoNFiscal() As Object
     Set TipoNFiscal = objUserControl.Controls("TipoNFiscal")
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

Public Property Get ImageStatus() As Object
     Set ImageStatus = objUserControl.Controls("ImageStatus")
End Property

Public Property Get FrameNFiscalOriginal() As Object
     Set FrameNFiscalOriginal = objUserControl.Controls("FrameNFiscalOriginal")
End Property

Public Property Get NFiscalPropria() As Object
     Set NFiscalPropria = objUserControl.Controls("NFiscalPropria")
End Property

Public Property Get SerieNFiscalOriginal() As Object
     Set SerieNFiscalOriginal = objUserControl.Controls("SerieNFiscalOriginal")
End Property

Public Property Get BotaoExibirDados() As Object
     Set BotaoExibirDados = objUserControl.Controls("Botao").Item(BOTAO_EXIBIRDADOS)
End Property

Public Property Get NFiscalOriginal() As Object
     Set NFiscalOriginal = objUserControl.Controls("NFiscalOriginal")
End Property

Public Property Get SerieOriginalLabel() As Object
     Set SerieOriginalLabel = objUserControl.Controls("SerieOriginalLabel")
End Property

Public Property Get NFOriginalLabel() As Object
     Set NFOriginalLabel = objUserControl.Controls("NFOriginalLabel")
End Property

Public Property Get CobrancaAutomatica() As Object
     Set CobrancaAutomatica = objUserControl.Controls("CobrancaAutomatica")
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

Public Property Get BotaoDataReferenciaDown() As Object
     Set BotaoDataReferenciaDown = objUserControl.Controls("BotaoDataReferenciaDown")
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

Public Property Get BotaoConsultaTitRec() As Object
     Set BotaoConsultaTitRec = objUserControl.Controls("Botao").Item(BOTAO_CONSULTATITREC)
End Property

Public Property Get ProdutoAlmox() As Object
     Set ProdutoAlmox = objUserControl.Controls("ProdutoAlmox")
End Property

Public Property Get GridAlocacao() As Object
     Set GridAlocacao = objUserControl.Controls("GridAlocacao")
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

Public Property Get BotaoLocalizacao() As Object
     Set BotaoLocalizacao = objUserControl.Controls("Botao").Item(BOTAO_LOCALIZACAO)
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

Public Property Get TipoFrete() As Object
     Set TipoFrete = objUserControl.Controls("TipoFrete")
End Property

Public Property Get ImprimeRomaneio() As Object
     Set ImprimeRomaneio = objUserControl.Controls("ImprimeRomaneio")
End Property

Public Property Get FilialEntrega() As Object
     Set FilialEntrega = objUserControl.Controls("FilialEntrega")
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

Public Property Get LabelPedidoTerc() As Object
     Set LabelPedidoTerc = objUserControl.Controls("LabelPedidoTerc")
End Property

Public Property Get MensagemLabel() As Object
     Set MensagemLabel = objUserControl.Controls("MensagemLabel")
End Property

Public Property Get BotaoItensNFiscalOrig() As Object
     Set BotaoItensNFiscalOrig = objUserControl.Controls("BotaoItensNFiscalOrig")
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

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("Botao").Item(BOTAO_PRODUTOS)
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

Public Property Get BotaoEstoqueProd() As Object
     Set BotaoEstoqueProd = objUserControl.Controls("Botao").Item(BOTAO_ESTOQUEPROD)
End Property

Public Property Get CTBSeqContraPartida() As Object
     Set CTBSeqContraPartida = objUserControl.Controls("CTBSeqContraPartida")
End Property

Public Property Get CTBListHistoricos() As Object
     Set CTBListHistoricos = objUserControl.Controls("CTBListHistoricos")
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

Public Property Get CTBHistorico() As Object
     Set CTBHistorico = objUserControl.Controls("CTBHistorico")
End Property

Public Property Get CTBAglutina() As Object
     Set CTBAglutina = objUserControl.Controls("CTBAglutina")
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

Public Property Get CTBTvwContas() As Object
     Set CTBTvwContas = objUserControl.Controls("CTBTvwContas")
End Property

Public Property Get CTBGridContabil() As Object
     Set CTBGridContabil = objUserControl.Controls("CTBGridContabil")
End Property

Public Property Get CTBTvwCcls() As Object
     Set CTBTvwCcls = objUserControl.Controls("CTBTvwCcls")
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

Public Property Get FrameTributacao() As Object
     Set FrameTributacao = objUserControl.Controls("FrameTributacao")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("Botao").Item(BOTAO_GRAVAR)
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("Botao").Item(BOTAO_LIMPAR)
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("Botao").Item(BOTAO_FECHAR)
End Property

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

Public Property Get Status() As Object
    Set Status = objUserControl.Controls("Status")
End Property

Public Property Get LabelFilial() As Object
    Set LabelFilial = objUserControl.Controls("LabelFilial")
End Property

Public Property Get LabelFilialEntrega() As Object
    Set LabelFilialEntrega = objUserControl.Controls("LabelFilialEntrega")
End Property

Public Property Get FilialL() As Object
    Set FilialL = objUserControl.Controls("FilialL")
End Property

Public Sub ImprimeRomaneio_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("Botao").Item(BOTAO_EXCLUIR)
End Property

'Rastreamento
Public Property Get GridRastro() As Object
     Set GridRastro = objUserControl.Controls("GridRastro")
End Property

'Rastreamento
Public Property Get EscaninhoRastro() As Object
     Set EscaninhoRastro = objUserControl.Controls("EscaninhoRastro")
End Property

Public Property Get BotaoMedicao() As Object
     Set BotaoMedicao = objUserControl.Controls("Botao").Item(BOTAO_MEDICAO)
End Property

Public Property Get BotaoItemContrato() As Object
     Set BotaoItemContrato = objUserControl.Controls("Botao").Item(BOTAO_ITEMCONTRATO)
End Property

Public Property Get Item() As Object
     Set Item = objUserControl.Controls("Item")
End Property

Public Property Get Contrato() As Object
     Set Contrato = objUserControl.Controls("Contrato")
End Property

Public Property Get DataCobranca() As Object
     Set DataCobranca = objUserControl.Controls("DataCobranca")
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

Public Property Get ContratoLabel() As Object
     Set ContratoLabel = objUserControl.Controls("ContratoLabel")
End Property

Public Property Get ContratoFat() As Object
     Set ContratoFat = objUserControl.Controls("ContratoFat")
End Property

Public Property Get BotaoContratoOrcamento() As Object
     Set BotaoContratoOrcamento = objUserControl.Controls("Botao").Item(BOTAO_CONTRATOORCAMENTO)
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

Public Property Get SerieNFOriginal() As Object
     Set SerieNFOriginal = objUserControl.Controls("SerieNFOriginal")
End Property

Public Property Get NumNFOriginal() As Object
     Set NumNFOriginal = objUserControl.Controls("NumNFOriginal")
End Property

Public Property Get ItemNFOriginal() As Object
     Set ItemNFOriginal = objUserControl.Controls("ItemNFOriginal")
End Property

Public Property Get EletronicaFed() As Object
     Set EletronicaFed = objUserControl.Controls("EletronicaFed")
End Property

Public Property Get NFiscalAux() As Object
     Set NFiscalAux = objUserControl.Controls("NFiscalAux")
End Property

'--- fim dos properties get dos controles da tela

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
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157148)

    End Select

    Exit Sub

End Sub

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
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157150)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimparNF_Click()

    NFiscal.Caption = ""
    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
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

Public Sub ImageStatus_Click(Index As Integer)

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

Public Sub NFiscal_GotFocus()
    Call MaskEdBox_TrataGotFocus(NFiscal, iAlterado)
End Sub

Public Sub NFiscalOriginal_GotFocus()
    Call MaskEdBox_TrataGotFocus(NFiscalOriginal, iAlterado)
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
        If lErro <> SUCESSO Then gError 46190
    End If

    If Len(Trim(TabelaPreco.Text)) > 0 Then

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 51620
    
    End If
    
    iPercAcresFinAlterado = 0
    
    Exit Sub
    
Erro_PercAcrescFin_Validate:

    Cancel = True

    Select Case gErr
    
        Case 46190

        Case 51620
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157152)
            
    End Select
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Call CF2(Me, "NFF_Form_Unload", Cancel)

    Set gobjInfoUsu = Nothing
    
    ' Liberar as variáveis globais
    Set gobjNFiscal = Nothing
    Set objEventoOrcamento = Nothing
    Set objGridParcelas = Nothing
    Set objGridItens = Nothing
    Set objGridComissoes = Nothing
    Set objGridAlocacoes = Nothing
    Set gobjInfoAdic = Nothing
    
    Set objEventoMensagem = Nothing
    Set objEventoCliente = Nothing
    Set objEventoCondPagto = Nothing
    Set objEventoEstoqueProd = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoNFiscalOriginal = Nothing
    Set objEventoProduto = Nothing
    Set objEventoSerie = Nothing
    Set objEventoSerieOriginal = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoTranspRedespacho = Nothing
    Set objEventoVendedor = Nothing
    Set objEventoContrato = Nothing
    Set objEventoItemContrato = Nothing
    Set gcolcolMedicoes = Nothing
    Set gobjTelaProjetoInfo = Nothing
    
    'se foi instanciada a classe de execucao de regras de comissoes => libera
    Set objComissoesRegrasCalc = Nothing
        
    'se foi instanciada a classe que calcula os mnemonicos => libera
    Set objMnemonicoComissCalc = Nothing
    Set objMnemonicoComissCalcAux = Nothing

    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
    'Eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    Set gobjRastreamento = Nothing
        
    'Fecha o comando de seta se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
    Set gobjAnotacao = Nothing
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157153)
    
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

Public Sub BotaoFechar_Click()
    Unload Me
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

Public Sub Cliente_Change()

Dim lErro As Long

On Error GoTo Erro_Cliente_Change

        'Seta iComissoesAlterada
        iComissoesAlterada = REGISTRO_ALTERADO

        iAlterado = REGISTRO_ALTERADO
        iClienteAlterado = REGISTRO_ALTERADO
        
        Call Cliente_Preenche
    
    Exit Sub

Erro_Cliente_Change:

    Select Case gErr

        Case 71653

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157154)

    End Select
    
    Exit Sub

End Sub

Public Sub CobrancaAutomatica_Click()

    iAlterado = REGISTRO_ALTERADO
    
    If CobrancaAutomatica.Value = vbChecked And Len(Trim(CondicaoPagamento.Text)) > 0 Then
        Call Cobranca_Automatica
    End If

End Sub

Public Sub ComissaoAutomatica_Click()

Dim lErro As Long

On Error GoTo Erro_ComissaoAutomatica_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Se a comissão automática estiver selecionada
    If ComissaoAutomatica.Value = vbChecked Then
        
        'Recalcula as comissoes
        lErro = objTabComissoes.Comissoes_Calcula()
        If lErro <> SUCESSO Then gError 51616
        
    End If
    
    Exit Sub

Erro_ComissaoAutomatica_Click:

    Select Case gErr
    
        Case 51616
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157155)
            
    End Select
    
    Exit Sub
End Sub

Public Sub CondicaoPagamento_Change()
    iAlterado = REGISTRO_ALTERADO
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

Public Sub FilialEntrega_Change()

    iAlterado = REGISTRO_ALTERADO
    giFilialEntregaAlterada = 1

End Sub

Public Sub FilialEntrega_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

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

        Call GridItens_InicializaColunas(Me)
        
        ''*** incluidos p/tratamento de tributacao *******************************
        Set gobjTribTab = New ClassTribTab
        lErro = gobjTribTab.Ativar(Me)
        If lErro <> SUCESSO Then gError 39030
    
        '*** fim tributacao
        
        'Carrega os Tipos de Notas Fiscais
        lErro = Carrega_TiposDocInfo()
        If lErro <> SUCESSO Then gError 42111
        
        'Carrega as Tabelas de Preço
        lErro = Carrega_TabelaPreco()
        If lErro <> SUCESSO Then gError 42112
    
        'Carrega as Transportadoras
        lErro = Carrega_Transportadora()
        If lErro <> SUCESSO Then gError 42113
    
        'Carrega a combo VolumeEspecie
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
        If lErro <> SUCESSO Then gError 102414
    
        'Carrega a combo VolumeMarca
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
        If lErro <> SUCESSO Then gError 102415
    
        'Carrega os Canais de Venda
        lErro = Carrega_CanalVenda()
        If lErro <> SUCESSO Then gError 42114
    
        'Carrega as Séries
        lErro = Carrega_Serie()
        If lErro <> SUCESSO Then gError 42115
    
        'Seleciona série padrão
        lErro = Seleciona_Serie_Padrao()
        If lErro <> SUCESSO Then gError 500137
         
        'Carrega na combo as Condições de Pagamento
        lErro = CF("Carrega_CondicaoPagamento", CondicaoPagamento, MODULO_CONTASARECEBER)
        If lErro <> SUCESSO Then gError 42123
    
        'Carrega os Estados
        lErro = Carrega_PlacaUF()
        If lErro <> SUCESSO Then gError 42116
    
        'Carrega os Tipos de Desconto
        lErro = Carrega_TipoDesconto()
        If lErro <> SUCESSO Then gError 42125
        
        PrecoUnitario.Format = gobjFAT.sFormatoPrecoUnitario
        
        Quantidade.Format = FORMATO_ESTOQUE
        QuantAlocada.Format = FORMATO_ESTOQUE
        QuantVendida.Format = FORMATO_ESTOQUE
    
        'Seta as Variáveis de Grid
        Set objGridItens = New AdmGrid
        Set objGridComissoes = New AdmGrid
        Set objGridParcelas = New AdmGrid
        Set objGridAlocacoes = New AdmGrid
        
        'Seta as Variáveis das Telas de browse
        Set objEventoMensagem = New AdmEvento
        Set objEventoCliente = New AdmEvento
        Set objEventoCondPagto = New AdmEvento
        Set objEventoEstoqueProd = New AdmEvento
        Set objEventoNFiscal = New AdmEvento
        Set objEventoNFiscalOriginal = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoSerie = New AdmEvento
        Set objEventoSerieOriginal = New AdmEvento
        Set objEventoTransportadora = New AdmEvento
        Set objEventoTranspRedespacho = New AdmEvento
        Set objEventoVendedor = New AdmEvento
        Set objEventoContrato = New AdmEvento
        giDataReferenciaAlterada = 0
        Set gcolcolMedicoes = New Collection
        Set objEventoItemContrato = New AdmEvento
        
        lErro = CF2(Me, "NFF_Form_Load")
        If lErro <> SUCESSO Then gError 185056
        
        Set objEventoOrcamento = New AdmEvento
        
        'Inicializa a máscara do Produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
        If lErro <> SUCESSO Then gError 42127
        
        'Inicializa a máscara do produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoAlmox)
        If lErro <> SUCESSO Then gError 42127
        
        'Inicializa o Grid de Ítens
        lErro = Inicializa_Grid_Itens(objGridItens)
        If lErro <> SUCESSO Then gError 42126
    
'        'Seleciona a única tabela de preço. Importante: GridItens já deve estar definido
'        If giTipoVersao = VERSAO_LIGHT And TabelaPreco.ListCount > 0 Then
'            TabelaPreco.ListIndex = 0
'        End If

        'Inicializa o Grid de Parcelas
        lErro = Inicializa_Grid_Parcelas(objGridParcelas)
        If lErro <> SUCESSO Then gError 42128

        'Inicializa o Grid de Comissões
        lErro = objTabComissoes.Inicializa_Grid_Comissoes(objGridComissoes)
        If lErro <> SUCESSO Then gError 42129

        'Inicializa o Grid de Alocações
        lErro = Inicializa_Grid_Alocacoes(objGridAlocacoes)
        If lErro <> SUCESSO Then gError 42130

        'Inicializa a contabilidade carregando apenas os modelos contábeis referentes ao modelo padrão
        If objContabil.objGridInt Is Nothing Then

            'Obtém o código do subtipo contábil referente ao TipoDocInfo default
            lErro = CF("Obtem_Subtipo_Contabil", "TiposDocInfo", DOCINFO_NFISFV, iSubTipo)
            If lErro <> SUCESSO And lErro <> 79671 Then gError 79682
            
            'Se não encontrou o subtipo => erro
            If lErro = 79671 Then gError 79683
            
            'Inicialização da parte de contabilidade
            lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_FATURAMENTO, iSubTipo)
            If lErro <> SUCESSO Then gError 39672

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
        
        Case 39030, 39672, 42111, 42112, 42113, 42114, 42123, 42125, 42126, 42127, 42128, _
             42129, 42130, 42115, 42116, 79682, 83383, 500137, 102414, 102415
        
        Case 79683
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPODOCINFO_NAO_ENCONTRADO", gErr, DOCINFO_NFISFV)

        Case 185056

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157156)
     
    End Select
     
    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        
        FrameOrcVenda.Visible = False
        Botao(BOTAO_GRADE).Visible = False
        BotaoContratoOrcamento.Visible = False
        Frame1(303).Visible = True 'frame de contratos na 1a aba no mesmo local do frame de orcvenda
        
        'sumir com abas de contabilidade e rastro
        TabStrip1.Tabs.Remove (TAB_RASTRO)
        TabStrip1.Tabs.Remove (TAB_Contabilizacao)
        
    End If
        
    Set objTabComissoes.objTela = Me
    iFrameAtual = 1
    
    'Preenche Data Referencia e Data de Emissão coma Data Atual
    DataReferencia.PromptInclude = False
    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataReferencia.PromptInclude = True
    
    Call CF2(Me, "NFCF_Default_DataSaida")
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    Call Carrega_FilialOrcamento
        
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        Filial.Visible = False
'        FilialEntrega.Visible = False
'        FilialL.Visible = False
'        LabelFilial.Visible = False
'        LabelFilialEntrega.Visible = False
'
'    End If
    
    Call CF2(Me, "NFF_Form_Load2")
    
    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207433
    
    iAlterado = 0
    
    Set gobjNFiscal = New ClassNFiscal
    
    MsgAutomatica.Value = vbChecked
    Mensagem.MaxLength = MENSAGEM_NF_MAXLENGTH
    MensagemCorpo.MaxLength = MENSAGEM_NF_MAXLENGTH
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 207433

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157157)

    End Select
    
    iAlterado = 0
    iComissoesAlterada = 0
    
    Exit Sub
    

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

Public Sub Serie_Change()
    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = REGISTRO_ALTERADO
End Sub

Public Sub Serie_Click()
    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = REGISTRO_ALTERADO
End Sub

Public Sub SerieNFiscalOriginal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SerieNFiscalOriginal_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TabelaPreco_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoNFiscal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Transportadora_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Transportadora_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UnidadeMed_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

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

Public Sub ValorProdutos_Change()

'    If Not gbCarregandoTela Then
'         Call DescontoGlobal_Recalcula
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

'    If Len(Trim(ValorProdutos.Caption)) <> 0 And IsNumeric(ValorProdutos.Caption) Then
'
'        dValorProdutos = CDbl(ValorProdutos.Caption)
'
'        'Se o cliente possui desconto e o campo desconto não foi alterado pelo usuário
'        If gdDesconto > 0 And iValorDescontoAlterado = 0 Then
'
'            Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)
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

Private Function Cobranca_Automatica() As Long
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
        If lErro <> SUCESSO And lErro <> 19205 Then gError 46187
        If lErro <> SUCESSO Then gError 46188
        'Preenche o grid parcelas de acordo com a condição de pagamento
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 157159)
    
    End Select
    
    Exit Function
        
End Function

Public Sub ValorSeguro_Change()

    'Seta iComissoesAlterada
    iComissoesAlterada = REGISTRO_ALTERADO

    iAlterado = REGISTRO_ALTERADO
    iValorSeguroAlterado = REGISTRO_ALTERADO

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157160)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157161)

    End Select

End Sub

Public Sub VolumeNumero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub VolumeQuant_Change()

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
                If objTipoDocInfo.iPadrao = Padrao Then TipoNFiscal.ListIndex = TipoNFiscal.NewIndex
            End If
        Next

'    'Se for versão LIGHT
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

Private Function Carrega_TabelaPreco() As Long
'Carrega na combo as Tabelas de Preço existentes

Dim lErro As Long
Dim objCodDescricao As AdmCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_TabelaPreco

    TabPrecoItemNF.AddItem ""
    TabPrecoItemNF.ItemData(TabPrecoItemNF.NewIndex) = 0
    
    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePrecoVenda", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao, "Ativo = 1")
    If lErro <> SUCESSO Then gError 42118

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o ítem na Lista de Tabela de Preços
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

        Case 42118

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157163)

    End Select

    Exit Function

End Function

Private Function Carrega_Transportadora() As Long
'Carrega na combo as Transportadoras existentes

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_Transportadora

    'Lê o Código e o NOme de Todas as Transportadoras do BD
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then gError 42119

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

        Case 42119

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157164)

    End Select

    Exit Function

End Function

Private Function Carrega_CanalVenda() As Long
'Carrega na combo os Canais de Venda existentes

Dim lErro As Long
Dim objCodDescricao As AdmCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_CanalVenda

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "CanalVenda", "Codigo", "NomeReduzido", STRING_CANAL_VENDA_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 42120

    For Each objCodDescricao In colCodigoDescricao
        'Adiciona o ítem na Lista de Tabela de Preços
        CanalVenda.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        CanalVenda.ItemData(CanalVenda.NewIndex) = objCodDescricao.iCodigo
    Next

    Carrega_CanalVenda = SUCESSO

    Exit Function

Erro_Carrega_CanalVenda:

    Carrega_CanalVenda = gErr

    Select Case gErr

        Case 42120

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157165)

    End Select

    Exit Function

End Function

Private Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim objSerie As ClassSerie
Dim colSerie As New colSerie
Dim sSerieAnt As String

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 42121

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

        Case 42121, 500115, 500116
        
        Case 500117

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157166)

    End Select

    Exit Function

End Function

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
    
        Case 500130, 500131
        
        Case 500132
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157167)
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157168)

    End Select

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
'    If lErro <> SUCESSO Then gError 42124
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
'        Case 42124
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157169)
'
'    End Select
'
'    Exit Function
'
'End Function

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

        Case 26492

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157170)

    End Select

    Exit Function

End Function

Private Function GridItens_InicializaColunas(ByVal objCT As Object) As Long

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
    iGrid_Contrato_Col = 11 - iColunasEscondidas
    iGrid_ItemContrato_Col = 12 - iColunasEscondidas
    iGrid_DataCobranca_Col = 13 - iColunasEscondidas
    iGrid_SerieNFOrig_Col = 14 - iColunasEscondidas
    iGrid_NumNFOrig_Col = 15 - iColunasEscondidas
    iGrid_ItemNFOrig_Col = 16 - iColunasEscondidas

    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        iGrid_ComissaoItem_Col = 17 - iColunasEscondidas
    Else
        iColunasEscondidas = iColunasEscondidas + 1
    End If

End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens
    Inicializa_Grid_Itens = CF2(Me, "NFF_Inicializa_Grid_Itens", objGridInt)
End Function

Public Function NFF_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
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
    objGridInt.colColuna.Add ("Contrato")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Data Cobrança")
    objGridInt.colColuna.Add ("Série NF Orig")
    objGridInt.colColuna.Add ("Núm NF Orig")
    objGridInt.colColuna.Add ("Item NF Orig")

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
    objGridInt.colCampo.Add (Contrato.Name)
    objGridInt.colCampo.Add (Item.Name)

    objGridInt.colCampo.Add (DataCobranca.Name)

    objGridInt.colCampo.Add (SerieNFOriginal.Name)
    objGridInt.colCampo.Add (NumNFOriginal.Name)
    objGridInt.colCampo.Add (ItemNFOriginal.Name)

    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        objGridInt.colCampo.Add (ComissaoItemNF.Name)
    Else
        ComissaoItemNF.left = POSICAO_FORA_TELA
    End If
    
    Call GridItens_InicializaColunas(Me)

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

    NFF_Inicializa_Grid_Itens = SUCESSO

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
    objGridInt.colColuna.Add ("Desconto 1 Tipo")
    objGridInt.colColuna.Add ("Desc.1 Data")
    objGridInt.colColuna.Add ("Desc.1 Valor")
    objGridInt.colColuna.Add ("Desc.1 %")
    objGridInt.colColuna.Add ("Desconto 2 Tipo")
    objGridInt.colColuna.Add ("Desc.2 Data")
    objGridInt.colColuna.Add ("Desc.2 Valor")
    objGridInt.colColuna.Add ("Desc.2 %")
    objGridInt.colColuna.Add ("Desconto 3 Tipo")
    objGridInt.colColuna.Add ("Desc.3 Data")
    objGridInt.colColuna.Add ("Desc.3 Valor")
    objGridInt.colColuna.Add ("Desc.3 %")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
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

    objGridInt.iProibidoExcluir = 1

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
    sTabela = "NFiscalFatura"

    'Lê os dados da Tela NFiscal
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 42132

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
    colCampoValor.Add "NumIntDocCPR", objNFiscal.lNumIntDocCPR, 0, "NumIntDocCPR"
    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_BUFFER_MAX_TEXTO, "VolumeNumero"
    colCampoValor.Add "FreteRespons", objNFiscal.iFreteRespons, 0, "FreteRespons"
    colCampoValor.Add "DataReferencia", objNFiscal.dtDataReferencia, 0, "DataReferencia"
    colCampoValor.Add "Status", objNFiscal.iStatus, 0, "Status"
    colCampoValor.Add "SemDataSaida", objNFiscal.iSemDataSaida, 0, "SemDataSaida"
    colCampoValor.Add "NaoImpCobranca", objNFiscal.iNaoImpCobranca, 0, "NaoImpCobranca"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
'    colSelecao.Add "Status", OP_DIFERENTE, STATUS_CANCELADO
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 42132

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157171)

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
        If lErro <> SUCESSO And lErro <> 31442 Then gError 180091

        'Coloca os dados da NFiscal na tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 42131

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 42131, 180091

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157172)

    End Select

    Exit Sub

End Sub

Public Sub ClienteLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    objCliente.sNomeReduzido = Cliente.Text

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

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim colSelecao As Collection

    'Recolhe os dados da Nota Fiscal
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then Exit Sub

    'Chama a Tela NFiscalFaturaLista
    Call Chama_Tela("NFiscalFaturaLista", colSelecao, objNFiscal, objEventoNFiscal)

    Exit Sub

End Sub

Private Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objNFiscal = obj1

    'Evita ter que alterar os browses a cada campo inserido na tabela de NF
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 180090

    'Coloca na Tela a Nota Fiscal escolhida
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 42133

    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case gErr

        Case 42133, 180090

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157173)

    End Select

    Exit Sub

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
        If GridItens.Row = 0 Then gError 43697

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 83085
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub
    
Erro_BotaoProdutos_Click:

    Select Case gErr
    
        Case 43697
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 83085
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157174)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoEnxuto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    'verifica se tem alguma linha do Grid selecionada
    If GridItens.Row = 0 Then gError 42134

    'Verifica se o Produto está preenchido
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then

        Set objProduto = obj1

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 42135

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
        
        If Not (Me.ActiveControl Is Produto) Then
        
            'Coloca o produto, a Descrição e a Unidade de Medida da tela
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
    
            lErro = Produto_Saida_Celula(objProduto)
            If lErro <> SUCESSO Then gError 42136

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    Select Case gErr

        Case 42134, 42136
        
        Case 42135
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157175)

    End Select

    Exit Sub

End Sub

Public Sub BotaoEstoqueProd_Click()

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim iAlmoxarifado As Integer
Dim objEstoqueProd As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto, sSelecao As String, sProdutoEnxuto As String

On Error GoTo Erro_BotaoEstoqueProd_Click

    'Verifica se a linha do Grid é válida
    If GridItens.Row = 0 Then gError 43698

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
            If lErro <> SUCESSO And lErro <> 23796 Then gError 58050
            
            If lErro = SUCESSO Then objEstoqueProd.iAlmoxarifado = iAlmoxarifado
            
            'Chama a tela EstoqueProdutoFilialLista
            Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProd, objEventoEstoqueProd)
        
        Else
        
            Call Mascara_RetornaProdutoEnxuto(sProdutoFormatado, sProdutoEnxuto)
            sSelecao = "Produto LIKE '" & sProdutoEnxuto & "%' AND FilialEmpresa = " & giFilialEmpresa
            Call Chama_Tela("EstProdLista_Consulta", colSelecao, objEstoqueProd, objEventoEstoqueProd, sSelecao)
        
        End If
        
    Else
        gError 43729
    End If

    Exit Sub
    
Erro_BotaoEstoqueProd_Click:

    Select Case gErr
    
        Case 101782
        
        Case 43698
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 43729
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 101783
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157176)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoEstoqueProd_evSelecao(obj1 As Object)

    Me.Show

End Sub

Public Sub TransportadoraLabel_Click()

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As Collection

    'Se a transportadora estiver preenchida recolhe
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

Public Sub SerieNFOriginalLabel_Click()

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

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente
Dim sSelecao As String

On Error GoTo Erro_NFiscalOriginalLabel_Click

    'Se o cliente não está preenchido => erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 101090
    
    'Recolhe os dados da Nota Fiscal
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 42137

    'Guarda a Serie e o Número da Nota Fiscal Original da Tela
    objNFiscal.sSerie = SerieNFiscalOriginal.Text
    If Len(Trim(NFiscalOriginal.ClipText)) > 0 Then
        objNFiscal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
    Else
        objNFiscal.lNumNotaFiscal = 0
    End If

    'Guarda o nome reduzido do Cliente
    objCliente.sNomeReduzido = Cliente.Text
    
    'Lê o Cliente
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 111080
    
    'Se não achou o Cliente --> erro
    If lErro = 12348 Then gError 111081
    
    'Faz o filtro por cliente
    colSelecao.Add objCliente.lCodigo
    sSelecao = "Cliente=?"
    
    'Chama a Tela NFiscalNormalLista
    Call Chama_Tela("NFiscalFaturaNormalLista", colSelecao, objNFiscal, objEventoNFiscalOriginal, sSelecao)

    Exit Sub

Erro_NFiscalOriginalLabel_Click:

    Select Case gErr
        
        Case 42137
        
        Case 111080
        
        Case 101090
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 111081
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157177)

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

Public Sub CondPagtoLabel_DblClick()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection

    'Se Condição de Pagto estiver preenchida, extrai o código
    objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)
    
    'Chama a Tela CondicoesPagamentoCRLista
    Call Chama_Tela("CondicaoPagtoCRLista", colSelecao, objCondicaoPagto, objEventoCondPagto)

End Sub

Private Sub objEventoCondPagto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCondicaoPagto As ClassCondicaoPagto

On Error GoTo Erro_objEventoCondPagto_evSelecao

    Set objCondicaoPagto = obj1

    'Preenche campo CondicaoPagamento
    CondicaoPagamento.Text = CondPagto_Traz(objCondicaoPagto)
    Call CondicaoPagamento_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoCondPagto_evSelecao:

    Select Case gErr

        Case 42138

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157178)

     End Select

     Exit Sub

End Sub

Public Sub BotaoVendedores_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoVendedores_Click
    
    'Chama a tela de browse de Vendedores
    lErro = objTabComissoes.BotaoVendedores_Click()
    If lErro <> SUCESSO Then gError 43696
    
    Exit Sub

Erro_BotaoVendedores_Click:

    Select Case gErr
    
        Case 43696
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157179)
    
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
            If lErro <> SUCESSO Then gError 51617
        
        'senao
        Else
            'verificar se o vendedor esta no grid e tratar o campo 'DiretoIndireto'
            lErro = objTabComissoes.VendedorLinha_Preenche1(objVendedor)
            If lErro <> SUCESSO Then gError 101516
        
        End If
    
    End If

    Me.Show
    
    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr

        Case 51617, 101516 'tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157180)

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
        If lErro <> SUCESSO Then gError 64005

        'Tenta ler a nota Fiscal passada como parâmetro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 42403
        
        'Se não achou a NFiscal --> erro
        If lErro <> SUCESSO Then gError 42406

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

        'Lê o Tipo de Documento com o Código de objNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 42405
        
        'Se não encontrar --> Erro
        If lErro = 31415 Then gError 42407

        If objTipoDocInfo.sNomeTelaNFiscal <> Me.Name Then gError 42408

        'Traz a nota para a tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 42404

    End If

    iAlterado = 0
    iComissoesAlterada = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 42403, 42404, 42405, 64005

        Case 42406
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case 42407
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 42408
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_DIFERENTE_NF_FATURA", gErr, objNFiscal.iTipoNFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157181)

    End Select
    
    iAlterado = 0
    iComissoesAlterada = 0

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
    
    If objTipoDocInfo.iCodigo = DOCINFO_NFISFVFE And gobjFAT.iVFEUsaTManual = MARCADO Then
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

    gobjRastreamento.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 42138

    'se não estiver cadastrado ==> erro
    If lErro = 31415 Then gError 42139
    
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Or (objTipoDocInfo.iTipoMovtoEstoque = 0 And objTipoDocInfo.iTipoMovtoEstoque = 0) Then
        Frame1(TAB_Almoxarifado).Enabled = False
    Else
        Frame1(TAB_Almoxarifado).Enabled = True
    End If
    
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then
        giComplemento = MARCADO
    Else
        giComplemento = DESMARCADO
    End If
    
    'If Not objGridAlocacoes Is Nothing Then Call Grid_Limpa(objGridAlocacoes)
    
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
    If Len(Trim(Cliente.Text)) = 0 Then Call gobjTribTab.NatOpInterna_AlterarTrocaTipoNF(NaturezaOp_Conv34(objTipoDocInfo.sNaturezaOperacaoPadrao, MaskedParaDate(DataEmissao)))

    ''*** incluidos p/tratamento de tributacao *******************************
    If gbLimpandoTela = False Then
        Call gobjTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
        If IsDate(DataEmissao.Text) Then Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
        Call gobjTribTab.Serie_Alterada(Serie.Text)
    End If
    '*** fim tributacao
    

    'subtipo contabil
    'ReInicializa a parte de contabilidade devido que o tipo implica na mudança dos modelos de contabilização
    lErro = objContabil.Contabil_Processa_Alteracao_Subtipo(objTipoDocInfo.iSubTipoContabil)
    If lErro <> SUCESSO Then gError 71690
    
    If objTipoDocInfo.iCodigo = DOCINFO_NFISFCP Then
        lErro = Trata_Complemento_Preco
        If lErro <> SUCESSO Then gError 71690
    End If

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

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objTipoCliente As New ClassTipoCliente
Dim objTipoFrete As ClassTipoFrete
Dim colTipoFrete As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Cliente_Validate

    'Verifica se o cliente foi alterado
    If iClienteAlterado = 0 Then Exit Sub
    'Se op cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        Transportadora.ListIndex = -1

        lErro = TP_Cliente_Le3(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 42142

        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 42143

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
        If lErro <> SUCESSO Then gError 131885
        
        giFilialEntregaManual = 0
        
        'Lê o tipo do cliente
        If objCliente.iTipo > 0 Then
            objTipoCliente.iCodigo = objCliente.iTipo
            lErro = CF("TipoCliente_Le", objTipoCliente)
            If lErro <> SUCESSO And lErro <> 19062 Then gError 42144
        End If
        
        lErro = CF2(Me, "NFCF_Cliente_Validate2", objCliente, objTipoCliente)
        If lErro <> SUCESSO Then gError 42144
        
        If Not gbCarregandoTela Then
        
            Call DescontoGlobal_Recalcula
    
            Call ValorTotal_Calcula
     
            'verifica se o cliente tem uma tabela default
            If objCliente.iTabelaPreco > 0 Then
                TabelaPreco.Text = objCliente.iTabelaPreco
            ElseIf objTipoCliente.iTabelaPreco > 0 Then
                TabelaPreco.Text = objTipoCliente.iTabelaPreco
            End If
            
            If Len(Trim(TabelaPreco.Text)) > 0 Then Call TabelaPreco_Validate(bSGECancelDummy)

            If objCliente.iCondicaoPagto > 0 Then
                CondicaoPagamento.Text = objCliente.iCondicaoPagto
            ElseIf objTipoCliente.iCondicaoPagto > 0 Then
                CondicaoPagamento.Text = objTipoCliente.iCondicaoPagto
            End If
            
            Call CondicaoPagamento_Validate(bSGECancelDummy)
        
            lErro = CF2(Me, "NFCF_Cliente_Validate3", objCliente)
            If lErro <> SUCESSO Then gError 42912
            
            'Tipo de frete
            lErro = CF("TipoFrete_Le", colTipoFrete)
            If lErro <> SUCESSO Then gError 207426
            
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
        
        iClienteAlterado = 0
    
        If Not gbCarregandoTela Then

    ''*** incluidos p/tratamento de tributacao *******************************
            If iCodFilial <> 0 Then Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, iCodFilial) '####
    '*** fim tributacao
        End If
    
    ElseIf Len(Trim(Cliente.Text)) = 0 Then 'Verifica preenchimento de Cliente

        Filial.Clear 'Se não estiver preenchido limpa as combos de filialis
        FilialEntrega.Clear

        Call CF2(Me, "NFCF_Limpa_Tela")

    End If

    iClienteAlterado = 0

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 42142, 131885, 207426
            
        Case 42143, 42144, 42912

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157184)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    'Verifica se algo foi selecionada
    If Filial.ListIndex = -1 Then Exit Sub

    'Faz o tratamento da Filial selecionada
    lErro = Trata_FilialCliente()
    If lErro <> SUCESSO Then gError 42145

    Exit Sub

Erro_Filial_Click:

    Select Case gErr

        Case 42145

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157185)

    End Select

    Exit Sub

End Sub

Private Function Trata_FilialCliente() As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim objVendedor As New ClassVendedor
Dim objTipoCliente As New ClassTipoCliente
Dim iVendedorLeu As Integer
Dim dValorTotal As Double
Dim dValorBase As Double
Dim objTransportadora As New ClassTransportadora
Dim dValorComissao As Double
Dim dValorEmissao As Double
Dim dComissaoVendas As Double
Dim objMensagem As New ClassMensagem
Dim objTipoFrete As ClassTipoFrete
Dim colTipoFrete As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Trata_FilialCliente

    objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
    'Lê a FilialCliente
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilialCliente)
    If lErro <> SUCESSO Then gError 42146

    objCliente.sNomeReduzido = Cliente.Text
    'Lê o Cliente
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 42147
        
    'Se não achou o Cliente --> erro
    If lErro = 12348 Then gError 42148

    Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, objFilialCliente.iCodFilial)
    
    'Comissões
    'se empresa nao utiliza as regras para
    'calculo de comissoes
    If Not gobjCRFAT.iUsaComissoesRegras = USA_REGRAS And ComissaoAutomatica.Value = 1 Then
        
        objTabComissoes.Limpa_Grid_Comissoes
        
        'calcular as comissoes no modelo antigo...
        lErro = objTabComissoes.Comissao_Automatica_FilialCli_Exibe(objFilialCliente)
        If lErro <> SUCESSO Then gError 59048
    
    End If

    'Tipo de frete
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207425
    
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


    'Verififca se a transporadora está preenchida
    If Len(Trim(Transportadora.Text)) = 0 Then
        
        lErro = CF("Transportadora_Padrao_FilCli", objFilialCliente, objTransportadora)
        If lErro <> SUCESSO Then gError 59049

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
    
    If giFilialEntregaManual = 0 Then
    
        FilialEntrega.Text = CStr(objFilialCliente.iCodFilial)
        Call FilialEntrega_Validate(bSGECancelDummy)
        
        giFilialEntregaManual = 0
        
    End If
    
    lErro = CF2(Me, "NFCF_Cliente_Filial_Validate", objFilialCliente)
    If lErro <> SUCESSO Then gError 131886
'
'    If objFilialCliente.iCodMensagem > 0 Then
'        objMensagem.iCodigo = objFilialCliente.iCodMensagem
'        'Lê a Mensagem "default" p/ o Cliente
'        lErro = CF("Mensagem_Le", objMensagem)
'        If lErro <> SUCESSO And lErro <> 19234 Then gError 140377
'
'        If lErro = SUCESSO Then Mensagem.Text = objMensagem.sDescricao
'
'    End If
    
    Trata_FilialCliente = SUCESSO

    Exit Function

Erro_Trata_FilialCliente:

    Trata_FilialCliente = gErr

    Select Case gErr

        Case 42146, 42147, 42148, 59048, 59049, 124056, 131886, 140377, 207425 'Inserido por Wagner
    
        Case 124057
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objTransportadora.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157186)
    
    End Select

    Exit Function

End Function

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
    If Filial.ListIndex > -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42152

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o Cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 42153

        sCliente = Cliente.Text

        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 42154

        If lErro = 17660 Then gError 42155

        'Coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        lErro = Trata_FilialCliente()
        If lErro <> SUCESSO Then gError 42156

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 42157

    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case gErr

        Case 42152, 42154

        Case 42153
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 42155
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If
            
        Case 42156

        Case 42157
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157187)

    End Select

    Exit Sub

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
            If lErro <> SUCESSO And lErro <> 12253 Then gError 42158
            'Se não está na combo
            If lErro <> SUCESSO Then
            
                objSerie.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
                'Busca a série no BD
                lErro = CF("Serie_Le", objSerie)
                If lErro <> SUCESSO And lErro <> 22202 Then gError 51621
                If lErro <> SUCESSO Then gError 42159 'Se não encontrou
            
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

        Case 42159
            'Pergunta se deseja criar nova série
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_SERIE", Serie.Text)
            'Se a resposta for afirmativa
            If vbMsgRes = vbYes Then
                'Chama a tela de cadastro de séries
                Call Chama_Tela("SerieNFiscal", objSerie)
            End If
            'segura o foco na série

        Case 42158, 51621

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157188)

    End Select

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    If iDataEmissaoAlterada = 1 Then

        'Verifica se a data de emissao foi digitada
        If Len(Trim(DataEmissao.ClipText)) <> 0 Then
    
            'Critica a data digitada
            lErro = Data_Critica(DataEmissao.Text)
            If lErro <> SUCESSO Then gError 42160
        
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

        Case 42160

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157189)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 42161

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 42161

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157190)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 42162

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 42162

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157191)

    End Select

    Exit Sub

End Sub

Public Sub DataSaida_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataSaida_Validate

    'Verifica se a data de Saida foi digitada
    If Len(Trim(DataSaida.ClipText)) <> 0 Then

        'Critica a data digitada
        lErro = Data_Critica(DataSaida.Text)
        If lErro <> SUCESSO Then gError 42163

    End If
    
    lErro = Cobranca_Automatica()
    If lErro <> SUCESSO Then gError 56909

    Exit Sub

Erro_DataSaida_Validate:

    Cancel = True

    Select Case gErr
        Case 42163, 56909

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157192)

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
    If lErro <> SUCESSO Then gError 89798

    Exit Sub

Erro_HoraSaida_Validate:

    Cancel = True

    Select Case gErr

        Case 89798

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157193)

    End Select

    Exit Sub

End Sub

Public Sub UpDownSaida_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownSaida_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataSaida, DIMINUI_DATA)
    If lErro Then gError 42164

    Exit Sub

Erro_UpDownSaida_DownClick:

    Select Case gErr

        Case 42164

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157194)

    End Select

    Exit Sub

End Sub

Public Sub UpDownSaida_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownSaida_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataSaida, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 42165

    Exit Sub

Erro_UpDownSaida_UpClick:

    Select Case gErr

        Case 42165

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157195)

    End Select

    Exit Sub

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
    If lErro <> SUCESSO Then gError 42166

    Exit Sub

Erro_TabelaPreco_Click:

    Select Case gErr

        Case 42166

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157196)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42167

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 42168
        
        If lErro <> SUCESSO Then gError 42170 'Não encontrou Tabela Preço no BD

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 42169

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 42171

    Exit Sub

Erro_TabelaPreco_Validate:

    Cancel = True


    Select Case gErr

        Case 42167, 42168, 42169
    
        Case 42170  'Não encontrou Tabela de Preço no BD
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")
    
            If vbMsgRes = vbYes Then
                'Chama a tela de Tabelas de Preço
                Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
            End If
            'Segura o foco
    
        Case 42171
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157197)

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

Public Sub PrecoTotal_Calcula(iLinha As Integer, Optional iCalculaTotal = True)

Dim dPrecoTotal As Double
Dim dPrecoTotalReal As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim lTamanho As Long
Dim dValorTotal As Double
Dim iIndice As Integer
Dim dValorProdutos As Double
Dim dValorDescontoGlobal As Double, dValorDescontoItens As Double
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

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
    If lErro <> SUCESSO Then gError 184285
    
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
'
'        'Limpa o valor de desconto
'        gdDesconto = 0
'
'        Call CF2(Me, "NFCF_Limpa_Desconto")
'
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
'        If lErro <> SUCESSO Then gError 56883
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
    If lErro <> SUCESSO Then gError 132029

    If iCalculaTotal <> False Then Call ValorTotal_Calcula
    
    Exit Sub

Erro_PrecoTotal_Calcula:

    Select Case gErr
    
        Case 56883, 184285
        
        Case 132029
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157199)
            
    End Select
            
    Exit Sub

End Sub

Function ValorTotal_Calcula() As Long

Dim dValorTotal As Double
Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    lErro = gobjTribTab.ValorTotal_Calcula(dValorTotal)
    If lErro <> SUCESSO Then gError 42176

    'Coloca o valor total na tela
    ValorTotal.Caption = Format(dValorTotal, "Standard")
    
    Call ValorDescontoTit_Validate(bSGECancelDummy)

    Call ValorDescontoItens_Calcula
    
    If Not gbCarregandoTela Then
        'Gera a cobranca em cima do novo valor total
        lErro = Cobranca_Automatica()
        If lErro <> SUCESSO Then gError 56908

        'Se a empresa nãu utiliza as regras para o cálculo de comissões
        If Not gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then
        
            'Significa que é necessário recalcular as comissões, utilizando o cálculo padrão
            lErro = objTabComissoes.Comissoes_Calcula_Padrao()
            If lErro <> SUCESSO Then gError 42177
        
        Else
            iComissoesAlterada = 1 'preventivamente, pois pode ter alterado algo que afete uma regra
        End If
    
    End If
    
    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 42176, 42177, 56883, 56884, 56908
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157200)

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

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
        'Produto
        Case Produto.Name
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 42184

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
            If lErro <> SUCESSO Then gError 42180

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                UnidadeMed.Enabled = False
            Else
                UnidadeMed.Enabled = True

                objProduto.sCodigo = sProdutoFormatado
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 42178
                If lErro = 28030 Then gError 42183 'Não achou

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lâ as Unidades de Medidas da Classe do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 42179
                'Carrega a combo de UM
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next
                'Seleciona na UM que está preenchida
                UnidadeMed.Text = sUM
                If Len(Trim(sUM)) > 0 Then
                    lErro = Combo_Item_Igual(UnidadeMed)
                    If lErro <> SUCESSO And lErro <> 12253 Then gError 42182
                End If
            End If
                
        'Nas demais
        Case TabPrecoItemNF.Name, ComissaoItemNF.Name, PrecoUnitario.Name, PercentDesc.Name, Desconto.Name, DescricaoItem.Name
            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 42181

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case Desconto1Ate.Name, Desconto1Valor.Name, Desconto1Percentual.Name
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

        Case Desconto2Codigo.Name, Desconto3Codigo.Name

            If Len(Trim(GridParcelas.TextMatrix(iLinha, GridParcelas.Col - 4))) = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
        Case ValorParcela.Name
            'Se o vencimento estiver preenchido, habilita o controle
            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))) = 0 Then
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

        Case Contrato.Name, Item.Name, DataCobranca.Name

            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 136294

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

     End Select

    'Rastreamento
    lErro = gobjRastreamento.Rotina_Grid_Enable(iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 83386

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 42178, 42179, 42180, 42181, 42182, 42184, 46713, 83386, 136294, 136295

        Case 42183
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157201)

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
        
        gcolcolMedicoes.Remove GridItens.Row

        For iLinhasExistentesAnterior = 1 To objGridItens.iLinhasExistentes
            If gobjNFiscal.ColItensNF(iLinhasExistentesAnterior).iPossuiGrade = MARCADO Then
                GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = "# " & iLinhasExistentesAnterior
            Else
                GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = iLinhasExistentesAnterior
            End If
        Next
        
        GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = iLinhasExistentesAnterior
        '*******************************
        
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
        If lErro <> SUCESSO Then gError 132030

''*** incluidos p/tratamento de tributacao *******************************

        lErro = gobjTribTab.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 42185

'*** fim tributacao

        Call ValorTotal_Calcula
    
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81517
    
        'Rastreamento
        lErro = gobjRastreamento.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 83387
    
        For iIndice = 1 To gobjNFiscal.ColItensNF.Count
            If gobjNFiscal.ColItensNF(iIndice).iPossuiGrade = MARCADO Then
                GridItens.TextMatrix(iIndice, 0) = "# " & CStr(iIndice)
            Else
                GridItens.TextMatrix(iIndice, 0) = CStr(iIndice)
            End If
            gobjNFiscal.ColItensNF(iIndice).iItem = iIndice 'Incluído para ao alterar a grade ou a reserva não fazer do item errado
        Next
        GridItens.TextMatrix(objGridItens.iLinhasExistentes + 1, 0) = CStr(objGridItens.iLinhasExistentes + 1)
    
    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr
    
        Case 42185, 81517, 83387
        
        Case 132030 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157202)
            
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

    'Rastreamento
    If bExcluiuLinha Then
        Call gobjRastreamento.Exclusao_Item_Grid(iItemAtual)
    End If

End Sub

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
Dim colMedicao As New Collection
Dim dPercDescontoTab As Double, dPrecoComDescontoTab As Double

On Error GoTo Erro_Produto_Saida_Celula

    'Critica o Produto
    If objProduto Is Nothing Then
        
        Set objProduto = New ClassProduto
        
        'Critica o Produto
        lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 26657
        'Se o produto é gerencial sem ser de grade ==> Erro
        If lErro = 86295 And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then gError 86296 'Alterado por Wagner
               
        'Se o produto não foi encontrado ==> Pergunta se deseja criar
        If lErro = 51381 Then gError 26658
        
        
    'incluido por tulio para poder ler a UM... 7/8/02
    Else

        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 101782
        If lErro <> SUCESSO Then gError 101783
        
        If objProduto.iGerencial = GERENCIAL And objProduto.iKitVendaComp <> MARCADO Then gError 86296 'Alterado por Wagner
        
    End If
    
    'Sinaliza que o produto possui grade
    If Len(Trim(objProduto.sGrade)) > 0 Then iPossuiGrade = MARCADO
    
    'Se não for um produto de grade
    If iPossuiGrade = DESMARCADO And objProduto.iKitVendaComp = DESMARCADO Then 'Alterado por Wagner
        
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
                    
        lErro = GrupoGrade_Valida_Itens(objRomaneioGrade)
        If lErro <> SUCESSO Then gError 86310
                    
        Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
        If giRetornoTela <> vbOK Then gError 86310

        lErro = GrupoGrade_Trata_Itens(objRomaneioGrade, objProduto)
        If lErro <> SUCESSO Then gError 86310
        
    End If
    
    'Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda
    'Descricao
    GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objProduto.sDescricao

    GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(gdDesconto, "Percent")

    lErro = CF2(Me, "NF_Produto_Saida_Celula", objProduto)
    If lErro <> SUCESSO Then gError 184286
    
    'Preco Unitário
    
    dPrecoUnitario = 0
    
    lErro = CF2(Me, "NFCF_PrecoUnitario_Calcula", giFilialEmpresa, objProduto.sCodigo, objProduto.sSiglaUMVenda, Cliente.Text, Codigo_Extrai(Filial.Text), Codigo_Extrai(TabelaPreco.Text), StrParaDbl(PercAcrescFin.Text) / 100, MOEDA_REAL, objRomaneioGrade, dPrecoUnitario, False, dPercDescontoTab, dPrecoComDescontoTab)
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

''*** incluidos p/tratamento de tributacao *******************************

        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 27832
        
        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItens.Row)
        
        gcolcolMedicoes.Add colMedicao
        
        objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
        
        gobjNFiscal.ColItensNF.Add1 objItemNF
        gobjNFiscal.ColItensNF(GridItens.Row).sUMEstoque = objProduto.sSiglaUMEstoque
        gobjNFiscal.ColItensNF(GridItens.Row).iItem = GridItens.Row
        gobjNFiscal.ColItensNF(GridItens.Row).sProduto = objProduto.sCodigo
        
        If iPossuiGrade = MARCADO Or objProduto.iKitVendaComp = MARCADO Then
        
            If objProduto.iKitVendaComp = MARCADO Then
                Set objItemNF.colItensRomaneioGrade = colItensRomaneioGrade
            End If
                       
            Set gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade = objItemNF.colItensRomaneioGrade
        
            If objProduto.iKitVendaComp <> MARCADO Then
            
            '************** GRADE ************
                gobjNFiscal.ColItensNF(GridItens.Row).iPossuiGrade = iPossuiGrade
                           
                GridItens.TextMatrix(GridItens.Row, 0) = "# " & GridItens.TextMatrix(GridItens.Row, 0)
                
                Call Atualiza_Grid_Itens(objItemNF)
                
            End If
            
            Call Atualiza_Grid_Alocacao(objItemNF)
            
        End If
        
'*** fim tributacao

    End If
    
    Call PrecoTotal_Calcula(GridItens.Row, False)
    Call ValorTotal_Calcula
    
    lErro = GrupoGrade_Separa_Itens(objRomaneioGrade)
    If lErro <> SUCESSO Then gError 177163
    
    Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_Produto_Saida_Celula:

    Produto_Saida_Celula = gErr

    Select Case gErr

        Case 26637, 26657, 27832, 51420, 86310, 86327, 86329, 101782, 35786, 177582, ERRO_SEM_MENSAGEM
        
        Case 26658
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)
                
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
            End If
            
        Case 35784
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, Produto.Text, iIndice)
            
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157203)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157204)

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
            If lErro <> SUCESSO Then gError 500093
            
            'Lê o produto
            objProduto.sCodigo = sProdutoFormatado
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 500094

            'Se o produto não está cadastrado, erro
            If lErro = 28030 Then gError 500096
                        
            'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then gError 500095
                                
            dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
            
            'Acumula o peso Bruto e o peso líquido
            dPesoBruto = dPesoBruto + objProduto.dPesoBruto * dFator * dQuantidade
            dPesoLiquido = dPesoLiquido + objProduto.dPesoLiq * dFator * dQuantidade
                        
        End If
        
    Next
    
    'Coloca os Pesos na tela
    PesoBruto.Text = Format(dPesoBruto, PesoBruto.Format) 'Alterado por Wagner
    PesoLiquido.Text = Format(dPesoLiquido, PesoLiquido.Format) 'Alterado por Wagner
    
    Pesos_Recalcula = SUCESSO
    
    Exit Function
    
Erro_Pesos_Recalcula:

    Pesos_Recalcula = gErr
    
    Select Case gErr
    
        Case 500093, 500094, 500095
        
        Case 500096
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157205)
    
    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidade que está deixando de ser a corrente

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
        If lErro <> SUCESSO Then gError 42196
        'Recolhe a quantidade informada
        dQuantidade = CDbl(Quantidade.Text)
        Quantidade = Formata_Estoque(dQuantidade)
        'Recolhe a quantidade anteriormente digitada
        dQuantidadeAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
        
        If dQuantidadeAnterior = dQuantidade Then iSairCelula = True

    End If

    GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)
    

    If (Not iSairCelula) And Frame1(TAB_Almoxarifado).Enabled = True Then
    
        objProduto.sCodigo = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177585
        
        If objProduto.iKitVendaComp <> MARCADO Then
            'Tenta fazer uma alocação p\ oitem de acordo com a nova UM
            lErro = Alocacao_Processa()
            If lErro <> SUCESSO Then gError 42198
    
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
       
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42197
    
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))

    If dPrecoUnitario > 0 Then Call PrecoTotal_Calcula(GridItens.Row, False)
    
    '*** incluidos p/tratamento de tributacao ****************
    If GridItens.Row <= objGridItens.iLinhasExistentes Then
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 42199
    End If
    '*** fim tributacao

    If Not iSairCelula Then
    
        Call ValorTotal_Calcula

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81521
    
    End If
    
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 42199, 177585, 177586, 177587
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 42196, 42197, 42198, 81521
            GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidadeAnterior)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157206)

    End Select

    Exit Function

End Function

Public Function Alocacao_Processa() As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dQuantidade As Double
Dim sUM As String, bFat As Boolean

On Error GoTo Erro_Alocacao_Processa

    'Remove as alocações do item selecionado do grid de locações
    Call Remove_Alocacoes(GridItens.Row)

    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 42200

    objProduto.sCodigo = sProduto

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 42201

    'Se não encontrou o Produto --> erro
    If lErro = 28030 Then gError 42202

    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    
    lErro = gobjTribTab.Testa_Item_Fat(GridItens.Row, bFat)
    If lErro <> SUCESSO Then gError 42203

    If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 And gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA And bFat Then

        sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

        lErro = AlocaAlmoxarifadoPradrao(dQuantidade, objProduto, sUM)
        If lErro <> SUCESSO Then gError 42203
    
    End If

    Alocacao_Processa = SUCESSO

    Exit Function

Erro_Alocacao_Processa:

    Alocacao_Processa = gErr

    Select Case gErr

        Case 42200, 42201, 42203

        Case 42202
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157207)

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
Dim iLinhas As Integer
Dim dAcrescimo As Double
Dim iNumCasasDec As Integer
Dim dTotal As Double
Dim iCodigo As Integer

On Error GoTo Erro_AlocaAlmoxarifadoPradrao

    'Faz a conversão da UM da tela para a UM de estoque
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 42204

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
        Case DOCINFO_NFISFBF, DOCINFO_NFFBFRS
        
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
                    If lErro <> SUCESSO Then gError 42210
                    'Adiciona na coleção de produtos
                    colOutrosProdutos.Add sProduto
                End If
            Next
    
            'Chama a tela de Localização de Produto
            Call Chama_Tela_Modal("LocalizacaoProduto1", objItemNF, colOutrosProdutos, dQuantAlocar, iCodigo)
            If giRetornoTela = vbCancel Then gError 42211
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
                    GridAlocacao.TextMatrix(iLinhas, iGrid_QuantVend_Col) = Formata_Estoque(dTotal)
                    GridAlocacao.TextMatrix(iLinhas, iGrid_UMAloc_Col) = objItemNF.sUMEstoque
                    
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
                    gError 42212
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157208)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoUnitario(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double
Dim dQuantCancelada As Double
Dim iValorUnitarioAlterado As Integer

On Error GoTo Erro_Saida_Celula_PrecoUnitario

    Set objGridInt.objControle = PrecoUnitario

    'Verifica o preenchimento de Preço Unitário
    If Len(Trim(PrecoUnitario.Text)) > 0 Then

        If giComplemento = MARCADO Then
            lErro = Valor_NaoNegativo_Critica(PrecoUnitario.Text)
        Else
            lErro = Valor_Positivo_Critica(StrParaDbl(PrecoUnitario.Text))
        End If
        If lErro <> SUCESSO Then gError 87190
        
        PrecoUnitario.Text = Format(PrecoUnitario.Text, gobjFAT.sFormatoPrecoUnitario)
        
        lErro = CF2(Me, "NFF_Saida_Celula_PrecoUnitario", StrParaDbl(PrecoUnitario.Text), GridItens.Row)
        If lErro <> SUCESSO Then gError 199808
        
    End If

    If PrecoUnitario.Text <> GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) Then iValorUnitarioAlterado = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42215

    If iValorUnitarioAlterado Then
    
        Call PrecoTotal_Calcula(GridItens.Row, False)
    
    ''*** incluidos p/tratamento de tributacao *******************************
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 42213
    '*** fim tributacao
    
        Call ValorTotal_Calcula

    End If
    
    Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = gErr


    Select Case gErr

        Case 42213, 42214, 42215, 87190, 199808
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157209)

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
        If lErro <> SUCESSO Then gError 42216

               
        If dPercentDesc <> dPercentDescAnterior Then
        
            'Verifica se o percentual é de 100%
            If dPercentDesc = 100 Then gError 42218
    
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
     If lErro <> SUCESSO Then gError 42217

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

        Case 42216, 42217
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 42218
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157210)

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
        dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
        dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))
        dPrecoTotal = dQuantidade * dPrecoUnitario
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

    Call PrecoTotal_Calcula(GridItens.Row, False)

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

Public Sub ValorFrete_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorFrete As Double

On Error GoTo Erro_ValorFrete_Validate

    If iValorFreteAlterado = 0 Then Exit Sub

    'Verifica se o Valor Frete foi preenchido
    If Len(Trim(ValorFrete.Text)) > 0 Then
    
        'Faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
        If lErro <> SUCESSO Then gError 42223

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

        Case 42223

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157212)

    End Select

    Exit Sub

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
        If lErro <> SUCESSO Then gError 42224

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

        Case 42224

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157213)

    End Select

    Exit Sub

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
        If lErro <> SUCESSO Then gError 42225

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

        Case 42225

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157214)

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
        If lErro <> SUCESSO Then gError 42226

        dValorDesconto = CDbl(ValorDesconto.Text)
        'Coloca o Valor formatado na tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")

        'Verifica se o Valor do Desconto é maior que a soma dos demais valores
        If dValorDesconto > dValorProdutos Then gError 42227

        dValorProdutos = dValorProdutos - dValorDesconto

    End If
    
    ValorProdutos.Caption = Format(dValorProdutos, "Standard")
    
    lErro = CF2(Me, "NFCF_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132031
    
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDesconto_Validate(Cancel, dValorDesconto)
'*** fim tributacao
    
    Call ValorTotal_Calcula

    iValorDescontoAlterado = 0
    
    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True


    Select Case gErr

        Case 42226

        Case 42227
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos)

        Case 132031 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157215)

    End Select

    Exit Sub

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
            If lErro <> SUCESSO Then gError 42228
        End If
'*** fim tributacao
        
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
                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_DADOS_PRINCIPAIS
            
            Case TAB_Itens
                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_ITENS
            
            Case TAB_Complemento
                Call Calcula_Msg
                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_COMPLEMENTO
            
            Case TAB_Cobranca
                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_COBRANCA
            
            'Alterado por Tulio em 27/03
            Case TAB_Comissoes
                'se for tab de comissoes e a empresa
                'utilizar regras para o calculo de comissoes
                lErro = objTabComissoes.Comissoes_Calcula
                If lErro <> SUCESSO Then gError 98971
                
                Parent.HelpContextID = IDH_PEDIDO_VENDA_COMISSOES
            
            Case TAB_Almoxarifado
                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_ALMOXARIFADO
            
            Case TAB_Tributacao
                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_TRIBUTACAO
            
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_CONTABILIZACAO
        
        End Select
        
    End If

    Exit Sub
    
Erro_TabStrip1_Click:

    Select Case gErr
    
        Case 42228, 98971
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157216)
            
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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42229

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o Cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 42232

        sCliente = Cliente.Text

        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe FilialEntrega com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 42230

        If lErro = 17660 Then gError 42231

        'coloca na tela
        FilialEntrega.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 42233

    giFilialEntregaAlterada = 0

    Exit Sub

Erro_FilialEntrega_Validate:

    Cancel = True


    Select Case gErr

        Case 42229, 42230

        Case 42231
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If
            
        Case 42232
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 42233
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialEntrega.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157217)

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
    'Verifica se está preenchida com o ítem selecionado na ComboBox Transportadora
    If Transportadora.Text = Transportadora.List(Transportadora.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Transportadora, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42234
    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo
        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 42235
        
        'Não encontrou Transportadora no BD
        If lErro = 19250 Then gError 42236

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    End If

    'Não existe o ítem com a String na List da Combo
    If lErro = 6731 Then gError 42237

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case gErr

        Case 42234, 42235

        Case 42236
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then
                Call Chama_Tela("Transportadora", objTransportadora)
            End If
            
        Case 42237
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157218)

    End Select

    Exit Sub

End Sub

Public Sub PlacaUF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PlacaUF_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(PlacaUF.Text)) = 0 Then Exit Sub

    'Verifica se existe o ítem na combo
    lErro = Combo_Item_Igual(PlacaUF)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 42238

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 42239

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True

    Select Case gErr

        Case 42238

        Case 42239
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157219)

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
    If lErro <> SUCESSO Then gError 42240

    'Coloca o valor formatado na Tela
    PesoBruto.Text = Format(PesoBruto.Text, PesoBruto.Format) 'Alterado por Wagner

    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True


    Select Case gErr

        Case 42240

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157220)

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
    If lErro <> SUCESSO Then gError 42241

    'Coloca o valor formatado na Tela
    PesoLiquido.Text = Format(PesoLiquido.Text, PesoLiquido.Format) 'Alterado por Wagner

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True


    Select Case gErr

        Case 42241

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157221)

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

    'Verifica se está preenchida com o ítem selecionado na ComboBox CanalVenda
    If CanalVenda.Text = CanalVenda.List(CanalVenda.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(CanalVenda, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42242

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCanalVenda.iCodigo = iCodigo

        'Tenta ler CanalVenda com esse código no BD
        lErro = CF("CanalVenda_Le", objCanalVenda)
        If lErro <> SUCESSO And lErro <> 23597 Then gError 42243
        
        'Não encontrou CanalVenda no BD
        If lErro <> SUCESSO Then gError 42244

        'Encontrou CanalVenda no BD, coloca no Text da Combo
        CanalVenda.Text = CStr(objCanalVenda.iCodigo) & SEPARADOR & objCanalVenda.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 42245

    Exit Sub

Erro_CanalVenda_Validate:

    Cancel = True


    Select Case gErr

        Case 42242, 42243

        Case 42244  'Não encontrou CanalVenda no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CANALVENDA", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CanalDeVenda", objCanalVenda)
            End If
            'Segura o foco

        Case 42245
            Call Rotina_Erro(vbOKOnly, "ERRO_CANALVENDA_NAO_ENCONTRADO", gErr, CanalVenda.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157222)

    End Select

    Exit Sub

End Sub

Public Sub SerieNFiscalOriginal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lTamanho As Long

On Error GoTo Erro_SerieNFiscalOriginal_Validate

    'Verifica se o tamanho da Serie é maior que 3
    If Len(Trim(SerieNFiscalOriginal.Text)) > STRING_SERIE Then gError 42246

    Exit Sub

Erro_SerieNFiscalOriginal_Validate:

    Cancel = True

    Select Case gErr

        Case 42246
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_SERIE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157223)

    End Select

    Exit Sub

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
    If lErro <> SUCESSO And lErro <> 19205 Then gError 42247
    
    'Não encontrou a Condição de Pagamento --> erro
    If lErro = 19205 Then gError 42248
    
    dPercAcresFin = StrParaDbl(PercAcrescFin.Text)
    
    If Len(Trim(PercAcrescFin.ClipText)) > 0 And dPercAcresFin <> objCondicaoPagto.dAcrescimoFinanceiro * 100 Then
        
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
            If lErro <> SUCESSO Then gError 42249
            
        End If
    End If
    
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_CondicaoPagamento_Click:

    Select Case gErr

        Case 42247, 42249

        Case 42248
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157224)
           
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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42250

    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condição Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 42251
        
        'Não encontrou a Condição de Pagamento
        If lErro = 19205 Then gError 42252

        'Testa se pode ser usada em Contas a Receber
        If objCondicaoPagto.iEmRecebimento = 0 Then gError 42253

        'Coloca na Tela
        CondicaoPagamento.Text = CondPagto_Traz(objCondicaoPagto)
        
        'Se ValorTotal e DataEmissao estiverem preenchidos, preenche GridParcelas
        If Len(Trim(ValorTotal)) > 0 Then
            If Len(Trim(DataReferencia.ClipText)) > 0 And CDbl(ValorTotal.Caption) > 0 Then

                'Preenche o GridParcelas
                lErro = Cobranca_Automatica()
                If lErro <> SUCESSO Then gError 42254

            End If
        End If

    End If

    'Não encontrou o valor que era STRING
    If lErro = 6731 Then gError 42255

    Exit Sub

Erro_Condicaopagamento_Validate:

    Cancel = True


    Select Case gErr

       Case 42250, 42251, 42254

       Case 42252
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If
            
        Case 42253
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", gErr, objCondicaoPagto.iCodigo)

        Case 42255
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157225)

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
Dim dValorPagarPrimParc As Double

On Error GoTo Erro_GridParcelas_Preenche

    'Limpa o GridParcelas
    Call Grid_Limpa(objGridParcelas)

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
    
        If dValorPagarPrimParc <> 0 Then
            objCondicaoPagto.dValorTotal = Arredonda_Moeda(dValorPagar - dValorPagarPrimParc)
        Else
            objCondicaoPagto.dValorTotal = dValorPagar
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
    
        objCondicaoPagto.dtDataRef = CDate(DataReferencia.Text)
        objCondicaoPagto.dtDataEntrega = StrParaDate(DataSaida.Text)
        If objCondicaoPagto.dtDataEntrega = DATA_NULA Then objCondicaoPagto.dtDataEntrega = StrParaDate(DataEmissao.Text)
        
        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True, True, True)
        If lErro <> SUCESSO Then gError 42257
        
        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.colParcelas.Count  'objCondicaoPagto.iNumeroParcelas
        
            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dtVencimento, "dd/mm/yyyy")
            
        Next
                        
    End If
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        lErro = Preenche_DescontoPadrao(iIndice)
        If lErro <> SUCESSO Then gError 51117
    Next
    GridParcelas_Preenche = SUCESSO
    
    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = gErr
    
    Select Case gErr
    
        Case 42256, 42257, 51117
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157226)
        
    End Select

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

On Error GoTo Erro_Saida_Celula_ValorParcela

    Set objGridInt.objControle = ValorParcela

    'Verifica se valor está preenchido
    If Len(ValorParcela.ClipText) > 0 Then
        
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(ValorParcela.Text)
        If lErro <> SUCESSO Then gError 42258
        
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
            'Coloca desconto padrao (le em CPRConfig)
            lErro = Preenche_DescontoPadrao(GridParcelas.Row)
            If lErro <> SUCESSO Then gError 51119
        
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
    If lErro <> SUCESSO Then gError 42259
                    
    Saida_Celula_ValorParcela = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_ValorParcela:

    Saida_Celula_ValorParcela = gErr
    
    Select Case gErr
    
        Case 42258, 42259, 51119
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157227)
        
    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoDesconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo de Desconto que está deixando de sser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer
Dim iTipo As Integer


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
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 42260
        
            'Não foi encontrado
            If lErro = 25085 Then gError 42262
            If lErro = 25086 Then gError 42263

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
        End If
    
    Else
        For iIndice = objGridInt.objGrid.Col To iGrid_Desc3Percentual_Col
            GridParcelas.TextMatrix(GridParcelas.Row, iIndice) = ""
        Next
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42261

    Saida_Celula_TipoDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDesconto:

    Saida_Celula_TipoDesconto = gErr

    Select Case gErr

        Case 42260, 42261
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 42262
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 42263
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", gErr, objGridInt.objControle.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157228)

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
        If lErro <> SUCESSO Then gError 42264
        'Se a data de vencimento estiver preenchida
        If Len(Trim(DataEmissao.ClipText)) = 0 Then
            'critica se DataDesconto ultrapassa DataVencimento
            If CDate(objGridInt.objControle.Text) < CDate(DataEmissao.Text) Then gError 42266
        End If

        If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))) > 0 Then
            dtDataVencimento = CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))
            If CDate(objGridInt.objControle) > dtDataVencimento Then gError 42267
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42265

    Saida_Celula_DescontoData = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoData:

    Saida_Celula_DescontoData = gErr

    Select Case gErr

        Case 42264, 42265
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 42266
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_INFERIOR_DATA_EMISSAO", gErr, CDate(objGridInt.objControle.Text))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 42267
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_SUPERIOR_DATA_VENCIMENTO", gErr, CDate(objGridInt.objControle.Text), dtDataVencimento)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157229)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoValor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Desconto Valor que está deixando de sser a corrente

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
        If lErro <> SUCESSO Then gError 42268
        
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42269
                    
    Saida_Celula_DescontoValor = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_DescontoValor:

    Saida_Celula_DescontoValor = gErr
    
    Select Case gErr
    
        Case 42268, 42269
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157230)
        
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
        'Critica porcentagem
        lErro = Porcentagem_Critica(objGridInt.objControle.Text)
        If lErro <> SUCESSO Then gError 42270

    
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
    If lErro <> SUCESSO Then gError 42271

    Saida_Celula_DescontoPerc = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoPerc:

    Saida_Celula_DescontoPerc = gErr

    Select Case gErr

        Case 42270, 42271
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157231)

    End Select

    Exit Function

End Function

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
        
                'Verifica se há filhos válidos com a grade preenchida
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157232)

    End Select

    Exit Sub

End Sub

Private Function PrecoUnitario_Calcula(sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long
    PrecoUnitario_Calcula = CF2(Me, "NFF_PrecoUnitario_Calcula", sUM, objTabelaPrecoItem, dPrecoUnitario)
End Function

Public Function NFF_PrecoUnitario_Calcula(ByVal objCT As Object, sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objUM As New ClassUnidadeDeMedida
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim dPercAcresFin As Double
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objCotacao As New ClassCotacaoMoeda, objTabelaPreco As New ClassTabelaPreco
Dim objCotacaoAnterior As New ClassCotacaoMoeda

On Error GoTo Erro_NFF_PrecoUnitario_Calcula

    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto
    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 42310
    If lErro = 28030 Then gError 42311 'Não encontrou

    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 42312

    'Converte o preço unitário para UM da tela
    dPrecoUnitario = objTabelaPrecoItem.dPreco / dFator
        
    'Aplica o acréscimo financeiro no preço unitário
    dPercAcresFin = StrParaDbl(PercAcrescFin.Text)
    dPrecoUnitario = dPrecoUnitario * (1 + dPercAcresFin / 100)
    
    If objTabelaPrecoItem.iCodTabela <> 0 Then
    
        objTabelaPreco.iCodigo = objTabelaPrecoItem.iCodTabela
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 42310
        If lErro = SUCESSO And objTabelaPreco.iMoeda <> MOEDA_REAL Then
            
            objCotacao.dtData = gdtDataAtual
            objCotacao.iMoeda = objTabelaPreco.iMoeda
            objCotacaoAnterior.iMoeda = objTabelaPreco.iMoeda
        
            'Chama função de leitura
            lErro = CF("CotacaoMoeda_Le_UltimasCotacoes", objCotacao, objCotacaoAnterior)
            If lErro <> SUCESSO Then gError 42310
            
            'Se nao existe cotacao para a data informada => Mostra a última.
            dPrecoUnitario = dPrecoUnitario * StrParaDbl(IIf(objCotacao.dValor <> 0, Format(objCotacao.dValor, "#.0000"), Format(objCotacaoAnterior.dValor, "#.0000")))
        
        End If
    
    End If
    
    NFF_PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_NFF_PrecoUnitario_Calcula:

    NFF_PrecoUnitario_Calcula = gErr

    Select Case gErr

        Case 42312, 42310

        Case 42311
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objTabelaPrecoItem.sCodProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157233)

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
        'se os percentuais (passado e calulado) forem diferentes calcula-se o desconto
        If dPercentDesc <> dPercDesc2 Then dDesconto = Arredonda_Moeda(dPercentDesc * dValorTotal)

    End If

    dValorReal = dValorTotal - dDesconto

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'Tratamento de saída de célula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 39674
        
        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridParcelas.Name

                lErro = Saida_Celula_GridParcelas(objGridInt)
                If lErro <> SUCESSO Then gError 42313

            'Se for o GridItens
            Case GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 42314

            'Se for o GridComissoes
            Case GridComissoes.Name

                lErro = objTabComissoes.Saida_Celula_GridComissoes(objGridInt)
                If lErro <> SUCESSO Then gError 42315

            'Rastreamento
            Case GridRastro.Name

                lErro = gobjRastreamento.Saida_Celula()
                If lErro <> SUCESSO Then gError 83388
                
                
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 42317

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 39674, 42313, 42314, 42315, 42317, 83388

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157234)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Produto_Col
            lErro = Saida_Celula_Produto(objGridInt)
            If lErro <> SUCESSO Then gError 42318

        Case iGrid_DescProduto_Col
            lErro = Saida_Celula_DescricaoItem(objGridInt)
            If lErro <> SUCESSO Then gError 42330

        Case iGrid_UnidadeMed_Col
            lErro = Saida_Celula_UM(objGridInt)
            If lErro <> SUCESSO Then gError 42319

        Case iGrid_Quantidade_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 42320

        Case iGrid_ValorUnitario_Col
            lErro = Saida_Celula_PrecoUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 42321

        Case iGrid_PercDesc_Col
            lErro = Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then gError 42322

        Case iGrid_Desconto_Col
            lErro = Saida_Celula_Desconto(objGridInt)
            If lErro <> SUCESSO Then gError 42323

        '###################################################
        'Inserido por Wagner
        Case iGrid_Contrato_Col
            lErro = Saida_Celula_Contrato(objGridInt)
            If lErro <> SUCESSO Then gError 132985
        
        Case iGrid_ItemContrato_Col
            lErro = Saida_Celula_ItemContrato(objGridInt)
            If lErro <> SUCESSO Then gError 132986
        '###################################################

        Case iGrid_DataCobranca_Col
            lErro = Saida_Celula_DataCobranca(objGridInt)
            If lErro <> SUCESSO Then gError 188448

        'NumNFOrig
        Case iGrid_NumNFOrig_Col
            lErro = Saida_Celula_NumNFOrig(objGridInt)
            If lErro <> SUCESSO Then gError 199733

        'SerieNFOrig
        Case iGrid_SerieNFOrig_Col
            lErro = Saida_Celula_SerieNFOriginal(objGridInt)
            If lErro <> SUCESSO Then gError 199734

        'ItemNFOrig
        Case iGrid_ItemNFOrig_Col
            lErro = Saida_Celula_ItemNFOrig(objGridInt)
            If lErro <> SUCESSO Then gError 199735

        Case iGrid_TabPrecoItem_Col
            If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
                lErro = Saida_Celula_TabPrecoItemNF(objGridInt)
                If lErro <> SUCESSO Then gError 199734
            End If

        Case iGrid_ComissaoItem_Col
            If gobjFAT.iExibeComissaoItensPV = MARCADO Then
                lErro = Saida_Celula_ComissaoItemNF(objGridInt)
                If lErro <> SUCESSO Then gError 199734
            End If

    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 42318, 42319, 42320, 42321, 42322, 42323, 42330, 132985, 132986, 132987, 188448, 199733 To 199735

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157235)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto que está deixando de ser a corrente

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
        If lErro <> SUCESSO Then gError 199359

        Produto.Text = sProduto
        
        'Executa a saida de célula
        lErro = Produto_Saida_Celula()
        If lErro <> SUCESSO Then gError 42325

    End If

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42326

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 42324, 42325, 42326, 199359
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 51087
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157236)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescricaoItem(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Descrição Item que está deixando de ser a corrente

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157237)

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
            If lErro <> SUCESSO Then gError 42336
        'VAlor da Parcela
        Case iGrid_ValorParcela_Col
            lErro = Saida_Celula_ValorParcela(objGridInt)
            If lErro <> SUCESSO Then gError 42337
        'Descontos da PArcela
        Case iGrid_Desc1Codigo_Col, iGrid_Desc2Codigo_Col, iGrid_Desc3Codigo_Col
            lErro = Saida_Celula_TipoDesconto(objGridInt)
            If lErro <> SUCESSO Then gError 42338
        'Datas de desconto da Parcela
        Case iGrid_Desc1Ate_Col, iGrid_Desc2Ate_Col, iGrid_Desc3Ate_Col
            lErro = Saida_Celula_DescontoData(objGridInt)
            If lErro <> SUCESSO Then gError 42339
        'VAlores dos descontos da parcela
        Case iGrid_Desc1Valor_Col, iGrid_Desc2Valor_Col, iGrid_Desc3Valor_Col
            lErro = Saida_Celula_DescontoValor(objGridInt)
            If lErro <> SUCESSO Then gError 42340
        'Percentuais de desconto da parcela.
        Case iGrid_Desc1Percentual_Col, iGrid_Desc2Percentual_Col, iGrid_Desc3Percentual_Col
            lErro = Saida_Celula_DescontoPerc(objGridInt)
            If lErro <> SUCESSO Then gError 42341
            
    End Select
    
    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42342
    
    Saida_Celula_GridParcelas = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_GridParcelas:

    Saida_Celula_GridParcelas = gErr
    
    Select Case gErr
    
        Case 42336, 42337, 42338, 42341, 42339, 42340
    
        Case 42342
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157238)
            
    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_DataVencimento(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Vencimento que está deixando de ser a corrente

Dim lErro As Long
Dim dtDataReferencia As Date
Dim dtDataVencimento As Date
Dim sDataVencimento As String
Dim bCriouLinha As Boolean

On Error GoTo Erro_Saida_Celula_DataVencimento

    Set objGridInt.objControle = DataVencimento

    bCriouLinha = False

    'Verifica se Data de Vencimento esta preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then gError 42343
        
         dtDataVencimento = CDate(DataVencimento.Text)
         
        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataReferencia.ClipText)) > 0 Then
            dtDataReferencia = CDate(DataReferencia.Text)
            If dtDataVencimento < dtDataReferencia Then gError 42345
        End If
        
        sDataVencimento = Format(dtDataVencimento, "dd/mm/yyyy")
        
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            bCriouLinha = True
        End If
        
    End If
    
    If sDataVencimento <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col) Then CobrancaAutomatica.Value = vbUnchecked
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42344

    If bCriouLinha Then
        'Coloca DescontoPadrao
        lErro = Preenche_DescontoPadrao(GridParcelas.Row)
        If lErro <> SUCESSO Then gError 51118
    End If

    Saida_Celula_DataVencimento = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_DataVencimento:

    Saida_Celula_DataVencimento = gErr
    
    Select Case gErr
    
        Case 42343, 42344
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 42345
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR_REFERENCIA", gErr, dtDataVencimento, GridParcelas.Row, dtDataReferencia)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 51118
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157239)
            
    End Select
    
    Exit Function

End Function


Private Sub Limpa_Tela_NFiscal2()
'Limpa a tela sem fechar setas

Dim iIndice As Integer

    '#####################################
    'Inserido por Wagner
    Set gcolcolMedicoes = New Collection
    '#####################################

    Call Limpa_Tela_NFiscal4

    
End Sub

Private Function Limpa_Tela_NFiscal() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_NFiscal

    'seta a variavel que limpa a tela como true, para que
    'o calculo de comissoes nao seja chamado na hora da limpeza da tela.
    gbLimpandoTela = True
    
    'Fecha o comando de seta se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call Limpa_Tela_NFiscal2
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    EletronicaFed.Value = vbUnchecked
    
    FilialOrcamento.Text = ""
    
    Set gobjInfoAdic = Nothing
    
    'Seleciona série padrão
    lErro = Seleciona_Serie_Padrao()
    If lErro <> SUCESSO Then gError 500133
        
    iAlterado = 0
    iComissoesAlterada = 0
    
    Call CF2(Me, "NFCF_Limpa_Tela")
    
    'volta com a variavel para false para que o calculo de comissoes
    'possa ser utilizado novamente.
    gbLimpandoTela = False
    
    Exit Function
    
Erro_Limpa_Tela_NFiscal:
    
    Limpa_Tela_NFiscal = gErr
    
    Select Case gErr
    
        Case 500133
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157240)
        
    End Select
    
    Exit Function
    
End Function

Private Function Limpa_Tela_NFiscal1() As Long
'Limpa a Tela NFiscalEntrada, mas mantém a natureza e o tipo
    
Dim sNatureza As String
Dim sTipoNFiscal As String
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_NFiscal1

    gbLimpandoTela = True
    
    sNatureza = gobjTribTab.sNatOpInterna
    sTipoNFiscal = TipoNFiscal.Text
    
    lErro = Limpa_Tela_NFiscal()
    If lErro <> SUCESSO Then gError 500134
    
    gobjTribTab.sNatOpInterna = sNatureza
    TipoNFiscal.Text = sTipoNFiscal
    Call TipoNFiscal_Validate(bSGECancelDummy)

    gbLimpandoTela = False

    Exit Function
    
    Limpa_Tela_NFiscal1 = gErr
    
Erro_Limpa_Tela_NFiscal1:

    Select Case gErr
        
        Case 500134
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157241)
        
    End Select
    
    Exit Function
    
End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 42346

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscal()
    If lErro <> SUCESSO Then gError 500135
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 42346, 500135

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157242)

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

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)
    
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

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 42348

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscal1()
    If lErro <> SUCESSO Then gError 500136
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 42348, 500136

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 157243)

    End Select

    Exit Sub

End Sub

Private Sub NFiscal_Imprime_Romaneio(objNFiscal As ClassNFiscal)

    Call CF("NFiscal_Imprime_Romaneio", objNFiscal)
    
End Sub

Public Function Gravar_Registro() As Long
'Valida os dados para gravação de uma NFiscal Fatura

Dim lErro As Long
Dim iIndice As Integer
Dim dValorTotal As Double, dFator As Double
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objNFiscal As New ClassNFiscal
Dim objNFiscalBD As New ClassNFiscal
Dim vbMsgRes As VbMsgBoxResult
Dim dValorDesconto As Double ', dValorIRRF As Double
Dim objProduto As New ClassProduto
Dim bNFNova As Boolean
Dim iNFSE As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 42349
    If Len(Trim(gobjTribTab.sNatOpInterna)) = 0 Then gError 42350
    If Len(Trim(Cliente.ClipText)) = 0 Then gError 42351
    If Len(Trim(Filial.Text)) = 0 Then gError 42352
    If Len(Trim(Serie.Text)) = 0 Then gError 42353
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 42355
    If Len(Trim(FilialEntrega.Text)) = 0 Then gError 51524
  
    'Verifica se o Peso Liqüido é maior que o Peso Bruto
    If Len(Trim(PesoLiquido.ClipText)) > 0 And Len(Trim(PesoBruto.ClipText)) > 0 Then
        If CDbl(PesoLiquido.Text) > CDbl(PesoBruto.Text) Then gError 42370
    End If

    'Verifica se a Data de Saida e menor que a Data de Emissão
    If Len(Trim(DataSaida.ClipText)) <> 0 Then
        If CDate(DataSaida.Text) < CDate(DataEmissao.Text) Then gError 42357
    End If
    
    dValorTotal = CDbl(IIf(Len(Trim(ValorTotal.Caption)) > 0, ValorTotal.Caption, 0))

    'Se o total for negativo --> Erro
    If dValorTotal < 0 Then gError 42369

    'Verifica se o Valor do Desconto é maior que a soma dos demais valores
    If Len(Trim(ValorDesconto.Text)) > 0 Then dValorDesconto = CDbl(ValorDesconto.Text)
    lErro = DescontoGlobal_Testa(dValorDesconto)
    If lErro <> SUCESSO Then gError 56900
    
    lErro = gobjTribTab.Valida_Dados
    If lErro <> SUCESSO Then gError 56900
    
    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 42358
    
    'Não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 42359

    If objTipoDocInfo.iComplementar = DOCINFO_NORMAL Then
        If objGridItens.iLinhasExistentes = 0 Then gError 42360
    End If
    
    'Valida os dados do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes
        'If objTipoDocInfo.iCodigo <> TIPODOCINFO_COMPLEMENTO_IPI Then
        If objTipoDocInfo.iComplementar = DOCINFO_NORMAL Then
            If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col)) = 0 Then gError 42366
        End If

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 42361
    Next

    'Chama Valida_Grid_Parcelas
    lErro = Valida_Grid_Parcelas()
    If lErro <> SUCESSO Then gError 42380

    'Valida os dados do grid de comissões
    lErro = objTabComissoes.Valida_Grid_Comissoes()
    If lErro <> SUCESSO Then gError 42390

    lErro = gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
    If lErro <> SUCESSO Then gError 83389

    Set objNFiscal.objRastreamento = gobjRastreamento

    'Recolhe os dados da tela
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 42371

    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Or objTipoDocInfo.iCodigo = DOCINFO_NFFISPC Or objTipoDocInfo.iCodigo = DOCINFO_NFISFVFE Then
        
        'Recolhe os dados da Nota Fiscal Original
        lErro = Move_Dados_NFOriginal_Memoria(objNFiscal)
        If lErro <> SUCESSO Then gError 42372
        
    End If

    'Move os Grids para memória
    lErro = Move_Grids_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 42402
    
    objNFiscalBD.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objNFiscalBD.sSerie = objNFiscal.sSerie
    objNFiscalBD.lNumNotaFiscal = objNFiscal.lNumNotaFiscal
    objNFiscalBD.iTipoNFiscal = objNFiscal.iTipoNFiscal
    objNFiscalBD.dtDataEmissao = objNFiscal.dtDataEmissao
    objNFiscalBD.lCliente = objNFiscal.lCliente
    objNFiscalBD.iFilialCli = objNFiscal.iFilialCli
    objNFiscalBD.lFornecedor = objNFiscal.lFornecedor
    objNFiscalBD.iFilialForn = objNFiscal.iFilialForn
    
    lErro = CF("NFiscal_Le_1", objNFiscalBD)
    If lErro <> SUCESSO And lErro <> 83971 Then gError 192923
    
    If lErro <> SUCESSO Then
        bNFNova = True
    Else
        bNFNova = False
    End If
    
    For iIndice = 1 To objNFiscal.ColItensNF.Count
        'Se for uma nota fiscal de serviço
        If objNFiscal.iTipoNFiscal = DOCINFO_NFISFS Then
                    
            objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 23080 Then gError 62102
            If lErro <> SUCESSO Then gError 62103
        
            If Not (objProduto.iCompras = PRODUTO_PRODUZIVEL And objProduto.iFaturamento = PRODUTO_VENDAVEL And objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE) Then gError 62104
               
        End If
    
        If bNFNova Then
            lErro = Valida_Contrato_Itens(objNFiscal.ColItensNF.Item(iIndice), iIndice)
            If lErro <> SUCESSO Then gError 136024
        End If
    
    Next
    
    lErro = gobjTribTab.Valida_TributacaoNF_Gravacao(objNFiscal)
    If lErro <> SUCESSO Then gError 46435
    
    lErro = Calcula_Msg
    If lErro <> SUCESSO Then gError 46435
    
    'tulio160103
    'se for pra verificar o limite de credito
    If gobjCRFAT.iCreditoVerificaLimite = VERIFICAR_LIMITE_CREDITO Then
    
        If Len(Trim(NFiscal.Caption)) = 0 Then
            'verifica se o cliente tem crédito.
            lErro = CF("NFiscal_Testa_Credito", objNFiscal)
            If lErro <> SUCESSO Then gError 45406
        End If
    
    End If
    
    'mover a parte do rastreamento
    lErro = gobjRastreamento.Move_Rastro_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 83390

    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataEmissao.Text))
    If lErro <> SUCESSO Then gError 92044
    
    lErro = CF("NFiscal_Valida_MargContr", objNFiscal)
    If lErro <> SUCESSO Then gError 201041
    
    lErro = CF("NFiscal_Valida_Diversos", objNFiscal)
    If lErro <> SUCESSO Then gError 201041
    
    'Grava a Nota Fiscal de Fatura(incluido os dados contábeis) (contabilidade)
    lErro = CF("NFiscalFatura_Grava", objNFiscal, objContabil)
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
    
    Call CF2(Me, "NFF_Grava_Registro", objNFiscal, objContabil)
    
    If Len(Trim(NFiscal.Caption)) = 0 Then vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_NUMERO_NOTA_GRAVADA", objNFiscal.lNumNotaFiscal)
    
'    If giTipoVersao = VERSAO_FULL Then
        If ImprimeRomaneio.Value = vbChecked Then
            Call CF("NFiscal_Imprime_Romaneio", objNFiscal)
        End If
'    End If
    
    If ISSerieEletronica(objNFiscal.sSerie) Then
        lErro = CF("NFE_Grava1", objNFiscal.lNumIntDoc, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 199996
    End If
    
    lErro = CF("Verifica_NFiscal_Servico_Eletronica", objNFiscal, iNFSE)
    If lErro <> SUCESSO Then gError 207051
    
    If iNFSE = 1 Then
        lErro = CF("NFSE_Grava1", objNFiscal.lNumIntDoc, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 207052
    End If
    
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 45406, 46435, 56900, 62102, 83389, 83390, 92044, 130406, 136024, 181786, 199996, 201041, 207051, 207052
                
        Case 42349
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 42350
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 42351
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 42352
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 42353
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 42355
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 42356
            Call Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_NAO_PREENCHIDA", gErr)

        Case 42357
            Call Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_ANTERIOR_DATAEMISSAO", gErr, DataSaida, DataEmissao)

        Case 42358, 42362, 42363, 42371, 42372, 42380, 42390, 42402, 35936, 41525

        Case 42359
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iTipo)

        Case 42360
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)

        Case 42361
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case 42366
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 42369
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)

        Case 42370
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr, CDbl(PesoLiquido.Text), CDbl(PesoBruto.Text))

        Case 51524
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)
        
'        Case 56905
'            Call Rotina_Erro(vbOKOnly, "ERRO_IR_FONTE_MAIOR_VALOR_TOTAL", gErr, dValorIRRF, dValorTotal)

        Case 62103
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                    
        Case 62104
           Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NOTA_SERVICO", gErr, GridItens.TextMatrix(iIndice, iGrid_Produto_Col))
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157244)

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
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 83973
    If Len(Trim(Serie.Text)) = 0 Then gError 83974
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 83975

    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
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
    
    'se origem for de empresa é interna
    If objTipoDocInfo.iOrigem = DOCINFO_EMPRESA Then
        'Verifica se a existe nota fiscal está cadastrada
        lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
        If lErro <> SUCESSO And lErro <> 43676 Then gError 99352
        
        If lErro = 43676 Then gError 89065
        
'        If gobjCRFAT.iUsaNFe = MARCADO And objNFiscal.iRecibo = 0 Then
'
'            lErro = CF("NFeFedLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'            If lErro <> SUCESSO Then gError 203016
'
'            If iEnviado = 1 Then gError 203017
'
'        ElseIf gobjCRFAT.iUsaNFSE = MARCADO And objNFiscal.iRecibo <> 0 Then
'
'            lErro = CF("RPSWEBLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'            If lErro <> SUCESSO Then gError 207074
'
'            If iEnviado = 1 Then gError 207075
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
        If lErro <> SUCESSO Then gError 83978

        'Faz o cancelamento de uma nota fiscal de Saida
        lErro = CF("NotaFiscalSaida_Excluir", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError 83979
    
        'Limpa a Tela
        lErro = Limpa_Tela_NFiscal1()
        If lErro <> SUCESSO Then gError 83980
    
        'fecha o comando de setas
        Call ComandoSeta_Fechar(Me.Name)

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 83973
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 83974
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 83975
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 84976, 83978, 83979, 83980, 99345, 99348, 99352, 99353, 203016, 207074
       
        Case 89065
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
            
        Case 99346
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)
        
        Case 99347
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)
        
        Case 99349
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objNFiscal.lCliente)
        
        Case 203017
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CADASTRADA_NAO_EXCLUI", gErr)

        Case 207075
            Call Rotina_Erro(vbOKOnly, "ERRO_NFSE_CADASTRADA_NAO_EXCLUI", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157245)
        
    End Select

    Exit Sub

End Sub

Private Function Move_Dados_NFOriginal_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os dados da NotaFiscalOriginal para a memória

Dim lErro As Long, ret As VbMsgBoxResult
Dim objNFOriginal As New ClassNFiscal
Dim vbResult As VbMsgBoxResult
Dim bCamposPreenchidos As Boolean

On Error GoTo Erro_Move_Dados_NFOriginal_Memoria

    'Verifica se a Série e o Número da NF original estão preenchidos
    If Len(Trim(NFiscalOriginal.ClipText)) = 0 And ISSerieEletronica(objNFiscal.sSerie) Then gError 42373
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 And ISSerieEletronica(objNFiscal.sSerie) Then gError 42374
    
    bCamposPreenchidos = True

    'Verifica se a Serie e o Número da NF original estão preenchidos
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Then
        bCamposPreenchidos = False
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_SERIE_NFISCAL_ORIGINAL_NAO_PREENCHIDA")
        If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
    End If
    
    If Len(Trim(NFiscalOriginal.ClipText)) = 0 And bCamposPreenchidos Then
        bCamposPreenchidos = False
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_NUMERO_NFISCAL_ORIGINAL_NAO_PREENCHIDO")
        If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
    End If

    If bCamposPreenchidos Then
    'If Len(Trim(NFiscalOriginal.ClipText)) <> 0 And Len(Trim(SerieNFiscalOriginal.Text)) <> 0 Then
    
        objNFOriginal.lCliente = objNFiscal.lCliente
        objNFOriginal.iFilialCli = objNFiscal.iFilialCli
        objNFOriginal.sSerie = SerieNFiscalOriginal.Text
        objNFOriginal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
        objNFOriginal.iFilialEmpresa = giFilialEmpresa
        
        lErro = CF("NFiscalFatura_Le_NFOriginal", objNFOriginal)
        If lErro <> SUCESSO And lErro <> 45347 Then gError 42920
        'Se não encontrar --> Erro
        If lErro = 45347 Then gError 42379
    
        'Guarda em objNfiscal o Número Interno da NF Original
        objNFiscal.lNumIntNotaOriginal = objNFOriginal.lNumIntDoc

    End If
    
    Move_Dados_NFOriginal_Memoria = SUCESSO

    Exit Function

Erro_Move_Dados_NFOriginal_Memoria:

    Move_Dados_NFOriginal_Memoria = gErr

    Select Case gErr

        Case 42373
'            ret = Rotina_Aviso(vbYesNo, "ERRO_NUMERO_NFISCAL_ORIGINAL_NAO_PREENCHIDO")
'            If ret = vbYes Then Resume Next
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NFISCAL_ORIGINAL_NAO_PREENCHIDO", gErr)
            
        Case 42374
'            ret = Rotina_Aviso(vbYesNo, "ERRO_SERIE_NFISCAL_ORIGINAL_NAO_PREENCHIDA")
'            If ret = vbYes Then Resume Next
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NFISCAL_ORIGINAL_NAO_PREENCHIDA", gErr)

        Case 42379
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_ASSOCIADA_CLIENTE", gErr, Cliente.Text, SerieNFiscalOriginal.Text, NFiscalOriginal.Text)

        Case 42920
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157246)

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
Dim dValorPagar As Double

On Error GoTo Erro_Valida_Grid_Parcelas

    'Verifica se alguma parcela foi informada
    If objGridParcelas.iLinhasExistentes = 0 Then gError 42381
    
    dSomaParcelas = 0
    
    'Para cada Parcela do grid de parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        dtDataEmissao = StrParaDate(DataEmissao.Text)
        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        
        'verifica se o vencimento e o valor da parcela estão preenchidos
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) = 0 Then gError 42382
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) = 0 Then gError 42385
        
        'verifica se a data de vencimento da parcela é menor que a data de emissão
        'If dtDataVencimento < dtDataEmissao Then gError 42383
        lErro = CF("NF_Valida_Data_Vencimento", dtDataVencimento, dtDataEmissao, iIndice)
        If lErro <> SUCESSO Then gError 140294
        
        'Se o desconto 1 da parcela está preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))) > 0 Then
            iDesconto = 1
            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
            'Verifica se a data do desconto está preenchida
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))) = 0 Then gError 51103
                        
            dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
            'Valida a data do desconto
            If dtDataDesconto > dtDataVencimento Then gError 51120
            'Se o tipo de desconto for por valor
            If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                'Verifica se o valor do desconto está preenchido
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))) = 0 Then gError 51106
            Else
                'Verifica se o percentual de desconto está preenchido
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Percentual_Col))) = 0 Then gError 51107
            End If
            'Se o desconto 2 está preenchido
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))) > 0 Then
                iDesconto = 2
                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
                'Verifica se a data de desconto está preenchida
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))) = 0 Then gError 51104
                'Verifica se a data de desconto está ordenada ou se é igual ao desconto anterior
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 51112
                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 51114
                'Se o desconto for do tipo valor
                If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                    'Verifica se o valor está preenchido
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))) = 0 Then gError 51108
                Else
                    'Verifica se o percentual está preenchido
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col))) = 0 Then gError 51109
                End If
                'Valida a data de desconto
                dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
                If dtDataDesconto > dtDataVencimento Then gError 51121
                'Se o desconto 3 está preenchido
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))) > 0 Then
                    iDesconto = 3
                    iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
                    'Verifica se a data de desconto está preenchida
                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))) = 0 Then gError 51105
                    'Verifica se a data de desconto está ordenada ou se é igual a do desconto anterior
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 51113
                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 51115
                    'Se o desconto for do tipo valor
                    If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
                        'Verifica se valor de desconto está preenchido
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))) = 0 Then gError 51110
                    Else
                        'verifica se o percentuial de desconto está preenchido
                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Percentual_Col))) = 0 Then gError 51111
                    End If
                    'Valida a data de desconto
                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
                    If dtDataDesconto > dtDataVencimento Then gError 51122
                End If
            End If
        End If
        'Verifica se as data de vencimentos das parcelas estão ordenadas
        If iIndice > 1 Then If CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col)) < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_Col)) Then gError 42384
       'Faz a soma do total das parcelas
        dSomaParcelas = dSomaParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))

    Next

    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then gError 140294
    
    If Abs(dValorPagar - dSomaParcelas) > DELTA_VALORMONETARIO Then gError 42386

    Valida_Grid_Parcelas = SUCESSO
    
    Exit Function
    
Erro_Valida_Grid_Parcelas:

    Valida_Grid_Parcelas = gErr
    
    Select Case gErr
    
        Case 42381
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_PARCELA_COBRANCA", gErr)
        
        Case 42382
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_NAO_INFORMADA", gErr, iIndice)
            
        Case 42383
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_MENOR", gErr, iIndice, dtDataVencimento, dtDataEmissao)
        
        Case 42384
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_COBRANCA_NAO_ORDENADA", gErr)
               
        Case 42385
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_PARCELA_COBRANCA_NAO_INFORMADO", gErr, iIndice)
        
        Case 42386
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_COBRANCA_INVALIDA", gErr)

        Case 51103, 51104, 51105
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_NAO_PREENCHIDA", gErr, iDesconto, iIndice)
        
        Case 51106, 51107, 51108, 51109, 51110, 51111
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_PARCELA_NAO_PREENCHIDO", gErr, iDesconto, iIndice)

        Case 51112, 51113
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTOS_DESORDENADAS", gErr, iIndice)

        Case 51114, 51115
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTO_IGUAIS", gErr, iDesconto - 1, iDesconto, iIndice)

        Case 51120, 51121, 51122
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DESC_PARCELA_SUP_DATA_VENCIMENTO", gErr, dtDataDesconto, iDesconto, iIndice)

        Case 140294

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157247)
            
    End Select
    
    Exit Function

End Function


Private Function Move_Grids_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os dados dos Grids para memória

Dim lErro As Long

On Error GoTo Erro_Move_Grids_Memoria
    
    'Chama Move_GridItens_Memoria
    lErro = Move_GridItens_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 42391

    If Frame1(TAB_Almoxarifado).Enabled = True Then
        'Chama Move_GridAlocacao_Memoria
        lErro = Move_GridAlocacao_Memoria(objNFiscal)
        If lErro <> SUCESSO Then gError 42392
    End If
    
    'Chama Move_GridComissoes_Memoria
    lErro = objTabComissoes.Move_TabComissoes_Memoria(objNFiscal, NOTA_FISCAL)
    If lErro <> SUCESSO Then gError 42393

    'Chama Move_GridParcelas_Memoria
    lErro = Move_GridParcelas_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 42394
    
    'Mover a parte de tributação
    lErro = gobjTribTab.PreencherNF(objNFiscal)
    If lErro <> SUCESSO Then gError 42395
    
    Move_Grids_Memoria = SUCESSO

    Exit Function

Erro_Move_Grids_Memoria:

    Move_Grids_Memoria = gErr

    Select Case gErr

        Case 42391, 42392, 42393, 42394, 42395

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157248)

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
Dim objItemNFOrig As ClassItemNF
Dim objItemNF1 As ClassItemNF
Dim iItem As Integer

On Error GoTo Erro_Move_GridItens_Memoria

    'Para cada linha existente do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemNF = New ClassItemNF

        'Verifica se o Produto está preenchido
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 42396

        'Armazena produto
        objItemNF.sProduto = sProdutoFormatado
              
        'Guarda os demais campos do Grid em objItemNF
        objItemNF.sDescricaoItem = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
        objItemNF.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        objItemNF.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objItemNF.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))
        objItemNF.dPercDesc = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col))

        objItemNF.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
        
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_SerieNFOrig_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_NumNFOrig_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ItemNFOrig_Col))) > 0 Then

            Set objItemNFOrig = New ClassItemNF

            objItemNFOrig.sSerieNFOrig = GridItens.TextMatrix(iIndice, iGrid_SerieNFOrig_Col)
            objItemNFOrig.lNumNFOrig = GridItens.TextMatrix(iIndice, iGrid_NumNFOrig_Col)
            objItemNFOrig.iItem = GridItens.TextMatrix(iIndice, iGrid_ItemNFOrig_Col)

'            If gbGravando Then
                lErro = CF("ItemNFiscalSaida_Le_NumNFItem1", objItemNFOrig)
                If lErro <> SUCESSO And lErro <> 35368 Then gError 199744
    
                If lErro <> SUCESSO Then gError 199745
                
                If objItemNFOrig.sProduto <> objItemNF.sProduto Then gError 199746
    
                objItemNF.lNumIntDocOrig = objItemNFOrig.lNumIntDoc
 '           End If
        End If
        
        'Adiciona na coleção de Ítens
        With objItemNF
            Set objItemNF1 = objNFiscal.ColItensNF.Add(0, iIndice, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, DATA_NULA, .sDescricaoItem, 0, 0, 0, 0, 0, colAlocacoes, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .sCcl, STATUS_LANCADO, .lNumIntDocOrig, "", 0, 0, 0, "", .iClasseUM, 0)
        End With

        Set objNFiscal.ColItensNF(iIndice).objInfoAdicDocItem = gobjNFiscal.ColItensNF(iIndice).objInfoAdicDocItem

        '********************* TRATAMENTO DE GRADE *****************
        Call Move_ItensGrade_Tela(objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade, gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade)
    
        lErro = Move_Contrato_Memoria(objNFiscal, objItemNF, iIndice)
        If lErro <> SUCESSO Then gError 132989
    
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

        Case 42396, 45242, 132989, 199744
        
        Case 199745
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NFORIGINAL_NAO_CADASTRADO", gErr, iIndice, objItemNFOrig.sSerieNFOrig, objItemNFOrig.lNumNFOrig)

        Case 199746
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CASADO", gErr, iIndice, objItemNF.sProduto, objItemNFOrig.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157249)

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
                objAlmoxarifado.sNomeReduzido = objItemAloc.sAlmoxarifado
                'Lê o Almoxarifado à partir do nome reduzido
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25060 Then gError 42398
                If lErro = 25060 Then gError 42399
                
                objItemAloc.iAlmoxarifado = objAlmoxarifado.iCodigo
                'Adiciona o almoxarifado na coleção de almoxarifados lidos
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

        Case 42398

        Case 42399
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157250)

    End Select

    Exit Function

End Function


Private Function Move_GridParcelas_Memoria(objNFiscal As ClassNFiscal) As Long
'Move as Parcelas do Grid para a Memória

Dim iIndice As Integer
Dim lTamanho As Long
Dim objParcela As ClassParcelaReceber
Dim dtDataReferencia As Date
Dim dtDataEmissao As Date
Dim lErro As Long

On Error GoTo Erro_Move_GridParcelas_Memoria

    dtDataReferencia = StrParaDate(DataReferencia.Text)
    dtDataEmissao = StrParaDate(DataEmissao.Text)

    If dtDataReferencia <> DATA_NULA Then
        lErro = CF("NF_Valida_Data_Referencia", dtDataReferencia, dtDataEmissao)
        If lErro <> SUCESSO Then gError 140330
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
            objNFiscal.ColParcelaReceber.Add 0, 0, iIndice, STATUS_ABERTO, .dtDataVencimento, .dtDataVencimento, .dValor, .dValor, 1, CARTEIRA_CARTEIRA, COBRADOR_PROPRIA_EMPRESA, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0, .iPrevisao, .sObservacao, .dValor
        End With

    Next

    objNFiscal.iNaoImpCobranca = NaoImpCobranca.Value
    
    objNFiscal.dValorDescontoTit = StrParaDbl(ValorDescontoTit.Text)

    Move_GridParcelas_Memoria = SUCESSO
    
    Exit Function

Erro_Move_GridParcelas_Memoria:

    Move_GridParcelas_Memoria = gErr
    
    Select Case gErr
    
        Case 51615
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAREFERENCIA", gErr)
        
        Case 140330 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157251)
            
    End Select

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

    For iIndice = 1 To objGridAlocacoes.iLinhasExistentes

        iItemLinha = StrParaInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))

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
                If lErro <> 25060 And lErro <> SUCESSO Then gError 42409
                
                'Se não encontrou o Almoxarifado --> erro
                If lErro = 25060 Then gError 42410
                            
                colCodigoNome.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido
            
            End If
            
            colAlocacoes.Add objAlmoxarifado.iCodigo, GridAlocacao.TextMatrix(iIndice, iGrid_AlmoxAloc_Col), StrParaDbl(GridAlocacao.TextMatrix(iIndice, iGrid_QuantAloc_Col))

        End If
    Next

    Recolhe_Alocacoes = SUCESSO

    Exit Function

Erro_Recolhe_Alocacoes:

    Recolhe_Alocacoes = gErr

    Select Case gErr

        Case 42409

        Case 42410
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157252)

    End Select

    Exit Function

End Function

Public Function Move_NF_Memoria(objNFiscal As ClassNFiscal) As Long
'Move a Nota Fiscal para a Memória

Dim lErro As Long
Dim objCliente As New ClassCliente
'Daniel 22/05/2002
Dim objOrcamentoVenda As ClassOrcamentoVenda
Dim objProjetoInfo As Object

On Error GoTo Erro_Move_NF_Memoria

    'Verifica preenchimento de Cliente
    If Len(Trim(Cliente.ClipText)) > 0 Then
    
        objCliente.sNomeReduzido = Cliente.Text
        'Lê o Cliente
        lErro = CF("Cliente_Le_Codigo_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 42411
        
        'Se não encontrou o Cliente --> erro
        If lErro = 12348 Then gError 42412
    
        objNFiscal.lCliente = objCliente.lCodigo
        
    End If
    
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
    
    'Daniel 22/05/2002 - Todo o IF
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

    objNFiscal.dtDataEmissao = MaskedParaDate(DataEmissao)
    objNFiscal.dtDataReferencia = StrParaDate(DataReferencia.Text)
    
    lErro = CF2(Me, "NFCF_Move_DataSaida_Memoria", objNFiscal)
    If lErro <> SUCESSO Then gError 132297

    'Verifica se os valores estão preenchidos antes de converter
    objNFiscal.dValorDesconto = StrParaDbl(ValorDesconto.Text)
    objNFiscal.dValorFrete = StrParaDbl(ValorFrete.Text)
    objNFiscal.dValorSeguro = StrParaDbl(ValorSeguro.Text)
    objNFiscal.dValorOutrasDespesas = StrParaDbl(ValorDespesas.Text)
    objNFiscal.dValorProdutos = StrParaDbl(ValorProdutos.Caption)
    objNFiscal.dValorTotal = StrParaDbl(ValorTotal.Caption)
    
    If Len(Trim(NFiscal.Caption)) > 0 Then
        objNFiscal.lNumNotaFiscal = CLng(NFiscal.Caption)
    Else
        objNFiscal.lNumNotaFiscal = 0
    End If
    
    If Len(Trim(VolumeQuant.ClipText)) > 0 Then objNFiscal.lVolumeQuant = StrParaLong(VolumeQuant.Text)

    'Armazena o responsável pelo frete
    objNFiscal.iFreteRespons = TipoFrete.ItemData(TipoFrete.ListIndex)
    
    'Recolhe os demais campos da tela
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
    objNFiscal.sVolumeNumero = VolumeNumero.Text
    objNFiscal.dtDataEntrada = DATA_NULA
    objNFiscal.dtDataVencimento = DATA_NULA
    objNFiscal.lNumIntNotaOriginal = 0
    objNFiscal.lNumIntDoc = 0
    objNFiscal.iCanal = Codigo_Extrai(CanalVenda.Text)
    objNFiscal.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)
    objNFiscal.iFilialEntrega = Codigo_Extrai(FilialEntrega.Text)
    
    objNFiscal.iCondicaoPagto = CondPagto_Extrai(CondicaoPagamento)

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
    
    lErro = CF2(Me, "NFF_Move_Tela_Memoria", objNFiscal)
    If lErro <> SUCESSO Then gError 185057
    
    Set objNFiscal.objInfoAdic = gobjInfoAdic
    objNFiscal.dValorItens = StrParaDbl(ValorProdutos2.Caption)
    objNFiscal.dValorDescontoItens = StrParaDbl(ValorDescontoItens.Text)
    
    Move_NF_Memoria = SUCESSO

    Exit Function

Erro_Move_NF_Memoria:

    Move_NF_Memoria = gErr

    Select Case gErr

        Case 42411, 132297, 181761, 185057 'Inserido por Wagner

        Case 42412
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case 94478
            Call Rotina_Erro(vbOKOnly, "ERRO_ORCAMENTOVENDA_NAO_CADASTRADO1", gErr, objOrcamentoVenda.lCodigo, objOrcamentoVenda.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157253)

    End Select

    Exit Function

End Function

'Caso a função esteja sendo chamada para carregar um OrcamentoVenda, receberá um objOrcamentoVenda.
Public Function Traz_NFiscal_Tela(objNFiscal As ClassNFiscal, Optional objOrcamentoVenda As ClassOrcamentoVenda) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long, objTituloRec As New ClassTituloReceber
Dim objNFiscalOriginal As New ClassNFiscal
Dim bCancel As Boolean
Dim bAlterouCT As Boolean
Dim objTipoDocInfo As New ClassTipoDocInfo

Dim lCodOrcamento As Long
Dim colTipoDocInfo As New colTipoDocInfo
Dim objTipoDocInfo1 As New ClassTipoDocInfo, iTipoNFiscal As Integer
Dim objItemNFiscal As ClassItemNF
Dim objItemNFiscalOriginal As New ClassItemNF
Dim bOV As Boolean, objFilialCliente As New ClassFilialCliente
Dim iIndice As Integer
Dim dPercDesc As Double

On Error GoTo Erro_Traz_NFiscal_Tela

    gbCarregandoTela = True

    'Limpa a tela sem fechar as setas
    Call Limpa_Tela_NFiscal2

    bOV = True
    'Se a função não for chamada para carregar os dados de um orçamento de venda. Ou seja, a NF já existe no BD.
    If (objOrcamentoVenda Is Nothing) Then
    
        bOV = False
    
        iTipoNFiscal = objNFiscal.iTipoNFiscal
        
        'Lê as mensagens
        lErro = CF("NFiscal_Le_Msg", objNFiscal)
        If lErro <> SUCESSO Then gError 35705
        
        'Lê os Ítens da Nota Fiscal
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 35705
    
        'Lê a parte de Tributação
        lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
        If lErro <> SUCESSO Then gError 39342
    
        'Lê a parte de tributacao dos Ítens da Nota Fiscal
        lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
        If lErro <> SUCESSO Then gError 39343
        
        'Lê as Comissões da Nota Fiscal
        lErro = CF("ComissoesNF_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 35703
    
        'Lê as Alocações dos Ítens da Nota Fiscal
        lErro = CF("AlocacoesNF_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 35704
        
        lErro = CF("ParcelasRecNF_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 42460
    
        'carrega dados do titulo a receber associado à nota fiscal para obter a condicao de pagto
        objTituloRec.lNumIntDoc = objNFiscal.lNumIntDocCPR
        lErro = CF("TituloReceber_Le", objTituloRec, 1)
        If lErro <> SUCESSO And lErro <> 26061 Then gError 56571
        If lErro <> SUCESSO Then
        
            lErro = CF("TituloReceberBaixado_Le", objTituloRec, 1)
            If lErro <> SUCESSO And lErro <> 56570 Then gError 56572
            If lErro <> SUCESSO Then gError 56573
        
        End If
        
        objNFiscal.iCondicaoPagto = objTituloRec.iCondicaoPagto
        
        Set objOrcamentoVenda = New ClassOrcamentoVenda
        
        objOrcamentoVenda.lNumIntNFiscal = objNFiscal.lNumIntDoc
        
        'Lê os dados do OrcamentoVenda através do seu NumIntNFiscal
        lErro = CF("OrcamentoVenda_Le_NumIntNFiscal", objOrcamentoVenda)
        If lErro <> SUCESSO And lErro <> 101305 Then gError 94485
        
    Else
    
        'Busca nos objs da Colecao o Tipo Padrao => guarda no obj
        For Each objTipoDocInfo1 In gobjCRFAT.colTiposDocInfo
            If objTipoDocInfo1.sNomeTelaNFiscal = Me.Name And objTipoDocInfo1.iPadrao = Padrao Then
                
                iTipoNFiscal = objTipoDocInfo1.iCodigo
                Exit For
            
            End If
        Next
    End If
    
    Set gobjNFiscal = objNFiscal
        
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
        giFilialAlterada = 0
        
            
    Else
            
        Cliente.Text = objOrcamentoVenda.sNomeCli
        Filial.Clear
        Filial.Text = objOrcamentoVenda.sNomeFilialCli
    
    End If

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
    
    'Preenche a Condicao de Pagto
    If objNFiscal.iCondicaoPagto > 0 Then
        PercAcrescFin.Text = ""
        CondicaoPagamento.Text = objNFiscal.iCondicaoPagto
        Call CondicaoPagamento_Validate(bSGECancelDummy)
    End If

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

    VolumeNumero.Text = objNFiscal.sVolumeNumero
    
    If objNFiscal.lNumOrcamentoVenda = 0 Then
        Call DateParaMasked(DataReferencia, objNFiscal.dtDataReferencia)
        giDataReferenciaAlterada = 0
        Call DateParaMasked(DataEmissao, objNFiscal.dtDataEmissao)
        If objNFiscal.iSemDataSaida = 0 Then
            Call DateParaMasked(DataSaida, objNFiscal.dtDataSaida)
        Else
            Call DateParaMasked(DataSaida, DATA_NULA)
        End If
    Else
        Call DateParaMasked(DataReferencia, gdtDataHoje)
        giDataReferenciaAlterada = 0
        Call DateParaMasked(DataEmissao, gdtDataHoje)
        Call DateParaMasked(DataSaida, gdtDataHoje)
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
    If lErro <> SUCESSO Then gError 132032

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
    
    If objNFiscal.lVolumeQuant > 0 Then VolumeQuant.Text = objNFiscal.lVolumeQuant
    NumPedidoTerc.Text = objNFiscal.sNumPedidoTerc

    Mensagem.Text = objNFiscal.sMensagemNota
    MensagemCorpo.Text = objNFiscal.sMensagemCorpoNota

    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = Format(objNFiscal.dPesoLiq, PesoLiquido.Format) 'Alterado por Wagner
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = Format(objNFiscal.dPesoBruto, PesoBruto.Format) 'Alterado por Wagner

    'Para cada Ítem da Nota Fiscal
    For Each objItemNFiscal In objNFiscal.ColItensNF

        objItemNFiscalOriginal.lNumIntDoc = objItemNFiscal.lNumIntDocOrig
        
        If objItemNFiscalOriginal.lNumIntDoc <> 0 Then
        
            'Lê o Ítem da NFOriginal em NFiscal
            lErro = CF("ItemNFiscal_Le", objItemNFiscalOriginal)
            If lErro <> SUCESSO And lErro <> 35225 Then gError 199747
    
            If lErro <> SUCESSO Then gError 199748
    
            objItemNFiscal.iItemNFOrig = objItemNFiscalOriginal.iItem
    
            objNFiscalOriginal.lNumIntDoc = objItemNFiscalOriginal.lNumIntNF
            
            'Lê a Nota Fiscal do Ítem Lido (Original)
            lErro = CF("NFiscal_Le", objNFiscalOriginal)
            If lErro <> SUCESSO Then gError 199749
    
            'Preenche a Série e o Número com os lidos da Nota
            objItemNFiscal.sSerieNFOrig = objNFiscalOriginal.sSerie
            objItemNFiscal.lNumNFOrig = objNFiscalOriginal.lNumNotaFiscal

        End If
        
    Next

    'Preenche o Grid com os Ítens da Nota Fiscal
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 39023
    
    objTipoDocInfo.iCodigo = iTipoNFiscal
     
    'Le o Tipo para saber se tem Movimento de Estoque
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO Then gError 83384
    
    If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
        'Carrega ItensNF com Rastreamentos
        lErro = gobjRastreamento.Carrega_RastroItensNF(objNFiscal)
        If lErro <> SUCESSO Then gError 83385
    End If
    
    'Carrega o Tab Comissões
    lErro = objTabComissoes.Carrega_Tab_Comissoes(objNFiscal)
    If lErro <> SUCESSO Then gError 39022
    
    'Preenche o Grid com as Alocações dos ítens da Nota Fiscal
    lErro = Preenche_GridAlocacoes(objNFiscal)
    If lErro <> SUCESSO Then gError 39026
    
    'Preenche o Grid de Parcelas
    NaoImpCobranca.Value = objNFiscal.iNaoImpCobranca
    lErro = Preenche_Grid_Parcelas(objNFiscal)
    If lErro <> SUCESSO Then gError 42467

    'Se o código da Nota Fiscal Original estiver preenchido
    If objNFiscal.lNumIntNotaOriginal <> 0 Then
        
        objNFiscalOriginal.lNumIntDoc = objNFiscal.lNumIntNotaOriginal
        'Lê a Nota Fiscal Original
        lErro = CF("NFiscal_Le", objNFiscalOriginal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 39024

        'Se não encontrou a NFiscal Original --> erro
        If lErro <> SUCESSO Then gError 39025
        
        'Coloca os dados relativos  Nota Fiscal Original na Tela
        SerieNFiscalOriginal.Text = objNFiscalOriginal.sSerie
        NFiscalOriginal.Text = objNFiscalOriginal.lNumNotaFiscal
        
    End If
    
    lErro = CF2(Me, "NFCF_Cliente_Traz_NF_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 131887
    
    lErro = CF2(Me, "NFF_Traz_NFiscal_Tela", objNFiscal, objTituloRec)
    If lErro <> SUCESSO Then gError 185058
    
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

    ''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 39031
    '*** fim tributacao
    
    'Traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 39673
    
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
    
    'Se veio do orçamento de venda
    If Not (objOrcamentoVenda Is Nothing) Then
        'E não usa regra de comissão
        If Not gobjCRFAT.iUsaComissoesRegras = USA_REGRAS Then

            objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
            'Lê a FilialCliente
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilialCliente)
            If lErro <> SUCESSO Then gError 39673

            objTabComissoes.Limpa_Grid_Comissoes
            'calcular as comissoes no modelo antigo...
            lErro = objTabComissoes.Comissao_Automatica_FilialCli_Exibe(objFilialCliente)
            If lErro <> SUCESSO Then gError 39673
        
        End If

    End If

    ValorDescontoTit.Text = Format(objNFiscal.dValorDescontoTit, "Standard")
    Call ValorDescontoTit_Validate(bSGECancelDummy)
    
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
    
    iComissoesAlterada = 0
    iAlterado = 0
    iPercAcresFinAlterado = 0
     
    gbCarregandoTela = False
     
    Traz_NFiscal_Tela = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela:

    gbCarregandoTela = False

    Traz_NFiscal_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 25821, 25822, 39031, 35705, 39023, 39024, 39022, 39342, 39343, 39673, 42460, 42467, 56571, 56572, 83384, 83385, 101305, 131887, 185058, 199747, 199749, 207551

        Case 56573
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_REC_INEXISTENTE", gErr)
        
        Case 39025
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntNotaOriginal)

        Case 132032

        Case 199748
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NFORIGINAL_NAO_CADASTRADO2", gErr, objItemNFiscalOriginal.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157254)

    End Select

    Exit Function

End Function

Public Function Preenche_GridItens(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com os Itens da Nota Fiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim sProdutoEnxuto As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objProduto As New ClassProduto
Dim iItem As Integer

On Error GoTo Erro_Preenche_GridItens

    iIndice = 0

    'Para cada ítem da Coleção
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
        
        If objItemNF.lNumNFOrig > 0 Then GridItens.TextMatrix(iIndice, iGrid_NumNFOrig_Col) = objItemNF.lNumNFOrig
        If Len(Trim(objItemNF.sSerieNFOrig)) > 0 Then GridItens.TextMatrix(iIndice, iGrid_SerieNFOrig_Col) = objItemNF.sSerieNFOrig
        If objItemNF.iItemNFOrig > 0 Then GridItens.TextMatrix(iIndice, iGrid_ItemNFOrig_Col) = objItemNF.iItemNFOrig


        lErro = CF2(Me, "NF_Preenche_GridItens", objNFiscal, objItemNF, objProduto, iIndice)
        If lErro <> SUCESSO Then gError 132988
        
        lErro = Traz_Contrato_Tela(objItemNF, iIndice)
        If lErro <> SUCESSO Then gError 132988
    
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

        Case 132988, 177926, 181676 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157255)

    End Select

    Exit Function

End Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157256)
            
    End Select

End Function

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
        
    Next
    
    objGridParcelas.iLinhasExistentes = iIndice
    
    Preenche_Grid_Parcelas = SUCESSO
    
    Exit Function

End Function

Public Sub Form_Activate()

Dim lErro As Long

On Error GoTo Erro_Form_Activate

    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError 59335
        
    Call TelaIndice_Preenche(Me)

    Exit Sub
     
Erro_Form_Activate:

    Select Case gErr
          
        Case 59335
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157257)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

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

'****
Public Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

'****
Public Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

'****
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
Dim objCliente As New ClassCliente
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim sContaMascarada As String
Dim iAlmoxPadrao As Integer
Dim bEncontrouProduto As Boolean
Dim iLinha2 As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim dQuantidade As Double
Dim dQuantidadeConvertida As Double
Dim objItemMovEstoque As New ClassItemMovEstoque
Dim sProduto As String
Dim sUM As String
Dim objProdutoCategoria As New ClassProdutoCategoria
Dim objCatProdItem As ClassCategoriaProdutoItem
Dim objPlanoConta As New ClassPlanoConta
Dim objProduto As New ClassProduto
Dim sContaTela As String
Dim objFilial As New ClassFilialCliente
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
                    If lErro <> SUCESSO Then gError 79009
                    
                    objItemMovEstoque.sProduto = sProdutoFormatado
                    
                    'Define o almoxarifado que será passado como parâmetro para MovEstoque_Le_QuantVendConsig
                    objAlmoxarifado.sNomeReduzido = GridAlocacao.TextMatrix(iLinha, iGrid_AlmoxAloc_Col)
                    
                    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25060 Then gError 79008
                    
                    objItemMovEstoque.iAlmoxarifado = objAlmoxarifado.iCodigo
                    
                    'Define o tipo de movimento, o DocOrigem e o TipoNumIntoDocOrigem que serão passados como parâmetros para MovEstoque_Le_QuantVendConsig
                                        
                    objItemMovEstoque.iTipoMov = MOV_EST_NF_VENDA_MAT_CONSIG
                    objItemMovEstoque.sDocOrigem = "NFISFV" & " " & Serie.Text & " " & objContabil.lNumNotaFiscalExt
                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                                                           
                    'Verifica se MovEstoque_Le_QuantVendConsig não encontrou erro
                    lErro = CF("MovEstoque_Le_QuantVendConsig", objItemMovEstoque)
                    If lErro <> SUCESSO And lErro <> 79003 Then gError 79007
                
                    'se não conseguiu encontrar mov. estoque para os parametros em questao ==> quantidade é zerada
                    If lErro = 79003 Then objItemMovEstoque.dQuantidade = 0
                
                    'Passa para o mnemônico o valor encontrado por MovEstoque_Le_QuantVendConsig
                    objMnemonicoValor.colValor.Add objItemMovEstoque.dQuantidade
                    
                Else
                    objMnemonicoValor.colValor.Add 0
                
                End If
            
            Next
                    
        
        Case QUANT_DISPONIVEL, QUANT_CONSIGTERC
        
            For iLinha = 1 To objGridItens.iLinhasExistentes
                    
                lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                If lErro <> SUCESSO Then gError 75467
                
                lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, sProdutoFormatado, iAlmoxPadrao)
                If lErro <> SUCESSO Then gError 75465
                    
                If iAlmoxPadrao > 0 Then
                    
                    'Lê as quantidades do Produto no almoxarifado
                    objEstoqueProduto.iAlmoxarifado = iAlmoxPadrao
                    objEstoqueProduto.sProduto = sProdutoFormatado
                    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
                    If lErro <> SUCESSO And lErro <> 21306 Then gError 75468
                    
                    'Se não encontrou Estoque Produto, erro
                    If lErro = 21306 Then gError 75469
                                        
                    If Len(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
                                            
                        If objMnemonicoValor.sMnemonico = QUANT_DISPONIVEL Then
                            dQuantidade = objEstoqueProduto.dQuantDisponivel
                        ElseIf objMnemonicoValor.sMnemonico = QUANT_CONSIGTERC Then
                            dQuantidade = objEstoqueProduto.dQuantConsig3
                        End If
                        
                        lErro = CF("UMEstoque_Conversao", sProdutoFormatado, GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), dQuantidade, dQuantidadeConvertida)
                        If lErro <> SUCESSO Then gError 75470

                        objMnemonicoValor.colValor.Add dQuantidadeConvertida
                    Else
                        objMnemonicoValor.colValor.Add 0
                    End If
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
            
        Case DESCONTO1, "Desc_Prod"
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
            If Len(Cliente.Text) > 0 Then
            
                objCliente.sNomeReduzido = Cliente.Text
                
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then gError 88028
                
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
            
        Case PERC_ACRESC_FIN
            If Len(PercAcrescFin.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(PercAcrescFin.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALOR_PARCELA
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                If Len(GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col))
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
                    If lErro <> SUCESSO Then gError 183767
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183768
        
                    sContaMascarada = String(STRING_CONTA, 0)

                    lErro = Mascara_MascararConta(objProduto.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError 183792

                    objMnemonicoValor.colValor.Add sContaMascarada
        
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case CONTACONTABILCATPROD
                
            For iLinha = 1 To objGridItens.iLinhasExistentes
                
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183769
                
                    objProdutoCategoria.sCategoria = objMnemonicoValor.vParam(1)
                    
                    objProdutoCategoria.sProduto = sProdutoFormatado
                    
                    Set objCatProdItem = New ClassCategoriaProdutoItem
                    
                    lErro = CF("ProdutoCategoria_Le_Item", objProdutoCategoria, objCatProdItem)
                    If lErro <> SUCESSO And lErro <> 183548 Then gError 183770
                    
                    If lErro = SUCESSO Then
                        
                        Set objPlanoConta = New ClassPlanoConta
                    
                        'le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta
                        lErro = CF("PlanoConta_Le_ContaSimples", objCatProdItem.dvalor1, objPlanoConta)
                        If lErro <> SUCESSO And lErro <> 5451 Then gError 183771
                        
                        If lErro = SUCESSO Then
                    
                            sContaMascarada = String(STRING_CONTA, 0)
        
                            lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
                            If lErro <> SUCESSO Then gError 183793
        
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
                If lErro <> SUCESSO Then gError 186134
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 186135
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
        
        Case Else
        
            lErro = gobjTribTab.Calcula_Mnemonico(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 39675

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39675
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case 64179, 64180, 64181, 64182, 64183, 64184, 64185, 75465, 75467, 75468, 75470, 79007, 79008, 79009, 88028, 183767 To 183771, 183792, 183793, 186134, 186135
                    
        Case 75469
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157258)

    End Select

    Exit Function

End Function

'Private Function TabelaPreco_Atualiza_PrecoUnitario(objProduto As ClassProduto, objRomaneioGrade As ClassRomaneioGrade) As Long
'
'Dim lErro As Long
'Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
'Dim dPrecoUnitario As Double
'
'On Error GoTo Erro_TabelaPreco_Atualiza_PrecoUnitario
'
'    'Preco Unitário
'    If Len(Trim(TabelaPreco.Text)) > 0 Then
'
'       If gobjCRFAT.iSeparaItensGradePrecoDif = MARCADO And Not (objRomaneioGrade Is Nothing) Then
'
'            Set objTabelaPrecoItem = objRomaneioGrade.objTabelaPrecoItem
'
'            If objTabelaPrecoItem.dPreco < DELTA_VALORMONETARIO Then lErro = 28014
'
'        Else
'
'            objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
'            objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
'            objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
'
'            'Lê o Item da Tabela de Preço
'            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
'            If lErro <> SUCESSO And lErro <> 28014 Then gError 42187
'
'            If lErro = 28014 Then
'                objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
'
'                'Lê o Item da Tabela de Preço da Empresa Toda
'                lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
'                If lErro <> SUCESSO And lErro <> 28014 Then gError 42188
'
'            End If
'
'        End If
'
'        If lErro <> 28014 Then
'
'            'Calcula o Preço Unitário
'            lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
'            If lErro <> SUCESSO Then gError 42189
'
'            GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
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
'        Case 42187, 42188, 42189
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157259)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Sub Seleciona_Tipo_Venda()
'Seleciona na combo de tipos de notas fiscais o tipo "Venda"

Dim iIndice As Integer

    For iIndice = 0 To TipoNFiscal.ListCount - 1
        If TipoNFiscal.ItemData(iIndice) = DOCINFO_NFISFV Then
            TipoNFiscal.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Exit Sub
        
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
        If lErro <> SUCESSO Then gError 51116
        
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
    
        Case 51116
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157260)
            
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'tulio261202
    'feito em virtude do redimensionamento da msg da nota p/ inpal
    Call CF("Redimensiona_Tam_MsgNF", Mensagem)

    Parent.HelpContextID = IDH_NF_FATURA_SAIDA_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Nota Fiscal Fatura de Saída"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "NFiscalFatura"
    
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
Dim objTipoDocInfo As New ClassTipoDocInfo
    
Dim lErro As Long

On Error GoTo Erro_DescontoGlobal_Testa

    If Len(Trim(ValorProdutos.Caption)) > 0 Then dValorProdutos = CDbl(ValorProdutos.Caption)
    If Len(Trim(ValorFrete.ClipText)) > 0 Then dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorSeguro.ClipText)) > 0 Then dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(ValorDespesas.ClipText)) > 0 Then dValorDespesas = CDbl(ValorDespesas.Text)

    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 130756

    'se não estiver cadastrado ==> erro
    If lErro = 31415 Then gError 130757
    
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then
    
        'Verifica se o Valor do Desconto é maior que a soma dos demais valores
        If dValorProdutos + dValorFrete + dValorSeguro + dValorDespesas < 0 Then gError 56899 'Alterado por Wagner
        
    Else
    
        'Verifica se o Valor do Desconto é maior que a soma dos demais valores
        If dValorProdutos + dValorFrete + dValorSeguro + dValorDespesas <= 0 Then gError 56899 'Alterado por Wagner
    
    End If
    
    DescontoGlobal_Testa = SUCESSO
     
    Exit Function
    
Erro_DescontoGlobal_Testa:

    DescontoGlobal_Testa = gErr
     
    Select Case gErr
                  
        Case 130756
        
        Case 130757
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 56899
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_ULTRAPASSOU_SOMA_VALORES", gErr, dValorDesconto, dValorProdutos + dValorDesconto + dValorFrete + dValorSeguro + dValorDespesas) 'Alterado por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157261)
     
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
        ElseIf Me.ActiveControl Is Contrato Then
            Call BotaoItemContrato_Click
        Else
            Call gobjRastreamento.objUserControl_KeyDown(KeyCode, Shift)
        End If
    
    End If

End Sub


Public Sub VolumeQuant_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VolumeQuant, iAlterado)

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
        If lErro <> SUCESSO Then gError 26713

        'Compara com data de emissão
        If Len(Trim(DataEmissao.ClipText)) > 0 Then

            dtDataEmissao = CDate(DataEmissao.Text)
            dtDataReferencia = CDate(DataReferencia.Text)

            lErro = CF("NF_Valida_Data_Referencia", dtDataReferencia, dtDataEmissao)
            If lErro <> SUCESSO Then gError 140329
            
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

        Case 26713, 140329

        Case 26714
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAREFERENCIA", gErr, dtDataReferencia, dtDataEmissao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157262)

    End Select

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

Public Sub BotaoConsultaTitRec_Click()
'Abre uma tela para consulta do DocCPR vinculado à nota fiscal

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoConsultaTitRec_Click

    'Verifica se todos os campos necessários para se efetuar a consulta foram preenchidos
    lErro = Critica_CamposNecessarios_ConsultaTitulo()
    If lErro <> SUCESSO Then gError 79795
    
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
    If lErro <> SUCESSO And lErro <> 12348 Then gError 79796
    
    'Se não encontrou o fornecedor = > erro
    If lErro = 12348 Then gError 79797
    
    'Guarda no objNFiscal o código do fornecedor
    objNFiscal.lCliente = objCliente.lCodigo
    
    'Lê o NumIntDocCPR da NFiscal e exibe o documento gerado no CPR por essa nota
    lErro = CF("NFiscal_Consulta_DocCPR", objNFiscal)
    If lErro <> SUCESSO And lErro <> 79717 Then gError 79798
    
    'Se não encontrou a nota => erro
    If lErro = 79717 Then gError 79799
    
    Exit Sub
    
Erro_BotaoConsultaTitRec_Click:

    Select Case gErr
    
        Case 79795, 79798, 79796
        
        Case 79799
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA2", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada)
        
        Case 79797
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157263)
    
    End Select
    
    Exit Sub

End Sub

Private Function Critica_CamposNecessarios_ConsultaTitulo() As Long
'Verifica se os campos necessários para encontrar consultar um título gerado por uma NFFatEntrada foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposNecessarios_ConsultaTitulo

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.Text)) = 0 Then gError 79800
    
    'Verifica se a filial do Fornecedor foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 79801
    
    'Verifica se a Série foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 79802
    
    'Verifica se o Número da Nota foi preenchido
    If Len(Trim(NFiscal.Caption)) = 0 Then gError 79803
    
    'Verifica se a data de emissão da nota foi preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 79804
    
    'Verifica se o tipo da Nota fiscal foi selecionado
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 79805
    
    Critica_CamposNecessarios_ConsultaTitulo = SUCESSO
    
    Exit Function

Erro_Critica_CamposNecessarios_ConsultaTitulo:

    Critica_CamposNecessarios_ConsultaTitulo = gErr
    
    Select Case gErr
    
        Case 79800
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 79801
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 79802
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 79803
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
        
        Case 79804
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 79805
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
        
        Case Else
            
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157264)
    
    End Select
        
    Exit Function
        
End Function

'Transferir para ClassCRFAT
Function Cliente_Pesquisa_NomeReduzido(objCliente As Object, sNomeReduzidoParte As String) As Long
'rotina para trazer o nome do cliente com uma parte dos caracteres digitados

Dim iCursorPos As Integer
Dim sNomeReduzido As String
Dim lErro As Long

On Error GoTo Erro_Cliente_Pesquisa_NomeReduzido

        If sNomeReduzidoParte <> objCliente.Text And Len(sNomeReduzidoParte) <= Len(objCliente.Text) Then
            
            sNomeReduzidoParte = objCliente.Text
            
            lErro = Cliente_Le_NomeReduzido_Parcial(sNomeReduzidoParte, sNomeReduzido)
            If lErro <> SUCESSO Then gError 71652
            
            If sNomeReduzido <> objCliente.Text Then
            
                sNomeReduzidoParte = sNomeReduzido
                iCursorPos = objCliente.SelStart
                objCliente.Text = sNomeReduzido
                objCliente.SelStart = iCursorPos
                objCliente.SelLength = Len(objCliente.Text) - (iCursorPos)
        
            End If
        
        Else
        
            sNomeReduzidoParte = objCliente.Text
        
        End If

    Cliente_Pesquisa_NomeReduzido = SUCESSO

    Exit Function

Erro_Cliente_Pesquisa_NomeReduzido:

    Cliente_Pesquisa_NomeReduzido = gErr

    Select Case gErr

        Case 71652

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157265)

    End Select

End Function

'Transferir para ClassCRFAT
Function Cliente_Le_NomeReduzido_Parcial(sNomeReduzidoParte As String, sNomeReduzido As String) As Long
'Le o NomeReduzido do cliente a partir de parte do nome reduzido. Se a parte não individualiza o cliente ==> retorna codigo de erro

Dim lComando As Long
Dim lErro As Long
Dim iTamanho As Integer
Dim sNomeReduzidoAux As String

On Error GoTo Erro_Cliente_Le_NomeReduzido_Parcial

    sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)

    iTamanho = Len(sNomeReduzidoParte)

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 71645

    'Le o NomeReduzido da tabela Clientes
    lErro = Comando_Executar(lComando, "SELECT NomeReduzido FROM Clientes WHERE NomeReduzido >= ?", sNomeReduzido, sNomeReduzidoParte)
    If lErro <> AD_SQL_SUCESSO Then gError 71646

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71647

    'se não encontrou nenhum cliente com este prefixo ==> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 71648

    'se encontrou um cliente, mas o prefixo difere ==> erro
    If UCase(sNomeReduzidoParte) <> UCase(left(sNomeReduzido, iTamanho)) Then gError 71649

    sNomeReduzidoAux = sNomeReduzido
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71650

    If lErro = AD_SQL_SUCESSO Then
        
        'se encontrou um segundo registro com o mesmo prefixo ==> ainda não pode trazer o nome pois ainda não conseguiu individualizar
        If UCase(sNomeReduzidoParte) = UCase(left(sNomeReduzido, iTamanho)) Then gError 71651
        
    End If
    
    sNomeReduzido = sNomeReduzidoAux

    lErro = Comando_Fechar(lComando)

    Cliente_Le_NomeReduzido_Parcial = SUCESSO

    Exit Function

Erro_Cliente_Le_NomeReduzido_Parcial:

    Cliente_Le_NomeReduzido_Parcial = gErr

    Select Case gErr

        Case 71645
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 71646, 71647, 71650
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)

        Case 71648, 71649, 71651

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157266)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'******************************************
'4 eventos do controle do Grid de Comissoes: DiretoIndireto
'******************************************

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157267)

    End Select

End Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157268)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157269)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157270)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157271)

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
        
'    'chamei essa func. para preencher o grid de comissoes tulio241002
'    Cliente.Text = objNFiscal.lCliente
'    Call Cliente_Validate(bSGECancelDummy)
'    Filial.Text = objNFiscal.iFilialCli
'    Call Filial_Validate(bSGECancelDummy)
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157272)

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
        '.dtDataSaida = gdtDataAtual 'Comentado por Wagner
        
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
    
    lErro = ParcelasOV_Gerando_ParcelasNF(objOrcamentoVenda, objNFiscal)
    If lErro <> SUCESSO Then gError 51562

    'Para cada item do Orcamento
    For Each objItemOV In objOrcamentoVenda.colItens
        
        iIndice = iIndice + 1
        
        If objItemOV.dQuantidade > 0 Then
            
            If objItemOV.dPrecoTotal > 0 Then
                
                'Calcula o percentual de desconto
                dPercDesc = objItemOV.dValorDesconto / (objItemOV.dPrecoTotal + objItemOV.dValorDesconto)
            End If

            Call objTribTab.ItemNFiscal_ObterTrib_ItemOrcamento(objItemOV, objItemTributacaoNF)
            
            With objItemOV
                                
                'Adiciona o item na coleção de itens da NF
                objNFiscal.ColItensNF.Add 0, iIndice, .sProduto, .sUnidadeMed, .dQuantidade, IIf(dCotacao = 1, .dPrecoUnitario, Arredonda_Moeda(.dPrecoUnitario * dCotacao)), dPercDesc, IIf(dCotacao = 1, .dValorDesconto, Arredonda_Moeda(.dValorDesconto * dCotacao)), .dtDataEntrega, .sDescricao, 0, .lNumIntDoc, 0, 0, .lNumIntDoc, Nothing, 0, "", "", .iStatus, 0, "", 0, 0, 0, "", .iClasseUM, 0
                
                Set objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).objTributacaoItemNF = objItemTributacaoNF
                
                Set objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).objInfoAdicDocItem = objItemOV.objInfoAdicDocItem
                
                If objItemOV.iPrazoEntrega <> 0 Then
                    dtDataBase = StrParaDate(DataEmissao.Text)
                    If dtDataBase = DATA_NULA Then dtDataBase = gdtDataAtual
                    
                    lErro = CF("Calcula_Data", dtDataBase, objItemOV.iPrazoEntrega, 0, CONDPAGTO_TIPOINTERVALO_DIAS_UTEIS, dtDataAux)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                    objNFiscal.ColItensNF.Item(objNFiscal.ColItensNF.Count).dtDataEntrega = dtDataAux
                End If
                
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

        Case 51562, 180034, 180035, 180036

        Case 211630
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_SEM_COTACAO", gErr, objCotacao.iMoeda)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157273)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157274)

    End Select

    Exit Sub

End Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157275)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157276)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157277)
            
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157278)
            
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157279)

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

Public Property Get Botao() As Object
     Set Botao = objUserControl.Controls("Botao")
End Property

Public Property Get BotaoGrade() As Object
     Set BotaoGrade = Botao(BOTAO_GRADE)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157280)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157281)

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
    If lErro <> SUCESSO Then gError 134008

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134008

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157282)

    End Select
    
    Exit Sub

End Sub

Public Sub Item_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Item_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Item_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Item_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Item
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Contrato_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Contrato_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub Contrato_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub Contrato_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Contrato
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataCobranca_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataCobranca_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub DataCobranca_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub DataCobranca_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataCobranca
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub BotaoItemContrato_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objItemDeContrato As New ClassItensDeContrato
Dim sProduto As String
Dim objCliente As New ClassCliente
Dim lCliente As Long
Dim iFilial As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sContrato As String
Dim iSeq As Integer

On Error GoTo Erro_BotaoItemContrato_Click

    If GridItens.Row = 0 Then gError 132930

    sProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
    
    If Me.ActiveControl Is Contrato Then
        sContrato = Contrato.Text
    Else
        sContrato = GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col)
    End If
    
    iSeq = StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_ItemContrato_Col))
      
    If Len(Trim(Cliente.Text)) = 0 Then gError 132932
    If Len(Trim(Filial.Text)) = 0 Then gError 132933
  
    objCliente.sNomeReduzido = Cliente.Text
    
    'Lê o fornecedor
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 132934
    
    If lErro = 12348 Then gError 132935 'Se nao encontrar --> erro

    lCliente = objCliente.lCodigo

    If Len(Trim(Filial.Text)) > 0 Then
        iFilial = Codigo_Extrai(Filial.Text)
    End If
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 132936
    
    colSelecao.Add lCliente
    colSelecao.Add lCliente
    colSelecao.Add iFilial
    colSelecao.Add iFilial
    colSelecao.Add sProdutoFormatado
    colSelecao.Add sProdutoFormatado
    colSelecao.Add sContrato
    colSelecao.Add sContrato
    
    objItemDeContrato.sContrato = sContrato
    objItemDeContrato.iSeq = iSeq
    
    Call Chama_Tela("ContratosCliItensLista", colSelecao, objItemDeContrato, objEventoItemContrato)

    Exit Sub
    
Erro_BotaoItemContrato_Click:
    
    Select Case gErr
    
        Case 132930
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
      
        Case 132932
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 132933
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 132934, 132936
        
        Case 132935
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, Cliente.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157283)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoItemContrato_evSelecao(obj1 As Object)

Dim lErro As Long
Dim obItensDeContrato As New ClassItensDeContrato
Dim bCancel As Boolean
Dim sProdutoEnxuto As String
Dim objProduto As New ClassProduto
Dim objItemMedicao As New ClassItensMedCtr
Dim iIndice As Integer

On Error GoTo Erro_objEventoItemContrato_evSelecao

    'verifica se tem alguma linha do Grid selecionada
    If GridItens.Row = 0 Then gError 136159

    Set obItensDeContrato = obj1

    'Se mudou o contrato ou o item
'    If GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col) <> obItensDeContrato.sContrato Or _
'        StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_ItemContrato_Col)) <> obItensDeContrato.iSeq Then
    
        'Busca repetições
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If iIndice <> GridItens.Row Then
                If GridItens.TextMatrix(iIndice, iGrid_Contrato_Col) = obItensDeContrato.sContrato And _
                    StrParaInt(GridItens.TextMatrix(iIndice, iGrid_ItemContrato_Col)) = obItensDeContrato.iSeq Then
                    gError 136225
                End If
            End If
        Next
        
        GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col) = obItensDeContrato.sContrato
        GridItens.TextMatrix(GridItens.Row, iGrid_ItemContrato_Col) = obItensDeContrato.iSeq
    
        If Me.ActiveControl Is Contrato Then Contrato.Text = obItensDeContrato.sContrato
        
        'Verifica se o Produto está preenchido
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then
    
            objProduto.sCodigo = obItensDeContrato.sProduto
    
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 136177
    
            'Se o produto não está cadastrado, erro
            If lErro = 28030 Then gError 136178
            
            lErro = Mascara_RetornaProdutoEnxuto(obItensDeContrato.sProduto, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 136160
    
            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True
            
            'Coloca o produto, a Descrição e a Unidade de Medida da tela
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
    
            lErro = Produto_Saida_Celula(objProduto)
            If lErro <> SUCESSO Then gError 136161
            
        End If
            
        'Se for um Item com medição
        If obItensDeContrato.iMedicao = MARCADO Then
        
            'Chama tela para obter as medições
            Call Chama_Tela_Modal("MedicoesNFiscal", gcolcolMedicoes.Item(GridItens.Row), obItensDeContrato)
            If giRetornoTela = vbCancel Then gError 136200
            If giRetornoTela = vbOK Then
            
                'Agrupa o Valor e quantidade das medições
                Call Obtem_ItemMedicao(objItemMedicao, GridItens.Row)
                
                If objItemMedicao.dQuantidade > 0 Then
                    GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(objItemMedicao.dVlrCobrar / objItemMedicao.dQuantidade, "STANDARD")
                    GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(objItemMedicao.dQuantidade)
                End If
                
            End If
            
        Else 'è um item de contrato que não tem medições
        
'            GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(obItensDeContrato.dValor, "STANDARD")
'            GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(obItensDeContrato.dQuantidade)
        
        End If
                
        GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = obItensDeContrato.sUM
                
        If Frame1(TAB_Almoxarifado).Enabled = True Then
            lErro = Alocacao_Processa()
            If lErro <> SUCESSO Then gError 136162
        End If
        
        Call PrecoTotal_Calcula(GridItens.Row, False)
        
        '*** incluidos p/tratamento de tributacao ****************
        If GridItens.Row <= objGridItens.iLinhasExistentes Then
            lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
            If lErro <> SUCESSO Then gError 136163
        End If
        '*** fim tributacao
        
        Call ValorTotal_Calcula
    
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 136164
        
        Me.Show
        
'    End If
    
    Exit Sub
    
Erro_objEventoItemContrato_evSelecao:

    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    Select Case gErr
    
        Case 136159, 136161, 136162, 136163, 136164, 136177, 136200
        
        Case 136160
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, obItensDeContrato.sProduto)

        Case 136178
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 136225
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_CONTRATO_REPETIDO", gErr, obItensDeContrato.sContrato, obItensDeContrato.iSeq, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157284)

    End Select

    Exit Sub
    
End Sub

Public Function Move_Contrato_Memoria(ByVal objNFiscal As ClassNFiscal, ByVal objItemNF As ClassItemNF, iLinha As Integer) As Long
'Move os Itens do Grid para a Memória

Dim lErro As Long
Dim objItemContrato As New ClassItensDeContrato
Dim objContrato As New ClassContrato
Dim sContrato As String
Dim iSeq As Integer
Dim objItemMedicao As New ClassItensMedCtr
Dim dtDataCobranca As Date

On Error GoTo Erro_Move_Contrato_Memoria

    Set objItemNF.objCobrItensContrato = New ClassCobrItensContrato

    sContrato = GridItens.TextMatrix(iLinha, iGrid_Contrato_Col)
    iSeq = StrParaInt(GridItens.TextMatrix(iLinha, iGrid_ItemContrato_Col))
    dtDataCobranca = StrParaDate(GridItens.TextMatrix(iLinha, iGrid_DataCobranca_Col))

    'Se está ligado a contrato
    If Len(Trim(sContrato)) <> 0 Then

        objContrato.sCodigo = sContrato
        objContrato.iFilialEmpresa = giFilialEmpresa
    
        'Le o contrato
        lErro = CF("Contrato_Le", objContrato)
        If lErro <> 129332 And lErro <> SUCESSO Then gError 132941
        
        'Se não está cadastrado = > Erro
        If lErro = 129332 Then gError 132942
        
        'Se mudou o cliente => Erro
        If objNFiscal.lCliente <> objContrato.lCliente Then gError 136180
    
        objItemContrato.lNumIntContrato = objContrato.lNumIntDoc
        objItemContrato.iSeq = iSeq
    
        'Le o item de contrato
        lErro = CF("ItensDeContrato_Le2", objItemContrato)
        If lErro <> 129266 And lErro <> SUCESSO Then gError 132943
        
        'Se não estiver cadastado => erro
        If lErro = 129266 Then gError 132944
                
        objItemNF.objCobrItensContrato.lNumIntItensContrato = objItemContrato.lNumIntDoc
        objItemNF.objCobrItensContrato.dtDataUltCobranca = dtDataCobranca
               
        'Se não tiver medições
        If objItemContrato.iMedicao <> MARCADO Then
    
            objItemNF.objCobrItensContrato.dtDataRefIni = objItemContrato.dtDataRefIni
            objItemNF.objCobrItensContrato.dtDataRefFim = objItemContrato.dtDataRefFim
            
        Else 'Com medições
    
            Set objItemNF.objCobrItensContrato.colMedicoes = gcolcolMedicoes.Item(iLinha)
        
        End If
        
    End If

    Set objNFiscal.ColItensNF.Item(iLinha).objCobrItensContrato = objItemNF.objCobrItensContrato

    Move_Contrato_Memoria = SUCESSO

    Exit Function

Erro_Move_Contrato_Memoria:

    Move_Contrato_Memoria = gErr

    Select Case gErr
    
        Case 132941, 132943, 136154
        
        Case 132942
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_CADASTRADO_GRID", gErr, sContrato, iLinha)
        
        Case 132944
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_CONTRATO_NAO_CADASTRADO_GRID", gErr, sContrato, iSeq, iLinha)

        Case 136180
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_CLIENTE_NF", gErr, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157285)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Contrato(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objContrato As New ClassContrato
Dim sContratoAnterior As String
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Contrato

    Set objGridInt.objControle = Contrato

    sContratoAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col)

    'Se o contrato estiver preenchido
    If Len(Trim(Contrato.Text)) > 0 Then

        objContrato.sCodigo = Contrato.Text
        objContrato.iFilialEmpresa = giFilialEmpresa
    
        'Le o contrato
        lErro = CF("Contrato_Le", objContrato)
        If lErro <> 129332 And lErro <> SUCESSO Then gError 132948
        
        'Se não estiver cadastrado => Erro
        If lErro = 129332 Then gError 132949
        
        'Se não for um contrato a receber => Erro
        If objContrato.iTipo <> CONTRATOS_RECEBER Then gError 132990
    
    End If
    
    'Se mudou o contrato
    If sContratoAnterior <> Contrato.Text Then
    
        'Limpa o item e as medições
        GridItens.TextMatrix(GridItens.Row, iGrid_ItemContrato_Col) = ""
        
        If GridItens.Row <= gcolcolMedicoes.Count Then
            For iIndice = gcolcolMedicoes.Item(GridItens.Row).Count To 1 Step -1
                gcolcolMedicoes.Item(GridItens.Row).Remove iIndice
            Next
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132950

    Saida_Celula_Contrato = SUCESSO

    Exit Function

Erro_Saida_Celula_Contrato:

    Saida_Celula_Contrato = gErr

    Select Case gErr

        Case 132948, 132850
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 132949
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_CADASTRADO", gErr, objContrato.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 132990
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_RECEBER", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157286)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ItemContrato(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objItemContrato As New ClassItensDeContrato
Dim objContrato As New ClassContrato
Dim iItemAnterior As Integer
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_ItemContrato

    Set objGridInt.objControle = Item
    
    iItemAnterior = StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_ItemContrato_Col))

    'Se o item estiver preenchida
    If Len(Trim(Item.Text)) > 0 Then
            
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col))) = 0 Then gError 136155
            
        'Busca repetições
        For iIndice = 1 To objGridItens.iLinhasExistentes
            If iIndice <> GridItens.Row Then
                If GridItens.TextMatrix(iIndice, iGrid_Contrato_Col) = GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col) And _
                    StrParaInt(GridItens.TextMatrix(iIndice, iGrid_ItemContrato_Col)) = StrParaInt(Item.Text) Then
                    gError 136226
                End If
            End If
        Next
                        
        'Critica o valor
        lErro = Valor_Inteiro_Critica(Item.Text)
        If lErro <> SUCESSO Then gError 132962
        
        objContrato.sCodigo = GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col)
        objContrato.iFilialEmpresa = giFilialEmpresa
    
        'Le o contrato
        lErro = CF("Contrato_Le", objContrato)
        If lErro <> 129332 And lErro <> SUCESSO Then gError 132951
        
        If lErro = 129332 Then gError 132952
        
        objItemContrato.lNumIntContrato = objContrato.lNumIntDoc
        objItemContrato.iSeq = StrParaInt(Item.Text)
    
        'Le o item
        lErro = CF("ItensDeContrato_Le2", objItemContrato)
        If lErro <> 129266 And lErro <> SUCESSO Then gError 132953
        
        If lErro = 129266 Then gError 132954
    
    End If
    
    'Se mudou
    If iItemAnterior <> StrParaInt(Item.Text) Then
           
        'Limpa as medições
        If GridItens.Row <= gcolcolMedicoes.Count Then
            For iIndice = gcolcolMedicoes.Item(GridItens.Row).Count To 1 Step -1
                gcolcolMedicoes.Item(GridItens.Row).Remove iIndice
            Next
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 132955

    Saida_Celula_ItemContrato = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemContrato:

    Saida_Celula_ItemContrato = gErr

    Select Case gErr

        Case 132951, 132853, 132955, 132962
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 132952
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_CADASTRADO", gErr, objContrato.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 132954
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMDECONTRATO_NAO_CADASTRADO", gErr, objItemContrato.iSeq)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 136155
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONTRATO_PREENCHIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 136226
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_CONTRATO_REPETIDO", gErr, GridItens.TextMatrix(iIndice, iGrid_Contrato_Col), GridItens.TextMatrix(iIndice, iGrid_ItemContrato_Col), iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157287)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataCobranca(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objContrato As New ClassContrato
Dim sContratoAnterior As String
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_DataCobranca

    Set objGridInt.objControle = DataCobranca

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 188449

    Saida_Celula_DataCobranca = SUCESSO

    Exit Function

Erro_Saida_Celula_DataCobranca:

    Saida_Celula_DataCobranca = gErr

    Select Case gErr

        Case 188449
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188450)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function


Public Function Traz_Contrato_Tela(ByVal objItemNF As ClassItemNF, iLinha As Integer) As Long
'Move os Itens do Grid para a Memória

Dim lErro As Long
Dim objItemContrato As New ClassItensDeContrato
Dim objContrato As New ClassContrato
Dim objItensDeContratoCob As New ClassItensDeContratoCob
Dim objItensMedCtr As New ClassItensMedCtr
Dim colItensMedCtr As New Collection

On Error GoTo Erro_Traz_Contrato_Tela

    lErro = CF("NFiscal_Le_Contrato", objItemNF)
    If lErro <> SUCESSO And lErro <> 132986 Then gError 132984

    If lErro = SUCESSO Then

        objItemContrato.lNumIntDoc = objItemNF.objCobrItensContrato.lNumIntItensContrato
        
        'Le os itens de contrato
        lErro = CF("ItensDeContrato_Le3", objItemContrato)
        If lErro <> SUCESSO And lErro <> 129266 Then gError 132970
    
        If lErro <> SUCESSO Then gError 132971
        
        'Le o Contrato
        objContrato.lNumIntDoc = objItemContrato.lNumIntContrato
    
        'Le o contrato
        lErro = CF("Contrato_Le2", objContrato)
        If lErro <> 129261 And lErro <> SUCESSO Then gError 132972
        
        If objItemContrato.iMedicao <> MARCADO Then
        
            objItensDeContratoCob.lNumIntItensContrato = objItemNF.objCobrItensContrato.lNumIntItensContrato
            objItensDeContratoCob.lNumIntItemNFiscal = objItemNF.lNumIntDoc
        
            lErro = CF("ItensDeContratoCobranca_Le", objItensDeContratoCob)
            If lErro <> SUCESSO And lErro <> 188444 Then gError 188451
        
            If objItensDeContratoCob.dtDataUltCobranca <> DATA_NULA Then GridItens.TextMatrix(iLinha, iGrid_DataCobranca_Col) = objItensDeContratoCob.dtDataUltCobranca
        
        Else
        
            objItensMedCtr.lNumIntItensContrato = objItemNF.objCobrItensContrato.lNumIntItensContrato
            objItensMedCtr.lNumIntItemNFiscal = objItemNF.lNumIntDoc
            
            lErro = CF("ItensDeMedicaoCobranca_Le", objItensMedCtr, colItensMedCtr)
            If lErro <> SUCESSO Then gError 193844
        
            If colItensMedCtr.Count > 0 Then
        
                Set objItensMedCtr = colItensMedCtr.Item(1)
        
                If objItensMedCtr.dtDataCobranca <> DATA_NULA Then GridItens.TextMatrix(iLinha, iGrid_DataCobranca_Col) = objItensMedCtr.dtDataCobranca
        
            End If
        
        End If
        
        GridItens.TextMatrix(iLinha, iGrid_Contrato_Col) = objContrato.sCodigo
        If objItemContrato.iSeq <> 0 Then GridItens.TextMatrix(iLinha, iGrid_ItemContrato_Col) = objItemContrato.iSeq

        gcolcolMedicoes.Add colItensMedCtr

    Else
    
        gcolcolMedicoes.Add colItensMedCtr
    
    End If

    Traz_Contrato_Tela = SUCESSO

    Exit Function

Erro_Traz_Contrato_Tela:

    Traz_Contrato_Tela = gErr

    Select Case gErr
    
        Case 132970, 132971, 132984, 188451, 193844
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157288)

    End Select

    Exit Function

End Function

Public Sub BotaoDocContrato_Click()

Dim lErro As Long
Dim objContrato As New ClassContrato
Dim sContrato As String

On Error GoTo Erro_BotaoDocContrato_Click

    If GridItens.Row = 0 Then gError 136006

    sContrato = GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col)

    If Len(Trim(sContrato)) = 0 Then gError 136007
    
    objContrato.sCodigo = sContrato
    objContrato.iFilialEmpresa = giFilialEmpresa

    Call Chama_Tela("ContratoCadastro", objContrato)

    Exit Sub
    
Erro_BotaoDocContrato_Click:
    
    Select Case gErr

        Case 136006
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 136007
             Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONTRATO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157289)

    End Select
    
    Exit Sub

End Sub

Public Function Valida_Contrato_Itens(ByVal objItemNF As ClassItemNF, iLinha As Integer) As Long

Dim lErro As Long
Dim objItemContrato As New ClassItensDeContrato
Dim objItemMedicao As New ClassItensMedCtr
Dim vbMsgRes As VbMsgBoxResult
Dim sAviso As String

On Error GoTo Erro_Valida_Contrato_Itens
    
    'Se está associado a um contrato
    If objItemNF.objCobrItensContrato.lNumIntItensContrato <> 0 Then
    
        objItemContrato.lNumIntDoc = objItemNF.objCobrItensContrato.lNumIntItensContrato
        
        'Le o item de Contrato
        lErro = CF("ItensDeContrato_Le3", objItemContrato)
        If lErro <> SUCESSO And lErro <> 129266 Then gError 136017
    
        'Se é um item de medição
        If objItemNF.objCobrItensContrato.colMedicoes.Count <> 0 Then
    
            lErro = Obtem_ItemMedicao(objItemMedicao, iLinha)
            If lErro <> SUCESSO Then gError 136018
            
            If gcolcolMedicoes.Item(iLinha).Count = 0 Then gError 136201
                              
            'Se tem diferença entre as quantidades do ItemNF para o ItemMedicao => Avisa
            If Abs(objItemNF.dQuantidade - objItemMedicao.dQuantidade) > QTDE_ESTOQUE_DELTA Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_QTD_MEDICAO_DIFERENTE", iLinha, Formata_Estoque(objItemMedicao.dQuantidade), Formata_Estoque(objItemNF.dQuantidade))
                If vbMsgRes = vbNo Then gError 136019
            End If
            
            'Se tem diferença entre os preços unitários do ItemNF para o ItemMedicao => Avisa
            If Abs(objItemNF.dPrecoUnitario - (objItemMedicao.dVlrCobrar / objItemMedicao.dQuantidade)) > DELTA_VALORMONETARIO Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PRECOUNIT_MEDICAO_DIFERENTE", iLinha, Format((objItemMedicao.dVlrCobrar / objItemMedicao.dQuantidade), "STANDARD"), Format(objItemNF.dPrecoUnitario, "STANDARD"))
                If vbMsgRes = vbNo Then gError 136020
            End If
            
        Else
        
            If objItemContrato.iMedicao = MARCADO Then gError 136153
        
            'Se tem diferença entre as quantidades do ItemNF para o ItemContrato => Avisa
            If Abs(objItemNF.dQuantidade - objItemContrato.dQuantidade) > QTDE_ESTOQUE_DELTA Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_QTD_CONTRATO_DIFERENTE", iLinha, Formata_Estoque(objItemContrato.dQuantidade), Formata_Estoque(objItemNF.dQuantidade))
                If vbMsgRes = vbNo Then gError 136021
            End If
            
            'Se tem diferença entre os preços unitários do ItemNF para o ItemContrato => Avisa
            If Abs(objItemNF.dPrecoUnitario - objItemContrato.dValor) > DELTA_VALORMONETARIO Then
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PRECOUNIT_CONTRATO_DIFERENTE", iLinha, Format(objItemContrato.dValor, "STANDARD"), Format(objItemNF.dPrecoUnitario, "STANDARD"))
                If vbMsgRes = vbNo Then gError 136022
            End If
            
        End If
        
        lErro = Testa_DataCobranca_ItemContrato(iLinha, objItemNF.objCobrItensContrato.dtDataUltCobranca, objItemContrato, sAviso)
        If lErro <> SUCESSO Then gError 188452
                    
        If Len(sAviso) > 0 Then gError 188453
        
    End If

    Valida_Contrato_Itens = SUCESSO

    Exit Function

Erro_Valida_Contrato_Itens:

    Valida_Contrato_Itens = gErr

    Select Case gErr
    
        Case 136017 To 136022, 188452

        Case 136201
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_MEDICAO_NAO_PREENCHIDO", gErr, iLinha)
        
        Case 136176
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_MEDICAO_NAO_CADASTRADO_GRID", gErr, objItemMedicao.lMedicao, objItemMedicao.iItem, iLinha)
        
        Case 136153
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_MEDICAO_SEM_MEDICAO", gErr, iLinha)
        
        Case 188453
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_DATA_COBRANCA", gErr, sAviso)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157290)

    End Select

    Exit Function

End Function

Private Function Obtem_ItemMedicao(ByVal objItemMedicao As ClassItensMedCtr, ByVal iLinha As Integer) As Long

Dim objItemAux As ClassItensMedCtr
Dim dValor As Double
Dim dQuantidade As Double

    For Each objItemAux In gcolcolMedicoes.Item(iLinha)
    
        dValor = dValor + objItemAux.dVlrCobrar
        dQuantidade = dQuantidade + objItemAux.dQuantidade
    
    Next
    
    objItemMedicao.dQuantidade = dQuantidade
    objItemMedicao.dVlrCobrar = dValor
    
End Function

Public Sub BotaoMedicao_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objItemContrato As New ClassItensDeContrato
Dim sContrato As String
Dim iSeq As Integer
Dim objContrato As New ClassContrato
Dim objItemMedicao As New ClassItensMedCtr

On Error GoTo Erro_BotaoMedicao_Click

    If GridItens.Row = 0 Then gError 132937

    sContrato = GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col)
    iSeq = StrParaInt(GridItens.TextMatrix(GridItens.Row, iGrid_ItemContrato_Col))

    If Len(Trim(sContrato)) = 0 Then gError 132938
    If iSeq = 0 Then gError 132939
    
    objContrato.sCodigo = sContrato
    objContrato.iFilialEmpresa = giFilialEmpresa

    lErro = CF("Contrato_Le", objContrato)
    If lErro <> 129332 And lErro <> SUCESSO Then gError 136207
    
    If lErro = 129332 Then gError 136208
    
    If objContrato.iTipo <> CONTRATOS_RECEBER Then gError 136209
    
    objItemContrato.lNumIntContrato = objContrato.lNumIntDoc
    objItemContrato.iSeq = iSeq

    lErro = CF("ItensDeContrato_Le2", objItemContrato)
    If lErro <> 129266 And lErro <> SUCESSO Then gError 136210
    
    If lErro = 129266 Then gError 136211
    
    If objItemContrato.iMedicao = 0 Then gError 136214
    
    lErro = Chama_Tela_Modal("MedicoesNFiscal", gcolcolMedicoes.Item(GridItens.Row), objItemContrato)
    If giRetornoTela = vbCancel Then gError 136206
    If giRetornoTela = vbOK Then
    
        'Agrupa o Valor e quantidade das medições
        Call Obtem_ItemMedicao(objItemMedicao, GridItens.Row)
        
        If objItemMedicao.dQuantidade > 0 Then
            GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(objItemMedicao.dVlrCobrar / objItemMedicao.dQuantidade, "STANDARD")
            GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(objItemMedicao.dQuantidade)
        End If
        
        If Frame1(TAB_Almoxarifado).Enabled = True Then
            lErro = Alocacao_Processa()
            If lErro <> SUCESSO Then gError 136290
        End If
        
        Call PrecoTotal_Calcula(GridItens.Row, False)
        
        '*** incluidos p/tratamento de tributacao ****************
        If GridItens.Row <= objGridItens.iLinhasExistentes Then
            lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
            If lErro <> SUCESSO Then gError 136291
        End If
        '*** fim tributacao
        
        Call ValorTotal_Calcula
    
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 136292
        
    End If
            
    Exit Sub
    
Erro_BotaoMedicao_Click:
    
    Select Case gErr

        Case 132937
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 132938
             Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONTRATO_PREENCHIDO", gErr)
        
        Case 132939
             Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_CODIGO_CONTRATO_PREENCHIDO", gErr)
             
        Case 136206, 136207, 136210, 136290 To 136292
        
        Case 136208
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_CADASTRADO", gErr, objContrato.sCodigo)
        
        Case 136209
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_RECEBER", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)
        
        Case 136211
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_CONTRATO_NAO_CADASTRADO_GRID", gErr, sContrato, iSeq, GridItens.Row)
        
        Case 136214
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_SEM_MEDICAO", gErr, iSeq)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157291)

    End Select
    
    Exit Sub

End Sub

Public Sub BotaoExibirDados_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objTipoDocInfoOrig As New ClassTipoDocInfo
Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoExibirDados_Click

    'Verifica se a Serie e o Número da Nota Fiscal original estão preenchidos
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Or Len(Trim(NFiscalOriginal.ClipText)) = 0 Then gError 126514

    If Len(Trim(Cliente.ClipText)) = 0 Then gError 42351
    If Len(Trim(Filial.Text)) = 0 Then gError 42352

    objCliente.sNomeReduzido = Cliente.Text
    'Lê o Cliente
    lErro = CF("Cliente_Le_Codigo_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 42411
    
    'Se não encontrou o Cliente --> erro
    If lErro = 12348 Then gError 42412

    objNFiscal.lCliente = objCliente.lCodigo
    
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
    objNFiscal.sSerie = SerieNFiscalOriginal.Text
    objNFiscal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("NFiscalFatura_Le_NFOriginal", objNFiscal)
    If lErro <> SUCESSO And lErro <> 45347 Then gError 42920
    'Se não encontrar --> Erro
    If lErro = 45347 Then gError 42379
    
    'Tenta ler a nota Fiscal passada como parâmetro
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 42403
    
    'Se não achou a NFiscal --> erro
    If lErro <> SUCESSO Then gError 42406

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 42358
    
    'Não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 42359

    If objTipoDocInfo.iComplementar = DOCINFO_NORMAL Then
    
        'Coloca na tela os dados encontados
        lErro = Traz_NFOrig_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 126517

    End If
    
    Exit Sub

Erro_BotaoExibirDados_Click:

    Select Case gErr

        Case 126514
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NUMERO_ORIGINAL_FALTANDO2", gErr)

        Case 126515, 126517, 42411, 42403, 42358

        Case 42359
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iTipo)

        Case 42406
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case 42412
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case 42351
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 42352
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 42379
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_ASSOCIADA_CLIENTE", gErr, Cliente.Text, SerieNFiscalOriginal.Text, NFiscalOriginal.Text)

        Case 126516
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_REMESSA_VENDA_FORA_ESTAB", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157292)

    End Select

    Exit Sub

End Sub

Public Function Traz_NFOrig_Tela(objNFiscal As ClassNFiscal) As Long
'Traz os dados da Nota Fiscal Original passada em objNFiscal

Dim lErro As Long, objItemNF As ClassItemNF
Dim objNFiscalOriginal As New ClassNFiscal
Dim bCancel As Boolean, iLinha As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim colTipoDocInfo As New colTipoDocInfo
Dim objTipoDocInfo1 As New ClassTipoDocInfo, iTipoNFiscal As Integer
Dim iIndice As Integer, objTituloRec As New ClassTituloReceber
Dim dPercDesc As Double

On Error GoTo Erro_Traz_NFOrig_Tela

    gbCarregandoTela = True

    'Limpa a tela sem fechar as setas
    Call Limpa_Tela_NFiscal2

    SerieNFiscalOriginal.Text = objNFiscal.sSerie
    NFiscalOriginal.Text = objNFiscal.lNumNotaFiscal

    'Lê os Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 35705

    'Lê a parte de Tributação
    lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
    If lErro <> SUCESSO Then gError 39342

    'Lê a parte de tributacao dos Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
    If lErro <> SUCESSO Then gError 39343
    
    'Lê as Comissões da Nota Fiscal
    lErro = CF("ComissoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 35703

    'Lê as Alocações dos Ítens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 35704
    
    If objNFiscal.lNumIntDocCPR <> 0 Then
        
        lErro = CF("ParcelasRecNF_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 35704
    
        'carrega dados do titulo a receber associado à nota fiscal para obter a condicao de pagto
        objTituloRec.lNumIntDoc = objNFiscal.lNumIntDocCPR
        lErro = CF("TituloReceber_Le", objTituloRec, 1)
        If lErro <> SUCESSO And lErro <> 26061 Then gError 35704
        If lErro <> SUCESSO Then
        
            lErro = CF("TituloReceberBaixado_Le", objTituloRec, 1)
            If lErro <> SUCESSO And lErro <> 56570 Then gError 35704
            If lErro <> SUCESSO Then gError 35704
        
        End If
        
        objNFiscal.iCondicaoPagto = objTituloRec.iCondicaoPagto
        
    End If
        
    Cliente.Text = objNFiscal.lCliente
    Call Cliente_Validate(bCancel)
    Filial.Text = objNFiscal.iFilialCli
    Call Filial_Validate(bSGECancelDummy)
    
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

    Select Case objNFiscal.iTipoNFiscal
    
        Case DOCINFO_NFISRVFE
            TipoNFiscal.Text = CStr(DOCINFO_NFISFVFE)
            objNFiscal.iTipoNFiscal = DOCINFO_NFISFVFE
            objNFiscal.objTributacaoNF.sNaturezaOpInterna = "5103"
            objNFiscal.objTributacaoNF.iTipoTributacao = 15
    
        Case DOCINFO_NFISC
            TipoNFiscal.Text = CStr(DOCINFO_NFFISPC)
            objNFiscal.iTipoNFiscal = DOCINFO_NFFISPC
'            objNFiscal.objTributacaoNF.sNaturezaOpInterna = "5113"
'            objNFiscal.objTributacaoNF.iTipoTributacao = 97
        
    End Select
    
    Call TipoNFiscal_Validate(bSGECancelDummy)
    
    'Preenche a Condicao de Pagto
    If objNFiscal.iCondicaoPagto > 0 Then
        PercAcrescFin.Text = ""
        CondicaoPagamento.Text = objNFiscal.iCondicaoPagto
        Call CondicaoPagamento_Validate(bSGECancelDummy)
    End If

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

    
    VolumeNumero.Text = objNFiscal.sVolumeNumero
    
    Call DateParaMasked(DataReferencia, gdtDataHoje)
    giDataReferenciaAlterada = 0
    Call DateParaMasked(DataEmissao, gdtDataHoje)
    Call DateParaMasked(DataSaida, gdtDataHoje)

    'Coloca os Valores na Tela
    ValorDesconto.Text = Format(objNFiscal.dValorDesconto, "Standard")
    ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Standard")
    ValorFrete.Text = Format(objNFiscal.dValorFrete, "Standard")
    ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Standard")
    ValorTotal.Caption = Format(objNFiscal.dValorTotal, "Standard")
    ValorProdutos.Caption = Format(objNFiscal.dValorProdutos, "Standard")

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
    
    If objNFiscal.lVolumeQuant > 0 Then VolumeQuant.Text = objNFiscal.lVolumeQuant
    NumPedidoTerc.Text = objNFiscal.sNumPedidoTerc

    Mensagem.Text = objNFiscal.sMensagemNota
    MensagemCorpo.Text = objNFiscal.sMensagemCorpoNota

    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = Format(objNFiscal.dPesoLiq, "Fixed")
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = Format(objNFiscal.dPesoBruto, "Fixed")

    'Preenche o Grid com os Ítens da Nota Fiscal
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 39023
    
    'Carrega o Tab Comissões
    lErro = objTabComissoes.Carrega_Tab_Comissoes(objNFiscal)
    If lErro <> SUCESSO Then gError 39022
    
    'Preenche o Grid de Parcelas
    lErro = Preenche_Grid_Parcelas(objNFiscal)
    If lErro <> SUCESSO Then gError 42467

    Set gobjNFiscal = objNFiscal
    
    iLinha = 0
    For Each objItemNF In objNFiscal.ColItensNF
        
        iLinha = iLinha + 1
        lErro = gobjTribTab.Inclusao_Item_Grid(iLinha, objItemNF.sProduto)
        If lErro <> SUCESSO Then gError 27832
        
    Next
    
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
    
    lErro = AlocaAlmoxarifadoPradrao1(objNFiscal)
    If lErro <> SUCESSO Then gError 133407
    
    'Faz a atualização dos valores da tributação
    lErro = gobjTribTab.AtualizarTributacao()
    If lErro <> SUCESSO Then gError 133407
    
    'Traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 39673
    
    lErro = ValorTotal_Calcula
    If lErro <> SUCESSO Then gError 133407

    iAlterado = 0
    iComissoesAlterada = 0
     
    gbCarregandoTela = False
     
    Traz_NFOrig_Tela = SUCESSO

    Exit Function

Erro_Traz_NFOrig_Tela:

    gbCarregandoTela = False

    Traz_NFOrig_Tela = gErr

    Select Case gErr

        Case 25821, 25822, 39031, 35705, 39023, 39024, 39022, 39342, 39343, 39673, 42460, 42467, 56571, 56572, 83384, 83385, 101305, 133407

        Case 56573
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_REC_INEXISTENTE", gErr)
        
        Case 39025
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntNotaOriginal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157293)

    End Select

    Exit Function

End Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157294)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157295)
     
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

Public Sub Botao_Click(iIndice As Integer)

On Error GoTo Erro_Botao_Click
    
    Select Case iIndice
    
        Case BOTAO_ORCAMENTO
            Call BotaoOrcamento_Click
        
        Case BOTAO_VERORCAMENTO
            Call BotaoVerOrcamento_Click
        
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
        
        Case BOTAO_EXIBIRDADOS
            Call BotaoExibirDados_Click
        
        Case BOTAO_GRADE
            Call BotaoGrade_Click
        
        Case BOTAO_KITVENDA
            Call BotaoKitVenda_Click
        
        Case BOTAO_PRODUTOS
            Call BotaoProdutos_Click
        
        Case BOTAO_ITEMCONTRATO
            Call BotaoItemContrato_Click
        
        Case BOTAO_MEDICAO
            Call BotaoMedicao_Click
        
        Case BOTAO_DOCCONTRATO
            Call BotaoDocContrato_Click
        
        Case BOTAO_ESTOQUEPROD
            Call BotaoEstoqueProd_Click
        
'        Case BOTAO_VENDEDORES
'            Call BotaoVendedores_Click
        
        Case BOTAO_LOCALIZACAO
            Call BotaoLocalizacao_Click
        
        Case BOTAO_SERIE
            Call BotaoSerie_Click
        
        Case BOTAO_LOTES
            Call BotaoLotes_Click
            
        Case BOTAO_PROJETOS
            Call BotaoProjetos_Click

        Case BOTAO_LIMPARNF
            Call BotaoLimparNF_Click

        Case BOTAO_CONTRATO
            Call BotaoContrato_Click
        
        Case BOTAO_VERCONTRATO
            Call BotaoVerContrato_Click
        

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

Public Function NFF_Form_Load(ByVal objCT As Object) As Long
    NFF_Form_Load = SUCESSO
End Function

Public Function NFF_Form_Load2(ByVal objCT As Object) As Long
    NFF_Form_Load2 = SUCESSO
End Function

Public Function NFF_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long
    NFF_Form_Unload = SUCESSO
End Function

Public Function NFF_Grava_Registro(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long
    NFF_Grava_Registro = SUCESSO
End Function

Public Function NFF_Limpa_Tela(ByVal objCT As Object) As Long
    NFF_Limpa_Tela = SUCESSO
End Function

Public Function NFF_Move_Tela_Memoria(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long
    NFF_Move_Tela_Memoria = SUCESSO
End Function

Public Function NFF_Traz_NFiscal_Tela(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal, ByVal objTituloRec As ClassTituloReceber) As Long
    NFF_Traz_NFiscal_Tela = SUCESSO
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

Function Testa_DataCobranca_ItemContrato(ByVal iItemGrid As Integer, ByVal dtDataCobrancaTela As Date, ByVal objItemContrato As ClassItensDeContrato, sAviso As String) As Long

Dim lErro As Long
Dim dtDataProxCobranca As Date
Dim dtRefFim As Date
Dim objContrato As New ClassContrato
Dim dtDataUltCobranca As Date

On Error GoTo Erro_Testa_DataCobranca_ItemContrato

    objContrato.lNumIntDoc = objItemContrato.lNumIntContrato

    'Le o Contrato pelo número interno
    lErro = CF("Contrato_Le2", objContrato)
    If lErro <> SUCESSO And lErro <> 129261 Then gError 188454

    If lErro <> SUCESSO Then gError 188455

    dtDataProxCobranca = objItemContrato.dtDataProxCobranca
    dtDataUltCobranca = DATA_NULA

    Do While dtDataProxCobranca < dtDataCobrancaTela And (dtDataProxCobranca < objContrato.dtDataFimContrato Or objContrato.dtDataFimContrato = DATA_NULA)

        dtDataUltCobranca = dtDataProxCobranca

        'retorna a data da proxima cobranca do item do contrato
        lErro = CF("Contrato_ProxCobranca", objItemContrato.iPeriodicidade, dtDataProxCobranca, objItemContrato.dtDataProxCobranca, dtDataProxCobranca, dtRefFim)
        If lErro <> SUCESSO Then gError 188456

    Loop

    If dtDataProxCobranca > dtDataCobrancaTela Then
        If dtDataProxCobranca > objContrato.dtDataFimContrato And dtDataUltCobranca = DATA_NULA And objContrato.dtDataFimContrato <> DATA_NULA Then
            sAviso = sAviso & "O contrato da linha " & CStr(iItemGrid) & " do grid já está encerrado" & Chr(10)
        ElseIf dtDataProxCobranca > objContrato.dtDataFimContrato And objContrato.dtDataFimContrato <> DATA_NULA Then
            sAviso = sAviso & "A data mais próxima válida da informada na linha " & CStr(iItemGrid) & " do grid é " & CStr(dtDataUltCobranca) & Chr(10)
        ElseIf dtDataUltCobranca = DATA_NULA Then
            sAviso = sAviso & "A data mais próxima válida da informada na linha " & CStr(iItemGrid) & " do grid é " & CStr(dtDataProxCobranca) & Chr(10)
        Else
            sAviso = sAviso & "As datas mais próximas válidas da informada na linha " & CStr(iItemGrid) & " do grid são " & CStr(dtDataUltCobranca) & " e " & CStr(dtDataProxCobranca) & Chr(10)
        End If
    ElseIf dtDataProxCobranca > objContrato.dtDataFimContrato And objContrato.dtDataFimContrato <> DATA_NULA Then
            If dtDataUltCobranca <> DATA_NULA Then
                sAviso = sAviso & "A data mais próxima válida da informada na linha " & CStr(iItemGrid) & " do grid é " & CStr(dtDataUltCobranca) & Chr(10)
            Else
                sAviso = sAviso & "O contrato da linha " & CStr(iItemGrid) & " do grid já está encerrado" & Chr(10)
            End If
    End If
                
    If Len(sAviso) > 0 Then sAviso = Chr(10) & sAviso
                
    Testa_DataCobranca_ItemContrato = SUCESSO

    Exit Function

Erro_Testa_DataCobranca_ItemContrato:

    Testa_DataCobranca_ItemContrato = gErr

    Select Case gErr
    
        Case 188454, 188456

        Case 184455
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_CADASTRADO1", gErr, objItemContrato.lNumIntContrato)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188457)

    End Select

    Exit Function

End Function

Public Function BotaoContratoOrcamento_Click() As Long
    
    If FrameOrcVenda.Visible = True Then
        FrameOrcVenda.Visible = False
        Frame1(303).Visible = True
        BotaoContratoOrcamento.Caption = "Orçamento"
    Else
        FrameOrcVenda.Visible = True
        Frame1(303).Visible = False
        BotaoContratoOrcamento.Caption = "Contrato"
    End If
    
End Function

Private Function NFiscalContrato_Gera2(objGeracaoFatContrato As ClassGeracaoFatContrato, objNFiscal As ClassNFiscal) As Long
'Gera as Notas Fiscais em cima dos itens de contratos/medição que vão ser faturados

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim objContabil As ClassContabil
Dim objContratoFat As ClassContratoFat
Dim objContratoFatItens As ClassContratoFatItens
Dim dValorTotal As Double, objTribTab As ClassTribTab
Dim objTipoDocInfo As New ClassTipoDocInfo, dValorLiquido As Double
Dim iItem As Integer
Dim objItemMedicao As ClassItensMedCtr
Dim colContFatItensAgrupado As New Collection
Dim objContFatItensAux1 As ClassContratoFatItens
Dim objContFatItensAux2 As ClassContratoFatItens
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim dQuantidade As Double
Dim dCusto As Double
Dim dValor As Double
Dim bAchou As Boolean
Dim colMedicoes As Collection
Dim colcolMedicoes As New Collection

On Error GoTo Erro_NFiscalContrato_Gera2

    Set gcolcolMedicoes = New Collection

    objGeracaoFatContrato.iFilialEmpresa = giFilialEmpresa
    objGeracaoFatContrato.iTipoNFiscal = DOCINFO_NFISFV

    'Lê o Tipo de Documento
    objTipoDocInfo.iCodigo = objGeracaoFatContrato.iTipoNFiscal
    
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 188547

    'se não estiver cadastrado ==> erro
    If lErro = 31415 Then gError 188548
    
    lErro = CF("NFiscalContrato_Le1", objGeracaoFatContrato, CONTRATOS_RECEBER)
    If lErro <> AD_SQL_SUCESSO Then gError 188549

    If objGeracaoFatContrato.colItens.Count > 0 Then

        Set objContratoFat = objGeracaoFatContrato.colItens.Item(1)
                      
        Set colcolMedicoes = New Collection
        Set colContFatItensAgrupado = New Collection
                      
        Set objNFiscal = New ClassNFiscal
                
        Call objNFiscal.Inicializa_Tributacao
        Call objNFiscal.objTributacaoNF.Coloca_Auto
        
        objNFiscal.iCondicaoPagto = objContratoFat.iCondPagto
        objNFiscal.iFilialCli = objContratoFat.iFilCli
        objNFiscal.lCliente = objContratoFat.lCliente
        objNFiscal.dtDataReferencia = objContratoFat.dtDataRef
        objNFiscal.dtDataEmissao = gdtDataAtual
        objNFiscal.dtDataSaida = gdtDataAtual
        objNFiscal.iFilialEmpresa = objGeracaoFatContrato.iFilialEmpresa
        objNFiscal.sSerie = objContratoFat.sSerie
        objNFiscal.sNaturezaOp = objContratoFat.sNaturezaOp
        objNFiscal.iTipoDocInfo = objGeracaoFatContrato.iTipoNFiscal
        objNFiscal.iTipoNFiscal = objGeracaoFatContrato.iTipoNFiscal
        objNFiscal.iStatus = STATUS_LANCADO
        
        If Len(Trim(objContratoFat.sNaturezaOp)) <> 0 Then
            objNFiscal.sNaturezaOp = objContratoFat.sNaturezaOp
            objNFiscal.objTributacaoNF.sNaturezaOpInterna = objContratoFat.sNaturezaOp
            objNFiscal.objTributacaoNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
        End If
        
        If objContratoFat.iTipoTributacao <> 0 Then
            objNFiscal.objTributacaoNF.iTipoTributacao = objContratoFat.iTipoTributacao
            objNFiscal.objTributacaoNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
        End If
           
        iItem = 0
        iIndice1 = 0
        
        'Agrupa itens
        For Each objContFatItensAux1 In objContratoFat.colItens
        
            iIndice1 = iIndice1 + 1
            iIndice2 = 0
            dQuantidade = 0
            dValor = 0
            dCusto = 0
            
            bAchou = False
            
            Set colMedicoes = New Collection

            For Each objContFatItensAux2 In colContFatItensAgrupado
                If objContFatItensAux1.sProduto = objContFatItensAux2.sProduto Then
                    bAchou = True
                    Exit For
                End If
            Next
            
            If Not bAchou Then
            
                'Soma os itens iguais
                For Each objContFatItensAux2 In objContratoFat.colItens
                    
                    iIndice2 = iIndice2 + 1
                
                    If iIndice1 <> iIndice2 Then
                        If objContFatItensAux1.sProduto = objContFatItensAux2.sProduto Then
                            dQuantidade = dQuantidade + objContFatItensAux2.dQuantidade
                            dValor = dValor + objContFatItensAux2.dVlrCobrar
                            dCusto = dCusto + objContFatItensAux2.dCusto
                            
                            If objContFatItensAux2.lMedicao <> 0 Then
                            
                                Set objItemMedicao = New ClassItensMedCtr
                            
                                objItemMedicao.lNumIntItensContrato = objContFatItensAux2.lNumIntItensContrato
                                objItemMedicao.iItem = objContFatItensAux2.iItem
                                objItemMedicao.lMedicao = objContFatItensAux2.lMedicao
                                objItemMedicao.dtDataCobranca = objContFatItensAux2.dtDataProxCobranca
                                objItemMedicao.dtDataRefIni = objContFatItensAux2.dtDataRefIni
                                objItemMedicao.dtDataRefFim = objContFatItensAux2.dtDataRefFim
                                
                                'Le a medição
                                lErro = CF("ItensDeMedicaoContrato_Le2", objItemMedicao)
                                If lErro <> SUCESSO And lErro <> 136173 Then gError 193850
                                
                                colMedicoes.Add objItemMedicao
                            
                            End If
                            
                        End If
                    End If
                
                Next
                
                If objContFatItensAux1.lMedicao <> 0 Then
                
                    Set objItemMedicao = New ClassItensMedCtr
                
                    objItemMedicao.lNumIntItensContrato = objContFatItensAux1.lNumIntItensContrato
                    objItemMedicao.iItem = objContFatItensAux1.iItem
                    objItemMedicao.lMedicao = objContFatItensAux1.lMedicao
                    objItemMedicao.dtDataCobranca = objContFatItensAux1.dtDataProxCobranca
                    objItemMedicao.dtDataRefIni = objContFatItensAux1.dtDataRefIni
                    objItemMedicao.dtDataRefFim = objContFatItensAux1.dtDataRefFim
                    
                    'Le a medição
                    lErro = CF("ItensDeMedicaoContrato_Le2", objItemMedicao)
                    If lErro <> SUCESSO And lErro <> 136173 Then gError 193851
                    
                    colMedicoes.Add objItemMedicao
                
                End If
                
                objContFatItensAux1.dCusto = objContFatItensAux1.dCusto + dCusto
                objContFatItensAux1.dQuantidade = objContFatItensAux1.dQuantidade + dQuantidade
                objContFatItensAux1.dVlrCobrar = objContFatItensAux1.dVlrCobrar + dValor
                
                colContFatItensAgrupado.Add objContFatItensAux1
                colcolMedicoes.Add colMedicoes
                gcolcolMedicoes.Add colMedicoes
                
            End If
        
        Next
        
        For Each objContratoFatItens In colContFatItensAgrupado
                  
            iItem = iItem + 1
        
            Set objItemNF = New ClassItemNF
            Call objItemNF.Inicializa_Tributacao
            Call objItemNF.objTributacaoItemNF.Coloca_Auto
            
            objItemNF.dCusto = objContratoFatItens.dCusto
            objItemNF.dQuantidade = objContratoFatItens.dQuantidade
            objItemNF.dPrecoUnitario = objContratoFatItens.dVlrCobrar / objContratoFatItens.dQuantidade
            objItemNF.dValorTotal = objContratoFatItens.dVlrCobrar
            objItemNF.sCcl = objContratoFatItens.sCcl
            objItemNF.sDescricaoItem = objContratoFatItens.sDescProd
            objItemNF.sProduto = objContratoFatItens.sProduto
            objItemNF.iItem = iItem
            objItemNF.sUnidadeMed = objContratoFatItens.sUM
            
            objItemNF.objItensDeContrato.sContrato = objGeracaoFatContrato.sContratoIni
            objItemNF.objItensDeContrato.iSeq = objContratoFatItens.iItem
            objItemNF.objItensDeContrato.dtDataProxCobranca = objContratoFatItens.dtDataProxCobranca
        
            objItemNF.objCobrItensContrato.lNumIntItensContrato = objContratoFatItens.lNumIntItensContrato
            objItemNF.objCobrItensContrato.dtDataUltCobranca = objContratoFatItens.dtDataProxCobranca
            objItemNF.objCobrItensContrato.lNumIntDocCobranca = objGeracaoFatContrato.lNumIntDoc
            objItemNF.objCobrItensContrato.dtDataRefIni = objContratoFatItens.dtDataRefIni
            objItemNF.objCobrItensContrato.dtDataRefFim = objContratoFatItens.dtDataRefFim
            
            Set objItemNF.objCobrItensContrato.colMedicoes = colcolMedicoes.Item(iItem)
        
            dValorTotal = dValorTotal + objItemNF.dValorTotal
        
            objNFiscal.ColItensNF.Add1 objItemNF
        
            If Len(Trim(objContratoFatItens.sNaturezaOp)) <> 0 Then
                objItemNF.objTributacaoItemNF.sNaturezaOp = objContratoFatItens.sNaturezaOp
                objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
            Else
                If Len(Trim(objContratoFat.sNaturezaOp)) <> 0 Then
                    objItemNF.objTributacaoItemNF.sNaturezaOp = objContratoFat.sNaturezaOp
                    objItemNF.objTributacaoItemNF.iNaturezaOpManual = VAR_PREENCH_MANUAL
                End If
            End If
            
            If objContratoFatItens.iTipoTributacao <> 0 Then
                objItemNF.objTributacaoItemNF.iTipoTributacao = objContratoFatItens.iTipoTributacao
                objItemNF.objTributacaoItemNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
            Else
                If objContratoFat.iTipoTributacao <> 0 Then
                    objItemNF.objTributacaoItemNF.iTipoTributacao = objContratoFat.iTipoTributacao
                    objItemNF.objTributacaoItemNF.iTipoTributacaoManual = VAR_PREENCH_MANUAL
                End If
            End If
            
        Next
        
        objNFiscal.dValorProdutos = dValorTotal
        
        Set objTribTab = New ClassTribTab
        
        lErro = objTribTab.Ativar(Me)
        If lErro <> SUCESSO Then gError 188550
        
        lErro = objTribTab.TributacaoNF_Reset(objNFiscal)
        If lErro <> SUCESSO Then gError 188550
        
        'obtem dados do contrato
        Call objTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
        Call objTribTab.DataEmissao_Alterada(objNFiscal.dtDataEmissao)
        Call objTribTab.Serie_Alterada(objNFiscal.sSerie)
        
        'forçado
        objTribTab.objNFTributacao.iISSInclusoManual = VAR_PREENCH_AUTOMATICO
        
        lErro = gobjTributacao.AtualizaImpostos(objTribTab.mvarobjNFTributacao, 0)
        If lErro <> SUCESSO Then gError 188551
        
        lErro = objTribTab.PreencherNF(objNFiscal)
        If lErro <> SUCESSO Then gError 188552
        
        Set objTribTab = Nothing
        
        objNFiscal.sNaturezaOp = objNFiscal.objTributacaoNF.sNaturezaOpInterna
        
        'acertar valor total e da parcela, considerando os tributos, retencóes,...
        objNFiscal.dValorTotal = dValorTotal + objNFiscal.objTributacaoNF.dICMSSubstValor + objNFiscal.objTributacaoNF.dIPIValor + objNFiscal.dValorFrete + objNFiscal.dValorSeguro + objNFiscal.dValorOutrasDespesas + IIf(objNFiscal.objTributacaoNF.iISSIncluso = 0, objNFiscal.objTributacaoNF.dISSValor, 0)
        'Adiciona a parcela na coleção de parcelas da Nota Fiscal
        dValorLiquido = objNFiscal.dValorTotal - (objNFiscal.objTributacaoNF.dCOFINSRetido + objNFiscal.objTributacaoNF.dCSLLRetido + objNFiscal.objTributacaoNF.dIRRFValor + objNFiscal.objTributacaoNF.dPISRetido + objNFiscal.objTributacaoNF.dISSRetido)
         
        lErro = NFiscalContrato2_Gera_Parcelas(objNFiscal)
        If lErro <> SUCESSO Then gError 188553
        
        lErro = NFiscalContrato2_Gera_Peso(objNFiscal)
        If lErro <> SUCESSO Then gError 188554
            
    End If

    NFiscalContrato_Gera2 = SUCESSO

    Exit Function

Erro_NFiscalContrato_Gera2:

    NFiscalContrato_Gera2 = gErr

    Select Case gErr

        Case 188547, 188549 To 188554, 193850, 193851
        
        Case 188548
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188555)

    End Select

    Exit Function

End Function

Private Function NFiscalContrato2_Gera_Parcelas(objNFiscal As ClassNFiscal)

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim dValorLiquido As Double
Dim iIndice As Integer
Dim dtDataVenctoReal As Date

On Error GoTo Erro_NFiscalContrato2_Gera_Parcelas

    dValorLiquido = objNFiscal.dValorTotal - (objNFiscal.objTributacaoNF.dCOFINSRetido + objNFiscal.objTributacaoNF.dCSLLRetido + objNFiscal.objTributacaoNF.dIRRFValor + objNFiscal.objTributacaoNF.dPISRetido + objNFiscal.objTributacaoNF.dISSRetido)

    objCondicaoPagto.iCodigo = objNFiscal.iCondicaoPagto
    
    'Lê a condição de pagamento
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 188556
    
    'Calcula os valores das Parcelas
    objCondicaoPagto.dValorTotal = dValorLiquido
    objCondicaoPagto.dtDataRef = objNFiscal.dtDataReferencia
    
    'Calcula os valores das Parcelas
    lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, True)
    If lErro <> SUCESSO Then gError 188557

    'Coloca os valores das Parcelas no Grid Parcelas
    For iIndice = 1 To objCondicaoPagto.colParcelas.Count
    
            'Calcula Data Vencimento Real
            lErro = CF("DataVencto_Real", objCondicaoPagto.colParcelas(iIndice).dtVencimento, dtDataVenctoReal)
            If lErro <> SUCESSO Then gError 188558
    
            objNFiscal.ColParcelaReceber.Add 0, 0, iIndice, STATUS_ABERTO, objCondicaoPagto.colParcelas(iIndice).dtVencimento, dtDataVenctoReal, objCondicaoPagto.colParcelas(iIndice).dValor, objCondicaoPagto.colParcelas(iIndice).dValor, 1, CARTEIRA_CARTEIRA, COBRADOR_PROPRIA_EMPRESA, "", 0, 0, 0, 0, 0, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, 0, "", objCondicaoPagto.colParcelas(iIndice).dValor
            
    Next
    
    NFiscalContrato2_Gera_Parcelas = SUCESSO
    
    Exit Function

Erro_NFiscalContrato2_Gera_Parcelas:

    NFiscalContrato2_Gera_Parcelas = gErr

    Select Case gErr
    
        Case 188556 To 188558

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188559)

    End Select
       
    Exit Function

End Function

Private Function NFiscalContrato2_Gera_Peso(objNFiscal As ClassNFiscal)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim objItemNF As ClassItemNF
Dim dPesoL As Double
Dim dPesoB As Double
Dim dFator As Double

On Error GoTo Erro_NFiscalContrato2_Gera_Peso

    For Each objItemNF In objNFiscal.ColItensNF
    
        Set objProduto = New ClassProduto
    
        objProduto.sCodigo = objItemNF.sProduto
    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 188560
    
        'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
        If lErro <> SUCESSO Then gError 188561
    
        dPesoL = dPesoL + (objProduto.dPesoLiq * objItemNF.dQuantidade * dFator)
        dPesoB = dPesoB + (objProduto.dPesoBruto * objItemNF.dQuantidade * dFator)
    
    Next
    
    objNFiscal.dPesoBruto = dPesoB
    objNFiscal.dPesoLiq = dPesoL
    
    NFiscalContrato2_Gera_Peso = SUCESSO
    
    Exit Function

Erro_NFiscalContrato2_Gera_Peso:

    NFiscalContrato2_Gera_Peso = gErr

    Select Case gErr
    
        Case 188560, 188561

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188562)

    End Select
       
    Exit Function

End Function

Public Sub ContratoFat_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContratoFat_GotFocus()

    Call MaskEdBox_TrataGotFocus(ContratoFat, iAlterado)

End Sub

Public Sub ContratoLabel_Click()

Dim objContrato As New ClassContrato
Dim colSelecao As New Collection

    If Len(Trim(ContratoFat.Text)) > 0 Then
        objContrato.sCodigo = ContratoFat.Text
        objContrato.iFilialEmpresa = giFilialEmpresa
    End If
    
    Call Chama_Tela("ContratosLista", colSelecao, objContrato, objEventoContrato)

End Sub

Private Sub objEventoContrato_evSelecao(obj1 As Object)

Dim objContrato As ClassContrato
Dim bCancel As Boolean
Dim objGeracaoFatContrato As New ClassGeracaoFatContrato
Dim objNFiscal As New ClassNFiscal
Dim lErro As Long

On Error GoTo Erro_objEventoContrato_evSelecao

    Set objContrato = obj1

    objGeracaoFatContrato.sContratoIni = objContrato.sCodigo

    lErro = NFiscalContrato_Gera2(objGeracaoFatContrato, objNFiscal)
    If lErro <> SUCESSO Then gError 188544

    lErro = Traz_NFiscal_Tela2(objNFiscal)
    If lErro <> SUCESSO Then gError 188545

    ContratoFat.Text = objContrato.sCodigo

    Exit Sub
    
Erro_objEventoContrato_evSelecao:
    
    Select Case gErr
    
        Case 188544, 188545

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188546)

    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoContrato_Click()

Dim objContrato As ClassContrato
Dim bCancel As Boolean
Dim objGeracaoFatContrato As New ClassGeracaoFatContrato
Dim objNFiscal As New ClassNFiscal
Dim lErro As Long

On Error GoTo Erro_BotaoContrato_Click

    If Len(Trim(ContratoFat.Text)) = 0 Then gError 188539

    objGeracaoFatContrato.sContratoIni = ContratoFat.Text

    lErro = NFiscalContrato_Gera2(objGeracaoFatContrato, objNFiscal)
    If lErro <> SUCESSO Then gError 188540

    lErro = Traz_NFiscal_Tela2(objNFiscal)
    If lErro <> SUCESSO Then gError 188541

    ContratoFat.Text = objGeracaoFatContrato.sContratoIni

    Exit Sub
    
Erro_BotaoContrato_Click:
    
    Select Case gErr
    
        Case 188539
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONTRATO_PREENCHIDO", gErr)

        Case 188540, 188541

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188542)

    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoVerContrato_Click()

Dim lErro As Long
Dim objContrato As New ClassContrato

On Error GoTo Erro_BotaoDocContrato_Click

    If Len(Trim(ContratoFat.Text)) = 0 Then gError 188543
    
    objContrato.sCodigo = ContratoFat.Text
    objContrato.iFilialEmpresa = giFilialEmpresa

    Call Chama_Tela("ContratoCadastro", objContrato)

    Exit Sub
    
Erro_BotaoDocContrato_Click:
    
    Select Case gErr

        Case 188543
             Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONTRATO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188544)

    End Select
    
    Exit Sub

End Sub

Private Function Traz_NFiscal_Tela2(objNFiscal As ClassNFiscal) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim objNFiscalOriginal As New ClassNFiscal
Dim lErro As Long, objTituloPag As New ClassTituloPagar
Dim bCancel As Boolean
Dim bAlterouCT As Boolean
Dim objItemNF As ClassItemNF
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim colItemPedCompraInfo As New Collection
Dim iLinha As Integer
Dim iIndice As Integer
Dim dPercDesc As Double

On Error GoTo Erro_Traz_NFiscal_Tela2
    
    gbCarregandoTela = True
    
    'Limpa a tela
    Call Limpa_Tela_NFiscal3

    'preenche o Fornecedor
    Cliente.Text = objNFiscal.lCliente
    Call Cliente_Validate(bCancel)
    
    'Preenche a Filial do Fornecedor
    Filial.Text = objNFiscal.iFilialCli
    Call Filial_Validate(bCancel)

    TipoNFiscal.Text = objNFiscal.iTipoNFiscal
    Call TipoNFiscal_Validate(bCancel)

    If objNFiscal.sNaturezaOp <> "" Then gobjTribTab.sNatOpInterna = objNFiscal.sNaturezaOp
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

    DataSaida.PromptInclude = False
    If objNFiscal.dtDataSaida <> DATA_NULA Then DataSaida.Text = Format(objNFiscal.dtDataSaida, "dd/mm/yy")
    DataSaida.PromptInclude = True

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
    If objNFiscal.dValorTotal > 0 Then ValorTotal.Caption = Format(objNFiscal.dValorTotal, "Standard")
    If objNFiscal.dValorProdutos > 0 Then ValorProdutos.Caption = Format(objNFiscal.dValorProdutos, "Standard")

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
    
    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = CStr(objNFiscal.dPesoLiq)
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = CStr(objNFiscal.dPesoBruto)
    
    'preenche o Grid com os Ítens da Nota Fiscal
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 188515

    'Para cada ítem da Coleção
    For Each objItemNF In objNFiscal.ColItensNF

        iLinha = iLinha + 1

        GridItens.TextMatrix(iLinha, iGrid_Contrato_Col) = objItemNF.objItensDeContrato.sContrato
        GridItens.TextMatrix(iLinha, iGrid_ItemContrato_Col) = objItemNF.objItensDeContrato.iSeq
        GridItens.TextMatrix(iLinha, iGrid_DataCobranca_Col) = objItemNF.objItensDeContrato.dtDataProxCobranca

    Next

    'Preenche a Condicao de Pagto
    If objNFiscal.iCondicaoPagto > 0 Then
        CondicaoPagamento.Text = objNFiscal.iCondicaoPagto
        Call CondicaoPagamento_Validate(bCancel)
    End If
        
    lErro = Preenche_Grid_Parcelas(objNFiscal)
    If lErro <> SUCESSO Then gError 188517

    lErro = AlocaAlmoxarifadoPradrao1(objNFiscal)
    If lErro <> SUCESSO Then gError 188518
    
    Set gobjNFiscal = objNFiscal

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 188519
'*** fim tributacao

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
    
    ValorDescontoTit.Text = Format(objNFiscal.dValorDescontoTit, "Standard")
    Call ValorDescontoTit_Validate(bSGECancelDummy)
    
    gbCarregandoTela = False
    
    lErro = objTabComissoes.Comissoes_Calcula()
    If lErro <> SUCESSO Then gError 188575
    
    MsgAutomatica.Value = vbChecked
    Set gobjInfoAdic = objNFiscal.objInfoAdic

    Traz_NFiscal_Tela2 = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela2:

    gbCarregandoTela = False
    
    Traz_NFiscal_Tela2 = gErr

    Select Case gErr

        Case 188515 To 188519, 188575

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188520)

    End Select

    Exit Function

End Function

Private Function AlocaAlmoxarifadoPradrao1(objNFiscal As ClassNFiscal) As Long
'Tenta fazer a alocação do produto no almoxarifado padrão. Caso não consiga chama a tela de Alocação de produto.

Dim lErro As Long
Dim dQuantAlocar As Double
Dim dQuantidade As Double
Dim dFator As Double
Dim iAlmoxarifado As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim dQuantDisponivel As Double
Dim objItemNF As ClassItemNF
Dim iLinhas As Integer
Dim objProduto As New ClassProduto
Dim sUM As String
Dim iLinhaItemNF As Integer

On Error GoTo Erro_AlocaAlmoxarifadoPradrao1

    For Each objItemNF In objNFiscal.ColItensNF

        iLinhaItemNF = iLinhaItemNF + 1

        objProduto.sCodigo = objItemNF.sProduto
    
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 188563
    
        'Se não encontrou o Produto --> erro
        If lErro = 28030 Then gError 188564
    
        dQuantidade = objItemNF.dQuantidade
    
       If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 And gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA Then
    
            sUM = objItemNF.sUnidadeMed
    
            'Faz a conversão da UM da tela para a UM de estoque
            lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 188565
    
            'Converte a quantidade para a UM de estoque
            dQuantAlocar = dQuantidade * dFator
    
            'Busca o Almoxarifado padrão
            lErro = CF("AlmoxarifadoPadrao_Le", objNFiscal.iFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 23796 Then gError 188566
    
            'Se não encontrou
            If iAlmoxarifado = 0 Then gError 188567
    
            objAlmoxarifado.iCodigo = iAlmoxarifado
            
            'Lê o Aloxarifado
            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> 25056 And lErro <> SUCESSO Then gError 188568
            
            If lErro = 25056 Then gError 188569
    
            objEstoqueProduto.sProduto = objProduto.sCodigo
            objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
        
            'Lê os estoques desse produto nesse almoxarifado
            lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 21306 Then gError 188570
            
            'Se não achou o Estoque do Produto --> erro
            If lErro = 21306 Then gError 188571
        
            dQuantDisponivel = objEstoqueProduto.dQuantDisponivel
        
            dQuantAlocar = Arredonda_Estoque(dQuantAlocar)
        
            'Verifica se a Quantidade disponível é maior que a quantidade a alocar
            If (dQuantAlocar - dQuantDisponivel) < QTDE_ESTOQUE_DELTA Then
            
                iLinhas = iLinhas + 1
                'Coloca no Grid alocações a alocação para o almoxarifado padrão
                GridAlocacao.TextMatrix(iLinhas, iGrid_Item_Col) = iLinhaItemNF
                GridAlocacao.TextMatrix(iLinhas, iGrid_ProdutoAloc_Col) = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
                GridAlocacao.TextMatrix(iLinhas, iGrid_AlmoxAloc_Col) = objAlmoxarifado.sNomeReduzido
                GridAlocacao.TextMatrix(iLinhas, iGrid_QuantAloc_Col) = Formata_Estoque(dQuantAlocar)
                GridAlocacao.TextMatrix(iLinhas, iGrid_QuantVend_Col) = Formata_Estoque(dQuantAlocar)
                GridAlocacao.TextMatrix(iLinhas, iGrid_UMAloc_Col) = objProduto.sSiglaUMEstoque
        
            End If
    
        End If
    
    Next
    
    'atualiza o numero de linhas existentes no grid de alocações
    objGridAlocacoes.iLinhasExistentes = iLinhas
        
    AlocaAlmoxarifadoPradrao1 = SUCESSO

    Exit Function

Erro_AlocaAlmoxarifadoPradrao1:

    AlocaAlmoxarifadoPradrao1 = gErr

    Select Case gErr

        Case 188563, 188565, 188566, 188568, 188570

        Case 188564
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 188567
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO1", gErr, objProduto.sCodigo, objNFiscal.iFilialEmpresa)

        Case 188569
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, iAlmoxarifado)

        Case 188571
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188572)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_NFiscal3() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_NFiscal3

    'seta a variavel que limpa a tela como true, para que
    'o calculo de comissoes nao seja chamado na hora da limpeza da tela.
    gbLimpandoTela = True
    
    'Fecha o comando de seta se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call Limpa_Tela_NFiscal4
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    FilialOrcamento.Text = ""
    
    'Seleciona série padrão
    lErro = Seleciona_Serie_Padrao()
    If lErro <> SUCESSO Then gError 193848
    
    iAlterado = 0
    iComissoesAlterada = 0
    
    Call CF2(Me, "NFCF_Limpa_Tela")
    
    'volta com a variavel para false para que o calculo de comissoes
    'possa ser utilizado novamente.
    gbLimpandoTela = False
    
    Exit Function
    
Erro_Limpa_Tela_NFiscal3:
    
    Limpa_Tela_NFiscal3 = gErr
    
    Select Case gErr
    
        Case 193848
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 193849)
        
    End Select
    
    Exit Function
    
End Function

Private Sub Limpa_Tela_NFiscal4()
'Limpa a tela sem fechar setas

Dim iIndice As Integer

    gbLimpandoTela = True
    
    Call Limpa_Tela(Me)
    
    Set gobjAnotacao = Nothing
    Set gobjInfoAdic = Nothing
    
    Set gobjNFiscal = New ClassNFiscal
    
    Status.Caption = ""
       
    TipoNFiscal.Text = ""
    Frame1(TAB_Almoxarifado).Enabled = True
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next
     
    NaoImpCobranca.Value = vbUnchecked
     
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridParcelas)
    Call Grid_Limpa(objGridAlocacoes)
    objTabComissoes.Limpa_Grid_Comissoes
    
'    If giTipoVersao = VERSAO_FULL Then
        Serie.Text = ""
        Call Serie_Validate(bSGECancelDummy)
        TabelaPreco.Text = ""
'    End If
    
    NFiscal.Caption = ""
    Filial.Clear
    FilialEntrega.Clear
    ValorProdutos.Caption = ""
    ValorProdutos2.Caption = ""
    ValorOriginalTit.Caption = ""
    ValorTit.Caption = ""
    ValorTotal.Caption = ""
    Transportadora.Text = ""
    TranspRedespacho.Text = ""
    RedespachoCli.Value = False
    PlacaUF.Text = ""
    CanalVenda.Text = ""
    SerieNFiscalOriginal.Text = ""
    CondicaoPagamento.Text = ""
    TotalPercentualComissao.Caption = ""
    TotalValorComissao.Caption = ""
    TotalValorBase.Caption = ""
    TipoFrete.ListIndex = 0
    Call Seleciona_Tipo_Venda
    
    VolumeMarca.Text = ""
    VolumeEspecie.Text = ""
    
    'Recarrega a combo VolumeEspecie e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)

    'Recarrega a combo VolumeMarca e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
    
    DataReferencia.PromptInclude = False
    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataReferencia.PromptInclude = True

    Call CF2(Me, "NFCF_Default_DataSaida")
    
    Call CF2(Me, "NFF_Limpa_Tela")
       
    iValorDescontoAlterado = 0
    gdDesconto = 0
    Call CF2(Me, "NFCF_Limpa_Desconto")
    iClienteAlterado = 0
    iValorDespesasAlterado = 0
    iValorFreteAlterado = 0
    iValorSeguroAlterado = 0
    giFilialEntregaAlterada = 0
    giFilialEntregaManual = 0
    giDataReferenciaAlterada = 0
    iPercAcresFinAlterado = 0
    '*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.Limpa_Tela
    '*** fim tributacao
   
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    Call DataEmissao_Validate(bSGECancelDummy)
    
    Recibo.Value = vbUnchecked
    
    NumNFe.Caption = ""
    CodVerificacao.Caption = ""
  
    'Limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
    MsgAutomatica.Value = vbChecked

    iAlterado = 0
    iComissoesAlterada = 0
    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0

    gbLimpandoTela = False
    
End Sub

Public Function NFF_Saida_Celula_PrecoUnitario(ByVal objCT As Object, ByVal dPrecoUnitario As Double, ByVal iLinha As Integer) As Long
    NFF_Saida_Celula_PrecoUnitario = SUCESSO
End Function

Public Sub NumNFOriginal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumNFOriginal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub NumNFOriginal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub NumNFOriginal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = NumNFOriginal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ItemNFOriginal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ItemNFOriginal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub ItemNFOriginal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub ItemNFOriginal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ItemNFOriginal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub SerieNFOriginal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SerieNFOriginal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridItens)

End Sub

Public Sub SerieNFOriginal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)

End Sub

Public Sub SerieNFOriginal_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridItens.objControle = SerieNFOriginal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_SerieNFOriginal(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objSerie As New ClassSerie

On Error GoTo Erro_Saida_Celula_SerieNFOriginal

    Set objGridInt.objControle = SerieNFOriginal

    'Verifica se está preenchida
    If Len(Trim(SerieNFOriginal.Text)) > 0 Then

        objSerie.sSerie = SerieNFOriginal.Text
        objSerie.iFilialEmpresa = giFilialEmpresa

        'Lê a Série
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO And lErro <> 22202 Then gError 199736
        
        'Se a série não está cadastrada ==> erro
        If lErro = 22202 Then gError 199737

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 199738

    Saida_Celula_SerieNFOriginal = SUCESSO

    Exit Function

Erro_Saida_Celula_SerieNFOriginal:

    Saida_Celula_SerieNFOriginal = gErr

    Select Case gErr

        Case 199736, 199738
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 199737
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199739)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_NumNFOrig(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_NumNFOrig

    Set objGridInt.objControle = NumNFOriginal

    'Verifica se está preenchido
    If Len(Trim(NumNFOriginal.ClipText)) > 0 Then

        'Verifica se é um Long
        lErro = Long_Critica(NumNFOriginal.Text)
        If lErro <> SUCESSO Then gError 199740

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 199741

    Saida_Celula_NumNFOrig = SUCESSO

    Exit Function

Erro_Saida_Celula_NumNFOrig:

    Saida_Celula_NumNFOrig = gErr

    Select Case gErr

        Case 199740, 199741
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199742)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ItemNFOrig(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ItemNFOrig

    Set objGridInt.objControle = ItemNFOriginal

    'Verifica se está preenchido
    If Len(Trim(ItemNFOriginal.ClipText)) > 0 Then

        'Verifica se é um Inteiro
        lErro = Inteiro_Critica(ItemNFOriginal.Text)
        If lErro <> SUCESSO Then gError 199743

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 199744

    Saida_Celula_ItemNFOrig = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemNFOrig:

    Saida_Celula_ItemNFOrig = gErr

    Select Case gErr

        Case 199743, 199744
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199745)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157400)
            
    End Select

    Exit Function
    
End Function

Function ParcelasOV_Gerando_ParcelasNF(objOV As ClassOrcamentoVenda, objNFiscal As ClassNFiscal) As Long
'Se a Data de referencia do Pedido estiver ultrapassada em regalação a Data Atual e se
'a cobrança automática do pedido estiver selecionada as  parcelas são atualizadas.
    
Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
'Dim colDataVencimento As New Collection
Dim iIndice As Integer
Dim objParcelaOV As New ClassParcelaOV
Dim objParcRec As ClassParcelaReceber
Dim dtDataVencimento As Date

On Error GoTo Erro_ParcelasOV_Gerando_ParcelasNF
    
    'Se a condição de pagamento estiver preenchida e cobrança automática selecionada e a data de referencia do pedido estiver vazia ou ultrapassada
    If objOV.iCondicaoPagto <> 0 And objOV.iCobrancaAutomatica = 1 And (objOV.dtDataReferencia = DATA_NULA Or objNFiscal.dtDataReferencia < gdtDataAtual) Then
        'A data de referencia da nota fiscal é atualizada
        objNFiscal.dtDataReferencia = gdtDataAtual
        
        objCondicaoPagto.iCodigo = objOV.iCondicaoPagto
        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 59095
        If lErro = 19205 Then gError 59096
        
        If objCondicaoPagto.iNumeroParcelas <> objOV.colParcela.Count Then gError 59097
        
        'Calcula Datas de Vencimento das Parcelas
        'lErro = CF("Parcelas_DatasVencimento", objCondicaoPagto, objNFiscal.dtDataReferencia, colDataVencimento)
        objCondicaoPagto.dtDataRef = objNFiscal.dtDataReferencia
        objCondicaoPagto.dtDataEntrega = StrParaDate(DataSaida.Text)
        If objCondicaoPagto.dtDataEntrega = DATA_NULA Then objCondicaoPagto.dtDataEntrega = StrParaDate(DataEmissao.Text)
        
        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True)
        If lErro <> SUCESSO Then gError 59098

        'Loop de preenchimento de Parcelas com Datas de Vencimento Recalculadas
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas

            Set objParcelaOV = objOV.colParcela.Item(iIndice)
            
            dtDataVencimento = objCondicaoPagto.colParcelas(iIndice).dtVencimento
            
            'Preenche o desconto padrão
            lErro = DescontoPadraoParcela_Gera(objParcelaOV, dtDataVencimento) 'colDataVencimento(iIndice))
            If lErro <> SUCESSO Then gError 51613
                        
            With objParcelaOV
                'Adiciona a parcela na coleção de parcelas
                Set objParcRec = objNFiscal.ColParcelaReceber.Add(0, 0, .iNumParcela, STATUS_ABERTO, dtDataVencimento, dtDataVencimento, .dValor, .dValor, 1, CARTEIRA_CARTEIRA, COBRADOR_PROPRIA_EMPRESA, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0, 0, "", .dValor)
            End With

        Next

    Else
    
        'transfere as parcelas do pedido de vendas para a nota fiscal
        For Each objParcelaOV In objOV.colParcela

            With objParcelaOV
                Set objParcRec = objNFiscal.ColParcelaReceber.Add(0, 0, .iNumParcela, STATUS_ABERTO, .dtDataVencimento, .dtDataVencimento, .dValor, .dValor, 1, CARTEIRA_CARTEIRA, COBRADOR_PROPRIA_EMPRESA, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0, 0, "", .dValor)
            End With

        Next
    
    End If
    
    ParcelasOV_Gerando_ParcelasNF = SUCESSO
    
    Exit Function

Erro_ParcelasOV_Gerando_ParcelasNF:

    ParcelasOV_Gerando_ParcelasNF = gErr
    
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
Dim objProduto As New ClassProduto, objItemNF As ClassItemNF
Dim dPrecoUnitario As Double, colMedicao As Collection

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
                    
                    lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(iItem, iGrid_UnidadeMed_Col), objRomaneioGradeAux.objTabelaPrecoItem, dPrecoUnitario)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                    If dPrecoUnitario > 0 Then
                        GridItens.TextMatrix(iItem, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
                    Else
                        GridItens.TextMatrix(iItem, iGrid_ValorUnitario_Col) = ""
                    End If
                    
                End If
               
                Call gobjTelaProjetoInfo.Adiciona_Linha(iItem)
                
                'permite que a tributacao reflita a inclusao de uma linha no grid
                lErro = gobjTribTab.Inclusao_Item_Grid(iItem, objProduto.sCodigo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Set colMedicao = New Collection
            
                gcolcolMedicoes.Add colMedicao
            
                Set objItemNF = New ClassItemNF
                
                objItemNF.iPossuiGrade = MARCADO
                
                objItemNF.sProduto = objProduto.sCodigo
                objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
                objItemNF.iItem = iItem
                objItemNF.lNumIntDoc = 0
                objItemNF.sDescricaoItem = GridItens.TextMatrix(iItem, iGrid_DescProduto_Col)
            
                gobjNFiscal.ColItensNF.Add1 objItemNF
                gobjNFiscal.ColItensNF(iItem).sUMEstoque = objProduto.sSiglaUMEstoque

                Set objItemNF.colItensRomaneioGrade = objRomaneioGradeAux.colItensRomaneioGrade
                
                GridItens.TextMatrix(iItem, 0) = "# " & GridItens.TextMatrix(iItem, 0)
                       
                Call Atualiza_Grid_Itens(objItemNF)
                
                Call PrecoTotal_Calcula(iItem, False)
            
                Call Atualiza_Grid_Alocacao(objItemNF)
            
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
    If lErro <> SUCESSO Then gError 207431

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
        Case 207431

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207432)

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
        
        lErro = CF2(Me, "NFCF_PrecoUnitario_Calcula", giFilialEmpresa, sProduto, GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), Cliente.Text, Codigo_Extrai(Filial.Text), iCodTabela, StrParaDbl(PercAcrescFin.Text) / 100, MOEDA_REAL, Nothing, dPrecoUnitario, False, dPercDescontoTab, dPrecoComDescontoTab)
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

Private Function Trata_Complemento_Preco() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim objItemGrade As ClassItemRomaneioGrade

On Error GoTo Erro_Trata_Complemento_Preco

    If gobjFAT.iZeraQtdeEmNFComplPreco = MARCADO Then
    
        'Para cada linha existente do Grid
        For iIndice = 1 To objGridItens.iLinhasExistentes
    
            GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col)
            GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(0)
            
            Set objItemNF = gobjNFiscal.ColItensNF(iIndice)
            objItemNF.dQuantidade = 0
            objItemNF.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))
            For Each objItemGrade In objItemNF.colItensRomaneioGrade
                objItemGrade.dQuantidade = 0
            Next
            
            Call Remove_Alocacoes(iIndice)
        
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) > 0 Then
                
                Call PrecoTotal_Calcula(iIndice)
                           
                lErro = gobjTribTab.Alteracao_Item_Grid(iIndice)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            End If
    
        Next
        
        lErro = ValorTotal_Calcula()
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    Trata_Complemento_Preco = SUCESSO

    Exit Function

Erro_Trata_Complemento_Preco:

    Trata_Complemento_Preco = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157457)

    End Select

    Exit Function
    
End Function
