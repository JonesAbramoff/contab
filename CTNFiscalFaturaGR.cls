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
''NAO POSSUI ALTERACOES DE COMISSOES FEITAS POR TULIO
''Mario falou que a GR nao usa essa tela, eles so usam NF simples....
''por tulio em 10/05/02
'
''Option Explicit
'
''19/03/01: modificacao p/gsilva: incluidos campos ValorMercadoria, ValorContainer, Origem, Destino
''16/10/01: Incluidos no Tab Parcelas as combos de Cobrador a CarteiraCobrador por Daniel
'
'Dim m_objUserControl As Object
'
'Dim giNaoRepeteEvento As String
'
''Property Variables:
'Dim m_Caption As String
'Event Unload()
'
''Flag que indica se a tela está sendo preenchida.
'Public gbCarregandoTela As Boolean
'Dim gbLimpandoTela As Boolean
'Public giPosCargaOk As Integer
'
''inicia objeto associado a GridComissoes
'Public objTabComissoes As New ClassTabComissoes
'
''inicio contabilidade
'Public objGrid1 As AdmGrid
'Public objContabil As New ClassContabil
'
'Public WithEvents objEventoLote As AdmEvento
'Public WithEvents objEventoDoc As AdmEvento
'
''mnemonicos
'Private Const TIPO_NOTA_FISCAL As String = "Tipo_Nota_Fiscal"
'Private Const NATUREZA_OP As String = "Natureza_OP"
'Private Const CLIENTE1 As String = "Cliente"
'Private Const CLIENTE_CODIGO As String = "Cliente_Codigo"
'Private Const FILIAL1 As String = "Filial_Cliente"
'Private Const Serie1 As String = "Serie"
'Private Const DATA_EMISSAO As String = "Data_Emissao"
'Private Const DATA_SAIDA As String = "Data_Saida"
'Private Const PRODUTO1 As String = "Produto_Codigo"
'Private Const UNIDADE_MED As String = "Unidade_Med"
'Private Const QUANTIDADE1 As String = "Quantidade"
'Private Const PRECO_UNITARIO As String = "Preco_Unitario"
'Private Const PRECO_TOTAL As String = "Preco_Total"
'Private Const DESCONTO1 As String = "Desconto_Produto"
'Private Const DESCRICAO_ITEM As String = "Descricao_Item"
'Private Const CTACONTABILEST1 As String = "ContaContabilEst"
''#######Valores####################
'Private Const ICMS As String = "ICMS_Valor"
'Private Const ICMSSUBST As String = "ICMSSubst_Valor"
'Private Const VALOR_PRODUTOS As String = "Valor_Produtos"
'Private Const VALOR_FRETE As String = "Valor_Frete"
'Private Const VALOR_SEGURO As String = "Valor_Seguro"
'Private Const VALOR_DESPESAS As String = "Valor_Despesas"
'Private Const IPI As String = "IPI_Valor"
'Private Const VALOR_DESCONTO As String = "Valor_Desconto_Glb"
'Private Const VALOR_TOTAL As String = "Valor_Total"
''###########Almoxarifado############
'Private Const PRODUTO_ALMOX As String = "Produto_Almox"
'Private Const ALMOX1 As String = "Almoxarifado"
'Private Const QUANT_ALOCADA As String = "Quant_Alocada"
'Private Const UNIDADE_MED_EST As String = "Unidade_Med_Est"
''###########Pagamento###############
'Private Const PERC_ACRESC_FIN As String = "Perc_Acresc_Fin"
'Private Const VALOR_PARCELA As String = "Valor_Parcela"
''###########Tributação##############
'Private Const ISS_VALOR As String = "ISS_Valor"
'Private Const ISS_INCLUSO As String = "ISS_Incluso"
'Private Const VALOR_IRRF As String = "Valor_IRRF"
'Private Const QUANT_DISPONIVEL As String = "Quant_Disponivel"
'Private Const QUANT_CONSIGTERC As String = "Quant_ConsigTerc"
'Private Const VALOR_INSS As String = "Valor_INSS"
'Private Const INSS_RETIDO As String = "INSS_Retido"
''fim da contabilidade
'
'Public iAlterado As Integer
'Public iPercAcresFinAlterado As Integer
'Public iFrameAtual As Integer
'Public gdDesconto As Double
'Public iClienteAlterado As Integer
'Public iValorDescontoAlterado As Integer  'É usado em ValorTotal_Change, CondicaoPagamento_Validate, GridParcelas_Preenche
'Public iValorFreteAlterado As Integer
'Public iValorSeguroAlterado As Integer
'Public iValorDespesasAlterado As Integer
'Public iValorIRRFAlterado As Integer
'Public giFilialEntregaAlterada As Integer
'Public giFilialEntregaManual As Integer
'Public giDataReferenciaAlterada As Integer
'
'Public objGridParcelas As AdmGrid
'Dim iGrid_Vencimento_col As Integer
'Dim iGrid_ValorParcela_Col As Integer
'Dim iGrid_Cobranca_Col As Integer
'Dim iGrid_CarteiraCobranca_Col As Integer
'Dim iGrid_Desc1Codigo_Col As Integer
'Dim iGrid_Desc1Ate_Col As Integer
'Dim iGrid_Desc1Valor_Col As Integer
'Dim iGrid_Desc1Percentual_Col As Integer
'Dim iGrid_Desc2Codigo_Col As Integer
'Dim iGrid_Desc2Ate_Col As Integer
'Dim iGrid_Desc2Valor_Col As Integer
'Dim iGrid_Desc2Percentual_Col As Integer
'Dim iGrid_Desc3Codigo_Col As Integer
'Dim iGrid_Desc3Ate_Col As Integer
'Dim iGrid_Desc3Valor_Col As Integer
'Dim iGrid_Desc3Percentual_Col As Integer
'
'Public objGridItens As AdmGrid
'Public iGrid_Produto_Col As Integer
'Public iGrid_DescProduto_Col As Integer
'Public iGrid_UnidadeMed_Col As Integer
'Public iGrid_Quantidade_Col As Integer
'Public iGrid_PrecoUnitario_Col As Integer
'Public iGrid_PercDesc_Col As Integer
'Public iGrid_Desconto_Col As Integer
'Public iGrid_PrecoTotal_Col As Integer
'
'Public objGridComissoes As AdmGrid
'
'Public objGridAlocacoes As AdmGrid
'Public iGrid_Item_Col As Integer
'Public iGrid_ProdutoAloc_Col As Integer
'Public iGrid_AlmoxAloc_Col As Integer
'Public iGrid_UMAloc_Col As Integer
'Public iGrid_QuantAloc_Col As Integer
'Public iGrid_QuantVend_Col As Integer
'
'Private WithEvents objEventoCliente As AdmEvento
'Private WithEvents objEventoCondPagto As AdmEvento
'Private WithEvents objEventoEstoqueProd As AdmEvento
'Public WithEvents objEventoNFiscal As AdmEvento
'Private WithEvents objEventoNFiscalOriginal As AdmEvento
'Private WithEvents objEventoProduto As AdmEvento
'Private WithEvents objEventoSerie As AdmEvento
'Private WithEvents objEventoSerieOriginal As AdmEvento
'Private WithEvents objEventoTransportadora As AdmEvento
'Public WithEvents objEventoVendedor As AdmEvento
'Private WithEvents objEventoMensagem As AdmEvento
'
'''*** incluidos p/tratamento de tributacao *******************************
'Public gobjTribTab As ClassTribTab
'
'Public gobjRastreamento As ClassRastreamento
'
''Constantes públicas dos tabs
'Private Const TAB_DadosPrincipais = 1
'Private Const TAB_Comprovantes = 2
'Private Const TAB_Itens = 3
'Private Const TAB_Complemento = 4
'Private Const TAB_Cobranca = 5
'Private Const TAB_Comissoes = 6
'Private Const TAB_Almoxarifado = 7
'Private Const TAB_Tributacao = 8
'Private Const TAB_Contabilizacao = 9
'Private Const TAB_RASTRO = 10
'
'Const TIPODOCINFO_COMPLEMENTO_IPI = 53
'
''Leo daqui p/ baixo
''???? Já existe na tela de Cotacao Moeda.
'Const MOEDA_DOLAR = 1
'Private Type typeCotacaoMoeda
'    dValor As Double
'    iMoeda As Integer
'    dtData As Date
'End Type
''***************************************
'
''****JÁ*EXISTE*EM**COMPROVANTESERVICO***
'Private Type typeItemComprovSer
'    iFilialEmpresa As Integer
'    lNumero As Long
'    dtData As Date
'    sEquipamento As String
'    dQuantidade As Double
'    dValor As Double
'    dQuantAdicional As Double
'    dValorAdicional As Double
'    dValorTotal As Double
'    iItem As Integer
'End Type
'
'
''****JÁ*EXISTE*EM**COMPROVANTESERVICO***
'Type typeComprovServ
'    iFilialEmpresa As Integer
'    lNumero As Long
'    dtData As Date
'    lCliente As Long
'    lFilialCliente As Long
'    iTabela As Integer
'    lNotaFiscal As Long
'    dPesoBruto As Double
'    dValorMerc As Double
'    dAdValoren As Double
'    dDespacho As Double
'    dManuseio As Double
'    dSabado As Double
'    dISSICMS As Double
'    dPedagio As Double
'    dTotal As Double
'    sObservacao As String
'    lNumIntNota As Long
'    dValorContainer As Double
'    dDesconto As Double
'    dFrete As Double
'End Type
'
'Dim iValorContainerAlterado As Integer
'
''Essa constante já existe em CompServico
'Const TIPO_NFISCAL = 2
'
'Const BOTAO_CONSULTATITREC = 1
'Const BOTAO_PRODUTOS = 2
'Const BOTAO_ESTOQUE_PRODUTOS = 3
'Const BOTAO_LOCALIZACAO_PRODUTOS = 4
'Const BOTAO_LOTES = 5
'Const BOTAO_GRAVAR = 6
'Const BOTAO_EXCLUIR = 7
'Const BOTAO_LIMPAR = 8
'Const BOTAO_FECHAR = 9
'Const BOTAO_LIMPARNF = 10
'Const BOTAO_DATAREFERENCIAUP = 11
'Const BOTAO_DATAREFERENCIADOWN = 12
'
'Const UPDOWN_EMISSAO = 1
'Const UPDOWN_SAIDA = 2
'
'
''Constantes correspondentes aos Index's dos controles Label1
'Const LABEL1_CLIENTELABEL = 41
'Const LABEL1_LBLNATOPINTERNA = 69
'Const LABEL1_SERIELABEL = 70
'Const LABEL1_TRANSPORTADORALABEL = 71
'Const LABEL1_MENSAGEMLABEL = 72
'Const LABEL1_CONDPAGTOLABEL = 73
'Const LABEL1_SERIENFORIGINALLABEL = 74
'Const LABEL1_NFISCALORIGINALLABEL = 75
'Const LABEL1_NFISCALLABEL = 76
'
''Variaveis e Constatanets relacionadas ao GridComprovServ
'Const NUM_MAXIMO_CONHECFRETE = 240
'
'Public objGridComprovServ As New AdmGrid
'
'Dim iGrid_ValorMercadoriaCon_col As Integer
''Dim iGrid_FretePesoCon_col As Integer
'Dim iGrid_AdValorCon_col As Integer
'Dim iGrid_SabadoCon_col As Integer
'Dim iGrid_ValorContainerCon_col As Integer
'Dim iGrid_ComprovServCon_col As Integer
'Dim iGrid_SubTotalCon_col As Integer
'Dim iGrid_ISSCon_col As Integer
'Dim iGrid_TotalCon_col As Integer
'Dim iGrid_SelecionaCon_col As Integer
'Dim iGrid_PedagioCon_Col As Integer
'Dim iGrid_DespachoCon_Col As Integer
'Dim iGrid_DataCon_Col As Integer
'Dim iGrid_Manuseio_Col As Integer
''Leo até aqui
'
''Constantes de Erro (leo)
'
'Const ERRO_COTACAOMOEDA_DOLAR_NAO_ENCONTRADA = 0 'Parametros : dtDataCotacao
''A Cotação do Dólar no dia %s não foi encontrada.
'
'
''--- inicio dos properties get dos controles da tela
'
''jones-15/03/01
'Public Property Get INSSRetido() As Object
'     Set INSSRetido = objUserControl.Controls("INSSRetido")
'End Property
'
'Public Property Get INSSDeducoes() As Object
'     Set INSSDeducoes = objUserControl.Controls("INSSDeducoes")
'End Property
'
'Public Property Get INSSBase() As Object
'     Set INSSBase = objUserControl.Controls("INSSBase")
'End Property
'
'Public Property Get INSSValor() As Object
'     Set INSSValor = objUserControl.Controls("INSSValor")
'End Property
''fim jones-15/03/01
'
'Public Property Get Frame1() As Object
'     Set Frame1 = objUserControl.Controls("Frame1")
'End Property
'
'Public Property Get FrameCliente() As Object
'     Set FrameCliente = objUserControl.Controls("FrameCliente")
'End Property
'
'Public Property Get Filial() As Object
'     Set Filial = objUserControl.Controls("Filial")
'End Property
'
'Public Property Get Fornecedor() As Object
'     Set Fornecedor = objUserControl.Controls("Fornecedor")
'End Property
'
'Public Property Get Cliente() As Object
'     Set Cliente = objUserControl.Controls("Cliente")
'End Property
'
''Public Property Get ClienteLabel() As Object
''     Set ClienteLabel = objUserControl.Controls("ClienteLabel")
''End Property
'
'Public Property Get FornecedorLabel() As Object
'     Set FornecedorLabel = objUserControl.Controls("FornecedorLabel")
'End Property
'
'Public Property Get FramePrecos() As Object
'     Set FramePrecos = objUserControl.Controls("FramePrecos")
'End Property
'
'Public Property Get CondicaoPagamento() As Object
'     Set CondicaoPagamento = objUserControl.Controls("CondicaoPagamento")
'End Property
'
'Public Property Get TabelaPreco() As Object
'     Set TabelaPreco = objUserControl.Controls("TabelaPreco")
'End Property
'
'Public Property Get PercAcrescFin() As Object
'     Set PercAcrescFin = objUserControl.Controls("PercAcrescFin")
'End Property
'
'Public Property Get CondPagtoLabel() As Object
'     Set CondPagtoLabel = objUserControl.Controls("CondPagtoLabel")
'End Property
'
'Public Property Get FrameDatas() As Object
'     Set FrameDatas = objUserControl.Controls("FrameDatas")
'End Property
'
'Public Property Get UpDownEmissao() As Object
'     Set UpDownEmissao = objUserControl.Controls("UpDownEmissao")
'End Property
'
'Public Property Get DataEmissao() As Object
'     Set DataEmissao = objUserControl.Controls("DataEmissao")
'End Property
'
'Public Property Get UpDownSaida() As Object
'     Set UpDownSaida = objUserControl.Controls("UpDownSaida")
'End Property
'
'Public Property Get DataSaida() As Object
'     Set DataSaida = objUserControl.Controls("DataSaida")
'End Property
'
''horasaida
'Public Property Get HoraSaida() As Object
'     Set HoraSaida = objUserControl.Controls("HoraSaida")
'End Property
'
'Public Property Get FrameIdentificacao() As Object
'     Set FrameIdentificacao = objUserControl.Controls("FrameIdentificacao")
'End Property
'
'Public Property Get BotaoLimparNF() As Object
'     Set BotaoLimparNF = objUserControl.Controls("BotaoLimparNF")
'End Property
'
'Public Property Get TipoNFiscal() As Object
'     Set TipoNFiscal = objUserControl.Controls("TipoNFiscal")
'End Property
'
'Public Property Get Serie() As Object
'     Set Serie = objUserControl.Controls("Serie")
'End Property
'
'Public Property Get NatOpInterna() As Object
'     Set NatOpInterna = objUserControl.Controls("NatOpInterna")
'End Property
'
'Public Property Get NFiscal() As Object
'     Set NFiscal = objUserControl.Controls("NFiscal")
'End Property
'
'Public Property Get NFiscalLabel() As Object
'     Set NFiscalLabel = objUserControl.Controls("NFiscalLabel")
'End Property
'
'Public Property Get SerieLabel() As Object
'     Set SerieLabel = objUserControl.Controls("SerieLabel")
'End Property
'
'Public Property Get LblNatOpInterna() As Object
'     Set LblNatOpInterna = objUserControl.Controls("LblNatOpInterna")
'End Property
'
'Public Property Get ImageStatus() As Object
'     Set ImageStatus = objUserControl.Controls("ImageStatus")
'End Property
'
'Public Property Get FrameNFiscalOriginal() As Object
'     Set FrameNFiscalOriginal = objUserControl.Controls("FrameNFiscalOriginal")
'End Property
'
'Public Property Get NFiscalPropria() As Object
'     Set NFiscalPropria = objUserControl.Controls("NFiscalPropria")
'End Property
'
'Public Property Get SerieNFiscalOriginal() As Object
'     Set SerieNFiscalOriginal = objUserControl.Controls("SerieNFiscalOriginal")
'End Property
'
'Public Property Get BotaoExibirDados() As Object
'     Set BotaoExibirDados = objUserControl.Controls("BotaoExibirDados")
'End Property
'
'Public Property Get NFiscalOriginal() As Object
'     Set NFiscalOriginal = objUserControl.Controls("NFiscalOriginal")
'End Property
'
'Public Property Get SerieOriginalLabel() As Object
'     Set SerieOriginalLabel = objUserControl.Controls("SerieOriginalLabel")
'End Property
'
'Public Property Get NFOriginalLabel() As Object
'     Set NFOriginalLabel = objUserControl.Controls("NFOriginalLabel")
'End Property
'
'Public Property Get CobrancaAutomatica() As Object
'     Set CobrancaAutomatica = objUserControl.Controls("CobrancaAutomatica")
'End Property
'
'Public Property Get BotaoDataReferenciaUp() As Object
'     Set BotaoDataReferenciaUp = objUserControl.Controls("BotaoDataReferenciaUp")
'End Property
'
'Public Property Get Desconto1Codigo() As Object
'     Set Desconto1Codigo = objUserControl.Controls("Desconto1Codigo")
'End Property
'
'Public Property Get Desconto2Codigo() As Object
'     Set Desconto2Codigo = objUserControl.Controls("Desconto2Codigo")
'End Property
'
'Public Property Get Desconto3Codigo() As Object
'     Set Desconto3Codigo = objUserControl.Controls("Desconto3Codigo")
'End Property
'
'Public Property Get BotaoDataReferenciaDown() As Object
'     Set BotaoDataReferenciaDown = objUserControl.Controls("BotaoDataReferenciaDown")
'End Property
'
'Public Property Get Desconto1Percentual() As Object
'     Set Desconto1Percentual = objUserControl.Controls("Desconto1Percentual")
'End Property
'
'Public Property Get Desconto3Valor() As Object
'     Set Desconto3Valor = objUserControl.Controls("Desconto3Valor")
'End Property
'
'Public Property Get Desconto3Ate() As Object
'     Set Desconto3Ate = objUserControl.Controls("Desconto3Ate")
'End Property
'
'Public Property Get Desconto2Valor() As Object
'     Set Desconto2Valor = objUserControl.Controls("Desconto2Valor")
'End Property
'
'Public Property Get Desconto2Ate() As Object
'     Set Desconto2Ate = objUserControl.Controls("Desconto2Ate")
'End Property
'
'Public Property Get Desconto1Valor() As Object
'     Set Desconto1Valor = objUserControl.Controls("Desconto1Valor")
'End Property
'
'Public Property Get Desconto1Ate() As Object
'     Set Desconto1Ate = objUserControl.Controls("Desconto1Ate")
'End Property
'
'Public Property Get DataVencimento() As Object
'     Set DataVencimento = objUserControl.Controls("DataVencimento")
'End Property
'
'Public Property Get ValorParcela() As Object
'     Set ValorParcela = objUserControl.Controls("ValorParcela")
'End Property
'
'Public Property Get Desconto2Percentual() As Object
'     Set Desconto2Percentual = objUserControl.Controls("Desconto2Percentual")
'End Property
'
'Public Property Get Desconto3Percentual() As Object
'     Set Desconto3Percentual = objUserControl.Controls("Desconto3Percentual")
'End Property
'
'Public Property Get GridParcelas() As Object
'     Set GridParcelas = objUserControl.Controls("GridParcelas")
'End Property
'
'Public Property Get DataReferencia() As Object
'     Set DataReferencia = objUserControl.Controls("DataReferencia")
'End Property
'
'Public Property Get BotaoConsultaTitRec() As Object
'     Set BotaoConsultaTitRec = objUserControl.Controls("BotaoConsultaTitRec")
'End Property
'
''Public Property Get Frame7() As Object
''     Set Frame7 = objUserControl.Controls("Frame7")
''End Property
'
'Public Property Get ProdutoAlmox() As Object
'     Set ProdutoAlmox = objUserControl.Controls("ProdutoAlmox")
'End Property
'
'Public Property Get GridAlocacao() As Object
'     Set GridAlocacao = objUserControl.Controls("GridAlocacao")
'End Property
'
'Public Property Get Almox() As Object
'     Set Almox = objUserControl.Controls("Almox")
'End Property
'
'Public Property Get QuantAlocada() As Object
'     Set QuantAlocada = objUserControl.Controls("QuantAlocada")
'End Property
'
'Public Property Get ItemNFiscal() As Object
'     Set ItemNFiscal = objUserControl.Controls("ItemNFiscal")
'End Property
'
'Public Property Get UnidadeMedEst() As Object
'     Set UnidadeMedEst = objUserControl.Controls("UnidadeMedEst")
'End Property
'
'Public Property Get QuantVendida() As Object
'     Set QuantVendida = objUserControl.Controls("QuantVendida")
'End Property
'
'Public Property Get BotaoLocalizacao() As Object
'     Set BotaoLocalizacao = objUserControl.Controls("BotaoLocalizacao")
'End Property
'
'Public Property Get BotaoVendedores() As Object
'     Set BotaoVendedores = objUserControl.Controls("BotaoVendedores")
'End Property
'
'Public Property Get ComissaoAutomatica() As Object
'     Set ComissaoAutomatica = objUserControl.Controls("ComissaoAutomatica")
'End Property
'
'Public Property Get ValorComissao() As Object
'     Set ValorComissao = objUserControl.Controls("ValorComissao")
'End Property
'
'Public Property Get ValorBase() As Object
'     Set ValorBase = objUserControl.Controls("ValorBase")
'End Property
'
'Public Property Get PercentualComissao() As Object
'     Set PercentualComissao = objUserControl.Controls("PercentualComissao")
'End Property
'
'Public Property Get Vendedor() As Object
'     Set Vendedor = objUserControl.Controls("Vendedor")
'End Property
'
'Public Property Get ValorEmissao() As Object
'     Set ValorEmissao = objUserControl.Controls("ValorEmissao")
'End Property
'
'Public Property Get PercentualEmissao() As Object
'     Set PercentualEmissao = objUserControl.Controls("PercentualEmissao")
'End Property
'
'Public Property Get ValorBaixa() As Object
'     Set ValorBaixa = objUserControl.Controls("ValorBaixa")
'End Property
'
'Public Property Get PercentualBaixa() As Object
'     Set PercentualBaixa = objUserControl.Controls("PercentualBaixa")
'End Property
'
'Public Property Get GridComissoes() As Object
'     Set GridComissoes = objUserControl.Controls("GridComissoes")
'End Property
'
'Public Property Get LabelTotaisComissoes() As Object
'     Set LabelTotaisComissoes = objUserControl.Controls("LabelTotaisComissoes")
'End Property
'
'Public Property Get TotalValorComissao() As Object
'     Set TotalValorComissao = objUserControl.Controls("TotalValorComissao")
'End Property
'
'Public Property Get TotalPercentualComissao() As Object
'     Set TotalPercentualComissao = objUserControl.Controls("TotalPercentualComissao")
'End Property
'
''Public Property Get Frame4() As Object
''     Set Frame4 = objUserControl.Controls("Frame4")
''End Property
'
''Public Property Get Frame6() As Object
''     Set Frame6 = objUserControl.Controls("Frame6")
''End Property
'
'Public Property Get Destinatario() As Object
'     Set Destinatario = objUserControl.Controls("Destinatario")
'End Property
'
'Public Property Get Emitente() As Object
'     Set Emitente = objUserControl.Controls("Emitente")
'End Property
'Public Property Get ImprimeRomaneio() As Object
'     Set ImprimeRomaneio = objUserControl.Controls("ImprimeRomaneio")
'End Property
'
'Public Property Get FilialEntrega() As Object
'     Set FilialEntrega = objUserControl.Controls("FilialEntrega")
'End Property
'
'Public Property Get Transportadora() As Object
'     Set Transportadora = objUserControl.Controls("Transportadora")
'End Property
'
'Public Property Get Placa() As Object
'     Set Placa = objUserControl.Controls("Placa")
'End Property
'
'Public Property Get PlacaUF() As Object
'     Set PlacaUF = objUserControl.Controls("PlacaUF")
'End Property
'
'Public Property Get TransportadoraLabel() As Object
'     Set TransportadoraLabel = objUserControl.Controls("TransportadoraLabel")
'End Property
'
''Public Property Get Frame8() As Object
''     Set Frame8 = objUserControl.Controls("Frame8")
''End Property
'
'Public Property Get VolumeNumero() As Object
'     Set VolumeNumero = objUserControl.Controls("VolumeNumero")
'End Property
'
'Public Property Get VolumeEspecie() As Object
'     Set VolumeEspecie = objUserControl.Controls("VolumeEspecie")
'End Property
'
'Public Property Get VolumeMarca() As Object
'     Set VolumeMarca = objUserControl.Controls("VolumeMarca")
'End Property
'
'Public Property Get VolumeQuant() As Object
'     Set VolumeQuant = objUserControl.Controls("VolumeQuant")
'End Property
'
''Public Property Get Frame5() As Object
''     Set Frame5 = objUserControl.Controls("Frame5")
''End Property
'
'Public Property Get Mensagem() As Object
'     Set Mensagem = objUserControl.Controls("Mensagem")
'End Property
'
'Public Property Get CanalVenda() As Object
'     Set CanalVenda = objUserControl.Controls("CanalVenda")
'End Property
'
'Public Property Get NumPedidoTerc() As Object
'     Set NumPedidoTerc = objUserControl.Controls("NumPedidoTerc")
'End Property
'
'Public Property Get PesoLiquido() As Object
'     Set PesoLiquido = objUserControl.Controls("PesoLiquido")
'End Property
'
'Public Property Get PesoBruto() As Object
'     Set PesoBruto = objUserControl.Controls("PesoBruto")
'End Property
'
'Public Property Get LabelPedidoTerc() As Object
'     Set LabelPedidoTerc = objUserControl.Controls("LabelPedidoTerc")
'End Property
'
'Public Property Get MensagemLabel() As Object
'     Set MensagemLabel = objUserControl.Controls("MensagemLabel")
'End Property
'
'Public Property Get BotaoItensNFiscalOrig() As Object
'     Set BotaoItensNFiscalOrig = objUserControl.Controls("BotaoItensNFiscalOrig")
'End Property
'
''Public Property Get Frame3() As Object
''     Set Frame3 = objUserControl.Controls("Frame3")
''End Property
'
'Public Property Get UnidadeMed() As Object
'     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
'End Property
'
'Public Property Get DescricaoItem() As Object
'     Set DescricaoItem = objUserControl.Controls("DescricaoItem")
'End Property
'
'Public Property Get Desconto() As Object
'     Set Desconto = objUserControl.Controls("Desconto")
'End Property
'
'Public Property Get PercentDesc() As Object
'     Set PercentDesc = objUserControl.Controls("PercentDesc")
'End Property
'
'Public Property Get PrecoUnitario() As Object
'     Set PrecoUnitario = objUserControl.Controls("PrecoUnitario")
'End Property
'
'Public Property Get Quantidade() As Object
'     Set Quantidade = objUserControl.Controls("Quantidade")
'End Property
'
'Public Property Get Produto() As Object
'     Set Produto = objUserControl.Controls("Produto")
'End Property
'
'Public Property Get PrecoTotal() As Object
'     Set PrecoTotal = objUserControl.Controls("PrecoTotal")
'End Property
'
'Public Property Get GridItens() As Object
'     Set GridItens = objUserControl.Controls("GridItens")
'End Property
'
'Public Property Get BotaoProdutos() As Object
'     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
'End Property
'
''Public Property Get Frame9() As Object
''     Set Frame9 = objUserControl.Controls("Frame9")
''End Property
'
'Public Property Get ValorFrete() As Object
'     Set ValorFrete = objUserControl.Controls("ValorFrete")
'End Property
'
'Public Property Get ValorDesconto() As Object
'     Set ValorDesconto = objUserControl.Controls("ValorDesconto")
'End Property
'
'Public Property Get ValorDespesas() As Object
'     Set ValorDespesas = objUserControl.Controls("ValorDespesas")
'End Property
'
'Public Property Get ValorSeguro() As Object
'     Set ValorSeguro = objUserControl.Controls("ValorSeguro")
'End Property
'
'Public Property Get ValorTotal() As Object
'     Set ValorTotal = objUserControl.Controls("ValorTotal")
'End Property
'
'Public Property Get LabelTotais() As Object
'     Set LabelTotais = objUserControl.Controls("LabelTotais")
'End Property
'
'Public Property Get IPIValor1() As Object
'     Set IPIValor1 = objUserControl.Controls("IPIValor1")
'End Property
'
'Public Property Get ValorProdutos() As Object
'     Set ValorProdutos = objUserControl.Controls("ValorProdutos")
'End Property
'
'Public Property Get ICMSBase1() As Object
'     Set ICMSBase1 = objUserControl.Controls("ICMSBase1")
'End Property
'
'Public Property Get ICMSValor1() As Object
'     Set ICMSValor1 = objUserControl.Controls("ICMSValor1")
'End Property
'
'Public Property Get ICMSSubstBase1() As Object
'     Set ICMSSubstBase1 = objUserControl.Controls("ICMSSubstBase1")
'End Property
'
'Public Property Get ICMSSubstValor1() As Object
'     Set ICMSSubstValor1 = objUserControl.Controls("ICMSSubstValor1")
'End Property
'
'Public Property Get BotaoEstoqueProd() As Object
'     Set BotaoEstoqueProd = objUserControl.Controls("BotaoEstoqueProd")
'End Property
'
'Public Property Get CTBSeqContraPartida() As Object
'     Set CTBSeqContraPartida = objUserControl.Controls("CTBSeqContraPartida")
'End Property
'
'Public Property Get CTBListHistoricos() As Object
'     Set CTBListHistoricos = objUserControl.Controls("CTBListHistoricos")
'End Property
'
'Public Property Get CTBBotaoImprimir() As Object
'     Set CTBBotaoImprimir = objUserControl.Controls("CTBBotaoImprimir")
'End Property
'
'Public Property Get CTBFrame7() As Object
'     Set CTBFrame7 = objUserControl.Controls("CTBFrame7")
'End Property
'
'Public Property Get CTBCclDescricao() As Object
'     Set CTBCclDescricao = objUserControl.Controls("CTBCclDescricao")
'End Property
'
'Public Property Get CTBContaDescricao() As Object
'     Set CTBContaDescricao = objUserControl.Controls("CTBContaDescricao")
'End Property
'
'Public Property Get CTBLabel7() As Object
'     Set CTBLabel7 = objUserControl.Controls("CTBLabel7")
'End Property
'
'Public Property Get CTBCclLabel() As Object
'     Set CTBCclLabel = objUserControl.Controls("CTBCclLabel")
'End Property
'
'Public Property Get CTBModelo() As Object
'     Set CTBModelo = objUserControl.Controls("CTBModelo")
'End Property
'
'Public Property Get CTBBotaoLimparGrid() As Object
'     Set CTBBotaoLimparGrid = objUserControl.Controls("CTBBotaoLimparGrid")
'End Property
'
'Public Property Get CTBBotaoModeloPadrao() As Object
'     Set CTBBotaoModeloPadrao = objUserControl.Controls("CTBBotaoModeloPadrao")
'End Property
'
'Public Property Get CTBHistorico() As Object
'     Set CTBHistorico = objUserControl.Controls("CTBHistorico")
'End Property
'
'Public Property Get CTBAglutina() As Object
'     Set CTBAglutina = objUserControl.Controls("CTBAglutina")
'End Property
'
'Public Property Get CTBLancAutomatico() As Object
'     Set CTBLancAutomatico = objUserControl.Controls("CTBLancAutomatico")
'End Property
'
'Public Property Get CTBConta() As Object
'     Set CTBConta = objUserControl.Controls("CTBConta")
'End Property
'
'Public Property Get CTBDebito() As Object
'     Set CTBDebito = objUserControl.Controls("CTBDebito")
'End Property
'
'Public Property Get CTBCredito() As Object
'     Set CTBCredito = objUserControl.Controls("CTBCredito")
'End Property
'
'Public Property Get CTBCcl() As Object
'     Set CTBCcl = objUserControl.Controls("CTBCcl")
'End Property
'
'Public Property Get CTBUpDown() As Object
'     Set CTBUpDown = objUserControl.Controls("CTBUpDown")
'End Property
'
'Public Property Get CTBDataContabil() As Object
'     Set CTBDataContabil = objUserControl.Controls("CTBDataContabil")
'End Property
'
'Public Property Get CTBLote() As Object
'     Set CTBLote = objUserControl.Controls("CTBLote")
'End Property
'
'Public Property Get CTBDocumento() As Object
'     Set CTBDocumento = objUserControl.Controls("CTBDocumento")
'End Property
'
'Public Property Get CTBTvwContas() As Object
'     Set CTBTvwContas = objUserControl.Controls("CTBTvwContas")
'End Property
'
'Public Property Get CTBGridContabil() As Object
'     Set CTBGridContabil = objUserControl.Controls("CTBGridContabil")
'End Property
'
'Public Property Get CTBTvwCcls() As Object
'     Set CTBTvwCcls = objUserControl.Controls("CTBTvwCcls")
'End Property
'
'Public Property Get CTBLabelLote() As Object
'     Set CTBLabelLote = objUserControl.Controls("CTBLabelLote")
'End Property
'
'Public Property Get CTBLabelDoc() As Object
'     Set CTBLabelDoc = objUserControl.Controls("CTBLabelDoc")
'End Property
'
'Public Property Get CTBLabel8() As Object
'     Set CTBLabel8 = objUserControl.Controls("CTBLabel8")
'End Property
'
'Public Property Get CTBTotalCredito() As Object
'     Set CTBTotalCredito = objUserControl.Controls("CTBTotalCredito")
'End Property
'
'Public Property Get CTBTotalDebito() As Object
'     Set CTBTotalDebito = objUserControl.Controls("CTBTotalDebito")
'End Property
'
'Public Property Get CTBLabelTotais() As Object
'     Set CTBLabelTotais = objUserControl.Controls("CTBLabelTotais")
'End Property
'
'Public Property Get CTBLabel1() As Object
'     Set CTBLabel1 = objUserControl.Controls("CTBLabel1")
'End Property
'
'Public Property Get CTBLabelCcl() As Object
'     Set CTBLabelCcl = objUserControl.Controls("CTBLabelCcl")
'End Property
'
'Public Property Get CTBLabelContas() As Object
'     Set CTBLabelContas = objUserControl.Controls("CTBLabelContas")
'End Property
'
'Public Property Get CTBLabelHistoricos() As Object
'     Set CTBLabelHistoricos = objUserControl.Controls("CTBLabelHistoricos")
'End Property
'
'Public Property Get CTBLabel5() As Object
'     Set CTBLabel5 = objUserControl.Controls("CTBLabel5")
'End Property
'
'Public Property Get CTBLabel13() As Object
'     Set CTBLabel13 = objUserControl.Controls("CTBLabel13")
'End Property
'
'Public Property Get CTBExercicio() As Object
'     Set CTBExercicio = objUserControl.Controls("CTBExercicio")
'End Property
'
'Public Property Get CTBPeriodo() As Object
'     Set CTBPeriodo = objUserControl.Controls("CTBPeriodo")
'End Property
'
'Public Property Get CTBLabel14() As Object
'     Set CTBLabel14 = objUserControl.Controls("CTBLabel14")
'End Property
'
'Public Property Get CTBOrigem() As Object
'     Set CTBOrigem = objUserControl.Controls("CTBOrigem")
'End Property
'
'Public Property Get CTBLabel21() As Object
'     Set CTBLabel21 = objUserControl.Controls("CTBLabel21")
'End Property
'
'Public Property Get FrameTributacao() As Object
'     Set FrameTributacao = objUserControl.Controls("FrameTributacao")
'End Property
'
'Public Property Get TributacaoRecalcular() As Object
'     Set TributacaoRecalcular = objUserControl.Controls("TributacaoRecalcular")
'End Property
'
''Public Property Get Frame12() As Object
''     Set Frame12 = objUserControl.Controls("Frame12")
''End Property
'
'Public Property Get IRAliquota() As Object
'     Set IRAliquota = objUserControl.Controls("IRAliquota")
'End Property
'
'Public Property Get ValorIRRF() As Object
'     Set ValorIRRF = objUserControl.Controls("ValorIRRF")
'End Property
'
'Public Property Get IRBase() As Object
'     Set IRBase = objUserControl.Controls("IRBase")
'End Property
'
''Public Property Get Frame11() As Object
''     Set Frame11 = objUserControl.Controls("Frame11")
''End Property
'
'Public Property Get IPIValor() As Object
'     Set IPIValor = objUserControl.Controls("IPIValor")
'End Property
'
'Public Property Get IPIBase() As Object
'     Set IPIBase = objUserControl.Controls("IPIBase")
'End Property
'
'Public Property Get IPICredito() As Object
'     Set IPICredito = objUserControl.Controls("IPICredito")
'End Property
'
''Public Property Get Frame10() As Object
''     Set Frame10 = objUserControl.Controls("Frame10")
''End Property
'
'Public Property Get ICMSSubstBase() As Object
'     Set ICMSSubstBase = objUserControl.Controls("ICMSSubstBase")
'End Property
'
'Public Property Get ICMSSubstValor() As Object
'     Set ICMSSubstValor = objUserControl.Controls("ICMSSubstValor")
'End Property
'
'Public Property Get ICMSBase() As Object
'     Set ICMSBase = objUserControl.Controls("ICMSBase")
'End Property
'
'Public Property Get ICMSValor() As Object
'     Set ICMSValor = objUserControl.Controls("ICMSValor")
'End Property
'
'Public Property Get ICMSCredito() As Object
'     Set ICMSCredito = objUserControl.Controls("ICMSCredito")
'End Property
'
''Public Property Get Frame13() As Object
''     Set Frame13 = objUserControl.Controls("Frame13")
''End Property
'
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
'Public Property Get TipoTributacao() As Object
'     Set TipoTributacao = objUserControl.Controls("TipoTributacao")
'End Property
'
'Public Property Get NatOpInternaEspelho() As Object
'     Set NatOpInternaEspelho = objUserControl.Controls("NatOpInternaEspelho")
'End Property
'
'Public Property Get DescTipoTrib() As Object
'     Set DescTipoTrib = objUserControl.Controls("DescTipoTrib")
'End Property
'
'Public Property Get LblTipoTrib() As Object
'     Set LblTipoTrib = objUserControl.Controls("LblTipoTrib")
'End Property
'
'Public Property Get DescNatOpInterna() As Object
'     Set DescNatOpInterna = objUserControl.Controls("DescNatOpInterna")
'End Property
'
'Public Property Get LblNatOpInternaEspelho() As Object
'     Set LblNatOpInternaEspelho = objUserControl.Controls("LblNatOpInternaEspelho")
'End Property
'
''Public Property Get Frame15() As Object
''     Set Frame15 = objUserControl.Controls("Frame15")
''End Property
'
'Public Property Get ComboICMSTipo() As Object
'     Set ComboICMSTipo = objUserControl.Controls("ComboICMSTipo")
'End Property
'
'Public Property Get Frame2() As Object
'     Set Frame2 = objUserControl.Controls("Frame2")
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
'Public Property Get ICMSCredita() As Object
'     Set ICMSCredita = objUserControl.Controls("ICMSCredita")
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
'Public Property Get IPIItemFrame() As Object
'     Set IPIItemFrame = objUserControl.Controls("IPIItemFrame")
'End Property
'
'Public Property Get ComboIPITipo() As Object
'     Set ComboIPITipo = objUserControl.Controls("ComboIPITipo")
'End Property
'
'Public Property Get IPICredita() As Object
'     Set IPICredita = objUserControl.Controls("IPICredita")
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
'Public Property Get NaturezaOpItem() As Object
'     Set NaturezaOpItem = objUserControl.Controls("NaturezaOpItem")
'End Property
'
'Public Property Get TipoTributacaoItem() As Object
'     Set TipoTributacaoItem = objUserControl.Controls("TipoTributacaoItem")
'End Property
'
'Public Property Get LblTipoTribItem() As Object
'     Set LblTipoTribItem = objUserControl.Controls("LblTipoTribItem")
'End Property
'
'Public Property Get NaturezaItemLabel() As Object
'     Set NaturezaItemLabel = objUserControl.Controls("NaturezaItemLabel")
'End Property
'
'Public Property Get LabelDescrNatOpItem() As Object
'     Set LabelDescrNatOpItem = objUserControl.Controls("LabelDescrNatOpItem")
'End Property
'
'Public Property Get DescTipoTribItem() As Object
'     Set DescTipoTribItem = objUserControl.Controls("DescTipoTribItem")
'End Property
'
'Public Property Get TribSobreItem() As Object
'     Set TribSobreItem = objUserControl.Controls("TribSobreItem")
'End Property
'
'Public Property Get TribSobreFrete() As Object
'     Set TribSobreFrete = objUserControl.Controls("TribSobreFrete")
'End Property
'
'Public Property Get TribSobreDesconto() As Object
'     Set TribSobreDesconto = objUserControl.Controls("TribSobreDesconto")
'End Property
'
'Public Property Get TribSobreSeguro() As Object
'     Set TribSobreSeguro = objUserControl.Controls("TribSobreSeguro")
'End Property
'
'Public Property Get TribSobreOutrasDesp() As Object
'     Set TribSobreOutrasDesp = objUserControl.Controls("TribSobreOutrasDesp")
'End Property
'
'Public Property Get FrameOutrosTrib() As Object
'     Set FrameOutrosTrib = objUserControl.Controls("FrameOutrosTrib")
'End Property
'
'Public Property Get LabelValorFrete() As Object
'     Set LabelValorFrete = objUserControl.Controls("LabelValorFrete")
'End Property
'
'Public Property Get LabelValorDesconto() As Object
'     Set LabelValorDesconto = objUserControl.Controls("LabelValorDesconto")
'End Property
'
'Public Property Get LabelValorSeguro() As Object
'     Set LabelValorSeguro = objUserControl.Controls("LabelValorSeguro")
'End Property
'
'Public Property Get LabelValorOutrasDespesas() As Object
'     Set LabelValorOutrasDespesas = objUserControl.Controls("LabelValorOutrasDespesas")
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
'Public Property Get LabelValorItem() As Object
'     Set LabelValorItem = objUserControl.Controls("LabelValorItem")
'End Property
'
'Public Property Get LabelQtdeItem() As Object
'     Set LabelQtdeItem = objUserControl.Controls("LabelQtdeItem")
'End Property
'
'Public Property Get LabelUMItem() As Object
'     Set LabelUMItem = objUserControl.Controls("LabelUMItem")
'End Property
'
'Public Property Get OpcaoTributacao() As Object
'     Set OpcaoTributacao = objUserControl.Controls("OpcaoTributacao")
'End Property
'
'Public Property Get BotaoGravar() As Object
'     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
'End Property
'
'Public Property Get BotaoLimpar() As Object
'     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
'End Property
'
'Public Property Get BotaoFechar() As Object
'     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
'End Property
'
'Public Property Get TabStrip1() As Object
'     Set TabStrip1 = objUserControl.Controls("TabStrip1")
'End Property
'
'Public Property Get Status() As Object
'    Set Status = objUserControl.Controls("Status")
'End Property
'
'Public Property Get LabelFilial() As Object
'    Set LabelFilial = objUserControl.Controls("LabelFilial")
'End Property
'
'Public Property Get LabelFilialEntrega() As Object
'    Set LabelFilialEntrega = objUserControl.Controls("LabelFilialEntrega")
'End Property
'
'Public Property Get FilialL() As Object
'    Set FilialL = objUserControl.Controls("FilialL")
'End Property
'
'Public Sub ImprimeRomaneio_Click()
'    iAlterado = REGISTRO_ALTERADO
'End Sub
'
'Public Property Get BotaoExcluir() As Object
'     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
'End Property
'
''Rastreamento
'Public Property Get GridRastro() As Object
'     Set GridRastro = objUserControl.Controls("GridRastro")
'End Property
'
''Rastreamento
'Public Property Get EscaninhoRastro() As Object
'     Set EscaninhoRastro = objUserControl.Controls("EscaninhoRastro")
'End Property
'
'Public Property Get ValorMercadoria() As Object
'     Set ValorMercadoria = objUserControl.Controls("ValorMercadoria")
'End Property
'
'Public Property Get ValorContainer() As Object
'     Set ValorContainer = objUserControl.Controls("ValorContainer")
'End Property
'
'Public Property Get Origem() As Object
'     Set Origem = objUserControl.Controls("Origem")
'End Property
'
'Public Property Get Destino() As Object
'     Set Destino = objUserControl.Controls("Destino")
'End Property
'
''Public Property Get BotaoRastreamento() As Object
''    Set BotaoRastreamento = objUserControl.Controls("BotaoRastreamento")
''End Property
'
''--- fim dos properties get dos controles da tela
'
'Public Sub ValorMercadoria_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ValorContainer_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Origem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Destino_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub BotaoDataReferenciaUp_Click()
'
'Dim lErro As Long
'Dim sData As String
'Dim bCancel As Boolean
'
'On Error GoTo Erro_BotaoDataReferenciaUp_Click
'
'    sData = DataReferencia.Text
'
'    'aumenta a data em um dia
'    lErro = Data_Up_Down_Click(DataReferencia, AUMENTA_DATA)
'    If lErro <> SUCESSO Then gError 26716
'
'    Call DataReferencia_Validate(bCancel)
'
'    If bCancel = True Then DataReferencia.Text = sData
'
'    Exit Sub
'
'Erro_BotaoDataReferenciaUp_Click:
'
'    Select Case gErr
'
'        Case 26716
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
''
''Public Sub BotaoImprimir_Click()
''
''Dim objNFiscal As New ClassNFiscal
''Dim lErro As Long
''Dim objRelatorio As New AdmRelatorio
''
''On Error GoTo Erro_BotaoImprimir_Click
''
''    'Verifica se os campos obrigatórios estão preenchidos
''    If Len(Trim(Serie.Text)) = 0 Then gError 57653
''    If Len(Trim(NFiscal.ClipText)) = 0 Then gError 57654
''    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 57655
''    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 57656
''
''    objNFiscal.sSerie = Serie.Text
''    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
''    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
''    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
''    objNFiscal.iFilialEmpresa = giFilialEmpresa
''
''    'Tenta ler a nota Fiscal passada por parametro
''    lErro = NFiscal_Interna_TestaExistencia(objNFiscal)
''    If lErro <> SUCESSO And lErro <> 57677 Then gError 57659
''    If lErro = 57677 Then gError 57660
''
''    lErro = objRelatorio.ExecutarDireto("Emissão das Notas Fiscais", "NotaFiscal = @NNFISCALINIC E Serie = @TSERIE", 1, "NFFatS", "TSERIE", objNFiscal.sSerie, "NNFISCALINIC", CStr(objNFiscal.lNumNotaFiscal), "NNFISCALFIM", CStr(objNFiscal.lNumNotaFiscal), "NENTRADASAIDA", "1", "NNFTIPO", "1")
''    If lErro <> SUCESSO Then gError 57661
''
''    Exit Sub
''
''Erro_BotaoImprimir_Click:
''
''    Select Case gErr
''
''        Case 57653
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
''
''        Case 57654
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
''
''        Case 57655
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
''
''        Case 57656
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
''
''        Case 57659, 57661
''
''        Case 57660
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", gErr, objNFiscal.lNumNotaFiscal)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
''
''    End Select
''
''    Exit Sub
''
''End Sub
'
'Public Sub BotaoDataReferenciaDown_Click()
'
'Dim lErro As Long
'Dim bCancel As Boolean
'Dim sData As String
'
'On Error GoTo Erro_BotaoDataReferenciaDown_Click
'
'    sData = DataReferencia.Text
'
'    'diminui a data em um dia
'    lErro = Data_Up_Down_Click(DataReferencia, DIMINUI_DATA)
'    If lErro <> SUCESSO Then gError 26715
'
'    Call DataReferencia_Validate(bCancel)
'
'    If bCancel = True Then DataReferencia.Text = sData
'
'    Exit Sub
'
'Erro_BotaoDataReferenciaDown_Click:
'
'    Select Case gErr
'
'        Case 26715
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub BotaoLimparNF_Click()
'
'    NFiscal.Caption = ""
'    'Limpa o Frame de Rastreamento
'    Call gobjRastreamento.Limpa_Tela_Rastreamento
'
'End Sub
'
'Public Sub DataEmissao_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
'
'End Sub
'
'Public Sub DataSaida_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(DataSaida, iAlterado)
'
'End Sub
'
'Public Sub Destinatario_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Emitente_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub GridComissoes_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    Call Grid_Trata_Tecla1(KeyCode, objGridComissoes)
'    'Se alguma comissao foi excluída
'    If KeyCode = vbKeyDelete Then
'        'atualiza os painéis totalizadores
'        Call objTabComissoes.Soma_Percentual
'        Call objTabComissoes.Soma_Valor
'    End If
'
'End Sub
'
'Public Sub ImageStatus_Click(Index As Integer)
'
'End Sub
'
'Public Sub MensagemLabel_Click()
'
'Dim objMensagem As New ClassMensagem
'Dim colSelecao As New Collection
'
'    'Coloca a mensagem da tela no objMensagem
'    objMensagem.sDescricao = Mensagem.Text
'
'    'Chama a tela de browse de Mensagem
'    Call Chama_Tela("MensagemLista", colSelecao, objMensagem, objEventoMensagem)
'
'    Exit Sub
'
'
'End Sub
'
'Public Sub NatOpInterna_GotFocus()
'
'    Call gobjTribTab.NatOpInterna_GotFocus(iAlterado)
'
'End Sub
'
'Public Sub NaturezaItemLabel_Click()
'
'    Call gobjTribTab.NaturezaItemLabel_Click
'
'End Sub
'
'Public Sub NaturezaOpItem_GotFocus()
'
'    Call gobjTribTab.NaturezaOpItem_GotFocus(iAlterado)
'
'End Sub
'
'Public Sub NFiscal_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(NFiscal, iAlterado)
'
'End Sub
'
'Public Sub NFiscalOriginal_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(NFiscalOriginal, iAlterado)
'
'End Sub
'
''Public Sub BotaoRastreamento_Click()
''
''Dim objNFiscal As New ClassNFiscal
''Dim objFornecedor As New ClassFornecedor
''Dim objCliente As New ClassCliente
''Dim objTipoDocInfo As New ClassTipoDocInfo
''Dim lErro As Long
''
''On Error GoTo Erro_BotaoRastreamento_Click
''
''    objNFiscal.sSerie = Serie.Text
''    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
''
''    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
''
''    'Lê o Tipo da NF
''    lErro = TipoDocInfo_Le_Codigo(objTipoDocInfo)
''    If lErro <> SUCESSO And lErro <> 31415 Then gError 62050
''
''    'Se não encontrou o Tipo de Documento --> erro
''    If lErro = 31415 Then gError 62053
''
''    objNFiscal.iTipoNFiscal = objTipoDocInfo.iCodigo
''
''    'de acordo com a sua Origem verifica se o Cliente ou Fornecedor estão preenchidos
''    If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Then
''
''        objCliente.sNomeReduzido = Cliente.Text
''        lErro = Cliente_Le_NomeReduzido(objCliente)
''        If lErro <> SUCESSO And lErro <> 12348 Then gError 62051
''
''        'Não encontrou o cliente
''        If lErro = 12348 Then gError 62052
''
''        objNFiscal.lCliente = objCliente.lCodigo
''        objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
''
''    End If
''
''    'Chama a tela de atualização de rastreamento
''    Call Chama_Tela("RastroItensNFFAT", objNFiscal)
''
''    Exit Sub
''
''Erro_BotaoRastreamento_Click:
''
''    Select Case gErr
''
''        Case 62050, 62051
''
''        Case 62052
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
''
''        Case 62053
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
''
''    End Select
''
''    Exit Sub
''
''End Sub
'
'Public Sub PercAcrescFin_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iPercAcresFinAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PercAcrescFin_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_PercAcrescFin_Validate
'
'    'Verific se o percentual foi alterado
'    If iPercAcresFinAlterado = 0 Then Exit Sub
'    'Se estiver preenchido
'    If Len(Trim(PercAcrescFin.Text)) > 0 Then
'        lErro = Porcentagem_Critica_Negativa(PercAcrescFin)
'        If lErro <> SUCESSO Then gError 46190
'    End If
'
'    If Len(Trim(TabelaPreco.Text)) > 0 Then
'
'        lErro = Trata_TabelaPreco()
'        If lErro <> SUCESSO Then gError 51620
'
'    End If
'
'    iPercAcresFinAlterado = 0
'
'    Exit Sub
'
'Erro_PercAcrescFin_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 46190
'
'        Case 51620
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub TipoTributacao_GotFocus()
'
'    Call gobjTribTab.TipoTributacao_GotFocus(iAlterado)
'
'End Sub
'
'Public Sub TipoTributacaoItem_GotFocus()
'
'    Call gobjTribTab.TipoTributacaoItem_GotFocus(iAlterado)
'
'End Sub
'
'Public Sub ValorIRRF_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.ValorIRRF_Validate(Cancel)
'
'End Sub
'
'Public Sub ComboICMSTipo_Click()
'
'    Call gobjTribTab.ComboICMSTipo_Click
'
'End Sub
'
'Public Sub ComboIPITipo_Click()
'
'    Call gobjTribTab.ComboIPITipo_Click
'
'End Sub
'
'Public Sub ComboItensTrib_Click()
'
'    Call gobjTribTab.ComboItensTrib_Click
'
'End Sub
'
'Public Sub LblNatOpInterna_Click()
'
'    Call gobjTribTab.LblNatOpInterna_Click(NF_SAIDA)
'
'End Sub
'
'Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
'
'    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
'
'End Sub
'
'Public Sub Form_Unload(Cancel As Integer)
'
'Dim lErro As Long
'
'    ' Liberar as variáveis globais
'    Set objGridParcelas = Nothing
'    Set objGridItens = Nothing
'    Set objGridComissoes = Nothing
'    Set objGridAlocacoes = Nothing
'
'    Set objEventoMensagem = Nothing
'    Set objEventoCliente = Nothing
'    Set objEventoCondPagto = Nothing
'    Set objEventoEstoqueProd = Nothing
'    Set objEventoNFiscal = Nothing
'    Set objEventoNFiscalOriginal = Nothing
'    Set objEventoProduto = Nothing
'    Set objEventoSerie = Nothing
'    Set objEventoSerieOriginal = Nothing
'    Set objEventoTransportadora = Nothing
'    Set objEventoVendedor = Nothing
'
'    Set objGrid1 = Nothing
'    Set objContabil = Nothing
'
'    'Eventos associados a contabilidade
'    Set objEventoLote = Nothing
'    Set objEventoDoc = Nothing
'
'    Set gobjRastreamento = Nothing
'
'    'Fecha o comando de seta se estiver aberto
'    lErro = ComandoSeta_Liberar(Me.Name)
'
'    Set objTabComissoes = Nothing
'    Set gobjTribTab = Nothing
'
'    Set objUserControl = Nothing
'
'End Sub
'
'Public Sub BotaoLotes_Click()
''Chama a tela de Lote de Rastreamento
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoLotes_Click
'
'    Call gobjRastreamento.BotaoLotes_Click
'
'    Exit Sub
'
'Erro_BotaoLotes_Click:
'
'    Select Case gErr
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ItemNFRastro_Change()
''Rastreamento
'
'    Call gobjRastreamento.ItemNFRastro_Change
'
'End Sub
'
'Public Sub ItemNFRastro_GotFocus()
''Rastreamento
'
'    Call gobjRastreamento.ItemNFRastro_GotFocus
'
'End Sub
'
'Public Sub ItemNFRastro_KeyPress(KeyAscii As Integer)
''Rastreamento
'
'    Call gobjRastreamento.ItemNFRastro_KeyPress(KeyAscii)
'
'End Sub
'
'Public Sub ItemNFRastro_Validate(Cancel As Boolean)
''Rastreamento
'
'    Call gobjRastreamento.ItemNFRastro_Validate(Cancel)
'
'End Sub
'
'Public Sub AlmoxRastro_Change()
''Rastreamento
'
'    Call gobjRastreamento.AlmoxRastro_Change
'
'End Sub
'
'Public Sub AlmoxRastro_GotFocus()
''Rastreamento
'
'    Call gobjRastreamento.AlmoxRastro_GotFocus
'
'End Sub
'
'Public Sub AlmoxRastro_KeyPress(KeyAscii As Integer)
''Rastreamento
'
'    Call gobjRastreamento.AlmoxRastro_KeyPress(KeyAscii)
'
'End Sub
'
'Public Sub AlmoxRastro_Validate(Cancel As Boolean)
''Rastreamento
'
'    Call gobjRastreamento.AlmoxRastro_Validate(Cancel)
'
'End Sub
'
'Public Sub EscaninhoRastro_Change()
''Rastreamento
'
'    Call gobjRastreamento.EscaninhoRastro_Change
'
'End Sub
'
'Public Sub EscaninhoRastro_Click()
''Rastreamento
'
'    Call gobjRastreamento.EscaninhoRastro_Click
'
'End Sub
'
'
'Public Sub EscaninhoRastro_GotFocus()
''Rastreamento
'
'    Call gobjRastreamento.EscaninhoRastro_GotFocus
'
'End Sub
'
'Public Sub EscaninhoRastro_KeyPress(KeyAscii As Integer)
''Rastreamento
'
'    Call gobjRastreamento.EscaninhoRastro_KeyPress(KeyAscii)
'
'End Sub
'
'Public Sub EscaninhoRastro_Validate(Cancel As Boolean)
''Rastreamento
'
'    Call gobjRastreamento.EscaninhoRastro_Validate(Cancel)
'
'End Sub
'
'Public Sub LoteRastro_Change()
''Rastreamento
'
'    Call gobjRastreamento.LoteRastro_Change
'
'End Sub
'
'Public Sub LoteRastro_GotFocus()
''Rastreamento
'
'    Call gobjRastreamento.LoteRastro_GotFocus
'
'End Sub
'
'Public Sub LoteRastro_KeyPress(KeyAscii As Integer)
''Rastreamento
'
'    Call gobjRastreamento.LoteRastro_KeyPress(KeyAscii)
'
'End Sub
'
'Public Sub LoteRastro_Validate(Cancel As Boolean)
''Rastreamento
'
'    Call gobjRastreamento.LoteRastro_Validate(Cancel)
'
'End Sub
'
'Public Sub FilialOPRastro_Change()
''Rastreamento
'
'    Call gobjRastreamento.FilialOPRastro_Change
'
'End Sub
'
'Public Sub FilialOPRastro_GotFocus()
''Rastreamento
'
'    Call gobjRastreamento.FilialOPRastro_GotFocus
'
'End Sub
'
'Public Sub FilialOPRastro_KeyPress(KeyAscii As Integer)
''Rastreamento
'
'    Call gobjRastreamento.FilialOPRastro_KeyPress(KeyAscii)
'
'End Sub
'
'Public Sub FilialOPRastro_Validate(Cancel As Boolean)
''Rastreamento
'
'    Call gobjRastreamento.FilialOPRastro_Validate(Cancel)
'
'End Sub
'
'Public Sub QuantLoteRastro_Change()
''Rastreamento
'
'    Call gobjRastreamento.QuantLoteRastro_Change
'
'End Sub
'
'Public Sub QuantLoteRastro_GotFocus()
''Rastreamento
'
'    Call gobjRastreamento.QuantLoteRastro_GotFocus
'
'End Sub
'
'Public Sub QuantLoteRastro_KeyPress(KeyAscii As Integer)
''Rastreamento
'
'    Call gobjRastreamento.QuantLoteRastro_KeyPress(KeyAscii)
'
'End Sub
'
'Public Sub QuantLoteRastro_Validate(Cancel As Boolean)
''Rastreamento
'
'    Call gobjRastreamento.QuantLoteRastro_Validate(Cancel)
'
'End Sub
'
'Public Sub LblTipoTrib_Click()
'
'    Call gobjTribTab.LblTipoTrib_Click
'
'End Sub
'
'Public Sub LblTipoTribItem_Click()
'
'    Call gobjTribTab.LblTipoTribItem_Click
'
'End Sub
'
'Public Sub NaturezaOpItem_Change()
'
'    Call gobjTribTab.NaturezaOpItem_Change
'
'End Sub
'
'Public Sub NaturezaOpItem_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.NaturezaOpItem_Validate(Cancel)
'
'End Sub
'
'Public Sub NatOpInterna_Change()
'
'    Call gobjTribTab.NatOpInterna_Change
'
'    NatOpInternaEspelho.Caption = NatOpInterna.Text
'
'End Sub
'
'Public Sub NatOpInterna_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.NatOpInterna_Validate(Cancel)
'
'End Sub
'
'Public Sub TipoTributacao_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'    Call gobjTribTab.TipoTributacao_Change
'
'End Sub
'
'Public Sub TipoTributacao_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.TipoTributacao_Validate(Cancel)
'
'End Sub
'
'Public Sub TipoTributacaoItem_Change()
'
'    Call gobjTribTab.TipoTributacaoItem_Change
'
'End Sub
'
'Public Sub TipoTributacaoItem_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.TipoTributacaoItem_Validate(Cancel)
'
'End Sub
'
'Public Sub TribSobreDesconto_Click()
'
'    Call gobjTribTab.TribSobreDesconto_Click
'
'End Sub
'
'Public Sub TribSobreFrete_Click()
'
'    Call gobjTribTab.TribSobreFrete_Click
'
'End Sub
'
'Public Sub TribSobreItem_Click()
'
'    Call gobjTribTab.TribSobreItem_Click
'
'End Sub
'
'Public Sub TribSobreOutrasDesp_Click()
'
'    Call gobjTribTab.TribSobreOutrasDesp_Click
'
'End Sub
'
'Public Sub TribSobreSeguro_Click()
'
'    Call gobjTribTab.TribSobreSeguro_Click
'
'End Sub
'
'Public Sub TributacaoRecalcular_Click()
'
'    Call gobjTribTab.TributacaoRecalcular_Click
'
'End Sub
'
'Public Sub OpcaoTributacao_Click()
'
'    Call gobjTribTab.OpcaoTributacao_Click
'
'End Sub
'
'Public Sub ValorIRRF_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iValorIRRFAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.ValorIRRF_Change
'
'End Sub
'
'Public Sub ICMSAliquotaItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.ICMSAliquotaItem_Change
'
'End Sub
'
'Public Sub ICMSAliquotaItem_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.ICMSAliquotaItem_Validate(Cancel)
'
'End Sub
'
'Public Sub ICMSBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.ICMSBaseItem_Change
'
'End Sub
'
'Public Sub ICMSBaseItem_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.ICMSBaseItem_Validate(Cancel)
'
'End Sub
'
'Public Sub ICMSPercRedBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.ICMSPercRedBaseItem_Change
'
'End Sub
'
'Public Sub ICMSPercRedBaseItem_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.ICMSPercRedBaseItem_Validate(Cancel)
'
'End Sub
'
'Public Sub ICMSSubstAliquotaItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.ICMSSubstAliquotaItem_Change
'
'End Sub
'
'Public Sub ICMSSubstAliquotaItem_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.ICMSSubstAliquotaItem_Validate(Cancel)
'
'End Sub
'
'Public Sub ICMSSubstBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.ICMSSubstBaseItem_Change
'
'End Sub
'
'Public Sub ICMSSubstBaseItem_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.ICMSSubstBaseItem_Validate(Cancel)
'
'End Sub
'
'Public Sub ICMSSubstValorItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.ICMSSubstValorItem_Change
'
'End Sub
'
'Public Sub ICMSSubstValorItem_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.ICMSSubstValorItem_Validate(Cancel)
'
'End Sub
'
'Public Sub ICMSValorItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.ICMSValorItem_Change
'
'End Sub
'
'Public Sub ICMSValorItem_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.ICMSValorItem_Validate(Cancel)
'
'End Sub
'
'Public Sub IPIAliquotaItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.IPIAliquotaItem_Change
'
'End Sub
'
'Public Sub IPIAliquotaItem_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.IPIAliquotaItem_Validate(Cancel)
'
'End Sub
'
'Public Sub IPIBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.IPIBaseItem_Change
'
'End Sub
'
'Public Sub IPIBaseItem_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.IPIBaseItem_Validate(Cancel)
'
'End Sub
'
'Public Sub IPIPercRedBaseItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.IPIPercRedBaseItem_Change
'
'End Sub
'
'Public Sub IPIPercRedBaseItem_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.IPIPercRedBaseItem_Validate(Cancel)
'
'End Sub
'
'Public Sub IPIValorItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.IPIValorItem_Change
'
'End Sub
'
'Public Sub IPIValorItem_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.IPIValorItem_Validate(Cancel)
'
'End Sub
'
'Public Sub IRAliquota_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.IRAliquota_Change
'
'End Sub
'
'Public Sub IRAliquota_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.IRAliquota_Validate(Cancel)
'
'End Sub
'
'Public Sub ISSAliquota_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.ISSAliquota_Change
'
'End Sub
'
'Public Sub ISSAliquota_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.ISSAliquota_Validate(Cancel)
'
'End Sub
'
'Public Sub ISSIncluso_Click()
'
'    Call gobjTribTab.ISSIncluso_Click
'
'End Sub
'
'Public Sub ISSValor_Change()
'
'    Call gobjTribTab.ISSValor_Change
'
'End Sub
'
'Public Sub ISSValor_Validate(Cancel As Boolean)
'
'    Call gobjTribTab.ISSValor_Validate(Cancel)
'
'End Sub
''*** fim tributacao
'
'Public Sub BotaoFechar_Click()
'
'    Unload Me
'
'End Sub
'
'Public Sub CanalVenda_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub CanalVenda_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Cliente_Change()
'
'Static sNomeReduzidoParte As String '*** rotina para trazer cliente
'Dim lErro As Long
'
'On Error GoTo Erro_Cliente_Change
'
'        iAlterado = REGISTRO_ALTERADO
'        iClienteAlterado = REGISTRO_ALTERADO
'
'        'rotina para trazer o nome do cliente com uma parte dos caracteres digitados
'        lErro = Cliente_Pesquisa_NomeReduzido(Cliente, sNomeReduzidoParte)
'        If lErro <> SUCESSO Then gError 71653
'
'    Exit Sub
'
'Erro_Cliente_Change:
'
'    Select Case gErr
'
'        Case 71653
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub CobrancaAutomatica_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'    If CobrancaAutomatica.Value = vbChecked And Len(Trim(CondicaoPagamento.Text)) > 0 And CobrancaAutomatica.Tag <> GRID_CHECKBOX_ATIVO Then
'        Call Cobranca_Automatica
'    Else
'        CobrancaAutomatica.Tag = ""
'    End If
'
'End Sub
'
'Public Sub ComissaoAutomatica_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_ComissaoAutomatica_Click
'
'    iAlterado = REGISTRO_ALTERADO
'
'    'Se a comissão automática estiver selecionada
'    If ComissaoAutomatica.Value = vbChecked Then
'        'Recalcula as comissoes
'        lErro = objTabComissoes.Comissoes_Automatica_Exibe()
'        If lErro <> SUCESSO Then gError 51616
'
'    End If
'
'    Exit Sub
'
'Erro_ComissaoAutomatica_Click:
'
'    Select Case gErr
'
'        Case 51616
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'End Sub
'
'Public Sub CondicaoPagamento_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub DataEmissao_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub DataSaida_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub DataVencimento_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub DataVencimento_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub DataVencimento_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub DataVencimento_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = DataVencimento
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridItens)
'
'End Sub
'
'Public Sub Desconto_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
'
'End Sub
'
'Public Sub Desconto_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridItens.objControle = Desconto
'    lErro = Grid_Campo_Libera_Foco(objGridItens)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'
'Public Sub Desconto1Ate_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto1Ate_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Ate_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Ate_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto1Ate
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto1Codigo_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Desconto1Codigo_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Codigo_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Codigo_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto1Codigo
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto1Percentual_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto1Percentual_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Percentual_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Percentual_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto1Percentual
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto1Valor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto1Valor_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Valor_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto1Valor_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto1Valor
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto2Ate_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto2Ate_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Ate_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Ate_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto2Ate
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto2Codigo_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Desconto2Codigo_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Codigo_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Codigo_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto2Codigo
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto2Percentual_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto2Percentual_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Percentual_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Percentual_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto2Percentual
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto2Valor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto2Valor_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Valor_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto2Valor_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto2Valor
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto3Ate_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto3Ate_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Ate_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Ate_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto3Ate
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto3Codigo_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto3Codigo_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Codigo_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Codigo_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto3Codigo
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto3Percentual_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto3Percentual_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Percentual_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Percentual_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto3Percentual
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Desconto3Valor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Desconto3Valor_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Valor_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Desconto3Valor_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Desconto3Valor
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub DescricaoItem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub DescricaoItem_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridItens)
'
'End Sub
'
'Public Sub DescricaoItem_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
'
'End Sub
'
'Public Sub DescricaoItem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridItens.objControle = DescricaoItem
'    lErro = Grid_Campo_Libera_Foco(objGridItens)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Filial_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub FilialEntrega_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giFilialEntregaAlterada = 1
'
'End Sub
'
'Public Sub FilialEntrega_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Function CargaPosFormLoad(Optional bTrazendoDoc As Boolean = False) As Long
'
'Dim iSubtipo As Integer
'Dim lErro As Long
'
'On Error GoTo Erro_CargaPosFormLoad
'
'    If (giPosCargaOk = 0) Then
'
'        'Rastreamento
'        Set gobjRastreamento = New ClassRastreamento
'        Set gobjRastreamento.objTela = Me
'        gobjRastreamento.bSaidaMaterial = True
'
'        'p/permitir o redesenho da tela
'        DoEvents
'
'        gbCarregandoTela = True
'
'        ''*** incluidos p/tratamento de tributacao *******************************
'
'        Call GridItens_InicializaColunas
'        Set gobjTribTab = New ClassTribTab
'        lErro = gobjTribTab.Ativar(Me)
'        If lErro <> SUCESSO Then gError 39030
'
'        '*** fim tributacao
'
'        'Carrega os Tipos de Notas Fiscais
'        lErro = Carrega_TiposDocInfo()
'        If lErro <> SUCESSO Then gError 42111
'
'        'Carrega as Tabelas de Preço
'        lErro = Carrega_TabelaPreco()
'        If lErro <> SUCESSO Then gError 42112
'
'        'Carrega as Transportadoras
'        lErro = Carrega_Transportadora()
'        If lErro <> SUCESSO Then gError 42113
'
'        'Carrega os Canais de Venda
'        lErro = Carrega_CanalVenda()
'        If lErro <> SUCESSO Then gError 42114
'
'        'Carrega as Séries
'        lErro = Carrega_Serie()
'        If lErro <> SUCESSO Then gError 42115
'
'        'Seleciona série padrão
'        lErro = Seleciona_Serie_Padrao()
'        If lErro <> SUCESSO Then gError 500137
'
'        'Carrega na combo as Condições de Pagamento
'        lErro = Carrega_CondicaoPagamento()
'        If lErro <> SUCESSO Then gError 42123
'
'        'Carrega na combo de Banco Cobradores
'        lErro = Carrega_Cobradores()
'        If lErro <> SUCESSO Then gError 95181
'
'        'Carrega os Estados
'        lErro = Carrega_PlacaUF()
'        If lErro <> SUCESSO Then gError 42116
'
'        'Carrega os Tipos de Desconto
'        lErro = Carrega_TipoDesconto()
'        If lErro <> SUCESSO Then gError 42125
'
'        PrecoUnitario.Format = gobjFAT.sFormatoPrecoUnitario
'
'        Quantidade.Format = FORMATO_ESTOQUE
'        QuantAlocada.Format = FORMATO_ESTOQUE
'        QuantVendida.Format = FORMATO_ESTOQUE
'
'        'Seta as Variáveis de Grid
'        Set objGridItens = New AdmGrid
'        Set objGridComissoes = New AdmGrid
'        Set objGridParcelas = New AdmGrid
'        Set objGridAlocacoes = New AdmGrid
'
'        'Seta as Variáveis das Telas de browse
'        Set objEventoMensagem = New AdmEvento
'        Set objEventoCliente = New AdmEvento
'        Set objEventoCondPagto = New AdmEvento
'        Set objEventoEstoqueProd = New AdmEvento
'        Set objEventoNFiscal = New AdmEvento
'        Set objEventoNFiscalOriginal = New AdmEvento
'        Set objEventoProduto = New AdmEvento
'        Set objEventoSerie = New AdmEvento
'        Set objEventoSerieOriginal = New AdmEvento
'        Set objEventoTransportadora = New AdmEvento
'        Set objEventoVendedor = New AdmEvento
'        giDataReferenciaAlterada = 0
'
'        'Inicializa o Grid de Ítens
'        lErro = Inicializa_Grid_Itens(objGridItens)
'        If lErro <> SUCESSO Then gError 42126
'
'        'Seleciona a única tabela de preço. Importante: GridItens já deve estar definido
'        If giTipoVersao = VERSAO_LIGHT And TabelaPreco.ListCount > 0 Then
'            TabelaPreco.ListIndex = 0
'        End If
'
'        'Inicializa a máscara do Produto
'        lErro = Inicializa_Mascara_Produto_MaskEd(Produto)
'        If lErro <> SUCESSO Then gError 42127
'
'        'Inicializa o Grid de Parcelas
'        lErro = Inicializa_Grid_Parcelas(objGridParcelas)
'        If lErro <> SUCESSO Then gError 42128
'
'        'Inicializa o Grid de Comissões
'
'        lErro = objTabComissoes.Inicializa_Grid_Comissoes(objGridComissoes)
'        If lErro <> SUCESSO Then gError 42129
'
'        'Inicializa o Grid de Alocações
'        lErro = Inicializa_Grid_Alocacoes(objGridAlocacoes)
'        If lErro <> SUCESSO Then gError 42130
'
'        'Inicializa a contabilidade carregando apenas os modelos contábeis referentes ao modelo padrão
'        If objContabil.objGridInt Is Nothing Then
'
'            'Obtém o código do subtipo contábil referente ao TipoDocInfo default
'            lErro = Obtem_Subtipo_Contabil("TiposDocInfo", DOCINFO_NFISFV, iSubtipo)
'            If lErro <> SUCESSO And lErro <> 79671 Then gError 79682
'
'            'Se não encontrou o subtipo => erro
'            If lErro = 79671 Then gError 79683
'
'            'Inicialização da parte de contabilidade
'            lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_FATURAMENTO, iSubtipo)
'            If lErro <> SUCESSO Then gError 39672
'
'        End If
'
'        gbCarregandoTela = False
'
'        iAlterado = 0
'
'        giPosCargaOk = 1
'
'    End If
'
'    CargaPosFormLoad = SUCESSO
'
'    Exit Function
'
'Erro_CargaPosFormLoad:
'
'    gbCarregandoTela = False
'
'    CargaPosFormLoad = gErr
'
'    Select Case gErr
'
'        Case 39030, 39672, 42111, 42112, 42113, 42114, 42123, 42125, 42126, 42127, 42128, _
'             42129, 42130, 42115, 42116, 79682, 83383, 500137, 95181
'
'        Case 79683
'            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPODOCINFO_NAO_ENCONTRADO", gErr, DOCINFO_NFISFV)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub Form_Load()
'
'Dim lErro As Long
'
'On Error GoTo Erro_Form_Load
'
'    Set objTabComissoes.objTela = Me
'    iFrameAtual = 1
'
'    'Preenche Data Referencia e Data de Emissão coma Data Atual
'    DataReferencia.PromptInclude = False
'    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
'    DataReferencia.PromptInclude = True
'    DataSaida.PromptInclude = False
'    DataSaida.Text = Format(gdtDataAtual, "dd/mm/yy")
'    DataSaida.PromptInclude = True
'    DataEmissao.PromptInclude = False
'    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
'    DataEmissao.PromptInclude = True
'
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        Filial.Visible = False
'        FilialEntrega.Visible = False
'        FilialL.Visible = False
'        LabelFilial.Visible = False
'        LabelFilialEntrega.Visible = False
'
'    End If
'
'    Call Inicializa_Grid_ComprovServ(objGridComprovServ)
'
'    iValorContainerAlterado = 0
'
'    lErro_Chama_Tela = SUCESSO
'
'    Exit Sub
'
'Erro_Form_Load:
'
'    lErro_Chama_Tela = gErr
'
'    Select Case gErr
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    iAlterado = 0
'
'    Exit Sub
'
'
'End Sub
'
'Public Sub Mensagem_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub NFiscal_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub NFiscalOriginal_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub NumPedidoTerc_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PercentDesc_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PercentDesc_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridItens)
'
'End Sub
'
'Public Sub PercentDesc_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
'
'End Sub
'
'Public Sub PercentDesc_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridItens.objControle = PercentDesc
'    lErro = Grid_Campo_Libera_Foco(objGridItens)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub PercentualBaixa_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PercentualComissao_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub PercentualComissao_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridComissoes)
'
'End Sub
'
'Public Sub PercentualComissao_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
'
'End Sub
'
'Public Sub PercentualComissao_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridComissoes.objControle = PercentualComissao
'    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub PercentualEmissao_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub PercentualEmissao_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridComissoes)
'
'End Sub
'
'Public Sub PercentualEmissao_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
'
'End Sub
'
'Public Sub PercentualEmissao_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridComissoes.objControle = PercentualEmissao
'    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub PesoBruto_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PesoLiquido_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Placa_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PlacaUF_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PlacaUF_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PrecoTotal_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub PrecoTotal_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridItens)
'
'End Sub
'
'Public Sub PrecoTotal_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
'
'End Sub
'
'Public Sub PrecoTotal_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridItens.objControle = PrecoTotal
'    lErro = Grid_Campo_Libera_Foco(objGridItens)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub PrecoUnitario_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub PrecoUnitario_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridItens)
'
'End Sub
'
'Public Sub PrecoUnitario_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
'
'End Sub
'
'Public Sub PrecoUnitario_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridItens.objControle = PrecoUnitario
'    lErro = Grid_Campo_Libera_Foco(objGridItens)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Produto_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Produto_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridItens)
'
'End Sub
'
'Public Sub Produto_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
'
'End Sub
'
'Public Sub Produto_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridItens.objControle = Produto
'    lErro = Grid_Campo_Libera_Foco(objGridItens)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Quantidade_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'Public Sub Quantidade_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridItens)
'
'End Sub
'
'Public Sub Quantidade_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
'
'End Sub
'
'Public Sub Quantidade_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridItens.objControle = Quantidade
'    lErro = Grid_Campo_Libera_Foco(objGridItens)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub Serie_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'
'Public Sub SerieNFiscalOriginal_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub SerieNFiscalOriginal_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub TabelaPreco_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub TipoNFiscal_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Transportadora_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Transportadora_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub UnidadeMed_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub UnidadeMed_Click()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub UnidadeMed_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridItens)
'
'End Sub
'
'Public Sub UnidadeMed_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
'
'End Sub
'
'Public Sub UnidadeMed_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridItens.objControle = UnidadeMed
'    lErro = Grid_Campo_Libera_Foco(objGridItens)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub ValorBaixa_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ValorBase_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ValorBase_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridComissoes)
'
'End Sub
'
'Public Sub ValorBase_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
'
'End Sub
'
'Public Sub ValorBase_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridComissoes.objControle = ValorBase
'    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub ValorComissao_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ValorComissao_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridComissoes)
'
'End Sub
'
'Public Sub ValorComissao_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
'
'End Sub
'
'Public Sub ValorComissao_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridComissoes.objControle = ValorComissao
'    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub ValorDesconto_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iValorDescontoAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ValorDespesas_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iValorDespesasAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ValorEmissao_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ValorEmissao_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridComissoes)
'
'End Sub
'
'Public Sub ValorEmissao_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
'
'End Sub
'
'Public Sub ValorEmissao_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridComissoes.objControle = ValorEmissao
'    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub ValorFrete_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iValorFreteAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ValorParcela_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub ValorParcela_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub ValorParcela_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub ValorParcela_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = ValorParcela
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub ValorProdutos_Change()
'
'    If Not gbCarregandoTela Then
'         Call DescontoGlobal_Recalcula
'    End If
'
'End Sub
'
'Private Sub DescontoGlobal_Recalcula()
'
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
'            Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)
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
'
'End Sub
'
'Private Function Cobranca_Automatica() As Long
''recalcula o tab de cobranca
'
'Dim lErro As Long
'Dim objCondicaoPagto As New ClassCondicaoPagto
'
'On Error GoTo Erro_Cobranca_Automatica
'
'    'Se automática estiver selecionada e a condicao de pagamento estiver preenchida
'    If CobrancaAutomatica.Value = 1 And Len(Trim(CondicaoPagamento.Text)) <> 0 Then
'        'Pega a condicao de pagamento da tela
'        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)
'        'Lê a condição de pagamento
'        lErro = CondicaoPagto_Le(objCondicaoPagto)
'        If lErro <> SUCESSO And lErro <> 19205 Then gError 46187
'        If lErro <> SUCESSO Then gError 46188
'        'Preenche o grid parcelas de acordo com a condição de pagamento
'        lErro = GridParcelas_Preenche(objCondicaoPagto)
'        If lErro <> SUCESSO Then gError 46189
'        CobrancaAutomatica.Tag = GRID_CHECKBOX_ATIVO
'        CobrancaAutomatica.Value = 1
'    End If
'
'    Cobranca_Automatica = SUCESSO
'
'    Exit Function
'
'Erro_Cobranca_Automatica:
'
'    Cobranca_Automatica = gErr
'
'    Select Case gErr
'
'        Case 46187, 46189
'
'        Case 46188
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub ValorSeguro_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    iValorSeguroAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'
'Public Sub Vendedor_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub Vendedor_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridComissoes)
'
'End Sub
'
'Public Sub Vendedor_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)
'
'End Sub
'
'Public Sub Vendedor_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridComissoes.objControle = Vendedor
'    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub VolumeEspecie_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub VolumeMarca_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub VolumeNumero_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub VolumeQuant_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Function Carrega_TiposDocInfo() As Long
''Carrega na combo os Tipo de Documentos relacionados com a tela
'
'Dim lErro As Long
'Dim objTipoDocInfo As New ClassTipoDocInfo
'Dim colTipoDocInfo As New colTipoDocInfo
'Dim vCodTipoDocInfo As Variant
'
'On Error GoTo Erro_Carrega_TiposDocInfo
'
'    Set colTipoDocInfo = gobjCRFAT.colTiposDocInfo
'
'    'Se for versão FULL
'    If giTipoVersao = VERSAO_FULL Then
'        'Carrega na combo só os Tipos ligados essa tela
'        For Each objTipoDocInfo In colTipoDocInfo
'            If objTipoDocInfo.sNomeTelaNFiscal = Me.Name Then
'                TipoNFiscal.AddItem CStr(objTipoDocInfo.iCodigo) & SEPARADOR & objTipoDocInfo.sNomeReduzido
'                TipoNFiscal.ItemData(TipoNFiscal.NewIndex) = objTipoDocInfo.iCodigo
'                'se for o tipo padrao, seleciona-o
'                If objTipoDocInfo.iPadrao = Padrao Then TipoNFiscal.ListIndex = TipoNFiscal.NewIndex
'            End If
'        Next
'
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
'
'    Carrega_TiposDocInfo = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_TiposDocInfo:
'
'    Carrega_TiposDocInfo = gErr
'
'    Select Case gErr
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Carrega_TabelaPreco() As Long
''Carrega na combo as Tabelas de Preço existentes
'
'Dim lErro As Long
'Dim objCodDescricao As AdmCodigoNome
'Dim colCodigoDescricao As New AdmColCodigoNome
'
'On Error GoTo Erro_Carrega_TabelaPreco
'
'    'Lê o código e a descrição de todas as Tabelas de Preços
'    lErro = Cod_Nomes_Le("TabelasDePreco", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao)
'    If lErro <> SUCESSO Then gError 42118
'
'    For Each objCodDescricao In colCodigoDescricao
'
'        'Adiciona o ítem na Lista de Tabela de Preços
'        TabelaPreco.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
'        TabelaPreco.ItemData(TabelaPreco.NewIndex) = objCodDescricao.iCodigo
'
'    Next
'
'    Carrega_TabelaPreco = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_TabelaPreco:
'
'    Carrega_TabelaPreco = gErr
'
'    Select Case gErr
'
'        Case 42118
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Carrega_Transportadora() As Long
''Carrega na combo as Transportadoras existentes
'
'Dim lErro As Long
'Dim objCodigoNome As New AdmCodigoNome
'Dim colCodigoNome As New AdmColCodigoNome
'
'On Error GoTo Erro_Carrega_Transportadora
'
'    'Lê o Código e o NOme de Todas as Transportadoras do BD
'    lErro = Cod_Nomes_Le("Transportadoras", "Codigo", "NomeReduzido", STRING_NOME_TABELA, colCodigoNome)
'    If lErro <> SUCESSO Then gError 42119
'
'    'Carrega a combo de Transportadoras
'    For Each objCodigoNome In colCodigoNome
'        Transportadora.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
'        Transportadora.ItemData(Transportadora.NewIndex) = objCodigoNome.iCodigo
'    Next
'
'    Carrega_Transportadora = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Transportadora:
'
'    Carrega_Transportadora = gErr
'
'    Select Case gErr
'
'        Case 42119
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Carrega_CanalVenda() As Long
''Carrega na combo os Canais de Venda existentes
'
'Dim lErro As Long
'Dim objCodDescricao As AdmCodigoNome
'Dim colCodigoDescricao As New AdmColCodigoNome
'
'On Error GoTo Erro_Carrega_CanalVenda
'
'    'Lê o código e a descrição de todas as Tabelas de Preços
'    lErro = Cod_Nomes_Le("CanalVenda", "Codigo", "NomeReduzido", STRING_CANAL_VENDA_NOME_REDUZIDO, colCodigoDescricao)
'    If lErro <> SUCESSO Then gError 42120
'
'    For Each objCodDescricao In colCodigoDescricao
'        'Adiciona o ítem na Lista de Tabela de Preços
'        CanalVenda.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
'        CanalVenda.ItemData(CanalVenda.NewIndex) = objCodDescricao.iCodigo
'    Next
'
'    Carrega_CanalVenda = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_CanalVenda:
'
'    Carrega_CanalVenda = gErr
'
'    Select Case gErr
'
'        Case 42120
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Carrega_Serie() As Long
''Carrega a combo de Séries com as séries lidas do BD
'
'Dim lErro As Long
'Dim objSerie As ClassSerie
'Dim colSerie As New colSerie
'
'On Error GoTo Erro_Carrega_Serie
'
'    'Lê as séries
'    lErro = Series_Le(colSerie)
'    If lErro <> SUCESSO Then gError 42121
'
'    'Carrega na combo
'    For Each objSerie In colSerie
'        Serie.AddItem objSerie.sSerie
'        SerieNFiscalOriginal.AddItem objSerie.sSerie
'    Next
'
''    'Seleciona a única série
''    If giTipoVersao = VERSAO_LIGHT Then
''
''        If Serie.ListCount > 0 Then
''            Serie.ListIndex = 0
''        End If
''
''    End If
'
'    Carrega_Serie = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Serie:
'
'    Carrega_Serie = gErr
'
'    Select Case gErr
'
'        Case 42121, 500115, 500116
'
'        Case 500117
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Seleciona_Serie_Padrao() As Long
'
'Dim lErro As Long
'Dim sSeriePadrao As String
'Dim objSerie As New ClassSerie
'
'On Error GoTo Erro_Seleciona_Serie_Padrao
'
'    'Lê série Padrão
'    lErro = Serie_Le_Padrao(sSeriePadrao)
'    If lErro <> SUCESSO Then gError 500130
'
'    'Se possui série padrão
'    If Len(Trim(sSeriePadrao)) > 0 Then
'
'        'Lê a série
'        objSerie.sSerie = sSeriePadrao
'        lErro = Serie_Le(objSerie)
'        If lErro <> SUCESSO And lErro <> 22202 Then gError 500131
'
'        'Série não cadastrada
'        If lErro = 22202 Then gError 500132
'
'        'Se a Série Padrão é para fatura
'        If objSerie.iTipoFormulario = TIPO_FORMULARIO_NFISCAL_FATURA Or objSerie.iTipoFormulario = TIPO_FORMULARIO_NFISCAL_FATURA_SERVICO Then
'
'            'Coloca ela na combo
'            Serie.Text = sSeriePadrao
'
'        End If
'
'    End If
'
'    Seleciona_Serie_Padrao = SUCESSO
'
'    Exit Function
'
'Erro_Seleciona_Serie_Padrao:
'
'    Seleciona_Serie_Padrao = gErr
'
'    Select Case gErr
'
'        Case 500130, 500131
'
'        Case 500132
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Carrega_PlacaUF() As Long
''Lê as Siglas dos Estados e alimenta a list da Combobox PlacaUF
'
'Dim lErro As Long
'Dim colSiglasUF As New Collection
'Dim iIndice As Integer
'
'On Error GoTo Erro_Carrega_PlacaUF
'
'    Set colSiglasUF = gcolUFs
'
'    'Adiciona na Combo PlacaUF
'    For iIndice = 1 To colSiglasUF.Count
'        PlacaUF.AddItem colSiglasUF.Item(iIndice)
'    Next
'
'    Carrega_PlacaUF = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_PlacaUF:
'
'    Carrega_PlacaUF = gErr
'
'    Select Case gErr
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'End Function
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
'    lErro = CondicoesPagto_Le_Recebimento(colCod_DescReduzida)
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
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Carrega_TipoDesconto() As Long
''Carrega na combo os Tipos de Desconto existentes
'
'Dim lErro As Long
'Dim objCodDescricao As AdmCodigoNome
'Dim colCodigoDescricao As New AdmColCodigoNome
'
'On Error GoTo Erro_Carrega_TipoDesconto
'
'    Set colCodigoDescricao = gobjCRFAT.colTiposDesconto
'
'    For Each objCodDescricao In colCodigoDescricao
'        'Adiciona o ítem nas List's das Combos de Tipos Desconto
'        Desconto1Codigo.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
'        Desconto1Codigo.ItemData(Desconto1Codigo.NewIndex) = objCodDescricao.iCodigo
'        Desconto2Codigo.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
'        Desconto2Codigo.ItemData(Desconto2Codigo.NewIndex) = objCodDescricao.iCodigo
'        Desconto3Codigo.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
'        Desconto3Codigo.ItemData(Desconto3Codigo.NewIndex) = objCodDescricao.iCodigo
'    Next
'
'    Carrega_TipoDesconto = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_TipoDesconto:
'
'    Carrega_TipoDesconto = gErr
'
'    Select Case gErr
'
'        Case 26492
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Sub GridItens_InicializaColunas()
'
'    'Colunas da Grid
'    iGrid_Produto_Col = 1
'    iGrid_DescProduto_Col = 2
'    iGrid_UnidadeMed_Col = 3
'    iGrid_Quantidade_Col = 4
'    iGrid_PrecoUnitario_Col = 5
'    iGrid_PercDesc_Col = 6
'    iGrid_Desconto_Col = 7
'    iGrid_PrecoTotal_Col = 8
'
'End Sub
'
'Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
''Inicializa o Grid de Itens
'
'    Set objGridInt.objForm = Me
'
'    'Títulos das colunas
'    objGridInt.colColuna.Add ("Item")
'    objGridInt.colColuna.Add ("Produto")
'    objGridInt.colColuna.Add ("Descrição")
'    objGridInt.colColuna.Add ("U.M.")
'    objGridInt.colColuna.Add ("Quantidade")
'    objGridInt.colColuna.Add ("Preço Unitário")
'    objGridInt.colColuna.Add ("% Desconto")
'    objGridInt.colColuna.Add ("Desconto")
'    objGridInt.colColuna.Add ("Preço Total")
'
'    'Controles que participam do Grid
'    objGridInt.colCampo.Add (Produto.Name)
'    objGridInt.colCampo.Add (DescricaoItem.Name)
'    objGridInt.colCampo.Add (UnidadeMed.Name)
'    objGridInt.colCampo.Add (Quantidade.Name)
'    objGridInt.colCampo.Add (PrecoUnitario.Name)
'    objGridInt.colCampo.Add (PercentDesc.Name)
'    objGridInt.colCampo.Add (Desconto.Name)
'    objGridInt.colCampo.Add (PrecoTotal.Name)
'
'    Call GridItens_InicializaColunas
'
'    'Grid do GridInterno
'    objGridInt.objGrid = GridItens
'
'    'Todas as linhas do grid
'    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1
'
'    'Linhas visíveis do grid
'    objGridInt.iLinhasVisiveis = 5
'
'    'Largura da primeira coluna
'    GridItens.ColWidth(0) = 500
'
'    'Largura automática para as outras colunas
'    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
'
'    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
'
'    'Chama função que inicializa o Grid
'    Call Grid_Inicializa(objGridInt)
'
'    Inicializa_Grid_Itens = SUCESSO
'
'    Exit Function
'
'End Function
'
'Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
''Inicializa o Grid de Parcelas
'
'    'Form do Grid
'    Set objGridInt.objForm = Me
'
'    'Títulos das colunas
'    objGridInt.colColuna.Add ("Parcela")
'    objGridInt.colColuna.Add ("Vencimento")
'    objGridInt.colColuna.Add ("Valor")
'    objGridInt.colColuna.Add ("Cobrador")
'    objGridInt.colColuna.Add ("Carteira")
'    objGridInt.colColuna.Add ("Desconto 1 Tipo")
'    objGridInt.colColuna.Add ("Desc.1 Data")
'    objGridInt.colColuna.Add ("Desc.1 Valor")
'    objGridInt.colColuna.Add ("Desc.1 %")
'    objGridInt.colColuna.Add ("Desconto 2 Tipo")
'    objGridInt.colColuna.Add ("Desc.2 Data")
'    objGridInt.colColuna.Add ("Desc.2 Valor")
'    objGridInt.colColuna.Add ("Desc.2 %")
'    objGridInt.colColuna.Add ("Desconto 3 Tipo")
'    objGridInt.colColuna.Add ("Desc.3 Data")
'    objGridInt.colColuna.Add ("Desc.3 Valor")
'    objGridInt.colColuna.Add ("Desc.3 %")
'
'    'Controles que participam do Grid
'    objGridInt.colCampo.Add (DataVencimento.Name)
'    objGridInt.colCampo.Add (ValorParcela.Name)
'    objGridInt.colCampo.Add (Cobrador.Name)
'    objGridInt.colCampo.Add (CarteiraCobrador.Name)
'    objGridInt.colCampo.Add (Desconto1Codigo.Name)
'    objGridInt.colCampo.Add (Desconto1Ate.Name)
'    objGridInt.colCampo.Add (Desconto1Valor.Name)
'    objGridInt.colCampo.Add (Desconto1Percentual.Name)
'    objGridInt.colCampo.Add (Desconto2Codigo.Name)
'    objGridInt.colCampo.Add (Desconto2Ate.Name)
'    objGridInt.colCampo.Add (Desconto2Valor.Name)
'    objGridInt.colCampo.Add (Desconto2Percentual.Name)
'    objGridInt.colCampo.Add (Desconto3Codigo.Name)
'    objGridInt.colCampo.Add (Desconto3Ate.Name)
'    objGridInt.colCampo.Add (Desconto3Valor.Name)
'    objGridInt.colCampo.Add (Desconto3Percentual.Name)
'
'    'Colunas do Grid
'    iGrid_Vencimento_col = 1
'    iGrid_ValorParcela_Col = 2
'    iGrid_Cobranca_Col = 3
'    iGrid_CarteiraCobranca_Col = 4
'    iGrid_Desc1Codigo_Col = 5
'    iGrid_Desc1Ate_Col = 6
'    iGrid_Desc1Valor_Col = 7
'    iGrid_Desc1Percentual_Col = 8
'    iGrid_Desc2Codigo_Col = 9
'    iGrid_Desc2Ate_Col = 10
'    iGrid_Desc2Valor_Col = 11
'    iGrid_Desc2Percentual_Col = 12
'    iGrid_Desc3Codigo_Col = 13
'    iGrid_Desc3Ate_Col = 14
'    iGrid_Desc3Valor_Col = 15
'    iGrid_Desc3Percentual_Col = 16
'
'    'Grid do GridInterno
'    objGridInt.objGrid = GridParcelas
'
'    'Todas as linhas do grid
'    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1
'
'    'Linhas visíveis do grid
'    objGridInt.iLinhasVisiveis = 6
'
'    'Largura da primeira coluna
'    GridParcelas.ColWidth(0) = 700
'
'    'Largura automática para as outras colunas
'    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
'
'    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
'
'    'Chama função que inicializa o Grid
'    Call Grid_Inicializa(objGridInt)
'
'    Inicializa_Grid_Parcelas = SUCESSO
'
'    Exit Function
'
'End Function
'
'
'Private Function Inicializa_Grid_Alocacoes(objGridInt As AdmGrid) As Long
''Inicializa o Grid de Alocações
'
'    Set objGridInt.objForm = Me
'
'    'Títulos das colunas
'    objGridInt.colColuna.Add (" ")
'    objGridInt.colColuna.Add ("Item")
'    objGridInt.colColuna.Add ("Produto")
'    objGridInt.colColuna.Add ("Almoxarifado")
'    objGridInt.colColuna.Add ("U.M.")
'    objGridInt.colColuna.Add ("Quant. Alocada")
'    objGridInt.colColuna.Add ("Quant. Vendida")
'
'    'Controles que participam do Grid
'    objGridInt.colCampo.Add (ItemNFiscal.Name)
'    objGridInt.colCampo.Add (ProdutoAlmox.Name)
'    objGridInt.colCampo.Add (Almox.Name)
'    objGridInt.colCampo.Add (UnidadeMedEst.Name)
'    objGridInt.colCampo.Add (QuantAlocada.Name)
'    objGridInt.colCampo.Add (QuantVendida.Name)
'
'    'Colunas da Grid
'    iGrid_Item_Col = 1
'    iGrid_ProdutoAloc_Col = 2
'    iGrid_AlmoxAloc_Col = 3
'    iGrid_UMAloc_Col = 4
'    iGrid_QuantAloc_Col = 5
'    iGrid_QuantVend_Col = 6
'
'    'Grid do GridInterno
'    objGridInt.objGrid = GridAlocacao
'
'    'Todas as linhas do grid
'    objGridInt.objGrid.Rows = NUM_MAX_ALOCACOES + 1
'
'    'Linhas visíveis do grid
'    objGridInt.iLinhasVisiveis = 10
'
'    'Largura da primeira coluna
'    GridAlocacao.ColWidth(0) = 500
'
'    'Largura automática para as outras colunas
'    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
'
'    objGridInt.iProibidoExcluir = 1
'
'    'Chama função que inicializa o Grid
'    Call Grid_Inicializa(objGridInt)
'
'    Inicializa_Grid_Alocacoes = SUCESSO
'
'    Exit Function
'
'End Function
'
'
''""""""""""""""""""""""""""""""""""""""""""""""
''"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
''""""""""""""""""""""""""""""""""""""""""""""""
'Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
''Extrai os campos da tela que correspondem aos campos no BD
'
'Dim lErro As Long
'Dim objNFiscal As New ClassNFiscal
'
'On Error GoTo Erro_Tela_Extrai
'
'    'Informa tabela associada à Tela
'    sTabela = "NFiscalFatura"
'
'    'Lê os dados da Tela NFiscal
'    lErro = Move_NF_Memoria(objNFiscal)
'    If lErro <> SUCESSO Then gError 42132
'
'    'Preenche a coleção colCampoValor, com nome do campo,
'    'valor atual (com a tipagem do BD), tamanho do campo
'    'no BD no caso de STRING e Key igual ao nome do campo
'    colCampoValor.Add "NumIntDoc", objNFiscal.lNumIntDoc, 0, "NumIntDoc"
'    colCampoValor.Add "TipoNFiscal", objNFiscal.iTipoNFiscal, 0, "TipoNFiscal"
'    colCampoValor.Add "NaturezaOp", objNFiscal.sNaturezaOp, STRING_BUFFER_MAX_TEXTO, "NaturezaOp"
'    colCampoValor.Add "Serie", objNFiscal.sSerie, STRING_BUFFER_MAX_TEXTO, "Serie"
'    colCampoValor.Add "NumNotaFiscal", objNFiscal.lNumNotaFiscal, 0, "NumNotaFiscal"
'    colCampoValor.Add "DataEmissao", objNFiscal.dtDataEmissao, 0, "DataEmissao"
'    colCampoValor.Add "DataSaida", objNFiscal.dtDataSaida, 0, "DataSaida"
'    colCampoValor.Add "Placa", objNFiscal.sPlaca, STRING_BUFFER_MAX_TEXTO, "Placa"
'    colCampoValor.Add "PlacaUF", objNFiscal.sPlacaUF, STRING_BUFFER_MAX_TEXTO, "PlacaUF"
'    colCampoValor.Add "VolumeQuant", objNFiscal.lVolumeQuant, 0, "VolumeQuant"
'    colCampoValor.Add "VolumeEspecie", objNFiscal.sVolumeEspecie, STRING_BUFFER_MAX_TEXTO, "VolumeEspecie"
'    colCampoValor.Add "VolumeMarca", objNFiscal.sVolumeMarca, STRING_BUFFER_MAX_TEXTO, "VolumeMarca"
'    colCampoValor.Add "MensagemNota", objNFiscal.sMensagemNota, STRING_BUFFER_MAX_TEXTO, "MensagemNota"
'
''horasaida
'    colCampoValor.Add "HoraSaida", CDbl(objNFiscal.dtHoraSaida), 0, "HoraSaida"
'    colCampoValor.Add "ValorProdutos", objNFiscal.dValorProdutos, 0, "ValorProdutos"
'    colCampoValor.Add "ValorFrete", objNFiscal.dValorFrete, 0, "ValorFrete"
'    colCampoValor.Add "ValorSeguro", objNFiscal.dValorSeguro, 0, "ValorSeguro"
'    colCampoValor.Add "ValorOutrasDespesas", objNFiscal.dValorOutrasDespesas, 0, "ValorOutrasDespesas"
'    colCampoValor.Add "ValorDesconto", objNFiscal.dValorDesconto, 0, "ValorDesconto"
'    colCampoValor.Add "ValorTotal", objNFiscal.dValorTotal, 0, "ValorTotal"
'    colCampoValor.Add "PesoLiq", objNFiscal.dPesoLiq, 0, "Pesoliq"
'    colCampoValor.Add "PesoBruto", objNFiscal.dPesoBruto, 0, "PesoBruto"
'    colCampoValor.Add "NumIntNotaOriginal", objNFiscal.lNumIntNotaOriginal, 0, "NumIntNotaOriginal"
'    colCampoValor.Add "Canal", objNFiscal.iCanal, 0, "Canal"
'    colCampoValor.Add "Cliente", objNFiscal.lCliente, 0, "Cliente"
'    colCampoValor.Add "FilialCli", objNFiscal.iFilialCli, 0, "FilialCli"
'    colCampoValor.Add "FilialEntrega", objNFiscal.iFilialEntrega, 0, "FilialEntrega"
'    colCampoValor.Add "NumPedidoTerc", objNFiscal.sNumPedidoTerc, STRING_BUFFER_MAX_TEXTO, "NumPedidoTerc"
'    colCampoValor.Add "CodTransportadora", objNFiscal.iCodTransportadora, 0, "CodTransportadora"
'    colCampoValor.Add "TabelaPreco", objNFiscal.iTabelaPreco, 0, "TabelaPreco"
'    colCampoValor.Add "NumIntDocCPR", objNFiscal.lNumIntDocCPR, 0, "NumIntDocCPR"
'    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_BUFFER_MAX_TEXTO, "VolumeNumero"
'    colCampoValor.Add "FreteRespons", objNFiscal.iFreteRespons, 0, "FreteRespons"
'    colCampoValor.Add "DataReferencia", objNFiscal.dtDataReferencia, 0, "DataReferencia"
'    colCampoValor.Add "Status", objNFiscal.iStatus, 0, "Status"
'    colCampoValor.Add "ValorMercadoria", objNFiscal.dValorMercadoria, 0, "ValorMercadoria"
'    colCampoValor.Add "ValorContainer", objNFiscal.dValorContainer, 0, "ValorContainer"
'    colCampoValor.Add "OrigemNF", objNFiscal.sOrigem, STRING_BUFFER_MAX_TEXTO, "OrigemNF"
'    colCampoValor.Add "DestinoNF", objNFiscal.sDestino, STRING_BUFFER_MAX_TEXTO, "DestinoNF"
'
'    'Filtros para o Sistema de Setas
'    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
'    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
''    colSelecao.Add "Status", OP_DIFERENTE, STATUS_CANCELADO
'
'    Exit Sub
'
'Erro_Tela_Extrai:
'
'    Select Case gErr
'
'        Case 42132
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
''Preenche os campos da tela com os correspondentes do BD
'
'Dim lErro As Long
'Dim objNFiscal As New ClassNFiscal
'
'On Error GoTo Erro_Tela_Preenche
'
'    objNFiscal.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
'
'    If objNFiscal.lNumIntDoc <> 0 Then
'
'        'Carrega objNFiscal com os dados passados em colCampoValor
'        objNFiscal.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
'        objNFiscal.iFilialEmpresa = giFilialEmpresa
'        objNFiscal.iTipoNFiscal = colCampoValor.Item("TipoNFiscal").vValor
'        objNFiscal.sNaturezaOp = colCampoValor.Item("NaturezaOP").vValor
'        objNFiscal.lCliente = colCampoValor.Item("Cliente").vValor
'        objNFiscal.iFilialCli = colCampoValor.Item("FilialCli").vValor
'        objNFiscal.sSerie = colCampoValor.Item("Serie").vValor
'        objNFiscal.lNumNotaFiscal = colCampoValor.Item("NumNotaFiscal").vValor
'        objNFiscal.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
'        objNFiscal.dtDataSaida = colCampoValor.Item("DataSaida").vValor
''horasaida
'        objNFiscal.dtHoraSaida = colCampoValor.Item("HoraSaida").vValor
'        objNFiscal.dValorFrete = colCampoValor.Item("ValorFrete").vValor
'        objNFiscal.dValorSeguro = colCampoValor.Item("ValorSeguro").vValor
'        objNFiscal.dValorOutrasDespesas = colCampoValor.Item("ValorOutrasDespesas").vValor
'        objNFiscal.dValorDesconto = colCampoValor.Item("ValorDesconto").vValor
'        objNFiscal.dValorTotal = colCampoValor.Item("ValorTotal").vValor
'        objNFiscal.iCodTransportadora = colCampoValor.Item("CodTransportadora").vValor
'        objNFiscal.sPlaca = colCampoValor.Item("Placa").vValor
'        objNFiscal.sPlacaUF = colCampoValor.Item("PlacaUF").vValor
'        objNFiscal.lVolumeQuant = colCampoValor.Item("VolumeQuant").vValor
'        objNFiscal.sVolumeEspecie = colCampoValor.Item("VolumeEspecie").vValor
'        objNFiscal.sVolumeMarca = colCampoValor.Item("VolumeMarca").vValor
'        objNFiscal.sMensagemNota = colCampoValor.Item("MensagemNota").vValor
'        objNFiscal.dPesoLiq = colCampoValor.Item("PesoLiq").vValor
'        objNFiscal.dPesoBruto = colCampoValor.Item("PesoBruto").vValor
'        objNFiscal.lNumIntNotaOriginal = colCampoValor.Item("NumIntNotaOriginal").vValor
'        objNFiscal.dValorProdutos = colCampoValor.Item("ValorProdutos").vValor
'        objNFiscal.iCanal = colCampoValor.Item("Canal").vValor
'        objNFiscal.dValorProdutos = colCampoValor.Item("ValorProdutos").vValor
'        objNFiscal.iFilialEntrega = colCampoValor.Item("FilialEntrega").vValor
'        objNFiscal.sNumPedidoTerc = colCampoValor.Item("NumPedidoTerc").vValor
'        objNFiscal.iTabelaPreco = colCampoValor.Item("TabelaPreco").vValor
'        objNFiscal.lNumIntDocCPR = colCampoValor.Item("NumIntDocCPR").vValor
'        objNFiscal.iFreteRespons = colCampoValor.Item("FreteRespons").vValor
'        objNFiscal.dtDataReferencia = colCampoValor.Item("DataReferencia").vValor
'        objNFiscal.sVolumeNumero = colCampoValor.Item("VolumeNumero").vValor
'        objNFiscal.iStatus = colCampoValor.Item("Status").vValor
'        objNFiscal.dValorMercadoria = colCampoValor.Item("ValorMercadoria").vValor
'        objNFiscal.dValorContainer = colCampoValor.Item("ValorContainer").vValor
'        objNFiscal.sOrigem = colCampoValor.Item("OrigemNF").vValor
'        objNFiscal.sDestino = colCampoValor.Item("DestinoNF").vValor
'
'        'Coloca os dados da NFiscal na tela
'        lErro = Traz_NFiscal_Tela(objNFiscal)
'        If lErro <> SUCESSO Then gError 42131
'
'    End If
'
'    Exit Sub
'
'Erro_Tela_Preenche:
'
'    Select Case gErr
'
'        Case 42131
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ClienteLabel_Click()
'
'Dim objCliente As New ClassCliente
'Dim colSelecao As Collection
'
'    'Preenche NomeReduzido com o cliente da tela
'    objCliente.sNomeReduzido = Cliente.Text
'
'    'Chama Tela ClienteLista
'    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)
'
'End Sub
'
'Private Sub objEventoCliente_evSelecao(obj1 As Object)
'
'Dim objCliente As ClassCliente
'Dim bCancel As Boolean
'
'    Set objCliente = obj1
'
'    'Preenche campo Cliente
'    Cliente.Text = objCliente.sNomeReduzido
'
'    'Executa o Validate
'    Call Cliente_Validate(bCancel)
'
'    Me.Show
'
'    Exit Sub
'
'End Sub
'
'Public Sub SerieLabel_Click()
'
'Dim objSerie As New ClassSerie
'Dim colSelecao As Collection
'
'    'recolhe a serie da tela
'    objSerie.sSerie = Serie.Text
'
'    'Chama a Tela de Browse SerieLista
'    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie)
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoSerie_evSelecao(obj1 As Object)
'
'Dim objSerie As ClassSerie
'
'    Set objSerie = obj1
'
'    'Coloca a Série na Tela
'    Serie.Text = objSerie.sSerie
'
'    Me.Show
'
'    Exit Sub
'
'End Sub
'
'Public Sub NFiscalLabel_Click()
'
'Dim lErro As Long
'Dim objNFiscal As New ClassNFiscal
'Dim colSelecao As Collection
'
'    'Recolhe os dados da Nota Fiscal
'    lErro = Move_NF_Memoria(objNFiscal)
'    If lErro <> SUCESSO Then Exit Sub
'
'    'Chama a Tela NFiscalFaturaLista
'    Call Chama_Tela("NFiscalFaturaLista", colSelecao, objNFiscal, objEventoNFiscal)
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoNFiscal_evSelecao(obj1 As Object)
'
'Dim lErro As Long
'Dim objNFiscal As ClassNFiscal
'
'On Error GoTo Erro_objEventoNFiscal_evSelecao
'
'    Set objNFiscal = obj1
'
'    'Coloca na Tela a Nota Fiscal escolhida
'    lErro = Traz_NFiscal_Tela(objNFiscal)
'    If lErro <> SUCESSO Then gError 42133
'
'    lErro = ComandoSeta_Fechar(Me.Name)
'
'    Me.Show
'
'    Exit Sub
'
'Erro_objEventoNFiscal_evSelecao:
'
'    Select Case gErr
'
'        Case 42133
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub BotaoProdutos_Click()
'
'Dim lErro As Long
'Dim sProduto As String
'Dim iPreenchido As Integer
'Dim objProduto As New ClassProduto
'Dim colSelecao As Collection
'Dim sProduto1 As String
'
'On Error GoTo Erro_BotaoProdutos_Click
'
'    If Me.ActiveControl Is Produto Then
'
'        sProduto1 = Produto.Text
'
'    Else
'
'        'Verifica se tem alguma linha selecionada no Grid
'        If GridItens.Row = 0 Then gError 43697
'
'        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
'
'    End If
'
'    lErro = Produto_Formata(sProduto1, sProduto, iPreenchido)
'    If lErro <> SUCESSO Then gError 83085
'
'    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
'
'    objProduto.sCodigo = sProduto
'
'    'Chama a Tela ProdutoVendaLista
'    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)
'
'    Exit Sub
'
'Erro_BotaoProdutos_Click:
'
'    Select Case gErr
'
'        Case 43697
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
'
'        Case 83085
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoProduto_evSelecao(obj1 As Object)
'
'Dim lErro As Long
'Dim objProduto As ClassProduto
'Dim sProdutoEnxuto As String
'
'On Error GoTo Erro_objEventoProduto_evSelecao
'
'    'verifica se tem alguma linha do Grid selecionada
'    If GridItens.Row = 0 Then gError 42134
'
'    'Verifica se o Produto está preenchido
'    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then
'
'        Set objProduto = obj1
'
'        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
'        If lErro <> SUCESSO Then gError 42135
'
'        Produto.PromptInclude = False
'        Produto.Text = sProdutoEnxuto
'        Produto.PromptInclude = True
'
'        'Coloca o produto, a Descrição e a Unidade de Medida da tela
'        GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
'
'        lErro = Produto_Saida_Celula(objProduto)
'        If lErro <> SUCESSO Then gError 42136
'
'    End If
'
'    Me.Show
'
'    Exit Sub
'
'Erro_objEventoProduto_evSelecao:
'
'    Select Case gErr
'
'        Case 42134, 42136
'
'        Case 42135
'            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub BotaoEstoqueProd_Click()
'
'Dim lErro As Long
'Dim iProdutoPreenchido As Integer
'Dim sProdutoFormatado As String
'Dim iAlmoxarifado As Integer
'Dim objEstoqueProd As New ClassEstoqueProduto
'Dim colSelecao As New Collection
'
'On Error GoTo Erro_BotaoEstoqueProd_Click
'
'    'Verifica se a linha do Grid é válida
'    If GridItens.Row = 0 Then gError 43698
'
'    'Verifica se o Produto na linha do Grid está preenchido
'    lErro = Produto_Formata(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'
'    If iProdutoPreenchido <> PRODUTO_VAZIO Then
'        objEstoqueProd.sProduto = sProdutoFormatado
'
'        colSelecao.Add sProdutoFormatado
'
'        'Lê o Almoxarifado Padrão do Produto
'        lErro = AlmoxarifadoPadrao_Le(giFilialEmpresa, sProdutoFormatado, iAlmoxarifado)
'        If lErro <> SUCESSO And lErro <> 23796 Then gError 58050
'
'        If lErro = SUCESSO Then objEstoqueProd.iAlmoxarifado = iAlmoxarifado
'
'        'Chama a tela EstoqueProdutoFilialLista
'        Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProd, objEventoEstoqueProd)
'
'    Else
'        gError 43729
'    End If
'
'    Exit Sub
'
'Erro_BotaoEstoqueProd_Click:
'
'    Select Case gErr
'
'        Case 43698
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
'
'        Case 43729
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoEstoqueProd_evSelecao(obj1 As Object)
'
'    Me.Show
'
'End Sub
'
'Public Sub TransportadoraLabel_Click()
'
'Dim objTransportadora As New ClassTransportadora
'Dim colSelecao As Collection
'
'    'Se a transportadora estiver preenchida recolhe
'    If Len(Trim(Transportadora.Text)) > 0 Then objTransportadora.iCodigo = Codigo_Extrai(Transportadora.Text)
'
'    'Chama a Tela de browse de Transportadora
'    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTransportadora)
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoTransportadora_evSelecao(obj1 As Object)
'
'Dim objTransportadora As ClassTransportadora
'
'    Set objTransportadora = obj1
'
'    'Coloca a Transportdora na Tela
'    Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido
'
'    Me.Show
'
'    Exit Sub
'
'End Sub
'
'Public Sub SerieNFOriginalLabel_Click()
'
'Dim objSerie As New ClassSerie
'Dim colSelecao As Collection
'
'    'Recolhe a Série da Nota Fiscal Original da tela
'    objSerie.sSerie = SerieNFiscalOriginal.Text
'
'    'Chama a Tela de Browse SerieLista
'    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerieOriginal)
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoSerieOriginal_evSelecao(obj1 As Object)
'
'Dim objSerie As ClassSerie
'
'    Set objSerie = obj1
'
'    'Coloca a Série da Nota Fiscal Original na tela
'    SerieNFiscalOriginal.Text = objSerie.sSerie
'
'    Me.Show
'
'End Sub
'
'Public Sub NFiscalOriginalLabel_Click()
'
'Dim lErro As Long
'Dim objNFiscal As New ClassNFiscal
'Dim colSelecao As New Collection
'Dim sCliente As String
'
'On Error GoTo Erro_NFiscalOriginalLabel_Click
'
'    'Se o cliente não está preenchido => erro
'    If Len(Trim(Cliente.Text)) = 0 Then gError 101090
'
'    'Recolhe os dados da Nota Fiscal
'    lErro = Move_NF_Memoria(objNFiscal)
'    If lErro <> SUCESSO Then gError 42137
'
'    'Guarda a Serie e o Número da Nota Fiscal Original da Tela
'    objNFiscal.sSerie = SerieNFiscalOriginal.Text
'    If Len(Trim(NFiscalOriginal.ClipText)) > 0 Then
'        objNFiscal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
'    Else
'        objNFiscal.lNumNotaFiscal = 0
'    End If
'
'    'Guarda o nome reduzido do Cliente
'    sCliente = Cliente.Text
'
'    'Faz o filtro por cliente
'    colSelecao.Add sCliente
'
'    'Chama a Tela NFiscalNormalLista
'    Call Chama_Tela("NFiscalFaturaNormalLista", colSelecao, objNFiscal, objEventoNFiscalOriginal)
'
'    Exit Sub
'
'Erro_NFiscalOriginalLabel_Click:
'
'    Select Case gErr
'
'        Case 42137
'
'        Case 101090
'            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoNFiscalOriginal_evSelecao(obj1 As Object)
'
'Dim objNFiscal As ClassNFiscal
'
'    Set objNFiscal = obj1
'
'    'Preenche a Série e o Número da Nota Fiscal Original
'    SerieNFiscalOriginal.Text = objNFiscal.sSerie
'    NFiscalOriginal.Text = objNFiscal.lNumNotaFiscal
'
'    Me.Show
'
'    Exit Sub
'
'End Sub
'
'Public Sub CondPagtoLabel_DblClick()
'
'Dim objCondicaoPagto As New ClassCondicaoPagto
'Dim colSelecao As New Collection
'
'    'Se Condição de Pagto estiver preenchida, extrai o código
'    If Len(Trim(CondicaoPagamento.Text)) > 0 Then
'        objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagamento.Text)
'    End If
'
'    'Chama a Tela CondicoesPagamentoCRLista
'    Call Chama_Tela("CondicaoPagtoCRLista", colSelecao, objCondicaoPagto, objEventoCondPagto)
'
'End Sub
'
'Private Sub objEventoCondPagto_evSelecao(obj1 As Object)
'
'Dim lErro As Long
'Dim objCondicaoPagto As ClassCondicaoPagto
'
'On Error GoTo Erro_objEventoCondPagto_evSelecao
'
'    Set objCondicaoPagto = obj1
'
'    'Preenche campo CondicaoPagamento
'    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida
'
'    'Se Valor nao estiver preenchido
'    If Len(Trim(ValorTotal.Caption)) = 0 Then Exit Sub
'
'    'Se DataEmissao estiver preenchida e Valor for positivo
'    If Len(Trim(DataEmissao.ClipText)) > 0 And (CDbl(ValorTotal.Caption) > 0) Then
'
'        'Preenche GridParcelas a partir da Condição de Pagto
'        lErro = Cobranca_Automatica()
'        If lErro <> SUCESSO Then gError 42138
'
'    End If
'
'    Me.Show
'
'    Exit Sub
'
'Erro_objEventoCondPagto_evSelecao:
'
'    Select Case gErr
'
'        Case 42138
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'     End Select
'
'     Exit Sub
'
'End Sub
'
'Public Sub BotaoVendedores_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoVendedores_Click
'
'    'Chama a tela de browse de Vendedores
'    lErro = objTabComissoes.BotaoVendedores_Click()
'    If lErro <> SUCESSO Then gError 43696
'
'    Exit Sub
'
'Erro_BotaoVendedores_Click:
'
'    Select Case gErr
'
'        Case 43696
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoVendedor_evSelecao(obj1 As Object)
'
'Dim objVendedor As ClassVendedor
'Dim lErro As Long
'
'On Error GoTo Erro_objEventoVendedor_evSelecao
'
'    Set objVendedor = obj1
'
'    If GridComissoes.Row > 0 Then
'
'        'Preenche a linha de Vendedor com dados default
'        lErro = objTabComissoes.VendedorLinha_Preenche(objVendedor)
'        If lErro <> SUCESSO Then gError 51617
'
'    End If
'
'
'    Me.Show
'
'    Exit Sub
'
'Erro_objEventoVendedor_evSelecao:
'
'    Select Case gErr
'
'        Case 51617  'tratado na rotina chamada
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long
'
'Dim lErro As Long
'Dim objTipoDocInfo As New ClassTipoDocInfo
'
'On Error GoTo Erro_Trata_Parametros
'
'    'Verifica se alguma nota foi passada por parametro
'    If Not (objNFiscal Is Nothing) Then
'
'        lErro = CargaPosFormLoad(True)
'        If lErro <> SUCESSO Then gError 64005
'
'        'Tenta ler a nota Fiscal passada como parâmetro
'        lErro = NFiscal_Le(objNFiscal)
'        If lErro <> SUCESSO And lErro <> 31442 Then gError 42403
'
'        'Se não achou a NFiscal --> erro
'        If lErro <> SUCESSO Then gError 42406
'
'        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
'
'        'Lê o Tipo de Documento com o Código de objNFiscal
'        lErro = TipoDocInfo_Le_Codigo(objTipoDocInfo)
'        If lErro <> SUCESSO And lErro <> 31415 Then gError 42405
'
'        'Se não encontrar --> Erro
'        If lErro = 31415 Then gError 42407
'
'        If objTipoDocInfo.sNomeTelaNFiscal <> Me.Name Then gError 42408
'
'        'Traz a nota para a tela
'        lErro = Traz_NFiscal_Tela(objNFiscal)
'        If lErro <> SUCESSO Then gError 42404
'
'    End If
'
'    iAlterado = 0
'
'    Trata_Parametros = SUCESSO
'
'    Exit Function
'
'Erro_Trata_Parametros:
'
'    Trata_Parametros = gErr
'
'    Select Case gErr
'
'        Case 42403, 42404, 42405, 64005
'
'        Case 42406
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)
'
'        Case 42407
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)
'
'        Case 42408
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_DIFERENTE_NF_FATURA", gErr, objNFiscal.iTipoNFiscal)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    iAlterado = 0
'
'    Exit Function
'
'End Function
'
'Public Sub TipoNFiscal_Click()
'
'Dim lErro As Long
'Dim iCodigo As Integer
'Dim objTipoDocInfo As New ClassTipoDocInfo
'
'On Error GoTo Erro_TipoNFiscal_Click
'
'    Frame1(TAB_Almoxarifado).Enabled = True
'
'    If TipoNFiscal.ListIndex = -1 Then Exit Sub
'
'    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
'
'    gobjRastreamento.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
'
'    'Lê o Tipo de Documento
'    lErro = TipoDocInfo_Le_Codigo(objTipoDocInfo)
'    If lErro <> SUCESSO And lErro <> 31415 Then gError 42138
'
'    'se não estiver cadastrado ==> erro
'    If lErro = 31415 Then gError 42139
'
'    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Or objTipoDocInfo.iCodigo = DOCINFO_NFISFVNE Or objTipoDocInfo.iCodigo = DOCINFO_NFISFS Then
'        Frame1(TAB_Almoxarifado).Enabled = False
'    Else
'        Frame1(TAB_Almoxarifado).Enabled = True
'    End If
'
'    If Not objGridAlocacoes Is Nothing Then Call Grid_Limpa(objGridAlocacoes)
'
'    If objTipoDocInfo.iCodigo = DOCINFO_NFFISPC Then
'        EscaninhoRastro.Enabled = False
'        gobjRastreamento.bEditavel = True
'    Else
'        EscaninhoRastro.Enabled = True
'        gobjRastreamento.bEditavel = False
'    End If
'
'    'Inicializa o grid de Rastreamento
'    lErro = gobjRastreamento.Inicializa_Grid_Rastreamento()
'    If lErro <> SUCESSO Then gError 83383
'
'    'se o cliente não estiver preenchido ==> preencher com a natureza padrao o campo natureza da operacao
'    If Len(Trim(Cliente.Text)) = 0 Then NatOpInterna.Text = objTipoDocInfo.sNaturezaOperacaoPadrao
'
'    ''*** incluidos p/tratamento de tributacao *******************************
'    If gbLimpandoTela = False Then Call gobjTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
'    '*** fim tributacao
'
'    'subtipo contabil
'    'ReInicializa a parte de contabilidade devido que o tipo implica na mudança dos modelos de contabilização
'    lErro = objContabil.Contabil_Processa_Alteracao_Subtipo(objTipoDocInfo.iSubTipoContabil)
'    If lErro <> SUCESSO Then gError 71690
'
'    Exit Sub
'
'Erro_TipoNFiscal_Click:
'
'    Select Case gErr
'
'        Case 42138, 71690, 83383
'
'        Case 42139
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub TipoNFiscal_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim iCodigo As Integer
'Dim objTipoDocInfo As New ClassTipoDocInfo
'
'On Error GoTo Erro_TipoNFiscal_Validate
'
'    'Verifica se o tipo está preenchido
'    If Len(Trim(TipoNFiscal.Text)) = 0 Then Exit Sub
'
'    'Verifica se foi selecionado
'    If TipoNFiscal.List(TipoNFiscal.ListIndex) = TipoNFiscal.Text Then Exit Sub
'
'    'Tenta selecionar na combo
'    lErro = Combo_Seleciona(TipoNFiscal, iCodigo)
'    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42140
'
'    'Se não conseguir --> Erro
'    If lErro <> SUCESSO Then gError 42141
'
'    Exit Sub
'
'Erro_TipoNFiscal_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42140
'
'        Case 42141
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub Cliente_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objCliente As New ClassCliente
'Dim iCodFilial As Integer
'Dim colCodigoNome As New AdmColCodigoNome
'Dim objTipoCliente As New ClassTipoCliente
'Dim objMensagem As New ClassMensagem
'
'On Error GoTo Erro_Cliente_Validate
'
'    'Verifica se o cliente foi alterado
'    If iClienteAlterado = 0 Then Exit Sub
'    'Se op cliente está preenchido
'    If Len(Trim(Cliente.Text)) > 0 Then
'
'        lErro = TP_Cliente_Le3(Cliente, objCliente, iCodFilial)
'        If lErro <> SUCESSO Then gError 42142
'
'        lErro = FiliaisClientes_Le_Cliente(objCliente, colCodigoNome)
'        If lErro <> SUCESSO Then gError 42143
'
'        'Preenche ComboBox de Filiais
'        Call Filial_Preenche(Filial, colCodigoNome)
'        Call Filial_Preenche(FilialEntrega, colCodigoNome)
'
'        If Not gbCarregandoTela Then
'
'            If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
'
'                If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
'
'                'Seleciona filial na Combo Filial
'                Call Filial_Seleciona(Filial, iCodFilial)
'                Call Filial_Seleciona(FilialEntrega, iCodFilial)
'
'            End If
'
'        End If
'
'        giFilialEntregaManual = 0
'
'        'Lê o tipo do cliente
'        If objCliente.iTipo > 0 Then
'            objTipoCliente.iCodigo = objCliente.iTipo
'            lErro = TipoCliente_Le(objTipoCliente)
'            If lErro <> SUCESSO And lErro <> 19062 Then gError 42144
'        End If
'        'Verifica se o cliente possui um desconto
'        If objCliente.dDesconto > 0 Then
'            gdDesconto = objCliente.dDesconto
'        ElseIf objTipoCliente.dDesconto > 0 Then
'            gdDesconto = objTipoCliente.dDesconto
'        Else
'            gdDesconto = 0
'        End If
'
'        If Not gbCarregandoTela Then
'
'            Call DescontoGlobal_Recalcula
'
'            Call ValorTotal_Calcula
'
'            'verifica se o cliente tem uma tabela default
'            If objCliente.iTabelaPreco > 0 Then
'                TabelaPreco.Text = objCliente.iTabelaPreco
'            ElseIf objTipoCliente.iTabelaPreco > 0 Then
'                TabelaPreco.Text = objTipoCliente.iTabelaPreco
'            End If
'
'            If Len(Trim(TabelaPreco.Text)) > 0 Then Call TabelaPreco_Validate(bSGECancelDummy)
'
'            If objCliente.iCondicaoPagto > 0 Then
'                CondicaoPagamento.Text = objCliente.iCondicaoPagto
'            ElseIf objTipoCliente.iCondicaoPagto > 0 Then
'                CondicaoPagamento.Text = objTipoCliente.iCondicaoPagto
'            End If
'
'            Call CondicaoPagamento_Validate(bSGECancelDummy)
'
'            If objCliente.iCodMensagem > 0 Then
'                objMensagem.iCodigo = objCliente.iCodMensagem
'                'Lê a Mensagem "default" p/ o Cliente
'                lErro = Mensagem_Le(objMensagem)
'                If lErro <> SUCESSO And lErro <> 19234 Then gError 42912
'
'                If lErro = SUCESSO Then Mensagem.Text = objMensagem.sDescricao
'
'            End If
'
'            'Tipo de frete
'            If objCliente.iTipoFrete = TIPO_FOB Then
'                Destinatario.Value = True
'            ElseIf objCliente.iTipoFrete = TIPO_CIF Then
'                Emitente.Value = True
'            End If
'
'        End If
'
'        iClienteAlterado = 0
'
'        If Not gbCarregandoTela Then
'
'    ''*** incluidos p/tratamento de tributacao *******************************
'            If iCodFilial <> 0 Then Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, iCodFilial) '####
'    '*** fim tributacao
'        End If
'
'    ElseIf Len(Trim(Cliente.Text)) = 0 Then 'Verifica preenchimento de Cliente
'
'        Filial.Clear 'Se não estiver preenchido limpa as combos de filialis
'        FilialEntrega.Clear
'
'    End If
'
'    iClienteAlterado = 0
'
'
'
'    Exit Sub
'
'Erro_Cliente_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 42142
'
'        Case 42143, 42144, 42912
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub Filial_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_Filial_Click
'
'    'Verifica se algo foi selecionada
'    If Filial.ListIndex = -1 Then Exit Sub
'
'    'Faz o tratamento da Filial selecionada
'    lErro = Trata_FilialCliente()
'    If lErro <> SUCESSO Then gError 42145
'
'    Exit Sub
'
'Erro_Filial_Click:
'
'    Select Case gErr
'
'        Case 42145
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Function Trata_FilialCliente() As Long
'
'Dim lErro As Long
'Dim objFilialCliente As New ClassFilialCliente
'Dim objCliente As New ClassCliente
'Dim objVendedor As New ClassVendedor
'Dim objTipoCliente As New ClassTipoCliente
'Dim iVendedorLeu As Integer
'Dim dValorTotal As Double
'Dim dValorBase As Double
'Dim objTransportadora As New ClassTransportadora
'Dim dValorComissao As Double
'Dim dValorEmissao As Double
'Dim dComissaoVendas As Double
'Dim objNFiscal As New ClassNFiscal
'
'On Error GoTo Erro_Trata_FilialCliente
'
'    objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
'    'Lê a FilialCliente
'    lErro = FilialCliente_Le_NomeRed_CodFilial(Cliente.Text, objFilialCliente)
'    If lErro <> SUCESSO Then gError 42146
'
'    objCliente.sNomeReduzido = Cliente.Text
'    'Lê o Cliente
'    lErro = Cliente_Le_NomeReduzido(objCliente)
'    If lErro <> SUCESSO And lErro <> 12348 Then gError 42147
'
'    'Se não achou o Cliente --> erro
'    If lErro = 12348 Then gError 42148
'
'    objNFiscal.lCliente = objCliente.lCodigo
'    objFilialCliente.iCodFilial = objFilialCliente.iCodFilial
'
'    lErro = Preenche_GridComprovServ(objNFiscal)
'    If lErro <> SUCESSO Then gError 94090
'
'    Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, objFilialCliente.iCodFilial)
'
'    If ComissaoAutomatica.Value = 1 Then
'
'        Call Grid_Limpa(objGridComissoes)
'
'        lErro = objTabComissoes.Comissao_Automatica_FilialCli_Exibe(objFilialCliente)
'        If lErro <> SUCESSO Then gError 59048
'
'    End If
'
'    'Tipo de frete
'    If objFilialCliente.iTipoFrete = TIPO_FOB Then
'        Destinatario.Value = True
'    ElseIf objFilialCliente.iTipoFrete = TIPO_CIF Then
'        Emitente.Value = True
'    End If
'
'    'Verififca se a transporadora está preenchida
'    If Len(Trim(Transportadora.Text)) = 0 Then
'
'        lErro = Transportadora_Padrao_FilCli(objFilialCliente, objTransportadora)
'        If lErro <> SUCESSO Then gError 59049
'
'        If objTransportadora.iCodigo <> 0 Then
'
'            'Coloca a transportadora na tela
'            Transportadora.Text = objTransportadora.sNomeReduzido
'            Call Transportadora_Validate(bSGECancelDummy)
'
'        End If
'
'    End If
'
'    If giFilialEntregaManual = 0 Then
'
'        FilialEntrega.Text = CStr(objFilialCliente.iCodFilial)
'        Call FilialEntrega_Validate(bSGECancelDummy)
'
'        giFilialEntregaManual = 0
'
'    End If
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
'        Case 42146, 42147, 42148, 59048, 59049, 94090
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub Filial_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim iCodigo As Integer
'Dim objFilialCliente As New ClassFilialCliente
'Dim sCliente As String
'Dim vbMsgRes As VbMsgBoxResult
'Dim objNFiscal As New ClassNFiscal
'
'On Error GoTo Erro_Filial_Validate
'
'    'Verifica se a filial foi preenchida
'    If Len(Trim(Filial.Text)) = 0 Then Exit Sub
'
'    'Verifica se é uma filial selecionada
'    If Filial.ListIndex > -1 Then Exit Sub
'
'    'Tenta selecionar na combo
'    lErro = Combo_Seleciona(Filial, iCodigo)
'    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42152
'
'    'Se nao encontra o ítem com o código informado
'    If lErro = 6730 Then
'
'        'Verifica de o Cliente foi digitado
'        If Len(Trim(Cliente.Text)) = 0 Then gError 42153
'
'        sCliente = Cliente.Text
'
'        objFilialCliente.iCodFilial = iCodigo
'
'        'Pesquisa se existe filial com o codigo extraido
'        lErro = FilialCliente_Le_NomeRed_CodFilial(sCliente, objFilialCliente)
'        If lErro <> SUCESSO And lErro <> 17660 Then gError 42154
'
'        If lErro = 17660 Then gError 42155
'
'        'Coloca na tela
'        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
'
'        lErro = Trata_FilialCliente()
'        If lErro <> SUCESSO Then gError 42156
'
'    End If
'
'    'Não encontrou valor informado que era STRING
'    If lErro = 6731 Then gError 42157
'
'
'
'    Exit Sub
'
'Erro_Filial_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42152, 42154
'
'        Case 42153
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
'
'        Case 42155
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)
'
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("FiliaisClientes", objFilialCliente)
'            End If
'
'        Case 42156
'
'        Case 42157
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub Serie_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objSerie As New ClassSerie
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_Serie_Validate
'
'    'Verifica se foi preenchida
'    If Len(Trim(Serie.Text)) = 0 Then Exit Sub
'    'Verifica se foi selecionada
'    If Serie.ListIndex <> -1 Then Exit Sub
'
'    'Tenta selecionar a serie
'    lErro = Combo_Item_Igual(Serie)
'    If lErro <> SUCESSO And lErro <> 12253 Then gError 42158
'    'Se não está na combo
'    If lErro <> SUCESSO Then
'
'        objSerie.sSerie = Serie.Text
'        'Busca a série no BD
'        lErro = Serie_Le(objSerie)
'        If lErro <> SUCESSO And lErro <> 22202 Then gError 51621
'        If lErro <> SUCESSO Then gError 42159 'Se não encontrou
'
'    End If
'
'    Exit Sub
'
'Erro_Serie_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42159
'            'Pergunta se deseja criar nova série
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_SERIE", Serie.Text)
'            'Se a resposta for afirmativa
'            If vbMsgRes = vbYes Then
'                'Chama a tela de cadastro de séries
'                Call Chama_Tela("SerieNFiscal", objSerie)
'            End If
'            'segura o foco na série
'
'        Case 42158, 51621
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'End Sub
'
'Public Sub DataEmissao_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_DataEmissao_Validate
'
'    'Verifica se a data de emissao foi digitada
'    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub
'
'    'Critica a data digitada
'    lErro = Data_Critica(DataEmissao.Text)
'    If lErro <> SUCESSO Then gError 42160
'
'    Exit Sub
'
'Erro_DataEmissao_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42160
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub UpDownEmissao_DownClick()
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_UpDownEmissao_DownClick
'
'    'Diminui a data em um dia
'    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
'    If lErro <> SUCESSO Then gError 42161
'
'    Exit Sub
'
'Erro_UpDownEmissao_DownClick:
'
'    Select Case gErr
'
'        Case 42161
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub UpDownEmissao_UpClick()
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_UpDownEmissao_UpClick
'
'    'Aumenta a data em um dia
'    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
'    If lErro <> SUCESSO Then gError 42162
'
'    Exit Sub
'
'Erro_UpDownEmissao_UpClick:
'
'    Select Case gErr
'
'        Case 42162
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub DataSaida_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_DataSaida_Validate
'
'    'Verifica se a data de Saida foi digitada
'    If Len(Trim(DataSaida.ClipText)) = 0 Then Exit Sub
'
'    'Critica a data digitada
'    lErro = Data_Critica(DataSaida.Text)
'    If lErro <> SUCESSO Then gError 42163
'
'    Exit Sub
'
'Erro_DataSaida_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'        Case 42163
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
''horasaida
'Public Sub HoraSaida_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(HoraSaida, iAlterado)
'
'End Sub
'
''horasaida
'Public Sub HoraSaida_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
''horasaida
'Public Sub HoraSaida_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_HoraSaida_Validate
'
'    'Verifica se a hora de saida foi digitada
'    If Len(Trim(HoraSaida.ClipText)) = 0 Then Exit Sub
'
'    'Critica a data digitada
'    lErro = Hora_Critica(HoraSaida.Text)
'    If lErro <> SUCESSO Then gError 89798
'
'    Exit Sub
'
'Erro_HoraSaida_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 89798
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub UpDownSaida_DownClick()
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_UpDownSaida_DownClick
'
'    'Diminui a data em um dia
'    lErro = Data_Up_Down_Click(DataSaida, DIMINUI_DATA)
'    If lErro Then gError 42164
'
'    Exit Sub
'
'Erro_UpDownSaida_DownClick:
'
'    Select Case gErr
'
'        Case 42164
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub UpDownSaida_UpClick()
'
'Dim lErro As Long
'Dim sData As String
'
'On Error GoTo Erro_UpDownSaida_UpClick
'
'    'Aumenta a data em um dia
'    lErro = Data_Up_Down_Click(DataSaida, AUMENTA_DATA)
'    If lErro <> SUCESSO Then gError 42165
'
'    Exit Sub
'
'Erro_UpDownSaida_UpClick:
'
'    Select Case gErr
'
'        Case 42165
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub TabelaPreco_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_TabelaPreco_Click
'
'    iAlterado = REGISTRO_ALTERADO
'
'    'Verifica se a Tabela de Preços foi informada
'    If TabelaPreco.ListIndex = -1 Then Exit Sub
'
'    'Verifica se existe algum Ítem no Grid
'    If objGridItens.iLinhasExistentes = 0 Then Exit Sub
'
'    'Faz o tratamento para a Tabela Preco selecionada
'    lErro = Trata_TabelaPreco()
'    If lErro <> SUCESSO Then gError 42166
'
'    Exit Sub
'
'Erro_TabelaPreco_Click:
'
'    Select Case gErr
'
'        Case 42166
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub TabelaPreco_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim vbMsgRes As VbMsgBoxResult
'Dim objTabelaPreco As New ClassTabelaPreco
'Dim iCodigo As Integer
'
'On Error GoTo Erro_TabelaPreco_Validate
'
'    'Verifica se foi preenchida a ComboBox TabelaPreco
'    If Len(Trim(TabelaPreco.Text)) = 0 Then Exit Sub
'
'    'Verifica se está preenchida com o ítem selecionado na ComboBox TabelaPreco
'    If TabelaPreco.Text = TabelaPreco.List(TabelaPreco.ListIndex) Then Exit Sub
'
'    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
'    lErro = Combo_Seleciona(TabelaPreco, iCodigo)
'    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42167
'
'    'Nao existe o ítem com o CÓDIGO na List da ComboBox
'    If lErro = 6730 Then
'
'        objTabelaPreco.iCodigo = iCodigo
'
'        'Tenta ler TabelaPreço com esse código no BD
'        lErro = TabelaPreco_Le(objTabelaPreco)
'        If lErro <> SUCESSO And lErro <> 28004 Then gError 42168
'
'        If lErro <> SUCESSO Then gError 42170 'Não encontrou Tabela Preço no BD
'
'        'Encontrou TabelaPreço no BD, coloca no Text da Combo
'        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao
'
'        lErro = Trata_TabelaPreco()
'        If lErro <> SUCESSO Then gError 42169
'
'    End If
'
'    'Não existe o ítem com a STRING na List da ComboBox
'    If lErro = 6731 Then gError 42171
'
'    Exit Sub
'
'Erro_TabelaPreco_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42167, 42168, 42169
'
'        Case 42170  'Não encontrou Tabela de Preço no BD
'
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")
'
'            If vbMsgRes = vbYes Then
'                'Chama a tela de Tabelas de Preço
'                Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
'            End If
'            'Segura o foco
'
'        Case 42171
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Function Trata_TabelaPreco() As Long
'
'Dim lErro As Long, iCodTabela As Integer
'Dim iLinha As Integer
'Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
'Dim dPrecoUnitario As Double
'Dim sProduto As String
'Dim iPreenchido As Integer
'
'On Error GoTo Erro_TabelaPreco_Click
'
'    If Not gbCarregandoTela Then
'
'        iCodTabela = Codigo_Extrai(TabelaPreco.Text)
'        For iLinha = 1 To objGridItens.iLinhasExistentes
'
'            'Verifica se o Produto está preenchido
'            lErro = Produto_Formata(GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
'            If lErro <> SUCESSO Then gError 42172
'
'            If iPreenchido <> PRODUTO_VAZIO And Len(Trim(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col))) > 0 Then
'
'                objTabelaPrecoItem.sCodProduto = sProduto
'                objTabelaPrecoItem.iCodTabela = iCodTabela
'                objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
'                'Lê a Tabela preço para filialEmpresa
'                lErro = TabelaPrecoItem_Le(objTabelaPrecoItem)
'                If lErro <> SUCESSO And lErro <> 28014 Then gError 42173
'
'                'Se não encontrar
'                If lErro = 28014 Then
'                    objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
'                    'Lê a Tabela de Preço a nível de Empresa toda
'                    lErro = TabelaPrecoItem_Le(objTabelaPrecoItem)
'                    If lErro <> SUCESSO And lErro <> 28014 Then gError 42174
'
'                End If
'
'                'Se  conseguir ler a Tabela de Preços
'                If lErro = SUCESSO Then
'                    'Calcula o Preco Unitário do Ítem
'                    lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
'                    If lErro <> SUCESSO Then gError 42175
'                    'Coloca no Grid
'                    If dPrecoUnitario > 0 Then
'                        GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
'
'                    Else
'                        GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col) = ""
'
'                    End If
'                    'Calcula o Preco Total do Ítem
'                    Call PrecoTotal_Calcula(iLinha, False)
'
'                    ''*** incluidos p/tratamento de tributacao *******************************
'                    lErro = gobjTribTab.Alteracao_Item_Grid(iLinha)
'                    If lErro <> SUCESSO Then gError 56859
'                    '*** fim tributacao
'
'                End If
'
'            End If
'        Next
'
'        'Calcula o VAlor Total da Nota
'        Call ValorTotal_Calcula
'    End If
'
'    Trata_TabelaPreco = SUCESSO
'
'    Exit Function
'
'Erro_TabelaPreco_Click:
'
'    Trata_TabelaPreco = gErr
'
'    Select Case gErr
'
'        Case 42172, 42173, 42174, 42175, 56859
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub PrecoTotal_Calcula(iLinha As Integer, Optional iCalculaTotal = True)
'
'Dim dPrecoTotal As Double
'Dim dPrecoTotalReal As Double
'Dim dPrecoUnitario As Double
'Dim dQuantidade As Double
'Dim dDesconto As Double
'Dim dPercentDesc As Double
'Dim lTamanho As Long
'Dim dValorTotal As Double
'Dim iIndice As Integer
'Dim dValorProdutos As Double
'Dim dValorDescontoGlobal As Double
'Dim vbMsgRes As VbMsgBoxResult
'Dim lErro As Long
'
'On Error GoTo Erro_PrecoTotal_Calcula
'
'    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col))
'    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
'
'    'Recolhe os valores Quantidade, Desconto, PerDesc e Valor Unitário da tela
'    If dPrecoUnitario = 0 Or dQuantidade = 0 Then
'        GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = ""
'        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
'    Else
'        dPrecoTotal = dPrecoUnitario * dQuantidade
'        dDesconto = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col))
'
'        dPercentDesc = PercentParaDbl(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col))
'
'        'Calcula o Valor Real
'        Call ValorReal_Calcula(dQuantidade, dPrecoUnitario, dPercentDesc, dDesconto, dPrecoTotalReal)
'
'        'Coloca o Desconto calculado na tela
'        If dDesconto > 0 Then
'            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
'        Else
'            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
'        End If
'
'        'Coloca o valor Real em Valor Total
'        GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(dPrecoTotalReal, "Standard")
'
'    End If
'
'    Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)
'
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
'
'    'Coloca o valor dos produtos na tela
'    ValorProdutos.Caption = Format(dValorProdutos, "Standard")
'
'    If iCalculaTotal <> False Then Call ValorTotal_Calcula
'
'    Exit Sub
'
'Erro_PrecoTotal_Calcula:
'
'    Select Case gErr
'
'        Case 56883
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Function ValorTotal_Calcula() As Long
'
'Dim dICMSSubstValor1 As Double
'Dim dIPIValor1 As Double
'Dim dValorProdutos As Double
'Dim dValorTotal As Double
'Dim dValorFrete As Double
'Dim dValorSeguro As Double
'Dim dValorDespesas As Double
'Dim vbMsgRes As VbMsgBoxResult
'Dim dValorIRRF As Double
'Dim dValorAposIR As Double
'Dim lErro As Long
'
'On Error GoTo Erro_ValorTotal_Calcula
'
'    If Not gbCarregandoTela Then
'        'Faz a atualização dos valores da tributação
'        lErro = gobjTribTab.AtualizarTributacao()
'        If lErro <> SUCESSO Then gError 42176
'    End If
'
'    'Recolhe os valores da tela
'    If Len(Trim(ValorProdutos.Caption)) > 0 And IsNumeric(ValorProdutos.Caption) Then dValorProdutos = CDbl(ValorProdutos.Caption)
'    If Len(Trim(ICMSSubstValor1.Caption)) > 0 And IsNumeric(ICMSSubstValor1.Caption) Then dICMSSubstValor1 = CDbl(ICMSSubstValor1.Caption)
'    If Len(Trim(IPIValor1.Caption)) And IsNumeric(IPIValor1.Caption) Then dIPIValor1 = CDbl(IPIValor1.Caption)
'    If Len(Trim(ValorFrete.Text)) > 0 And IsNumeric(ValorFrete.Text) Then dValorFrete = CDbl(ValorFrete.Text)
'    If Len(Trim(ValorSeguro.Text)) > 0 And IsNumeric(ValorSeguro.Text) Then dValorSeguro = CDbl(ValorSeguro.Text)
'    If Len(Trim(ValorDespesas.ClipText)) > 0 And IsNumeric(ValorDespesas.Text) Then dValorDespesas = CDbl(ValorDespesas.Text)
'    If Len(Trim(ValorIRRF.Text)) > 0 And IsNumeric(ValorIRRF.Text) Then dValorIRRF = CDbl(ValorIRRF.Text)
'
'    'Calcula o valor total
'
'    dValorTotal = dValorProdutos + dICMSSubstValor1 + dIPIValor1 + dValorFrete + dValorSeguro + dValorDespesas
'
'    dValorAposIR = dValorTotal - dValorIRRF
'
'    If dValorAposIR < 0 And dValorIRRF > 0 Then
'        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_IR_FONTE_MAIOR_VALOR_TOTAL", dValorIRRF, dValorTotal)
'        ValorIRRF.Text = ""
'
'        Call gobjTribTab.ValorIRRF_Validate(bSGECancelDummy)
'
'        'Faz a atualização dos valores da tributação
'        lErro = gobjTribTab.AtualizarTributacao()
'        If lErro <> SUCESSO Then gError 56884
'
'    End If
'
'    'Coloca o valor total na tela
'    '???? Shirley
''    dValorTotal = dValorTotal + StrParaDbl(TotalComprovServ.Caption)
'    ValorTotal.Caption = Format(dValorTotal, "Standard")
'
'    If Not gbCarregandoTela Then
'        'Gera a cobranca em cima do novo valor total
'        lErro = Cobranca_Automatica()
'        If lErro <> SUCESSO Then gError 56908
'
'        'Faz o cálculo automático das comissões
'        lErro = objTabComissoes.Comissoes_Automatica_Exibe()
'        If lErro <> SUCESSO Then gError 42177
'
'    End If
'
'    ValorTotal_Calcula = SUCESSO
'
'    Exit Function
'
'Erro_ValorTotal_Calcula:
'
'    ValorTotal_Calcula = gErr
'
'    Select Case gErr
'
'        Case 42176, 42177, 56883, 56884, 56908
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)
'
'Dim sProdutoFormatado As String
'Dim iProdutoPreenchido As Integer
'Dim objProduto As New ClassProduto
'Dim lErro As Long
'Dim objClasseUM As New ClassClasseUM
'Dim colSiglas As New Collection
'Dim objUM As ClassUnidadeDeMedida
'Dim sUM As String
'Dim iTipo As Integer
'Dim sCarteira As String
'Dim iIndice As Integer
'
'On Error GoTo Erro_Rotina_Grid_Enable
'
'    'Pesquisa a controle da coluna em questão
'    Select Case objControl.Name
'        'Produto
'        Case Produto.Name
'            lErro = Produto_Formata(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'            If lErro <> SUCESSO Then gError 42184
'
'            If iProdutoPreenchido = PRODUTO_VAZIO Then
'                Produto.Enabled = True
'            Else
'                Produto.Enabled = False
'            End If
'        'Unidade de Medida
'        Case UnidadeMed.Name
'
'            UnidadeMed.Clear
'
'            'Guarda a UM que está no Grid
'            sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
'
'            lErro = Produto_Formata(GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'            If lErro <> SUCESSO Then gError 42180
'
'            If iProdutoPreenchido = PRODUTO_VAZIO Then
'                UnidadeMed.Enabled = False
'            Else
'                UnidadeMed.Enabled = True
'
'                objProduto.sCodigo = sProdutoFormatado
'                'Lê o Produto
'                lErro = Produto_Le(objProduto)
'                If lErro <> SUCESSO And lErro <> 28030 Then gError 42178
'                If lErro = 28030 Then gError 42183 'Não achou
'
'                objClasseUM.iClasse = objProduto.iClasseUM
'                'Lâ as Unidades de Medidas da Classe do produto
'                lErro = UnidadesDeMedidas_Le_ClasseUM(objClasseUM, colSiglas)
'                If lErro <> SUCESSO Then gError 42179
'                'Carrega a combo de UM
'                For Each objUM In colSiglas
'                    UnidadeMed.AddItem objUM.sSigla
'                Next
'
'                'Seleciona na UM que está preenchida
'                UnidadeMed.Text = sUM
'                If Len(Trim(sUM)) > 0 Then
'                    lErro = Combo_Item_Igual(UnidadeMed)
'                    If lErro <> SUCESSO And lErro <> 12253 Then gError 42182
'                End If
'
'            End If
'
'        'Nas demais
'
'        Case Cobrador.Name
'
'            'Verifica se a data de vencimento OU a parcela foi(-i+ram) preenchida(s)
'            If Not ((Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_col))) > 0) Or (Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col))) > 0)) Then
'                objControl.Enabled = False
'            Else
'                objControl.Enabled = True
'            End If
'
'        Case CarteiraCobrador.Name
'
'            'Guarda a Carteira que está no Grid
'            sCarteira = GridParcelas.TextMatrix(iLinha, iGrid_CarteiraCobranca_Col)
'
'            objControl.Enabled = True
'
'            CarteiraCobrador.Clear
'            'Verifica se o cobrador foi preenchido
'
'            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Cobranca_Col))) > 0 Then
'                Call Carrega_CarteiraCobrador(Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Cobranca_Col)))
'            End If
'
'            'Seleciona na Combo
'            For iIndice = 0 To CarteiraCobrador.ListCount - 1
'                If CarteiraCobrador.List(iIndice) = sCarteira Then
'                    CarteiraCobrador.ListIndex = iIndice
'                    Exit For
'                End If
'            Next
'
'        Case PrecoUnitario.Name, PercentDesc.Name, Desconto.Name, Quantidade.Name, DescricaoItem.Name
'            'Verifica se o produto está preenchido
'            lErro = Produto_Formata(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'            If lErro <> SUCESSO Then gError 42181
'
'            If iProdutoPreenchido = PRODUTO_VAZIO Then
'                objControl.Enabled = False
'            Else
'                objControl.Enabled = True
'            End If
'
'        Case Desconto1Ate.Name, Desconto1Valor.Name, Desconto1Percentual.Name
'            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc1Codigo_Col))) = 0 Then
'                objControl.Enabled = False
'            Else
'                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc1Codigo_Col))
'                If objControl.Name = Desconto1Ate.Name Then
'                    objControl.Enabled = True
'                ElseIf objControl.Name = Desconto1Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
'                    Desconto1Valor.Enabled = True
'                ElseIf objControl.Name = Desconto1Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
'                    Desconto1Percentual.Enabled = True
'                Else
'                    objControl.Enabled = False
'                End If
'            End If
'
'        Case Desconto2Ate.Name, Desconto2Valor.Name, Desconto2Percentual.Name
'            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc2Codigo_Col))
'            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc2Codigo_Col))) = 0 Then
'                objControl.Enabled = False
'            Else
'                If objControl.Name = Desconto2Ate.Name Then
'                    objControl.Enabled = True
'                ElseIf objControl.Name = Desconto2Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
'                    Desconto2Valor.Enabled = True
'                ElseIf objControl.Name = Desconto2Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
'                    Desconto2Percentual.Enabled = True
'                Else
'                    objControl.Enabled = False
'                End If
'            End If
'
'        Case Desconto3Ate.Name, Desconto3Valor.Name, Desconto3Percentual.Name
'            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Desc3Codigo_Col))
'            If Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_Desc3Codigo_Col))) = 0 Then
'                objControl.Enabled = False
'            Else
'                If objControl.Name = Desconto3Ate.Name Then
'                    objControl.Enabled = True
'                ElseIf objControl.Name = Desconto3Valor.Name And (iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO) Then
'                    Desconto3Valor.Enabled = True
'                ElseIf objControl.Name = Desconto3Percentual.Name And (iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual) Then
'                    Desconto3Percentual.Enabled = True
'                Else
'                    objControl.Enabled = False
'                End If
'            End If
'
'        Case Desconto2Codigo.Name, Desconto3Codigo.Name
'
'            If Len(Trim(GridParcelas.TextMatrix(iLinha, GridParcelas.Col - 4))) = 0 Then
'                objControl.Enabled = False
'            Else
'                objControl.Enabled = True
'            End If
'
'        Case ValorParcela.Name
'            'Se o vencimento estiver preenchido, habilita o controle
'            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_col))) = 0 Then
'                objControl.Enabled = False
'            Else
'                objControl.Enabled = True
'            End If
'
'    End Select
'
'    'Rastreamento
'    lErro = gobjRastreamento.Rotina_Grid_Enable(iLinha, objControl, iCaminho)
'    If lErro <> SUCESSO Then gError 83386
'
'    Exit Sub
'
'Erro_Rotina_Grid_Enable:
'
'    Select Case gErr
'
'        Case 42178, 42179, 42180, 42181, 42182, 42184, 83386
'
'        Case 42183
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub GridItens_Click()
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Click(objGridItens, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridItens, iAlterado)
'    End If
'
'End Sub
'
'Public Sub GridItens_EnterCell()
'
'    Call Grid_Entrada_Celula(objGridItens, iAlterado)
'
'End Sub
'
'Public Sub GridItens_GotFocus()
'
'    Call Grid_Recebe_Foco(objGridItens)
'
'End Sub
'
'Public Sub GridItens_KeyPress(KeyAscii As Integer)
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)
'
'   If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridItens, iAlterado)
'    End If
'
'End Sub
'
'Public Sub GridItens_LeaveCell()
'
'    Call Saida_Celula(objGridItens)
'
'End Sub
'
'Public Sub GridItens_Validate(Cancel As Boolean)
'
'    Call Grid_Libera_Foco(objGridItens)
'
'End Sub
'
'Public Sub GridItens_RowColChange()
'
'    Call Grid_RowColChange(objGridItens)
'
'End Sub
'
'Public Sub GridItens_Scroll()
'
'    Call Grid_Scroll(objGridItens)
'
'End Sub
'
'Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
'
'Dim iLinhasExistentesAnterior As Integer
'Dim iItemAtual As Integer
'Dim iIndice As Integer
'Dim dValorProdutos As Double
'Dim lErro As Long
'
'On Error GoTo Erro_GridItens_KeyDown
'    'Guardo o item atual e o número de linhas existente
'    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
'    iItemAtual = GridItens.Row
'
'    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
'    'se alguma linha ofi excluída
'    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then
'        'Se existem alocacões no grid de alocações
'        If objGridAlocacoes.iLinhasExistentes > 0 Then
'            'Remove as alocações do item exluído
'            Call Remove_Alocacoes(iItemAtual)
'            'Atualiza o número do item nas alocações restantes
'            For iIndice = 1 To objGridAlocacoes.iLinhasExistentes
'                If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))) > 0 Then If CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) > iItemAtual Then GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col) = CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) - 1
'            Next
'        End If
'        'Recalcula o total da nota
'        Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)
'
'        ValorProdutos.Caption = Format(dValorProdutos, "Standard")
'
'''*** incluidos p/tratamento de tributacao *******************************
'
'        lErro = gobjTribTab.Exclusao_Item_Grid(iItemAtual)
'        If lErro <> SUCESSO Then gError 42185
'
''*** fim tributacao
'
'        Call ValorTotal_Calcula
'
'        'Recalcula o peso bruto e o peso líquido
'        lErro = Pesos_Recalcula()
'        If lErro <> SUCESSO Then gError 81517
'
'        'Rastreamento
'        lErro = gobjRastreamento.Exclusao_Item_Grid(iItemAtual)
'        If lErro <> SUCESSO Then gError 83387
'
'    End If
'
'    Exit Sub
'
'Erro_GridItens_KeyDown:
'
'    Select Case gErr
'
'        Case 42185, 81517, 83387
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub GridRastro_Click()
''Rastreamento
'
'    Call gobjRastreamento.GridRastro_Click
'
'End Sub
'
'Public Sub GridRastro_EnterCell()
''Rastreamento
'
'    Call gobjRastreamento.GridRastro_EnterCell
'
'End Sub
'
'Public Sub GridRastro_GotFocus()
''Rastreamento
'
'    Call gobjRastreamento.GridRastro_GotFocus
'
'End Sub
'
'Public Sub GridRastro_KeyPress(KeyAscii As Integer)
''Rastreamento
'
'    Call gobjRastreamento.GridRastro_KeyPress(KeyAscii)
'
'End Sub
'
'Public Sub GridRastro_LeaveCell()
''Rastreamento
'
'    Call gobjRastreamento.GridRastro_LeaveCell
'
'End Sub
'
'Public Sub GridRastro_Validate(Cancel As Boolean)
''Rastreamento
'
'    Call gobjRastreamento.GridRastro_Validate(Cancel)
'
'End Sub
'
'Public Sub GridRastro_RowColChange()
''Rastreamento
'
'    Call gobjRastreamento.GridRastro_RowColChange
'
'End Sub
'
'Public Sub GridRastro_Scroll()
''Rastreamento
'
'    Call gobjRastreamento.GridRastro_Scroll
'
'End Sub
'
'Public Sub GridRastro_KeyDown(KeyCode As Integer, Shift As Integer)
''Rastreamento
'
'    Call gobjRastreamento.GridRastro_KeyDown(KeyCode, Shift)
'
'End Sub
'
'Private Sub Remove_Alocacoes(iItemAtual As Integer)
''Remove do grid de alocações as alocações do item passado
'
'Dim iIndice As Integer
'Dim iItemLinha As Integer
'
'    iIndice = 1
'
'    Do While iIndice <= objGridAlocacoes.iLinhasExistentes
'        If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))) > 0 Then
'            'Recolhe o item a qual pertence a alocação
'            iItemLinha = GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)
'            'Se for o item passado em iItemAtual
'            If iItemLinha = iItemAtual Then
'                'Exclui do grid de alocações as alocações daquele ItemNF
'                Do While CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col)) = iItemAtual
'                    Call Grid_Exclui_Linha(objGridAlocacoes, iIndice)
'                    If Len(Trim(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))) = 0 Then Exit Do
'                Loop
'            End If
'        End If
'        iIndice = iIndice + 1
'    Loop
'
'    'Rastreamento
'    Call gobjRastreamento.Exclusao_Item_Grid(iItemAtual)
'
'End Sub
'
'Function Produto_Saida_Celula(Optional objProduto As ClassProduto) As Long
''Faz a crítica da célula Produto que está deixando de ser a corrente
'
'Dim lErro As Long
'Dim iProdutoPreenchido As Integer
'Dim iIndice As Integer
'Dim sProdutoEnxuto As String
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_Produto_Saida_Celula
'
'    'Critica o Produto
'    If objProduto Is Nothing Then
'        Set objProduto = New ClassProduto
'        lErro = Produto_Critica(Produto.Text, objProduto, iProdutoPreenchido)
'        If lErro <> SUCESSO And lErro <> 25041 Then gError 42186
'        If lErro <> SUCESSO Then gError 42191
'    End If
'
'    'Verifica se é de Faturamento
'    If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 42193
'
'    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
'    If lErro <> SUCESSO Then gError 42190
'
'    Produto.PromptInclude = False
'    Produto.Text = sProdutoEnxuto
'    Produto.PromptInclude = True
'
'    'Verifica se está no Grid
'    For iIndice = 1 To objGridItens.iLinhasExistentes
'        If iIndice <> GridItens.Row Then If GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 42192
'    Next
'
'    'Unidade de Medida
'    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda
'    'Descricao
'    GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objProduto.sDescricao
'
'    lErro = TabelaPreco_Atualiza_PrecoUnitario(objProduto)
'    If lErro <> SUCESSO Then gError 42915
'
'    If GridItens.Row - GridItens.FixedRows = objGridItens.iLinhasExistentes Then
'        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
'
'''*** incluidos p/tratamento de tributacao *******************************
'
'        'permite que a tributacao reflita a inclusao de uma linha no grid
'        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objProduto.sCodigo)
'        If lErro <> SUCESSO Then gError 27832
''*** fim tributacao
'    End If
'
'    Produto_Saida_Celula = SUCESSO
'
'    Exit Function
'
'Erro_Produto_Saida_Celula:
'
'    Produto_Saida_Celula = gErr
'
'    Select Case gErr
'
'        Case 42186, 42915
'
'        Case 42190
'            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
'
'        Case 42191
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
'            If vbMsgRes = vbYes Then
'
'                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)
'
'                Call Chama_Tela("Produto", objProduto)
'            Else
'                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
'            End If
'
'
'        Case 42192
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, Produto.Text, iIndice)
'
'        Case 42193
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_UM(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Unidade de Medida que está deixando de ser a corrente
'
'Dim lErro As Long
'Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
'Dim sProduto As String
'Dim iPreenchido As Integer
'Dim dPrecoUnitario As Double
'Dim sUmAnterior As String
'
'On Error GoTo Erro_Saida_Celula_UM
'
'    Set objGridInt.objControle = UnidadeMed
'    'recolhe a UM anteriormente escolhida
'    sUmAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
'    'coloca no grid a UM atual selecionda
'    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text
'    'se houve alteração na UM
'    If sUmAnterior <> UnidadeMed.Text And Frame1(TAB_Almoxarifado).Enabled = True Then
'        'Tenta fazer uma alocação p\ oitem de acordo com a nova UM
'        lErro = Alocacao_Processa()
'        If lErro <> SUCESSO Then gError 42195
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 42194
'
'    'Se a Um selecionada agora é diferente da anterior
'    If sUmAnterior <> UnidadeMed.Text Then
'
'        'Recalcula o peso bruto e o peso líquido
'        lErro = Pesos_Recalcula()
'        If lErro <> SUCESSO Then gError 500092
'
'    End If
'
'    Saida_Celula_UM = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_UM:
'
'    Saida_Celula_UM = gErr
'
'    Select Case gErr
'
'        Case 42194, 42195, 500092
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Pesos_Recalcula() As Long
'
'Dim lErro As Long
'Dim dFator As Double
'Dim iIndice As Integer
'Dim sProdutoFormatado As String
'Dim iProdutoPreenchido As Integer
'Dim objProduto As New ClassProduto
'Dim dPesoBruto As Double
'Dim dPesoLiquido As Double
'Dim dQuantidade As Double
'
'On Error GoTo Erro_Pesos_Recalcula
'
'    dPesoBruto = 0
'    dPesoLiquido = 0
'
'    'Para cada linha do Grid de Itens
'    For iIndice = 1 To objGridItens.iLinhasExistentes
'
'        'Se o produto e a UM estiverem preenchidos
'        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Produto_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) > 0 Then
'
'            'Formata o produto
'            lErro = Produto_Formata(GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'            If lErro <> SUCESSO Then gError 500093
'
'            'Lê o produto
'            objProduto.sCodigo = sProdutoFormatado
'            lErro = Produto_Le(objProduto)
'            If lErro <> SUCESSO And lErro <> 28030 Then gError 500094
'
'            'Se o produto não está cadastrado, erro
'            If lErro = 28030 Then gError 500096
'
'            'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
'            lErro = UM_Conversao_Trans(objProduto.iClasseUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), objProduto.sSiglaUMVenda, dFator)
'            If lErro <> SUCESSO Then gError 500095
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
'    PesoBruto.Text = Format(dPesoBruto, "Standard")
'    PesoLiquido.Text = Format(dPesoLiquido, "Standard")
'
'    Pesos_Recalcula = SUCESSO
'
'    Exit Function
'
'Erro_Pesos_Recalcula:
'
'    Pesos_Recalcula = gErr
'
'    Select Case gErr
'
'        Case 500093, 500094, 500095
'
'        Case 500096
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Quantidade que está deixando de ser a corrente
'
'Dim lErro As Long
'Dim dQuantidade As Double
'Dim iSairCelula As Integer
'Dim dPrecoUnitario As Double
'Dim dQuantidadeAnterior As Double
'
'On Error GoTo Erro_Saida_Celula_Quantidade
'
'    Set objGridInt.objControle = Quantidade
'
'    iSairCelula = False
'
'    'Verifica o preenchimento de Quantidade
'    If Len(Trim(Quantidade.ClipText)) > 0 Then
'        'VAlida a quantidae informada
'        lErro = Valor_Positivo_Critica(Quantidade.Text)
'        If lErro <> SUCESSO Then gError 42196
'        'Recolhe a quantidade informada
'        dQuantidade = CDbl(Quantidade.Text)
'        Quantidade = Formata_Estoque(dQuantidade)
'        'Recolhe a quantidade anteriormente digitada
'        dQuantidadeAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
'
'        If dQuantidadeAnterior = dQuantidade Then iSairCelula = True
'
'    End If
'
'    GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)
'
'
'    If (Not iSairCelula) And Frame1(TAB_Almoxarifado).Enabled = True Then
'        lErro = Alocacao_Processa()
'        If lErro <> SUCESSO Then gError 42198
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 42197
'
'    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))
'
'    If dPrecoUnitario > 0 Then Call PrecoTotal_Calcula(GridItens.Row, False)
'
'    '*** incluidos p/tratamento de tributacao ****************
'    If GridItens.Row <= objGridItens.iLinhasExistentes Then
'        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
'        If lErro <> SUCESSO Then gError 42199
'    End If
'    '*** fim tributacao
'
'    If Not iSairCelula Then
'
'        Call ValorTotal_Calcula
'
'        'Recalcula o peso bruto e o peso líquido
'        lErro = Pesos_Recalcula()
'        If lErro <> SUCESSO Then gError 81521
'
'    End If
'
'    Saida_Celula_Quantidade = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_Quantidade:
'
'    Saida_Celula_Quantidade = gErr
'
'    Select Case gErr
'
'        Case 42199
'
'        Case 42196, 42197, 42198, 81521
'            GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidadeAnterior)
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Alocacao_Processa() As Long
'
'Dim lErro As Long
'Dim sProduto As String
'Dim iPreenchido As Integer
'Dim objProduto As New ClassProduto
'Dim dQuantidade As Double
'Dim sUM As String
'
'On Error GoTo Erro_Alocacao_Processa
'
'    'Remove as alocações do item selecionado do grid de locações
'    Call Remove_Alocacoes(GridItens.Row)
'
'    lErro = Produto_Formata(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
'    If lErro <> SUCESSO Then gError 42200
'
'    objProduto.sCodigo = sProduto
'
'    'Lê o Produto
'    lErro = Produto_Le(objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then gError 42201
'
'    'Se não encontrou o Produto --> erro
'    If lErro = 28030 Then gError 42202
'
'    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
'
'   If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE And dQuantidade > 0 And gobjFAT.iNFiscalAlocacaoAutomatica = NFISCAL_ALOCA_AUTOMATICA Then
'
'        sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
'
'        lErro = AlocaAlmoxarifadoPradrao(dQuantidade, objProduto, sUM)
'        If lErro <> SUCESSO Then gError 42203
'
'    End If
'
'    Alocacao_Processa = SUCESSO
'
'    Exit Function
'
'Erro_Alocacao_Processa:
'
'    Alocacao_Processa = gErr
'
'    Select Case gErr
'
'        Case 42200, 42201, 42203
'
'        Case 42202
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function AlocaAlmoxarifadoPradrao(dQuantidade As Double, objProduto As ClassProduto, sUM As String) As Long
''Tenta fazer a alocação do produto no almoxarifado padrão. Caso não consiga chama a tela de Alocação de produto.
'
'Dim lErro As Long
'Dim dQuantAlocar As Double
'Dim dFator As Double
'Dim iAlmoxarifado As Integer
'Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim objEstoqueProduto As New ClassEstoqueProduto
'Dim dQuantDisponivel As Double
'Dim objItemNF As ClassItemNF
'Dim colOutrosProdutos As New Collection
'Dim sProduto As String
'Dim iPreenchido As Integer
'Dim iIndice As Integer
'Dim objItemNFAloc As ClassItemNFAlocacao
'Dim sProdutoEnxuto As String
'Dim iLinhas As Integer
'Dim dAcrescimo As Double
'Dim iNumCasasDec As Integer
'Dim dTotal As Double
'Dim iCodigo As Integer
'
'On Error GoTo Erro_AlocaAlmoxarifadoPradrao
'
'    'Faz a conversão da UM da tela para a UM de estoque
'    lErro = UM_Conversao(objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
'    If lErro <> SUCESSO Then gError 42204
'
'    'Converte a quantidade para a UM de estoque
'    dQuantAlocar = dQuantidade * dFator
'
'    'Busca o Almoxarifado padrão
'    lErro = AlmoxarifadoPadrao_Le(giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
'    If lErro <> SUCESSO And lErro <> 23796 Then gError 42205
'
'    'Se não encontrou
'    If iAlmoxarifado = 0 Then gError 51416
'
'    objAlmoxarifado.iCodigo = iAlmoxarifado
'    'Lê o Aloxarifado
'    lErro = Almoxarifado_Le(objAlmoxarifado)
'    If lErro <> 25056 And lErro <> SUCESSO Then gError 42206
'    If lErro = 25056 Then gError 42207
'
'    objEstoqueProduto.sProduto = objProduto.sCodigo
'    objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
'    'Lê os estoques desse produto nesse almoxarifado
'    lErro = EstoqueProduto_Le(objEstoqueProduto)
'    If lErro <> SUCESSO And lErro <> 21306 Then gError 42208
'    'Se não achou o Estoque do Produto --> erro
'    If lErro = 21306 Then gError 42209
'
'    iCodigo = Codigo_Extrai(TipoNFiscal.Text)
'
'    'Seleciona a origem da quantidade disponível
'    Select Case iCodigo
'
'        'Se o tipo da nota for cobrança de mat. consignado
'        Case DOCINFO_NFFISPC
'
'            'A quantidade disponível deve ser igual a quantidade no escaninho de mat. em Consignação (Consig)
'            dQuantDisponivel = objEstoqueProduto.dQuantConsig
'
'        'Se o tipo da nota for mat. beneficiado de 3º´s
'        Case DOCINFO_NFISFBF
'
'            'A quantidade disponível deve ser igual a quantidade no escaninho Mat.de 3º´s em Beneficiamento (Benef3)
'            dQuantDisponivel = objEstoqueProduto.dQuantBenef3
'
'        'Se for outro tipo de nota
'        Case Else
'
'            'A quantidade disponível deve ser igual a quantidade do escaninho mat. nosso disponível (DispNossa)
'            dQuantDisponivel = objEstoqueProduto.dQuantDisponivel
'
'    End Select
'
'    'Remove as alocações do produto do grid de alocações
'    Call Remove_Alocacoes(GridItens.Row)
'
'    iLinhas = objGridAlocacoes.iLinhasExistentes
'
'    'Calcula o número de casas decimais do Formato de Estoque
'    iNumCasasDec = Len(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
'    If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'    iCodigo = Codigo_Extrai(TipoNFiscal.Text)
'
'    If StrParaDbl(Formata_Estoque(dQuantAlocar)) < dQuantAlocar Then
'        dQuantAlocar = StrParaDbl(Formata_Estoque(dQuantAlocar)) + dAcrescimo
'    End If
'
'    'Verifica se a Quantidade disponível é maior que a quantidade a alocar
'    If dQuantAlocar <= dQuantDisponivel Then
'
'        iLinhas = iLinhas + 1
'        'Coloca no Grid alocações a alocação para o almoxarifado padrão
'        GridAlocacao.TextMatrix(iLinhas, iGrid_Item_Col) = GridItens.Row
'        GridAlocacao.TextMatrix(iLinhas, iGrid_ProdutoAloc_Col) = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
'        GridAlocacao.TextMatrix(iLinhas, iGrid_AlmoxAloc_Col) = objAlmoxarifado.sNomeReduzido
'        GridAlocacao.TextMatrix(iLinhas, iGrid_QuantAloc_Col) = Formata_Estoque(dQuantAlocar)
'        GridAlocacao.TextMatrix(iLinhas, iGrid_QuantVend_Col) = Formata_Estoque(dQuantAlocar)
'        GridAlocacao.TextMatrix(iLinhas, iGrid_UMAloc_Col) = objProduto.sSiglaUMEstoque
'
'    'Se não for
'    Else
'        'Se a função não estiver sendo usada pelo carregamento da tela
'        'se prepara para chamar a tela de localização
'        If Not gbCarregandoTela Then
'
'            Set objItemNF = New ClassItemNF
'            'Recolhe os dados do item
'            objItemNF.iItem = GridItens.Row
'            objItemNF.sProduto = objProduto.sCodigo
'            objItemNF.sDescricaoItem = objProduto.sDescricao
'            objItemNF.dQuantidade = dQuantidade
'            objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
'
'            'Recolhe todos os produtos dos outros itens
'            For iIndice = 1 To objGridItens.iLinhasExistentes
'                If iIndice <> GridItens.Row Then
'                    lErro = Produto_Formata(GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
'                    If lErro <> SUCESSO Then gError 42210
'                    'Adiciona na coleção de produtos
'                    colOutrosProdutos.Add sProduto
'                End If
'            Next
'
'            'Chama a tela de Localização de Produto
'            Call Chama_Tela_Modal("LocalizacaoProduto1", objItemNF, colOutrosProdutos, dQuantAlocar, iCodigo)
'            If giRetornoTela = vbCancel Then gError 42211
'            If giRetornoTela = vbOK Then
'
'                sProdutoEnxuto = String(STRING_PRODUTO, 0)
'
'                lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
'                If lErro <> SUCESSO Then gError 35820
'
'                Produto.PromptInclude = False
'                Produto.Text = sProdutoEnxuto
'                Produto.PromptInclude = True
'
'                'Calcula o Total Alocado
'                For Each objItemNFAloc In objItemNF.colAlocacoes
'                    dTotal = dTotal + objItemNFAloc.dQuantidade
'                Next
'
'                'Para cada alocação feita para o item
'                For Each objItemNFAloc In objItemNF.colAlocacoes
'
'                    iLinhas = iLinhas + 1
'                    'coloca no grid de alocações os dados das alocações
'                    GridAlocacao.TextMatrix(iLinhas, iGrid_Item_Col) = objItemNF.iItem
'                    GridAlocacao.TextMatrix(iLinhas, iGrid_ProdutoAloc_Col) = Produto.Text
'                    GridAlocacao.TextMatrix(iLinhas, iGrid_AlmoxAloc_Col) = objItemNFAloc.sAlmoxarifado
'                    GridAlocacao.TextMatrix(iLinhas, iGrid_QuantAloc_Col) = Formata_Estoque(objItemNFAloc.dQuantidade)
'                    GridAlocacao.TextMatrix(iLinhas, iGrid_QuantVend_Col) = Formata_Estoque(dTotal)
'                    GridAlocacao.TextMatrix(iLinhas, iGrid_UMAloc_Col) = objItemNF.sUMEstoque
'
'                Next
'
'                Quantidade.Text = Formata_Estoque(dTotal)
'
'                'Se o produto foi substituido
'                If objProduto.sCodigo <> objItemNF.sProduto Then
'                    'Coloca o produto no Grid de itens
'                    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
'                    'Chama a rotina que trata um novo produto
'                    lErro = Produto_Saida_Celula()
'                    objGridAlocacoes.iLinhasExistentes = iLinhas
'                    gError 42212
'                End If
'            End If
'        End If
'    End If
'
'    'atualiza o numero de linhas existentes no grid de alocações
'    objGridAlocacoes.iLinhasExistentes = iLinhas
'
'    AlocaAlmoxarifadoPradrao = SUCESSO
'
'    Exit Function
'
'Erro_AlocaAlmoxarifadoPradrao:
'
'    AlocaAlmoxarifadoPradrao = gErr
'
'    Select Case gErr
'
'        Case 35820
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)
'
'        Case 42204, 42208, 42205, 42206, 42210
'
'        Case 42207
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, iAlmoxarifado)
'
'        Case 42209
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
'
'        Case 42212
'
'        Case 42211
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_LOCALIZACAO", gErr, objProduto.sCodigo)
'
'        Case 51416
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO1", gErr, objProduto.sCodigo, giFilialEmpresa)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_PrecoUnitario(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Preço Unitário que está deixando de ser a corrente
'
'Dim lErro As Long
'Dim dQuantidade As Double
'Dim dQuantCancelada As Double
'Dim iValorUnitarioAlterado As Integer
'Dim iCodigo As Integer
'
'On Error GoTo Erro_Saida_Celula_PrecoUnitario
'
'    Set objGridInt.objControle = PrecoUnitario
'
'    'Verifica o preenchimento de Preço Unitário
'    If Len(Trim(PrecoUnitario.Text)) > 0 Then
'
'        iCodigo = Codigo_Extrai(TipoNFiscal.Text)
'
'        If iCodigo = TIPODOCINFO_COMPLEMENTO_IPI Then
'
'            lErro = Valor_NaoNegativo_Critica(PrecoUnitario.Text)
'            If lErro <> SUCESSO Then gError 87190
'
'        Else
'
'            lErro = Valor_Positivo_Critica(PrecoUnitario.Text)
'            If lErro <> SUCESSO Then gError 42214
'
'            PrecoUnitario.Text = Format(PrecoUnitario.Text, "Standard")
'
'        End If
'
'    End If
'
'    If PrecoUnitario.Text <> GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col) Then iValorUnitarioAlterado = True
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 42215
'
'    If iValorUnitarioAlterado Then
'
'        Call PrecoTotal_Calcula(GridItens.Row, False)
'
'    ''*** incluidos p/tratamento de tributacao *******************************
'        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
'        If lErro <> SUCESSO Then gError 42213
'    '*** fim tributacao
'
'        Call ValorTotal_Calcula
'
'    End If
'
'    Saida_Celula_PrecoUnitario = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_PrecoUnitario:
'
'    Saida_Celula_PrecoUnitario = gErr
'
'
'    Select Case gErr
'
'        Case 42213, 42214, 42215, 87190
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Saida_Celula_PercentDesc(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Percentual de Desconto que está deixando de ser a corrente
'
'Dim lErro As Long
'Dim dPercentDesc As Double
'Dim dPrecoUnitario As Double
'Dim dDesconto As Double
'Dim dValorTotal As Double
'Dim lTamanho As Long
'Dim dPercentDescAnterior As Double
'
'On Error GoTo Erro_Saida_Celula_PercentDesc
'
'    Set objGridInt.objControle = PercentDesc
'
'    'verifica se o percentual está preenchido
'    If Len(Trim(PercentDesc.Text)) > 0 Then
'        'Critica a procentagem
'        lErro = Porcentagem_Critica(PercentDesc.Text)
'        If lErro <> SUCESSO Then gError 42216
'
'        dPercentDesc = CDbl(PercentDesc.Text)
'
'        lTamanho = Len(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col))
'        If lTamanho > 0 Then dPercentDescAnterior = StrParaDbl(Left(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col), lTamanho - 1))
'
'        If dPercentDesc <> dPercentDescAnterior Then
'
'            'Verifica se o percentual é de 100%
'            If dPercentDesc = 100 Then gError 42218
'
'            PercentDesc.Text = Format(dPercentDesc, "Fixed")
'
'        End If
'    Else
'        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = ""
'    End If
'
'    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))) > 0 Then dPrecoUnitario = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'     If lErro <> SUCESSO Then gError 42217
'
'    If dPrecoUnitario > 0 Then
'
'        Call PrecoTotal_Calcula(GridItens.Row, False)
'
'''*** incluidos p/tratamento de tributacao *******************************
'        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
'        If lErro <> SUCESSO Then gError 39039
''*** fim tributacao
'
'        Call ValorTotal_Calcula
'
'    End If
'
'    Saida_Celula_PercentDesc = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_PercentDesc:
'
'    Saida_Celula_PercentDesc = gErr
'
'    Select Case gErr
'
'        Case 42216, 42217
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 42218
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Desconto que está deixando de ser a corrente
'
'Dim lErro As Long
'Dim dPrecoUnitario As Double
'Dim dQuantidade As Double
'Dim dPrecoTotal As Double
'Dim dDesconto As Double
'Dim dPercentDesc As Double
'
'On Error GoTo Erro_Saida_Celula_Desconto
'
'    Set objGridInt.objControle = Desconto
'    'Verifica se o o desconto foi digitado
'    If Len(Trim(Desconto.ClipText)) > 0 Then
'
'        'Critica o valor digitado
'        lErro = Valor_NaoNegativo_Critica(Desconto.Text)
'        If lErro <> SUCESSO Then gError 42219
'
'        dDesconto = CDbl(Desconto.Text)
'
'    End If
'
'    If dDesconto <> StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col)) Then
'        dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
'        dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col))
'        dPrecoTotal = dQuantidade * dPrecoUnitario
'        'Se o Preço total e positivo
'        If dPrecoTotal > 0 Then
'            'Verifica se o Valor do desconto é superior ao Preço Total
'            If dDesconto >= dPrecoTotal Then gError 42220
'
'            'Recalcula o percentual de desconto
'            dPercentDesc = dDesconto / dPrecoTotal
'
'            GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")
'        End If
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 42221
'
'    Call PrecoTotal_Calcula(GridItens.Row, False)
'
'''*** incluidos p/tratamento de tributacao *******************************
'    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
'    If lErro <> SUCESSO Then gError 42222
''*** fim tributacao
'
'    Call ValorTotal_Calcula
'
'    Saida_Celula_Desconto = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_Desconto:
'
'    Saida_Celula_Desconto = gErr
'
'    Select Case gErr
'
'        Case 42219, 42221
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 42220
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCONTO_MAIOR_OU_IGUAL_PRECO_TOTAL", gErr, GridItens.Row, dDesconto, dPrecoTotal)
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 42222
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub ValorFrete_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValorFrete As Double
'
'On Error GoTo Erro_ValorFrete_Validate
'
'    If iValorFreteAlterado = 0 Then Exit Sub
'
'    'Verifica se o Valor Frete foi preenchido
'    If Len(Trim(ValorFrete.Text)) > 0 Then
'
'        'Faz a crítica do valor
'        lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
'        If lErro <> SUCESSO Then gError 42223
'
'        dValorFrete = CDbl(ValorFrete.Text)
'
'        'Coloca o Valor Formatado na tela
'        ValorFrete.Text = Format(dValorFrete, "Fixed")
'
'    End If
'
'    iValorFreteAlterado = 0
'
'''*** incluidos p/tratamento de tributacao *******************************
'    Call gobjTribTab.ValorFrete_Validate(Cancel, dValorFrete)
''*** fim tributacao
'
'    Call ValorTotal_Calcula
'
'    Exit Sub
'
'Erro_ValorFrete_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42223
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ValorSeguro_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValorSeguro As Double
'
'On Error GoTo Erro_Valorseguro_Validate
'
'    If iValorSeguroAlterado = 0 Then Exit Sub
'
'    'Se estiver preenchido
'    If Len(Trim(ValorSeguro.Text)) > 0 Then
'
'        'Critica o valor
'        lErro = Valor_NaoNegativo_Critica(ValorSeguro.Text)
'        If lErro <> SUCESSO Then gError 42224
'
'        dValorSeguro = CDbl(ValorSeguro.Text)
'
'        'Coloca o valor formatado na tela
'        ValorSeguro.Text = Format(dValorSeguro, "Fixed")
'
'    End If
'
'''*** incluidos p/tratamento de tributacao *******************************
'    Call gobjTribTab.ValorSeguro_Validate(Cancel, dValorSeguro)
''*** fim tributacao
'
'    Call ValorTotal_Calcula
'
'    iValorSeguroAlterado = 0
'
'    Exit Sub
'
'Erro_Valorseguro_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42224
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ValorDespesas_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dValorDespesas As Double
'
'On Error GoTo Erro_ValorDespesas_Validate
'
'    If iValorDespesasAlterado = 0 Then Exit Sub
'
'    'Se  estiver preenchido
'    If Len(Trim(ValorDespesas.Text)) > 0 Then
'
'        'Faz a crítica do valor
'        lErro = Valor_NaoNegativo_Critica(ValorDespesas.Text)
'        If lErro <> SUCESSO Then gError 42225
'
'        dValorDespesas = CDbl(ValorDespesas.Text)
'
'        'coloca o valor formatado na tela
'        ValorDespesas.Text = Format(dValorDespesas, "Fixed")
'
'    End If
'
'    iValorDespesasAlterado = 0
'
'''*** incluidos p/tratamento de tributacao *******************************
'    Call gobjTribTab.ValorDespesas_Validate(Cancel, dValorDespesas)
''*** fim tributacao
'
'    Call ValorTotal_Calcula
'
'    Exit Sub
'
'Erro_ValorDespesas_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42225
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub ValorDesconto_Validate(Cancel As Boolean)
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
'    Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)
'
'    'Verifica se o Valor está preenchido
'    If Len(Trim(ValorDesconto.Text)) > 0 Then
'
'        'Faz a Crítica do Valor digitado
'        lErro = Valor_NaoNegativo_Critica(ValorDesconto.Text)
'        If lErro <> SUCESSO Then gError 42226
'
'        dValorDesconto = CDbl(ValorDesconto.Text)
'        'Coloca o Valor formatado na tela
'        ValorDesconto.Text = Format(dValorDesconto, "Standard")
'
'        'Verifica se o Valor do Desconto é maior que a soma dos demais valores
'        If dValorDesconto > dValorProdutos Then gError 42227
'
'        dValorProdutos = dValorProdutos - dValorDesconto
'
'    End If
'
'    ValorProdutos.Caption = Format(dValorProdutos, "Standard")
'
'''*** incluidos p/tratamento de tributacao *******************************
'    Call gobjTribTab.ValorDesconto_Validate(Cancel, dValorDesconto)
''*** fim tributacao
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
'
'    Select Case gErr
'
'        Case 42226
'
'        Case 42227
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub TabStrip1_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_TabStrip1_Click
'
'    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
'    If TabStrip1.SelectedItem.Index <> iFrameAtual Then
'
'        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
'
'''*** incluidos p/tratamento de tributacao *******************************
'        'se abriu o tab de tributacao
'        If TabStrip1.SelectedItem.Index = TAB_Tributacao Then
'            lErro = gobjTribTab.TabClick
'            If lErro <> SUCESSO Then gError 42228
'        End If
''*** fim tributacao
'
'        'Torna Frame correspondente ao Tab selecionado visivel
'        Frame1(TabStrip1.SelectedItem.Index).Visible = True
'        'Torna Frame atual visivel
'        Frame1(iFrameAtual).Visible = False
'        'Armazena novo valor de iFrameAtual
'        iFrameAtual = TabStrip1.SelectedItem.Index
'
'        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
'        If UCase(Left(TabStrip1.SelectedItem.Caption, 6)) = UCase(Left(TITULO_TAB_CONTABILIDADE, 6)) Then Call objContabil.Contabil_Carga_Modelo_Padrao
'
'        Select Case iFrameAtual
'
'            Case TAB_DadosPrincipais
'                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_DADOS_PRINCIPAIS
'
'            Case TAB_Itens
'                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_ITENS
'
'            Case TAB_Complemento
'                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_COMPLEMENTO
'
'            Case TAB_Cobranca
'                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_COBRANCA
'
'            Case TAB_Comissoes
'                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_COMISSOES
'
'            Case TAB_Almoxarifado
'                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_ALMOXARIFADO
'
'            Case TAB_Tributacao
'                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_TRIBUTACAO
'
'            Case TAB_Contabilizacao
'                Parent.HelpContextID = IDH_NF_FATURA_SAIDA_CONTABILIZACAO
'
'        End Select
'
'    End If
'
'    Exit Sub
'
'Erro_TabStrip1_Click:
'
'    Select Case gErr
'
'        Case 42228
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub FilialEntrega_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim iCodigo As Integer
'Dim objFilialCliente As New ClassFilialCliente
'Dim sCliente As String
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_FilialEntrega_Validate
'
'    'Verifica se a FilialEntrega foi preenchida
'    If Len(Trim(FilialEntrega.Text)) = 0 Or giFilialEntregaAlterada = 0 Then Exit Sub
'
'    'Verifica se é uma FilialEntrega selecionada
'    If FilialEntrega.ListIndex <> -1 Then Exit Sub
'
'    'Tenta selecionar na combo
'    lErro = Combo_Seleciona(FilialEntrega, iCodigo)
'    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42229
'
'    'Se nao encontra o ítem com o código informado
'    If lErro = 6730 Then
'
'        'Verifica de o Cliente foi digitado
'        If Len(Trim(Cliente.Text)) = 0 Then gError 42232
'
'        sCliente = Cliente.Text
'
'        objFilialCliente.iCodFilial = iCodigo
'
'        'Pesquisa se existe FilialEntrega com o codigo extraido
'        lErro = FilialCliente_Le_NomeRed_CodFilial(sCliente, objFilialCliente)
'        If lErro <> SUCESSO And lErro <> 17660 Then gError 42230
'
'        If lErro = 17660 Then gError 42231
'
'        'coloca na tela
'        FilialEntrega.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
'
'    End If
'
'    'Não encontrou valor informado que era STRING
'    If lErro = 6731 Then gError 42233
'
'    giFilialEntregaAlterada = 0
'
'    Exit Sub
'
'Erro_FilialEntrega_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42229, 42230
'
'        Case 42231
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)
'
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("FiliaisClientes", objFilialCliente)
'            End If
'
'        Case 42232
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
'
'        Case 42233
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialEntrega.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub Transportadora_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objTransportadora As New ClassTransportadora
'Dim iCodigo As Integer
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_Transportadora_Validate
'
'    'Verifica se foi preenchida a ComboBox Transportadora
'    If Len(Trim(Transportadora.Text)) = 0 Then Exit Sub
'    'Verifica se está preenchida com o ítem selecionado na ComboBox Transportadora
'    If Transportadora.Text = Transportadora.List(Transportadora.ListIndex) Then Exit Sub
'
'    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
'    lErro = Combo_Seleciona(Transportadora, iCodigo)
'    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42234
'    'Nao existe o ítem com o CÓDIGO na List da ComboBox
'    If lErro = 6730 Then
'
'        objTransportadora.iCodigo = iCodigo
'        'Tenta ler Transportadora com esse código no BD
'        lErro = Transportadora_Le(objTransportadora)
'        If lErro <> SUCESSO And lErro <> 19250 Then gError 42235
'
'        'Não encontrou Transportadora no BD
'        If lErro = 19250 Then gError 42236
'
'        'Encontrou Transportadora no BD, coloca no Text da Combo
'        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido
'
'    End If
'
'    'Não existe o ítem com a String na List da Combo
'    If lErro = 6731 Then gError 42237
'
'    Exit Sub
'
'Erro_Transportadora_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42234, 42235
'
'        Case 42236
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")
'
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("Transportadora", objTransportadora)
'            End If
'
'        Case 42237
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub PlacaUF_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_PlacaUF_Validate
'
'    'verifica se tem alguma Coisa preenchida
'    If Len(Trim(PlacaUF.Text)) = 0 Then Exit Sub
'
'    'Verifica se existe o ítem na combo
'    lErro = Combo_Item_Igual(PlacaUF)
'    If lErro <> SUCESSO And lErro <> 12253 Then gError 42238
'
'    'Se não encontrar --> Erro
'    If lErro = 12253 Then gError 42239
'
'    Exit Sub
'
'Erro_PlacaUF_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42238
'
'        Case 42239
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub PesoBruto_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_PesoBruto_Validate
'
'    'Verifica se está preenchido
'    If Len(Trim(PesoBruto.ClipText)) = 0 Then Exit Sub
'
'    'Criticao valor
'    lErro = Valor_Positivo_Critica(PesoBruto.Text)
'    If lErro <> SUCESSO Then gError 42240
'
'    'Coloca o valor formatado na Tela
'    PesoBruto.Text = Format(PesoBruto.Text, "Fixed")
'
'    Exit Sub
'
'Erro_PesoBruto_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42240
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub PesoLiquido_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_PesoLiquido_Validate
'
'    'Verifica se foi preenchido
'    If Len(Trim(PesoLiquido.ClipText)) = 0 Then Exit Sub
'
'    'Critica o valor
'    lErro = Valor_NaoNegativo_Critica(PesoLiquido.Text)
'    If lErro <> SUCESSO Then gError 42241
'
'    'Coloca o valor formatado na Tela
'    PesoLiquido.Text = Format(PesoLiquido.Text, "Fixed")
'
'    Exit Sub
'
'Erro_PesoLiquido_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42241
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub CanalVenda_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim vbMsgRes As VbMsgBoxResult
'Dim objCanalVenda As New ClassCanalVenda
'Dim iCodigo As Integer
'
'On Error GoTo Erro_CanalVenda_Validate
'
'    'Verifica se foi preenchida a ComboBox CanalVenda
'    If Len(Trim(CanalVenda.Text)) = 0 Then Exit Sub
'
'    'Verifica se está preenchida com o ítem selecionado na ComboBox CanalVenda
'    If CanalVenda.Text = CanalVenda.List(CanalVenda.ListIndex) Then Exit Sub
'
'    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
'    lErro = Combo_Seleciona(CanalVenda, iCodigo)
'    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42242
'
'    'Nao existe o ítem com o CÓDIGO na List da ComboBox
'    If lErro = 6730 Then
'
'        objCanalVenda.iCodigo = iCodigo
'
'        'Tenta ler CanalVenda com esse código no BD
'        lErro = CanalVenda_Le(objCanalVenda)
'        If lErro <> SUCESSO And lErro <> 23597 Then gError 42243
'
'        'Não encontrou CanalVenda no BD
'        If lErro <> SUCESSO Then gError 42244
'
'        'Encontrou CanalVenda no BD, coloca no Text da Combo
'        CanalVenda.Text = CStr(objCanalVenda.iCodigo) & SEPARADOR & objCanalVenda.sNome
'
'    End If
'
'    'Não existe o ítem com a STRING na List da ComboBox
'    If lErro = 6731 Then gError 42245
'
'    Exit Sub
'
'Erro_CanalVenda_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42242, 42243
'
'
'        Case 42244  'Não encontrou CanalVenda no BD
'
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CANALVENDA", iCodigo)
'
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("CanalDeVenda", objCanalVenda)
'            End If
'            'Segura o foco
'
'        Case 42245
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CANALVENDA_NAO_ENCONTRADO", gErr, CanalVenda.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub SerieNFiscalOriginal_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim lTamanho As Long
'
'On Error GoTo Erro_SerieNFiscalOriginal_Validate
'
'    'Verifica se o tamanho da Serie é maior que 3
'    If Len(Trim(SerieNFiscalOriginal.Text)) > 3 Then gError 42246
'
'    Exit Sub
'
'Erro_SerieNFiscalOriginal_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'        Case 42246
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_SERIE", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub CondicaoPagamento_Click()
'
'Dim lErro As Long
'Dim objCondicaoPagto As New ClassCondicaoPagto
'Dim vbMsgRes As VbMsgBoxResult
'Dim dPercAcresFin As Double
'
'On Error GoTo Erro_CondicaoPagamento_Click
'
'    'Verifica se alguma Condição foi selecionada
'    If CondicaoPagamento.ListIndex = -1 Then Exit Sub
'
'    'Passa o código da Condição para objCondicaoPagto
'    objCondicaoPagto.iCodigo = CondicaoPagamento.ItemData(CondicaoPagamento.ListIndex)
'
'    'Lê Condição de Pagamento à partir do código
'    lErro = CondicaoPagto_Le(objCondicaoPagto)
'    If lErro <> SUCESSO And lErro <> 19205 Then gError 42247
'
'    'Não encontrou a Condição de Pagamento --> erro
'    If lErro = 19205 Then gError 42248
'
'    dPercAcresFin = StrParaDbl(PercAcrescFin.Text)
'
'    If Len(Trim(PercAcrescFin.ClipText)) > 0 And dPercAcresFin <> objCondicaoPagto.dAcrescimoFinanceiro * 100 Then
'
'        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_SUBSTITUICAO_PERC_ACRESCIMO_FINANCEIRO")
'        If vbMsgRes = vbYes Then
'            PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
'            Call PercAcrescFin_Validate(bSGECancelDummy)
'        End If
'    Else
'        PercAcrescFin.Text = Format(objCondicaoPagto.dAcrescimoFinanceiro * 100, "Fixed")
'        Call PercAcrescFin_Validate(bSGECancelDummy)
'    End If
'
'    'Testa se ValorTotal está preenchido
'    If Len(Trim(ValorTotal)) > 0 Then
'
'        'Testa se DataEmissao está preenchida e ValorTotal é positivo
'        If Len(Trim(DataReferencia.ClipText)) > 0 And (CDbl(ValorTotal.Caption) > 0) Then
'
'            'Preenche o GridParcelas
'            lErro = Cobranca_Automatica()
'            If lErro <> SUCESSO Then gError 42249
'
'        End If
'    End If
'
'    iAlterado = REGISTRO_ALTERADO
'
'    Exit Sub
'
'Erro_CondicaoPagamento_Click:
'
'    Select Case gErr
'
'        Case 42247, 42249
'
'        Case 42248
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'      End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub CondicaoPagamento_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim iCodigo As Integer
'Dim objCondicaoPagto As New ClassCondicaoPagto
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_Condicaopagamento_Validate
'
'    'Verifica se a Condicaopagamento foi preenchida
'    If Len(Trim(CondicaoPagamento.Text)) = 0 Then Exit Sub
'
'    'Verifica se é uma Condicaopagamento selecionada
'    If CondicaoPagamento.Text = CondicaoPagamento.List(CondicaoPagamento.ListIndex) Then Exit Sub
'
'    'Tenta selecionar na combo
'    lErro = Combo_Seleciona(CondicaoPagamento, iCodigo)
'    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 42250
'
'    'Se não encontra valor que contém CÓDIGO, mas extrai o código
'    If lErro = 6730 Then
'
'        objCondicaoPagto.iCodigo = iCodigo
'
'        'Lê Condição Pagamento no BD
'        lErro = CondicaoPagto_Le(objCondicaoPagto)
'        If lErro <> SUCESSO And lErro <> 19205 Then gError 42251
'
'        'Não encontrou a Condição de Pagamento
'        If lErro = 19205 Then gError 42252
'
'        'Testa se pode ser usada em Contas a Receber
'        If objCondicaoPagto.iEmRecebimento = 0 Then gError 42253
'
'        'Coloca na Tela
'        CondicaoPagamento.Text = iCodigo & SEPARADOR & objCondicaoPagto.sDescReduzida
'
'        'Se ValorTotal e DataEmissao estiverem preenchidos, preenche GridParcelas
'        If Len(Trim(ValorTotal)) > 0 Then
'            If Len(Trim(DataReferencia.ClipText)) > 0 And CDbl(ValorTotal.Caption) > 0 Then
'
'                'Preenche o GridParcelas
'                lErro = Cobranca_Automatica()
'                If lErro <> SUCESSO Then gError 42254
'
'            End If
'        End If
'
'    End If
'
'    'Não encontrou o valor que era STRING
'    If lErro = 6731 Then gError 42255
'
'    Exit Sub
'
'Erro_Condicaopagamento_Validate:
'
'    Cancel = True
'
'
'    Select Case gErr
'
'       Case 42250, 42251, 42254
'
'       Case 42252
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)
'
'            If vbMsgRes = vbYes Then
'                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
'            End If
'
'        Case 42253
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", gErr, objCondicaoPagto.iCodigo)
'
'        Case 42255
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto) As Long
''Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas
'
'Dim lErro As Long
'Dim dValorPagar As Double
'Dim colValorParcelas As New Collection
'Dim colDataVencimento As New Collection
'Dim dtDataReferencia As Date
'Dim dValorIRRF As Double, dValorINSSRetido As Double
'Dim iIndice As Integer
'Dim dPercAcrescFin As Double
'Dim iTamanho As Integer
'Dim iPosicao As Integer
'Dim objFilialCliente As New ClassFilialCliente
'Dim iContador As Integer
'Dim objCliente As New ClassCliente
'
'On Error GoTo Erro_GridParcelas_Preenche
'
'    'Limpa o GridParcelas
'    Call Grid_Limpa(objGridParcelas)
'
'    'Número de Parcelas
'    objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas
'
'    dValorIRRF = StrParaDbl(ValorIRRF.Text)
'    If INSSRetido.Value = vbChecked Then
'        dValorINSSRetido = StrParaDbl(INSSValor.Text)
'    End If
'
'    'Valor a Pagar
'    dValorPagar = StrParaDbl(ValorTotal.Caption) - dValorIRRF - dValorINSSRetido
'
'    'Se Valor a Pagar for positivo
'    If dValorPagar > 0 Then
'
'        'Calcula os valores das Parcelas
'        lErro = Parcelas_Calcula(dValorPagar, objCondicaoPagto.iNumeroParcelas, colValorParcelas)
'        If lErro <> SUCESSO Then gError 42256
'
'        'Atribui o cliente e sua filial para futura leitura
'        'Verifica preenchimento de Cliente
'        If Len(Trim(Cliente.ClipText)) > 0 Then
'
'            objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
'            lErro = FilialCliente_Le_NomeRed_CodFilial(Cliente.Text, objFilialCliente)
'            If lErro <> SUCESSO Then gError 95205
'
''            objCliente.sNomeReduzido = Cliente.Text
''            'Lê o Cliente
''            lErro = Cliente_Le_Codigo_NomeReduzido(objCliente)
''            If lErro <> SUCESSO And lErro <> 12348 Then gError 95201
''
''            'Se não encontrou o Cliente --> erro
''            If lErro = 12348 Then gError 95202
''
''            objFilialCliente.lCodCliente = objCliente.lCodigo
''            objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
''
'        End If
'
''        If objFilialCliente.lCodCliente <> 0 And objFilialCliente.iCodFilial <> 0 Then
''
''            'Le o Cobrador Default do cliente selecionado
''            lErro = FilialCliente_Le(objFilialCliente)
''            If lErro <> SUCESSO Then gError 95200
''
'        If objFilialCliente.iCodCobrador = 0 Then objFilialCliente.iCodCobrador = COBRADOR_PROPRIA_EMPRESA
'
''        End If
'
'        'Coloca os valores das Parcelas no Grid Parcelas
'        For iIndice = 1 To objGridParcelas.iLinhasExistentes
'
'            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(colValorParcelas(iIndice), "Standard")
'
'            'Alteracao Daniel
'
'            For iContador = 0 To Cobrador.ListCount - 1
'                If Cobrador.ItemData(iContador) = objFilialCliente.iCodCobrador Then
'                    GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col) = Cobrador.List(iContador)
'                    Exit For
'                End If
'            Next
'
'            Carrega_CarteiraCobrador (objFilialCliente.iCodCobrador)
'            GridParcelas.TextMatrix(iIndice, iGrid_CarteiraCobranca_Col) = CarteiraCobrador.List(0)
'            'Fim da Alteracao Daniel em 11/10/2001
'
'        Next
'
'    End If
'
'    'Se Data Emissão estiver preenchida
'    If Len(Trim(DataReferencia.ClipText)) > 0 Then
'
'        dtDataReferencia = CDate(DataReferencia.Text)
'
'        'Calcula Datas de Vencimento das Parcelas
'        lErro = Parcelas_DatasVencimento(objCondicaoPagto, dtDataReferencia, colDataVencimento)
'        If lErro <> SUCESSO Then gError 42257
'
'        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
'        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas
'
'            'Coloca Data de Vencimento no Grid Parcelas
'            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col) = Format(colDataVencimento(iIndice), "dd/mm/yyyy")
'
'        Next
'
'    End If
'
'    For iIndice = 1 To objGridParcelas.iLinhasExistentes
'
'        lErro = Preenche_DescontoPadrao(iIndice)
'        If lErro <> SUCESSO Then gError 51117
'    Next
'    GridParcelas_Preenche = SUCESSO
'
'    Exit Function
'
'Erro_GridParcelas_Preenche:
'
'    GridParcelas_Preenche = gErr
'
'    Select Case gErr
'
'        Case 42256, 42257, 51117
'
'        '??? Daniel
'        Case 95200
'
'        Case 95201
'
'        Case 95202
'            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'End Function
'
'Private Function Saida_Celula_ValorParcela(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Valor da Parcela que está deixando de sser a corrente
'
'Dim lErro As Long
'Dim dColunaSoma As Double
'Dim iIndice As Integer
'Dim iColDescPerc As Integer
'Dim iColTipoDesconto As Integer
'Dim lTamanho As Long
'Dim dPercentual As Double
'Dim dValorParcela As Double
'Dim sValorDesconto As String
'Dim iTipoDesconto As Integer
'
'On Error GoTo Erro_Saida_Celula_ValorParcela
'
'    Set objGridInt.objControle = ValorParcela
'
'    'Verifica se valor está preenchido
'    If Len(ValorParcela.ClipText) > 0 Then
'
'        'Critica se valor é positivo
'        lErro = Valor_Positivo_Critica(ValorParcela.Text)
'        If lErro <> SUCESSO Then gError 42258
'
'        ValorParcela.Text = Format(ValorParcela.Text, "Standard")
'
'        If ValorParcela.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) Then
'
'            CobrancaAutomatica.Value = vbUnchecked
'
'            '***Código para colocar valores de desconto
'            dValorParcela = StrParaDbl(ValorParcela.Text)
'            If dValorParcela > 0 Then
'
'                'Vai varrer todos os 3 descontos para colocar valores
'                For iIndice = 1 To 3
'
'                    Select Case iIndice
'                        Case 1
'                            iColDescPerc = iGrid_Desc1Percentual_Col
'                            iColTipoDesconto = iGrid_Desc1Codigo_Col
'                        Case 2
'                            iColDescPerc = iGrid_Desc2Percentual_Col
'                            iColTipoDesconto = iGrid_Desc2Codigo_Col
'                        Case 3
'                            iColDescPerc = iGrid_Desc3Percentual_Col
'                            iColTipoDesconto = iGrid_Desc3Codigo_Col
'                    End Select
'
'                    iTipoDesconto = Codigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iColTipoDesconto))
'                    lTamanho = Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc)))
'
'                    'Coloca valor de desconto na tela
'                    If (iTipoDesconto = Percentual Or iTipoDesconto = PERC_ANT_DIA Or iTipoDesconto = PERC_ANT_DIA_UTIL) And lTamanho > 0 Then
'                        dPercentual = PercentParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc))
'                        sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
'                        GridParcelas.TextMatrix(GridParcelas.Row, iColDescPerc - 1) = sValorDesconto
'                    End If
'
'                Next
'
'            End If
'            '***Fim Código para colocar valores de desconto
'
'        End If
'
'        'Acrescenta uma linha no Grid se for o caso
'        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
'            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
'            'Coloca desconto padrao (le em CPRConfig)
'            lErro = Preenche_DescontoPadrao(GridParcelas.Row)
'            If lErro <> SUCESSO Then gError 51119
'
'        End If
'
'    Else
'        '***Código para colocar valores de desconto
'        'Limpa Valores de Desconto
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc1Valor_Col) = ""
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc2Valor_Col) = ""
'        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desc3Valor_Col) = ""
'        '***Fim Código para colocar valores de desconto
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 42259
'
'    Saida_Celula_ValorParcela = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_ValorParcela:
'
'    Saida_Celula_ValorParcela = gErr
'
'    Select Case gErr
'
'        Case 42258, 42259, 51119
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_TipoDesconto(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Tipo de Desconto que está deixando de sser a corrente
'
'Dim lErro As Long
'Dim iCodigo As Integer
'Dim iIndice As Integer
'Dim iTipo As Integer
'
'
'On Error GoTo Erro_Saida_Celula_TipoDesconto
'
'    If GridParcelas.Col = iGrid_Desc1Codigo_Col Then
'        Set objGridInt.objControle = Desconto1Codigo
'    ElseIf GridParcelas.Col = iGrid_Desc2Codigo_Col Then
'        Set objGridInt.objControle = Desconto2Codigo
'    ElseIf GridParcelas.Col = iGrid_Desc3Codigo_Col Then
'        Set objGridInt.objControle = Desconto3Codigo
'    End If
'
'    'Verifica se o Tipo foi preenchido
'    If Len(Trim(objGridInt.objControle.Text)) > 0 Then
'
'        'Verifica se ele foi selecionado
'        If objGridInt.objControle.Text <> objGridInt.objControle.List(objGridInt.objControle.ListIndex) Then
'
'            'Tenta selecioná-lo na combo
'            lErro = Combo_Seleciona_Grid(objGridInt.objControle, iCodigo)
'            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 42260
'
'            'Não foi encontrado
'            If lErro = 25085 Then gError 42262
'            If lErro = 25086 Then gError 42263
'
'        End If
'
'        iTipo = Codigo_Extrai(objGridInt.objControle.Text)
'
'        If (iTipo = VALOR_ANT_DIA) Or (iTipo = VALOR_ANT_DIA_UTIL) Or (iTipo = VALOR_FIXO) Then
'            GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 3) = ""
'        ElseIf iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual Then
'            '*** Acrescentado + 1 If para contabilizar com colocação de valores de desconto
'            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 3))) = 0 Then
'                GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col + 2) = ""
'            End If
'        End If
'
'        'Acrescenta uma linha no Grid se for o caso
'        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
'            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
'        End If
'
'    Else
'        For iIndice = objGridInt.objGrid.Col To iGrid_Desc3Percentual_Col
'            GridParcelas.TextMatrix(GridParcelas.Row, iIndice) = ""
'        Next
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 42261
'
'    Saida_Celula_TipoDesconto = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_TipoDesconto:
'
'    Saida_Celula_TipoDesconto = gErr
'
'    Select Case gErr
'
'        Case 42260, 42261
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 42262
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", gErr, iCodigo)
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 42263
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", gErr, objGridInt.objControle.Text)
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_DescontoData(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Desconto Data que está deixando de sser a corrente
'
'Dim lErro As Long
'Dim iCodigo As Integer
'Dim dtDataVencimento As Date
'
'On Error GoTo Erro_Saida_Celula_DescontoData
'
'    If GridParcelas.Col = iGrid_Desc1Ate_Col Then
'        Set objGridInt.objControle = Desconto1Ate
'    ElseIf GridParcelas.Col = iGrid_Desc2Ate_Col Then
'        Set objGridInt.objControle = Desconto2Ate
'    ElseIf GridParcelas.Col = iGrid_Desc3Ate_Col Then
'        Set objGridInt.objControle = Desconto3Ate
'    End If
'
'    If Len(Trim(objGridInt.objControle.ClipText)) > 0 Then
'
'        lErro = Data_Critica(objGridInt.objControle.Text)
'        If lErro <> SUCESSO Then gError 42264
'        'Se a data de vencimento estiver preenchida
'        If Len(Trim(DataEmissao.ClipText)) = 0 Then
'            'critica se DataDesconto ultrapassa DataVencimento
'            If CDate(objGridInt.objControle.Text) < CDate(DataEmissao.Text) Then gError 42266
'        End If
'
'        If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_col))) > 0 Then
'            dtDataVencimento = CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_col))
'            If CDate(objGridInt.objControle) > dtDataVencimento Then gError 42267
'        End If
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 42265
'
'    Saida_Celula_DescontoData = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_DescontoData:
'
'    Saida_Celula_DescontoData = gErr
'
'    Select Case gErr
'
'        Case 42264, 42265
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 42266
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_INFERIOR_DATA_EMISSAO", gErr, CDate(objGridInt.objControle.Text))
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 42267
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_SUPERIOR_DATA_VENCIMENTO", gErr, CDate(objGridInt.objControle.Text), dtDataVencimento)
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_DescontoValor(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Desconto Valor que está deixando de sser a corrente
'
'Dim lErro As Long
'Dim dColunaSoma As Double
'
'On Error GoTo Erro_Saida_Celula_DescontoValor
'
'    If GridParcelas.Col = iGrid_Desc1Valor_Col Then
'        Set objGridInt.objControle = Desconto1Valor
'    ElseIf GridParcelas.Col = iGrid_Desc2Valor_Col Then
'        Set objGridInt.objControle = Desconto2Valor
'    ElseIf GridParcelas.Col = iGrid_Desc3Valor_Col Then
'        Set objGridInt.objControle = Desconto3Valor
'    End If
'
'    'Verifica se valor está preenchido
'    If Len(objGridInt.objControle.ClipText) > 0 Then
'
'        'Critica se valor é positivo
'        lErro = Valor_Positivo_Critica(objGridInt.objControle.Text)
'        If lErro <> SUCESSO Then gError 42268
'
'        'Acrescenta uma linha no Grid se for o caso
'        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
'            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
'        End If
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 42269
'
'    Saida_Celula_DescontoValor = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_DescontoValor:
'
'    Saida_Celula_DescontoValor = gErr
'
'    Select Case gErr
'
'        Case 42268, 42269
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_DescontoPerc(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Desconto Percentual que está deixando de sser a corrente
'
'Dim lErro As Long
'Dim iCodigo As Integer
'Dim dPercentual As Double
'Dim dValorParcela As Double
'Dim sValorDesconto As String
'
'On Error GoTo Erro_Saida_Celula_DescontoPerc
'
'    If GridParcelas.Col = iGrid_Desc1Percentual_Col Then
'        Set objGridInt.objControle = Desconto1Percentual
'    ElseIf GridParcelas.Col = iGrid_Desc2Percentual_Col Then
'        Set objGridInt.objControle = Desconto2Percentual
'    ElseIf GridParcelas.Col = iGrid_Desc3Percentual_Col Then
'        Set objGridInt.objControle = Desconto3Percentual
'    End If
'
'    'Se a Porcentagem estiver preenchida
'    If Len(Trim(objGridInt.objControle.Text)) > 0 Then
'        'Critica porcentagem
'        lErro = Porcentagem_Critica(objGridInt.objControle.Text)
'        If lErro <> SUCESSO Then gError 42270
'
'
'        '***Código para colocar valores de desconto
'        dPercentual = CDbl(objGridInt.objControle.Text) / 100
'        dValorParcela = StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col))
'
'        'Coloca Valor do Desconto na tela
'        If dValorParcela > 0 Then
'            sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
'            GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col - 1) = sValorDesconto
'        End If
'
'    Else
'
'        'Limpa Valor de Desconto
'        GridParcelas.TextMatrix(GridParcelas.Row, GridParcelas.Col - 1) = ""
'        '***Fim Código para colocar valores de desconto
'
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 42271
'
'    Saida_Celula_DescontoPerc = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_DescontoPerc:
'
'    Saida_Celula_DescontoPerc = gErr
'
'    Select Case gErr
'
'        Case 42270, 42271
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Public Sub BotaoLocalizacao_Click()
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim iPreenchido As Integer
'Dim sProduto As String
'Dim dQuantidade As Double
'Dim ColItensNF As New ColItensNF
'Dim colAlocacoes As ColAlocacoesItemNF
'Dim objProduto As New ClassProduto
'Dim objItemNF As New ClassItemNF
'Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
'Dim lbLabel As Label
'Dim dPrecoUnitario As Double
'Dim objNFiscal As New ClassNFiscal
'
'On Error GoTo Erro_BotaoLocalizacao_Click:
'    'Verifica se há alguma linha selecionada
'    If objGridItens.iLinhasExistentes = 0 Then gError 43695
'    'Para cada ítem do grid
'    For iIndice = 1 To objGridItens.iLinhasExistentes
'        'Critica o produto
'        lErro = Produto_Critica(GridItens.TextMatrix(iIndice, iGrid_Produto_Col), objProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError 42309
'        'recolhe a quantidade
'        dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
'        'Se o produto fizer controle de estoque
'        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
'
'            Set colAlocacoes = New ColAlocacoesItemNF
'            'Recolhe as alocações do item
'            lErro = Recolhe_Alocacoes(iIndice, colAlocacoes)
'
'            ColItensNF.Add 0, iIndice, objProduto.sCodigo, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), dQuantidade, 0, 0, 0, DATA_NULA, GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col), 0, 0, 0, 0, 0, colAlocacoes, 0, "", "", 0, 0, "", 0, 0, 0, objProduto.sSiglaUMEstoque, 0
'        End If
'
'    Next
'    'Chama a tela de localização de produtos
'    Call Chama_Tela_Modal("LocalizacaoProduto", ColItensNF, Codigo_Extrai(TipoNFiscal.Text))
'    'Limpa o grid de Alocações
'    Call Grid_Limpa(objGridAlocacoes)
'    Call gobjRastreamento.Limpa_Tela_Rastreamento
'
'    'para cada item
'    For Each objItemNF In ColItensNF
'        'Se as UM de estoque e de venda do itemNF não estiverem preenchidas
'        If Len(Trim(objItemNF.sUMEstoque)) = 0 Or Len(Trim(objItemNF.sUnidadeMed)) = 0 Then
'
'            objProduto.sCodigo = objItemNF.sProduto
'            'Lê o produto
'            lErro = Produto_Le(objProduto)
'            If lErro <> SUCESSO And lErro <> 28030 Then gError 42303
'            If lErro <> SUCESSO Then gError 42304
'            'Preenche as UMs do ItemNF
'            objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
'            objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
'        End If
'        'Formata o Produto
'        lErro = Produto_Formata(GridItens.TextMatrix(objItemNF.iItem, iGrid_Produto_Col), sProduto, iPreenchido)
'        If lErro <> SUCESSO Then gError 42305
'        'Se o produto foi substituído
'        If sProduto <> objItemNF.sProduto Then
'
'            lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProduto)
'            If lErro <> SUCESSO Then gError 42306
'            'Coloca o produto na tela
'            Produto.PromptInclude = False
'            Produto.Text = sProduto
'            Produto.PromptInclude = True
'            '' Coloca no grid alguns dados do produto
'            GridItens.TextMatrix(objItemNF.iItem, iGrid_Produto_Col) = Produto.Text
'            GridItens.TextMatrix(objItemNF.iItem, iGrid_DescProduto_Col) = objItemNF.sDescricaoItem
'            GridItens.TextMatrix(objItemNF.iItem, iGrid_Quantidade_Col) = objItemNF.dQuantidade
'            GridItens.TextMatrix(objItemNF.iItem, iGrid_UnidadeMed_Col) = objItemNF.sUnidadeMed
'
'            'Preco Unitário
'            lErro = TabelaPreco_Atualiza_PrecoUnitario(objProduto)
'            If lErro <> SUCESSO Then gError 42307
'
'            Call PrecoTotal_Calcula(objItemNF.iItem, False)
'
'            Call ValorTotal_Calcula
'
'        End If
'    Next
'
'    Set objNFiscal.ColItensNF = ColItensNF
'
'    Call Preenche_GridAlocacoes(objNFiscal)
'
'    Exit Sub
'
'Erro_BotaoLocalizacao_Click:
'
'    Select Case gErr
'
'        Case 42303, 42305, 42307
'
'        Case 42304
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
'
'        Case 42306
'            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)
'
'        Case 43695
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Function PrecoUnitario_Calcula(sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long
'
'Dim lErro As Long
'Dim objProduto As New ClassProduto
'Dim objUM As New ClassUnidadeDeMedida
'Dim objUMEst As New ClassUnidadeDeMedida
'Dim dFator As Double
'Dim dPercAcresFin As Double
'Dim objCondicaoPagto As New ClassCondicaoPagto
'
'On Error GoTo Erro_PrecoUnitario_Calcula
'
'    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto
'    'Lê o Produto
'    lErro = Produto_Le(objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then gError 42310
'    If lErro = 28030 Then gError 42311 'Não encontrou
'
'    lErro = UM_Conversao(objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
'    If lErro <> SUCESSO Then gError 42312
'
'    'Converte o preço unitário para UM da tela
'    dPrecoUnitario = objTabelaPrecoItem.dPreco / dFator
'
'    'Aplica o acréscimo financeiro no preço unitário
'    dPercAcresFin = StrParaDbl(PercAcrescFin.Text)
'    dPrecoUnitario = dPrecoUnitario * (1 + dPercAcresFin / 100)
'
'    PrecoUnitario_Calcula = SUCESSO
'
'    Exit Function
'
'Erro_PrecoUnitario_Calcula:
'
'    PrecoUnitario_Calcula = gErr
'
'    Select Case gErr
'
'        Case 42312, 42310
'
'        Case 42311
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objTabelaPrecoItem.sCodProduto)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Private Sub SubTotal_Calcula(objGridInt As AdmGrid, iGrid_Coluna As Integer, dSubTotal As Double)
''Faz a soma da Coluna passado no Grid passado e devolve em dValorTotal
'
'Dim iIndice As Integer
'
'    dSubTotal = 0
'
'    For iIndice = 1 To objGridInt.iLinhasExistentes
'        'Acumula em dSubTotal
'        dSubTotal = dSubTotal + CDbl(IIf(Len(Trim(objGridInt.objGrid.TextMatrix(iIndice, iGrid_Coluna))) > 0, objGridInt.objGrid.TextMatrix(iIndice, iGrid_Coluna), 0))
'
'    Next
'
'    Exit Sub
'
'End Sub
'
'Private Sub ValorReal_Calcula(dQuantidade As Double, dValorUnitario As Double, dPercentDesc As Double, dDesconto As Double, dValorReal As Double)
''Calcula o Valor Real
'
'Dim dValorTotal As Double
'Dim dPercDesc1 As Double
'Dim dPercDesc2 As Double
'
'    dValorTotal = dValorUnitario * dQuantidade
'
'    'Se o Percentual Desconto estiver preenchido
'    If dPercentDesc > 0 Then
'
'        'Testa se o desconto está preenchido
'        If dDesconto = 0 Then
'            dPercDesc2 = 0
'        Else
'            'Calcula o Percentual em cima dos valores passados
'            dPercDesc2 = dDesconto / dValorTotal
'            dPercDesc2 = CDbl(Format(dPercDesc2, "0.0000"))
'        End If
'        'se os percentuais (passado e calulado) forem diferentes calcula-se o desconto
'        If dPercentDesc <> dPercDesc2 Then dDesconto = dPercentDesc * dValorTotal
'
'    End If
'
'    dValorReal = dValorTotal - dDesconto
'
'End Sub
'
'Public Function Saida_Celula(objGridInt As AdmGrid) As Long
''Faz a crítica da ceula do grid que está deixando de ser a corrente
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula
'
'    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
'
'    If lErro = SUCESSO Then
'
'        'Tratamento de saída de célula da contabilidade
'        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
'        If lErro <> SUCESSO Then gError 39674
'
'        'Verifica qual o Grid em questão
'        Select Case objGridInt.objGrid.Name
'
'            'Se for o GridParcelas
'            Case GridParcelas.Name
'
'                lErro = Saida_Celula_GridParcelas(objGridInt)
'                If lErro <> SUCESSO Then gError 42313
'
'            'Se for o GridItens
'            Case GridItens.Name
'
'                lErro = Saida_Celula_GridItens(objGridInt)
'                If lErro <> SUCESSO Then gError 42314
'
'            'Se for o GridComissoes
'            Case GridComissoes.Name
'
'                lErro = objTabComissoes.Saida_Celula_GridComissoes(objGridInt)
'                If lErro <> SUCESSO Then gError 42315
'
'            'Rastreamento
'            Case GridRastro.Name
'
'                lErro = gobjRastreamento.Saida_Celula()
'                If lErro <> SUCESSO Then gError 83388
'
'        End Select
'
'        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
'        If lErro <> SUCESSO Then gError 42317
'
'    End If
'
'    Saida_Celula = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula:
'
'    Saida_Celula = gErr
'
'    Select Case gErr
'
'        Case 39674, 42313, 42314, 42315, 42317, 83388
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_GridItens
'
'    'Verifica qual a coluna atual do Grid
'    Select Case objGridInt.objGrid.Col
'
'        Case iGrid_Produto_Col
'            lErro = Saida_Celula_Produto(objGridInt)
'            If lErro <> SUCESSO Then gError 42318
'
'        Case iGrid_DescProduto_Col
'            lErro = Saida_Celula_DescricaoItem(objGridInt)
'            If lErro <> SUCESSO Then gError 42330
'
'        Case iGrid_UnidadeMed_Col
'            lErro = Saida_Celula_UM(objGridInt)
'            If lErro <> SUCESSO Then gError 42319
'
'        Case iGrid_Quantidade_Col
'            lErro = Saida_Celula_Quantidade(objGridInt)
'            If lErro <> SUCESSO Then gError 42320
'
'        Case iGrid_PrecoUnitario_Col
'            lErro = Saida_Celula_PrecoUnitario(objGridInt)
'            If lErro <> SUCESSO Then gError 42321
'
'        Case iGrid_PercDesc_Col
'            lErro = Saida_Celula_PercentDesc(objGridInt)
'            If lErro <> SUCESSO Then gError 42322
'
'        Case iGrid_Desconto_Col
'            lErro = Saida_Celula_Desconto(objGridInt)
'            If lErro <> SUCESSO Then gError 42323
'
'    End Select
'
'    Saida_Celula_GridItens = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_GridItens:
'
'    Saida_Celula_GridItens = gErr
'
'    Select Case gErr
'
'        Case 42318, 42319, 42320, 42321, 42322, 42323, 42330
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Produto que está deixando de ser a corrente
'
'Dim lErro As Long
'Dim objProduto As New ClassProduto
'Dim iProdutoPreenchido As Integer
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_Saida_Celula_Produto
'
'    Set objGridInt.objControle = Produto
'
'    'Faz a crítica do Produto
'    lErro = Produto_Critica(Produto.Text, objProduto, iProdutoPreenchido)
'    If lErro <> SUCESSO And lErro <> 25041 Then gError 42324
'    If lErro <> SUCESSO Then gError 51087
'
'    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
'
'        'Faz a crítica do Produto digitado
'        lErro = Produto_Saida_Celula(objProduto)
'        If lErro <> SUCESSO Then gError 42325
'
'    End If
'
'    'Necessário para o funcionamento da Rotina_Grid_Enable
'    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 42326
'
'    Saida_Celula_Produto = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_Produto:
'
'    Saida_Celula_Produto = gErr
'
'    Select Case gErr
'
'        Case 42324, 42325, 42326
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 51087
'            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
'            If vbMsgRes = vbYes Then
'
'                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)
'                Call Chama_Tela("Produto", objProduto)
'            Else
'                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
'            End If
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_DescricaoItem(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Descrição Item que está deixando de ser a corrente
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_DescricaoItem
'
'    Set objGridInt.objControle = DescricaoItem
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 42329
'
'    Saida_Celula_DescricaoItem = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_DescricaoItem:
'
'    Saida_Celula_DescricaoItem = gErr
'
'    Select Case gErr
'
'        Case 42329
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Private Function Saida_Celula_GridParcelas(objGridInt As AdmGrid) As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_GridParcelas
'
'    'Verifica qual a coluna atual do Grid
'    Select Case objGridInt.objGrid.Col
'        'Data de Vencimento
'        Case iGrid_Vencimento_col
'            lErro = Saida_Celula_DataVencimento(objGridInt)
'            If lErro <> SUCESSO Then gError 42336
'        'VAlor da Parcela
'        Case iGrid_ValorParcela_Col
'            lErro = Saida_Celula_ValorParcela(objGridInt)
'            If lErro <> SUCESSO Then gError 42337
'        'Descontos da PArcela
'        Case iGrid_Desc1Codigo_Col, iGrid_Desc2Codigo_Col, iGrid_Desc3Codigo_Col
'            lErro = Saida_Celula_TipoDesconto(objGridInt)
'            If lErro <> SUCESSO Then gError 42338
'        'Datas de desconto da Parcela
'        Case iGrid_Desc1Ate_Col, iGrid_Desc2Ate_Col, iGrid_Desc3Ate_Col
'            lErro = Saida_Celula_DescontoData(objGridInt)
'            If lErro <> SUCESSO Then gError 42339
'        'VAlores dos descontos da parcela
'        Case iGrid_Desc1Valor_Col, iGrid_Desc2Valor_Col, iGrid_Desc3Valor_Col
'            lErro = Saida_Celula_DescontoValor(objGridInt)
'            If lErro <> SUCESSO Then gError 42340
'        'Percentuais de desconto da parcela.
'        Case iGrid_Desc1Percentual_Col, iGrid_Desc2Percentual_Col, iGrid_Desc3Percentual_Col
'            lErro = Saida_Celula_DescontoPerc(objGridInt)
'            If lErro <> SUCESSO Then gError 42341
'        'Banco Cobrador
'        Case iGrid_Cobranca_Col
'            lErro = Saida_Celula_Cobrador(objGridInt)
'            If lErro <> SUCESSO Then gError 95194
'        'Carteira de Cobranca
'        Case iGrid_CarteiraCobranca_Col
'            lErro = Saida_Celula_CarteiraCobrador(objGridInt)
'            If lErro <> SUCESSO Then gError 95195
'
'    End Select
'
'    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 42342
'
'    Saida_Celula_GridParcelas = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_GridParcelas:
'
'    Saida_Celula_GridParcelas = gErr
'
'    Select Case gErr
'
'        Case 42336, 42337, 42338, 42341, 42339, 42340, 95194, 95195
'
'        Case 42342
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_Cobrador(objGridInt As AdmGrid) As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_Cobrador
'
'    Set objGridInt.objControle = Cobrador
'
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 95193
'
'    Carrega_CarteiraCobrador (Codigo_Extrai(Cobrador))
'    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_CarteiraCobranca_Col) = CarteiraCobrador.List(0)
'
'    Saida_Celula_Cobrador = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_Cobrador:
'
'    Saida_Celula_Cobrador = gErr
'
'    Select Case gErr
'
'        Case 95193
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_CarteiraCobrador(objGridInt As AdmGrid) As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_CarteiraCobrador
'
'    Set objGridInt.objControle = CarteiraCobrador
'
'    'Guarda a opcao selecionada
'    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_CarteiraCobranca_Col) = CarteiraCobrador.Text
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 95193
'
'    Saida_Celula_CarteiraCobrador = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_CarteiraCobrador:
'
'    Saida_Celula_CarteiraCobrador = gErr
'
'    Select Case gErr
'
'        Case 95193
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Private Function Saida_Celula_DataVencimento(objGridInt As AdmGrid) As Long
''Faz a crítica da célula Data Vencimento que está deixando de ser a corrente
'
'Dim lErro As Long
'Dim dtDataReferencia As Date
'Dim dtDataVencimento As Date
'Dim sDataVencimento As String
'Dim bCriouLinha As Boolean
'Dim objFilialCliente As New ClassFilialCliente
'Dim iContador As Integer
'
'On Error GoTo Erro_Saida_Celula_DataVencimento
'
'    Set objGridInt.objControle = DataVencimento
'
'    bCriouLinha = False
'
'    'Verifica se Data de Vencimento esta preenchida
'    If Len(Trim(DataVencimento.ClipText)) > 0 Then
'
'        'Critica a data
'        lErro = Data_Critica(DataVencimento.Text)
'        If lErro <> SUCESSO Then gError 42343
'
'         dtDataVencimento = CDate(DataVencimento.Text)
'
'        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
'        If Len(Trim(DataReferencia.ClipText)) > 0 Then
'            dtDataReferencia = CDate(DataReferencia.Text)
'            If dtDataVencimento < dtDataReferencia Then gError 42345
'        End If
'
'        sDataVencimento = Format(dtDataVencimento, "dd/mm/yyyy")
'
'        'Acrescenta uma linha no Grid se for o caso
'        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
'            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
'            bCriouLinha = True
'        End If
'
'    End If
'
'    If sDataVencimento <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_col) Then CobrancaAutomatica.Value = vbUnchecked
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 42344
'
'    If bCriouLinha Then
'        'Coloca DescontoPadrao
'        lErro = Preenche_DescontoPadrao(GridParcelas.Row)
'        If lErro <> SUCESSO Then gError 51118
'
'        If Len(Trim(Cliente.ClipText)) > 0 Then
'
'            objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
'            lErro = FilialCliente_Le_NomeRed_CodFilial(Cliente.Text, objFilialCliente)
'            If lErro <> SUCESSO Then gError 95206
'
'            If objFilialCliente.iCodCobrador = 0 Then objFilialCliente.iCodCobrador = COBRADOR_PROPRIA_EMPRESA
'
'            For iContador = 0 To Cobrador.ListCount - 1
'                If Cobrador.ItemData(iContador) = objFilialCliente.iCodCobrador Then
'                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) = Cobrador.List(iContador)
'                    Exit For
'                End If
'            Next
'
'            Carrega_CarteiraCobrador (objFilialCliente.iCodCobrador)
'            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_CarteiraCobranca_Col) = CarteiraCobrador.List(0)
'        End If
'
'    End If
'    Saida_Celula_DataVencimento = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_DataVencimento:
'
'    Saida_Celula_DataVencimento = gErr
'
'    Select Case gErr
'
'        Case 42343, 42344
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 42345
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR_REFERENCIA", gErr, dtDataVencimento, GridParcelas.Row, dtDataReferencia)
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 51118, 95206
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Private Sub Limpa_Tela_NFiscal2()
''Limpa a tela sem fechar setas
'
'    gbLimpandoTela = True
'
'    Call Limpa_Tela(Me)
'
'    Status.Caption = ""
'
'    TipoNFiscal.Text = ""
'    Frame1(TAB_Almoxarifado).Enabled = True
'
'    Call Grid_Limpa(objGridItens)
'    Call Grid_Limpa(objGridParcelas)
'    Call Grid_Limpa(objGridAlocacoes)
'    Call Grid_Limpa(objGridComissoes)
'    Call Grid_Limpa(objGridComprovServ)
'
'    If giTipoVersao = VERSAO_FULL Then
'        Serie.Text = ""
'        TabelaPreco.Text = ""
'    End If
'
'    NFiscal.Caption = ""
'    Filial.Clear
'    FilialEntrega.Clear
'    ValorProdutos.Caption = ""
'    ValorTotal.Caption = ""
'    Transportadora.Text = ""
'    PlacaUF.Text = ""
'    CanalVenda.Text = ""
'    SerieNFiscalOriginal.Text = ""
'    CondicaoPagamento.Text = ""
'    TotalPercentualComissao.Caption = ""
'    TotalValorComissao.Caption = ""
'    Emitente.Value = True
'    Call Seleciona_Tipo_Venda
'
'    DataReferencia.PromptInclude = False
'    DataReferencia.Text = Format(gdtDataAtual, "dd/mm/yy")
'    DataReferencia.PromptInclude = True
'    DataSaida.PromptInclude = False
'    DataSaida.Text = Format(gdtDataAtual, "dd/mm/yy")
'    DataSaida.PromptInclude = True
'    DataEmissao.PromptInclude = False
'    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
'    DataEmissao.PromptInclude = True
'
'    iValorDescontoAlterado = 0
'    iClienteAlterado = 0
'    iValorDespesasAlterado = 0
'    iValorFreteAlterado = 0
'    iValorIRRFAlterado = 0
'    iValorSeguroAlterado = 0
'    giFilialEntregaAlterada = 0
'    giFilialEntregaManual = 0
'    giDataReferenciaAlterada = 0
'    iPercAcresFinAlterado = 0
'    '*** incluidos p/tratamento de tributacao *******************************
'    Call gobjTribTab.Limpa_Tela
'    '*** fim tributacao
'
'    'Limpeza da área relativa à contabilidade
'    Call objContabil.Contabil_Limpa_Contabilidade
'
'    'Limpa o Frame de Rastreamento
'    Call gobjRastreamento.Limpa_Tela_Rastreamento
'
'    iAlterado = 0
'
'    gbLimpandoTela = False
'
'End Sub
'
'Private Function Limpa_Tela_NFiscal() As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Limpa_Tela_NFiscal
'
'    'Fecha o comando de seta se estiver aberto
'    lErro = ComandoSeta_Fechar(Me.Name)
'
'    Call Limpa_Tela_NFiscal2
'
'    'Desabilita o Rastreamento
''    BotaoRastreamento.Enabled = False
'
'    'Seleciona série padrão
'    lErro = Seleciona_Serie_Padrao()
'    If lErro <> SUCESSO Then gError 500133
'
'    iAlterado = 0
'
'    Exit Function
'
'Erro_Limpa_Tela_NFiscal:
'
'    Limpa_Tela_NFiscal = gErr
'
'    Select Case gErr
'
'        Case 500133
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Limpa_Tela_NFiscal1() As Long
''Limpa a Tela NFiscalEntrada, mas mantém a natureza e o tipo
'
'Dim sNatureza As String
'Dim sTipoNFiscal As String
'Dim lErro As Long
'
'On Error GoTo Erro_Limpa_Tela_NFiscal1
'
'    sNatureza = NatOpInterna.Text
'    sTipoNFiscal = TipoNFiscal.Text
'
'    lErro = Limpa_Tela_NFiscal()
'    If lErro <> SUCESSO Then gError 500134
'
'    NatOpInterna.Text = sNatureza
'    TipoNFiscal.Text = sTipoNFiscal
'
'    Exit Function
'
'    Limpa_Tela_NFiscal1 = gErr
'
'Erro_Limpa_Tela_NFiscal1:
'
'    Select Case gErr
'
'        Case 500134
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub BotaoLimpar_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoLimpar_Click
'
'    'Testa se deseja salvar mudanças
'    lErro = Teste_Salva(Me, iAlterado)
'    If lErro <> SUCESSO Then gError 42346
'
'    'Limpa a Tela
'    lErro = Limpa_Tela_NFiscal()
'    If lErro <> SUCESSO Then gError 500135
'
'    Exit Sub
'
'Erro_BotaoLimpar_Click:
'
'    Select Case gErr
'
'        Case 42346, 500135
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub GridComissoes_Click()
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Click(objGridComissoes, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
'    End If
'
'End Sub
'
'Public Sub GridComissoes_EnterCell()
'
'    Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
'
'End Sub
'
'Public Sub GridComissoes_GotFocus()
'
'    Call Grid_Recebe_Foco(objGridComissoes)
'
'End Sub
'
'Public Sub GridComissoes_KeyPress(KeyAscii As Integer)
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Trata_Tecla(KeyAscii, objGridComissoes, iExecutaEntradaCelula)
'
'   If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
'    End If
'
'End Sub
'
'Public Sub GridComissoes_LeaveCell()
'
'    Call Saida_Celula(objGridComissoes)
'
'End Sub
'
'Public Sub GridComissoes_Validate(Cancel As Boolean)
'
'    Call Grid_Libera_Foco(objGridComissoes)
'
'End Sub
'
'Public Sub GridComissoes_RowColChange()
'
'    Call Grid_RowColChange(objGridComissoes)
'
'End Sub
'
'Public Sub GridComissoes_Scroll()
'
'    Call Grid_Scroll(objGridComissoes)
'
'End Sub
'
'Public Sub GridParcelas_Click()
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
'    End If
'
'End Sub
'
'Public Sub GridParcelas_GotFocus()
'
'    Call Grid_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub GridParcelas_EnterCell()
'
'    Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
'
'End Sub
'
'Public Sub GridParcelas_LeaveCell()
'
'    Call Saida_Celula(objGridParcelas)
'
'End Sub
'
'Public Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)
'
'End Sub
'
'Public Sub GridParcelas_KeyPress(KeyAscii As Integer)
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
'    End If
'
'End Sub
'
'Public Sub GridParcelas_Validate(Cancel As Boolean)
'
'    Call Grid_Libera_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub GridParcelas_RowColChange()
'
'    Call Grid_RowColChange(objGridParcelas)
'
'End Sub
'
'Public Sub GridParcelas_Scroll()
'
'    Call Grid_Scroll(objGridParcelas)
'
'End Sub
'
'Public Sub BotaoGravar_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoGravar_Click
'
'    'Chama rotina de Gravação
'    lErro = Gravar_Registro()
'    If lErro <> SUCESSO Then gError 42348
'
'    'Limpa a Tela
'    lErro = Limpa_Tela_NFiscal1()
'    If lErro <> SUCESSO Then gError 500136
'
'    Exit Sub
'
'Erro_BotaoGravar_Click:
'
'    Select Case gErr
'
'        Case 42348, 500136
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub NFiscal_Imprime_Romaneio(lNumIntNF As Long)
'
'Dim objRelatorio As New AdmRelatorio
'Dim sNomeTsk As String, sBuffer As String
'Dim lErro As Long
'
'On Error GoTo Erro_NFiscal_Imprime_Romaneio
'
'    lErro = objRelatorio.ExecutarDireto("Romaneio de Separação de Nota Fiscal", "", 0, "", "NNOTAFISCAL", CStr(lNumIntNF))
'    If lErro <> SUCESSO Then gError 64458
'
'    Exit Sub
'
'Erro_NFiscal_Imprime_Romaneio:
'
'    Select Case gErr
'
'        Case 64458
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Function Gravar_Registro() As Long
''Valida os dados para gravação de uma NFiscal Fatura
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim dValorTotal As Double, dFator As Double
'Dim objTipoDocInfo As New ClassTipoDocInfo
'Dim objNFiscal As New ClassNFiscal
'Dim vbMsgRes As VbMsgBoxResult
'Dim dValorDesconto As Double, dValorIRRF As Double
'Dim objProduto As New ClassProduto
'
'On Error GoTo Erro_Gravar_Registro
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    'Verifica se os campos obrigatórios estão preenchidos
'    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 42349
'    If Len(Trim(NatOpInterna.Text)) = 0 Then gError 42350
'    If Len(Trim(Cliente.ClipText)) = 0 Then gError 42351
'    If Len(Trim(Filial.Text)) = 0 Then gError 42352
'    If Len(Trim(Serie.Text)) = 0 Then gError 42353
'    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 42355
'    If Len(Trim(DataSaida.ClipText)) = 0 Then gError 42356
'    If Len(Trim(FilialEntrega.Text)) = 0 Then gError 51524
'
'    'Verifica se o Peso Liqüido é maior que o Peso Bruto
'    If Len(Trim(PesoLiquido.ClipText)) > 0 And Len(Trim(PesoBruto.ClipText)) > 0 Then
'        If CDbl(PesoLiquido.Text) > CDbl(PesoBruto.Text) Then gError 42370
'    End If
'
'    'Verifica se a Data de Saida e menor que a Data de Emissão
'    If CDate(DataSaida.Text) < CDate(DataEmissao.Text) Then gError 42357
'
'    dValorTotal = CDbl(IIf(Len(Trim(ValorTotal.Caption)) > 0, ValorTotal.Caption, 0))
'
'    'Se o total for negativo --> Erro
'    If dValorTotal < 0 Then gError 42369
'
'    'Verifica se o Valor do Desconto é maior que a soma dos demais valores
'    If Len(Trim(ValorDesconto.Text)) > 0 Then dValorDesconto = CDbl(ValorDesconto.Text)
'    lErro = DescontoGlobal_Testa(dValorDesconto)
'    If lErro <> SUCESSO Then gError 56900
'
'    dValorIRRF = StrParaDbl(ValorIRRF.Text)
'    If dValorTotal < dValorIRRF Then gError 56905
'
'    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
'    'Lê o Tipo da Nota Fiscal
'    lErro = TipoDocInfo_Le_Codigo(objTipoDocInfo)
'    If lErro <> SUCESSO And lErro <> 31415 Then gError 42358
'
'    'Não encontrou o Tipo de Documento --> erro
'    If lErro = 31415 Then gError 42359
'
'    If objTipoDocInfo.iComplementar = DOCINFO_NORMAL Then
'        If objGridItens.iLinhasExistentes = 0 And StrParaDbl(TotalComprovServ.Caption) <= 0 Then gError 42360
'    End If
'
'    'Valida os dados do Grid
'    For iIndice = 1 To objGridItens.iLinhasExistentes
'        If objTipoDocInfo.iCodigo <> TIPODOCINFO_COMPLEMENTO_IPI Then
'            If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col)) = 0 Then gError 42366
'        End If
'
'       If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 42361
'    Next
'
'    'Chama Valida_Grid_Parcelas
'    lErro = Valida_Grid_Parcelas()
'    If lErro <> SUCESSO Then gError 42380
'
'    'Valida os dados do grid de comissões
'    lErro = objTabComissoes.Valida_Grid_Comissoes()
'    If lErro <> SUCESSO Then gError 42390
'
'    lErro = gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
'    If lErro <> SUCESSO Then gError 83389
'
'    Set objNFiscal.objRastreamento = gobjRastreamento
'
'    'Recolhe os dados da tela
'    lErro = Move_NF_Memoria(objNFiscal)
'    If lErro <> SUCESSO Then gError 42371
'
'    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Or objTipoDocInfo.iCodigo = DOCINFO_NFFISPC Then
'
'        'Recolhe os dados da Nota Fiscal Original
'        lErro = Move_Dados_NFOriginal_Memoria(objNFiscal)
'        If lErro <> SUCESSO Then gError 42372
'
'    End If
'
'    'Move os Grids para memória
'    lErro = Move_Grids_Memoria(objNFiscal)
'    If lErro <> SUCESSO Then gError 42402
'
'    For iIndice = 1 To objNFiscal.ColItensNF.Count
'        'Se for uma nota fiscal de serviço
'        If objNFiscal.iTipoNFiscal = DOCINFO_NFISFS Then
'
'            objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
'
'            lErro = Produto_Le(objProduto)
'            If lErro <> SUCESSO And lErro <> 23080 Then gError 62102
'            If lErro <> SUCESSO Then gError 62103
'
'            If Not (objProduto.iCompras = PRODUTO_PRODUZIVEL And objProduto.iFaturamento = PRODUTO_VENDAVEL And objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE) Then gError 62104
'
'        End If
'    Next
'
'    lErro = gobjTribTab.Valida_TributacaoNF_Gravacao(objNFiscal)
'    If lErro <> SUCESSO Then gError 46435
'
'    If Len(Trim(NFiscal.Caption)) = 0 Then
'        'verifica se o cliente tem crédito.
'        lErro = NFiscal_Testa_Credito(objNFiscal)
'        If lErro <> SUCESSO Then gError 45406
'    End If
'
'    'mover a parte do rastreamento
'    lErro = gobjRastreamento.Move_Rastro_Memoria(objNFiscal)
'    If lErro <> SUCESSO Then gError 83390
'
''**** Colocado por Leonardo em 31/10/01 para igualar a versão 2******
''**** Para isso tomei como base AlteracoesSistema ****
'    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
'    lErro = objContabil.Contabil_Testa_Data(CDate(DataEmissao.Text))
'    If lErro <> SUCESSO Then gError 92044
''*********
'
'    'Grava a Nota Fiscal de Fatura(incluido os dados contábeis) (contabilidade)
'    lErro = NFiscalFatura_Grava(objNFiscal, objContabil)
'    If lErro <> SUCESSO Then gError 35936
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    If Len(Trim(NFiscal.Caption)) = 0 Then vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_NUMERO_NOTA_GRAVADA", objNFiscal.lNumNotaFiscal)
'
'    If giTipoVersao = VERSAO_FULL Then
'        If ImprimeRomaneio.Value = vbChecked Then
'            Call NFiscal_Imprime_Romaneio(objNFiscal.lNumIntDoc)
'        End If
'    End If
'
'    Gravar_Registro = SUCESSO
'
'    Exit Function
'
'Erro_Gravar_Registro:
'
'    Gravar_Registro = gErr
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case gErr
'
'        Case 45406, 46435, 56900, 62102, 83389, 83390, 92044
'
'        Case 42349
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
'
'        Case 42350
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)
'
'        Case 42351
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
'
'        Case 42352
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
'
'        Case 42353
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
'
'        Case 42355
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
'
'        Case 42356
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_NAO_PREENCHIDA", gErr)
'
'        Case 42357
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_ANTERIOR_DATAEMISSAO", gErr, DataSaida, DataEmissao)
'
'        Case 42358, 42362, 42363, 42371, 42372, 42380, 42390, 42402, 35936, 41525
'
'        Case 42359
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iTipo)
'
'        Case 42360
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)
'
'        Case 42361
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)
'
'        Case 42366
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
'
'        Case 42369
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)
'
'        Case 42370
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr, CDbl(PesoLiquido.Text), CDbl(PesoBruto.Text))
'
'        Case 51524
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)
'
'        Case 56905
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_IR_FONTE_MAIOR_VALOR_TOTAL", gErr, dValorIRRF, dValorTotal)
'
'        Case 62103
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
'
'        Case 62104
'           lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NOTA_SERVICO", gErr, GridItens.TextMatrix(iIndice, iGrid_Produto_Col))
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub BotaoExcluir_Click()
'
'Dim lErro As Long
'Dim vbMsg As VbMsgBoxResult
'Dim objNFiscal As New ClassNFiscal
'
'On Error GoTo Erro_BotaoExcluir_Click
'
'    GL_objMDIForm.MousePointer = vbHourglass
'
'    'Verifica se os campos obrigatórios estão preenchidos
'    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 83973
'    If Len(Trim(Serie.Text)) = 0 Then gError 83974
'    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 83975
'
'    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
'    objNFiscal.sSerie = Serie.Text
'    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
'    objNFiscal.iFilialEmpresa = giFilialEmpresa
'    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
'
'    'Verifica se a existe nota fiscal está cadastrada
'    lErro = NFiscal_Le_1(objNFiscal)
'    If lErro <> SUCESSO And lErro <> 83971 Then gError 83976
'
'    'se a nota não está cadastrada ==> erro
'    If lErro = 83971 Then gError 83977
'
'    'pede confirmacao
'    vbMsg = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_NFISCAL", objNFiscal.lNumNotaFiscal)
'    If vbMsg = vbYes Then
'
'        'Lê os itens da nota fiscal
'        lErro = NFiscalItens_Le(objNFiscal)
'        If lErro <> SUCESSO Then gError 83978
'
'        'Faz o cancelamento de uma nota fiscal de Saida
'        lErro = NotaFiscalSaida_Excluir(objNFiscal, objContabil)
'        If lErro <> SUCESSO Then gError 83979
'
'        'Limpa a Tela
'        lErro = Limpa_Tela_NFiscal1()
'        If lErro <> SUCESSO Then gError 83980
'
'        'fecha o comando de setas
'        Call ComandoSeta_Fechar(Me.Name)
'
'    End If
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Exit Sub
'
'Erro_BotaoExcluir_Click:
'
'    GL_objMDIForm.MousePointer = vbDefault
'
'    Select Case gErr
'
'        Case 83973
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
'
'        Case 83974
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
'
'        Case 83975
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
'
'        Case 84976, 83978, 83979, 83980
'
'        Case 83977
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Function Move_Dados_NFOriginal_Memoria(objNFiscal As ClassNFiscal) As Long
''Move os dados da NotaFiscalOriginal para a memória
'
'Dim lErro As Long
'Dim objNFOriginal As New ClassNFiscal
'
'On Error GoTo Erro_Move_Dados_NFOriginal_Memoria
'
'    'Verifica se a Série e o Número da NF original estão preenchidos
'    If Len(Trim(NFiscalOriginal.ClipText)) = 0 Then gError 42373
'    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Then gError 42374
'
'    objNFOriginal.lCliente = objNFiscal.lCliente
'    objNFOriginal.iFilialCli = objNFiscal.iFilialCli
'    objNFOriginal.sSerie = SerieNFiscalOriginal.Text
'    objNFOriginal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
'
'    lErro = NFiscalFatura_Le_NFOriginal(objNFOriginal)
'    If lErro <> SUCESSO And lErro <> 45347 Then gError 42920
'    'Se não encontrar --> Erro
'    If lErro = 45347 Then gError 42379
'
'    'Guarda em objNfiscal o Número Interno da NF Original
'    objNFiscal.lNumIntNotaOriginal = objNFOriginal.lNumIntDoc
'
'    Move_Dados_NFOriginal_Memoria = SUCESSO
'
'    Exit Function
'
'Erro_Move_Dados_NFOriginal_Memoria:
'
'    Move_Dados_NFOriginal_Memoria = gErr
'
'    Select Case gErr
'
'        Case 42373
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NFISCAL_ORIGINAL_NAO_PREENCHIDO", gErr)
'
'        Case 42374
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NFISCAL_ORIGINAL_NAO_PREENCHIDA", gErr)
'
'        Case 42379
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_ASSOCIADA_CLIENTE", gErr, Cliente.Text, SerieNFiscalOriginal.Text, NFiscalOriginal.Text)
'
'        Case 42920
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Valida_Grid_Parcelas() As Long
''Valida os dados do Grid de Parcelas
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim dSomaParcelas As Double
'Dim dValorIRRF As Double
'Dim dValorTotal As Double
'Dim dtDataEmissao As Date
'Dim dtDataVencimento As Date
'Dim iTamanho As Integer
'Dim iTipo As Integer
'Dim dPercAcrecFin As Double
'Dim iDesconto As Integer
'Dim dtDataDesconto As Date, dValorINSSRetido As Double
'
'On Error GoTo Erro_Valida_Grid_Parcelas
'
'    'Verifica se alguma parcela foi informada
'    If objGridParcelas.iLinhasExistentes = 0 Then gError 42381
'
'    dSomaParcelas = 0
'
'    'Para cada Parcela do grid de parcelas
'    For iIndice = 1 To objGridParcelas.iLinhasExistentes
'
'        dtDataEmissao = StrParaDate(DataEmissao.Text)
'        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col))
'
'        'verifica se o vencimento e o valor da parcela estão preenchidos
'        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col))) = 0 Then gError 42382
'        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) = 0 Then gError 42385
'        'verifica se a data de vencimento da parcela é menor que a data de emissão
'        If dtDataVencimento < dtDataEmissao Then gError 42383
'        'Se o desconto 1 da parcela está preenchido
'        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))) > 0 Then
'            iDesconto = 1
'            iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
'            'Verifica se a data do desconto está preenchida
'            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))) = 0 Then gError 51103
'
'            dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
'            'Valida a data do desconto
'            If dtDataDesconto > dtDataVencimento Then gError 51120
'            'Se o tipo de desconto for por valor
'            If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
'                'Verifica se o valor do desconto está preenchido
'                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))) = 0 Then gError 51106
'            Else
'                'Verifica se o percentual de desconto está preenchido
'                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Percentual_Col))) = 0 Then gError 51107
'            End If
'            'Se o desconto 2 está preenchido
'            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))) > 0 Then
'                iDesconto = 2
'                iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
'                'Verifica se a data de desconto está preenchida
'                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))) = 0 Then gError 51104
'                'Verifica se a data de desconto está ordenada ou se é igual ao desconto anterior
'                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 51112
'                If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col)) Then gError 51114
'                'Se o desconto for do tipo valor
'                If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
'                    'Verifica se o valor está preenchido
'                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))) = 0 Then gError 51108
'                Else
'                    'Verifica se o percentual está preenchido
'                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col))) = 0 Then gError 51109
'                End If
'                'Valida a data de desconto
'                dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
'                If dtDataDesconto > dtDataVencimento Then gError 51121
'                'Se o desconto 3 está preenchido
'                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))) > 0 Then
'                    iDesconto = 3
'                    iTipo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
'                    'Verifica se a data de desconto está preenchida
'                    If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))) = 0 Then gError 51105
'                    'Verifica se a data de desconto está ordenada ou se é igual a do desconto anterior
'                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) < StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 51113
'                    If StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col)) = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col)) Then gError 51115
'                    'Se o desconto for do tipo valor
'                    If iTipo = VALOR_FIXO Or iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Then
'                        'Verifica se valor de desconto está preenchido
'                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))) = 0 Then gError 51110
'                    Else
'                        'verifica se o percentuial de desconto está preenchido
'                        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Percentual_Col))) = 0 Then gError 51111
'                    End If
'                    'Valida a data de desconto
'                    dtDataDesconto = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
'                    If dtDataDesconto > dtDataVencimento Then gError 51122
'                End If
'            End If
'        End If
'        'Verifica se as data de vencimentos das parcelas estão ordenadas
'        If iIndice > 1 Then If CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col)) < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_col)) Then gError 42384
'       'Faz a soma do total das parcelas
'        dSomaParcelas = dSomaParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
'
'    Next
'
'    dValorTotal = StrParaDbl(ValorTotal.Caption)
'    dValorIRRF = StrParaDbl(ValorIRRF.Text)
'    If INSSRetido.Value = vbChecked Then
'        dValorINSSRetido = StrParaDbl(INSSValor.Text)
'    End If
'
'    'Verifica se o total das parcelas cobre o valor da nota fiscal
'    If Format((dValorTotal - dValorIRRF - dValorINSSRetido), "Standard") <> Format(dSomaParcelas, "Standard") Then gError 42386
'
'    Valida_Grid_Parcelas = SUCESSO
'
'    Exit Function
'
'Erro_Valida_Grid_Parcelas:
'
'    Valida_Grid_Parcelas = gErr
'
'    Select Case gErr
'
'        Case 42381
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_PARCELA_COBRANCA", gErr)
'
'        Case 42382
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_NAO_INFORMADA", gErr, iIndice)
'
'        Case 42383
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_COBRANCA_MENOR", gErr, iIndice, dtDataVencimento, dtDataEmissao)
'
'        Case 42384
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_COBRANCA_NAO_ORDENADA", gErr)
'
'        Case 42385
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_PARCELA_COBRANCA_NAO_INFORMADO", gErr, iIndice)
'
'        Case 42386
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_COBRANCA_INVALIDA", gErr)
'
'        Case 51103, 51104, 51105
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_NAO_PREENCHIDA", gErr, iDesconto, iIndice)
'
'        Case 51106, 51107, 51108, 51109, 51110, 51111
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_PARCELA_NAO_PREENCHIDO", gErr, iDesconto, iIndice)
'
'        Case 51112, 51113
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTOS_DESORDENADAS", gErr, iIndice)
'
'        Case 51114, 51115
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTO_IGUAIS", gErr, iDesconto - 1, iDesconto, iIndice)
'
'        Case 51120, 51121, 51122
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_SUPERIOR_DATA_VENCIMENTO", gErr, dtDataDesconto, iDesconto, iIndice)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Private Function Move_Grids_Memoria(objNFiscal As ClassNFiscal) As Long
''Move os dados dos Grids para memória
'
'Dim lErro As Long
'
'On Error GoTo Erro_Move_Grids_Memoria
'
'    'Chama Move_GridItens_Memoria
'    lErro = Move_GridItens_Memoria(objNFiscal)
'    If lErro <> SUCESSO Then gError 42391
'
'    If Frame1(TAB_Almoxarifado).Enabled = True Then
'        'Chama Move_GridAlocacao_Memoria
'        lErro = Move_GridAlocacao_Memoria(objNFiscal)
'        If lErro <> SUCESSO Then gError 42392
'    End If
'
'    'Chama Move_GridComissoes_Memoria
'    lErro = objTabComissoes.Move_TabComissoes_Memoria_NF(objNFiscal)
'    If lErro <> SUCESSO Then gError 42393
'
'    'Chama Move_GridParcelas_Memoria
'    lErro = Move_GridParcelas_Memoria(objNFiscal)
'    If lErro <> SUCESSO Then gError 42394
'
'    'move o grid de comprovantes de serviços para a memória
'    Call Move_GridComprovServ_Memoria(objNFiscal)
'
'    'Mover a parte de tributação
'    lErro = gobjTribTab.PreencherNF(objNFiscal)
'    If lErro <> SUCESSO Then gError 42395
'
'    Move_Grids_Memoria = SUCESSO
'
'    Exit Function
'
'Erro_Move_Grids_Memoria:
'
'    Move_Grids_Memoria = gErr
'
'    Select Case gErr
'
'        Case 42391, 42392, 42393, 42394, 42395
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function Move_GridItens_Memoria(objNFiscal As ClassNFiscal) As Long
''Move os Itens do Grid para a Memória
'
'Dim iIndice As Integer
'Dim lErro As Long
'Dim objItemNF As ClassItemNF
'Dim sProdutoFormatado As String
'Dim iProdutoPreenchido As Integer
'Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim lTamanho As Long
'Dim colAlocacoes As ColAlocacoesItemNF
'Dim objProduto As New ClassProduto
'
'On Error GoTo Erro_Move_GridItens_Memoria
'
'    'Para cada linha existente do Grid
'    For iIndice = 1 To objGridItens.iLinhasExistentes
'
'        Set objItemNF = New ClassItemNF
'
'        'Verifica se o Produto está preenchido
'        lErro = Produto_Formata(GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'        If lErro <> SUCESSO Then gError 42396
'
'        'Armazena produto
'        objItemNF.sProduto = sProdutoFormatado
'
'        'Guarda os demais campos do Grid em objItemNF
'        objItemNF.sDescricaoItem = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
'        objItemNF.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
'        objItemNF.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
'        objItemNF.dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col))
'        objItemNF.dPercDesc = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col))
'
'        objItemNF.dValorDesconto = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
'
'        'Adiciona na coleção de Ítens
'        With objItemNF
'            objNFiscal.ColItensNF.Add 0, iIndice, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, DATA_NULA, .sDescricaoItem, 0, 0, 0, 0, 0, colAlocacoes, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .sCcl, STATUS_LANCADO, 0, "", 0, 0, 0, "", .iClasseUM
'        End With
'
'    Next
'
'    Move_GridItens_Memoria = SUCESSO
'
'    Exit Function
'
'Erro_Move_GridItens_Memoria:
'
'    Move_GridItens_Memoria = gErr
'
'    Select Case gErr
'
'        Case 42396, 45242
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Move_GridAlocacao_Memoria(objNFiscal As ClassNFiscal) As Long
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim iItem As Integer
'Dim objItemAloc As ClassItemNFAlocacao, objItemAloc2 As ClassItemNFAlocacao
'Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim colCodigoNome As New AdmColCodigoNome
'Dim objCodigoNome As AdmCodigoNome
'
'On Error GoTo Erro_Move_GridAlocacao_Memoria
'
'    'Se o número de linhas existentes e  maior que zero
'    If objGridAlocacoes.iLinhasExistentes > 0 Then
'
'        iIndice = 0
'        'Percorre o grid de alocações
'        For iIndice = 1 To objGridAlocacoes.iLinhasExistentes
'
'            iItem = CInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))
'
'            Set objItemAloc = New ClassItemNFAlocacao
'
'            'Preenche o objItemAloc
'            objItemAloc.dQuantidade = CDbl(GridAlocacao.TextMatrix(iIndice, iGrid_QuantAloc_Col))
'            objItemAloc.sAlmoxarifado = GridAlocacao.TextMatrix(iIndice, iGrid_AlmoxAloc_Col)
'            objItemAloc.sUnidadeMed = GridAlocacao.TextMatrix(iIndice, iGrid_UMAloc_Col)
'
'            'Verifica se o Almoxarifado já foi lido
'            If colCodigoNome.Count > 0 Then
'                For Each objCodigoNome In colCodigoNome
'                    If objCodigoNome.sNome = objItemAloc.sAlmoxarifado Then
'                        objItemAloc.iAlmoxarifado = objCodigoNome.iCodigo
'                        Exit For
'                    End If
'                Next
'            End If
'            'Se o almoxarifado ainda não foi lido
'            If objItemAloc.iAlmoxarifado = 0 Then
'                objAlmoxarifado.sNomeReduzido = objItemAloc.sAlmoxarifado
'                'Lê o Almoxarifado à partir do nome reduzido
'                lErro = Almoxarifado_Le_NomeReduzido(objAlmoxarifado)
'                If lErro <> SUCESSO And lErro <> 25060 Then gError 42398
'                If lErro = 25060 Then gError 42399
'
'                objItemAloc.iAlmoxarifado = objAlmoxarifado.iCodigo
'                'Adiciona o almoxarifado na coleção de almoxarifados lidos
'                colCodigoNome.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido
'
'            End If
'            'Adiciona a alocação do grid na coleção de itens
'            With objItemAloc
'                Set objItemAloc2 = objNFiscal.ColItensNF(iItem).colAlocacoes.Add(.iAlmoxarifado, .sAlmoxarifado, .dQuantidade)
'                objItemAloc2.sUnidadeMed = .sUnidadeMed
'            End With
'
'            objNFiscal.ColItensNF(iItem).iAlmoxarifado = objItemAloc.iAlmoxarifado
'
'        Next
'
'    End If
'
'    Move_GridAlocacao_Memoria = SUCESSO
'
'    Exit Function
'
'Erro_Move_GridAlocacao_Memoria:
'
'    Move_GridAlocacao_Memoria = gErr
'
'    Select Case gErr
'
'        Case 42398
'
'        Case 42399
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Private Function Move_GridParcelas_Memoria(objNFiscal As ClassNFiscal) As Long
''Move as Parcelas do Grid para a Memória
'
'Dim iIndice As Integer
'Dim lTamanho As Long
'Dim objParcela As ClassParcelaReceber
'Dim dtDataReferencia As Date
'Dim dtDataEmissao As Date
'Dim lErro As Long
'
'On Error GoTo Erro_Move_GridParcelas_Memoria
'
'    dtDataReferencia = StrParaDate(DataReferencia.Text)
'    dtDataEmissao = StrParaDate(DataEmissao.Text)
'
'    If dtDataReferencia <> DATA_NULA Then
'        If dtDataReferencia < dtDataEmissao Then gError 51615
'    End If
'
'    'Se não há parcelas a recolher, sai da função
'    If objGridParcelas.iLinhasExistentes = 0 Then Exit Function
'
'    'Para cada parcela do grid
'    For iIndice = 1 To objGridParcelas.iLinhasExistentes
'
'        Set objParcela = New ClassParcelaReceber
'
'        objParcela.iNumParcela = iIndice
'
'        'recolhe os dados da parcela
'        objParcela.dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col))
'        objParcela.dValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
'
'        objParcela.iCobrador = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col))
'        objParcela.iCarteiraCobranca = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_CarteiraCobranca_Col))
'
'        objParcela.iDesconto1Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col))
'        objParcela.iDesconto2Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col))
'        objParcela.iDesconto3Codigo = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col))
'        objParcela.dtDesconto1Ate = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col))
'        objParcela.dtDesconto2Ate = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col))
'        objParcela.dtDesconto3Ate = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col))
'
'        'Se o tipo de desconto for de Valor recolhe a coluna valor
'        'Senão recolhe a coluna percentual
'        If objParcela.iDesconto1Codigo = VALOR_FIXO Or objParcela.iDesconto1Codigo = VALOR_ANT_DIA Or objParcela.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
'            objParcela.dDesconto1Valor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col))
'        ElseIf objParcela.iDesconto1Codigo = Percentual Or objParcela.iDesconto1Codigo = PERC_ANT_DIA Or objParcela.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
'            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Percentual_Col)))
'            If lTamanho > 0 Then objParcela.dDesconto1Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc1Percentual_Col))
'        End If
'
'        'Se o tipo de desconto for de Valor recolhe a coluna valor
'        'Senão recolhe a coluna percentual
'        If objParcela.iDesconto2Codigo = VALOR_FIXO Or objParcela.iDesconto2Codigo = VALOR_ANT_DIA Or objParcela.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
'            objParcela.dDesconto2Valor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col))
'        ElseIf objParcela.iDesconto2Codigo = Percentual Or objParcela.iDesconto2Codigo = PERC_ANT_DIA Or objParcela.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
'            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col)))
'            If lTamanho > 0 Then objParcela.dDesconto2Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col))
'        End If
'
'        'Se o tipo de desconto for de Valor recolhe a coluna valor
'        'Senão recolhe a coluna percentual
'        If objParcela.iDesconto3Codigo = VALOR_FIXO Or objParcela.iDesconto3Codigo = VALOR_ANT_DIA Or objParcela.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
'            objParcela.dDesconto3Valor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col))
'        ElseIf objParcela.iDesconto3Codigo = Percentual Or objParcela.iDesconto3Codigo = PERC_ANT_DIA Or objParcela.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
'            lTamanho = Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Percentual_Col)))
'            If lTamanho > 0 Then objParcela.dDesconto3Valor = PercentParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desc3Percentual_Col))
'        End If
'
'        'Adiciona a parcela na coleção de parcelas da Nota Fiscal
'        With objParcela
'            objNFiscal.ColParcelaReceber.Add 0, 0, iIndice, STATUS_ABERTO, .dtDataVencimento, .dtDataVencimento, .dValor, .dValor, 1, .iCarteiraCobranca, .iCobrador, "", 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, 0, 0
'        End With
'
'    Next
'
'    Move_GridParcelas_Memoria = SUCESSO
'
'    Exit Function
'
'Erro_Move_GridParcelas_Memoria:
'
'    Move_GridParcelas_Memoria = gErr
'
'    Select Case gErr
'
'        Case 51615
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAREFERENCIA", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'End Function
'
'Private Function Recolhe_Alocacoes(iItem As Integer, colAlocacoes As ColAlocacoesItemNF) As Long
''Recolhe as Alocações
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim iItemLinha As Integer
'Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim objCodigoNome As New AdmCodigoNome
'Dim colCodigoNome As New AdmColCodigoNome
'
'On Error GoTo Erro_Recolhe_Alocacoes
'
'    For iIndice = 1 To objGridAlocacoes.iLinhasExistentes
'
'        iItemLinha = StrParaInt(GridAlocacao.TextMatrix(iIndice, iGrid_Item_Col))
'
'        If iItemLinha = iItem Then
'
'            objAlmoxarifado.sNomeReduzido = GridAlocacao.TextMatrix(iIndice, iGrid_AlmoxAloc_Col)
'            objAlmoxarifado.iCodigo = 0
'
'            For Each objCodigoNome In colCodigoNome
'                If objCodigoNome.sNome = objAlmoxarifado.sNomeReduzido Then
'                    objAlmoxarifado.iCodigo = objCodigoNome.iCodigo
'                End If
'            Next
'
'            If objAlmoxarifado.iCodigo = 0 Then
'
'                'Lê o Almoxarifado
'                lErro = Almoxarifado_Le_NomeReduzido(objAlmoxarifado)
'                If lErro <> 25060 And lErro <> SUCESSO Then gError 42409
'
'                'Se não encontrou o Almoxarifado --> erro
'                If lErro = 25060 Then gError 42410
'
'                colCodigoNome.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido
'
'            End If
'
'            colAlocacoes.Add objAlmoxarifado.iCodigo, GridAlocacao.TextMatrix(iIndice, iGrid_AlmoxAloc_Col), StrParaDbl(GridAlocacao.TextMatrix(iIndice, iGrid_QuantAloc_Col))
'
'        End If
'    Next
'
'    Recolhe_Alocacoes = SUCESSO
'
'    Exit Function
'
'Erro_Recolhe_Alocacoes:
'
'    Recolhe_Alocacoes = gErr
'
'    Select Case gErr
'
'        Case 42409
'
'        Case 42410
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function Move_NF_Memoria(objNFiscal As ClassNFiscal) As Long
''Move a Nota Fiscal para a Memória
'
'Dim lErro As Long
'Dim objCliente As New ClassCliente
'
'On Error GoTo Erro_Move_NF_Memoria
'
'    'Verifica preenchimento de Cliente
'    If Len(Trim(Cliente.ClipText)) > 0 Then
'
'        objCliente.sNomeReduzido = Cliente.Text
'        'Lê o Cliente
'        lErro = Cliente_Le_Codigo_NomeReduzido(objCliente)
'        If lErro <> SUCESSO And lErro <> 12348 Then gError 42411
'
'        'Se não encontrou o Cliente --> erro
'        If lErro = 12348 Then gError 42412
'
'        objNFiscal.lCliente = objCliente.lCodigo
'
'    End If
'
'    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
'
'    'Preenche o objNFiscal com os demais campos da tela
'    If Len(Trim(PesoBruto.ClipText)) > 0 Then objNFiscal.dPesoBruto = CDbl(PesoBruto.Text)
'    If Len(Trim(PesoLiquido.ClipText)) > 0 Then objNFiscal.dPesoLiq = CDbl(PesoLiquido.Text)
'
'    objNFiscal.dtDataEmissao = MaskedParaDate(DataEmissao)
'    objNFiscal.dtDataSaida = MaskedParaDate(DataSaida)
'    objNFiscal.dtDataReferencia = StrParaDate(DataReferencia.Text)
'
''horasaida
'    If Len(Trim(HoraSaida.ClipText)) > 0 Then
'        objNFiscal.dtHoraSaida = CDate(HoraSaida.Text)
'    Else
'        objNFiscal.dtHoraSaida = Time
'    End If
'
'    'Verifica se os valores estão preenchidos antes de converter
'    objNFiscal.dValorDesconto = StrParaDbl(ValorDesconto.Text)
'    objNFiscal.dValorFrete = StrParaDbl(ValorFrete.Text)
'    objNFiscal.dValorSeguro = StrParaDbl(ValorSeguro.Text)
'    objNFiscal.dValorOutrasDespesas = StrParaDbl(ValorDespesas.Text)
'    objNFiscal.dValorProdutos = StrParaDbl(ValorProdutos.Caption)
'    objNFiscal.dValorTotal = StrParaDbl(ValorTotal.Caption)
'
'    If Len(Trim(NFiscal.Caption)) > 0 Then
'        objNFiscal.lNumNotaFiscal = CLng(NFiscal.Caption)
'    Else
'        objNFiscal.lNumNotaFiscal = 0
'    End If
'
'    If Len(Trim(VolumeQuant.ClipText)) > 0 Then objNFiscal.lVolumeQuant = CInt(VolumeQuant.Text)
'
'    'Armazena o responsável pelo frete
'    If Emitente.Value Then
'        objNFiscal.iFreteRespons = FRETE_EMITENTE
'    Else
'        objNFiscal.iFreteRespons = FRETE_DESTINATARIO
'    End If
'
'    'Recolhe os demais campos da tela
'    objNFiscal.sSerie = Serie.Text
'    objNFiscal.sNumPedidoTerc = NumPedidoTerc.Text
'    objNFiscal.iCodTransportadora = Codigo_Extrai(Transportadora.Text)
'    objNFiscal.iFilialEmpresa = giFilialEmpresa
'    objNFiscal.iStatus = STATUS_LANCADO
'    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
'    objNFiscal.sMensagemNota = Mensagem.Text
'    objNFiscal.sNaturezaOp = NatOpInterna.Text
'    objNFiscal.sPlaca = Placa.Text
'    objNFiscal.sPlacaUF = PlacaUF.Text
'    objNFiscal.sVolumeEspecie = VolumeEspecie.Text
'    objNFiscal.sVolumeMarca = VolumeMarca.Text
'    objNFiscal.sVolumeNumero = VolumeNumero.Text
'    objNFiscal.dtDataEntrada = DATA_NULA
'    objNFiscal.dtDataVencimento = DATA_NULA
'    objNFiscal.lNumIntNotaOriginal = 0
'    objNFiscal.lNumIntDoc = 0
'    objNFiscal.iCanal = Codigo_Extrai(CanalVenda.Text)
'    objNFiscal.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)
'    objNFiscal.iFilialEntrega = Codigo_Extrai(FilialEntrega.Text)
'    objNFiscal.iCondicaoPagto = Codigo_Extrai(CondicaoPagamento.Text)
'    objNFiscal.dValorMercadoria = StrParaDbl(ValorMercadoria.Text)
'    objNFiscal.dValorContainer = StrParaDbl(ValorContainer.Text)
'    objNFiscal.sOrigem = Origem.Text
'    objNFiscal.sDestino = Destino.Text
'
'    Move_NF_Memoria = SUCESSO
'
'    Exit Function
'
'Erro_Move_NF_Memoria:
'
'    Move_NF_Memoria = gErr
'
'    Select Case gErr
'
'        Case 42411
'
'        Case 42412
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function Traz_NFiscal_Tela(objNFiscal As ClassNFiscal) As Long
''Traz os dados da Nota Fiscal passada em objNFiscal
'
'Dim lErro As Long, objTituloRec As New ClassTituloReceber
'Dim objNFiscalOriginal As New ClassNFiscal
'Dim bCancel As Boolean
'Dim bAlterouCT As Boolean
'Dim objTipoDocInfo As New ClassTipoDocInfo
'
'On Error GoTo Erro_Traz_NFiscal_Tela
'
'    'Se for versao LIGHT testa chaveamento
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        lErro = objUserControl.CT_Chavea_Trazendo(objNFiscal, bAlterouCT)
'        If lErro <> SUCESSO Then gError 25822
'
'        If bAlterouCT Then
'
'            'Chama o Traz_NFiscal_Tela de objCT chaveado
'            lErro = objUserControl.objCT.Traz_NFiscal_Tela(objNFiscal)
'            If lErro <> SUCESSO Then gError 25821
'
'            iAlterado = 0
'            gbCarregandoTela = False
'            Traz_NFiscal_Tela = SUCESSO
'
'            Exit Function
'
'        End If
'
'    End If
'
'    gbCarregandoTela = True
'
'    'Limpa a tela sem fechar as setas
'    Call Limpa_Tela_NFiscal2
'
'    'Lê os Ítens da Nota Fiscal
'    lErro = NFiscalItens_Le(objNFiscal)
'    If lErro <> SUCESSO Then gError 35705
'
'    'Lê a parte de Tributação
'    lErro = NFiscal_Le_Tributacao(objNFiscal)
'    If lErro <> SUCESSO Then gError 39342
'
'    'Lê a parte de tributacao dos Ítens da Nota Fiscal
'    lErro = NFiscalItens_Le_Trib(objNFiscal)
'    If lErro <> SUCESSO Then gError 39343
'
'    'Lê as Comissões da Nota Fiscal
'    lErro = ComissoesNF_Le(objNFiscal)
'    If lErro <> SUCESSO Then gError 35703
'
'    'Lê as Alocações dos Ítens da Nota Fiscal
'    lErro = AlocacoesNF_Le(objNFiscal)
'    If lErro <> SUCESSO Then gError 35704
'
'    lErro = ParcelasRecNF_Le(objNFiscal)
'    If lErro <> SUCESSO Then gError 42460
'
'    'Chamada de função incluida por Leo.
'    'Lê os comprovantes de Serviços Vinculados a Nota Fiscal
'    lErro = ComprovServ_Le_Vinculados(objNFiscal)
'    If lErro <> SUCESSO Then gError 94093
'
'    'Chamada de função incluida por Leo.
'    'Lê os comprovantes de Serviços Desvinculados a Nota Fiscal
'    lErro = ComprovServ_Le_Desvinculados(objNFiscal, TIPO_NFISCAL)
'    If lErro <> SUCESSO Then gError 94094
'
'    'carrega dados do titulo a receber associado à nota fiscal para obter a condicao de pagto
'    objTituloRec.lNumIntDoc = objNFiscal.lNumIntDocCPR
'    lErro = TituloReceber_Le(objTituloRec, 1)
'    If lErro <> SUCESSO And lErro <> 26061 Then gError 56571
'    If lErro <> SUCESSO Then
'
'        lErro = TituloReceberBaixado_Le(objTituloRec, 1)
'        If lErro <> SUCESSO And lErro <> 56570 Then gError 56572
'        If lErro <> SUCESSO Then gError 56573
'
'    End If
'
'    objNFiscal.iCondicaoPagto = objTituloRec.iCondicaoPagto
'
'    'Preenche o Status da Nota Fiscal
'    If objNFiscal.iStatus = STATUS_LANCADO Then
'        Status.Caption = STRING_STATUS_LANCADO
'    ElseIf objNFiscal.iStatus = STATUS_BAIXADO Then
'        Status.Caption = STRING_STATUS_BAIXADO
'    ElseIf objNFiscal.iStatus = STATUS_CANCELADO Then
'        Status.Caption = STRING_STATUS_CANCELADO
'    End If
'
'    'Preenche o Cliente
'    Cliente.Text = objNFiscal.lCliente
'    Call Cliente_Validate(bCancel)
'
'    'Preenche a Filial do Cliente
'    Filial.Text = objNFiscal.iFilialCli
'    Call Filial_Validate(bSGECancelDummy)
'
'    'Coloca a FilialEntrega na tela
'    If objNFiscal.iFilialEntrega > 0 Then
'        FilialEntrega.Text = objNFiscal.iFilialEntrega
'        Call FilialEntrega_Validate(bSGECancelDummy)
'    End If
'
'    'Coloca TabelaPreco na Tela
'    If objNFiscal.iTabelaPreco > 0 Then
'        TabelaPreco.Text = objNFiscal.iTabelaPreco
'        Call TabelaPreco_Validate(bSGECancelDummy)
'    End If
'
'    'Coloca o Tipo na tela
'    TipoNFiscal.Text = objNFiscal.iTipoNFiscal
'    Call TipoNFiscal_Validate(bSGECancelDummy)
'
'    'Preenche a Condicao de Pagto
'    If objNFiscal.iCondicaoPagto > 0 Then
'        PercAcrescFin.Text = ""
'        CondicaoPagamento.Text = objNFiscal.iCondicaoPagto
'        Call CondicaoPagamento_Validate(bSGECancelDummy)
'    End If
'
'    'Preenche os demais dados da tela
'    NFiscal.Caption = objNFiscal.lNumNotaFiscal
'    Serie.Text = objNFiscal.sSerie
'
'    If objNFiscal.iCanal > 0 Then
'        CanalVenda.Text = objNFiscal.iCanal
'        Call CanalVenda_Validate(bSGECancelDummy)
'    End If
'
'    If objNFiscal.iFreteRespons = FRETE_EMITENTE Then
'        Emitente.Value = True
'    Else
'        Destinatario.Value = True
'    End If
'
'    VolumeNumero.Text = objNFiscal.sVolumeNumero
'
'    Call DateParaMasked(DataReferencia, objNFiscal.dtDataReferencia)
'    giDataReferenciaAlterada = 0
'    Call DateParaMasked(DataEmissao, objNFiscal.dtDataEmissao)
'    Call DateParaMasked(DataSaida, objNFiscal.dtDataSaida)
'
''horasaida
'    HoraSaida.PromptInclude = False
'    'este teste está correto
'    If objNFiscal.dtDataSaida <> DATA_NULA Then HoraSaida.Text = Format(objNFiscal.dtHoraSaida, "hh:mm:ss")
'    HoraSaida.PromptInclude = True
'
'    'Coloca os Valores na Tela
'    If objNFiscal.dValorDesconto > 0 Then ValorDesconto.Text = Format(objNFiscal.dValorDesconto, "Standard")
'    If objNFiscal.dValorOutrasDespesas > 0 Then ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Standard")
'    If objNFiscal.dValorFrete > 0 Then ValorFrete.Text = Format(objNFiscal.dValorFrete, "Standard")
'    If objNFiscal.dValorSeguro > 0 Then ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Standard")
'    If objNFiscal.dValorTotal > 0 Then ValorTotal.Caption = Format(objNFiscal.dValorTotal, "Standard")
'    If objNFiscal.dValorProdutos > 0 Then ValorProdutos.Caption = Format(objNFiscal.dValorProdutos, "Standard")
'    If objNFiscal.dValorMercadoria > 0 Then ValorMercadoria.Text = Format(objNFiscal.dValorMercadoria, "Standard")
'    If objNFiscal.dValorContainer > 0 Then ValorContainer.Text = Format(objNFiscal.dValorContainer, "Standard")
'
'    If objNFiscal.iCodTransportadora <> 0 Then
'        Transportadora.Text = objNFiscal.iCodTransportadora
'        Call Transportadora_Validate(bSGECancelDummy)
'    End If
'
'    Placa.Text = objNFiscal.sPlaca
'    PlacaUF.Text = objNFiscal.sPlacaUF
'    VolumeEspecie = objNFiscal.sVolumeEspecie
'    VolumeMarca = objNFiscal.sVolumeMarca
'    If objNFiscal.lVolumeQuant > 0 Then VolumeQuant.Text = objNFiscal.lVolumeQuant
'    NumPedidoTerc.Text = objNFiscal.sNumPedidoTerc
'    Origem.Text = objNFiscal.sOrigem
'    Destino.Text = objNFiscal.sDestino
'
'    Mensagem.Text = objNFiscal.sMensagemNota
'    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = Format(objNFiscal.dPesoLiq, "Fixed")
'    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = Format(objNFiscal.dPesoBruto, "Fixed")
'
'    'Preenche o Grid com os Ítens da Nota Fiscal
'    lErro = Preenche_GridItens(objNFiscal)
'    If lErro <> SUCESSO Then gError 39023
'
'    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
'
'    'Le o Tipo para saber se tem Movimento de Estoque
'    lErro = TipoDocInfo_Le_Codigo(objTipoDocInfo)
'    If lErro <> SUCESSO Then gError 83384
'
'    If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
'        'Carrega ItensNF com Rastreamentos
'        lErro = gobjRastreamento.Carrega_RastroItensNF(objNFiscal)
'        If lErro <> SUCESSO Then gError 83385
'    End If
'
'    'Habilita o Rastreamento
''    BotaoRastreamento.Enabled = True
'
'    'Carrega o Tab Comissões
'    lErro = objTabComissoes.Carrega_Tab_Comissoes_NF(objNFiscal)
'    If lErro <> SUCESSO Then gError 39022
'
'    'Preenche o Grid com as Alocações dos ítens da Nota Fiscal
'    lErro = Preenche_GridAlocacoes(objNFiscal)
'    If lErro <> SUCESSO Then gError 39026
'
'    'Preenche o Grid de Parcelas
'    lErro = Preenche_Grid_Parcelas(objNFiscal)
'    If lErro <> SUCESSO Then gError 42467
'
'    'Se o código da Nota Fiscal Original estiver preenchido
'    If objNFiscal.lNumIntNotaOriginal <> 0 Then
'
'        objNFiscalOriginal.lNumIntDoc = objNFiscal.lNumIntNotaOriginal
'        'Lê a Nota Fiscal Original
'        lErro = NFiscal_Le(objNFiscalOriginal)
'        If lErro <> SUCESSO And lErro <> 31442 Then gError 39024
'
'        'Se não encontrou a NFiscal Original --> erro
'        If lErro <> SUCESSO Then gError 39025
'
'        'Coloca os dados relativos  Nota Fiscal Original na Tela
'        SerieNFiscalOriginal.Text = objNFiscalOriginal.sSerie
'        NFiscalOriginal.Text = objNFiscalOriginal.lNumNotaFiscal
'
'    End If
'
'    'Preenche o Grid de Comprovantes de Serviços
'    lErro = Preenche_GridComprovServ(objNFiscal)
'    If lErro <> SUCESSO Then gError 94092
'
'    ''*** incluidos p/tratamento de tributacao *******************************
'    lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
'    If lErro <> SUCESSO Then gError 39031
'    '*** fim tributacao
'
'    'Traz os dados contábeis para a tela (contabilidade)
'    lErro = objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
'    If lErro <> SUCESSO And lErro <> 36326 Then gError 39673
'
'    iAlterado = 0
'
'    gbCarregandoTela = False
'
'    Traz_NFiscal_Tela = SUCESSO
'
'    Exit Function
'
'Erro_Traz_NFiscal_Tela:
'
'    gbCarregandoTela = False
'
'    Traz_NFiscal_Tela = gErr
'
'    Select Case gErr
'
'        Case 25821, 25822, 39031, 35705, 39023, 39024, 39022, 39342, 39343, 39673, 42460, 42467, 56571, 56572, 83384, 83385
'
'        Case 56573
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_REC_INEXISTENTE", gErr)
'
'        Case 39025
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntNotaOriginal)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'
'Public Function Preenche_GridItens(objNFiscal As ClassNFiscal) As Long
''Preenche o Grid com os Itens da Nota Fiscal
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim objItemNF As ClassItemNF
'Dim sProdutoEnxuto As String
'Dim objAlmoxarifado As New ClassAlmoxarifado
'
'On Error GoTo Erro_Preenche_GridItens
'
'    iIndice = 0
'
'    'Para cada ítem da Coleção
'    For Each objItemNF In objNFiscal.ColItensNF
'
'        iIndice = iIndice + 1
'        objItemNF.iItem = iIndice
'        'Formata o Produto
'        lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
'        If lErro <> SUCESSO Then gError 35943
'
'        Produto.PromptInclude = False
'        Produto.Text = sProdutoEnxuto
'        Produto.PromptInclude = True
'
'        'Preenche o Grid
'        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
'        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItemNF.sDescricaoItem
'        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemNF.sUnidadeMed
'        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemNF.dQuantidade)
'        GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col) = Format(objItemNF.dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
'        GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(objItemNF.dPercDesc, "Percent")
'        GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objItemNF.dValorDesconto, "Standard")
'        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objItemNF.dValorTotal, "Standard")
'
'    Next
'
'    'Atualiza o número de linhas existentes
'    objGridItens.iLinhasExistentes = iIndice
'
'    Preenche_GridItens = SUCESSO
'
'    Exit Function
'
'Erro_Preenche_GridItens:
'
'    Preenche_GridItens = gErr
'
'    Select Case gErr
'
'        Case 35943
'            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
'Private Function Preenche_GridAlocacoes(objNFiscal As ClassNFiscal) As Long
''Preenche o Grid com as Alocações da Nota Fiscal
'
'Dim objItemAloc As ClassItemNFAlocacao
'Dim iIndice As Integer
'Dim iIndice1 As Integer
'Dim lErro As Long
'Dim objProduto As New ClassProduto
'Dim dFator As Double
'Dim dQuantidade As Double
'Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim iNumCasasDec As Integer
'Dim dAcrescimo As Double
'
'On Error GoTo Erro_Preenche_GridAlocacoes
'
'    'Limpa o grid de alocações
'    Call Grid_Limpa(objGridAlocacoes)
'
'    'Almox mesmo Produto
'
'    iIndice1 = 0
'    'Para cada item da NF
'    For iIndice = 1 To objNFiscal.ColItensNF.Count
'
'        Call AlocacoesNF_Agrupa(objNFiscal.ColItensNF.Item(iIndice).colAlocacoes)
'
'        'Para cada alocação do Item de NF
'        For Each objItemAloc In objNFiscal.ColItensNF.Item(iIndice).colAlocacoes
'
'            iIndice1 = iIndice1 + 1
'
'            objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
'            'Lê o Produto
'            lErro = Produto_Le(objProduto)
'            If lErro <> SUCESSO And lErro <> 28030 Then gError 51087
'            If lErro <> SUCESSO Then gError 51088
'
'            lErro = UM_Conversao_Trans(objProduto.iClasseUM, objNFiscal.ColItensNF(iIndice).sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
'            If lErro <> SUCESSO Then gError 51089
'            'Coloca os dados da alocação na tela
'            GridAlocacao.TextMatrix(iIndice1, iGrid_Item_Col) = objNFiscal.ColItensNF(iIndice).iItem
'            GridAlocacao.TextMatrix(iIndice1, iGrid_ProdutoAloc_Col) = GridItens.TextMatrix(objNFiscal.ColItensNF(iIndice).iItem, iGrid_Produto_Col)
'            GridAlocacao.TextMatrix(iIndice1, iGrid_AlmoxAloc_Col) = objItemAloc.sAlmoxarifado
'            GridAlocacao.TextMatrix(iIndice1, iGrid_UMAloc_Col) = objProduto.sSiglaUMEstoque
'            GridAlocacao.TextMatrix(iIndice1, iGrid_QuantAloc_Col) = Formata_Estoque(objItemAloc.dQuantidade)
'            dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
'            dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
'            dQuantidade = dQuantidade * dFator
'            If StrParaDbl(Formata_Estoque(dQuantidade)) < dQuantidade Then
'
'                'Calcula o número de casas decimais do Formato de Estoque
'                iNumCasasDec = Len(Mid(FORMATO_ESTOQUE, (InStr(FORMATO_ESTOQUE, ".")) + 1))
'
'                dAcrescimo = 0
'                If iNumCasasDec > 0 Then dAcrescimo = 10 ^ -iNumCasasDec
'
'                dQuantidade = dQuantidade + dAcrescimo
'
'            End If
'
'            GridAlocacao.TextMatrix(iIndice1, iGrid_QuantVend_Col) = Formata_Estoque(dQuantidade)
'
'
'        Next
'
'    Next
'
'    objGridAlocacoes.iLinhasExistentes = iIndice1
'
'    Preenche_GridAlocacoes = SUCESSO
'
'    Exit Function
'
'Erro_Preenche_GridAlocacoes:
'
'    Preenche_GridAlocacoes = gErr
'
'    Select Case gErr
'
'        Case 51087, 51089
'
'        Case 51088
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'End Function
'
'Function Preenche_Grid_Parcelas(objNFiscal As ClassNFiscal) As Long
''Preenche o Grid com as Parcelas da Nota Fiscal
'
'Dim objParcela As New ClassParcelaReceber
'Dim iIndice As Integer
'Dim iIndice2 As Integer
'Dim dValorDesconto As Double
'Dim iContador As Integer
'Dim objCarteiraCobranca As New ClassCarteiraCobranca
'Dim lErro As Long
'
'On Error GoTo Erro_Preenche_Grid_Parcelas
'
'    Call Grid_Limpa(objGridParcelas)
'
'    iIndice = 0
'
'    'PAra cada parcela da coleção de parcelas
'    For Each objParcela In objNFiscal.ColParcelaReceber
'
'        iIndice = iIndice + 1
'        'Preenche o grid com os dados da parcela
'        GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_col) = Format(objParcela.dtDataVencimento, "dd/mm/yyyy")
'        GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objParcela.dValor, "Standard")
'
'        'Inicio da Alteracao Daniel em 08/10/2001
'
'        'Preenche o Cobrador
'        For iContador = 0 To Cobrador.ListCount - 1
'            If Cobrador.ItemData(iContador) = objParcela.iCobrador Then
'                GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col) = Cobrador.List(iContador)
'                Exit For
'            End If
'        Next
'
'        'Atribui o codigo para o obj a ser passado como parametro
'        objCarteiraCobranca.iCodigo = objParcela.iCarteiraCobranca
'
'        'Le a carteira de cobranca a partir de um código
'        lErro = CarteiraDeCobranca_Le(objCarteiraCobranca)
'        If lErro <> SUCESSO And lErro <> 23413 Then gError 95190
'
'        'Preenche a carteira de cobranca
'        GridParcelas.TextMatrix(iIndice, iGrid_CarteiraCobranca_Col) = objCarteiraCobranca.iCodigo & SEPARADOR & objCarteiraCobranca.sDescricao
'
'        If objParcela.dtDesconto1Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col) = Format(objParcela.dtDesconto1Ate, "dd/mm/yyyy")
'        If objParcela.dtDesconto2Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col) = Format(objParcela.dtDesconto2Ate, "dd/mm/yyyy")
'        If objParcela.dtDesconto3Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col) = Format(objParcela.dtDesconto3Ate, "dd/mm/yyyy")
'        If objParcela.iDesconto1Codigo = VALOR_FIXO Or objParcela.iDesconto1Codigo = VALOR_ANT_DIA Or objParcela.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
'            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(objParcela.dDesconto1Valor, "Standard")
'        ElseIf objParcela.iDesconto1Codigo = Percentual Or objParcela.iDesconto1Codigo = PERC_ANT_DIA Or objParcela.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
'            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Percentual_Col) = Format(objParcela.dDesconto1Valor, "Percent")
'            '*** Inicio código p/ colocar Valor Desconto
'            If objParcela.dValor > 0 Then
'                dValorDesconto = objParcela.dDesconto1Valor * objParcela.dValor
'                GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(dValorDesconto, "Standard")
'            End If
'            '*** Fim
'        End If
'        If objParcela.iDesconto2Codigo = VALOR_FIXO Or objParcela.iDesconto2Codigo = VALOR_ANT_DIA Or objParcela.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
'            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(objParcela.dDesconto2Valor, "Standard")
'        ElseIf objParcela.iDesconto2Codigo = Percentual Or objParcela.iDesconto2Codigo = PERC_ANT_DIA Or objParcela.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
'            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Percentual_Col) = Format(objParcela.dDesconto2Valor, "Percent")
'            '*** Inicio código p/ colocar Valor Desconto
'            If objParcela.dValor > 0 Then
'                dValorDesconto = objParcela.dDesconto2Valor * objParcela.dValor
'                GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(dValorDesconto, "Standard")
'            End If
'            '*** Fim
'        End If
'        If objParcela.iDesconto3Codigo = VALOR_FIXO Or objParcela.iDesconto3Codigo = VALOR_ANT_DIA Or objParcela.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
'            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(objParcela.dDesconto3Valor, "Standard")
'        ElseIf objParcela.iDesconto3Codigo = Percentual Or objParcela.iDesconto3Codigo = PERC_ANT_DIA Or objParcela.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
'            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Percentual_Col) = Format(objParcela.dDesconto3Valor, "Percent")
'            '*** Inicio código p/ colocar Valor Desconto
'            If objParcela.dValor > 0 Then
'                dValorDesconto = objParcela.dDesconto3Valor * objParcela.dValor
'                GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(dValorDesconto, "Standard")
'            End If
'            '*** Fim
'        End If
'        For iIndice2 = 0 To Desconto1Codigo.ListCount - 1
'            If Desconto1Codigo.ItemData(iIndice2) = objParcela.iDesconto1Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col) = Desconto1Codigo.List(iIndice2)
'            If Desconto2Codigo.ItemData(iIndice2) = objParcela.iDesconto2Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col) = Desconto2Codigo.List(iIndice2)
'            If Desconto3Codigo.ItemData(iIndice2) = objParcela.iDesconto3Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col) = Desconto3Codigo.List(iIndice2)
'        Next
'
'    Next
'
'    objGridParcelas.iLinhasExistentes = iIndice
'
'    Preenche_Grid_Parcelas = SUCESSO
'
'    Exit Function
'
'Erro_Preenche_Grid_Parcelas:
'
'    Preenche_Grid_Parcelas = gErr
'
'    Select Case gErr
'
'        Case 95190
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub Form_Activate()
'
'Dim lErro As Long
'
'On Error GoTo Erro_Form_Activate
'
'    lErro = CargaPosFormLoad
'    If lErro <> SUCESSO Then gError 59335
'
'    Call TelaIndice_Preenche(Me)
'
'    Exit Sub
'
'Erro_Form_Activate:
'
'    Select Case gErr
'
'        Case 59335
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub Form_Deactivate()
'
'    gi_ST_SetaIgnoraClick = 1
'
'End Sub
'
''Início contabilidade
'Public Sub CTBBotaoModeloPadrao_Click()
'
'    Call objContabil.Contabil_BotaoModeloPadrao_Click
'
'End Sub
'
'Public Sub CTBModelo_Click()
'
'    Call objContabil.Contabil_Modelo_Click
'
'End Sub
'
'Public Sub CTBGridContabil_Click()
'
'    Call objContabil.Contabil_GridContabil_Click
'
'    If giTipoVersao = VERSAO_LIGHT Then
'        Call objContabil.Contabil_GridContabil_Consulta_Click
'    End If
'
'End Sub
'
'Public Sub CTBGridContabil_EnterCell()
'
'    Call objContabil.Contabil_GridContabil_EnterCell
'
'End Sub
'
'Public Sub CTBGridContabil_GotFocus()
'
'    Call objContabil.Contabil_GridContabil_GotFocus
'
'End Sub
'
'Public Sub CTBGridContabil_KeyPress(KeyAscii As Integer)
'
'    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)
'
'End Sub
'
'Public Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)
'
'End Sub
'
'
'Public Sub CTBGridContabil_LeaveCell()
'
'        Call objContabil.Contabil_GridContabil_LeaveCell
'
'End Sub
'
'Public Sub CTBGridContabil_Validate(Cancel As Boolean)
'
'    Call objContabil.Contabil_GridContabil_Validate(Cancel)
'
'End Sub
'
'Public Sub CTBGridContabil_RowColChange()
'
'    Call objContabil.Contabil_GridContabil_RowColChange
'
'End Sub
'
'Public Sub CTBGridContabil_Scroll()
'
'    Call objContabil.Contabil_GridContabil_Scroll
'
'End Sub
'
'Public Sub CTBConta_Change()
'
'    Call objContabil.Contabil_Conta_Change
'
'End Sub
'
'Public Sub CTBConta_GotFocus()
'
'    Call objContabil.Contabil_Conta_GotFocus
'
'End Sub
'
'Public Sub CTBConta_KeyPress(KeyAscii As Integer)
'
'    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)
'
'End Sub
'
'Public Sub CTBConta_Validate(Cancel As Boolean)
'
'    Call objContabil.Contabil_Conta_Validate(Cancel)
'
'End Sub
'
'Public Sub CTBCcl_Change()
'
'    Call objContabil.Contabil_Ccl_Change
'
'End Sub
'
'Public Sub CTBCcl_GotFocus()
'
'    Call objContabil.Contabil_Ccl_GotFocus
'
'End Sub
'
'Public Sub CTBCcl_KeyPress(KeyAscii As Integer)
'
'    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)
'
'End Sub
'
'Public Sub CTBCcl_Validate(Cancel As Boolean)
'
'    Call objContabil.Contabil_Ccl_Validate(Cancel)
'
'End Sub
'
'Public Sub CTBCredito_Change()
'
'    Call objContabil.Contabil_Credito_Change
'
'End Sub
'
'Public Sub CTBCredito_GotFocus()
'
'    Call objContabil.Contabil_Credito_GotFocus
'
'End Sub
'
'Public Sub CTBCredito_KeyPress(KeyAscii As Integer)
'
'    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)
'
'End Sub
'
'Public Sub CTBCredito_Validate(Cancel As Boolean)
'
'    Call objContabil.Contabil_Credito_Validate(Cancel)
'
'End Sub
'
'Public Sub CTBDebito_Change()
'
'    Call objContabil.Contabil_Debito_Change
'
'End Sub
'
'Public Sub CTBDebito_GotFocus()
'
'    Call objContabil.Contabil_Debito_GotFocus
'
'End Sub
'
'Public Sub CTBDebito_KeyPress(KeyAscii As Integer)
'
'    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)
'
'End Sub
'
'Public Sub CTBDebito_Validate(Cancel As Boolean)
'
'    Call objContabil.Contabil_Debito_Validate(Cancel)
'
'End Sub
'
'Public Sub CTBSeqContraPartida_Change()
'
'    Call objContabil.Contabil_SeqContraPartida_Change
'
'End Sub
'
''****
'Public Sub CTBSeqContraPartida_GotFocus()
'
'    Call objContabil.Contabil_SeqContraPartida_GotFocus
'
'End Sub
'
''****
'Public Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)
'
'    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)
'
'End Sub
'
''****
'Public Sub CTBSeqContraPartida_Validate(Cancel As Boolean)
'
'    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)
'
'End Sub
'
'Public Sub CTBHistorico_Change()
'
'    Call objContabil.Contabil_Historico_Change
'
'End Sub
'
'Public Sub CTBHistorico_GotFocus()
'
'    Call objContabil.Contabil_Historico_GotFocus
'
'End Sub
'
'Public Sub CTBHistorico_KeyPress(KeyAscii As Integer)
'
'    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)
'
'End Sub
'
'Public Sub CTBHistorico_Validate(Cancel As Boolean)
'
'    Call objContabil.Contabil_Historico_Validate(Cancel)
'
'End Sub
'
'Public Sub CTBLancAutomatico_Click()
'
'    Call objContabil.Contabil_LancAutomatico_Click
'
'End Sub
'
'Public Sub CTBAglutina_Click()
'
'    Call objContabil.Contabil_Aglutina_Click
'
'End Sub
'
'Public Sub CTBAglutina_GotFocus()
'
'    Call objContabil.Contabil_Aglutina_GotFocus
'
'End Sub
'
'Public Sub CTBAglutina_KeyPress(KeyAscii As Integer)
'
'    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)
'
'End Sub
'
'Public Sub CTBAglutina_Validate(Cancel As Boolean)
'
'    Call objContabil.Contabil_Aglutina_Validate(Cancel)
'
'End Sub
'
'Public Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
'
'    Call objContabil.Contabil_TvwContas_NodeClick(Node)
'
'End Sub
'
'Public Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)
'
'    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)
'
'End Sub
'
'Public Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
'
'    Call objContabil.Contabil_TvwCcls_NodeClick(Node)
'
'End Sub
'
'Public Sub CTBListHistoricos_DblClick()
'
'    Call objContabil.Contabil_ListHistoricos_DblClick
'
'End Sub
'
'Public Sub CTBBotaoLimparGrid_Click()
'
'    Call objContabil.Contabil_Limpa_GridContabil
'
'End Sub
'
'Public Sub CTBLote_Change()
'
'    Call objContabil.Contabil_Lote_Change
'
'End Sub
'
'Public Sub CTBLote_GotFocus()
'
'    Call objContabil.Contabil_Lote_GotFocus
'
'End Sub
'
'Public Sub CTBLote_Validate(Cancel As Boolean)
'
'    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)
'
'End Sub
'
'Public Sub CTBDataContabil_Change()
'
'    Call objContabil.Contabil_DataContabil_Change
'
'End Sub
'
'Public Sub CTBDataContabil_GotFocus()
'
'    Call objContabil.Contabil_DataContabil_GotFocus
'
'End Sub
'
'Public Sub CTBDataContabil_Validate(Cancel As Boolean)
'
'    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)
'
'End Sub
'
'Private Sub objEventoLote_evSelecao(obj1 As Object)
''Traz o lote selecionado para a tela
'
'    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)
'
'End Sub
'
'Private Sub objEventoDoc_evSelecao(obj1 As Object)
'
'    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)
'
'End Sub
'
'Public Sub CTBDocumento_Change()
'
'    Call objContabil.Contabil_Documento_Change
'
'End Sub
'
'Public Sub CTBDocumento_GotFocus()
'
'    Call objContabil.Contabil_Documento_GotFocus
'
'End Sub
'
'Public Sub CTBBotaoImprimir_Click()
'
'    Call objContabil.Contabil_BotaoImprimir_Click
'
'End Sub
'
'Public Sub CTBUpDown_DownClick()
'
'    Call objContabil.Contabil_UpDown_DownClick
'
'End Sub
'
'Public Sub CTBUpDown_UpClick()
'
'    Call objContabil.Contabil_UpDown_UpClick
'
'End Sub
'
'Public Sub CTBLabelDoc_Click()
'
'    Call objContabil.Contabil_LabelDoc_Click
'
'End Sub
'
'Public Sub CTBLabelLote_Click()
'
'    Call objContabil.Contabil_LabelLote_Click
'
'End Sub
'
'Public Sub Cobrador_Click()
'    iAlterado = REGISTRO_ALTERADO
'    CobrancaAutomatica.Value = vbUnchecked
'End Sub
'Public Sub CarteiraCobrador_Click()
'    iAlterado = REGISTRO_ALTERADO
'    CobrancaAutomatica.Value = vbUnchecked
'End Sub
'
'Public Sub CarteiraCobrador_Change()
'    iAlterado = REGISTRO_ALTERADO
'    CobrancaAutomatica.Value = vbUnchecked
'End Sub
'Public Sub Cobrador_Change()
'    iAlterado = REGISTRO_ALTERADO
'    CobrancaAutomatica.Value = vbUnchecked
'End Sub
'Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long
'
'Dim lErro As Long
'Dim iLinha As Integer
'Dim objCliente As New ClassCliente
'Dim objEstoqueProduto As New ClassEstoqueProduto
'Dim sProdutoFormatado As String
'Dim iPreenchido As Integer
'Dim sContaMascarada As String
'Dim iAlmoxPadrao As Integer
'Dim bEncontrouProduto As Boolean
'Dim iLinha2 As Integer
'Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim dQuantidade As Double
'Dim dQuantidadeConvertida As Double
'Dim objItemMovEstoque As New ClassItemMovEstoque
'Dim sProduto As String
'Dim sUM As String
'
'On Error GoTo Erro_Calcula_Mnemonico
'
'    Select Case objMnemonicoValor.sMnemonico
'
'        Case ESCANINHO_CUSTO_CONSIG
'            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
'                If Len(GridAlocacao.TextMatrix(iLinha, iGrid_ProdutoAloc_Col)) > 0 Then
'                    objMnemonicoValor.colValor.Add ESCANINHO_3_EM_CONSIGNACAO
'                Else
'                    objMnemonicoValor.colValor.Add 0
'                End If
'
'            Next
'
'        Case ESCANINHO_CUSTO
'            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
'                If Len(GridAlocacao.TextMatrix(iLinha, iGrid_ProdutoAloc_Col)) > 0 Then
'                    objMnemonicoValor.colValor.Add ESCANINHO_NOSSO
'                Else
'                    objMnemonicoValor.colValor.Add 0
'                End If
'
'            Next
'
'        Case QUANT_ALOCADA_CONSIG
'            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
'                If Len(GridAlocacao.TextMatrix(iLinha, iGrid_QuantAloc_Col)) > 0 Then
'
'                    'Define o produto que será passado como parâmetro para MovEstoque_Le_QuantVendConsig
'                    lErro = Produto_Formata(GridAlocacao.TextMatrix(iLinha, iGrid_ProdutoAloc_Col), sProdutoFormatado, iPreenchido)
'                    If lErro <> SUCESSO Then gError 79009
'
'                    objItemMovEstoque.sProduto = sProdutoFormatado
'
'                    'Define o almoxarifado que será passado como parâmetro para MovEstoque_Le_QuantVendConsig
'                    objAlmoxarifado.sNomeReduzido = GridAlocacao.TextMatrix(iLinha, iGrid_AlmoxAloc_Col)
'
'                    lErro = Almoxarifado_Le_NomeReduzido(objAlmoxarifado)
'                    If lErro <> SUCESSO And lErro <> 25060 Then gError 79008
'
'                    objItemMovEstoque.iAlmoxarifado = objAlmoxarifado.iCodigo
'
'                    'Define o tipo de movimento, o DocOrigem e o TipoNumIntoDocOrigem que serão passados como parâmetros para MovEstoque_Le_QuantVendConsig
'
'                    objItemMovEstoque.iTipoMov = MOV_EST_NF_VENDA_MAT_CONSIG
'                    objItemMovEstoque.sDocOrigem = "NFISFV" & " " & Serie.Text & " " & objContabil.lNumNotaFiscalExt
'                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
'
'                    'Verifica se MovEstoque_Le_QuantVendConsig não encontrou erro
'                    lErro = MovEstoque_Le_QuantVendConsig(objItemMovEstoque)
'                    If lErro <> SUCESSO And lErro <> 79003 Then gError 79007
'
'                    'se não conseguiu encontrar mov. estoque para os parametros em questao ==> quantidade é zerada
'                    If lErro = 79003 Then objItemMovEstoque.dQuantidade = 0
'
'                    'Passa para o mnemônico o valor encontrado por MovEstoque_Le_QuantVendConsig
'                    objMnemonicoValor.colValor.Add objItemMovEstoque.dQuantidade
'
'                Else
'                    objMnemonicoValor.colValor.Add 0
'
'                End If
'
'            Next
'
'
'        Case QUANT_DISPONIVEL, QUANT_CONSIGTERC
'
'            For iLinha = 1 To objGridItens.iLinhasExistentes
'
'                lErro = Produto_Formata(GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
'                If lErro <> SUCESSO Then gError 75467
'
'                lErro = AlmoxarifadoPadrao_Le(giFilialEmpresa, sProdutoFormatado, iAlmoxPadrao)
'                If lErro <> SUCESSO Then gError 75465
'
'                If iAlmoxPadrao > 0 Then
'
'                    'Lê as quantidades do Produto no almoxarifado
'                    objEstoqueProduto.iAlmoxarifado = iAlmoxPadrao
'                    objEstoqueProduto.sProduto = sProdutoFormatado
'                    lErro = EstoqueProduto_Le(objEstoqueProduto)
'                    If lErro <> SUCESSO And lErro <> 21306 Then gError 75468
'
'                    'Se não encontrou Estoque Produto, erro
'                    If lErro = 21306 Then gError 75469
'
'                    If Len(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
'
'                        If objMnemonicoValor.sMnemonico = QUANT_DISPONIVEL Then
'                            dQuantidade = objEstoqueProduto.dQuantDisponivel
'                        ElseIf objMnemonicoValor.sMnemonico = QUANT_CONSIGTERC Then
'                            dQuantidade = objEstoqueProduto.dQuantConsig3
'                        End If
'
'                        lErro = UMEstoque_Conversao(sProdutoFormatado, GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), dQuantidade, dQuantidadeConvertida)
'                        If lErro <> SUCESSO Then gError 75470
'
'                        objMnemonicoValor.colValor.Add dQuantidadeConvertida
'                    Else
'                        objMnemonicoValor.colValor.Add 0
'                    End If
'                Else
'                    objMnemonicoValor.colValor.Add 0
'                End If
'            Next
'
'        Case ALMOX1
'            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
'                If Len(GridAlocacao.TextMatrix(iLinha, iGrid_AlmoxAloc_Col)) > 0 Then
'                    objMnemonicoValor.colValor.Add GridAlocacao.TextMatrix(iLinha, iGrid_AlmoxAloc_Col)
'                Else
'                    objMnemonicoValor.colValor.Add ""
'                End If
'            Next
'
'        Case DATA_EMISSAO
'            If Len(DataEmissao.ClipText) > 0 Then
'                objMnemonicoValor.colValor.Add CDate(DataEmissao.FormattedText)
'            Else
'                objMnemonicoValor.colValor.Add DATA_NULA
'            End If
'
'        Case DATA_SAIDA
'            If Len(DataSaida.ClipText) > 0 Then
'                objMnemonicoValor.colValor.Add CDate(DataSaida.FormattedText)
'            Else
'                objMnemonicoValor.colValor.Add DATA_NULA
'            End If
'
'        Case DESCONTO1
'            For iLinha = 1 To objGridItens.iLinhasExistentes
'                If Len(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col)) > 0 Then
'                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col))
'                Else
'                    objMnemonicoValor.colValor.Add 0
'                End If
'            Next
'
'        Case DESCRICAO_ITEM
'            For iLinha = 1 To objGridItens.iLinhasExistentes
'                If Len(GridItens.TextMatrix(iLinha, iGrid_DescProduto_Col)) > 0 Then
'                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_DescProduto_Col)
'                Else
'                    objMnemonicoValor.colValor.Add ""
'                End If
'            Next
'
'        Case FILIAL1
'            If Len(Filial.Text) > 0 Then
'                objMnemonicoValor.colValor.Add Filial.Text
'            Else
'                objMnemonicoValor.colValor.Add ""
'            End If
'
'        Case CLIENTE1
'            If Len(Cliente.Text) > 0 Then
'                objMnemonicoValor.colValor.Add Cliente.Text
'            Else
'                objMnemonicoValor.colValor.Add ""
'            End If
'
'        Case CLIENTE_CODIGO
'            If Len(Cliente.Text) > 0 Then
'
'                objCliente.sNomeReduzido = Cliente.Text
'
'                lErro = Cliente_Le_NomeReduzido(objCliente)
'                If lErro <> SUCESSO Then gError 88028
'
'                objMnemonicoValor.colValor.Add objCliente.lCodigo
'            Else
'                objMnemonicoValor.colValor.Add 0
'            End If
'
'        Case NATUREZA_OP
'            If Len(NatOpInterna.Text) > 0 Then
'                objMnemonicoValor.colValor.Add NatOpInterna.Text
'            Else
'                objMnemonicoValor.colValor.Add ""
'            End If
'
'        Case PRODUTO1
'            For iLinha = 1 To objGridItens.iLinhasExistentes
'                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
'                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_Produto_Col)
'                Else
'                    objMnemonicoValor.colValor.Add ""
'                End If
'            Next
'
'        Case QUANTIDADE1
'            For iLinha = 1 To objGridItens.iLinhasExistentes
'                If Len(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
'
'                    'Guarda os valores que serão passados como parâmetros em UMEstoque_Conversao
'                    sProduto = GridItens.TextMatrix(iLinha, iGrid_Produto_Col)
'                    sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
'                    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
'
'                    'Converte a quantidade para UM padrão estoque
'                    lErro = UMEstoque_Conversao(sProduto, sUM, dQuantidade, dQuantidadeConvertida)
'                    If lErro <> SUCESSO Then gError 79943
'
'                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
'
'                Else
'                    objMnemonicoValor.colValor.Add 0
'                End If
'            Next
'
'        Case Serie1
'            If Len(Serie.Text) > 0 Then
'                objMnemonicoValor.colValor.Add Serie.Text
'            Else
'                objMnemonicoValor.colValor.Add ""
'            End If
'
'        Case TIPO_NOTA_FISCAL
'            If Len(TipoNFiscal.Text) > 0 Then
'                objMnemonicoValor.colValor.Add TipoNFiscal.Text
'            Else
'                objMnemonicoValor.colValor.Add ""
'            End If
'
'        Case UNIDADE_MED
'            For iLinha = 1 To objGridItens.iLinhasExistentes
'                If Len(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
'                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
'                Else
'                    objMnemonicoValor.colValor.Add ""
'                End If
'            Next
'
'        Case VALOR_TOTAL
'            If Len(ValorTotal.Caption) > 0 Then
'                objMnemonicoValor.colValor.Add CDbl(ValorTotal.Caption)
'            Else
'                objMnemonicoValor.colValor.Add 0
'            End If
'
'        Case PRECO_UNITARIO
'            For iLinha = 1 To objGridItens.iLinhasExistentes
'                If Len(GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col)) > 0 Then
'                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoUnitario_Col))
'                Else
'                    objMnemonicoValor.colValor.Add 0
'                End If
'            Next
'
'        Case PRECO_TOTAL
'            For iLinha = 1 To objGridItens.iLinhasExistentes
'                If Len(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col)) > 0 Then
'                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col))
'                Else
'                    objMnemonicoValor.colValor.Add 0
'                End If
'            Next
'
'        Case ICMS
'            If Len(ICMSValor1.Caption) > 0 Then
'                objMnemonicoValor.colValor.Add CDbl(ICMSValor1.Caption)
'            Else
'                objMnemonicoValor.colValor.Add 0
'            End If
'
'        Case ICMSSUBST
'            If Len(ICMSSubstValor1.Caption) > 0 Then
'                objMnemonicoValor.colValor.Add CDbl(ICMSSubstValor1.Caption)
'            Else
'                objMnemonicoValor.colValor.Add 0
'            End If
'
'        Case VALOR_FRETE
'            If Len(ValorFrete.Text) > 0 Then
'                objMnemonicoValor.colValor.Add CDbl(ValorFrete.Text)
'            Else
'                objMnemonicoValor.colValor.Add 0
'            End If
'
'        Case VALOR_SEGURO
'            If Len(ValorSeguro.Text) > 0 Then
'                objMnemonicoValor.colValor.Add CDbl(ValorSeguro.Text)
'            Else
'                objMnemonicoValor.colValor.Add 0
'            End If
'
'        Case VALOR_DESPESAS
'            If Len(ValorDespesas.Text) > 0 Then
'                objMnemonicoValor.colValor.Add CDbl(ValorDespesas.Text)
'            Else
'                objMnemonicoValor.colValor.Add 0
'            End If
'
'        Case IPI
'            If Len(IPIValor1.Caption) > 0 Then
'                objMnemonicoValor.colValor.Add CDbl(IPIValor1.Caption)
'            Else
'                objMnemonicoValor.colValor.Add 0
'            End If
'
'        Case VALOR_DESCONTO
'            If Len(ValorDesconto.Text) > 0 Then
'                objMnemonicoValor.colValor.Add CDbl(ValorDesconto.Text)
'            Else
'                objMnemonicoValor.colValor.Add 0
'            End If
'
'        Case ISS_VALOR
'            If Len(ISSValor.Text) > 0 Then
'                objMnemonicoValor.colValor.Add CDbl(ISSValor.Text)
'            Else
'                objMnemonicoValor.colValor.Add 0
'            End If
'
'        Case ISS_INCLUSO
'            objMnemonicoValor.colValor.Add ISSIncluso.Value
'
'        Case VALOR_IRRF
'            If Len(ValorIRRF.Text) > 0 Then
'                objMnemonicoValor.colValor.Add CDbl(ValorIRRF.Text)
'            Else
'                objMnemonicoValor.colValor.Add 0
'            End If
'
'        Case VALOR_INSS
'            If Len(INSSValor.Text) > 0 Then
'                objMnemonicoValor.colValor.Add CDbl(INSSValor.Text)
'            Else
'                objMnemonicoValor.colValor.Add 0
'            End If
'
'        Case INSS_RETIDO
'            objMnemonicoValor.colValor.Add INSSRetido.Value
'
'        Case VALOR_PRODUTOS
'            If Len(ValorProdutos.Caption) > 0 Then
'                objMnemonicoValor.colValor.Add CDbl(ValorProdutos.Caption)
'            Else
'                objMnemonicoValor.colValor.Add 0
'            End If
'
'        Case PRODUTO_ALMOX
'            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
'                If Len(GridAlocacao.TextMatrix(iLinha, iGrid_ProdutoAloc_Col)) > 0 Then
'                    objMnemonicoValor.colValor.Add GridAlocacao.TextMatrix(iLinha, iGrid_ProdutoAloc_Col)
'                Else
'                    objMnemonicoValor.colValor.Add ""
'                End If
'            Next
'
'        Case QUANT_ALOCADA
'                For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
'                    If Len(Trim(GridAlocacao.TextMatrix(iLinha, iGrid_QuantAloc_Col))) > 0 Then
'                        If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
'                            objMnemonicoValor.colValor.Add CDbl(GridAlocacao.TextMatrix(iLinha, iGrid_QuantAloc_Col))
'                        Else
'                            objMnemonicoValor.colValor.Add 0
'                        End If
'                    Else
'                        objMnemonicoValor.colValor.Add 0
'                    End If
'                Next
'
'        Case UNIDADE_MED_EST
'            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
'                If Len(GridAlocacao.TextMatrix(iLinha, iGrid_UMAloc_Col)) > 0 Then
'                    objMnemonicoValor.colValor.Add GridAlocacao.TextMatrix(iLinha, iGrid_UMAloc_Col)
'                Else
'                    objMnemonicoValor.colValor.Add ""
'                End If
'            Next
'
'        Case PERC_ACRESC_FIN
'            If Len(PercAcrescFin.Text) > 0 Then
'                objMnemonicoValor.colValor.Add CDbl(PercAcrescFin.Text)
'            Else
'                objMnemonicoValor.colValor.Add 0
'            End If
'
'        Case VALOR_PARCELA
'            For iLinha = 1 To objGridParcelas.iLinhasExistentes
'                If Len(GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col)) > 0 Then
'                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col))
'                Else
'                    objMnemonicoValor.colValor.Add 0
'                End If
'            Next
'
'        Case CTACONTABILEST1
'
'            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
'                If Len(Trim(GridAlocacao.TextMatrix(iLinha, iGrid_QuantAloc_Col))) > 0 Then
'
'                    If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
'
'                        objAlmoxarifado.sNomeReduzido = CStr(objMnemonicoValor.vParam(3))
'
'                        'Lê o Nome Reduzido do Almoxarifado
'                        lErro = Almoxarifado_Le_NomeReduzido(objAlmoxarifado)
'                        If lErro <> SUCESSO And lErro <> 25060 Then gError 64180
'
'                        'Se não encontrou ===> Erro
'                        If lErro = SUCESSO Then
'
'                            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
'
'                            lErro = Produto_Formata(objMnemonicoValor.vParam(1), sProdutoFormatado, iPreenchido)
'                            If iPreenchido <> PRODUTO_PREENCHIDO Then gError 64179
'
'                            objEstoqueProduto.sProduto = sProdutoFormatado
'
'                            'Lê a conta contábil do Produto no Almoxarifado
'                            lErro = EstoqueProdutoCC_Le(objEstoqueProduto)
'                            If lErro <> SUCESSO And lErro <> 49991 Then gError 64181
'
'                            If Len(Trim(objEstoqueProduto.sContaContabil)) > 0 Then
'
'                                sContaMascarada = String(STRING_CONTA, 0)
'
'                                lErro = Mascara_MascararConta(objEstoqueProduto.sContaContabil, sContaMascarada)
'                                If lErro <> SUCESSO Then gError 64182
'
'                                objMnemonicoValor.colValor.Add sContaMascarada
'                            Else
'                                objMnemonicoValor.colValor.Add ""
'                            End If
'                        Else
'                            objMnemonicoValor.colValor.Add ""
'                        End If
'                    Else
'                        objMnemonicoValor.colValor.Add ""
'                    End If
'                Else
'                    objMnemonicoValor.colValor.Add ""
'                End If
'            Next
'
'        Case Else
'            gError 39675
'
'    End Select
'
'    Calcula_Mnemonico = SUCESSO
'
'    Exit Function
'
'Erro_Calcula_Mnemonico:
'
'    Calcula_Mnemonico = gErr
'
'    Select Case gErr
'
'        Case 39675
'            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
'
'        Case 64179, 64180, 64181, 64182, 64183, 64184, 64185, 75465, 75467, 75468, 75470, 79007, 79008, 79009, 88028
'
'        Case 75469
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function TabelaPreco_Atualiza_PrecoUnitario(objProduto As ClassProduto) As Long
'
'Dim lErro As Long
'Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
'Dim dPrecoUnitario As Double
'
'On Error GoTo Erro_TabelaPreco_Atualiza_PrecoUnitario
'
'    'Preco Unitário
'    If Len(Trim(TabelaPreco.Text)) > 0 Then
'        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
'        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
'        objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
'
'        'Lê o Item da Tabela de Preço
'        lErro = TabelaPrecoItem_Le(objTabelaPrecoItem)
'        If lErro <> SUCESSO And lErro <> 28014 Then gError 42187
'
'        If lErro = 28014 Then
'            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
'
'            'Lê o Item da Tabela de Preço da Empresa Toda
'            lErro = TabelaPrecoItem_Le(objTabelaPrecoItem)
'            If lErro <> SUCESSO And lErro <> 28014 Then gError 42188
'
'        End If
'
'        If lErro <> 28014 Then
'
'            'Calcula o Preço Unitário
'            lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
'            If lErro <> SUCESSO Then gError 42189
'
'            GridItens.TextMatrix(GridItens.Row, iGrid_PrecoUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
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
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Sub Seleciona_Tipo_Venda()
''Seleciona na combo de tipos de notas fiscais o tipo "Venda"
'
'Dim iIndice As Integer
'
'    For iIndice = 0 To TipoNFiscal.ListCount - 1
'        If TipoNFiscal.ItemData(iIndice) = DOCINFO_NFISFV Then
'            TipoNFiscal.ListIndex = iIndice
'            Exit For
'        End If
'    Next
'
'    Exit Sub
'
'End Sub
'
'Function Preenche_DescontoPadrao(iLinha As Integer) As Long
'
'Dim lErro As Long
'Dim colDescontoPadrao As New colDesconto
'Dim iIndice1 As Integer
'Dim iIndice2 As Integer
'Dim iColuna  As Integer
'Dim dtDataVencimento As Date
'Dim dPercentual As Double
'Dim dValorParcela As Double
'Dim sValorDesconto As String
'
'On Error GoTo Erro_Preenche_DescontoPadrao
'
'    'Se a data de referencia estiver preenchida
'    If Len(Trim(DataReferencia.ClipText)) > 0 Then
'
'        dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_col))
'        lErro = Parcela_GeraDescontoPadrao(colDescontoPadrao, dtDataVencimento)
'        If lErro <> SUCESSO Then gError 51116
'
'        If colDescontoPadrao.Count > 0 Then
'
'            'Para cada um dos desontos padrão
'            For iIndice1 = 1 To colDescontoPadrao.Count
'
'                'Seleciona a coluna correspondente ao Desconto
'                If iIndice1 = 1 Then iColuna = iGrid_Desc1Codigo_Col
'                If iIndice1 = 2 Then iColuna = iGrid_Desc2Codigo_Col
'                If iIndice1 = 3 Then iColuna = iGrid_Desc3Codigo_Col
'
'                'Seleciona o tipo de desconto
'                For iIndice2 = 0 To Desconto1Codigo.ListCount - 1
'                    If colDescontoPadrao.Item(iIndice1).iCodigo = Desconto1Codigo.ItemData(iIndice2) Then
'                        GridParcelas.TextMatrix(iLinha, iColuna) = Desconto1Codigo.List(iIndice2)
'                        GridParcelas.TextMatrix(iLinha, iColuna + 1) = Format(colDescontoPadrao.Item(iIndice1).dtData, "dd/mm/yyyy")
'                        GridParcelas.TextMatrix(iLinha, iColuna + 3) = Format(colDescontoPadrao.Item(iIndice1).dValor, "Percent")
'
'                        '*** Inicio colocacao Valor Desconto na tela
'                        dPercentual = colDescontoPadrao.Item(iIndice1).dValor
'                        dValorParcela = StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col))
'
'                        'Coloca Valor do Desconto na tela
'                        If dValorParcela > 0 Then
'                            sValorDesconto = Format(dPercentual * dValorParcela, "Standard")
'                            GridParcelas.TextMatrix(iLinha, iColuna + 2) = sValorDesconto
'                        End If
'                        '*** Fim colocacao Valor Desconto na tela
'
'                    End If
'                Next
'            Next
'
'        End If
'
'    End If
'
'    Preenche_DescontoPadrao = SUCESSO
'
'    Exit Function
'
'Erro_Preenche_DescontoPadrao:
'
'    Preenche_DescontoPadrao = gErr
'
'    Select Case gErr
'
'        Case 51116
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Function
'
'End Function
'
''**** inicio do trecho a ser copiado *****
'Public Function Form_Load_Ocx() As Object
'
'    Parent.HelpContextID = IDH_NF_FATURA_SAIDA_DADOS_PRINCIPAIS
'    Set Form_Load_Ocx = Me
'    Caption = "Nota Fiscal Fatura de Saída"
'    Call Form_Load
'
'End Function
'
'Public Function Name() As String
'
'    Name = "NFiscalFatura"
'
'End Function
'
'Public Sub Show()
'    Parent.Show
'    Parent.SetFocus
'End Sub
'
'
'
'Private Sub Unload(objme As Object)
'   ' Parent.UnloadDoFilho
'
'   RaiseEvent Unload
'
'End Sub
'
'Public Property Get Caption() As String
'    Caption = m_Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As String)
'    Parent.Caption = New_Caption
'    m_Caption = New_Caption
'End Property
'
''***** fim do trecho a ser copiado ******
'
'Private Function DescontoGlobal_Testa(dValorDesconto As Double) As Long
'
'Dim dValorProdutos As Double
'Dim dValorFrete As Double
'Dim dValorSeguro As Double
'Dim dValorDespesas As Double
'
'Dim lErro As Long
'
'On Error GoTo Erro_DescontoGlobal_Testa
'
'    If Len(Trim(ValorProdutos.Caption)) > 0 Then dValorProdutos = CDbl(ValorProdutos.Caption)
'    If Len(Trim(ValorFrete.ClipText)) > 0 Then dValorFrete = CDbl(ValorFrete.Text)
'    If Len(Trim(ValorSeguro.ClipText)) > 0 Then dValorSeguro = CDbl(ValorSeguro.Text)
'    If Len(Trim(ValorDespesas.ClipText)) > 0 Then dValorDespesas = CDbl(ValorDespesas.Text)
'
'    'Verifica se o Valor do Desconto é maior que a soma dos demais valores
'    If dValorDesconto > dValorProdutos + dValorFrete + dValorSeguro + dValorDespesas Then gError 56899
'
'    DescontoGlobal_Testa = SUCESSO
'
'    Exit Function
'
'Erro_DescontoGlobal_Testa:
'
'    DescontoGlobal_Testa = gErr
'
'    Select Case gErr
'
'        Case 56899
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_ULTRAPASSOU_SOMA_VALORES", gErr, dValorDesconto, dValorProdutos + dValorFrete + dValorSeguro + dValorDespesas)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = KEYCODE_BROWSER Then
'
'        If Me.ActiveControl Is NatOpInterna Then
'            Call LblNatOpInterna_Click
'        ElseIf Me.ActiveControl Is Serie Then
'            Call SerieLabel_Click
'        ElseIf Me.ActiveControl Is NFiscal Then
'            Call NFiscalLabel_Click
'        ElseIf Me.ActiveControl Is Cliente Then
'            Call ClienteLabel_Click
'        ElseIf Me.ActiveControl Is Produto Then
'            Call BotaoProdutos_Click
'        ElseIf Me.ActiveControl Is Transportadora Then
'            Call TransportadoraLabel_Click
'        ElseIf Me.ActiveControl Is SerieNFiscalOriginal Then
'            Call SerieNFOriginalLabel_Click
'        ElseIf Me.ActiveControl Is NFiscalOriginal Then
'            Call NFiscalOriginalLabel_Click
'        ElseIf Me.ActiveControl Is Vendedor Then
'            Call BotaoVendedores_Click
'        Else
'            Call gobjRastreamento.objUserControl_KeyDown(KeyCode, Shift)
'        End If
'
'    End If
'
'End Sub
'
'
'Public Sub VolumeQuant_GotFocus()
'
'    Call MaskEdBox_TrataGotFocus(VolumeQuant, iAlterado)
'
'End Sub
'
'Public Sub DataReferencia_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'    giDataReferenciaAlterada = REGISTRO_ALTERADO
'
'End Sub
'
'Public Sub DataReferencia_GotFocus()
'
'Dim iDataAux As Integer
'
'    iDataAux = giDataReferenciaAlterada
'    Call MaskEdBox_TrataGotFocus(DataReferencia, iAlterado)
'    giDataReferenciaAlterada = iDataAux
'
'End Sub
'
'Public Sub DataReferencia_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim dtDataEmissao As Date
'Dim dtDataReferencia As Date
'Dim objCondicaoPagto As New ClassCondicaoPagto
'
'On Error GoTo Erro_DataReferencia_Validate
'
'    If giDataReferenciaAlterada <> REGISTRO_ALTERADO Then Exit Sub
'
'    If Len(Trim(DataReferencia.ClipText)) > 0 Then
'
'        'Critica a data digitada
'        lErro = Data_Critica(DataReferencia.Text)
'        If lErro <> SUCESSO Then gError 26713
'
'        'Compara com data de emissão
'        If Len(Trim(DataEmissao.ClipText)) > 0 Then
'
'            dtDataEmissao = CDate(DataEmissao.Text)
'            dtDataReferencia = CDate(DataReferencia.Text)
'
'            If dtDataEmissao > dtDataReferencia Then gError 26714
'
'        End If
'
'
'    End If
'
'    giDataReferenciaAlterada = 0
'
'    'Preenche o GridParcelas
'    lErro = Cobranca_Automatica()
'    If lErro <> SUCESSO Then gError 25436
'
'    Exit Sub
'
'Erro_DataReferencia_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 25436 'Tratado na rotina chamada
'
'        Case 26713
'
'        Case 26714
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR_DATAREFERENCIA", gErr, dtDataReferencia, dtDataEmissao)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub objEventoMensagem_evSelecao(obj1 As Object)
'
'Dim objMensagem As ClassMensagem
'
'    Set objMensagem = obj1
'
'    'Traz a mensagem NFiscal para a tela
'    Mensagem.Text = objMensagem.sDescricao
'
'    Me.Show
'
'End Sub
'
'Public Property Get objUserControl() As Object
'    Set objUserControl = m_objUserControl
'End Property
'
'Public Property Set objUserControl(ByVal vData As Object)
'    Set m_objUserControl = vData
'End Property
'
''Devolve Parent do User Control
'Public Property Get Parent() As Object
'    Set Parent = objUserControl.Parent
'End Property
'
'Public Property Get Controls() As Object
'    Set Controls = objUserControl.Controls
'End Property
'
'Public Property Get ActiveControl() As Object
'    Set ActiveControl = objUserControl.ActiveControl
'End Property
'
'Public Property Get Enabled() As Boolean
'    Enabled = objUserControl.Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    objUserControl.Enabled = New_Enabled
'End Property
'
'
'Private Sub AlocacoesNF_Agrupa(colAlocacoes As ColAlocacoesItemNF)
'
'Dim iIndice As Integer
'Dim iIndice1 As Integer
'
'    If colAlocacoes.Count <= 1 Then Exit Sub
'
'    'se a alocação se referir a uma transferencia ==> não leva em consideração
'    For iIndice = colAlocacoes.Count To 1 Step -1
'        If colAlocacoes.Item(iIndice).iTransferencia = TIPOMOV_EST_TRANSFERENCIA Then
'            colAlocacoes.Remove iIndice
'        End If
'    Next
'
'    For iIndice = colAlocacoes.Count To 2 Step -1
'
'        For iIndice1 = 1 To iIndice - 1
'            If (colAlocacoes.Item(iIndice).iAlmoxarifado = colAlocacoes.Item(iIndice1).iAlmoxarifado) Then
'                colAlocacoes.Item(iIndice1).dQuantidade = colAlocacoes.Item(iIndice1).dQuantidade + colAlocacoes.Item(iIndice).dQuantidade
'                colAlocacoes.Remove iIndice
'                Exit For
'            End If
'        Next
'
'    Next
'
'End Sub
'
'Public Sub BotaoConsultaTitRec_Click()
''Abre uma tela para consulta do DocCPR vinculado à nota fiscal
'
'Dim lErro As Long
'Dim objNFiscal As New ClassNFiscal
'Dim objCliente As New ClassCliente
'
'On Error GoTo Erro_BotaoConsultaTitRec_Click
'
'    'Verifica se todos os campos necessários para se efetuar a consulta foram preenchidos
'    lErro = Critica_CamposNecessarios_ConsultaTitulo()
'    If lErro <> SUCESSO Then gError 79795
'
'    'Guarda no objNFiscal os dados necessários para consultar o título
'    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
'    objNFiscal.sSerie = Serie.Text
'    objNFiscal.iTipoDocInfo = Codigo_Extrai(TipoNFiscal.Text)
'    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
'    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
'    objNFiscal.dtDataEntrada = DATA_NULA
'
'    'Guarda em objFornecedor o nome reduzido do Fornecedor
'    objCliente.sNomeReduzido = Cliente.Text
'
'    'Lê o código do Fornecedor a partir do nome reduzido obtido na tela
'    lErro = Cliente_Le_NomeReduzido(objCliente)
'    If lErro <> SUCESSO And lErro <> 12348 Then gError 79796
'
'    'Se não encontrou o fornecedor = > erro
'    If lErro = 12348 Then gError 79797
'
'    'Guarda no objNFiscal o código do fornecedor
'    objNFiscal.lCliente = objCliente.lCodigo
'
'    'Lê o NumIntDocCPR da NFiscal e exibe o documento gerado no CPR por essa nota
'    lErro = NFiscal_Consulta_DocCPR(objNFiscal)
'    If lErro <> SUCESSO And lErro <> 79717 Then gError 79798
'
'    'Se não encontrou a nota => erro
'    If lErro = 79717 Then gError 79799
'
'    Exit Sub
'
'Erro_BotaoConsultaTitRec_Click:
'
'    Select Case gErr
'
'        Case 79795, 79798, 79796
'
'        Case 79799
'            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA2", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada)
'
'        Case 79797
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Function Critica_CamposNecessarios_ConsultaTitulo() As Long
''Verifica se os campos necessários para encontrar consultar um título gerado por uma NFFatEntrada foram preenchidos
'
'Dim lErro As Long
'
'On Error GoTo Erro_Critica_CamposNecessarios_ConsultaTitulo
'
'    'Verifica se o Cliente foi preenchido
'    If Len(Trim(Cliente.Text)) = 0 Then gError 79800
'
'    'Verifica se a filial do Fornecedor foi preenchida
'    If Len(Trim(Filial.Text)) = 0 Then gError 79801
'
'    'Verifica se a Série foi preenchida
'    If Len(Trim(Serie.Text)) = 0 Then gError 79802
'
'    'Verifica se o Número da Nota foi preenchido
'    If Len(Trim(NFiscal.Caption)) = 0 Then gError 79803
'
'    'Verifica se a data de emissão da nota foi preenchida
'    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 79804
'
'    'Verifica se o tipo da Nota fiscal foi selecionado
'    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 79805
'
'    Critica_CamposNecessarios_ConsultaTitulo = SUCESSO
'
'    Exit Function
'
'Erro_Critica_CamposNecessarios_ConsultaTitulo:
'
'    Critica_CamposNecessarios_ConsultaTitulo = gErr
'
'    Select Case gErr
'
'        Case 79800
'            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
'
'        Case 79801
'            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
'
'        Case 79802
'            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
'
'        Case 79803
'            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
'
'        Case 79804
'            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
'
'        Case 79805
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
'
'        Case Else
'
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
''Transferir para ClassCRFAT
'Function Cliente_Pesquisa_NomeReduzido(objCliente As Object, sNomeReduzidoParte As String) As Long
''rotina para trazer o nome do cliente com uma parte dos caracteres digitados
'
'Dim iCursorPos As Integer
'Dim sNomeReduzido As String
'Dim lErro As Long
'
'On Error GoTo Erro_Cliente_Pesquisa_NomeReduzido
'
'        If sNomeReduzidoParte <> objCliente.Text And Len(sNomeReduzidoParte) <= Len(objCliente.Text) Then
'
'            sNomeReduzidoParte = objCliente.Text
'
'            lErro = Cliente_Le_NomeReduzido_Parcial(sNomeReduzidoParte, sNomeReduzido)
'            If lErro <> SUCESSO Then gError 71652
'
'            If sNomeReduzido <> objCliente.Text Then
'
'                sNomeReduzidoParte = sNomeReduzido
'                iCursorPos = objCliente.SelStart
'                objCliente.Text = sNomeReduzido
'                objCliente.SelStart = iCursorPos
'                objCliente.SelLength = Len(objCliente.Text) - (iCursorPos)
'
'            End If
'
'        Else
'
'            sNomeReduzidoParte = objCliente.Text
'
'        End If
'
'    Cliente_Pesquisa_NomeReduzido = SUCESSO
'
'    Exit Function
'
'Erro_Cliente_Pesquisa_NomeReduzido:
'
'    Cliente_Pesquisa_NomeReduzido = gErr
'
'    Select Case gErr
'
'        Case 71652
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'End Function
'
''Transferir para ClassCRFAT
'Function Cliente_Le_NomeReduzido_Parcial(sNomeReduzidoParte As String, sNomeReduzido As String) As Long
''Le o NomeReduzido do cliente a partir de parte do nome reduzido. Se a parte não individualiza o cliente ==> retorna codigo de erro
'
'Dim lComando As Long
'Dim lErro As Long
'Dim iTamanho As Integer
'Dim sNomeReduzidoAux As String
'
'On Error GoTo Erro_Cliente_Le_NomeReduzido_Parcial
'
'    sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
'
'    iTamanho = Len(sNomeReduzidoParte)
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 71645
'
'    'Le o NomeReduzido da tabela Clientes
'    lErro = Comando_Executar(lComando, "SELECT NomeReduzido FROM Clientes WHERE NomeReduzido >= ?", sNomeReduzido, sNomeReduzidoParte)
'    If lErro <> AD_SQL_SUCESSO Then gError 71646
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71647
'
'    'se não encontrou nenhum cliente com este prefixo ==> erro
'    If lErro = AD_SQL_SEM_DADOS Then gError 71648
'
'    'se encontrou um cliente, mas o prefixo difere ==> erro
'    If UCase(sNomeReduzidoParte) <> UCase(Left(sNomeReduzido, iTamanho)) Then gError 71649
'
'    sNomeReduzidoAux = sNomeReduzido
'
'    lErro = Comando_BuscarProximo(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 71650
'
'    If lErro = AD_SQL_SUCESSO Then
'
'        'se encontrou um segundo registro com o mesmo prefixo ==> ainda não pode trazer o nome pois ainda não conseguiu individualizar
'        If UCase(sNomeReduzidoParte) = UCase(Left(sNomeReduzido, iTamanho)) Then gError 71651
'
'    End If
'
'    sNomeReduzido = sNomeReduzidoAux
'
'    lErro = Comando_Fechar(lComando)
'
'    Cliente_Le_NomeReduzido_Parcial = SUCESSO
'
'    Exit Function
'
'Erro_Cliente_Le_NomeReduzido_Parcial:
'
'    Cliente_Le_NomeReduzido_Parcial = gErr
'
'    Select Case gErr
'
'        Case 71645
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 71646, 71647, 71650
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)
'
'        Case 71648, 71649, 71651
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
''*** Fernando, essa função também está na tela BaixasRecOcx e deve ser apagada daqui depois que for transportada para o RotinasContab
'
'Private Function Obtem_Subtipo_Contabil(sTabela As String, iCodigo As Integer, iSubtipo As Integer) As Long
'
'Dim lErro As Long
'Dim lComando As Long
'Dim sSQL As String
'
'On Error GoTo Erro_Obtem_Subtipo_Contabil
'
'    'Abre o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 79668
'
'    'Define o comando SQL que será utilizado
'    sSQL = "SELECT SubTipoContabil FROM " & sTabela & " WHERE Codigo = ?"
'
'    'Lê o subtipo na tabela passada como parâmetro
'    lErro = Comando_Executar(lComando, sSQL, iSubtipo, iCodigo)
'    If lErro <> AD_SQL_SUCESSO Then gError 79669
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79670
'
'    'Se não encontrou => erro
'    If lErro = AD_SQL_SEM_DADOS Then gError 79671
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    Obtem_Subtipo_Contabil = SUCESSO
'
'    Exit Function
'
'Erro_Obtem_Subtipo_Contabil:
'
'    Obtem_Subtipo_Contabil = gErr
'
'    Select Case gErr
'
'        Case 79668
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 79669, 79670
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA", gErr, sTabela)
'
'        Case 79671 'Subtipo não encontrado. Deve ser tratado na rotina chamadora
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
''*** Fim do trecho a ser apagado ***
'
''*** Fernando, essas funções foram copiadas do CTNFiscalFatEntrada e devem ser apagadas depois que forem copiadas para o Rotinas ***
'
'Public Function NFiscal_Consulta_DocCPR(objNFiscal As ClassNFiscal) As Long
''Exibe na tela um documento (título a pagar, título a receber, crédito a pagar, etc.), vinculado à uma nota fiscal
'
'Dim lErro As Long
'Dim sNomeTela As String
'Dim objTipoDocInfo As New ClassTipoDocInfo
'Dim objTituloPagar As New ClassTituloPagar
'Dim objTituloReceber As New ClassTituloReceber
'Dim objCreditoPagar As New ClassCreditoPagar
'Dim objDebitoReceber As New ClassDebitoRecCli
'Dim objNFsPag As New ClassNFsPag
'
'On Error GoTo Erro_NFiscal_Consulta_DocCPR
'
'    'Obtem ClasseDocCPR e NumIntDocCPR da Nota Fiscal
'    lErro = NFiscal_Le_SemNumIntDoc(objNFiscal)
'    If lErro <> SUCESSO And lErro <> 79714 Then gError 79716
'
'    'Se não encontrou a Nota Fiscal => erro
'    If lErro = 79714 Then gError 79717
'
'    'Define o parâmetro que será utilizado para ler a sigla do DocCPR
'    objTipoDocInfo.iCodigo = objNFiscal.iTipoDocInfo
'
'    'Lê os dados do TipoDocInfo no BD (a intenção é pegar a sigla do DocCPR)
'    lErro = TipoDocInfo_Le_Codigo(objTipoDocInfo)
'    If lErro <> SUCESSO And lErro <> 31415 Then gError 79729
'
'    'Se não encontrou => erro
'    If lErro = 31415 Then gError 79730
'
'
'    'Seleciona o nome da tela que será aberta
'    Select Case objNFiscal.iClasseDocCPR
'
'        'Consulta de títulos a pagar
'        Case CPR_TITULO_PAGAR
'
'            'Chama a tela de consulta de Títulos a Pagar
'            lErro = NFiscal_Consulta_TituloPagar(objNFiscal, objTipoDocInfo)
'            If lErro <> SUCESSO Then gError 79731
'
'        'Consulta de títulos a receber
'        Case CPR_TITULO_RECEBER
'
'            'Chama a tela de consulta de Títulos a Receber
'            lErro = NFiscal_Consulta_TituloReceber(objNFiscal)
'            If lErro <> SUCESSO Then gError 79732
'
'        'Consulta de notas fiscais no cp (daqui o usuário pode ir para a tela de faturas)
'        Case CPR_NF_PAGAR
'
'            lErro = NFiscal_Consulta_NFPag(objNFiscal)
'            If lErro <> SUCESSO Then gError 79733
'
'        'Tela de Cadastro de Débitos a Receber
'        Case CPR_DEBITO_RECEBER
'
'            lErro = NFiscal_Consulta_DebitosReceb(objNFiscal)
'            If lErro <> SUCESSO Then gError 79734
'
'        'Tela de Cadastro de Débitos a Receber
'        Case CPR_CREDITO_PAGAR
'
'            lErro = NFiscal_Consulta_CreditoPagar(objNFiscal)
'            If lErro <> SUCESSO Then gError 79735
'
'        'O título não tem uma tela de Doc Correspondente
'        Case Else
'            gError 79715
'
'    End Select
'
'    NFiscal_Consulta_DocCPR = SUCESSO
'
'    Exit Function
'
'Erro_NFiscal_Consulta_DocCPR:
'
'    NFiscal_Consulta_DocCPR = gErr
'
'    Select Case gErr
'
'        Case 79716, 79717, 79729, 79731 To 79735
'
'        Case 79715
'            Call Rotina_Erro(vbOKOnly, "ERRO_NF_SEM_DOCCPR_VINCULADO", gErr)
'
'        Case 79730
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO2", gErr, objTipoDocInfo.iCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function NFiscal_Le_SemNumIntDoc(objNFiscal As ClassNFiscal) As Long
''Lê os dados da nota fiscal a partir dos dados Numero, Serie, FilialEmpresa, Tipo, Fornecedor ou Cliente, FilialForn ou FilialCli, DataEmissao, DataEntrada
'
'Dim lErro As Long
'Dim lComando As Long
'Dim sSelecaoSQL As String
'Dim tNFiscal As typeNFiscal
'
'On Error GoTo Erro_NFiscal_Le_SemNumIntDoc
'
'    'Abre o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 79711
'
'    With tNFiscal
'
'        'Inicializa a string que receberá a Série
'        .sNumPedidoTerc = String(STRING_NUM_PEDIDO_TERC, 0)
'        .sMensagemNota = String(STRING_NFISCAL_MENSAGEM, 0)
'        .sNaturezaOp = String(STRING_NATUREZAOP_CODIGO, 0)
'        .sPlaca = String(STRING_NFISCAL_PLACA, 0)
'        .sPlacaUF = String(STRING_NFISCAL_PLACA_UF, 0)
'        .sVolumeEspecie = String(STRING_NFISCAL_VOLUME_ESPECIE, 0)
'        .sVolumeMarca = String(STRING_NFISCAL_VOLUME_MARCA, 0)
'        .sVolumeNumero = String(STRING_NFISCAL_VOLUME_NUMERO, 0)
'        .sObservacao = String(STRING_NFISCAL_OBSERVACAO, 0)
'        .sCodUsuarioCancel = String(STRING_USUARIO_CODIGO, 0)
'        .sMotivoCancel = String(STRING_NFISCAL_MOTIVOCANCEL, 0)
'
'        'Define o comando SQL que será passado no select
'        sSelecaoSQL = "SELECT NumIntDoc, Status, FilialEmpresa, FilialEntrega, DataVencimento, DataReferencia, FilialPedido, NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, VolumeNumero, Canal, NumIntNotaOriginal, ClienteBenef, FilialCliBenef, FornecedorBenef, FilialFornBenef, FreteRespons, NumRecebimento, Observacao, CodUsuarioCancel, MotivoCancel FROM NFiscal WHERE NumNotaFiscal = ? AND FilialEmpresa = ? AND Serie = ? AND Fornecedor = ? AND Cliente = ? AND FilialForn = ? AND FilialCli = ? AND DataEmissao = ? AND DataEntrada = ? AND TipoNFiscal = ? AND Status <> ?"
'
'        'Busca no BD os campos necessários para se definir a tela e o doc que será exibido
'        lErro = Comando_Executar(lComando, sSelecaoSQL, .lNumIntDoc, .iStatus, .iFilialEmpresa, .iFilialEntrega, .dtDataVencimento, .dtDataReferencia, .iFilialPedido, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .sVolumeEspecie, .sVolumeMarca, .sVolumeNumero, .iCanal, .lNumIntNotaOriginal, .lClienteBenef, .iFilialCliBenef, .lFornecedorBenef, .iFilialFornBenef, .iFreteRespons, .lNumRecebimento, .sObservacao, .sCodUsuarioCancel, .sMotivoCancel, objNFiscal.lNumNotaFiscal, giFilialEmpresa, objNFiscal.sSerie, objNFiscal.lFornecedor, objNFiscal.lCliente, objNFiscal.iFilialForn, objNFiscal.iFilialCli, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada, objNFiscal.iTipoDocInfo, STATUS_CANCELADO)
'        If lErro <> AD_SQL_SUCESSO Then gError 79712
'
'        lErro = Comando_BuscarPrimeiro(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 79713
'
'        'Se não encontrou a NFiscal = > erro
'        If lErro = AD_SQL_SEM_DADOS Then gError 79714
'
'    End With
'
'    'Guarda no objNFiscal os dados obtidos no select
'    Call Move_NFiscal_Obj(objNFiscal, tNFiscal)
'
'    Call Comando_Fechar(lComando)
'
'    NFiscal_Le_SemNumIntDoc = SUCESSO
'
'    Exit Function
'
'Erro_NFiscal_Le_SemNumIntDoc:
'
'    NFiscal_Le_SemNumIntDoc = gErr
'
'    Select Case gErr
'
'        Case 79711
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 79712, 79713
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
'
'        Case 79714
'        'Sem dados
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Private Sub Move_NFiscal_Obj(objNFiscal As ClassNFiscal, tNFiscal As typeNFiscal)
'
'    With tNFiscal
'
'        objNFiscal.lNumIntDoc = .lNumIntDoc
'        objNFiscal.iStatus = .iStatus
'        objNFiscal.iFilialEmpresa = .iFilialEmpresa
'        objNFiscal.iFilialEntrega = .iFilialEntrega
'        objNFiscal.dtDataVencimento = .dtDataVencimento
'        objNFiscal.dtDataReferencia = .dtDataReferencia
'        objNFiscal.iFilialPedido = .iFilialPedido
'        objNFiscal.lNumPedidoVenda = .lNumPedidoVenda
'        objNFiscal.sNumPedidoTerc = .sNumPedidoTerc
'        objNFiscal.iClasseDocCPR = .iClasseDocCPR
'        objNFiscal.lNumIntDocCPR = .lNumIntDocCPR
'        objNFiscal.dValorTotal = .dValorTotal
'        objNFiscal.dValorProdutos = .dValorProdutos
'        objNFiscal.dValorFrete = .dValorFrete
'        objNFiscal.dValorOutrasDespesas = .dValorOutrasDespesas
'        objNFiscal.dValorDesconto = .dValorDesconto
'        objNFiscal.iCodTransportadora = .iCodTransportadora
'        objNFiscal.sMensagemNota = .sMensagemNota
'        objNFiscal.iTabelaPreco = .iTabelaPreco
'        objNFiscal.sNaturezaOp = .sNaturezaOp
'        objNFiscal.dPesoLiq = .dPesoLiq
'        objNFiscal.dPesoBruto = .dPesoBruto
'        objNFiscal.lNumIntTrib = .lNumIntTrib
'        objNFiscal.sPlaca = .sPlaca
'        objNFiscal.sPlacaUF = .sPlacaUF
'        objNFiscal.lVolumeQuant = .lVolumeQuant
'        objNFiscal.sVolumeEspecie = .sVolumeEspecie
'        objNFiscal.sVolumeMarca = .sVolumeMarca
'        objNFiscal.sVolumeNumero = .sVolumeNumero
'        objNFiscal.iCanal = .iCanal
'        objNFiscal.lNumIntNotaOriginal = .lNumIntNotaOriginal
'        objNFiscal.lClienteBenef = .lClienteBenef
'        objNFiscal.iFilialCliBenef = .iFilialCliBenef
'        objNFiscal.lFornecedorBenef = .lFornecedorBenef
'        objNFiscal.iFilialFornBenef = .iFilialFornBenef
'        objNFiscal.iFreteRespons = .iFreteRespons
'        objNFiscal.lNumRecebimento = .lNumRecebimento
'        objNFiscal.sObservacao = .sObservacao
'        objNFiscal.sCodUsuarioCancel = .sCodUsuarioCancel
'        objNFiscal.sMotivoCancel = .sMotivoCancel
'
'    End With
'
'End Sub
'
'Public Function NFiscal_Consulta_TituloPagar(objNFiscal As ClassNFiscal, objTipoDocInfo As ClassTipoDocInfo) As Long
'
'Dim objTituloPagar As New ClassTituloPagar
'
'On Error GoTo Erro_NFiscal_Consulta_TituloPagar
'
'    With objTituloPagar
'
'        .lNumTitulo = objNFiscal.lNumNotaFiscal
'        .lFornecedor = objNFiscal.lFornecedor
'        .iFilial = objNFiscal.iFilialForn
'        .dtDataEmissao = objNFiscal.dtDataEmissao
'        .sSiglaDocumento = objTipoDocInfo.sTipoDocCPR
'
'    End With
'
'    Call Chama_Tela("TituloPagar_Consulta", objTituloPagar)
'
'    NFiscal_Consulta_TituloPagar = SUCESSO
'
'    Exit Function
'
'Erro_NFiscal_Consulta_TituloPagar:
'
'    NFiscal_Consulta_TituloPagar = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function NFiscal_Consulta_TituloReceber(objNFiscal As ClassNFiscal) As Long
'
'Dim objTituloReceber As New ClassTituloReceber
'
'On Error GoTo Erro_NFiscal_Consulta_TituloReceber
'
'    objTituloReceber.lNumIntDoc = objNFiscal.lNumIntDocCPR
'
'    Call Chama_Tela("TituloReceber_Consulta", objTituloReceber)
'
'    NFiscal_Consulta_TituloReceber = SUCESSO
'
'    Exit Function
'
'Erro_NFiscal_Consulta_TituloReceber:
'
'    NFiscal_Consulta_TituloReceber = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function NFiscal_Consulta_NFPag(objNFiscal As ClassNFiscal) As Long
'
'Dim objNFsPag As New ClassNFsPag
'
'On Error GoTo Erro_NFiscal_Consulta_NFPag
'
'    With objNFsPag
'
'        .lNumNotaFiscal = objNFiscal.lNumNotaFiscal
'        .lFornecedor = objNFiscal.lFornecedor
'        .iFilial = objNFiscal.iFilialForn
'        .iFilialEmpresa = objNFiscal.iFilialEmpresa
'        .dtDataEmissao = objNFiscal.dtDataEmissao
'
'    End With
'
'    Call Chama_Tela("NFPag_Consulta", objNFsPag)
'
'    NFiscal_Consulta_NFPag = SUCESSO
'
'    Exit Function
'
'Erro_NFiscal_Consulta_NFPag:
'
'    NFiscal_Consulta_NFPag = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function NFiscal_Consulta_DebitosReceb(objNFiscal As ClassNFiscal) As Long
'
'Dim objDebitoReceber As New ClassDebitoRecCli
'
'On Error GoTo Erro_NFiscal_Consulta_DebitosReceb
'
'    objDebitoReceber.lNumIntDoc = objNFiscal.lNumIntDocCPR
'
'    Call Chama_Tela("DebitosReceb", objDebitoReceber)
'
'    NFiscal_Consulta_DebitosReceb = SUCESSO
'
'    Exit Function
'
'Erro_NFiscal_Consulta_DebitosReceb:
'
'    NFiscal_Consulta_DebitosReceb = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Function NFiscal_Consulta_CreditoPagar(objNFiscal As ClassNFiscal) As Long
'
'Dim objCreditoPagar As New ClassCreditoPagar
'
'On Error GoTo Erro_NFiscal_Consulta_CreditoPagar
'
'    objCreditoPagar.lNumIntDoc = objNFiscal.lNumIntDocCPR
'
'    Call Chama_Tela("CreditoPagar", objCreditoPagar)
'
'    NFiscal_Consulta_CreditoPagar = SUCESSO
'
'    Exit Function
'
'Erro_NFiscal_Consulta_CreditoPagar:
'
'    NFiscal_Consulta_CreditoPagar = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
''*** Fim do trecho que deverá ser apagado ***
'
''jones-15/03/01
'Public Sub INSSBase_Change()
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.INSSBase_Change
'End Sub
'
'Public Sub INSSBase_Validate(Cancel As Boolean)
'    Call gobjTribTab.INSSBase_Validate(Cancel)
'End Sub
'
'Public Sub INSSDeducoes_Change()
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.INSSDeducoes_Change
'End Sub
'
'Public Sub INSSDeducoes_Validate(Cancel As Boolean)
'    Call gobjTribTab.INSSDeducoes_Validate(Cancel)
'End Sub
'
'Public Sub INSSRetido_Click()
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.INSSRetido_Click
'End Sub
'
'Public Sub INSSValor_Change()
'    iAlterado = REGISTRO_ALTERADO
'    Call gobjTribTab.INSSValor_Change
'End Sub
'
'Public Sub INSSValor_Validate(Cancel As Boolean)
'    Call gobjTribTab.INSSValor_Validate(Cancel)
'End Sub
''fim jones-15/03/01
'
''Leo daqui p/ baixo
'Function Inicializa_Grid_ComprovServ(objGridInt As AdmGrid) As Long
'
''Inicializa o Grid de ComprovServ
'
'    'Form do Grid
'    Set objGridInt.objForm = Me
'
'    'Títulos das colunas
'    objGridInt.colColuna.Add (" ")
'    objGridInt.colColuna.Add ("Seleciona")
'    objGridInt.colColuna.Add ("Número")
'    objGridInt.colColuna.Add ("Data")
'    objGridInt.colColuna.Add ("Valor Mercadoria")
'    'objGridInt.colColuna.Add ("Frete Peso")
'    objGridInt.colColuna.Add ("Ad Valoren")
'    objGridInt.colColuna.Add ("Sábado")
'    objGridInt.colColuna.Add ("Valor Container")
'    objGridInt.colColuna.Add ("Despacho")
'    objGridInt.colColuna.Add ("Manuseio")
'    objGridInt.colColuna.Add ("Sub Total")
'    objGridInt.colColuna.Add ("ICMS")
'    objGridInt.colColuna.Add ("Pedagio")
'    objGridInt.colColuna.Add ("Total")
'
'
'    'Controles que participam do Grid
'    objGridInt.colCampo.Add (SelecionaCon.Name)
'    objGridInt.colCampo.Add (ComprovServCon.Name)
'    objGridInt.colCampo.Add (DataCon.Name)
'    objGridInt.colCampo.Add (ValorMercadCon.Name)
'    'objGridInt.colCampo.Add (FretePesoCon.Name)
'    objGridInt.colCampo.Add (AdValorenCon.Name)
'    objGridInt.colCampo.Add (SabadoCon.Name)
'    objGridInt.colCampo.Add (ValorContainerCon.Name)
'    objGridInt.colCampo.Add (DespachoCon.Name)
'    objGridInt.colCampo.Add (Manuseio.Name)
'    objGridInt.colCampo.Add (SubTotalCon.Name)
'    objGridInt.colCampo.Add (ISSCon.Name)
'    objGridInt.colCampo.Add (PedagioCon.Name)
'    objGridInt.colCampo.Add (TotalCon.Name)
'
'    'Colunas do Grid
'    iGrid_SelecionaCon_col = 1
'    iGrid_ComprovServCon_col = 2
'    iGrid_DataCon_Col = 3
'    iGrid_ValorMercadoriaCon_col = 4
'    'iGrid_FretePesoCon_col
'    iGrid_AdValorCon_col = 5
'    iGrid_SabadoCon_col = 6
'    iGrid_ValorContainerCon_col = 7
'    iGrid_DespachoCon_Col = 8
'    iGrid_Manuseio_Col = 9
'    iGrid_SubTotalCon_col = 10
'    iGrid_ISSCon_col = 11
'    iGrid_PedagioCon_Col = 12
'    iGrid_TotalCon_col = 13
'
'    'Grid do GridInterno
'    objGridInt.objGrid = GridComprovServ
'
'    'Todas as linhas do grid
'    objGridInt.objGrid.Rows = NUM_MAXIMO_CONHECFRETE
'
'    'Linhas visíveis do grid
'    objGridInt.iLinhasVisiveis = 12
'
'    'Largura da primeira coluna
'    GridComprovServ.ColWidth(0) = 300
'
'    'Largura automática para as outras colunas
'    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
'
'    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
'
'    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
'
'    'Chama função que inicializa o Grid
'    Call Grid_Inicializa(objGridInt)
'
'    Inicializa_Grid_ComprovServ = SUCESSO
'
'    Exit Function
'
'End Function
'
'
'Sub BotoesNFFatura_Click(Index As Integer)
'
''Leo
'    Select Case Index
'
'        Case BOTAO_DATAREFERENCIAUP
'            Call BotaoDataReferenciaUp_Click
'
'        Case BOTAO_DATAREFERENCIADOWN
'            Call BotaoDataReferenciaDown_Click
'
'        Case BOTAO_LIMPARNF
'            Call BotaoLimparNF_Click
'
'        Case BOTAO_FECHAR
'            Call BotaoFechar_Click
'
'        Case BOTAO_LIMPAR
'            Call BotaoLimpar_Click
'
'        Case BOTAO_EXCLUIR
'            Call BotaoExcluir_Click
'
'        Case BOTAO_GRAVAR
'            Call BotaoGravar_Click
'
'        Case BOTAO_CONSULTATITREC
'            Call BotaoConsultaTitRec_Click
'
'        Case BOTAO_PRODUTOS
'            Call BotaoProdutos_Click
'
'        Case BOTAO_ESTOQUE_PRODUTOS
'            Call BotaoEstoqueProd_Click
'
'        Case BOTAO_LOCALIZACAO_PRODUTOS
'            Call BotaoLocalizacao_Click
'
'        Case BOTAO_LOTES
'            Call BotaoLotes_Click
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub GridComprovServ_Click()
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Click(objGridComprovServ, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridComprovServ, iAlterado)
'    End If
'
'End Sub
'
'Public Sub GridComprovServ_EnterCell()
'
'    Call Grid_Entrada_Celula(objGridComprovServ, iAlterado)
'
'End Sub
'
'Public Sub GridComprovServ_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    Call Grid_Trata_Tecla1(KeyCode, objGridComprovServ)
'
'End Sub
'
'Public Sub GridComprovServ_GotFocus()
'
'    Call Grid_Recebe_Foco(objGridComprovServ)
'
'End Sub
'
'Public Sub GridComprovServ_KeyPress(KeyAscii As Integer)
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Trata_Tecla(KeyAscii, objGridComprovServ, iExecutaEntradaCelula)
'
'   If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridComprovServ, iAlterado)
'    End If
'
'End Sub
'
'Public Sub GridComprovServ_LeaveCell()
'
'    Call Saida_Celula(objGridComprovServ)
'
'End Sub
'
'
'Public Sub GridComprovServ_Validate(Cancel As Boolean)
'
'    Call Grid_Libera_Foco(objGridComprovServ)
'
'End Sub
'
'Public Sub GridComprovServ_RowColChange()
'
'    Call Grid_RowColChange(objGridComprovServ)
'
'End Sub
'
'Public Sub GridComprovServ_Scroll()
'
'    Call Grid_Scroll(objGridComprovServ)
'
'End Sub
'
'Public Sub SelecionaCon_Click()
'
'Dim lErro As Long
'
'    iAlterado = REGISTRO_ALTERADO
'
'    Call Soma_Coluna
'
'    Exit Sub
'
'End Sub
'
'Public Sub SelecionaCon_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridComprovServ)
'
'End Sub
'
'Public Sub SelecionaCon_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComprovServ)
'
'End Sub
'
'Public Sub SelecionaCon_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridComprovServ.objControle = SelecionaCon
'    lErro = Grid_Campo_Libera_Foco(objGridComprovServ)
'    If lErro <> SUCESSO Then Cancel = True
'
'
'End Sub
'
''Faz a soma das colunas do grid e atualiza os
''valores dos campos na tela
'Sub Soma_Coluna()
'
'Dim iIndice As Integer
'Dim dValorMerc As Double
'Dim dValorContainer As Double
'Dim bValidateICMS As Boolean
'Dim dValorTotal As Double
'Dim lErro As Long
'
'On Error GoTo Erro_Soma_Coluna
'
'    iIndice = 0
'
'
'
'    'Para o nº de colunas existentes do grid
'    For iIndice = 1 To objGridComprovServ.iLinhasExistentes
'
'        'Se a linha estiver selecionada
'        If StrParaInt(GridComprovServ.TextMatrix(iIndice, iGrid_SelecionaCon_col)) = MARCADO Then
'
'            'Acumula nas variáveis, a soma dos valores selecionados
'            'em cada coluna
'            dValorContainer = dValorContainer + StrParaDbl(GridComprovServ.TextMatrix(iIndice, iGrid_ValorContainerCon_col))
'            dValorTotal = dValorTotal + StrParaDbl(GridComprovServ.TextMatrix(iIndice, iGrid_TotalCon_col))
'            dValorMerc = dValorMerc + StrParaDbl(GridComprovServ.TextMatrix(iIndice, iGrid_ValorMercadoriaCon_col))
'
'        End If
'
'    Next
'
'    TotalComprovServ.Caption = Format(dValorTotal, "STANDARD")
'
'    ValorMercadoria.Text = Format(dValorMerc, "STANDARD")
'
'    ValorContainer.Text = Format(dValorContainer, "STANDARD")
'
'    bValidateICMS = True
'
'    lErro = ValorTotal_Calcula()
'    If lErro <> SUCESSO Then gError 94101
'
'    Exit Sub
'
'Erro_Soma_Coluna:
'
'    Select Case gErr
'
'        Case 94101
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Function Preenche_GridComprovServ(objNFiscal As ClassNFiscal) As Long
''Preenche o Grid com as Parcelas da Nota Fiscal
'
'Dim dSubTotal As Double
'Dim objComprovServ As ClassComprovServ
'Dim iIndice As Integer
'Dim lErro As Long
'Dim objCodigoNome As New ClassCliente
'Dim objCotacaoMoeda As New ClassCotacaoMoeda
'Dim objEquipamento As New ClassEquipamentos
'Dim objItemComprov As ClassItemComprovServ
'
'On Error GoTo Erro_Preenche_GridComprovServ
'
'    'Limpa o Grid
'    Call Grid_Limpa(objGridComprovServ)
'
'    If objNFiscal.lNumIntDoc <> 0 Then
'
'        'Preenche o Grid Conheciento de Frete com Comprovantes de
'        'Servicos Desvinculados
'        lErro = ComprovServ_Le_Vinculados(objNFiscal)
'        If lErro <> SUCESSO Then gError 94082
'    End If
'
'    If Len(Trim(NFiscal.Caption)) = 0 Then
'
'        objCodigoNome.sNomeReduzido = Cliente.Text
'
'        lErro = Cliente_Le_NomeReduzido(objCodigoNome)
'        If lErro <> SUCESSO And lErro <> 12348 Then gError 94083
'        If lErro = 12348 Then gError 94089
'
'        objNFiscal.lCliente = objCodigoNome.lCodigo
'        objNFiscal.iFilialCli = Codigo_Extrai(Filial)
'
'    End If
'
'    'Preenche o Grid Conheciento de Frete com os Comprovante de
'    'Servicos Vinculados a NF (se for uma NF existente)
'    lErro = ComprovServ_Le_Desvinculados(objNFiscal, TIPO_NFISCAL)
'    If lErro <> SUCESSO Then gError 94084
'
''    If Len(Trim(DataEmissao.ClipText)) > 0 Then
''
''        objCotacaoMoeda.dtData = StrParaDate(DataEmissao.Text)
''        objCotacaoMoeda.iMoeda = MOEDA_DOLAR
''
''        lErro = CotacaoMoeda_Le(objCotacaoMoeda)
''        If lErro <> SUCESSO And lErro <> 80267 Then gError 94085
''        If lErro = 80267 Then
''
''            lErro = Rotina_Aviso(vbOKOnly, "ERRO_COTACAOMOEDA_DOLAR_NAO_ENCONTRADA", objCotacaoMoeda.dtData)
''
''        End If
'
''    End If
'
'    iIndice = 0
'
'    If objNFiscal.colComprovServ.Count > objGridComprovServ.objGrid.Rows - 1 Then
'        objGridComprovServ.objGrid.Rows = objNFiscal.colComprovServ.Count + 2
'        Call Grid_Inicializa(objGridComprovServ)
'    End If
'
'    'Preenche o Grid
'    For Each objComprovServ In objNFiscal.colComprovServ
'
'        iIndice = iIndice + 1
'
'        lErro = ItensComprovServ_Le(objComprovServ)
'        If lErro <> SUCESSO Then gError 94086
'
'
'        GridComprovServ.TextMatrix(iIndice, iGrid_ValorMercadoriaCon_col) = Format(objComprovServ.dValorMerc, "Standard")
'       ' GridComprovServ.TextMatrix(iIndice, iGrid_FretePesoCon_col) = Format(objComprovServ.dFrete, "Standard")
'        GridComprovServ.TextMatrix(iIndice, iGrid_AdValorCon_col) = Format(objComprovServ.dAdValoren, "Standard")
'        GridComprovServ.TextMatrix(iIndice, iGrid_SabadoCon_col) = Format(objComprovServ.dSabado, "Standard")
'        GridComprovServ.TextMatrix(iIndice, iGrid_ValorContainerCon_col) = Format(objComprovServ.dValorContainer, "Standard")
'        GridComprovServ.TextMatrix(iIndice, iGrid_ComprovServCon_col) = objComprovServ.lNumero
'        GridComprovServ.TextMatrix(iIndice, iGrid_ISSCon_col) = Format(objComprovServ.dISSICMS, "Standard")
'        GridComprovServ.TextMatrix(iIndice, iGrid_TotalCon_col) = Format(objComprovServ.dTotal, "Standard")
'        GridComprovServ.TextMatrix(iIndice, iGrid_PedagioCon_Col) = Format(objComprovServ.dPedagio, "Standard")
'        GridComprovServ.TextMatrix(iIndice, iGrid_DespachoCon_Col) = Format(objComprovServ.dDespacho, "Standard")
'        GridComprovServ.TextMatrix(iIndice, iGrid_Manuseio_Col) = Format(objComprovServ.dManuseio, "Standard")
'
'        If objComprovServ.dtData <> DATA_NULA Then GridComprovServ.TextMatrix(iIndice, iGrid_DataCon_Col) = Format(objComprovServ.dtData, "dd/mm/yyyy")
'
'        If objComprovServ.lNumIntNota = 0 Then
'            GridComprovServ.TextMatrix(iIndice, iGrid_SelecionaCon_col) = DESMARCADO
'        Else
'            GridComprovServ.TextMatrix(iIndice, iGrid_SelecionaCon_col) = MARCADO
'        End If
'
'        dSubTotal = objComprovServ.dTotal - objComprovServ.dISSICMS - objComprovServ.dPedagio
'        GridComprovServ.TextMatrix(iIndice, iGrid_SubTotalCon_col) = Format(dSubTotal, "Standard")
'
'    Next
'
'    objGridComprovServ.iLinhasExistentes = iIndice
'
'    Call Grid_Refresh_Checkbox(objGridComprovServ)
'
'    Preenche_GridComprovServ = SUCESSO
'
'    Exit Function
'
'Erro_Preenche_GridComprovServ:
'
'    Preenche_GridComprovServ = gErr
'
'    Select Case gErr
'
''       Case 81183
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_EQUIPAMENTO_NAO_EXISTENTE", gErr, objEquipamento.sSigla)
'
'        Case 94089
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCodigoNome.sNomeReduzido)
'
'        Case 94082, 94084, 94083, 94085, 94087, 94086
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'   End Select
'
'    Exit Function
'
'End Function
'
''****JÁ*EXISTE*EM**COMPROVANTESERVICO***
'
'Public Function ComprovServ_Le_Desvinculados(objNFiscal As ClassNFiscal, iTipoComprovante As Integer) As Long
''Lê na Tabela ComprovantesServicos, registros desvinculados em uma
''filial cliente de uma filial fornecedor
'
'Dim lErro As Long
'Dim objComprovantesServicos As ClassComprovServ
'Dim tComprovServ As typeComprovServ
'Dim lComando As Long
'
'On Error GoTo Erro_ComprovServ_Le_Desvinculados
'
'    'Abre o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 84747
'
'    'Seleciona em ComprovantesServicos, registros cujo Cliente e
'    'filialCliente iguais ao passado em objNfiscal, FilialEmpresa
'    '= giFilialEmpresa e NumIntNota = 0(não vinculado)
'    lErro = Comando_Executar(lComando, "SELECT Frete, Manuseio, Desconto, ValorMerc, PesoBruto, AdValoren, Sabado, Numero, ISSICMS, Total, ValorContainer, Pedagio, Despacho, Data, NumIntNota FROM ComprovantesServicos WHERE Cliente = ? AND FilialCliente = ? AND FilialEmpresa = ? AND NumIntNota = 0 AND TipoComprovante = ?", tComprovServ.dFrete, tComprovServ.dManuseio, tComprovServ.dDesconto, tComprovServ.dValorMerc, tComprovServ.dPesoBruto, tComprovServ.dAdValoren, tComprovServ.dSabado, tComprovServ.lNumero, tComprovServ.dISSICMS, tComprovServ.dTotal, tComprovServ.dValorContainer, tComprovServ.dPedagio, tComprovServ.dDespacho, tComprovServ.dtData, tComprovServ.lNumIntNota, objNFiscal.lCliente, objNFiscal.iFilialCli, giFilialEmpresa, iTipoComprovante)
'    If lErro <> AD_SQL_SUCESSO Then gError 84748
'
'    'busca o 1º
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 84749
'
'    'Enquanto encontrar compovantes, os carrega em colComprovServ
'    Do While lErro = AD_SQL_SUCESSO
'
'        Set objComprovantesServicos = New ClassComprovServ
'
'        objComprovantesServicos.dFrete = tComprovServ.dFrete
'        objComprovantesServicos.dtData = tComprovServ.dtData
'        objComprovantesServicos.lNumIntNota = tComprovServ.lNumIntNota
'        objComprovantesServicos.dValorMerc = tComprovServ.dValorMerc
'        objComprovantesServicos.dPesoBruto = tComprovServ.dPesoBruto
'        objComprovantesServicos.dAdValoren = tComprovServ.dAdValoren
'        objComprovantesServicos.dSabado = tComprovServ.dSabado
'        objComprovantesServicos.lNumero = tComprovServ.lNumero
'        objComprovantesServicos.dISSICMS = tComprovServ.dISSICMS
'        objComprovantesServicos.dTotal = tComprovServ.dTotal
'        objComprovantesServicos.dValorContainer = tComprovServ.dValorContainer
'        objComprovantesServicos.dPedagio = tComprovServ.dPedagio
'        objComprovantesServicos.dDespacho = tComprovServ.dDespacho
'        objComprovantesServicos.dDesconto = tComprovServ.dDesconto
'        objComprovantesServicos.dManuseio = tComprovServ.dManuseio
'
'        objNFiscal.colComprovServ.Add objComprovantesServicos
'
'        'tenta buscar um próximo Comprovante
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SEM_DADOS And AD_SQL_SUCESSO Then gError 84750
'
'    Loop
'
'    ' Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    ComprovServ_Le_Desvinculados = SUCESSO
'
'    Exit Function
'
'Erro_ComprovServ_Le_Desvinculados:
'
'    ComprovServ_Le_Desvinculados = gErr
'
'    Select Case gErr
'
'
'        Case 84747
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 84748, 84749, 84750
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPROVANTESSERVICOS", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
''****JÁ*EXISTE*EM**COMPROVANTESERVICO***
'Public Function ComprovServ_Le_Vinculados(objNFiscal As ClassNFiscal) As Long
''Lê os Comprovantes vinculados a nota
'
'Dim lErro As Long
'Dim lComando As Long
'Dim tComprovServ As typeComprovServ
'Dim objComprovantesServicos As ClassComprovServ
'
'On Error GoTo Erro_ComprovServ_Le_Vinculados
'
'    'Abre comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 84743
'
'    'Seleciona Comprovantes Vinculados a nota
'    lErro = Comando_Executar(lComando, "SELECT Frete, Manuseio, Desconto, ValorMerc, PesoBruto, AdValoren, Sabado, Numero, ISSICMS, Total, ValorContainer, Pedagio, Despacho, Data FROM ComprovantesServicos WHERE Cliente = ? AND FilialCliente = ? AND FilialEmpresa = ? AND NumIntNota = ? ", tComprovServ.dFrete, tComprovServ.dManuseio, tComprovServ.dDesconto, tComprovServ.dValorMerc, tComprovServ.dPesoBruto, tComprovServ.dAdValoren, tComprovServ.dSabado, tComprovServ.lNumero, tComprovServ.dISSICMS, tComprovServ.dTotal, tComprovServ.dValorContainer, tComprovServ.dPedagio, tComprovServ.dDespacho, tComprovServ.dtData, objNFiscal.lCliente, objNFiscal.iFilialCli, giFilialEmpresa, objNFiscal.lNumIntDoc)
'    If lErro <> AD_SQL_SUCESSO Then gError 84744
'
'    'Tenta ler a primeira Parcela
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 84745
'
'    'Enquanto encontrar compovantes, os carrega em colComprovServ
'    Do While lErro = AD_SQL_SUCESSO
'
'        Set objComprovantesServicos = New ClassComprovServ
'
'        objComprovantesServicos.dFrete = tComprovServ.dFrete
'        objComprovantesServicos.dtData = tComprovServ.dtData
'        objComprovantesServicos.lNumIntNota = objNFiscal.lNumIntDoc
'        objComprovantesServicos.dValorMerc = tComprovServ.dValorMerc
'        objComprovantesServicos.dPesoBruto = tComprovServ.dPesoBruto
'        objComprovantesServicos.dAdValoren = tComprovServ.dAdValoren
'        objComprovantesServicos.dSabado = tComprovServ.dSabado
'        objComprovantesServicos.lNumero = tComprovServ.lNumero
'        objComprovantesServicos.dISSICMS = tComprovServ.dISSICMS
'        objComprovantesServicos.dTotal = tComprovServ.dTotal
'        objComprovantesServicos.dValorContainer = tComprovServ.dValorContainer
'        objComprovantesServicos.dPedagio = tComprovServ.dPedagio
'        objComprovantesServicos.dDespacho = tComprovServ.dDespacho
'        objComprovantesServicos.dManuseio = tComprovServ.dManuseio
'        objComprovantesServicos.dDesconto = tComprovServ.dDesconto
'
'        objNFiscal.colComprovServ.Add objComprovantesServicos
'
'        'tenta encontrar um próximo Comprovante
'        lErro = Comando_BuscarProximo(lComando)
'        If lErro <> AD_SQL_SEM_DADOS And AD_SQL_SUCESSO Then gError 84746
'
'    Loop
'
'    'Fecha comando
'    Call Comando_Fechar(lComando)
'
'    ComprovServ_Le_Vinculados = SUCESSO
'
'    Exit Function
'
'Erro_ComprovServ_Le_Vinculados:
'
'    ComprovServ_Le_Vinculados = gErr
'
'    Select Case gErr
'
'       Case 84743
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 84744, 84745, 84746
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPROVANTESSERVICOS", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
''Move os dados de gridComprovServ para a memória
'Sub Move_GridComprovServ_Memoria(objNFiscal As ClassNFiscal)
'
'Dim iIndice As Integer
'Dim objComprovServ As ClassComprovServ
'
'    'Se não há comprovantes a recolher, sai da função
'    If objGridComprovServ.iLinhasExistentes = 0 Then Exit Sub
'
'    'Para cada comprovante do grid
'    For iIndice = 1 To objGridComprovServ.iLinhasExistentes
'
'        Set objComprovServ = New ClassComprovServ
'
'        If StrParaInt(GridComprovServ.TextMatrix(iIndice, iGrid_SelecionaCon_col)) = MARCADO Then
'
'            'recolhe os dados do grid de comprovantes de serviços
'            objComprovServ.lNumero = StrParaLong(GridComprovServ.TextMatrix(iIndice, iGrid_ComprovServCon_col))
'
'            objComprovServ.dtData = StrParaDate(GridComprovServ.TextMatrix(iIndice, iGrid_DataCon_Col))
'
'
'            objComprovServ.iFilialEmpresa = giFilialEmpresa
'
'            'Carrega o obj na coleção
'            objNFiscal.colComprovServ.Add objComprovServ
'
'        End If
'
'    Next
'
'    Exit Sub
'
'End Sub
'
''****JÁ*EXISTE*EM**COMPROVANTESERVICO***
'
''Altera os vínculos dos comprovantes de serviços
'Public Function ComprovServ_Vincula(objNFiscal As ClassNFiscal) As Long
'
'Dim lErro As Long
'Dim lComando As Long
'Dim lComando1 As Long
'Dim objComprovServ As ClassComprovServ
'Dim lNumIntNota As Long
'
'On Error GoTo Erro_ComprovServ_Vincula
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 84738
'
'    lComando1 = Comando_Abrir()
'    If lComando1 = 0 Then gError 84739
'
'    For Each objComprovServ In objNFiscal.colComprovServ
'
'        objComprovServ.lNumIntNota = objNFiscal.lNumIntDoc
'
'        'Seleciona em comprovantes de serviços o NumIntNota cujo
'        'filial empresa, Numero e Data sejam iguais aos recebidos no objNFiscal
'        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntNota FROM ComprovantesServicos WHERE FilialEmpresa = ? AND Numero = ? AND Data = ?", 0, lNumIntNota, objComprovServ.iFilialEmpresa, objComprovServ.lNumero, objComprovServ.dtData)
'        If lErro <> AD_SQL_SUCESSO Then gError 84740
'
'        'Busca o 1º
'        lErro = Comando_BuscarPrimeiro(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 84741
'
'        If lErro = AD_SQL_SUCESSO Then
'
'            If lNumIntNota <> DESMARCADO And lNumIntNota <> objNFiscal.lNumIntDoc Then gError 84756
'
'            'Atualiza o Campo NumIntNota na Tabela ComprovantesServicos
'            lErro = Comando_ExecutarPos(lComando1, "UPDATE ComprovantesServicos SET NumIntNota = ? ", lComando, objComprovServ.lNumIntNota)
'            If lErro <> AD_SQL_SUCESSO Then gError 84742
'
'4        End If
'
'    Next
'
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando1)
'
'    ComprovServ_Vincula = SUCESSO
'
'    Exit Function
'
'Erro_ComprovServ_Vincula:
'
'    ComprovServ_Vincula = gErr
'
'    Select Case gErr
'
'        Case 84756
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMPROVSERV_VINCULADO", gErr, objComprovServ.lNumero)
'
'        Case 84738, 84739
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 84740, 84741
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMPROVANTESSERVICOS", gErr)
'
'        Case 84742
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COMPROVANTESSERVICOS", gErr)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'
'    Call Comando_Fechar(lComando)
'    Call Comando_Fechar(lComando1)
'
'    Exit Function
'
'End Function
'
'
'
''******************JÁ EXISTE NA TELA DE COTACAOMOEDA *****************
'Public Function CotacaoMoeda_Le(objCotacaoMoeda As ClassCotacaoMoeda) As Long
'
'Dim lComando As Long
'Dim lErro As Long
'Dim tCotacaoMoeda As typeCotacaoMoeda
'
'On Error GoTo Erro_CotacaoMoeda_Le
'
'    'Abre Comandos
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 80264
'
'    'Verifica se existe cotação para a data informada.
'    lErro = Comando_Executar(lComando, "SELECT Valor FROM CotacoesMoeda WHERE Data = ? AND Moeda = ?", tCotacaoMoeda.dValor, objCotacaoMoeda.dtData, objCotacaoMoeda.iMoeda)
'    If lErro <> AD_SQL_SUCESSO Then gError 80265
'
'    '===> Caso exista seleciona
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 80266
'
'    '===> Se não exite Erro
'    If lErro = AD_SQL_SEM_DADOS Then gError 80267
'
'    'Carrega objCotacaoMoeda
'    objCotacaoMoeda.dValor = tCotacaoMoeda.dValor
'
'    'Fecha Comando
'    Call Comando_Fechar(lComando)
'
'    CotacaoMoeda_Le = SUCESSO
'
'Exit Function
'
'Erro_CotacaoMoeda_Le:
'
'    CotacaoMoeda_Le = gErr
'
'    Select Case gErr
'
'        Case 80264
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 80265, 80266
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COTACOESMOEDA", gErr, objCotacaoMoeda.dtData, objCotacaoMoeda.iMoeda)
'
'        Case 80267 'CotacaoMoeda não cadastrada
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
''''''????? FERNANDO
''Função já existe, deve ser removida.
'Public Function Equipamento_Le(objEquipamento As ClassEquipamentos) As Long
''Lê os dados de Equipamento cujo código foi passado.
''Devolve os dados em objEquipamento.
'
'Dim lErro As Long
'Dim lComando As Long
'Dim tEquipamento As typeEquipamentos
'
'On Error GoTo Erro_Equipamento_Le
'
'    'Abre o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 81180
'
'    'Inicializa as strings que receberão dados do BD
'    tEquipamento.sDescricao = String(STRING_EQUIPAMENTO_DESCRICA0, 0)
'
'    'Busca o Equipamento no BD através da  sigla passada como parâmetro
'    lErro = Comando_Executar(lComando, "SELECT Descricao, TaxaMinima, ValorDolar FROM Equipamentos WHERE Sigla = ? ", tEquipamento.sDescricao, tEquipamento.dTaxaMinima, tEquipamento.dValorDolar, objEquipamento.sSigla)
'    If lErro <> AD_SQL_SUCESSO Then gError 81181
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81182
'
'    'Se não encontrou o Equipamento => erro
'    If lErro = AD_SQL_SEM_DADOS Then gError 81183
'
'     'Transfere para objEquipamentos os dados obtidos do BD
'    objEquipamento.sDescricao = tEquipamento.sDescricao
'    objEquipamento.dTaxaMinima = tEquipamento.dTaxaMinima
'    objEquipamento.dValorDolar = tEquipamento.dValorDolar
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    Equipamento_Le = SUCESSO
'
'    Exit Function
'
'Erro_Equipamento_Le:
'
'    Equipamento_Le = gErr
'
'    Select Case gErr
'
'        Case 81180
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 81181, 81182
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EQUIPAMENTO", gErr, objEquipamento.sSigla)
'
'        Case 81183
'            'Equipamento não encontrado. Esse erro será tratado na rotina chamadora
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
''??????? Essa função já existe na tela de comprovante de serviço
'Public Function ItensComprovServ_Le(objComprovServ As ClassComprovServ) As Long
'
'Dim lErro As Long
'Dim lComando As Long
'Dim objItemCompServ As ClassItemComprovServ
'Dim tItemComprovServ As typeItemComprovSer
'
'On Error GoTo Erro_ItensComprovServ_Le
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError 86195
'
'    With tItemComprovServ
'
'        .sEquipamento = String(STRING_EQUIPAMENTO_SIGLA, 0)
'
'        lErro = Comando_Executar(lComando, "SELECT Equipamento, Quantidade, Valor, QuantAdicional, ValorAdicional, ValorTotal FROM ItensComprovanteServ WHERE FilialEmpresa = ? AND Numero = ? AND Data = ?", .sEquipamento, .dQuantidade, .dValor, .dQuantAdicional, .dValorAdicional, .dValorTotal, objComprovServ.iFilialEmpresa, objComprovServ.lNumero, objComprovServ.dtData)
'        If lErro <> AD_SQL_SUCESSO Then gError 86196
'
'        lErro = Comando_BuscarPrimeiro(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86197
'
'        Do While lErro = AD_SQL_SUCESSO
'
'            Set objItemCompServ = New ClassItemComprovServ
'
'            objItemCompServ.sEquipamento = .sEquipamento
'            objItemCompServ.dQuantAdicional = .dQuantAdicional
'            objItemCompServ.dQuantidade = .dQuantidade
'            objItemCompServ.dValor = .dValor
'            objItemCompServ.dValorAdicional = .dValorAdicional
'            objItemCompServ.dValorTotal = .dValorTotal
'
'            objComprovServ.colItens.Add objItemCompServ
'
'            lErro = Comando_BuscarProximo(lComando)
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86198
'
'        Loop
'
'    End With
'
'    Call Comando_Fechar(lComando)
'
'    ItensComprovServ_Le = SUCESSO
'
'    Exit Function
'
'Erro_ItensComprovServ_Le:
'
'    ItensComprovServ_Le = gErr
'
'    Select Case gErr
'
'        Case 86195
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 86196, 86198
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITEMCOMPROVSERV", gErr)
'
'        Case 86197
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function
'
'Public Property Get GridComprovServ() As Object
'     Set GridComprovServ = objUserControl.Controls("GridComprovServ")
'End Property
'
'Public Property Get SelecionaCon() As Object
'     Set SelecionaCon = objUserControl.Controls("SelecionaCon")
'End Property
'
'Public Property Get ComprovServCon() As Object
'     Set ComprovServCon = objUserControl.Controls("ComprovServCon")
'End Property
'
'Public Property Get DataCon() As Object
'     Set DataCon = objUserControl.Controls("DataCon")
'End Property
'
'Public Property Get ValorMercadCon() As Object
'     Set ValorMercadCon = objUserControl.Controls("ValorMercadCon")
'End Property
'
''Public Property Get FretePesoCon() As Object
''     Set FretePesoCon = objUserControl.Controls("FretePesoCon")
''End Property
'
'Public Property Get AdValorenCon() As Object
'     Set AdValorenCon = objUserControl.Controls("AdValorenCon")
'End Property
'
'Public Property Get SabadoCon() As Object
'     Set SabadoCon = objUserControl.Controls("SabadoCon")
'End Property
'
'Public Property Get ValorContainerCon() As Object
'     Set ValorContainerCon = objUserControl.Controls("ValorContainerCon")
'End Property
'
'Public Property Get DespachoCon() As Object
'     Set DespachoCon = objUserControl.Controls("DespachoCon")
'End Property
'
'Public Property Get Manuseio() As Object
'     Set Manuseio = objUserControl.Controls("Manuseio")
'End Property
'
'Public Property Get SubTotalCon() As Object
'     Set SubTotalCon = objUserControl.Controls("SubTotalCon")
'End Property
'
'Public Property Get ISSCon() As Object
'     Set ISSCon = objUserControl.Controls("ISSCon")
'End Property
'
'Public Property Get PedagioCon() As Object
'     Set PedagioCon = objUserControl.Controls("PedagioCon")
'End Property
'
'Public Property Get TotalCon() As Object
'     Set TotalCon = objUserControl.Controls("TotalCon")
'End Property
'
''Seleciona o evento clique apropriado através do índice.
'Public Function SelecionaLabel1_Click(Index As Integer)
'
'    Select Case Index
'
'        Case LABEL1_CLIENTELABEL
'            Call ClienteLabel_Click
'
'        Case LABEL1_LBLNATOPINTERNA
'            Call LblNatOpInterna_Click
'
'        Case LABEL1_SERIELABEL
'            Call SerieLabel_Click
'
'        Case LABEL1_TRANSPORTADORALABEL
'            Call TransportadoraLabel_Click
'
'        Case LABEL1_MENSAGEMLABEL
'            Call MensagemLabel_Click
'
'        Case LABEL1_SERIENFORIGINALLABEL
'            Call SerieNFOriginalLabel_Click
'
'        Case LABEL1_NFISCALORIGINALLABEL
'            Call NFiscalOriginalLabel_Click
'
'        Case LABEL1_CONDPAGTOLABEL
'            Call CondPagtoLabel_DblClick
'
'        Case LABEL1_NFISCALLABEL
'            Call NFiscalLabel_Click
'
'    End Select
'
'End Function
'
'Public Sub SelecionaUpDown_DownClick(Index As Integer)
''Seleciona o Botão UpDown através do Indice.
'
'    Select Case Index
'
'        Case UPDOWN_EMISSAO
'            Call UpDownEmissao_DownClick
'
'        Case UPDOWN_SAIDA
'            Call UpDownSaida_DownClick
'
'    End Select
'
'End Sub
'
'Public Sub SelecionaUpDown_UpClick(Index As Integer)
''Seleciona o Botão UpDown através do Indice.
'
'    Select Case Index
'
'        Case UPDOWN_EMISSAO
'            Call UpDownEmissao_UpClick
'
'        Case UPDOWN_SAIDA
'            Call UpDownSaida_UpClick
'
'    End Select
'
'End Sub
'
'Private Sub ValorContainercom_Change()
'
'    iAlterado = 0
'    iValorContainerAlterado = REGISTRO_ALTERADO
'
'End Sub
'
''William
'Private Sub ValorContainercom_GotFocus()
'
'    iValorContainerAlterado = 0
'
'End Sub
'
''William
'Private Sub ValorContainercom_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'On Error GoTo Erro_ValorContainer_Validate
'
'    If iValorContainerAlterado = 0 Then Exit Sub
'
'    'Verifica se algum valor foi digitado
'    If Len(Trim(ValorContainerCom.ClipText)) > 0 Then
'
'        'Critica se é valor Positivo
'        lErro = Valor_Positivo_Critica_Double(ValorContainerCom.Text)
'        If lErro <> SUCESSO Then gError 87540
'
'        'Põe o valor formatado na tela
'        ValorContainerCom.Text = Format(ValorContainerCom.Text, "Fixed")
'
'    End If
'
'    Exit Sub
'
'Erro_ValorContainer_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case 87450
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Property Get ValorContainerCom() As Object
'     Set ValorContainerCom = objUserControl.Controls("ValorContainerCom")
'End Property
'
'Public Property Get TotalComprovServ() As Object
'     Set TotalComprovServ = objUserControl.Controls("TotalComprovServ")
'End Property
'
'Private Function Carrega_Cobradores() As Long
''Carrega a combo de Cobrador
'
'Dim lErro As Long
'Dim colCodigoNome As New AdmColCodigoNome
'Dim objCodigoNome As New AdmCodigoNome
'Dim iIndice As Integer
'
'On Error GoTo Erro_Carrega_Cobradores
'
'    'Leitura dos códigos e descrições dos Bancos BD
'    '?????? Achar constante para nomered de banco com valor = 20
'    lErro = Cod_Nomes_Le("Cobradores", "Codigo", "NomeReduzido", 20, colCodigoNome)
'    If lErro <> SUCESSO Then Error 95180
'
'   'Preenche ComboBox com código e nome dos Bancos Cobradores
'    For iIndice = 1 To colCodigoNome.Count
'        Set objCodigoNome = colCodigoNome(iIndice)
'        Cobrador.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
'        Cobrador.ItemData(Cobrador.NewIndex) = objCodigoNome.iCodigo
'    Next
'
'    Carrega_Cobradores = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Cobradores:
'
'    Carrega_Cobradores = gErr
'
'    Select Case gErr
'
'        Case 95180 'Tratado na rotina chamada
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Public Property Get Cobrador() As Object
'     Set Cobrador = objUserControl.Controls("Cobrador")
'End Property
'
'Public Property Get CarteiraCobrador() As Object
'     Set CarteiraCobrador = objUserControl.Controls("CarteiraCobrador")
'End Property
'
'Private Sub Carrega_CarteiraCobrador(iCobrador As Integer)
''Carrega a combo de carteiras de Cobradores
'
'Dim objCobrador As New ClassCobrador
'Dim lErro As Long
'Dim objCarteiraCobrador As New ClassCarteiraCobrador
'Dim objCarteiraCobranca As New ClassCarteiraCobranca
'Dim sListBoxItem As String
'Dim colCarteirasCobrador As New Collection
'
'On Error GoTo Erro_Carrega_CarteiraCobrador
'
'    'Limpa a Combo de Carteiras
'    CarteiraCobrador.Clear
'
'    'Passa o Código do Cobrador que está na tela para o Obj
'    objCobrador.iCodigo = iCobrador
'
'    'Lê os dados do Cobrador
'    lErro = Cobrador_Le(objCobrador)
'    If lErro <> SUCESSO And lErro <> 19294 Then gError 95183
'
'    'Se o Cobrador não estiver cadastrado
'    If lErro = 19294 Then objCobrador.iCodigo = COBRADOR_PROPRIA_EMPRESA
'
'    'Le as carteiras associadas ao Cobrador
'    lErro = Cobrador_Le_Carteiras(objCobrador, colCarteirasCobrador)
'    If lErro <> SUCESSO And lErro <> 23500 Then gError 95185
'
'    If lErro = SUCESSO Then
'
'        'Preencher a Combo
'        For Each objCarteiraCobrador In colCarteirasCobrador
'
'            objCarteiraCobranca.iCodigo = objCarteiraCobrador.iCodCarteiraCobranca
'
'            lErro = CarteiraDeCobranca_Le(objCarteiraCobranca)
'            If lErro <> SUCESSO And lErro <> 23413 Then gError 95186
'
'            'Carteira não está cadastrado
'            If lErro = 23413 Then gError 95187
'
'            'Concatena Código e a Descricao da carteira
'            sListBoxItem = CStr(objCarteiraCobranca.iCodigo)
'            sListBoxItem = sListBoxItem & SEPARADOR & objCarteiraCobranca.sDescricao
'
'            CarteiraCobrador.AddItem sListBoxItem
'            CarteiraCobrador.ItemData(CarteiraCobrador.NewIndex) = objCarteiraCobranca.iCodigo
'
'        Next
'
'    End If
'    '??????
'    'Seleciona uma das Carteiras
'    If objCobrador.iCodigo = COBRADOR_PROPRIA_EMPRESA Then Cobrador.Text = Cobrador.List(0)
'    If CarteiraCobrador.ListCount <> 0 Then CarteiraCobrador.ListIndex = 0
'
'    Exit Sub
'
'Erro_Carrega_CarteiraCobrador:
'
'    Select Case gErr
'
'        Case 95183, 15836, 95185, 95186
'            Cobrador.SetFocus
'
'        Case 95184
'            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", gErr, Cobrador.Text)
'            Cobrador.SetFocus
'
'        Case 95187
'            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA", gErr, objCarteiraCobranca.iCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Public Sub Cobrador_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub Cobrador_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub Cobrador_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = Cobrador
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Public Sub CarteiraCobrador_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridParcelas)
'
'End Sub
'
'Public Sub CarteiraCobrador_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
'
'End Sub
'
'Public Sub CarteiraCobrador_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridParcelas.objControle = CarteiraCobrador
'    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
