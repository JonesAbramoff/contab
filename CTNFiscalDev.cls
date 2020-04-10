VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalDev"
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
Public giValorDescontoManual As Integer 'Alterado por Wagner

Dim gobjNFDevBenef As New ClassNFDevBenef

'distribuicao
Public gobjDistribuicao As Object

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

'************ GRADE **********************
Public gobjNFiscal As ClassNFiscal
'*****************************************

'Flag que indica se a tela está sendo preenchida.
Public gbCarregandoTela As Boolean

Public giPosCargaOk As Integer

Public gobjTelaProjetoInfo As ClassTelaPRJInfo
Dim giComplemento As Integer
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
Private Const UNIDADE_MED As String = "Unidade_Medida"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const PRECO_UNITARIO As String = "Preco_Unitario"
Private Const PRECO_TOTAL As String = "Preco_Total"
Private Const ALMOXARIFADO1 As String = "Almoxarifado"
Private Const DESCONTO1 As String = "Desconto_Produto"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"
'#######Valores####################
Private Const ICMS As String = "ICMS_Valor"
Private Const ICMSSUBST As String = "ICMSSubst_Valor"
Private Const VALOR_PRODUTOS As String = "Valor_Produtos"
Private Const VALOR_FRETE As String = "Valor_Frete"
Private Const VALOR_SEGURO As String = "Valor_Seguro"
Private Const VALOR_DESPESAS As String = "Valor_Despesas"
Private Const IPI As String = "IPI_Valor"
Private Const VALOR_DESCONTO As String = "Valor_Desconto_Glb"
Private Const VALOR_TOTAL As String = "Valor_Total"
'###########Tributação##############
Private Const ISS_VALOR As String = "ISS_Valor"
Private Const ISS_INCLUSO As String = "ISS_Incluso"
Private Const VALOR_IRRF As String = "Valor_IRRF"
Private Const VALOR_INSS As String = "Valor_INSS"
Private Const INSS_RETIDO As String = "INSS_Retido"
Private Const FORNECEDOR1 As String = "Fornecedor"
Private Const FILIAL_FORNECEDOR1 As String = "Filial_Fornecedor"
Private Const CTACONTABILEST1 As String = "ContaContabilEst"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque"
Private Const ICMS_DEBITO_VALOR As String = "ICMS_Debito_Valor"
Private Const IPI_DEBITO_VALOR As String = "IPI_Debito_Valor"
Private Const ICMS_DEBITO_FRETE As String = "ICMS_Debito_Frete"
Private Const ICMS_DEBITO_SEGURO As String = "ICMS_Debito_Seguro"
Private Const ICMS_DEBITO_DESPESAS As String = "ICMS_Debito_Despesas"
Private Const IPI_DEBITO_FRETE As String = "IPI_Debito_Frete"
Private Const IPI_DEBITO_SEGURO As String = "IPI_Debito_Seguro"
Private Const IPI_DEBITO_DESPESAS As String = "IPI_Debito_Despesas"
Private Const QUANT_DISPONIVEL As String = "Quant_Disponivel"
Private Const QUANT_CONSIGTERC As String = "Quant_ConsigTerc"
Private Const QUANT_CONSERTOTERC As String = "Quant_ConsertoTerc"
Private Const QUANT_DEMONSTERC As String = "Quant_DemonsTerc"
Private Const QUANT_OUTRASTERC As String = "Quant_OutrasTerc"
Private Const CTAFORNCONSIG As String = "CtaFornConsig"
Private Const ICMS_DEBITO_TOTAL As String = "ICMS_Debito_Total"
Private Const IPI_DEBITO_TOTAL As String = "IPI_Debito_Total"
Private Const PIS_DEBITO_TOTAL As String = "PIS_Debito_Total"
Private Const COFINS_DEBITO_TOTAL As String = "COFINS_Debito_Total"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const ISS_RETIDO As String = "ISS_Retido"
Private Const COFINS_RETIDO As String = "COFINS_Retido"
Private Const CSLL_RETIDO As String = "CSLL_Retido"
Private Const CONTACONTABILAPLIC As String = "ContaContabilAplic"
Private Const CONTACONTABILCATPROD As String = "ContaContabilCatProd"
Private Const CTAFILIALFORNECEDOR1 As String = "FilialForn_Conta_Ctb"
Private Const FilialEmpresa As String = "FilialEmpresa"
Private Const PIS_VALOR_TOTAL As String = "PIS_Valor_Total"
Private Const COFINS_VALOR_TOTAL As String = "COFINS_Valor_Total"
'fim da contabilidade

Public iAlterado As Integer
Public iClienteAlterado As Integer
Public iFornecedorAlterado As Integer
Public iFrameAtual As Integer
Public iValorDescontoAlterado As Integer
Public iValorFreteAlterado As Integer
Public iValorSeguroAlterado As Integer
Public iValorDespesasAlterado As Integer
Public iFornecedorAtivo As Integer
Public giFilialEntregaAlterada As Integer
Public giFilialEntregaManual As Integer
Public iDataEmissaoAlterada As Integer
Public iSerieAlterada As Integer

Public objGridItens As AdmGrid
Public iGrid_Produto_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
'Public iGrid_Almoxarifado_Col As Integer
Public iGrid_ValorUnitario_Col As Integer
Public iGrid_ValorUnitarioMoeda_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PrecoTotal_Col As Integer
Public iGrid_PrecoTotalB_Col As Integer
Public iGrid_SerieNFOrig_Col As Integer
Public iGrid_CodNFOrig_Col As Integer
Public iGrid_ItemNFOrig_Col As Integer

Private WithEvents objEventoSerieOriginal As AdmEvento
Attribute objEventoSerieOriginal.VB_VarHelpID = -1
Private WithEvents objEventoNFOriginal As AdmEvento
Attribute objEventoNFOriginal.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Public WithEvents objEventoNFiscal As AdmEvento
Attribute objEventoNFiscal.VB_VarHelpID = -1
Private WithEvents objEventoItensNFiscalOrig As AdmEvento
Attribute objEventoItensNFiscalOrig.VB_VarHelpID = -1
Private WithEvents objEventoLocalizacao As AdmEvento
Attribute objEventoLocalizacao.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoTranspRedespacho As AdmEvento
Attribute objEventoTranspRedespacho.VB_VarHelpID = -1
Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Private WithEvents objEventoMensagem As AdmEvento
Attribute objEventoMensagem.VB_VarHelpID = -1


''*** incluidos p/tratamento de tributacao *******************************
Public gobjTribTab As ClassTribTab

Public gobjRastreamento As ClassRastreamento

'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_Itens = 2
Private Const TAB_Complemento = 3
Private Const TAB_Distribuicao = 4
Private Const TAB_Tributacao = 5
Private Const TAB_Contabilizacao = 6
Private Const TAB_RASTRO = 7

'--- inicio dos properties get dos controles da tela

Public Property Get BotaoOPBenef() As Object
     Set BotaoOPBenef = objUserControl.Controls("BotaoOPBenef")
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

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get TipoFrete() As Object
     Set TipoFrete = objUserControl.Controls("TipoFrete")
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

Public Property Get PesoLiquido() As Object
     Set PesoLiquido = objUserControl.Controls("PesoLiquido")
End Property

Public Property Get PesoBruto() As Object
     Set PesoBruto = objUserControl.Controls("PesoBruto")
End Property

Public Property Get MensagemLabel() As Object
     Set MensagemLabel = objUserControl.Controls("MensagemLabel")
End Property

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
End Property

Public Property Get BotaoItensNFiscalOrig() As Object
     Set BotaoItensNFiscalOrig = objUserControl.Controls("BotaoItensNFiscalOrig")
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

Public Property Get Total() As Object
     Set Total = objUserControl.Controls("Total")
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

Public Property Get DescricaoItem() As Object
     Set DescricaoItem = objUserControl.Controls("DescricaoItem")
End Property

Public Property Get UnidadeMed() As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
End Property

Public Property Get PercentDesc() As Object
     Set PercentDesc = objUserControl.Controls("PercentDesc")
End Property

Public Property Get Desconto() As Object
     Set Desconto = objUserControl.Controls("Desconto")
End Property

Public Property Get Almoxarifado() As Object
     Set Almoxarifado = objUserControl.Controls("Almoxarifado")
End Property

Public Property Get ItemNFOriginal() As Object
     Set ItemNFOriginal = objUserControl.Controls("ItemNFOriginal")
End Property

Public Property Get CodigoNFOriginal() As Object
     Set CodigoNFOriginal = objUserControl.Controls("CodigoNFOriginal")
End Property

Public Property Get SerieNFOriginal() As Object
     Set SerieNFOriginal = objUserControl.Controls("SerieNFOriginal")
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

Public Property Get BotaoExibirDados() As Object
     Set BotaoExibirDados = objUserControl.Controls("BotaoExibirDados")
End Property

Public Property Get SerieNFiscalOriginal() As Object
     Set SerieNFiscalOriginal = objUserControl.Controls("SerieNFiscalOriginal")
End Property

Public Property Get NFiscalPropria() As Object
     Set NFiscalPropria = objUserControl.Controls("NFiscalPropria")
End Property

Public Property Get NFiscalOriginal() As Object
     Set NFiscalOriginal = objUserControl.Controls("NFiscalOriginal")
End Property

Public Property Get NFOriginalLabel() As Object
     Set NFOriginalLabel = objUserControl.Controls("NFOriginalLabel")
End Property

Public Property Get SerieOriginalLabel() As Object
     Set SerieOriginalLabel = objUserControl.Controls("SerieOriginalLabel")
End Property

Public Property Get BotaoLimparNF() As Object
     Set BotaoLimparNF = objUserControl.Controls("BotaoLimparNF")
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

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get Fornecedor() As Object
     Set Fornecedor = objUserControl.Controls("Fornecedor")
End Property

Public Property Get ClienteLabel() As Object
     Set ClienteLabel = objUserControl.Controls("ClienteLabel")
End Property

Public Property Get FornecedorLabel() As Object
     Set FornecedorLabel = objUserControl.Controls("FornecedorLabel")
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

Public Property Get TabelaPreco() As Object
     Set TabelaPreco = objUserControl.Controls("TabelaPreco")
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

Public Property Get Status() As Object
    Set Status = objUserControl.Controls("Status")
End Property

Public Property Get LabelFilial() As Object
     Set LabelFilial = objUserControl.Controls("LabelFilial")
End Property

Public Property Get LabelFilialEntrega() As Object
     Set LabelFilialEntrega = objUserControl.Controls("LabelFilialEntrega")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

'Rastreamento
Public Property Get GridRastro() As Object
     Set GridRastro = objUserControl.Controls("GridRastro")
End Property

Public Property Get BotaoConsultaDocCPR() As Object
     Set BotaoConsultaDocCPR = objUserControl.Controls("BotaoConsultaDocCPR")
End Property

Public Property Get ImprimeRomaneio() As Object
     Set ImprimeRomaneio = objUserControl.Controls("ImprimeRomaneio")
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

Public Property Get BotaoItensNFOrigAuto() As Object
     Set BotaoItensNFOrigAuto = objUserControl.Controls("BotaoItensNFOrigAuto")
End Property

Public Property Get NumNFe() As Object
     Set NumNFe = objUserControl.Controls("NumNFe")
End Property

Public Property Get EletronicaFed() As Object
     Set EletronicaFed = objUserControl.Controls("EletronicaFed")
End Property

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

Public Sub NFiscal_GotFocus()
    Call MaskEdBox_TrataGotFocus(NFiscal, iAlterado)
End Sub

Public Sub NFiscalOriginal_GotFocus()
    Call MaskEdBox_TrataGotFocus(NFiscalOriginal, iAlterado)
End Sub

Public Sub ImprimeRomaneio_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub Form_Activate()

Dim lErro As Long

On Error GoTo Erro_Form_Activate

    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError 59334

    Call TelaIndice_Preenche(Me)

    Exit Sub

Erro_Form_Activate:

    Select Case gErr

        Case 59334

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156293)

    End Select

    Exit Sub

End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Public Function CargaPosFormLoad(Optional bTrazendoDoc As Boolean = False) As Long

Dim iSubTipo As Integer
Dim lErro As Long

On Error GoTo Erro_CargaPosFormLoad

    If (giPosCargaOk = 0) Then

        'p/permitir o redesenho da tela
        DoEvents

        gbCarregandoTela = True
        
        'Rastreamento
        Set gobjRastreamento = New ClassRastreamento
        Set gobjRastreamento.objTela = Me
        gobjRastreamento.bDevolucaoMaterial = True
        gobjRastreamento.bSaidaMaterial = True

    ''*** incluidos p/tratamento de tributacao *******************************
        Call GridItens_InicializaColunas
        Set gobjTribTab = New ClassTribTab
        lErro = gobjTribTab.Ativar(Me)
        If lErro <> SUCESSO Then gError 39303
    '*** fim tributacao

        Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
        Set gobjTelaProjetoInfo.objUserControl = objUserControl
        Set gobjTelaProjetoInfo.objTela = Me

        'Carrega a combo de Tipos de Documento
        lErro = Carrega_TiposDocInfo()
        If lErro <> SUCESSO Then gError 39045

        'Carrega a combo de Séries
        lErro = Carrega_Serie()
        If lErro <> SUCESSO Then gError 39047

        'Seleciona série padrão
        lErro = Seleciona_Serie_Padrao()
        If lErro <> SUCESSO Then gError 500121

        'Carrega a combo de Tabelas de Preço
        lErro = Carrega_TabelaPreco()
        If lErro <> SUCESSO Then gError 39049

        'Carrega a combo de Transportadora
        lErro = Carrega_Transportadora()
        If lErro <> SUCESSO Then gError 39051

        'Incluído por Luiz Nogueira em 21/08/03
        'Carrega a combo VolumeEspecie
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
        If lErro <> SUCESSO Then gError 102472
    
        'Carrega a combo VolumeMarca
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
        If lErro <> SUCESSO Then gError 102473
        
        'Carrega a Combo de Estados
        lErro = Carrega_PlacaUF()
        If lErro <> SUCESSO Then gError 39053

        'PrecoUnitario.Format = gobjFAT.sFormatoPrecoUnitario
        PrecoUnitario.Format = FORMATO_PRECO_UNITARIO_EXTERNO

        Quantidade.Format = FORMATO_ESTOQUE

        'Inicializa as Variáveis de browse
        Set objGridItens = New AdmGrid
        Set objEventoSerie = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoCliente = New AdmEvento
        Set objEventoNFiscal = New AdmEvento
        Set objEventoMensagem = New AdmEvento
        Set objEventoNatureza = New AdmEvento
        Set objEventoNFOriginal = New AdmEvento
        Set objEventoFornecedor = New AdmEvento
        Set objEventoLocalizacao = New AdmEvento
        Set objEventoSerieOriginal = New AdmEvento
        Set objEventoTransportadora = New AdmEvento
        Set objEventoTranspRedespacho = New AdmEvento
        Set objEventoItensNFiscalOrig = New AdmEvento

        lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
        If lErro <> SUCESSO Then gError 39055

        'Inicializa o Grid de itens
        lErro = Inicializa_Grid_Itens(objGridItens)
        If lErro <> SUCESSO Then gError 39054

        'Inicializa a contabilidade carregando apenas os modelos contábeis referentes ao TipoDocInfo default
        If objContabil.objGridInt Is Nothing Then

            'Obtém o código do subtipo contábil referente ao TipoDocInfo default
            lErro = CF("Obtem_Subtipo_Contabil", "TiposDocInfo", DOCINFO_NFISDCM, iSubTipo)
            If lErro <> SUCESSO And lErro <> 79671 Then gError 79686

            'Se não encontrou o subtipo => erro
            If lErro = 79671 Then gError 79687

            'Inicialização da parte de contabilidade
            lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_FATURAMENTO, iSubTipo)
            If lErro <> SUCESSO Then gError 79330

        End If

        'distribuicao
        Set gobjDistribuicao = CreateObject("RotinasMat.ClassMATDist")
        Set gobjDistribuicao.objTela = Me
        Set gobjDistribuicao.objRastreamento = gobjRastreamento

        'Inicializa o grid de Rastreamento
        lErro = gobjRastreamento.Inicializa_Grid_Rastreamento()
        If lErro <> SUCESSO Then gError 83410

        lErro = gobjDistribuicao.Inicializa_GridDist()
        If lErro <> SUCESSO Then gError 83410

'        TipoNFiscal.ListIndex = 1

        gbCarregandoTela = False

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

        Case 39045, 39047, 39049, 39051, 39053, 39054, 39055, 39303, 39664, 79330, 79687, 83410, 500121, 102472, 102473

        Case 79687
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPODOCINFO_NAO_ENCONTRADO", gErr, DOCINFO_NFISDCM)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156294)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    
    If InStr(UCase(gsNomeEmpresa), "SHOWLUZ") <> 0 Then Total.Visible = True
    
    Call CF2(Me, "NFCF_Default_DataSaida")
    
    Call CF2(Me, "NFDev_Form_Load")
    
    If giTipoVersao = VERSAO_LIGHT Then
        
        BotaoGrade.Visible = False
        
        'sumir com abas de contabilidade e rastro
        TabStrip1.Tabs.Remove (TAB_RASTRO)
        TabStrip1.Tabs.Remove (TAB_Contabilizacao)
        
    End If
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    MsgAutomatica.Value = vbChecked
    Mensagem.MaxLength = MENSAGEM_NF_MAXLENGTH
    MensagemCorpo.MaxLength = MENSAGEM_NF_MAXLENGTH
    
    Set gobjNFiscal = New ClassNFiscal
        
    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207428
    
    If gobjFAT.iTrataNFDevSimbPorOP = MARCADO Then
        BotaoOPBenef.Visible = True
    Else
        BotaoOPBenef.Visible = False
    End If
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 207428

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156295)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Function Form_Load_Light()

Dim lErro As Long

On Error GoTo Erro_Form_Load_Light

    'obs.: a carga pos form_load será feita pela classe que trata nffat
    giPosCargaOk = 1

    'Inicializa as Variáveis de browse
    Set objEventoSerie = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoNFiscal = New AdmEvento
    Set objEventoMensagem = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    Set objEventoNFOriginal = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoLocalizacao = New AdmEvento
    Set objEventoSerieOriginal = New AdmEvento
    Set objEventoTransportadora = New AdmEvento
    Set objEventoItensNFiscalOrig = New AdmEvento

    Call GridItens_InicializaColunas

    iAlterado = 0

    Form_Load_Light = SUCESSO

    Exit Function

Erro_Form_Load_Light:

    Form_Load_Light = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156296)

    End Select

    iAlterado = 0

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

Public Sub Almoxarifado_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Almoxarifado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub Almoxarifado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Almoxarifado
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO

    Call Cliente_Preenche

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objTipoCliente As New ClassTipoCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objMensagem As New ClassMensagem
Dim objTipoFrete As ClassTipoFrete
Dim colTipoFrete As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Cliente_Validate

    'Verifica se o cliente foi alterado
    If iClienteAlterado = 1 Then

        If Len(Trim(Cliente.Text)) > 0 Then
            'Lê o Cliente digitado
            lErro = TP_Cliente_Le3(Cliente, objCliente, iCodFilial)
            If lErro <> SUCESSO Then gError 39123

            'Le as Filiais do Cliente
            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then gError 39124

            'Preenche as ComboBoxes de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)
            Call CF("Filial_Preenche", FilialEntrega, colCodigoNome)

            If Not gbCarregandoTela Then

                If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then

                    If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ

                    'Seleciona filial nas Combos de Filial
                    Call CF("Filial_Seleciona", Filial, iCodFilial)
                    Call CF("Filial_Seleciona", FilialEntrega, iCodFilial)

                End If

            End If

            lErro = CF2(Me, "NFCF_Cliente_Validate", objCliente.lCodigo, iCodFilial)
            If lErro <> SUCESSO Then gError 131879
           
            giFilialEntregaManual = 0

            If Not gbCarregandoTela Then
                If objCliente.iTabelaPreco > 0 Then
                    TabelaPreco.Text = objCliente.iTabelaPreco
                ElseIf objCliente.iTipo > 0 Then
                    objTipoCliente.iCodigo = objCliente.iTipo
                    lErro = CF("TipoCliente_Le", objTipoCliente)
                    If lErro <> SUCESSO And lErro <> 19062 Then gError 39125
                    If lErro = SUCESSO And objTipoCliente.iTabelaPreco > 0 Then TabelaPreco.Text = objTipoCliente.iTabelaPreco
                End If

                If Len(Trim(TabelaPreco.Text)) > 0 Then TabelaPreco_Validate (bSGECancelDummy)

                lErro = CF2(Me, "NFCF_Cliente_Validate3", objCliente)
                If lErro <> SUCESSO Then gError 51628
            
                'Tipo de frete
                lErro = CF("TipoFrete_Le", colTipoFrete)
                If lErro <> SUCESSO Then gError 207422
                
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

    ''*** incluidos p/tratamento de tributacao *******************************
                If iCodFilial <> 0 Then
                    Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, iCodFilial) '####
                    Call Trata_FilialCli_Dev(objCliente.lCodigo, iCodFilial)
                End If
    '*** fim tributacao
            End If

        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            Filial.Clear
            FilialEntrega.Clear

            Call CF2(Me, "NFCF_Limpa_Tela")

        End If

        iClienteAlterado = 0

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 39123, 39124, 51628, 131879, 207422
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156297)

    End Select

    Exit Sub

End Sub

Public Sub CodigoNFOriginal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CodigoNFOriginal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub CodigoNFOriginal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub CodigoNFOriginal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = CodigoNFOriginal
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataEmissao_Change()
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
            If lErro <> SUCESSO Then gError 39140
        
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

        Case 39140

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156298)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 39141

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 39141

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156299)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 39142

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 39142

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156300)

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
    If lErro <> SUCESSO Then gError 39143

    Exit Sub

Erro_DataSaida_Validate:

    Cancel = True


    Select Case gErr

        Case 39143

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156301)

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
    If lErro <> SUCESSO Then gError 89797

    Exit Sub

Erro_HoraSaida_Validate:

    Cancel = True

    Select Case gErr

        Case 89797

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156302)

    End Select

    Exit Sub

End Sub

Public Sub UpDownSaida_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownSaida_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataSaida, DIMINUI_DATA)
    If lErro Then gError 39144

    Exit Sub

Erro_UpDownSaida_DownClick:

    Select Case gErr

        Case 39144

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156303)

    End Select

    Exit Sub

End Sub

Public Sub UpDownSaida_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownSaida_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataSaida, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 39145

    Exit Sub

Erro_UpDownSaida_UpClick:

    Select Case gErr

        Case 39145

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156304)

    End Select

    Exit Sub

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

Public Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    iAlterado = REGISTRO_ALTERADO

    'Verifica se alguma filial foi selecionada
    If Filial.ListIndex = -1 Then Exit Sub

    If iFornecedorAtivo = FORN_INATIVO And Len(Trim(Transportadora.Text)) = 0 Then
        lErro = Trata_FilialCliente()
        If lErro <> SUCESSO Then gError 39132
    Else
        lErro = Trata_FilialForn()
        If lErro <> SUCESSO Then gError 56964
    End If

    Exit Sub

Erro_Filial_Click:

    Select Case gErr

        Case 39132, 56964

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156305)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFilialCliente As New ClassFilialCliente
Dim sNomeRed As String
Dim vbMsgRes As VbMsgBoxResult
Dim objTipoFrete As ClassTipoFrete
Dim colTipoFrete As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 39130

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        If iFornecedorAtivo = FORN_ATIVO Then

            'Verifica de o fornecedor foi digitado
            If Len(Trim(Fornecedor.Text)) = 0 Then gError 39131

            sNomeRed = Fornecedor.Text

            objFilialFornecedor.iCodFilial = iCodigo

            'Pesquisa se existe filial com o codigo extraido
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 39133

            If lErro = 18272 Then gError 39134

            'Coloca na tela
            Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

            Call gobjTribTab.FilialFornecedor_Alterada(objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)

        Else

            If Len(Trim(Cliente.ClipText)) = 0 Then gError 39135

            sNomeRed = Cliente.Text

            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeRed, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 39136

            If lErro = 17660 Then gError 42853

            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

            'Seleciona tipo de frete da Filial do Cliente
            lErro = CF("TipoFrete_Le", colTipoFrete)
            If lErro <> SUCESSO Then gError 207423
            
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

            Call gobjTribTab.FilialCliente_Alterada(objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
            Call Trata_FilialCli_Dev(objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    
        End If

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 42854

    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case gErr

        Case 39133, 39136, 39130, 207423

        Case 39134
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 39131
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 39135
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 42853
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 42854
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156306)

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
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFilialCliente As New ClassFilialCliente
Dim sNomeRed As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialEntrega_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialEntrega.Text)) = 0 Or giFilialEntregaAlterada = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialEntrega.Text = FilialEntrega.List(FilialEntrega.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialEntrega, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 39204

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        If iFornecedorAtivo = FORN_ATIVO Then

            'Verifica de o fornecedor foi digitado
            If Len(Trim(Fornecedor.Text)) = 0 Then gError 39205

            sNomeRed = Fornecedor.Text

            objFilialFornecedor.iCodFilial = iCodigo

            'Pesquisa se existe filial com o codigo extraido
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 39206

            If lErro = 18272 Then gError 39207

            'Coloca na tela
            FilialEntrega.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

        Else

            'Verifica se Cliente foi preenchido
            If Len(Trim(Cliente.ClipText)) = 0 Then gError 39208

            sNomeRed = Cliente.Text

            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeRed, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 42855

            'Se não achou a Filial Cliente --> erro
            If lErro = 17660 Then gError 42856

            FilialEntrega.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        End If

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 42857

    giFilialEntregaAlterada = 0

    Exit Sub

Erro_FilialEntrega_Validate:

    Cancel = True

    Select Case gErr

        Case 39206, 42855, 39204

        Case 39207
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 39205
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 39208
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 42856
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 42857
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, FilialEntrega.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156307)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set gobjInfoUsu = Nothing
    
    Set gobjTelaProjetoInfo = Nothing
    Set gobjInfoAdic = Nothing
    Set gobjDistribuicao = Nothing
    Set gobjNFiscal = Nothing

    'Fechaocomando se setas
    lErro = ComandoSeta_Liberar(Me.Name)

    Set gobjAnotacao = Nothing
    
    'Encerra as variáveis ligadas aos browses
    Set objEventoSerie = Nothing
    Set objEventoProduto = Nothing
    Set objEventoCliente = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoNatureza = Nothing
    Set objEventoMensagem = Nothing
    Set objEventoNFOriginal = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoLocalizacao = Nothing
    Set objEventoSerieOriginal = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoTranspRedespacho = Nothing
    Set objEventoItensNFiscalOrig = Nothing

    Set objGrid1 = Nothing
    Set objContabil = Nothing
    Set objGridItens = Nothing

    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    Set gobjRastreamento = Nothing
    Set gobjNFDevBenef = Nothing

    Call CF2(Me, "NFDev_Form_UnLoad")

''*** incluidos p/tratamento de tributacao *******************************
    If Not (gobjTribTab Is Nothing) Then Call gobjTribTab.Desativar
    Set gobjTribTab = Nothing
'*** fim tributacao

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156308)

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

Public Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO

    Call Fornecedor_Preenche

End Sub

Public Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le3(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO And lErro <> 6664 Then gError 39126
            If lErro = 6664 Then gError 51627

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 39127

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)
            Call CF("Filial_Preenche", FilialEntrega, colCodigoNome)

            If Not gbCarregandoTela Then

                If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then

                    If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ

                    'Seleciona filial na Combo Filial
                    Call CF("Filial_Seleciona", Filial, iCodFilial)
                    Call CF("Filial_Seleciona", FilialEntrega, iCodFilial)

                    ''*** incluidos p/tratamento de tributacao *******************************
                    Call gobjTribTab.FilialFornecedor_Alterada(objFornecedor.lCodigo, iCodFilial)
                    '*** fim tributacao

                End If

            End If

            lErro = CF2(Me, "NFCF_Fornecedor_Validate", objFornecedor.lCodigo, iCodFilial)
            If lErro <> SUCESSO Then gError 131880
        
        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

            Filial.Clear
            FilialEntrega.Clear

            Call CF2(Me, "NFCF_Limpa_Tela")

        End If

        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 39126, 39127

        Case 51627, 131880

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156309)

    End Select

    Exit Sub

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

Public Sub Mensagem_Change()

    iAlterado = REGISTRO_ALTERADO
    MsgAutomatica.Value = vbUnchecked

End Sub

Private Sub NaturezaOp_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NFiscal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NFiscalOriginal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NFiscalPropria_Click()
    iAlterado = REGISTRO_ALTERADO
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
    If lErro <> SUCESSO Then gError 39216

    'Coloca o valor formatado na Tela
    PesoBruto.Text = Format(PesoBruto.Text, PesoBruto.Format)

    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True

    Select Case gErr

        Case 39216

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156310)

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
    If lErro <> SUCESSO Then gError 39215

    'Coloca o valor formatado na Tela
    PesoLiquido.Text = Format(PesoLiquido.Text, PesoLiquido.Format)

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True

    Select Case gErr

        Case 39215

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156311)

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

    'Verifica se tem alguma PlacaUF preenchida
    If Len(Trim(PlacaUF.Text)) = 0 Then Exit Sub

    'Verifica se existe o ítem na combo
    lErro = Combo_Item_Igual(PlacaUF)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 39213

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 39214

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True


    Select Case gErr

        Case 39213

        Case 39214
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156312)

    End Select

    Exit Sub

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
            If lErro <> SUCESSO And lErro <> 12253 Then gError 39139
            'Se não está na combo
            If lErro <> SUCESSO Then
        
                objSerie.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
                'Busca a série no BD
                lErro = CF("Serie_Le", objSerie)
                If lErro <> SUCESSO And lErro <> 22202 Then gError 51622
                If lErro <> SUCESSO Then gError 39138
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

        Case 39138
            'Pergunta se deseja criar nova série
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_SERIE", Serie.Text)
            'Se a resposta for afirmativa
            If vbMsgRes = vbYes Then
                'Chama a tela de cadastro de séries
                Call Chama_Tela("SerieNFiscal", objSerie)
            End If
            'segura o foco na série

        Case 39139, 51622

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156313)

    End Select

End Sub

Public Sub SerieNFiscalOriginal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SerieNFiscalOriginal_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SerieNFiscalOriginal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SerieNFiscalOriginal_Validate

    'Verififca se está preenchida
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Then Exit Sub

    'Verifica se foi alguma Série selecionada
    If SerieNFiscalOriginal.Text = SerieNFiscalOriginal.List(SerieNFiscalOriginal.ListIndex) Then Exit Sub

    'Tenta achar a Série na combo
    lErro = Combo_Item_Igual(SerieNFiscalOriginal)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 39096

    'Não encontrou
    If lErro = 12253 Then
        If NFiscalPropria.Value = vbChecked Then
            gError 39097
        Else
            If Len(Trim(SerieNFiscalOriginal.Text)) > STRING_SERIE Then gError 42858
        End If
    End If

    Exit Sub

Erro_SerieNFiscalOriginal_Validate:

    Cancel = True

    Select Case gErr

        Case 39096

        Case 39097
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, SerieNFiscalOriginal.Text)

        Case 42858
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_SERIE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156314)

    End Select

    Exit Sub

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
    If lErro <> SUCESSO Then gError 39146

    Exit Sub

Erro_TabelaPreco_Click:

    Select Case gErr

        Case 39146

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156315)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 39151

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTabelaPreco.iCodigo = iCodigo

        'Tenta ler TabelaPreço com esse código no BD
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 39152

        'Não encontrou Tabela Preço no BD
        If lErro <> SUCESSO Then gError 39153

        'Encontrou TabelaPreço no BD, coloca no Text da Combo
        TabelaPreco.Text = CStr(objTabelaPreco.iCodigo) & SEPARADOR & objTabelaPreco.sDescricao

        lErro = Trata_TabelaPreco()
        If lErro <> SUCESSO Then gError 39154

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 39155

    Exit Sub

Erro_TabelaPreco_Validate:

    Cancel = True


    Select Case gErr

        Case 39151, 39152, 39154

        Case 39153  'Não encontrou Tabela de Preço no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TABELA_PRECO")

            If vbMsgRes = vbYes Then
                'Chama a tela de Tabelas de Preço
                Call Chama_Tela("TabelaPrecoCriacao", objTabelaPreco)
            End If
            'Segura o foco

        Case 39155
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_ENCONTRADA", gErr, TabelaPreco.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156316)

    End Select

    Exit Sub

End Sub

Public Sub TipoNFiscal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoNFiscal_Click()

Dim objTipoDocInfo As New ClassTipoDocInfo
Dim lErro As Long

On Error GoTo Erro_TipoNFiscal_Click

    iAlterado = REGISTRO_ALTERADO

    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
    gobjRastreamento.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 39119

    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 39120
    
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then
        giComplemento = MARCADO
    Else
        giComplemento = DESMARCADO
    End If

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
    If IsDate(DataEmissao.Text) Then Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
    Call gobjTribTab.Serie_Alterada(Serie.Text)
'*** fim tributacao

    'verifica a Origem do Tipo da Nota Fiscal
    If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Then
        Call Habilita_Cliente
    ElseIf objTipoDocInfo.iDestinatario = DOCINFO_FORNECEDOR Then
        Call Habilita_Fornecedor
    End If

    'Processa a alteração do subtipo para que sejam carregados os modelos contábeis correspondentes ao novo TipoDocInfo selecionado
    lErro = objContabil.Contabil_Processa_Alteracao_Subtipo(objTipoDocInfo.iSubTipoContabil)
    If lErro <> SUCESSO Then gError 79329

    Exit Sub

Erro_TipoNFiscal_Click:

    Select Case gErr

        Case 39119, 79329

        Case 39120
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156317)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 39121
    'Se não conseguir --> Erro
    If lErro <> SUCESSO Then gError 39122

    Exit Sub

Erro_TipoNFiscal_Validate:

    Cancel = True


    Select Case gErr

        Case 39121

        Case 39122
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156318)

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
Dim vbMsgRes As VbMsgBoxResult
Dim objTransportadora As New ClassTransportadora
Dim iCodigo As Integer

On Error GoTo Erro_Transportadora_Validate

    'Verifica se foi preenchida a ComboBox Transportadora
    If Len(Trim(Transportadora.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Transportadora
    If Transportadora.Text = Transportadora.List(Transportadora.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Transportadora, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 39209
    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo
        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 39210
        If lErro <> SUCESSO Then gError 39212 'Não encontrou Transportadora no BD

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 39211

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True


    Select Case gErr

        Case 39209, 39210


        Case 39212  'Não encontrou Transportadora no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TRANSPORTADORA")

            If vbMsgRes = vbYes Then

                Call Chama_Tela("Transportadora", objTransportadora)
            End If
            'Segura o foco

        Case 39211
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)

        Case Else

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156319)

    End Select

    Exit Sub

End Sub

Public Sub UnidadeMed_Click()
    iAlterado = REGISTRO_ALTERADO
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

Public Sub ValorSeguro_Change()
    iAlterado = REGISTRO_ALTERADO
    iValorSeguroAlterado = REGISTRO_ALTERADO
End Sub

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
    If lErro <> SUCESSO Then gError 102474
    
    Exit Sub

Erro_VolumeEspecie_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102474
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156320)

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
    If lErro <> SUCESSO Then gError 102475
    
    Exit Sub

Erro_VolumeMarca_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102475
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156321)

    End Select

End Sub

Public Sub VolumeNumero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub VolumeQuant_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Carrega_TiposDocInfo() As Long
'Carrega na os Tipo de Documentos relacionados com a tela

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
            If objTipoDocInfo.iPadrao = Padrao Then
                TipoNFiscal.ListIndex = TipoNFiscal.NewIndex
            End If
        End If
    Next

    Carrega_TiposDocInfo = SUCESSO

    Exit Function

Erro_Carrega_TiposDocInfo:

    Carrega_TiposDocInfo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156322)

    End Select

    Exit Function

End Function

Private Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie
Dim sSerieAnt As String

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 39046

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

        Case 39046, 500114

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156323)

    End Select

    Exit Function

End Function

Function Seleciona_Serie_Padrao() As Long

Dim lErro As Long
Dim sSeriePadrao As String

On Error GoTo Erro_Seleciona_Serie_Padrao

    'Lê série Padrão
    lErro = CF("Serie_Le_Padrao", sSeriePadrao)
    If lErro <> SUCESSO Then gError 500120

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

        Case 500120

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156324)

    End Select

    Exit Function

End Function

Private Function Carrega_TabelaPreco() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TabelaPreco

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TabelasDePrecoVenda", "Codigo", "Descricao", STRING_TABELA_PRECO_DESCRICAO, colCodigoDescricao, "Ativo = 1")
    If lErro <> SUCESSO Then gError 39048

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o ítem na Lista de Tabela de Preços
        TabelaPreco.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TabelaPreco.ItemData(TabelaPreco.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = gErr

    Select Case gErr

        Case 39048

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156325)

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
    If lErro <> SUCESSO Then gError 39050

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

        Case 39050

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156326)

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

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156327)

    End Select

End Function

Sub GridItens_InicializaColunas()
    Call CF2(Me, "NFDEV_GridItens_InicializaColunas")
End Sub

Public Function NFDEV_GridItens_InicializaColunas(ByVal objCT As Object) As Long

    'Colunas da Grid
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_ValorUnitario_Col = 5
    iGrid_PercDesc_Col = 6
    iGrid_Desconto_Col = 7
    iGrid_PrecoTotal_Col = 8
    iGrid_PrecoTotalB_Col = 9
    iGrid_SerieNFOrig_Col = 10
    iGrid_CodNFOrig_Col = 11
    iGrid_ItemNFOrig_Col = 12
    
    NFDEV_GridItens_InicializaColunas = SUCESSO

End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
    Inicializa_Grid_Itens = CF2(Me, "NFDEV_Inicializa_Grid_Itens", objGridInt)
End Function

Public Function NFDEV_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")
    objGridInt.colColuna.Add ("Série NF Orig")
    objGridInt.colColuna.Add ("Num NF Orig")
    objGridInt.colColuna.Add ("Item NF Orig")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    objGridInt.colCampo.Add (PrecoTotalB.Name)
    objGridInt.colCampo.Add (SerieNFOriginal.Name)
    objGridInt.colCampo.Add (CodigoNFOriginal.Name)
    objGridInt.colCampo.Add (ItemNFOriginal.Name)

    Call CF2(Me, "NFDEV_GridItens_InicializaColunas")

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_NF + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NFDEV_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Sub SerieOriginalLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'Recolhe a Série da Nota Fiscal Original da tela
    objSerie.sSerie = SerieNFiscalOriginal.Text

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerieOriginal)

    Exit Sub

End Sub

Private Sub Class_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
End Sub

Private Sub objEventoSerieOriginal_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    'Coloca a Série da Nota Fiscal Original na tela
    SerieNFiscalOriginal.Text = objSerie.sSerie

    Me.Show

End Sub

Public Sub NFOriginalLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As Collection

    'Guarda a Serie e o Número da Nota Fiscal Original da Tela
    objNFiscal.sSerie = SerieNFiscalOriginal.Text
    If Len(Trim(NFiscalOriginal.ClipText)) > 0 Then
        objNFiscal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
    Else
        objNFiscal.lNumNotaFiscal = 0
    End If

    'Chama a Tela NFEntNaoDevLista
    If giComplemento = MARCADO Then
        Call Chama_Tela("NFiscalDevLista", colSelecao, objNFiscal, objEventoNFOriginal)
    Else
        Call Chama_Tela("NFEntNaoDevLista", colSelecao, objNFiscal, objEventoNFOriginal)
    End If
    
    Exit Sub

End Sub

Private Sub objEventoNFOriginal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo

    Set objNFiscal = obj1

    'Preenche a Série e o Número da Nota Fiscal Original
    SerieNFiscalOriginal.Text = objNFiscal.sSerie
    NFiscalOriginal.Text = objNFiscal.lNumNotaFiscal

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

    'Lê o tipo de documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro = SUCESSO Then

        If objTipoDocInfo.iEmitente = DOCINFO_EMPRESA Then
            NFiscalPropria.Value = vbChecked
        Else
            NFiscalPropria.Value = vbUnchecked
        End If

    End If

    Me.Show

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

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    Fornecedor.Text = objFornecedor.lCodigo
    Call Fornecedor_Validate(bSGECancelDummy)

    Me.Show

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

On Error GoTo Erro_NFiscalLabel_Click

    'Recolhe os dados da NFiscal que estão na tela
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 39056

    'Chama a tela de Browse Nota Fiscal
    Call Chama_Tela("NFiscalDevLista", colSelecao, objNFiscal, objEventoNFiscal)

    Exit Sub

Erro_NFiscalLabel_Click:

    Select Case gErr

        Case 39056

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156328)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objNFiscal = obj1

    'Evita ter que alterar os browses a cada campo inserido na tabela de NF
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 180100

    'Traz a NotaFiscal de entrada para a a tela
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 39057

    'Fecha o Comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case gErr

        Case 39057, 180100

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156329)

    End Select

    Exit Sub

End Sub

Public Sub BotaoItensNFiscalOrig_Click()

Dim objItemNFiscalOriginal As ClassItemNF
Dim colSelecao As Collection

    'Chama a tela ItensNFiscalSaidaLista
    Call Chama_Tela("ItensNFiscalEntNaoDevLista", colSelecao, objItemNFiscalOriginal, objEventoItensNFiscalOrig)

End Sub

Private Sub objEventoItensNFiscalOrig_evSelecao(obj1 As Object)

Dim lErro As Long
Dim sProduto As String
Dim objItemNFiscalOriginal As ClassItemNF
Dim objNFiscalOrig As New ClassNFiscal
Dim objProduto As New ClassProduto, objKitVenda As New ClassKitVenda

On Error GoTo Erro_objEventoItensNFiscalOrig_evSelecao

    Set objItemNFiscalOriginal = obj1

    'Se nenhuma linha está selecionada , sai.
    If GridItens.Row = 0 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objItemNFiscalOriginal.sProduto, sProduto)
    If lErro <> SUCESSO Then gError 55140

    If lErro = SUCESSO And Len(Trim(sProduto)) > 0 Then
        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True

        lErro = Produto_Saida_Celula()
        If lErro <> SUCESSO Then gError 39058
        'Coloca os dados do item da Nf original no grid
        GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objItemNFiscalOriginal.sDescricaoItem
        GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objItemNFiscalOriginal.sUnidadeMed
        GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(objItemNFiscalOriginal.dQuantidade)
'        GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(objItemNFiscalOriginal.dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(objItemNFiscalOriginal.dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
        GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(objItemNFiscalOriginal.dPercDesc, "Percent")
        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = Format(objItemNFiscalOriginal.dValorDesconto, "Standard")
        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(objItemNFiscalOriginal.dValorTotal, "Standard")

        objNFiscalOrig.lNumIntDoc = objItemNFiscalOriginal.lNumIntNF

        lErro = CF("NFiscal_Le", objNFiscalOrig)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 199320

        If lErro = SUCESSO Then

            GridItens.TextMatrix(GridItens.Row, iGrid_SerieNFOrig_Col) = objNFiscalOrig.sSerie
            GridItens.TextMatrix(GridItens.Row, iGrid_CodNFOrig_Col) = objNFiscalOrig.lNumNotaFiscal
            GridItens.TextMatrix(GridItens.Row, iGrid_ItemNFOrig_Col) = objItemNFiscalOriginal.iItem

        End If

        lErro = CF2(Me, "NFDEV_Preenche_GridItens", objItemNFiscalOriginal, GridItens.Row)
        If lErro <> SUCESSO Then gError 181281

        Call PrecoTotal_Calcula(GridItens.Row, False)

        'Ajusta quantidade e preço no obj
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Call ValorTotal_Calcula
               
        objProduto.sCodigo = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
        If gobjNFiscal.ColItensNF(GridItens.Row).iPossuiGrade <> MARCADO And objProduto.iKitVendaComp <> MARCADO Then
            
            lErro = gobjDistribuicao.Distribuicao_Processa()
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = gobjRastreamento.Atualiza_Troca_Quantidade(GridItens.Row)
            
        Else
        
            If objProduto.iKitVendaComp = MARCADO Then
            
                objKitVenda.sProduto = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
                objKitVenda.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
                objKitVenda.dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
            
                lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            End If
    
            lErro = gobjDistribuicao.Distribuicao_Processa_Grade()
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            Call gobjDistribuicao.Atualiza_Grid_Distribuicao(gobjNFiscal.ColItensNF.Item(GridItens.Row))
            
            lErro = gobjRastreamento.Atualiza_Troca_Quantidade_Grade(GridItens.Row, 1)
        End If
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoItensNFiscalOrig_evSelecao:

    Select Case gErr

        Case 39058, 181281, 199320

        Case 55140
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNFiscalOriginal.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156330)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLocalizacao_Click()

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim colSelecao As New Collection
Dim objEstoqueProduto As New ClassEstoqueProduto

On Error GoTo Erro_BotaoLocalizacao_Click:

    'Verifica se tem alguma linha do Grid Selecionada
    If GridItens.Row = 0 Then gError 43690

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 39060

    If iProdutoPreenchido <> PRODUTO_VAZIO Then
        
        If gobjNFiscal.ColItensNF(GridItens.Row).iPossuiGrade = MARCADO Then
        
            Call BotaoGrade_Click
            Exit Sub
        End If
        
        colSelecao.Add sProdutoFormatado

        'Chama a Tela de Browse EstoqueProdutoTerceirosLista
        Call Chama_Tela("EstProdTercLista", colSelecao, objEstoqueProduto, objEventoLocalizacao)

    Else
        gError 43727

    End If

    Exit Sub

Erro_BotaoLocalizacao_Click:

    Select Case gErr

        Case 39060

        Case 43690
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 43727
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156331)

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
        If GridItens.Row = 0 Then gError 43689

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    End If

    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 42859

    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoLista_Consulta
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 42859

        Case 43689
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156332)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoEnxuto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    'verifica se tem alguma linha do Grid selecionada
    If GridItens.Row = 0 Then gError 39061

    'Verifica se o Produto está preenchido
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then

        Set objProduto = obj1

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 39062

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
        
        If Not (Me.ActiveControl Is Produto) Then
        
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
    
            lErro = Produto_Saida_Celula(objProduto)
            If lErro <> SUCESSO Then gError 39063
    
        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    Select Case gErr

        Case 39061, 39063

        Case 39062
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156333)

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

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim colTabelaPrecoItem As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "NFiscalDev"

    'Lê os dados da Tela
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 39064

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objNFiscal.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "TipoNFiscal", objNFiscal.iTipoNFiscal, 0, "TipoNFiscal"
    colCampoValor.Add "NaturezaOp", objNFiscal.sNaturezaOp, STRING_BUFFER_MAX_TEXTO, "NaturezaOp"
    colCampoValor.Add "Fornecedor", objNFiscal.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "FilialForn", objNFiscal.iFilialForn, 0, "FilialForn"
    colCampoValor.Add "Cliente", objNFiscal.lCliente, 0, "Cliente"
    colCampoValor.Add "FilialCli", objNFiscal.iFilialCli, 0, "FilialCli"
    colCampoValor.Add "Serie", objNFiscal.sSerie, STRING_BUFFER_MAX_TEXTO, "Serie"
    colCampoValor.Add "NumNotaFiscal", objNFiscal.lNumNotaFiscal, 0, "NumNotaFiscal"
    colCampoValor.Add "DataEmissao", objNFiscal.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "DataSaida", objNFiscal.dtDataSaida, 0, "DataSaida"
'horasaida
    colCampoValor.Add "HoraSaida", CDbl(objNFiscal.dtHoraSaida), 0, "HoraSaida"
    colCampoValor.Add "ValorFrete", objNFiscal.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorSeguro", objNFiscal.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "ValorOutrasDespesas", objNFiscal.dValorOutrasDespesas, 0, "ValorOutrasDespesas"
    colCampoValor.Add "ValorDesconto", objNFiscal.dValorDesconto, 0, "ValorDesconto"
    colCampoValor.Add "ValorTotal", objNFiscal.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "CodTransportadora", objNFiscal.iCodTransportadora, 0, "CodTransportadora"
    colCampoValor.Add "CodTranspRedesp", objNFiscal.iCodTranspRedesp, 0, "CodTranspRedesp"
    colCampoValor.Add "DetPagFrete", objNFiscal.iDetPagFrete, 0, "DetPagFrete"
    colCampoValor.Add "Placa", objNFiscal.sPlaca, STRING_BUFFER_MAX_TEXTO, "Placa"
    colCampoValor.Add "PlacaUF", objNFiscal.sPlacaUF, STRING_BUFFER_MAX_TEXTO, "PlacaUF"
    colCampoValor.Add "VolumeQuant", objNFiscal.lVolumeQuant, 0, "VolumeQuant"
    colCampoValor.Add "VolumeEspecie", objNFiscal.lVolumeEspecie, 0, "VolumeEspecie" 'Alterado por Luiz Nogueira em 21/08/03
    colCampoValor.Add "VolumeMarca", objNFiscal.lVolumeMarca, 0, "VolumeMarca" 'Alterado por Luiz Nogueira em 21/08/03
    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_BUFFER_MAX_TEXTO, "VolumeNumero"
    colCampoValor.Add "FreteRespons", objNFiscal.iFreteRespons, 0, "FreteRespons"
    colCampoValor.Add "MensagemNota", objNFiscal.sMensagemNota, STRING_BUFFER_MAX_TEXTO, "MensagemNota"
    colCampoValor.Add "PesoLiq", objNFiscal.dPesoLiq, 0, "Pesoliq"
    colCampoValor.Add "PesoBruto", objNFiscal.dPesoBruto, 0, "PesoBruto"
    colCampoValor.Add "NumIntNotaOriginal", objNFiscal.lNumIntNotaOriginal, 0, "NumIntNotaOriginal"
    colCampoValor.Add "ValorProdutos", objNFiscal.dValorProdutos, 0, "ValorProdutos"
    colCampoValor.Add "TabelaPreco", objNFiscal.iTabelaPreco, 0, "TabelaPreco"
    colCampoValor.Add "FilialEntrega", objNFiscal.iFilialEntrega, 0, "FilialEntrega"
    colCampoValor.Add "Status", objNFiscal.iStatus, 0, "Status"
    colCampoValor.Add "SemDataSaida", objNFiscal.iSemDataSaida, 0, "SemDataSaida"

    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
'    colSelecao.Add "Status", OP_DIFERENTE, STATUS_CANCELADO

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 39064

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156334)

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
        If lErro <> SUCESSO And lErro <> 31442 Then gError 180094

        'Coloca os dados da NFiscal na tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 39065

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 39065, 180094

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156335)

    End Select

    Exit Sub

End Sub

Public Function Move_NF_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os dados de NFiscal da tela para memória

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objCliente As New ClassCliente
Dim objTabelaPreco As New ClassTabelaPreco
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objProjetoInfo As Object

On Error GoTo Erro_Move_NF_Memoria

    If iFornecedorAtivo = FORN_ATIVO Then

        'Verifica se o Fornecedor está preenchido
        If Len(Trim(Fornecedor.ClipText)) > 0 Then

            objFornecedor.sNomeReduzido = Fornecedor.Text

            'Lê o Fornecedor
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 39066

            'Não encontrou o Fornecedor --> erro
            If lErro = 6681 Then gError 39067

            objNFiscal.lFornecedor = objFornecedor.lCodigo
            objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)

        End If

    Else

        If Len(Trim(Cliente.ClipText)) > 0 Then

            objCliente.sNomeReduzido = Cliente.Text

            'Lê o Cliente
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 39068

            'Se não encontrou o Cliente --> erro
            If lErro = 12348 Then gError 39069

            objNFiscal.lCliente = objCliente.lCodigo
            objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
        End If

    End If

    'Armazena o responsável pelo frete
    objNFiscal.iFreteRespons = TipoFrete.ItemData(TipoFrete.ListIndex)


    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 69121

    If lErro <> SUCESSO Then gError 69122

    If objTipoDocInfo.sTipoDocCPR = "" Then
        objNFiscal.iStatus = STATUS_BAIXADO
    Else
        objNFiscal.iStatus = STATUS_LANCADO
    End If

    objNFiscal.sVolumeNumero = VolumeNumero.Text
    objNFiscal.dPesoBruto = StrParaDbl(PesoBruto.Text)
    objNFiscal.dPesoLiq = StrParaDbl(PesoLiquido.Text)
    objNFiscal.dtDataEntrada = DATA_NULA
    objNFiscal.dtDataVencimento = DATA_NULA
    objNFiscal.dValorDesconto = StrParaDbl(ValorDesconto.Text)
    objNFiscal.dValorFrete = StrParaDbl(ValorFrete.Text)
    objNFiscal.dValorOutrasDespesas = StrParaDbl(ValorDespesas.Text)
    objNFiscal.dValorProdutos = StrParaDbl(ValorProdutos.Caption)
    objNFiscal.dValorSeguro = StrParaDbl(ValorSeguro.Text)
    objNFiscal.dValorTotal = StrParaDbl(ValorTotal.Caption)
    objNFiscal.iCanal = 0
    objNFiscal.iClasseDocCPR = 0
    objNFiscal.iCodTransportadora = Codigo_Extrai(Transportadora.Text)
    objNFiscal.iCodTranspRedesp = Codigo_Extrai(TranspRedespacho.Text)
    objNFiscal.iDetPagFrete = RedespachoCli.Value
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.iFilialEntrega = Codigo_Extrai(FilialEntrega.Text)
    objNFiscal.iTabelaPreco = Codigo_Extrai(TabelaPreco.Text)
    objNFiscal.lVolumeQuant = StrParaLong(VolumeQuant.Text)
    objNFiscal.lNumIntDoc = 0
    objNFiscal.lNumIntDocCPR = 0
    objNFiscal.lNumIntNotaOriginal = 0
    objNFiscal.lNumIntTrib = 0
    objNFiscal.lNumPedidoVenda = 0
    objNFiscal.sMensagemNota = Mensagem.Text
    objNFiscal.sMensagemCorpoNota = MensagemCorpo.Text
    objNFiscal.sNaturezaOp = gobjTribTab.sNatOpInterna
    objNFiscal.sNumPedidoTerc = ""
    objNFiscal.sPlaca = Placa.Text
    objNFiscal.sPlacaUF = PlacaUF.Text
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    
    'Incluído por Luiz Nogueira em 21/08/03
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objNFiscal.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objNFiscal.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
    
    objNFiscal.dtDataReferencia = DATA_NULA

    If Len(Trim(NFiscal.Caption)) > 0 Then
        objNFiscal.lNumNotaFiscal = CLng(NFiscal.Caption)
    Else
        objNFiscal.lNumNotaFiscal = 0
    End If

    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
   
    lErro = CF2(Me, "NFCF_Move_DataSaida_Memoria", objNFiscal)
    If lErro <> SUCESSO Then gError 132292

    If objNFiscal.iTabelaPreco <> 0 Then

        objTabelaPreco.iCodigo = objNFiscal.iTabelaPreco
        
        'Verifica se Tabela de Preço é uma informada é uma tabela de preço válida
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 43118

        'Se não achou a Tabela de Preço --> erro
        If lErro = 28004 Then gError 43119

    End If

    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181761
    
    Set objNFiscal.objProjetoInfo = objProjetoInfo
    Set objNFiscal.objInfoAdic = gobjInfoAdic
    objNFiscal.dValorItens = StrParaDbl(ValorProdutos2.Caption)
    objNFiscal.dValorDescontoItens = StrParaDbl(ValorDescontoItens.Text)
    
    Move_NF_Memoria = SUCESSO

    Exit Function

Erro_Move_NF_Memoria:

    Move_NF_Memoria = gErr

    Select Case gErr

        Case 39066, 39068, 43118, 69121, 132292, 181761

        Case 39067
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case 39069
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)

        Case 43119
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELA_PRECO_NAO_CADASTRADA", gErr, objTabelaPreco.iCodigo)

        Case 69122
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO2", gErr, objNFiscal.iTipoNFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156336)

    End Select

    Exit Function

End Function

Public Function Traz_NFiscal_Tela(objNFiscal As ClassNFiscal) As Long
'Mostra os dados da NFiscal na tela

Dim lErro As Long
Dim objItemNFiscal As New ClassItemNF
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemNFiscalOriginal As New ClassItemNF
Dim objNFiscalOriginal As New ClassNFiscal
Dim bAlterouCT As Boolean
Dim iIndice As Integer
Dim dPercDesc As Double

On Error GoTo Erro_Traz_NFiscal_Tela

'    'Se for versao LIGHT testa chaveamento
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        lErro = objUserControl.CT_Chavea_Trazendo(objNFiscal, bAlterouCT)
'        If lErro <> SUCESSO Then gError 25830
'
'        If bAlterouCT Then
'
'            'Chama o Traz_NFiscal_Tela de objCT chaveado
'            lErro = objUserControl.objCT.Traz_NFiscal_Tela(objNFiscal)
'            If lErro <> SUCESSO Then gError 25831
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

    gbCarregandoTela = True

    lErro = Limpa_Tela_NFiscal()
    If lErro <> SUCESSO Then gError 500124
    
    Set gobjNFiscal = objNFiscal

    'Lê as mensagens
    lErro = CF("NFiscal_Le_Msg", objNFiscal)
    If lErro <> SUCESSO Then gError 39070

    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 39070

    lErro = CF("NFiscalDev_Le_Dados_Benef", objNFiscal)
    If lErro <> SUCESSO Then gError 39070

    'distribuicao
    'Lê a Distribuição dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 39070
    
    'Lê a parte de Tributação
    lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
    If lErro <> SUCESSO Then gError 39311

    'Lê a parte de tributacao dos Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
    If lErro <> SUCESSO Then gError 39312

    'Para cada Ítem da Nota Fiscal
    For Each objItemNFiscal In objNFiscal.ColItensNF

        If objItemNFiscal.lNumIntDocOrig > 0 Then
        
            objItemNFiscalOriginal.lNumIntDoc = objItemNFiscal.lNumIntDocOrig
            'Lê o Ítem da NFOriginal em NFiscal
            lErro = CF("ItemNFiscal_Le", objItemNFiscalOriginal)
            If lErro <> SUCESSO And lErro <> 35225 Then gError 39071
            'Se não encontrou --> Erro
            If lErro <> SUCESSO Then gError 42860

            objItemNFiscal.iItemNFOrig = objItemNFiscalOriginal.iItem

            objNFiscalOriginal.lNumIntDoc = objItemNFiscalOriginal.lNumIntNF

            'Lê a Nota Fiscal do Ítem Lido
            lErro = CF("NFiscal_Le", objNFiscalOriginal)
            If lErro <> SUCESSO Then gError 39072

            'Preenche a Série e o Número com os lidos da Nota
            objItemNFiscal.sSerieNFOrig = objNFiscalOriginal.sSerie
            objItemNFiscal.lNumNFOrig = objNFiscalOriginal.lNumNotaFiscal

        End If
    Next

    'Preenche o Status da Nota Fiscal
    If objNFiscal.iStatus = STATUS_LANCADO Then
        Status.Caption = STRING_STATUS_LANCADO
    ElseIf objNFiscal.iStatus = STATUS_BAIXADO Then
        Status.Caption = STRING_STATUS_BAIXADO
    ElseIf objNFiscal.iStatus = STATUS_CANCELADO Then
        Status.Caption = STRING_STATUS_CANCELADO
    End If

    'Coloca o Tipo na tela
    TipoNFiscal.Text = objNFiscal.iTipoNFiscal
    Call TipoNFiscal_Validate(bSGECancelDummy)

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    'Lê o tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 39075
    If lErro = 31415 Then gError 39076

    'De acordo com a Origem do tipo Coloca o Cliente ou o fornecedor na tela
    If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Then
    
        Call Habilita_Cliente
        Call Cliente_Formata(objNFiscal.lCliente)
        Call Filial_Formata(Filial, objNFiscal.iFilialCli)
    
        lErro = CF2(Me, "NFCF_Cliente_Traz_NF_Tela", objNFiscal)
        If lErro <> SUCESSO Then gError 131881
    
    ElseIf objTipoDocInfo.iDestinatario = DOCINFO_FORNECEDOR Then
        Call Habilita_Fornecedor
        Fornecedor.Text = objNFiscal.lFornecedor
        Call Fornecedor_Validate(bSGECancelDummy)
        Filial.Text = objNFiscal.iFilialForn
        
        lErro = CF2(Me, "NFCF_Fornecedor_Traz_NF_Tela", objNFiscal)
        If lErro <> SUCESSO Then gError 131882
    
    End If

    Call Filial_Validate(bSGECancelDummy)

    'Se a nota fiscal original estiver preenchida
    If objNFiscal.lNumIntNotaOriginal > 0 Then

        Set objNFiscalOriginal = New ClassNFiscal
        objNFiscalOriginal.lNumIntDoc = objNFiscal.lNumIntNotaOriginal
        
        'Lâ Nota Fiscal Original
        lErro = CF("NFiscal_Le", objNFiscalOriginal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 39077
        If lErro <> SUCESSO Then gError 39079 'Não encontrou

        'Coloca na Tela o Número e a Série
        If objNFiscalOriginal.lNumNotaFiscal > 0 Then
            NFiscalOriginal.Text = objNFiscalOriginal.lNumNotaFiscal
            SerieNFiscalOriginal.Text = objNFiscalOriginal.sSerie
        End If

        objTipoDocInfo.iCodigo = objNFiscalOriginal.iTipoNFiscal

        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 42861
        If lErro <> SUCESSO Then gError 42862

        If objTipoDocInfo.iEmitente = DOCINFO_EMPRESA Then NFiscalPropria.Value = vbChecked

    End If

    'Preenche os Demais campos da Tela
    lErro = Preenche_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 39080

    If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
        'Carrega ItensNF com Rastreamentos
        lErro = gobjRastreamento.Carrega_RastroItensNF(objNFiscal)
        If lErro <> SUCESSO Then gError 83411
    End If

    NumNFe.Caption = objNFiscal.sStatusNFeFed

    Set gobjNFiscal = objNFiscal

    ''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 39031
    '*** fim tributacao

    'Traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 39665
    
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
    iAlterado = 0

    gbCarregandoTela = False
    
    MsgAutomatica.Value = vbUnchecked
    Set gobjInfoAdic = objNFiscal.objInfoAdic

    Traz_NFiscal_Tela = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela:

    gbCarregandoTela = False

    Traz_NFiscal_Tela = gErr

    Select Case gErr

        Case 25830, 25831, 39070, 39071, 39072, 39073, 39074, 39075, 39077, 39078, 39080, 39311, 39312, 42861, 83411, 500124

        Case 39076, 42862
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case 39079
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA", gErr, objNFiscalOriginal.lNumIntDoc)

        Case 39665, 131881, 131882

        Case 42860
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NFORIGINAL_NAO_CADASTRADO2", gErr, objItemNFiscalOriginal.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156337)

    End Select

    Exit Function

End Function

Private Sub Habilita_Cliente()
'Desabilita o Fornecededor e Habilita o Cliente

Dim bCancel As Boolean

'    If giTipoVersao = VERSAO_LIGHT Then
'        LabelFilialEntrega.Visible = False
'        Filial.Visible = False
'        LabelFilial.Visible = False
'        FilialEntrega.Visible = False
'    End If

    ClienteLabel.Visible = True
    Cliente.Visible = True
    FornecedorLabel.Visible = False
    Fornecedor.Visible = False
    iFornecedorAtivo = FORN_INATIVO

    'será que isto é usado p/forçar a carga da combo de filiais
    iClienteAlterado = REGISTRO_ALTERADO
    Call Cliente_Validate(bCancel)

    Call CF2(Me, "NFCF_Limpa_Tela")

End Sub

Private Sub Habilita_Fornecedor()
'Desabilita o Cliente e habilita o Fornecedor

'    If giTipoVersao = VERSAO_LIGHT Then
'        LabelFilialEntrega.Visible = True
'        Filial.Visible = True
'        LabelFilial.Visible = True
'        FilialEntrega.Visible = True
'    End If

    FornecedorLabel.Visible = True
    Fornecedor.Visible = True
    iFornecedorAtivo = FORN_ATIVO
    Cliente.Visible = False
    ClienteLabel.Visible = False

    'p/forcar a carga da combo de filiais
    iFornecedorAlterado = REGISTRO_ALTERADO
    Call Fornecedor_Validate(bSGECancelDummy)

    Call CF2(Me, "NFCF_Limpa_Tela")

End Sub

Private Function Preenche_Tela(objNFiscal As ClassNFiscal) As Long
'Preenche a tela com os dados de NFiscal Remessa de objNFiscal passado com parâmetro

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Tela

    'Preenche a tela com os dados vindos em objNFiscal
    gobjTribTab.sNatOpInterna = objNFiscal.sNaturezaOp
    Serie.Text = Desconverte_Serie_Eletronica(objNFiscal.sSerie)
    If ISSerieEletronica(objNFiscal.sSerie) Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    Call Serie_Validate(bSGECancelDummy)
    NFiscal.Caption = objNFiscal.lNumNotaFiscal

    DataEmissao.PromptInclude = False
    If objNFiscal.dtDataEmissao <> DATA_NULA Then DataEmissao.Text = Format(objNFiscal.dtDataEmissao, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    DataSaida.PromptInclude = False
    If objNFiscal.dtDataSaida <> DATA_NULA And objNFiscal.iSemDataSaida = 0 Then DataSaida.Text = Format(objNFiscal.dtDataSaida, "dd/mm/yy")
    DataSaida.PromptInclude = True

'horasaida
    HoraSaida.PromptInclude = False
    'este teste está correto
    If objNFiscal.dtDataSaida <> DATA_NULA Then HoraSaida.Text = Format(objNFiscal.dtHoraSaida, "hh:mm:ss")
    HoraSaida.PromptInclude = True

    'Valores
    ValorProdutos.Caption = Format(objNFiscal.dValorProdutos, "Standard")
    ValorFrete.Text = Format(objNFiscal.dValorFrete, "Fixed")
    ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Fixed")
    ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Fixed")
    ValorDesconto.Text = Format(objNFiscal.dValorDesconto, "Fixed")
    ValorTotal.Caption = Format(objNFiscal.dValorTotal, "Standard")
    
    lErro = CF2(Me, "NFCF_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132024
    
    'Transportadora
    If objNFiscal.iCodTransportadora <> 0 Then
        Transportadora.Text = objNFiscal.iCodTransportadora
        Call Transportadora_Validate(bSGECancelDummy)
    End If
    If objNFiscal.iCodTranspRedesp <> 0 Then
        TranspRedespacho.Text = CStr(objNFiscal.iCodTranspRedesp)
        Call TranspRedespacho_Validate(bSGECancelDummy)
    End If

    RedespachoCli.Value = objNFiscal.iDetPagFrete
    
    'Filial entrega
    If objNFiscal.iFilialEntrega > 0 Then
        FilialEntrega.Text = objNFiscal.iFilialEntrega
        Call FilialEntrega_Validate(bSGECancelDummy)
        giFilialEntregaAlterada = 0
    End If
    'Tabela de preço
    If objNFiscal.iTabelaPreco > 0 Then
        TabelaPreco.Text = objNFiscal.iTabelaPreco
        Call TabelaPreco_Validate(bSGECancelDummy)
    End If

    For iIndice = 0 To TipoFrete.ListCount - 1
        If TipoFrete.ItemData(iIndice) = objNFiscal.iFreteRespons Then
            TipoFrete.ListIndex = iIndice
            Exit For
        End If
    Next



    Placa.Text = objNFiscal.sPlaca
    PlacaUF.Text = objNFiscal.sPlacaUF
    If objNFiscal.lVolumeQuant > 0 Then VolumeQuant.Text = objNFiscal.lVolumeQuant
    VolumeEspecie.Text = objNFiscal.lVolumeEspecie
    
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
    
    Mensagem.Text = objNFiscal.sMensagemNota
    MensagemCorpo.Text = objNFiscal.sMensagemCorpoNota

    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = Format(objNFiscal.dPesoLiq, PesoLiquido.Format) 'Alterado por Wagner
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = Format(objNFiscal.dPesoBruto, PesoBruto.Format) 'Alterado por Wagner

    'Preenhce o Grid com os Ítens da Nota Fiscal em questão
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 39081

    'distribuicao
    'Preenche o Grid com as Distribuições dos itens da Nota Fiscal
    lErro = CF2(Me, "NFDEV_Preenche_GridDistribuicao", objNFiscal)
    If lErro <> SUCESSO Then gError 199179
    
    iAlterado = 0

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = gErr

    Select Case gErr

        Case 39081, 132024
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156338)

    End Select

    Exit Function

End Function

Public Function Preenche_GridItens(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com os ítens da Nota Fiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim sProdutoEnxuto As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim colCodigoNome As New AdmColCodigoNome
Dim iIndice1 As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Preenche_GridItens

    iIndice = 0

    'Para cada ítem da Coleção
    For Each objItemNF In objNFiscal.ColItensNF

        iIndice = iIndice + 1

        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 39082

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
'        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(objItemNF.dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(objItemNF.dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
        GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(objItemNF.dPercDesc, "Percent")
        GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objItemNF.dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objItemNF.dValorTotal, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col) = Format(objItemNF.dValorTotal + objItemNF.dValorDesconto, "Standard")
        If objItemNF.lNumNFOrig > 0 Then GridItens.TextMatrix(iIndice, iGrid_CodNFOrig_Col) = objItemNF.lNumNFOrig
        If Len(Trim(objItemNF.sSerieNFOrig)) > 0 Then GridItens.TextMatrix(iIndice, iGrid_SerieNFOrig_Col) = objItemNF.sSerieNFOrig
        If objItemNF.iItemNFOrig > 0 Then GridItens.TextMatrix(iIndice, iGrid_ItemNFOrig_Col) = objItemNF.iItemNFOrig

        lErro = CF2(Me, "NFDEV_Preenche_GridItens", objItemNF, iIndice)
        If lErro <> SUCESSO Then gError 181281

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

        Case 39082
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)

        Case 39083, 177926, 181281, 181676

        Case 39084
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objItemNF.iAlmoxarifado)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156339)

    End Select

    Exit Function

End Function

Private Function Preenche_Tela_NFOriginal(objNFiscal As ClassNFiscal, Optional iDestinatario As Integer) As Long
'Preenche a tela com alguns dados da Nota Fiscal Original passada em objNFiscal

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemNF As ClassItemNF
Dim iTipoNFiscal As Integer
Dim iIndice As Integer
Dim bCancel As Boolean, dPercDesc As Double

On Error GoTo Erro_Preenche_Tela_NFOriginal

    'Lê os Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 39085

    If objNFiscal.ColItensNF.Count = 0 Then

        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 39086

    End If

    'distribuicao
    'Lê a Distribuição dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 39086
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 39087

    If lErro = 31415 Then gError 39088

    iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscal_ComandoSeta()
    If lErro <> SUCESSO Then gError 500125

    Set gobjNFiscal = objNFiscal

    If iDestinatario = objTipoDocInfo.iOrigem Then Call CF("Filial_Seleciona", TipoNFiscal, iTipoNFiscal)

    'Verifica a Origem do Tipo de Nota Fiscal
    If objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
        Call Habilita_Cliente
        Cliente.Text = objNFiscal.lCliente
        Call Cliente_Validate(bCancel)
        Filial.Text = objNFiscal.iFilialCli
    ElseIf objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Or giComplemento = MARCADO Then
        Call Habilita_Fornecedor
        Fornecedor.Text = objNFiscal.lFornecedor
        Call Fornecedor_Validate(bSGECancelDummy)
        Filial.Text = objNFiscal.iFilialForn
    End If

    Call Filial_Validate(bSGECancelDummy)

    If objTipoDocInfo.iEmitente = DOCINFO_EMPRESA Then NFiscalPropria.Value = vbChecked

    SerieNFiscalOriginal.Text = objNFiscal.sSerie
    NFiscalOriginal.Text = objNFiscal.lNumNotaFiscal

    Call CF2(Me, "NFCF_Default_DataSaida")

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    Call DataEmissao_Validate(bSGECancelDummy)

    'Coloca nos Ítens a serie e o Numero da Nota Fiscal Original
    For Each objItemNF In objNFiscal.ColItensNF
        objItemNF.sSerieNFOrig = objNFiscal.sSerie
        objItemNF.lNumNFOrig = objNFiscal.lNumNotaFiscal
        objItemNF.iItemNFOrig = objItemNF.iItem
        objItemNF.lNumIntDocOrig = objItemNF.lNumIntDoc
    Next

    lErro = CF2(Me, "NFDEV_Preenche_GridDistribuicao", objNFiscal)
    If lErro <> SUCESSO Then gError 199174

    For iIndice = 1 To objNFiscal.ColItensNF.Count
        lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, objNFiscal.ColItensNF.Item(iIndice).sProduto)
        If lErro <> SUCESSO Then gError 27832
    Next

    ValorProdutos.Caption = Format(objNFiscal.dValorProdutos, "Standard")

    'Preenche o Grid Com os Ítens lidos
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 39089
    
    For iIndice = 1 To objNFiscal.ColItensNF.Count
        lErro = gobjTribTab.Alteracao_Item_Grid(iIndice, , objNFiscal.ColItensNF.Item(iIndice).lNumIntDocOrig)
        If lErro <> SUCESSO Then gError 27832
    Next

'    lErro = gobjTribTab.Traz_NFiscalOrig_Tela(objNFiscal)
'    If lErro <> SUCESSO Then gError 27832
    
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
    
    lErro = CF2(Me, "NFCF_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132027
    
    Call ValorTotal_Calcula

    lErro = CF2(Me, "NFDEV_Calcula_PesoKg_Total")
    If lErro <> SUCESSO Then gError 177053

    iAlterado = 0

    Preenche_Tela_NFOriginal = SUCESSO

    Exit Function

Erro_Preenche_Tela_NFOriginal:

    Preenche_Tela_NFOriginal = gErr

    Select Case gErr

        Case 39085, 39086, 39087, 39089, 500125, 177053, 199174
        
        Case 132027

        Case 39088
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156340)

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
        If lErro <> SUCESSO Then gError 64004

        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 39090
        If lErro <> SUCESSO Then gError 39091

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

        'Lê o Tipo de Documento com o Código de objNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 39092
        If lErro = 31415 Then gError 39093 'Se não encontrar --> Erro

'        If giTipoVersao = VERSAO_FULL Then
            If objTipoDocInfo.sNomeTelaNFiscal <> Me.Name Then gError 39094
'        End If

        'Traz a nota para a tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 39095

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 39090, 39095, 39092, 64004

        Case 39091
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case 39093
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 39094
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_DIFERENTE_NF_SAIDA_DEVOLUCAO", gErr, objNFiscal.iTipoNFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156341)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub BotaoExibirDados_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objTipoDocInfoOrig As New ClassTipoDocInfo

On Error GoTo Erro_BotaoExibirDados_Click

    'Verifica se a Serie e o Número da Nota Fiscal original estão preenchidos
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Or Len(Trim(NFiscalOriginal.ClipText)) = 0 Then gError 39102

    objNFiscal.sSerie = SerieNFiscalOriginal.Text
    objNFiscal.lNumNotaFiscal = NFiscalOriginal.Text
    objNFiscal.iFilialEmpresa = giFilialEmpresa

    If NFiscalPropria.Value = vbChecked Then

        'Tenta lêr a Nota com esses dados
        lErro = CF("NFiscalInternaEntNaoDev_Le_Numero", objNFiscal)
        
        If giComplemento = MARCADO Then
            lErro = CF("NFiscalInternaSaida_Le_Numero2", objNFiscal)
        Else
            lErro = CF("NFiscalInternaEntNaoDev_Le_Numero", objNFiscal)
        End If
        If lErro <> SUCESSO And lErro <> 39101 Then gError 39118

        If lErro = 39101 Then gError 39103 'Não encontrou

    Else
        If TipoNFiscal.ListIndex = -1 Then gError 39104

        objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)

        'Lê o tipo de documento
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 39105

        'Se não encontrou --> erro
        If lErro <> SUCESSO Then gError 39106

        If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Then
            If Len(Trim(Cliente.ClipText)) = 0 Then gError 39107
        Else
            If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 39108
        End If

        'Verifica se Filial do Fornecedor foi preenchida
        If Len(Trim(Filial.Text)) = 0 Then gError 39109

        'Move os dados de NFiscal para memória
        lErro = Move_NF_Memoria(objNFiscal)
        If lErro <> SUCESSO Then gError 39110

        objNFiscal.sSerie = SerieNFiscalOriginal.Text
        objNFiscal.lNumNotaFiscal = NFiscalOriginal.Text

        lErro = CF("NFiscalExternaNaoDev_Le_Numero", objNFiscal)
        If lErro <> SUCESSO And lErro <> 39116 Then gError 39111

        'Se não encontrou --> erro
        If lErro <> SUCESSO Then gError 39112

        objTipoDocInfoOrig.iCodigo = objNFiscal.iTipoNFiscal

        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfoOrig)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 62298
        If lErro <> SUCESSO Then gError 62299

        'Se for uma devolução simbólica ou a devolução real de material beneficiado, a nf original pode ser
        'uma remessa para beneficiamento emitida pelo cliente ou por conta e ordem do cliente
        If (objTipoDocInfo.iCodigo = DOCINFO_NFISDBF Or objTipoDocInfo.iCodigo = DOCINFO_NFISDBFNE) And (objTipoDocInfoOrig.iCodigo = DOCINFO_NFIE3BF Or objTipoDocInfoOrig.iCodigo = DOCINFO_NFEE3BF Or objTipoDocInfoOrig.iCodigo = DOCINFO_NFEERMB3 Or objTipoDocInfoOrig.iCodigo = DOCINFO_NFIERMB3) Then
            If objTipoDocInfoOrig.iCodigo = DOCINFO_NFIE3BF Then objNFiscal.iTipoNFiscal = DOCINFO_NFIERMB3 'Altera para a remessa direta do cliente
            If objTipoDocInfoOrig.iCodigo = DOCINFO_NFEE3BF Then objNFiscal.iTipoNFiscal = DOCINFO_NFEERMB3 'Altera para a remessa direta do cliente
        Else
            
            'para tratar casos de mais de um tipo de nf de devolucao para um tipo de entrada
            Select Case objTipoDocInfoOrig.iCodigo
        
                Case DOCINFO_NFEEC
                    If objTipoDocInfo.sSigla <> "NFISDSC" And objTipoDocInfo.sSigla <> "NFISDC" Then gError 62300
        
                Case Else
                    If objTipoDocInfo.sSigla <> objTipoDocInfoOrig.sSiglaNFOriginal Then gError 62300
                    
            End Select
                    
        End If

    End If

    'Coloca na tela os dados encontados
    lErro = Preenche_Tela_NFOriginal(objNFiscal, objTipoDocInfo.iDestinatario)
    If lErro <> SUCESSO Then gError 39117

    Exit Sub

Erro_BotaoExibirDados_Click:

    Select Case gErr

        Case 39102
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NUMERO_ORIGINAL_FALTANDO2", gErr)

        Case 39103
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_INTERNA_ENTRADA_NAO_CADASTRADA", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 39104
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 39106, 62299
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case 39107
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 39108
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 39109
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 39112
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_EXTERNA_NAO_CADASTRADA", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.lFornecedor, Codigo_Extrai(Filial.Text))

        Case 39105, 39110, 39111, 39117, 39118, 62298

        Case 62300
            If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Then
                Call Rotina_Erro(vbOKOnly, "ERRO_TIPONFISCALORIGINAL_DIFERENTE_TELA", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.lCliente, objNFiscal.iFilialCli, objTipoDocInfo.sSigla, objTipoDocInfoOrig.sSiglaNFOriginal)
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_TIPONFISCALORIGINAL_DIFERENTE_TELA", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objTipoDocInfo.sSigla, objTipoDocInfoOrig.sSiglaNFOriginal)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156342)

    End Select

    Exit Sub

End Sub

Private Function Trata_FilialCliente() As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim objTipoCliente As New ClassTipoCliente
Dim objTransportadora As New ClassTransportadora
Dim objMensagem As New ClassMensagem
Dim objTipoFrete As ClassTipoFrete
Dim colTipoFrete As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Trata_FilialCliente

    objFilialCliente.iCodFilial = Filial.ItemData(Filial.ListIndex)
    'Lê a FilialCliente
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilialCliente)
    If lErro <> SUCESSO Then gError 39128

    Call gobjTribTab.FilialCliente_Alterada(objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    Call Trata_FilialCli_Dev(objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

    'Se a Transportadora da Filial estiver preenchida
    If objFilialCliente.iCodTransportadora <> 0 Then
        Transportadora.Text = objFilialCliente.iCodTransportadora
    Else
        objCliente.sNomeReduzido = Cliente.Text
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO Then gError 39129

        If objCliente.iCodTransportadora <> 0 Then
            Transportadora.Text = objCliente.iCodTransportadora
        ElseIf objCliente.iTipo > 0 Then
            objTipoCliente.iCodigo = objCliente.iTipo

            lErro = CF("TipoCliente_Le", objTipoCliente)
            If lErro = SUCESSO Then
                If objTipoCliente.iCodTransportadora <> 0 Then Transportadora.Text = objTipoCliente.iCodTransportadora
            End If
        End If
    End If

    If Len(Trim(Transportadora.Text)) > 0 Then Transportadora_Validate (bSGECancelDummy)

    'Seleciona tipo de frete da Filial do Cliente
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207424
    
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
    If lErro <> SUCESSO Then gError 131883
    
'    If objFilialCliente.iCodMensagem > 0 Then
'        objMensagem.iCodigo = objFilialCliente.iCodMensagem
'        'Lê a Mensagem "default" p/ o Cliente
'        lErro = CF("Mensagem_Le", objMensagem)
'        If lErro <> SUCESSO And lErro <> 19234 Then gError 140376
'
'        If lErro = SUCESSO Then Mensagem.Text = objMensagem.sDescricao
'
'    End If

    Trata_FilialCliente = SUCESSO

    Exit Function

Erro_Trata_FilialCliente:

    Trata_FilialCliente = gErr

    Select Case gErr

        Case 39128, 39129, 124056, 131883, 140376, 207424

        Case 124057
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objTransportadora.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156343)

    End Select

    Exit Function

End Function

Private Function Trata_TabelaPreco() As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_TabelaPreco_Click

    If Not gbCarregandoTela Then

        For iLinha = 1 To objGridItens.iLinhasExistentes

            'Verifica se o Produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProduto, iPreenchido)
            If lErro <> SUCESSO Then gError 39147

            If iPreenchido <> PRODUTO_VAZIO And Len(Trim(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col))) > 0 Then
                objTabelaPrecoItem.sCodProduto = sProduto
                objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
                objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
                'Lê a Tabela preço para filialEmpresa
                lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                If lErro <> SUCESSO And lErro <> 28014 Then gError 39148
                'Se não encontrar
                If lErro = 28014 Then
                    objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA
                    'Lê a Tabela de Preço a nível de Empresa toda
                    lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
                    If lErro <> SUCESSO And lErro <> 28014 Then gError 39149

                End If
                'Se  conseguir ler a Tabela de Preços
                If lErro = SUCESSO Then
                    'Calcula o Preco Unitário do Ítem
                    lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
                    If lErro <> SUCESSO Then gError 39150
                    'Coloca no Grid
                    If dPrecoUnitario > 0 Then
                        GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
                    Else
                        GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col) = ""
                    End If
                    'Calcula o Preco Total do Ítem
                    Call PrecoTotal_Calcula(iLinha, False)

                    ''*** incluidos p/tratamento de tributacao *******************************
                    lErro = gobjTribTab.Alteracao_Item_Grid(iLinha)
                    If lErro <> SUCESSO Then gError 56858
                    '*** fim tributacao

                End If

            End If

        Next

    End If

    'Calcula o Valor Total da Nota
    Call ValorTotal_Calcula

    Trata_TabelaPreco = SUCESSO

    Exit Function

Erro_TabelaPreco_Click:

    Trata_TabelaPreco = gErr

    Select Case gErr

        Case 39147, 39148, 39149, 39150, 56858

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156344)

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

    lErro = CF2(Me, "NFDEV_Calcula_Valores", dQuantidade, iLinha)
    If lErro <> SUCESSO Then gError 177058

    'Recolhe os valores Quantidade, Desconto, PerDesc e Valor Unitário da tela
    If dPrecoUnitario = 0 Or dQuantidade = 0 Then
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = ""
        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(0, "Standard")
    Else
        dPrecoTotal = dPrecoUnitario * dQuantidade
        dDesconto = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col))

        lTamanho = Len(Trim(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col)))
        If lTamanho > 0 Then
            dPercentDesc = PercentParaDbl(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col))
        Else
            dPercentDesc = 0
        End If

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

    Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)

'    'Obtem o valor do Desconto global
'    If Len(Trim(ValorDesconto.Text)) > 0 And IsNumeric(ValorDesconto.Text) Then dValorDescontoGlobal = CDbl(ValorDesconto.Text)
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
'        If lErro <> SUCESSO Then gError 56891
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
    If lErro <> SUCESSO Then gError 132028

    If iCalculaTotal <> False Then Call ValorTotal_Calcula

    Exit Sub

Erro_PrecoTotal_Calcula:

    Select Case gErr

        Case 56891
        
        Case 132028, 177058 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156345)

    End Select

    Exit Sub

End Sub

Public Sub ValorTotal_Calcula()

Dim dValorTotal As Double
Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    lErro = gobjTribTab.ValorTotal_Calcula(dValorTotal)
    If lErro <> SUCESSO Then gError 39305

    ValorTotal.Caption = Format(dValorTotal, "Standard")

    Call ValorDescontoItens_Calcula

    Exit Sub

Erro_ValorTotal_Calcula:

    Select Case gErr

        Case 39305 ', 56892

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156346)

    End Select

    Exit Sub

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim lErro As Long
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
            If lErro <> SUCESSO Then gError 42863

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
            If lErro <> SUCESSO Then gError 39157

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                UnidadeMed.Enabled = False
            Else
                UnidadeMed.Enabled = True

                objProduto.sCodigo = sProdutoFormatado
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 39158
                If lErro = 28030 Then gError 39159 'Não achou

                objClasseUM.iClasse = objProduto.iClasseUM
                'Lâ as Unidades de Medidas da Classe do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 39160
                'Carrega a combo de UM
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next
                'Seleciona na UM que está preenchida
                UnidadeMed.Text = sUM
                If Len(Trim(sUM)) > 0 Then
                    lErro = Combo_Item_Igual(UnidadeMed)
                    If lErro <> SUCESSO And lErro <> 12253 Then gError 39161
                End If
            End If

        'Nas demais
        Case PrecoUnitario.Name, PercentDesc.Name, Desconto.Name, SerieNFOriginal.Name, CodigoNFOriginal.Name, ItemNFOriginal.Name, DescricaoItem.Name
            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 39162

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case Quantidade.Name 'm, Almoxarifado.Name

            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 46713

            If iProdutoPreenchido = PRODUTO_VAZIO Or left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

    End Select
    
    lErro = CF2(Me, "NFDEV_Rotina_Grid_Enable", iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 177036
    
    'Rastreamento
    lErro = gobjRastreamento.Rotina_Grid_Enable(iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 83412

    'distribuicao
    lErro = gobjDistribuicao.Rotina_Grid_Enable_Dist(iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 83412
    
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 39158, 39160, 39157, 39162, 39161, 42863, 46713, 83412

        Case 39159
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156347)

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

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer
Dim iIndice As Integer
Dim dValorProdutos As Double
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_GridItens_KeyDown

    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes
    iItemAtual = GridItens.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

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
        
        If objGridItens.iLinhasExistentes <> 0 Then
            Call PrecoTotal_Calcula(objGridItens.iLinhasExistentes, False)
        Else
            If StrParaDbl(ValorDesconto.Text) <> 0 Then
                'Avisa ao usuário
                vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", ValorDesconto.Text, 0)
                
                'Limpa o valor de desconto
                ValorDesconto.Text = ""
                iValorDescontoAlterado = 0
            
                Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, 0)
            End If
        End If
        
        Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotal_Col, dValorProdutos)

        ValorProdutos.Caption = Format(dValorProdutos - StrParaDbl(ValorDesconto.Text), "Standard")

        Call SubTotal_Calcula(objGridItens, iGrid_PrecoTotalB_Col, dValorProdutos)

        ValorProdutos2.Caption = Format(dValorProdutos, "Standard")

        lErro = CF2(Me, "NFCF_ValorDesconto_Validate")
        If lErro <> SUCESSO Then gError 132025
        
        ''*** incluidos p/tratamento de tributacao *******************************
        lErro = gobjTribTab.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 39036
        '*** fim tributacao

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81518

        lErro = CF2(Me, "NFDEV_Calcula_PesoKg_Total")
        If lErro <> SUCESSO Then gError 177051

        Call ValorTotal_Calcula

        lErro = gobjDistribuicao.Exclusao_Item_GridDist(iItemAtual)
        If lErro <> SUCESSO Then gError 83413

        For iIndice = 1 To gobjNFiscal.ColItensNF.Count
            gobjNFiscal.ColItensNF(iIndice).iItem = iIndice 'Incluído para ao alterar a grade ou a reserva não fazer do item errado
        Next
        
    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr

        Case 39036, 81518, 83413
        
        Case 132025

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156348)

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

    lErro = CF2(Me, "NFDEV_Produto_Saida_Celula")
    If lErro <> SUCESSO Then gError 26659

'    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
'    If lErro <> SUCESSO Then gError 39021
'
'    Produto.PromptInclude = False
'    Produto.Text = sProdutoEnxuto
'    Produto.PromptInclude = True


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
    If iFornecedorAtivo = FORN_INATIVO Then
        GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMVenda
    Else
        GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMCompra
    End If
    
    
    'Descricao
    GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objProduto.sDescricao

    'Preco Unitário
    lErro = TabelaPreco_Atualiza_PrecoUnitario(objProduto)
    If lErro <> SUCESSO Then gError 35786
    
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
        gobjNFiscal.ColItensNF(GridItens.Row).sUnidadeMed = objProduto.sSiglaUMVenda
    
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
        
            Call gobjDistribuicao.Atualiza_Grid_Distribuicao(objItemNF)
        
        End If
    
'*** fim tributacao

    End If

    lErro = CF2(Me, "NFDEV_ProdutoLinha_Preenche", objProduto)
    If lErro <> SUCESSO Then gError 177054

    Call PrecoTotal_Calcula(GridItens.Row, False)
    Call ValorTotal_Calcula
    
    Produto_Saida_Celula = SUCESSO

    Exit Function

Erro_Produto_Saida_Celula:

    Produto_Saida_Celula = gErr

    Select Case gErr

        Case 26637, 26657, 27832, 35786, 51420, 86310, 86327, 86329, 101782, 177582, 177054
        
        Case 26658
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)
                
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
            End If
            
        Case 26659
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE", gErr, Produto.Text, iIndice)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156349)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 39666

        If objGridInt.objGrid.Name = GridItens.Name Then

            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                'Produto
                Case iGrid_Produto_Col
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 39171
                Case iGrid_DescProduto_Col
                    lErro = Saida_Celula_DescricaoItem(objGridInt)
                    If lErro <> SUCESSO Then gError 42873

                'Unidade Medida
                Case iGrid_UnidadeMed_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError 39172
                    
                'Quantidade
                Case iGrid_Quantidade_Col
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 39173

                'ValorUnitario
                Case iGrid_ValorUnitario_Col
                    lErro = Saida_Celula_PrecoUnitario(objGridInt)
                    If lErro <> SUCESSO Then gError 39175

                'PercDesc
                Case iGrid_PercDesc_Col
                    lErro = Saida_Celula_PercentDesc(objGridInt)
                    If lErro <> SUCESSO Then gError 39176
                    
                'Desconto
                Case iGrid_Desconto_Col
                    lErro = Saida_Celula_Desconto(objGridInt)
                    If lErro <> SUCESSO Then gError 39177

                Case iGrid_CodNFOrig_Col
                    lErro = Saida_Celula_CodigoNFOriginal(objGridInt)
                    If lErro <> SUCESSO Then gError 39298

                Case iGrid_SerieNFOrig_Col
                    lErro = Saida_Celula_SerieNFOriginal(objGridInt)
                    If lErro <> SUCESSO Then gError 39299

                Case iGrid_ItemNFOrig_Col
                    lErro = Saida_Celula_ItemNFOrig(objGridInt)
                    If lErro <> SUCESSO Then gError 39302

            End Select

        'Rastreamento
        ElseIf objGridInt.objGrid.Name = GridRastro.Name Then

            lErro = gobjRastreamento.Saida_Celula()
            If lErro <> SUCESSO Then gError 83414

        'distribuicao
        ElseIf objGridInt.objGrid.Name = GridDist.Name Then

            lErro = gobjDistribuicao.Saida_Celula_Dist()
            If lErro <> SUCESSO Then gError 89561

        End If
        
        lErro = CF2(Me, "NFDEV_Saida_Celula", objGridInt)
        If lErro <> SUCESSO Then gError 177035

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 39178

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 39178, 89561, 177035
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 39171 To 39177, 39298, 39302, 39666, 42873, 83414
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156350)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

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
        If lErro <> SUCESSO Then gError 199358
    
        Produto.Text = sProduto
      
        'Executa a saida de célula
        lErro = Produto_Saida_Celula()
        If lErro <> SUCESSO Then gError 39180
    End If

    'Necessário para o funcionamento da Rotina_Grid_Enable
    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 39181

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 39179
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridItens)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridItens)
            End If

        Case 39180, 39181, 39492, 199358
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156352)

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
Dim dQuantidadeAtual As Double
Dim objProduto As New ClassProduto
Dim objKitVenda As New ClassKitVenda
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim dFator As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    'distribuicao
    dQuantidadeAnterior = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    'fim  distribuicao
    
    'verifica se a quantidade está preenchida
    If Len(Trim(Quantidade.ClipText)) > 0 Then
        'Critica o Valor digitado
        If giComplemento = DESMARCADO Then
            lErro = Valor_Positivo_Critica(Quantidade.Text)
        Else
            lErro = Valor_NaoNegativo_Critica(Quantidade.Text)
        End If
        If lErro <> SUCESSO Then gError 39182

        dQuantidade = CDbl(Quantidade.Text)
        Quantidade = Formata_Estoque(dQuantidade)

        'distribuicao
        dQuantidadeAtual = StrParaDbl(Quantidade.Text)
        'fim  distribuicao
    
    End If

    If dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col)) Then iSairCelula = True
    
    lErro = CF2(Me, "NFDev_AlocaAlmoxarifadoPadrao", gobjDistribuicao)
    If lErro <> SUCESSO Then gError 141077
    
    'inicio distribuicao
    If dQuantidadeAnterior <> dQuantidadeAtual Then
        
        objProduto.sCodigo = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 180040
        
        GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)
            
        If gobjNFiscal.ColItensNF(GridItens.Row).iPossuiGrade <> MARCADO And objProduto.iKitVendaComp <> MARCADO Then
            
            'altera as quantidades no tab de distribuicao
            lErro = gobjDistribuicao.Distribuicao_Processa()
            If lErro <> SUCESSO Then gError 89562
        
        Else
        
            If objProduto.iKitVendaComp = MARCADO Then
            
                objKitVenda.sProduto = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
                objKitVenda.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
                objKitVenda.dQuantidade = dQuantidade
            
                lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 180041
                
            End If

            lErro = gobjDistribuicao.Distribuicao_Processa_Grade()
            If lErro <> SUCESSO Then gError 89602
        
            Call gobjDistribuicao.Atualiza_Grid_Distribuicao(gobjNFiscal.ColItensNF.Item(GridItens.Row))
        
        End If
    End If
    'fim distribuicao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 39183

    If Not iSairCelula Then

        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81519

        lErro = CF2(Me, "NFDEV_Calcula_PesoKg", GridItens.Row)
        If lErro <> SUCESSO Then gError 177055

        dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))

        If dPrecoUnitario > 0 Then

            Call PrecoTotal_Calcula(GridItens.Row, False)

            '*** incluidos p/tratamento de tributacao ****************
            If GridItens.Row <= objGridItens.iLinhasExistentes Then
                lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
                If lErro <> SUCESSO Then gError 39037
            End If
            '*** fim tributacao

            Call ValorTotal_Calcula

            If GridItens.Row <= objGridItens.iLinhasExistentes Then
            
                If gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade.Count = 0 Then
            
                    lErro = gobjRastreamento.Atualiza_Troca_Quantidade(GridItens.Row)
                    If lErro <> SUCESSO Then gError 83847
                
                Else
                
                    If dQuantidadeAnterior <> 0 Then
                        dFator = dQuantidadeAtual \ dQuantidadeAnterior
                    Else
                        dFator = 1
                    End If
                    
                    lErro = gobjRastreamento.Atualiza_Troca_Quantidade_Grade(GridItens.Row, dFator)
                    If lErro <> SUCESSO Then gError 83847
                
                End If
                
            End If

        End If
        
        gobjNFiscal.ColItensNF(GridItens.Row).dQuantidade = dQuantidadeAtual
        
        If gobjFAT.iTrataNFDevSimbPorOP = MARCADO Then
            lErro = CF("NFDevBenef_Ajusta_Qtde", gobjNFiscal.ColItensNF(GridItens.Row).sUnidadeMed, dQuantidadeAnterior, gobjNFiscal.ColItensNF(GridItens.Row))
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If

    End If

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 39037, 39182, 39183, 81519, 83847, 141077, 89562, 180040, 180041, 177055, ERRO_SEM_MENSAGEM 'Inserido por Wagner
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156353)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrecoUnitario(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double
Dim dQuantCancelada As Double
Dim iPrecoUnitarioAlterado As Integer
Dim sValorUnitario As String

On Error GoTo Erro_Saida_Celula_PrecoUnitario

    Set objGridInt.objControle = PrecoUnitario

    iPrecoUnitarioAlterado = False

    'Verifica se tem alguma preço digitado
    If Len(Trim(PrecoUnitario.ClipText)) > 0 Then

        If giComplemento = MARCADO Then
            lErro = Valor_NaoNegativo_Critica(PrecoUnitario.Text)
        Else
            lErro = Valor_Positivo_Critica(PrecoUnitario.Text)
        End If
        If lErro <> SUCESSO Then gError 39190

'        PrecoUnitario.Text = Format(PrecoUnitario.Text, gobjFAT.sFormatoPrecoUnitario)
        PrecoUnitario.Text = Format(PrecoUnitario.Text, FORMATO_PRECO_UNITARIO_EXTERNO)

    End If

    If PrecoUnitario.Text <> GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) Then iPrecoUnitarioAlterado = True

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 39191

    If iPrecoUnitarioAlterado = True Then

        Call PrecoTotal_Calcula(GridItens.Row, False)

    ''*** incluidos p/tratamento de tributacao *******************************
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 39038
    '*** fim tributacao

        Call ValorTotal_Calcula

    End If

    Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = gErr

    Select Case gErr

        Case 39190, 39191, 39038
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156355)

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

Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_PercentDesc

    Set objGridInt.objControle = PercentDesc

    dPercentDesc = StrParaDbl(PercentDesc.Text)
    
    lTamanho = Len(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col))
    If lTamanho > 0 Then dPercentDescAnterior = StrParaDbl(left(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col), lTamanho - 1))

    'verifica se o percentual está preenchido
    If Len(Trim(PercentDesc.Text)) > 0 Then
        'Critica a procentagem
        lErro = Porcentagem_Critica(PercentDesc.Text)
        If lErro <> SUCESSO Then gError 39192

        If dPercentDesc <> dPercentDescAnterior Then

            'Verifica se o percentual é de 100%
            If dPercentDesc = 100 Then gError 39193

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
    If lErro <> SUCESSO Then gError 39194

    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))

    If dPrecoUnitario > 0 And dQuantidade > 0 Then

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

        Case 39192, 39194, 39039
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 39193
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156356)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidade de Medida que está deixando de ser a corrente

Dim lErro As Long
Dim sUmAnterior As String
Dim objProduto As New ClassProduto
Dim objKitVenda As New ClassKitVenda

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

'inicio distribuicao
    'recolhe a UM anteriormente escolhida
    sUmAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
    
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text
    
    'coloca no grid a UM atual selecionda
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    gobjNFiscal.ColItensNF(GridItens.Row).sUnidadeMed = UnidadeMed.Text
    
    lErro = CF2(Me, "NFDev_AlocaAlmoxarifadoPadrao", gobjDistribuicao)
    If lErro <> SUCESSO Then gError 141078
    
    If sUmAnterior <> UnidadeMed.Text Then
        
        'Recalcula o peso bruto e o peso líquido
        lErro = Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 89511
                
        objProduto.sCodigo = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177583
        
        If gobjNFiscal.ColItensNF(GridItens.Row).iPossuiGrade <> MARCADO And objProduto.iKitVendaComp <> MARCADO Then
        
            'Tenta fazer a distribuição automatica p\ o item de acordo com a nova UM
            lErro = gobjDistribuicao.Distribuicao_Processa()
            If lErro <> SUCESSO Then gError 89511
            
        Else

            If objProduto.iKitVendaComp = MARCADO Then
    
                objKitVenda.sProduto = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
                objKitVenda.sUM = UnidadeMed.Text
                objKitVenda.dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
            
                lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 177584
            
            End If

            lErro = gobjDistribuicao.Distribuicao_Processa_Grade()
            If lErro <> SUCESSO Then gError 89602
        
            Call gobjDistribuicao.Atualiza_Grid_Distribuicao(gobjNFiscal.ColItensNF.Item(GridItens.Row))
        End If
    End If
    
'fim distribuicao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 35083

    If GridItens.Row <= objGridItens.iLinhasExistentes Then
        
        If gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade.Count = 0 Then
            lErro = gobjRastreamento.Atualiza_Troca_UM(GridItens.Row)
            If lErro <> SUCESSO Then gError 83206
        End If
    End If
    
    lErro = CF2(Me, "NFDEV_Calcula_PesoKg", GridItens.Row)
    If lErro <> SUCESSO Then gError 177057
    
    If gobjFAT.iTrataNFDevSimbPorOP = MARCADO And sUmAnterior <> UnidadeMed.Text Then
        lErro = CF("NFDevBenef_Ajusta_Qtde", sUmAnterior, gobjNFiscal.ColItensNF(GridItens.Row).dQuantidade, gobjNFiscal.ColItensNF(GridItens.Row))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 35083, 83206, 89511, 89602, 141078, 177057, ERRO_SEM_MENSAGEM 'Inserido por Wagner
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156357)

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
            If lErro <> SUCESSO Then gError 500087

            'Lê o produto
            objProduto.sCodigo = sProdutoFormatado
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 500088

            'Se o produto não está cadastrado, erro
            If lErro = 28030 Then gError 500090

            'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then gError 500089

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

        Case 500087, 500088, 500089

        Case 500090
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156358)

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

    If Trim(Desconto.Text) <> GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) Then

        'Verifica se o o desconto foi digitado
        If Len(Trim(Desconto.ClipText)) > 0 Then

            'Critica o valor digitado
            lErro = Valor_NaoNegativo_Critica(Desconto.Text)
            If lErro <> SUCESSO Then gError 39195

            dDesconto = CDbl(Desconto.Text)

        End If
        'Verifica se Desconto foi alterado
        If dDesconto <> StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col)) Then
            iDescontoAlterado = REGISTRO_ALTERADO
            dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
            dPrecoUnitario = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))
            dPrecoTotal = dQuantidade * dPrecoUnitario
            'Se o Preço total e positivo
            If dPrecoTotal > 0 Then

                'Verifica se o Valor do desconto é superior ao Preço Total
                If dDesconto >= dPrecoTotal Then gError 39196

                'Recalcula o percentual de desconto
                dPercentDesc = dDesconto / dPrecoTotal

                GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")
            End If
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 39197

    Call PrecoTotal_Calcula(GridItens.Row, False)

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 39040
'*** fim tributacao

    Call ValorTotal_Calcula

    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = gErr

    Select Case gErr

        Case 39195, 39197, 39040
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 39196
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCONTO_MAIOR_OU_IGUAL_PRECO_TOTAL", gErr, GridItens.Row, dDesconto, dPrecoTotal)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156359)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ItemNFOrig(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Iten de Nota Fiscal Original que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ItemNFOrig

    Set objGridInt.objControle = ItemNFOriginal

    'Verifica se está preenchido
    If Len(Trim(ItemNFOriginal.ClipText)) > 0 Then

        'Verifica se é um Inteiro
        lErro = Inteiro_Critica(ItemNFOriginal.Text)
        If lErro <> SUCESSO Then gError 39300

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 39301

    Saida_Celula_ItemNFOrig = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemNFOrig:

    Saida_Celula_ItemNFOrig = gErr

    Select Case gErr

        Case 39300, 39301
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156360)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_SerieNFOriginal(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Série de Nota Fiscal Original que está deixando de ser a corrente

Dim lErro As Long
Dim objSerie As New ClassSerie

On Error GoTo Erro_Saida_Celula_SerieNFOriginal

    Set objGridInt.objControle = SerieNFOriginal

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 39199

    Saida_Celula_SerieNFOriginal = SUCESSO

    Exit Function

Erro_Saida_Celula_SerieNFOriginal:

    Saida_Celula_SerieNFOriginal = gErr

    Select Case gErr

        Case 39199
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156361)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CodigoNFOriginal(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Código de Nota Fiscal Original que está deixando de ser a corrente

Dim lErro As Long
Dim objSerie As New ClassSerie

On Error GoTo Erro_Saida_Celula_CodigoNFOriginal

    Set objGridInt.objControle = CodigoNFOriginal

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 39297

    Saida_Celula_CodigoNFOriginal = SUCESSO

    Exit Function

Erro_Saida_Celula_CodigoNFOriginal:

    Saida_Celula_CodigoNFOriginal = gErr

    Select Case gErr

        Case 39297
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156362)

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

        'faz a crítica do valor
        lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
        If lErro <> SUCESSO Then gError 39200

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

        Case 39200

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156363)

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
        If lErro <> SUCESSO Then gError 39201

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

        Case 39201

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156364)

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
        If lErro <> SUCESSO Then gError 39202

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

        Case 39202

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156365)

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
        If lErro <> SUCESSO Then gError 39203

        dValorDesconto = CDbl(ValorDesconto.Text)
        'Coloca o Valor formatado na tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")

        'Verifica se o Valor do Desconto é maior que a soma dos demais valores
        If dValorDesconto > dValorProdutos Then gError 39313

        dValorProdutos = dValorProdutos - dValorDesconto

    End If

    ValorProdutos.Caption = Format(dValorProdutos, "Standard")

    lErro = CF2(Me, "NFCF_ValorDesconto_Validate")
    If lErro <> SUCESSO Then gError 132026

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDesconto_Validate(Cancel, dValorDesconto)
'*** fim tributacao

    Call ValorTotal_Calcula

    iValorDescontoAlterado = 0

    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True


    Select Case gErr

        Case 39203

        Case 39313
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_MAIOR", gErr, dValorDesconto, dValorProdutos)

        Case 132026

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156366)

    End Select

    Exit Sub

End Sub

Public Sub TabStrip1_Click()

Dim lErro As Long
Dim FrameL(1 To 5) As Frame

On Error GoTo Erro_TabStrip1_Click

'    If giTipoVersao = VERSAO_LIGHT Then
'
'        Set FrameL(1) = Frame1(1)
'        Set FrameL(2) = Frame1(2)
'        Set FrameL(3) = Frame1(3)
'        Set FrameL(4) = Frame1(7)
'        Set FrameL(5) = Frame1(8)
'
'    End If

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

''*** incluidos p/tratamento de tributacao *******************************
        'se abriu o tab de tributacao
        If TabStrip1.SelectedItem.Index = TAB_Tributacao Then
            lErro = gobjTribTab.TabClick
            If lErro <> SUCESSO Then gError 39304
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
                Parent.HelpContextID = IDH_NF_SAIDA_DEVOLUCAO_DADOS_PRINCIPAIS

            Case TAB_Itens
                Parent.HelpContextID = IDH_NF_SAIDA_DEVOLUCAO_ITENS

            Case TAB_Complemento
                Call Calcula_Msg
                Parent.HelpContextID = IDH_NF_SAIDA_DEVOLUCAO_COMPLEMENTO

            Case TAB_Tributacao
                Parent.HelpContextID = IDH_NF_SAIDA_DEVOLUCAO_TRIBUTACAO

            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_NF_SAIDA_DEVOLUCAO_CONTABILIZACAO

        End Select

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 39304

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156367)

    End Select

    Exit Sub

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 39217

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscal_ComandoSeta1()
    If lErro <> SUCESSO Then gError 500126

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 39217, 500126

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156368)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Dispara o processo de gravação de uma Nota Fiscal de Devolução

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim vbMsgRes As VbMsgBoxResult
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Chama Valida_Dados_Tela
    lErro = Valida_Dados_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 39169

    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 83450

    If lErro <> SUCESSO Then gError 83449

    'distribuicao
    
    lErro = gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
    If lErro <> SUCESSO Then gError 83415

    Set objNFiscal.objRastreamento = gobjRastreamento

    'Recolhe os Ítens do Grid
    lErro = Move_GridItens_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 39244

    lErro = gobjDistribuicao.Move_GridDist_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 89557
    
    'Mover a parte de tributação
    lErro = gobjTribTab.PreencherNF(objNFiscal)
    If lErro <> SUCESSO Then gError 39308

    lErro = gobjTribTab.Valida_TributacaoNF_Gravacao(objNFiscal)
    If lErro <> SUCESSO Then gError 46434

    lErro = Calcula_Msg
    If lErro <> SUCESSO Then gError 46434

    'mover a parte do rastreamento
    lErro = gobjRastreamento.Move_Rastro_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 83416

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataEmissao.Text))
    If lErro <> SUCESSO Then gError 92043

    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786
    
    lErro = CF("NFiscal_Valida_Diversos", objNFiscal)
    If lErro <> SUCESSO Then gError 181786
    
    'Chama a rotina para iniciar a Gravação(gravando inclusive os dados contábeis) (contabilidade)
    lErro = CF("NFiscalDevolucao_Grava", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 39245

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
        If lErro <> SUCESSO Then gError 199995
    End If
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 39169, 39245, 39244, 39308, 46434, 83415, 83416, 83450, 92043, 130406, 181786, 199995

        Case 83449
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156369)

    End Select

    Exit Function

End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objNFiscal As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objCliente As New ClassCliente
Dim objFornecedor As New ClassFornecedor
Dim snProt As String
Dim iEnviado As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 83997
    If Len(Trim(Serie.Text)) = 0 Then gError 83998
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 83999

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
        
'        If gobjCRFAT.iUsaNFe = MARCADO Then
'
'            lErro = CF("NFeFedLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'            If lErro <> SUCESSO Then gError 203014
'
'            If iEnviado = 1 Then gError 203015
'
'        End If
        
    Else
        'De acordo com a sua Origem verifica se o Cliente ou Fornecedor estão preenchidos
        If objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
        
            If Len(Trim(Cliente.ClipText)) = 0 Then gError 99347
            
            objCliente.sNomeReduzido = Cliente.Text
            'Lê o Cliente
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 99348
            
            If lErro = 12348 Then gError 99349

            objNFiscal.lCliente = objCliente.lCodigo
            objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
            
        ElseIf objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        
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
                
        End If
        
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
        If lErro <> SUCESSO Then gError 89002

        'Faz o cancelamento de uma nota fiscal de Saida
        lErro = CF("NotaFiscalSaida_Excluir", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError 89003

        'Limpa a Tela
        lErro = Limpa_Tela_NFiscal_ComandoSeta1()
        If lErro <> SUCESSO Then gError 89004

        'fecha o comando de setas
        Call ComandoSeta_Fechar(Me.Name)

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 83997
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 83998
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 83999
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 89000, 89002, 89003, 89004, 99345, 99348, 99351, 99352, 99353, 203014

        Case 89065
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
            
        Case 99346
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)
        
        Case 99347
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)
        
        Case 99349
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objNFiscal.lCliente)
        
        Case 99350
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
        
        Case 99352
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)

        Case 203015
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CADASTRADA_NAO_EXCLUI", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156370)

    End Select

    Exit Sub

End Sub

Public Function Move_GridItens_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os Itens do Grid para a Memória

Dim iIndice As Integer
Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lTamanho As Long
Dim objItemNFOrig As New ClassItemNF
Dim objNFOrig As New ClassNFiscal
Dim colAlocacoes As ColAlocacoesItemNF
Dim colCodigoNome As New AdmColCodigoNome
Dim iIndice1 As Integer
Dim objItemNFNovo As ClassItemNF

On Error GoTo Erro_Move_GridItens_Memoria

    'Para cada linha existente do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemNF = New ClassItemNF

        'Verifica se o Produto está preenchido
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 39314

        'Armazena produto
        objItemNF.sProduto = sProdutoFormatado

        'Guarda os demais campos do Grid em objItemNF
        objItemNF.sDescricaoItem = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
        objItemNF.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then objItemNF.dQuantidade = CDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) > 0 Then objItemNF.dPrecoUnitario = CDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))
        lTamanho = Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col)))
        If lTamanho > 0 Then objItemNF.dPercDesc = CDbl(Format(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col), "General Number"))

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))) > 0 Then objItemNF.dValorDesconto = CDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))

        'Se os dados do itemNF original estiverem preenchidos
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_SerieNFOrig_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_CodNFOrig_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ItemNFOrig_Col))) > 0 Then

            Set objItemNFOrig = New ClassItemNF
            'recolhe os dados
            objItemNFOrig.sSerieNFOrig = GridItens.TextMatrix(iIndice, iGrid_SerieNFOrig_Col)
            objItemNFOrig.lNumNFOrig = GridItens.TextMatrix(iIndice, iGrid_CodNFOrig_Col)
            objItemNFOrig.iItem = GridItens.TextMatrix(iIndice, iGrid_ItemNFOrig_Col)
            'BUsca no Bd Item Nf com os dados informados
            
            If giComplemento = MARCADO Then 'ItemNFiscalSaida_Le_NumNFItem1
                lErro = CF("ItemNFiscalSaida_Le_NumNFItem1", objItemNFOrig)
                If lErro <> SUCESSO And lErro <> 35368 Then gError 39316
    
                If lErro <> SUCESSO Then gError 199745
                
                If objItemNFOrig.sProduto <> objItemNF.sProduto Then gError 199746
            Else
                lErro = CF("ItemNFEntradaNaoDev_Le_NumNFItem", objNFiscal, objItemNFOrig)
                If lErro <> SUCESSO And lErro <> 39323 Then gError 39316
                If lErro = 39323 Then gError 39318
            End If
            
            'se o item nÃo corresponder ao item informado --> erro
            If objItemNFOrig.sProduto <> objItemNF.sProduto Then gError 39319

            objItemNF.lNumIntDocOrig = objItemNFOrig.lNumIntDoc
            
            objNFOrig.lNumIntDoc = objItemNFOrig.lNumIntNF
        
            lErro = CF("NFiscal_Le", objNFOrig)
            If lErro <> SUCESSO And lErro <> 31442 Then gError 39316
            
            If objNFOrig.dtDataEmissao > objNFiscal.dtDataEmissao Then gError 206614
        
        End If

        'Adiciona na coleção de Ítens
        With objItemNF
            Set objItemNFNovo = objNFiscal.ColItensNF.Add(0, iIndice, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, DATA_NULA, .sDescricaoItem, 0, 0, 0, 0, 0, colAlocacoes, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .sCcl, STATUS_BAIXADO, .lNumIntDocOrig, "", 0, 0, 0, "", 0, 0)
        End With
        
        Set objItemNFNovo.colItensNFDevBenefItemOP = gobjNFiscal.ColItensNF(iIndice).colItensNFDevBenefItemOP
        Set objItemNFNovo.colItensNFDevBenefRem = gobjNFiscal.ColItensNF(iIndice).colItensNFDevBenefRem

        lErro = gobjDistribuicao.Move_DistribuicaoGrade_Memoria(gobjNFiscal.ColItensNF(iIndice))
        If lErro <> SUCESSO Then gError 39316
    
        Set objNFiscal.ColItensNF(iIndice).objInfoAdicDocItem = gobjNFiscal.ColItensNF(iIndice).objInfoAdicDocItem
    
        '********************* TRATAMENTO DE GRADE *****************
        Call Move_ItensGrade_Tela(objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade, gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade)
        
        If objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade.Count > 0 Then objNFiscal.ColItensNF(iIndice).iPossuiGrade = MARCADO

        lErro = CF2(Me, "NFDEV_Move_GridItens_Memoria", objItemNFNovo, iIndice)
        If lErro <> SUCESSO Then gError 181283

    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 39316, 39317, 39314, 39315, 181283

        Case 39318
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NFORIGINAL_NAO_CADASTRADO", gErr, iIndice, objItemNFOrig.sSerieNFOrig, objItemNFOrig.lNumNFOrig)

        Case 39319
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CASADO", gErr, iIndice, objItemNF.sProduto, objItemNFOrig.sProduto)

        Case 199745
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NFORIGINAL_NAO_CADASTRADO", gErr, iIndice, objItemNFOrig.sSerieNFOrig, objItemNFOrig.lNumNFOrig)

        Case 199746
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CASADO", gErr, iIndice, objItemNF.sProduto, objItemNFOrig.sProduto)

        Case 206614
            Call Rotina_Erro(vbOKOnly, "ERRO_DEV_DATA_MENOR_ORIG", gErr, objNFiscal.dtDataEmissao, objNFOrig.dtDataEmissao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156371)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_NFiscal() As Long
'limpa a tela sem fechar o comando de setas

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_NFiscal

    Call Limpa_Tela(Me)
    Set gobjAnotacao = Nothing
    Set gobjNFDevBenef = New ClassNFDevBenef
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    Call Grid_Limpa(objGridItens)

'    If giTipoVersao = VERSAO_FULL Then
        Serie.Text = ""
        Call Serie_Validate(bSGECancelDummy)
        TabelaPreco.Text = ""
'    End If

    Set gobjNFiscal = New ClassNFiscal
    Set gobjInfoAdic = Nothing
    
    NFiscal.Caption = ""
    SerieNFiscalOriginal.Text = ""
    Filial.Clear
    Status.Caption = ""
    FilialEntrega.Clear
    ICMSBase1.Caption = ""
    ICMSValor1.Caption = ""
    ICMSSubstBase1.Caption = ""
    ICMSSubstValor1.Caption = ""
    ValorProdutos.Caption = ""
    ValorProdutos2.Caption = ""
    IPIValor1.Caption = ""
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
    
    NFiscalPropria.Value = vbUnchecked
    TipoFrete.ListIndex = 0

    Call CF2(Me, "NFCF_Default_DataSaida")

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    Call DataEmissao_Validate(bSGECancelDummy)

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.Limpa_Tela
'*** fim tributacao

    giFilialEntregaAlterada = 0
    giFilialEntregaManual = 0
    iClienteAlterado = 0
    iFornecedorAlterado = 0
    iValorDescontoAlterado = 0
    iValorFreteAlterado = 0
    iValorSeguroAlterado = 0
    iValorDespesasAlterado = 0

    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    'Seleciona série padrão
    lErro = Seleciona_Serie_Padrao()
    If lErro <> SUCESSO Then gError 500122

    'distribuicao
    Call gobjDistribuicao.Limpa_Tela_Distribuicao
    
    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento

    Call CF2(Me, "NFCF_Limpa_Tela")

    MsgAutomatica.Value = vbChecked

    iAlterado = 0

    Exit Function

Erro_Limpa_Tela_NFiscal:

    Limpa_Tela_NFiscal = gErr

    Select Case gErr

        Case 500122

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156372)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_NFiscal1() As Long
'Limpa a Tela mas mantém o tipo

Dim iTipoNFiscal As Integer
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_NFiscal1

    iTipoNFiscal = TipoNFiscal.ListIndex

    lErro = Limpa_Tela_NFiscal()
    If lErro <> SUCESSO Then gError 500123

    If iTipoNFiscal <> -1 Then TipoNFiscal.ListIndex = iTipoNFiscal

    Exit Function

Erro_Limpa_Tela_NFiscal1:

    Limpa_Tela_NFiscal1 = gErr

    Select Case gErr

        Case 500123

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156373)

    End Select

    Exit Function

End Function

Private Function PrecoUnitario_Calcula(sUM As String, objTabelaPrecoItem As ClassTabelaPrecoItem, dPrecoUnitario As Double) As Long

Dim objProduto As New ClassProduto
Dim objUM As New ClassUnidadeDeMedida
Dim objUMEst As New ClassUnidadeDeMedida
Dim dFator As Double
Dim lErro As Long

On Error GoTo Erro_PrecoUnitario_Calcula

    objProduto.sCodigo = objTabelaPrecoItem.sCodProduto

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 39339
    If lErro = 28030 Then gError 39340

    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
    If lErro <> SUCESSO Then gError 39341
    'Converte o preço para a unidade de venda
    dPrecoUnitario = objTabelaPrecoItem.dPreco / dFator

    PrecoUnitario_Calcula = SUCESSO

    Exit Function

Erro_PrecoUnitario_Calcula:

    PrecoUnitario_Calcula = gErr

    Select Case gErr

        Case 39341, 39339

        Case 39340
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objTabelaPrecoItem.sCodProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156374)

    End Select

    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 39296

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscal_ComandoSeta()
    If lErro <> SUCESSO Then gError 500127

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 39296, 500127

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156375)

    End Select

    Exit Sub

End Sub


'inicio contabilidade
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
'traz o lote selecionado para a tela
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
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim sContaMascarada As String
Dim dQuantidadeConvertida As Double
Dim objItem As New ClassTribItemNFAux
Dim dQuantidade As Double
Dim iTipoNF As Integer
Dim sProduto As String
Dim sUM As String
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objProdutoCategoria As New ClassProdutoCategoria
Dim objCatProdItem As ClassCategoriaProdutoItem
Dim objPlanoConta As New ClassPlanoConta
Dim objProduto As New ClassProduto
Dim objFilial As New ClassFilialCliente
Dim sContaTela As String
Dim objTribItemNFAux As ClassTribItemNFAux, dValorAux As Double

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case ESCANINHO_CUSTO
            For iLinha = 1 To objGridItens.iLinhasExistentes

                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                    iTipoNF = Codigo_Extrai(TipoNFiscal.Text)

                    Select Case iTipoNF

                        Case DOCINFO_NFISDC
                            objMnemonicoValor.colValor.Add ESCANINHO_3_EM_CONSIGNACAO

                        Case DOCINFO_NFISDCNT
                            objMnemonicoValor.colValor.Add ESCANINHO_3_EM_CONSERTO

                        Case DOCINFO_NFISDD
                            objMnemonicoValor.colValor.Add ESCANINHO_3_EM_DEMO

                        Case DOCINFO_NFISODC, DOCINFO_NFISODF
                            objMnemonicoValor.colValor.Add ESCANINHO_3_EM_OUTROS

                        Case Else
                            objMnemonicoValor.colValor.Add ESCANINHO_NOSSO

                    End Select

                Else
                    objMnemonicoValor.colValor.Add 0

                End If

            Next

        Case ICMS_DEBITO_VALOR

            For Each objItem In gobjTribTab.objNFTributacao.colItens
                objMnemonicoValor.colValor.Add objItem.dICMSCredito
            Next


        Case IPI_DEBITO_VALOR

            For Each objItem In gobjTribTab.objNFTributacao.colItens
                objMnemonicoValor.colValor.Add objItem.dIPICredito
            Next

        Case ICMS_DEBITO_FRETE
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribFrete.dICMSCredito

        Case ICMS_DEBITO_SEGURO
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribSeguro.dICMSCredito

        Case ICMS_DEBITO_DESPESAS
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribOutras.dICMSCredito

        Case IPI_DEBITO_FRETE
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribFrete.dIPICredito

        Case IPI_DEBITO_SEGURO
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribSeguro.dIPICredito

        Case IPI_DEBITO_DESPESAS
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribOutras.dIPICredito

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
            If Len(Filial.Text) > 0 And iFornecedorAtivo = FORN_INATIVO Then
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
                If lErro <> SUCESSO Then gError 88027

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

        Case FORNECEDOR1
            If Len(Trim(Fornecedor.Text)) > 0 Then
                objMnemonicoValor.colValor.Add Fornecedor.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case FILIAL_FORNECEDOR1
            If Len(Trim(Filial.Text)) > 0 And iFornecedorAtivo = FORN_ATIVO Then
                objMnemonicoValor.colValor.Add Filial.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case CTAFORNCONSIG

            'Se o fornecedor foi preenchido
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                'Lê o fornecedor informado pelo usuáro e guarda em objFornecedor.sNomeReduzido
                objFornecedor.sNomeReduzido = Trim(Fornecedor.Text)
    
                'Lê os dados de fornecedor com o NomeReduzido passado
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO And lErro <> 6682 Then gError 84902
                
                'Atribui o codigo do fornecedor encontrado a objFilialFornecedor.lCodFornecedor
                objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
                
                'Lê a filial informada pelo usuáro e guarda em objFilialFornecedor.iCodFilial
                objFilialFornecedor.iCodFilial = Codigo_Extrai(Filial.Text)
                
                'Lê a filial com o código do fornecedor e codfilial passados
                lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
                If lErro <> SUCESSO And lErro <> 12929 Then gError 84903
                
                'Se encontrou a Conta de Fornecedores em Consignação
                If Len(Trim(objFilialFornecedor.sContaFornConsig)) > 0 Then
    
                    'Inicializa a string
                    sContaMascarada = String(STRING_CONTA, 0)
    
                    'Formata a conta que será exibida na tela
                    lErro = Mascara_MascararConta(objFilialFornecedor.sContaFornConsig, sContaMascarada)
                    If lErro <> SUCESSO Then gError 90564
    
                    objMnemonicoValor.colValor.Add sContaMascarada
                
                'Senão (se não encontrou a conta de fornecedores em consignação)
                Else
                    
                    objMnemonicoValor.colValor.Add ""
                
                End If
            
            'Senão (se o fornecedor não foi preenchido)
            Else
                
                objMnemonicoValor.colValor.Add ""
                
            End If
            
        Case ICMS_DEBITO_TOTAL
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dICMSCredito
        
        Case IPI_DEBITO_TOTAL
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dIPICredito
        
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
                    If lErro <> SUCESSO Then gError 183762
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183763
        
                    sContaMascarada = String(STRING_CONTA, 0)

                    lErro = Mascara_MascararConta(objProduto.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError 183796

                    objMnemonicoValor.colValor.Add sContaMascarada
                    
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case CONTACONTABILCATPROD
                
            For iLinha = 1 To objGridItens.iLinhasExistentes
                
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183764
                
                    objProdutoCategoria.sCategoria = objMnemonicoValor.vParam(1)
                    
                    objProdutoCategoria.sProduto = sProdutoFormatado
                    
                    Set objCatProdItem = New ClassCategoriaProdutoItem
                    
                    lErro = CF("ProdutoCategoria_Le_Item", objProdutoCategoria, objCatProdItem)
                    If lErro <> SUCESSO And lErro <> 183548 Then gError 183765
                    
                    If lErro = SUCESSO Then
                        
                        Set objPlanoConta = New ClassPlanoConta
                    
                        'le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta
                        lErro = CF("PlanoConta_Le_ContaSimples", objCatProdItem.dvalor1, objPlanoConta)
                        If lErro <> SUCESSO And lErro <> 5451 Then gError 183766
                    
                        If lErro = SUCESSO Then
                    
                            sContaMascarada = String(STRING_CONTA, 0)
        
                            lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
                            If lErro <> SUCESSO Then gError 183797
        
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
                If lErro <> SUCESSO Then gError 186136
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 186137
                
                Else
                    sContaTela = ""
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case CTAFILIALFORNECEDOR1
            
            If Len(Trim(Filial.Text)) > 0 Then
                
                objFilialFornecedor.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilialFornecedor)
                If lErro <> SUCESSO Then gError 186146
                
                If objFilialFornecedor.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilialFornecedor.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 186147
                
                Else
                    sContaTela = ""
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case FilialEmpresa
            objMnemonicoValor.colValor.Add giFilialEmpresa
        
        Case Else
            lErro = gobjTribTab.Calcula_Mnemonico(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 39667

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39667
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case 61817, 61818, 61819, 61820, 64213, 75459, 75461, 75462, 75464, 88027, 84900, 84901, 84902, 183762 To 183766, 183796, 183797, 186136, 186137, 186146, 186147

        Case 75460
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO1", gErr, objAlmoxarifado.sNomeReduzido)

        Case 75463
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156376)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_NFiscal_ComandoSeta() As Long
'limpa a tela fechando o comando de setas
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_NFiscal_ComandoSeta

    lErro = ComandoSeta_Fechar(Me.Name)

    lErro = Limpa_Tela_NFiscal()
    If lErro <> SUCESSO Then gError 500128

    Exit Function

Erro_Limpa_Tela_NFiscal_ComandoSeta:

    Limpa_Tela_NFiscal_ComandoSeta = gErr

    Select Case gErr

        Case 500128

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156377)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_NFiscal_ComandoSeta1() As Long
'limpa a tela fechando o comando de setas, mas mantém a natureza e o tipo
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_NFiscal_ComandoSeta1

    lErro = ComandoSeta_Fechar(Me.Name)

    lErro = Limpa_Tela_NFiscal1()
    If lErro <> SUCESSO Then gError 500129

    Exit Function

Erro_Limpa_Tela_NFiscal_ComandoSeta1:

    Limpa_Tela_NFiscal_ComandoSeta1 = gErr

    Select Case gErr

        Case 500129

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156378)

    End Select

    Exit Function

End Function

Public Function TabelaPreco_Atualiza_PrecoUnitario(objProduto As ClassProduto) As Long
'Se a Tabela de Preços estiver selecionada tenta calcular o preço unitário para o produto passado

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim dPrecoUnitario As Double

On Error GoTo Erro_TabelaPreco_Atualiza_PrecoUnitario

    'Verifica se a TabelaPreco está selecionada
    If Len(Trim(TabelaPreco.Text)) > 0 Then

        objTabelaPrecoItem.iCodTabela = Codigo_Extrai(TabelaPreco.Text)
        objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
        objTabelaPrecoItem.sCodProduto = objProduto.sCodigo

        'Lê o preço do produto para a filialempresa
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 42864
        'Se não encontrar
        If lErro = 28014 Then
            objTabelaPrecoItem.iFilialEmpresa = EMPRESA_TODA

            'Lê o preço para a empresa toda
            lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
            If lErro <> SUCESSO And lErro <> 28014 Then gError 42865

        End If

        'Se encontrou o valor na Tabela de preço
        If lErro <> 28014 Then
            'Calcula o preço unitário em relação a UM da tela
            lErro = PrecoUnitario_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col), objTabelaPrecoItem, dPrecoUnitario)
            If lErro <> SUCESSO Then gError 42866

            'Coloca o preço unitário na tela
            If dPrecoUnitario > 0 Then
                GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
            Else
                GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = ""
            End If
        End If
    End If

    TabelaPreco_Atualiza_PrecoUnitario = SUCESSO

    Exit Function

Erro_TabelaPreco_Atualiza_PrecoUnitario:

    TabelaPreco_Atualiza_PrecoUnitario = gErr

    Select Case gErr

        Case 42864, 42865, 42866

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156379)

    End Select

    Exit Function

End Function

Private Function Valida_Dados_Tela(objNFiscal As ClassNFiscal) As Long
'Valida os dados de NFiscal da tela

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim dValorTotal As Double, dValorDesconto As Double
Dim objNFiscalOriginal As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objTipoDocInfoOrig As New ClassTipoDocInfo
Dim iIndiceAux As Integer

On Error GoTo Erro_Valida_Dados_Tela

    'Verifica se os campos obrigatórios da tela estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 39218
    If Len(Trim(gobjTribTab.sNatOpInterna)) = 0 Then gError 39219
    If Len(Trim(FilialEntrega.Text)) = 0 Then gError 51523

    'Lê o Tipo da NF
    If iFornecedorAtivo <> FORN_ATIVO Then
        If Len(Trim(Cliente.ClipText)) = 0 Then gError 39222
    Else
        If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 39223
    End If

    If Len(Trim(Filial.Text)) = 0 Then gError 39224
    If Len(Trim(Serie.Text)) = 0 Then gError 39225
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 39227

    If Len(Trim(PesoLiquido.ClipText)) > 0 And Len(Trim(PesoBruto.ClipText)) > 0 Then
        If CDbl(PesoLiquido.Text) > CDbl(PesoBruto.Text) Then gError 39240
    End If

    'Verifica se a Data de Entrada é menor que a Data de Emissao
    If Len(Trim(DataSaida.ClipText)) <> 0 Then
        If StrParaDate(DataSaida.Text) < StrParaDate(DataEmissao.Text) Then gError 39229
    End If

    If Len(Trim(ValorTotal.Caption)) > 0 Then dValorTotal = CDbl(ValorTotal.Caption)

    'Verifica se o Total é positivo
    If dValorTotal < 0 Then gError 39237

    'Verifica se o Valor do Desconto é maior que a soma dos demais valores
    If Len(Trim(ValorDesconto.Text)) > 0 Then dValorDesconto = CDbl(ValorDesconto.Text)
    lErro = DescontoGlobal_Testa(dValorDesconto)
    If lErro <> SUCESSO Then gError 56901

    lErro = gobjTribTab.Valida_Dados
    If lErro <> SUCESSO Then gError 56901
    
    'Verifica se o Grid está vazio
    If objGridItens.iLinhasExistentes = 0 Then gError 39230

    For iIndice = 1 To objGridItens.iLinhasExistentes

        'Verifica se os campos obrigatorios do Grid estão todos preenchidos

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 39231

        For iIndiceAux = 0 To gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade.Count

            If gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade.Count = 0 Then
                lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProduto, iPreenchido)
                If lErro <> SUCESSO Then gError 39232

            Else
                If iIndiceAux = 0 Then iIndiceAux = 1
                sProduto = gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade(iIndiceAux).sProduto
            
            End If
            
            objProduto.sCodigo = sProduto
    
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 39233
    
            'Se não achou o Produto --> erro
            If lErro = 28030 Then gError 39234
        
        Next
        
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) = 0 Then gError 39236

    Next

    'Guarda todos os campos da tela em objNFiscal
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 39241

    'Se os dados da Nota Fiscal Original estão preenchidos
    If Len(Trim(SerieNFiscalOriginal.Text)) > 0 And Len(Trim(NFiscalOriginal.ClipText)) > 0 Then

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 62301
        If lErro <> SUCESSO Then gError 62302

        If NFiscalPropria.Value = vbChecked Then

            objNFiscalOriginal.sSerie = SerieNFiscalOriginal.Text
            objNFiscalOriginal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
            objNFiscalOriginal.iFilialEmpresa = giFilialEmpresa

            'Lê a Nota Fiscal Original
            
            If giComplemento = MARCADO Then
                lErro = CF("NFiscalInternaSaida_Le_Numero2", objNFiscalOriginal)
            Else
                lErro = CF("NFiscalInternaEntNaoDev_Le_Numero", objNFiscalOriginal)
            End If
            If lErro <> AD_SQL_SUCESSO And lErro <> 39101 Then gError 39238

            'Não encontrou
            If lErro = 39101 Then gError 39239

        ElseIf NFiscalPropria.Value = vbUnchecked Then
            
            objNFiscalOriginal.sSerie = SerieNFiscalOriginal.Text
            objNFiscalOriginal.lNumNotaFiscal = CLng(NFiscalOriginal)
            objNFiscalOriginal.lCliente = objNFiscal.lCliente
            objNFiscalOriginal.lFornecedor = objNFiscal.lFornecedor
            objNFiscalOriginal.iFilialCli = objNFiscal.iFilialCli
            objNFiscalOriginal.iFilialForn = objNFiscal.iFilialForn

            lErro = CF("NFiscalExternaNaoDev_Le_Numero", objNFiscalOriginal)
            If lErro <> SUCESSO And lErro <> 39116 Then gError 39242
            If lErro <> SUCESSO Then gError 39243

            objTipoDocInfoOrig.iCodigo = objNFiscalOriginal.iTipoNFiscal

            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfoOrig)
            If lErro <> SUCESSO And lErro <> 31415 Then gError 62303
            If lErro <> SUCESSO Then gError 62304

            If Not ((objNFiscal.iTipoNFiscal = DOCINFO_NFISDBF Or objNFiscal.iTipoNFiscal = DOCINFO_NFISDBFNE) And (objTipoDocInfoOrig.iCodigo = DOCINFO_NFIE3BF Or objTipoDocInfoOrig.iCodigo = DOCINFO_NFEE3BF Or objTipoDocInfoOrig.iCodigo = DOCINFO_NFEERMB3 Or objTipoDocInfoOrig.iCodigo = DOCINFO_NFIERMB3)) Then
                
                'para tratar casos de mais de um tipo de nf de devolucao para um tipo de entrada
                Select Case objTipoDocInfoOrig.iCodigo
            
                    Case DOCINFO_NFEEC
                        If objTipoDocInfo.sSigla <> "NFISDSC" And objTipoDocInfo.sSigla <> "NFISDC" Then gError 62305
            
                    Case Else
                        If objTipoDocInfo.sSigla <> objTipoDocInfoOrig.sSiglaNFOriginal Then gError 62305
                        
                End Select
            
            End If

        End If

        'Guarda em objNFiscal o Número Interno da Nota Fiscal Original
        objNFiscal.lNumIntNotaOriginal = objNFiscalOriginal.lNumIntDoc

    End If

    Valida_Dados_Tela = SUCESSO

    Exit Function

Erro_Valida_Dados_Tela:

    Valida_Dados_Tela = gErr

    Select Case gErr

        Case 39218
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 39219
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 39232, 39233, 39238, 39241, 39242, 39245, 56901, 62301, 62303

        Case 39222
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 39223
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 39224
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 39225
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 39227
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 39228
            Call Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_NAO_PREENCHIDA", gErr)

        Case 39229
            Call Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_ANTERIOR_DATAEMISSAO", gErr, DataSaida.Text, DataEmissao.Text)

        Case 39230
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)

        Case 39231
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case 39234
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 39235
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 39236
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 39237
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)

        Case 39239
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_INTERNA_ENTRADA_NAO_CADASTRADA", gErr, objNFiscalOriginal.sSerie, objNFiscalOriginal.lNumNotaFiscal)

        Case 39240
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr, CDbl(PesoLiquido.Text), CDbl(PesoBruto.Text))

        Case 39243
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_EXTERNA_NAO_CADASTRADA", gErr, objNFiscalOriginal.sSerie, objNFiscalOriginal.lNumNotaFiscal, objNFiscalOriginal.lFornecedor, Codigo_Extrai(Filial.Text))

        Case 51523
            If iFornecedorAtivo = FORN_ATIVO Then
                Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_FORNECEDOR_NAO_INFORMADA", gErr)
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)
            End If

        Case 56903
            'Call Rotina_Erro(vbOKOnly, "ERRO_IR_FONTE_MAIOR_VALOR_TOTAL", gErr, dValorIRRF, dValorTotal)

        Case 62302
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case 62305
            If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Then
                Call Rotina_Erro(vbOKOnly, "ERRO_TIPONFISCALORIGINAL_DIFERENTE_TELA", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.lCliente, objNFiscal.iFilialCli, objTipoDocInfo.sSigla, objTipoDocInfoOrig.sSiglaNFOriginal)
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_TIPONFISCALORIGINAL_DIFERENTE_TELA", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objTipoDocInfo.sSigla, objTipoDocInfoOrig.sSiglaNFOriginal)
            End If

        Case 62304
            Call Rotina_Erro(gErr, "ERRO_TIPONFISCALORIGINAL_NAO_ENCONTRADO", gErr, objTipoDocInfoOrig.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156380)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescricaoItem(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoItem

    Set objGridInt.objControle = DescricaoItem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42872

    Saida_Celula_DescricaoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoItem:

    Saida_Celula_DescricaoItem = gErr

    Select Case gErr

        Case 42872
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156381)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    'feito em virtude do redimensionamento da msg da nota p/ inpal
    Call CF("Redimensiona_Tam_MsgNF", Mensagem)

    Parent.HelpContextID = IDH_NF_SAIDA_DEVOLUCAO_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Nota Fiscal de Saída - Devolução"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "NFiscalDev"
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

    If Len(Trim(ValorProdutos.Caption)) > 0 Then dValorProdutos = CDbl(ValorProdutos.Caption)
    If Len(Trim(ValorFrete.ClipText)) > 0 Then dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorSeguro.ClipText)) > 0 Then dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(ValorDespesas.ClipText)) > 0 Then dValorDespesas = CDbl(ValorDespesas.Text)

    'Verifica se o Valor do Desconto é maior que a soma dos demais valores
    If dValorProdutos + dValorFrete + dValorSeguro + dValorDespesas < 0 Then gError 56897

    DescontoGlobal_Testa = SUCESSO

    Exit Function

Erro_DescontoGlobal_Testa:

    DescontoGlobal_Testa = gErr

    Select Case gErr

        Case 56897
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_ULTRAPASSOU_SOMA_VALORES", gErr, dValorDesconto, dValorProdutos + dValorDesconto + dValorFrete + dValorSeguro + dValorDespesas) 'Alterado por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156382)

    End Select

    Exit Function

End Function

Private Function Trata_FilialForn() As Long

Dim lErro As Long
Dim objFilialForn As New ClassFilialFornecedor

On Error GoTo Erro_Trata_FilialForn

    objFilialForn.iCodFilial = Codigo_Extrai(Filial.Text)

    'Lê a Filial
    lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilialForn)
    If lErro <> SUCESSO Then gError 56965

    Call gobjTribTab.FilialFornecedor_Alterada(objFilialForn.lCodFornecedor, objFilialForn.iCodFilial)

    lErro = CF2(Me, "NFCF_Fornecedor_Filial_Validate", objFilialForn)
    If lErro <> SUCESSO Then gError 131884

    Trata_FilialForn = SUCESSO

    Exit Function

Erro_Trata_FilialForn:

    Trata_FilialForn = gErr

    Select Case gErr

        Case 56965, 131884

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156383)

    End Select

    Exit Function

End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

        If Me.ActiveControl Is Serie Then
            Call SerieLabel_Click
        ElseIf Me.ActiveControl Is NFiscal Then
            Call NFiscalLabel_Click
        ElseIf Me.ActiveControl Is SerieNFiscalOriginal Then
            Call SerieOriginalLabel_Click
        ElseIf Me.ActiveControl Is NFiscalOriginal Then
            Call NFOriginalLabel_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is ItemNFOriginal Then
            Call BotaoItensNFiscalOrig_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        ElseIf Me.ActiveControl Is TranspRedespacho Then
            Call TranspRedLabel_Click
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        Else
            Call gobjRastreamento.objUserControl_KeyDown(KeyCode, Shift)
        End If

    ElseIf KeyCode = KEYCODE_CODBARRAS Then
        Call Trata_CodigoBarras
        
    End If
End Sub

Public Sub Filial_Formata(objFilial As Object, iFilial As Integer)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim sFornecedor As String
Dim objFilialFornecedor As New ClassFilialFornecedor

On Error GoTo Erro_Filial_Formata


    objFilial.Text = CStr(iFilial)

    If iFornecedorAtivo = FORN_INATIVO Then
        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iFilial

        'Pesquisa se existe Filial de cliente com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 51521

        If lErro = 17660 Then gError 51522

        'Coloca na tela a Filial de Cliente lida
        objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Else
        sFornecedor = Fornecedor.Text
        objFilialFornecedor.iCodFilial = iFilial

        'Pesquisa se existe Filial de fornecedor com o código extraído
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 51625
        If lErro <> SUCESSO Then gError 51626

        'Coloca na tela a Filial de fornecedor lida
        objFilial.Text = iFilial & SEPARADOR & objFilialFornecedor.sNome

    End If

    Exit Sub

Erro_Filial_Formata:

    Select Case gErr

        Case 51521, 51625

        Case 51522
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)

        Case 51626
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, objFilial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156384)

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
    If lErro <> SUCESSO Then gError 56915

    lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 56916

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", Filial, colCodigoNome)
    Call CF("Filial_Preenche", FilialEntrega, colCodigoNome)

    iClienteAlterado = 0

    Exit Sub

Erro_Cliente_Formata:

    Select Case gErr

        Case 56915, 56916

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156385)

    End Select

    Exit Sub

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

Private Sub NFiscal_Imprime_Romaneio(objNFiscal As ClassNFiscal)

    Call CF("NFiscal_Imprime_Romaneio", objNFiscal)
    
End Sub

Public Sub BotaoConsultaDocCPR_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objFornecedor As New ClassFornecedor
Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoConsultaDocCPR_Click

    'Verifica se todos os campos necessários para se efetuar a consulta foram preenchidos
    lErro = Critica_CamposNecessarios_ConsultaTitulo()
    If lErro <> SUCESSO Then gError 79838

    'Guarda no objNFiscal os dados necessários para consultar o título
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.iTipoDocInfo = Codigo_Extrai(TipoNFiscal.Text)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.dtDataEntrada = DATA_NULA

    If Fornecedor.Visible = True Then

        'Guarda em objFornecedor o nome reduzido do Fornecedor
        objFornecedor.sNomeReduzido = Fornecedor.Text

        'Lê o código do Fornecedor a partir do nome reduzido obtido na tela
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 79839

        'Se não encontrou o fornecedor = > erro
        If lErro = 6681 Then gError 79840

        'Guarda no objNFiscal o código do fornecedor
        objNFiscal.lFornecedor = objFornecedor.lCodigo
        objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)

    ElseIf Cliente.Visible = True Then

        'Guarda em objCliente o nome reduzido do Cliente
        objCliente.sNomeReduzido = Cliente.Text

        'Lê o código do Fornecedor a partir do nome reduzido obtido na tela
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 79841

        'Se não encontrou o cliente = > erro
        If lErro = 12348 Then gError 79842

        'Guarda no objNFiscal o código do fornecedor
        objNFiscal.lCliente = objCliente.lCodigo
        objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)

    End If

    'Lê o NumIntDocCPR da NFiscal e exibe o documento gerado no CPR por essa nota
    lErro = CF("NFiscal_Consulta_DocCPR", objNFiscal)
    If lErro <> SUCESSO And lErro <> 79717 Then gError 79843

    'Se não encontrou a nota
    If lErro = 79717 Then

        'Testa se a origem da nota é um fornecedor e exibe mensagem de erro
        If Fornecedor.Visible = True Then gError 79844

        'Testa se a origem da nota é um cliente e exibe mensagem de erro
        If Cliente.Visible = True Then gError 79845

    End If

    Exit Sub

Erro_BotaoConsultaDocCPR_Click:

    Select Case gErr

        Case 79838, 79843, 79839, 79841

        Case 79845
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA2", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada)

        Case 79844
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA4", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iTipoNFiscal, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada)

        Case 79840
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)

        Case 79842
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objNFiscal.lCliente)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156386)

    End Select

    Exit Sub

End Sub

Private Function Critica_CamposNecessarios_ConsultaTitulo() As Long
'Verifica se os campos necessários para consultar um título gerado por uma NFEntrada foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposNecessarios_ConsultaTitulo

    'Se o campo Fornecedor está visível
    If Fornecedor.Visible = True Then

        'Verifica se ele foi preenchido
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 79846

    'Se o campo Cliente está visível
    ElseIf Cliente.Visible = True Then

        'Verifica se o cliente foi preenchido
        If Len(Trim(Cliente.Text)) = 0 Then gError 79847

    End If

    'Verifica se a filial do Fornecedor / Cliente foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 79848

    'Verifica se a Série foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 79849

    'Verifica se o Número da Nota foi preenchido
    If Len(Trim(NFiscal.Caption)) = 0 Then gError 79850

    'Verifica se a data de emissão da nota foi preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 79851

    'Verifica se o tipo da Nota fiscal foi selecionado
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 79852

    Critica_CamposNecessarios_ConsultaTitulo = SUCESSO

    Exit Function

Erro_Critica_CamposNecessarios_ConsultaTitulo:

    Critica_CamposNecessarios_ConsultaTitulo = gErr

    Select Case gErr

        Case 79846
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 79847
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 79848
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 79849
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 79850
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)

        Case 79851
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 79852
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156387)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156388)

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

Public Property Get BotaoGrade() As Object
     Set BotaoGrade = objUserControl.Controls("BotaoGrade")
End Property

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
                        
            Call gobjDistribuicao.Move_DistribuicaoGrade_Memoria(objItemNF)
            
            Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
            If lErro <> SUCESSO Then gError 86360
        
            Call Atualiza_Grid_Itens(objItemNF)
        
            Call gobjDistribuicao.Atualiza_Grid_Distribuicao(objItemNF)
            
            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) > 0 Then
                Call PrecoTotal_Calcula(GridItens.Row)
                
                lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
                If lErro <> SUCESSO Then gError 86360
                
                Call ValorTotal_Calcula
            End If
        
        End If
    
    End If
    
    Exit Sub

Erro_BotaoGrade_Click:

    Select Case gErr
    
        Case 51037, 86360
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156389)
            
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

Function Move_ItensGrade_Tela(colItensRomaneio As Collection, colItensRomaneioTela As Collection, Optional bTrazPedido As Boolean = False) As Long

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradeTela As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim objReservaItemTela As ClassReservaItem
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lErro As Long, iFilialEmpresa As Integer

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
            
            If objReservaItemTela.iAlmoxarifado > 0 Then
            
                objAlmoxarifado.iCodigo = objReservaItemTela.iAlmoxarifado
                            
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                
                objAlmoxarifado.sNomeReduzido = objAlmoxarifado.sNomeReduzido
            
            Else
                objAlmoxarifado.sNomeReduzido = objReservaItemTela.sAlmoxarifado
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            
                objReservaItemTela.iAlmoxarifado = objAlmoxarifado.iCodigo
            End If
            
            iFilialEmpresa = giFilialEmpresa
            Call CF("FilialEmpresaAlmox_Customiza", iFilialEmpresa)
                
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
                objItemRomaneioGrade.iAlmoxarifado = objItemRomaneioGradeTela.iAlmoxarifado
                objItemRomaneioGrade.sAlmoxarifado = objItemRomaneioGradeTela.sAlmoxarifado
                
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

Public Sub BotaoLocalizacaoDist_Click()
'distribuicao
    Call gobjDistribuicao.BotaoLocalizacaoDist_Click
End Sub

Public Property Get ProdutoAlmoxDist() As Object
     Set ProdutoAlmoxDist = objUserControl.Controls("ProdutoAlmoxDist")
End Property

Public Sub ProdutoAlmoxDist_Change()
'distribuicao
    Call gobjDistribuicao.ProdutoAlmoxDist_Change
End Sub

Public Sub ProdutoAlmoxDist_GotFocus()
'distribuicao
    Call gobjDistribuicao.ProdutoAlmoxDist_GotFocus
End Sub

Public Sub ProdutoAlmoxDist_KeyPress(KeyAscii As Integer)
'distribuicao
    Call gobjDistribuicao.ProdutoAlmoxDist_KeyPress(KeyAscii)
End Sub

Public Sub ProdutoAlmoxDist_Validate(Cancel As Boolean)
'distribuicao
    Call gobjDistribuicao.ProdutoAlmoxDist_Validate(Cancel)
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

'Distribuicao
Public Property Get GridDist() As Object
     Set GridDist = objUserControl.Controls("GridDist")
End Property

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

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134016

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134016

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156390)

    End Select
    
    Exit Sub

End Sub

Private Sub Fornecedor_Preenche()
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134057

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134057

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156391)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156392)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156393)
     
    End Select
     
    Exit Sub

End Sub

Public Function NFDev_AlocaAlmoxarifadoPadrao(ByVal objCT As Object, ByVal objDistribuicao As Object) As Long
    NFDev_AlocaAlmoxarifadoPadrao = SUCESSO
End Function

Public Function NFDev_Form_Load(ByVal objCT As Object) As Long
    NFDev_Form_Load = SUCESSO
End Function

Public Function NFDev_Form_UnLoad(ByVal objCT As Object) As Long
    NFDev_Form_UnLoad = SUCESSO
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

Public Function NFDEV_Calcula_PesoKg(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    NFDEV_Calcula_PesoKg = SUCESSO
End Function

Public Function NFDEV_Calcula_PesoKg_Total(ByVal objCT As Object) As Long
    NFDEV_Calcula_PesoKg_Total = SUCESSO
End Function

Public Function NFDEV_Calcula_Valores(ByVal objCT As Object, dQuantidade As Double, ByVal iLinha As Integer) As Long
    NFDEV_Calcula_Valores = SUCESSO
End Function

Public Function NFDEV_Move_GridItens_Memoria(ByVal objCT As Object, ByVal objItemNF As ClassItemNF, ByVal iLinha As Integer) As Long
    NFDEV_Move_GridItens_Memoria = SUCESSO
End Function

Public Function NFDEV_Preenche_GridItens(ByVal objCT As Object, ByVal objItemNF As ClassItemNF, ByVal iLinha As Integer) As Long
    NFDEV_Preenche_GridItens = SUCESSO
End Function

Public Function NFDEV_ProdutoLinha_Preenche(ByVal objCT As Object, objProduto As ClassProduto) As Long
    NFDEV_ProdutoLinha_Preenche = SUCESSO
End Function

Public Function NFDEV_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iLocalChamada As Integer) As Long
    NFDEV_Rotina_Grid_Enable = SUCESSO
End Function

Public Function NFDEV_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    NFDEV_Saida_Celula = SUCESSO
End Function

Public Function NFDEV_ValorTotal(ByVal objItemNF As ClassItemNF) As Long
    NFDEV_ValorTotal = SUCESSO
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

Public Function NFDEV_Produto_Saida_Celula(ByVal objCT As Object) As Long

Dim iIndice As Integer

On Error GoTo Erro_NFDEV_Produto_Saida_Celula

    'Verifica se já está em outra linha do Grid
    For iIndice = 1 To objCT.objGridItens.iLinhasExistentes
        If iIndice <> objCT.GridItens.Row Then
            If objCT.GridItens.TextMatrix(iIndice, objCT.iGrid_Produto_Col) = objCT.Produto.Text Then
                gError 177164
                Exit For
            End If
        End If
    Next

    NFDEV_Produto_Saida_Celula = SUCESSO
    
    Exit Function

Erro_NFDEV_Produto_Saida_Celula:

    NFDEV_Produto_Saida_Celula = gErr

    Select Case gErr
    
        Case 177164
            If gobjCRFAT.iProdRepetidoNF = PRODUTO_REPETICAO_NAO_PERMITE Then
                Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_EXISTENTE3", gErr, Produto.Text, iIndice)
            ElseIf gobjCRFAT.iProdRepetidoNF = PRODUTO_REPETICAO_AVISA Then
                If Rotina_Aviso(vbYesNo, "AVISO_PRODUTO_JA_EXISTENTE", Produto.Text, iIndice) = vbYes Then
                    Resume Next
                End If
            Else
                Resume Next
            End If
            
         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177165)

    End Select

    Exit Function
    
End Function

Public Function NFDEV_Produto_Saida_Celula2(ByVal objCT As Object) As Long
    NFDEV_Produto_Saida_Celula2 = SUCESSO
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

Public Sub BotaoItensNFOrigAuto_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoItensNFOrigAuto_Click

    lErro = CF2(Me, "NFDEV_BotaoItensNFOrigAuto_Click")
    If lErro <> SUCESSO Then gError 199155
    
    Exit Sub

Erro_BotaoItensNFOrigAuto_Click:

    Select Case gErr
    
        Case 199155

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199156)

    End Select

    Exit Sub
 
End Sub

Private Sub Trata_CodigoBarras()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoEnxuto As String

On Error GoTo Erro_Trata_CodigoBarras

        If objGridItens.iLinhasExistentes + 1 = GridItens.Row Then
        
            'Verifica se o Produto está preenchido
            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then
        
                Call Chama_Tela_Modal("CodigoBarras", objProduto)
        
                If objProduto.sCodigoBarras <> "Cancel" Then
        
                    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
                    If lErro <> SUCESSO Then gError 199065
            
                    Produto.PromptInclude = False
                    Produto.Text = sProdutoEnxuto
                    Produto.PromptInclude = True
                    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
            
                    lErro = Produto_Saida_Celula(objProduto)
                    If lErro <> SUCESSO Then gError 199066
        
                End If
        
            End If
        
        End If

    Exit Sub

Erro_Trata_CodigoBarras:

    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    Select Case gErr

        Case 199065
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 199066

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199067)

    End Select

    Exit Sub

End Sub

Public Function NFDEV_BotaoItensNFOrigAuto_Click(ByVal objCT As Object) As Long

Dim lErro As Long
Dim dQuantResto As Double
Dim sProduto As String
Dim dQuantidade As Double
Dim ColItensNF As New Collection
Dim sUM As String
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim objItemNF As ClassItemNF
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer
Dim iProdutoPreenchido As Integer
Dim objCliente As New ClassCliente
Dim lCliente As Long
Dim iFilialCli As Integer
Dim iLinhaOriginal As Integer
Dim objFornecedor As New ClassFornecedor
Dim lFornecedor As Long
Dim iFilialForn As Integer
    
On Error GoTo Erro_NFDEV_BotaoItensNFOrigAuto_Click

    If GridItens.Row <> objGridItens.iLinhasExistentes Then gError 199044

    dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
    sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

    If dQuantidade = 0 Then gError 199045
    
    If Len(Trim(sUM)) = 0 Then gError 199046

    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 199157

    objProduto.sCodigo = sProduto
    
    'ler dados do produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 199047
    
    If lErro <> SUCESSO Then gError 199048
    
    'converter qtde a faturar p/UMEstoque do produto
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 199049

    dQuantResto = dQuantidade * dFator

    If iFornecedorAtivo <> FORN_ATIVO Then
        
        If Len(Trim(Cliente.ClipText)) = 0 Then gError 199053
    
        objCliente.sNomeReduzido = Cliente.Text
    
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 199054
    
        'Se não encontrou o Cliente --> erro
        If lErro = 12348 Then gError 199055
    
        lCliente = objCliente.lCodigo
    
        iFilialCli = Codigo_Extrai(Filial.Text)
        
        lErro = CF("Retornar_ItensNF_Orig", ColItensNF, sProduto, lCliente, iFilialCli, dQuantResto)
        If lErro <> SUCESSO Then gError 199050

    Else
        
        If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 39223
    
        objFornecedor.sNomeReduzido = Fornecedor.Text

        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 39066

        'Não encontrou o Fornecedor --> erro
        If lErro = 6681 Then gError 39067

        lFornecedor = objFornecedor.lCodigo
        iFilialForn = Codigo_Extrai(Filial.Text)
        
        lErro = CF("Retornar_ItensNF_Orig_Forn", ColItensNF, sProduto, lFornecedor, iFilialForn, dQuantResto)
        If lErro <> SUCESSO Then gError 199050

    End If

    iIndice = GridItens.Row
    iLinhaOriginal = GridItens.Row

    For Each objItemNF In ColItensNF

        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = GridItens.TextMatrix(iLinhaOriginal, iGrid_Produto_Col)
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = GridItens.TextMatrix(iLinhaOriginal, iGrid_DescProduto_Col)
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = GridItens.TextMatrix(iLinhaOriginal, iGrid_UnidadeMed_Col)
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemNF.dQuantidade / dFator)
'        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(objItemNF.dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(objItemNF.dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
        GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(objItemNF.dPercDesc, "Percent")
        GridItens.TextMatrix(iIndice, iGrid_CodNFOrig_Col) = objItemNF.lNumNFOrig
        GridItens.TextMatrix(iIndice, iGrid_SerieNFOrig_Col) = objItemNF.sSerieNFOrig
        GridItens.TextMatrix(iIndice, iGrid_ItemNFOrig_Col) = objItemNF.iItemNFOrig
        
        'altera as quantidades no tab de distribuicao
        lErro = gobjDistribuicao.Distribuicao_Processa()
        If lErro <> SUCESSO Then gError 199158
        
        If iIndice <> iLinhaOriginal Then objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
        
        If iIndice = iLinhaOriginal Then

            Call PrecoTotal_Calcula(iIndice, False)

            lErro = gobjTribTab.Alteracao_Item_Grid(iIndice)
            If lErro <> SUCESSO Then gError 199051
        
        Else

            gobjNFiscal.ColItensNF.Add1 objItemNF
            
            gobjNFiscal.ColItensNF(GridItens.Row).sUMEstoque = objProduto.sSiglaUMEstoque
            gobjNFiscal.ColItensNF(GridItens.Row).iItem = iIndice
            gobjNFiscal.ColItensNF(GridItens.Row).sProduto = objProduto.sCodigo
            gobjNFiscal.ColItensNF(GridItens.Row).sUnidadeMed = GridItens.TextMatrix(iLinhaOriginal, iGrid_UnidadeMed_Col)
    
            Call PrecoTotal_Calcula(iIndice, False)

            ''*** incluidos p/tratamento de tributacao *******************************
            lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, sProduto)
            If lErro <> SUCESSO Then gError 199056
                '*** fim tributacao
    
        End If
    
        iIndice = iIndice + 1
        
        GridItens.Row = iIndice
    
    Next
    
    Call ValorTotal_Calcula
    
    If dQuantResto > 0 Then vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_QUANT_A_SER_DEVOLVIDA_INSUFICIENTE", dQuantResto / dFator)
    
    NFDEV_BotaoItensNFOrigAuto_Click = SUCESSO
    
    Exit Function

Erro_NFDEV_BotaoItensNFOrigAuto_Click:

    NFDEV_BotaoItensNFOrigAuto_Click = gErr
    
    Select Case gErr
    
        Case 199044
            Call Rotina_Erro(vbOKOnly, "NAO_E_ULTIMA_LINHA_GRID", gErr)

        Case 199045
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_QUANTIDADE_ZERADA", gErr)

        Case 199046
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_UM_NAO_PREENCHIDO", gErr)
            
        Case 199047, 199049, 199050, 199051, 199054, 199056, 199157, 199158, 39066
        
        Case 199048
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 199053
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 199055
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)

        Case 39223
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 39067
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

         Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199052)

    End Select

    Exit Function

End Function
    
Public Function NFDEV_Preenche_GridDistribuicao(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_NFDEV_Preenche_GridDistribuicao

    'distribuicao
    'Preenche o Grid com as Distribuições dos itens da Nota Fiscal
    lErro = gobjDistribuicao.Preenche_GridDistribuicao(objNFiscal)
    If lErro <> SUCESSO Then gError 199175

    NFDEV_Preenche_GridDistribuicao = SUCESSO
    
    Exit Function
    
Erro_NFDEV_Preenche_GridDistribuicao:

    NFDEV_Preenche_GridDistribuicao = gErr
    
    Select Case gErr
    
        Case 199175
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 199176)

    End Select

    Exit Function

End Function

Public Sub ValorTotal_Change()
    
    If InStr(UCase(gsNomeEmpresa), "SHOWLUZ") <> 0 Then
    
        If Abs(StrParaDbl(Total.Text) - StrParaDbl(ValorTotal.Caption)) > DELTA_VALORMONETARIO Then
            If Len(Trim(ValorTotal.Caption)) = 0 Then
                Total.Text = ""
            Else
                Total.Text = Format(StrParaDbl(ValorTotal.Caption), "Standard")
            End If
        End If
    
    End If
    
End Sub

Public Sub Total_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_Total_Validate
    
    If InStr(UCase(gsNomeEmpresa), "SHOWLUZ") <> 0 Then
    
        'Verifica se o Valor está preenchido
        If Len(Trim(Total.Text)) > 0 Then
    
            'Faz a Crítica do Valor digitado
            lErro = Valor_NaoNegativo_Critica(Total.Text)
            If lErro <> SUCESSO Then gError 39203
    
        End If

        If Abs(StrParaDbl(Total.Text) - StrParaDbl(ValorTotal.Caption)) > DELTA_VALORMONETARIO Then
        
            If Len(Trim(Total.Text)) = 0 Then
                ValorTotal.Caption = ""
            Else
                ValorTotal.Caption = Format(StrParaDbl(Total.Text), "Standard")
            End If
            
        End If
    
    End If
    
    Exit Sub

Erro_Total_Validate:

    Cancel = True

    Select Case gErr

        Case 39203

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156366)

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

Private Function Carrega_TipoFrete() As Long

Dim lErro As Long
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer
Dim colTipoFrete As Collection

On Error GoTo Erro_Carrega_TipoFrete
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207429

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
        Case 207429

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207430)

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
    Call ValorTotal_Calcula
        
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

Public Sub BotaoOPBenef_Click()

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim iIndice As Integer

On Error GoTo Erro_BotaoOPBenef_Click

    'Se é uma devolução do material usado no beneficiamento
    If Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_NFISDBFNE Then
    
        gobjNFDevBenef.dtData = StrParaDate(DataEmissao.Text)
        If gobjNFDevBenef.dtData = DATA_NULA Then gobjNFDevBenef.dtData = gdtDataAtual
        
        If gobjNFDevBenef.lCliente <> 0 And gobjNFDevBenef.iFilialCliente <> 0 Then
    
            lErro = Chama_Tela_Modal("NFDevBenefOP", gobjNFDevBenef)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If gobjNFDevBenef.iRetorno = vbOK Then
            
                'Limpa os dados dos itens atuais do grid
                For iIndice = objGridItens.iLinhasExistentes To 1 Step -1
                    gobjNFiscal.ColItensNF.Remove iIndice
                    Call gobjTelaProjetoInfo.Remove_Linha(iIndice)
                    lErro = gobjTribTab.Exclusao_Item_Grid(iIndice)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                Next
                Call Grid_Limpa(objGridItens)
            
                For Each objItemNF In gobjNFDevBenef.colItens
                    gobjNFiscal.ColItensNF.Add1 objItemNF
                Next
                
                'Preenhce o Grid com os Ítens da Nota Fiscal em questão
                lErro = Preenche_GridItens(gobjNFiscal)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                For iIndice = 1 To objGridItens.iLinhasExistentes
                    
                    Call gobjTelaProjetoInfo.Adiciona_Linha(iIndice)
                    
                    Set objItemNF = gobjNFiscal.ColItensNF(iIndice)
                    
                    'permite que a tributacao reflita a inclusao de uma linha no grid
                    lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, objItemNF.sProduto)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    Call PrecoTotal_Calcula(iIndice, False)
                    
                Next
    
                'Recalcula o peso bruto e o peso líquido
                lErro = Pesos_Recalcula()
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Call ValorTotal_Calcula
                
            End If
            
        End If
        
    End If
        
    Exit Sub

Erro_BotaoOPBenef_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209956)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoItensNFOrigDet_Click()

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim iIndice As Integer

On Error GoTo Erro_BotaoItensNFOrigDet_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridItens.Row = 0 Then gError 209957
    
    If GridItens.Row <= gobjNFiscal.ColItensNF.Count Then
    
        Set objItemNF = gobjNFiscal.ColItensNF(GridItens.Row)
        
        If objItemNF.colItensNFDevBenefItemOP.Count = 0 And objItemNF.colItensNFDevBenefRem.Count = 0 Then gError 209958 'Não existe dados a serem exibidos
    
        lErro = Chama_Tela_Modal("NFDevBenefRem", objItemNF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
        
    Exit Sub

Erro_BotaoItensNFOrigDet_Click:

    Select Case gErr

        Case 209957
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 209958
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_DADOS_A_EXIBIR", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209959)

    End Select

    Exit Sub
    
End Sub

Private Sub Trata_FilialCli_Dev(ByVal lCliente As Long, ByVal iFilial As Integer)

Dim objItemNF As ClassItemNF

On Error GoTo Erro_Trata_FilialCli_Dev

    If gobjFAT.iTrataNFDevSimbPorOP = MARCADO And Not gbCarregandoTela Then
    
        'Se trocou o cliente ou filial limpa os dados calculados anteriormente
        If gobjNFDevBenef.lCliente <> lCliente Or gobjNFDevBenef.iFilialCliente <> iFilial Then
            Set gobjNFDevBenef = New ClassNFDevBenef
            gobjNFDevBenef.lCliente = lCliente
            gobjNFDevBenef.iFilialCliente = iFilial
            For Each objItemNF In gobjNFiscal.ColItensNF
                Set objItemNF.colItensNFDevBenefItemOP = New Collection
                Set objItemNF.colItensNFDevBenefRem = New Collection
            Next
            
            'Se é uma devolução do material usado no beneficiamento
            If Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_NFISDBFNE Then
                   
                'Se é para sugerir o cálculo automaticamente
                'abre a tela com as OPs
                If gobjFAT.iNFDevSimbSugOP = MARCADO Then
                    Call BotaoOPBenef_Click
                End If
                
            End If
        End If
    End If
    
    Exit Sub

Erro_Trata_FilialCli_Dev:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209992)

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
