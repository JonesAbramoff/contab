VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalEntDev"
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

'Property Variables:
Dim m_Caption As String
Event Unload()
Dim FrameL(1 To 5) As Frame

'Flag que indica se a tela está sendo preenchida.
Public gbCarregandoTela As Boolean
Public giPosCargaOk As Integer

Public gbGravando As Boolean
Dim giComplemento As Integer

'************ GRADE **********************
Public gobjNFiscal As ClassNFiscal
'*****************************************

'###########Tributação##############
Private Const ISS_VALOR As String = "ISS_Valor"
Private Const ISS_INCLUSO As String = "ISS_Incluso"
Private Const VALOR_IRRF As String = "Valor_IRRF"
Private Const VALOR_INSS As String = "Valor_INSS"
Private Const INSS_RETIDO As String = "INSS_Retido"
Private Const ICMS_CREDITO_VALOR As String = "ICMS_Credito_Valor"
Private Const IPI_CREDITO_VALOR As String = "IPI_Credito_Valor"
Private Const ICMS_CREDITO_FRETE As String = "ICMS_Credito_Frete"
Private Const ICMS_CREDITO_SEGURO As String = "ICMS_Credito_Seguro"
Private Const ICMS_CREDITO_OUTRASDESP As String = "ICMS_Cred_OutrasDesp"
Private Const ICMS_CREDITO_DESCONTO As String = "ICMS_Cred_Desconto"
Private Const IPI_CREDITO_FRETE As String = "IPI_Credito_Frete"
Private Const IPI_CREDITO_SEGURO As String = "IPI_Credito_Seguro"
Private Const IPI_CREDITO_OUTRASDESP As String = "IPI_Cred_OutrasDesp"
Private Const IPI_CREDITO_DESCONTO As String = "IPI_Cred_Desconto"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque"
Private Const IPI_PRODUTO As String = "IPI_Produto"

'fim dos mneumonicos da tributação

'inicio contabilidade
Public objGrid1 As AdmGrid
Public objContabil As New ClassContabil

Public WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Public WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

'mnemonicos
Private Const FILIAL_CONTA As String = "FilialCli_Conta_Ctb"
Private Const ALMOXARIFADO1 As String = "Almoxarifado"
Private Const CLIENTE1 As String = "Cliente"
Private Const DATA_EMISSAO As String = "Data_Emissao"
Private Const DATA_ENTRADA As String = "Data_Entrada"
Private Const DESCONTO1 As String = "Desconto_Produto"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"
Private Const FILIAL1 As String = "Filial"
Private Const FORNECEDOR1 As String = "Fornecedor"
Private Const NATUREZA_OP As String = "Natureza_OP"
Private Const NFISCAL1 As String = "Nota_Fiscal"
Private Const PRODUTO1 As String = "Produto_Codigo"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const SERIE1 As String = "Serie"
Private Const TIPO_NOTA_FISCAL As String = "Tipo_Nota_Fiscal"
Private Const UNIDADE_MED As String = "Unidade_Med"
Private Const VALOR_TOTAL As String = "Valor_Total_Produto"
Private Const VALOR_UNITARIO As String = "Valor_Unitario"
Private Const PRODUTO_ITENS As String = "Produto_Itens"
Private Const PRODUTO_DIST As String = "Produto_Dist"
Private Const QUANT_DIST  As String = "Quant_Dist"
Private Const FATOR_ITEM_EST As String = "F1"
Private Const PRODUTO_TIPO_EST As String = "Prod_Dist_Tipo"
Private Const PRODUTO_TIPO As String = "Prod_Tipo"
'#######Valores####################
Private Const ICMS As String = "ICMS_Valor"
Private Const ICMSSUBST As String = "ICMSSubst_Valor"
Private Const SUBTOTAL1 As String = "Valor_Produtos"
Private Const VALOR_FRETE As String = "Valor_Frete"
Private Const VALOR_SEGURO As String = "Valor_Seguro"
Private Const VALOR_DESPESAS As String = "Valor_Despesas"
Private Const IPI As String = "IPI_Valor"
Private Const VALOR_DESCONTO As String = "Valor_Desconto_Glb"
Private Const TOTAL1 As String = "Total_Nota_Fiscal"
Private Const CTACONTABILEST1 As String = "ContaContabilEst"
Private Const CTAFILIALFORNECEDOR1 As String = "FilialForn_Conta_Ctb"
Private Const QUANT_DISPONIVEL As String = "Quant_Disponivel"
Private Const QUANT_CONSIGNADA As String = "Quant_Consignada"
Private Const QUANT_CONSERTO As String = "Quant_Conserto"
Private Const QUANT_DEMONS As String = "Quant_Demons"
Private Const QUANT_OUTRAS As String = "Quant_Outras"
Private Const ICMS_CREDITO_TOTAL As String = "ICMS_Credito_Total"
Private Const IPI_CREDITO_TOTAL As String = "IPI_Credito_Total"
Private Const PIS_CREDITO_TOTAL As String = "PIS_Credito_Total"
Private Const COFINS_CREDITO_TOTAL As String = "COFINS_Credito_Total"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const COFINS_RETIDO As String = "COFINS_Retido"
Private Const CSLL_RETIDO As String = "CSLL_Retido"
Private Const CTACONTABILDIST As String = "ContaContabilDist"
Private Const VALOR_UNIT_DIST As String = "Valor_Unit_Dist"
Private Const ICMS_CRED_VALOR_DIST As String = "ICMS_Cred_Valor_Dist"
Private Const IPI_CRED_VALOR_DIST As String = "IPI_Cred_Valor_Dist"
Private Const IPI_PRODUTO_DIST As String = "IPI_Produto_Dist"
Private Const QUANT_DIST1 As String = "Quant_Dist1"
Private Const CONTROLA_EST As String = "Controla_Est"
Private Const CONTA_DESP_EST_FORN As String = "Conta_Desp_Est_Forn"
Private Const CONTACONTABILAPLIC As String = "ContaContabilAplic"
Private Const CONTACONTABILPROD As String = "ContaContabilProd"
Private Const CONTACONTABILCATPROD As String = "ContaContabilCatProd"
'fim da contabilidade

Public gobjTelaProjetoInfo As ClassTelaPRJInfo

'Rastreamento
'Public gcolItensNF As Collection
'Public objGenerico As AdmGenerico
Public gobjRastreamento As ClassRastreamento

'distribuicao
Public gobjDistribuicao As Object

Public iAlterado As Integer
Public iFornecedorAlterado As Integer
Public iClienteAlterado As Integer
Public iNaturezaOpAlterada As Integer
Public iFrameAtual As Integer
Private iEmitenteNF As Integer
Public iValorDescontoAlterado As Integer
Public iValorFreteAlterado As Integer
Public iValorSeguroAlterado As Integer
Public iValorDespesasAlterado As Integer
Public iDataEntradaAlterada As Integer
Public iSerieAlterada As Integer

Public objGridItens As AdmGrid
Public iGrid_Produto_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
'distribuicao
'Public iGrid_Almoxarifado_Col As Integer
Public iGrid_ValorUnitario_Col As Integer
Public iGrid_ValorUnitarioMoeda_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PrecoTotal_Col As Integer
Public iGrid_PrecoTotalB_Col As Integer
Public iGrid_SerieNFOrig_Col As Integer
Public iGrid_NumNFOrig_Col As Integer
Public iGrid_ItemNFOrig_Col As Integer
Public iGrid_Ccl_Col As Integer

'Grid relacionado com Beneficiamento
Public objGridBenef As AdmGrid

Dim iGrid_ItemNFBenef_Col As Integer
Dim iGrid_ProdutoBenef_Col As Integer
Dim iGrid_DescricaoItemBenef_Col As Integer
Dim iGrid_AlmoxarifadoBenef_Col As Integer
Dim iGrid_QuantInsumoBenef_Col As Integer
Dim iGrid_ContaContabilEstBenef_Col As Integer
Dim iGrid_ContaContabilProducaoBenef_Col As Integer

Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoProdutoBenef As AdmEvento
Attribute objEventoProdutoBenef.VB_VarHelpID = -1
Private WithEvents objEventoEstoqueBenef As AdmEvento
Attribute objEventoEstoqueBenef.VB_VarHelpID = -1

Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Public WithEvents objEventoNFiscal As AdmEvento
Attribute objEventoNFiscal.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoLocalizacao As AdmEvento
Attribute objEventoLocalizacao.VB_VarHelpID = -1
Private WithEvents objEventoRecebimento As AdmEvento
Attribute objEventoRecebimento.VB_VarHelpID = -1
Private WithEvents objEventoSerieOriginal As AdmEvento
Attribute objEventoSerieOriginal.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoNFiscalOriginal As AdmEvento
Attribute objEventoNFiscalOriginal.VB_VarHelpID = -1
Private WithEvents objEventoItensNFiscalOriginal As AdmEvento
Attribute objEventoItensNFiscalOriginal.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoChvNFe As AdmEvento
Attribute objEventoChvNFe.VB_VarHelpID = -1

Dim gsChvAnt As String

''*** incluidos p/tratamento de tributacao *******************************

Public gobjTribTab As ClassTribTab

'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_Itens = 2
Private Const TAB_Complemento = 3
'distribuicao
Private Const TAB_Distribuicao = 4
Private Const TAB_Tributacao = 5
Private Const TAB_Contabilizacao = 6
'Rastreamento
Private Const TAB_RASTRO = 7
Private TAB_BENEFICIAMENTO As Integer

'--- inicio dos properties get dos controles da tela
Public Property Get ChvNFe() As Object
     Set ChvNFe = objUserControl.Controls("ChvNFe")
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

Public Property Get MensagemCorpo() As Object
     Set MensagemCorpo = objUserControl.Controls("MensagemCorpo")
End Property

Public Property Get EletronicaFed() As Object
     Set EletronicaFed = objUserControl.Controls("EletronicaFed")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get Serie() As Object
     Set Serie = objUserControl.Controls("Serie")
End Property

Public Property Get TipoNFiscal() As Object
     Set TipoNFiscal = objUserControl.Controls("TipoNFiscal")
End Property

Public Property Get Recebimento() As Object
     Set Recebimento = objUserControl.Controls("Recebimento")
End Property

Public Property Get RecebimentoLabel() As Object
     Set RecebimentoLabel = objUserControl.Controls("RecebimentoLabel")
End Property

Public Property Get BotaoLimparNF() As Object
     Set BotaoLimparNF = objUserControl.Controls("BotaoLimparNF")
End Property

Public Property Get NFiscal() As Object
     Set NFiscal = objUserControl.Controls("NFiscal")
End Property

Public Property Get NaturezaOp() As Object
     Set NaturezaOp = objUserControl.Controls("NaturezaOp")
End Property

Public Property Get NumRecebimento() As Object
     Set NumRecebimento = objUserControl.Controls("NumRecebimento")
End Property

Public Property Get NaturezaLabel() As Object
     Set NaturezaLabel = objUserControl.Controls("NaturezaLabel")
End Property

Public Property Get SerieLabel() As Object
     Set SerieLabel = objUserControl.Controls("SerieLabel")
End Property

Public Property Get NFiscalLabel() As Object
     Set NFiscalLabel = objUserControl.Controls("NFiscalLabel")
End Property

Public Property Get NFiscalInterna() As Object
     Set NFiscalInterna = objUserControl.Controls("NFiscalInterna")
End Property

Public Property Get SerieNFiscalOriginal() As Object
     Set SerieNFiscalOriginal = objUserControl.Controls("SerieNFiscalOriginal")
End Property

Public Property Get BotaoExibirDados() As Object
     Set BotaoExibirDados = objUserControl.Controls("BotaoExibirDados")
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

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get Fornecedor() As Object
     Set Fornecedor = objUserControl.Controls("Fornecedor")
End Property

Public Property Get FornecedorLabel() As Object
     Set FornecedorLabel = objUserControl.Controls("FornecedorLabel")
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

Public Property Get UpDownEntrada() As Object
     Set UpDownEntrada = objUserControl.Controls("UpDownEntrada")
End Property

Public Property Get DataEntrada() As Object
     Set DataEntrada = objUserControl.Controls("DataEntrada")
End Property

'horaentrada
Public Property Get HoraEntrada() As Object
     Set HoraEntrada = objUserControl.Controls("HoraEntrada")
End Property

Public Property Get BotaoRecebimentos() As Object
     Set BotaoRecebimentos = objUserControl.Controls("BotaoRecebimentos")
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

Public Property Get TipoFrete() As Object
     Set TipoFrete = objUserControl.Controls("TipoFrete")
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

Public Property Get Mensagem() As Object
     Set Mensagem = objUserControl.Controls("Mensagem")
End Property

Public Property Get PesoLiquido() As Object
     Set PesoLiquido = objUserControl.Controls("PesoLiquido")
End Property

Public Property Get PesoBruto() As Object
     Set PesoBruto = objUserControl.Controls("PesoBruto")
End Property

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
End Property

Public Property Get BotaoItensNFiscalOrig() As Object
     Set BotaoItensNFiscalOrig = objUserControl.Controls("BotaoItensNFiscalOrig")
End Property

Public Property Get Total() As Object
     Set Total = objUserControl.Controls("Total")
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

Public Property Get SubTotal() As Object
     Set SubTotal = objUserControl.Controls("SubTotal")
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

Public Property Get UnidadeMed() As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
End Property

Public Property Get DescricaoItem() As Object
     Set DescricaoItem = objUserControl.Controls("DescricaoItem")
End Property

Public Property Get Ccl() As Object
     Set Ccl = objUserControl.Controls("Ccl")
End Property

Public Property Get ItemNFOriginal() As Object
     Set ItemNFOriginal = objUserControl.Controls("ItemNFOriginal")
End Property

Public Property Get NumNFOriginal() As Object
     Set NumNFOriginal = objUserControl.Controls("NumNFOriginal")
End Property

Public Property Get SerieNFOriginal() As Object
     Set SerieNFOriginal = objUserControl.Controls("SerieNFOriginal")
End Property

Public Property Get Desconto() As Object
     Set Desconto = objUserControl.Controls("Desconto")
End Property

Public Property Get PercentDesc() As Object
     Set PercentDesc = objUserControl.Controls("PercentDesc")
End Property

Public Property Get ValorUnitario() As Object
     Set ValorUnitario = objUserControl.Controls("ValorUnitario")
End Property

Public Property Get Quantidade() As Object
     Set Quantidade = objUserControl.Controls("Quantidade")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

Public Property Get ValorTotal() As Object
     Set ValorTotal = objUserControl.Controls("ValorTotal")
End Property

Public Property Get GridItens() As Object
     Set GridItens = objUserControl.Controls("GridItens")
End Property

Public Property Get BotaoCcls() As Object
     Set BotaoCcls = objUserControl.Controls("BotaoCcls")
End Property

Public Property Get BotaoConsultaDocCPR() As Object
     Set BotaoConsultaDocCPR = objUserControl.Controls("BotaoConsultaDocCPR")
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

Public Property Get LabelFilial() As Object
    Set LabelFilial = objUserControl.Controls("LabelFilial")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

'Rastreamento
Public Property Get GridRastro() As Object
     Set GridRastro = objUserControl.Controls("GridRastro")
End Property

'Distribuicao
Public Property Get GridDist() As Object
     Set GridDist = objUserControl.Controls("GridDist")
End Property

Public Property Get DescricaoItemBenef() As Object
     Set DescricaoItemBenef = objUserControl.Controls("DescricaoItemBenef")
End Property

Public Property Get BotaoProdutosBenef() As Object
     Set BotaoProdutosBenef = objUserControl.Controls("BotaoProdutosBenef")
End Property

Public Property Get BotaoPlanoConta() As Object
     Set BotaoPlanoConta = objUserControl.Controls("BotaoPlanoConta")
End Property

Public Property Get BotaoEstoqueBenef() As Object
     Set BotaoEstoqueBenef = objUserControl.Controls("BotaoEstoqueBenef")
End Property

Public Property Get ContaContabilProducaoBenef() As Object
     Set ContaContabilProducaoBenef = objUserControl.Controls("ContaContabilProducaoBenef")
End Property

Public Property Get ContaContabilEstBenef() As Object
     Set ContaContabilEstBenef = objUserControl.Controls("ContaContabilEstBenef")
End Property

Public Property Get AlmoxarifadoBenef() As Object
     Set AlmoxarifadoBenef = objUserControl.Controls("AlmoxarifadoBenef")
End Property

Public Property Get QuantInsumoBenef() As Object
     Set QuantInsumoBenef = objUserControl.Controls("QuantInsumoBenef")
End Property

Public Property Get ProdutoBenef() As Object
     Set ProdutoBenef = objUserControl.Controls("ProdutoBenef")
End Property

Public Property Get ItemNFBenef() As Object
     Set ItemNFBenef = objUserControl.Controls("ItemNFBenef")
End Property

Public Property Get GridMovimentos() As Object
     Set GridMovimentos = objUserControl.Controls("GridMovimentos")
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

Public Property Get NumNFe() As Object
     Set NumNFe = objUserControl.Controls("NumNFe")
End Property

Public Sub UnidadeMed_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UnidadeMed_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ItemNFOriginal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumNFOriginal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Desconto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercentDesc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorUnitario_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Quantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorTotal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub GridMovimentos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridBenef, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBenef, iAlterado)

    End If

End Sub

Public Sub GridMovimentos_EnterCell()
    Call Grid_Entrada_Celula(objGridBenef, iAlterado)
End Sub

Public Sub GridMovimentos_GotFocus()
    Call Grid_Recebe_Foco(objGridBenef)
End Sub

Public Sub GridMovimentos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer
Dim lNumIntDoc As Long
Dim lErro As Long

On Error GoTo Erro_GridMovimentos_KeyDown

    'Verifica se a Tecla apertada foi Del
    If KeyCode = vbKeyDelete Then
    
        'Guarda iLinhasExistentes
        iLinhasExistentesAnterior = objGridBenef.iLinhasExistentes

        'Guarda o índice da Linha a ser Excluída
        iLinhaAnterior = GridMovimentos.Row

    End If

    Call Grid_Trata_Tecla1(KeyCode, objGridBenef)

    Exit Sub

Erro_GridMovimentos_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156418)

    End Select

    Exit Sub

End Sub

Public Sub GridMovimentos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridBenef, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBenef, iAlterado)
    End If


End Sub

Public Sub GridMovimentos_LeaveCell()
    Call Saida_Celula(objGridBenef)
End Sub

Public Sub GridMovimentos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridBenef)
End Sub

Public Sub GridMovimentos_Scroll()
    Call Grid_Scroll(objGridBenef)
End Sub

Public Sub GridMovimentos_RowColChange()

Dim lErro As Long

On Error GoTo Erro_GridMovimentos_RowColChange

    Call Grid_RowColChange(objGridBenef)

    Exit Sub

Erro_GridMovimentos_RowColChange:

    Select Case gErr

        Case 126854

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156419)

    End Select

    Exit Sub

End Sub

Public Sub ItemNFBenef_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ItemNFBenef_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBenef)
End Sub

Public Sub ItemNFBenef_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBenef)
End Sub

Public Sub ItemNFBenef_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridBenef.objControle = ItemNFBenef
    lErro = Grid_Campo_Libera_Foco(objGridBenef)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ProdutoBenef_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ProdutoBenef_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBenef)
End Sub

Public Sub ProdutoBenef_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBenef)
End Sub

Public Sub ProdutoBenef_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridBenef.objControle = ProdutoBenef
    lErro = Grid_Campo_Libera_Foco(objGridBenef)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub AlmoxarifadoBenef_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub AlmoxarifadoBenef_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBenef)
End Sub

Public Sub AlmoxarifadoBenef_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBenef)
End Sub

Public Sub AlmoxarifadoBenef_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBenef.objControle = AlmoxarifadoBenef
    lErro = Grid_Campo_Libera_Foco(objGridBenef)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantInsumoBenef_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub QuantInsumoBenef_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBenef)
End Sub

Public Sub QuantInsumoBenef_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBenef)
End Sub

Public Sub QuantInsumoBenef_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBenef.objControle = QuantInsumoBenef
    lErro = Grid_Campo_Libera_Foco(objGridBenef)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ContaContabilEstBenef_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ContaContabilEstBenef_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBenef)
End Sub

Public Sub ContaContabilEstBenef_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBenef)
End Sub

Public Sub ContaContabilEstBenef_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBenef.objControle = ContaContabilEstBenef
    lErro = Grid_Campo_Libera_Foco(objGridBenef)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ContaContabilProducaoBenef_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ContaContabilProducaoBenef_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBenef)
End Sub

Public Sub ContaContabilProducaoBenef_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBenef)
End Sub

Public Sub ContaContabilProducaoBenef_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBenef.objControle = ContaContabilProducaoBenef
    lErro = Grid_Campo_Libera_Foco(objGridBenef)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Ccl_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Ccl_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub Ccl_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub BotaoCcls_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As ClassCcl

On Error GoTo Erro_BotaoCcls_Click

    If GridItens.Row = 0 Then gError 52838

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then gError 52839
    
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

    Exit Sub
    
Erro_BotaoCcls_Click:
    
    Select Case gErr
    
        Case 52838
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 52839
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156420)

    End Select
    
    Exit Sub

End Sub

Public Sub DataEmissao_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
End Sub

Public Sub DataEntrada_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataEntrada, iAlterado)
End Sub

Public Sub Destinatario_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoFrete_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NaturezaOp_GotFocus()
Dim iNaturezaAux As Integer
    
    iNaturezaAux = iNaturezaOpAlterada
    Call MaskEdBox_TrataGotFocus(NaturezaOp, iAlterado)
    iNaturezaOpAlterada = iNaturezaAux

End Sub

Public Sub NFiscal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NFiscal_GotFocus()
    Call MaskEdBox_TrataGotFocus(NFiscal, iAlterado)
End Sub

Public Sub NFiscalOriginal_GotFocus()
    Call MaskEdBox_TrataGotFocus(NFiscalOriginal, iAlterado)
End Sub

Private Sub Class_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
End Sub

Public Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) <> 0 And GridItens.Row <> 0 Then

        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 52840

        'Coloca o valor do Ccl na coluna correspondente
        GridItens.TextMatrix(GridItens.Row, iGrid_Ccl_Col) = sCclMascarado
                
        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 52840 'Tratado na rotina chamadora
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156421)

    End Select

    Exit Sub

End Sub

Public Sub DescricaoItem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NatOpInterna_Change()
    Call gobjTribTab.NatOpInterna_Change
End Sub

Public Sub NatOpInterna_Validate(Cancel As Boolean)
    Call gobjTribTab.NatOpInterna_Validate(Cancel)
End Sub

Public Sub NatOpInterna_GotFocus()
    Call gobjTribTab.NatOpInterna_GotFocus(iAlterado)
End Sub

Public Sub Total_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Total_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dTotal As Double

On Error GoTo Erro_Total_Validate

    'Se estiver preenchido
    If Len(Trim(Total.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(Total.Text)
        If lErro <> SUCESSO Then gError 43785

        dTotal = CDbl(Total.Text)

        'Coloca o valor formatado na tela
        Total.Text = Format(dTotal, "Standard")

    End If

    Exit Sub

Erro_Total_Validate:

    Cancel = True

    Select Case gErr

        Case 43785

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156422)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

Dim lErro As Long

On Error GoTo Erro_Form_Activate

    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError 59339
        
    Call TelaIndice_Preenche(Me)

    Exit Sub
     
Erro_Form_Activate:

    Select Case gErr
          
        Case 59339
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156423)
     
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

        ''*** incluidos p/tratamento de tributacao *******************************
        Call GridItens_InicializaColunas
        Set gobjTribTab = New ClassTribTab
        lErro = gobjTribTab.Ativar(Me, 1)
        If lErro <> SUCESSO Then gError 55664
    
        '*** fim tributacao
    
        'Rastreamento
        Set gobjRastreamento = New ClassRastreamento
        Set gobjRastreamento.objTela = Me
        
        'distribuicao
        Set gobjDistribuicao = CreateObject("RotinasMat.ClassMATDist")
        Set gobjDistribuicao.objTela = Me
        Set gobjDistribuicao.objRastreamento = gobjRastreamento
        
        'Carrega as combos de Séries
        lErro = Carrega_Serie()
        If lErro <> SUCESSO Then gError 35151
    
        'Carrega as combos com os tipos de Notas Fiscais
        lErro = Carrega_TiposDocInfo()
        If lErro <> SUCESSO Then gError 35149
    
        'Inicializa as variáveis relacionadas c\ as telas de Browse
        Set objEventoSerie = New AdmEvento
        Set objEventoNFiscal = New AdmEvento
        Set objEventoCliente = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoNatureza = New AdmEvento
        Set objEventoFornecedor = New AdmEvento
        Set objEventoRecebimento = New AdmEvento
        Set objEventoLocalizacao = New AdmEvento
        Set objEventoSerieOriginal = New AdmEvento
        Set objEventoTransportadora = New AdmEvento
        Set objEventoNFiscalOriginal = New AdmEvento
        Set objEventoItensNFiscalOriginal = New AdmEvento
        Set objEventoCcl = New AdmEvento
        Set objEventoEstoqueBenef = New AdmEvento
        Set objEventoContaContabil = New AdmEvento
        Set objEventoProdutoBenef = New AdmEvento
        Set objEventoChvNFe = New AdmEvento
        
        '########################################
        'Inserido por Wagner 10/08/2006
        Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
        Set gobjTelaProjetoInfo.objUserControl = objUserControl
        Set gobjTelaProjetoInfo.objTela = Me
        '########################################
    
        'Carrega a combo Placa UF com os Estados cadastrados no BD
        lErro = Carrega_PlacaUF()
        If lErro <> SUCESSO Then gError 35153
    
        'Carrega a combo com o Código e o Nome Reduzido das Transportadoras do BD
        lErro = Carrega_Transportadora()
        If lErro <> SUCESSO Then gError 35155
    
        'Incluído por Luiz Nogueira em 21/08/03
        'Carrega a combo VolumeEspecie
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
        If lErro <> SUCESSO Then gError 102442
    
        'Incluído por Luiz Nogueira em 21/08/03
        'Carrega a combo VolumeMarca
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
        If lErro <> SUCESSO Then gError 102443
    
        'Inicializa a Máscara de Produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
        If lErro <> SUCESSO Then gError 35156
    
        'Inicializa a Mascara de Ccl
        lErro = Inicializa_MascaraCcl()
        If lErro <> SUCESSO Then gError 49385
    
        Quantidade.Format = FORMATO_ESTOQUE
    
        Set objGridItens = New AdmGrid
    
        'Faz a Inicialização do Grid de Itens
        lErro = Inicializa_Grid_Itens(objGridItens)
        If lErro <> SUCESSO Then gError 35157
    
        'Inicializa a contabilidade carregando apenas os modelos contábeis referentes ao TipoDocInfo default
        If objContabil.objGridInt Is Nothing Then

            'Obtém o código do subtipo contábil referente ao TipoDocInfo default
            lErro = CF("Obtem_Subtipo_Contabil", "TiposDocInfo", DOCINFO_NFIEDV, iSubTipo)
            If lErro <> SUCESSO And lErro <> 79671 Then gError 79690
            
            'Se não encontrou o subtipo => erro
            If lErro = 79671 Then gError 79691
            
            'Inicialização da parte de contabilidade
            lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_ESTOQUE, iSubTipo)
            If lErro <> SUCESSO Then gError 79367

        End If

        'Inicializa o grid de Rastreamento
        lErro = gobjRastreamento.Inicializa_Grid_Rastreamento()
        If lErro <> SUCESSO Then gError 83283

        'Inicializa o grid de Distribuicao
        lErro = gobjDistribuicao.Inicializa_GridDist()
        If lErro <> SUCESSO Then gError 89555

        Set objGridBenef = New AdmGrid
        
        'Inicializa a Máscara de Produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoBenef)
        If lErro <> SUCESSO Then gError 126913
                    
        'Inicializa mascara de contaContabilEst
        lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilEstBenef)
        If lErro <> SUCESSO Then gError 126914
        
        'Inicializa mascara de contaContabilEst
        lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilProducaoBenef)
        If lErro <> SUCESSO Then gError 126915
        
        'Faz a inicialização do grid de Movimentos
        lErro = Inicializa_GridMovimentos(objGridBenef)
        If lErro <> SUCESSO Then gError 126852

        gbCarregandoTela = False
        
        iAlterado = 0
        iDescontoAlterado = 0
    
        giPosCargaOk = 1
        
        Call ValorTotal_Calcula
        
        Set gobjNFiscal = New ClassNFiscal
    
    End If

    CargaPosFormLoad = SUCESSO
    
    Exit Function
     
Erro_CargaPosFormLoad:

    gbCarregandoTela = False
    
    CargaPosFormLoad = gErr
    
    Select Case gErr

        Case 35149, 35151, 35153, 35155, 35156, 35157, 39644, 55664, 79690, 83283, 89555, 102442, 102443, 126852, 126913, 126914, 126915

        Case 79691
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPODOCINFO_NAO_ENCONTRADO", gErr, DOCINFO_NFIEDV)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156424)
     
    End Select
     
    Exit Function

End Function

Public Function Inicializa_GridMovimentos(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Rastreamento

Dim iIndice As Integer
Dim iEscaninho As Integer
Dim lErro As Long

On Error GoTo Erro_Inicializa_GridMovimentos

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("ItemNF")
    objGridInt.colColuna.Add ("Produto Produzido")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Qtd. de Insumos Utilizados")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Conta Contábil Estoque")
    objGridInt.colColuna.Add ("Conta Contábil Produção")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ItemNFBenef.Name)
    objGridInt.colCampo.Add (ProdutoBenef.Name)
    objGridInt.colCampo.Add (DescricaoItemBenef.Name)
    objGridInt.colCampo.Add (QuantInsumoBenef.Name)
    objGridInt.colCampo.Add (AlmoxarifadoBenef.Name)
    objGridInt.colCampo.Add (ContaContabilEstBenef.Name)
    objGridInt.colCampo.Add (ContaContabilProducaoBenef.Name)

    'Colunas da Grid
    
    iGrid_ItemNFBenef_Col = 1
    iGrid_ProdutoBenef_Col = 2
    iGrid_DescricaoItemBenef_Col = 3
    iGrid_QuantInsumoBenef_Col = 4
    iGrid_AlmoxarifadoBenef_Col = 5
    iGrid_ContaContabilEstBenef_Col = 6
    iGrid_ContaContabilProducaoBenef_Col = 7
    
    'Grid do GridInterno
    objGridInt.objGrid = GridMovimentos

    'Largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGridInt.iLinhasVisiveis = 10
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridMovimentos = SUCESSO

    Exit Function

Erro_Inicializa_GridMovimentos:

    Inicializa_GridMovimentos = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156425)
        
    End Select
    
    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Inicializa o FrameAtual com o 1
    iFrameAtual = 1
        
    TAB_BENEFICIAMENTO = 8
    
    If giTipoVersao = VERSAO_LIGHT Then
        
        BotaoGrade.Visible = False
        BotaoCcls.Visible = False
        RecebimentoLabel.Visible = False
        NumRecebimento.Visible = False
        Recebimento.Visible = False
        
        'sumir com abas de contabilidade e rastro
        TabStrip1.Tabs.Remove (TAB_RASTRO)
        TabStrip1.Tabs.Remove (TAB_Contabilizacao)
        
        TAB_BENEFICIAMENTO = TAB_BENEFICIAMENTO - 2
        
    End If
    
    'inicializa a Data de Entrada com a Data Atual do Sistema
    DataEntrada.PromptInclude = False
    DataEntrada.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEntrada.PromptInclude = True
    
    gbGravando = False

    Set gobjNFiscal = New ClassNFiscal
    
    MsgAutomatica.Value = vbChecked
    Mensagem.MaxLength = MENSAGEM_NF_MAXLENGTH
    MensagemCorpo.MaxLength = MENSAGEM_NF_MAXLENGTH
    
    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207453
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 207453

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156426)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Function Inicializa_MascaraCcl() As Long
'Inicializa a mascara do centro de custo

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_MascaraCcl

    sMascaraCcl = String(STRING_CCL, 0)

    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError 49386

    Ccl.Mask = sMascaraCcl

    Inicializa_MascaraCcl = SUCESSO

    Exit Function

Erro_Inicializa_MascaraCcl:

    Inicializa_MascaraCcl = gErr

    Select Case gErr

        Case 49386

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156427)

    End Select

    Exit Function

End Function

Public Function Carrega_TiposDocInfo() As Long
'Carrega na os Tipo de Documentos relacionados com a tela

Dim lErro As Long
Dim colTipoDocInfo As New colTipoDocInfo
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iIndice As Integer

On Error GoTo Erro_Carrega_TiposDocInfo

    Set colTipoDocInfo = gobjCRFAT.colTiposDocInfo

    'Carrega na combo só os Tipos ligados essa tela
    For Each objTipoDocInfo In colTipoDocInfo
    
        If objTipoDocInfo.sNomeTelaNFiscal = Me.Name Then
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156428)

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
    If lErro <> SUCESSO Then gError 35152

    'Carrega nas combos
    For Each objSerie In colSerie
        If UCase(sSerieAnt) <> UCase(Desconverte_Serie_Eletronica(objSerie.sSerie)) Then Serie.AddItem Desconverte_Serie_Eletronica(objSerie.sSerie)
        SerieNFiscalOriginal.AddItem objSerie.sSerie
        sSerieAnt = Desconverte_Serie_Eletronica(objSerie.sSerie)
    Next

    If gobjEST.iSerieEletronicaPadrao = MARCADO Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    
    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case 35152

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156429)

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

            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156430)

    End Select

End Function

Private Function Carrega_Transportadora() As Long
'Carrega a combo todas as Transportadoras existentes

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_Transportadora

    'Lê o Código e o NOme de Todas as Transportadoras do BD
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then gError 35156

    'Carrega a combo de Transportadoras
    For Each objCodigoNome In colCodigoNome
        Transportadora.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Transportadora.ItemData(Transportadora.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_Transportadora = SUCESSO

    Exit Function

Erro_Carrega_Transportadora:

    Carrega_Transportadora = gErr

    Select Case gErr

        Case 35156

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156431)

    End Select

    Exit Function

End Function

Sub GridItens_InicializaColunas()
    Call CF2(Me, "NFEDEV_GridItens_InicializaColunas")
End Sub

Public Function NFEDEV_GridItens_InicializaColunas(ByVal objCT As Object) As Long
    
    'Colunas da Grid
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
    iGrid_Ccl_Col = 5
    iGrid_ValorUnitario_Col = 6
    iGrid_PercDesc_Col = 7
    iGrid_Desconto_Col = 8
    iGrid_PrecoTotal_Col = 9
    iGrid_PrecoTotalB_Col = 10
    iGrid_SerieNFOrig_Col = 11
    iGrid_NumNFOrig_Col = 12
    iGrid_ItemNFOrig_Col = 13

End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long

    Inicializa_Grid_Itens = CF2(Me, "NFEDEV_Inicializa_Grid_Itens", objGridInt)
    
End Function

Public Function NFEDEV_Inicializa_Grid_Itens(ByVal obCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = obCT

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Ccl")
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
    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (ValorUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (ValorTotal.Name)
    objGridInt.colCampo.Add (PrecoTotalB.Name)
    objGridInt.colCampo.Add (SerieNFOriginal.Name)
    objGridInt.colCampo.Add (NumNFOriginal.Name)
    objGridInt.colCampo.Add (ItemNFOriginal.Name)

    Call GridItens_InicializaColunas
    
    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_NF + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 5

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NFEDEV_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim iIndice As Integer
Dim objNFiscal As New ClassNFiscal
Dim colTabelaPrecoItem As New Collection
Dim objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "NFiscalEntradaDev"

    'Lê os dados da Tela Notas Fiscais a Pagar
    lErro = Move_NF_Memoria(objNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO Then gError 35158

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objNFiscal.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "TipoNFiscal", objNFiscal.iTipoNFiscal, 0, "TipoNFiscal"
    colCampoValor.Add "NaturezaOp", objNFiscal.sNaturezaOp, STRING_BUFFER_MAX_TEXTO, "NaturezaOp"
    colCampoValor.Add "Cliente", objNFiscal.lCliente, 0, "Cliente"
    colCampoValor.Add "FilialCli", objNFiscal.iFilialCli, 0, "FilialCli"
    colCampoValor.Add "Fornecedor", objNFiscal.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "FilialForn", objNFiscal.iFilialForn, 0, "FilialForn"
    colCampoValor.Add "Serie", objNFiscal.sSerie, STRING_BUFFER_MAX_TEXTO, "Serie"
    colCampoValor.Add "NumNotaFiscal", objNFiscal.lNumNotaFiscal, 0, "NumNotaFiscal"
    colCampoValor.Add "DataEmissao", objNFiscal.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "DataEntrada", objNFiscal.dtDataEntrada, 0, "DataEntrada"
'horaentrada
    colCampoValor.Add "HoraEntrada", CDbl(objNFiscal.dtHoraEntrada), 0, "HoraEntrada"
    colCampoValor.Add "DataVencimento", objNFiscal.dtDataVencimento, 0, "DataVencimento"
    colCampoValor.Add "ValorFrete", objNFiscal.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorSeguro", objNFiscal.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "ValorOutrasDespesas", objNFiscal.dValorOutrasDespesas, 0, "ValorOutrasDespesas"
    colCampoValor.Add "ValorDesconto", objNFiscal.dValorDesconto, 0, "ValorDesconto"
    colCampoValor.Add "ValorTotal", objNFiscal.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "CodTransportadora", objNFiscal.iCodTransportadora, 0, "CodTransportadora"
    colCampoValor.Add "Placa", objNFiscal.sPlaca, STRING_BUFFER_MAX_TEXTO, "Placa"
    colCampoValor.Add "PlacaUF", objNFiscal.sPlacaUF, STRING_BUFFER_MAX_TEXTO, "PlacaUF"
    colCampoValor.Add "VolumeQuant", objNFiscal.lVolumeQuant, 0, "VolumeQuant"
    colCampoValor.Add "VolumeEspecie", objNFiscal.lVolumeEspecie, 0, "VolumeEspecie" 'Alterado por Luiz Nogueira em 21/08/03
    colCampoValor.Add "VolumeMarca", objNFiscal.lVolumeMarca, 0, "VolumeMarca" 'Alterado por Luiz Nogueira em 21/08/03
    colCampoValor.Add "MensagemNota", objNFiscal.sMensagemNota, STRING_BUFFER_MAX_TEXTO, "MensagemNota"
    colCampoValor.Add "PesoLiq", objNFiscal.dPesoLiq, 0, "Pesoliq"
    colCampoValor.Add "PesoBruto", objNFiscal.dPesoBruto, 0, "PesoBruto"
    colCampoValor.Add "NumIntNotaOriginal", objNFiscal.lNumIntNotaOriginal, 0, "NumIntNotaOriginal"
    colCampoValor.Add "ValorProdutos", objNFiscal.dValorProdutos, 0, "ValorProdutos"
    colCampoValor.Add "FreteRespons", objNFiscal.iFreteRespons, 0, "FreteRespons"
    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_BUFFER_MAX_TEXTO, "VolumeNumero"
    colCampoValor.Add "Status", objNFiscal.iStatus, 0, "Status"

    'Adiciona filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
'    colSelecao.Add "Status", OP_DIFERENTE, STATUS_CANCELADO

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 35158

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156432)

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

        'Coloca os dados da NFiscal na tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 35159

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 35159

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156433)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO

    'Verifica se alguma filial foi selecionada
    If Filial.ListIndex = -1 Then Exit Sub
     
    If Cliente.Visible = True Then
        Call Trata_FilialCliente
    Else
        Call Trata_FilialForn
    End If
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    Call CF2(Me, "NFEnt_Form_Unload", Cancel)
    
    Set objGridItens = Nothing

    'Finaliza as variáveis das telas de Browse
    Set gobjNFiscal = Nothing
    Set objEventoSerie = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoCliente = Nothing
    Set objEventoProduto = Nothing
    Set objEventoNatureza = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoLocalizacao = Nothing
    Set objEventoRecebimento = Nothing
    Set objEventoSerieOriginal = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoNFiscalOriginal = Nothing
    Set objEventoItensNFiscalOriginal = Nothing
    Set objEventoCcl = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoEstoqueBenef = Nothing
    Set objEventoProdutoBenef = Nothing
    Set objEventoChvNFe = Nothing

    Set gobjTelaProjetoInfo = Nothing
    Set gobjInfoAdic = Nothing

    'Eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    Set objGridBenef = Nothing
    Set gobjTribTab = Nothing
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
'   Rastreamento
    Set gobjRastreamento = Nothing
    
    'distribuicao
    Set gobjDistribuicao = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas
    lErro = ComandoSeta_Liberar(Me.Name)

    Set gobjInfoUsu = Nothing

    Set gobjAnotacao = Nothing
    
    ''*** incluidos p/tratamento de tributacao *******************************
    If Not (gobjTribTab Is Nothing) Then Call gobjTribTab.Desativar
    Set gobjTribTab = Nothing
    '*** fim tributacao

End Sub

Public Sub BotaoLocalizacaoDist_Click()
'distribuicao
    Call gobjDistribuicao.BotaoLocalizacaoDist_Click
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156434)
    
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

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentesAnterior As Integer
Dim dSubTotal As Double
Dim iItemAtual As Integer
Dim iIndice As Integer

On Error GoTo Erro_GridItens_KeyDown

    iLinhasExistentesAnterior = objGridItens.iLinhasExistentes

    iItemAtual = GridItens.Row
    
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    If objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then
       
        GridItens.TextMatrix(GridItens.Row, 0) = GridItens.Row
        gobjNFiscal.ColItensNF.Remove GridItens.Row
                
        For iIndice = GridItens.Row To objGridItens.iLinhasExistentes
            gobjNFiscal.ColItensNF(iIndice).iItem = iIndice
        Next
                
        Call gobjTelaProjetoInfo.Remove_Linha(iItemAtual)
                
        For iLinhasExistentesAnterior = 1 To objGridItens.iLinhasExistentes
            If gobjNFiscal.ColItensNF(iLinhasExistentesAnterior).iPossuiGrade = MARCADO Then
                GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = "# " & iLinhasExistentesAnterior
            Else
                GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = iLinhasExistentesAnterior
            End If
            
        Next
        GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = iLinhasExistentesAnterior
                
        lErro = SubTotal_Calcula()
        If lErro <> SUCESSO Then gError 55596

        ''*** incluidos p/tratamento de tributacao *******************************
        lErro = gobjTribTab.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 55665
        '*** fim tributacao

        Call ValorTotal_Calcula
        
        If Not gbCarregandoTela Then
            lErro = gobjTribTab.Pesos_Recalcula()
            If lErro <> SUCESSO Then gError 177051
        End If

        lErro = CF2(Me, "NFEDEV_Calcula_PesoKg_Total")
        If lErro <> SUCESSO Then gError 177051

        'distribuicao
        lErro = gobjDistribuicao.Exclusao_Item_GridDist(iItemAtual)
        If lErro <> SUCESSO Then gError 89556

    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr

        Case 55596, 55665, 83284, 89556, 177051

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156435)

    End Select

    Exit Sub

End Sub

Public Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
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

Public Sub Mensagem_Change()
    iAlterado = REGISTRO_ALTERADO
    MsgAutomatica.Value = vbUnchecked
End Sub

Public Sub MensagemCorpo_Change()
    iAlterado = REGISTRO_ALTERADO
    MsgAutomatica.Value = vbUnchecked
End Sub

Public Sub NaturezaOp_Change()
    iAlterado = REGISTRO_ALTERADO
    iNaturezaOpAlterada = 1
End Sub

Public Sub NaturezaOp_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNaturezaOp As New ClassNaturezaOp

On Error GoTo Erro_NaturezaOp_Validate

    If iNaturezaOpAlterada = 0 Then Exit Sub
    
    'Verifica se Natureza de Operação foi preenchida
    If Len(Trim(NaturezaOp.ClipText)) = 0 Then Exit Sub

    objNaturezaOp.sCodigo = NaturezaOp.Text
    
    'Lê a Natureza de Operação
    lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
    If lErro <> SUCESSO And lErro <> 17958 Then gError 35716

    'Se não encontrou a Natureza de Operação --> erro
    If lErro = 17958 Then gError 35717

    'se for nota interna garantir que tenha codigo < 500
    If iEmitenteNF = DOCINFO_EMPRESA Then
    
        If objNaturezaOp.sCodigo >= NATUREZA_SAIDA_COD_INICIAL Then gError 56984

        If (gobjTribTab.sNatOpInterna <> objNaturezaOp.sCodigo) Then
        
            gobjTribTab.sNatOpInterna = objNaturezaOp.sCodigo
            Call NatOpInterna_Validate(Cancel)
            
        End If

    End If
    
    'se for nota externa garantir que tenha codigo > 500
    If iEmitenteNF <> DOCINFO_EMPRESA And objNaturezaOp.sCodigo <= NATUREZA_SAIDA_COD_INICIAL Then gError 56985

    iNaturezaOpAlterada = 0
    
    Exit Sub

Erro_NaturezaOp_Validate:

    Cancel = True


    Select Case gErr

        Case 35716

        Case 35717
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", gErr, NaturezaOp.Text)

        Case 56984
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_ENTRADA", gErr)
        
        Case 56985
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156436)

    End Select

    Exit Sub

End Sub

Public Sub NFiscalOriginal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub objEventoItensNFiscalOriginal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim sProduto As String
Dim objItemNFiscalOriginal As ClassItemNF
Dim objItemNF As New ClassItemNF
Dim objProduto As New ClassProduto

On Error GoTo Erro_objEventoItensNFiscalOriginal_evSelecao

    Set objItemNFiscalOriginal = obj1

    If GridItens.Row = 0 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objItemNFiscalOriginal.sProduto, sProduto)
    If lErro <> SUCESSO Then gError 55142
    
    If lErro = SUCESSO And Len(Trim(sProduto)) > 0 Then

        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True

        GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(GridItens.Row, iGrid_DescProduto_Col) = objItemNFiscalOriginal.sDescricaoItem
        GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objItemNFiscalOriginal.sUnidadeMed
        GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(objItemNFiscalOriginal.dQuantidade)
        GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col) = Format(objItemNFiscalOriginal.dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
        GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(objItemNFiscalOriginal.dPercDesc, "Percent")
        GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col) = Format(objItemNFiscalOriginal.dValorDesconto, "Standard")
        GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col) = Format(objItemNFiscalOriginal.dValorTotal, "Standard")

    End If

    If GridItens.Row > objGridItens.iLinhasExistentes Then
       
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
       
        objProduto.sCodigo = objItemNFiscalOriginal.sProduto
    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 56995
        
        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItens.Row)
    
        ''*** incluidos p/tratamento de tributacao *******************************
        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objItemNFiscalOriginal.sProduto)
        If lErro <> SUCESSO Then gError 56995
        '*** fim tributacao
        
        gobjNFiscal.ColItensNF.Add1 objItemNF
        gobjNFiscal.ColItensNF(GridItens.Row).sUMEstoque = objProduto.sSiglaUMEstoque
        gobjNFiscal.ColItensNF(GridItens.Row).iItem = GridItens.Row
        gobjNFiscal.ColItensNF(GridItens.Row).sProduto = objItemNFiscalOriginal.sProduto
        gobjNFiscal.ColItensNF(GridItens.Row).sUnidadeMed = objItemNFiscalOriginal.sUnidadeMed
        gobjNFiscal.ColItensNF(GridItens.Row).sDescricaoItem = objItemNFiscalOriginal.sDescricaoItem
        gobjNFiscal.ColItensNF(GridItens.Row).dQuantidade = objItemNFiscalOriginal.dQuantidade

        'Tenta fazer a distribuição automatica p\ o item de acordo com a nova UM
        lErro = gobjDistribuicao.Distribuicao_Processa()
        If lErro <> SUCESSO Then gError 56995

    End If

    Me.Show

    'recalcula os valores de desconto, percentual de desconto e valor total
    Call Calcula_Valores(GridItens.Row)

    Call ValorTotal_Calcula
    
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 56995
    End If
    
    Exit Sub

Erro_objEventoItensNFiscalOriginal_evSelecao:

    Select Case gErr

        Case 56995
        
        Case 55142
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNFiscalOriginal.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156437)

    End Select

    Exit Sub

End Sub

Public Sub objEventoRecebimento_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_objEventoRecebimento_evSelecao

    Set objNFiscal = obj1

    'Preenche a Tela com alguns dados de objNFiscal
    lErro = Preenche_Tela_Recebimento(objNFiscal)
    If lErro <> SUCESSO Then gError 35248

    'Fecha o Comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Call gobjTribTab.TributacaoRecalcular_Click
    
    Exit Sub

Erro_objEventoRecebimento_evSelecao:

    Select Case gErr

        Case 35248

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156438)

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

    'Verifica se tem alguma Placa U.F. foi preenchida
    If Len(Trim(PlacaUF.Text)) = 0 Then Exit Sub

    'Verifica se existe o ítem na combo
    lErro = Combo_Item_Igual(PlacaUF)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 35318

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 35319

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True

    Select Case gErr

        Case 35318

        Case 35319
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156439)

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

Public Sub SerieNFiscalOriginal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SerieNFiscalOriginal_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SerieOriginalLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As New Collection

    'Recolhe a Série da Nota Fiscal Original da tela
    objSerie.sSerie = SerieNFiscalOriginal.Text

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerieOriginal)

End Sub

Public Sub objEventoSerieOriginal_evSelecao(obj1 As Object)

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

    'Chama a Tela NFiscalInternaSaidalLista
    Call Chama_Tela("NFiscalInternaSaidaLista", colSelecao, objNFiscal, objEventoNFiscalOriginal)

End Sub

Public Sub objEventoNFiscalOriginal_evSelecao(obj1 As Object)

Dim objNFiscal As ClassNFiscal

    Set objNFiscal = obj1

    'Preenche a Série e o Número da Nota Fiscal Original
    SerieNFiscalOriginal.Text = objNFiscal.sSerie
    NFiscalOriginal.Text = objNFiscal.lNumNotaFiscal

    Me.Show

End Sub

Public Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Public Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    Fornecedor.Text = objFornecedor.lCodigo

    Call Fornecedor_Validate(bCancel)

    Me.Show

End Sub

Public Sub SerieLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'recolhe a serie da tela
    objSerie.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie)

End Sub

Public Sub objEventoSerie_evSelecao(obj1 As Object)

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

End Sub

Public Sub NFiscalLabel_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim colSelecao As Collection
Dim objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_NFiscalLabel_Click

    'Recolhe os dados da NFiscal que estão na tela
    lErro = Move_NF_Memoria(objNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO Then gError 35160

    'Chama a tela de Browse Nota Fiscal de Entrada
    Call Chama_Tela("NFiscalEntDevLista", colSelecao, objNFiscal, objEventoNFiscal)

    Exit Sub

Erro_NFiscalLabel_Click:

    Select Case gErr

        Case 35160

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156440)

    End Select

    Exit Sub

End Sub

Public Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objNFiscal = obj1

    'Traz a NotaFiscal de entrada para a a tela
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 35162

    'Fecha o Comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case gErr

        Case 35162

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156441)

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

End Sub

Public Sub objEventoTransportadora_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora

    Set objTransportadora = obj1

    'Coloca a Transportadora na Tela
    Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    Me.Show

End Sub

Public Sub NaturezaLabel_Click()

Dim objNaturezaOp As New ClassNaturezaOp
Dim colSelecao As New Collection
Dim dtDataref As Date

    'Se NaturezaOP estiver preenchida coloca no Obj
    If Len(Trim(NaturezaOp.ClipText)) > 0 Then objNaturezaOp.sCodigo = NaturezaOp.Text

    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        dtDataref = MaskedParaDate(DataEmissao)
    Else
        dtDataref = DATA_NULA
    End If
        
    If dtDataref < DATA_INICIO_CFOP4 Then
    
        If Len(Trim(DataEntrada)) > 0 Then
            dtDataref = MaskedParaDate(DataEntrada)
        End If
    
    End If
            
    'Adiciona os limites de Natureza
    Select Case iEmitenteNF
    
        Case DOCINFO_EMPRESA
            colSelecao.Add NATUREZA_ENTRADA_COD_INICIAL
            colSelecao.Add NATUREZA_ENTRADA_COD_FINAL
            
        Case Else
            colSelecao.Add NATUREZA_SAIDA_COD_INICIAL
            colSelecao.Add NATUREZA_SAIDA_COD_FINAL
            
    End Select
    
    'Chama a Tela de browse de NaturezaOp
    Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, objEventoNatureza, "{fn LENGTH(Codigo) } = " & IIf(dtDataref < DATA_INICIO_CFOP4, "3", "4"))

End Sub

Public Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim objNaturezaOp As ClassNaturezaOp

    Set objNaturezaOp = obj1

    'Coloca a natureza Operação na tela
    NaturezaOp.Text = objNaturezaOp.sCodigo

    Call NaturezaOp_Validate(bSGECancelDummy)
    
    Me.Show

End Sub

Public Sub BotaoLocalizacao_Click()

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim iAlmoxarifado As Integer

On Error GoTo Erro_BotaoLocalizacao_Click:

    'Verifica se tem alguma linha do Grid Selecionada
    If GridItens.Row = 0 Then gError 43756

    'Verifica se o Produto está preenchido
    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 35163

    If iProdutoPreenchido <> PRODUTO_VAZIO Then
        colSelecao.Add sProdutoFormatado

        objEstoqueProduto.sProduto = sProdutoFormatado
        'Lê o Almoxarifado Padrão
        lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objEstoqueProduto.sProduto, iAlmoxarifado)
        If iAlmoxarifado <> 0 Then objEstoqueProduto.iAlmoxarifado = iAlmoxarifado

        'Chama a Tela de Browse EstoqueProdutoFilialLista
        Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoLocalizacao)

    Else
        Error 43757
    End If

    Exit Sub

Erro_BotaoLocalizacao_Click:

    Select Case gErr

        Case 35163

        Case 43756
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 43757
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156442)

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
        If GridItens.Row = 0 Then gError 43754

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 83083
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoLista_Consulta
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 43754
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 83083
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156443)

    End Select

    Exit Sub

End Sub

Public Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProdutoEnxuto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Verifica se o Produto está preenchido
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then

        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 43788

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 126890

            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 43790

            'Se não encontrou o Produto --> Erro
            If lErro = 28030 Then gError 43791

            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True

            If Not (Me.ActiveControl Is Produto) Then

                GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
    
                'Preenche a Linha do Grid
                lErro = ProdutoLinha_Preenche(objProduto)
                If lErro <> SUCESSO Then gError 43792

            End If
            
        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 43788, 43790, 43792

        Case 43789
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)

        Case 43791
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156444)

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
        If lErro <> SUCESSO Then gError 64009

        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 35164
        
        'Se não encontrou a NFiscal --> erro
        If lErro <> SUCESSO Then gError 35165

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

        'Lê o Tipo de Documento com o Código de objNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 35166

        'Se não encontrou o Tipo de Documento --> erro
        If lErro = 31415 Then gError 35167
        
'        If giTipoVersao = VERSAO_FULL Then
            If objTipoDocInfo.sNomeTelaNFiscal <> Me.Name Then gError 35168
'        End If
        
        'Traz NFiscal para a tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 35169

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 35164, 35169, 35166, 64009

        Case 35165
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)
            Call Limpa_Tela_NFiscal
            iAlterado = 0

        Case 35167
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 35168
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_DIFERENTE_NF_ENTRADA_DEVOLUCAO", gErr, objNFiscal.iTipoNFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156445)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub SerieNFiscalOriginal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SerieNFiscalOriginal_Validate

    'Verififca se está preenchida
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Then Exit Sub

    'Verifica se foi alguma Série selecionada
    If SerieNFiscalOriginal.Text = SerieNFiscalOriginal.List(SerieNFiscalOriginal.ListIndex) Then Exit Sub

    'Tenta achar a Série na combo
    lErro = Combo_Item_Igual(SerieNFiscalOriginal)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 35174

    'Não encontrou a Série
    If lErro = 12253 Then gError 35175

    Exit Sub

Erro_SerieNFiscalOriginal_Validate:

    Cancel = True

    Select Case gErr

        Case 35174

        Case 35175
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, SerieNFiscalOriginal.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156446)

    End Select

    Exit Sub

End Sub

Public Sub TipoNFiscal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoNFiscal_Click()

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TipoNFiscal_Click

    iAlterado = REGISTRO_ALTERADO

    If TipoNFiscal.ListIndex = -1 Then Exit Sub

    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 35178

    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 35179
    
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then
        giComplemento = MARCADO
    Else
        giComplemento = DESMARCADO
    End If

    gobjRastreamento.iCodigo = objTipoDocInfo.iCodigo

    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
        NFiscal.Visible = False
        NFiscalInterna.Visible = True
        BotaoLimparNF.Visible = True
        'ChvNFe.Enabled = False
        'Cor Preta para label pois não é obrigatorio
        NFiscalLabel.ForeColor = vbBlack
    Else
        NFiscal.Visible = True
        NFiscalInterna.Visible = False
        BotaoLimparNF.Visible = False
        'ChvNFe.Enabled = True
        'Cor Vermelha para label pois é obrigatorio
        NFiscalLabel.ForeColor = &H80&
    End If
    
    If NFiscalInterna.Visible Then
        'Seleciona série padrão
        lErro = Seleciona_Serie_Padrao()
        If lErro <> SUCESSO Then gError 35178
    End If

    iEmitenteNF = objTipoDocInfo.iEmitente
    
    'Coloca na Tela a Natureza Operação Padrão
    If objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
    
        If iEmitenteNF <> DOCINFO_EMPRESA Then
        
            NaturezaOp.Text = NaturezaOp_Conv34(objTipoDocInfo.sNatOpExtPadrao, IIf(MaskedParaDate(DataEntrada) = DATA_NULA, MaskedParaDate(DataEmissao), MaskedParaDate(DataEntrada)))
            
        End If
            
        Call Habilita_Cliente
    
    End If
        
    If objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        
        If iEmitenteNF <> DOCINFO_EMPRESA Then
        
            NaturezaOp.Text = NaturezaOp_Conv34(objTipoDocInfo.sNatOpExtPadrao, IIf(MaskedParaDate(DataEntrada) = DATA_NULA, MaskedParaDate(DataEmissao), MaskedParaDate(DataEntrada)))
            
        End If
            
        Call Habilita_Fornecedor
    
    End If
    
    If TipoNFiscal.ListIndex <> -1 Then
        If (TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFIEDSB) Or (TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFEEDSB) Then
            If TabStrip1.Tabs.Count = TAB_BENEFICIAMENTO - 1 Then
                TabStrip1.Tabs.Add TAB_BENEFICIAMENTO, , "Benef."
            End If
        ElseIf TabStrip1.Tabs.Count = TAB_BENEFICIAMENTO Then
               TabStrip1.Tabs.Remove TAB_BENEFICIAMENTO
        End If
    End If
    
    ''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
    If IsDate(DataEntrada.Text) Then Call gobjTribTab.DataEntrada_Alterada(StrParaDate(DataEntrada.Text))
    Call gobjTribTab.Serie_Alterada(Serie.Text)
    '*** fim tributacao

    'tratamento dado ao subtipo contábil: de acordo com o DOCINFO selecionado, serão carregados os modelos pertencentes ao subtipo contábil correspondente
    lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_ESTOQUE, objTipoDocInfo.iSubTipoContabil)
    If lErro <> SUCESSO Then gError 79368
    
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 79368
    End If

    Exit Sub

Erro_TipoNFiscal_Click:

    Select Case gErr

        Case 35178, 79368

        Case 35179
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156447)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimparNF_Click()
    NFiscalInterna.Caption = ""
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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 35176

    'Se não conseguir --> Erro
    If lErro <> SUCESSO Then gError 35177

    'tenta ler a natureza de operacao
    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
        
    'Lê o Tipo da NF
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 55544
    
    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 55545

    'Se a Natureza de Operação ainda não foi preenchida
    If Len(Trim(NaturezaOp.ClipText)) = 0 Then

        NaturezaOp.Text = NaturezaOp_Conv34(objTipoDocInfo.sNaturezaOperacaoPadrao, IIf(MaskedParaDate(DataEntrada) = DATA_NULA, MaskedParaDate(DataEmissao), MaskedParaDate(DataEntrada)))

    End If
    
    'Verifica se é externa ou interna
    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
        NFiscal.Visible = False
        NFiscalInterna.Visible = True
        BotaoLimparNF.Visible = True
        'ChvNFe.Enabled = False
        'Cor Preta para label pois não é obrigatorio
        NFiscalLabel.ForeColor = vbBlack
    Else
        NFiscal.Visible = True
        NFiscalInterna.Visible = False
        BotaoLimparNF.Visible = False
        'ChvNFe.Enabled = True
        'Cor Vermelha para label pois é obrigatorio
        NFiscalLabel.ForeColor = &H80&
    End If
    
    If NFiscalInterna.Visible Then
        'Seleciona série padrão
        lErro = Seleciona_Serie_Padrao()
        If lErro <> SUCESSO Then gError 55544
    End If
    
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 55544
    End If

    Exit Sub

Erro_TipoNFiscal_Validate:

    Cancel = True

    Select Case gErr

        Case 35176, 55544

        Case 35177
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case 55545
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156448)

    End Select

    Exit Sub

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
            If lErro <> SUCESSO Then gError 35180

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 35181

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
                If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
                'Seleciona filial na Combo Filial
                Call CF("Filial_Seleciona", Filial, iCodFilial)
                
''*** incluidos p/tratamento de tributacao *******************************
                Call gobjTribTab.FilialFornecedor_Alterada(objFornecedor.lCodigo, iCodFilial)
'*** fim tributacao

            End If

            lErro = CF2(Me, "NFCF_Fornecedor_Validate", objFornecedor.lCodigo, iCodFilial)
            If lErro <> SUCESSO Then gError 131930

        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

            Filial.Clear

            Call CF2(Me, "NFCF_Limpa_Tela")

        End If

        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 35180, 35181, 131930

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156449)

    End Select

    Exit Sub

End Sub

Public Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO

    Call Cliente_Preenche

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFilialCliente As New ClassFilialCliente
Dim sNomeRed As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 35189

    'Se não encontra o ítem com o código informado
    If lErro = 6730 Then

        If Fornecedor.Visible = True Then

            'Verifica de o fornecedor foi digitado
            If Len(Trim(Fornecedor.Text)) = 0 Then gError 35182

            sNomeRed = Fornecedor.Text

            objFilialFornecedor.iCodFilial = iCodigo
            
            'Pesquisa se existe Filial com o código extraído
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 35183

            'Se não encontrou a Filial Fornecedor --> erro
            If lErro = 18272 Then gError 35184

            'Mostra a Filial Fornecedor na tela
            Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

            Call gobjTribTab.FilialFornecedor_Alterada(objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)

        Else

            'Verifica se o Cliente foi preenchido
            If Len(Trim(Cliente.ClipText)) = 0 Then gError 35185

            sNomeRed = Cliente.Text
            
            'Lê a Filial Cliente
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeRed, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 35186

            'Se não encontrou a Filial Cliente --> erro
            If lErro = 17660 Then gError 35187

            'Mostra a Filial Cliente na tela
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

            Call gobjTribTab.FilialCliente_Alterada(objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
        
        End If

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 35188

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 35183, 35186, 35189

        Case 35184
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 35182
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 35185
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 35187
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 35188
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156450)

    End Select

    Exit Sub

End Sub

Public Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Serie_Validate

    If iSerieAlterada = 0 Then Exit Sub
    
    'Verifica se foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub

    'Verifica se foi selecionada
    If Serie.ListIndex = -1 Then

        'Tenta selecionar a serie
        lErro = Combo_Item_Igual(Serie)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 35190
    
        'Se a serie não está cadastrada --> verifica se é uma nota interna, se for ==> erro, pois a serie não está cadastrada
        If lErro <> SUCESSO Then
    
            'Se o tipo está preenchido
            If Len(Trim(TipoNFiscal.Text)) > 0 Then
        
                objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
                'Lê o Tipo de Documento
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO Then gError 35191
        
                If objTipoDocInfo.iTipo = DOCINFO_NF_INT_ENTRADA Then gError 35192
        
            End If
            
            'Verifica se tamanho da série é maior do que o espaço no bd ==> erro
            If Len(Trim(Serie.Text)) > STRING_SERIE Then gError 35193
    
        End If
    
    End If
        
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.Serie_Alterada(Serie.Text)
'*** fim tributacao
        
    iSerieAlterada = 0
    
    Exit Sub

Erro_Serie_Validate:

    Cancel = True


    Select Case gErr

        Case 35190, 35191

        Case 35192
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, Serie.Text)

        Case 35193
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_SERIE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156451)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a data de emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 35195

    Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
        
    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case gErr

        Case 35195

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156452)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 35196

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 35196

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156453)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 35197

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 35197

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156454)

    End Select

    Exit Sub

End Sub

Public Sub DataEntrada_Change()
    iAlterado = REGISTRO_ALTERADO
    iDataEntradaAlterada = REGISTRO_ALTERADO
End Sub

Public Sub DataEntrada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntrada_Validate

    If iDataEntradaAlterada <> 0 Then
        
        'Verifica se a data de Entrada foi digitada
        If Len(Trim(DataEntrada.ClipText)) = 0 Then Exit Sub
    
        'Critica a data digitada
        lErro = Data_Critica(DataEntrada.Text)
        If lErro <> SUCESSO Then gError 35198

''*** incluidos p/tratamento de tributacao *******************************
        Call gobjTribTab.DataEntrada_Alterada(StrParaDate(DataEntrada.Text))
'*** fim tributacao
        
        iDataEntradaAlterada = 0

    End If

    Exit Sub

Erro_DataEntrada_Validate:

    Cancel = True

    Select Case gErr

        Case 35198

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156455)

    End Select

    Exit Sub

End Sub

'horaentrada
Public Sub HoraEntrada_GotFocus()
    Call MaskEdBox_TrataGotFocus(HoraEntrada, iAlterado)
End Sub

'horaentrada
Public Sub HoraEntrada_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

'horaentrada
Public Sub HoraEntrada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_HoraEntrada_Validate

    'Verifica se a hora de Entrada foi digitada
    If Len(Trim(HoraEntrada.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(HoraEntrada.Text)
    If lErro <> SUCESSO Then gError 89779

    Exit Sub

Erro_HoraEntrada_Validate:

    Cancel = True

    Select Case gErr

        Case 89779

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156456)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntrada_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntrada_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEntrada, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 35199

    Exit Sub

Erro_UpDownEntrada_DownClick:

    Select Case gErr

        Case 35199

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156457)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntrada_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntrada_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEntrada, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 35200

    Exit Sub

Erro_UpDownEntrada_UpClick:

    Select Case gErr

        Case 35200

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156458)

    End Select

    Exit Sub

End Sub

Public Function Total_Calcula(Optional iTesta As Integer = 1) As Long

Dim lErro As Long
Dim dTotal As Double

On Error GoTo Erro_Total_Calcula

    lErro = gobjTribTab.ValorTotal_Calcula(dTotal, iTesta)
    If lErro <> SUCESSO Then gError 43786

    Total.Text = Format(dTotal, "Standard")

    Total_Calcula = SUCESSO

    Exit Function

Erro_Total_Calcula:

    Total_Calcula = gErr

    Select Case gErr

        Case 43786

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156459)

    End Select

    Exit Function

End Function

Private Function IndTab_ObterIndFrame(ByVal iIndTab As Integer) As Integer
'obtem indice nos frames a partir do indice na tabstrip

    IndTab_ObterIndFrame = iIndTab
    
    If giTipoVersao = VERSAO_LIGHT And iIndTab = TAB_BENEFICIAMENTO Then IndTab_ObterIndFrame = iIndTab + 2 'contab e rastro

End Function

Public Sub TabStrip1_Click()

Dim lErro As Long, iFrameSelecionado As Integer, iTabSelecionado As Integer

On Error GoTo Erro_TabStrip1_Click

    iTabSelecionado = TabStrip1.SelectedItem.Index
    iFrameSelecionado = IndTab_ObterIndFrame(iTabSelecionado)

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If iFrameSelecionado <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(IndTab_ObterIndFrame(iFrameAtual), TabStrip1, Me) <> SUCESSO Then Exit Sub

''*** incluidos p/tratamento de tributacao *******************************

        'se abriu o tab de tributacao
        If iTabSelecionado = TAB_Tributacao Then
            lErro = gobjTribTab.TabClick()
            If lErro <> SUCESSO Then gError 35263
        End If

        'Torna Frame correspondente ao Tab selecionado visivel
'        If giTipoVersao = VERSAO_FULL Then
            Frame1(iFrameSelecionado).Visible = True
            'Torna Frame atual invisivel
            Frame1(iFrameAtual).Visible = False
'        ElseIf giTipoVersao = VERSAO_LIGHT Then
'            FrameL(TabStrip1.SelectedItem.Index).Visible = True
'            FrameL(iFrameAtual).Visible = False
'        End If
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = iFrameSelecionado

        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If TabStrip1.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao

        Select Case iTabSelecionado
        
            Case TAB_DadosPrincipais
                Parent.HelpContextID = IDH_ENTRADA_NF_DEVOLUCAO_DADOS_PRINCIPAIS
                
            Case TAB_Itens
                Parent.HelpContextID = IDH_ENTRADA_NF_DEVOLUCAO_ITENS
                        
            Case TAB_Complemento
                Call Calcula_Msg
                Parent.HelpContextID = IDH_ENTRADA_NF_DEVOLUCAO_COMPLEMENTO
            
            Case TAB_Tributacao
                Parent.HelpContextID = IDH_ENTRADA_NF_DEVOLUCAO_TRIBUTACAO
            
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_ENTRADA_NF_DEVOLUCAO_CONTABILIZACAO
        
        End Select

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 35263

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156460)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()
    Unload Me
End Sub

Public Sub ValorDespesas_Change()
    iAlterado = REGISTRO_ALTERADO
    iValorDespesasAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorFrete_Change()
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
        If lErro <> SUCESSO Then gError 35201

        dValorFrete = CDbl(ValorFrete.Text)
        'Coloca o Valor Formatado na tela
        ValorFrete.Text = Format(dValorFrete, "Standard")

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

        Case 35201

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156461)

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
        If lErro <> SUCESSO Then gError 35203

        dValorDespesas = CDbl(ValorDespesas.Text)

        'coloca o valor formatado na tela
        ValorDespesas.Text = Format(dValorDespesas, "Standard")

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

        Case 35203

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156462)

    End Select

    Exit Sub

End Sub

Public Sub ValorSeguro_Change()
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
        If lErro <> SUCESSO Then gError 35202

        dValorSeguro = CDbl(ValorSeguro.Text)
        'Coloca o valor formatado na tela
        ValorSeguro.Text = Format(dValorSeguro, "Standard")

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

        Case 35202

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156463)

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

On Error GoTo Erro_Transportadora_Validate

    'Verifica se foi preenchida
    If Len(Trim(Transportadora.Text)) = 0 Then Exit Sub

    'Verifica se foi selecionada
    If Transportadora.Text = Transportadora.List(Transportadora.ListIndex) Then Exit Sub

    'Tenta Selecionar
    lErro = Combo_Seleciona(Transportadora, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 35205

    'Não conseguiu e retornou o Código
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo
        
        'Lê a Transportadora com o Código digitado
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 35206

        'Se não encontrou a Transportadora --> erro
        If lErro = 19250 Then gError 35207

        'Mostra a Transportadora na tela
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    End If

    'Não está cadstrado
    If lErro = 6731 Then gError 35208

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case gErr

        Case 35205, 35206

        Case 35207
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr, objTransportadora.iCodigo)
            
        Case 35208
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156464)

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
    If lErro <> SUCESSO Then gError 35209

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True

    Select Case gErr

        Case 35209

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156465)

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
    lErro = Valor_NaoNegativo_Critica(PesoBruto.Text)
    If lErro <> SUCESSO Then gError 35211

    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True

    Select Case gErr

        Case 35211

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156466)

    End Select

    Exit Sub

End Sub

Public Function Move_NF_Memoria(objNFiscal As ClassNFiscal, objTipoDocInfo As ClassTipoDocInfo) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objCliente As New ClassCliente
Dim objProjetoInfo As Object

On Error GoTo Erro_Move_NF_Memoria
    
    If objTipoDocInfo Is Nothing Then

        Set objTipoDocInfo = New ClassTipoDocInfo
    
        objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)

        'Lê o Tipo de documento
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 35274
        
        'Se não encontrar --> erro
        If lErro = 31415 Then gError 35275

    End If

    If objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then

        'Verifica se o Fornecedor está preenchido
        If Len(Trim(Fornecedor.ClipText)) > 0 Then

            objFornecedor.sNomeReduzido = Fornecedor.Text
            'Lê o fornecedor
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 35213
            
            If lErro = 6681 Then gError 35214

            objNFiscal.lFornecedor = objFornecedor.lCodigo
            objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)
            
        End If

    ElseIf objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then

        If Len(Trim(Cliente.ClipText)) > 0 Then

            objCliente.sNomeReduzido = Cliente.Text
            'Lê o Cliente
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 35215
            
            If lErro = 12348 Then gError 35216

            objNFiscal.lCliente = objCliente.lCodigo
            objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
            objNFiscal.iFilialEntrega = Codigo_Extrai(Filial.Text)
            
        End If
        
    End If

    'Preenche o objNFiscal com os demais campos da tela
    If Len(Trim(PesoBruto.ClipText)) > 0 Then objNFiscal.dPesoBruto = CDbl(PesoBruto.Text)
    If Len(Trim(PesoLiquido.ClipText)) > 0 Then objNFiscal.dPesoLiq = CDbl(PesoLiquido.Text)

    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        objNFiscal.dtDataEmissao = CDate(DataEmissao.Text)
    Else
        objNFiscal.dtDataEmissao = DATA_NULA
    End If

    If Len(Trim(DataEntrada.ClipText)) > 0 Then
        objNFiscal.dtDataEntrada = CDate(DataEntrada.Text)
    Else
        objNFiscal.dtDataEntrada = DATA_NULA
    End If

'horaentrada
    If objNFiscal.dtDataEntrada <> DATA_NULA Then
        If Len(Trim(HoraEntrada.ClipText)) > 0 Then
            objNFiscal.dtHoraEntrada = CDate(HoraEntrada.Text)
        Else
            objNFiscal.dtHoraEntrada = Time
        End If
    Else
        objNFiscal.dtHoraEntrada = 0
    End If

    objNFiscal.dtDataRegistro = gdtDataAtual

    'Verifica se os valores estão preenchidos antes de converter
    If Len(Trim(ValorDesconto.ClipText)) > 0 Then objNFiscal.dValorDesconto = CDbl(ValorDesconto.Text)
    If Len(Trim(ValorFrete.ClipText)) > 0 Then objNFiscal.dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorSeguro.ClipText)) > 0 Then objNFiscal.dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(ValorDespesas.ClipText)) > 0 Then objNFiscal.dValorOutrasDespesas = CDbl(ValorDespesas.Text)
    If Len(Trim(SubTotal.Caption)) > 0 Then objNFiscal.dValorProdutos = CDbl(SubTotal.Caption)
    If Len(Trim(Total.Text)) > 0 Then objNFiscal.dValorTotal = CDbl(Total.Text)
    If Len(Trim(VolumeQuant.ClipText)) > 0 Then objNFiscal.lVolumeQuant = StrParaLong(VolumeQuant.Text)

    'Armazena o responsável pelo frete
    objNFiscal.iFreteRespons = TipoFrete.ItemData(TipoFrete.ListIndex)
    
    objNFiscal.dtDataReferencia = DATA_NULA
    objNFiscal.sVolumeNumero = VolumeNumero.Text
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    If Not NFiscalInterna.Visible Then objNFiscal.sChvNFe = ChvNFe.ClipText
    
    If objNFiscal.iTipoDocInfo = TIPODOCINFO_TIPO_NFEXT Then
        objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
    Else
        objNFiscal.lNumNotaFiscal = StrParaLong(NFiscalInterna.Caption)
    End If
    
    If Len(Trim(NumRecebimento.Text)) > 0 Then objNFiscal.lNumRecebimento = CLng(NumRecebimento.Text)
    objNFiscal.iCodTransportadora = Codigo_Extrai(Transportadora.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
        
    If objTipoDocInfo.sTipoDocCPR = "" Then
        objNFiscal.iStatus = STATUS_BAIXADO
    Else
        objNFiscal.iStatus = STATUS_LANCADO
    End If
    
    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
    objNFiscal.sMensagemNota = Mensagem.Text
    objNFiscal.sMensagemCorpoNota = MensagemCorpo.Text
    objNFiscal.sNaturezaOp = NaturezaOp.Text
    objNFiscal.sPlaca = Placa.Text
    objNFiscal.sPlacaUF = PlacaUF.Text
    
    'Incluído por Luiz Nogueira em 21/08/03
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objNFiscal.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objNFiscal.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
    
    objNFiscal.dtDataSaida = DATA_NULA
    objNFiscal.dtDataVencimento = DATA_NULA

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

        Case 35213, 35215, 35274, 181761

        Case 35214
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case 35216
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case 35275
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156467)

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
    If lErro <> SUCESSO Then gError 102444
    
    Exit Sub

Erro_VolumeEspecie_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102444
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156468)

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
    If lErro <> SUCESSO Then gError 102445
    
    Exit Sub

Erro_VolumeMarca_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102445
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156469)

    End Select

End Sub

Public Sub VolumeNumero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub VolumeQuant_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Limpa_Tela_NFiscal2()
'Limpa a Tela NFiscalEntrada, mas mantém a natureza e o tipo
    
Dim sNatureza As String
Dim sTipoNFiscal As String

    sNatureza = NaturezaOp.Text
    sTipoNFiscal = TipoNFiscal.Text
    
    Call Limpa_Tela_NFiscal

    NaturezaOp.Text = sNatureza
    TipoNFiscal.Text = sTipoNFiscal

End Sub

Public Sub Limpa_Tela_NFiscal()
'Limpa a Tela NFiscalEntrada

Dim iIndice As Integer

    'Chama o Limpa tela
    Call Limpa_Tela(Me)
    
    gsChvAnt = ""

    Set gobjAnotacao = Nothing
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
        
    Set gobjInfoAdic = Nothing
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next
    
    'Limpa os campos que não são limpos pelo Limpa_Tela
    Filial.Clear
'    If giTipoVersao = VERSAO_FULL Then
        Serie.Text = ""
        Call Serie_Validate(bSGECancelDummy)
'    End If
    'EletronicaFed.Value = vbUnchecked
    
    If gobjEST.iSerieEletronicaPadrao = MARCADO Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    
    Status.Caption = ""
    
    'limpa outros campos da tela
    Call Limpa_Tela_NFiscal1
    
    Set gobjNFiscal = New ClassNFiscal

    'Preenche a Data de Entrada com a Data Atual
    DataEntrada.PromptInclude = False
    DataEntrada.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEntrada.PromptInclude = True
    Call DataEntrada_Validate(bSGECancelDummy)

    Call CF2(Me, "NFCF_Limpa_Tela")
    
    MsgAutomatica.Value = vbChecked
    
    iAlterado = 0
    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0
    
End Sub

Private Function Preenche_Tela(objNFiscal As ClassNFiscal) As Long
'Complementa o carregamento da tela de Notas Fiscais de Entrada e Devolução

Dim lErro As Long
Dim bCancel As Boolean

On Error GoTo Erro_Preenche_Tela

    'Preenhce  a tela com os dados vindos em objNFiscal
    NaturezaOp.Text = objNFiscal.sNaturezaOp
    Serie.Text = Desconverte_Serie_Eletronica(objNFiscal.sSerie)
    If ISSerieEletronica(objNFiscal.sSerie) Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    Call Serie_Validate(bSGECancelDummy)
    NFiscal.Text = objNFiscal.lNumNotaFiscal
    NFiscalInterna.Caption = objNFiscal.lNumNotaFiscal
    
    'Se não trouxer para tela ao regravar vai excluir a NF anterior e apagar
    'qualquer rastro de uso do recebimento deixando-o perdido conforme ocorreu na DLE
    If objNFiscal.lNumRecebimento <> 0 Then NumRecebimento.Text = objNFiscal.lNumRecebimento

    DataEmissao.PromptInclude = False
    If objNFiscal.dtDataEmissao <> DATA_NULA Then DataEmissao.Text = Format(objNFiscal.dtDataEmissao, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    DataEntrada.PromptInclude = False
    If objNFiscal.dtDataEntrada <> DATA_NULA Then DataEntrada.Text = Format(objNFiscal.dtDataEntrada, "dd/mm/yy")
    DataEntrada.PromptInclude = True

'horaentrada
    HoraEntrada.PromptInclude = False
    'este teste está correto
    If objNFiscal.dtDataEntrada <> DATA_NULA Then HoraEntrada.Text = Format(objNFiscal.dtHoraEntrada, "hh:mm:ss")
    HoraEntrada.PromptInclude = True

    VolumeNumero = objNFiscal.sVolumeNumero
    SubTotal.Caption = Format(objNFiscal.dValorProdutos, "Standard")
    ValorFrete.Text = Format(objNFiscal.dValorFrete, "Standard")
    ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Standard")
    ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Standard")
    ValorDesconto.Text = Format(objNFiscal.dValorDesconto, "Standard")
    Total.Text = Format(objNFiscal.dValorTotal, "Standard")

    If objNFiscal.iCodTransportadora > 0 Then
        Transportadora.Text = objNFiscal.iCodTransportadora
        Transportadora_Validate bCancel
    End If

    Placa.Text = objNFiscal.sPlaca
    PlacaUF.Text = objNFiscal.sPlacaUF
    If objNFiscal.lVolumeQuant > 0 Then VolumeQuant.Text = objNFiscal.lVolumeQuant
    
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
    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = CStr(objNFiscal.dPesoLiq)
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = CStr(objNFiscal.dPesoBruto)

    'Preenhce o Grid com os Ítens da Nota Fiscal em questão
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 35239

    'distribuicao
    'Preenche o Grid com as Distribuições dos itens da Nota Fiscal
    lErro = gobjDistribuicao.Preenche_GridDistribuicao(objNFiscal)
    If lErro <> SUCESSO Then gError 89559

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = gErr

    Select Case gErr

        Case 35239, 89559

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156470)

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
Dim sCclMascarado As String
Dim dDesconto As Double, dQuantidade As Double, dValorUnitario As Double, dValorTotal As Double, dPercentDesc  As Double, dValorTotalItens As Double, dValorReal As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_Preenche_GridItens

    'Faz o somátorio dos valores dos itens, pois o Valor da Nota pode não ser o Somatório dos Itens (ex. Fretes, ...)
    For Each objItemNF In objNFiscal.ColItensNF

        dValorTotalItens = dValorTotalItens + objItemNF.dValorTotal
        
    Next
    
    'Para cada ítem da Coleção
    For Each objItemNF In objNFiscal.ColItensNF

        iIndice = iIndice + 1

        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 35240

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177926

        'Preenche o Grid
        '****** IF INCLUÍDO PARA TRATAMENTO DE GRADE ***************
        If objItemNF.colItensRomaneioGrade.Count > 0 And objProduto.iKitVendaComp <> MARCADO Then GridItens.TextMatrix(iIndice, 0) = "# " & GridItens.TextMatrix(iIndice, 0)
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItemNF.sDescricaoItem
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemNF.sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemNF.dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(objItemNF.dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
        
        dDesconto = objItemNF.dValorDesconto
    
        GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(dDesconto, "Standard")

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then dQuantidade = CDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        
        lErro = CF2(Me, "NFEDEV_Preenche_GridItens", objItemNF, iIndice, dQuantidade)
        If lErro <> SUCESSO Then gError 177052
        
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) > 0 Then dValorUnitario = CDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))

        dValorTotal = dValorUnitario * dQuantidade

        'Calcula o novo Percentual de Desconto
        If dValorTotal > 0 Then

            dPercentDesc = dDesconto / dValorTotal

            If dPercentDesc >= 1 Then gError 108725

        End If

        GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")
        
        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dValorUnitario, dPercentDesc, dDesconto, dValorReal)
        
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(dValorReal, "Standard")
        
        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dValorUnitario, 0, 0, dValorReal)
        
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col) = Format(dValorReal, "Standard")
        
        If objItemNF.lNumNFOrig > 0 Then GridItens.TextMatrix(iIndice, iGrid_NumNFOrig_Col) = objItemNF.lNumNFOrig
        If Len(Trim(objItemNF.sSerieNFOrig)) > 0 Then GridItens.TextMatrix(iIndice, iGrid_SerieNFOrig_Col) = objItemNF.sSerieNFOrig
        If objItemNF.iItemNFOrig > 0 Then GridItens.TextMatrix(iIndice, iGrid_ItemNFOrig_Col) = objItemNF.iItemNFOrig


        sCclMascarado = ""
        
        'mascara Ccl , se estiver informada
        If objItemNF.sCcl <> "" Then

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_RetornaCclEnxuta(objItemNF.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 49416
        
            'Preenche o campo Ccl
            Ccl.PromptInclude = False
            Ccl.Text = sCclMascarado
            Ccl.PromptInclude = True
    
            'Joga o Ccl no Grid
            GridItens.TextMatrix(iIndice, iGrid_Ccl_Col) = Ccl.Text
            
        End If
    
    Next

    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objNFiscal.objProjetoInfo, objNFiscal.lNumIntDoc, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181676

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice

    Call Grid_Refresh_Checkbox(objGridItens)

    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr

        Case 35240
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)
        
        Case 35719, 177052, 177926, 181676
        
        Case 108725
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
        
        Case 35720
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objItemNF.iAlmoxarifado)

        Case 49416
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, objItemNF.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156471)

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

Public Sub ValorUnitario_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub ValorUnitario_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub ValorUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = ValorUnitario
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

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

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 39646

        'Verifica se é o GridItens
        If objGridInt.objGrid.Name = GridItens.Name Then

            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                'Produto
                Case iGrid_Produto_Col
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 35325

                'DescricaoItem
                Case iGrid_DescProduto_Col
                    lErro = Saida_Celula_DescricaoItem(objGridInt)
                    If lErro <> SUCESSO Then gError 43295

                'Unidade Medida
                Case iGrid_UnidadeMed_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError 35326

                'Quantidade
                Case iGrid_Quantidade_Col
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 35327

                'Ccl
                Case iGrid_Ccl_Col
                    lErro = Saida_Celula_Ccl(objGridInt)
                    If lErro <> SUCESSO Then gError 49388

                'ValorUnitario
                Case iGrid_ValorUnitario_Col
                    lErro = Saida_Celula_ValorUnitario(objGridInt)
                    If lErro <> SUCESSO Then gError 35329

                'PercDesc
                Case iGrid_PercDesc_Col
                    lErro = Saida_Celula_PercDesc(objGridInt)
                    If lErro <> SUCESSO Then gError 35330

                'Desconto
                Case iGrid_Desconto_Col
                    lErro = Saida_Celula_Desconto(objGridInt)
                    If lErro <> SUCESSO Then gError 35331

                'NumNFOrig
                Case iGrid_NumNFOrig_Col
                    lErro = Saida_Celula_NumNFOrig(objGridInt)
                    If lErro <> SUCESSO Then gError 35333

                'SerieNFOrig
                Case iGrid_SerieNFOrig_Col
                    lErro = Saida_Celula_SerieNFOriginal(objGridInt)
                    If lErro <> SUCESSO Then gError 35334

                'ItemNFOrig
                Case iGrid_ItemNFOrig_Col
                    lErro = Saida_Celula_ItemNFOrig(objGridInt)
                    If lErro <> SUCESSO Then gError 35335

            End Select

        'beneficiamento
        ElseIf objGridInt.objGrid.Name = GridMovimentos.Name Then
        
            lErro = Saida_Celula_GridMovimentos(objGridInt)
            If lErro <> SUCESSO Then gError 126855

        'Rastreamento
        ElseIf objGridInt.objGrid.Name = GridRastro.Name Then

            lErro = gobjRastreamento.Saida_Celula()
            If lErro <> SUCESSO Then gError 83290

        'distribuicao
        ElseIf objGridInt.objGrid.Name = GridDist.Name Then

            lErro = gobjDistribuicao.Saida_Celula_Dist()
            If lErro <> SUCESSO Then gError 89561

        End If
        
        lErro = CF2(Me, "NFEDEV_Saida_Celula", objGridInt)
        If lErro <> SUCESSO Then gError 177035

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 35336

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 35325, 35326, 35327, 35328, 35329, 35330, 35331, 35332, 35333, 35334, 35335, 39646, 43295, 49388, 83290, 89561, 126855, 177035

        Case 35336
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156472)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl
    
    'Verifica se Ccl foi preenchido
    If Len(Trim(Ccl.ClipText)) > 0 Then

        'Critica o Ccl
        lErro = CF("Ccl_Critica", Ccl, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 49389

        If lErro = 5703 Then gError 49390

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 49424
        
    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = gErr

    Select Case gErr

        Case 49389
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 49390
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, Ccl.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 49424
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156473)
            
    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescricaoItem(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoItem

    Set objGridInt.objControle = DescricaoItem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 43296

    Saida_Celula_DescricaoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoItem:

    Saida_Celula_DescricaoItem = gErr

    Select Case gErr

        Case 43296
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156474)

    End Select

    Exit Function

End Function

Public Sub ClienteLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    objCliente.sNomeReduzido = Cliente.Text

    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Public Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

End Sub

Public Sub BotaoItensNFiscalOrig_Click()

Dim objItemNFiscalOriginal As ClassItemNF
Dim colSelecao As Collection

    'Chama a tela ItensNFiscalSaidaLista
    Call Chama_Tela("ItensNFiscalSaidaLista", colSelecao, objItemNFiscalOriginal, objEventoItensNFiscalOriginal)

End Sub

Public Function Preenche_Tela_NFOriginal(objNFiscal As ClassNFiscal) As Long
'Preenche a tela com alguns dados da Nota Fiscal Original passada em objNFiscal

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objItemNF As ClassItemNF, iIndice As Integer
Dim bCancel As Boolean, dPercDesc As Double

On Error GoTo Erro_Preenche_Tela_NFOriginal

    'Lê os Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 35243

    lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
    If lErro <> SUCESSO Then gError 46431

    lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
    If lErro <> SUCESSO Then gError 46432

    'Lê o Tipo da Nota Fiscal
    lErro = CF("TiposDocInfo_Le_SiglaNFOriginal", objNFiscal.iTipoNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 35245
    If lErro = 31415 Then gError 35246

    'Lê a Distribuição dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 20867

    gobjNFiscal.lNumIntDoc = 0

    'Limpa a Tela
    Call Limpa_Tela_NFiscal

    'Fecha o comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Set gobjNFiscal = objNFiscal
    
    TipoNFiscal.Text = objTipoDocInfo.iCodigo
    Call TipoNFiscal_Validate(bCancel)

    'Verifica a Origem do Tipo de Nota Fiscal
    If objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
        Call Habilita_Cliente
        Cliente.Text = objNFiscal.lCliente
        Call Cliente_Validate(bCancel)
        Filial.Text = objNFiscal.iFilialCli
    ElseIf objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        Call Habilita_Fornecedor
        Fornecedor.Text = objNFiscal.lFornecedor
        Call Fornecedor_Validate(bCancel)
        Filial.Text = objNFiscal.iFilialForn
    End If

    Call Filial_Validate(bCancel)

    SerieNFiscalOriginal.Text = objNFiscal.sSerie
    NFiscalOriginal.Text = objNFiscal.lNumNotaFiscal

    'Coloca nos Ítens a serie e o Numero da Nota Fiscal Original
    For Each objItemNF In objNFiscal.ColItensNF
        objItemNF.sSerieNFOrig = objNFiscal.sSerie
        objItemNF.lNumNFOrig = objNFiscal.lNumNotaFiscal
        objItemNF.iItemNFOrig = objItemNF.iItem
    Next

    ValorFrete.Text = Format(objNFiscal.dValorFrete, "Standard")
    ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Standard")
    ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Standard")
    ValorDesconto.Text = Format(objNFiscal.dValorDesconto, "Standard")
    
    Total.Text = Format(objNFiscal.dValorTotal, "Standard")
    
    'Preenche o Grid Com os Ítens lidos
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 35247
        
    lErro = gobjDistribuicao.Preenche_GridDistribuicao(objNFiscal)
    If lErro <> SUCESSO Then gError 20868
        
'    If giTipoVersao = VERSAO_FULL Then
        SubTotal.Caption = Format(objNFiscal.dValorProdutos, "Standard")
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'        SubTotal.Caption = Format(objNFiscal.dValorProdutos + objNFiscal.dValorDesconto, "Standard")
'        objNFiscal.dValorDesconto = 0
'    End If
    
'    DataEmissao.PromptInclude = False
'    If objNFiscal.dtDataEmissao <> DATA_NULA Then DataEmissao.Text = Format(objNFiscal.dtDataEmissao, "dd/mm/yy")
'    DataEmissao.PromptInclude = True

    DataEntrada.PromptInclude = False
    'If objNFiscal.dtDataSaida <> DATA_NULA Then DataEntrada.Text = Format(objNFiscal.dtDataSaida, "dd/mm/yy")
    DataEntrada.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEntrada.PromptInclude = True
    
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.DataEntrada_Alterada(StrParaDate(DataEntrada.Text))
'*** fim tributacao
    
    lErro = gobjTribTab.Traz_NFiscalOrig_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 59299

    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = CStr(objNFiscal.dPesoLiq)
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = CStr(objNFiscal.dPesoBruto)
    
    lErro = CF2(Me, "NFEDEV_Calcula_PesoKg_Total")
    If lErro <> SUCESSO Then gError 177053

    VolumeNumero = objNFiscal.sVolumeNumero
    
    If objNFiscal.lVolumeQuant > 0 Then VolumeQuant.Text = objNFiscal.lVolumeQuant
    
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
    
    If objNFiscal.iCodTransportadora > 0 Then
        Transportadora.Text = objNFiscal.iCodTransportadora
        Transportadora_Validate bCancel
    End If
    
    If objNFiscal.dValorItens = 0 Then
        For iIndice = 1 To objGridItens.iLinhasExistentes
            objNFiscal.dValorItens = objNFiscal.dValorItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
            objNFiscal.dValorDescontoItens = objNFiscal.dValorDescontoItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
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
        
    Preenche_Tela_NFOriginal = SUCESSO

    Exit Function

Erro_Preenche_Tela_NFOriginal:

    Preenche_Tela_NFOriginal = gErr

    Select Case gErr

        Case 20867, 20868, 35243, 35244, 35245, 35247, 46431, 46432, 59299, 177053

        Case 35246
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156475)

    End Select

    Exit Function

End Function

Public Sub Habilita_Cliente()
'Desabilita o Fornecededor e Habilita o Cliente

Dim bCancel As Boolean
    
'    If giTipoVersao = VERSAO_LIGHT Then
'        LabelFilial.Visible = False
'        Filial.Visible = False
'    End If
    
    ClienteLabel.Visible = True
    Cliente.Visible = True
    FornecedorLabel.Visible = False
    Fornecedor.Visible = False
    iClienteAlterado = REGISTRO_ALTERADO
    Call Cliente_Validate(bCancel)

    Call CF2(Me, "NFCF_Limpa_Tela")

End Sub

Public Sub Habilita_Fornecedor()
'Desabilita o Cliente e habilita o Fornecedor

Dim bCancel As Boolean

'    If giTipoVersao = VERSAO_LIGHT Then
'        LabelFilial.Visible = True
'        Filial.Visible = True
'    End If

    FornecedorLabel.Visible = True
    Fornecedor.Visible = True
    Cliente.Visible = False
    ClienteLabel.Visible = False
    iFornecedorAlterado = REGISTRO_ALTERADO
    Call Fornecedor_Validate(bCancel)
    
    Call CF2(Me, "NFCF_Limpa_Tela")

End Sub

Public Sub BotaoRecebimentos_Click()

Dim colSelecao As Collection
Dim objNFiscal As New ClassNFiscal

    'Chama a tela de browse RecebMaterialLista
    Call Chama_Tela("RecebMaterialLista", colSelecao, objNFiscal, objEventoRecebimento)

    Exit Sub

End Sub

Public Sub BotaoExibirDados_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_BotaoExibirDados_Click

    'Verifica se a Serie e o Número da Nota Fiscal original estão preenchidos
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Or Len(Trim(NFiscalOriginal.ClipText)) = 0 Then gError 35256

    objNFiscal.sSerie = SerieNFiscalOriginal.Text
    objNFiscal.lNumNotaFiscal = NFiscalOriginal.Text
    objNFiscal.iFilialEmpresa = giFilialEmpresa

    'Tenta lêr a Nota com esses dados
    'lErro = CF("NFiscalInternaSaida_Le_Numero", objNFiscal)
    'If lErro <> SUCESSO And lErro <> 30765 Then gError 35257
    If Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_NFIEDCSM Or Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_NFEEDCSM Then
        'Aceita NFs de devolução saída como NF original
        lErro = CF("NFiscalInternaSaida_Le_Numero2", objNFiscal)
        If lErro <> AD_SQL_SUCESSO And lErro <> 62144 Then gError 35257
    Else
        lErro = CF("NFiscalInternaSaida_Le_Numero", objNFiscal)
        If lErro <> AD_SQL_SUCESSO And lErro <> 30765 Then gError 35257
    End If
    If lErro <> SUCESSO Then gError 35258 'Não encontrou

    'Coloca na tela os dados encontrados
    lErro = Preenche_Tela_NFOriginal(objNFiscal)
    If lErro <> SUCESSO Then gError 35259

    Exit Sub

Erro_BotaoExibirDados_Click:

    Select Case gErr

        Case 35256
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NUMERO_ORIGINAL_FALTANDO2", gErr)

        Case 35257, 35259

        Case 35258
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_INTERNA_SAIDA_NAO_CADASTRADA", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156476)

    End Select

    Exit Sub

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    If iClienteAlterado = 1 Then

        If Len(Trim(Cliente.Text)) > 0 Then

            lErro = TP_Cliente_Le3(Cliente, objCliente, iCodFilial)
            If lErro <> SUCESSO Then gError 35265

            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then gError 35266

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
                If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
                Call CF("Filial_Seleciona", Filial, iCodFilial)
            
            ''*** incluidos p/tratamento de tributacao *******************************
                Call gobjTribTab.FilialCliente_Alterada(objCliente.lCodigo, iCodFilial)
            '*** fim tributacao
            
            End If
            
            lErro = CF2(Me, "NFCF_Cliente_Validate", objCliente.lCodigo, iCodFilial)
            If lErro <> SUCESSO Then gError 131929
            
        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            Filial.Clear

            Call CF2(Me, "NFCF_Limpa_Tela")

        End If

        iClienteAlterado = 0

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True
    
    Select Case gErr

        Case 35265, 131929 'Inserido por Wagner
            
        Case 35266

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156477)

    End Select

    Exit Sub

End Sub

Public Sub Recebimento_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iTipoNFiscal As Integer

On Error GoTo Erro_Recebimento_Click

    'Se o Número do Recebimento estiver preenchido
    If Len(Trim(NumRecebimento.Text)) > 0 Then
        
        objNFiscal.lNumRecebimento = NumRecebimento.Text
        
        'Le o Recebimento
        lErro = CF("NFiscal_Le_Recebimento", objNFiscal)
        If lErro <> SUCESSO And lErro <> 61050 Then gError 61061
        
        'Se não encontrou --> Erro
        If lErro = 61050 Then gError 61062

        'Lê o Tipo de documento
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 27623 Then gError 75423
        
        'Se o Tipo de documento é de compras
        If objTipoDocInfo.sNomeTelaNFiscal = RECEBIMENTO_MATERIAL_FCOM Or objTipoDocInfo.sNomeTelaNFiscal = NOTA_FISCAL_COMPRAS Or objTipoDocInfo.sNomeTelaNFiscal = NOTA_FISCAL_FATURA_COMPRAS Then gError 75424

    Else
        
        'Verifica se o tipo está Preenchido
        If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 42533
        
        objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
    
        'Lê o Tipo de documento
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 27623 Then gError 61077
        
        'Se não encontrar --> erro
        If lErro = 27623 Then gError 61078
        
        'Se o Tipo de documento é de compras
        If objTipoDocInfo.sNomeTelaNFiscal = RECEBIMENTO_MATERIAL_FCOM Or objTipoDocInfo.sNomeTelaNFiscal = NOTA_FISCAL_COMPRAS Or objTipoDocInfo.sNomeTelaNFiscal = NOTA_FISCAL_FATURA_COMPRAS Then gError 75425
        
        objNFiscal.iTipoDocInfo = objTipoDocInfo.iTipo
        
        'Se for Nota Fiscal Interna então o NumRecebimento tem que está Preenchido
        If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then gError 61063

        If Fornecedor.Visible = True Then
            If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 35268
        Else
            If Len(Trim(Cliente.ClipText)) = 0 Then gError 35269
        End If
    
        If Len(Trim(Filial.Text)) = 0 Then gError 35270
    
        If Len(Trim(Serie.Text)) = 0 Then gError 35271
    
        If Len(Trim(NFiscal.ClipText)) = 0 Then gError 35272
        
        'Recolhe os dados da Tela
        lErro = Move_NF_Memoria(objNFiscal, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 35273

        Call TipoNFiscal_Correspondente_Recebimento(objTipoDocInfo, iTipoNFiscal)
        objNFiscal.iTipoNFiscal = iTipoNFiscal

        'Lê  a Nota Fiscal
        lErro = CF("NFiscal_Le_NumFornCli", objNFiscal)
        If lErro <> SUCESSO And lErro <> 35279 Then gError 35280
    
        'se a nota não está cadastrada ==> erro
        If lErro = 35279 Then gError 35281

    End If
    
    'Preenche a tela com os dados do Recebimento
    lErro = Preenche_Tela_Recebimento(objNFiscal)
    If lErro <> SUCESSO Then gError 35282

    Call gobjTribTab.TributacaoRecalcular_Click
    
    Exit Sub

Erro_Recebimento_Click:

    Select Case gErr

        Case 35267
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 35268
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 35269
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 35270
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 35271
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 35272
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)

        Case 35273, 35280, 35282, 61061, 61077, 75423

        Case 35281
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_MATERIAL_NAO_CADASTRADO2", gErr, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 61062
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NAO_CADASTRADO", gErr, objNFiscal.lNumRecebimento)
                
        Case 61063
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NFINT_NAO_PREENCHIDO", gErr)
        
        Case 61078
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO2", gErr, objTipoDocInfo.iCodigo)
        
        Case 75424, 75425
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_RECEBIMENTO_COMPRAS", gErr, objTipoDocInfo.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156478)

    End Select

    Exit Sub

End Sub

Private Function Preenche_Tela_Recebimento(objNFiscal As ClassNFiscal) As Long
'Preenche alguns campos da tela com os dados do recebimento passado em objNFiscal
'Só não altera o campo de cabeçalho

Dim lErro As Long
Dim objNFiscalOriginal As New ClassNFiscal
Dim objItemNF As New ClassItemNF
Dim iIndice As Integer
Dim bCancel As Boolean
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Preenche_Tela_Recebimento

    'Lê os Ítens
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 61090
    
    'distribuicao
    'Lê a Distribuição dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 89610
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 61091

    'Se não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 61092

    'Se possui Nota Fiscal Original
    If objNFiscal.lNumIntNotaOriginal > 0 Then

        objNFiscalOriginal.lNumIntDoc = objNFiscal.lNumIntNotaOriginal

        'Le a Nota Fiscal Original
        lErro = CF("NFiscal_Le", objNFiscalOriginal)
        If lErro <> SUCESSO And lErro <> 35285 Then gError 61093

        If lErro <> SUCESSO Then gError 61094

        'Coloca na Tela os dados encontrados
        SerieNFiscalOriginal.Text = objNFiscalOriginal.sSerie
        NFiscalOriginal.Text = objNFiscalOriginal.lNumNotaFiscal

        'Preenche a SerieNFOrignal e NumNFOriginal dos Ítens com os dados da Nota Fiscal Original
        For Each objItemNF In objNFiscal.ColItensNF
            objItemNF.sSerieNFOrig = objNFiscalOriginal.sSerie
            objItemNF.lNumNFOrig = objNFiscalOriginal.lNumNotaFiscal
        Next
    End If

    'Limpa Todos os compos da Tela menos os do Cabeçalho
    Call Limpa_Tela_NFiscal
    
    Set gobjNFiscal = objNFiscal
        
    NumRecebimento.Text = objNFiscal.lNumRecebimento

    DataEntrada.PromptInclude = False
    If objNFiscal.dtDataEntrada <> DATA_NULA Then DataEntrada.Text = Format(objNFiscal.dtDataEntrada, "dd/mm/yy")
    DataEntrada.PromptInclude = True
    
    'Preenche a série
    If Len(Trim(objNFiscal.sSerie)) > 0 Then
        Serie.Text = Desconverte_Serie_Eletronica(objNFiscal.sSerie)
        If ISSerieEletronica(objNFiscal.sSerie) Then
            EletronicaFed.Value = vbChecked
        Else
            EletronicaFed.Value = vbUnchecked
        End If
    Else
        Serie.Text = ""
    End If
    Call Serie_Validate(bSGECancelDummy)

    'Preeche o Número da Nota
    If objNFiscal.lNumNotaFiscal = 0 Then
        NFiscalInterna.Caption = ""
        NFiscal.Text = ""
    Else
        NFiscalInterna.Caption = CStr(objNFiscal.lNumNotaFiscal)
        NFiscal.Text = CStr(objNFiscal.lNumNotaFiscal)
    End If
    
    'Preenche a Tela de Acordo com a Origem do Tipo da Nota Fiscal
    If objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
        Cliente.Text = objNFiscal.lCliente
        Call Cliente_Validate(bCancel)
        Filial.Text = objNFiscal.iFilialCli
    ElseIf objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        Fornecedor.Text = objNFiscal.lFornecedor
        Call Fornecedor_Validate(bCancel)
        Filial.Text = objNFiscal.iFilialForn
    End If

    'Preenche a Filial
    Call Filial_Validate(bCancel)
    
    'Modifica o Tipo Verificando se é Externa ou interna e Cliente ou Fornecedor
    If objNFiscal.iTipoNFiscal = DOCINFO_NRFF Then
        objNFiscal.iTipoNFiscal = DOCINFO_NFEEDCNT
    ElseIf objNFiscal.iTipoNFiscal = DOCINFO_NRFP Then
        objNFiscal.iTipoNFiscal = DOCINFO_NFIEDCNT
    ElseIf objNFiscal.iTipoNFiscal = DOCINFO_NRCC Then
        objNFiscal.iTipoNFiscal = DOCINFO_NFEEDV
    ElseIf objNFiscal.iTipoNFiscal = DOCINFO_NRCP Then
        objNFiscal.iTipoNFiscal = DOCINFO_NFIEDV
    End If
    
    'Preenche o Tipo
    TipoNFiscal.Text = objNFiscal.iTipoNFiscal
    Call TipoNFiscal_Validate(bCancel)
    
    If objNFiscal.iCodTransportadora <> 0 Then
        Transportadora.Text = objNFiscal.iCodTransportadora
        Transportadora_Validate bCancel
    End If
    
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

    Mensagem.Text = objNFiscal.sMensagemNota
    MensagemCorpo.Text = objNFiscal.sMensagemCorpoNota
    
    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = CStr(objNFiscal.dPesoLiq)
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = CStr(objNFiscal.dPesoBruto)

    If objNFiscal.dValorDesconto > 0 Then ValorDesconto.Text = Format(objNFiscal.dValorDesconto, "Standard")
    If objNFiscal.dValorOutrasDespesas > 0 Then ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Standard")
    If objNFiscal.dValorFrete > 0 Then ValorFrete.Text = Format(objNFiscal.dValorFrete, "Standard")
    If objNFiscal.dValorSeguro > 0 Then ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Standard")
    If objNFiscal.dValorTotal > 0 Then Total.Text = Format(objNFiscal.dValorTotal, "Standard")
    If objNFiscal.dValorProdutos > 0 Then SubTotal.Caption = Format(objNFiscal.dValorProdutos, "Standard")

    'Preenche o Grid com os ítens lidos
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 61095

    'distribuicao
    'Preenche o Grid com as Distribuições dos itens da Nota Fiscal
    lErro = gobjDistribuicao.Preenche_GridDistribuicao(objNFiscal)
    If lErro <> SUCESSO Then gError 89612

    Call gobjTribTab.ValorFrete_Validate(bSGECancelDummy, objNFiscal.dValorFrete)
    Call gobjTribTab.ValorSeguro_Validate(bSGECancelDummy, objNFiscal.dValorSeguro)
    Call gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, objNFiscal.dValorDesconto)
    Call gobjTribTab.ValorDespesas_Validate(bSGECancelDummy, objNFiscal.dValorOutrasDespesas)
        
    For iIndice = 1 To objNFiscal.ColItensNF.Count
        lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, objNFiscal.ColItensNF.Item(iIndice).sProduto)
        If lErro <> SUCESSO Then gError 61096
    Next
    
    Preenche_Tela_Recebimento = SUCESSO

    Exit Function

Erro_Preenche_Tela_Recebimento:

    Preenche_Tela_Recebimento = gErr

    Select Case gErr

        Case 61090, 61091, 61093, 61095, 61096, 89610, 89612

        Case 61092
           Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case 61094
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscalOriginal.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156479)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim iProdutoPreenchidoBenef As Integer
Dim sProdutoFormatadoBenef As String

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name

        Case Produto.Name
        
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 42922
            
            If iProdutoPreenchido <> PRODUTO_VAZIO Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
            End If

        Case UnidadeMed.Name

            UnidadeMed.Clear

            'Guarda a UM que está no Grid
            sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 35289
            
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                UnidadeMed.Enabled = False

            Else
                UnidadeMed.Enabled = True
                objProduto.sCodigo = sProdutoFormatado
                
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 35289
                
                'se o produto não está cadastrado ==> erro
                If lErro = 28030 Then gError 35290

                objClasseUM.iClasse = objProduto.iClasseUM
                
                'Lâ as Unidades de Medidas da Classe do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 35291
                
                'Carrega a combo de UM
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next
                
                'Seleciona na UM que está preenchida
                UnidadeMed.Text = sUM
                If Len(Trim(sUM)) > 0 Then
                    lErro = Combo_Item_Igual(UnidadeMed)
                    If lErro <> SUCESSO And lErro <> 12253 Then gError 35292
                End If
            End If


        Case ValorUnitario.Name, PercentDesc.Name, Desconto.Name, SerieNFOriginal.Name, NumNFOriginal.Name, ItemNFOriginal.Name, DescricaoItem.Name
        
            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 35293

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        
        Case Quantidade.Name
    
            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 31474
            
            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Or left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
    
        Case ItemNFBenef.Name
            
            If Len(Trim(GridMovimentos.TextMatrix(iLinha, iGrid_ItemNFBenef_Col))) > 0 Then
                ItemNFBenef.Enabled = False
            Else
                ItemNFBenef.Enabled = True
            End If
    
        Case ProdutoBenef.Name, AlmoxarifadoBenef.Name, QuantInsumoBenef.Name, ContaContabilEstBenef.Name, ContaContabilProducaoBenef.Name
            
            If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ItemNFBenef_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
    
    End Select
    
    lErro = CF2(Me, "NFEDEV_Rotina_Grid_Enable", iLinha, objControl, iLocalChamada)
    If lErro <> SUCESSO Then gError 177036

    'Rastreamento
    lErro = gobjRastreamento.Rotina_Grid_Enable(iLinha, objControl, iLocalChamada)
    If lErro <> SUCESSO Then gError 83289

    'distribuicao
    lErro = gobjDistribuicao.Rotina_Grid_Enable_Dist(iLinha, objControl, iLocalChamada)
    If lErro <> SUCESSO Then gError 89560

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 31474, 35289, 35291, 35292, 35289, 35293, 35294, 42922, 83289, 89560, 126853, 177036

        Case 35290
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156480)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iPossuiGrade As Integer

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    'Verifica se o Produto esta preenchida
    If Len(Trim(Produto.ClipText)) > 0 Then

        sProduto = Produto.Text

        lErro = CF("Trata_Segmento_Produto", sProduto)
        If lErro <> SUCESSO Then gError 199351

        Produto.Text = sProduto

        lErro = CF("Produto_Critica2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 31479
        
        If lErro = 25041 Then gError 31480
        
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 55139
        
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True

        'Coloca as demais características do produto na tela
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then gError 35657
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31481

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 31479, 31481, 35657, 199351
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 31480
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 55139
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156481)

    End Select

    Exit Function

End Function

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long
'Preenche os dados do Produto da linha do grid selecionada

Dim lErro As Long
Dim iAlmoxarifadoPadrao As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim iPossuiGrade As Integer
Dim sProdutoPai As String
Dim colItensRomaneioGrade As New Collection
Dim objItensRomaneio As ClassItemRomaneioGrade
Dim sProduto As String
Dim objRomaneioGrade As New ClassRomaneioGrade
Dim iIndice As Integer
Dim objItemNF As ClassItemNF

On Error GoTo Erro_ProdutoLinha_Preenche

    If objProduto.iGerencial And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then gError 86296
    
    If Len(Trim(objProduto.sGrade)) > 0 Then iPossuiGrade = MARCADO

    If iPossuiGrade = DESMARCADO And objProduto.iKitVendaComp <> MARCADO Then
    
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
            
    Set objItemNF = New ClassItemNF
    
    objItemNF.iPossuiGrade = iPossuiGrade
        
    objItemNF.sProduto = objProduto.sCodigo
    objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
    objItemNF.iItem = GridItens.Row
    objItemNF.lNumIntDoc = 0
    objItemNF.sDescricaoItem = objProduto.sDescricao
                
    If objItemNF.iPossuiGrade = MARCADO Then
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

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridItens.Row - GridItens.FixedRows) = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    
''*** incluidos p/tratamento de tributacao *******************************
        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 55653
'*** fim tributacao

        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItens.Row)

        gobjNFiscal.ColItensNF.Add1 objItemNF
        gobjNFiscal.ColItensNF(GridItens.Row).sUMEstoque = objProduto.sSiglaUMEstoque
        gobjNFiscal.ColItensNF(GridItens.Row).iItem = GridItens.Row
        gobjNFiscal.ColItensNF(GridItens.Row).sProduto = objProduto.sCodigo
        gobjNFiscal.ColItensNF(GridItens.Row).sUnidadeMed = objProduto.sSiglaUMCompra
       
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
        
    End If
    
    lErro = CF2(Me, "NFEDEV_ProdutoLinha_Preenche", objProduto)
    If lErro <> SUCESSO Then gError 177054

    Call Calcula_Valores(GridItens.Row)
    
    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case 30317, 30318, 55653, 86310, 86327, 86329, 177054, 177582

        Case 30319
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)

        Case 86296
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)

        Case 86328
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_GRID", gErr, Trim(sProdutoPai), Trim(Produto.Text))
                    
        Case 86330
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, Produto.Text)
        
        Case 86331
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILHO_GRADE_GRID", gErr, Trim(objProduto.sCodigo), Trim(gobjNFiscal.ColItensNF(iIndice).sProduto))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156482)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
'distribuicao
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

    'Se quantidade estiver preenchida
    If Len(Trim(Quantidade.ClipText)) > 0 Then
    
        'Critica o valor
        If giComplemento = DESMARCADO Then
            lErro = Valor_Positivo_Critica(Quantidade.Text)
        Else
            lErro = Valor_NaoNegativo_Critica(Quantidade.Text)
        End If
        If lErro <> SUCESSO Then gError 35299

        dQuantidade = CDbl(Quantidade.Text)

        'Coloca o valor Formatado na tela
        Quantidade.Text = Formata_Estoque(dQuantidade)

        'distribuicao
        dQuantidadeAtual = StrParaDbl(Quantidade.Text)
        'fim  distribuicao

    End If

    'inicio distribuicao
    If dQuantidadeAnterior <> dQuantidadeAtual Then
        
        GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)
        
        objProduto.sCodigo = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 180040
        
        GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)
        
        If gobjNFiscal.ColItensNF(GridItens.Row).iPossuiGrade <> MARCADO And objProduto.iKitVendaComp <> MARCADO Then
        
            'Tenta fazer a distribuição automatica p\ o item de acordo com a nova UM
            lErro = gobjDistribuicao.Distribuicao_Processa()
            If lErro <> SUCESSO Then gError 89562
            
        Else
        
            objKitVenda.sProduto = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
            objKitVenda.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
            objKitVenda.dQuantidade = dQuantidade
        
            lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade)
            If lErro <> SUCESSO Then gError 180041

            lErro = gobjDistribuicao.Distribuicao_Processa_Grade()
            If lErro <> SUCESSO Then gError 89562
        
            Call gobjDistribuicao.Atualiza_Grid_Distribuicao(gobjNFiscal.ColItensNF.Item(GridItens.Row))
        End If
        
    End If
    'fim distribuicao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 35300

    lErro = CF2(Me, "NFEDEV_Calcula_PesoKg", GridItens.Row)
    If lErro <> SUCESSO Then gError 177055

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 55586

    '*** incluidos p/tratamento de tributacao ****************
    If GridItens.Row <= objGridItens.iLinhasExistentes Then
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 22843
    End If
    '*** fim tributacao

    If GridItens.Row <= objGridItens.iLinhasExistentes Then
    
        If gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade.Count = 0 Then
    
            lErro = gobjRastreamento.Atualiza_Troca_Quantidade(GridItens.Row)
            If lErro <> SUCESSO Then gError 83291
        
        Else
        
            If dQuantidadeAnterior <> 0 Then
                dFator = dQuantidadeAtual \ dQuantidadeAnterior
            Else
                dFator = 1
            End If
            
            lErro = gobjRastreamento.Atualiza_Troca_Quantidade_Grade(GridItens.Row, dFator)
            If lErro <> SUCESSO Then gError 83291
        
        End If
    
    End If

    Call ValorTotal_Calcula
    
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 83291
    End If
    
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 22843, 35299, 35300, 55586, 83291, 89562, 177055, 180040, 180041
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156483)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorUnitario(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim dValorUnitario As Double
Dim dValorReal As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim dSubTotal As Double
Dim lTamanho As Long
'Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Saida_Celula_ValorUnitario

    Set objGridInt.objControle = ValorUnitario

    'Se estiver preenchido
    If Len(Trim(ValorUnitario.ClipText)) > 0 Then
        
        If giComplemento = MARCADO Then
            lErro = Valor_NaoNegativo_Critica(ValorUnitario.Text)
        Else
            lErro = Valor_Positivo_Critica(ValorUnitario.Text)
        End If
        
        If lErro <> SUCESSO Then gError 35306
            
        dValorUnitario = CDbl(ValorUnitario.Text)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 35307

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 55577
    
    ''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 22844
    '*** fim tributacao

    Call ValorTotal_Calcula
    
    Saida_Celula_ValorUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorUnitario:

    Saida_Celula_ValorUnitario = gErr

    Select Case gErr

        Case 22844, 35306, 35307, 55577, 130758
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 130759
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156485)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercDesc(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim lTamanho As Long
Dim dPercentDescAnterior As Double

On Error GoTo Erro_Saida_Celula_PercDesc

    Set objGridInt.objControle = PercentDesc

    dPercentDesc = StrParaDbl(PercentDesc.Text)
    
    lTamanho = Len(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col))
    If lTamanho > 0 Then dPercentDescAnterior = StrParaDbl(left(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col), lTamanho - 1))

    'verifica se o percentual está preenchido
    If Len(Trim(PercentDesc.Text)) > 0 Then
    
        'Critica a procentagem
        lErro = Porcentagem_Critica(PercentDesc.Text)
        If lErro <> SUCESSO Then gError 52266

'        dPercentDesc = CDbl(PercentDesc.Text)
'
'        lTamanho = Len(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col))
'        If lTamanho > 0 Then dPercentDescAnterior = PercentParaDbl(left(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col), lTamanho - 1))

        If dPercentDesc <> dPercentDescAnterior Then

            'Verifica se o percentual é de 100%
            If dPercentDesc = 100 Then gError 52267

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
     If lErro <> SUCESSO Then gError 52268

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 55580

    ''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 52269
    '*** fim tributacao

    Call ValorTotal_Calcula
    
    Saida_Celula_PercDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercDesc:

    Saida_Celula_PercDesc = gErr

    Select Case gErr

        Case 52266, 52268, 52269, 55580
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52267
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156486)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercentDesc As Double
Dim dValorTotal As Double
Dim dDesconto As Double
Dim dValorUnitario As Double
Dim dValorReal As Double
Dim dQuantidade As Double
Dim dSubTotal As Double
Dim lTamanho As Long

On Error GoTo Erro_Saida_Celula_Desconto

    Set objGridInt.objControle = Desconto
    
    dDesconto = 0

    'Se estiver preenchido
    If Len(Trim(Desconto.ClipText)) > 0 Then
    
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(Desconto.Text)
        If lErro <> SUCESSO Then gError 35311

        dDesconto = CDbl(Desconto.Text)

        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) > 0 Then dQuantidade = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) > 0 Then dValorUnitario = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))

        dValorTotal = dValorUnitario * dQuantidade

        'Calcula o novo Percentual de Desconto
        If dValorTotal > 0 Then

            dPercentDesc = dDesconto / (dValorTotal)

            If dPercentDesc >= 1 Then gError 35798

        End If

        GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")

    Else
    
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) <> 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) <> 0 Then

            GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col) = ""
        
        End If

    End If
    
    'Verifica se o valor de desconto foi alterado
    If Abs(StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Desconto_Col)) - dDesconto) > DELTA_VALORMONETARIO Then
        iDescontoAlterado = REGISTRO_ALTERADO
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 35312

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 55583

    ''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 22846
    '*** fim tributacao

    Call ValorTotal_Calcula
    
    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = gErr

    Select Case gErr

        Case 22846, 35311, 35312, 55583
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 35798
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156487)

    End Select

    Exit Function

End Function

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
        If lErro <> SUCESSO And lErro <> 22202 Then gError 35313
        
        'Se a série não está cadastrada ==> erro
        If lErro = 22202 Then gError 35314

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 35315

    Saida_Celula_SerieNFOriginal = SUCESSO

    Exit Function

Erro_Saida_Celula_SerieNFOriginal:

    Saida_Celula_SerieNFOriginal = gErr

    Select Case gErr

        Case 35313, 35315
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 35314
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156488)

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
        If lErro <> SUCESSO Then gError 35320

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 35321

    Saida_Celula_NumNFOrig = SUCESSO

    Exit Function

Erro_Saida_Celula_NumNFOrig:

    Saida_Celula_NumNFOrig = gErr

    Select Case gErr

        Case 35320, 35321
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156489)

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
        If lErro <> SUCESSO Then gError 35322

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 35323

    Saida_Celula_ItemNFOrig = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemNFOrig:

    Saida_Celula_ItemNFOrig = gErr

    Select Case gErr

        Case 35322, 35323
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156490)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidade de Medida que está deixando de ser a corrente

Dim lErro As Long
Dim sUMAnterior As String
Dim objProduto As New ClassProduto
Dim objKitVenda As New ClassKitVenda

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

'inicio distribuicao
    'recolhe a UM anteriormente escolhida
    sUMAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
    
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text
    
    'coloca no grid a UM atual selecionda
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    gobjNFiscal.ColItensNF(GridItens.Row).sUnidadeMed = UnidadeMed.Text
    
    If sUMAnterior <> UnidadeMed.Text Then
        
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
        Else
            lErro = gobjRastreamento.Atualiza_Troca_UM_Grade(GridItens.Row, sUMAnterior, UnidadeMed.Text, objProduto)
            If lErro <> SUCESSO Then gError 83206
        End If
    End If
    
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 177057
    End If

    lErro = CF2(Me, "NFEDEV_Calcula_PesoKg", GridItens.Row)
    If lErro <> SUCESSO Then gError 177057

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 35083, 83206, 89511, 89602, 177057, 177583, 177584
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156491)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 35337

    'Limpa a Tela
    Call Limpa_Tela_NFiscal

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 35337

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156492)

    End Select

    Exit Sub

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 35338

    'Limpa a Tela
    Call Limpa_Tela_NFiscal2
    
    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 35338

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156493)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Dispara o processo de gravação de uma Nota Fiscal de Entrada e Devolução

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iIndice As Integer
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dTotal As Double
Dim objNFiscalOriginal As New ClassNFiscal
Dim vbMsgRes As VbMsgBoxResult
Dim dQuantLote As Double
Dim objItemNF As New ClassItemNF
Dim objRastroItemNF As ClassRastroItemNF

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    gbGravando = True
    
    'Verifica se os campos obrigatórios da tela estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 35339

    If Len(Trim(NaturezaOp.ClipText)) = 0 Then gError 35340

    If Len(Trim(Total.Text)) > 0 Then dTotal = CDbl(Total.Text)

    'Verifica se o Total é negativo
    If dTotal < 0 Then gError 35360

    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
    
    'Lê o Tipo da NF
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 35341

    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 35342

    'De acordo com a sua Origem verifica se o Cliente ou Fornecedor estão preenchidos
    If objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
        If Len(Trim(Cliente.ClipText)) = 0 Then gError 35343
    ElseIf objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 35344
    End If

    If Len(Trim(Filial.Text)) = 0 Then gError 35345
    If Len(Trim(Serie.Text)) = 0 Then gError 35346
    
    'Critica se é uma Nota Fiscal Interna
    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
        'Se for Interna --> não pode preencher o campo Número
'        If Len(Trim(NFiscalInterna.Caption)) > 0 Then gError 60451
        If Len(Trim(NFiscalInterna.Caption)) > 0 Then
            objNFiscal.lNumNotaFiscal = CLng(NFiscalInterna.Caption)
        Else
            objNFiscal.lNumNotaFiscal = 0
        End If
    Else
        'Se for Externa --> O campo Número tem que estar preenchido
        If Len(Trim(NFiscal.ClipText)) = 0 Then gError 35347
        objNFiscal.lNumNotaFiscal = CLng(NFiscal.Text)
    End If
    
    objNFiscal.iTipoDocInfo = objTipoDocInfo.iTipo
    
    If Len(Trim(DataEntrada.ClipText)) = 0 Then gError 35348
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 35349

    'Verifica se a Data de Entrada é menor que a Data de Emissao
    If CDate(DataEntrada.Text) < CDate(DataEmissao.Text) Then gError 35350

    If Len(Trim(PesoLiquido.ClipText)) > 0 And Len(Trim(PesoBruto.ClipText)) > 0 Then
        If CDbl(PesoLiquido.Text) > CDbl(PesoBruto.Text) Then gError 35210
    End If

    'se o tab de beneficiamento existe
    If TabStrip1.Tabs.Count = TAB_BENEFICIAMENTO Then
    
        'Valida os dados do grid de beneficiamento
        For iIndice = 1 To objGridBenef.iLinhasExistentes
            
            'Verifica se a Quantidade foi informada
            If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_QuantInsumoBenef_Col))) = 0 Then gError 126848

            'Verifica se o Produto foi informado
            If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_ProdutoBenef_Col))) = 0 Then gError 126850

            'Verifica se o Almoxarifado foi informado
            If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxarifadoBenef_Col))) = 0 Then gError 126910

        Next
        
    End If

    'Guarda todos os campos da tela em objNFiscal
    lErro = Move_NF_Memoria(objNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO Then gError 35363

    'Recolhe os Ítens do Grid
    lErro = Move_GridItens_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 35377

    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786

    lErro = gobjTribTab.Valida_Dados
    If lErro <> SUCESSO Then gError 181786

    'distribuicao
    lErro = gobjDistribuicao.Move_GridDist_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 89557

    'Verifica se o Grid está vazio
    If objGridItens.iLinhasExistentes = 0 Then gError 35353

    For iIndice = 1 To objGridItens.iLinhasExistentes

        'Verifica se os campos do Grid estão todos preenchidos
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 35354
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) = 0 Then gError 35359
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col))) = 0 Then gError 49387

        Set objItemNF = objNFiscal.ColItensNF(iIndice)
        
        objProduto.sCodigo = objItemNF.sProduto

        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 35356

        'Se não achou o Produto --> erro
        If lErro = 28030 Then gError 35357

        objItemNF.iClasseUM = objProduto.iClasseUM
        objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
        objItemNF.sUMVenda = objProduto.sSiglaUMVenda

        'Verifica se a Unidade de Medida foi preenchida
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) = 0 Then gError 55631

    Next

    'Se os dados da Nota Fiscal Original estão preenchidos
    If Len(Trim(SerieNFiscalOriginal.Text)) > 0 And Len(Trim(NFiscalOriginal.ClipText)) > 0 Then

        objNFiscalOriginal.sSerie = SerieNFiscalOriginal.Text
        objNFiscalOriginal.lNumNotaFiscal = NFiscalOriginal.Text
        objNFiscalOriginal.iFilialEmpresa = objNFiscal.iFilialEmpresa
        
        'Lê a Nota Fiscal Original
        If objTipoDocInfo.iCodigo = DOCINFO_NFIEDCSM Or objTipoDocInfo.iCodigo = DOCINFO_NFEEDCSM Then
            'Aceita NFs de devolução saída como NF original
            lErro = CF("NFiscalInternaSaida_Le_Numero2", objNFiscalOriginal)
            If lErro <> AD_SQL_SUCESSO And lErro <> 62144 Then gError 35361
        Else
            lErro = CF("NFiscalInternaSaida_Le_Numero", objNFiscalOriginal)
            If lErro <> AD_SQL_SUCESSO And lErro <> 30765 Then gError 35361
        End If
        'Se não encontrou a NFiscal Original --> erro
        If lErro <> SUCESSO Then gError 35362

        'Guarda em objNFiscal o Número Interno da Nota Fiscal Original
        objNFiscal.lNumIntNotaOriginal = objNFiscalOriginal.lNumIntDoc

    End If

    lErro = gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
    If lErro <> SUCESSO Then gError 20808

    'Recolhe os dados do Grid Movimentos
    If (objNFiscal.iTipoNFiscal = DOCINFO_NFIEDSB) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEDSB) Then
    
        lErro = Move_GridMovimentos_Memoria(objNFiscal)
        If lErro <> SUCESSO Then gError 126896
        
    End If

    Set objNFiscal.objRastreamento = gobjRastreamento

    lErro = Total_Calcula()
    If lErro <> SUCESSO Then gError 43787

    'Mover a parte de tributação
    lErro = gobjTribTab.PreencherNF(objNFiscal)
    If lErro <> SUCESSO Then gError 35262

    lErro = gobjTribTab.Valida_TributacaoNF_Gravacao(objNFiscal)
    If lErro <> SUCESSO Then gError 59001
    
    lErro = Calcula_Msg
    If lErro <> SUCESSO Then gError 59001
    
    'mover a parte do rastreamento
    lErro = gobjRastreamento.Move_Rastro_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 20809
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataEntrada.Text))
    If lErro <> SUCESSO Then gError 92032
    
    lErro = CF2(Me, "NFEnt_Move_Tela_Memoria", objNFiscal)
    If lErro <> SUCESSO Then gError 124253
    
    lErro = CF("NFiscal_Valida_Diversos", objNFiscal)
    If lErro <> SUCESSO Then gError 124253
    
    'Chama a rotina para iniciar a Gravação (gravando inclusive os dados contábeis)
    lErro = CF("NFiscalEntradaDevolucao_Grava", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 35364
    
    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_NFISCAL
        gobjAnotacao.sID = CStr(objNFiscal.lNumIntDoc)
        gobjAnotacao.dtDataAlteracao = gdtDataHoje
        
        lErro = CF("Anotacoes_Grava", gobjAnotacao)
        If lErro <> SUCESSO Then gError 130406
        
    End If
    
    gbGravando = False
    GL_objMDIForm.MousePointer = vbDefault
    
    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
    
        If Len(Trim(NFiscalInterna.Caption)) = 0 Then vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_NUMERO_NOTA_GRAVADA", objNFiscal.lNumNotaFiscal)
    
    End If
    
    If ISSerieEletronica(objNFiscal.sSerie) Then
        lErro = CF("NFE_Grava1", objNFiscal.lNumIntDoc, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 203001
    End If

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    gbGravando = False
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 35210
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr, CDbl(PesoLiquido.Text), CDbl(PesoBruto.Text))

        Case 20808, 20809, 35262, 35341, 35355, 35356, 35361, 35363, 35364, 35377, 43787, 59001, 89557, 92032, 124253, 126896, 130406, 181786, 203001

        Case 35339
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 35340
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 35342
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case 35343
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 35344
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 35345
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 35346
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 35347
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)

        Case 35348
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_NAO_PREENCHIDA", gErr)

        Case 35349
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 35350
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_ANTERIOR_DATAEMISSAO", gErr, DataEntrada.Text, DataEmissao.Text)

        Case 35353
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)

        Case 35354
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case 35357
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 35358
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 35359
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 35360
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)

        Case 35362
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_INTERNA_SAIDA_NAO_CADASTRADA", gErr, objNFiscalOriginal.sSerie, objNFiscalOriginal.lNumNotaFiscal)

        Case 49387
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAOITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case 55631
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDA", gErr, iIndice)

        Case 75925
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTTOTAL_LOTE_MAIOR_ALMOXARIFADO", gErr, dQuantLote, StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)))
                
        Case 126848
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_BENEF_NAO_PREENCHIDA", gErr, iIndice)
            
        Case 126850
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_BENEF_NAO_PREENCHIDO", gErr, iIndice)
            
        Case 126910
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_BENEF_NAO_PREENCHIDO", gErr, iIndice)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156494)

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
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 89062
    If Len(Trim(Serie.Text)) = 0 Then gError 89063
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 89064

    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
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
    
    'se for nfiscal interna
    If objTipoDocInfo.iOrigem = DOCINFO_EMPRESA Then
        'Verifica se a existe nota fiscal está cadastrada
        lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
        If lErro <> SUCESSO And lErro <> 43676 Then gError 99352

        If lErro = 43676 Then gError 89065

'        If gobjCRFAT.iUsaNFe = MARCADO Then
'
'            lErro = CF("NFeFedLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'            If lErro <> SUCESSO Then gError 203038
'
'            If iEnviado = 1 Then gError 203039
'
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
        If lErro <> SUCESSO Then gError 89067

        'Faz o cancelamento de uma nota fiscal de Saida
        lErro = CF("NotaFiscalEntrada_Excluir", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError 89068
    
        'Limpa a Tela
        Call Limpa_Tela_NFiscal2
    
        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 89062
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 89063
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 89064
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 89065, 89067, 89068, 99345, 99348, 99351, 99352, 99353, 203038
       
        Case 89065
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
            
        Case 99346
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
        
        Case 99347
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 99349
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objNFiscal.lCliente)
        
        Case 99350
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 99352
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)
       
        Case 203039
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CADASTRADA_NAO_EXCLUI", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156495)
        
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
Dim lTamanho As Long
Dim colAlocacoes As ColAlocacoesItemNF
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objItemNFOrig As New ClassItemNF
Dim objNFOrig As New ClassNFiscal
Dim objItemNFAux As ClassItemNF 'Inserido por Wagner 21/03/2006

On Error GoTo Erro_Move_GridItens_Memoria

    'Para cada linha existente do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemNF = New ClassItemNF

        'Verifica se o Produto está preenchido
        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 35380
        'Armazena produto
        objItemNF.sProduto = sProdutoFormatado

        'Guarda os demais campos do Grid em objItemNF
        objItemNF.sDescricaoItem = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
        objItemNF.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then objItemNF.dQuantidade = CDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) > 0 Then objItemNF.dPrecoUnitario = CDbl(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))
        lTamanho = Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col)))
        If lTamanho > 0 Then objItemNF.dPercDesc = PercentParaDbl(GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col))

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))) > 0 Then objItemNF.dValorDesconto = CDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))

        'Verifica se Ccl foi preenchido
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Ccl_Col))) <> 0 Then

            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", GridItens.TextMatrix(iIndice, iGrid_Ccl_Col), sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError 49415

        Else
            sCclFormatada = ""
        End If

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_SerieNFOrig_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_NumNFOrig_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ItemNFOrig_Col))) > 0 Then

            Set objItemNFOrig = New ClassItemNF

            objItemNFOrig.sSerieNFOrig = GridItens.TextMatrix(iIndice, iGrid_SerieNFOrig_Col)
            objItemNFOrig.lNumNFOrig = GridItens.TextMatrix(iIndice, iGrid_NumNFOrig_Col)
            objItemNFOrig.iItem = GridItens.TextMatrix(iIndice, iGrid_ItemNFOrig_Col)

            If gbGravando Then
            
                If objNFiscal.iTipoNFiscal = DOCINFO_NFIEDCSM Or objNFiscal.iTipoNFiscal = DOCINFO_NFEEDCSM Then
                    lErro = CF("ItemNFiscalSaida_Le_NumNFItem1", objItemNFOrig)
                    If lErro <> SUCESSO And lErro <> 199746 Then gError 35369
                Else
                    lErro = CF("ItemNFiscalSaida_Le_NumNFItem", objItemNFOrig)
                    If lErro <> SUCESSO And lErro <> 35368 Then gError 35369
                End If
                If lErro <> SUCESSO Then gError 35375
                
                If objItemNFOrig.sProduto <> objItemNF.sProduto Then gError 35376
    
                objItemNF.lNumIntDocOrig = objItemNFOrig.lNumIntDoc
            
                objNFOrig.lNumIntDoc = objItemNFOrig.lNumIntNF
            
                lErro = CF("NFiscal_Le", objNFOrig)
                If lErro <> SUCESSO And lErro <> 31442 Then gError 35376
                
                If objNFOrig.dtDataEmissao > objNFiscal.dtDataEmissao Then gError 206614
            
            End If
        End If

        'Adiciona na coleção de Ítens
        With objItemNF
            Set objItemNFAux = objNFiscal.ColItensNF.Add(0, iIndice, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, DATA_NULA, .sDescricaoItem, 0, 0, 0, 0, 0, colAlocacoes, .iAlmoxarifado, .sAlmoxarifadoNomeRed, sCclFormatada, STATUS_LANCADO, .lNumIntDocOrig, "", 0, 0, 0, "", 0, 0)
        End With
        
        lErro = CF2(Me, "NFEDEV_Move_GridItens_Memoria", objItemNFAux, iIndice)
        If lErro <> SUCESSO Then gError 177037

            '********************* TRATAMENTO DE GRADE *****************
        lErro = gobjDistribuicao.Move_DistribuicaoGrade_Memoria(gobjNFiscal.ColItensNF(iIndice))
        If lErro <> SUCESSO Then gError 86375
        
        Call Move_ItensGrade_Tela(objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade, gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade)

        Set objItemNFAux.objInfoUsu = gobjNFiscal.ColItensNF(iIndice).objInfoUsu

        Set objItemNFAux.objInfoAdicDocItem = gobjNFiscal.ColItensNF(iIndice).objInfoAdicDocItem

    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 35369, 35374, 35380, 35381, 49415, 86375, 177037

        Case 35375
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NFORIGINAL_NAO_CADASTRADO", gErr, iIndice, objItemNFOrig.sSerieNFOrig, objItemNFOrig.lNumNFOrig)

        Case 35376
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CASADO", gErr, iIndice, objItemNF.sProduto, objItemNFOrig.sProduto)
            
        Case 206614
            Call Rotina_Erro(vbOKOnly, "ERRO_DEV_DATA_MENOR_ORIG", gErr, objNFiscal.dtDataEmissao, objNFOrig.dtDataEmissao)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156496)

    End Select

    Exit Function

End Function

Public Sub Limpa_Tela_NFiscal1()
'Limpa Todos os compos da Tela menos os do Cabeçalho

Dim objTipoDocInfo As New ClassTipoDocInfo
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_NFiscal1

    'Limpa o Grid
    Call Grid_Limpa(objGridItens)

    'Limpa o Grid
    Call Grid_Limpa(objGridBenef)

    'Limpa os campos que não são limpos pelo Limpa_Tela
    SerieNFiscalOriginal.Text = ""
    SubTotal.Caption = ""
    ValorProdutos2.Caption = ""
    Transportadora.Text = ""
    Placa.Text = ""
    PlacaUF.Text = ""
    NFiscalInterna.Caption = ""
    ValorFrete.Text = ""
    ValorSeguro.Text = ""
    ValorDespesas.Text = ""
    VolumeQuant.Text = ""
    
    VolumeMarca.Text = ""
    VolumeEspecie.Text = ""
    
    'Recarrega a combo VolumeEspecie e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)

    'Recarrega a combo VolumeMarca e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
    
    Mensagem.Text = ""
    MensagemCorpo.Text = ""
    PesoLiquido.Text = ""
    PesoBruto.Text = ""
    TipoFrete.ListIndex = 0

    Call gobjTribTab.Limpa_Tela
    
    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)

    'Lê o Tipo de documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 55887
    
    'Se não encontrar --> erro
    If lErro = 31415 Then gError 55888
    
    If objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        Call Trata_FilialForn
    ElseIf objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then

        Call Trata_FilialCliente
    End If
    
    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
    'distribuicao
    Call gobjDistribuicao.Limpa_Tela_Distribuicao
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade
    
    'retira o tab
    If TabStrip1.Tabs.Count = TAB_BENEFICIAMENTO Then
        TabStrip1.Tabs.Remove (TAB_BENEFICIAMENTO)
    End If
    
    Exit Sub

Erro_Limpa_Tela_NFiscal1:

    Select Case gErr
    
        Case 55887
    
        Case 55888
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156497)

    End Select

    Exit Sub

End Sub

Public Function Traz_NFiscal_Tela(objNFiscal As ClassNFiscal) As Long
'Coloca na tela os dados da Nota Fiscal passada em objNfiscal

Dim lErro As Long
Dim objItemNFiscal As ClassItemNF
Dim objItemNFiscalOriginal As New ClassItemNF
Dim objNFiscalOriginal As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim bCancel As Boolean
Dim bAlterouCT As Boolean
Dim objItemNF As ClassItemNF
Dim iIndice As Integer, dPercDesc As Double

On Error GoTo Erro_Traz_NFiscal_Tela

    'Se for versao LIGHT testa chaveamento
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

    'Limpa a Tela (nao pode fechar as setas)
    Call Limpa_Tela_NFiscal

    'Tenta ler a nota Fiscal passada por parametro
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 35217
    
    'Le os Ítens da Nota Fiscal passada
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 35217
    
    'Lê as mensagens
    lErro = CF("NFiscal_Le_Msg", objNFiscal)
    If lErro <> SUCESSO Then gError 35217

    'distribuicao
    'Lê a Distribuição dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 89558
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Le o Tipo para saber se tem Movimento de Estoque
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO Then gError 83286

    If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
        'Carrega ItensNF com Rastreamentos
        lErro = gobjRastreamento.Carrega_RastroItensNF(objNFiscal)
        If lErro <> SUCESSO Then gError 83287
    End If

    'Lê a parte de Tributação
    lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
    If lErro <> SUCESSO Then gError 35261

    'Lê a parte de tributacao dos Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
    If lErro <> SUCESSO Then gError 35260

    'Para cada Ítem da Nota Fiscal
    For Each objItemNFiscal In objNFiscal.ColItensNF

        objItemNFiscalOriginal.lNumIntDoc = objItemNFiscal.lNumIntDocOrig
        
        If objItemNFiscalOriginal.lNumIntDoc <> 0 Then
        
            'Lê o Ítem da NFOriginal em NFiscal
            lErro = CF("ItemNFiscal_Le", objItemNFiscalOriginal)
            If lErro <> SUCESSO And lErro <> 35225 Then gError 35226
    
            If lErro <> SUCESSO Then gError 55175
    
            objItemNFiscal.iItemNFOrig = objItemNFiscalOriginal.iItem
    
            objNFiscalOriginal.lNumIntDoc = objItemNFiscalOriginal.lNumIntNF
            
            'Lê a Nota Fiscal do Ítem Lido (Original)
            lErro = CF("NFiscal_Le", objNFiscalOriginal)
            If lErro <> SUCESSO Then gError 35227
    
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

    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 35234

    'Se não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 35235

    Set gobjNFiscal = objNFiscal
    
    Set gobjNFiscal = objNFiscal
    
    lErro = CF("NFiscal_Le_ChvNFe", objNFiscal)
    If lErro <> SUCESSO Then gError 42745
    
    ChvNFe.PromptInclude = False
    ChvNFe.Text = objNFiscal.sChvNFe
    ChvNFe.PromptInclude = True
    gsChvAnt = objNFiscal.sChvNFe

    'Coloca o Tipo na tela
    TipoNFiscal.Text = objTipoDocInfo.iCodigo
    Call TipoNFiscal_Validate(bCancel)
    
    'De acordo com a Origem do tipo Coloca o Cliente ou o fornecedor na tela
    If objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
    
        Call Habilita_Cliente
        Cliente.Text = objNFiscal.lCliente
        Call Cliente_Validate(bCancel)
        Filial.Text = objNFiscal.iFilialCli
        
        lErro = CF2(Me, "NFCF_Cliente_Traz_NF_Tela", objNFiscal)
        If lErro <> SUCESSO Then gError 131931
        
    ElseIf objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
    
        Call Habilita_Fornecedor
        Fornecedor.Text = objNFiscal.lFornecedor
        Call Fornecedor_Validate(bCancel)
        Filial.Text = objNFiscal.iFilialForn
        
        lErro = CF2(Me, "NFCF_Fornecedor_Traz_NF_Tela", objNFiscal)
        If lErro <> SUCESSO Then gError 131932
        
    End If
    
    For iIndice = 0 To TipoFrete.ListCount - 1
        If TipoFrete.ItemData(iIndice) = objNFiscal.iFreteRespons Then
            TipoFrete.ListIndex = iIndice
            Exit For
        End If
    Next
    

    'Lê os ítens do Movimento se for Benefiamento e preenche o grid
    If (objNFiscal.iTipoNFiscal = DOCINFO_NFIEDSB) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFEEDSB) Then
        
        lErro = CF("MovEstoqueItensBenef_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 52384 Then gError 55648
        
        lErro = Preenche_GridMovimentos(objNFiscal)
        If lErro <> SUCESSO Then gError 55649
            
    End If

    'Se a nota fiscal original estiver preenchida
    If objNFiscal.lNumIntNotaOriginal > 0 Then

        Set objNFiscalOriginal = New ClassNFiscal
        objNFiscalOriginal.lNumIntDoc = objNFiscal.lNumIntNotaOriginal

        'Lê Nota Fiscal Original
        lErro = CF("NFiscal_Le", objNFiscalOriginal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 35236

        'Se não encontrou a Nota Fiscal Original --> erro
        If lErro <> SUCESSO Then gError 35238

        'Coloca na Tela o Número e a Série
        If objNFiscalOriginal.lNumNotaFiscal > 0 Then
            NFiscalOriginal.Text = objNFiscalOriginal.lNumNotaFiscal
            SerieNFiscalOriginal.Text = objNFiscalOriginal.sSerie
        End If

    End If

    'Preenche os Demais campos da Tela
    lErro = Preenche_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 35241

    ''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 22842
    '*** fim tributacao

    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 39645

    lErro = CF2(Me, "NFEnt_Traz_NFiscal_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 124252

    gbCarregandoTela = False

    NumNFe.Caption = objNFiscal.sStatusNFeFed
    
    MsgAutomatica.Value = vbUnchecked
    
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

    Traz_NFiscal_Tela = SUCESSO

    iAlterado = 0

    Exit Function

Erro_Traz_NFiscal_Tela:

    gbCarregandoTela = False

    Traz_NFiscal_Tela = gErr

    Select Case gErr

        Case 35217, 35226, 35227, 35228, 35233, 35234, 35136, 35237, 124252
        
        Case 35241, 35260, 35261, 75931, 83286, 83287, 89558, 131931, 131932 'Inserido por Wagner

        Case 35235
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case 35238
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA", gErr, objNFiscalOriginal.lNumIntDoc)

        Case 39645

        Case 55175
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NFORIGINAL_NAO_CADASTRADO2", gErr, objItemNFiscalOriginal.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156498)

    End Select

    Exit Function

End Function

Private Sub ValorReal_Calcula(dQuantidade As Double, dValorUnitario As Double, dPercentDesc As Double, dDesconto As Double, dValorReal As Double)
'Calcula o Valor Real

Dim dValorTotal As Double
Dim dPercDesc1 As Double
Dim dPercDesc2 As Double

    dValorTotal = Arredonda_Moeda(dValorUnitario * dQuantidade)

     'Se o Percentual Desconto estiver preenchido e ValorTotal positivo
    If dPercentDesc > 0 And dValorTotal > 0 Then

        'Testa se o desconto está preenchido
        If dDesconto = 0 Then
            dPercDesc2 = 0
        Else
            'Calcula o Percentual em cima dos valores passados
            dPercDesc2 = (dDesconto / dValorTotal)
            dPercDesc2 = CDbl(Format(dPercDesc2, "0.0000"))
        End If
        'se os percentuais (passado e calulado) forem diferentes calcula-se o desconto
        If dPercentDesc <> dPercDesc2 Then dDesconto = Arredonda_Moeda(dPercentDesc * dValorTotal)

    End If

    dValorReal = dValorTotal - dDesconto

End Sub

Private Function SubTotal_Calcula() As Long
'Soma a coluna de Valor Total e acumula em SubTotal

Dim lErro As Long
Dim dSubTotal As Double
Dim dValorProdutos As Double, dValorDescontoItens As Double
Dim iIndice As Integer
Dim dValorDesconto As Double
Dim objItemNF As ClassItemNF

On Error GoTo Erro_SubTotal_Calcula

    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemNF = gobjNFiscal.colItens.Item(iIndice)

        If objItemNF.objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))) <> 0 Then
            
                dSubTotal = dSubTotal + CDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
    
            End If
            dValorProdutos = dValorProdutos + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col))
            dValorDescontoItens = dValorDescontoItens + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Desconto_Col))
        End If
    Next

'    If Len(Trim(ValorDesconto.ClipText)) > 0 And IsNumeric(ValorDesconto) Then dValorDesconto = CDbl(ValorDesconto.Text)
'
'    dSubTotal = Arredonda_Moeda(dSubTotal - dValorDesconto, 2)
    
    SubTotal.Caption = Format(CStr(dSubTotal), "Standard")
    ValorProdutos2.Caption = Format(CStr(dValorProdutos), "Standard")

    dValorDescontoItensAnt = dValorDescontoItens
    If dValorProdutos > 0 Then
        PercDescontoItens.Text = Format(dValorDescontoItens / dValorProdutos * 100, "FIXED")
        dPercDescontoItensAnt = Arredonda_Moeda(dValorDescontoItens / dValorProdutos, 4)
    Else
        PercDescontoItens.Text = Format(0, "FIXED")
        dPercDescontoItensAnt = 0
    End If
    
    SubTotal_Calcula = SUCESSO

    Exit Function

Erro_SubTotal_Calcula:

    SubTotal_Calcula = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156500)

    End Select

    Exit Function

End Function

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

Public Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)
    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)
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

Public Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela
    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)
End Sub

Public Sub objEventoDoc_evSelecao(obj1 As Object)
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
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim objFilial As New ClassFilialFornecedor
Dim sContaTela As String
Dim sContaMascarada As String
Dim objItem As New ClassTribItemNFAux
Dim dQuantidadeConvertida As Double
Dim dQuantidade As Double
Dim iCodigo As Integer
Dim sProduto As String
Dim sUM As String
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim iLinha1 As Integer
Dim iAchou As Integer
Dim sCategoriaProduto As String
Dim objProdutoCategoria As New ClassProdutoCategoria
Dim objCatProdItem As ClassCategoriaProdutoItem
Dim objPlanoConta As New ClassPlanoConta
Dim objFilialCliente As New ClassFilialCliente
Dim objItemNFAux As ClassItemNF, objItemNFOrig As ClassItemNF, dValor As Double

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case QUANT_ESTOQUE
            For iLinha = 1 To objGridItens.iLinhasExistentes
            
                If Len(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
                    
                    lErro = CF("UMEstoque_Conversao", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), CDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)), dQuantidadeConvertida)
                    If lErro <> SUCESSO Then gError 64209

                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case ICMS_CREDITO_FRETE
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribFrete.dICMSCredito
            
        Case ICMS_CREDITO_SEGURO
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribSeguro.dICMSCredito
            
        Case ICMS_CREDITO_OUTRASDESP
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribOutras.dICMSCredito
            
        Case ICMS_CREDITO_DESCONTO
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribDesconto.dICMSCredito
            
        Case IPI_CREDITO_FRETE
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribFrete.dIPICredito
            
        Case IPI_CREDITO_SEGURO
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribSeguro.dIPICredito
            
        Case IPI_CREDITO_OUTRASDESP
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribOutras.dIPICredito
            
        Case IPI_CREDITO_DESCONTO
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.objTribDesconto.dIPICredito

        Case ICMS_CREDITO_VALOR
            
            For Each objItem In gobjTribTab.objNFTributacao.colItens
                
                objMnemonicoValor.colValor.Add objItem.dICMSCredito
                
            Next

        Case IPI_CREDITO_VALOR
            
            For Each objItem In gobjTribTab.objNFTributacao.colItens
                    
                objMnemonicoValor.colValor.Add objItem.dIPICredito
                
            Next
            
        Case IPI_PRODUTO
            
            For Each objItem In gobjTribTab.objNFTributacao.colItens
                    
                objMnemonicoValor.colValor.Add objItem.dIPIValor
                
            Next
            
        Case CLIENTE1
            If Len(Cliente.Text) > 0 Then
                objMnemonicoValor.colValor.Add Cliente.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case DATA_EMISSAO
            If Len(DataEmissao.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataEmissao.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If

        Case DATA_ENTRADA
            If Len(DataEntrada.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataEntrada.FormattedText)
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

        Case FORNECEDOR1
            If Len(Fornecedor.Text) > 0 Then
                objMnemonicoValor.colValor.Add Fornecedor.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case NATUREZA_OP
            If Len(NaturezaOp.Text) > 0 Then
                objMnemonicoValor.colValor.Add NaturezaOp.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case NFISCAL1
            If Len(NFiscal.Text) > 0 Then
                objMnemonicoValor.colValor.Add CLng(NFiscal.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

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

        Case SERIE1
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
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case VALOR_UNITARIO
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))
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
        
            dValor = 0
        
            If Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_NFIEDV Then
                
                For iLinha = 1 To objGridItens.iLinhasExistentes
                'For Each objItemNFAux In gobjNFiscal.colItens
                
                    If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_SerieNFOrig_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iLinha, iGrid_NumNFOrig_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iLinha, iGrid_ItemNFOrig_Col))) > 0 Then
                    'If objItemNFAux.lNumIntDocOrig <> 0 Then
                    
                        Set objItemNFOrig = New ClassItemNF
                        
                        Set objItemNFOrig = New ClassItemNF
            
                        objItemNFOrig.sSerieNFOrig = GridItens.TextMatrix(iLinha, iGrid_SerieNFOrig_Col)
                        objItemNFOrig.lNumNFOrig = GridItens.TextMatrix(iLinha, iGrid_NumNFOrig_Col)
                        objItemNFOrig.iItem = GridItens.TextMatrix(iLinha, iGrid_ItemNFOrig_Col)

                        lErro = CF("ItemNFiscalSaida_Le_NumNFItem", objItemNFOrig)
                        If lErro <> SUCESSO And lErro <> 35368 Then gError ERRO_SEM_MENSAGEM
                
'                        objItemNFOrig.lNumIntDoc = objItemNFAux.lNumIntDocOrig
'
'                        lErro = CF("ItemNFiscal_Le", objItemNFOrig)
'                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                           
                        Set objItemNFOrig.objTributacao = New ClassTributacaoDocItem
                        objItemNFOrig.objTributacao.iTipoDoc = TIPODOC_TRIB_NF
                        objItemNFOrig.objTributacao.lNumIntDocItem = objItemNFOrig.lNumIntDoc
                        lErro = CF("TributacaoDocItem_Le", objItemNFOrig.objTributacao)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                          
                        dValor = dValor + (objItemNFOrig.objTributacao.dICMSSubstValor / objItemNFOrig.dQuantidade) * StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
                    
                    End If
                    
                Next
            End If
            If Len(ICMSSubstValor1.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ICMSSubstValor1.Caption) + dValor
            Else
                objMnemonicoValor.colValor.Add 0 + dValor
            End If

        Case SUBTOTAL1
            If Len(SubTotal.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(SubTotal.Caption)
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
        
            dValor = 0
            If Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_NFIEDV Then
            
                For iLinha = 1 To objGridItens.iLinhasExistentes
                'For Each objItemNFAux In gobjNFiscal.colItens
                
                    If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_SerieNFOrig_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iLinha, iGrid_NumNFOrig_Col))) > 0 And Len(Trim(GridItens.TextMatrix(iLinha, iGrid_ItemNFOrig_Col))) > 0 Then
                    'If objItemNFAux.lNumIntDocOrig <> 0 Then
                    
                        Set objItemNFOrig = New ClassItemNF
            
                        objItemNFOrig.sSerieNFOrig = GridItens.TextMatrix(iLinha, iGrid_SerieNFOrig_Col)
                        objItemNFOrig.lNumNFOrig = GridItens.TextMatrix(iLinha, iGrid_NumNFOrig_Col)
                        objItemNFOrig.iItem = GridItens.TextMatrix(iLinha, iGrid_ItemNFOrig_Col)

                        lErro = CF("ItemNFiscalSaida_Le_NumNFItem", objItemNFOrig)
                        If lErro <> SUCESSO And lErro <> 35368 Then gError ERRO_SEM_MENSAGEM
                
'                        objItemNFOrig.lNumIntDoc = objItemNFAux.lNumIntDocOrig
'
'                        lErro = CF("ItemNFiscal_Le", objItemNFOrig)
'                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                           
                        Set objItemNFOrig.objTributacao = New ClassTributacaoDocItem
                        objItemNFOrig.objTributacao.iTipoDoc = TIPODOC_TRIB_NF
                        objItemNFOrig.objTributacao.lNumIntDocItem = objItemNFOrig.lNumIntDoc
                        lErro = CF("TributacaoDocItem_Le", objItemNFOrig.objTributacao)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                        dValor = dValor + (objItemNFOrig.objTributacao.dICMSSubstValor / objItemNFOrig.dQuantidade) * StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
                    
                    End If
                    
                Next
            End If
            If Len(ValorDespesas.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorDespesas.Text) - dValor
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

        Case TOTAL1
            If Len(Total.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(Total.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case ALMOXARIFADO1
            
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                
                If Len(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_AlmoxDist_Col)) > 0 Then
                    
                    objMnemonicoValor.colValor.Add GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_AlmoxDist_Col)
                
                Else
                    
                    objMnemonicoValor.colValor.Add ""
                
                End If
            
            Next

        
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
            
        
        Case PRODUTO_ITENS
            
            For iLinha = 1 To objGridItens.iLinhasExistentes
                
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                    
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_Produto_Col)
                
                Else
                    
                    objMnemonicoValor.colValor.Add ""
                
                End If
            
            Next
        
        Case PRODUTO_DIST
            
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                
                If Len(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case QUANT_DIST
                
                For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                    
                    If Len(Trim(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col))) > 0 Then
                        If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
                            objMnemonicoValor.colValor.Add CDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col))
                        Else
                            objMnemonicoValor.colValor.Add 0
                        End If
                    Else
                        objMnemonicoValor.colValor.Add 0
                    End If
                Next
        
        'Conta Contabil de Estoque do par Produto x ALmoxarifado
        Case CTACONTABILEST1
                        
            'Para cada linha do grid distribuição
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                
                'Se iLinha do Grid Distribuição tem a quantidade preenchida
                If Len(Trim(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col))) > 0 Then
                    
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

        Case CTAFILIALFORNECEDOR1
            
            If Len(Trim(Filial.Text)) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 57880
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 57881
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case ESCANINHO_CUSTO
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                    iCodigo = Codigo_Extrai(TipoNFiscal.Text)
                    
                    Select Case iCodigo
                    
                        Case DOCINFO_NFEEDC, DOCINFO_NFIEDC
                            objMnemonicoValor.colValor.Add ESCANINHO_NOSSO_EM_CONSIGNACAO
                            
                        Case DOCINFO_NFEEDCNT, DOCINFO_NFIEDCNT
                            objMnemonicoValor.colValor.Add ESCANINHO_NOSSO_EM_CONSERTO
                        
                        Case DOCINFO_NFEEDD, DOCINFO_NFIEDD
                            objMnemonicoValor.colValor.Add ESCANINHO_NOSSO_EM_DEMO
                        
                        Case DOCINFO_NFEEODC, DOCINFO_NFEEODF, DOCINFO_NFIEODC, DOCINFO_NFIEODF
                            objMnemonicoValor.colValor.Add ESCANINHO_NOSSO_EM_OUTROS
                       
                        Case Else
                            objMnemonicoValor.colValor.Add ESCANINHO_NOSSO
                            
                    End Select
                
                Else
                
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case ICMS_CREDITO_TOTAL
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dICMSCredito
        
        Case IPI_CREDITO_TOTAL
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dIPICredito
        
        Case PIS_CREDITO_TOTAL
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dPISCredito
        
        Case COFINS_CREDITO_TOTAL
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dCOFINSCredito
        
        Case ISS_VALOR
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.dISSValor

        Case ISS_INCLUSO
            objMnemonicoValor.colValor.Add gobjTribTab.objNFTributacao.iISSIncluso <> 0

        Case FATOR_ITEM_EST
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                If Len(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_UMDist_Col)) > 0 Then
                    If Len(GridItens.TextMatrix(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ItemNFDist_Col), iGrid_UnidadeMed_Col)) > 0 Then
                
                        lErro = CF("Produto_Formata", GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iPreenchido)
                        If lErro <> SUCESSO Then gError 183082
                        
                        'Lê o produto
                        objProduto.sCodigo = sProdutoFormatado
                        lErro = CF("Produto_Le", objProduto)
                        If lErro <> SUCESSO And lErro <> 28030 Then gError 183083
            
                        'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, GridItens.TextMatrix(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ItemNFDist_Col), iGrid_UnidadeMed_Col), GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_UMDist_Col), dFator)
                        If lErro <> SUCESSO Then gError 183084
                
                        objMnemonicoValor.colValor.Add dFator
                    Else
                        objMnemonicoValor.colValor.Add 1
                    End If
                        
                Else
                    objMnemonicoValor.colValor.Add 1
                End If
            Next
        
        Case PRODUTO_TIPO_EST
        
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                If Len(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183085
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183086
        
                    objMnemonicoValor.colValor.Add objProduto.iTipo
                    
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        Case PRODUTO_TIPO
        
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183152
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183153
        
                    objMnemonicoValor.colValor.Add objProduto.iTipo
                    
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        
        'Conta Contabil de Estoque do par Produto x ALmoxarifado
        Case CTACONTABILDIST
                        
            'Para cada linha do grid distribuição
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                
                'Se iLinha do Grid Distribuição tem a quantidade preenchida
                If Len(Trim(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col))) > 0 Then
                    
                    objAlmoxarifado.sNomeReduzido = GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_AlmoxDist_Col)

                    'Lê o Nome Reduzido do Almoxarifado
                    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25060 Then gError 183173

                    'Se não encontrou ===> Erro
                    If lErro = SUCESSO Then

                        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
                        
                        lErro = CF("Produto_Formata", GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iPreenchido)
                        If iPreenchido <> PRODUTO_PREENCHIDO Then gError 183174

                        objEstoqueProduto.sProduto = sProdutoFormatado
                        
                        'Lê a conta contábil do Produto no Almoxarifado
                        lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
                        If lErro <> SUCESSO And lErro <> 49991 Then gError 183175

                        If Len(Trim(objEstoqueProduto.sContaContabil)) > 0 Then

                            sContaMascarada = String(STRING_CONTA, 0)

                            lErro = Mascara_MascararConta(objEstoqueProduto.sContaContabil, sContaMascarada)
                            If lErro <> SUCESSO Then gError 183176

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
        
        Case VALOR_UNIT_DIST
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                
                For iLinha1 = 1 To objGridItens.iLinhasExistentes
                
                    If GridItens.TextMatrix(iLinha1, iGrid_Produto_Col) = GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col) Then
                    
                        If Len(GridItens.TextMatrix(iLinha1, iGrid_ValorUnitario_Col)) > 0 Then
                            objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha1, iGrid_ValorUnitario_Col))
                        Else
                            objMnemonicoValor.colValor.Add 0
                        End If
                    
                        Exit For
                    
                    End If
                
                Next
                
            Next
        
        Case ICMS_CRED_VALOR_DIST
            
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
            
                lErro = CF("Produto_Formata", GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iPreenchido)
                If iPreenchido <> PRODUTO_PREENCHIDO Then gError 183181

                iAchou = 0
                
                For Each objItem In gobjTribTab.objNFTributacao.colItens
                
                    If objItem.sProduto = sProdutoFormatado Then
                
                        If StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col)) > 0 And _
                           StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantItemNFDist_Col)) > 0 Then
                
                            objMnemonicoValor.colValor.Add objItem.dICMSCredito * StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col)) / StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantItemNFDist_Col))
                            iAchou = 1
                            Exit For
                        
                        End If
                    
                    End If
                    
                Next
                
                If iAchou = 0 Then
                    objMnemonicoValor.colValor.Add 0
                End If
                
            Next
        
        Case IPI_CRED_VALOR_DIST
            
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
            
                lErro = CF("Produto_Formata", GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iPreenchido)
                If iPreenchido <> PRODUTO_PREENCHIDO Then gError 183182

                iAchou = 0
                
                For Each objItem In gobjTribTab.objNFTributacao.colItens
                
                    If StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col)) > 0 And _
                       StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantItemNFDist_Col)) > 0 Then
            
                        objMnemonicoValor.colValor.Add objItem.dIPICredito * StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col)) / StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantItemNFDist_Col))
                        iAchou = 1
                        Exit For
                    
                    End If
                    
                Next
                    
                If iAchou = 0 Then
                    objMnemonicoValor.colValor.Add 0
                End If
                
            Next
                    

        Case IPI_PRODUTO_DIST
            
            
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
            
                lErro = CF("Produto_Formata", GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_ProdutoAlmoxDist_Col), sProdutoFormatado, iPreenchido)
                If iPreenchido <> PRODUTO_PREENCHIDO Then gError 183183

                iAchou = 0
                
                For Each objItem In gobjTribTab.objNFTributacao.colItens
                
                    If StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col)) > 0 And _
                       StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantItemNFDist_Col)) > 0 Then
            
                        objMnemonicoValor.colValor.Add objItem.dIPIValor * StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col)) / StrParaDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantItemNFDist_Col))
                        iAchou = 1
                        Exit For
                    
                    End If
                    
                Next
                    
                If iAchou = 0 Then
                    objMnemonicoValor.colValor.Add 0
                End If
                
            Next
            
        Case QUANT_DIST1
                
            For iLinha = 1 To gobjDistribuicao.objGridDist.iLinhasExistentes
                
                If Len(Trim(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col))) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridDist.TextMatrix(iLinha, gobjDistribuicao.iGrid_QuantDist_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        Case CONTROLA_EST
        
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183420
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183421
        
                    objMnemonicoValor.colValor.Add objProduto.iControleEstoque
                    
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        Case CONTA_DESP_EST_FORN
            
            If Len(Trim(Filial.Text)) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 183422
                
                If objFilial.sContaDespesa <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaDespesa, sContaTela)
                    If lErro <> SUCESSO Then gError 183423
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
        
        Case CONTACONTABILAPLIC
        
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183528
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183529
        
                    If objProduto.sContaContabil <> "" Then
                    
                        lErro = Mascara_RetornaContaTela(objProduto.sContaContabil, sContaTela)
                        If lErro <> SUCESSO Then gError 183806
                    
                    Else
                    
                        sContaTela = ""
                        
                    End If
        
                    objMnemonicoValor.colValor.Add sContaTela
                    
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case CONTACONTABILPROD
        
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183528
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183529
        
                    If objProduto.sContaContabilProducao <> "" Then
                    
                        lErro = Mascara_RetornaContaTela(objProduto.sContaContabilProducao, sContaTela)
                        If lErro <> SUCESSO Then gError 183806
                    
                    Else
                    
                        sContaTela = ""
                        
                    End If
        
                    objMnemonicoValor.colValor.Add sContaTela
                    
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case CONTACONTABILCATPROD
                
            For iLinha = 1 To objGridItens.iLinhasExistentes
                
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183546
                
                    objProdutoCategoria.sCategoria = objMnemonicoValor.vParam(1)
                    
                    objProdutoCategoria.sProduto = sProdutoFormatado
                    
                    Set objCatProdItem = New ClassCategoriaProdutoItem
                    
                    lErro = CF("ProdutoCategoria_Le_Item", objProdutoCategoria, objCatProdItem)
                    If lErro <> SUCESSO And lErro <> 183548 Then gError 183547
                    
                    If lErro = SUCESSO Then
                        
                        Set objPlanoConta = New ClassPlanoConta
                    
                        'le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta
                        lErro = CF("PlanoConta_Le_ContaSimples", objCatProdItem.dvalor1, objPlanoConta)
                        If lErro <> SUCESSO And lErro <> 5451 Then gError 183548
                    
                        If lErro = SUCESSO Then
                    
                            If objPlanoConta.sConta <> "" Then
                            
                                lErro = Mascara_RetornaContaTela(objPlanoConta.sConta, sContaTela)
                                If lErro <> SUCESSO Then gError 183807
                            
                            Else
                            
                                sContaTela = ""
                                
                            End If
                    
                            objMnemonicoValor.colValor.Add sContaTela
                            
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
                
                objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilialCliente)
                If lErro <> SUCESSO Then gError 210809
                
                If objFilialCliente.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilialCliente.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 210810
                
                Else
                    sContaTela = ""
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        
        
        Case Else
            gError 39648

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39648
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case 57880, 57881, 57966, 57987, 57988, 57989, 64209, 75453, 75455, 75456, 75458, 84902, 84903, 183067 To 183071, 183141, 183142, 183419 To 183423, 183528, 183529, 183546, 183547, 183548, 183806, 183807, 210809, 210810
        
        Case 75454
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO1", gErr, objAlmoxarifado.sNomeReduzido)
            
        Case 75457
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156501)

    End Select

    Exit Function

End Function

Public Function ValorTotal_Calcula() As Long

Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    lErro = gobjTribTab.AtualizarTributacao()
    If lErro <> SUCESSO Then gError 42068
    
    Call ValorDescontoItens_Calcula

    If iEmitenteNF = DOCINFO_EMPRESA Then
    
        'atualizar o total da nota
        Call Total_Calcula(0)
        
    End If
        
    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 42068

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156502)

    End Select

    Exit Function

End Function

Public Sub TipoNFiscal_Correspondente_Recebimento(objTipoDocInfo As ClassTipoDocInfo, iTipoRecebimento As Integer)

    If objTipoDocInfo.iTipo = DOCINFO_NF_INT_ENTRADA And objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
        iTipoRecebimento = DOCINFO_NRCP
    ElseIf objTipoDocInfo.iTipo = DOCINFO_NF_INT_ENTRADA And objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        iTipoRecebimento = DOCINFO_NRFP
    ElseIf objTipoDocInfo.iTipo = DOCINFO_NF_EXTERNA And objTipoDocInfo.iOrigem = DOCINFO_CLIENTE Then
        iTipoRecebimento = DOCINFO_NRCC
    ElseIf objTipoDocInfo.iTipo = DOCINFO_NF_EXTERNA And objTipoDocInfo.iOrigem = DOCINFO_FORNECEDOR Then
        iTipoRecebimento = DOCINFO_NRFF
    End If

End Sub

Public Function Calcula_Valores(iLinha As Integer) As Long
'recalcula os valores de desconto, percentual de desconto e valor total

Dim sProduto As String
Dim lErro As Long
Dim lTamanho As Long
Dim dPercentDesc As Double
Dim dValorUnitario As Double
Dim dDesconto As Double
Dim dValorReal As Double
Dim dQuantidade As Double

On Error GoTo Erro_Calcula_Valores

    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))

    lErro = CF2(Me, "NFEDEV_Calcula_Valores", iLinha, dQuantidade)
    If lErro <> SUCESSO Then gError 177058

    'Recolhe os valores Quantidade, Desconto, PerDesc e Valor Unitário da tela
    If dQuantidade = 0 Or Len(Trim(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))) = 0 Then

        GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = ""
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(0, "Standard")
        
    Else

        lTamanho = Len(Trim(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col)))

        If lTamanho > 0 Then
            dPercentDesc = PercentParaDbl(GridItens.TextMatrix(iLinha, iGrid_PercDesc_Col))
        Else
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        End If

        If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))) > 0 Then dValorUnitario = CDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))
        If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col))) > 0 Then dDesconto = CDbl(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col))

        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dValorUnitario, dPercentDesc, dDesconto, dValorReal)

        'Coloca o Desconto calculado na tela
        If dDesconto > 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
        Else
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        End If

        'Coloca o valor Real em Valor Total
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(dValorReal, "Standard")

        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dValorUnitario, 0, 0, dValorReal)

        'Coloca o valor Real em Valor Total
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(dValorReal, "Standard")
    End If

    lErro = SubTotal_Calcula()
    If lErro <> SUCESSO Then gError 55572

    Calcula_Valores = SUCESSO
    
    Exit Function
    
Erro_Calcula_Valores:

    Calcula_Valores = gErr
    
    Select Case gErr

        Case 55572, 177058

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156503)

    End Select

    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ENTRADA_NF_DEVOLUCAO_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Entrada - Nota Fiscal Devolução"
    Call Form_Load
    
End Function

Public Function Name() As String
    Name = "NFiscalEntDev"
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Public Sub Unload(objme As Object)
   RaiseEvent Unload
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property
'**** fim do trecho a ser copiado *****

Private Function Trata_FilialCliente() As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Trata_FilialCliente
    
    If Filial.ListIndex <> -1 Then
    
        objFilialCliente.iCodFilial = Filial.ItemData(Filial.ListIndex)
        'Lê a FilialCliente
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilialCliente)
        If lErro <> SUCESSO Then gError 56977
        
        Call gobjTribTab.FilialCliente_Alterada(objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    
        lErro = CF2(Me, "NFCF_Cliente_Filial_Validate", objFilialCliente)
        If lErro <> SUCESSO Then gError 131933
    
    End If
        
    Trata_FilialCliente = SUCESSO
    
    Exit Function
    
Erro_Trata_FilialCliente:

    Trata_FilialCliente = gErr
    
    Select Case gErr
    
        Case 56977, 131933 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156504)
            
    End Select
    
    Exit Function
    
End Function

Private Function Trata_FilialForn() As Long

Dim lErro As Long
Dim objFilialForn As New ClassFilialFornecedor

On Error GoTo Erro_Trata_FilialForn
    
    objFilialForn.iCodFilial = Codigo_Extrai(Filial.Text)
    
    If objFilialForn.iCodFilial <> 0 Then
    
        'Lê a Filial
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilialForn)
        If lErro <> SUCESSO Then gError 56978
        
        Call gobjTribTab.FilialFornecedor_Alterada(objFilialForn.lCodFornecedor, objFilialForn.iCodFilial)
    
        lErro = CF2(Me, "NFCF_Fornecedor_Filial_Validate", objFilialForn)
        If lErro <> SUCESSO Then gError 131934
    
    End If
    
    Trata_FilialForn = SUCESSO
     
    Exit Function
    
Erro_Trata_FilialForn:

    Trata_FilialForn = gErr
     
    Select Case gErr
          
        Case 56978, 131934 'Inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156505)
     
    End Select
     
    Exit Function

End Function

Public Sub BotaoConsultaDocCPR_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objFornecedor As New ClassFornecedor
Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoConsultaDocCPR_Click

    'Verifica se todos os campos necessários para se efetuar a consulta foram preenchidos
    lErro = Critica_CamposNecessarios_ConsultaTitulo()
    If lErro <> SUCESSO Then gError 79783
        
    'Guarda no objNFiscal os dados necessários para consultar o título
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.iTipoDocInfo = Codigo_Extrai(TipoNFiscal.Text)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.dtDataEntrada = StrParaDate(DataEntrada.Text)
    
    If Fornecedor.Visible = True Then
    
        'Guarda em objFornecedor o nome reduzido do Fornecedor
        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        'Lê o código do Fornecedor a partir do nome reduzido obtido na tela
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 79784
        
        'Se não encontrou o fornecedor = > erro
        If lErro = 6681 Then gError 79785
        
        'Guarda no objNFiscal o código do fornecedor
        objNFiscal.lFornecedor = objFornecedor.lCodigo
        objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)
    
    ElseIf Cliente.Visible = True Then
    
        'Guarda em objCliente o nome reduzido do Cliente
        objCliente.sNomeReduzido = Cliente.Text
        
        'Lê o código do Fornecedor a partir do nome reduzido obtido na tela
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 79780
        
        'Se não encontrou o cliente = > erro
        If lErro = 12348 Then gError 79781
        
        'Guarda no objNFiscal o código do fornecedor
        objNFiscal.lCliente = objCliente.lCodigo
        objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
    
    End If
    
    'Lê o NumIntDocCPR da NFiscal e exibe o documento gerado no CPR por essa nota
    lErro = CF("NFiscal_Consulta_DocCPR", objNFiscal)
    If lErro <> SUCESSO And lErro <> 79717 Then gError 79786
    
    'Se não encontrou a nota
    If lErro = 79717 Then
    
        'Testa se a origem da nota é um fornecedor e exibe mensagem de erro
        If Fornecedor.Visible = True Then gError 79787
                
        'Testa se a origem da nota é um cliente e exibe mensagem de erro
        If Cliente.Visible = True Then gError 79837
    
    End If
    
    Exit Sub
    
Erro_BotaoConsultaDocCPR_Click:

    Select Case gErr
    
        Case 79783, 79786, 79784, 79780
        
        Case 79787
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA1", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada)
        
        Case 79837
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA3", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iTipoNFiscal, objNFiscal.lCliente, objNFiscal.iFilialCli, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada)
        
        Case 79785
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)
        
        Case 79781
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objNFiscal.lCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156506)
    
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
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 79788
    
    'Se o campo Cliente está visível
    ElseIf Cliente.Visible = True Then
    
        'Verifica se o cliente foi preenchido
        If Len(Trim(Cliente.Text)) = 0 Then gError 79782
    
    End If
    
    'Verifica se a filial do Fornecedor / Cliente foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 79789
    
    'Verifica se a Série foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 79790
    
    'Verifica se o Número da Nota foi preenchido
    If Len(Trim(NFiscal.Text)) = 0 Then gError 79791
    
    'Verifica se a data de emissão da nota foi preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 79792
    
    'Verifica se a data de entrada da nota foi preenchida
    If Len(Trim(DataEntrada.ClipText)) = 0 Then gError 79793
    
    'Verifica se o tipo da Nota fiscal foi selecionado
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 79794
    
    Critica_CamposNecessarios_ConsultaTitulo = SUCESSO
    
    Exit Function

Erro_Critica_CamposNecessarios_ConsultaTitulo:

    Critica_CamposNecessarios_ConsultaTitulo = gErr
    
    Select Case gErr
    
        Case 79788
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 79782
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 79789
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 79790
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 79791
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
        
        Case 79792
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 79793
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_NAO_PREENCHIDA", gErr)
        
        Case 79794
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156507)
    
    End Select
        
    Exit Function
        
End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is SerieNFiscalOriginal Then
            Call SerieOriginalLabel_Click
        ElseIf Me.ActiveControl Is NFiscalOriginal Then
            Call NFOriginalLabel_Click
        ElseIf Me.ActiveControl Is NaturezaOp Then
            Call NaturezaLabel_Click
        ElseIf Me.ActiveControl Is Serie Then
            Call SerieLabel_Click
        ElseIf Me.ActiveControl Is NFiscal Then
            Call NFiscalLabel_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
'distribuicao
        ElseIf Me.ActiveControl Is gobjDistribuicao.AlmoxDist Then
            Call gobjDistribuicao.BotaoLocalizacaoDist_Click
        ElseIf Me.ActiveControl Is ItemNFOriginal Then
            Call BotaoItensNFiscalOrig_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        ElseIf Me.ActiveControl Is ProdutoBenef Then
            Call BotaoProdutosBenef_Click
        ElseIf Me.ActiveControl Is AlmoxarifadoBenef Then
            Call BotaoEstoqueBenef_Click
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

Public Function Form_Load_Light() As Long

Dim lErro As Long

On Error GoTo Erro_Form_Load_Light

    'obs.: a carga pos form_load será feita pela classe que trata nffat
    giPosCargaOk = 1
    
    iFrameAtual = 1
    gbGravando = False
    
    'Inicializa as variáveis relacionadas c\ as telas de Browse
    Set objEventoSerie = New AdmEvento
    Set objEventoNFiscal = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoRecebimento = New AdmEvento
    Set objEventoLocalizacao = New AdmEvento
    Set objEventoSerieOriginal = New AdmEvento
    Set objEventoTransportadora = New AdmEvento
    Set objEventoNFiscalOriginal = New AdmEvento
    Set objEventoItensNFiscalOriginal = New AdmEvento
    Set objEventoCcl = New AdmEvento

    'Inicializa FrameL para versão LIGHT
'    If giTipoVersao = VERSAO_LIGHT Then
'        Set FrameL(1) = Frame1(1)
'        Set FrameL(2) = Frame1(2)
'        Set FrameL(3) = Frame1(3)
'        Set FrameL(4) = Frame1(5)
'        Set FrameL(5) = Frame1(6)
'    End If
    
    Call GridItens_InicializaColunas
    
    iAlterado = 0

    Form_Load_Light = SUCESSO

    Exit Function

Erro_Form_Load_Light:

    Form_Load_Light = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156508)

    End Select

    iAlterado = 0
    
    Exit Function

End Function

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Sub Observacao_Change()
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
        
            Call Atualiza_Grid_Itens(objItemNF)
            
            Call gobjDistribuicao.Atualiza_Grid_Distribuicao(objItemNF)
        
            Call Calcula_Valores(GridItens.Row)
        
        End If
    
    End If
    
    Exit Sub

Erro_BotaoGrade_Click:

    Select Case gErr
    
        Case 51037, 86360
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156509)
            
    End Select
    
    Exit Sub

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
            
'            If objAlmoxarifado.iFilialEmpresa = giFilialEmpresa Then
            
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

Private Function Saida_Celula_GridMovimentos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridMovimentos
        
    Select Case GridMovimentos.Col

        Case iGrid_ItemNFBenef_Col

            lErro = Saida_Celula_ItemNFBenef(objGridInt)
            If lErro <> SUCESSO Then gError 126858

        Case iGrid_ProdutoBenef_Col

            lErro = Saida_Celula_ProdutoBenef(objGridInt)
            If lErro <> SUCESSO Then gError 126859

        Case iGrid_AlmoxarifadoBenef_Col

            lErro = Saida_Celula_AlmoxarifadoBenef(objGridInt)
            If lErro <> SUCESSO Then gError 126898

        Case iGrid_QuantInsumoBenef_Col

            lErro = Saida_Celula_QuantInsumoBenef(objGridInt)
            If lErro <> SUCESSO Then gError 126860

        Case iGrid_ContaContabilEstBenef_Col
            lErro = Saida_Celula_ContaContabilEstBenef(objGridInt)
            If lErro <> SUCESSO Then gError 126861

        Case iGrid_ContaContabilProducaoBenef_Col
            lErro = Saida_Celula_ContaContabilProducaoBenef(objGridInt)
            If lErro <> SUCESSO Then gError 126862

    End Select
    
    Saida_Celula_GridMovimentos = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_GridMovimentos:

    Saida_Celula_GridMovimentos = gErr

    Select Case gErr

        Case 126858 To 126862

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156510)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ItemNFBenef(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iItem As Integer

On Error GoTo Erro_Saida_Celula_ItemNFBenef

    Set objGridInt.objControle = ItemNFBenef

    'Verifica se o Item esta preenchido
    If Len(Trim(ItemNFBenef.Text)) = 0 And Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFBenef_Col))) > 0 Then gError 126856
    
    If Len(Trim(ItemNFBenef.Text)) > 0 Then

        iItem = CInt(ItemNFBenef.Text)
        
        If iItem < 1 Or iItem > objGridItens.iLinhasExistentes Then gError 126857
        
        'Se necessário cria uma nova linha no Grid
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 126863

    Saida_Celula_ItemNFBenef = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemNFBenef:

    Saida_Celula_ItemNFBenef = gErr

    Select Case gErr
        
        Case 126856, 126857
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMNFBENEF_NAO_ITEMNF", gErr, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 126863
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156511)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ProdutoBenef(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_ProdutoBenef

    Set objGridInt.objControle = ProdutoBenef

    'Verifica se o Produto esta preenchida
    If Len(Trim(ProdutoBenef.ClipText)) > 0 Then

        lErro = CF("Produto_Critica", ProdutoBenef.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 126864
        
        If lErro = 25041 Then gError 126865
        
        If objProduto.iPCP = PRODUTO_PCP_NAOPODE Then gError 126886

        lErro = ProdutoLinha_PreencheBenef(objProduto)
        If lErro <> SUCESSO Then gError 126887

        lErro = Preenche_ContaContabilEst(ProdutoBenef.Text)
        If lErro <> SUCESSO Then gError 126918

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 126867

    Saida_Celula_ProdutoBenef = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdutoBenef:

    Saida_Celula_ProdutoBenef = gErr

    Select Case gErr

        Case 126864, 126867, 126887, 126888, 126918
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 126865
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", ProdutoBenef.Text)

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 126886
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PCP", gErr, ProdutoBenef.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156512)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_AlmoxarifadoBenef(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_AlmoxarifadoBenef

    Set objGridInt.objControle = AlmoxarifadoBenef

    If Len(Trim(AlmoxarifadoBenef.ClipText)) > 0 Then

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 126899

        lErro = TP_Almoxarifado_Produto_Grid(sProdutoFormatado, AlmoxarifadoBenef, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25130 And lErro <> 25135 Then gError 126900

        If lErro = 25130 Then gError 126901

        If lErro = 25135 Then gError 126902

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col) = objAlmoxarifado.sNomeReduzido
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 126904

    'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
    lErro = Preenche_ContaContabilEst(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col))
    If lErro <> SUCESSO Then gError 126905

    Saida_Celula_AlmoxarifadoBenef = SUCESSO

    Exit Function

Erro_Saida_Celula_AlmoxarifadoBenef:

    Saida_Celula_AlmoxarifadoBenef = gErr

    Select Case gErr

        Case 126899, 126900, 126903 To 126905
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 126901

            vbMsg = Rotina_Aviso(vbYesNo, "ERRO_ALMOXARIFADO_INEXISTENTE", AlmoxarifadoBenef.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = AlmoxarifadoBenef.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 126902

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(AlmoxarifadoBenef.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(AlmoxarifadoBenef.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156513)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantInsumoBenef(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim dQuantAtual As Double
Dim iLinha As Integer
Dim dQuantItem As Double

On Error GoTo Erro_Saida_Celula_QuantInsumoBenef

    Set objGridInt.objControle = QuantInsumoBenef

    'Se quantidade estiver preenchida
    If Len(Trim(QuantInsumoBenef.ClipText)) > 0 Then
    
        'Critica o valor
        lErro = Valor_Positivo_Critica(QuantInsumoBenef.Text)
        If lErro <> SUCESSO Then gError 126867

        dQuantidade = CDbl(QuantInsumoBenef.Text)

        'Coloca o valor Formatado na tela
        QuantInsumoBenef.Text = Formata_Estoque(dQuantidade)

        dQuantItem = StrParaDbl(objGridItens.objGrid.TextMatrix(StrParaInt(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFBenef_Col)), iGrid_Quantidade_Col))

        For iLinha = 1 To objGridInt.iLinhasExistentes
            If iLinha <> objGridInt.objGrid.Row And objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFBenef_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFBenef_Col) Then
                dQuantAtual = dQuantAtual + StrParaDbl(objGridInt.objGrid.TextMatrix(iLinha, iGrid_QuantInsumoBenef_Col))
            End If
        Next

        If dQuantItem < dQuantidade + dQuantAtual Then gError 126868

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 126869

    Saida_Celula_QuantInsumoBenef = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantInsumoBenef:

    Saida_Celula_QuantInsumoBenef = gErr

    Select Case gErr

        Case 126867, 126869
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 126868
            Call Rotina_Erro(vbOKOnly, "ERRO_GRIDITENS_DIFERE_GRIDMOVS", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156514)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaContabilEstBenef(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabilEstBenef
    
    Set objGridInt.objControle = ContaContabilEstBenef
    
    If Len(Trim(ContaContabilEstBenef.ClipText)) > 0 Then
    
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabilEstBenef.ClipText, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 126870
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 126871
            
            ContaContabilEstBenef.PromptInclude = False
            ContaContabilEstBenef.Text = sContaMascarada
            ContaContabilEstBenef.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilEstBenef.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 126872
    
            'conta não cadastrada
            If lErro = 5700 Then gError 126873
             
        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 126874
    
    Saida_Celula_ContaContabilEstBenef = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilEstBenef:

    Saida_Celula_ContaContabilEstBenef = gErr

    Select Case gErr

        Case 126870, 126872, 126874
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 126871
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 126873
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilEstBenef.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156515)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ContaContabilProducaoBenef(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabilProducaoBenef

    Set objGridInt.objControle = ContaContabilProducaoBenef

    If Len(Trim(ContaContabilProducaoBenef.ClipText)) > 0 Then
    
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabilProducaoBenef.ClipText, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 126875
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 126876
            
            ContaContabilProducaoBenef.PromptInclude = False
            ContaContabilProducaoBenef.Text = sContaMascarada
            ContaContabilProducaoBenef.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilProducaoBenef.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 126877
    
            'conta não cadastrada
            If lErro = 5700 Then gError 126878
             
        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 126879
    
   Saida_Celula_ContaContabilProducaoBenef = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilProducaoBenef:

    Saida_Celula_ContaContabilProducaoBenef = gErr

    Select Case gErr

        Case 126875, 126877, 126879
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 126976
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 126878
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilProducaoBenef.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156516)

    End Select

    Exit Function

End Function

Public Sub BotaoPlanoConta_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPlanoConta_Click

    If GridMovimentos.Row = 0 Then gError 126880

    sConta = String(STRING_CONTA, 0)

    If GridMovimentos.Col = iGrid_ContaContabilEstBenef_Col Then

        lErro = CF("Conta_Formata", ContaContabilEstBenef.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 126881

    ElseIf GridMovimentos.Col = iGrid_ContaContabilProducaoBenef_Col Then

        lErro = CF("Conta_Formata", ContaContabilProducaoBenef.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 126882

    End If

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaESTLista
    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 126880
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 126881, 126882

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156517)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objGridBenef.objControle.Name <> "ContaContabilEstBenef" And objGridBenef.objControle.Name <> "ContaContabilProducaoBenef" Then
        Me.Show
        Exit Sub
    End If

    If objPlanoConta.sConta <> "" Then

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 126883

        objGridBenef.objControle.PromptInclude = False
        objGridBenef.objControle.Text = sContaEnxuta
        objGridBenef.objControle.PromptInclude = True

        GridMovimentos.TextMatrix(GridMovimentos.Row, GridMovimentos.Col) = objGridBenef.objControle.Text

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 126883
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156518)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProdutosBenef_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim objItemOP As New ClassItemOP

On Error GoTo Erro_BotaoProdutosBenef_Click

    'se não selecionou nenhuma linha do grid entao - - - -> Erro
    If GridMovimentos.Row = 0 Then gError 126884

    ' Formata o Produto
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 126885

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objProduto.sCodigo = sProdutoFormatado

    'Lista de produtos que  podem participar da producao e inventariados
    Call Chama_Tela("ProdutoEstoquePCPLista", colSelecao, objProduto, objEventoProdutoBenef)

   Exit Sub

Erro_BotaoProdutosBenef_Click:

    Select Case gErr

        Case 126884
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 126885

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156519)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoBenef_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoMascarado As String
Dim objTipoDeProduto As New ClassTipoDeProduto
Dim sProdutoEnxuto As String

On Error GoTo Erro_objEventoProdutoBenef_evSelecao

    Set objProduto = obj1

    If (GridMovimentos.Row > 0) Then

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 126890

        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 126891

        If lErro = 28030 Then gError 126892

        ProdutoBenef.PromptInclude = False
        ProdutoBenef.Text = sProdutoEnxuto
        ProdutoBenef.PromptInclude = True

        If Not (Me.ActiveControl Is ProdutoBenef) Then

            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col) = ProdutoBenef.Text
    
            If Trim(objProduto.sContaContabilProducao) = "" Then
    
                objTipoDeProduto.iTipo = objProduto.iTipo
    
                lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
                If lErro <> SUCESSO And lErro <> 22531 Then gError 126893
    
                If lErro = 22531 Then gError 126894
    
                objProduto.sContaContabilProducao = objTipoDeProduto.sContaContabil
    
            End If
    
            'Preenche a Linha do Grid
            lErro = ProdutoLinha_PreencheBenef(objProduto)
            If lErro <> SUCESSO Then gError 126895
    
            lErro = Preenche_ContaContabilEst(ProdutoBenef.Text)
            If lErro <> SUCESSO Then gError 126919
    
        End If
            
    End If

    Me.Show

    Exit Sub

Erro_objEventoProdutoBenef_evSelecao:

    Select Case gErr

        Case 126891, 126893, 126895, 126919

        Case 126890
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)
        
        Case 126892
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 126894
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156520)

    End Select

    Exit Sub

End Sub

Private Function ProdutoLinha_PreencheBenef(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iCclPreenchida As Integer
Dim sContaEnxuta As String
Dim sAlmoxarifadoPadrao As String

On Error GoTo Erro_ProdutoLinha_PreencheBenef

    If Trim(Len(objProduto.sContaContabilProducao)) > 0 Then

        lErro = Mascara_RetornaContaEnxuta(objProduto.sContaContabilProducao, sContaEnxuta)
        If lErro <> SUCESSO Then gError 126889

        'preenche  a ContaContabilProducao
        ContaContabilProducaoBenef.PromptInclude = False
        ContaContabilProducaoBenef.Text = sContaEnxuta
        ContaContabilProducaoBenef.PromptInclude = True

        'preenche Conta De Producao
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilProducaoBenef_Col) = ContaContabilProducaoBenef.Text

    End If

    'Descricao
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_DescricaoItemBenef_Col) = objProduto.sDescricao

    'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
    lErro = CF("AlmoxarifadoPadrao_Le_NomeReduzido", objProduto.sCodigo, sAlmoxarifadoPadrao)
    If lErro <> SUCESSO Then gError 126906

    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col) = sAlmoxarifadoPadrao

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridMovimentos.Row - GridMovimentos.FixedRows) = objGridBenef.iLinhasExistentes Then
        objGridBenef.iLinhasExistentes = objGridBenef.iLinhasExistentes + 1
    End If

    ProdutoLinha_PreencheBenef = SUCESSO

    Exit Function

Erro_ProdutoLinha_PreencheBenef:

    ProdutoLinha_PreencheBenef = gErr

    Select Case gErr

        Case 126889
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objProduto.sContaContabilProducao)
        
        Case 126906
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156521)

    End Select

    Exit Function

End Function

Private Function Move_GridMovimentos_Memoria(objNFiscal As ClassNFiscal) As Long
'Move os Itens do Grid de Movimentos para a Memória - - -> isto é para a coleção que está dentro do objMovEstoque que esta dentro do objNFiscal

Dim iIndice As Integer
Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sContaFormatadaProducao As String
Dim sContaFormatadaEst As String
Dim iContaPreenchida As Integer
Dim objItemNF As ClassItemNF
Dim dCusto As Double
Dim iItemNF As Integer
Dim objProduto As New ClassProduto
Dim objItemMovEstoque As ClassItemMovEstoque

On Error GoTo Erro_Move_GridMovimentos_Memoria

    Set objNFiscal.objMovEstoque = New ClassMovEstoque

    'Para cada linha existente do Grid
    For iIndice = 1 To objGridBenef.iLinhasExistentes

        'Formata o Produto para BD
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_ProdutoBenef_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 126911

        objProduto.sCodigo = sProdutoFormatado

        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 126925

        'Se não achou o Produto --> erro
        If lErro = 28030 Then gError 126926

        'Pega o Código do Almoxarifado
        Set objAlmoxarifado = New ClassAlmoxarifado

        objAlmoxarifado.sNomeReduzido = GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxarifadoBenef_Col)

        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then gError 126912

        If lErro = 25056 Then gError 126913

        If GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilEstBenef_Col) <> "" Then

            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilEstBenef_Col), sContaFormatadaEst, iContaPreenchida)
            If lErro <> SUCESSO Then gError 126914

        Else
            sContaFormatadaEst = ""
        End If

        If GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilProducaoBenef_Col) <> "" Then

            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilProducaoBenef_Col), sContaFormatadaProducao, iContaPreenchida)
            If lErro <> SUCESSO Then gError 126915

        Else
            sContaFormatadaProducao = ""
        End If

        iItemNF = GridMovimentos.TextMatrix(iIndice, iGrid_ItemNFBenef_Col)
        Set objItemNF = objNFiscal.ColItensNF(iItemNF)

        dCusto = (((objItemNF.dQuantidade * objItemNF.dPrecoUnitario) - objItemNF.dValorDesconto) / objItemNF.dQuantidade) * StrParaDbl(GridMovimentos.TextMatrix(iIndice, iGrid_QuantInsumoBenef_Col))
    
        Set objItemMovEstoque = objNFiscal.objMovEstoque.colItens.Add(0, 0, dCusto, 0, sProdutoFormatado, GridMovimentos.TextMatrix(iIndice, iGrid_DescricaoItemBenef_Col), objProduto.sSiglaUMEstoque, 0, objAlmoxarifado.iCodigo, GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxarifadoBenef_Col), 0, "", 0, "", "", sContaFormatadaProducao, sContaFormatadaEst, 0, Nothing, Nothing, DATA_NULA)

        objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE
        objItemMovEstoque.iClasseUM = objProduto.iClasseUM
        objItemMovEstoque.iControleEstoque = objProduto.iControleEstoque
        objItemMovEstoque.sSiglaUMEst = objProduto.sSiglaUMEstoque
        objItemMovEstoque.dQuantInsumos = StrParaDbl(GridMovimentos.TextMatrix(iIndice, iGrid_QuantInsumoBenef_Col))
        objItemMovEstoque.iItemNF = StrParaInt(GridMovimentos.TextMatrix(iIndice, iGrid_ItemNFBenef_Col))

    Next

    Move_GridMovimentos_Memoria = SUCESSO

    Exit Function

Erro_Move_GridMovimentos_Memoria:

    Move_GridMovimentos_Memoria = gErr

    Select Case gErr

        Case 126911, 126912, 126914, 126915

        Case 126913
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case 126926
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156522)

    End Select

    Exit Function

End Function

Private Function Preenche_ContaContabilEst(sProduto As String) As Long
'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto
'Caso nao encontre -----> não tratar erro

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sContaEnxuta As String

On Error GoTo Erro_Preenche_ContaContabilEst

    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col))) = 0 Or Len(Trim(sProduto)) = 0 Then Exit Function

    'preenche o objEstoqueProduto
    objAlmoxarifado.sNomeReduzido = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col)

    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25060 Then gError 126914

    If lErro = 25060 Then gError 126915

    'Formata o Produto para BD
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 126916

    objEstoqueProduto.sProduto = sProdutoFormatado
    objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo

    lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 49991 Then gError 126917

    If lErro <> 49991 Then

        lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then gError 126918

        ContaContabilEstBenef.PromptInclude = False
        ContaContabilEstBenef.Text = sContaEnxuta
        ContaContabilEstBenef.PromptInclude = True

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilEstBenef_Col) = ContaContabilEstBenef.Text

    Else

        'Preenche em branco a conta de estoque no grid
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilEstBenef_Col) = ""

    End If

    Preenche_ContaContabilEst = SUCESSO

    Exit Function

Erro_Preenche_ContaContabilEst:

    Preenche_ContaContabilEst = gErr

        Select Case gErr

            Case 126914, 126916, 126917

            Case 126918
                Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objEstoqueProduto.sContaContabil)

            Case 126915
                Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

            Case Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156523)

        End Select

        Exit Function

End Function

Public Sub BotaoEstoqueBenef_Click()

Dim lErro As Long
Dim objEstoqueProduto As ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoqueBenef_Click

    If (GridMovimentos.Row = 0) Then gError 126920
    
    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 126921

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 126922

    colSelecao.Add sProdutoFormatado

    Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoqueBenef)

    Exit Sub

Erro_BotaoEstoqueBenef_Click:

    Select Case gErr

        Case 126920
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 126921
                
        Case 126922
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156524)

    End Select

    Exit Sub

End Sub

Private Sub objEventoEstoqueBenef_evselecao(obj1 As Object)

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sCodProduto As String

On Error GoTo Erro_objEventoEstoqueBenef_evselecao

    Set objEstoqueProduto = obj1

    If (GridMovimentos.Row > 0) Then

        sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col)

        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 126923

        'Verifica se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Preenche o Nome do Almoxarifado
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            AlmoxarifadoBenef.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoqueBenef_evselecao:

    Select Case gErr

        Case 126923, 126924

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156525)

    End Select

    Exit Sub

End Sub

Function Preenche_GridMovimentos(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com os ítens da Nota Fiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objItemMovEstoque As ClassItemMovEstoque
Dim sProdutoEnxuto As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sContaEnxutaEst As String
Dim sContaEnxutaProducao As String

On Error GoTo Erro_Preenche_GridMovimentos

    'Para cada ítem da Coleção
    For Each objItemMovEstoque In objNFiscal.objMovEstoque.colItens

        iIndice = iIndice + 1

        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemMovEstoque.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 52496

        ProdutoBenef.PromptInclude = False
        ProdutoBenef.Text = sProdutoEnxuto
        ProdutoBenef.PromptInclude = True

        'Preenche o Grid
        GridMovimentos.TextMatrix(iIndice, iGrid_ItemNFBenef_Col) = objItemMovEstoque.iItemNF
        GridMovimentos.TextMatrix(iIndice, iGrid_ProdutoBenef_Col) = ProdutoBenef.Text
        GridMovimentos.TextMatrix(iIndice, iGrid_DescricaoItemBenef_Col) = objItemMovEstoque.sProdutoDesc
        GridMovimentos.TextMatrix(iIndice, iGrid_QuantInsumoBenef_Col) = Formata_Estoque(objItemMovEstoque.dQuantInsumos)
        GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxarifadoBenef_Col) = objItemMovEstoque.sAlmoxarifadoNomeRed

        'preenche contaEst no grid
        If objItemMovEstoque.sContaContabilEst <> "" Then

            sContaEnxutaEst = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilEst, sContaEnxutaEst)
            If lErro <> SUCESSO Then gError 52497

            ContaContabilEstBenef.PromptInclude = False
            ContaContabilEstBenef.Text = sContaEnxutaEst
            ContaContabilEstBenef.PromptInclude = True

            GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilEstBenef_Col) = ContaContabilEstBenef.Text

        End If

        'preenche contaEst no grid
        If objItemMovEstoque.sContaContabilAplic <> "" Then

            sContaEnxutaProducao = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilAplic, sContaEnxutaProducao)
            If lErro <> SUCESSO Then gError 52498

            ContaContabilProducaoBenef.PromptInclude = False
            ContaContabilProducaoBenef.Text = sContaEnxutaProducao
            ContaContabilProducaoBenef.PromptInclude = True

            GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilProducaoBenef_Col) = ContaContabilProducaoBenef.Text

        End If

    Next

    'Atualiza o número de linhas existentes
    objGridBenef.iLinhasExistentes = iIndice

    Preenche_GridMovimentos = SUCESSO

    Exit Function

Erro_Preenche_GridMovimentos:

    Preenche_GridMovimentos = gErr

    Select Case gErr

        Case 52496
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemMovEstoque.sProduto)

        Case 52497
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objItemMovEstoque.sContaContabilEst)

        Case 52498
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objItemMovEstoque.sContaContabilAplic)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156526)

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
    If lErro <> SUCESSO Then gError 134034

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134034

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156527)

    End Select
    
    Exit Sub

End Sub

Private Sub Fornecedor_Preenche()
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134065

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134065

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156528)

    End Select
    
    Exit Sub

End Sub

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long, sID As String
Dim objNFiscal As New ClassNFiscal
Dim objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_Anotacao_Extrai

    'Lê os dados da Tela NFiscal
    lErro = Move_NF_Memoria(objNFiscal, objTipoDocInfo)
    If lErro <> SUCESSO Then gError 130404

    If objNFiscal.lNumNotaFiscal <> 0 And objNFiscal.sSerie <> "" And objNFiscal.dtDataEmissao <> DATA_NULA Then
    
        lErro = CF("NFiscal_Le_NumFornCli", objNFiscal)
        If lErro <> SUCESSO And lErro <> 35279 Then gError 130405
        
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156529)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156530)
     
    End Select
     
    Exit Sub

End Sub

Public Sub ValorDesconto_Validate(Cancel As Boolean)

Dim dValorDesconto As Double
Dim lErro As Long
Dim dSubTotal As Double
Dim dValorFrete As Double
Dim dValorSeguro As Double

On Error GoTo Erro_ValorDesconto_Validate

    'Verifica se o valor foi alterado
    If iValorDescontoAlterado = 0 Then Exit Sub

    'Se estiver preenchido
    If Len(Trim(ValorDesconto.Text)) > 0 Then
        
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(ValorDesconto.Text)
        If lErro <> SUCESSO Then gError 65777

        dValorDesconto = CDbl(ValorDesconto.Text)
        'Coloca o valor formatado na tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")

    End If

    Call SubTotal_Calcula

''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDesconto_Validate(Cancel, dValorDesconto)
'*** fim tributacao
   
    Call ValorTotal_Calcula

    iValorDescontoAlterado = 0

    Exit Sub

Erro_ValorDesconto_Validate:

    Cancel = True


    Select Case gErr

        Case 65777

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156531)

    End Select

    Exit Sub

End Sub

Public Sub ValorDesconto_Change()
    iAlterado = REGISTRO_ALTERADO
    iValorDescontoAlterado = REGISTRO_ALTERADO
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

Public Function NFEDEV_Preenche_GridItens(ByVal objCT As Object, ByVal objItemNF As ClassItemNF, ByVal iLinha As Integer, dQuantidade As Double) As Long
    NFEDEV_Preenche_GridItens = SUCESSO
End Function

Public Function NFEDEV_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iLocalChamada As Integer) As Long
    NFEDEV_Rotina_Grid_Enable = SUCESSO
End Function

Public Function NFEDEV_Calcula_PesoKg(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    NFEDEV_Calcula_PesoKg = SUCESSO
End Function

Public Function NFEDEV_Calcula_PesoKg_Total(ByVal objCT As Object) As Long
    NFEDEV_Calcula_PesoKg_Total = SUCESSO
End Function

Public Function NFEDEV_Calcula_Valores(ByVal objCT As Object, iLinha As Integer, dQuantidade As Double) As Long
    NFEDEV_Calcula_Valores = SUCESSO
End Function

Public Function NFEDEV_Move_GridItens_Memoria(ByVal objCT As Object, ByVal objItemNF As ClassItemNF, ByVal iLinha As Integer) As Long
    NFEDEV_Move_GridItens_Memoria = SUCESSO
End Function

Public Function NFEDEV_ProdutoLinha_Preenche(ByVal objCT As Object, objProduto As ClassProduto) As Long
    NFEDEV_ProdutoLinha_Preenche = SUCESSO
End Function

Public Function NFEDEV_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    NFEDEV_Saida_Celula = SUCESSO
End Function

Public Function NFEDEV_ValorTotal(ByVal objItemNF As ClassItemNF) As Long
    NFEDEV_ValorTotal = SUCESSO
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
        objItensPRJCR.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) & SEPARADOR & GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)

        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objItensPRJCR.dValorOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))

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
    
        Case 500130, 500131
        
        Case 500132
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157167)
    
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

Dim lErro As Long
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer
Dim colTipoFrete As Collection

On Error GoTo Erro_Carrega_TipoFrete
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207451

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

        Case 207451

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207452)

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
        
        If Len(Trim(ValorDescontoItens.Text)) > 0 Then
            dValorDesconto = CDbl(ValorDescontoItens.Text)
        Else
            dValorDesconto = 0
        End If
        
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
        If dTotal > 0 Then
            dFator = dDesconto / dTotal
        Else
            dFator = 0
        End If
        
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
                Call Calcula_Valores(iIndice)
                
                lErro = gobjTribTab.Alteracao_Item_Grid(iIndice)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                dValorTotal = dValorTotal + StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
            End If
        Next
    
    End If
        
    'Coloca valor total dos produtos na tela
    SubTotal.Caption = Format(dValorTotal, "Standard")

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
        
        Call Chama_Tela_Modal("InfoAdicDocItem", objItemNF.objInfoAdicDocItem, objTela, TIPO_ENTRADA)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Call SubTotal_Calcula
        
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

Public Sub BotaoTrazerNFe_Click()

Dim lErro  As Long
Dim objTela As Object, objNF As New ClassNFiscal

On Error GoTo Erro_BotaoTrazerNFe_Click

    If Len(Trim(ChvNFe.ClipText)) = 0 Then gError 211286

    Set objTela = Me
    Set gobjNFiscal = objNF
    
    lErro = CF("NFe_Chv_Monta_NF", objTela, objNF, ChvNFe.ClipText)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If objNF.lNumNotaFiscal > 0 Then
        
        lErro = Traz_NFiscal_Tela1(objNF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        Call ValorTotal_Calcula
    
    End If
    
    Exit Sub

Erro_BotaoTrazerNFe_Click:

    Select Case gErr
    
        Case 211286
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CHAVE_NAO_PREENCHIDA", gErr)
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211287)
            
    End Select
    
    Exit Sub

End Sub

Public Sub ChvNFeLabel_Click()

Dim lErro As Long, objNF As New ClassNFiscal
Dim colSelecao As New Collection
Dim iFilialEmpresa As Integer, objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_ChvNFeLabel_Click

    objNF.sChvNFe = ChvNFe.ClipText
    
    iFilialEmpresa = giFilialEmpresa
    If iFilialEmpresa > 50 Then iFilialEmpresa = iFilialEmpresa - 50
    
    objFilialEmpresa.iCodFilial = iFilialEmpresa
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    colSelecao.Add objFilialEmpresa.sCgc
    
    Call Chama_Tela("NFeXmlLista", colSelecao, objNF, objEventoChvNFe, "CGCEmitente <> ?")

    Exit Sub
    
Erro_ChvNFeLabel_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201452)

    End Select
    
    Exit Sub

End Sub

Private Sub objEventoChvNFe_evSelecao(obj1 As Object)

Dim bCancel As Boolean
Dim objNFiscal As New ClassNFiscal
Dim lErro As Long

On Error GoTo Erro_objEventoChvNFe_evSelecao

    Set objNFiscal = obj1

    ChvNFe.PromptInclude = False
    ChvNFe.Text = objNFiscal.sChvNFe
    ChvNFe.PromptInclude = True
    
    Call BotaoTrazerNFe_Click

    Exit Sub
    
Erro_objEventoChvNFe_evSelecao:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211288)

    End Select
    
    Exit Sub
    
End Sub

Public Sub ChvNFe_Validate(Cancel As Boolean)

Dim lErro As Long, objNF As New ClassNFiscal
Dim bExisteNF As Boolean, bExisteXml As Boolean, bTrazer As Boolean, lNumIntNF As Long

On Error GoTo Erro_ChvNFe_Validate

    If gsChvAnt <> ChvNFe.ClipText Then

        lErro = CF("NFe_Chv_Valida", ChvNFe.ClipText, bExisteNF, bExisteXml, bTrazer, lNumIntNF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If bTrazer Then
            If bExisteNF Then
                objNF.lNumIntDoc = lNumIntNF
                lErro = Traz_NFiscal_Tela(objNF)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            ElseIf bExisteXml Then
                Call BotaoTrazerNFe_Click
            End If
        End If
        
        gsChvAnt = ChvNFe.ClipText
        
    End If
       
    Exit Sub

Erro_ChvNFe_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211289)

    End Select

    Exit Sub

End Sub

Public Sub ChvNFe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ChvNFe_GotFocus()
    Call MaskEdBox_TrataGotFocus(ChvNFe, iAlterado)
End Sub

Private Function Traz_NFiscal_Tela1(objNFiscal As ClassNFiscal) As Long
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

On Error GoTo Erro_Traz_NFiscal_Tela1
    
    gbCarregandoTela = True
    
    'Limpa a tela
    Call Limpa_Tela_NFiscal1
    
    Set gobjNFiscal = objNFiscal
    
    ChvNFe.PromptInclude = False
    ChvNFe.Text = objNFiscal.sChvNFe
    ChvNFe.PromptInclude = True
    gsChvAnt = objNFiscal.sChvNFe
    
    If objNFiscal.lNumNotaFiscal <> 0 Then NFiscal.Text = objNFiscal.lNumNotaFiscal

    If objNFiscal.lFornecedor <> 0 Then
        'preenche o Fornecedor
        Fornecedor.Text = objNFiscal.lFornecedor
        Call Fornecedor_Validate(bCancel)
        
        'Preenche a Filial do Fornecedor
        Filial.Text = objNFiscal.iFilialForn
        Call Filial_Validate(bCancel)
    ElseIf objNFiscal.lCliente <> 0 Then
        'preenche o Cliente
        Cliente.Text = objNFiscal.lCliente
        Call Cliente_Validate(bCancel)
        
        'Preenche a Filial do Cliente
        Filial.Text = objNFiscal.iFilialCli
        Call Filial_Validate(bCancel)
    End If

    TipoNFiscal.Text = objNFiscal.iTipoNFiscal
    Call TipoNFiscal_Validate(bCancel)

    If objNFiscal.sNaturezaOp <> "" Then NaturezaOp.Text = objNFiscal.sNaturezaOp
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

    DataEntrada.PromptInclude = False
    If objNFiscal.dtDataEntrada <> DATA_NULA Then DataEntrada.Text = Format(objNFiscal.dtDataEntrada, "dd/mm/yy")
    DataEntrada.PromptInclude = True

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
    If objNFiscal.dValorTotal > 0 Then Total.Text = Format(objNFiscal.dValorTotal, "Standard")
    If objNFiscal.dValorProdutos > 0 Then SubTotal.Caption = Format(objNFiscal.dValorProdutos, "Standard")

    If objNFiscal.iCodTransportadora <> 0 Then
        Transportadora.Text = objNFiscal.iCodTransportadora
        Call Transportadora_Validate(bCancel)
    End If

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

    Mensagem.Text = objNFiscal.sMensagemNota
    MensagemCorpo.Text = objNFiscal.sMensagemCorpoNota
'    Observacao.Text = objNFiscal.sObservacao
    
    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = CStr(objNFiscal.dPesoLiq)
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = CStr(objNFiscal.dPesoBruto)
    
    'preenche o Grid com os Ítens da Nota Fiscal
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        GridItens.Row = iIndice
        
        lErro = gobjDistribuicao.Distribuicao_Processa()
        'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Next
    
    'distribuicao
    'Preenche o Grid com as Distribuições dos itens da Nota Fiscal
'    lErro = gobjDistribuicao.Preenche_GridDistribuicao(objNFiscal)
'    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

'    'Preenche a Condicao de Pagto
'    If objNFiscal.iCondicaoPagto > 0 Then
'        CondicaoPagamento.Text = objNFiscal.iCondicaoPagto
'        Call CondicaoPagamento_Validate(bCancel)
'    End If
'
'    lErro = Preenche_GridParcelas(objNFiscal)
'    If lErro <> SUCESSO Then gError 188517
'
'    'Tenta fazer a distribuição automatica
'    lErro = gobjDistribuicao.Preenche_GridDistribuicaoPC1(colItemPedCompraInfo)
'    If lErro <> SUCESSO And lErro <> 89627 Then gError 188518

    'Set gobjNFiscal = objNFiscal

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'*** fim tributacao

    gbCarregandoTela = False
    
    Traz_NFiscal_Tela1 = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela1:

    gbCarregandoTela = False
    
    Traz_NFiscal_Tela1 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211290)

    End Select

    Exit Function

End Function

