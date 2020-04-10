VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalFatEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim giTeste As Integer

Public gobjAnotacao As ClassAnotacoes
Public gobjInfoAdic As ClassInfoAdic
Dim dValorDescontoItensAnt As Double
Dim dPercDescontoItensAnt As Double
Dim iDescontoAlterado As Integer

Public gobjInfoUsu As Object

Dim m_objUserControl As Object

'Flag que indica se a tela está sendo preenchida.
Public gbCarregandoTela As Boolean
Public giPosCargaOk As Integer
Dim giComplemento As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

'************ GRADE **********************
Public gobjNFiscal As ClassNFiscal
'*****************************************

'inicio contabilidade
Public objGrid1 As AdmGrid
Public objContabil As New ClassContabil

Public WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Public WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

'Mnemônicos
Private Const TIPO_NOTA_FISCAL As String = "Tipo_Nota_Fiscal"
Private Const NATUREZA_OP As String = "Natureza_OP"
Private Const FORNECEDOR1 As String = "Fornecedor"
Private Const FILIAL1 As String = "Filial_Fornecedor"
Private Const SERIE1 As String = "Serie"
Private Const NFISCAL1 As String = "Nota_Fiscal"
Private Const DATA_EMISSAO As String = "Data_Emissao"
Private Const DATA_ENTRADA As String = "Data_Entrada"
Private Const PRODUTO1 As String = "Produto_Codigo"
Private Const UNIDADE_MED As String = "Unidade_Med"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const PRECO_UNITARIO As String = "Valor_Unitario_Prod"
Private Const PRECO_TOTAL As String = "Valor_Total_Produto"
Private Const ALMOXARIFADO1 As String = "Almoxarifado"
Private Const DESCONTO1 As String = "Desconto_Produto"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"
Private Const PRODUTO_BENEFICIADO As String = "Produto_Beneficiado"
Private Const PRODUTO_DIST As String = "Produto_Dist"
Private Const PRODUTO_ITENS As String = "Produto_Itens"
Private Const QUANT_DIST As String = "Quant_Dist"
Private Const FATOR_ITEM_EST As String = "F1"
Private Const PRODUTO_TIPO_EST As String = "Prod_Dist_Tipo"
Private Const PRODUTO_TIPO As String = "Prod_Tipo"
'#######Valores####################
Private Const ICMS As String = "ICMS_Valor"
Private Const ICMSSUBST As String = "ICMSSubst_Valor"
Private Const VALOR_PRODUTOS As String = "Valor_Produtos"
Private Const VALOR_FRETE As String = "Valor_Frete"
Private Const VALOR_SEGURO As String = "Valor_Seguro"
Private Const VALOR_DESPESAS As String = "Valor_Despesas"
Private Const IPI As String = "IPI_Valor"
Private Const VALOR_DESCONTO As String = "Valor_Desconto_Glb"
Private Const VALOR_TOTAL As String = "Total_Nota_Fiscal"
'###########Pagamento###############
Private Const VALOR_PARCELA As String = "Valor_Parcela"
Private Const SUSPENSO1 As String = "Suspenso"
'###########Tributação##############
Private Const ISS_VALOR As String = "ISS_Valor"
Private Const ISS_INCLUSO As String = "ISS_Incluso"
Private Const VALOR_IRRF As String = "Valor_IRRF"
Private Const VALOR_INSS As String = "Valor_INSS"
Private Const INSS_RETIDO As String = "INSS_Retido"
Private Const ICMS_CREDITO_VALOR As String = "ICMS_Credito_Valor"
Private Const IPI_CREDITO_VALOR As String = "IPI_Credito_Valor"
Private Const CTACONTABILEST1 As String = "ContaContabilEst"
Private Const CTAFILIALFORNECEDOR1 As String = "FilialForn_Conta_Ctb"
Private Const CTACONTABILESTBENEF As String = "CtaContabilEstBenef"
Private Const CTACONTABILPRODBENEF As String = "CtaContabilProdBenef"
Private Const ICMS_CREDITO_FRETE As String = "ICMS_Credito_Frete"
Private Const ICMS_CREDITO_SEGURO As String = "ICMS_Credito_Seguro"
Private Const ICMS_CREDITO_OUTRASDESP As String = "ICMS_Cred_OutrasDesp"
Private Const ICMS_CREDITO_DESCONTO As String = "ICMS_Cred_Desconto"
Private Const IPI_CREDITO_FRETE As String = "IPI_Credito_Frete"
Private Const IPI_CREDITO_SEGURO As String = "IPI_Credito_Seguro"
Private Const IPI_CREDITO_OUTRASDESP As String = "IPI_Cred_OutrasDesp"
Private Const IPI_CREDITO_DESCONTO As String = "IPI_Cred_Desconto"
Private Const IPI_PRODUTO As String = "IPI_Produto"
Private Const QUANT_DISPONIVEL As String = "Quant_Disponivel"
Private Const QUANT_BENEF As String = "Quant_Benef"
Private Const ICMS_CREDITO_TOTAL As String = "ICMS_Credito_Total"
Private Const IPI_CREDITO_TOTAL As String = "IPI_Credito_Total"
Private Const PIS_CREDITO_TOTAL As String = "PIS_Credito_Total"
Private Const COFINS_CREDITO_TOTAL As String = "COFINS_Credito_Total"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const ISS_RETIDO As String = "ISS_Retido"
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
Private Const CONTACONTABILCATPROD As String = "ContaContabilCatProd"
'Fim da Contabilidade

'Rastreamento
'Public gcolItensNF As Collection
'Public objGenerico As AdmGenerico
Public gobjRastreamento As ClassRastreamento

'distribuicao
Public gobjDistribuicao As Object

Public iAlterado As Integer

Public iFrameAtual As Integer
Public iLinhaAntiga As Integer
Public iEmitenteNF As Integer
Public iFornecedorAlterado As Integer
Public iFornNFOrigAlterado As Integer 'Incluído por Luiz em 25/01/02
Public iFornecedorBenefAlterado As Integer
Public iNaturezaOpAlterada As Integer
Public iTotalAlterado As Integer
Public iValorDescontoAlterado As Integer
Public iValorFreteAlterado As Integer
Public iValorSeguroAlterado As Integer
Public iValorDespesasAlterado As Integer
Public iDataEntradaAlterada As Integer
Public iSerieAlterada As Integer

Public objGridItens As AdmGrid
Public objGridParcelas As AdmGrid

Public iGrid_Produto_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer

Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Quantidade_Col As Integer
Public iGrid_PrecoTotal_Col As Integer
Public iGrid_PrecoTotalB_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_ValorUnitario_Col As Integer
Public iGrid_ValorUnitarioMoeda_Col As Integer
'distribuicao
'Public iGrid_Almoxarifado_Col As Integer
Public iGrid_Ccl_Col As Integer

Public iGrid_Parcela_Col  As Integer
Public iGrid_Vencimento_Col As Integer
Public iGrid_VenctoReal_Col As Integer
Public iGrid_Valor_Col  As Integer
Public iGrid_Cobranca_Col As Integer
Public iGrid_Banco_Col As Integer
Public iGrid_Portador_Col As Integer
Public iGrid_Suspenso_Col As Integer
Public iGrid_CodigoDeBarras_Col As Integer

'Grid relacionado com Beneficiamento
Public objGridBenef As AdmGrid

Dim iGrid_ProdutoBenef_Col As Integer
Dim iGrid_UnidadeMedBenef_Col As Integer
Dim iGrid_QuantidadeBenef_Col As Integer
Dim iGrid_DescricaoBenef_Col As Integer
Dim iGrid_AlmoxarifadoBenef_Col As Integer
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
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Public WithEvents objEventoNFiscal As AdmEvento
Attribute objEventoNFiscal.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1
Private WithEvents objEventoCondPagto As AdmEvento
Attribute objEventoCondPagto.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorBenef As AdmEvento
Attribute objEventoFornecedorBenef.VB_VarHelpID = -1
Private WithEvents objEventoFornNFOrig As AdmEvento 'Incluído por Luiz em 25/01/02
Attribute objEventoFornNFOrig.VB_VarHelpID = -1
Private WithEvents objEventoRecebimento As AdmEvento
Attribute objEventoRecebimento.VB_VarHelpID = -1
Private WithEvents objEventoSerieOriginal As AdmEvento
Attribute objEventoSerieOriginal.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1
Private WithEvents objEventoNFiscalOriginal As AdmEvento
Attribute objEventoNFiscalOriginal.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoContrato As AdmEvento
Attribute objEventoContrato.VB_VarHelpID = -1
Private WithEvents objEventoChvNFe As AdmEvento
Attribute objEventoChvNFe.VB_VarHelpID = -1

Dim gsChvAnt As String

Private WithEvents objEventoItemContrato As AdmEvento
Attribute objEventoItemContrato.VB_VarHelpID = -1

Public iGrid_Contrato_Col As Integer
Public iGrid_ItemContrato_Col As Integer
Public iGrid_DataCobranca_Col As Integer

Dim gcolcolMedicoes As Collection

Public gobjTelaProjetoInfo As ClassTelaPRJInfo

Public gobjTribTab As ClassTribTab

'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_Itens = 2
Private Const TAB_Complemento = 3
'distribuicao
Private Const TAB_Distribuicao = 4
Private Const TAB_Pagamento = 5
Private Const TAB_Tributacao = 6
Private Const TAB_Contabilizacao = 7
'Rastreamento
Private Const TAB_RASTRO = 8
Private TAB_BENEFICIAMENTO As Integer

Const TIPODOCINFO_COMPLEMFAT_INT_IPI = 63
Const TIPODOCINFO_COMPLEMFAT_EXT_IPI = 60

'Constante copiada da tela CTNFiscalEntrada
Const CONFIRMA_VUNITARIO_GRAVACAO = 0 'Parametro iIndice
'O Valor Unitário do item %s não esta preenchido. Deseja Prosseguir? - iIndice

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

Public Property Get Frame15() As Object
     Set Frame15 = objUserControl.Controls("Frame15")
End Property

Public Property Get Frame10() As Object
     Set Frame10 = objUserControl.Controls("Frame10")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get Frame13() As Object
     Set Frame13 = objUserControl.Controls("Frame13")
End Property

Public Property Get Frame9() As Object
     Set Frame9 = objUserControl.Controls("Frame9")
End Property

Public Property Get Frame11() As Object
     Set Frame11 = objUserControl.Controls("Frame11")
End Property

Public Property Get Frame12() As Object
     Set Frame12 = objUserControl.Controls("Frame12")
End Property

Public Property Get CTBSeqContraPartida() As Object
     Set CTBSeqContraPartida = objUserControl.Controls("CTBSeqContraPartida")
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

Public Property Get CTBBotaoImprimir() As Object
     Set CTBBotaoImprimir = objUserControl.Controls("CTBBotaoImprimir")
End Property

Public Property Get CTBFrame7() As Object
     Set CTBFrame7 = objUserControl.Controls("CTBFrame7")
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

Public Property Get CTBCclDescricao() As Object
     Set CTBCclDescricao = objUserControl.Controls("CTBCclDescricao")
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

Public Property Get BotaoLimparNF() As Object
     Set BotaoLimparNF = objUserControl.Controls("BotaoLimparNF")
End Property

Public Property Get Recebimento() As Object
     Set Recebimento = objUserControl.Controls("Recebimento")
End Property

Public Property Get RecebimentoLabel() As Object
     Set RecebimentoLabel = objUserControl.Controls("RecebimentoLabel")
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

Public Property Get NaturezaOp() As Object
     Set NaturezaOp = objUserControl.Controls("NaturezaOp")
End Property

Public Property Get NumRecebimento() As Object
     Set NumRecebimento = objUserControl.Controls("NumRecebimento")
End Property

Public Property Get NFiscalInterna() As Object
     Set NFiscalInterna = objUserControl.Controls("NFiscalInterna")
End Property

Public Property Get NFiscalLabel() As Object
     Set NFiscalLabel = objUserControl.Controls("NFiscalLabel")
End Property

Public Property Get SerieLabel() As Object
     Set SerieLabel = objUserControl.Controls("SerieLabel")
End Property

Public Property Get NaturezaLabel() As Object
     Set NaturezaLabel = objUserControl.Controls("NaturezaLabel")
End Property

Public Property Get Frame7() As Object
     Set Frame7 = objUserControl.Controls("Frame7")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Fornecedor() As Object
     Set Fornecedor = objUserControl.Controls("Fornecedor")
End Property

Public Property Get FornecedorLabel() As Object
     Set FornecedorLabel = objUserControl.Controls("FornecedorLabel")
End Property

Public Property Get Frame14() As Object
     Set Frame14 = objUserControl.Controls("Frame14")
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

Public Property Get UnidadeMedBenef() As Object
     Set UnidadeMedBenef = objUserControl.Controls("UnidadeMedBenef")
End Property

Public Property Get DescricaoItemBenef() As Object
     Set DescricaoItemBenef = objUserControl.Controls("DescricaoItemBenef")
End Property

Public Property Get BotaoProdutosBenef() As Object
     Set BotaoProdutosBenef = objUserControl.Controls("BotaoProdutosBenef")
End Property

Public Property Get BotaoEstoqueBenef() As Object
     Set BotaoEstoqueBenef = objUserControl.Controls("BotaoEstoqueBenef")
End Property

Public Property Get BotaoPlanoConta() As Object
     Set BotaoPlanoConta = objUserControl.Controls("BotaoPlanoConta")
End Property

Public Property Get ContaContabilProducao() As Object
     Set ContaContabilProducao = objUserControl.Controls("ContaContabilProducao")
End Property

Public Property Get ContaContabilEst() As Object
     Set ContaContabilEst = objUserControl.Controls("ContaContabilEst")
End Property

Public Property Get AlmoxarifadoBenef() As Object
     Set AlmoxarifadoBenef = objUserControl.Controls("AlmoxarifadoBenef")
End Property

Public Property Get QuantidadeBenef() As Object
     Set QuantidadeBenef = objUserControl.Controls("QuantidadeBenef")
End Property

Public Property Get ProdutoBenef() As Object
     Set ProdutoBenef = objUserControl.Controls("ProdutoBenef")
End Property

Public Property Get GridMovimentos() As Object
     Set GridMovimentos = objUserControl.Controls("GridMovimentos")
End Property

Public Property Get QuantDisponivelBenef() As Object
     Set QuantDisponivelBenef = objUserControl.Controls("QuantDisponivelBenef")
End Property

Public Property Get PagamentoAutomatico() As Object
     Set PagamentoAutomatico = objUserControl.Controls("PagamentoAutomatico")
End Property

Public Property Get Suspenso() As Object
     Set Suspenso = objUserControl.Controls("Suspenso")
End Property

Public Property Get TipoCobranca() As Object
     Set TipoCobranca = objUserControl.Controls("TipoCobranca")
End Property

Public Property Get CondicaoPagamento() As Object
     Set CondicaoPagamento = objUserControl.Controls("CondicaoPagamento")
End Property

Public Property Get DataVencimentoReal() As Object
     Set DataVencimentoReal = objUserControl.Controls("DataVencimentoReal")
End Property

Public Property Get ValorParcela() As Object
     Set ValorParcela = objUserControl.Controls("ValorParcela")
End Property

Public Property Get DataVencimento() As Object
     Set DataVencimento = objUserControl.Controls("DataVencimento")
End Property

Public Property Get GridParcelas() As Object
     Set GridParcelas = objUserControl.Controls("GridParcelas")
End Property

Public Property Get CondPagtoLabel() As Object
     Set CondPagtoLabel = objUserControl.Controls("CondPagtoLabel")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get Frame6() As Object
     Set Frame6 = objUserControl.Controls("Frame6")
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

Public Property Get SerieNFiscalOriginal() As Object
     Set SerieNFiscalOriginal = objUserControl.Controls("SerieNFiscalOriginal")
End Property

Public Property Get NFiscalOriginal() As Object
     Set NFiscalOriginal = objUserControl.Controls("NFiscalOriginal")
End Property

Public Property Get SerieOriginalLabel() As Object
     Set SerieOriginalLabel = objUserControl.Controls("SerieOriginalLabel")
End Property

Public Property Get NFiscalOriginalLabel() As Object
     Set NFiscalOriginalLabel = objUserControl.Controls("NFiscalOriginalLabel")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get Mensagem() As Object
     Set Mensagem = objUserControl.Controls("Mensagem")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get PesoLiquido() As Object
     Set PesoLiquido = objUserControl.Controls("PesoLiquido")
End Property

Public Property Get PesoBruto() As Object
     Set PesoBruto = objUserControl.Controls("PesoBruto")
End Property

Public Property Get Frame16() As Object
     Set Frame16 = objUserControl.Controls("Frame16")
End Property

Public Property Get FilialFornBenef() As Object
     Set FilialFornBenef = objUserControl.Controls("FilialFornBenef")
End Property

Public Property Get FornecedorBenef() As Object
     Set FornecedorBenef = objUserControl.Controls("FornecedorBenef")
End Property

Public Property Get FornecedorBenefLabel() As Object
     Set FornecedorBenefLabel = objUserControl.Controls("FornecedorBenefLabel")
End Property

'distribuicao
'Public Property Get BotaoLocalizacao() As Object
'     Set BotaoLocalizacao = objUserControl.Controls("BotaoLocalizacao")
'End Property

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
End Property

Public Property Get Total() As Object
     Set Total = objUserControl.Controls("Total")
End Property

Public Property Get ValorFrete() As Object
     Set ValorFrete = objUserControl.Controls("ValorFrete")
End Property

Public Property Get ValorDespesas() As Object
     Set ValorDespesas = objUserControl.Controls("ValorDespesas")
End Property

Public Property Get ValorSeguro() As Object
     Set ValorSeguro = objUserControl.Controls("ValorSeguro")
End Property

Public Property Get SubTotal() As Object
     Set SubTotal = objUserControl.Controls("SubTotal")
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

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get DescricaoItem() As Object
     Set DescricaoItem = objUserControl.Controls("DescricaoItem")
End Property

Public Property Get UnidadeMed() As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
End Property

Public Property Get Ccl() As Object
     Set Ccl = objUserControl.Controls("Ccl")
End Property

'distribuicao
'Public Property Get Almoxarifado() As Object
'     Set Almoxarifado = objUserControl.Controls("Almoxarifado")
'End Property

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

Public Property Get ValorDesconto() As Object
     Set ValorDesconto = objUserControl.Controls("ValorDesconto")
End Property

Public Property Get BotaoConsultaTitPag() As Object
     Set BotaoConsultaTitPag = objUserControl.Controls("BotaoConsultaTitPag")
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

'################################################
'Inserido por Wagner
Public Property Get BotaoMedicao() As Object
     Set BotaoMedicao = objUserControl.Controls("BotaoMedicao")
End Property

Public Property Get BotaoItemContrato() As Object
     Set BotaoItemContrato = objUserControl.Controls("BotaoItemContrato")
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

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

Public Property Get ContratoLabel() As Object
     Set ContratoLabel = objUserControl.Controls("ContratoLabel")
End Property

Public Property Get ContratoPag() As Object
     Set ContratoPag = objUserControl.Controls("ContratoPag")
End Property

Public Property Get BotaoContrato() As Object
     Set BotaoContrato = objUserControl.Controls("BotaoContrato")
End Property

Public Property Get BotaoVerContrato() As Object
     Set BotaoVerContrato = objUserControl.Controls("BotaoVerContrato")
End Property

Public Property Get CodBarrasAvulso() As Object
     Set CodBarrasAvulso = objUserControl.Controls("CodBarrasAvulso")
End Property

Public Property Get NumNFe() As Object
     Set NumNFe = objUserControl.Controls("NumNFe")
End Property

'################################################

'--- fim dos properties get dos controles da tela

'Private Sub BotaoImprimir_Click()
'
'Dim objNFiscal As New ClassNFiscal
'Dim lErro As Long
'Dim objRelatorio As New AdmRelatorio
'
'On Error GoTo Erro_BotaoImprimir_Click
'
'    'Se a nota fiscal for externa --> Erro
'    If iEmitenteNF <> DOCINFO_EMPRESA Then Error 57622
'
'    'Verifica se os campos obrigatórios estão preenchidos
'    If Len(Trim(Serie.Text)) = 0 Then Error 57623
'    If Len(Trim(NFiscal.ClipText)) = 0 Then Error 57624
'    If Len(Trim(TipoNFiscal.Text)) = 0 Then Error 57625
'    If Len(Trim(DataEmissao.ClipText)) = 0 Then Error 57627
'
'    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
'    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
'    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
'    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
'    objNFiscal.iFilialEmpresa = giFilialEmpresa
'
'    'Tenta ler a nota Fiscal passada por parametro
'    lErro = CF("NFiscal_Interna_TestaExistencia",objNFiscal)
'    If lErro <> SUCESSO And lErro <> 57677 Then Error 57629
'    If lErro = 57677 Then Error 57630
'
'    lErro = objRelatorio.ExecutarDireto("Emissão das Notas Fiscais", "NotaFiscal = @NNFISCALINIC E Serie = @TSERIE", 1, "NFFatE", "TSERIE", objNFiscal.sSerie, "NNFISCALINIC", CStr(objNFiscal.lNumNotaFiscal), "NNFISCALFIM", CStr(objNFiscal.lNumNotaFiscal), "NENTRADASAIDA", "0", "NNFTIPO", "1")
'    If lErro <> SUCESSO Then Error 57631
'
'    Exit Sub
'
'Erro_BotaoImprimir_Click:
'
'    Select Case Err
'
'        Case 57622
'            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_EXTERNA", Err)
'
'        Case 57623
'            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", Err)
'
'        Case 57624
'            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", Err)
'
'        Case 57625
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", Err)
'
'        Case 57627
'            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", Err)
'
'        Case 57629, 57631
'
'        Case 57630
'            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", Err, objNFiscal.lNumNotaFiscal)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156866)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub BotaoLimparNF_Click()

    NFiscalInterna.Caption = ""
    
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

'Rastreamento
'Public Sub BotaoRastro_Click()
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim iIndice2 As Integer
'Dim objProduto As New ClassProduto
'Dim iProdutoPreenchido As Integer
'Dim objItemNF As New ClassItemNF
'
'On Error GoTo Erro_BotaoRastro_Click
'
'    'Para cada ítem do grid
'    For iIndice = 1 To objGridItens.iLinhasExistentes
'
'        'Formata o produto
'        lErro = CF("Produto_Critica",GridItens.TextMatrix(iIndice, iGrid_Produto_Col), objProduto, iProdutoPreenchido)
'        If lErro <> SUCESSO Then gError 75928
'
'        'Verifica se o ItemNF já está no coleção global
'        For iIndice2 = 1 To gcolItensNF.Count
'            If gcolItensNF(iIndice2).iItem = iIndice Then
'                Exit For
'            End If
'        Next
'
'        'Se o ItemNF não estava na coleção
'        If iIndice2 > gcolItensNF.Count Then
'
'            Set objItemNF = New ClassItemNF
'            objItemNF.sProduto = objProduto.sCodigo
'            objItemNF.iItem = iIndice
'            objItemNF.sDescricaoItem = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
'            objItemNF.sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
'            objItemNF.dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
'            objItemNF.sAlmoxarifadoNomeRed = GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
'
'            'Adiciona o ItemNF na coleção
'            gcolItensNF.Add objItemNF
'
'        'Atualiza dados da coleção
'        Else
'
'            gcolItensNF(iIndice2).sDescricaoItem = GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)
'            gcolItensNF(iIndice2).sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
'            gcolItensNF(iIndice2).dQuantidade = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
'            gcolItensNF(iIndice2).sAlmoxarifadoNomeRed = GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
'
'        End If
'
'    Next
'
'    'Chama a tela de Rasteamento dos Produtos
'    Call Chama_Tela_Nova_Instancia("RastroProdNFEST", gcolItensNF, objGenerico)
'
'    'Desabilita a tela de nota fiscal
'    objGenerico.vVariavel = DESABILITA_TELA
'
'    Exit Sub
'
'Erro_BotaoRastro_Click:
'
'    Select Case gErr
'
'        Case 75928
'
'        Case Else
'             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156867)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub BotaoConsultaTitPag_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoConsultaTitPag_Click

    'Verifica se todos os campos necessários para se efetuar a consulta foram preenchidos
    lErro = Critica_CamposNecessarios_ConsultaTitulo()
    If lErro <> SUCESSO Then gError 79710
    
    'Guarda no objNFiscal os dados necessários para consultar o título
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.iTipoDocInfo = Codigo_Extrai(TipoNFiscal.Text)
    objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.dtDataEntrada = StrParaDate(DataEntrada.Text)
    
    'Guarda em objFornecedor o nome reduzido do Fornecedor
    objFornecedor.sNomeReduzido = Fornecedor.Text
    
    'Lê o código do Fornecedor a partir do nome reduzido obtido na tela
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 79727
    
    'Se não encontrou o fornecedor = > erro
    If lErro = 6681 Then gError 79728
    
    'Guarda no objNFiscal o código do fornecedor
    objNFiscal.lFornecedor = objFornecedor.lCodigo
    
    'Lê o NumIntDocCPR da NFiscal e exibe o documento gerado no CPR por essa nota
    lErro = CF("NFiscal_Consulta_DocCPR", objNFiscal)
    If lErro <> SUCESSO And lErro <> 79717 Then gError 79718
    
    'Se não encontrou a nota => erro
    If lErro = 79717 Then gError 79719
    
    Exit Sub
    
Erro_BotaoConsultaTitPag_Click:

    Select Case gErr
    
        Case 79710, 79718, 79727
        
        Case 79719
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA1", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada)
        
        Case 79728
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156868)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoCcls_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As ClassCcl

On Error GoTo Erro_BotaoCcls_Click

    If GridItens.Row = 0 Then Error 52835

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then Error 52836
    
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

    Exit Sub
    
Erro_BotaoCcls_Click:
    
    Select Case Err
    
        Case 52835
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)

        Case 52836
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156869)

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

Public Sub FilialFornBenef_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub FornecedorBenefLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = FornecedorBenef.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedorBenef)

    Exit Sub

End Sub

Public Sub NaturezaOp_GotFocus()
Dim iNaturezaAux As Integer

    iNaturezaAux = iNaturezaOpAlterada
    Call MaskEdBox_TrataGotFocus(NaturezaOp, iAlterado)
    iNaturezaOpAlterada = iNaturezaAux

End Sub

Public Sub NFiscal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NFiscal, iAlterado)

End Sub

Public Sub NFiscalOriginal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NFiscalOriginal, iAlterado)

End Sub

Public Sub NumRecebimento_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumRecebimento, iAlterado)
    
End Sub

Private Sub Class_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) <> 0 And GridItens.Row <> 0 Then

        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then Error 52837

        'Coloca o valor do Ccl na coluna correspondente
        GridItens.TextMatrix(GridItens.Row, iGrid_Ccl_Col) = sCclMascarado

        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case Err

        Case 52837 'Tratado na rotina chamadora
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 156870)

    End Select

    Exit Sub

End Sub

Public Sub Serie_Click()

    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = REGISTRO_ALTERADO

End Sub

Public Sub Total_Change()

    iAlterado = REGISTRO_ALTERADO
    iTotalAlterado = 1
    
End Sub

Public Sub Total_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dTotal As Double

On Error GoTo Erro_Total_Validate

    If iTotalAlterado = 0 Then Exit Sub
    
    'Se estiver preenchido
    If Len(Trim(Total.Text)) > 0 Then

        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(Total.Text)
        If lErro <> SUCESSO Then Error 43804

        dTotal = CDbl(Total.Text)

        'Coloca o valor formatado na tela
        Total.Text = Format(dTotal, "Standard")
        
    End If

    Call Cobranca_Automatica

    iTotalAlterado = 0
    
    Exit Sub

Erro_Total_Validate:

    Cancel = True


    Select Case Err

        Case 43804

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156871)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

Dim lErro As Long

On Error GoTo Erro_Form_Activate

    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then Error 59342
        
    Call TelaIndice_Preenche(Me)

'    'Retorno da tela de RastroProdNF
'    If objGenerico.vVariavel = HABILITA_TELA Then
'
'        'Habilita a tela de nota fiscal
'        Me.Enabled = True
'
'    Else
'
'        'Desabilita a tela de nota fiscal
'        Me.Enabled = False
'
'    End If

    Exit Sub
     
Erro_Form_Activate:

    Select Case Err
          
        Case 59342
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156872)
     
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

        ValorUnitario.Format = FORMATO_PRECO_UNITARIO_EXTERNO

        Call GridItens_InicializaColunas
        Set gobjTribTab = New ClassTribTab
        lErro = gobjTribTab.Ativar(Me, 1)
        If lErro <> SUCESSO Then gError 42669
    
        'Rastreamento
        Set gobjRastreamento = New ClassRastreamento
        Set gobjRastreamento.objTela = Me

        'distribuicao
        Set gobjDistribuicao = CreateObject("RotinasMat.ClassMATDist")
        Set gobjDistribuicao.objTela = Me
        Set gobjDistribuicao.objRastreamento = gobjRastreamento

    'Inicio alteracao Daniel em 03/10/2001
    
        'Carrega na combo de Banco de Cobradores
        lErro = Carrega_Bancos()
        If lErro <> SUCESSO Then gError 95162
        
        'Carrega na combo de Portadores
        lErro = Carrega_Portadores()
        If lErro <> SUCESSO Then gError 95163
        
    'Fim alteracao Daniel em 03/10/2001
    
        lErro = Carrega_TiposDocInfo()
        If lErro <> SUCESSO Then gError 42471
    
        lErro = Carrega_Serie()
        If lErro <> SUCESSO Then gError 42472
    
        lErro = Carrega_PlacaUF()
        If lErro <> SUCESSO Then gError 42473
    
        lErro = Carrega_Transportadora()
        If lErro <> SUCESSO Then gError 42474
    
        'Incluído por Luiz Nogueira em 21/08/03
        'Carrega a combo VolumeEspecie
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
        If lErro <> SUCESSO Then gError 102460
    
        'Incluído por Luiz Nogueira em 21/08/03
        'Carrega a combo VolumeMarca
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
        If lErro <> SUCESSO Then gError 102461
        
        lErro = CF("Carrega_CondicaoPagamento", CondicaoPagamento, MODULO_CONTASAPAGAR)
        If lErro <> SUCESSO Then gError 42475
    
        lErro = Carrega_TipoCobranca()
        If lErro <> SUCESSO Then gError 42476
    
        Quantidade.Format = FORMATO_ESTOQUE
    
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
        If lErro <> SUCESSO Then gError 42477
    
        'Inicializa a Mascara de Ccl
        lErro = Inicializa_MascaraCcl()
        If lErro <> SUCESSO Then gError 49394
        
        Set objEventoSerie = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoNFiscal = New AdmEvento
        Set objEventoNatureza = New AdmEvento
        Set objEventoCondPagto = New AdmEvento
        Set objEventoFornecedor = New AdmEvento
        Set objEventoFornecedorBenef = New AdmEvento
        Set objEventoFornNFOrig = New AdmEvento
        Set objEventoRecebimento = New AdmEvento
        Set objEventoSerieOriginal = New AdmEvento
        Set objEventoTransportadora = New AdmEvento
        Set objEventoNFiscalOriginal = New AdmEvento
        Set objEventoContaContabil = New AdmEvento
        Set objEventoProdutoBenef = New AdmEvento
        Set objEventoEstoqueBenef = New AdmEvento
        Set objEventoCcl = New AdmEvento
        Set objEventoContrato = New AdmEvento
        Set objEventoChvNFe = New AdmEvento
        
        '########################################
        'Inserido por Wagner
        Set objEventoItemContrato = New AdmEvento
        
        Set gcolcolMedicoes = New Collection
        Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
        Set gobjTelaProjetoInfo.objUserControl = objUserControl
        Set gobjTelaProjetoInfo.objTela = Me
        '########################################
        
        Set objGridItens = New AdmGrid
        Set objGridParcelas = New AdmGrid
    
        'Rastreamento
'        Set gcolItensNF = New Collection
'        Set objGenerico = New AdmGenerico
'        objGenerico.vVariavel = HABILITA_TELA
    
        lErro = Inicializa_Grid_Itens(objGridItens)
        If lErro <> SUCESSO Then gError 42478
    
        lErro = CF2(Me, "NFFE_Inicializa_Grid_Parcelas", objGridParcelas)
        If lErro <> SUCESSO Then gError 42508
        
'        If giTipoVersao = VERSAO_FULL Then
            Set objGridBenef = New AdmGrid
            'Inicializa a Máscara de Produto
            lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoBenef)
            If lErro <> SUCESSO Then gError 52504
                        
            'Inicializa mascara de contaContabilEst
            lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilEst)
            If lErro <> SUCESSO Then gError 52505
            
            'Inicializa mascara de contaContabilEst
            lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilProducao)
            If lErro <> SUCESSO Then gError 52506
            
            'Faz a inicialização do grid de Movimentos
            lErro = Inicializa_GridMovimentos(objGridBenef)
            If lErro <> SUCESSO Then gError 52507
 '       End If
        
    '    'Seleciona a única tabela de preço. Importante: GridItens já deve estar definido
    '    If giTipoVersao = VERSAO_LIGHT And TabelaPreco.ListCount > 0 Then
    '        TabelaPreco.ListIndex = 0
    '    End If
    
        'Inicializa a contabilidade carregando apenas os modelos contábeis referentes ao TipoDocInfo default
        If objContabil.objGridInt Is Nothing Then

            'Obtém o código do subtipo contábil referente ao TipoDocInfo default
            lErro = CF("Obtem_Subtipo_Contabil", "TiposDocInfo", DOCINFO_NFEEFN, iSubTipo)
            If lErro <> SUCESSO And lErro <> 79671 Then gError 79698
            
            'Se não encontrou o subtipo => erro
            If lErro = 79671 Then gError 79699
            
            'Inicialização da parte de contabilidade
            lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_ESTOQUE, iSubTipo)
            If lErro <> SUCESSO Then gError 79346

        End If

        
        'Inicializa o grid de Rastreamento
        lErro = gobjRastreamento.Inicializa_Grid_Rastreamento()
        If lErro <> SUCESSO Then gError 83265

        'Inicializa o grid de Distribuicao
        lErro = gobjDistribuicao.Inicializa_GridDist()
        If lErro <> SUCESSO Then gError 89527

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

        Case 39656, 42471 To 42478, 42508, 42669, 49394, 52504, 52505, 52506, 52507, 79346, 79698, 83265, 89527, 95162, 95163, 102460, 102461
        
        Case 79699
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPODOCINFO_NAO_ENCONTRADO", gErr, DOCINFO_NFEEFN)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156873)
     
    End Select
     
    iAlterado = 0
    
    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    TAB_BENEFICIAMENTO = 9

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
    
    DataEntrada.PromptInclude = False
    DataEntrada.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEntrada.PromptInclude = True

    Set gobjNFiscal = New ClassNFiscal
    
    Call CF2(Me, "NFEntC_Form_Load")
    
    MsgAutomatica.Value = vbChecked
    Mensagem.MaxLength = MENSAGEM_NF_MAXLENGTH
    MensagemCorpo.MaxLength = MENSAGEM_NF_MAXLENGTH
    
    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207467
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 207467

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156874)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_MascaraCcl() As Long
'Inicializa a mascara do centro de custo

Dim sMascaraCcl As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_MascaraCcl

    sMascaraCcl = String(STRING_CCL, 0)

    'le a mascara dos centros de custo/lucro
    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then Error 49395

    Ccl.Mask = sMascaraCcl

    Inicializa_MascaraCcl = SUCESSO

    Exit Function

Erro_Inicializa_MascaraCcl:

    Inicializa_MascaraCcl = Err

    Select Case Err

        Case 49395

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156875)

    End Select

    Exit Function

End Function

Private Function Carrega_TiposDocInfo() As Long
'Carrega na os Tipo de Documentos relacionados com a tela

Dim lErro As Long
Dim colTipoDocInfo As New colTipoDocInfo
Dim objTipoDocInfo As New ClassTipoDocInfo
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
                If objTipoDocInfo.iPadrao = Padrao Then TipoNFiscal.ListIndex = TipoNFiscal.NewIndex
            End If
        Next

    'Se for versão LIGHT
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'        'Altera NomeReduzido de tipos c\ o mesmo nome red
'        colTipoDocInfo(CStr(DOCINFO_NFEEN)).sNomeReduzido = "Externa - Normal (NFiscal)"
'        colTipoDocInfo(CStr(DOCINFO_NFIEN)).sNomeReduzido = "Interna - Normal (NFiscal)"
'        colTipoDocInfo(CStr(DOCINFO_NFEEFN)).sNomeReduzido = "Externa - Normal (NFiscalFatura)"
'        colTipoDocInfo(CStr(DOCINFO_NFIEFN)).sNomeReduzido = "Interna - Normal (NFiscalFatura)"
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
'                    If objTipoDocInfo.iCodigo = DOCINFO_NFEEFN Then
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

    Carrega_TiposDocInfo = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156876)

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
    If lErro <> SUCESSO Then Error 42480

    'Carrega na combo
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

    Carrega_Serie = Err

    Select Case Err

        Case 42480

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156877)

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

    Carrega_PlacaUF = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156878)

    End Select

End Function

Private Function Carrega_Transportadora() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_Transportadora

    'Lê o Código e o NOme de Todas as Transportadoras do BD
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then Error 42482

    'Carrega a combo de Transportadoras
    For Each objCodigoNome In colCodigoNome
        Transportadora.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        Transportadora.ItemData(Transportadora.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_Transportadora = SUCESSO

    Exit Function

Erro_Carrega_Transportadora:

    Carrega_Transportadora = Err

    Select Case Err

        Case 42482

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156879)

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
'    lErro = CF("CondicoesPagto_Le_Pagamento", colCod_DescReduzida)
'    If lErro <> SUCESSO Then Error 42483
'
'    For Each objCod_DescReduzida In colCod_DescReduzida
'
'        'Adiciona novo ítem na List da Combo CondicaoPagamento
'        If gobjCRFAT.iCondPagtoSemCodigo = 0 Then
'            CondicaoPagamento.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
'        Else
'            CondicaoPagamento.AddItem objCod_DescReduzida.sNome
'        End If
'
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
'    Carrega_CondicaoPagamento = Err
'
'    Select Case Err
'
'        Case 42483
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156880)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Carrega_TipoCobranca() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoCobranca

    'Lê o código e a descrição de todos os Tipos de Cobrança
    lErro = CF("Cod_Nomes_Le", "TiposDeCobranca", "Codigo", "Descricao", STRING_TIPOSDECOBRANCA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 42489

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o ítem na List da Combo TipoCobranca
        TipoCobranca.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TipoCobranca.ItemData(TipoCobranca.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TipoCobranca = SUCESSO

    Exit Function

Erro_Carrega_TipoCobranca:

    Carrega_TipoCobranca = Err

    Select Case Err

        Case 42489

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156881)

    End Select

    Exit Function

End Function

Sub GridItens_InicializaColunas()
    Call CF2(Me, "NFFE_GridItens_InicializaColunas")
End Sub

Function NFFE_GridItens_InicializaColunas(ByVal objCT As Object) As Long
    
    'Colunas da Grid
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_Quantidade_Col = 4
'distribuicao
'    iGrid_Almoxarifado_Col = 5
    iGrid_Ccl_Col = 5
    iGrid_ValorUnitario_Col = 6
    iGrid_PercDesc_Col = 7
    iGrid_Desconto_Col = 8
    iGrid_PrecoTotal_Col = 9
    iGrid_PrecoTotalB_Col = 10
    
    '#########################
    'Inserido por Wagner
    iGrid_Contrato_Col = 11
    iGrid_ItemContrato_Col = 12
    '#########################
    iGrid_DataCobranca_Col = 13
    
    NFFE_GridItens_InicializaColunas = SUCESSO
    
End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
    Inicializa_Grid_Itens = CF2(Me, "NFFE_Inicializa_Grid_Itens", objGridInt)
End Function

Public Function NFFE_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
'distribuicao
'    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Valor Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")

    '##############################################
    'Inserido por Wagner
    objGridInt.colColuna.Add ("Contrato")
    objGridInt.colColuna.Add ("Item")
    '##############################################

    objGridInt.colColuna.Add ("Data Cobrança")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
'distribuicao
'    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (ValorUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (ValorTotal.Name)
    objGridInt.colCampo.Add (PrecoTotalB.Name)

    '##############################################
    'Inserido por Wagner
    objGridInt.colCampo.Add (Contrato.Name)
    objGridInt.colCampo.Add (Item.Name)
    '##############################################
    objGridInt.colCampo.Add (DataCobranca.Name)
    
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

    NFFE_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Function NFFE_Inicializa_Grid_Parcelas(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Vencto Real")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Cobrança")
    objGridInt.colColuna.Add ("Banco Cobrador")
    objGridInt.colColuna.Add ("Portador")
    objGridInt.colColuna.Add ("Suspenso")
    objGridInt.colColuna.Add ("Código de Barras")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (DataVencimentoReal.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (TipoCobranca.Name)
    objGridInt.colCampo.Add (comboCobrador.Name)
    objGridInt.colCampo.Add (comboPortador.Name)
    objGridInt.colCampo.Add (Suspenso.Name)
    objGridInt.colCampo.Add (CodigodeBarras.Name)
    
    'Colunas do Grid
    iGrid_Parcela_Col = 0
    iGrid_Vencimento_Col = 1
    iGrid_VenctoReal_Col = 2
    iGrid_Valor_Col = 3
    iGrid_Cobranca_Col = 4
    iGrid_Banco_Col = 5
    iGrid_Portador_Col = 6
    iGrid_Suspenso_Col = 7
    iGrid_CodigoDeBarras_Col = 8

    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 900

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NFFE_Inicializa_Grid_Parcelas = SUCESSO

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
    sTabela = "NFiscalFatEntrada"

    'Lê os dados da Tela NFiscal
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then Error 42484

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objNFiscal.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "TipoNFiscal", objNFiscal.iTipoNFiscal, 0, "TipoNFiscal"
    colCampoValor.Add "NaturezaOp", objNFiscal.sNaturezaOp, STRING_BUFFER_MAX_TEXTO, "NaturezaOp"
    colCampoValor.Add "Serie", objNFiscal.sSerie, STRING_BUFFER_MAX_TEXTO, "Serie"
    colCampoValor.Add "NumNotaFiscal", objNFiscal.lNumNotaFiscal, 0, "NumNotaFiscal"
    colCampoValor.Add "DataEmissao", objNFiscal.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "DataEntrada", objNFiscal.dtDataEntrada, 0, "DataEntrada"
'horaentrada
    colCampoValor.Add "HoraEntrada", CDbl(objNFiscal.dtHoraEntrada), 0, "HoraEntrada"
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
    colCampoValor.Add "Fornecedor", objNFiscal.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "FilialForn", objNFiscal.iFilialForn, 0, "FilialForn"
    colCampoValor.Add "FilialEntrega", objNFiscal.iFilialEntrega, 0, "FilialEntrega"
    colCampoValor.Add "NumPedidoTerc", objNFiscal.sNumPedidoTerc, STRING_BUFFER_MAX_TEXTO, "NumPedidoTerc"
    colCampoValor.Add "NumPedidoVenda", objNFiscal.lNumPedidoVenda, 0, "NumPedidoVenda"
    colCampoValor.Add "CodTransportadora", objNFiscal.iCodTransportadora, 0, "CodTransportadora"
    colCampoValor.Add "TabelaPreco", objNFiscal.iTabelaPreco, 0, "TabelaPreco"
    colCampoValor.Add "NumIntDocCPR", objNFiscal.lNumIntDocCPR, 0, "NumIntDocCPR"
    colCampoValor.Add "FornecedorBenef", objNFiscal.lFornecedorBenef, 0, "FornecedorBenef"
    colCampoValor.Add "FilialFornBenef", objNFiscal.iFilialFornBenef, 0, "FilialFornBenef"
    colCampoValor.Add "FreteRespons", objNFiscal.iFreteRespons, 0, "FreteRespons"
    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_BUFFER_MAX_TEXTO, "VolumeNumero"
    colCampoValor.Add "Observacao", objNFiscal.sObservacao, STRING_NFISCAL_OBSERVACAO, "Observacao"
    colCampoValor.Add "Status", objNFiscal.iStatus, 0, "Status"
    
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
'    colSelecao.Add "Status", OP_DIFERENTE, STATUS_CANCELADO

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 42484

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156882)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Tela_Preenche

    objNFiscal.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objNFiscal.lNumIntDoc <> 0 Then

        'Carrega objNFiscal com os dados passados em colCampoValor
        objNFiscal.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
        objNFiscal.iTipoNFiscal = colCampoValor.Item("TipoNFiscal").vValor
        objNFiscal.sNaturezaOp = colCampoValor.Item("NaturezaOP").vValor
        objNFiscal.lFornecedor = colCampoValor.Item("Fornecedor").vValor
        objNFiscal.iFilialForn = colCampoValor.Item("FilialForn").vValor
        objNFiscal.sSerie = colCampoValor.Item("Serie").vValor
        objNFiscal.lNumNotaFiscal = colCampoValor.Item("NumNotaFiscal").vValor
        objNFiscal.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
        objNFiscal.dtDataEntrada = colCampoValor.Item("DataEntrada").vValor
'horaentrada
        objNFiscal.dtHoraEntrada = colCampoValor.Item("HoraEntrada").vValor
        objNFiscal.dValorFrete = colCampoValor.Item("ValorFrete").vValor
        objNFiscal.dValorSeguro = colCampoValor.Item("ValorSeguro").vValor
        objNFiscal.dValorOutrasDespesas = colCampoValor.Item("ValorOutrasDespesas").vValor
        objNFiscal.dValorDesconto = colCampoValor.Item("ValorDesconto").vValor
        objNFiscal.dValorTotal = colCampoValor.Item("ValorTotal").vValor
        objNFiscal.iCodTransportadora = colCampoValor.Item("CodTransportadora").vValor
        objNFiscal.sPlaca = colCampoValor.Item("Placa").vValor
        objNFiscal.sPlacaUF = colCampoValor.Item("PlacaUF").vValor
        objNFiscal.lVolumeQuant = colCampoValor.Item("VolumeQuant").vValor
        objNFiscal.lVolumeEspecie = colCampoValor.Item("VolumeEspecie").vValor
        objNFiscal.lVolumeMarca = colCampoValor.Item("VolumeMarca").vValor
        objNFiscal.sMensagemNota = colCampoValor.Item("MensagemNota").vValor
        objNFiscal.dPesoLiq = colCampoValor.Item("PesoLiq").vValor
        objNFiscal.dPesoBruto = colCampoValor.Item("PesoBruto").vValor
        objNFiscal.lNumIntNotaOriginal = colCampoValor.Item("NumIntNotaOriginal").vValor
        objNFiscal.dValorProdutos = colCampoValor.Item("ValorProdutos").vValor
        objNFiscal.iCanal = colCampoValor.Item("Canal").vValor
        objNFiscal.dValorProdutos = colCampoValor.Item("ValorProdutos").vValor
        objNFiscal.iFilialEntrega = colCampoValor.Item("FilialEntrega").vValor
        objNFiscal.sNumPedidoTerc = colCampoValor.Item("NumPedidoTerc").vValor
        objNFiscal.lNumPedidoVenda = colCampoValor.Item("NumPedidoVenda").vValor
        objNFiscal.iTabelaPreco = colCampoValor.Item("TabelaPreco").vValor
        objNFiscal.lNumIntDocCPR = colCampoValor.Item("NumIntDocCPR").vValor
        objNFiscal.iFilialEmpresa = giFilialEmpresa
        objNFiscal.lFornecedorBenef = colCampoValor.Item("FornecedorBenef").vValor
        objNFiscal.iFilialFornBenef = colCampoValor.Item("FilialFornBenef").vValor
        objNFiscal.iFreteRespons = colCampoValor.Item("FreteRespons").vValor
        objNFiscal.sVolumeNumero = colCampoValor.Item("VolumeNumero").vValor
        objNFiscal.sObservacao = colCampoValor.Item("Observacao").vValor
        objNFiscal.iStatus = colCampoValor.Item("Status").vValor

        objNFiscal.iFilialEmpresa = giFilialEmpresa

        'Coloca os dados da NFiscal na tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then Error 42485

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 42485

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156883)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    Call CF2(Me, "NFEnt_Form_Unload", Cancel)
    Set gobjInfoUsu = Nothing
    
    Set gobjNFiscal = Nothing
    Set objEventoSerie = Nothing
    Set objEventoProduto = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoNatureza = Nothing
    Set objEventoCondPagto = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoFornecedorBenef = Nothing
    Set objEventoFornNFOrig = Nothing 'Incluído por Luiz em 25/01/02
    Set objEventoRecebimento = Nothing
    Set objEventoSerieOriginal = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoNFiscalOriginal = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoEstoqueBenef = Nothing
    Set objEventoProdutoBenef = Nothing
    Set objEventoCcl = Nothing
    Set objEventoContrato = Nothing
    Set objEventoChvNFe = Nothing
    
    Set objGridBenef = Nothing
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    Set objGridItens = Nothing
    Set objGridParcelas = Nothing

    '########################################
    'Inserido por Wagner
    Set objEventoItemContrato = Nothing
    
    Set gcolcolMedicoes = Nothing
    Set gobjTelaProjetoInfo = Nothing
    Set gobjInfoAdic = Nothing
    '########################################

    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

'Rastreamento
'    Set gcolItensNF = Nothing
'    Set objGenerico = Nothing


    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

    Set gobjAnotacao = Nothing
    
''*** incluidos p/tratamento de tributacao *******************************
    If Not (gobjTribTab Is Nothing) Then Call gobjTribTab.Desativar
    Set gobjTribTab = Nothing
'*** fim tributacao

    'Rastreamento
    Set gobjRastreamento = Nothing

    'distribuicao
    Set gobjDistribuicao = Nothing

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156884)
    
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

    Exit Sub

End Sub

Private Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim objNaturezaOp As ClassNaturezaOp

    Set objNaturezaOp = obj1

    'Coloca a natureza Operação na tela
    NaturezaOp.Text = objNaturezaOp.sCodigo

    Me.Show

    Call NaturezaOp_Validate(bSGECancelDummy)
    
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
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    Fornecedor.Text = objFornecedor.lCodigo
    Call Fornecedor_Validate(bCancel)

    Me.Show

End Sub

Private Sub objEventoFornecedorBenef_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    FornecedorBenef.Text = objFornecedor.lCodigo
    Call FornecedorBenef_Validate(bCancel)

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

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As Collection
Dim lErro As Long

On Error GoTo Erro_NFiscalLabel_Click

    'Recolhe os dados da NFiscal que estão na tela
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then Error 42486

    'Chama a tela de Brow sw Nota Fiscal de Entrada
    Call Chama_Tela("NFiscalFatEntradaLista", colSelecao, objNFiscal, objEventoNFiscal)

    Exit Sub

Erro_NFiscalLabel_Click:

    Select Case Err

        Case 42486

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156885)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim objNFiscal As ClassNFiscal
Dim lErro As Long

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objNFiscal = obj1

    'Traz a NotaFiscal de entrada para a a tela
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then Error 42487

    'Fecha o Comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case Err

        Case 42487

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156886)

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

Public Sub SerieOriginalLabel_Click()

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
Dim lErro As Long
Dim sTela As String
Dim iTipoDocInfo As Integer
Dim sFiltro As String

On Error GoTo Erro_NFiscalOriginalLabel_Click

    'Recolhe os dados da Nota Fiscal Entrada
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 42490

    'Guarda a Serie e o Número da Nota Fiscal Original da Tela
    objNFiscal.sSerie = SerieNFiscalOriginal.Text
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscalOriginal.Text)
    
    iTipoDocInfo = Codigo_Extrai(Trim(TipoNFiscal.Text))
    
    If iTipoDocInfo = DOCINFO_FDDAI Then
            
        'Adiciona à coleção o filtro de TipoNFiscal
        colSelecao.Add DOCINFO_NFIEIMP
        colSelecao.Add DOCINFO_NFIEIMPSE
        colSelecao.Add DOCINFO_NFIEIMPPC
        
        'Guarda o nome do browser que será chamado
        'sTela = "NFiscalTipoDocInfoLista"
        sTela = "NFiscalInternaEntradaLista"
        sFiltro = "TipoNFiscal IN (?,?,?)"
        
    Else
        
        'Guarda o nome do browser que será chamado
        sTela = "NFiscalEntradaTodasLista"
        
    End If
    
    'Chama a Tela NFiscalFatEntradaNormalLista
    Call Chama_Tela(sTela, colSelecao, objNFiscal, objEventoNFiscalOriginal, sFiltro)

    Exit Sub

Erro_NFiscalOriginalLabel_Click:

    Select Case gErr

        Case 42490

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156887)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNFiscalOriginal_evSelecao(obj1 As Object)

Dim objNFiscal As ClassNFiscal
Dim bCancel As Boolean

    Set objNFiscal = obj1

    'Preenche a Série e o Número da Nota Fiscal Original
    SerieNFiscalOriginal.Text = objNFiscal.sSerie
    NFiscalOriginal.Text = objNFiscal.lNumNotaFiscal

    If (Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_FDDAI) Or (Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_NFSEFT) Then

        'Preenche o Fornecedor da NF Orig
        FornNFOrig.Text = objNFiscal.lFornecedor
        Call FornNFOrig_Validate(bCancel)
        
        'Preenche a Filial do Fornecedor
        FilialFornNFOrig.Text = objNFiscal.iFilialForn
        Call FilialFornNFOrig_Validate(bCancel)

    End If

    Me.Show

    Exit Sub

End Sub

Public Sub CondPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection
Dim iCodigo As Integer
Dim lErro As Long

On Error GoTo Erro_CondPagtoLabel_Click

    'Se Condição de Pagto estiver preenchida, extrai o código
    objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)


    'Chama a Tela CondicoesPagamentoCPLista
    Call Chama_Tela("CondicaoPagtoCPLista", colSelecao, objCondicaoPagto, objEventoCondPagto)

    Exit Sub

Erro_CondPagtoLabel_Click:

    Select Case gErr

        Case 126963

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 176321)

     End Select

     Exit Sub

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

    Select Case Err

        Case 42488

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156888)

     End Select

     Exit Sub

End Sub

Public Sub BotaoRecebimentos_Click()

Dim colSelecao As Collection
Dim objNFiscal As New ClassNFiscal
Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoRecebimentos_Click

    'Se o Recebimento estiver preenchido
    If Len(Trim(NumRecebimento.Text)) > 0 Then
        objNFiscal.lNumRecebimento = CLng(NumRecebimento.Text)
    Else
        objNFiscal.lNumRecebimento = 0
    End If
    
    'Verifica preenchimento de Fornecedor
    If Len(Trim(Fornecedor.Text)) <> 0 Then

        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 89313

        'Se não achou o Fornecedor --> erro
        If lErro = 6681 Then gError 89314

        objNFiscal.lFornecedor = objFornecedor.lCodigo

    End If

    If Len(Trim(Filial.Text)) <> 0 Then objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa

    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)

    If Len(Trim(NFiscal.Text)) <> 0 Then objNFiscal.lNumNotaFiscal = CLng(NFiscal.Text)

    objNFiscal.dtDataEntrada = MaskedParaDate(DataEntrada)

    'Chama a tela de browse RecebMaterialFLista
    Call Chama_Tela("RecebMaterialFLista", colSelecao, objNFiscal, objEventoRecebimento)
    
    Exit Sub

Erro_BotaoRecebimentos_Click:

    Select Case gErr
    
        Case 89313

        Case 89314
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, Fornecedor.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156889)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoRecebimento_evSelecao(obj1 As Object)

Dim objNFiscal As ClassNFiscal
Dim lErro As Long

On Error GoTo Erro_objEventoRecebimento_evSelecao

    Set objNFiscal = obj1

    If objNFiscal.iTipoNFiscal = DOCINFO_NRFF Then
        objNFiscal.iTipoNFiscal = DOCINFO_NFEEFN
    ElseIf objNFiscal.iTipoNFiscal = DOCINFO_NRFP Then
        objNFiscal.iTipoNFiscal = DOCINFO_NFIEFN
    End If

    lErro = Preenche_Tela_Recebimento(objNFiscal)
    If lErro <> SUCESSO Then Error 42491

    'Fecha o Comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Call gobjTribTab.TributacaoRecalcular_Click
    
    Call objContabil.Contabil_Carga_Modelo_Padrao
    
    Exit Sub

Erro_objEventoRecebimento_evSelecao:

    Select Case Err

        Case 42491

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156890)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLocalizacaoDist_Click()
'distribuicao

    Call gobjDistribuicao.BotaoLocalizacaoDist_Click

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
        If GridItens.Row = 0 Then gError 43717

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 83081
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 43717
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 83081
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156891)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoEnxuto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    'Verifica se o Produto está preenchido
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then

        Set objProduto = obj1

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 42494

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        If Not (Me.ActiveControl Is Produto) Then

            'Coloca o produto, a Descrição e a Unidade de Medida da tela
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
    
            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then 'Alterado por Wagner
                GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""
                Error 42495
                
            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case Err

        Case 42494
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, objProduto.sCodigo)
        
        Case 42495

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156892)

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
        If lErro <> SUCESSO Then Error 64012
            
        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then Error 42496
        
        If lErro <> SUCESSO Then Error 42497

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

        'Lê o Tipo de Documento com o Código de objNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then Error 42498
        
        'Se não encontrar --> Erro
        If lErro = 31415 Then Error 42499
        
'        If giTipoVersao = VERSAO_FULL Then
            If objTipoDocInfo.sNomeTelaNFiscal <> Me.Name Then Error 42286
'        End If
        
        'Traz a nota para a tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then Error 42404

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 42404, 42496, 42498, 64012

        Case 42286
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_DIFERENTE_NF_FATURA_ENTRADA", Err, objNFiscal.iTipoNFiscal)

        Case 42497
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", Err, objNFiscal.lNumIntDoc)
            Call Limpa_Tela_NFiscal
            iAlterado = 0

        Case 42499
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objNFiscal.iTipoNFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156893)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub TipoNFiscal_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TipoNFiscal_Click

    If TipoNFiscal.ListIndex = -1 Then Exit Sub

    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 42520

    'se não estiver cadastrado ==> erro
    If lErro = 31415 Then gError 42521

    gobjRastreamento.iCodigo = objTipoDocInfo.iCodigo
    
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then
        giComplemento = MARCADO
    Else
        giComplemento = DESMARCADO
    End If
    
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
        'ChvNFe.Enabled = True
        BotaoLimparNF.Visible = False
        'Cor Vermelha para label pois é obrigatorio
        NFiscalLabel.ForeColor = &H80&
    End If
    
    If NFiscalInterna.Visible Then
        'Seleciona série padrão
        lErro = Seleciona_Serie_Padrao()
        If lErro <> SUCESSO Then gError 42520
    End If

    'habilita/desabilita os campos relativos a fornecedor associados a nota fiscal original
    Call Habilita_Desabilita_Forn_NFOrig
    
    iEmitenteNF = objTipoDocInfo.iEmitente
    
    If iEmitenteNF <> DOCINFO_EMPRESA Then
    
        NaturezaOp.Text = NaturezaOp_Conv34(objTipoDocInfo.sNatOpExtPadrao, IIf(MaskedParaDate(DataEntrada) = DATA_NULA, MaskedParaDate(DataEmissao), MaskedParaDate(DataEntrada)))
        
    End If
    
    If TipoNFiscal.ListIndex <> -1 Then
        If (TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFEEFBEN) Or (TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFIEFBEN) Then
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
    If lErro <> SUCESSO Then gError 79347

    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 79347
    End If
        
    Exit Sub

Erro_TipoNFiscal_Click:

    Select Case gErr

        Case 42520, 79347

        Case 42521
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156894)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 42505
    
    'Se não conseguir --> Erro
    If lErro <> SUCESSO Then Error 42506

    'tenta ler a natureza de operacao
    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
    
    'Lê o Tipo da NF
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 55550

    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then Error 55551

    'Se a Natureza de Operação ainda não foi preenchida
    If Len(Trim(NaturezaOp.ClipText)) = 0 Then

        NaturezaOp.Text = NaturezaOp_Conv34(objTipoDocInfo.sNaturezaOperacaoPadrao, IIf(MaskedParaDate(DataEntrada) = DATA_NULA, MaskedParaDate(DataEmissao), MaskedParaDate(DataEntrada)))

    End If

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
        If lErro <> SUCESSO Then Error 55550
    End If
    
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then Error 55550
    End If

    Exit Sub

Erro_TipoNFiscal_Validate:

    Cancel = True

    Select Case Err

        Case 42505, 55550

        Case 42506
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, TipoNFiscal.Text)

        Case 55551
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156895)

    End Select

    Exit Sub

End Sub

Public Sub NaturezaOp_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNaturezaOp As New ClassNaturezaOp
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NaturezaOp_Validate

    If iNaturezaOpAlterada = 0 Then Exit Sub
    
    If Len(Trim(NaturezaOp.ClipText)) = 0 Then Exit Sub

    objNaturezaOp.sCodigo = NaturezaOp.Text

    lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
    If lErro <> SUCESSO And lErro <> 17958 Then Error 42522

    If lErro = 17958 Then Error 42523

    'se for nota interna garantir que tenha codigo < 500
    If iEmitenteNF = DOCINFO_EMPRESA Then
    
        If objNaturezaOp.sCodigo >= NATUREZA_SAIDA_COD_INICIAL Then Error 56982

        If (gobjTribTab.sNatOpInterna <> objNaturezaOp.sCodigo) Then
        
            gobjTribTab.sNatOpInterna = objNaturezaOp.sCodigo
            Call NatOpInterna_Validate(Cancel)
            
        End If

    End If
    
    'se for nota externa garantir que tenha codigo > 500
    If iEmitenteNF <> DOCINFO_EMPRESA And objNaturezaOp.sCodigo <= NATUREZA_SAIDA_COD_INICIAL Then Error 56983

    iNaturezaOpAlterada = 0
    
    Exit Sub

Erro_NaturezaOp_Validate:

    Cancel = True

    Select Case Err

        Case 42522

        Case 56982
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_ENTRADA", Err)
        
        Case 56983
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", Err)
        
        Case 42523
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_NATUREZA_OPERACAO", NaturezaOp.Text)
            If vbMsgRes = vbYes Then

                Call Chama_Tela("NaturezaOperacao", objNaturezaOp)

            Else
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156896)

    End Select

    Exit Sub

End Sub

Public Sub NatOpInterna_Validate(Cancel As Boolean)
    Call gobjTribTab.NatOpInterna_Validate(Cancel)
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
Dim bCancel As Boolean

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le3(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 42302

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 42316

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
            
                If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                
                'Seleciona filial na Combo Filial
                Call CF("Filial_Seleciona", Filial, iCodFilial)
                
            End If
            
            If PagamentoAutomatico.Value = vbChecked Then
                If objFornecedor.iCondicaoPagto > 0 Then
                    CondicaoPagamento.Text = objFornecedor.iCondicaoPagto
                    Call CondicaoPagamento_Validate(bCancel)
                End If
            End If

            ''*** incluidos p/tratamento de tributacao *******************************
            If iCodFilial <> 0 Then Call gobjTribTab.FilialFornecedor_Alterada(objFornecedor.lCodigo, iCodFilial)
            '*** fim tributacao

            If gobjInfoAdic Is Nothing Then Set gobjInfoAdic = New ClassInfoAdic
            gobjInfoAdic.sNatureza = objFornecedor.sNatureza

            '########################################################
            'Inserido por Wagner
            lErro = CF2(Me, "NFCF_Fornecedor_Validate", objFornecedor.lCodigo, iCodFilial)
            If lErro <> SUCESSO Then gError 131951
            '########################################################

        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

            Filial.Clear

            '#############################################
            'Inserido por Wagner
            Call CF2(Me, "NFCF_Limpa_Tela")
            '#############################################

        End If

        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 42302, 42316, 131951 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156897)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.ListIndex <> -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 42500

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 42501

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 42502

        If lErro = 18272 Then Error 42503

        'coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

        Call gobjTribTab.FilialFornecedor_Alterada(objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 42504

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case Err

        Case 42500, 42502

        Case 42501
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 42503
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 42504
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156898)

    End Select

    Exit Sub

End Sub

Public Sub DataEntrada_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEntrada_Validate

    If iDataEntradaAlterada <> 0 Then
        
        'Verifica se a data de Entrada foi digitada
        If Len(Trim(DataEntrada.ClipText)) = 0 Then Exit Sub
    
        'Critica a data digitada
        lErro = Data_Critica(DataEntrada.Text)
        If lErro <> SUCESSO Then gError 42507

''*** incluidos p/tratamento de tributacao *******************************
        Call gobjTribTab.DataEntrada_Alterada(StrParaDate(DataEntrada.Text))
'*** fim tributacao
        
        iDataEntradaAlterada = 0

    End If
    
    Exit Sub

Erro_DataEntrada_Validate:

    Cancel = True

    Select Case gErr

        Case 42507

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156899)

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
    If lErro <> SUCESSO Then gError 89782

    Exit Sub

Erro_HoraEntrada_Validate:

    Cancel = True

    Select Case gErr

        Case 89782

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156900)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntrada_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntrada_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEntrada, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 42508

    Exit Sub

Erro_UpDownEntrada_DownClick:

    Select Case Err

        Case 42508

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156901)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntrada_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntrada_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEntrada, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 42509

    Exit Sub

Erro_UpDownEntrada_UpClick:

    Select Case Err

        Case 42509

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156902)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCondPagamento As New ClassCondicaoPagto

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a data de emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) <> 0 Then

        'Critica a data digitada
        lErro = Data_Critica(DataEmissao.Text)
        If lErro <> SUCESSO Then gError 42510
    
        Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
        
        Call Cobranca_Automatica
    
    End If
        
    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 42510

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156903)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 42511

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case Err

        Case 42511

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156904)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 42512

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case Err

        Case 42512

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156905)

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
        If lErro <> SUCESSO And lErro <> 12253 Then Error 42528
        
        'Se a serie não está cadastrada --> verifica se é uma nota interna, se for ==> erro, pois a serie não está cadastrada
        If lErro <> SUCESSO Then
    
            'Se o tipo está preenchido
            If Len(Trim(TipoNFiscal.Text)) > 0 Then
        
                objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
                
                'Lê o Tipo de Documento
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO Then Error 42529
        
                'se a nota interna e a serie não está cadastrada ==> erro
                If objTipoDocInfo.iTipo = DOCINFO_NF_INT_ENTRADA Then Error 42530
        
            End If
        
            'Verifica se tamanho da série é maior do que o espaço no bd ==> erro
            If Len(Trim(Serie.Text)) > STRING_SERIE Then Error 42531
    
        End If
    
    End If
    
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.Serie_Alterada(Serie.Text)
'*** fim tributacao
        
    iSerieAlterada = 0
    
    Exit Sub

Erro_Serie_Validate:

    Cancel = True


    Select Case Err

        Case 42528, 42529

        Case 42530
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, Serie.Text)

        Case 42531
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_SERIE", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156906)

    End Select

    Exit Sub

End Sub

Public Sub Recebimento_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Recebimento_Click
        
    'Se o Número do Recebimento estiver preenchido
    If Len(Trim(NumRecebimento.Text)) > 0 Then
        
        objNFiscal.lNumRecebimento = CLng(NumRecebimento.Text)
        
        'Le o Recebimento
        lErro = CF("NFiscal_Le_Recebimento", objNFiscal)
        If lErro <> SUCESSO And lErro <> 61050 Then gError 61051
        
        'Se não encontrou --> Erro
        If lErro = 61050 Then gError 61052
        
        If objNFiscal.iTipoNFiscal = DOCINFO_NRCP Or objNFiscal.iTipoNFiscal = DOCINFO_NRCC Then gError 61120
                
        If objNFiscal.iStatus = STATUS_BAIXADO Then gError 64169
                
        'Lê o Tipo de documento
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 27623 Then gError 75432
        
        'Se o Tipo de documento é de compras
        If objTipoDocInfo.sNomeTelaNFiscal = RECEBIMENTO_MATERIAL_FCOM Or objTipoDocInfo.sNomeTelaNFiscal = NOTA_FISCAL_COMPRAS Or objTipoDocInfo.sNomeTelaNFiscal = NOTA_FISCAL_FATURA_COMPRAS Then gError 75433
                
    Else
        
        'Verifica se o tipo está Preenchido
        If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 42533
        
        objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
    
        'Lê o Tipo de documento
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 27623 Then gError 42542
        
        'Se não encontrar --> erro
        If lErro = 27623 Then gError 42544
        
        'Se o Tipo de documento é de compras
        If objTipoDocInfo.sNomeTelaNFiscal = RECEBIMENTO_MATERIAL_FCOM Or objTipoDocInfo.sNomeTelaNFiscal = NOTA_FISCAL_COMPRAS Or objTipoDocInfo.sNomeTelaNFiscal = NOTA_FISCAL_FATURA_COMPRAS Then gError 75434
        
        objNFiscal.iTipoDocInfo = objTipoDocInfo.iTipo
        
        'Se for Nota Fiscal Interna então o NumRecebimento tem que está Preenchido
        If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then gError 61053
                            
        'Verifica se os campos essenciais estão preenchidos
        If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 42534
        If Len(Trim(Filial.Text)) = 0 Then gError 42535
        If Len(Trim(Serie.Text)) = 0 Then gError 42536
        If Len(Trim(NFiscal.ClipText)) = 0 Then gError 42537
        
        'Recolhe os dados da Tela
        lErro = Move_NF_Memoria(objNFiscal)
        If lErro <> SUCESSO Then gError 42541
        
        If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then gError 42538
            
        If objNFiscal.iTipoNFiscal = DOCINFO_NFEEFN Then
            objNFiscal.iTipoNFiscal = DOCINFO_NRFF
        End If
        
        'Lê o Recebimento com estas características
        lErro = CF("NFiscal_Le_Num_Forn1", objNFiscal)
        If lErro <> SUCESSO And lErro <> 55885 Then gError 42539
    
        'Se não encontrar -->Erro
        If lErro = 55885 Then gError 42543
        
        If objNFiscal.iStatus = STATUS_BAIXADO Then gError 64170
        
    End If
    
    'Se encontrar traz para a tela
    lErro = Preenche_Tela_Recebimento2(objNFiscal)
    If lErro <> SUCESSO Then gError 42540
        
    Call gobjTribTab.TributacaoRecalcular_Click
        
    Call objContabil.Contabil_Carga_Modelo_Padrao
    
    Exit Sub
    
Erro_Recebimento_Click:

    Select Case gErr
        
        Case 42533
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 42534
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 42535
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_CHQPRE_NAO_PREENCHIDA", gErr)

        Case 42536
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 42537
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)

        Case 42538
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_NORMAL", gErr)

        Case 42539, 42540, 42541, 42542, 61051, 75432

        Case 42543
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_MATERIAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 42544
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO2", gErr, objTipoDocInfo.iCodigo)
        
        Case 61052
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NAO_CADASTRADO", gErr, objNFiscal.lNumRecebimento)
                
        Case 61053
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NFINT_NAO_PREENCHIDO", gErr)
        
        Case 61120
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_CLIENTE", gErr, objNFiscal.lNumRecebimento)
        
        Case 64169, 64170
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_VINCULADO_NF", gErr, objNFiscal.lNumRecebimento, objNFiscal.lNumNotaFiscal)

        Case 75433, 75434
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_RECEBIMENTO_COMPRAS", gErr, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156907)

    End Select

    Exit Sub

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim sProdutoFormatadoBenef As String
Dim iProdutoPreenchidoBenef As Integer
Dim sUnidadeMed  As String
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
    
        'Produto
        Case Produto.Name
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Produto_Col))) > 0 Then
                Produto.Enabled = False
            Else
                Produto.Enabled = True
            End If
            
        'Unidade de Medida
        Case UnidadeMed.Name

            UnidadeMed.Clear

            'Guarda a UM que está no Grid
            sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

            lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 42545

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                UnidadeMed.Enabled = False
            Else
                UnidadeMed.Enabled = True

                objProduto.sCodigo = sProdutoFormatado
                
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 42546
                
                If lErro = 28030 Then gError 42552 'Não achou

                objClasseUM.iClasse = objProduto.iClasseUM
                
                'Le as Unidades de Medidas da Classe do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 42548
                
                'Carrega a combo de UM
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next
                
                'Seleciona na UM que está preenchida
                UnidadeMed.Text = sUM
                
                If Len(Trim(sUM)) > 0 Then
                    lErro = Combo_Item_Igual(UnidadeMed)
                    If lErro <> SUCESSO And lErro <> 12253 Then gError 42549
                End If
                
            End If

        'Nas demais
'distribuicao
        Case ValorUnitario.Name, PercentDesc.Name, DescricaoItem.Name, Ccl.Name
'        Case ValorUnitario.Name, PercentDesc.Name, Quantidade.Name, Almoxarifado.Name, DescricaoItem.Name, Ccl.Name
        
            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 42550

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case Desconto.Name

            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 42547

            If iProdutoPreenchido = PRODUTO_VAZIO Or Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))) = 0 Then
                objControl.Enabled = False
            ElseIf iProdutoPreenchido = PRODUTO_PREENCHIDO And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_PrecoTotal_Col))) > 0 Then
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
   
        '######################################
        'Inserido por Wagner
        Case Contrato.Name, Item.Name, DataCobranca.Name

            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 136294

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

'        Case Item.Name
'
'            'Verifica se o produto está preenchido
'            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'            If lErro <> SUCESSO Then gError 136295
'
'            If iProdutoPreenchido = PRODUTO_VAZIO Then
'                objControl.Enabled = False
'            Else
'                objControl.Enabled = True
'            End If
        '########################################
        
        Case CodigodeBarras.Name
            
            If iLocalChamada = 3 Then Exit Sub
        
            If Len(GridParcelas.TextMatrix(iLinha, iGrid_CodigoDeBarras_Col)) > 0 Then
        
                If left(GridParcelas.TextMatrix(iLinha, iGrid_CodigoDeBarras_Col), 1) = "8" Then
                    giTeste = 1
                    CodigodeBarras.Mask = "###########.#.###########.#.###########.#.###########.#"
                    CodigodeBarras.PromptInclude = False
                    CodigodeBarras.Text = GridParcelas.TextMatrix(iLinha, iGrid_CodigoDeBarras_Col)
                    CodigodeBarras.PromptInclude = True
                    giTeste = 0
                Else
                    giTeste = 1
                    CodigodeBarras.Mask = "#####.#####.#####.######.#####.######.#.#################"
                    CodigodeBarras.PromptInclude = False
                    CodigodeBarras.Text = GridParcelas.TextMatrix(iLinha, iGrid_CodigoDeBarras_Col)
                    CodigodeBarras.PromptInclude = True
                    giTeste = 0
                End If
        
            End If
        
        Case Else
            '###############################################################
            'Inserido por Wagner 31/07/2006
            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 136295

            lErro = CF2(Me, "NFE_Rotina_Grid_Enable", iLinha, objControl, iLocalChamada, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 124178
            '###############################################################
        
    End Select
    
'    If giTipoVersao = VERSAO_FULL Then
    
        Select Case objControl.Name
        
            Case ProdutoBenef.Name
                lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iLinha, iGrid_ProdutoBenef_Col), sProdutoFormatadoBenef, iProdutoPreenchidoBenef)
                If lErro <> SUCESSO Then gError 52521
                
                If iProdutoPreenchidoBenef = PRODUTO_PREENCHIDO Then
                    objControl.Enabled = False
                Else
                    objControl.Enabled = True
                End If
    
            Case UnidadeMedBenef.Name
                
                lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iLinha, iGrid_ProdutoBenef_Col), sProdutoFormatadoBenef, iProdutoPreenchidoBenef)
                If lErro <> SUCESSO Then gError 52522
               
                If iProdutoPreenchidoBenef <> PRODUTO_PREENCHIDO Then
                    objControl.Enabled = False
                Else
                    objControl.Enabled = True
    
                    objProduto.sCodigo = sProdutoFormatadoBenef
    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 52523
    
                    If lErro = 28030 Then gError 52524
    
                    objClasseUM.iClasse = objProduto.iClasseUM
    
                    'Preenche a List da Combo UnidadeMed com as UM's do Produto
                    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                    If lErro <> SUCESSO Then gError 52525
    
                    'Guardo o valor da Unidade de Medida da Linha
                    sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMedBenef_Col)
    
                    'Limpar as Unidades utilizadas anteriormente
                    UnidadeMedBenef.Clear
    
                    For Each objUM In colSiglas
                        UnidadeMedBenef.AddItem objUM.sSigla
                    Next
    
                    UnidadeMedBenef.AddItem ""
    
                    'Tento selecionar na Combo a Unidade anterior
                    If UnidadeMedBenef.ListCount <> 0 Then
                        For iIndice = 0 To UnidadeMedBenef.ListCount - 1
    
                            If UnidadeMedBenef.List(iIndice) = sUnidadeMed Then
                                UnidadeMedBenef.ListIndex = iIndice
                                Exit For
                            End If
                        Next
                    End If
    
                End If
    
            Case QuantidadeBenef.Name, AlmoxarifadoBenef.Name, ContaContabilEst.Name, ContaContabilProducao.Name
                
                lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iLinha, iGrid_ProdutoBenef_Col), sProdutoFormatadoBenef, iProdutoPreenchidoBenef)
                If lErro <> SUCESSO Then gError 52526
    
                If iProdutoPreenchidoBenef = PRODUTO_PREENCHIDO Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
        
        End Select
    
        'Ratreamento
        lErro = gobjRastreamento.Rotina_Grid_Enable(iLinha, objControl, iLocalChamada)
        If lErro <> SUCESSO Then gError 83268
    
        'distribuicao
        lErro = gobjDistribuicao.Rotina_Grid_Enable_Dist(iLinha, objControl, iLocalChamada)
        If lErro <> SUCESSO Then gError 89532
    
 '   End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 31474, 42545 To 42550, 83268, 89532

        Case 42552, 52524
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 52521, 52522, 52523, 52525, 52526, 124178
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156908)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto do grid que está deixando de ser a corrente
    Saida_Celula_Produto = CF2(Me, "NFFE_Saida_Celula_Produto", objGridInt)
End Function

Function NFFE_Saida_Celula_Produto(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iPossuiGrade As Integer

On Error GoTo Erro_NFFE_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    'Verifica se o Produto esta preenchida
    If Len(Trim(Produto.ClipText)) > 0 Then


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

    NFFE_Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_NFFE_Saida_Celula_Produto:

    NFFE_Saida_Celula_Produto = gErr

    Select Case gErr

        Case 31479, 31481, 35657, 199354
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156909)

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
        If lErro <> SUCESSO Then gError 42558

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
        
         '####################################################################
        'Inserido por Wagner 18/05/2006
        objProduto.sCodigo = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 180040
        '####################################################################
        
        GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)
        
        If gobjNFiscal.ColItensNF(GridItens.Row).iPossuiGrade <> MARCADO And objProduto.iKitVendaComp <> MARCADO Then
        
            'Tenta fazer a distribuição automatica p\ o item de acordo com a nova UM
            lErro = gobjDistribuicao.Distribuicao_Processa()
            If lErro <> SUCESSO Then gError 89512
            
        Else
        
            '####################################################################
            'Inserido por Wagner 18/05/2006
            objKitVenda.sProduto = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
            objKitVenda.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
            objKitVenda.dQuantidade = dQuantidade
        
            lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade)
            If lErro <> SUCESSO Then gError 180041
            '####################################################################

            lErro = gobjDistribuicao.Distribuicao_Processa_Grade()
            If lErro <> SUCESSO Then gError 89512
        
            Call gobjDistribuicao.Atualiza_Grid_Distribuicao(gobjNFiscal.ColItensNF.Item(GridItens.Row))
        End If
        
    End If
    'fim distribuicao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42559

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 55585

    '*** incluidos p/tratamento de tributacao ****************
    If GridItens.Row <= objGridItens.iLinhasExistentes Then
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 42560
    End If
    '*** fim tributacao

    Call ValorTotal_Calcula
        
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 181262
    End If
    
    '#################################################
    'Inserido por Wagner 22/03/2006
    If Abs(dQuantidadeAnterior - dQuantidadeAtual) > QTDE_ESTOQUE_DELTA Then
        lErro = CF2(Me, "NFE_Calcula_PesoKg", GridItens.Row)
        If lErro <> SUCESSO Then gError 181262
    End If
    '#################################################
    
    If GridItens.Row <= objGridItens.iLinhasExistentes Then
    
        If gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade.Count = 0 Then
    
            lErro = gobjRastreamento.Atualiza_Troca_Quantidade(GridItens.Row)
            If lErro <> SUCESSO Then gError 83270
        
        Else
        
            If dQuantidadeAnterior <> 0 Then
                dFator = dQuantidadeAtual \ dQuantidadeAnterior
            Else
                dFator = 1
            End If
            
            lErro = gobjRastreamento.Atualiza_Troca_Quantidade_Grade(GridItens.Row, dFator)
            If lErro <> SUCESSO Then gError 83270
        
        End If
    
    End If
    
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 42558 To 42560, 55585, 83270, 89512, 180040, 180041, 181262
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156910)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

'distribuicao
'Private Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long
'
'Dim lErro As Long
'Dim iProdutoPreenchido As Integer
'Dim sProdutoFormatado As String
'Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim vbMsg As VbMsgBoxResult
'
'On Error GoTo Erro_Saida_Celula_Almoxarifado
'
'    Set objGridInt.objControle = Almoxarifado
'
'    'Se o Almoxarifado está preenchido
'    If Len(Trim(Almoxarifado.Text)) > 0 Then
'
'        lErro = CF("Produto_Formata",GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'        If lErro <> SUCESSO Then gError 42562
'
'        'Valida o ALmoxarifado
'        lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
'        If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 42563
'
'        'Se não for encontrado --> Erro
'        If lErro = 25157 Then gError 42564
'        If lErro = 25162 Then gError 42565
'
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then gError 42561
'
'    If GridItens.Row <= objGridItens.iLinhasExistentes Then
'        lErro = gobjRastreamento.Atualiza_Troca_Almoxarifado(GridItens.Row)
'        If lErro <> SUCESSO Then gError 83271
'    End If
'
'    Saida_Celula_Almoxarifado = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_Almoxarifado:
'
'    Saida_Celula_Almoxarifado = gErr
'
'    Select Case gErr
'
'        Case 42561, 42562, 42563, 83271
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 42564
'            'Pergunta de deseja criar o Almoxarifado
'            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", Almoxarifado.Text)
'            'Se a resposta for sim
'            If vbMsg = vbYes Then
'
'                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
'
'                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text
'
'                'Chama a Tela Almoxarifados
'                Call Chama_Tela("Almoxarifado", objAlmoxarifado)
'
'            Else
'                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'            End If
'
'        Case 42565
'
'            'Pergunta se deseja criar o Almoxarifado
'            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO1", Codigo_Extrai(Almoxarifado.Text))
'            'Se a resposta for positiva
'            If vbMsg = vbYes Then
'
'                objAlmoxarifado.iCodigo = Codigo_Extrai(Almoxarifado.Text)
'
'                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
'
'                'Chama a tela de Almoxarifados
'                Call Chama_Tela("Almoxarifado", objAlmoxarifado)
'
'
'            Else
'                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'            End If
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156911)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Saida_Celula_ValorUnitario(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPrecoUnitario As Double

On Error GoTo Erro_Saida_Celula_ValorUnitario

'    If giTipoVersao = VERSAO_FULL Then
        Set objGridInt.objControle = ValorUnitario
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'        Set objGridInt.objControle = ValorUnitario
'    End If

    'Se estiver preenchido
    If Len(Trim(objGridInt.objControle.Text)) > 0 Then

        If giComplemento = MARCADO Then
            lErro = Valor_NaoNegativo_Critica(objGridInt.objControle.Text)
        Else
            lErro = Valor_Positivo_Critica(objGridInt.objControle.Text)
        End If
        If lErro <> SUCESSO Then gError 87193
        
        dPrecoUnitario = StrParaDbl(objGridInt.objControle.Text)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 42567

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 55578

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 42568
'*** fim tributacao

    Call ValorTotal_Calcula
    
    Saida_Celula_ValorUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorUnitario:

    Saida_Celula_ValorUnitario = gErr

    Select Case gErr

        Case 42566, 42567, 42568, 55578, 87193
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156912)

    End Select

    Exit Function

End Function

Function Saida_Celula_PercentDesc(objGridInt As AdmGrid) As Long

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
        If lErro <> SUCESSO Then Error 42569

'        dPercentDesc = CDbl(PercentDesc.Text)
'
'        lTamanho = Len(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col))
'        If lTamanho > 0 Then dPercentDescAnterior = PercentParaDbl(left(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col), lTamanho - 1))

        If dPercentDesc <> dPercentDescAnterior Then

            'Verifica se o percentual é de 100%
            If dPercentDesc = 100 Then Error 42571

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
    If lErro <> SUCESSO Then Error 42570

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then Error 55579

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then Error 39039
'*** fim tributacao

    Call ValorTotal_Calcula
    
    Saida_Celula_PercentDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDesc:

    Saida_Celula_PercentDesc = Err

    Select Case Err

        Case 39039, 42569, 42570, 55579
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 42571
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156913)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dPercentDesc As Double
Dim dValorTotal As Double
Dim dDesconto As Double
Dim dPrecoUnitario As Double
Dim dValorReal As Double
Dim dQuantidade As Double
Dim dSubTotal As Double
Dim lTamanho As Long

On Error GoTo Erro_Saida_Celula_Desconto

    Set objGridInt.objControle = Desconto

    'Se estiver preenchido
    If Len(Trim(Desconto.ClipText)) > 0 Then
    
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(Desconto.Text)
        If lErro <> SUCESSO Then gError 52258

        dDesconto = CDbl(Desconto.Text)

        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) > 0 Then dQuantidade = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
        
        '########################################
        'Inserido por Wagner 20/07/2006
        lErro = CF2(Me, "NFE_PrecoTotal_Calcula", GridItens.Row, dQuantidade)
        If lErro <> SUCESSO Then gError 181270
        '########################################
        
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) > 0 Then dPrecoUnitario = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))

        dValorTotal = dPrecoUnitario * dQuantidade

        If dValorTotal > 0 Then

            dPercentDesc = dDesconto / (dValorTotal)

            ' se dPercentDesc>=1("100%"), erro.
            If dPercentDesc >= 1 Then gError 52259

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
    If lErro <> SUCESSO Then gError 52260

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 55582

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 52261
'*** fim tributacao

    Call ValorTotal_Calcula
    
    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = gErr

    Select Case gErr

        Case 52258, 52260, 52261, 55582, 181270
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52259
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156914)

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
        If lErro <> SUCESSO Then Error 42582

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


    Select Case Err

        Case 42582

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156915)

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
        If lErro <> SUCESSO Then Error 42583

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


    Select Case Err

        Case 42583

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156916)

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
        If lErro <> SUCESSO Then Error 42584

        dValorDespesas = CDbl(ValorDespesas.Text)

        'coloca o valor formatado na tela
        ValorDespesas.Text = Format(dValorDespesas, "Standard")

    End If

    Call gobjTribTab.ValorDespesas_Validate(Cancel, dValorDespesas)

    Call ValorTotal_Calcula

    iValorDespesasAlterado = 0
    
    Exit Sub

Erro_ValorDespesas_Validate:

    Cancel = True


    Select Case Err

        Case 42584

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156917)

    End Select

    Exit Sub

End Sub

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
            lErro = gobjTribTab.TabClick
            If lErro <> SUCESSO Then Error 42586
        End If
        '*** fim tributacao

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(iFrameSelecionado).Visible = True
        
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = iFrameSelecionado

        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If TabStrip1.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao

        Select Case iFrameSelecionado
        
            Case TAB_DadosPrincipais
                Parent.HelpContextID = IDH_ENTRADA_NF_FATURA_DADOS_PRINCIPAIS
                
            Case TAB_Itens
                Parent.HelpContextID = IDH_ENTRADA_NF_FATURA_ITENS
                        
            Case TAB_Complemento
                Call Calcula_Msg
                Parent.HelpContextID = IDH_ENTRADA_NF_FATURA_COMPLEMENTOS
            
            Case TAB_Pagamento
                Parent.HelpContextID = IDH_ENTRADA_NF_FATURA_PAGAMENTO
                
            Case TAB_Tributacao
                Parent.HelpContextID = IDH_ENTRADA_NF_FATURA_TRIBUTACAO
            
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_ENTRADA_NF_FATURA_CONTABILIZACAO
                        
        End Select

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case Err

        Case 42586

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156918)

    End Select

    Exit Sub

End Sub

Public Sub Transportadora_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTransportadora As New ClassTransportadora
Dim iCodigo As Integer

On Error GoTo Erro_Transportadora_Validate

    'Verifica se foi preenchida a ComboBox Transportadora
    If Len(Trim(Transportadora.Text)) = 0 Then Exit Sub
    
    'Verifica se está preenchida com o ítem selecionado na ComboBox Transportadora
    If Transportadora.Text = Transportadora.List(Transportadora.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Transportadora, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 42587
    
    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo
        
        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 42588
        
        If lErro = 19250 Then Error 42589 ''Não encontrou Transportadora no BD

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    End If

    'Não existe o ítem com a String na List da Combo
    If lErro = 6731 Then Error 42590

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case Err

        Case 42587, 42588

        Case 42589
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objTransportadora.iCodigo)

        Case 42590
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", Err, Transportadora.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156919)

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
    If lErro <> SUCESSO And lErro <> 12253 Then Error 42591

    'Se não encontrar --> Erro
    If lErro = 12253 Then Error 42592

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True


    Select Case Err

        Case 42591

        Case 42592
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", Err, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156920)

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
    If lErro <> SUCESSO Then Error 42593

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True


    Select Case Err

        Case 42593

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156921)

    End Select

    Exit Sub

End Sub

Public Sub PesoBruto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoBruto_Validate

    'Verifica se está preenchido
    If Len(Trim(PesoBruto.ClipText)) = 0 Then Exit Sub

    'Criticao valor
    lErro = Valor_NaoNegativo_Critica(PesoBruto.Text)
    If lErro <> SUCESSO Then Error 42594

    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True


    Select Case Err

        Case 42594

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156922)

    End Select

    Exit Sub

End Sub

Public Sub SerieNFiscalOriginal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lTamanho As Long

On Error GoTo Erro_SerieNFiscalOriginal_Validate

    'Verifica se o tamanho da Serie é maior que 3
    If Len(Trim(SerieNFiscalOriginal.Text)) > STRING_SERIE Then Error 42595

    Exit Sub

Erro_SerieNFiscalOriginal_Validate:

    Cancel = True


    Select Case Err

        Case 42595
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_SERIE", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156923)

    End Select

    Exit Sub

End Sub

Public Sub PagamentoAutomatico_Click()

Dim bCancel As Boolean

    iAlterado = REGISTRO_ALTERADO

    If PagamentoAutomatico.Value = vbChecked And Len(Trim(CondicaoPagamento.Text)) > 0 Then Call Cobranca_Automatica

End Sub

Public Sub CondicaoPagamento_Click()

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_CondicaoPagamento_Click

    If gbCarregandoTela = True Then Exit Sub
    
    'Verifica se alguma Condição foi selecionada
    If CondicaoPagamento.ListIndex = -1 Then Exit Sub

    'Passa o código da Condição para objCondicaoPagto
    objCondicaoPagto.iCodigo = CondicaoPagamento.ItemData(CondicaoPagamento.ListIndex)

    'Lê Condição a partir do código
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then Error 42596
    
    If lErro = 19205 Then Error 42597
    
    If PagamentoAutomatico.Value = vbChecked Then
        
        'Preenche o GridParcelas
        lErro = GridParcelas_Preenche(objCondicaoPagto)
        If lErro <> SUCESSO Then Error 42598
    
    End If
    
    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_CondicaoPagamento_Click:

    Select Case Err

        Case 42596, 42598

        Case 42597
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156924)

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
    If CondicaoPagamento.ListIndex > -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(CondicaoPagamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 42250

    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro <> SUCESSO Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 42251
        
        If lErro = 19205 Then Error 42252

        'Testa se não pode ser usada em Contas a Receber
        If objCondicaoPagto.iEmPagamento = 0 Then Error 42253

        'Coloca na Tela
        CondicaoPagamento.Text = CondPagto_Traz(objCondicaoPagto)

        If gbCarregandoTela = False Then
                    
            If PagamentoAutomatico.Value = vbChecked Then
        
                'Preenche o GridParcelas
                lErro = GridParcelas_Preenche(objCondicaoPagto)
                If lErro <> SUCESSO Then Error 42254
            
            End If
            
        End If
                
        
    End If

    'Não encontrou o valor que era STRING
    If lErro = 6731 Then Error 42255

    Exit Sub

Erro_Condicaopagamento_Validate:

    Cancel = True

    Select Case Err

       Case 42250, 42251, 42254

       Case 42252
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If

        Case 42253
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", Err, objCondicaoPagto.iCodigo)

        Case 42255
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", Err, CondicaoPagamento.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156925)

    End Select

    Exit Sub

End Sub

Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto) As Long
'Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas

Dim lErro As Long
Dim dValorPagar As Double
Dim iIndice As Integer
Dim dtDataVenctoReal As Date

On Error GoTo Erro_GridParcelas_Preenche

    'Limpa o GridParcelas
    Call Grid_Limpa(objGridParcelas)

    'Número de Parcelas
    objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas

    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then Error 42599

    'Se Valor a Pagar for positivo
    If dValorPagar > 0 Then

        objCondicaoPagto.dValorTotal = dValorPagar
        
        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, False)
        If lErro <> SUCESSO Then Error 42599

        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        Next

    End If

    'Se Data Emissão estiver preenchida
    If Len(Trim(DataEmissao.ClipText)) > 0 Then

        objCondicaoPagto.dtDataEmissao = CDate(DataEmissao.Text)
        objCondicaoPagto.dtDataEntrega = MaskedParaDate(DataEntrada)
        
        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True)
        If lErro <> SUCESSO Then Error 42600

        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas

            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dtVencimento, "dd/mm/yyyy")

            'Calcula Data Vencimento Real
            lErro = CF("DataVencto_Real", objCondicaoPagto.colParcelas(iIndice).dtVencimento, dtDataVenctoReal)
            If lErro <> SUCESSO Then Error 42741

            'Coloca Data de Vencimento Real no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col) = Format(dtDataVenctoReal, "dd/mm/yyyy")

        Next

    End If

    GridParcelas_Preenche = SUCESSO

    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = Err

    Select Case Err

        Case 42599, 42600, 42741

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156926)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorParcela(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_ValorParcela

    Set objGridInt.objControle = ValorParcela

    'Verifica se valor está preenchido
    If Len(ValorParcela.ClipText) > 0 Then
    
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(ValorParcela.Text)
        If lErro <> SUCESSO Then Error 42601

        If CDbl(ValorParcela) <> StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Valor_Col)) Then PagamentoAutomatico.Value = vbUnchecked

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 42602

    Saida_Celula_ValorParcela = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorParcela:

    Saida_Celula_ValorParcela = Err

    Select Case Err

        Case 42601, 42602
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156927)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataVencimento(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim dtDataVencimentoAnterior As Date
Dim dtDataVenctoReal As Date

On Error GoTo Erro_Saida_Celula_DataVencimento

    Set objGridInt.objControle = DataVencimento

    dtDataVencimento = DATA_NULA
    dtDataVencimentoAnterior = DATA_NULA

    If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))) > 0 Then
        dtDataVencimentoAnterior = StrParaDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col))
    End If

    'Verifica se Data de Vencimento esta preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then Error 42603

         dtDataVencimento = CDate(DataVencimento.Text)

        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataEmissao.ClipText)) > 0 Then
            dtDataEmissao = CDate(DataEmissao.Text)
            If dtDataVencimento < dtDataEmissao Then Error 42605
        End If

        'Calcula a Data de Vencimento Real
        lErro = CF("DataVencto_Real", dtDataVencimento, dtDataVenctoReal)
        If lErro <> SUCESSO Then Error 42740

        'Coloca data de Vencimento Real no Grid
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_VenctoReal_Col) = Format(dtDataVenctoReal, "dd/mm/yyyy")

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    If dtDataVencimento <> dtDataVencimentoAnterior Then PagamentoAutomatico.Value = vbUnchecked

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 42604

    Saida_Celula_DataVencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_DataVencimento:

    Saida_Celula_DataVencimento = Err

    Select Case Err

        Case 42603, 42604, 42740
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 42605
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", Err, dtDataVencimento, GridParcelas.Row, dtDataEmissao)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156928)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cobranca(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoCobranca

On Error GoTo Erro_Saida_Celula_Cobranca

    Set objGridInt.objControle = TipoCobranca

    'Verifica se o Tipo de Cobrança foi preenchido
    If Len(Trim(TipoCobranca.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoCobranca.Text <> TipoCobranca.List(TipoCobranca.ListIndex) Then

            'Seleciona o Tipo de Cobrança
            lErro = Combo_Seleciona(TipoCobranca, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 42609

            If lErro = 6730 Then Error 42607
            If lErro = 6731 Then Error 42608

        End If

        If TipoCobranca.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) Then PagamentoAutomatico.Value = vbUnchecked

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 42606

    Saida_Celula_Cobranca = SUCESSO

    Exit Function

Erro_Saida_Celula_Cobranca:

    Saida_Celula_Cobranca = Err

    Select Case Err

        Case 42606
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 42607
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_CADASTRADO", iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 42608
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_ENCONTRADO", TipoCobranca.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 42609
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_ENCONTRADO", Err, TipoCobranca.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156929)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 42613

    'Limpa a Tela
    Call Limpa_Tela_NFiscal2

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 42613

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 156930)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iIndice As Integer
Dim dTotal As Double
Dim objNFiscal As New ClassNFiscal
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim dQuantLote As Double
Dim objItemNF As New ClassItemNF
Dim objRastroItemNF As ClassRastroItemNF

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 42614
    If Len(Trim(NaturezaOp.Text)) = 0 Then gError 42615
    If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 42616
    
    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 42623
    
    If lErro = 31415 Then gError 42624 'Não encontrou
    
    'Critica e Preenche a NumNota se for uma Nota Externa
    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
        'Se for Interna --> não pode preencher o campo Número
        'If Len(Trim(NFiscalInterna.Caption)) > 0 Then gError 60453
        If Len(Trim(NFiscalInterna.Caption)) > 0 Then
            objNFiscal.lNumNotaFiscal = CLng(NFiscalInterna.Caption)
        Else
            objNFiscal.lNumNotaFiscal = 0
        End If
    Else
        If Len(Trim(NFiscal.ClipText)) = 0 Then gError 42619
        objNFiscal.lNumNotaFiscal = CLng(NFiscal.Text)
    End If
    
    objNFiscal.iTipoDocInfo = objTipoDocInfo.iTipo
    
    If Len(Trim(Serie.Text)) = 0 Then gError 42618
    If Len(Trim(DataEntrada.ClipText)) = 0 Then gError 42620
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 42621
    
    'se o tipo de Nfiscal for do tipo "Externa - Mat. A ser beneficiado" entao obtem o Fornecedor a ser beneficiado
    If TipoNFiscal.ListIndex <> -1 Then
        
        If TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFEEFBF Or TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFIEFBF Or TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFEFCCO Then
            
            If Len(Trim(FornecedorBenef.Text)) = 0 Then gError 52508
            If Len(Trim(FilialFornBenef.List(FilialFornBenef.ListIndex))) = 0 Then gError 52509
        
        End If
    End If
    
    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Then
        If CDate(DataEntrada.Text) < CDate(DataEmissao.Text) Then
            vbMsgRes = Rotina_Aviso(vbYesNo, "CONFIRMA_NF_ENTRADA_ANTERIOR_EMISSAO")
        
            If vbMsgRes = vbNo Then gError 87194
        End If
    Else
        'Verifica se a Data de Entrada e menor que a Data de Emissão
        If CDate(DataEntrada.Text) < CDate(DataEmissao.Text) Then gError 42622
    End If
    
    If objTipoDocInfo.iComplementar = DOCINFO_NORMAL Then
        If objGridItens.iLinhasExistentes = 0 Then gError 42627
    End If

    'Recolhe os dados da tela
    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 42636

    lErro = Move_GridItens_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 42742
    
    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786

    lErro = gobjTribTab.Valida_Dados
    If lErro <> SUCESSO Then gError 181786

    'distribuicao
    lErro = gobjDistribuicao.Move_GridDist_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 89529

    'Valida os dados do Grid
    If objGridItens.iLinhasExistentes > 0 Then
        
        For iIndice = 1 To objGridItens.iLinhasExistentes

            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 42628
            
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col))) = 0 Then gError 49396

            Set objItemNF = objNFiscal.ColItensNF(iIndice)
            
            objProduto.sCodigo = objItemNF.sProduto

            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 42630
            
            If lErro = 28030 Then gError 42631

            objItemNF.iClasseUM = objProduto.iClasseUM
            objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
            objItemNF.sUMVenda = objProduto.sSiglaUMVenda

'            If TipoNFiscal.ListIndex <> -1 Then
'
'                If TipoNFiscal.ItemData(TipoNFiscal.ListIndex) <> DOCINFO_NFEEBEN And TipoNFiscal.ItemData(TipoNFiscal.ListIndex) <> DOCINFO_NFIEEBEN Then
'
'                    'se a nota é de entrada de material beneficiado por terceiros, pode ser um produto não comprável
'                    If objProduto.iCompras = PRODUTO_NAO_COMPRAVEL Then gError 55553
'
'                End If
'
'            End If

            'Verifica se a Unidade de Medida foi preenchida
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) = 0 Then gError 55629
            
            If objTipoDocInfo.iCodigo <> TIPODOCINFO_COMPLEMFAT_INT_IPI And objTipoDocInfo.iCodigo <> TIPODOCINFO_COMPLEMFAT_EXT_IPI Then
                If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) = 0 Then gError 42633
            
                ElseIf objTipoDocInfo.iCodigo = TIPODOCINFO_COMPLEMFAT_EXT_IPI Then
                    vbMsgRes = Rotina_Aviso(vbYesNo, "CONFIRMA_VUNITARIO_GRAVACAO")
                
                    If vbMsgRes = vbNo Then gError 87194
            End If
            
            '####################################
            'Inserido por Wagner
            lErro = Valida_Contrato_Itens(objItemNF, iIndice)
            If lErro <> SUCESSO Then gError 136023
            '####################################
            
            '############################################################
            'Inserido por Wagner 31/07/2006
            lErro = CF2(Me, "NFE_Valida_Grid_Itens", iIndice)
            If lErro <> SUCESSO Then gError 181418
            '############################################################
            
        Next
        
    End If

    dTotal = CDbl(IIf(Len(Trim(Total.Text)) > 0, Total.Text, 0))

    'Se o total for negativo --> Erro
    If dTotal < 0 Then gError 42634

    If Len(Trim(PesoLiquido.ClipText)) > 0 And Len(Trim(PesoBruto.ClipText)) > 0 Then
        If CDbl(PesoLiquido.Text) > CDbl(PesoBruto.Text) Then gError 42635
    End If

    lErro = Total_Calcula()
    If lErro <> SUCESSO Then gError 43806

    'se o tab de beneficiamento existe
    If TabStrip1.Tabs.Count = TAB_BENEFICIAMENTO Then
    
        'Valida os dados do grid de beneficiamento
        For iIndice = 1 To objGridBenef.iLinhasExistentes
            
            'Verifica se a Quantidade foi informada
            If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_QuantidadeBenef_Col))) = 0 Then gError 55635

            'Verifica se a Unidade de Medida foi preenchida
            If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMedBenef_Col))) = 0 Then gError 55636
            
            'Verifica se o Almoxarifado foi informado
            If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxarifadoBenef_Col))) = 0 Then gError 55637

        Next
        
    End If

    lErro = gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
    If lErro <> SUCESSO Then gError 83272

    Set objNFiscal.objRastreamento = gobjRastreamento

    If (objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO) Or (objTipoDocInfo.iCodigo = DOCINFO_FDDAI) Or (objTipoDocInfo.iCodigo = DOCINFO_NFIEFCNFP) Or (objTipoDocInfo.iCodigo = DOCINFO_NFSEFT) Then
    
        'Recolhe os dados da Nota Fiscal Original
        lErro = Move_Dados_NFOriginal_Memoria(objNFiscal, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 42637
        
    End If

    'Recolhe os dados do Grid Movimentos
    If (objNFiscal.iTipoNFiscal = DOCINFO_NFEEFBEN) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEFBEN) Then
    
        lErro = Move_GridMovimentos_Memoria(objNFiscal)
        If lErro <> SUCESSO Then gError 52512
        
    End If
    
    'mover a parte de tributacao
    lErro = gobjTribTab.PreencherNF(objNFiscal)
    If lErro <> SUCESSO Then gError 42526

    lErro = gobjTribTab.Valida_TributacaoNF_Gravacao(objNFiscal)
    If lErro <> SUCESSO Then gError 59002
    
    lErro = Calcula_Msg
    If lErro <> SUCESSO Then gError 59002
    
    'Valida o Grid de Parcelas
    lErro = Valida_Grid_Parcela()
    If lErro <> SUCESSO Then gError 64136
    
    lErro = Move_GridParcelas_Memoria(objNFiscal.ColParcelaPagar)
    If lErro <> SUCESSO Then gError 42743

    'verifica se o total das parcelas coincide com o valor total da nota menos retencoes
    lErro = Valida_Total_Parcelas(objNFiscal)
    If lErro <> SUCESSO Then gError 55869

    'mover a parte do rastreamento
    lErro = gobjRastreamento.Move_Rastro_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 83267

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataEntrada.Text))
    If lErro <> SUCESSO Then gError 92036

    lErro = CF2(Me, "NFEntC_Move_Tela_Memoria", objNFiscal)
    If lErro <> SUCESSO Then gError 124253
    
    lErro = CF("NFiscal_Valida_Diversos", objNFiscal)
    If lErro <> SUCESSO Then gError 124253
    
    'Grava a Nota Fiscal Fatura de Entrada(inclusive os dados contábeis) (contabilidade)
    lErro = CF("NFiscalFatEntrada_Grava", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 42638

    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_NFISCAL
        gobjAnotacao.sID = CStr(objNFiscal.lNumIntDoc)
        gobjAnotacao.dtDataAlteracao = gdtDataHoje
        
        lErro = CF("Anotacoes_Grava", gobjAnotacao)
        If lErro <> SUCESSO Then gError 130406
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
    
        If Len(Trim(NFiscalInterna.Caption)) = 0 Then vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_NUMERO_NOTA_GRAVADA", objNFiscal.lNumNotaFiscal)
    
    End If
    
    If ISSerieEletronica(objNFiscal.sSerie) Then
        lErro = CF("NFE_Grava1", objNFiscal.lNumIntDoc, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 203005
    End If
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 42614
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 42615
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 42616
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 42617
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 42618
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 42619
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)

        Case 42620
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_NAO_PREENCHIDA", gErr)

        Case 42621
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 42622
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_ANTERIOR_DATAEMISSAO", gErr, DataEntrada, DataEmissao)

        Case 42623, 42629, 42630, 42636, 42638, 42637, 42526, 42742, 42743, 43806, 52512, 59002, 55869, 64136, 83267, 83272, 87194, 89529, 92036, 124253, 136023, 130406, 181418, 203005

        Case 42624
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iTipo)

        Case 42627
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)

        Case 42628
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case 42631
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 42633
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 42634
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)

        Case 42635
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr, CDbl(PesoLiquido.Text), CDbl(PesoBruto.Text))

        Case 49396
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAOITEM_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 52508
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_BENEF_NAO_PREENCHIDA", gErr)
        
        Case 52509
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORNECEDOR_BENEF_NAO_PREENCHIDA", gErr)
                        
        Case 55553
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", gErr, GridItens.TextMatrix(iIndice, iGrid_Produto_Col))
                        
        Case 55629
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDA", gErr, iIndice)
                        
        Case 55635
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_BENEF_NAO_PREENCHIDA", gErr, iIndice)
            
        Case 55636
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_BENEF_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 55637
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_BENEF_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 75929
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTTOTAL_LOTE_MAIOR_ALMOXARIFADO", gErr, dQuantLote, StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)))
        
        Case 181786
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156931)

    End Select

    Exit Function

End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objNFiscal As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objFornecedor As New ClassFornecedor
Dim snProt As String
Dim iEnviado As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 89055
    If Len(Trim(Serie.Text)) = 0 Then gError 89056
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 89057

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
'            If lErro <> SUCESSO Then gError 203046
'
'            If iEnviado = 1 Then gError 203047
'
'        End If
        
    Else
        
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
        If lErro <> SUCESSO Then gError 89060

        'Faz o cancelamento de uma nota fiscal de Saida
        lErro = CF("NotaFiscalEntrada_Excluir", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError 89061
    
        'Limpa a Tela
        Call Limpa_Tela_NFiscal2
    
        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 89055
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 89056
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 89057
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 89058, 89060, 89061, 99345, 99351, 99352, 99353, 203046
       
        Case 89065
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
            
        Case 99346
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)
        
        Case 99350
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
        
        Case 99352
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)
       
        Case 203047
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CADASTRADA_NAO_EXCLUI", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156932)
        
    End Select

    Exit Sub

End Sub

Private Function Valida_Total_Parcelas(objNFiscal As ClassNFiscal) As Long
'verifica se o total das parcelas coincide com o valor total da nota menos retencoes de IR e INSS

Dim lErro As Long
Dim objParcelaPag As ClassParcelaPagar
Dim dValorTotal As Double, dValorPagar As Double

On Error GoTo Erro_Valida_Total_Parcelas

    For Each objParcelaPag In objNFiscal.ColParcelaPagar
        
        dValorTotal = dValorTotal + objParcelaPag.dValor

    Next
    
    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then Error 42646
    
    'se o valor da nota fiscal menos retencoes não coincide com o total das parcelas a pagar ==> erro
    'If Format(objNFiscal.dValorTotal - objNFiscal.objTributacaoNF.dTotalRetencoes, "0.00") <> Format(dValorTotal, "0.00") Then Error 55870
    If Format(dValorPagar, "0.00") <> Format(dValorTotal, "0.00") Then Error 55870
    
    Valida_Total_Parcelas = SUCESSO
    
    Exit Function
    
Erro_Valida_Total_Parcelas:

    Valida_Total_Parcelas = Err
    
    Select Case Err
    
        Case 42646

        Case 55870
            Call Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_DIFERE_PARCELASPAGAR", Err, objNFiscal.dValorTotal - objNFiscal.objTributacaoNF.dTotalRetencoes, dValorTotal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156933)

    End Select

    Exit Function

End Function

Public Sub CondicaoPagamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEntrada_Change()

    iAlterado = REGISTRO_ALTERADO
    iDataEntradaAlterada = REGISTRO_ALTERADO

End Sub

Public Sub DataVencimentoReal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Desconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DescricaoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO
    
    'Verifica se alguma filial foi selecionada
    If Filial.ListIndex = -1 Then Exit Sub
     
    Call Trata_FilialForn
    
End Sub

Public Sub Mensagem_Change()
    iAlterado = REGISTRO_ALTERADO
    MsgAutomatica.Value = vbUnchecked
End Sub

Public Sub MensagemCorpo_Change()
    iAlterado = REGISTRO_ALTERADO
    MsgAutomatica.Value = vbUnchecked
End Sub

Public Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NaturezaOp_Change()

    iAlterado = REGISTRO_ALTERADO
    iNaturezaOpAlterada = 1

End Sub

Public Sub NFiscal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NFiscalOriginal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentDesc_Change()

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

Public Sub ValorTotal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorUnitario_Change()

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

Public Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Serie_Change()

    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = REGISTRO_ALTERADO

End Sub

Public Sub SerieNFiscalOriginal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub SerieNFiscalOriginal_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Suspenso_Click()

    iAlterado = REGISTRO_ALTERADO
    PagamentoAutomatico.Value = vbUnchecked

End Sub

Public Sub Suspenso_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Suspenso_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Suspenso_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Suspenso
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TipoCobranca_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoCobranca_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub TipoCobranca_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub TipoCobranca_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = TipoCobranca
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

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

Public Sub ValorDespesas_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorDespesasAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorDesconto_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorDescontoAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorDesconto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorDesconto As Double
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
        If lErro <> SUCESSO Then Error 35204

        dValorDesconto = CDbl(ValorDesconto.Text)
        'Coloca o valor formatado na tela
        ValorDesconto.Text = Format(dValorDesconto, "Standard")

    End If

    ''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.ValorDesconto_Validate(Cancel, dValorDesconto)
    '*** fim tributacao

    Call ValorTotal_Calcula

    iValorDescontoAlterado = 0
    
    Exit Sub

Erro_ValorDesconto_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 35204

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156934)

    End Select

    Exit Sub

End Sub

Public Sub ValorFrete_Change()

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

    Set objGridItens.objControle = ValorParcela
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

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
    If lErro <> SUCESSO Then gError 102462
    
    Exit Sub

Erro_VolumeEspecie_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102462
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156935)

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
    If lErro <> SUCESSO Then gError 102463
    
    Exit Sub

Erro_VolumeMarca_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102463
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156936)

    End Select

End Sub

Public Sub VolumeNumero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VolumeQuant_Change()

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

    Set objGridItens.objControle = DataVencimento
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Move_NF_Memoria(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objNFiscalOriginal As New ClassNFiscal
Dim objProjetoInfo As Object
Dim iCodigo As Integer

On Error GoTo Erro_Move_NF_Memoria

    'Se o Fornecedor estiver preenchido lê o código no BD
    If Len(Trim(Fornecedor.Text)) > 0 Then

        objFornecedor.sNomeReduzido = Fornecedor.Text
        'Lê o fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 42524
        
        If lErro = 6681 Then gError 42527 'Se nao encontrar --> erro

        objNFiscal.lFornecedor = objFornecedor.lCodigo

    End If

    'Recolhe os demais campos da NFiscal
    If Len(Trim(Filial.Text)) > 0 Then objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)

    If Len(Trim(TipoNFiscal.Text)) > 0 Then objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    objNFiscal.dtDataReferencia = DATA_NULA
    objNFiscal.lNumIntDoc = 0
    objNFiscal.sNaturezaOp = NaturezaOp.Text
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    If Not NFiscalInterna.Visible Then objNFiscal.sChvNFe = ChvNFe.ClipText
    
    If objNFiscal.iTipoDocInfo = TIPODOCINFO_TIPO_NFEXT Then
        objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
    Else
        objNFiscal.lNumNotaFiscal = StrParaLong(NFiscalInterna.Caption)
    End If
    
    objNFiscal.sPlaca = Placa.Text
    objNFiscal.sPlacaUF = PlacaUF.Text
    
    'Incluído por Luiz Nogueira em 21/08/03
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objNFiscal.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objNFiscal.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
    
    objNFiscal.sMensagemNota = Mensagem.Text
    objNFiscal.sMensagemCorpoNota = MensagemCorpo.Text
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.dtDataSaida = DATA_NULA
    objNFiscal.iStatus = STATUS_LANCADO
    objNFiscal.sObservacao = Observacao.Text
    
    
    'se o tipo de Nfiscal for do tipo "Externa - Mat. A ser beneficiado" entao obtem o Fornecedor a ser beneficiado
    If TipoNFiscal.ListIndex <> -1 Then
        
        If TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFEEFBF Or TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFIEFBF Or TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFEFCCO Then
                
            If Len(Trim(FornecedorBenef.Text)) > 0 Then
            
                Set objFornecedor = New ClassFornecedor
            
                objFornecedor.sNomeReduzido = FornecedorBenef.Text
            
                'Lê o fornecedor
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO And lErro <> 6681 Then gError 52510
            
                If lErro = 6681 Then gError 52511 'Se nao encontrar --> erro
    
                objNFiscal.lFornecedorBenef = objFornecedor.lCodigo
    
                If Len(Trim(FilialFornBenef.List(FilialFornBenef.ListIndex))) > 0 Then objNFiscal.iFilialFornBenef = Codigo_Extrai(FilialFornBenef.List(FilialFornBenef.ListIndex))
            End If
        End If
    End If
    
    
    
    'Armazena o responsável pelo frete
    objNFiscal.iFreteRespons = TipoFrete.ItemData(TipoFrete.ListIndex)
    
    If Len(Trim(NumRecebimento.Text)) > 0 Then objNFiscal.lNumRecebimento = CLng(NumRecebimento.Text)
    
    objNFiscal.sVolumeNumero = VolumeNumero.Text
    If Len(Trim(ValorDespesas.ClipText)) > 0 Then objNFiscal.dValorOutrasDespesas = CDbl(ValorDespesas.Text)
    If Len(Trim(ValorFrete.ClipText)) > 0 Then objNFiscal.dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorSeguro.ClipText)) > 0 Then objNFiscal.dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(Total.Text)) > 0 Then objNFiscal.dValorTotal = CDbl(Total.Text)
    If Len(Trim(SubTotal.Caption)) > 0 Then objNFiscal.dValorProdutos = CDbl(SubTotal.Caption)

    If Len(Trim(PesoBruto.ClipText)) > 0 Then objNFiscal.dPesoBruto = CDbl(PesoBruto.Text)
    If Len(Trim(PesoLiquido.ClipText)) > 0 Then objNFiscal.dPesoLiq = CDbl(PesoLiquido.Text)

    If Len(Trim(VolumeQuant.ClipText)) > 0 Then objNFiscal.lVolumeQuant = StrParaLong(VolumeQuant.Text)

    If Len(Trim(Transportadora.Text)) > 0 Then objNFiscal.iCodTransportadora = Codigo_Extrai(Transportadora.Text)

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
    
    If Len(Trim(DataVencimento.ClipText)) > 0 Then
        objNFiscal.dtDataVencimento = CDate(DataVencimento.Text)
    Else
        objNFiscal.dtDataVencimento = DATA_NULA
    End If
    
    objNFiscal.dtDataRegistro = gdtDataAtual
    
    objNFiscal.iCondicaoPagto = CondPagto_Extrai(CondicaoPagamento)
    
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181761
    
    Set objNFiscal.objProjetoInfo = objProjetoInfo
    
    Set objNFiscal.objInfoAdic = gobjInfoAdic
    objNFiscal.dValorItens = StrParaDbl(ValorProdutos2.Caption)
    objNFiscal.dValorDescontoItens = StrParaDbl(ValorDescontoItens.Text)

    Move_NF_Memoria = SUCESSO

    Exit Function

Erro_Move_NF_Memoria:

    Move_NF_Memoria = Err

    Select Case Err

        Case 42524, 126965, 181761

        Case 42527
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, Fornecedor.Text)
        
        Case 52510
        
        Case 52511
          Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, FornecedorBenef.Text)
        
        Case 126966
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156937)

    End Select

    Exit Function

End Function

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long
'Preenche os dados do Produto da linha do grid selecionada

Dim lErro As Long

On Error GoTo Erro_ProdutoLinha_Preenche

    lErro = CF2(Me, "NFFE_ProdutoLinha_Preenche", objProduto)
    If lErro <> SUCESSO Then gError 199277
    
    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr
    
        Case 199277

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199278)

    End Select

    Exit Function
    
End Function

Function NFFE_ProdutoLinha_Preenche(ByVal objCT As Object, objProduto As ClassProduto) As Long
'Faz a crítica da célula Produto do grid que está deixando de ser a corrente

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
Dim colMedicao As New Collection 'Inserido por Wagner

On Error GoTo Erro_NFFE_ProdutoLinha_Preenche

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 180063

    If objProduto.iGerencial And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then gError 86296
    
    If Len(Trim(objProduto.sGrade)) > 0 Then iPossuiGrade = MARCADO

    If iPossuiGrade = DESMARCADO And objProduto.iKitVendaComp <> MARCADO Then
    
        If Grid_Possui_Grade(objCT) Then
        
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

    '####################################################
    'Inserido por Wagner 19/07/2006
    lErro = CF2(Me, "NFE_ProdutoLinha_Preenche", objProduto)
    If lErro <> SUCESSO Then gError 181258
    '####################################################

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridItens.Row - GridItens.FixedRows) = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    
''*** incluidos p/tratamento de tributacao *******************************
        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 55653
'*** fim tributacao

        gcolcolMedicoes.Add colMedicao 'Inserido por Wagner
        
        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItens.Row)

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
        
    End If
    
    lErro = Preco_Traz_Padrao(GridItens.Row, objProduto.sCodigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Calcula_Valores(GridItens.Row)
    
    NFFE_ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_NFFE_ProdutoLinha_Preenche:

    NFFE_ProdutoLinha_Preenche = gErr
    
    Select Case gErr

        Case 30317, 30318, 55653, 86310, 86327, 86329, 177582, 180063, 181258

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
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156938)

    End Select

    Exit Function

End Function


Public Function Traz_NFiscal_Tela(objNFiscal As ClassNFiscal) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim objNFiscalOriginal As New ClassNFiscal
Dim lErro As Long, objTituloPag As New ClassTituloPagar
Dim bCancel As Boolean
Dim bAlterouCT As Boolean
Dim objItemNF As ClassItemNF
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iIndice As Integer, dPercDesc As Double

On Error GoTo Erro_Traz_NFiscal_Tela

    'Se for versao LIGHT testa chaveamento
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
    
    gbCarregandoTela = True
    
    'Limpa a tela
    Call Limpa_Tela_NFiscal
    
    'carrega dados do titulo a pagar associado à nota fiscal para obter a condicao de pagto
    objTituloPag.lNumIntDoc = objNFiscal.lNumIntDocCPR
    
    lErro = CF("TituloPagar_Le", objTituloPag)
    If lErro <> SUCESSO And lErro <> 18372 Then gError 56662
    
    If lErro <> SUCESSO Then
    
        lErro = CF("TituloPagarBaixado_Le", objTituloPag)
        If lErro <> SUCESSO And lErro <> 56661 Then gError 56663
        
'        If lErro <> SUCESSO Then gError 56664
    
    End If
    
    objNFiscal.iCondicaoPagto = objTituloPag.iCondicaoPagto
    
    'Lê a parte de Tributação
    lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
    If lErro <> SUCESSO Then gError 42744

    'Tenta ler a nota Fiscal passada por parametro
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 42745
    
    'Lê os Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 42745
    
    'Lê as mensagens
    lErro = CF("NFiscal_Le_Msg", objNFiscal)
    If lErro <> SUCESSO Then gError 42745

    'distribuicao
    'Lê a Distribuição dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 89530
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Le o Tipo para saber se tem Movimento de Estoque
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO Then gError 20807
    
    If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
        'Carrega ItensNF com Rastreamentos
        lErro = gobjRastreamento.Carrega_RastroItensNF(objNFiscal)
        If lErro <> SUCESSO Then gError 20806
    End If

    'Lê a parte de tributacao dos Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
    If lErro <> SUCESSO Then gError 42746

    lErro = CF("ParcelasPagNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 42747
    
    Set gobjNFiscal = objNFiscal
    
    lErro = CF("NFiscal_Le_ChvNFe", objNFiscal)
    If lErro <> SUCESSO Then gError 42745
    
    ChvNFe.PromptInclude = False
    ChvNFe.Text = objNFiscal.sChvNFe
    ChvNFe.PromptInclude = True
    gsChvAnt = objNFiscal.sChvNFe

    'Lê os ítens do Movimento se for Benefiamento e preenche o grid
    If (objNFiscal.iTipoNFiscal = DOCINFO_NFEEBEN) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEEBEN) Then
        
        lErro = CF("MovEstoqueItensBenef_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 52384 Then gError 55648
        
        lErro = Preenche_GridMovimentos(objNFiscal)
        If lErro <> SUCESSO Then gError 55649
            
    End If

    'preenche o Fornecedor
    Fornecedor.Text = objNFiscal.lFornecedor
    Call Fornecedor_Validate(bCancel)
    
    'Preenche a Filial do Fornecedor
    Filial.Text = objNFiscal.iFilialForn
    Call Filial_Validate(bCancel)

    TipoNFiscal.Text = objNFiscal.iTipoNFiscal
    Call TipoNFiscal_Validate(bCancel)

    If objNFiscal.sNaturezaOp <> "" Then NaturezaOp.Text = objNFiscal.sNaturezaOp
    NFiscal.Text = objNFiscal.lNumNotaFiscal
    NFiscalInterna.Caption = objNFiscal.lNumNotaFiscal
    Serie.Text = Desconverte_Serie_Eletronica(objNFiscal.sSerie)
    If ISSerieEletronica(objNFiscal.sSerie) Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    Call Serie_Validate(bSGECancelDummy)
    
    'Se não trouxer para tela ao regravar vai excluir a NF anterior e apagar
    'qualquer rastro de uso do recebimento deixando-o perdido conforme ocorreu na DLE
    If objNFiscal.lNumRecebimento <> 0 Then NumRecebimento.Text = objNFiscal.lNumRecebimento

    'Preenche o Status da Nota Fiscal
    If objNFiscal.iStatus = STATUS_LANCADO Then
        Status.Caption = STRING_STATUS_LANCADO
    ElseIf objNFiscal.iStatus = STATUS_BAIXADO Then
        Status.Caption = STRING_STATUS_BAIXADO
    ElseIf objNFiscal.iStatus = STATUS_CANCELADO Then
        Status.Caption = STRING_STATUS_CANCELADO
    End If

    'Se a data não for nula coloca na Tela
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
    Observacao.Text = objNFiscal.sObservacao
    
    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = CStr(objNFiscal.dPesoLiq)
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = CStr(objNFiscal.dPesoBruto)
    
    If objNFiscal.iTipoNFiscal = DOCINFO_NFEEFBF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEFBF Or objNFiscal.iTipoNFiscal = DOCINFO_NFEFCCO Then

        'Preenche o Fornecedor Benef
        If objNFiscal.lFornecedorBenef <> 0 Then
            FornecedorBenef.Text = objNFiscal.lFornecedorBenef
        Else
            FornecedorBenef.Text = ""
        End If
        Call FornecedorBenef_Validate(bCancel)
        
        'Preenche a Filial do Fornecedor
        If objNFiscal.iFilialFornBenef <> 0 Then
            FilialFornBenef.Text = objNFiscal.iFilialFornBenef
        Else
            FilialFornBenef.Text = ""
        End If
        Call FilialFornBenef_Validate(bCancel)
        
    End If
    
    'preenche o Grid com os Ítens da Nota Fiscal
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 42748

    'distribuicao
    'Preenche o Grid com as Distribuições dos itens da Nota Fiscal
    lErro = gobjDistribuicao.Preenche_GridDistribuicao(objNFiscal)
    If lErro <> SUCESSO Then gError 89531

    'Preenche a Condicao de Pagto
    If objNFiscal.iCondicaoPagto > 0 Then
        CondicaoPagamento.Text = objNFiscal.iCondicaoPagto
        Call CondicaoPagamento_Validate(bCancel)
    End If
        
    lErro = Preenche_GridParcelas(objNFiscal)
    If lErro <> SUCESSO Then gError 42749

    'Se o código da Nota Fiscal Original estiver preenchido
    If objNFiscal.lNumIntNotaOriginal <> 0 Then
    
        objNFiscalOriginal.lNumIntDoc = objNFiscal.lNumIntNotaOriginal
        
        'Lê a Nota Fiscal original
        lErro = CF("NFiscal_Le", objNFiscalOriginal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 42750

        If lErro <> SUCESSO Then gError 42751
        
        'Coloca os dados relativos  Nota Fiscal Original na Tela
        SerieNFiscalOriginal.Text = objNFiscalOriginal.sSerie
        NFiscalOriginal.Text = objNFiscalOriginal.lNumNotaFiscal
        
        If (objNFiscal.iTipoNFiscal = DOCINFO_FDDAI) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFSEFT) Then
    
            'Preenche o Fornecedor da NF Orig
            FornNFOrig.Text = objNFiscalOriginal.lFornecedor
            Call FornNFOrig_Validate(bCancel)
            
            'Preenche a Filial do Fornecedor
            FilialFornNFOrig.Text = objNFiscalOriginal.iFilialForn
            Call FilialFornNFOrig_Validate(bCancel)
    
        End If
        
    Else

        If objNFiscal.lNumNFPOrig <> 0 Then
    
            'Coloca os dados relativos  Nota Fiscal Original na Tela
            SerieNFiscalOriginal.Text = objNFiscal.sSerieNFPOrig
            NFiscalOriginal.Text = objNFiscal.lNumNFPOrig
    
        End If

    End If

    'Set gobjNFiscal = objNFiscal

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 42752
'*** fim tributacao

    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 39657

    lErro = CF2(Me, "NFEntC_Traz_NFiscal_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 124252

    '#############################################################
    'Inserido por Wagner
    lErro = CF2(Me, "NFCF_Fornecedor_Traz_NF_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 131950
    '#############################################################

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

    iAlterado = 0

    gbCarregandoTela = False
    
    Traz_NFiscal_Tela = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela:

    gbCarregandoTela = False
    
    Traz_NFiscal_Tela = gErr

    Select Case gErr

        Case 20806, 20807, 25822, 39657, 42744, 42745, 42746, 42747, 42748, 42749, 42750, 42752, 55648, 55649, 56662, 56663, 89530, 89531, 124252, 131950 'Inserido por Wagner

        Case 42751
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntNotaOriginal)

        Case 56664
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156939)

    End Select

    Exit Function

End Function

'*** Retirar futuramento pois foi colocado em ClassRastreamento já que todas as notas utilizam o mesmo codigo
'Function Carrega_RastroItensNF(objNFiscal As ClassNFiscal) As Long
''Carrega rastreamento dos itens da nota fiscal
'
'Dim lErro As Long
'Dim objItemNF As ClassItemNF
'Dim colRastroMovto As New Collection
'Dim objRastroMovto As ClassRastreamentoMovto
'Dim objRastroItemNF As ClassRastroItemNF
'Dim objAlmoxarifado As New ClassAlmoxarifado
'Dim objRastreamentoLote As New ClassRastreamentoLote
'Dim objItemMovEstoque As New ClassItemMovEstoque
'
'On Error GoTo Erro_Carrega_RastroItensNF
'
'    'Para cada item da nota fiscal
'    For Each objItemNF In objNFiscal.ColItensNF
'
'        'Lê o Almoxarifado
'        objAlmoxarifado.iCodigo = objItemNF.iAlmoxarifado
'        lErro = CF("Almoxarifado_Le",objAlmoxarifado)
'        If lErro <> SUCESSO And lErro <> 25056 Then gError 75949
'
'        'Se não encontrou Almoxarifado --> Erro
'        If lErro = 25056 Then gError 75952
'
'        'Lê item de movimento de estoque
'        objItemMovEstoque.lNumIntDocOrigem = objItemNF.lNumIntDoc
'        objItemMovEstoque.iTipoNumIntDocOrigem = TIPO_ORIGEM_ITEMNF
'        objItemMovEstoque.iFilialEmpresa = giFilialEmpresa
'        lErro = CF("MovEstoque_Le_ItemNF",objItemMovEstoque)
'        If lErro <> SUCESSO And lErro <> 75796 Then gError 75951
'
'        'Se não encontrou, erro
'        If lErro = 75796 Then gError 75954
'
'        'Lê movimentos de rastreamento vinculados ao itemNF passado ao ItemNF
'        Set colRastroMovto = New Collection
'        lErro = CF("RastreamentoMovto_Le_DocOrigem",objItemMovEstoque.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRastroMovto)
'        If lErro <> SUCESSO Then gError 75948
'
'        Set objItemNF.colRastreamento = New Collection
'
'        'Guarda as quantidades alocadas dos lotes
'        For Each objRastroMovto In colRastroMovto
'
'            Set objRastroItemNF = New ClassRastroItemNF
'
'            objRastroItemNF.dLoteQdtAlocada = objRastroMovto.dQuantidade
'            objRastroItemNF.sLote = objRastroMovto.sLote
'            objRastroItemNF.iAlmoxCodigo = objItemNF.iAlmoxarifado
'            objRastroItemNF.sAlmoxNomeRed = objAlmoxarifado.sNomeReduzido
'            objRastroItemNF.dAlmoxQtdAlocada = objItemNF.dQuantidade
'
'            'procura RastreamentoLote
'            objRastreamentoLote.sProduto = objItemNF.sProduto
'            objRastreamentoLote.iFilialOP = objRastroMovto.iFilialOP
'            objRastreamentoLote.sCodigo = objRastroMovto.sLote
'            lErro = CF("RastreamentoLote_Le",objRastreamentoLote)
'            If lErro <> SUCESSO And lErro <> 75710 Then gError 75950
'
'            'Se não encontrou, erro
'            If lErro = 75710 Then gError 75953
'
'            objRastroItemNF.dtLoteData = objRastreamentoLote.dtDataEntrada
'            objRastroItemNF.iLoteFilialOP = objRastreamentoLote.iFilialOP
'
'            'Adiciona na coleção de rastreamento
'            objItemNF.colRastreamento.Add objRastroItemNF
'
'        Next
'
'    Next
'
'    Carrega_RastroItensNF = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_RastroItensNF:
'
'    Carrega_RastroItensNF = gErr
'
'    Select Case gErr
'
'        Case 75948, 75949, 75950, 75951
'
'        Case 75952
'            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)
'
'        Case 75953
'            Call Rotina_Erro(vbOKOnly, "ERRO_RASTREAMENTOLOTE_NAO_CADASTRADO1", gErr, objRastreamentoLote.sProduto, objRastreamentoLote.sCodigo, objRastreamentoLote.iFilialOP)
'
'        Case 75954
'            Call Rotina_Erro(vbOKOnly, "ERRO_MOVESTOQUE_NAO_CADASTRADO", gErr, objItemNF.lNumIntDoc)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156940)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Preenche_Tela_Recebimento(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim iIndice As Integer
Dim bCancel As Boolean

On Error GoTo Erro_Preenche_Tela_Recebimento

    'Lê os Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 42518
    
    'Lê a Distribuição dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 108732

    Call Limpa_Tela_NFiscal

    lErro = ComandoSeta_Fechar(Me.Name)

    '###################################
    'Inserido por Wagner
    'Tratamento de Erro 5 ao trazer o recebimento pelo Browse
    Set gobjNFiscal = objNFiscal
    '###################################

    TipoNFiscal.Text = objNFiscal.iTipoNFiscal
    Call TipoNFiscal_Validate(bCancel)

    Fornecedor.Text = objNFiscal.lFornecedor
    Call Fornecedor_Validate(bCancel)

    Filial.Text = objNFiscal.iFilialForn
    Call Filial_Validate(bCancel)

    Serie.Text = Desconverte_Serie_Eletronica(objNFiscal.sSerie)
    If ISSerieEletronica(objNFiscal.sSerie) Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    Call Serie_Validate(bSGECancelDummy)

    NFiscal.Text = objNFiscal.lNumNotaFiscal
    NumRecebimento.Text = objNFiscal.lNumRecebimento
    
    'Se a data não for nula coloca na Tela
    DataEmissao.PromptInclude = False
    If objNFiscal.dtDataEmissao <> DATA_NULA Then DataEmissao.Text = Format(objNFiscal.dtDataEmissao, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    DataEntrada.PromptInclude = False
    If objNFiscal.dtDataEntrada <> DATA_NULA Then DataEntrada.Text = Format(objNFiscal.dtDataEntrada, "dd/mm/yy")
    DataEntrada.PromptInclude = True

    For Each objItemNF In objNFiscal.ColItensNF
        objItemNF.dPercDesc = 0
        objItemNF.dValorDesconto = 0
    Next

    If objNFiscal.iCodTransportadora <> 0 Then
        Transportadora.Text = objNFiscal.iCodTransportadora
        Call Transportadora_Validate(bCancel)
    End If

    Placa.Text = objNFiscal.sPlaca
    PlacaUF.Text = objNFiscal.sPlacaUF
    VolumeEspecie = objNFiscal.lVolumeEspecie
    VolumeMarca = objNFiscal.lVolumeMarca
    If objNFiscal.lVolumeQuant > 0 Then VolumeQuant.Text = objNFiscal.lVolumeQuant

    Mensagem.Text = objNFiscal.sMensagemNota
    MensagemCorpo.Text = objNFiscal.sMensagemCorpoNota
    Observacao.Text = objNFiscal.sObservacao
    
    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = CStr(objNFiscal.dPesoLiq)
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = CStr(objNFiscal.dPesoBruto)
    
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 42519

    If objNFiscal.dValorOutrasDespesas > 0 Then ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Standard")
    If objNFiscal.dValorFrete > 0 Then ValorFrete.Text = Format(objNFiscal.dValorFrete, "Standard")
    If objNFiscal.dValorSeguro > 0 Then ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Standard")
    If objNFiscal.dValorTotal > 0 Then Total.Text = Format(objNFiscal.dValorTotal, "Standard")
    If objNFiscal.dValorProdutos > 0 Then SubTotal.Caption = Format(objNFiscal.dValorProdutos, "Standard")
    
    Call gobjTribTab.ValorFrete_Validate(bSGECancelDummy, objNFiscal.dValorFrete)
    Call gobjTribTab.ValorSeguro_Validate(bSGECancelDummy, objNFiscal.dValorSeguro)
    Call gobjTribTab.ValorDespesas_Validate(bSGECancelDummy, objNFiscal.dValorOutrasDespesas)
        
    For iIndice = 1 To objNFiscal.ColItensNF.Count
        lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, objNFiscal.ColItensNF.Item(iIndice).sProduto)
        If lErro <> SUCESSO Then gError 55654
    Next
    
    'Preenche o Grid com as Distribuições dos itens da Nota Fiscal
    lErro = gobjDistribuicao.Preenche_GridDistribuicao(objNFiscal)
    If lErro <> SUCESSO Then gError 108730

    iAlterado = 0

    Preenche_Tela_Recebimento = SUCESSO

    Exit Function

Erro_Preenche_Tela_Recebimento:

    Preenche_Tela_Recebimento = gErr

    Select Case gErr

        Case 42518, 42519, 55654, 108730, 108732

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156941)

    End Select

    Exit Function

End Function

Private Function Preenche_Tela_Recebimento2(objNFiscal As ClassNFiscal) As Long
'Preenche alguns campos da tela com os dados do recebimento passado em objNFiscal

Dim lErro As Long
Dim objItemNF As New ClassItemNF
Dim iIndice As Integer
Dim bCancel As Boolean

On Error GoTo Erro_Preenche_Tela_Recebimento2

    'Lê os Ítens
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 42611
    
    'Lê a Distribuição dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 108733

    'Limpa Todos os compos da Tela menos os do Cabeçalho
    Call Limpa_Tela_NFiscal1
    
    If objNFiscal.lNumRecebimento = 0 Then
        NumRecebimento.Text = ""
    Else
        NumRecebimento.Text = objNFiscal.lNumRecebimento
    End If
    
    Set gobjNFiscal = objNFiscal
    
    DataEntrada.PromptInclude = False
    If objNFiscal.dtDataEntrada <> DATA_NULA Then DataEntrada.Text = Format(objNFiscal.dtDataEntrada, "dd/mm/yy")
    DataEntrada.PromptInclude = True

    If objNFiscal.iCodTransportadora <> 0 Then
        Transportadora.Text = objNFiscal.iCodTransportadora
        Call Transportadora_Validate(bCancel)
    End If
    
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

    If objNFiscal.lNumNotaFiscal = 0 Then
        NFiscalInterna.Caption = ""
        NFiscal.Text = ""
    Else
        NFiscalInterna.Caption = CStr(objNFiscal.lNumNotaFiscal)
        NFiscal.Text = CStr(objNFiscal.lNumNotaFiscal)
    End If
    
    If objNFiscal.iTipoNFiscal = DOCINFO_NRFF Then
        objNFiscal.iTipoNFiscal = DOCINFO_NFEEFN
    ElseIf objNFiscal.iTipoNFiscal = DOCINFO_NRFP Then
        objNFiscal.iTipoNFiscal = DOCINFO_NFIEFN
    End If

    If objNFiscal.dValorTotal > 0 Then Total.Text = Format(objNFiscal.dValorTotal, "Standard")

    TipoNFiscal.Text = objNFiscal.iTipoNFiscal
    Call TipoNFiscal_Validate(bCancel)
    
    Fornecedor.Text = objNFiscal.lFornecedor
    Call Fornecedor_Validate(bCancel)
        
    Filial.Text = objNFiscal.iFilialForn
    Call Filial_Validate(bCancel)
    
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
    Observacao.Text = objNFiscal.sObservacao
    
    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = CStr(objNFiscal.dPesoLiq)
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = CStr(objNFiscal.dPesoBruto)

    'Preenche o Grid com os ítens lidos
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 42612

    If objNFiscal.dValorOutrasDespesas > 0 Then ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Standard")
    If objNFiscal.dValorFrete > 0 Then ValorFrete.Text = Format(objNFiscal.dValorFrete, "Standard")
    If objNFiscal.dValorSeguro > 0 Then ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Standard")
    If objNFiscal.dValorProdutos > 0 Then SubTotal.Caption = Format(objNFiscal.dValorProdutos, "Standard")
    
    Call gobjTribTab.ValorFrete_Validate(bSGECancelDummy, objNFiscal.dValorFrete)
    Call gobjTribTab.ValorSeguro_Validate(bSGECancelDummy, objNFiscal.dValorSeguro)
    Call gobjTribTab.ValorDespesas_Validate(bSGECancelDummy, objNFiscal.dValorOutrasDespesas)
        
    For iIndice = 1 To objNFiscal.ColItensNF.Count
        lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, objNFiscal.ColItensNF.Item(iIndice).sProduto)
        If lErro <> SUCESSO Then gError 55655
    Next
    
    'Preenche o Grid com as Distribuições dos itens da Nota Fiscal
    lErro = gobjDistribuicao.Preenche_GridDistribuicao(objNFiscal)
    If lErro <> SUCESSO Then gError 108731

    iAlterado = 0

    Preenche_Tela_Recebimento2 = SUCESSO

    Exit Function

Erro_Preenche_Tela_Recebimento2:

    Preenche_Tela_Recebimento2 = gErr

    Select Case gErr

        Case 42611, 42612, 55655, 108731, 108733

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156942)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_NFiscal1()
'Limpa a Tela NFiscalEntrada

Dim lErro As Long

    'Limpa o Grid
    Call Grid_Limpa(objGridItens)

    Call Grid_Limpa(objGridBenef)
    
    'retira o tab
    If TabStrip1.Tabs.Count = TAB_BENEFICIAMENTO Then
        TabStrip1.Tabs.Remove (TAB_BENEFICIAMENTO)
    End If
    
    'Limpa os campos que não são limpos pelo Limpa_Tela
    FilialFornBenef.Clear
    
    TipoFrete.ListIndex = 0
    ValorProdutos2.Caption = ""
    SubTotal.Caption = ""
    Transportadora.Text = ""
    Placa.Text = ""
    PlacaUF.Text = ""
    ValorFrete.Text = ""
    ValorSeguro.Text = ""
    ValorDespesas.Text = ""
    VolumeQuant.Text = ""
    
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
    
    Mensagem.Text = ""
    MensagemCorpo.Text = ""
    PesoLiquido.Text = ""
    PesoBruto.Text = ""
    Observacao.Text = ""
    
''*** incluidos p/tratamento de tributacao *******************************

    Call gobjTribTab.Limpa_Tela
    Call Trata_FilialForn
    
'*** fim tributacao
    
    lErro = ComandoSeta_Fechar(Me.Name)

End Sub

Private Sub Limpa_Tela_NFiscal()
'Limpa a Tela NFiscalEntrada

    Set gcolcolMedicoes = New Collection
    
    Call Limpa_Tela_NFiscal3

End Sub

Private Sub Limpa_Tela_NFiscal2()
'Limpa a Tela NFiscalEntrada, mas mantém a natureza e o tipo
    
Dim sNatureza As String
Dim sTipoNFiscal As String

    sNatureza = NaturezaOp.Text
    sTipoNFiscal = TipoNFiscal.Text
    
    Call Limpa_Tela_NFiscal

    NaturezaOp.Text = sNatureza
    TipoNFiscal.Text = sTipoNFiscal

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

    SubTotal_Calcula = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156943)

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

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGridItens)

End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer
Dim iIndice As Integer
Dim dSubTotal As Double
Dim lErro As Long

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
        
        gcolcolMedicoes.Remove GridItens.Row 'Inserido por Wagner
       
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
        If lErro <> SUCESSO Then gError 55588

        ''*** incluidos p/tratamento de tributacao *******************************

        lErro = gobjTribTab.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 42610

        '*** fim tributacao

        Call ValorTotal_Calcula
        
        If Not gbCarregandoTela Then
            lErro = gobjTribTab.Pesos_Recalcula()
            If lErro <> SUCESSO Then gError 42610
        End If

'distribuicao
'        'Rastreamento
'        lErro = gobjRastreamento.Exclusao_Item_Grid(iItemAtual)
'        If lErro <> SUCESSO Then gError 83266

        'distribuicao
        lErro = gobjDistribuicao.Exclusao_Item_GridDist(iItemAtual)
        If lErro <> SUCESSO Then gError 89528

        '#############################################################
        'Inserido por Wagner 23/03/2006
        lErro = CF2(Me, "NFE_Calcula_PesoKg_Total")
        If lErro <> SUCESSO Then gError 177160
        '#############################################################

    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr

        Case 42610, 55588, 83266, 89528, 177160

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156944)

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

Private Function Move_Dados_NFOriginal_Memoria(objNFiscal As ClassNFiscal, objTipoDocInfo As ClassTipoDocInfo) As Long
'Move os dados da NotaFiscalOriginal para a memória

'Alterado por Luiz em 25/01/02 - incluído tratamento para recolher dados da nota fiscal original
'quando a nota que está sendo gerada é uma nota de despesa com despacho aduaneiro

Dim lErro As Long
Dim objNFOriginal As New ClassNFiscal
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Dados_NFOriginal_Memoria

    'verifica se a Serie e o Número da NF original estão preenchidos
    If Len(Trim(NFiscalOriginal.ClipText)) = 0 Then gError 42758
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 And Codigo_Extrai(Trim(TipoNFiscal.Text)) <> DOCINFO_NFIEFCNFP Then gError 42757
    
    'Se for uma nota de despesas c/despacho aduaneiro
    If Codigo_Extrai(Trim(TipoNFiscal.Text)) = DOCINFO_FDDAI Or Codigo_Extrai(Trim(TipoNFiscal.Text)) = DOCINFO_NFSEFT Then
    
        'Verifica se o código do Fornecedor e da Filial estão preenchidos
        If Len(Trim(FornNFOrig.Text)) = 0 Then gError 94876
        If Len(Trim(FilialFornNFOrig.Text)) = 0 Then gError 94877
        
        'Guarda a série da nota fiscal original no obj
        objNFOriginal.sSerie = SerieNFiscalOriginal.Text
        'Guarda o número da nota fiscal original no obj
        objNFOriginal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
        
        'Guarda o código do fornecedor da nota fiscal original =>
        'o fornecedor da nota de importação é diferente do fornecedor da nota que está sendo lançada
        'por isso é necessário ler os seus dados e armazenar o código no obj
        
        'Guarda o nome reduzido em obj fornecedor
        objFornecedor.sNomeReduzido = FornNFOrig.Text
        
        'Lê os dados do fornecedor a partir do nome passado no obj
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 94907
        
        'Se não encontrou o fornecedor => erro
        If lErro = 6681 Then gError 94908
        
        'Guarda no obj o código do fornecedor da NF original
        objNFOriginal.lFornecedor = objFornecedor.lCodigo
        
        'Guarda a filial do fornecedor da nota fiscal original
        objNFOriginal.iFilialForn = Codigo_Extrai(FilialFornNFOrig.Text)
        
    End If
    
    If Codigo_Extrai(Trim(TipoNFiscal.Text)) = DOCINFO_FDDAI Then

        'O tipo da nota fiscal original só pode ser nota de importação
        objNFOriginal.iTipoNFiscal = DOCINFO_NFIEIMP

        'Lê a Nota Fiscal Original
        lErro = CF("NFiscal_Le_Num_Forn1", objNFOriginal)
        'If lErro <> SUCESSO And lErro <> 31467 Then gError 94878
        If lErro <> SUCESSO And lErro <> 55885 Then gError 94878
    
        If lErro <> SUCESSO Then
        
            'O tipo da nota fiscal original só pode ser nota de importação
            objNFOriginal.iTipoNFiscal = DOCINFO_NFIEIMPSE
    
            'Lê a Nota Fiscal Original
            lErro = CF("NFiscal_Le_Num_Forn1", objNFOriginal)
            If lErro <> SUCESSO And lErro <> 55885 Then gError 94878
        
'            'Se não encontrar --> Erro
'            If lErro = 55885 Then gError 42756
        
            If lErro <> SUCESSO Then
            
                'O tipo da nota fiscal original só pode ser nota de importação
                objNFOriginal.iTipoNFiscal = DOCINFO_NFIEIMPPC
        
                'Lê a Nota Fiscal Original
                lErro = CF("NFiscal_Le_Num_Forn1", objNFOriginal)
                If lErro <> SUCESSO And lErro <> 55885 Then gError 94878
            
                'Se não encontrar --> Erro
                If lErro = 55885 Then gError 42756
            
            End If
        
        End If
    
    ElseIf Codigo_Extrai(Trim(TipoNFiscal.Text)) = DOCINFO_NFIEFCNFP Then
    
        objNFiscal.sSerieNFPOrig = SerieNFiscalOriginal.Text
        objNFiscal.lNumNFPOrig = CLng(NFiscalOriginal.Text)
    
    Else
        
        objNFOriginal.sSerie = SerieNFiscalOriginal.Text
        objNFOriginal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
        
        If Codigo_Extrai(Trim(TipoNFiscal.Text)) <> DOCINFO_NFSEFT Then
        
            objNFOriginal.lFornecedor = objNFiscal.lFornecedor
            objNFOriginal.iFilialForn = objNFiscal.iFilialForn
            
            'Lê a Nota Fiscal Original
            lErro = CF("NFiscal_Le_Num_Forn", objNFOriginal, objTipoDocInfo.iTipo)
            If lErro <> SUCESSO And lErro <> 31467 Then gError 42753
            
        Else
        
            lErro = CF("NFiscal_Le_Num_Forn2", objNFOriginal)
            If lErro <> SUCESSO And lErro <> 92086 Then gError 42753
        
        End If
    
    '    objTipoDocInfo.iComplementar = DOCINFO_NORMAL
    '    lErro = CF("TiposDocInfo_Le_NomeTelaComplemetarTipo",objTipoDocInfo)
    '    If lErro <> SUCESSO And lErro <> 35103 Then gError 42754
    '    If lErro = 35103 Then gError 42759
    '
    '    objNFOriginal.iTipoNFiscal = objTipoDocInfo.iCodigo
    
        
        'Se não encontrar --> Erro
        If lErro <> SUCESSO Then gError 42756
    
    End If
    
    'guarda em objNfiscal o Numero Interno da NF Original
    objNFiscal.lNumIntNotaOriginal = objNFOriginal.lNumIntDoc

    Move_Dados_NFOriginal_Memoria = SUCESSO

    Exit Function

Erro_Move_Dados_NFOriginal_Memoria:

    Move_Dados_NFOriginal_Memoria = gErr

    Select Case gErr

        Case 42753, 42754, 42755, 94878, 94907

        Case 94876
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NFORIGINAL_NAO_PREENCHIDO", gErr)

        Case 94908
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case 94877
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NFORIGINAL_NAO_PREENCHIDO", gErr)
        
        Case 42756
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA1", gErr, objNFOriginal.sSerie, objNFOriginal.lNumNotaFiscal)

        Case 42757
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NFISCAL_ORIGINAL_NAO_PREENCHIDA", gErr)

        Case 42758
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NFISCAL_ORIGINAL_NAO_PREENCHIDO", gErr)

'        Case 42759
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPONFISCALORIGINAL_NAO_ENCONTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156945)

    End Select

    Exit Function

End Function

Private Function Valida_Grid_Parcela() As Long

Dim iIndice As Integer
Dim lErro As Long
Dim dSomaParcelas As Double
Dim dtDataVencimento As Date
Dim dValorPagar As Double

On Error GoTo Erro_Valida_Grid_Parcela

    'Verifica se GridParcelas foi preenchido
    If objGridParcelas.iLinhasExistentes = 0 Then Error 42642

    dSomaParcelas = 0

    'Loop no GridParcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        'Verifica se DataVencimento foi preenchida
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) = 0 Then Error 42648

        dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))

        'Verifica se DataVencimento é maior ou igual a DataEmissao
        If Len(Trim(DataEmissao.ClipText)) > 0 Then
            If dtDataVencimento < CDate(DataEmissao) Then Error 42643
        End If

        'Verifica a ordenação das Datas de Vencimento das Parcelas
        If iIndice > 1 Then
            If dtDataVencimento < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_Col)) Then Error 42644
        End If

        'Verifica se Valor da Parcela foi preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))) = 0 Then Error 42645

        'Verifica se Valor da Parcela é positivo
        lErro = Valor_Positivo_Critica(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
        If lErro <> SUCESSO Then Error 42646

        'Acumula Valor Parcela em dSomaParcelas
        dSomaParcelas = dSomaParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))

    Next

    lErro = gobjTribTab.ValorPagar_Calcula(dValorPagar)
    If lErro <> SUCESSO Then Error 42646

    'Compara Soma das Parcelas com Valor a Pagar
    If Abs(dSomaParcelas - dValorPagar) > DELTA_VALORMONETARIO Then Error 42647

    Valida_Grid_Parcela = SUCESSO

    Exit Function

Erro_Valida_Grid_Parcela:

    Valida_Grid_Parcela = Err

    Select Case Err

        Case 42642
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PARCELAS_GRAVAR", Err)

        Case 42643
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", Err, dtDataVencimento, DataEmissao.Text, iIndice)

        Case 42644
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_NAO_ORDENADA", Err)

        Case 42645
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORPARCELA_NAO_INFORMADO", Err, iIndice)

        Case 42646

        Case 42647
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_INVALIDA", Err, dSomaParcelas, dValorPagar)

        Case 42648
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_NAO_INFORMADA", Err, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156946)

    End Select
    
    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 39658

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridParcelas.Name

                lErro = Saida_Celula_GridParcelas(objGridInt)
                If lErro <> SUCESSO Then gError 42649

            'Se for o GridItens
            Case GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 42650
            
            Case GridMovimentos.Name
            
                lErro = Saida_Celula_GridMovimentos(objGridInt)
                If lErro <> SUCESSO Then gError 52520
        
            Case GridRastro.Name
            
                lErro = gobjRastreamento.Saida_Celula()
                If lErro <> SUCESSO Then gError 83269
        
            'distribuicao
            Case GridDist.Name
                lErro = gobjDistribuicao.Saida_Celula_Dist()
                If lErro <> SUCESSO Then gError 89533
        
        
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 42651

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 39658, 42649, 42650, 52520, 83269, 89533
        
        Case 42651
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156947)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridParcelas(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridParcelas

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Vencimento_Col
            lErro = Saida_Celula_DataVencimento(objGridInt)
            If lErro <> SUCESSO Then gError 42652

        Case iGrid_Valor_Col
            lErro = Saida_Celula_ValorParcela(objGridInt)
            If lErro <> SUCESSO Then gError 42653

        Case iGrid_Cobranca_Col
            lErro = Saida_Celula_Cobranca(objGridInt)
            If lErro <> SUCESSO Then gError 42654
            
        'Inicio da alteracao Daniel em 03/10/2001
        Case iGrid_Banco_Col
            lErro = Saida_Celula_Banco(objGridInt)
            If lErro <> SUCESSO Then gError 95160
                  
        Case iGrid_Portador_Col
            lErro = Saida_Celula_Portador(objGridInt)
            If lErro <> SUCESSO Then gError 95161
        'Fim da alteracao Daniel em 03/10/2001
            
        Case iGrid_CodigoDeBarras_Col
            lErro = Saida_Celula_CodigodeBarras(objGridInt)
            If lErro <> SUCESSO Then gError 95161
    
    End Select

    Saida_Celula_GridParcelas = SUCESSO

    Exit Function

Erro_Saida_Celula_GridParcelas:

    Saida_Celula_GridParcelas = gErr

    Select Case gErr

        Case 42652, 42653, 42654, 95160, 95161

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156948)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridMovimentos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridMovimentos
        
    Select Case GridMovimentos.Col

        Case iGrid_ProdutoBenef_Col

            lErro = Saida_Celula_ProdutoBenef(objGridInt)
            If lErro <> SUCESSO Then Error 52513

        Case iGrid_QuantidadeBenef_Col

            lErro = Saida_Celula_QuantidadeBenef(objGridInt)
            If lErro <> SUCESSO Then Error 52514

        Case iGrid_AlmoxarifadoBenef_Col

            lErro = Saida_Celula_AlmoxarifadoBenef(objGridInt)
            If lErro <> SUCESSO Then Error 52515

        Case iGrid_UnidadeMedBenef_Col

            lErro = Saida_Celula_UnidadeMedBenef(objGridInt)
            If lErro <> SUCESSO Then Error 52516

        Case iGrid_ContaContabilEstBenef_Col
            lErro = Saida_Celula_ContaContabilEst(objGridInt)
            If lErro <> SUCESSO Then Error 52517

        Case iGrid_ContaContabilProducaoBenef_Col
            lErro = Saida_Celula_ContaContabilProducao(objGridInt)
            If lErro <> SUCESSO Then Error 52518

    End Select
    
    Saida_Celula_GridMovimentos = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_GridMovimentos:

    Saida_Celula_GridMovimentos = Err

    Select Case Err

        Case 52513, 52514, 52515, 52516, 52517, 52518

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156949)

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
            If lErro <> SUCESSO Then gError 42656

        'DescricaoItem
        Case iGrid_DescProduto_Col
            lErro = Saida_Celula_DescricaoItem(objGridInt)
            If lErro <> SUCESSO Then gError 43300

        'UnidadeMedida
        Case iGrid_UnidadeMed_Col
            lErro = Saida_Celula_UM(objGridInt)
            If lErro <> SUCESSO Then gError 42658

        'Quantidade
        Case iGrid_Quantidade_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 42659

'distribuicao
'        'Almoxarifado
'        Case iGrid_Almoxarifado_Col
'            lErro = Saida_Celula_Almoxarifado(objGridInt)
'            If lErro <> SUCESSO Then Error 42660
        
        'Ccl
        Case iGrid_Ccl_Col
            lErro = Saida_Celula_Ccl(objGridInt)
            If lErro <> SUCESSO Then gError 49400

        'ValorUnitario
        Case iGrid_ValorUnitario_Col
            lErro = Saida_Celula_ValorUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 42661

        'PercDesc
        Case iGrid_PercDesc_Col
            lErro = Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then gError 42662

        'Desconto
        Case iGrid_Desconto_Col
            lErro = Saida_Celula_Desconto(objGridInt)
            If lErro <> SUCESSO Then gError 42663
            
        '###################################################
        'Inserido por Wagner
        Case iGrid_Contrato_Col
            lErro = Saida_Celula_Contrato(objGridInt)
            If lErro <> SUCESSO Then gError 132947
        
        Case iGrid_ItemContrato_Col
            lErro = Saida_Celula_ItemContrato(objGridInt)
            If lErro <> SUCESSO Then gError 132945
        '###################################################

        Case iGrid_DataCobranca_Col
            lErro = Saida_Celula_DataCobranca(objGridInt)
            If lErro <> SUCESSO Then gError 188438

        '############################################################
        'Inserido por Wagner 31/07/2006
        Case Else
            lErro = CF2(Me, "NFE_Saida_Celula_GridItens", objGridInt)
            If lErro <> SUCESSO Then gError 181417
        '############################################################

    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = Err

    Select Case gErr

        Case 42656, 42658, 42659, 42660, 42661, 42662, 42663, 43300, 49400, 132945, 132946, 132947, 181417, 188438 'Alterado por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156950)

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
        If lErro <> SUCESSO And lErro <> 5703 Then Error 49401

        If lErro = 5703 Then Error 49402

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 49426
    
    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = Err

    Select Case Err

        Case 49401
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 49402
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", Err, Ccl.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 49426
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156951)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescricaoItem(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescricaoItem

    Set objGridInt.objControle = DescricaoItem

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 43301

    Saida_Celula_DescricaoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoItem:

    Saida_Celula_DescricaoItem = Err

    Select Case Err

        Case 43301
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156952)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UM(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Unidade de Medida que está deixando de ser a corrente

Dim lErro As Long
Dim sUMAnterior As String
Dim objProduto As New ClassProduto
Dim objKitVenda As New ClassKitVenda

On Error GoTo Erro_Saida_Celula_UM

    Set objGridInt.objControle = UnidadeMed

'inicio distribuicao
    'recolhe a UM anteriormente escolhida
    sUMAnterior = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
    
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text
    
    'coloca no grid a UM atual selecionda
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    gobjNFiscal.ColItensNF(GridItens.Row).sUnidadeMed = UnidadeMed.Text
    
    If sUMAnterior <> UnidadeMed.Text Then
        
         '####################################################################
        'Inserido por Wagner 18/05/2006
        objProduto.sCodigo = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177583
        '####################################################################
        
        If gobjNFiscal.ColItensNF(GridItens.Row).iPossuiGrade <> MARCADO And objProduto.iKitVendaComp <> MARCADO Then
        
            'Tenta fazer a distribuição automatica p\ o item de acordo com a nova UM
            lErro = gobjDistribuicao.Distribuicao_Processa()
            If lErro <> SUCESSO Then gError 89511
            
        Else
        
            If objProduto.iKitVendaComp = MARCADO Then
            
                '####################################################################
                'Inserido por Wagner 18/05/2006
                objKitVenda.sProduto = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
                objKitVenda.sUM = UnidadeMed.Text
                objKitVenda.dQuantidade = StrParaDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
            
                lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 177584
                '####################################################################
    
            End If
            lErro = gobjDistribuicao.Distribuicao_Processa_Grade()
            If lErro <> SUCESSO Then gError 89602
        
            Call gobjDistribuicao.Atualiza_Grid_Distribuicao(gobjNFiscal.ColItensNF.Item(GridItens.Row))
        End If
    End If
    
'fim distribuicao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 35083
    
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 177161
    End If

    '#############################################################
    'Inserido por Wagner 23/03/2006
    If sUMAnterior <> UnidadeMed.Text Then
        lErro = CF2(Me, "NFE_Calcula_PesoKg", GridItens.Row)
        If lErro <> SUCESSO Then gError 177161
    End If
    '#############################################################

    If GridItens.Row <= objGridItens.iLinhasExistentes Then
        If gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade.Count = 0 Then
            lErro = gobjRastreamento.Atualiza_Troca_UM(GridItens.Row)
            If lErro <> SUCESSO Then gError 83206
        Else
            lErro = gobjRastreamento.Atualiza_Troca_UM_Grade(GridItens.Row, sUMAnterior, UnidadeMed.Text, objProduto)
            If lErro <> SUCESSO Then gError 83206
        End If
    End If

    Saida_Celula_UM = SUCESSO

    Exit Function

Erro_Saida_Celula_UM:

    Saida_Celula_UM = gErr

    Select Case gErr

        Case 35083, 83206, 89511, 89602, 177583, 177584, 177161
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156953)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 42739

    'Limpa a Tela
    Call Limpa_Tela_NFiscal

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 42739

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156954)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Move_GridParcelas_Memoria(colParcelas As ColParcelaPagar) As Long
'Move para a memória os dados existentes no Grid

Dim iIndice As Integer
Dim objParcelaPag As ClassParcelaPagar
Dim lErro As Long

On Error GoTo Erro_Move_GridParcelas_Memoria

    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        Set objParcelaPag = New ClassParcelaPagar

        'Preenche objParcelaPag com a linha do GridParcelas
        objParcelaPag.iNumParcela = iIndice
        objParcelaPag.dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        objParcelaPag.dtDataVencimentoReal = CDate(GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col))
        objParcelaPag.dValor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
        objParcelaPag.iBancoCobrador = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Banco_Col))
        objParcelaPag.iPortador = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Portador_Col))

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col))) = 0 Then
            'Se estiver vazio usamos o Tipo Cobrança DEFAULT
            objParcelaPag.iTipoCobranca = TIPO_COBRANCA_CARTEIRA
        Else
            objParcelaPag.iTipoCobranca = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col))
        End If

        If GridParcelas.TextMatrix(iIndice, iGrid_Suspenso_Col) = "1" Then
            objParcelaPag.iStatus = STATUS_SUSPENSO
        Else
            objParcelaPag.iStatus = STATUS_ABERTO
        End If

        'retirar os pontos separadores para gravar sem eles
        objParcelaPag.sCodigoDeBarras = Trim(Replace(GridParcelas.TextMatrix(iIndice, iGrid_CodigoDeBarras_Col), ".", ""))
        
        'Adiciona objParcelaPag à coleção colParcelas
        With objParcelaPag
            colParcelas.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iPortador, .iProxSeqBaixa, .iTipoCobranca, .iBancoCobrador, .sNossoNumero, .sCodigoDeBarras
        End With
    Next

    Move_GridParcelas_Memoria = SUCESSO

    Exit Function

Erro_Move_GridParcelas_Memoria:

    Move_GridParcelas_Memoria = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156955)

    End Select

    Exit Function

End Function

Private Function Preenche_GridParcelas(objNFiscal As ClassNFiscal) As Long

Dim iLinha As Integer
Dim objParcelaPagar As ClassParcelaPagar
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Preenche_GridParcelas

    'Preenche as linhas do Grid Parcelas com os dados de cada Parcela
    For Each objParcelaPagar In objNFiscal.ColParcelaPagar

        iLinha = iLinha + 1

        GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col) = objParcelaPagar.iNumParcela
        GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_Col) = Format(objParcelaPagar.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_VenctoReal_Col) = Format(objParcelaPagar.dtDataVencimentoReal, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objParcelaPagar.dValor, "Standard")

        For iIndice = 0 To TipoCobranca.ListCount - 1
            If TipoCobranca.ItemData(iIndice) = objParcelaPagar.iTipoCobranca Then
                GridParcelas.TextMatrix(iLinha, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                Exit For
            End If
        Next
        
        'Inicio da alteracao Daniel
        For iIndice = 0 To comboCobrador.ListCount - 1
            If comboCobrador.ItemData(iIndice) = objParcelaPagar.iBancoCobrador Then
                GridParcelas.TextMatrix(iLinha, iGrid_Banco_Col) = comboCobrador.List(iIndice)
                Exit For
            End If
        Next
        
        For iIndice = 0 To comboPortador.ListCount - 1
            If comboPortador.ItemData(iIndice) = objParcelaPagar.iPortador Then
                GridParcelas.TextMatrix(iLinha, iGrid_Portador_Col) = comboPortador.List(iIndice)
                Exit For
            End If
        Next
        'Fim da alteracao Daniel

        If objParcelaPagar.iStatus = STATUS_SUSPENSO Then
            GridParcelas.TextMatrix(iLinha, iGrid_Suspenso_Col) = "1"
        Else
            GridParcelas.TextMatrix(iLinha, iGrid_Suspenso_Col) = "0"
        End If

        CodigodeBarras.PromptInclude = False
        CodigodeBarras.Text = objParcelaPagar.sCodigoDeBarras
        CodigodeBarras.PromptInclude = True
        GridParcelas.TextMatrix(iLinha, iGrid_CodigoDeBarras_Col) = CodigodeBarras.Text
        
    Next

    'Faz o número de linhas existentes do Grid ser igual ao número de Parcelas
    objGridParcelas.iLinhasExistentes = iLinha

    lErro = CF2(Me, "NFFE_Preenche_GridParcelas", objNFiscal)
    If lErro <> SUCESSO Then gError 181415

    'Faz refresh nas checkboxes
    Call Grid_Refresh_Checkbox(objGridParcelas)

    Preenche_GridParcelas = SUCESSO

    Exit Function

Erro_Preenche_GridParcelas:

    Preenche_GridParcelas = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156956)

    End Select

    Exit Function

End Function

''Private Function Parcelas_Calcula(dValor As Double, iNumParcelas As Integer, colValorParcelas As Collection) As Long
'''Calcula os valores das Parcelas a partir do Valor e Número de Parcelas
'''Devolve em colValorParcelas
''
''Dim dValorParcela As Double
''Dim iResto As Integer
''Dim iIndice As Integer
''
''On Error GoTo Erro_Parcelas_Calcula
''
''    'Faz divisão inteira do valor multiplicado por 100 pelo número
''    'de parcelas para trabalhar com os centésimos
''    dValorParcela = Fix((dValor * 100) / iNumParcelas)
''
''    'Calcula o resto da divisão inteira
''    iResto = (dValor * 100) Mod iNumParcelas
''
''    'Atualiza dValorParcela para o valor definitivo
''    dValorParcela = dValorParcela / 100
''
''    'Acrescentar valores das parcelas na coleção
''    For iIndice = 1 To iNumParcelas - iResto
''
''        colValorParcelas.Add dValorParcela
''
''    Next
''
''    'Soma 0.01 ao Valor da Parcela
''    dValorParcela = dValorParcela + 0.01
''
''    'Se a divisão não foi exata acrescentar as "iResto" últimas parcelas adicionadas de 0.01
''    For iIndice = 1 To iResto
''
''        colValorParcelas.Add dValorParcela
''
''    Next
''
''    Parcelas_Calcula = SUCESSO
''
''    Exit Function
''
''Erro_Parcelas_Calcula:
''
''    Parcelas_Calcula = Err
''
''    Select Case Err
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156957)
''
''    End Select
''
''    Exit Function
''
''End Function
''
''Function Parcelas_DatasVencimento(objCondicaoPagto As ClassCondicaoPagto, dtDataEmissao As Date, colDataVencimento As Collection) As Long
'''Calcula as datas de vencimento de cada uma das Parcelas do Grid
''
''Dim lErro As Long
''Dim dtDataVencimento As Date
''Dim iIndice As Integer
''Dim iDiferenca As Integer
''
''On Error GoTo Erro_Parcelas_DatasVencimento
''
''    'Se a condição for mensal
''    If objCondicaoPagto.iMensal <> False Then
''
''        'Calcula a data com intervalo para a primeira parcela
''        dtDataVencimento = dtDataEmissao + objCondicaoPagto.iDiasParaPrimeiraParcela
''
''        'Se o dia dessa data for menor que o dia do mês de pagamento da Parcela incrementa o mês
''        iIndice = 1
''
''        Do While iIndice <= objCondicaoPagto.iNumeroParcelas
''
''            If objCondicaoPagto.iDiaDoMes > Day(dtDataVencimento) Then
''                Do While (Month(dtDataVencimento + 1) = Month(dtDataVencimento)) And (Day(dtDataVencimento) <> objCondicaoPagto.iDiaDoMes)
''                    dtDataVencimento = dtDataVencimento + 1
''                Loop
''            Else
''                iDiferenca = Abs(Day(dtDataVencimento - objCondicaoPagto.iDiaDoMes))
''                dtDataVencimento = dtDataVencimento - iDiferenca
''                dtDataVencimento = DateAdd("m", 1, dtDataVencimento)
''
''            End If
''
''            'Guarda a data de vencimento em colDataVencimento
''            colDataVencimento.Add dtDataVencimento
''
''            dtDataVencimento = DateAdd("m", 1, dtDataVencimento)
''            iIndice = iIndice + 1
''        Loop
''
''    Else 'Se a condição não for mensal
''
''        'Calcula a primeira data de vencimento
''        dtDataVencimento = dtDataEmissao + objCondicaoPagto.iDiasParaPrimeiraParcela
''
''        'Loop para acrescentar as datas na coleção
''        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas
''            colDataVencimento.Add dtDataVencimento
''            dtDataVencimento = dtDataVencimento + objCondicaoPagto.iIntervaloParcelas
''        Next
''
''    End If
''
''    Parcelas_DatasVencimento = SUCESSO
''
''    Exit Function
''
''Erro_Parcelas_DatasVencimento:
''
''    Parcelas_DatasVencimento = Err
''
''    Select Case Err
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156958)
''
''    End Select
''
''    Exit Function
''
''End Function

Public Function Preenche_GridItens(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com os ítens da Nota Fiscal

Dim lErro  As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF
Dim sProdutoEnxuto As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sCclMascarado As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_Preenche_GridItens

    iIndice = 0

    'Para cada ítem da Coleção
    For Each objItemNF In objNFiscal.ColItensNF

        iIndice = iIndice + 1

        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 42760

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
        
        objProduto.sCodigo = objNFiscal.ColItensNF(iIndice).sProduto
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177926
        
        sCclMascarado = ""
        
        'Formata Ccl
        If Trim(objItemNF.sCcl) <> "" Then

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_RetornaCclEnxuta(objItemNF.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then Error 52876
        
            'Preenche o campo Ccl
            Ccl.PromptInclude = False
            Ccl.Text = sCclMascarado
            Ccl.PromptInclude = True
                    
            'Joga o Ccl no grid
            GridItens.TextMatrix(iIndice, iGrid_Ccl_Col) = Ccl.Text

        End If

        
        'Preenche o Grid
        '****** IF INCLUÍDO PARA TRATAMENTO DE GRADE ***************
        If objItemNF.colItensRomaneioGrade.Count > 0 And objProduto.iKitVendaComp <> MARCADO Then GridItens.TextMatrix(iIndice, 0) = "# " & GridItens.TextMatrix(iIndice, 0)
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objItemNF.sDescricaoItem
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemNF.sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemNF.dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col) = Format(objItemNF.dPrecoUnitario, FORMATO_PRECO_UNITARIO_EXTERNO)
        If objItemNF.dPercDesc <> 0 Then GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(objItemNF.dPercDesc, "Percent")
        If objItemNF.dValorDesconto <> 0 Then GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objItemNF.dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objItemNF.dValorTotal, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotalB_Col) = Format(objItemNF.dValorTotal + objItemNF.dValorDesconto, "Standard")

        '###################################################
        'Inserido por Wagner
        lErro = Traz_Contrato_Tela(objItemNF, iIndice)
        If lErro <> SUCESSO Then gError 132969
        '###################################################
        
    Next

    '###################################################
    'Inserido por Wagner 04/08/2006
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objNFiscal.objProjetoInfo, objNFiscal.lNumIntDoc, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181676
    '###################################################

    '###################################################
    'Inserido por Wagner 31/07/2006
    lErro = CF2(Me, "NFE_Preenche_GridItens", objNFiscal)
    If lErro <> SUCESSO Then gError 181415
    '###################################################

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice

    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = Err

    Select Case Err

        Case 42760
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, objItemNF.sProduto)
        
        Case 42761, 42762, 132969, 177926, 181415, 181676 'Alterado por Wagner
        
        Case 52876
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", Err, objItemNF.sCcl)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156959)

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
Dim lTamanho As Long
Dim colAlocacoes As ColAlocacoesItemNF
Dim sCclFormatada As String
Dim iCclPreenchida As Integer
Dim objItemNFGlb As ClassItemNF
Dim objItemNFNovo As ClassItemNF

On Error GoTo Erro_Move_GridItens_Memoria

    'Para cada linha existente do Grid
    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItemNF = New ClassItemNF

        lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 42763

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
            If lErro <> SUCESSO Then gError 49412

        Else
            sCclFormatada = ""
        End If

        With objItemNF
            Set objItemNFNovo = objNFiscal.ColItensNF.Add(0, iIndice, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, DATA_NULA, .sDescricaoItem, 0, 0, 0, 0, 0, colAlocacoes, .iAlmoxarifado, .sAlmoxarifadoNomeRed, sCclFormatada, STATUS_LANCADO, 0, "", 0, 0, 0, "", 0, 0)
        End With
        
        '********************* TRATAMENTO DE GRADE *****************
        lErro = gobjDistribuicao.Move_DistribuicaoGrade_Memoria(gobjNFiscal.ColItensNF(iIndice))
        If lErro <> SUCESSO Then gError 86375
        
        Call Move_ItensGrade_Tela(objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade, gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade)

        '#################################################
        'Inserido por Wagner
        lErro = Move_Contrato_Memoria(objNFiscal, objItemNF, iIndice)
        If lErro <> SUCESSO Then gError 132940
        '#################################################

        '###################################################
        'Inserido por Wagner 31/07/2006
        lErro = CF2(Me, "NFE_Move_GridItem_Memoria", objItemNFNovo, iIndice)
        If lErro <> SUCESSO Then gError 181416
        '###################################################
        
        Set objItemNFNovo.objInfoAdicDocItem = gobjNFiscal.ColItensNF(iIndice).objInfoAdicDocItem

    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 42763, 42764, 49412, 86375, 132940, 181416 'Inserido por Wagner

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156960)

    End Select

    Exit Function

End Function

Private Sub ValorReal_Calcula(dQuantidade As Double, dPrecoUnitario As Double, dPercentDesc As Double, dDesconto As Double, dValorReal As Double)
'Calcula o Valor Real

Dim dValorTotal As Double
Dim dPercDesc1 As Double
Dim dPercDesc2 As Double

    dValorTotal = Arredonda_Moeda(dPrecoUnitario * dQuantidade)

     'Se o Percentual Desconto estiver preenchido e ValorTotal positivo
    If dPercentDesc > 0 And dValorTotal > 0 Then

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
Dim objFilial As New ClassFilialFornecedor
Dim sContaTela As String
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim sContaMascarada As String
Dim objItem As New ClassTribItemNFAux
Dim dQuantidadeConvertida As Double
Dim dQuantidade As Double
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

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
        
        'Caso o mnemônico encontrado na coluna débito ou na coluna crédito, seja Quant_Beneficiada ==> o sistema utilizará o escaninho Beneficiamento para calcular o custo da transação
        Case ESCANINHO_CUSTO_BENEF
            For iLinha = 1 To objGridBenef.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add ESCANINHO_NOSSO_EM_BENEF
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        'Caso o mnemônico encontrado na coluna débito ou na coluna crédito NÃO seja Quant_Beneficiada ==> o sistema utilizará o escaninho Nossa Disponível para calcular o custo da transação
        Case ESCANINHO_CUSTO
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add ESCANINHO_NOSSO
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        'Retorna a quantidade dos produtos localizados no tab Beneficiamento
        Case QUANT_BENEFICIADA
            For iLinha = 1 To objGridBenef.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_QuantidadeBenef_Col)) > 0 Then
                    
                    'converte a UM utilizada no grid Beneficiamento para a UMEstoque
                    'define os parâmetros que serão passados para a função UMEstoque_Conversao
                    sProduto = GridMovimentos.TextMatrix(iLinha, iGrid_ProdutoBenef_Col)
                    sUM = GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMedBenef_Col)
                    dQuantidade = CDbl(GridMovimentos.TextMatrix(iLinha, iGrid_QuantidadeBenef_Col))
                    
                    'testa se a unidade de medida no grid Beneficiamento foi informada
                    If Len(Trim(sUM)) > 0 Then
                    
                        'chama a função que converte a UM
                        lErro = CF("UMEstoque_Conversao", sProduto, sUM, dQuantidade, dQuantidadeConvertida)
                        If lErro <> SUCESSO Then gError 79016
                    
                    'se a unidade de media não foi informada=> erro
                    Else
                        gError 79019
                    End If
                                        
                    'passa a quantidade convertida para o mnemônico
                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        'Retorna o código dos produtos localizados no tab Beneficiamento
        Case PRODUTO_BENEFICIADO
            For iLinha = 1 To objGridBenef.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_ProdutoBenef_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_ProdutoBenef_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        'Retorna a conta contábil de estoque de cada produto no tab Beneficiamento
        Case CTACONTABILESTBENEF
            For iLinha = 1 To objGridBenef.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilEstBenef_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilEstBenef_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        'Retorna a conta contábil de produção de cada produto no tab Beneficiamento
        Case CTACONTABILPRODBENEF
            For iLinha = 1 To objGridBenef.iLinhasExistentes
                If Len(GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilProducaoBenef_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridMovimentos.TextMatrix(iLinha, iGrid_ContaContabilProducaoBenef_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
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
                        If lErro <> SUCESSO And lErro <> 25060 Then gError 183177

                        'Se não encontrou ===> Erro
                        If lErro = SUCESSO Then

                            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
                            
                            lErro = CF("Produto_Formata", objMnemonicoValor.vParam(1), sProdutoFormatado, iPreenchido)
                            If iPreenchido <> PRODUTO_PREENCHIDO Then gError 183178

                            objEstoqueProduto.sProduto = sProdutoFormatado
                            
                            'Lê a conta contábil do Produto no Almoxarifado
                            lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
                            If lErro <> SUCESSO And lErro <> 49991 Then gError 183179

                            If Len(Trim(objEstoqueProduto.sContaContabil)) > 0 Then

                                sContaMascarada = String(STRING_CONTA, 0)

                                lErro = Mascara_MascararConta(objEstoqueProduto.sContaContabil, sContaMascarada)
                                If lErro <> SUCESSO Then gError 183180

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

        Case PRECO_TOTAL
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case PRECO_UNITARIO
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
            If Len(ICMSSubstValor1.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ICMSSubstValor1.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case VALOR_PRODUTOS
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

        Case VALOR_TOTAL
            If Len(Total.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(Total.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case VALOR_PARCELA
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                If Len(GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_Valor_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case SUSPENSO1
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                'Coloca em objMnemonicoValor.colValor 1 ou 0
                objMnemonicoValor.colValor.Add CInt(GridParcelas.TextMatrix(iLinha, iGrid_Suspenso_Col))
            Next


        Case CTAFILIALFORNECEDOR1
                 
            If Len(Trim(Filial.Text)) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 57990
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 57991
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
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
                    If lErro <> SUCESSO Then gError 183432
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183433
        
                    objMnemonicoValor.colValor.Add objProduto.iControleEstoque
                    
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        Case CONTA_DESP_EST_FORN
            
            If Len(Trim(Filial.Text)) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 183434
                
                If objFilial.sContaDespesa <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaDespesa, sContaTela)
                    If lErro <> SUCESSO Then gError 183435
                
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
                    If lErro <> SUCESSO Then gError 183534
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183535
        
                    If objProduto.sContaContabil <> "" Then
                    
                        lErro = Mascara_RetornaContaTela(objProduto.sContaContabil, sContaTela)
                        If lErro <> SUCESSO Then gError 183814
                    
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
                    If lErro <> SUCESSO Then gError 183558
                
                    objProdutoCategoria.sCategoria = objMnemonicoValor.vParam(1)
                    
                    objProdutoCategoria.sProduto = sProdutoFormatado
                    
                    Set objCatProdItem = New ClassCategoriaProdutoItem
                    
                    lErro = CF("ProdutoCategoria_Le_Item", objProdutoCategoria, objCatProdItem)
                    If lErro <> SUCESSO And lErro <> 183548 Then gError 183559
                    
                    If lErro = SUCESSO Then
                        
                        Set objPlanoConta = New ClassPlanoConta
                    
                        'le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta
                        lErro = CF("PlanoConta_Le_ContaSimples", objCatProdItem.dvalor1, objPlanoConta)
                        If lErro <> SUCESSO And lErro <> 5451 Then gError 183560
                    
                        If lErro = SUCESSO Then
                    
                            If objPlanoConta.sConta <> "" Then
                            
                                lErro = Mascara_RetornaContaTela(objPlanoConta.sConta, sContaTela)
                                If lErro <> SUCESSO Then gError 183815
                            
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
        
        Case Else
            lErro = gobjTribTab.Calcula_Mnemonico(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 39659

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 79019
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_BENEF_NAO_PREENCHIDA", gErr, iLinha)
        
        Case 39659
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case 57920, 57939, 57961, 57972, 57990, 57991, 75447, 75449, 75450, 75452, 79943, 183082 To 183086, 183152, 183153, 183173 To 183183, 183432 To 183435, 183534, 183535, 183558, 183559, 183560, 183814, 183815
        
        Case 75448
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO1", gErr, objAlmoxarifado.sNomeReduzido)
        
        Case 75451
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO", objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156961)

    End Select

    Exit Function

End Function

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

Public Sub BotaoPlanoConta_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPlanoConta_Click

    If GridMovimentos.Row = 0 Then Error 52415

    If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col) = "" Then Error 52416

    sConta = String(STRING_CONTA, 0)

    If GridMovimentos.Col = iGrid_ContaContabilEstBenef_Col Then

        lErro = CF("Conta_Formata", ContaContabilEst.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then Error 52417

    ElseIf GridMovimentos.Col = iGrid_ContaContabilProducaoBenef_Col Then


        lErro = CF("Conta_Formata", ContaContabilProducao.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then Error 52418

    End If

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaESTLista
    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case Err

        Case 52415
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)

        Case 52416
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case 52417, 52418

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156962)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objGridBenef.objControle.Name <> "ContaContabilEst" And objGridBenef.objControle.Name <> "ContaContabilProducao" Then
        Me.Show
        Exit Sub
    End If

    If objPlanoConta.sConta <> "" Then

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 52419

        objGridBenef.objControle.PromptInclude = False
        objGridBenef.objControle.Text = sContaEnxuta
        objGridBenef.objControle.PromptInclude = True

        GridMovimentos.TextMatrix(GridMovimentos.Row, GridMovimentos.Col) = objGridBenef.objControle.Text

    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case Err

        Case 52419
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156963)

    End Select

    Exit Sub

End Sub

Public Sub ContaContabilEst_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaContabilEst_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBenef)

End Sub

Public Sub ContaContabilEst_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBenef)

End Sub

Public Sub ContaContabilEst_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBenef.objControle = ContaContabilEst
    lErro = Grid_Campo_Libera_Foco(objGridBenef)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Function Saida_Celula_ContaContabilEst(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabilEst

    Set objGridInt.objControle = ContaContabilEst

    If Len(Trim(ContaContabilEst.ClipText)) > 0 Then

        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabilEst.ClipText, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 52420

        If lErro = SUCESSO Then

            sContaFormatada = objPlanoConta.sConta

            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 52421

            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaMascarada
            ContaContabilEst.PromptInclude = True

        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then

            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilEst.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then Error 52422

            'conta não cadastrada
            If lErro = 5700 Then Error 52423

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 52424

    Saida_Celula_ContaContabilEst = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilEst:

    Saida_Celula_ContaContabilEst = Err

    Select Case Err

        Case 52420, 52422, 52424
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52421
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52423
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilEst.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156964)

    End Select

    Exit Function

End Function

Public Sub ContaContabilProducao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaContabilProducao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBenef)

End Sub

Public Sub ContaContabilProducao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBenef)

End Sub

Public Sub ContaContabilProducao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBenef.objControle = ContaContabilProducao
    lErro = Grid_Campo_Libera_Foco(objGridBenef)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ContaContabilProducao(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabilProducao

    Set objGridInt.objControle = ContaContabilProducao

    If Len(Trim(ContaContabilProducao.ClipText)) > 0 Then

        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabilProducao.ClipText, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 52425

        If lErro = SUCESSO Then

            sContaFormatada = objPlanoConta.sConta

            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 52426

            ContaContabilProducao.PromptInclude = False
            ContaContabilProducao.Text = sContaMascarada
            ContaContabilProducao.PromptInclude = True

        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then

            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilProducao.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then Error 52427

            'conta não cadastrada
            If lErro = 5700 Then Error 52428

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 52429

    Saida_Celula_ContaContabilProducao = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilProducao:

    Saida_Celula_ContaContabilProducao = Err

    Select Case Err

        Case 52425, 52427, 52429
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52426
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52428
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilProducao.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156965)

    End Select

    Exit Function

End Function

Public Sub BotaoProdutosBenef_Click()

Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim objItemOP As New ClassItemOP

On Error GoTo Erro_BotaoProdutosBenef_Click

    'se não selecionou nenhuma linha do grid entao - - - -> Erro
    If GridMovimentos.Row = 0 Then Error 52430

    ' Formata o Produto
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 52431

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objProduto.sCodigo = sProdutoFormatado

    'Lista de produtos que  podem participar da producao e inventariados
    Call Chama_Tela("ProdutoEstoquePCPLista", colSelecao, objProduto, objEventoProdutoBenef)

   Exit Sub

Erro_BotaoProdutosBenef_Click:

    Select Case Err

        Case 52430
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)

        Case 52431

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156966)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoBenef_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim objTipoDeProduto As New ClassTipoDeProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    If (GridMovimentos.Row > 0) Then

        'Verifica se o Produto está preenchido
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col))) = 0 Then


            lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then Error 52432

            sProdutoMascarado = String(STRING_PRODUTO, 0)

            lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then Error 52433

            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 52434

            If lErro = 28030 Then Error 52435

            Produto.PromptInclude = False
            ProdutoBenef.Text = sProdutoMascarado
            Produto.PromptInclude = True

            If Not (Me.ActiveControl Is ProdutoBenef) Then

                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col) = sProdutoMascarado
        
                If Trim(objProduto.sContaContabilProducao) = "" Then
        
                    objTipoDeProduto.iTipo = objProduto.iTipo
        
                    lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
                    If lErro <> SUCESSO And lErro <> 22531 Then Error 52436
        
                    If lErro = 22531 Then Error 52437
        
                    objProduto.sContaContabilProducao = objTipoDeProduto.sContaContabil
        
                End If
        
                'Preenche a Linha do Grid
                lErro = ProdutoLinha_PreencheBenef(objProduto)
                If lErro <> SUCESSO Then Error 52438
        
                lErro = Preenche_ContaContabilEst(ProdutoBenef.Text)
                If lErro <> SUCESSO Then Error 52439
        
                lErro = QuantDisponivel_Calcula1(sProdutoMascarado, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col), objProduto)
                If lErro <> SUCESSO Then Error 52440

            End If
            
        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case Err

        Case 52432, 52434, 52436, 52438, 52439, 52440

        Case 52433
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)
        
        Case 52435
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case 52437
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", Err, objTipoDeProduto.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156967)

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
        If lErro <> SUCESSO Then Error 52440

        'preenche  a ContaContabilProducao
        ContaContabilProducao.PromptInclude = False
        ContaContabilProducao.Text = sContaEnxuta
        ContaContabilProducao.PromptInclude = True

        'preenche Conta De Producao
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilProducaoBenef_Col) = ContaContabilProducao.Text

    End If

    'Unidade de Medida
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMedBenef_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_DescricaoBenef_Col) = objProduto.sDescricao

    'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
    lErro = CF("AlmoxarifadoPadrao_Le_NomeReduzido", objProduto.sCodigo, sAlmoxarifadoPadrao)
    If lErro <> SUCESSO Then Error 52441

    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col) = sAlmoxarifadoPadrao

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridMovimentos.Row - GridMovimentos.FixedRows) = objGridBenef.iLinhasExistentes Then
        objGridBenef.iLinhasExistentes = objGridBenef.iLinhasExistentes + 1
    End If

    ProdutoLinha_PreencheBenef = SUCESSO

    Exit Function

Erro_ProdutoLinha_PreencheBenef:

    ProdutoLinha_PreencheBenef = Err

    Select Case Err

        Case 52440
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objProduto.sContaContabilProducao)
        
        Case 52441
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156968)

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

    If (GridMovimentos.Row = 0) Then Error 52442
    
    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 52443

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then Error 52444

    colSelecao.Add sProdutoFormatado

    Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoqueBenef)

    Exit Sub

Erro_BotaoEstoqueBenef_Click:

    Select Case Err

        Case 52442
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)
        
        Case 52443
                
        Case 52444
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156969)

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
        If lErro <> SUCESSO Then Error 52445

        'Verifica se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Preenche o Nome do Almoxarifado
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            AlmoxarifadoBenef.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            'Calcula a Quantidade Disponível nesse Almoxarifado
            lErro = QuantDisponivel_Calcula(sCodProduto, objEstoqueProduto.sAlmoxarifadoNomeReduzido)
            If lErro <> SUCESSO Then Error 52446

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoqueBenef_evselecao:

    Select Case Err

        Case 52445, 52446

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156970)

    End Select

    Exit Sub

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

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156971)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_ProdutoBenef(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim objTipoDeProduto As New ClassTipoDeProduto

On Error GoTo Erro_Saida_Celula_ProdutoBenef

    Set objGridInt.objControle = ProdutoBenef

    If Len(Trim(ProdutoBenef.ClipText)) <> 0 Then

        lErro = CF("Produto_Critica_Estoque", ProdutoBenef.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25077 Then Error 52448

        If lErro = 25077 Then Error 52449

        If Trim(objProduto.sContaContabilProducao) = "" Then

            objTipoDeProduto.iTipo = objProduto.iTipo

            lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
            If lErro <> SUCESSO And lErro <> 22531 Then Error 52450

            If lErro = 22531 Then Error 52451

            objProduto.sContaContabilProducao = objTipoDeProduto.sContaContabil

        End If


    End If

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        If objProduto.iPCP = PRODUTO_PCP_NAOPODE Then Error 52452

        lErro = ProdutoLinha_PreencheBenef(objProduto)
        If lErro <> SUCESSO Then Error 52453

        lErro = Preenche_ContaContabilEst(ProdutoBenef.Text)
        If lErro <> SUCESSO Then Error 52454

        lErro = QuantDisponivel_Calcula1(ProdutoBenef.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col), objProduto)
        If lErro <> SUCESSO Then Error 52455

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 52456

    Saida_Celula_ProdutoBenef = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdutoBenef:

    Saida_Celula_ProdutoBenef = Err

    Select Case Err

        Case 52448, 52450, 52453, 52454, 52455, 52456
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52449
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", ProdutoBenef.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = ProdutoBenef.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 52451
             Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", Err, objTipoDeProduto.iTipo)
             Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52452
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PCP", Err, ProdutoBenef.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156972)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantidadeBenef(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_QuantidadeBenef

    Set objGridInt.objControle = QuantidadeBenef

    If Len(Trim(QuantidadeBenef.ClipText)) > 0 Then

        lErro = Valor_Positivo_Critica(QuantidadeBenef.Text)
        If lErro <> SUCESSO Then Error 52457

            If Len(Trim(QuantDisponivelBenef.Caption)) <> 0 Then

                lErro = Testa_QuantRequisitada(CDbl(QuantidadeBenef.Text))
                If lErro <> SUCESSO Then Error 52458

            End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 52460

    Saida_Celula_QuantidadeBenef = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadeBenef:

    Saida_Celula_QuantidadeBenef = Err

    Select Case Err

        Case 52457, 52458, 52460
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156973)

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
        If lErro <> SUCESSO Then Error 52461

        lErro = TP_Almoxarifado_Produto_Grid(sProdutoFormatado, AlmoxarifadoBenef, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25130 And lErro <> 25135 Then Error 52462

        If lErro = 25130 Then Error 52463

        If lErro = 25135 Then Error 52464

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col) = objAlmoxarifado.sNomeReduzido
        
        lErro = QuantDisponivel_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col), AlmoxarifadoBenef.Text)
        If lErro <> SUCESSO Then Error 52465

    Else

        'Limpa a Quantidade Disponível da Tela
        QuantDisponivelBenef.Caption = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 52466

    'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
    lErro = Preenche_ContaContabilEst(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col))
    If lErro <> SUCESSO Then Error 52467

    Saida_Celula_AlmoxarifadoBenef = SUCESSO

    Exit Function

Erro_Saida_Celula_AlmoxarifadoBenef:

    Saida_Celula_AlmoxarifadoBenef = Err

    Select Case Err

        Case 52461, 52462, 52465, 52466, 52467
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52463

            vbMsg = Rotina_Aviso(vbYesNo, "ERRO_ALMOXARIFADO_INEXISTENTE", AlmoxarifadoBenef.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = AlmoxarifadoBenef.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 52464

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(AlmoxarifadoBenef.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(AlmoxarifadoBenef.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156974)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UnidadeMedBenef(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_Saida_Celula_UnidadeMedBenef

    Set objGridInt.objControle = UnidadeMedBenef

    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_UnidadeMedBenef_Col) = UnidadeMedBenef.Text

    If Len(UnidadeMedBenef.Text) > 0 Then

        lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col))
        If lErro <> SUCESSO Then Error 55638
    
        'Se a quantidade está preenchida e não se trata de estorno
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_QuantidadeBenef_Col))) <> 0 Then

            dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_QuantidadeBenef_Col))

            'Testa a Quantidade requisitada
            lErro = Testa_QuantRequisitada(dQuantidade)
            If lErro <> SUCESSO Then Error 55639

        End If
    
    Else
    
        QuantDisponivelBenef.Caption = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 52468

    Saida_Celula_UnidadeMedBenef = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMedBenef:

    Saida_Celula_UnidadeMedBenef = Err

    Select Case Err

        Case 52468, 55638, 55639
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156975)

    End Select

    Exit Function

End Function

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

    If (GridMovimentos.Row <> iLinhaAntiga) Then

        'Guarda a Linha usada
        iLinhaAntiga = GridMovimentos.Row

        lErro = QuantDisponivel_Calcula1(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col))
        If lErro <> SUCESSO Then Error 52470


    End If

    Exit Sub

Erro_GridMovimentos_RowColChange:

    Select Case Err

        Case 52470

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156976)

    End Select

    Exit Sub

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

Public Sub QuantidadeBenef_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBenef)

End Sub

Public Sub QuantidadeBenef_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBenef)

End Sub

Public Sub QuantidadeBenef_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBenef.objControle = QuantidadeBenef
    lErro = Grid_Campo_Libera_Foco(objGridBenef)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UnidadeMedBenef_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBenef)

End Sub

Public Sub UnidadeMedBenef_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBenef)

End Sub

Public Sub UnidadeMedBenef_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBenef.objControle = UnidadeMedBenef
    lErro = Grid_Campo_Libera_Foco(objGridBenef)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function QuantDisponivel_Calcula(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_QuantDisponivel_Calcula

    If (objProduto Is Nothing) Then

        lErro = QuantDisponivel_Calcula1(sProduto, sAlmoxarifado)
        If lErro <> SUCESSO Then Error 55614
        
    Else
    
        lErro = QuantDisponivel_Calcula1(sProduto, sAlmoxarifado, objProduto)
        If lErro <> SUCESSO Then Error 55615

    End If

    lErro = Testa_Quantidade()
    If lErro <> SUCESSO Then Error 55616

    QuantDisponivel_Calcula = SUCESSO

    Exit Function

Erro_QuantDisponivel_Calcula:

    QuantDisponivel_Calcula = Err

    Select Case Err

        Case 55614, 55615, 55616

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 156977)

    End Select

    Exit Function

End Function

Private Function QuantDisponivel_Calcula1(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long
'descobre a quantidade disponivel e coloca na tela

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim dQuantTotal As Double
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto

On Error GoTo Erro_QuantDisponivel_Calcula1

    QuantDisponivelBenef.Caption = ""

    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMedBenef_Col))) > 0 Then

        'Verifica se o produto está preenchido
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 55617
    
        If GridMovimentos.Row >= GridMovimentos.FixedRows And Len(Trim(sAlmoxarifado)) <> 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            If (objProduto Is Nothing) Then
            
                Set objProduto = New ClassProduto
    
                objProduto.sCodigo = sProdutoFormatado
    
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then Error 55618
    
                If lErro = 28030 Then Error 55619
    
            End If
    
            objAlmoxarifado.sNomeReduzido = sAlmoxarifado
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then Error 55620
    
            If lErro = 25056 Then Error 55621
    
            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
            objEstoqueProduto.sProduto = sProdutoFormatado
    
            'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
            lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 21306 Then Error 55622
    
            'Se não encontrou EstoqueProduto no Banco de Dados
            If lErro = 21306 Then
            
                 QuantDisponivelBenef.Caption = Formata_Estoque(0)
    
            Else
                sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMedBenef_Col)
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
                If lErro <> SUCESSO Then Error 55623
        
                QuantDisponivelBenef.Caption = Formata_Estoque(objEstoqueProduto.dQuantBenef * dFator)
    
            End If
    
        Else
    
            'Limpa a Quantidade Disponível da Tela
            QuantDisponivelBenef.Caption = ""
    
        End If

    End If
    
    QuantDisponivel_Calcula1 = SUCESSO

    Exit Function

Erro_QuantDisponivel_Calcula1:

    QuantDisponivel_Calcula1 = Err

    Select Case Err

        Case 55617, 55618, 55620, 55622, 55623

        Case 55619
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case 55621
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", Err, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 156978)

    End Select

    Exit Function

End Function

Private Function Testa_Quantidade() As Long

Dim dQuantidade As Double
Dim lErro As Long

On Error GoTo Erro_Testa_Quantidade

    If GridMovimentos.Row >= GridMovimentos.FixedRows Then

        'Se a quantidade está preenchida e não se trata de linha estornada
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_QuantidadeBenef_Col))) <> 0 Then

            dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_QuantidadeBenef_Col))

            'Calcula a Quantidade requisitada
            lErro = Testa_QuantRequisitada(dQuantidade)
            If lErro <> SUCESSO Then Error 55624

        End If
    
    End If

    Testa_Quantidade = SUCESSO

    Exit Function

Erro_Testa_Quantidade:

    Testa_Quantidade = Err

    Select Case Err

        Case 55624

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 156979)

    End Select
    
    Exit Function
    
End Function

Private Function Testa_QuantRequisitada(ByVal dQuantAtual As Double) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProdutoAtual As String
Dim sAlmoxarifado As String
Dim sAlmoxarifadoAtual As String
Dim sUnidadeAtual As String
Dim sUnidadeProd As String
Dim dQuantidadeProd As String
Dim dFator As Double
Dim objProduto As New ClassProduto
Dim dQuantTotal As Double

On Error GoTo Erro_Testa_QuantRequisitada

    sProdutoAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col)
    sAlmoxarifadoAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col)
    sUnidadeAtual = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMedBenef_Col)

    If Len(sProdutoAtual) > 0 And Len(sAlmoxarifadoAtual) > 0 And Len(sUnidadeAtual) > 0 Then

        lErro = CF("Produto_Formata", sProdutoAtual, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 55625

        objProduto.sCodigo = sProdutoFormatado

        'Lê o produto para saber qual é a sua ClasseUM
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 55626
    
        If lErro = 28030 Then Error 55627
    
        For iIndice = 1 To objGridBenef.iLinhasExistentes
    
            'Não pode somar a Linha atual
            If GridMovimentos.Row <> iIndice Then
    
                sCodProduto = GridMovimentos.TextMatrix(iIndice, iGrid_ProdutoBenef_Col)
                sAlmoxarifado = GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxarifadoBenef_Col)
    
                lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then Error 55628
    
                'Verifica se há outras Requisições de Produto no mesmo Almoxarifado
                If UCase(sAlmoxarifado) = UCase(sAlmoxarifadoAtual) And UCase(objProduto.sCodigo) = UCase(sProdutoFormatado) Then
    
                    'Verifica se há alguma QuanTidade informada
                    If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_QuantidadeBenef_Col))) <> 0 Then
    
                        sUnidadeProd = GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMedBenef_Col)
    
                        dQuantidadeProd = CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_QuantidadeBenef_Col))
    
                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUnidadeProd, sUnidadeAtual, dFator)
                        If lErro <> SUCESSO Then Error 55629
    
                        dQuantTotal = dQuantTotal + (dQuantidadeProd * dFator)
    
                    End If
    
                End If
    
            End If
    
        Next
    
        dQuantTotal = dQuantTotal + dQuantAtual

        If dQuantTotal > CDbl(QuantDisponivelBenef.Caption) Then Error 55630

    End If

    Testa_QuantRequisitada = SUCESSO

    Exit Function

Erro_Testa_QuantRequisitada:

    Testa_QuantRequisitada = Err

    Select Case Err

        Case 55625, 55626, 55628, 55629

        Case 55627
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, sCodProduto)

        Case 55630
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_REQ_MAIOR", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 156980)

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
    If lErro <> SUCESSO And lErro <> 25060 Then Error 52484

    If lErro = 25060 Then Error 52485

    'Formata o Produto para BD
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 52486

    objEstoqueProduto.sProduto = sProdutoFormatado
    objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo

    lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 49991 Then Error 52487

    If lErro <> 49991 Then

        lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then Error 52488

        ContaContabilEst.PromptInclude = False
        ContaContabilEst.Text = sContaEnxuta
        ContaContabilEst.PromptInclude = True

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilEstBenef_Col) = ContaContabilEst.Text

    Else

        'Preenche em branco a conta de estoque no grid
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilEstBenef_Col) = ""

    End If

    Preenche_ContaContabilEst = SUCESSO

    Exit Function

Erro_Preenche_ContaContabilEst:

    Preenche_ContaContabilEst = Err

        Select Case Err

            Case 52484, 52486, 52487

            Case 52488
                Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objEstoqueProduto.sContaContabil)

            Case 52485
                 Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", Err, objAlmoxarifado.sNomeReduzido)

            Case Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156981)

        End Select

        Exit Function

End Function

Public Sub FornecedorBenef_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorBenefAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FornecedorBenef_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedorBenef As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_FornecedorBenef_Validate

    If iFornecedorBenefAlterado = 1 Then

        If Len(Trim(FornecedorBenef.Text)) > 0 Then
    
            'Tenta ler o FornecedorBenef (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(FornecedorBenef, objFornecedorBenef, iCodFilial)
            If lErro <> SUCESSO Then Error 52489
    
            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedorBenef, colCodigoNome)
            If lErro <> SUCESSO Then Error 52490
    
            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", FilialFornBenef, colCodigoNome)
    
            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", FilialFornBenef, iCodFilial)
    
    
        ElseIf Len(Trim(FornecedorBenef.Text)) = 0 Then
    
            FilialFornBenef.Clear
    
        End If
        
        iFornecedorBenefAlterado = 0

    End If
    
    Exit Sub

Erro_FornecedorBenef_Validate:

    Cancel = True

    Select Case Err

        Case 52489, 52490

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156982)

    End Select

    Exit Sub

End Sub

Public Sub FilialFornBenef_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialFornBenef_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialFornBenef.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialFornBenef.Text = FilialFornBenef.List(FilialFornBenef.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialFornBenef, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 52533

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(FornecedorBenef.Text)) = 0 Then Error 52534

        sFornecedor = FornecedorBenef.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 52535

        If lErro = 18272 Then Error 52536

        'coloca na tela
        FilialFornBenef.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 52537

    Exit Sub

Erro_FilialFornBenef_Validate:

    Cancel = True
    
    Select Case Err

        Case 52533, 52535
        
        Case 52536
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 52534
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, FornecedorBenef.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 52537
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", Err, FilialFornBenef.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156983)

    End Select

    Exit Sub

End Sub

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

On Error GoTo Erro_Move_GridMovimentos_Memoria

    Set objNFiscal.objMovEstoque = New ClassMovEstoque

    'Para cada linha existente do Grid
    For iIndice = 1 To objGridBenef.iLinhasExistentes

        'Formata o Produto para BD
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_ProdutoBenef_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 52491

        'Pega o Código do Almoxarifado
        Set objAlmoxarifado = New ClassAlmoxarifado

        objAlmoxarifado.sNomeReduzido = GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxarifadoBenef_Col)

        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then Error 52492

        If lErro = 25056 Then Error 52493

        If GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilEstBenef_Col) <> "" Then

            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilEstBenef_Col), sContaFormatadaEst, iContaPreenchida)
            If lErro <> SUCESSO Then Error 52494

        Else
            sContaFormatadaEst = ""
        End If

        If GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilProducaoBenef_Col) <> "" Then

            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilProducaoBenef_Col), sContaFormatadaProducao, iContaPreenchida)
            If lErro <> SUCESSO Then Error 52495

        Else
            sContaFormatadaProducao = ""
        End If

        objNFiscal.objMovEstoque.colItens.Add 0, 0, 0, 0, sProdutoFormatado, GridMovimentos.TextMatrix(iIndice, iGrid_DescricaoBenef_Col), GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMedBenef_Col), CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_QuantidadeBenef_Col)), objAlmoxarifado.iCodigo, GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxarifadoBenef_Col), 0, "", 0, "", "", sContaFormatadaProducao, sContaFormatadaEst, 0, Nothing, Nothing, DATA_NULA

    Next

    Move_GridMovimentos_Memoria = SUCESSO

    Exit Function

Erro_Move_GridMovimentos_Memoria:

    Move_GridMovimentos_Memoria = Err

    Select Case Err

        Case 52491, 52492, 52494, 52495

        Case 52493
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", Err, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156984)

    End Select

    Exit Function

End Function

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
        If lErro <> SUCESSO Then Error 52496

        ProdutoBenef.PromptInclude = False
        ProdutoBenef.Text = sProdutoEnxuto
        ProdutoBenef.PromptInclude = True

        'Preenche o Grid
        GridMovimentos.TextMatrix(iIndice, iGrid_ProdutoBenef_Col) = ProdutoBenef.Text
        GridMovimentos.TextMatrix(iIndice, iGrid_DescricaoBenef_Col) = objItemMovEstoque.sProdutoDesc
        GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMedBenef_Col) = objItemMovEstoque.sSiglaUM
        GridMovimentos.TextMatrix(iIndice, iGrid_QuantidadeBenef_Col) = Formata_Estoque(objItemMovEstoque.dQuantidade)
        GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxarifadoBenef_Col) = objItemMovEstoque.sAlmoxarifadoNomeRed

        'preenche contaEst no grid
        If objItemMovEstoque.sContaContabilEst <> "" Then

            sContaEnxutaEst = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilEst, sContaEnxutaEst)
            If lErro <> SUCESSO Then Error 52497

            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxutaEst
            ContaContabilEst.PromptInclude = True

            GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilEstBenef_Col) = ContaContabilEst.Text

        End If

        'preenche contaEst no grid
        If objItemMovEstoque.sContaContabilAplic <> "" Then

            sContaEnxutaProducao = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilAplic, sContaEnxutaProducao)
            If lErro <> SUCESSO Then Error 52498

            ContaContabilProducao.PromptInclude = False
            ContaContabilProducao.Text = sContaEnxutaProducao
            ContaContabilProducao.PromptInclude = True

            GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilProducaoBenef_Col) = ContaContabilProducao.Text

        End If

    Next

    'Atualiza o número de linhas existentes
    objGridBenef.iLinhasExistentes = iIndice

    Preenche_GridMovimentos = SUCESSO

    Exit Function

Erro_Preenche_GridMovimentos:

    Preenche_GridMovimentos = Err

    Select Case Err

        Case 52496
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, objItemMovEstoque.sProduto)

        Case 52497
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objItemMovEstoque.sContaContabilEst)

        Case 52498
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objItemMovEstoque.sContaContabilAplic)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156985)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridMovimentos(objGridInt As AdmGrid) As Long
'Inicializa o Grid
Dim lErro As Long

On Error GoTo Erro_Inicializa_GridMovimentos
    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Conta Contábil de Estoque")
    objGridInt.colColuna.Add ("Conta Contábil de Produção")
    

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ProdutoBenef.Name)
    objGridInt.colCampo.Add (DescricaoItemBenef.Name)
    objGridInt.colCampo.Add (UnidadeMedBenef.Name)
    objGridInt.colCampo.Add (QuantidadeBenef.Name)
    objGridInt.colCampo.Add (AlmoxarifadoBenef.Name)
    objGridInt.colCampo.Add (ContaContabilEst.Name)
    objGridInt.colCampo.Add (ContaContabilProducao.Name)
    
    'Colunas do Grid
    iGrid_ProdutoBenef_Col = 1
    iGrid_DescricaoBenef_Col = 2
    iGrid_UnidadeMedBenef_Col = 3
    iGrid_QuantidadeBenef_Col = 4
    iGrid_AlmoxarifadoBenef_Col = 5
    iGrid_ContaContabilEstBenef_Col = 6
    iGrid_ContaContabilProducaoBenef_Col = 7
    
    'Grid do GridInterno
    objGridInt.objGrid = GridMovimentos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 9

    'Largura da primeira coluna
    GridMovimentos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridMovimentos = SUCESSO

    Exit Function

Erro_Inicializa_GridMovimentos:

    Inicializa_GridMovimentos = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156986)

    End Select

    Exit Function

End Function

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

Public Function Calcula_Valores(iLinha As Integer) As Long
'recalcula os valores de desconto, percentual de desconto e valor total

Dim sProduto As String
Dim lErro As Long
Dim lTamanho As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorReal As Double
Dim dQuantidade As Double

On Error GoTo Erro_Calcula_Valores

    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))

    lErro = CF2(Me, "NFE_PrecoTotal_Calcula", iLinha, dQuantidade)
    If lErro <> SUCESSO Then gError 181259

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

        If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))) > 0 Then dPrecoUnitario = CDbl(GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col))
        If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col))) > 0 Then dDesconto = CDbl(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col))

        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dPrecoUnitario, dPercentDesc, dDesconto, dValorReal)

        'Coloca o Desconto calculado na tela
        If dDesconto > 0 Then
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(dDesconto, "Standard")
        Else
            GridItens.TextMatrix(iLinha, iGrid_Desconto_Col) = ""
        End If

        'Coloca o valor Real em Valor Total
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col) = Format(dValorReal, "Standard")

        'Calcula o Valor Real
        Call ValorReal_Calcula(dQuantidade, dPrecoUnitario, 0, 0, dValorReal)

        'Coloca o valor Real em Valor Total
        GridItens.TextMatrix(iLinha, iGrid_PrecoTotalB_Col) = Format(dValorReal, "Standard")
    End If

    lErro = SubTotal_Calcula()
    If lErro <> SUCESSO Then gError 55570

    Calcula_Valores = SUCESSO
    
    Exit Function
    
Erro_Calcula_Valores:

    Calcula_Valores = gErr
    
    Select Case gErr

        Case 55570, 181259

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 156988)

    End Select

    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ENTRADA_NF_FATURA_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    
'    If giTipoVersao = VERSAO_FULL Then
        Caption = "Nota Fiscal Fatura de Entrada"
'    Else
'        Caption = "Nota Fiscal de Entrada"
'    End If
    
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "NFiscalFatEntrada"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Sub Unload(objme As Object)
    
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

Public Function ValorTotal_Calcula() As Long

Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    If gbCarregandoTela = False Then
    
        lErro = gobjTribTab.AtualizarTributacao()
        If lErro <> SUCESSO Then Error 56969
        
        Call ValorDescontoItens_Calcula
    
        If iEmitenteNF = DOCINFO_EMPRESA Then
        
            'atualizar o total da nota
            Call Total_Calcula(0)
            
        End If
    
        'atualizar o tab de pagamentos
        Call Cobranca_Automatica
            
    End If
    
    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = Err

    Select Case Err

        Case 56969

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156989)

    End Select

    Exit Function

End Function

Private Function Cobranca_Automatica() As Long

Dim lErro As Long, objCondicaoPagto As New ClassCondicaoPagto
Dim iCodigo As Integer

On Error GoTo Erro_Cobranca_Automatica

    If PagamentoAutomatico.Value = vbChecked And Len(Trim(Total.Text)) <> 0 Then
    
        objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)
        
        If objCondicaoPagto.iCodigo <> 0 Then
        
            'Lê Condicao Pagamento no BD
            lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
            If lErro <> SUCESSO And lErro <> 19205 Then Error 56973
            If lErro <> SUCESSO Then Error 56974
            
            lErro = GridParcelas_Preenche(objCondicaoPagto)
            If lErro <> SUCESSO Then Error 56975
    
        End If
    
    End If
    
    Cobranca_Automatica = SUCESSO
     
    Exit Function
    
Erro_Cobranca_Automatica:

    Cobranca_Automatica = Err
     
    Select Case Err
          
        Case 56973, 56975
        
        Case 56974 'nao achou no bd a condicao de pagto selecionada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156990)
     
    End Select
     
    Exit Function

End Function

Private Function Trata_FilialForn() As Long

Dim lErro As Long
Dim objFilialForn As New ClassFilialFornecedor

On Error GoTo Erro_Trata_FilialForn
    
    If gbCarregandoTela = False Then
    
        objFilialForn.iCodFilial = Codigo_Extrai(Filial.Text)
        
        If objFilialForn.iCodFilial <> 0 Then
        
            'Lê a Filial
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilialForn)
            If lErro <> SUCESSO Then gError 56976
            
            Call gobjTribTab.FilialFornecedor_Alterada(objFilialForn.lCodFornecedor, objFilialForn.iCodFilial)
        
            lErro = CF2(Me, "NFCF_Fornecedor_Filial_Validate", objFilialForn)
            If lErro <> SUCESSO Then gError 131952
        
        End If
        
    End If
    
    Trata_FilialForn = SUCESSO
     
    Exit Function
    
Erro_Trata_FilialForn:

    Trata_FilialForn = gErr
     
    Select Case gErr
          
        Case 56976, 131952 'inserido por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156991)
     
    End Select
     
    Exit Function

End Function

Private Function Critica_CamposNecessarios_ConsultaTitulo() As Long
'Verifica se os campos necessários para encontrar consultar um título gerado por uma NFFatEntrada foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposNecessarios_ConsultaTitulo

    'Verifica se o Fornecedor foi preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 79720
    
    'Verifica se a filial do Fornecedor foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 79721
    
    'Verifica se a Série foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 79722
    
    'Verifica se o Número da Nota foi preenchido
    If Len(Trim(NFiscal.Text)) = 0 Then gError 79723
    
    'Verifica se a data de emissão da nota foi preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 79724
    
    'Verifica se a data de entrada da nota foi preenchida
    If Len(Trim(DataEntrada.ClipText)) = 0 Then gError 79725
    
    'Verifica se o tipo da Nota fiscal foi selecionado
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 79726
    
    Critica_CamposNecessarios_ConsultaTitulo = SUCESSO
    
    Exit Function

Erro_Critica_CamposNecessarios_ConsultaTitulo:

    Critica_CamposNecessarios_ConsultaTitulo = gErr
    
    Select Case gErr
    
        Case 79720
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 79721
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 79722
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 79723
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
        
        Case 79724
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 79725
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_NAO_PREENCHIDA", gErr)
        
        Case 79726
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156992)
    
    End Select
        
    Exit Function
        
End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is NaturezaOp Then
            Call NaturezaLabel_Click
        ElseIf Me.ActiveControl Is Serie Then
            Call SerieLabel_Click
        ElseIf Me.ActiveControl Is NFiscal Then
            Call NFiscalLabel_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is FornNFOrig Then
            Call LabelFornNFOrig_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
'distribuicao
        ElseIf Me.ActiveControl Is gobjDistribuicao.AlmoxDist Then
            Call gobjDistribuicao.BotaoLocalizacaoDist_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        ElseIf Me.ActiveControl Is SerieNFiscalOriginal Then
            Call SerieOriginalLabel_Click
        ElseIf Me.ActiveControl Is NFiscalOriginal Then
            Call NFiscalOriginalLabel_Click
        ElseIf Me.ActiveControl Is FornecedorBenef Then
            Call FornecedorBenefLabel_Click
        ElseIf Me.ActiveControl Is ProdutoBenef Then
            Call BotaoProdutosBenef_Click
        ElseIf Me.ActiveControl Is AlmoxarifadoBenef Then
            Call BotaoEstoqueBenef_Click
        ElseIf Me.ActiveControl Is ContaContabilProducao Or Me.ActiveControl Is ContaContabilEst Then
            Call BotaoPlanoConta_Click
        ElseIf Me.ActiveControl Is Contrato Then
            Call BotaoItemContrato_Click
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        Else
            Call gobjRastreamento.objUserControl_KeyDown(KeyCode, Shift)
        End If
    Else
    
        Call CF2(Me, "NFFE_UserControl_KeyDown", KeyCode, Shift)
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

Public Property Get ImageStatus() As Object
     Set ImageStatus = objUserControl.Controls("ImageStatus")
End Property

Public Sub comboPortador_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub comboCobrador_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Saida_Celula_Portador(objGridInt As AdmGrid) As Long
    
Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_Portador

    Set objGridInt.objControle = comboPortador

    'Verifica se o Portador foi preenchido
    If Len(Trim(comboPortador.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If comboPortador.Text <> comboPortador.List(comboPortador.ListIndex) Then

            'Seleciona o Portador
            lErro = Combo_Seleciona(comboPortador, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 57868

            If lErro = 6730 Then Error 57866
            If lErro = 6731 Then Error 57867

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 57869

    Saida_Celula_Portador = SUCESSO

    Exit Function

Erro_Saida_Celula_Portador:

    Saida_Celula_Portador = Err

    Select Case Err

        Case 57866
            Call Rotina_Erro(vbOKOnly, "ERRO_PORTADOR_NAO_CADASTRADO1", Err, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 57867
            Call Rotina_Erro(vbOKOnly, "ERRO_PORTADOR_NAO_CADASTRADO2", Err, comboPortador.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 57868, 57869
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156993)

    End Select

    Exit Function

End Function

Private Function Carrega_Portadores() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_Portadores

    'Lê o código e a descrição de todos os Portadores Ativos
    lErro = CF("Portadores_Le_CodigosNomesRed", colCodigoDescricao)
    If lErro <> SUCESSO Then Error 57858

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o ítem na Combo de Portadores
        comboPortador.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        comboPortador.ItemData(comboPortador.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_Portadores = SUCESSO

    Exit Function

Erro_Carrega_Portadores:

    Carrega_Portadores = Err

    Select Case Err

        Case 57858

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156994)

    End Select

    Exit Function

End Function

Public Sub comboPortador_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub comboPortador_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub comboPortador_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub comboPortador_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = comboPortador
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Property Get comboPortador() As Object
     Set comboPortador = objUserControl.Controls("comboPortador")
End Property

Public Property Get comboCobrador() As Object
     Set comboCobrador = objUserControl.Controls("comboCobrador")
End Property

Public Property Get CodigodeBarras() As Object
     Set CodigodeBarras = objUserControl.Controls("CodigodeBarras")
End Property

Private Function Saida_Celula_Banco(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_Banco

    Set objGridInt.objControle = comboCobrador

    'Verifica se o Banco foi preenchido
    If Len(Trim(comboCobrador.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If comboCobrador.Text <> comboCobrador.List(comboCobrador.ListIndex) Then

            'Seleciona o Banco
            lErro = Combo_Seleciona(comboCobrador, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 57864

            If lErro = 6730 Then Error 57862
            If lErro = 6731 Then Error 57863

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 57865

    Saida_Celula_Banco = SUCESSO

    Exit Function

Erro_Saida_Celula_Banco:

    Saida_Celula_Banco = Err

    Select Case Err

        Case 57862
            Call Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_CADASTRADO", Err, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 57863
            Call Rotina_Erro(vbOKOnly, "ERRO_BANCO_INEXISTENTE1", Err, comboCobrador.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 57864, 57865
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156995)

    End Select

    Exit Function

End Function

Private Function Carrega_Bancos() As Long
'Carrega a combo de Cobrador

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Bancos

    'Leitura dos códigos e descrições dos Bancos BD
    lErro = CF("Cod_Nomes_Le", "Bancos", "CodBanco", "NomeReduzido", STRING_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then Error 57856

   'Preenche ComboBox com código e nome dos Bancos
    For iIndice = 1 To colCodigoNome.Count
        Set objCodigoNome = colCodigoNome(iIndice)
        comboCobrador.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        comboCobrador.ItemData(comboCobrador.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_Bancos = SUCESSO

    Exit Function

Erro_Carrega_Bancos:

    Carrega_Bancos = Err

    Select Case Err

        Case 57856 'Tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156996)

    End Select

    Exit Function

End Function

Public Sub comboCobrador_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Public Sub comboCobrador_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub comboCobrador_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub comboCobrador_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = comboCobrador
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Habilita_Desabilita_Forn_NFOrig()
'habilita/desabilita os campos relativos a cliente/fornecedor associados a nota fiscal original

    If (Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_FDDAI) Or (Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_NFSEFT) Then
    
        FornNFOrig.Enabled = True
        LabelFornNFOrig.Enabled = True
        FilialFornNFOrig.Enabled = True
        LabelFilialFornNFOrig.Enabled = True
    
    Else
    
        FornNFOrig.Enabled = False
        LabelFornNFOrig.Enabled = False
        FilialFornNFOrig.Enabled = False
        LabelFilialFornNFOrig.Enabled = False
        
    End If

End Sub

Public Sub FornNFOrig_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornNFOrigAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub FornNFOrig_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_FornNFOrig_Validate
   
    If iFornNFOrigAlterado = 1 Then
   
        If Len(Trim(FornNFOrig.Text)) > 0 Then
    
            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(FornNFOrig, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 94870
    
            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 94871
    
            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", FilialFornNFOrig, colCodigoNome)
    
            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", FilialFornNFOrig, iCodFilial)
    
    
        ElseIf Len(Trim(FornNFOrig.Text)) = 0 Then
    
            FilialFornNFOrig.Clear
    
        End If
    
        iFornNFOrigAlterado = 0
    
    End If
    
    Exit Sub

Erro_FornNFOrig_Validate:

    Cancel = True

    Select Case gErr

        Case 94870, 94871

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156997)

    End Select

    Exit Sub

End Sub

Public Sub LabelFornNFOrig_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = FornNFOrig.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornNFOrig)

    Exit Sub

End Sub

Private Sub objEventoFornNFOrig_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    FornNFOrig.Text = objFornecedor.lCodigo
    Call FornNFOrig_Validate(bCancel)

    Me.Show

End Sub

Public Sub FilialFornNFOrig_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub FilialFornNFOrig_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialFornNFOrig_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialFornNFOrig.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialFornNFOrig.Text = FilialFornNFOrig.List(FilialFornNFOrig.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialFornNFOrig, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 94872

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(FornNFOrig.Text)) = 0 Then gError 94873

        sFornecedor = FornNFOrig.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 94874

        If lErro = 18272 Then gError 94875

        'coloca na tela
        FilialFornNFOrig.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 94876

    Exit Sub

Erro_FilialFornNFOrig_Validate:

    Cancel = True

    Select Case gErr

        Case 94872, 94874
        
        Case 94873
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 94875
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, FornNFOrig.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 94876
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, FilialFornNFOrig.Text)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156998)

    End Select

    Exit Sub

End Sub

Public Property Get FilialFornNFOrig() As Object
     Set FilialFornNFOrig = objUserControl.Controls("FilialFornNFOrig")
End Property

Public Property Get LabelFilialFornNFOrig() As Object
     Set LabelFilialFornNFOrig = objUserControl.Controls("LabelFilialFornNFOrig")
End Property

Public Property Get FornNFOrig() As Object
     Set FornNFOrig = objUserControl.Controls("FornNFOrig")
End Property

Public Property Get LabelFornNFOrig() As Object
     Set LabelFornNFOrig = objUserControl.Controls("LabelFornNFOrig")
End Property

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156999)
            
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


Function Grid_Possui_Grade(ByVal objCT As Object) As Boolean

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim iIndice As Integer

    For iIndice = 1 To objCT.gobjNFiscal.ColItensNF.Count
        If objCT.gobjNFiscal.ColItensNF(iIndice).iPossuiGrade = MARCADO Then
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
Dim lErro As Long

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
            
            If objAlmoxarifado.iFilialEmpresa = giFilialEmpresa Then
            
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

Private Sub Fornecedor_Preenche()
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134062

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134062

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157000)

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
Dim objFornecedor As New ClassFornecedor
Dim lFornecedor As Long
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
      
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 132932
    If Len(Trim(Filial.Text)) = 0 Then gError 132933
  
    objFornecedor.sNomeReduzido = Fornecedor.Text
    
    'Lê o fornecedor
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 132934
    
    If lErro = 6681 Then gError 132935 'Se nao encontrar --> erro

    lFornecedor = objFornecedor.lCodigo

    If Len(Trim(Filial.Text)) > 0 Then
        iFilial = Codigo_Extrai(Filial.Text)
    End If
    
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 132936
    
    colSelecao.Add lFornecedor
    colSelecao.Add iFilial
    colSelecao.Add sProdutoFormatado
    colSelecao.Add sProdutoFormatado
    colSelecao.Add sContrato
    colSelecao.Add sContrato
    
    objItemDeContrato.sContrato = sContrato
    objItemDeContrato.iSeq = iSeq
    
    Call Chama_Tela("ContratosFornItensLista", colSelecao, objItemDeContrato, objEventoItemContrato)

    Exit Sub
    
Erro_BotaoItemContrato_Click:
    
    Select Case gErr
    
        Case 132930
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 132932
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 132933
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 132934, 132936
        
        Case 132935
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, Fornecedor.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157001)

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

        lErro = ProdutoLinha_Preenche(objProduto)
        If lErro <> SUCESSO Then gError 136161
        
    End If
    
     'Se for um Item com medição
    If obItensDeContrato.iMedicao = MARCADO Then
    
        'Chama tela para obter as medições
        Call Chama_Tela_Modal("MedicoesNFiscal", gcolcolMedicoes.Item(GridItens.Row), obItensDeContrato)
        If giRetornoTela = vbCancel Then gError 136200
    
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoItemContrato_evSelecao:

    GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = ""

    Select Case gErr
    
        Case 136159, 136161, 136177, 136200
        
        Case 136160
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, obItensDeContrato.sProduto)

        Case 136178
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 136225
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_CONTRATO_REPETIDO", gErr, obItensDeContrato.sContrato, obItensDeContrato.iSeq, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157002)

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
    
        'Le contrato
        lErro = CF("Contrato_Le", objContrato)
        If lErro <> 129332 And lErro <> SUCESSO Then gError 132941
        
        'Se não está cadastrado = > Erro
        If lErro = 129332 Then gError 132942
        
        'Se mudou o fornecedor => Erro
        If objNFiscal.lFornecedor <> objContrato.lFornecedor Then gError 136180
    
        objItemContrato.lNumIntContrato = objContrato.lNumIntDoc
        objItemContrato.iSeq = iSeq
    
        'Le o item
        lErro = CF("ItensDeContrato_Le2", objItemContrato)
        If lErro <> 129266 And lErro <> SUCESSO Then gError 132943
        
        'Se não estiver cadastrado => Erro
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
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_CODIGO_NAO_CADASTRADO_GRID", gErr, sContrato, iSeq, iLinha)

        Case 136175
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_MEDICAO_NAO_CADASTRADO_GRID", gErr, objItemMedicao.lMedicao, objItemMedicao.iItem, iLinha)

        Case 136180
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_FORNECEDOR_NF", gErr, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157003)

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

    'Se contrato estiver preenchida
    If Len(Trim(Contrato.Text)) > 0 Then
    
        objContrato.sCodigo = Contrato.Text
        objContrato.iFilialEmpresa = giFilialEmpresa
    
        'Le o contrato
        lErro = CF("Contrato_Le", objContrato)
        If lErro <> 129332 And lErro <> SUCESSO Then gError 132948
        
        'Se não estiver cadastrado => Erro
        If lErro = 129332 Then gError 132949
        
        'Se não for um contrato a pagar + Erro
        If objContrato.iTipo <> CONTRATOS_PAGAR Then gError 132990

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
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_PAGAR", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157004)
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
    
        lErro = CF("Contrato_Le", objContrato)
        If lErro <> 129332 And lErro <> SUCESSO Then gError 132951
        
        If lErro = 129332 Then gError 132952

        objItemContrato.lNumIntContrato = objContrato.lNumIntDoc
        objItemContrato.iSeq = StrParaInt(Item.Text)
    
        lErro = CF("ItensDeContrato_Le2", objItemContrato)
        If lErro <> 129266 And lErro <> SUCESSO Then gError 132953
        
        If lErro = 129266 Then gError 132954
        
    End If
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157005)
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
    If lErro <> SUCESSO Then gError 188439

    Saida_Celula_DataCobranca = SUCESSO

    Exit Function

Erro_Saida_Celula_DataCobranca:

    Saida_Celula_DataCobranca = gErr

    Select Case gErr

        Case 188439
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188440)
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
    
        lErro = CF("Contrato_Le2", objContrato)
        If lErro <> 129261 And lErro <> SUCESSO Then gError 132972
           
        objItensDeContratoCob.lNumIntItensContrato = objItemNF.objCobrItensContrato.lNumIntItensContrato
        objItensDeContratoCob.lNumIntItemNFiscal = objItemNF.lNumIntDoc
    
        lErro = CF("ItensDeContratoCobranca_Le", objItensDeContratoCob)
        If lErro <> SUCESSO And lErro <> 188444 Then gError 188446
    
        If lErro = SUCESSO Then
        
            If objItensDeContratoCob.dtDataUltCobranca <> DATA_NULA Then GridItens.TextMatrix(iLinha, iGrid_DataCobranca_Col) = objItensDeContratoCob.dtDataUltCobranca
        
        Else
    
            objItensMedCtr.lNumIntItensContrato = objItemNF.objCobrItensContrato.lNumIntItensContrato
            objItensMedCtr.lNumIntItemNFiscal = objItemNF.lNumIntDoc
            
            lErro = CF("ItensDeMedicaoCobranca_Le", objItensMedCtr, colItensMedCtr)
            If lErro <> SUCESSO Then gError 193843
        
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
    
        Case 132970, 132971, 132984, 188446, 193843
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157006)

    End Select

    Exit Function

End Function

Public Sub BotaoDocContrato_Click()

Dim lErro As Long
Dim objMedicaoContrato As New ClassMedicaoContrato
Dim objContrato As New ClassContrato
Dim sContrato As String

On Error GoTo Erro_BotaoDocContrato_Click

    If GridItens.Row = 0 Then gError 136006

    sContrato = GridItens.TextMatrix(GridItens.Row, iGrid_Contrato_Col)

    If Len(Trim(sContrato)) = 0 Then gError 136007
    
    objContrato.sCodigo = sContrato
    objContrato.iFilialEmpresa = giFilialEmpresa

    Call Chama_Tela("ContratoPagar", objContrato)

    Exit Sub
    
Erro_BotaoDocContrato_Click:
    
    Select Case gErr

        Case 136006
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 136007
             Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONTRATO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157007)

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
        If lErro <> SUCESSO Then gError 188424
                    
        If Len(sAviso) > 0 Then gError 188437
        
    End If

    Valida_Contrato_Itens = SUCESSO

    Exit Function

Erro_Valida_Contrato_Itens:

    Valida_Contrato_Itens = gErr

    Select Case gErr
    
        Case 136017 To 136022, 188424

        Case 136201
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_MEDICAO_NAO_PREENCHIDO", gErr, iLinha)
        
        Case 136176
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_MEDICAO_NAO_CADASTRADO_GRID", gErr, objItemMedicao.lMedicao, objItemMedicao.iItem, iLinha)
        
        Case 136153
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_MEDICAO_SEM_MEDICAO", gErr, iLinha)
        
        Case 188437
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_DATA_COBRANCA", gErr, sAviso)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157008)

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
    
    If objContrato.iTipo <> CONTRATOS_PAGAR Then gError 136209
    
    objItemContrato.lNumIntContrato = objContrato.lNumIntDoc
    objItemContrato.iSeq = iSeq

    lErro = CF("ItensDeContrato_Le2", objItemContrato)
    If lErro <> 129266 And lErro <> SUCESSO Then gError 136210
    
    If lErro = 129266 Then gError 136211
    
    If objItemContrato.iMedicao = 0 Then gError 136214
    
    lErro = Chama_Tela_Modal("MedicoesNFiscal", gcolcolMedicoes.Item(GridItens.Row), objItemContrato)
    If lErro <> SUCESSO Then gError 136206

    Exit Sub
    
Erro_BotaoMedicao_Click:
    
    Select Case gErr

        Case 132937
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 132938
             Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONTRATO_PREENCHIDO", gErr)
        
        Case 132939
             Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_CODIGO_CONTRATO_PREENCHIDO", gErr)
             
        Case 136206, 136207, 136210
        
        Case 136208
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_CADASTRADO", gErr, objContrato.sCodigo)
        
        Case 136209
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_PAGAR", gErr, objContrato.sCodigo, objContrato.iFilialEmpresa)
        
        Case 136211
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_CONTRATO_NAO_CADASTRADO_GRID", gErr, sContrato, iSeq, GridItens.Row)
        
        Case 136214
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_SEM_MEDICAO", gErr, iSeq)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157009)

    End Select
    
    Exit Sub

End Sub

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long, sID As String
Dim objNFiscal As New ClassNFiscal
Dim objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_Anotacao_Extrai

    'Lê os dados da Tela NFiscal
    lErro = Move_NF_Memoria(objNFiscal)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157010)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157011)
     
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

Public Sub CodigodeBarras_Change()
Dim sTexto As String

    If giTeste = 1 Then Exit Sub

    iAlterado = REGISTRO_ALTERADO
    
    If left(CodigodeBarras.Text, 1) = "8" Then
        giTeste = 1
        sTexto = CodigodeBarras.ClipText
        CodigodeBarras.Mask = ""
        CodigodeBarras.Mask = "###########.#.###########.#.###########.#.###########.#"
        CodigodeBarras.PromptInclude = False
        CodigodeBarras.Text = sTexto
        CodigodeBarras.PromptInclude = True
        giTeste = 0
    Else
        giTeste = 1
        sTexto = CodigodeBarras.ClipText
        CodigodeBarras.Mask = ""
        CodigodeBarras.Mask = "#####.#####.#####.######.#####.######.#.#################"
        CodigodeBarras.PromptInclude = False
        CodigodeBarras.Text = sTexto
        CodigodeBarras.PromptInclude = True
        giTeste = 0
    End If
End Sub

Public Sub CodigodeBarras_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub CodigodeBarras_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub CodigodeBarras_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridParcelas.objControle = CodigodeBarras
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_CodigodeBarras(objGridInt As AdmGrid) As Long
'Faz a crítica da celula CodigodeBarras do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim dColunaSoma As Double
Dim bCancel As Boolean

On Error GoTo Erro_Saida_Celula_CodigodeBarras

    Set objGridInt.objControle = CodigodeBarras
    
    CodBarrasAvulso.Mask = CodigodeBarras.Mask
    
    CodBarrasAvulso.PromptInclude = False
    CodBarrasAvulso.Text = CodigodeBarras.Text
    CodBarrasAvulso.PromptInclude = True
    
    bCancel = False
    
    Call CodBarrasAvulso_Validate(bCancel)
    If bCancel Then Error 18426
    
    CodBarrasAvulso.PromptInclude = False
    CodBarrasAvulso.Text = ""
    CodBarrasAvulso.PromptInclude = True
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 18427

    Saida_Celula_CodigodeBarras = SUCESSO

    Exit Function

Erro_Saida_Celula_CodigodeBarras:

    Saida_Celula_CodigodeBarras = Err

    Select Case Err

        Case 18426, 18427
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 156130)

    End Select

    Exit Function

End Function

Public Sub CodBarrasAvulso_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodBanco As Integer
Dim dNumDias As Double
Dim dtDataVenc As Date
Dim iIndice As Integer
Dim dtDataVencReal As Date
Dim dValor As Double
Dim sLinDig As String

On Error GoTo Erro_CodBarrasAvulso_Validate

    If Len(Trim(CodBarrasAvulso.ClipText)) <> 0 Then

        'Verifica se tem alguma linha selecionada no Grid
        If GridParcelas.Row = 0 Then gError 188323
        
        lErro = CB_Converte_LinDig(CodBarrasAvulso.ClipText, sLinDig)
        If lErro <> SUCESSO Then gError 188324
        
        CodBarrasAvulso.PromptInclude = False
        CodBarrasAvulso.Text = sLinDig
        CodBarrasAvulso.PromptInclude = True
    
        'contas de serviços publicos (concessionarias e impostos)
        If StrParaInt(left(CodBarrasAvulso.Text, 1)) = 8 Then
            
            dValor = StrParaDbl(Mid(CodBarrasAvulso.ClipText, 5, 7) & Mid(CodBarrasAvulso.ClipText, 13, 4)) / 100
        
            If dValor <> 0 Then
        
                GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Valor_Col) = Format(dValor, "Standard")
            
            End If
        Else
        
            'Preenche a Combo de Banco Cobrador com os 3 primeiros dígitos do Código de barras
            iCodBanco = StrParaInt(left(CodBarrasAvulso.Text, 3))
            
            For iIndice = 0 To comboCobrador.ListCount - 1
                If comboCobrador.ItemData(iIndice) = iCodBanco Then
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Banco_Col) = comboCobrador.List(iIndice)
                    Exit For
                End If
            Next
            
            dNumDias = StrParaDbl(Mid(CodBarrasAvulso.ClipText, 34, 4))
        
            If dNumDias <> 0 Then
        
                dtDataVenc = DateAdd("d", dNumDias, CDate("7/10/1997"))
                GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col) = Format(dtDataVenc, "dd/mm/yyyy")
                
                'Calcula a Data de Vencimento Real
                lErro = CF("DataVencto_Real", dtDataVenc, dtDataVencReal)
                If lErro <> SUCESSO Then gError 188325
        
                'Coloca data de Vencimento Real no Grid
                GridParcelas.TextMatrix(GridParcelas.Row, iGrid_VenctoReal_Col) = Format(dtDataVencReal, "dd/mm/yyyy")
                
            End If

            dValor = StrParaDbl(Mid(CodBarrasAvulso.ClipText, 38, 10)) / 100
        
            If dValor <> 0 Then
        
                GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Valor_Col) = Format(dValor, "Standard")
            
            End If
                
        End If
        
        For iIndice = 0 To TipoCobranca.ListCount - 1
            If TipoCobranca.ItemData(iIndice) = TIPO_COBRANCA_BANCARIA Then
                GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                Exit For
            End If
        Next
        
        giTeste = 1
        CodigodeBarras.Mask = CodBarrasAvulso.Mask
        
        CodigodeBarras.PromptInclude = False
        CodigodeBarras.Text = CodBarrasAvulso.Text
        CodigodeBarras.PromptInclude = True
        
        giTeste = 0
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_CodigoDeBarras_Col) = CodBarrasAvulso.Text
    
    End If
    
    Exit Sub

Erro_CodBarrasAvulso_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 188323
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 188324, 188325
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 188326)
    
    End Select
    
    Exit Sub

End Sub

Public Sub CodBarrasAvulso_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodBarrasAvulso, iAlterado)

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
    If lErro <> SUCESSO And lErro <> 129261 Then gError 188425

    If lErro <> SUCESSO Then gError 188426

    dtDataProxCobranca = objItemContrato.dtDataProxCobranca
    dtDataUltCobranca = DATA_NULA

    Do While dtDataProxCobranca < dtDataCobrancaTela And dtDataProxCobranca < objContrato.dtDataFimContrato

        dtDataUltCobranca = dtDataProxCobranca

        'retorna a data da proxima cobranca do item do contrato
        lErro = CF("Contrato_ProxCobranca", objItemContrato.iPeriodicidade, dtDataProxCobranca, objItemContrato.dtDataProxCobranca, dtDataProxCobranca, dtRefFim)
        If lErro <> SUCESSO Then gError 188427

        If dtDataUltCobranca = dtDataProxCobranca Then Exit Do

    Loop

    If dtDataProxCobranca > dtDataCobrancaTela Then
        If dtDataProxCobranca > objContrato.dtDataFimContrato And dtDataUltCobranca = DATA_NULA Then
            sAviso = sAviso & "O contrato da linha %s do grid já está encerrado" & Chr(10)
        ElseIf dtDataProxCobranca > objContrato.dtDataFimContrato Then
            sAviso = sAviso & "A data mais próxima válida da informada na linha " & CStr(iItemGrid) & " do grid é " & CStr(dtDataUltCobranca) & Chr(10)
        ElseIf dtDataUltCobranca = DATA_NULA Then
            sAviso = sAviso & "A data mais próxima válida da informada na linha " & CStr(iItemGrid) & " do grid é " & CStr(dtDataProxCobranca) & Chr(10)
        Else
            sAviso = sAviso & "As datas mais próximas válidas da informada na linha " & CStr(iItemGrid) & " do grid são " & CStr(dtDataUltCobranca) & " e " & CStr(dtDataProxCobranca) & Chr(10)
        End If
    ElseIf dtDataProxCobranca > objContrato.dtDataFimContrato And objContrato.dtDataFimContrato <> DATA_NULA Then
    
        'Se tem a data de referência, mesmo cobrando depois do contrato encerrado, se for referente a algo ainda do contrato (mesmo que parcial) não pode dar erro
        If objItemContrato.dtDataRefIni = DATA_NULA Or objItemContrato.dtDataRefIni > objContrato.dtDataFimContrato Then
            If dtDataUltCobranca <> DATA_NULA Then
                sAviso = sAviso & "A data mais próxima válida da informada na linha " & CStr(iItemGrid) & " do grid é " & CStr(dtDataUltCobranca) & Chr(10)
            Else
                sAviso = sAviso & "O contrato da linha " & CStr(iItemGrid) & " do grid já está encerrado" & Chr(10)
            End If
        End If
    End If
                
    If Len(sAviso) > 0 Then sAviso = Chr(10) & sAviso
                
    Testa_DataCobranca_ItemContrato = SUCESSO

    Exit Function

Erro_Testa_DataCobranca_ItemContrato:

    Testa_DataCobranca_ItemContrato = gErr

    Select Case gErr
    
        Case 188425, 188427

        Case 184426
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTRATO_NAO_CADASTRADO1", gErr, objItemContrato.lNumIntContrato)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188428)

    End Select

    Exit Function

End Function

Private Function NFiscalContrato_Gera1(objGeracaoFatContrato As ClassGeracaoFatContrato, objNFiscal As ClassNFiscal) As Long
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

On Error GoTo Erro_NFiscalContrato_Gera1

    Set gcolcolMedicoes = New Collection

    objGeracaoFatContrato.iFilialEmpresa = giFilialEmpresa
    objGeracaoFatContrato.iTipoNFiscal = DOCINFO_NFEEFN

    'Lê o Tipo de Documento
    objTipoDocInfo.iCodigo = objGeracaoFatContrato.iTipoNFiscal
    
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 188527

    'se não estiver cadastrado ==> erro
    If lErro = 31415 Then gError 188528
    
    lErro = CF("NFiscalContrato_Le1", objGeracaoFatContrato, CONTRATOS_PAGAR)
    If lErro <> AD_SQL_SUCESSO Then gError 188529

    If objGeracaoFatContrato.colItens.Count > 0 Then

        Set objContratoFat = objGeracaoFatContrato.colItens.Item(1)
                      
        Set colcolMedicoes = New Collection
        Set colContFatItensAgrupado = New Collection
                      
        Set objNFiscal = New ClassNFiscal
                
        Call objNFiscal.Inicializa_Tributacao
        Call objNFiscal.objTributacaoNF.Coloca_Auto
        
        objNFiscal.iCondicaoPagto = objContratoFat.iCondPagto
        objNFiscal.iFilialForn = objContratoFat.iFilialFornecedor
        objNFiscal.lFornecedor = objContratoFat.lFornecedor
        objNFiscal.dtDataReferencia = objContratoFat.dtDataref
        objNFiscal.dtDataEmissao = gdtDataAtual
        objNFiscal.dtDataEntrada = gdtDataAtual
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
                                If lErro <> SUCESSO And lErro <> 136173 Then gError 193845
        
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
                    If lErro <> SUCESSO And lErro <> 136173 Then gError 193846
                    
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
            
            objItemNF.dCusto = Arredonda_Moeda(objContratoFatItens.dCusto, 4)
            objItemNF.dQuantidade = objContratoFatItens.dQuantidade
            objItemNF.dPrecoUnitario = Arredonda_Moeda(objContratoFatItens.dVlrCobrar / objContratoFatItens.dQuantidade, 4)
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
        If lErro <> SUCESSO Then gError 188530
        
        lErro = objTribTab.TributacaoNF_Reset(objNFiscal)
        If lErro <> SUCESSO Then gError 188530
        
        'obtem dados do contrato
        Call objTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
        Call objTribTab.DataEmissao_Alterada(objNFiscal.dtDataEmissao)
        Call objTribTab.Serie_Alterada(objNFiscal.sSerie)
        
        lErro = gobjTributacao.AtualizaImpostos(objTribTab.mvarobjNFTributacao, 0)
        If lErro <> SUCESSO Then gError 188531
        
        lErro = objTribTab.PreencherNF(objNFiscal)
        If lErro <> SUCESSO Then gError 188532
        
        Set objTribTab = Nothing
        
        objNFiscal.sNaturezaOp = objNFiscal.objTributacaoNF.sNaturezaOpInterna
        
        'acertar valor total e da parcela, considerando os tributos, retencóes,...
        objNFiscal.dValorTotal = dValorTotal + objNFiscal.objTributacaoNF.dICMSSubstValor + objNFiscal.objTributacaoNF.dIPIValor + objNFiscal.dValorFrete + objNFiscal.dValorSeguro + objNFiscal.dValorOutrasDespesas + IIf(objNFiscal.objTributacaoNF.iISSIncluso = 0, objNFiscal.objTributacaoNF.dISSValor, 0)
        'Adiciona a parcela na coleção de parcelas da Nota Fiscal
        dValorLiquido = objNFiscal.dValorTotal - (objNFiscal.objTributacaoNF.dCOFINSRetido + objNFiscal.objTributacaoNF.dCSLLRetido + objNFiscal.objTributacaoNF.dIRRFValor + objNFiscal.objTributacaoNF.dPISRetido + objNFiscal.objTributacaoNF.dISSRetido)
         
        lErro = NFiscalContrato1_Gera_Parcelas(objNFiscal)
        If lErro <> SUCESSO Then gError 188533
        
        lErro = NFiscalContrato1_Gera_Peso(objNFiscal)
        If lErro <> SUCESSO Then gError 188534
            
    End If

    NFiscalContrato_Gera1 = SUCESSO

    Exit Function

Erro_NFiscalContrato_Gera1:

    NFiscalContrato_Gera1 = gErr

    Select Case gErr

        Case 188527, 188529 To 188534, 193845, 193846
        
        Case 188528
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188535)

    End Select

    Exit Function

End Function

Private Function NFiscalContrato1_Gera_Parcelas(objNFiscal As ClassNFiscal)

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim dValorLiquido As Double
Dim iIndice As Integer
Dim dtDataVenctoReal As Date

On Error GoTo Erro_NFiscalContrato1_Gera_Parcelas

    dValorLiquido = objNFiscal.dValorTotal - (objNFiscal.objTributacaoNF.dCOFINSRetido + objNFiscal.objTributacaoNF.dCSLLRetido + objNFiscal.objTributacaoNF.dIRRFValor + objNFiscal.objTributacaoNF.dPISRetido + objNFiscal.objTributacaoNF.dISSRetido)

    objCondicaoPagto.iCodigo = objNFiscal.iCondicaoPagto
    
    'Lê a condição de pagamento
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 188527
    
    'Calcula os valores das Parcelas
    objCondicaoPagto.dValorTotal = dValorLiquido
    objCondicaoPagto.dtDataref = objNFiscal.dtDataReferencia
    
    'Calcula os valores das Parcelas
    lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, True)
    If lErro <> SUCESSO Then gError 188528

    'Coloca os valores das Parcelas no Grid Parcelas
    For iIndice = 1 To objCondicaoPagto.colParcelas.Count
    
            'Calcula Data Vencimento Real
            lErro = CF("DataVencto_Real", objCondicaoPagto.colParcelas(iIndice).dtVencimento, dtDataVenctoReal)
            If lErro <> SUCESSO Then gError 188526
    
            objNFiscal.ColParcelaPagar.Add 0, 0, iIndice, STATUS_ABERTO, objCondicaoPagto.colParcelas(iIndice).dtVencimento, dtDataVenctoReal, objCondicaoPagto.colParcelas(iIndice).dValor, objCondicaoPagto.colParcelas(iIndice).dValor, 1, 0, 0, 0, "", ""
            
    Next
    
    NFiscalContrato1_Gera_Parcelas = SUCESSO
    
    Exit Function

Erro_NFiscalContrato1_Gera_Parcelas:

    NFiscalContrato1_Gera_Parcelas = gErr

    Select Case gErr
    
        Case 188526 To 188528

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188529)

    End Select
       
    Exit Function

End Function

Private Function NFiscalContrato1_Gera_Peso(objNFiscal As ClassNFiscal)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim objItemNF As ClassItemNF
Dim dPesoL As Double
Dim dPesoB As Double
Dim dFator As Double

On Error GoTo Erro_NFiscalContrato1_Gera_Peso

    For Each objItemNF In objNFiscal.ColItensNF
    
        Set objProduto = New ClassProduto
    
        objProduto.sCodigo = objItemNF.sProduto
    
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 188536
    
        'Converte a unidade de medida de UM de Venda para a UM da linha do Grid
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemNF.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
        If lErro <> SUCESSO Then gError 188537
    
        dPesoL = dPesoL + (objProduto.dPesoLiq * objItemNF.dQuantidade * dFator)
        dPesoB = dPesoB + (objProduto.dPesoBruto * objItemNF.dQuantidade * dFator)
    
    Next
    
    objNFiscal.dPesoBruto = dPesoB
    objNFiscal.dPesoLiq = dPesoL
    
    NFiscalContrato1_Gera_Peso = SUCESSO
    
    Exit Function

Erro_NFiscalContrato1_Gera_Peso:

    NFiscalContrato1_Gera_Peso = gErr

    Select Case gErr
    
        Case 188536, 188537

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188538)

    End Select
       
    Exit Function

End Function

Public Sub ContratoPag_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContratoPag_GotFocus()

    Call MaskEdBox_TrataGotFocus(ContratoPag, iAlterado)

End Sub

Public Sub ContratoLabel_Click()

Dim objContrato As New ClassContrato
Dim colSelecao As New Collection

    If Len(Trim(ContratoPag.Text)) > 0 Then
        objContrato.sCodigo = ContratoPag.Text
        objContrato.iFilialEmpresa = giFilialEmpresa
    End If
    
    Call Chama_Tela("ContratosPagLista", colSelecao, objContrato, objEventoContrato)

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

    lErro = NFiscalContrato_Gera1(objGeracaoFatContrato, objNFiscal)
    If lErro <> SUCESSO Then gError 188512

    lErro = Traz_NFiscal_Tela1(objNFiscal)
    If lErro <> SUCESSO Then gError 188514

    ContratoPag.Text = objContrato.sCodigo

    Exit Sub
    
Erro_objEventoContrato_evSelecao:
    
    Select Case gErr
    
        Case 188512, 188514

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188513)

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

    If Len(Trim(ContratoPag.Text)) = 0 Then gError 188521

    objGeracaoFatContrato.sContratoIni = ContratoPag.Text

    lErro = NFiscalContrato_Gera1(objGeracaoFatContrato, objNFiscal)
    If lErro <> SUCESSO Then gError 188522

    lErro = Traz_NFiscal_Tela1(objNFiscal)
    If lErro <> SUCESSO Then gError 188523

    ContratoPag.Text = objGeracaoFatContrato.sContratoIni

    Exit Sub
    
Erro_BotaoContrato_Click:
    
    Select Case gErr
    
        Case 188521
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONTRATO_PREENCHIDO", gErr)

        Case 188522, 188523

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188524)

    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoVerContrato_Click()

Dim lErro As Long
Dim objContrato As New ClassContrato

On Error GoTo Erro_BotaoDocContrato_Click

    If Len(Trim(ContratoPag.Text)) = 0 Then gError 188525
    
    objContrato.sCodigo = ContratoPag.Text
    objContrato.iFilialEmpresa = giFilialEmpresa

    Call Chama_Tela("ContratoPagar", objContrato)

    Exit Sub
    
Erro_BotaoDocContrato_Click:
    
    Select Case gErr

        Case 188525
             Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CONTRATO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188526)

    End Select
    
    Exit Sub

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
Dim dPercDesc As Double

On Error GoTo Erro_Traz_NFiscal_Tela1
    
    gbCarregandoTela = True
    
    'Limpa a tela
    Call Limpa_Tela_NFiscal3
    
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
'    Else
'        'preenche o Cliente
'        Cliente.Text = objNFiscal.lCliente
'        Call Cliente_Validate(bCancel)
'
'        'Preenche a Filial do Cliente
'        Filial.Text = objNFiscal.iFilialCli
'        Call Filial_Validate(bCancel)
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
    Observacao.Text = objNFiscal.sObservacao
    
    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = CStr(objNFiscal.dPesoLiq)
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = CStr(objNFiscal.dPesoBruto)
    
    'preenche o Grid com os Ítens da Nota Fiscal
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 188515

    'Para cada ítem da Coleção
    For Each objItemNF In objNFiscal.ColItensNF

        iLinha = iLinha + 1

        GridItens.TextMatrix(iLinha, iGrid_Contrato_Col) = objItemNF.objItensDeContrato.sContrato
        If objItemNF.objItensDeContrato.iSeq <> 0 Then GridItens.TextMatrix(iLinha, iGrid_ItemContrato_Col) = objItemNF.objItensDeContrato.iSeq
        If objItemNF.objItensDeContrato.dtDataProxCobranca <> DATA_NULA Then GridItens.TextMatrix(iLinha, iGrid_DataCobranca_Col) = objItemNF.objItensDeContrato.dtDataProxCobranca

    Next

    'Preenche a Condicao de Pagto
    If objNFiscal.iCondicaoPagto > 0 Then
        CondicaoPagamento.Text = objNFiscal.iCondicaoPagto
        Call CondicaoPagamento_Validate(bCancel)
    End If
        
    lErro = Preenche_GridParcelas(objNFiscal)
    If lErro <> SUCESSO Then gError 188517

    'Tenta fazer a distribuição automatica
    lErro = gobjDistribuicao.Preenche_GridDistribuicaoPC1(colItemPedCompraInfo)
    If lErro <> SUCESSO And lErro <> 89627 Then gError 188518
    
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

    'Set gobjNFiscal = objNFiscal

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 188519
'*** fim tributacao

    For iIndice = 1 To objGridItens.iLinhasExistentes
    
        GridItens.Row = iIndice
        
        lErro = gobjDistribuicao.Distribuicao_Processa()
        'If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Next
    
'    'distribuicao
'    'Preenche o Grid com as Distribuições dos itens da Nota Fiscal
'    lErro = gobjDistribuicao.Preenche_GridDistribuicao(objNFiscal)
'    If lErro <> SUCESSO Then gError 188516

    lErro = Preenche_GridParcelas(objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    gbCarregandoTela = False
    
    Traz_NFiscal_Tela1 = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela1:

    gbCarregandoTela = False
    
    Traz_NFiscal_Tela1 = gErr

    Select Case gErr

        Case 188515 To 188519, ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 188520)

    End Select

    Exit Function

End Function

Public Sub CodBarrasAvulso_Change()

Dim sTexto As String
Static iTeste As Integer
Dim iPos As Integer

    If iTeste = 1 Then Exit Sub
     If left(CodBarrasAvulso.Text, 1) = "8" Then

        iPos = CodBarrasAvulso.SelStart
        iTeste = 1
        sTexto = CodBarrasAvulso.ClipText
        CodBarrasAvulso.Mask = "###########.#.###########.#.###########.#.###########.#"
        CodBarrasAvulso.PromptInclude = False
        CodBarrasAvulso.Text = sTexto
        CodBarrasAvulso.PromptInclude = True
        CodBarrasAvulso.SelStart = iPos
        iTeste = 0

    Else
        iPos = CodBarrasAvulso.SelStart
        iTeste = 1
        sTexto = CodBarrasAvulso.ClipText
        CodBarrasAvulso.Mask = ""
        CodBarrasAvulso.Mask = "#####.#####.#####.######.#####.######.#.#################"
        CodBarrasAvulso.PromptInclude = False
        CodBarrasAvulso.Text = sTexto
        CodBarrasAvulso.PromptInclude = True
        CodBarrasAvulso.SelStart = iPos
        iTeste = 0
    End If

End Sub

Private Sub Limpa_Tela_NFiscal3()
'Limpa a Tela NFiscalEntrada

Dim lErro As Long
Dim iIndice As Integer
    
    'Chama o Limpa tela
    Call Limpa_Tela(Me)
    
    gsChvAnt = ""
    
    Set gobjAnotacao = Nothing
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    Set gobjInfoAdic = Nothing

    Set gobjNFiscal = New ClassNFiscal
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next
    
    'Limpa o Grid
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridParcelas)
    
'    If giTipoVersao = VERSAO_FULL Then
        Call Grid_Limpa(objGridBenef)
'    End If
    
    TipoFrete.ListIndex = 0
            
    'Limpa os campos que não são limpos pelo Limpa_Tela
    Filial.Clear
    
    
'    If giTipoVersao = VERSAO_FULL Then
        Serie.Text = ""
        Call Serie_Validate(bSGECancelDummy)
        FilialFornBenef.Clear
        NumRecebimento.Text = ""
'    End If
    
    Status.Caption = ""
    SubTotal.Caption = ""
    Transportadora.Text = ""
    PlacaUF.Text = ""
    'EletronicaFed.Value = vbUnchecked
    If gobjEST.iSerieEletronicaPadrao = MARCADO Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    
    VolumeMarca.Text = ""
    VolumeEspecie.Text = ""
    
    'Recarrega a combo VolumeEspecie e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)

    'Recarrega a combo VolumeMarca e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
    
    NFiscal.Text = ""
    SerieNFiscalOriginal.Text = ""
    FilialFornNFOrig.Clear

    Call gobjTribTab.DataEmissao_Alterada(DATA_NULA)

    'Preenche a Data de Entrada com a Data Atual
    DataEntrada.PromptInclude = False
    DataEntrada.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEntrada.PromptInclude = True
    Call DataEntrada_Validate(bSGECancelDummy)

    PagamentoAutomatico.Value = vbChecked

''*** incluidos p/tratamento de tributacao *******************************

    Call gobjTribTab.Limpa_Tela

'*** fim tributacao

    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento

    'distribuicao
    Call gobjDistribuicao.Limpa_Tela_Distribuicao

    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    Call CF2(Me, "NFCF_Limpa_Tela")

    Call CF2(Me, "NFEntC_Limpa_Tela")
    
    MsgAutomatica.Value = vbChecked
    
    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0

    Call ComandoSeta_Fechar(Me.Name)

End Sub

Public Sub NFFE_UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
'
End Sub

Function NFFE_Preenche_GridParcelas(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long
    NFFE_Preenche_GridParcelas = SUCESSO
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

Private Function Preco_Traz_Padrao(ByVal iLinha As Integer, ByVal sProduto As String) As Long

Dim lErro As Long
Dim lFornecedor As Long
Dim iFilial As Integer
Dim objFornecedor As ClassFornecedor
Dim dPreco As Double, sUM As String

On Error GoTo Erro_Preco_Traz_Padrao

    If gobjEST.iTrazPrecoUltCompr = MARCADO Then
    
        'Se o Fornecedor estiver preenchido lê o código no BD
        If Len(Trim(Fornecedor.Text)) > 0 Then
            
            Set objFornecedor = New ClassFornecedor
            
            objFornecedor.sNomeReduzido = Fornecedor.Text
            
            'Lê o fornecedor
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError ERRO_SEM_MENSAGEM
    
            lFornecedor = objFornecedor.lCodigo
    
        End If
        
        If Len(Trim(Filial.Text)) > 0 Then iFilial = Codigo_Extrai(Filial.Text)
        
        If iLinha <= objGridItens.iLinhasExistentes And iLinha > 0 And lFornecedor > 0 And iFilial > 0 And Len(Trim(sProduto)) > 0 Then
        
            lErro = CF("FilialForn_Le_UltCompra_Prod", lFornecedor, iFilial, sProduto, sUM, dPreco)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If dPreco > DELTA_VALORMONETARIO2 Then
                           
                GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col) = sUM
                GridItens.TextMatrix(iLinha, iGrid_ValorUnitario_Col) = Format(dPreco, FORMATO_PRECO_UNITARIO_EXTERNO)

                lErro = Calcula_Valores(iLinha)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                lErro = gobjTribTab.Alteracao_Item_Grid(iLinha)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Call ValorTotal_Calcula

            End If
        
        End If
        
    End If
    
    Preco_Traz_Padrao = SUCESSO
    
    Exit Function
    
Erro_Preco_Traz_Padrao:
    
    Preco_Traz_Padrao = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208275)
    
    End Select
    
    Exit Function
    
End Function

Private Function Carrega_TipoFrete() As Long

Dim lErro As Long
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer
Dim colTipoFrete As Collection

On Error GoTo Erro_Carrega_TipoFrete
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207465

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

        Case 207465

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207466)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211304)
            
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoTrazerNFe_Click()

Dim lErro  As Long
Dim objTela As Object, objNF As New ClassNFiscal

On Error GoTo Erro_BotaoTrazerNFe_Click

    If Len(Trim(ChvNFe.ClipText)) = 0 Then gError 211305

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
    
        Case 211305
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CHAVE_NAO_PREENCHIDA", gErr)
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211306)
            
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211307)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211308)

    End Select

    Exit Sub

End Sub

Public Sub ChvNFe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ChvNFe_GotFocus()
    Call MaskEdBox_TrataGotFocus(ChvNFe, iAlterado)
End Sub

Public Sub BotaoAnexos_Click()
    If gobjInfoAdic Is Nothing Then Set gobjInfoAdic = New ClassInfoAdic
    Call Chama_Tela_Modal("Anexos", gobjInfoAdic.objAnexos)
End Sub
