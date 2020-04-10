VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalEntrada"
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
Dim FrameL(1 To 5) As Frame

'************ GRADE **********************
Public gobjNFiscal As ClassNFiscal
'*****************************************

'Property Variables:
Dim m_Caption As String
Event Unload()

'inicio contabilidade

Public objGrid1 As AdmGrid
Public objContabil As New ClassContabil
Dim giComplemento As Integer

Public WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Public WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

'Flag que indica se a tela está sendo preenchida.
Public gbCarregandoTela As Boolean
Public giPosCargaOk As Integer

Public gobjTelaProjetoInfo As ClassTelaPRJInfo

'mnemonicos
Private Const TIPO_NOTA_FISCAL As String = "Tipo_Nota_Fiscal"
Private Const NATUREZA_OP As String = "Natureza_OP"
Private Const FORNECEDOR1 As String = "Fornecedor"
Private Const FILIAL1 As String = "Filial_Fornecedor"
Private Const SERIE1 As String = "Serie"
Private Const NFISCAL1 As String = "Nota_Fiscal"
Private Const DATA_EMISSAO As String = "Data_Emissao"
Private Const DATA_ENTRADA As String = "Data_Entrada"
Private Const PRODUTO1 As String = "Produto_Codigo"
Private Const UNIDADE_MED As String = "Unidade_Medida"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const VALOR_UNITARIO As String = "Valor_Unitario_Prod"
Private Const VALOR_TOTAL As String = "Valor_Total_Produto"
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
Private Const SUBTOTAL1 As String = "Valor_Produtos"
Private Const VALOR_FRETE As String = "Valor_Frete"
Private Const VALOR_SEGURO As String = "Valor_Seguro"
Private Const VALOR_DESPESAS As String = "Valor_Despesas"
Private Const IPI As String = "IPI_Valor"
Private Const VALOR_DESCONTO As String = "Valor_Desconto_Glb"
Private Const TOTAL1 As String = "Total_Nota_Fiscal"
'###########Tributação##############
Private Const ISS_VALOR As String = "ISS_Valor"
Private Const ISS_INCLUSO As String = "ISS_Incluso"
Private Const VALOR_IRRF As String = "Valor_IRRF"
Private Const VALOR_INSS As String = "Valor_INSS"
Private Const INSS_RETIDO As String = "INSS_Retido"
Private Const CTACONTABILEST1 As String = "ContaContabilEst"
Private Const CTACONTABILESTBENEF As String = "CtaContabilEstBenef"
Private Const CTACONTABILPRODBENEF As String = "CtaContabilProdBenef"
Private Const CTAFILIALFORNECEDOR1 As String = "FilialForn_Conta_Ctb"
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
Private Const IPI_PRODUTO As String = "IPI_Produto"
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
Private Const CCL_LCTO_NFORIG As String = "Ccl_Lcto_NF_Orig"
'fim da contabilidade

'Rastreamento
'Public gcolItensNF As Collection
'Public objGenerico As AdmGenerico
Public gobjRastreamento As ClassRastreamento

'distribuicao
Public gobjDistribuicao As Object


Public iAlterado As Integer
Dim iLinhaAntiga As Integer
Public iFornecedorAlterado As Integer
Public iFornecedorBenefAlterado As Integer
Public iFornNFOrigAlterado As Integer
Public iClienteNFOrigAlterado As Integer
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
Public iGrid_Ccl_Col As Integer
Public iGrid_ValorUnitario_Col As Integer
Public iGrid_ValorUnitarioMoeda_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PrecoTotal_Col As Integer
Public iGrid_PrecoTotalB_Col As Integer

Dim giRecebimento As Integer

'Grid relacionado com Beneficiamento
Public objGridBenef As AdmGrid


Public iGrid_ProdutoBenef_Col As Integer
Public iGrid_UnidadeMedBenef_Col As Integer
Public iGrid_QuantidadeBenef_Col As Integer
Public iGrid_DescricaoBenef_Col As Integer
Public iGrid_AlmoxarifadoBenef_Col As Integer
Public iGrid_ContaContabilEstBenef_Col As Integer
Public iGrid_ContaContabilProducaoBenef_Col As Integer

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
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoFornecedorBenef As AdmEvento
Attribute objEventoFornecedorBenef.VB_VarHelpID = -1
Private WithEvents objEventoFornNFOrig As AdmEvento
Attribute objEventoFornNFOrig.VB_VarHelpID = -1
Private WithEvents objEventoClienteNFOrig As AdmEvento
Attribute objEventoClienteNFOrig.VB_VarHelpID = -1
Private WithEvents objEventoChvNFe As AdmEvento
Attribute objEventoChvNFe.VB_VarHelpID = -1

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

Dim gsChvAnt As String

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

'Fernando subir Constante
Const CONFIRMA_VUNITARIO_GRAVACAO = 0 'Parametro iIndice
'O Valor Unitário do item %s não esta preenchido. Deseja Prosseguir? - iIndice

''*** incluidos p/tratamento de tributacao *******************************
Public gobjTribTab As ClassTribTab

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

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Fornecedor() As Object
     Set Fornecedor = objUserControl.Controls("Fornecedor")
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

Public Property Get UpDownVencimento() As Object
     Set UpDownVencimento = objUserControl.Controls("UpDownVencimento")
End Property

Public Property Get dtVencimento() As Object
     Set dtVencimento = objUserControl.Controls("DtVencimento")
End Property

Public Property Get BotaoRecebimentos() As Object
     Set BotaoRecebimentos = objUserControl.Controls("BotaoRecebimentos")
End Property

Public Property Get BotaoPlanoConta() As Object
     Set BotaoPlanoConta = objUserControl.Controls("BotaoPlanoConta")
End Property

Public Property Get BotaoEstoqueBenef() As Object
     Set BotaoEstoqueBenef = objUserControl.Controls("BotaoEstoqueBenef")
End Property

Public Property Get BotaoProdutosBenef() As Object
     Set BotaoProdutosBenef = objUserControl.Controls("BotaoProdutosBenef")
End Property

Public Property Get DescricaoItemBenef() As Object
     Set DescricaoItemBenef = objUserControl.Controls("DescricaoItemBenef")
End Property

Public Property Get UnidadeMedBenef() As Object
     Set UnidadeMedBenef = objUserControl.Controls("UnidadeMedBenef")
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

Public Property Get PlacaUF() As Object
     Set PlacaUF = objUserControl.Controls("PlacaUF")
End Property

Public Property Get Placa() As Object
     Set Placa = objUserControl.Controls("Placa")
End Property

Public Property Get Transportadora() As Object
     Set Transportadora = objUserControl.Controls("Transportadora")
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

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get PesoLiquido() As Object
     Set PesoLiquido = objUserControl.Controls("PesoLiquido")
End Property

Public Property Get PesoBruto() As Object
     Set PesoBruto = objUserControl.Controls("PesoBruto")
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

Public Property Get SerieOriginalLabel() As Object
     Set SerieOriginalLabel = objUserControl.Controls("SerieOriginalLabel")
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

Public Property Get FilialFornNFOrig() As Object
     Set FilialFornNFOrig = objUserControl.Controls("FilialFornNFOrig")
End Property

Public Property Get FilialFornNFOrigLabel() As Object
     Set FilialFornNFOrigLabel = objUserControl.Controls("FilialFornNFOrigLabel")
End Property

Public Property Get FornNFOrig() As Object
     Set FornNFOrig = objUserControl.Controls("FornNFOrig")
End Property

Public Property Get FornNFOrigLabel() As Object
     Set FornNFOrigLabel = objUserControl.Controls("FornNFOrigLabel")
End Property

Public Property Get FilialCliNFOrig() As Object
     Set FilialCliNFOrig = objUserControl.Controls("FilialCliNFOrig")
End Property

Public Property Get FilialCliNFOrigLabel() As Object
     Set FilialCliNFOrigLabel = objUserControl.Controls("FilialCliNFOrigLabel")
End Property

Public Property Get ClienteNFOrig() As Object
     Set ClienteNFOrig = objUserControl.Controls("ClienteNFOrig")
End Property

Public Property Get ClienteNFOrigLabel() As Object
     Set ClienteNFOrigLabel = objUserControl.Controls("ClienteNFOrigLabel")
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

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
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

Public Property Get BotaoConsultaNFPag() As Object
     Set BotaoConsultaNFPag = objUserControl.Controls("BotaoConsultaNFPag")
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

'Distribuicao
Public Property Get GridDist() As Object
     Set GridDist = objUserControl.Controls("GridDist")
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

Public Sub BotaoCcls_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCcl As ClassCcl

On Error GoTo Erro_BotaoCcls_Click

    If GridItens.Row = 0 Then gError 52796

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then gError 52797
    
    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

    Exit Sub
    
Erro_BotaoCcls_Click:
    
    Select Case gErr
    
        Case 52796
             Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 52797
             Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156533)

    End Select
    
    Exit Sub

End Sub

Public Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Public Sub DtVencimento_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(dtVencimento, iAlterado)

End Sub

Public Sub Destinatario_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoFrete_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
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
        If lErro <> SUCESSO Then gError 52798

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

        Case 52798 'Tratado na rotina chamadora
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156535)

    End Select

    Exit Sub

End Sub

Public Sub BotaoPlanoConta_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPlanoConta_Click

    If GridMovimentos.Row = 0 Then gError 52289
    
    If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col) = "" Then gError 52290

    sConta = String(STRING_CONTA, 0)

    If GridMovimentos.Col = iGrid_ContaContabilEstBenef_Col Then
        
        lErro = CF("Conta_Formata", ContaContabilEst.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 52291
    
    ElseIf GridMovimentos.Col = iGrid_ContaContabilProducaoBenef_Col Then


        lErro = CF("Conta_Formata", ContaContabilProducao.Text, sConta, iContaPreenchida)
        If lErro <> SUCESSO Then gError 52292

    End If
    
    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaESTLista
    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)
    
    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 52289
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 52290
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 52291, 52292

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156536)

    End Select

    Exit Sub

End Sub

Public Sub objEventoContaContabil_evSelecao(obj1 As Object)

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
        If lErro <> SUCESSO Then gError 52293
        
        objGridBenef.objControle.PromptInclude = False
        objGridBenef.objControle.Text = sContaEnxuta
        objGridBenef.objControle.PromptInclude = True

        GridMovimentos.TextMatrix(GridMovimentos.Row, GridMovimentos.Col) = objGridBenef.objControle.Text
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 52293
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156537)

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
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 52294
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 52295
            
            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaMascarada
            ContaContabilEst.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilEst.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 52296
    
            'conta não cadastrada
            If lErro = 5700 Then gError 52297
             
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 52298
    
    Saida_Celula_ContaContabilEst = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilEst:

    Saida_Celula_ContaContabilEst = gErr

    Select Case gErr

        Case 52294, 52296, 52298
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52295
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 52297
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilEst.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156538)

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
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 52299
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 52300
            
            ContaContabilProducao.PromptInclude = False
            ContaContabilProducao.Text = sContaMascarada
            ContaContabilProducao.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabilProducao.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 52301
    
            'conta não cadastrada
            If lErro = 5700 Then gError 52302
             
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 52303
    
    Saida_Celula_ContaContabilProducao = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabilProducao:

    Saida_Celula_ContaContabilProducao = gErr

    Select Case gErr

        Case 52299, 52301, 52303
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52300
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 52302
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabilProducao.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
                            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156539)

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
    If GridMovimentos.Row = 0 Then gError 52304
    
    ' Formata o Produto
    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 52305
            
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then objProduto.sCodigo = sProdutoFormatado
    
    'Lista de produtos que  podem participar da producao e inventariados
    Call Chama_Tela("ProdutoEstoquePCPLista", colSelecao, objProduto, objEventoProdutoBenef)
   
   Exit Sub
   
Erro_BotaoProdutosBenef_Click:

    Select Case gErr
        
        Case 52304
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
                
        Case 52305
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156540)
        
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
            If lErro <> SUCESSO Then gError 52306
    
            sProdutoMascarado = String(STRING_PRODUTO, 0)
    
            lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 52307
    
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col) = sProdutoMascarado
    
            ProdutoBenef.PromptInclude = False
            ProdutoBenef.Text = sProdutoMascarado
            ProdutoBenef.PromptInclude = True
    
            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 52308
    
            If lErro = 28030 Then gError 52309
            
            If Trim(objProduto.sContaContabilProducao) = "" Then
                
                objTipoDeProduto.iTipo = objProduto.iTipo
                
                lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
                If lErro <> SUCESSO And lErro <> 22531 Then gError 52310
                
                If lErro = 22531 Then gError 52311
                
                objProduto.sContaContabilProducao = objTipoDeProduto.sContaContabil
                            
            End If
    
            'Preenche a Linha do Grid
            lErro = ProdutoLinha_PreencheBenef(objProduto)
            If lErro <> SUCESSO Then gError 52312
            
            lErro = Preenche_ContaContabilEst(ProdutoBenef.Text)
            If lErro <> SUCESSO Then gError 52313
            
            lErro = QuantDisponivel_Calcula1(sProdutoMascarado, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col), objProduto)
            If lErro <> SUCESSO Then gError 52314

        End If
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 52306, 52308, 52310, 52312, 52313, 52314

        Case 52307
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)

        Case 52309
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 52311
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156541)

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
        If lErro <> SUCESSO Then gError 52315
        
        'preenche  a ContaContabilProducao
        ContaContabilProducao.PromptInclude = False
        ContaContabilProducao.Text = sContaEnxuta
        ContaContabilProducao.PromptInclude = True
        
        'preenche Conta De Producao
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ContaContabilProducaoBenef_Col) = ContaContabilProducao.Text
            
    End If
    
    'Unidade de Medida
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_DescricaoBenef_Col) = objProduto.sDescricao

    'Almoxarifado
        
    'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
    lErro = CF("AlmoxarifadoPadrao_Le_NomeReduzido", objProduto.sCodigo, sAlmoxarifadoPadrao)
    If lErro <> SUCESSO Then gError 52230

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
        
        Case 52230
        
        Case 52315
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objProduto.sContaContabilProducao)
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156542)

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

    If (GridMovimentos.Row = 0) Then gError 52315

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 52316

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 52317

    colSelecao.Add sProdutoFormatado

    Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoqueBenef)

    Exit Sub

Erro_BotaoEstoqueBenef_Click:

    Select Case gErr

        Case 52315
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 52316
                    
        Case 52317
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156543)

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
        If lErro <> SUCESSO Then gError 52318

        'Verifica se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Preenche o Nome do Almoxarifado
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            AlmoxarifadoBenef.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            'Calcula a Quantidade Disponível nesse Almoxarifado
            lErro = QuantDisponivel_Calcula(sCodProduto, objEstoqueProduto.sAlmoxarifadoNomeReduzido)
            If lErro <> SUCESSO Then gError 52319

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoqueBenef_evselecao:

    Select Case gErr

        Case 52318, 52319

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156544)

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

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156545)
    
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
        If lErro <> SUCESSO And lErro <> 25077 Then gError 52320

        If lErro = 25077 Then gError 52321
        
        If Trim(objProduto.sContaContabilProducao) = "" Then
                
            objTipoDeProduto.iTipo = objProduto.iTipo
            
            lErro = CF("TipoDeProduto_Le", objTipoDeProduto)
            If lErro <> SUCESSO And lErro <> 22531 Then gError 52322
            
            If lErro = 22531 Then gError 52323
            
            objProduto.sContaContabilProducao = objTipoDeProduto.sContaContabil
                        
        End If

    End If

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        If objProduto.iPCP = PRODUTO_PCP_NAOPODE Then gError 52324

        lErro = ProdutoLinha_PreencheBenef(objProduto)
        If lErro <> SUCESSO Then gError 52325
        
        lErro = Preenche_ContaContabilEst(ProdutoBenef.Text)
        If lErro <> SUCESSO Then gError 52326
            
        lErro = QuantDisponivel_Calcula1(ProdutoBenef.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col), objProduto)
        If lErro <> SUCESSO Then gError 52402
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 52327

    Saida_Celula_ProdutoBenef = SUCESSO

    Exit Function

Erro_Saida_Celula_ProdutoBenef:

    Saida_Celula_ProdutoBenef = gErr

    Select Case gErr

        Case 52320, 52322, 52325, 52326, 52327, 52402
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52321
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = ProdutoBenef.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If
        
        Case 52323
             Call Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_NAO_CADASTRADO", gErr, objTipoDeProduto.iTipo)
             Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52324
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PCP", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156546)

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
        If lErro <> SUCESSO Then gError 52328

        If Len(Trim(QuantDisponivelBenef.Caption)) <> 0 Then

            lErro = Testa_QuantRequisitada(CDbl(QuantidadeBenef.Text))
            If lErro <> SUCESSO Then gError 52329
            
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 52331

    Saida_Celula_QuantidadeBenef = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadeBenef:

    Saida_Celula_QuantidadeBenef = gErr

    Select Case gErr

        Case 52328, 52329, 52331
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156547)

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
        If lErro <> SUCESSO Then gError 52332

        lErro = TP_Almoxarifado_Produto_Grid(sProdutoFormatado, AlmoxarifadoBenef, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25130 And lErro <> 25135 Then gError 52333

        If lErro = 25130 Then gError 52334

        If lErro = 25135 Then gError 52335

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_AlmoxarifadoBenef_Col) = objAlmoxarifado.sNomeReduzido
        
        lErro = QuantDisponivel_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col), AlmoxarifadoBenef.Text)
        If lErro <> SUCESSO Then gError 52336
        
    Else

        'Limpa a Quantidade Disponível da Tela
        QuantDisponivelBenef.Caption = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 52337
    
    'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
    lErro = Preenche_ContaContabilEst(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_ProdutoBenef_Col))
    If lErro <> SUCESSO Then gError 52338

    Saida_Celula_AlmoxarifadoBenef = SUCESSO

    Exit Function

Erro_Saida_Celula_AlmoxarifadoBenef:

    Saida_Celula_AlmoxarifadoBenef = gErr

    Select Case gErr

        Case 52332, 52333, 52336, 52337, 52338
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52334

            vbMsg = Rotina_Aviso(vbYesNo, "ERRO_ALMOXARIFADO_INEXISTENTE", AlmoxarifadoBenef.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = AlmoxarifadoBenef.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 52335

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(AlmoxarifadoBenef.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(AlmoxarifadoBenef.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156548)

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
        If lErro <> SUCESSO Then gError 55626
    
        'Se a quantidade está preenchida e não se trata de estorno
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_QuantidadeBenef_Col))) <> 0 Then

            dQuantidade = CDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_QuantidadeBenef_Col))

            'Testa a Quantidade requisitada
            lErro = Testa_QuantRequisitada(dQuantidade)
            If lErro <> SUCESSO Then gError 55627

        End If
    
    Else
    
        QuantDisponivelBenef.Caption = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 52339

    Saida_Celula_UnidadeMedBenef = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMedBenef:

    Saida_Celula_UnidadeMedBenef = gErr

    Select Case gErr

        Case 52339, 55626, 55627
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156549)

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
        If lErro <> SUCESSO Then gError 52341

    End If

    Exit Sub

Erro_GridMovimentos_RowColChange:

    Select Case gErr

        Case 52341
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156550)

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
        If lErro <> SUCESSO Then gError 55614
        
    Else
    
        lErro = QuantDisponivel_Calcula1(sProduto, sAlmoxarifado, objProduto)
        If lErro <> SUCESSO Then gError 55615

    End If

    lErro = Testa_Quantidade()
    If lErro <> SUCESSO Then gError 55616

    QuantDisponivel_Calcula = SUCESSO

    Exit Function

Erro_QuantDisponivel_Calcula:

    QuantDisponivel_Calcula = gErr

    Select Case gErr

        Case 55614, 55615, 55616

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156551)

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
        If lErro <> SUCESSO Then gError 55617
    
        If GridMovimentos.Row >= GridMovimentos.FixedRows And Len(Trim(sAlmoxarifado)) <> 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            If (objProduto Is Nothing) Then
            
                Set objProduto = New ClassProduto
    
                objProduto.sCodigo = sProdutoFormatado
    
                'Lê o produto no BD para obter UM de estoque
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 55618
    
                If lErro = 28030 Then gError 55619
    
            End If
    
            objAlmoxarifado.sNomeReduzido = sAlmoxarifado
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 55620
    
            If lErro = 25056 Then gError 55621
    
            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
            objEstoqueProduto.sProduto = sProdutoFormatado
    
            'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
            lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 21306 Then gError 55622
    
            'Se não encontrou EstoqueProduto no Banco de Dados
            If lErro = 21306 Then
            
                 QuantDisponivelBenef.Caption = Formata_Estoque(0)
    
            Else
                sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMedBenef_Col)
        
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError 55623
        
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

    QuantDisponivel_Calcula1 = gErr

    Select Case gErr

        Case 55617, 55618, 55620, 55622, 55623

        Case 55619
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 55621
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156552)

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
            If lErro <> SUCESSO Then gError 55624

        End If
    
    End If

    Testa_Quantidade = SUCESSO

    Exit Function

Erro_Testa_Quantidade:

    Testa_Quantidade = gErr

    Select Case gErr

        Case 55624

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156553)

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
        If lErro <> SUCESSO Then gError 55625

        objProduto.sCodigo = sProdutoFormatado

        'Lê o produto para saber qual é a sua ClasseUM
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 55626
    
        If lErro = 28030 Then gError 55627
    
        For iIndice = 1 To objGridBenef.iLinhasExistentes
    
            'Não pode somar a Linha atual
            If GridMovimentos.Row <> iIndice Then
    
                sCodProduto = GridMovimentos.TextMatrix(iIndice, iGrid_ProdutoBenef_Col)
                sAlmoxarifado = GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxarifadoBenef_Col)
    
                lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 55628
    
                'Verifica se há outras Requisições de Produto no mesmo Almoxarifado
                If UCase(sAlmoxarifado) = UCase(sAlmoxarifadoAtual) And UCase(objProduto.sCodigo) = UCase(sProdutoFormatado) Then
    
                    'Verifica se há alguma QuanTidade informada
                    If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_QuantidadeBenef_Col))) <> 0 Then
    
                        sUnidadeProd = GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMedBenef_Col)
    
                        dQuantidadeProd = CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_QuantidadeBenef_Col))
    
                        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUnidadeProd, sUnidadeAtual, dFator)
                        If lErro <> SUCESSO Then gError 55629
    
                        dQuantTotal = dQuantTotal + (dQuantidadeProd * dFator)
    
                    End If
    
                End If
    
            End If
    
        Next
    
        dQuantTotal = dQuantTotal + dQuantAtual

        If dQuantTotal > CDbl(QuantDisponivelBenef.Caption) Then gError 55630

    End If

    Testa_QuantRequisitada = SUCESSO

    Exit Function

Erro_Testa_QuantRequisitada:

    Testa_QuantRequisitada = gErr

    Select Case gErr

        Case 55625, 55626, 55628, 55629

        Case 55627
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, sCodProduto)

        Case 55630
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_REQ_MAIOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156554)

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
    If lErro <> SUCESSO And lErro <> 25060 Then gError 52354
    
    If lErro = 25060 Then gError 52355
    
    'Formata o Produto para BD
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 52356
    
    objEstoqueProduto.sProduto = sProdutoFormatado
    objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
    
    lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 49991 Then gError 52357
    
    If lErro <> 49991 Then
        
        lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then gError 52358
    
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

    Preenche_ContaContabilEst = gErr
    
        Select Case gErr
            
            Case 52354, 52356, 52357
            
            Case 52355
                 Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)
        
            Case 52358
                Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objEstoqueProduto.sContaContabil)
                
            Case Else
                Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156555)
        
        End Select
        
        Exit Function
        
End Function

Public Sub DescricaoItem_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

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
            If lErro <> SUCESSO Then gError 52287
    
            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedorBenef, colCodigoNome)
            If lErro <> SUCESSO Then gError 52288
    
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

    Select Case gErr

        Case 52287, 52288

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156556)

    End Select

    Exit Sub

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
            If lErro <> SUCESSO Then gError 92052
    
            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 92053
    
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

        Case 92052, 92053

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156557)

    End Select

    Exit Sub

End Sub

Public Sub FornNFOrigLabel_Click()

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 92056

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(FornNFOrig.Text)) = 0 Then gError 92057

        sFornecedor = FornNFOrig.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 92058

        If lErro = 18272 Then gError 92059

        'coloca na tela
        FilialFornNFOrig.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 92060

    Exit Sub

Erro_FilialFornNFOrig_Validate:

    Cancel = True

    Select Case gErr

        Case 92056, 92058
        
        Case 92057
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 92059
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, FornNFOrig.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 92060
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, FilialFornNFOrig.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156558)

    End Select

    Exit Sub

End Sub

Public Sub ClienteNFOrig_Change()

    iAlterado = REGISTRO_ALTERADO
    iClienteNFOrigAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ClienteNFOrig_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_ClienteNFOrig_Validate
   
    If iClienteNFOrigAlterado = 1 Then
   
        If Len(Trim(ClienteNFOrig.Text)) > 0 Then
    
            'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Cliente_Le(ClienteNFOrig, objCliente, iCodFilial)
            If lErro <> SUCESSO Then gError 92054
    
            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then gError 92055
    
            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", FilialCliNFOrig, colCodigoNome)
    
            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", FilialCliNFOrig, iCodFilial)
    
        ElseIf Len(Trim(ClienteNFOrig.Text)) = 0 Then
    
            FilialCliNFOrig.Clear
    
        End If
    
        iClienteNFOrigAlterado = 0
    
    End If
    
    Exit Sub

Erro_ClienteNFOrig_Validate:

    Cancel = True

    Select Case gErr

        Case 92054, 92055

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156559)

    End Select

    Exit Sub

End Sub

Public Sub ClienteNFOrigLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(ClienteNFOrig.Text)) > 0 Then objCliente.sNomeReduzido = ClienteNFOrig.Text

    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteNFOrig)

End Sub

Private Sub objEventoClienteNFOrig_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    ClienteNFOrig.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call ClienteNFOrig_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub FilialCliNFOrig_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub FilialCliNFOrig_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialCliNFOrig_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialCliNFOrig.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialCliNFOrig.Text = FilialCliNFOrig.List(FilialCliNFOrig.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialCliNFOrig, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 92061

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(ClienteNFOrig.Text)) = 0 Then gError 92062

        sCliente = ClienteNFOrig.Text

        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 92063

        If lErro = 17660 Then gError 92064

        'Coloca na tela
        FilialFornNFOrig.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 92065

    Exit Sub

Erro_FilialCliNFOrig_Validate:

    Cancel = True

    Select Case gErr

        Case 92061, 92063

        Case 92062
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 92064
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, ClienteNFOrig.Text)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 92065
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialCliNFOrig.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156560)

    End Select

    Exit Sub

End Sub

Public Sub Serie_Click()

    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = REGISTRO_ALTERADO

End Sub

Public Sub Total_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dTotal As Double

On Error GoTo Erro_Total_Validate

    'Se estiver preenchido
    If Len(Trim(Total.Text)) > 0 Then
    
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(Total.Text)
        If lErro <> SUCESSO Then gError 43803

        dTotal = CDbl(Total.Text)
        
        'Coloca o valor formatado na tela
        Total.Text = Format(dTotal, "Standard")

    End If

    Exit Sub

Erro_Total_Validate:

    Cancel = True


    Select Case gErr

        Case 43803

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156561)

    End Select

    Exit Sub

End Sub

Public Sub TipoNFiscal_Click()

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TipoNFiscal_Click

    If TipoNFiscal.ListIndex = -1 Then Exit Sub
    
    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 31411

    'se não estiver cadastrado ==> erro
    If lErro = 31415 Then gError 31416
    
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
        BotaoLimparNF.Visible = False
        'ChvNFe.Enabled = True
        'Cor Vermelha para label pois é obrigatorio
        NFiscalLabel.ForeColor = &H80&
    End If
    
    If NFiscalInterna.Visible Then
        'Seleciona série padrão
        lErro = Seleciona_Serie_Padrao()
        If lErro <> SUCESSO Then gError 31411
    End If
    
    'habilita/desabilita os campos relativos a cliente/fornecedor associados a nota fiscal original
    Call Habilita_Desabilita_CliForn_NFOrig
        
    iEmitenteNF = objTipoDocInfo.iEmitente
    
    If iEmitenteNF <> DOCINFO_EMPRESA Then
    
        NaturezaOp.Text = NaturezaOp_Conv34(objTipoDocInfo.sNatOpExtPadrao, IIf(MaskedParaDate(DataEntrada) = DATA_NULA, MaskedParaDate(DataEmissao), MaskedParaDate(DataEntrada)))
        
    End If
    
    'Rastreamento
    If TipoNFiscal.ListIndex <> -1 Then
        If (TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFEEBEN) Or (TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFIEEBEN) Then
            If TabStrip1.Tabs.Count = TAB_BENEFICIAMENTO - 1 Then
                TabStrip1.Tabs.Add TAB_BENEFICIAMENTO, , "Beneficiamento"
            End If
        ElseIf TabStrip1.Tabs.Count = TAB_BENEFICIAMENTO Then
               TabStrip1.Tabs.Remove (TAB_BENEFICIAMENTO)
        End If
    End If
    
''*** incluidos p/tratamento de tributacao *******************************
    Call gobjTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
    If IsDate(DataEntrada.Text) Then Call gobjTribTab.DataEntrada_Alterada(StrParaDate(DataEntrada.Text))
    Call gobjTribTab.Serie_Alterada(Serie.Text)
'*** fim tributacao

    'Seleciona os subtipos contábeis existentes para a transação correspondente ao tipo de Nota Fiscal selecionada
    lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_ESTOQUE, objTipoDocInfo.iSubTipoContabil)
    If lErro <> SUCESSO Then gError 79333
    
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 79333
    End If
    
    Exit Sub

Erro_TipoNFiscal_Click:

    Select Case gErr

        Case 31411, 79333
        
        Case 31416
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156562)

    End Select

    Exit Sub

End Sub

Private Sub Habilita_Desabilita_CliForn_NFOrig()
'habilita/desabilita os campos relativos a cliente/fornecedor associados a nota fiscal original

    If Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_CFEV Then
    
        ClienteNFOrig.Visible = True
        ClienteNFOrigLabel.Visible = True
        FilialCliNFOrig.Visible = True
        FilialCliNFOrigLabel.Visible = True
        FornNFOrig.Visible = False
        FornNFOrigLabel.Visible = False
        FilialFornNFOrig.Visible = False
        FilialFornNFOrigLabel.Visible = False
        ClienteNFOrig.Enabled = True
        ClienteNFOrigLabel.Enabled = True
        FilialCliNFOrig.Enabled = True
        FilialCliNFOrigLabel.Enabled = True
        FornNFOrig.Enabled = False
        FornNFOrigLabel.Enabled = False
        FilialFornNFOrig.Enabled = False
        FilialFornNFOrigLabel.Enabled = False
    
    ElseIf (Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_CFECT) Or (Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_DDAI) Or (Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_NFSET) Then
    
        ClienteNFOrig.Visible = False
        ClienteNFOrigLabel.Visible = False
        FilialCliNFOrig.Visible = False
        FilialCliNFOrigLabel.Visible = False
        FornNFOrig.Visible = True
        FornNFOrigLabel.Visible = True
        FilialFornNFOrig.Visible = True
        FilialFornNFOrigLabel.Visible = True
        ClienteNFOrig.Enabled = False
        ClienteNFOrigLabel.Enabled = False
        FilialCliNFOrig.Enabled = False
        FilialCliNFOrigLabel.Enabled = False
        FornNFOrig.Enabled = True
        FornNFOrigLabel.Enabled = True
        FilialFornNFOrig.Enabled = True
        FilialFornNFOrigLabel.Enabled = True
    
    Else
    
        ClienteNFOrig.Enabled = False
        ClienteNFOrigLabel.Enabled = False
        FilialCliNFOrig.Enabled = False
        FilialCliNFOrigLabel.Enabled = False
        FornNFOrig.Enabled = False
        FornNFOrigLabel.Enabled = False
        FilialFornNFOrig.Enabled = False
        FilialFornNFOrigLabel.Enabled = False
        
    End If

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Call CF2(Me, "NFEnt_Form_Unload", Cancel)
    Set gobjInfoUsu = Nothing
    
    Set gobjNFiscal = Nothing
    Set objEventoSerie = Nothing
    Set objEventoProduto = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoNatureza = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoFornecedorBenef = Nothing
    Set objEventoFornNFOrig = Nothing
    Set objEventoClienteNFOrig = Nothing
    
    Set objEventoRecebimento = Nothing
    Set objEventoSerieOriginal = Nothing
    Set objEventoNFiscalOriginal = Nothing
    Set objEventoTransportadora = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoEstoqueBenef = Nothing
    Set objEventoProdutoBenef = Nothing
    Set objEventoCcl = Nothing
    Set objEventoChvNFe = Nothing

    Set gobjTelaProjetoInfo = Nothing
    Set gobjInfoAdic = Nothing

    Set objGrid1 = Nothing
    Set objGridBenef = Nothing
    Set objContabil = Nothing
    Set objGridItens = Nothing
    
'   Rastreamento
    Set gobjRastreamento = Nothing
    
    'distribuicao
    Set gobjDistribuicao = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
''*** incluidos p/tratamento de tributacao *******************************
    If Not (gobjTribTab Is Nothing) Then Call gobjTribTab.Desativar
    Set gobjTribTab = Nothing
'*** fim tributacao

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)
    
    Set gobjAnotacao = Nothing
    
    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156563)

    End Select

    Exit Sub

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156564)
    
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

Public Sub NaturezaOp_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNaturezaOp As New ClassNaturezaOp
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_NaturezaOp_Validate
   
    If iNaturezaOpAlterada = 0 Then Exit Sub
    
    If Len(Trim(NaturezaOp.ClipText)) = 0 Then Exit Sub
    
    objNaturezaOp.sCodigo = NaturezaOp.Text
    
    lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
    If lErro <> SUCESSO And lErro <> 17958 Then gError 35652
    
    If lErro = 17958 Then gError 35653
    
    'se for nota interna garantir que tenha codigo < 500
    If iEmitenteNF = DOCINFO_EMPRESA Then
    
        If objNaturezaOp.sCodigo >= NATUREZA_SAIDA_COD_INICIAL Then gError 56986
        
        If (gobjTribTab.sNatOpInterna <> objNaturezaOp.sCodigo) Then
        
            gobjTribTab.sNatOpInterna = objNaturezaOp.sCodigo
            Call NatOpInterna_Validate(Cancel)
            
        End If

    End If
    
    'se for nota externa garantir que tenha codigo > 500
    If iEmitenteNF <> DOCINFO_EMPRESA And objNaturezaOp.sCodigo <= NATUREZA_SAIDA_COD_INICIAL Then gError 56987

    iNaturezaOpAlterada = 0
    
    Exit Sub

Erro_NaturezaOp_Validate:

    Cancel = True


    Select Case gErr

        Case 35652

        Case 35653
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_NATUREZA_OPERACAO", NaturezaOp.Text)
            If vbMsgRes = vbYes Then
                
                Call Chama_Tela("NaturezaOperacao", objNaturezaOp)
                
            Else
            End If
        
        
        Case 56986
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_ENTRADA", gErr)
        
        Case 56987
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156565)

    End Select

    Exit Sub

End Sub

Public Sub NatOpInterna_Validate(Cancel As Boolean)
    Call gobjTribTab.NatOpInterna_Validate(Cancel)
End Sub

Public Sub Form_Activate()

Dim lErro As Long

On Error GoTo Erro_Form_Activate

    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError 59340
        
    Call TelaIndice_Preenche(Me)

    Exit Sub
     
Erro_Form_Activate:

    Select Case gErr
          
        Case 59340
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156566)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Function CargaPosFormLoad(Optional bTrazendoDoc As Boolean = False) As Long

Dim iSubTipo As Integer
Dim lErro As Long
Dim sMascaraCcl As String

On Error GoTo Erro_CargaPosFormLoad

    If (giPosCargaOk = 0) Then

        'p/permitir o redesenho da tela
        DoEvents
        
        gbCarregandoTela = True

        ValorUnitario.Format = FORMATO_PRECO_UNITARIO_EXTERNO

        ''*** incluidos p/tratamento de tributacao *******************************
    
        Call GridItens_InicializaColunas
        
        Set gobjTribTab = New ClassTribTab
        lErro = gobjTribTab.Ativar(Me, 1)
        If lErro <> SUCESSO Then gError 55663
    
        '*** fim tributacao
    
        'Rastreamento
        Set gobjRastreamento = New ClassRastreamento
        Set gobjRastreamento.objTela = Me

        Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
        Set gobjTelaProjetoInfo.objUserControl = objUserControl
        Set gobjTelaProjetoInfo.objTela = Me

        'distribuicao
        Set gobjDistribuicao = CreateObject("RotinasMat.ClassMATDist")
        Set gobjDistribuicao.objTela = Me
        Set gobjDistribuicao.objRastreamento = gobjRastreamento

        'Carrega os Tipos de Documentos relacionadas à tela
        lErro = Carrega_TiposDocInfo()
        If lErro <> SUCESSO Then gError 31392
    
        'Inicializa as variáveis globais da tela(objGrid, objEventos)
        Set objGridBenef = New AdmGrid
        Set objGridItens = New AdmGrid
        Set objEventoSerie = New AdmEvento
        Set objEventoProduto = New AdmEvento
        Set objEventoNFiscal = New AdmEvento
        Set objEventoNatureza = New AdmEvento
        Set objEventoFornecedor = New AdmEvento
        Set objEventoFornecedorBenef = New AdmEvento
        Set objEventoFornNFOrig = New AdmEvento
        Set objEventoClienteNFOrig = New AdmEvento
        Set objEventoRecebimento = New AdmEvento
        Set objEventoSerieOriginal = New AdmEvento
        Set objEventoNFiscalOriginal = New AdmEvento
        Set objEventoTransportadora = New AdmEvento
        Set objEventoContaContabil = New AdmEvento
        Set objEventoProdutoBenef = New AdmEvento
        Set objEventoEstoqueBenef = New AdmEvento
        Set objEventoCcl = New AdmEvento
        Set objEventoChvNFe = New AdmEvento
        
        'Carrega as Combos que armazenam as Séries
        lErro = Carrega_Serie()
        If lErro <> SUCESSO Then gError 31393
    
        'Carrega a combo PlacaUF com as siglas das Estados
        lErro = Carrega_PlacaUF()
        If lErro <> SUCESSO Then gError 31395
    
        'Carrega a combo com o Código e o Nome Reduzido das Transportadoras do BD
        lErro = Carrega_Transportadora()
        If lErro <> SUCESSO Then gError 31402
    
        'Incluído por Luiz Nogueira em 21/08/03
        'Carrega a combo VolumeEspecie
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
        If lErro <> SUCESSO Then gError 102448
    
        'Incluído por Luiz Nogueira em 21/08/03
        'Carrega a combo VolumeMarca
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
        If lErro <> SUCESSO Then gError 102449
    
        'Inicializa a Máscara de Produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
        If lErro <> SUCESSO Then gError 31406
        
        'Inicializa a Máscara de Produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoBenef)
        If lErro <> SUCESSO Then gError 52361
        
        'Inicializa mascara de contaContabilEst
        lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilEst)
        If lErro <> SUCESSO Then gError 52362
        
        'Inicializa mascara de contaContabilEst
        lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabilProducao)
        If lErro <> SUCESSO Then gError 52363
        
        'Inicializa a mascara do CCL
        sMascaraCcl = String(STRING_CCL, 0)
    
        'le a mascara dos centros de custo/lucro
        lErro = MascaraCcl(sMascaraCcl)
        If lErro <> SUCESSO Then gError 52872
    
        Ccl.Mask = sMascaraCcl
        
        Quantidade.Format = FORMATO_ESTOQUE
    
        'Faz a Inicialização do Grid de Itens
        lErro = Inicializa_Grid_Itens(objGridItens)
        If lErro <> SUCESSO Then gError 31403
        
        'Faz a inicialização do grid de Movimentos
        lErro = Inicializa_GridMovimentos(objGridBenef)
        If lErro <> SUCESSO Then gError 52364
    
        'Inicializa a contabilidade carregando apenas os modelos contábeis referentes ao TipoDocInfo default
        If objContabil.objGridInt Is Nothing Then

            'Obtém o código do subtipo contábil referente ao TipoDocInfo default
            lErro = CF("Obtem_Subtipo_Contabil", "TiposDocInfo", DOCINFO_NFEEN, iSubTipo)
            If lErro <> SUCESSO And lErro <> 79671 Then gError 79692
            
            'Se não encontrou o subtipo => erro
            If lErro = 79671 Then gError 79693
            
            'Inicialização da parte de contabilidade
            lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_ESTOQUE, iSubTipo)
            If lErro <> SUCESSO Then gError 79345

        End If

        'Inicializa o grid de Rastreamento
        lErro = gobjRastreamento.Inicializa_Grid_Rastreamento()
        If lErro <> SUCESSO Then gError 83145

        'Inicializa o grid de Distribuicao
        lErro = gobjDistribuicao.Inicializa_GridDist()
        If lErro <> SUCESSO Then gError 89525
       
        gbCarregandoTela = False
        
        iAlterado = 0
        iDescontoAlterado = 0
    
        giPosCargaOk = 1
        
        Set gobjNFiscal = New ClassNFiscal
    
        Call ValorTotal_Calcula
    
    End If
    
    CargaPosFormLoad = SUCESSO
    
    Exit Function
     
Erro_CargaPosFormLoad:

    gbCarregandoTela = False
    
    CargaPosFormLoad = gErr
    
    Select Case gErr

        Case 27838, 27839, 31392, 31393, 31395, 31402, 31403, 31406, 39615, 52361, 52362, 52363, 52364, 52872, 55663, 79345, 79692, 83145, 89460, 89525, 102448, 102449

        Case 79693
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPODOCINFO_NAO_ENCONTRADO", gErr, DOCINFO_NFEEN)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156567)
     
    End Select
     
    iAlterado = 0
    
    Exit Function

End Function


Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Inicializa o Frame atual como o 1
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
    
    'Coloca na Data de Entrada a Data Atual
    DataEntrada.PromptInclude = False
    DataEntrada.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEntrada.PromptInclude = True
    
    Set gobjNFiscal = New ClassNFiscal
    
    Call CF2(Me, "NFEnt_Form_Load")
    
    MsgAutomatica.Value = vbChecked
    Mensagem.MaxLength = MENSAGEM_NF_MAXLENGTH
    MensagemCorpo.MaxLength = MENSAGEM_NF_MAXLENGTH
    
    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207457
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 207457

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156568)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

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

    Inicializa_GridMovimentos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156569)

    End Select

    Exit Function

End Function

Private Function Carrega_TiposDocInfo() As Long
'Carrega na os Tipo de Documentos relacionados com a tela de Nota Fiscal de Entrada

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
            If objTipoDocInfo.iPadrao = Padrao Then TipoNFiscal.ListIndex = TipoNFiscal.NewIndex
        End If
    Next

    Carrega_TiposDocInfo = SUCESSO

    Exit Function

Erro_Carrega_TiposDocInfo:

    Carrega_TiposDocInfo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156570)

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
    If lErro <> SUCESSO Then gError 31394

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

    Carrega_Serie = gErr

    Select Case gErr

        Case 31394

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156571)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156572)

    End Select

End Function

Sub GridItens_InicializaColunas()
    
    Call CF2(Me, "NFE_GridItens_InicializaColunas")

End Sub

Function NFE_GridItens_InicializaColunas(ByVal objCT As Object) As Long
    
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
    
    NFE_GridItens_InicializaColunas = SUCESSO

End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
    Inicializa_Grid_Itens = CF2(Me, "NFE_Inicializa_Grid_Itens", objGridInt)
End Function

Public Function NFE_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    
'distribuicao
'    objGridInt.colColuna.Add ("Almoxarifado")

    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Valor Unitário")
    objGridInt.colColuna.Add ("% Desc.")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Líquido")
    objGridInt.colColuna.Add ("Preço Bruto")

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

    Call GridItens_InicializaColunas
    
    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_NF + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6
    
    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NFE_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim colTabelaPrecoItem As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "NFiscalEntrada"

    'Lê os dados da Tela Notas Fiscais a Pagar
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 31423

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objNFiscal.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "TipoNFiscal", objNFiscal.iTipoNFiscal, 0, "TipoNFiscal"
    colCampoValor.Add "NaturezaOp", objNFiscal.sNaturezaOp, STRING_BUFFER_MAX_TEXTO, "NaturezaOp"
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
    colCampoValor.Add "FornecedorBenef", objNFiscal.lFornecedorBenef, 0, "FornecedorBenef"
    colCampoValor.Add "FilialFornBenef", objNFiscal.iFilialFornBenef, 0, "FilialFornBenef"
    colCampoValor.Add "FreteRespons", objNFiscal.iFreteRespons, 0, "FreteRespons"
    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_BUFFER_MAX_TEXTO, "VolumeNumero"
    colCampoValor.Add "Observacao", objNFiscal.sObservacao, STRING_NFISCAL_OBSERVACAO, "Observacao"
    colCampoValor.Add "Status", objNFiscal.iStatus, 0, "Status"

    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
'    colSelecao.Add "Status", OP_DIFERENTE, STATUS_CANCELADO
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 31423

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156573)

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
        If lErro <> SUCESSO Then gError 31424

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 31424

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156574)

    End Select

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

Public Sub objEventoFornecedor_evSelecao(obj1 As Object)

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

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 35020

    'Limpa a Tela
    Call Limpa_Tela_NFiscal1

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 35020

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156575)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 35042

    'Limpa a Tela
    Call Limpa_Tela_NFiscal

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 35042

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156576)

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
        If lErro <> SUCESSO Then gError 64010

        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 31443
        
        If lErro <> SUCESSO Then gError 31444

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

        'Lê o Tipo de Documento com o Código de objNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 31486
        
        'Se não encontrar --> Erro
        If lErro = 31415 Then gError 31487
        
'        If giTipoVersao = VERSAO_FULL Then
            If objTipoDocInfo.sNomeTelaNFiscal <> Me.Name Then gError 31396
'        End If
        
        'Traz a nota para a tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 31445

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 31443, 31445, 31486, 64010

        Case 31444
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)
            Call Limpa_Tela_NFiscal
            iAlterado = 0

        Case 31487
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 31396
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_DIFERENTE_NF_ENTRADA", gErr, objNFiscal.iTipoNFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156577)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO

    Call Fornecedor_Preenche

End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer, lErro As Long
Dim iIndice As Integer
'Rastreamento
'Dim sProduto As String
'Dim sAlmoxarifado As String

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
        If lErro <> SUCESSO Then gError 55567
        
''*** incluidos p/tratamento de tributacao *******************************
        
        lErro = gobjTribTab.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 22884
        
'*** fim tributacao
    
        Call ValorTotal_Calcula
        
        If Not gbCarregandoTela Then
            lErro = gobjTribTab.Pesos_Recalcula()
            If lErro <> SUCESSO Then gError 177160
        End If

'distribuicao
'        'Rastreamento
'        lErro = gobjRastreamento.Exclusao_Item_Grid(iItemAtual)
'        If lErro <> SUCESSO Then gError 83203

        'distribuicao
        lErro = gobjDistribuicao.Exclusao_Item_GridDist(iItemAtual)
        If lErro <> SUCESSO Then gError 89514
        
        lErro = CF2(Me, "NFE_Calcula_PesoKg_Total")
        If lErro <> SUCESSO Then gError 177160
        
    End If

    Exit Sub
    
Erro_GridItens_KeyDown:

    Select Case gErr

        Case 22884, 55567, 83203, 89514, 177160

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156578)

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

Public Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

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

Public Sub NaturezaOp_Change()

    iAlterado = REGISTRO_ALTERADO
    iNaturezaOpAlterada = 1
    
End Sub

Public Sub NFiscal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NFiscalLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As Collection
Dim lErro As Long

On Error GoTo Erro_NFiscalLabel_Click

    'Recolhe os dados da NFiscal que estão na tela
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 31436

    'Chama a tela de Brow sw Nota Fiscal de Entrada
    Call Chama_Tela("NFiscalEntradaLista", colSelecao, objNFiscal, objEventoNFiscal)

    Exit Sub

Erro_NFiscalLabel_Click:

    Select Case gErr

        Case 31436

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156579)

    End Select

    Exit Sub

End Sub

Public Sub NFiscalOriginal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NFiscalOriginalLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection
Dim lErro As Long
Dim iTipoDocInfo As Integer
Dim sTela As String, sFiltro As String

On Error GoTo Erro_NFiscalOriginalLabel_Click

    'Recolhe os dados da Nota Fiscal Entrada
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 31437

    'Guarda a Serie e o Número da Nota Fiscal Original da Tela
    objNFiscal.sSerie = SerieNFiscalOriginal.Text
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscalOriginal.Text)
    
    'Obtém o código do tipo da nota fiscal
    iTipoDocInfo = Codigo_Extrai(TipoNFiscal.Text)
    
    'Se a nota for do tipo de conhecimento de frete sobre venda
    If iTipoDocInfo = DOCINFO_CFEV Then
    
        'Adiciona à coleção o filtro de NomeCliente
        colSelecao.Add ClienteNFOrig.Text
        
        'Adiciona à coleção o filtro de NomeCliente
        colSelecao.Add ClienteNFOrig.Text
        
        'Guarda o nome do browser que será chamado
        sTela = "NFiscalNormalLista"
    
    'Se for do tipo de despesas c/despacho aduaneiro
    ElseIf iTipoDocInfo = DOCINFO_DDAI Then
    
        'Adiciona à coleção o filtro de TipoNFiscal
        colSelecao.Add DOCINFO_NFIEIMP
        colSelecao.Add DOCINFO_NFIEIMPSE
        colSelecao.Add DOCINFO_NFIEIMPPC
        
        'Guarda o nome do browser que será chamado
        'sTela = "NFiscalTipoDocInfoLista"
        sTela = "NFiscalInternaEntradaLista"
        sFiltro = "TipoNFiscal IN (?,?,?)"

    'Se for outro tipo
    Else

        'Guarda o nome do browser que será chamado
        sTela = "NFiscalEntradaTodasLista"
    
    End If
        
    'Chama o Browser
    Call Chama_Tela(sTela, colSelecao, objNFiscal, objEventoNFiscalOriginal, sFiltro)

    Exit Sub

Erro_NFiscalOriginalLabel_Click:

    Select Case gErr

        Case 31437

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156580)

    End Select

    Exit Sub

End Sub

Public Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim objNaturezaOp As ClassNaturezaOp

    Set objNaturezaOp = obj1

    'Coloca a natureza Operação na tela
    NaturezaOp.Text = objNaturezaOp.sCodigo

    Call NaturezaOp_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Public Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objNFiscal = obj1

    'Traz a NotaFiscal de entrada para a a tela
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 31418

    'Fecha o Comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case gErr

        Case 31418

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156581)

    End Select

    Exit Sub

End Sub

Public Sub objEventoNFiscalOriginal_evSelecao(obj1 As Object)

Dim bCancel As Boolean
Dim objNFiscal As ClassNFiscal

    Set objNFiscal = obj1

    'Preenche a Série e o Número da Nota Fiscal Original
    SerieNFiscalOriginal.Text = objNFiscal.sSerie
    NFiscalOriginal.Text = objNFiscal.lNumNotaFiscal

    If Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_CFEV Then
    
        'Preenche o Cliente da NF Orig
        ClienteNFOrig.Text = objNFiscal.lCliente
        Call ClienteNFOrig_Validate(bCancel)
        
        'Preenche a Filial da Cliente
        FilialCliNFOrig.Text = objNFiscal.iFilialCli
        Call FilialCliNFOrig_Validate(bCancel)
        
    ElseIf (Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_CFECT) Or (Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_DDAI) Or (Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_NFSET) Then

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

Public Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoEnxuto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    'Verifica se o Produto está preenchido
    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then

        Set objProduto = obj1

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 31409

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
        
        If Not (Me.ActiveControl Is Produto) Then
    
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
            
            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then gError 35655
            
            'mario inicio
            DescricaoItem.Text = objProduto.sDescricao
            'mario fim
        
        End If
        
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoProduto_evSelecao:

    Select Case gErr
    
        Case 31409
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 35655
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156582)
            
    End Select

    Exit Sub
    
End Sub

Private Sub objEventoRecebimento_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoRecebimento_evSelecao

    Set objNFiscal = obj1

    giRecebimento = True
    
    'Coloca a Nota Fiscal na Tela1
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 31485

    giRecebimento = False
    
    'Fecha o Comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoRecebimento_evSelecao:

    giRecebimento = False
    
    Select Case gErr

        Case 31485

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156583)

    End Select

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

Public Sub objEventoSerieOriginal_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    'Coloca a Série da Nota Fiscal Original na tela
    SerieNFiscalOriginal.Text = objSerie.sSerie

    Me.Show

End Sub

Public Sub objEventoTransportadora_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora

    Set objTransportadora = obj1

    'Coloca a Transportdora na Tela
    Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    Me.Show

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
    If lErro <> SUCESSO Then gError 35018

    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True


    Select Case gErr

        Case 35018

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156584)

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
    If lErro <> SUCESSO Then gError 35017

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True


    Select Case gErr

        Case 35017

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156585)

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

    'Verifica se existe o ítem na combo
    lErro = Combo_Item_Igual(PlacaUF)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 35853

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 35854

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True


    Select Case gErr

        Case 35853

        Case 35854
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156586)

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
        
        objNFiscal.lNumRecebimento = NumRecebimento.Text

        'Le o Recebimento
        lErro = CF("NFiscal_Le_Recebimento", objNFiscal)
        If lErro <> SUCESSO And lErro <> 61050 Then gError 61058
        
        'Se não encontrou --> Erro
        If lErro = 61050 Then gError 61059
        
        If objNFiscal.iTipoNFiscal = DOCINFO_NRCP Or objNFiscal.iTipoNFiscal = DOCINFO_NRCC Then gError 61121
                
        If objNFiscal.iStatus = STATUS_BAIXADO Then gError 64167
        
        'Lê o Tipo de documento
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 27623 Then gError 75426
        
        'Se o Tipo de documento é de compras
        If objTipoDocInfo.sNomeTelaNFiscal = RECEBIMENTO_MATERIAL_FCOM Or objTipoDocInfo.sNomeTelaNFiscal = NOTA_FISCAL_COMPRAS Or objTipoDocInfo.sNomeTelaNFiscal = NOTA_FISCAL_FATURA_COMPRAS Then gError 75427
        
    Else
        
        'Verifica se o tipo está Preenchido
        If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 42533
        
        objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
    
        'Lê o Tipo de documento
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 27623 Then gError 31493
        
        'Se não encontrar --> erro
        If lErro = 27623 Then gError 31494
        
        'Se o Tipo de documento é de compras
        If objTipoDocInfo.sNomeTelaNFiscal = RECEBIMENTO_MATERIAL_FCOM Or objTipoDocInfo.sNomeTelaNFiscal = NOTA_FISCAL_COMPRAS Or objTipoDocInfo.sNomeTelaNFiscal = NOTA_FISCAL_FATURA_COMPRAS Then gError 75428
        
        objNFiscal.iTipoDocInfo = objTipoDocInfo.iTipo
        
        'Se for Nota Fiscal Interna então o NumRecebimento tem que está Preenchido
        If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then gError 61060

        'Verifica se os campos essenciais estão preenchidos
        If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 31458
        If Len(Trim(Filial.Text)) = 0 Then gError 31459
        If Len(Trim(Serie.Text)) = 0 Then gError 31460
        If Len(Trim(NFiscal.ClipText)) = 0 Then gError 31461
        
        'Recolhe os dados da Tela
        lErro = Move_NF_Memoria(objNFiscal)
        If lErro <> SUCESSO Then gError 31470

        If objNFiscal.iTipoNFiscal = DOCINFO_NFIEN Then
            objNFiscal.iTipoNFiscal = DOCINFO_NRFP
        ElseIf objNFiscal.iTipoNFiscal = DOCINFO_NFEEN Then
            objNFiscal.iTipoNFiscal = DOCINFO_NRFF
        End If

        If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then gError 31462

        'Lê o Recebimento com estas características
        lErro = CF("NFiscal_Le_Num_Forn1", objNFiscal)
        If lErro <> SUCESSO And lErro <> 55885 Then gError 31463
    
        'Se não encontrar -->Erro
        If lErro = 55885 Then gError 31468

        If objNFiscal.iStatus = STATUS_BAIXADO Then gError 64168
    
    End If
    
    giRecebimento = True
    
    'Traz para a tela
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 31469

    giRecebimento = False
    
    Exit Sub

Erro_Recebimento_Click:

    giRecebimento = False
    
    Select Case gErr

        Case 31457
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 31458
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 31459
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_CHQPRE_NAO_PREENCHIDA", gErr)

        Case 31460
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 31461
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)

        Case 31462
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_NORMAL", gErr)

        Case 31463, 31469, 31470, 31493, 61058, 75426

        Case 31468
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_MATERIAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 31494
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO2", gErr, objTipoDocInfo.iCodigo)
        
        Case 61059
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NAO_CADASTRADO", gErr, objNFiscal.lNumRecebimento)
                
        Case 61060
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_NFINT_NAO_PREENCHIDO", gErr)
        
        Case 61121
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_CLIENTE", gErr, objNFiscal.lNumRecebimento)
        
        Case 64167, 64168
            Call Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_VINCULADO_NF", gErr, objNFiscal.lNumRecebimento, objNFiscal.lNumNotaFiscal)
        
        Case 75427, 75428
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_RECEBIMENTO_COMPRAS", gErr, objTipoDocInfo.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156587)

    End Select

    Exit Sub

End Sub

Public Sub Serie_Change()

    iAlterado = REGISTRO_ALTERADO
    iSerieAlterada = REGISTRO_ALTERADO
    
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
        If lErro <> SUCESSO And lErro <> 12253 Then gError 31488
        
        'Se a serie não está cadastrada --> verifica se é uma nota interna, se for ==> erro, pois a serie não está cadastrada
        If lErro <> SUCESSO Then
    
            'Se o tipo está preenchido
            If Len(Trim(TipoNFiscal.Text)) > 0 Then
        
                objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
                
                'Lê o Tipo de Documento
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO Then gError 31489
        
                'se a nota interna e a serie não está cadastrada ==> erro
                If objTipoDocInfo.iTipo = DOCINFO_NF_INT_ENTRADA Then gError 31490
        
            End If
    
            'Verifica se tamanho da série é maior do que o espaço no bd ==> erro
            If Len(Trim(Serie.Text)) > STRING_SERIE Then gError 31491
    
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

        Case 31488, 31489

        Case 31490
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, Serie.Text)

        Case 31491
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_SERIE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156588)

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

Public Sub SerieNFiscalOriginal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub SerieNFiscalOriginal_Click()

Dim lErro As Long
Dim lNumAuto As Long

On Error GoTo Erro_SerieNFiscalOriginal_Click

    iAlterado = REGISTRO_ALTERADO
        
    Exit Sub

Erro_SerieNFiscalOriginal_Click:

    Select Case gErr
    
        Case 41787
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156589)
    
    End Select
    
    Exit Sub
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156590)

    End Select

    Exit Sub

End Sub

Public Sub SerieOriginalLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'Recolhe a Série da Nota Fiscal Original da tela
    objSerie.sSerie = SerieNFiscalOriginal.Text

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerieOriginal)

End Sub

Public Sub SubTotal_Change()

    iAlterado = REGISTRO_ALTERADO

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
            If lErro <> SUCESSO Then gError 27840
        End If
        
'*** fim tributacao

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
                Parent.HelpContextID = IDH_ENTRADA_NF_SIMPLES_DADOS_PRINCIPAIS
                
            Case TAB_Itens
                Parent.HelpContextID = IDH_ENTRADA_NF_SIMPLES_ITENS
                        
            Case TAB_Complemento
                Call Calcula_Msg
                Parent.HelpContextID = IDH_ENTRADA_NF_SIMPLES_COMPLEMENTOS
            
            Case TAB_Tributacao
                Parent.HelpContextID = IDH_ENTRADA_NF_SIMPLES_TRIBUTACAO
            
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_ENTRADA_NF_SIMPLES_CONTABILIZACAO
                        
        End Select

    End If
    
    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 27840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156591)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Transportadora() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_Transportadora

    'Lê o Código e o NOme de Todas as Transportadoras do BD
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_NOME_TABELA, colCodigoNome)
    If lErro <> SUCESSO Then gError 31404

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

        Case 31404

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156592)

    End Select

    Exit Function

End Function

Public Sub TipoNFiscal_Change()

    iAlterado = REGISTRO_ALTERADO

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 31410
    
    'Se não conseguir --> Erro
    If lErro <> SUCESSO Then gError 31417
   
    'tenta ler a natureza de operacao
    objTipoDocInfo.iCodigo = TipoNFiscal.ItemData(TipoNFiscal.ListIndex)
    
    'Lê o Tipo da NF
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 55546

    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 55547
   
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
        If lErro <> SUCESSO Then gError 55546
    End If
    
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 55546
    End If
    
    Exit Sub

Erro_TipoNFiscal_Validate:

    Cancel = True

    Select Case gErr

        Case 31410, 55546

        Case 31417
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case 55547
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156593)

    End Select

    Exit Sub

End Sub

Public Sub Total_Change()

    iAlterado = REGISTRO_ALTERADO

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

    'Verifica se foi preenchida a ComboBox Transportadora
    If Len(Trim(Transportadora.Text)) = 0 Then Exit Sub
    
    'Verifica se está preenchida com o ítem selecionado na ComboBox Transportadora
    If Transportadora.Text = Transportadora.List(Transportadora.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Transportadora, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 31432
    
    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objTransportadora.iCodigo = iCodigo
        
        'Tenta ler Transportadora com esse código no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then gError 31433
        
        ''Não encontrou Transportadora no BD
        If lErro = 19250 Then gError 31434

        'Encontrou Transportadora no BD, coloca no Text da Combo
        Transportadora.Text = CStr(objTransportadora.iCodigo) & SEPARADOR & objTransportadora.sNomeReduzido

    End If

    'Não existe o ítem com a String na List da Combo
    If lErro = 6731 Then gError 31435

    Exit Sub

Erro_Transportadora_Validate:

    Cancel = True

    Select Case gErr

        Case 31432, 31433

        Case 31434
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr, objTransportadora.iCodigo)
             
        Case 31435
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_ENCONTRADA", gErr, Transportadora.Text)
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156594)

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
        If GridItens.Row = 0 Then gError 43715

        sProduto1 = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 83078
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr
    
        Case 43715
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 83078
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156595)
    
    End Select
    
    Exit Sub
    
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
            If lErro <> SUCESSO Then gError 35095

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 35096

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
            
            If gobjInfoAdic Is Nothing Then Set gobjInfoAdic = New ClassInfoAdic
            gobjInfoAdic.sNatureza = objFornecedor.sNatureza
            
            lErro = CF2(Me, "NFCF_Fornecedor_Validate", objFornecedor.lCodigo, iCodFilial)
            If lErro <> SUCESSO Then gError 131936
            
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

        Case 35095, 35096, 131936

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156596)

    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    lErro = Move_NF_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 56994
    
    'Recolhe os dados do Grid
    If objGridItens.iLinhasExistentes > 0 Then
        lErro = Move_GridItens_Memoria(objNFiscal)
        If lErro <> SUCESSO Then gError 31421
    End If
    
    'distribuicao
    lErro = gobjDistribuicao.Move_GridDist_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 89466
    
    'Recolhe os dados do Grid Movimentos
    If (objNFiscal.iTipoNFiscal = DOCINFO_NFEEBEN) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEEBEN) Then
        lErro = Move_GridMovimentos_Memoria(objNFiscal)
        If lErro <> SUCESSO Then gError 52375
    End If
       
    'mover a parte de tributacao
    lErro = gobjTribTab.PreencherNF(objNFiscal)
    If lErro <> SUCESSO Then gError 22854

    lErro = CF2(Me, "NFEnt_Move_Tela_Memoria", objNFiscal)
    If lErro <> SUCESSO Then gError 124253
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 22854, 31421, 52375, 56994, 89466, 124253

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156597)

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
        If lErro <> SUCESSO Then gError 31422
        
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
        
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Ccl_Col))) > 0 Then

            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", GridItens.TextMatrix(iIndice, iGrid_Ccl_Col), sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError 52824
            
            objItemNF.sCcl = sCclFormatada
            
        Else

            objItemNF.sCcl = ""

        End If
        
        'Adiciona na coleção de Ítens
        With objItemNF
            Set objItemNFNovo = objNFiscal.ColItensNF.Add(0, iIndice, .sProduto, .sUnidadeMed, .dQuantidade, .dPrecoUnitario, .dPercDesc, .dValorDesconto, DATA_NULA, .sDescricaoItem, 0, 0, 0, 0, 0, colAlocacoes, .iAlmoxarifado, .sAlmoxarifadoNomeRed, .sCcl, STATUS_LANCADO, 0, "", 0, 0, 0, "", 0, 0)
        End With
        
        '********************* TRATAMENTO DE GRADE *****************
        lErro = gobjDistribuicao.Move_DistribuicaoGrade_Memoria(gobjNFiscal.ColItensNF(iIndice))
        If lErro <> SUCESSO Then gError 86375
        
        Call Move_ItensGrade_Tela(objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade, gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade)
    
        lErro = CF2(Me, "NFE_Move_GridItem_Memoria", objItemNFNovo, iIndice)
        If lErro <> SUCESSO Then gError 181416

        Set objItemNFNovo.objInfoAdicDocItem = gobjNFiscal.ColItensNF(iIndice).objInfoAdicDocItem

    Next

    Move_GridItens_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItens_Memoria:

    Move_GridItens_Memoria = gErr

    Select Case gErr

        Case 31422, 35087, 52824, 86375, 181416

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156598)

    End Select

    Exit Function

End Function

Public Function Traz_NFiscal_Tela(objNFiscal As ClassNFiscal) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim objNFiscalOriginal As New ClassNFiscal
Dim iIndice As Integer
Dim bCancel As Boolean
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim bAlterouCT As Boolean
Dim objItemNF As ClassItemNF
Dim dPercDesc As Double

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

    'Limpa a tela
    Call Limpa_Tela_NFiscal
    
    If objNFiscal.lNumRecebimento = 0 Then
        NumRecebimento.Text = ""
    Else
        NumRecebimento.Text = objNFiscal.lNumRecebimento
    End If
        
    'Tenta ler a nota Fiscal passada por parametro
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 31425
    
    'Lê os Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 31425
    
    'Lê as mensagens
    lErro = CF("NFiscal_Le_Msg", objNFiscal)
    If lErro <> SUCESSO Then gError 31425
    
    'distribuicao
    'Lê a Distribuição dos itens da Nota Fiscal
    lErro = CF("AlocacoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 89499
    
    objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
    'Le o Tipo para saber se tem Movimento de Estoque
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO Then gError 78835
    
    If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
        'Carrega ItensNF com Rastreamentos
        lErro = gobjRastreamento.Carrega_RastroItensNF(objNFiscal)
        If lErro <> SUCESSO Then gError 75910
    End If
    
    'Lê os ítens do Movimento se for Benefiamento e preenche o grid
    If (objNFiscal.iTipoNFiscal = DOCINFO_NFEEBEN) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFIEEBEN) Then
        
        lErro = CF("MovEstoqueItensBenef_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 52384 Then gError 52376
        
        lErro = Preenche_GridMovimentos(objNFiscal)
        If lErro <> SUCESSO Then gError 52377
            
    End If
    
    Set gobjNFiscal = objNFiscal
    
    lErro = CF("NFiscal_Le_ChvNFe", objNFiscal)
    If lErro <> SUCESSO Then gError 42745
    
    ChvNFe.PromptInclude = False
    ChvNFe.Text = objNFiscal.sChvNFe
    ChvNFe.PromptInclude = True
    gsChvAnt = objNFiscal.sChvNFe
    
    'preenche o Fornecedor
    Fornecedor.Text = objNFiscal.lFornecedor
    Call Fornecedor_Validate(bCancel)
    
    'Preenche a Filial do Fornecedor
    Filial.Text = objNFiscal.iFilialForn
    Call Filial_Validate(bCancel)

    lErro = CF2(Me, "NFCF_Fornecedor_Traz_NF_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 131937

    If objNFiscal.iTipoNFiscal = DOCINFO_NRFF Then
        objNFiscal.iTipoNFiscal = DOCINFO_NFEEN
    ElseIf objNFiscal.iTipoNFiscal = DOCINFO_NRFP Then
        objNFiscal.iTipoNFiscal = DOCINFO_NFIEN
    End If

    TipoNFiscal.Text = objNFiscal.iTipoNFiscal
    Call TipoNFiscal_Validate(bCancel)

    If objNFiscal.lNumNotaFiscal = 0 Then
        NFiscalInterna.Caption = ""
        NFiscal.Text = ""
    Else
        NFiscalInterna.Caption = CStr(objNFiscal.lNumNotaFiscal)
        NFiscal.Text = CStr(objNFiscal.lNumNotaFiscal)
    End If
    
    If objNFiscal.sNaturezaOp <> "" Then NaturezaOp.Text = objNFiscal.sNaturezaOp

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
    
    'Se não trouxer para tela ao regravar vai excluir a NF anterior e apagar
    'qualquer rastro de uso do recebimento deixando-o perdido conforme ocorreu na DLE
    If objNFiscal.lNumRecebimento <> 0 Then NumRecebimento.Text = objNFiscal.lNumRecebimento

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

    dtVencimento.PromptInclude = False
    If objNFiscal.dtDataVencimento <> DATA_NULA Then dtVencimento.Text = Format(objNFiscal.dtDataVencimento, "dd/mm/yy")
    dtVencimento.PromptInclude = True

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
    
    If objNFiscal.iTipoNFiscal = DOCINFO_NFEEBF Or objNFiscal.iTipoNFiscal = DOCINFO_NFIEBF Or objNFiscal.iTipoNFiscal = DOCINFO_NFECCO Then

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
    MensagemCorpo.Text = objNFiscal.sMensagemCorpoNota
    Observacao.Text = objNFiscal.sObservacao
    
    If objNFiscal.dPesoLiq > 0 Then PesoLiquido.Text = CStr(objNFiscal.dPesoLiq)
    If objNFiscal.dPesoBruto > 0 Then PesoBruto.Text = CStr(objNFiscal.dPesoBruto)

    'preenche o Grid com os Ítens da Nota Fiscal
    lErro = Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 31431
    
    'distribuicao
    'Preenche o Grid com as Distribuições dos itens da Nota Fiscal
    lErro = gobjDistribuicao.Preenche_GridDistribuicao(objNFiscal)
    If lErro <> SUCESSO Then gError 89500
    
    'Set gobjNFiscal = objNFiscal
      
    'Se o código da Nota Fiscal Original estiver preenchido
    If objNFiscal.lNumIntNotaOriginal <> 0 Then
    
        objNFiscalOriginal.lNumIntDoc = objNFiscal.lNumIntNotaOriginal
        
        'Lê a Nota Fiscal original
        lErro = CF("NFiscal_Le", objNFiscalOriginal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 35132

        If lErro = 35132 Then gError 35133
        
        'Coloca os dados relativos  Nota Fiscal Original na Tela
        SerieNFiscalOriginal.Text = objNFiscalOriginal.sSerie
        NFiscalOriginal.Text = objNFiscalOriginal.lNumNotaFiscal
        
        If objNFiscal.iTipoNFiscal = DOCINFO_CFEV Then
        
            'Preenche o Cliente da NF Orig
            ClienteNFOrig.Text = objNFiscalOriginal.lCliente
            Call ClienteNFOrig_Validate(bCancel)
            
            'Preenche a Filial da Cliente
            FilialCliNFOrig.Text = objNFiscalOriginal.iFilialCli
            Call FilialCliNFOrig_Validate(bCancel)
            
        ElseIf (objNFiscal.iTipoNFiscal = DOCINFO_CFECT) Or (objNFiscal.iTipoNFiscal = DOCINFO_DDAI) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFSET) Then
    
            'Preenche o Fornecedor da NF Orig
            FornNFOrig.Text = objNFiscalOriginal.lFornecedor
            Call FornNFOrig_Validate(bCancel)
            
            'Preenche a Filial do Fornecedor
            FilialFornNFOrig.Text = objNFiscalOriginal.iFilialForn
            Call FilialFornNFOrig_Validate(bCancel)
    
        End If
        
    End If

    If Not giRecebimento Then
    
        'traz os dados contábeis para a tela (contabilidade)
        lErro = objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
        If lErro <> SUCESSO And lErro <> 36326 Then gError 39616

        'Preenche o Status da Nota Fiscal
        If objNFiscal.iStatus = STATUS_LANCADO Then
            Status.Caption = STRING_STATUS_LANCADO
        ElseIf objNFiscal.iStatus = STATUS_BAIXADO Then
            Status.Caption = STRING_STATUS_BAIXADO
        ElseIf objNFiscal.iStatus = STATUS_CANCELADO Then
            Status.Caption = STRING_STATUS_CANCELADO
        End If
    
        'Lê a parte de Tributação
        lErro = CF("NFiscal_Le_Tributacao", objNFiscal)
        If lErro <> SUCESSO Then gError 22875
        
        'Lê a parte de tributacao dos Ítens da Nota Fiscal
        lErro = CF("NFiscalItens_Le_Trib", objNFiscal)
        If lErro <> SUCESSO Then gError 22880

    ''*** incluidos p/tratamento de tributacao *******************************
        lErro = gobjTribTab.Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 22842
    '*** fim tributacao
    
    Else
    
        Call gobjTribTab.ValorFrete_Validate(bSGECancelDummy, objNFiscal.dValorFrete)
        Call gobjTribTab.ValorSeguro_Validate(bSGECancelDummy, objNFiscal.dValorSeguro)
        Call gobjTribTab.ValorDespesas_Validate(bSGECancelDummy, objNFiscal.dValorOutrasDespesas)
        
        For iIndice = 1 To objNFiscal.ColItensNF.Count
            lErro = gobjTribTab.Inclusao_Item_Grid(iIndice, objNFiscal.ColItensNF.Item(iIndice).sProduto)
            If lErro <> SUCESSO Then gError 51030
        Next
        
        Call gobjTribTab.TributacaoRecalcular_Click
    
        Call objContabil.Contabil_Carga_Modelo_Padrao
        
    End If
    
    lErro = CF2(Me, "NFEnt_Traz_NFiscal_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 124252
    
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
        
    NumNFe.Caption = objNFiscal.sStatusNFeFed
                
    Traz_NFiscal_Tela = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela:

    gbCarregandoTela = False
    
    Traz_NFiscal_Tela = gErr

    Select Case gErr

        Case 22842, 22875, 22880, 27833, 27834, 31425, 31431, 35132, 52376, 52377, 75910, 78835, 89499, 89500, 124252
        
        Case 51030, 131937 'Inserido por Wagner
        
        Case 35133
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntNotaOriginal)
        
        Case 39616
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156599)

    End Select

    Exit Function

End Function

Public Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO

    'Verifica se alguma filial foi selecionada
    If Filial.ListIndex = -1 Then Exit Sub
     
    Call Trata_FilialForn
    
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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 52410

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(FornecedorBenef.Text)) = 0 Then gError 52411

        sFornecedor = FornecedorBenef.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 52412

        If lErro = 18272 Then gError 52413

        'coloca na tela
        FilialFornBenef.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 52414

    Exit Sub

Erro_FilialFornBenef_Validate:

    Cancel = True

    Select Case gErr

        Case 52410, 52412
        
        Case 52411
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 52413
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, FornecedorBenef.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 52414
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, FilialFornBenef.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156600)

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
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 31446

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 31449

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 31447

        If lErro = 18272 Then gError 31448

        'coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

        Call gobjTribTab.FilialFornecedor_Alterada(objFilialFornecedor.lCodFornecedor, objFilialFornecedor.iCodFilial)
    
    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 35136
    
    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 31446, 31447

        Case 31448
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 31449
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 35136
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156601)

    End Select

    Exit Sub

End Sub

Public Function Preenche_GridItens(objNFiscal As ClassNFiscal) As Long
'Preenche o Grid com os ítens da Nota Fiscal

Dim lErro As Long
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
        If lErro <> SUCESSO Then gError 31430

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
            If lErro <> SUCESSO Then gError 52875
        
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

    Next
    
    lErro = CF2(Me, "NFE_Preenche_GridItens", objNFiscal)
    If lErro <> SUCESSO Then gError 181415

    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objNFiscal.objProjetoInfo, objNFiscal.lNumIntDoc, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181676

    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = iIndice

    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr
        
        Case 31430
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemNF.sProduto)
        
        Case 35683, 177926, 181415, 181676
        
        Case 35684
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.iCodigo)
        
        Case 52875
            Call Rotina_Erro(vbOKOnly, "Erro_Mascara_MascararCcl", gErr, objItemNF.sCcl)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156602)

    End Select

    Exit Function

End Function

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
    
    Set gobjNFiscal = New ClassNFiscal
    Set gobjInfoAdic = Nothing
    
    For iIndice = 1 To objGridItens.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next

    'Limpa o Grid
    Call Grid_Limpa(objGridItens)
    
'    If giTipoVersao = VERSAO_FULL Then
        Call Grid_Limpa(objGridBenef)
        Serie.Text = ""
        Call Serie_Validate(bSGECancelDummy)
'    End If

    'Limpa os campos que não são limpos pelo Limpa_Tela
    NFiscal.Text = ""
    NFiscalInterna.Caption = ""
    Status.Caption = ""
    FilialFornBenef.Clear
    Filial.Clear
    SerieNFiscalOriginal.Text = ""
    FilialFornNFOrig.Clear
    FilialCliNFOrig.Clear
    SubTotal.Caption = ""
    ValorProdutos2.Caption = ""
    Transportadora.Text = ""
    PlacaUF.Text = ""
    TipoFrete.ListIndex = 0

    VolumeMarca.Text = ""
    VolumeEspecie.Text = ""
    'EletronicaFed.Value = vbUnchecked
    
    If gobjEST.iSerieEletronicaPadrao = MARCADO Then
        EletronicaFed.Value = vbChecked
    Else
        EletronicaFed.Value = vbUnchecked
    End If
    
    'Recarrega a combo VolumeEspecie e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)

    'Recarrega a combo VolumeMarca e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
    
    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
    'distribuicao
    Call gobjDistribuicao.Limpa_Tela_Distribuicao
    
    Call gobjTribTab.DataEmissao_Alterada(DATA_NULA)
    
    'Preenche a Data de Entrada com a Data Atual
    DataEntrada.PromptInclude = False
    DataEntrada.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEntrada.PromptInclude = True
    Call DataEntrada_Validate(bSGECancelDummy)

''*** incluidos p/tratamento de tributacao *******************************

    Call gobjTribTab.Limpa_Tela

'*** fim tributacao
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    Call CF2(Me, "NFCF_Limpa_Tela")
    
    MsgAutomatica.Value = vbChecked
    
    iDescontoAlterado = 0
    dValorDescontoItensAnt = 0
    dPercDescontoItensAnt = 0

End Sub

Public Sub Limpa_Tela_NFiscal1()
'Limpa a Tela NFiscalEntrada, mas mantém a natureza e o tipo
    
Dim sNatureza As String
Dim sTipoNFiscal As String

    sNatureza = NaturezaOp.Text
    sTipoNFiscal = TipoNFiscal.Text
    
    Call Limpa_Tela_NFiscal

    NaturezaOp.Text = sNatureza
    TipoNFiscal.Text = sTipoNFiscal

End Sub

Public Sub DataEntrada_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEntrada, iAlterado)

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
        If lErro <> SUCESSO Then gError 31451

''*** incluidos p/tratamento de tributacao *******************************
        Call gobjTribTab.DataEntrada_Alterada(StrParaDate(DataEntrada.Text))
'*** fim tributacao
        
        iDataEntradaAlterada = 0

    End If

    Exit Sub

Erro_DataEntrada_Validate:

    Cancel = True


    Select Case gErr

        Case 31451

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156603)

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
    If lErro <> SUCESSO Then gError 89667

    Exit Sub

Erro_HoraEntrada_Validate:

    Cancel = True

    Select Case gErr

        Case 89667

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156604)

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
    If Len(Trim(DataEmissao.ClipText)) <> 0 Then

        'Critica a data digitada
        lErro = Data_Critica(DataEmissao.Text)
        If lErro <> SUCESSO Then gError 31452
    
        Call gobjTribTab.DataEmissao_Alterada(StrParaDate(DataEmissao.Text))
        
    End If

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case gErr

        Case 31452

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156605)

    End Select

    Exit Sub

End Sub

Public Sub DtVencimento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DtVencimento_Validate(Cancel As Boolean)

Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim lErro As Long

On Error GoTo Erro_DtVencimento_Validate

    'Verifica se a data de vencimento foi digitada
    If Len(Trim(dtVencimento.ClipText)) = 0 Then Exit Sub

    'Faz a critica da data digitada
    lErro = Data_Critica(dtVencimento.Text)
    If lErro <> SUCESSO Then gError 35097

    'Verifica se a data de emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    dtDataEmissao = CDate(DataEmissao.Text)
    dtDataVencimento = CDate(dtVencimento.Text)

    'Verifica se a data de vencimento é menor que a data de emissão
    If dtDataVencimento < dtDataEmissao Then gError 35098

    Exit Sub

Erro_DtVencimento_Validate:

    Cancel = True


    Select Case gErr

        Case 35097

        Case 35098
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_MENOR", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156606)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntrada_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntrada_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEntrada, DIMINUI_DATA)
    If lErro Then gError 31453

    Exit Sub

Erro_UpDownEntrada_DownClick:

    Select Case gErr

        Case 31453

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156607)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntrada_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEntrada_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEntrada, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 31454

    Exit Sub

Erro_UpDownEntrada_UpClick:

    Select Case gErr

        Case 31454

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156608)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 31456

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 31456

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156609)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 31455

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 31455

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156610)

    End Select

    Exit Sub

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatadoBenef As String
Dim iProdutoPreenchidoBenef As Integer
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String
Dim sUnidadeMed As String
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
            If lErro <> SUCESSO Then gError 31473
           
            If iProdutoPreenchido = PRODUTO_VAZIO Then
                UnidadeMed.Enabled = False
            Else
                UnidadeMed.Enabled = True

                objProduto.sCodigo = sProdutoFormatado
                
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 31471
                
                'Não achou o Produto
                If lErro = 28030 Then gError 31476

                objClasseUM.iClasse = objProduto.iClasseUM
                
                'Lâ as Unidades de Medidas da Classe do produto
                lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                If lErro <> SUCESSO Then gError 31472
                
                'Carrega a combo de UM
                For Each objUM In colSiglas
                    UnidadeMed.AddItem objUM.sSigla
                Next
                
                'Seleciona na UM que está preenchida
                UnidadeMed.Text = sUM
                
                If Len(Trim(sUM)) > 0 Then
                    lErro = Combo_Item_Igual(UnidadeMed)
                    If lErro <> SUCESSO And lErro <> 12253 Then gError 35104
                End If
            
            End If
            
        'Nas demais
        
'distribuicao
    'mario consulta produto
        Case ValorUnitario.Name, PercentDesc.Name, Desconto.Name, Ccl.Name
    
'        Case ValorUnitario.Name, PercentDesc.Name, Desconto.Name, Quantidade.Name, DescricaoItem.Name, Ccl.Name
'        Case ValorUnitario.Name, PercentDesc.Name, Desconto.Name, Quantidade.Name, Almoxarifado.Name, DescricaoItem.Name, Ccl.Name
        
            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 31474

            If iProdutoPreenchido = PRODUTO_VAZIO Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case "Quantidade"
    
            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 31474
            
            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Or left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case Else
            'Verifica se o produto está preenchido
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 31474

            lErro = CF2(Me, "NFE_Rotina_Grid_Enable", iLinha, objControl, iLocalChamada, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 124178

    End Select
    
'    If giTipoVersao = VERSAO_FULL Then
        Select Case objControl.Name
        
            Case ProdutoBenef.Name
                 
                lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iLinha, iGrid_ProdutoBenef_Col), sProdutoFormatadoBenef, iProdutoPreenchidoBenef)
                If lErro <> SUCESSO Then gError 52365
                
                If iProdutoPreenchidoBenef = PRODUTO_PREENCHIDO Then
                    objControl.Enabled = False
                Else
                    objControl.Enabled = True
                End If
    
            Case UnidadeMedBenef.Name
                
                lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iLinha, iGrid_ProdutoBenef_Col), sProdutoFormatadoBenef, iProdutoPreenchidoBenef)
                If lErro <> SUCESSO Then gError 52366
               
                If iProdutoPreenchidoBenef <> PRODUTO_PREENCHIDO Then
                    objControl.Enabled = False
                Else
                    objControl.Enabled = True
    
                    objProduto.sCodigo = sProdutoFormatadoBenef
    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 34777
    
                    If lErro = 28030 Then gError 34778
    
                    objClasseUM.iClasse = objProduto.iClasseUM
    
                    'Preenche a List da Combo UnidadeMed com as UM's do Produto
                    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
                    If lErro <> SUCESSO Then gError 34779
    
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
                If lErro <> SUCESSO Then gError 52367
    
                If iProdutoPreenchidoBenef = PRODUTO_PREENCHIDO Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
                
        End Select
        
        'Rastreamento
        lErro = gobjRastreamento.Rotina_Grid_Enable(iLinha, objControl, iLocalChamada)
        If lErro <> SUCESSO Then gError 83226
        
        'distribuicao
        lErro = gobjDistribuicao.Rotina_Grid_Enable_Dist(iLinha, objControl, iLocalChamada)
        If lErro <> SUCESSO Then gError 89501
        
 '   End If
    
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 31471, 31472, 31473, 31474, 34779, 35104, 35658, 52365, 52366, 52367, 83226, 89501

        Case 31476
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156611)

    End Select

    Exit Sub

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

Public Sub ValorDespesas_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorDespesasAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorFrete_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorFreteAlterado = REGISTRO_ALTERADO

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
        If lErro <> SUCESSO Then gError 35013

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

        Case 35013

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156612)

    End Select

    Exit Sub

End Sub

Public Sub ValorSeguro_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorSeguroAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorUnitario_Change()
    iAlterado = REGISTRO_ALTERADO
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

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
        
        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 39617
        
        'Verifica se é o GridItens
        If objGridInt.objGrid.Name = GridItens.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                'Produto
                Case iGrid_Produto_Col
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 35140
                
                'DescricaoItem
                Case iGrid_DescProduto_Col
                    lErro = Saida_Celula_DescricaoItem(objGridInt)
                    If lErro <> SUCESSO Then gError 43297
                
                'Unidade Medida
                Case iGrid_UnidadeMed_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError 35084
                
                'Quantidade
                Case iGrid_Quantidade_Col
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 35656
 
                'ValorUnitario
                Case iGrid_ValorUnitario_Col
                    lErro = Saida_Celula_ValorUnitario(objGridInt)
                    If lErro <> SUCESSO Then gError 35000
                
                'PercDesc
                Case iGrid_PercDesc_Col
                    lErro = Saida_Celula_PercDesc(objGridInt)
                    If lErro <> SUCESSO Then gError 35003
                
                'Desconto
                Case iGrid_Desconto_Col
                    lErro = Saida_Celula_Desconto(objGridInt)
                    If lErro <> SUCESSO Then gError 35007
                
                'CCL
                Case iGrid_Ccl_Col
                
                    lErro = Saida_Celula_Ccl(objGridInt)
                    If lErro <> SUCESSO Then gError 52850
                    
                Case Else
                    lErro = CF2(Me, "NFE_Saida_Celula_GridItens", objGridInt)
                    If lErro <> SUCESSO Then gError 181417
                    
            End Select
        
        ElseIf objGridInt.objGrid.Name = GridMovimentos.Name Then

            Select Case GridMovimentos.Col

                Case iGrid_ProdutoBenef_Col

                    lErro = Saida_Celula_ProdutoBenef(objGridInt)
                    If lErro <> SUCESSO Then gError 52368

                Case iGrid_QuantidadeBenef_Col

                    lErro = Saida_Celula_QuantidadeBenef(objGridInt)
                    If lErro <> SUCESSO Then gError 52369

                Case iGrid_AlmoxarifadoBenef_Col

                    lErro = Saida_Celula_AlmoxarifadoBenef(objGridInt)
                    If lErro <> SUCESSO Then gError 52370

                Case iGrid_UnidadeMedBenef_Col

                    lErro = Saida_Celula_UnidadeMedBenef(objGridInt)
                    If lErro <> SUCESSO Then gError 52371

                Case iGrid_ContaContabilEstBenef_Col
                    lErro = Saida_Celula_ContaContabilEst(objGridInt)
                    If lErro <> SUCESSO Then gError 52372

                Case iGrid_ContaContabilProducaoBenef_Col
                    lErro = Saida_Celula_ContaContabilProducao(objGridInt)
                    If lErro <> SUCESSO Then gError 52373

            End Select
        
        'Rastreamento
        ElseIf objGridInt.objGrid.Name = GridRastro.Name Then

            lErro = gobjRastreamento.Saida_Celula()
            If lErro <> SUCESSO Then gError 83196

        'distribuicao
        ElseIf objGridInt.objGrid.Name = GridDist.Name Then

            lErro = gobjDistribuicao.Saida_Celula_Dist()
            If lErro <> SUCESSO Then gError 89492

        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 31477

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 31477
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 31495, 35000, 35003, 35007, 35084, 35140, 35656, 39617, 43297, 181417
            
        Case 52368, 52369, 52370, 52371, 52372, 52373, 52850, 83196, 89492

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156613)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl

    If Len(Trim(Ccl.ClipText)) <> 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 52847

        If lErro = 5703 Then gError 52848

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 52849

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = gErr

    Select Case gErr

        Case 52847, 52849
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52848
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, Ccl.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156614)

    End Select

    Exit Function

End Function

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
        If lErro <> SUCESSO Then gError 199352

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

        Case 31479, 31481, 35657, 199352
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 31480
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 35659, 55612
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 55139
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156615)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescricaoItem(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_DescricaoItem

    Set objGridInt.objControle = DescricaoItem

'mario inicio

    If Len(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)) = 0 Then

        Select Case left(LTrim(DescricaoItem.Text), 1)
        
            Case "*"
            
                objProduto.sNomeReduzido = Mid(LTrim(DescricaoItem.Text), 2)
            
                'Lê o Produto com a Descricao Passada
                lErro = CF("Produto_Le_NomeReduzido", objProduto)
                If lErro <> SUCESSO And lErro <> 26927 Then gError 89869
            
                'se o produto não estiver cadastrado ==> erro
                If lErro = 26927 Then gError 89870
            
        End Select
    
        If Len(objProduto.sCodigo) > 0 Then
            
            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
            If lErro <> SUCESSO Then gError 89873
            
            Produto.PromptInclude = False
            Produto.Text = sProduto
            Produto.PromptInclude = True
        
            GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text
        
            'Coloca as demais características do produto na tela
            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then gError 89874
            
            DescricaoItem.Text = Mid(LTrim(DescricaoItem.Text), 2)
            
        End If
    
    End If
    
'mario fim

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 43298
    
    Saida_Celula_DescricaoItem = SUCESSO

    Exit Function

Erro_Saida_Celula_DescricaoItem:

    Saida_Celula_DescricaoItem = gErr

    Select Case gErr

        Case 43298, 89869
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

'mario inicio
        
        Case 89870
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO_NOMEREDUZIDO", DescricaoItem.Text)

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
            
        Case 89871, 89872
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 89873
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 89874
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
'mario fim

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156616)

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

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 180063

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

    lErro = CF2(Me, "NFE_ProdutoLinha_Preenche", objProduto)
    If lErro <> SUCESSO Then gError 181258

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridItens.Row - GridItens.FixedRows) = objGridItens.iLinhasExistentes Then
        objGridItens.iLinhasExistentes = objGridItens.iLinhasExistentes + 1
    
        Call gobjTelaProjetoInfo.Adiciona_Linha(GridItens.Row)
    
''*** incluidos p/tratamento de tributacao *******************************
        'permite que a tributacao reflita a inclusao de uma linha no grid
        lErro = gobjTribTab.Inclusao_Item_Grid(GridItens.Row, objProduto.sCodigo)
        If lErro <> SUCESSO Then gError 55653
'*** fim tributacao

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
    
    lErro = Preco_Traz_Padrao(GridItens.Row, objProduto.sCodigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Calcula_Valores(GridItens.Row)
    
    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156617)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim iIndice As Integer
Dim iIndice2 As Integer
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
        If lErro <> SUCESSO Then gError 31483

        dQuantidade = CDbl(Quantidade.Text)

        'Coloca o valor Formatado na tela
        Quantidade.Text = Formata_Estoque(dQuantidade)
    
        'distribuicao
        dQuantidadeAtual = StrParaDbl(Quantidade.Text)
        'fim  distribuicao
    
    End If
            
    lErro = CF2(Me, "NFEnt_AlocaAlmoxarifadoPadrao", gobjDistribuicao)
    If lErro <> SUCESSO Then gError 141072
    
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
            If lErro <> SUCESSO Then gError 89512
            
        Else
        
            objKitVenda.sProduto = gobjNFiscal.ColItensNF(GridItens.Row).sProduto
            objKitVenda.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
            objKitVenda.dQuantidade = dQuantidade
        
            lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade)
            If lErro <> SUCESSO Then gError 180041

            lErro = gobjDistribuicao.Distribuicao_Processa_Grade()
            If lErro <> SUCESSO Then gError 89512
        
            Call gobjDistribuicao.Atualiza_Grid_Distribuicao(gobjNFiscal.ColItensNF.Item(GridItens.Row))
        End If
    End If
    'fim distribuicao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31484
    
    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 55557

    '*** incluidos p/tratamento de tributacao ****************
    If GridItens.Row <= objGridItens.iLinhasExistentes Then
        lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
        If lErro <> SUCESSO Then gError 22843
    End If
    '*** fim tributacao
    
    Call ValorTotal_Calcula
    
    If Not gbCarregandoTela Then
        lErro = gobjTribTab.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 22843
    End If
    
    If Abs(dQuantidadeAnterior - dQuantidadeAtual) > QTDE_ESTOQUE_DELTA Then
        lErro = CF2(Me, "NFE_Calcula_PesoKg", GridItens.Row)
        If lErro <> SUCESSO Then gError 181262
    End If
    
    If GridItens.Row <= objGridItens.iLinhasExistentes Then
    
        If gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade.Count = 0 Then
    
            lErro = gobjRastreamento.Atualiza_Troca_Quantidade(GridItens.Row)
            If lErro <> SUCESSO Then gError 83205
        
        Else
        
            If dQuantidadeAnterior <> 0 Then
                dFator = dQuantidadeAtual \ dQuantidadeAnterior
            Else
                dFator = 1
            End If
            
            lErro = gobjRastreamento.Atualiza_Troca_Quantidade_Grade(GridItens.Row, dFator)
            If lErro <> SUCESSO Then gError 83205
        
        End If
    
    End If
    
    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 31483, 31484, 22843, 55557, 83205, 89512, 141072, 180040, 180041, 181262 'Inserido por Wagner
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156618)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156619)

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
Dim dValorUnitario As Double
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
    If lErro <> SUCESSO Then gError 55554

    Calcula_Valores = SUCESSO
    
    Exit Function
    
Erro_Calcula_Valores:

    Calcula_Valores = gErr
    
    Select Case gErr

        Case 55554, 181259

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156621)

    End Select

    Exit Function
    
End Function

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
        If lErro <> SUCESSO And lErro <> 6681 Then gError 89311

        'Se não achou o Fornecedor --> erro
        If lErro = 6681 Then gError 89312

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
    
        Case 89311

        Case 89312
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, Fornecedor.Text)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156622)
    
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

Private Function Saida_Celula_ValorUnitario(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorUnitario As Double
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_ValorUnitario

    Set objGridInt.objControle = ValorUnitario

    'Se estiver preenchido
    If Len(Trim(ValorUnitario.ClipText)) > 0 Then
    
        iCodigo = Codigo_Extrai(TipoNFiscal.Text)
        
        If giComplemento = MARCADO Then
            lErro = Valor_NaoNegativo_Critica(ValorUnitario.Text)
        Else
            lErro = Valor_Positivo_Critica(ValorUnitario.Text)
        End If
        If lErro <> SUCESSO Then gError 35001
        dValorUnitario = CDbl(ValorUnitario.Text)
            
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 35002
    
    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 55556

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

        Case 22844, 35001, 35002, 55556, 87191
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156624)

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
        If lErro <> SUCESSO Then gError 52262

'        dPercentDesc = CDbl(PercentDesc.Text)
'
'        lTamanho = Len(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col))
'        If lTamanho > 0 Then dPercentDescAnterior = PercentParaDbl(left(GridItens.TextMatrix(GridItens.Row, iGrid_PercDesc_Col), lTamanho - 1))

        If dPercentDesc <> dPercentDescAnterior Then
        
            'Verifica se o percentual é de 100%
            If dPercentDesc = 100 Then gError 52263
    
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
     If lErro <> SUCESSO Then gError 52264

    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 55558

''*** incluidos p/tratamento de tributacao *******************************
    lErro = gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
    If lErro <> SUCESSO Then gError 52265
'*** fim tributacao

    Call ValorTotal_Calcula
    
    Saida_Celula_PercDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercDesc:

    Saida_Celula_PercDesc = gErr

    Select Case gErr

        Case 52262, 52264, 52265, 55558
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52263
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156625)

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

    'Se estiver preenchido
    If Len(Trim(Desconto.ClipText)) > 0 Then
    
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(Desconto.Text)
        If lErro <> SUCESSO Then gError 35008

        dDesconto = CDbl(Desconto.Text)

        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))) > 0 Then dQuantidade = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col))
        
        lErro = CF2(Me, "NFE_PrecoTotal_Calcula", GridItens.Row, dQuantidade)
        If lErro <> SUCESSO Then gError 181270
        
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))) > 0 Then dValorUnitario = CDbl(GridItens.TextMatrix(GridItens.Row, iGrid_ValorUnitario_Col))
        
        dValorTotal = dValorUnitario * dQuantidade

        If dValorTotal > 0 Then
        
            dPercentDesc = dDesconto / (dValorTotal)
           
            If dPercentDesc >= 1 Then gError 35789
            
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
    If lErro <> SUCESSO Then gError 35009
   
    'recalcula os valores de desconto, percentual de desconto e valor total
    lErro = Calcula_Valores(GridItens.Row)
    If lErro <> SUCESSO Then gError 55559

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

        Case 35008, 35009, 22846, 55559, 181270
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 35789
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156626)

    End Select

    Exit Function

End Function

Public Sub ValorSeguro_Validate(Cancel As Boolean)

Dim dValorSeguro As Double
Dim lErro As Long

On Error GoTo Erro_Valorseguro_Validate

    If iValorSeguroAlterado = 0 Then Exit Sub

    'Se estiver preenchido
    If Len(Trim(ValorSeguro.Text)) > 0 Then
        'Critica o valor
        lErro = Valor_NaoNegativo_Critica(ValorSeguro.Text)
        If lErro <> SUCESSO Then gError 35014

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

        Case 35014

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156627)

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
        If lErro <> SUCESSO Then gError 35015

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


    Select Case gErr

        Case 35015

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156628)

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
Dim iPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objItemNF As ClassItemNF
Dim dQuantLote As Double

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 35021
    If Len(Trim(NaturezaOp.Text)) = 0 Then gError 35022
    If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 35023
    If Len(Trim(Filial.Text)) = 0 Then gError 35024
    If Len(Trim(Serie.Text)) = 0 Then gError 35025
            
    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 35031
    
    If lErro = 31415 Then gError 35032 'Não encontrou

    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS Then
        'If Len(Trim(NFiscalInterna.Caption)) > 0 Then gError 60450
        If Len(Trim(NFiscalInterna.Caption)) > 0 Then
            objNFiscal.lNumNotaFiscal = CLng(NFiscalInterna.Caption)
        Else
            objNFiscal.lNumNotaFiscal = 0
        End If

    Else
        If Len(Trim(NFiscal.ClipText)) = 0 Then gError 35026
    End If
        
    objNFiscal.iTipoDocInfo = objTipoDocInfo.iTipo
        
    If Len(Trim(DataEntrada.ClipText)) = 0 Then gError 35027
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 35028
    
    'se o tipo de Nfiscal for do tipo "Externa - Mat. A ser beneficiado" entao obtem o Fornecedor a ser beneficiado
    If TipoNFiscal.ListIndex <> -1 Then
        
        If TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFEEBF Or TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFIEBF Or TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFECCO Then
            
            If Len(Trim(FornecedorBenef.Text)) = 0 Then gError 52359
            If Len(Trim(FilialFornBenef.List(FilialFornBenef.ListIndex))) = 0 Then gError 52360
        
        End If
    End If
    
    'Verifica se a Data de Entrada e menor que a Data de Emissão
    If CDate(DataEntrada.Text) < CDate(DataEmissao.Text) Then gError 35029
    
    'Verifica de a Data de Vencimento e Menor que a de Emissão
    If Len(Trim(dtVencimento.ClipText)) <> 0 Then
        If CDate(dtVencimento.Text) < CDate(DataEmissao.Text) Then gError 35030
    End If

    If objTipoDocInfo.iComplementar = DOCINFO_NORMAL Then
        If objGridItens.iLinhasExistentes = 0 Then gError 35035
    End If

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 35040

    'Valida os dados do Grid
    If objGridItens.iLinhasExistentes > 0 Then
    
        For iIndice = 1 To objGridItens.iLinhasExistentes

            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 35036
            
            Set objItemNF = objNFiscal.ColItensNF(iIndice)
        
            objProduto.sCodigo = objItemNF.sProduto
        
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 35661
            
            If lErro = 28030 Then gError 35662
        
            objItemNF.iClasseUM = objProduto.iClasseUM
            objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
            objItemNF.sUMVenda = objProduto.sSiglaUMVenda
            
            'Verifica se a Unidade de Medida foi preenchida
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col))) = 0 Then gError 55628
            
            If objTipoDocInfo.iCodigo <> DOCINFO_NFEEICM And objTipoDocInfo.iCodigo <> DOCINFO_NFIEICM And _
                objTipoDocInfo.iCodigo <> DOCINFO_NFEEIPI And objTipoDocInfo.iCodigo <> DOCINFO_NFIEIPI Then
                If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) = 0 Then gError 35038
            
            ElseIf objTipoDocInfo.iCodigo = DOCINFO_NFEEICM Then
                
                If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ValorUnitario_Col))) = 0 Then
                    vbMsgRes = Rotina_Aviso(vbYesNo, "CONFIRMA_VUNITARIO_GRAVACAO")

                    If vbMsgRes = vbNo Then gError 87192
                End If
                
            End If
            
            lErro = CF2(Me, "NFE_Valida_Grid_Itens", iIndice)
            If lErro <> SUCESSO Then gError 181418
           
        Next
        
    End If
    
    dTotal = CDbl(IIf(Len(Trim(Total.Text)) > 0, Total.Text, 0))

    'Se o total for negativo --> Erro
    If dTotal < 0 Then gError 35039

    If Len(Trim(PesoLiquido.ClipText)) > 0 And Len(Trim(PesoBruto.ClipText)) > 0 Then
        If CDbl(PesoLiquido.Text) > CDbl(PesoBruto.Text) Then gError 35134
    End If
    
    lErro = Total_Calcula()
    If lErro <> SUCESSO Then gError 43802

    'se o tab de beneficiamento existe
    'Rastreamento
    If TabStrip1.Tabs.Count = TAB_BENEFICIAMENTO Then
    
        'Valida os dados do grid de beneficiamento
        For iIndice = 1 To objGridBenef.iLinhasExistentes
            
            'Verifica se a Quantidade foi informada
            If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_QuantidadeBenef_Col))) = 0 Then gError 55632

            'Verifica se a Unidade de Medida foi preenchida
            If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMedBenef_Col))) = 0 Then gError 55633
            
            'Verifica se o Almoxarifado foi informado
            If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxarifadoBenef_Col))) = 0 Then gError 55634

        Next
        
    End If

    lErro = gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
    If lErro <> SUCESSO Then gError 83220

    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786

    lErro = gobjTribTab.Valida_Dados
    If lErro <> SUCESSO Then gError 181786

    Set objNFiscal.objRastreamento = gobjRastreamento

    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Or objTipoDocInfo.iCodigo = DOCINFO_DDAI Then
    
        'Recolhe os dados da Nota Fiscal Original
        lErro = Move_Dados_NFOriginal_Memoria(objNFiscal, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 35091
        
    ElseIf objTipoDocInfo.iCodigo = DOCINFO_CFEV Then
    
        'Recolhe os dados da Nota Fiscal Original
        lErro = Move_Dados_NFOriginal_Cli(objNFiscal, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 92082
    
    ElseIf objTipoDocInfo.iCodigo = DOCINFO_CFECT Or objTipoDocInfo.iCodigo = DOCINFO_NFSET Then
        
        'Recolhe os dados da Nota Fiscal Original
        lErro = Move_Dados_NFOriginal_Forn(objNFiscal, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 92083
        
    End If

    lErro = gobjTribTab.Valida_TributacaoNF_Gravacao(objNFiscal)
    If lErro <> SUCESSO Then gError 59003
    
    lErro = Calcula_Msg
    If lErro <> SUCESSO Then gError 59003
    
    'mover a parte do rastreamento
    lErro = gobjRastreamento.Move_Rastro_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 83221
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataEntrada.Text))
    If lErro <> SUCESSO Then gError 92033
    
    lErro = CF("NFiscal_Valida_Diversos", objNFiscal)
    If lErro <> SUCESSO Then gError 92033
    
    'Grava a Nota Fiscal de Entrada(inclusive os dados contábeis)
    lErro = CF("NFiscalEntrada_Grava", objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 35041
    
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
        If lErro <> SUCESSO Then gError 203002
    End If
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 35021
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 35022
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 35023
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 35024
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 35025
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 35026
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)

        Case 35027
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_NAO_PREENCHIDA", gErr)

        Case 35028
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)

        Case 35029
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_ANTERIOR_DATAEMISSAO", gErr, DataEntrada, DataEmissao)

        Case 35030
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR", gErr, DataEmissao.Text, dtVencimento.Text)

        Case 35031, 35040, 35041, 35091, 43802, 59003, 83220, 83221, 87192, 92033, 92082, 92083, 130406, 181418, 181786, 203002

        Case 35032
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iTipo)

        Case 35035
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)

        Case 35036
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ITEM_NAO_PREENCHIDA", gErr, iIndice)

        Case 35038
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORUNITARIO_ITEM_NAO_PREENCHIDO", gErr, iIndice)

        Case 35039
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)

        Case 35134
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr, CDbl(PesoLiquido.Text), CDbl(PesoBruto.Text))
              
        Case 35660, 35661
        
        Case 35662
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 52359
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_BENEF_NAO_PREENCHIDA", gErr)
        
        Case 52360
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORNECEDOR_BENEF_NAO_PREENCHIDA", gErr)
        
        Case 55552
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_COMPRAVEL", gErr, GridItens.TextMatrix(iIndice, iGrid_Produto_Col))
            
        Case 55628
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 55632
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_BENEF_NAO_PREENCHIDA", gErr, iIndice)
            
        Case 55633
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_BENEF_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 55634
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_BENEF_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 75906
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTTOTAL_LOTE_MAIOR_ALMOXARIFADO", gErr, dQuantLote, StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)))
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156629)

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
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 89069
    If Len(Trim(Serie.Text)) = 0 Then gError 89070
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 89071

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
'            If lErro <> SUCESSO Then gError 203040
'
'            If iEnviado = 1 Then gError 203041
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
        If lErro <> SUCESSO Then gError 89074

        'Faz a exclusão da nota fiscal
        lErro = CF("NotaFiscalEntrada_Excluir", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError 89075
    
        'Limpa a Tela
        Call Limpa_Tela_NFiscal1
    
        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 89069
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 89070
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 89071
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 89072, 89074, 89075, 99345, 99351, 99352, 99353, 203040
       
        Case 89065
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
            
        Case 99346
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)
        
        Case 99350
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
        
        Case 99352
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)

        Case 203041
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CADASTRADA_NAO_EXCLUI", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156630)
        
    End Select

    Exit Sub

End Sub

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
    
    lErro = CF2(Me, "NFEnt_AlocaAlmoxarifadoPadrao", gobjDistribuicao)
    If lErro <> SUCESSO Then gError 141073
    
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
    
    If sUMAnterior <> UnidadeMed.Text Then
        If Not gbCarregandoTela Then
            lErro = gobjTribTab.Pesos_Recalcula()
            If lErro <> SUCESSO Then gError 177057
        End If

        lErro = CF2(Me, "NFE_Calcula_PesoKg", GridItens.Row)
        If lErro <> SUCESSO Then gError 177161
    End If

    If GridItens.Row <= objGridItens.iLinhasExistentes Then
        If gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade.Count = 0 Then
            lErro = gobjRastreamento.Atualiza_Troca_UM(GridItens.Row)
            If lErro <> SUCESSO Then gError 83206
        Else
            lErro = gobjRastreamento.Atualiza_Troca_UM_Grade(GridItens.Row, sUMAnterior, UnidadeMed.Text, objProduto)
            If lErro <> SUCESSO Then gError 83206
        End If
    End If

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 35083, 83206, 89511, 89602, 141073, 177583, 177584, 177161
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156631)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Sub UpDownVencimento_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVencimento_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(dtVencimento, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 35085

    Exit Sub

Erro_UpDownVencimento_DownClick:

    Select Case gErr

        Case 35085

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156632)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencimento_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVencimento_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(dtVencimento, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 35086

    Exit Sub

Erro_UpDownVencimento_UpClick:

    Select Case gErr

        Case 35086

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156633)

    End Select

    Exit Sub

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
    If lErro <> SUCESSO Then gError 102450
    
    Exit Sub

Erro_VolumeEspecie_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102450
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156634)

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
    If lErro <> SUCESSO Then gError 102451
    
    Exit Sub

Erro_VolumeMarca_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102451
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156635)

    End Select

End Sub

Public Sub VolumeNumero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VolumeQuant_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Move_Dados_NFOriginal_Memoria(objNFiscal As ClassNFiscal, objTipoDocInfo As ClassTipoDocInfo) As Long
'Move os dados da NotaFiscalOriginal para a memória

Dim lErro As Long
Dim objNFOriginal As New ClassNFiscal
Dim objFornecedor As New ClassFornecedor
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Move_Dados_NFOriginal_Memoria

    If Len(Trim(NFiscalOriginal.ClipText)) <> 0 Or Len(Trim(SerieNFiscalOriginal.Text)) <> 0 Then
    
        'verifica se a Serie e o Número da NF original estão preenchidos
        If Len(Trim(NFiscalOriginal.ClipText)) = 0 Then gError 35090
        If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Then gError 35099
    
        'Se for uma nota de despesas c/despacho aduaneiro
        If Codigo_Extrai(Trim(TipoNFiscal.Text)) = DOCINFO_DDAI Then
        
            'Verifica se o código do Fornecedor e da Filial estão preenchidos
            If Len(Trim(FornNFOrig.Text)) = 0 Then gError 94880
            If Len(Trim(FilialFornNFOrig.Text)) = 0 Then gError 94881
            
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
            If lErro <> SUCESSO And lErro <> 6681 Then gError 94909
            
            'Se não encontrou o fornecedor => erro
            If lErro = 6681 Then gError 94910
            
            'Guarda no obj o código do fornecedor da NF original
            objNFOriginal.lFornecedor = objFornecedor.lCodigo
            
            'Guarda a filial do fornecedor da nota fiscal original
            objNFOriginal.iFilialForn = Codigo_Extrai(FilialFornNFOrig.Text)
            
            'O tipo da nota fiscal original só pode ser nota de importação
            objNFOriginal.iTipoNFiscal = DOCINFO_NFIEIMP
    
            'Lê a Nota Fiscal Original
            lErro = CF("NFiscal_Le_Num_Forn1", objNFOriginal)
            If lErro <> SUCESSO And lErro <> 55885 Then gError 94878
        
            If lErro <> SUCESSO Then
            
                'O tipo da nota fiscal original só pode ser nota de importação
                objNFOriginal.iTipoNFiscal = DOCINFO_NFIEIMPSE
        
                'Lê a Nota Fiscal Original
                lErro = CF("NFiscal_Le_Num_Forn1", objNFOriginal)
                If lErro <> SUCESSO And lErro <> 55885 Then gError 94878
                
'                'Se não encontrar --> Erro
'                If lErro = 55885 Then gError 35089
                    
                If lErro <> SUCESSO Then
                
                    'O tipo da nota fiscal original só pode ser nota de importação
                    objNFOriginal.iTipoNFiscal = DOCINFO_NFIEIMPPC
            
                    'Lê a Nota Fiscal Original
                    lErro = CF("NFiscal_Le_Num_Forn1", objNFOriginal)
                    If lErro <> SUCESSO And lErro <> 55885 Then gError 94878
                    
                    'Se não encontrar --> Erro
                    If lErro = 55885 Then gError 35089
                
                End If
                    
            End If
        
        Else
            
             objNFOriginal.lFornecedor = objNFiscal.lFornecedor
             objNFOriginal.iFilialForn = objNFiscal.iFilialForn
             objNFOriginal.sSerie = SerieNFiscalOriginal.Text
             objNFOriginal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
            
             'Lê a Nota Fiscal Original
             lErro = CF("NFiscal_Le_Num_Forn", objNFOriginal, objTipoDocInfo.iTipo)
             If lErro <> SUCESSO And lErro <> 31467 Then gError 35088
             
             If lErro = 31467 Then gError 35089 'Se não encontrar --> Erro
        
        End If
        
        'guarda em objNfiscal o Numero Interno da NF Original
        objNFiscal.lNumIntNotaOriginal = objNFOriginal.lNumIntDoc

    Else
    
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NFORIGINAL_NAO_PREENCHIDO")
        If vbMsgRes <> vbYes Then gError 184159
        
    End If
    
    Move_Dados_NFOriginal_Memoria = SUCESSO

    Exit Function

Erro_Move_Dados_NFOriginal_Memoria:

    Move_Dados_NFOriginal_Memoria = gErr

    Select Case gErr

        Case 35088, 35670, 94909, 184159

        Case 94880
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NFORIGINAL_NAO_PREENCHIDO", gErr)

        Case 94881
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NFORIGINAL_NAO_PREENCHIDO", gErr)
        
        Case 94910
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case 35089
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA1", gErr, objNFOriginal.sSerie, objNFOriginal.lNumNotaFiscal)

        Case 35099
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NFISCAL_ORIGINAL_NAO_PREENCHIDA", gErr)

        Case 35090
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NFISCAL_ORIGINAL_NAO_PREENCHIDO", gErr)

'        Case 35669
'            Call Rotina_Erro(vbOKOnly, "ERRO_TIPONFISCALORIGINAL_NAO_ENCONTRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156636)

    End Select

    Exit Function

End Function

Private Function Move_Dados_NFOriginal_Forn(objNFiscal As ClassNFiscal, objTipoDocInfo As ClassTipoDocInfo) As Long
'Move os dados da NotaFiscalOriginal para a memória

Dim lErro As Long
Dim objNFOriginal As New ClassNFiscal
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Dados_NFOriginal_Forn

    If Len(Trim(NFiscalOriginal.ClipText)) <> 0 Or Len(Trim(SerieNFiscalOriginal.Text)) <> 0 Or Len(Trim(FornNFOrig.Text)) <> 0 Or Len(Trim(FilialFornNFOrig.Text)) <> 0 Then
    
        'verifica se a Serie e o Número da NF original estão preenchidos
        If Len(Trim(NFiscalOriginal.ClipText)) = 0 Then gError 92066
        If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Then gError 92067
        If Len(Trim(FornNFOrig.Text)) = 0 Then gError 92068
        If Len(Trim(FilialFornNFOrig.Text)) = 0 Then gError 92069
        
        objNFOriginal.sSerie = SerieNFiscalOriginal.Text
        objNFOriginal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
        
        objFornecedor.sNomeReduzido = FornNFOrig.Text
    
        'Lê o fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 92070
    
        If lErro = 6681 Then gError 92071 'Se nao encontrar --> erro
    
        objNFOriginal.lFornecedor = objFornecedor.lCodigo
    
        If Len(Trim(FilialFornNFOrig.List(FilialFornNFOrig.ListIndex))) > 0 Then objNFOriginal.iFilialForn = Codigo_Extrai(FilialFornNFOrig.List(FilialFornNFOrig.ListIndex))
        
        'Lê a Nota Fiscal Original
        lErro = CF("NFiscal_Le_Num_Forn2", objNFOriginal)
        If lErro <> SUCESSO And lErro <> 92086 Then gError 92072
        
        If lErro <> SUCESSO Then gError 92073 'Se não encontrar --> Erro
        
        'guarda em objNfiscal o Numero Interno da NF Original
        objNFiscal.lNumIntNotaOriginal = objNFOriginal.lNumIntDoc

    End If
    
    Move_Dados_NFOriginal_Forn = SUCESSO

    Exit Function

Erro_Move_Dados_NFOriginal_Forn:

    Move_Dados_NFOriginal_Forn = gErr

    Select Case gErr

        Case 92066
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NFISCAL_ORIGINAL_NAO_PREENCHIDO", gErr)

        Case 92067
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NFISCAL_ORIGINAL_NAO_PREENCHIDA", gErr)

        Case 92068
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NFORIGINAL_NAO_PREENCHIDO", gErr)

        Case 92069
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NFORIGINAL_NAO_PREENCHIDO", gErr)

        Case 92070, 92072
        
        Case 92071
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, FornNFOrig.Text)

        Case 92073
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA1", gErr, objNFOriginal.sSerie, objNFOriginal.lNumNotaFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156637)

    End Select

    Exit Function

End Function

Private Function Move_Dados_NFOriginal_Cli(objNFiscal As ClassNFiscal, objTipoDocInfo As ClassTipoDocInfo) As Long
'Move os dados da NotaFiscalOriginal para a memória

Dim lErro As Long
Dim objNFOriginal As New ClassNFiscal
Dim objCliente As New ClassCliente

On Error GoTo Erro_Move_Dados_NFOriginal_Cli

    If Len(Trim(NFiscalOriginal.ClipText)) <> 0 Or Len(Trim(SerieNFiscalOriginal.Text)) <> 0 Or Len(Trim(ClienteNFOrig.Text)) <> 0 Or Len(Trim(FilialCliNFOrig.Text)) <> 0 Then
    
        'verifica se a Serie e o Número da NF original estão preenchidos
        If Len(Trim(NFiscalOriginal.ClipText)) = 0 Then gError 92074
        If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Then gError 92075
        If Len(Trim(ClienteNFOrig.Text)) = 0 Then gError 92076
        If Len(Trim(FilialCliNFOrig.Text)) = 0 Then gError 92077
        
        objNFOriginal.sSerie = SerieNFiscalOriginal.Text
        objNFOriginal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
        
        objCliente.sNomeReduzido = ClienteNFOrig.Text
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 92078
        
        'Não encontrou o Cliente --> erro
        If lErro = 12348 Then gError 92079
    
        objNFOriginal.lCliente = objCliente.lCodigo
        
        If Len(Trim(FilialCliNFOrig.List(FilialCliNFOrig.ListIndex))) > 0 Then objNFOriginal.iFilialCli = Codigo_Extrai(FilialCliNFOrig.List(FilialCliNFOrig.ListIndex))
        
        'Lê a Nota Fiscal Original
        lErro = CF("NFiscal_Le_Num_Cli2", objNFOriginal)
        If lErro <> SUCESSO And lErro <> 92090 Then gError 92080
            
        If lErro = 92090 Then gError 92081 'Se não encontrar --> Erro
        
        'guarda em objNfiscal o Numero Interno da NF Original
        objNFiscal.lNumIntNotaOriginal = objNFOriginal.lNumIntDoc

    End If
    
    Move_Dados_NFOriginal_Cli = SUCESSO

    Exit Function

Erro_Move_Dados_NFOriginal_Cli:

    Move_Dados_NFOriginal_Cli = gErr

    Select Case gErr

        Case 92074
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NFISCAL_ORIGINAL_NAO_PREENCHIDO", gErr)

        Case 92075
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NFISCAL_ORIGINAL_NAO_PREENCHIDA", gErr)

        Case 92076
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NFORIGINAL_NAO_PREENCHIDO", gErr)

        Case 92077
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NFORIGINAL_NAO_PREENCHIDO", gErr)

        Case 92078, 92080
        
        Case 92079
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, ClienteNFOrig.Text)

        Case 92081
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA1", gErr, objNFOriginal.sSerie, objNFOriginal.lNumNotaFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156638)

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
Dim objFilial As New ClassFilialFornecedor
Dim sContaTela As String
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim sContaMascarada As String
Dim objItem As New ClassTribItemNFAux
Dim sProduto As String
Dim sUM As String
Dim dQuantidade As Double
Dim dQuantidadeConvertida As Double
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim iLinha1 As Integer
Dim iAchou As Integer
Dim sCategoriaProduto As String
Dim objProdutoCategoria As New ClassProdutoCategoria
Dim objCatProdItem As ClassCategoriaProdutoItem
Dim objPlanoConta As New ClassPlanoConta
Dim sCclCtaResultNFOrig As String, objNF As ClassNFiscal
Dim objTipoDocInfo As ClassTipoDocInfo

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
                        gError 79017
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
                    objMnemonicoValor.colValor.Add 0
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
            If Len(ICMSSubstValor1.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ICMSSubstValor1.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
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
                If lErro <> SUCESSO Then gError 57992
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 57993
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case CCL_LCTO_NFORIG
        
            sCclCtaResultNFOrig = ""
            Set objNF = New ClassNFiscal
            Set objTipoDocInfo = New ClassTipoDocInfo
            
            objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
            
            'Lê o Tipo da Nota Fiscal
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
        
            If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Or objTipoDocInfo.iCodigo = DOCINFO_DDAI Then
            
                'Recolhe os dados da Nota Fiscal Original
                lErro = Move_Dados_NFOriginal_Memoria(objNF, objTipoDocInfo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            ElseIf objTipoDocInfo.iCodigo = DOCINFO_CFEV Then
            
                'Recolhe os dados da Nota Fiscal Original
                lErro = Move_Dados_NFOriginal_Cli(objNF, objTipoDocInfo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            ElseIf objTipoDocInfo.iCodigo = DOCINFO_CFECT Or objTipoDocInfo.iCodigo = DOCINFO_NFSET Then
                
                'Recolhe os dados da Nota Fiscal Original
                lErro = Move_Dados_NFOriginal_Forn(objNF, objTipoDocInfo)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            End If
            
            If objNF.lNumIntNotaOriginal <> 0 Then
            
                lErro = CF("Lcto_Le_Ccl_Cta_Result_Doc", TRANSACAOCTBORIGEM_NFISCAL, objNF.lNumIntNotaOriginal, sCclCtaResultNFOrig)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            End If
            
            objMnemonicoValor.colValor.Add sCclCtaResultNFOrig
            
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
                    If lErro <> SUCESSO Then gError 183416
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183417
        
                    objMnemonicoValor.colValor.Add objProduto.iControleEstoque
                    
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        Case CONTA_DESP_EST_FORN
            
            If Len(Trim(Filial.Text)) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 183418
                
                If objFilial.sContaDespesa <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaDespesa, sContaTela)
                    If lErro <> SUCESSO Then gError 183419
                
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
                    If lErro <> SUCESSO Then gError 183526
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183527
        
                    If objProduto.sContaContabil <> "" Then
                    
                        lErro = Mascara_RetornaContaTela(objProduto.sContaContabil, sContaTela)
                        If lErro <> SUCESSO Then gError 183808
                    
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
                    If lErro <> SUCESSO Then gError 183549
                
                    objProdutoCategoria.sCategoria = objMnemonicoValor.vParam(1)
                    
                    objProdutoCategoria.sProduto = sProdutoFormatado
                    
                    Set objCatProdItem = New ClassCategoriaProdutoItem
                    
                    lErro = CF("ProdutoCategoria_Le_Item", objProdutoCategoria, objCatProdItem)
                    If lErro <> SUCESSO And lErro <> 183548 Then gError 183550
                    
                    If lErro = SUCESSO Then
                        
                        Set objPlanoConta = New ClassPlanoConta
                    
                        'le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta
                        lErro = CF("PlanoConta_Le_ContaSimples", objCatProdItem.dvalor1, objPlanoConta)
                        If lErro <> SUCESSO And lErro <> 5451 Then gError 183551
                    
                        If lErro = SUCESSO Then
                    
                            If objPlanoConta.sConta <> "" Then
                            
                                lErro = Mascara_RetornaContaTela(objPlanoConta.sConta, sContaTela)
                                If lErro <> SUCESSO Then gError 183809
                            
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
            If lErro <> SUCESSO Then gError 39649

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 79017
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_BENEF_NAO_PREENCHIDA", gErr, iLinha)
            
        Case 39649
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case 57818, 57846, 57859, 57968, 57992, 57993, 79016, 183062 To 183066, 183139, 183140, 183416 To 183419, 183526, 183527, 183549, 183550, 183551, 183808, 183809
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156639)

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

On Error GoTo Erro_Move_GridMovimentos_Memoria
    
    Set objNFiscal.objMovEstoque = New ClassMovEstoque
    
    'Para cada linha existente do Grid
    For iIndice = 1 To objGridBenef.iLinhasExistentes
 
        'Formata o Produto para BD
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_ProdutoBenef_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 52385

        'Pega o Código do Almoxarifado
        Set objAlmoxarifado = New ClassAlmoxarifado

        objAlmoxarifado.sNomeReduzido = GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxarifadoBenef_Col)

        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then gError 52386

        If lErro = 25056 Then gError 52387
        
        If GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilEstBenef_Col) <> "" Then
        
            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilEstBenef_Col), sContaFormatadaEst, iContaPreenchida)
            If lErro <> SUCESSO Then gError 52388
        
        Else
            sContaFormatadaEst = ""
        End If
                
        If GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilProducaoBenef_Col) <> "" Then
        
            'Formata as Contas para o Bd
            lErro = CF("Conta_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilProducaoBenef_Col), sContaFormatadaProducao, iContaPreenchida)
            If lErro <> SUCESSO Then gError 52389
        
        Else
            sContaFormatadaProducao = ""
        End If
        
        objNFiscal.objMovEstoque.colItens.Add 0, 0, 0, 0, sProdutoFormatado, GridMovimentos.TextMatrix(iIndice, iGrid_DescricaoBenef_Col), GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMedBenef_Col), CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_QuantidadeBenef_Col)), objAlmoxarifado.iCodigo, GridMovimentos.TextMatrix(iIndice, iGrid_AlmoxarifadoBenef_Col), 0, "", 0, "", "", sContaFormatadaProducao, sContaFormatadaEst, 0, Nothing, Nothing, DATA_NULA
    
    Next
    
    Move_GridMovimentos_Memoria = SUCESSO

    Exit Function

Erro_Move_GridMovimentos_Memoria:

    Move_GridMovimentos_Memoria = gErr

    Select Case gErr
        
        Case 52385, 52386, 52388, 52389
        
        Case 52387
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)
           
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156640)

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
        If lErro <> SUCESSO Then gError 52390

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
            If lErro <> SUCESSO Then gError 52391

            ContaContabilEst.PromptInclude = False
            ContaContabilEst.Text = sContaEnxutaEst
            ContaContabilEst.PromptInclude = True

            GridMovimentos.TextMatrix(iIndice, iGrid_ContaContabilEstBenef_Col) = ContaContabilEst.Text

        End If

        'preenche contaEst no grid
        If objItemMovEstoque.sContaContabilAplic <> "" Then

            sContaEnxutaProducao = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objItemMovEstoque.sContaContabilAplic, sContaEnxutaProducao)
            If lErro <> SUCESSO Then gError 52392

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

    Preenche_GridMovimentos = gErr

    Select Case gErr

        Case 52390
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objItemMovEstoque.sProduto)
        
        Case 52391
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objItemMovEstoque.sContaContabilEst)

        Case 52392
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objItemMovEstoque.sContaContabilAplic)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156641)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ENTRADA_NF_SIMPLES_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Entrada - Nota Fiscal Simples"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "NFiscalEntrada"
    
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

Public Function ValorTotal_Calcula() As Long

Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    lErro = gobjTribTab.AtualizarTributacao()
    If lErro <> SUCESSO Then gError 56970
    
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

        Case 56970

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156642)

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
        If lErro <> SUCESSO Then gError 56979
        
        Call gobjTribTab.FilialFornecedor_Alterada(objFilialForn.lCodFornecedor, objFilialForn.iCodFilial)
    
        lErro = CF2(Me, "NFCF_Fornecedor_Filial_Validate", objFilialForn)
        If lErro <> SUCESSO Then gError 131935
    
    End If
    
    Trata_FilialForn = SUCESSO
     
    Exit Function
    
Erro_Trata_FilialForn:

    Trata_FilialForn = gErr
     
    Select Case gErr
          
        Case 56979, 131935
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156643)
     
    End Select
     
    Exit Function

End Function

Public Function Move_NF_Memoria(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objFornecedor As ClassFornecedor
Dim objNFiscalOriginal As New ClassNFiscal
Dim objProjetoInfo As Object

On Error GoTo Erro_Move_NF_Memoria

    'Se o Fornecedor estiver preenchido lê o código no BD
    If Len(Trim(Fornecedor.Text)) > 0 Then
        
        Set objFornecedor = New ClassFornecedor
        
        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        'Lê o fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 31419
        
        If lErro = 6681 Then gError 31420 'Se nao encontrar --> erro

        objNFiscal.lFornecedor = objFornecedor.lCodigo

    End If
    
    'Recolhe os demais campos da Tela
    If Len(Trim(NumRecebimento.Text)) > 0 Then objNFiscal.lNumRecebimento = CLng(NumRecebimento.Text)
    
    If Len(Trim(Filial.Text)) > 0 Then objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)

    If Len(Trim(TipoNFiscal.Text)) > 0 Then objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    If objNFiscal.iTipoDocInfo = TIPODOCINFO_TIPO_NFEXT Then
        objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
    Else
        objNFiscal.lNumNotaFiscal = StrParaLong(NFiscalInterna.Caption)
    End If
    
    objNFiscal.iFreteRespons = TipoFrete.ItemData(TipoFrete.ListIndex)
    
    objNFiscal.dtDataReferencia = DATA_NULA
    objNFiscal.sVolumeNumero = VolumeNumero.Text
    objNFiscal.lNumIntDoc = 0
    objNFiscal.sNaturezaOp = NaturezaOp.Text
    objNFiscal.sSerie = Converte_Serie_Eletronica(Serie.Text, EletronicaFed.Value)
    objNFiscal.sPlaca = Placa.Text
    objNFiscal.sPlacaUF = PlacaUF.Text
    If Not NFiscalInterna.Visible Then objNFiscal.sChvNFe = ChvNFe.ClipText
    
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objNFiscal.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objNFiscal.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
    
    objNFiscal.sMensagemNota = Mensagem.Text
    objNFiscal.sMensagemCorpoNota = MensagemCorpo.Text
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.dtDataSaida = DATA_NULA
    objNFiscal.iStatus = STATUS_LANCADO
    objNFiscal.sObservacao = Observacao.Text

    If Len(Trim(ValorDespesas.ClipText)) > 0 Then objNFiscal.dValorOutrasDespesas = CDbl(ValorDespesas.Text)
    If Len(Trim(ValorFrete.ClipText)) > 0 Then objNFiscal.dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorSeguro.ClipText)) > 0 Then objNFiscal.dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(Total.Text)) > 0 Then objNFiscal.dValorTotal = CDbl(Total.Text)
    If Len(Trim(SubTotal.Caption)) > 0 Then objNFiscal.dValorProdutos = CDbl(SubTotal.Caption)

    If Len(Trim(PesoBruto.ClipText)) > 0 Then objNFiscal.dPesoBruto = CDbl(PesoBruto.Text)
    If Len(Trim(PesoLiquido.ClipText)) > 0 Then objNFiscal.dPesoLiq = CDbl(PesoLiquido.Text)

    If Len(Trim(VolumeQuant.ClipText)) > 0 Then objNFiscal.lVolumeQuant = StrParaLong(VolumeQuant.Text)

    If Len(Trim(Transportadora.Text)) > 0 Then objNFiscal.iCodTransportadora = Codigo_Extrai(Transportadora.Text)
    
    'se o tipo de Nfiscal for do tipo "Externa - Mat. A ser beneficiado" entao obtem o Fornecedor a ser beneficiado
    If TipoNFiscal.ListIndex <> -1 Then
        
        If TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFEEBF Or TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFIEBF Or TipoNFiscal.ItemData(TipoNFiscal.ListIndex) = DOCINFO_NFECCO Then
                
            If Len(Trim(FornecedorBenef.Text)) > 0 Then
            
                Set objFornecedor = New ClassFornecedor
            
                objFornecedor.sNomeReduzido = FornecedorBenef.Text
            
                'Lê o fornecedor
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO And lErro <> 6681 Then gError 52288
            
                If lErro = 6681 Then gError 52289 'Se nao encontrar --> erro
    
                objNFiscal.lFornecedorBenef = objFornecedor.lCodigo
    
                If Len(Trim(FilialFornBenef.List(FilialFornBenef.ListIndex))) > 0 Then objNFiscal.iFilialFornBenef = Codigo_Extrai(FilialFornBenef.List(FilialFornBenef.ListIndex))
            End If
        End If
    End If
    
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
    
    If Len(Trim(dtVencimento.ClipText)) > 0 Then
        objNFiscal.dtDataVencimento = CDate(dtVencimento.Text)
    Else
        objNFiscal.dtDataVencimento = DATA_NULA
    End If
        
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181761
    
    Set objNFiscal.objProjetoInfo = objProjetoInfo
        
    'esta da indica quando está sendo registrada a nota fiscal no sistema. Será utilizado na gravacao do estoque
    objNFiscal.dtDataRegistro = gdtDataAtual
    
    Set objNFiscal.objInfoAdic = gobjInfoAdic
    objNFiscal.dValorItens = StrParaDbl(ValorProdutos2.Caption)
    objNFiscal.dValorDescontoItens = StrParaDbl(ValorDescontoItens.Text)
        
    Move_NF_Memoria = SUCESSO

    Exit Function

Erro_Move_NF_Memoria:

    Move_NF_Memoria = gErr

    Select Case gErr

        Case 31419, 52288, 181761

        Case 31420
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, Fornecedor.Text)
        
        Case 52289
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, FornecedorBenef.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156644)

    End Select

    Exit Function

End Function

Public Sub BotaoConsultaNFPag_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoConsultaNFPag_Click

    'Verifica se todos os campos necessários para se efetuar a consulta foram preenchidos
    lErro = Critica_CamposNecessarios_ConsultaTitulo()
    If lErro <> SUCESSO Then gError 79748
    
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
    If lErro <> SUCESSO And lErro <> 6681 Then gError 79749
    
    'Se não encontrou o fornecedor = > erro
    If lErro = 6681 Then gError 79750
    
    'Guarda no objNFiscal o código do fornecedor
    objNFiscal.lFornecedor = objFornecedor.lCodigo
    
    'Lê o NumIntDocCPR da NFiscal e exibe o documento gerado no CPR por essa nota
    lErro = CF("NFiscal_Consulta_DocCPR", objNFiscal)
    If lErro <> SUCESSO And lErro <> 79717 Then gError 79751
    
    'Se não encontrou a nota => erro
    If lErro = 79717 Then gError 79752
    
    Exit Sub
    
Erro_BotaoConsultaNFPag_Click:

    Select Case gErr
    
        Case 79748, 79751, 79749
        
        Case 79752
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_CADASTRADA1", gErr, objNFiscal.lNumNotaFiscal, objNFiscal.sSerie, objNFiscal.iTipoNFiscal, objNFiscal.lFornecedor, objNFiscal.iFilialForn, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada)
        
        Case 79750
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objNFiscal.lFornecedor)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156645)
    
    End Select
    
    Exit Sub

End Sub

Private Function Critica_CamposNecessarios_ConsultaTitulo() As Long
'Verifica se os campos necessários para consultar um título gerado por uma NFEntrada foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposNecessarios_ConsultaTitulo

    'Verifica se o Fornecedor foi preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 79753
    
    'Verifica se a filial do Fornecedor foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 79754
    
    'Verifica se a Série foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 79755
    
    'Verifica se o Número da Nota foi preenchido
    If Len(Trim(NFiscal.Text)) = 0 Then gError 79756
    
    'Verifica se a data de emissão da nota foi preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 79757
    
    'Verifica se a data de entrada da nota foi preenchida
    If Len(Trim(DataEntrada.ClipText)) = 0 Then gError 79758
    
    'Verifica se o tipo da Nota fiscal foi selecionado
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 79759
    
    Critica_CamposNecessarios_ConsultaTitulo = SUCESSO
    
    Exit Function

Erro_Critica_CamposNecessarios_ConsultaTitulo:

    Critica_CamposNecessarios_ConsultaTitulo = gErr
    
    Select Case gErr
    
        Case 79753
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 79754
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 79755
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 79756
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
        
        Case 79757
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 79758
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAENTRADA_NAO_PREENCHIDA", gErr)
        
        Case 79759
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156646)
    
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
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
'distribuicao
        ElseIf Me.ActiveControl Is gobjDistribuicao.AlmoxDist Then
            Call gobjDistribuicao.BotaoLocalizacaoDist_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
        ElseIf Me.ActiveControl Is Transportadora Then
            Call TransportadoraLabel_Click
        ElseIf Me.ActiveControl Is FornecedorBenef Then
            Call FornecedorBenefLabel_Click
        ElseIf Me.ActiveControl Is NFiscalOriginal Then
            Call NFiscalOriginalLabel_Click
        ElseIf Me.ActiveControl Is SerieNFiscalOriginal Then
            Call NFiscalOriginalLabel_Click
        ElseIf Me.ActiveControl Is FornNFOrig Then 'Incluído por Luiz - 24/01/02
            Call FornNFOrigLabel_Click
        ElseIf Me.ActiveControl Is ProdutoBenef Then
            Call BotaoProdutosBenef_Click
        ElseIf Me.ActiveControl Is AlmoxarifadoBenef Then
            Call BotaoEstoqueBenef_Click
        ElseIf Me.ActiveControl Is ContaContabilEst Then
            Call BotaoPlanoConta_Click
        ElseIf Me.ActiveControl Is DescricaoItem Then
            Call Produto_ChamaBrowser
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
Dim sMascaraCcl As String

On Error GoTo Erro_Form_Load_Light

    'obs.: a carga pos form_load será feita pela classe que trata nffat
    giPosCargaOk = 1
    
    iFrameAtual = 1
    
    'Inicializa as variáveis globais da tela(objGrid, objEventos)
    Set objGridBenef = New AdmGrid
    Set objGridItens = New AdmGrid
    Set objEventoSerie = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoNFiscal = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoFornecedorBenef = New AdmEvento
    Set objEventoFornNFOrig = New AdmEvento
    Set objEventoClienteNFOrig = New AdmEvento
    Set objEventoRecebimento = New AdmEvento
    Set objEventoSerieOriginal = New AdmEvento
    Set objEventoNFiscalOriginal = New AdmEvento
    Set objEventoTransportadora = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoProdutoBenef = New AdmEvento
    Set objEventoEstoqueBenef = New AdmEvento
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156647)

    End Select

    iAlterado = 0
    
    Exit Function

End Function

'mario inicio
Private Sub Produto_ChamaBrowser()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection
Dim sProduto1 As String
Dim sOrdenacao As String

On Error GoTo Erro_Produto_ChamaBrowser

    Select Case left(Trim(DescricaoItem.Text), 1)
    
        Case "*"
            objProduto.sNomeReduzido = Mid(LTrim(DescricaoItem.Text), 2)
            sOrdenacao = "NomeReduzido"

    End Select

    'Chama a Tela ProdutoCompraLista
    Call Chama_Tela("ProdutoCompraLista", colSelecao, objProduto, objEventoProduto, "", sOrdenacao)

    Exit Sub

Erro_Produto_ChamaBrowser:

    Select Case gErr
    
        Case 43715
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 83078
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156648)
    
    End Select
    
    Exit Sub
    
End Sub

'mario fim

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156649)
            
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
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134060

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134060

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156650)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156651)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156652)
     
    End Select
     
    Exit Sub

End Sub

Public Function NFEnt_AlocaAlmoxarifadoPadrao(ByVal objCT As Object, ByVal gobjDistribuicao As Object) As Long
    NFEnt_AlocaAlmoxarifadoPadrao = SUCESSO
End Function

Public Function NFEnt_Form_Unload(ByVal objCT As Object) As Long
    NFEnt_Form_Unload = SUCESSO
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

Public Function NFE_Calcula_PesoKg(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    NFE_Calcula_PesoKg = SUCESSO
End Function

Public Function NFE_Calcula_PesoKg_Total(ByVal objCT As Object) As Long
    NFE_Calcula_PesoKg_Total = SUCESSO
End Function

Public Function NFE_Move_GridItem_Memoria(ByVal objCT As Object, ByVal objItemNF As ClassItemNF, ByVal iLinha As Integer) As Long
    NFE_Move_GridItem_Memoria = SUCESSO
End Function

Public Function NFE_PrecoTotal_Calcula(ByVal objCT As Object, ByVal iLinha As Integer, dQuantidade As Double) As Long
    NFE_PrecoTotal_Calcula = SUCESSO
End Function

Function NFE_Preenche_GridItens(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long
    NFE_Preenche_GridItens = SUCESSO
End Function

Public Function NFE_ProdutoLinha_Preenche(ByVal objCT As Object, objProduto As ClassProduto) As Long
    NFE_ProdutoLinha_Preenche = SUCESSO
End Function

Public Function NFE_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iLocalChamada As Integer, ByVal sProdutoFormatado As String, ByVal iProdutoPreenchido As Integer) As Long
    NFE_Rotina_Grid_Enable = SUCESSO
End Function

Public Function NFE_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    NFE_Saida_Celula_GridItens = SUCESSO
End Function

Public Function NFE_Valida_Grid_Itens(ByVal objCT As Object, ByVal iLinha As Integer) As Long
    NFE_Valida_Grid_Itens = SUCESSO
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
        objItensPRJCR.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) & SEPARADOR & GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)

        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        objItensPRJCR.dValorOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))

        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NF
        
        colItensPRJ.Add objItensPRJCR

    Next

    Call gobjTribTab.Preenche_Itens_Projeto_NF(colItensPRJ, True)

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
    If lErro <> SUCESSO Then gError 207455

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

        Case 207455

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207456)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211291)
            
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoTrazerNFe_Click()

Dim lErro  As Long
Dim objTela As Object, objNF As New ClassNFiscal

On Error GoTo Erro_BotaoTrazerNFe_Click

    If Len(Trim(ChvNFe.ClipText)) = 0 Then gError 211292

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
    
        Case 211292
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CHAVE_NAO_PREENCHIDA", gErr)
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211293)
            
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211294)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211295)

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
Dim dPercDesc As Double

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
    
    'Se o código da Nota Fiscal Original estiver preenchido
    If objNFiscal.lNumIntNotaOriginal <> 0 Then
    
        objNFiscalOriginal.lNumIntDoc = objNFiscal.lNumIntNotaOriginal
        
        'Lê a Nota Fiscal original
        lErro = CF("NFiscal_Le", objNFiscalOriginal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
        
        'Coloca os dados relativos  Nota Fiscal Original na Tela
        SerieNFiscalOriginal.Text = objNFiscalOriginal.sSerie
        NFiscalOriginal.Text = objNFiscalOriginal.lNumNotaFiscal
        
        If objNFiscal.iTipoNFiscal = DOCINFO_CFEV Then
        
            'Preenche o Cliente da NF Orig
            ClienteNFOrig.Text = objNFiscalOriginal.lCliente
            Call ClienteNFOrig_Validate(bCancel)
            
            'Preenche a Filial da Cliente
            FilialCliNFOrig.Text = objNFiscalOriginal.iFilialCli
            Call FilialCliNFOrig_Validate(bCancel)
            
        ElseIf (objNFiscal.iTipoNFiscal = DOCINFO_CFECT) Or (objNFiscal.iTipoNFiscal = DOCINFO_DDAI) Or (objNFiscal.iTipoNFiscal = DOCINFO_NFSET) Then
    
            'Preenche o Fornecedor da NF Orig
            FornNFOrig.Text = objNFiscalOriginal.lFornecedor
            Call FornNFOrig_Validate(bCancel)
            
            'Preenche a Filial do Fornecedor
            FilialFornNFOrig.Text = objNFiscalOriginal.iFilialForn
            Call FilialFornNFOrig_Validate(bCancel)
    
        End If
        
    End If
    
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO And gobjMAT.iContabPelaDtEntXml = MARCADO Then
        Call objContabil.Contabil_Carga_Modelo_Padrao
        
        Call DateParaMasked(CTBDataContabil, objNFiscal.dtDataEntrada)
        Call CTBDataContabil_Validate(bSGECancelDummy)
    End If

    gbCarregandoTela = False
    
    Traz_NFiscal_Tela1 = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela1:

    gbCarregandoTela = False
    
    Traz_NFiscal_Tela1 = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211296)

    End Select

    Exit Function

End Function
