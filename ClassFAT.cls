VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvariPedidoReservaAutomatica As Integer
Private mvariNFiscalAlocacaoAutomatica As Integer
Private mvariPedidoVendaEditaComissao As Integer
Private mvariNFiscalEditaComissao As Integer
Private mvariAglutinaLancamPorDia As Integer
Private mvariGeraLoteAutomatico As Integer
Private mvariFilialFaturamento As Integer
Private mvarsFormatoPrecoUnitario As String
Private mvarsCategMoeda As String
Private mvarsCategCustoFixo As String
Private mvariPVPreenchCondPagto As Integer
Private mvariPVPreenchDataEntrega As Integer
Private mvariCliPreenchCondPagto As Integer
Private mvariCliPreenchTransp As Integer
Private mvariCliPreenchCNPJ As Integer
Private mvariCliPreenchInscEst As Integer
Private mvariPedidoValidaEmbalagem As Integer
Private mvariTemVersaoOV As Integer
Private mvariBrowseFecha As Integer
Private mvariTipoValidacaoPRJ As Integer
Private mvariCotarItensOrcVenda As Integer
Private mvariTemFaixaReceb As Integer
Private mvariRecebForaFaixa As Integer
Private mvardPercentMaisReceb As Double
Private mvardPercentMenosReceb As Double
Private mvariExibePercFatPV As Integer
Private mvariAbateICMSSubstVlrCobr As Integer
Private mvariFaltaEstObrigaDecisao As Integer
Private mvariImpedeImpPVComBloq As Integer
Private mvariVFEUsaTManual As Integer
Private mvariGeraFatTodasFiliais As Integer
Private mvariExibeRotasCliente As Integer
Private mvariPossuiIntMapLink As Integer
Private mvariImpDadosInterno As Integer
Private mvariPulaLinhaGrade As Integer
Private mvariImprimePVGravacao As Integer
Private mvarlCodCliGov As Long
Private mvarlCodCliNFCeSemCPF As Long
Private mvarsNatOpGov As String
Private mvariTrazDescPVParaNF As Integer
Private mvariExibeOPImpPreco As Integer
Private mvariInativaCliDiaSemCompra As Integer
Private mvariExibeDataSaida As Integer
Private mvarsDirImportPVSOC As String
Private mvariPedVendaCodBarra As Integer
Private mvariBloqEditPrecoPV As Integer
Private mvariExibeOPImpCodProd As Integer
Private mvariExibeOPCalcST As Integer
Private mvariCalcSTPadrao As Integer
Private mvariPrimParcICMSST As Integer
Private mvariPrimParcIPI As Integer
Private mvariPrimParcPISST As Integer
Private mvariPrimParcCOFINSST As Integer
Private mvariTrataNFDevSimbPorOP As Integer
Private mvariNFDevSimbSugOP As Integer
Private mvariNFDevSimbAlmSobra As Integer
Private mvariNFDevSimbUM As Integer

Private mvariCliPreenchVend As Integer
Private mvariCliPreenchDtUltVi As Integer
Private mvariCliPreenchEnd As Integer
Private mvariCliPreenchEndNum As Integer
Private mvariCliPreenchEndCompl As Integer
Private mvariCliPreenchCEP As Integer
Private mvariCliPreenchTelefone As Integer
Private mvariCliPreenchEmail As Integer
Private mvariOVExibeOrdItens As Integer
Private mvariOVExibeTS As Integer
Private mvariCliPreenchContato As Integer
Private mvariOVAcompFiltraVend As Integer
Private mvariOVAcompFiltroVendObr As Integer
Private mvariUsaParcDePV As Integer

Private mvariExibeTabPrecoItensPV As Integer
Private mvariExibeComissaoItensPV As Integer
Private mvariExibeQtdeCanceladaItensPV As Integer
Private mvariAtualizaPlacaPVMapaEntrega As Integer
Private mvariUsaComissaoPorItem As Integer
Private mvariKitVendaExibeComp As Integer
Private mvariExibeBtnExportaPV As Integer
Private mvariZeraQtdeEmNFComplPreco As Integer
Private mvariParcRecCobraSTNovaParc As Integer
Private mvariParcSTNumDias As Integer
Private mvariOVSemFormPreco As Integer
Private mvariPesosRecalcula As Integer
Private mvariPRJObrigatorio As Integer
Private mvariOVBloqEditDescProd As Integer
Private mvariNFFaturaApenasQuantReservPV As Integer
Private mvariPRJExibeVistorias As Integer
Private mvariPRJTipoPrazoValidVist As Integer
Private mvariPRJPrazoValidVist As Integer
Private mvariChvNFeValidaCNPJ As Integer

Public Function Inicializa() As Long

Dim lErro As Long, colConfig As New ColFATConfig

On Error GoTo Erro_Inicializa

    colConfig.Add FATCFG_PVPREENCH_CONDPAGTO, EMPRESA_TODA, "", 0, "", FATCFG_PVPREENCH_CONDPAGTO
    colConfig.Add FATCFG_CLIPREENCH_CONDPAGTO, EMPRESA_TODA, "", 0, "", FATCFG_CLIPREENCH_CONDPAGTO
    colConfig.Add FATCFG_CLIPREENCH_TRANSP, EMPRESA_TODA, "", 0, "", FATCFG_CLIPREENCH_TRANSP
    colConfig.Add FATCFG_CLIPREENCH_CNPJ, EMPRESA_TODA, "", 0, "", FATCFG_CLIPREENCH_CNPJ
    colConfig.Add FATCFG_CLIPREENCH_INSCEST, EMPRESA_TODA, "", 0, "", FATCFG_CLIPREENCH_INSCEST
    
    colConfig.Add FATCFG_CLIPREENCH_VEND, EMPRESA_TODA, "", 0, "", FATCFG_CLIPREENCH_VEND
    colConfig.Add FATCFG_CLIPREENCH_DTULTVI, EMPRESA_TODA, "", 0, "", FATCFG_CLIPREENCH_DTULTVI
    colConfig.Add FATCFG_CLIPREENCH_END, EMPRESA_TODA, "", 0, "", FATCFG_CLIPREENCH_END
    colConfig.Add FATCFG_CLIPREENCH_ENDNUM, EMPRESA_TODA, "", 0, "", FATCFG_CLIPREENCH_ENDNUM
    colConfig.Add FATCFG_CLIPREENCH_CEP, EMPRESA_TODA, "", 0, "", FATCFG_CLIPREENCH_CEP
    colConfig.Add FATCFG_CLIPREENCH_TELEFONE, EMPRESA_TODA, "", 0, "", FATCFG_CLIPREENCH_TELEFONE
    colConfig.Add FATCFG_CLIPREENCH_EMAIL, EMPRESA_TODA, "", 0, "", FATCFG_CLIPREENCH_EMAIL
    colConfig.Add FATCFG_CLIPREENCH_ENDCOMPL, EMPRESA_TODA, "", 0, "", FATCFG_CLIPREENCH_ENDCOMPL
    colConfig.Add FATCFG_CLIPREENCH_CONTATO, EMPRESA_TODA, "", 0, "", FATCFG_CLIPREENCH_CONTATO

    colConfig.Add FATCFG_PEDIDO_RESERVA_AUTOMATICA, EMPRESA_TODA, "", 0, "", FATCFG_PEDIDO_RESERVA_AUTOMATICA
    colConfig.Add FATCFG_NFISCAL_ALOCA_AUTOMATICA, EMPRESA_TODA, "", 0, "", FATCFG_NFISCAL_ALOCA_AUTOMATICA
    colConfig.Add FATCFG_PEDIDO_VENDA_EDITA_COMISSAO, EMPRESA_TODA, "", 0, "", FATCFG_PEDIDO_VENDA_EDITA_COMISSAO
    colConfig.Add FATCFG_NFISCAL_EDITA_COMISSAO, EMPRESA_TODA, "", 0, "", FATCFG_NFISCAL_EDITA_COMISSAO
    colConfig.Add FATCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, "", FATCFG_AGLUTINA_LANCAM_POR_DIA
    colConfig.Add FATCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, "", FATCFG_GERA_LOTE_AUTOMATICO
    If giFilialEmpresa <> EMPRESA_TODA Then colConfig.Add FATCFG_FILIAL_FATURAMENTO, giFilialEmpresa, "", 0, "", FATCFG_FILIAL_FATURAMENTO
    colConfig.Add FATCFG_FORMATO_PRECO_UNITARIO, EMPRESA_TODA, "", 0, "", FATCFG_FORMATO_PRECO_UNITARIO
    colConfig.Add FATCFG_CATEG_CUSTO_FIXO, EMPRESA_TODA, "", 0, "", FATCFG_CATEG_CUSTO_FIXO
    colConfig.Add FATCFG_CATEG_MOEDA, EMPRESA_TODA, "", 0, "", FATCFG_CATEG_MOEDA
    colConfig.Add FATCFG_PEDIDO_VALIDA_EMBALAGEM, EMPRESA_TODA, "", 0, "", FATCFG_PEDIDO_VALIDA_EMBALAGEM
    colConfig.Add FATCFG_TEM_VERSAO_OV, EMPRESA_TODA, "", 0, "", FATCFG_TEM_VERSAO_OV 'Inserido por Wagner
    colConfig.Add FATCFG_BROWSE_FECHA, EMPRESA_TODA, "", 0, "", FATCFG_BROWSE_FECHA 'Inserido por Wagner
    colConfig.Add FATCFG_TIPO_VALIDACAO_PRJ, EMPRESA_TODA, "", 0, "", FATCFG_TIPO_VALIDACAO_PRJ 'Inserido por Wagner
    colConfig.Add FATCFG_COTAR_ITENS_ORCVENDA, EMPRESA_TODA, "", 0, "", FATCFG_COTAR_ITENS_ORCVENDA
 
    colConfig.Add FATCFG_TEM_FAIXA_RECEB, EMPRESA_TODA, "", 0, "", FATCFG_TEM_FAIXA_RECEB
    colConfig.Add FATCFG_PERCENT_MAIS_RECEB, EMPRESA_TODA, "", 0, "", FATCFG_PERCENT_MAIS_RECEB
    colConfig.Add FATCFG_PERCENT_MENOS_RECEB, EMPRESA_TODA, "", 0, "", FATCFG_PERCENT_MENOS_RECEB
    colConfig.Add FATCFG_RECEB_FORA_FAIXA, EMPRESA_TODA, "", 0, "", FATCFG_RECEB_FORA_FAIXA
    colConfig.Add FATCFG_EXIBE_PERCFAT_PV, EMPRESA_TODA, "", 0, "", FATCFG_EXIBE_PERCFAT_PV
 
    colConfig.Add FATCFG_ABATE_ICMSSUBST_VALOR_COBRANCA, EMPRESA_TODA, "", 0, "", FATCFG_ABATE_ICMSSUBST_VALOR_COBRANCA
    colConfig.Add FATCFG_FALTA_ESTOQUE_OBRIGA_TOMADA_DECISAO, EMPRESA_TODA, "", 0, "", FATCFG_FALTA_ESTOQUE_OBRIGA_TOMADA_DECISAO
    colConfig.Add FATCFG_IMPEDE_IMP_PV_COM_BLOQ, EMPRESA_TODA, "", 0, "", FATCFG_IMPEDE_IMP_PV_COM_BLOQ
    colConfig.Add FATCFG_VENDA_FORA_EST_USA_TALAO_MANUAL, EMPRESA_TODA, "", 0, "", FATCFG_VENDA_FORA_EST_USA_TALAO_MANUAL
    colConfig.Add FATCFG_GERACAO_FATURA_TODAS_FILIAIS, EMPRESA_TODA, "", 0, "", FATCFG_GERACAO_FATURA_TODAS_FILIAIS
    colConfig.Add FATCFG_EXIBE_ROTAS_CLIENTE, EMPRESA_TODA, "", 0, "", FATCFG_EXIBE_ROTAS_CLIENTE
    colConfig.Add FATCFG_POSSUI_INTEGRACAO_MAPLINK, EMPRESA_TODA, "", 0, "", FATCFG_POSSUI_INTEGRACAO_MAPLINK
    colConfig.Add FATCFG_IMPORTACAO_COM_DADOS_INTERNOS, EMPRESA_TODA, "", 0, "", FATCFG_IMPORTACAO_COM_DADOS_INTERNOS
    colConfig.Add FATCFG_PULA_LINHA_GRADE, EMPRESA_TODA, "", 0, "", FATCFG_PULA_LINHA_GRADE
    colConfig.Add FATCFG_IMPRIME_PV_GRAVACAO, EMPRESA_TODA, "", 0, "", FATCFG_IMPRIME_PV_GRAVACAO
    colConfig.Add FATCFG_NATUREZAOP_NF_TIPO_226_GOVERNO, EMPRESA_TODA, "", 0, "", FATCFG_NATUREZAOP_NF_TIPO_226_GOVERNO
    colConfig.Add FATCFG_CLIENTE_NF_TIPO_226_GOVERNO, EMPRESA_TODA, "", 0, "", FATCFG_CLIENTE_NF_TIPO_226_GOVERNO
    
    colConfig.Add FATCFG_CLIENTE_NFCE_SEM_CPF, EMPRESA_TODA, "", 0, "", FATCFG_CLIENTE_NFCE_SEM_CPF
    
    colConfig.Add FATCFG_TRAZ_DESCONTO_PV_PARA_NF, EMPRESA_TODA, "", 0, "", FATCFG_TRAZ_DESCONTO_PV_PARA_NF
    colConfig.Add FATCFG_EXIBE_OPCAO_IMPRIMIR_OV_SEM_PRECO, EMPRESA_TODA, "", 0, "", FATCFG_EXIBE_OPCAO_IMPRIMIR_OV_SEM_PRECO
    colConfig.Add FATCFG_INATIVA_CLIENTE_DIAS_SEM_COMPRA, EMPRESA_TODA, "", 0, "", FATCFG_INATIVA_CLIENTE_DIAS_SEM_COMPRA
    colConfig.Add FATCFG_EXIBE_DATA_SAIDA_NF, EMPRESA_TODA, "", 0, "", FATCFG_EXIBE_DATA_SAIDA_NF
    
    colConfig.Add FATCFG_DIR_IMPORT_PV_SOC, EMPRESA_TODA, "", 0, "", FATCFG_DIR_IMPORT_PV_SOC
    colConfig.Add FATCFG_PEDVENDA_CODBARRA, EMPRESA_TODA, "", 0, "", FATCFG_PEDVENDA_CODBARRA
    colConfig.Add FATCFG_BLOQUEIA_EDICAO_PRECO_PV, EMPRESA_TODA, "", 0, "", FATCFG_BLOQUEIA_EDICAO_PRECO_PV

    colConfig.Add FATCFG_EXIBE_OPCAO_IMPRIMIR_OV_COM_CODPROD, EMPRESA_TODA, "", 0, "", FATCFG_EXIBE_OPCAO_IMPRIMIR_OV_COM_CODPROD
    colConfig.Add FATCFG_EXIBE_OPCAO_OV_CALCULAR_ST, EMPRESA_TODA, "", 0, "", FATCFG_EXIBE_OPCAO_OV_CALCULAR_ST
    colConfig.Add FATCFG_CALCULAR_ST_OP_PADRAO, EMPRESA_TODA, "", 0, "", FATCFG_CALCULAR_ST_OP_PADRAO

    colConfig.Add FATCFG_PRIM_PARC_ICMSST, EMPRESA_TODA, "", 0, "", FATCFG_PRIM_PARC_ICMSST
    colConfig.Add FATCFG_PRIM_PARC_IPI, EMPRESA_TODA, "", 0, "", FATCFG_PRIM_PARC_IPI
    colConfig.Add FATCFG_PRIM_PARC_PISST, EMPRESA_TODA, "", 0, "", FATCFG_PRIM_PARC_PISST
    colConfig.Add FATCFG_PRIM_PARC_COFINSST, EMPRESA_TODA, "", 0, "", FATCFG_PRIM_PARC_COFINSST

    colConfig.Add FATCFG_TRATA_NF_DEV_SIMB_BENEF_POR_OP, EMPRESA_TODA, "", 0, "", FATCFG_TRATA_NF_DEV_SIMB_BENEF_POR_OP
    colConfig.Add FATCFG_NF_DEV_SIMB_BENEF_SUGERIR_OPS, EMPRESA_TODA, "", 0, "", FATCFG_NF_DEV_SIMB_BENEF_SUGERIR_OPS
    colConfig.Add FATCFG_NF_DEV_SIMB_BENEF_ALM_SOBRA, EMPRESA_TODA, "", 0, "", FATCFG_NF_DEV_SIMB_BENEF_ALM_SOBRA
    colConfig.Add FATCFG_NF_DEV_SIMB_UM, EMPRESA_TODA, "", 0, "", FATCFG_NF_DEV_SIMB_UM

    colConfig.Add FATCFG_OV_EXIBE_ORDENACAO_ITENS, EMPRESA_TODA, "", 0, "", FATCFG_OV_EXIBE_ORDENACAO_ITENS
    colConfig.Add FATCFG_OV_EXIBE_TOTAL_SELECIONADO, EMPRESA_TODA, "", 0, "", FATCFG_OV_EXIBE_TOTAL_SELECIONADO

    colConfig.Add FATCFG_OV_ACOMP_FILTRA_VENDEDOR, EMPRESA_TODA, "", 0, "", FATCFG_OV_ACOMP_FILTRA_VENDEDOR
    colConfig.Add FATCFG_OV_SEM_FORM_PRECO, EMPRESA_TODA, "", 0, "", FATCFG_OV_SEM_FORM_PRECO
    colConfig.Add FATCFG_PESOS_RECALCULA, EMPRESA_TODA, "", 0, "", FATCFG_PESOS_RECALCULA
    colConfig.Add FATCFG_OV_ACOMP_FILTRO_VENDEDOR_OBRIGATORIO, EMPRESA_TODA, "", 0, "", FATCFG_OV_ACOMP_FILTRO_VENDEDOR_OBRIGATORIO
    colConfig.Add FATCFG_USA_PARCELAMENTO_DO_PV, EMPRESA_TODA, "", 0, "", FATCFG_USA_PARCELAMENTO_DO_PV

    colConfig.Add FATCFG_EXIBE_TABPRECO_ITENSPV, EMPRESA_TODA, "", 0, "", FATCFG_EXIBE_TABPRECO_ITENSPV
    colConfig.Add FATCFG_EXIBE_COMISSAO_ITENSPV, EMPRESA_TODA, "", 0, "", FATCFG_EXIBE_COMISSAO_ITENSPV
    colConfig.Add FATCFG_EXIBE_QTDECANCELADA_ITENSPV, EMPRESA_TODA, "", 0, "", FATCFG_EXIBE_QTDECANCELADA_ITENSPV
    colConfig.Add FATCFG_ATUALIZA_PLACAPV_MAPAENTREGA, EMPRESA_TODA, "", 0, "", FATCFG_ATUALIZA_PLACAPV_MAPAENTREGA
    colConfig.Add FATCFG_USA_COMISSAO_POR_ITEM, EMPRESA_TODA, "", 0, "", FATCFG_USA_COMISSAO_POR_ITEM
    colConfig.Add FATCFG_KIT_VENDA_EXIBE_COMPONENTES, EMPRESA_TODA, "", 0, "", FATCFG_KIT_VENDA_EXIBE_COMPONENTES
    
    colConfig.Add FATCFG_EXIBE_BTN_EXPORTA_PV, EMPRESA_TODA, "", 0, "", FATCFG_EXIBE_BTN_EXPORTA_PV
    colConfig.Add FATCFG_ZERA_QTDE_EM_NF_COMPL_PRECO, EMPRESA_TODA, "", 0, "", FATCFG_ZERA_QTDE_EM_NF_COMPL_PRECO
    
    colConfig.Add FATCFG_PACREC_COBRA_ST_NOVA_PARC, EMPRESA_TODA, "", 0, "", FATCFG_PACREC_COBRA_ST_NOVA_PARC
    colConfig.Add FATCFG_PARC_ST_NUM_DIAS, EMPRESA_TODA, "", 0, "", FATCFG_PARC_ST_NUM_DIAS
    
    colConfig.Add FATCFG_OBRIGA_PREENCH_PRJ, EMPRESA_TODA, "", 0, "", FATCFG_OBRIGA_PREENCH_PRJ
    
    colConfig.Add FATCFG_OV_BLOQUEIA_EDIT_DESCPROD, EMPRESA_TODA, "", 0, "", FATCFG_OV_BLOQUEIA_EDIT_DESCPROD
    
    colConfig.Add FATCFG_NF_FATURA_APENAS_O_RESERVADO_NO_PV, EMPRESA_TODA, "", 0, "", FATCFG_NF_FATURA_APENAS_O_RESERVADO_NO_PV
    
    colConfig.Add FATCFG_PRJ_EXIBE_VISTORIA_ETAPAS, EMPRESA_TODA, "", 0, "", FATCFG_PRJ_EXIBE_VISTORIA_ETAPAS
    colConfig.Add FATCFG_PRJ_PRAZO_VALIDADE_VISTORIA, EMPRESA_TODA, "", 0, "", FATCFG_PRJ_PRAZO_VALIDADE_VISTORIA
    colConfig.Add FATCFG_PRJ_TIPO_PRAZO_VALIDADE_VISTORIA, EMPRESA_TODA, "", 0, "", FATCFG_PRJ_TIPO_PRAZO_VALIDADE_VISTORIA
    colConfig.Add FATCFG_CHVNFE_VALIDA_CNPJ, EMPRESA_TODA, "", 0, "", FATCFG_CHVNFE_VALIDA_CNPJ
    
    colConfig.Add FATCFG_PVPREENCH_DATAENTREGA, EMPRESA_TODA, "", 0, "", FATCFG_PVPREENCH_DATAENTREGA
    
    'Lê as configurações em FATConfig
    lErro = CF("FATConfig_Le_Configs", colConfig)
    If lErro <> SUCESSO Then Error 25082
    
    mvariPVPreenchDataEntrega = CInt(colConfig.Item(FATCFG_PVPREENCH_DATAENTREGA).sConteudo)
    
    mvariChvNFeValidaCNPJ = CInt(colConfig.Item(FATCFG_CHVNFE_VALIDA_CNPJ).sConteudo)
    mvariPRJExibeVistorias = CInt(colConfig.Item(FATCFG_PRJ_EXIBE_VISTORIA_ETAPAS).sConteudo)
    mvariPRJPrazoValidVist = CInt(colConfig.Item(FATCFG_PRJ_PRAZO_VALIDADE_VISTORIA).sConteudo)
    mvariPRJTipoPrazoValidVist = CInt(colConfig.Item(FATCFG_PRJ_TIPO_PRAZO_VALIDADE_VISTORIA).sConteudo)
    
    mvariNFFaturaApenasQuantReservPV = CInt(colConfig.Item(FATCFG_NF_FATURA_APENAS_O_RESERVADO_NO_PV).sConteudo)
    
    mvariOVBloqEditDescProd = CInt(colConfig.Item(FATCFG_OV_BLOQUEIA_EDIT_DESCPROD).sConteudo)
    
    mvariPRJObrigatorio = CInt(colConfig.Item(FATCFG_OBRIGA_PREENCH_PRJ).sConteudo)
    
    mvariParcRecCobraSTNovaParc = CInt(colConfig.Item(FATCFG_PACREC_COBRA_ST_NOVA_PARC).sConteudo)
    mvariParcSTNumDias = CInt(colConfig.Item(FATCFG_PARC_ST_NUM_DIAS).sConteudo)
    
    mvariZeraQtdeEmNFComplPreco = CInt(colConfig.Item(FATCFG_ZERA_QTDE_EM_NF_COMPL_PRECO).sConteudo)
    
    mvariExibeBtnExportaPV = CInt(colConfig.Item(FATCFG_EXIBE_BTN_EXPORTA_PV).sConteudo)
    
    mvariKitVendaExibeComp = CInt(colConfig.Item(FATCFG_KIT_VENDA_EXIBE_COMPONENTES).sConteudo)
    
    mvariOVAcompFiltraVend = CInt(colConfig.Item(FATCFG_OV_ACOMP_FILTRA_VENDEDOR).sConteudo)
    mvariOVSemFormPreco = CInt(colConfig.Item(FATCFG_OV_SEM_FORM_PRECO).sConteudo)
    mvariPesosRecalcula = CInt(colConfig.Item(FATCFG_PESOS_RECALCULA).sConteudo)
    mvariOVAcompFiltroVendObr = CInt(colConfig.Item(FATCFG_OV_ACOMP_FILTRO_VENDEDOR_OBRIGATORIO).sConteudo)

    mvariOVExibeOrdItens = CInt(colConfig.Item(FATCFG_OV_EXIBE_ORDENACAO_ITENS).sConteudo)
    mvariOVExibeTS = CInt(colConfig.Item(FATCFG_OV_EXIBE_TOTAL_SELECIONADO).sConteudo)

    mvariUsaParcDePV = CInt(colConfig.Item(FATCFG_USA_PARCELAMENTO_DO_PV).sConteudo)

    mvarsDirImportPVSOC = colConfig.Item(FATCFG_DIR_IMPORT_PV_SOC).sConteudo
    
    mvariPVPreenchCondPagto = CInt(colConfig.Item(FATCFG_PVPREENCH_CONDPAGTO).sConteudo)
    mvariCliPreenchCondPagto = CInt(colConfig.Item(FATCFG_CLIPREENCH_CONDPAGTO).sConteudo)
    mvariCliPreenchTransp = CInt(colConfig.Item(FATCFG_CLIPREENCH_TRANSP).sConteudo)
    mvariCliPreenchCNPJ = CInt(colConfig.Item(FATCFG_CLIPREENCH_CNPJ).sConteudo)
    mvariCliPreenchInscEst = CInt(colConfig.Item(FATCFG_CLIPREENCH_INSCEST).sConteudo)
    
    mvariCliPreenchVend = CInt(colConfig.Item(FATCFG_CLIPREENCH_VEND).sConteudo)
    mvariCliPreenchDtUltVi = CInt(colConfig.Item(FATCFG_CLIPREENCH_DTULTVI).sConteudo)
    mvariCliPreenchEnd = CInt(colConfig.Item(FATCFG_CLIPREENCH_END).sConteudo)
    mvariCliPreenchEndNum = CInt(colConfig.Item(FATCFG_CLIPREENCH_ENDNUM).sConteudo)
    mvariCliPreenchCEP = CInt(colConfig.Item(FATCFG_CLIPREENCH_CEP).sConteudo)
    mvariCliPreenchTelefone = CInt(colConfig.Item(FATCFG_CLIPREENCH_TELEFONE).sConteudo)
    mvariCliPreenchEmail = CInt(colConfig.Item(FATCFG_CLIPREENCH_EMAIL).sConteudo)
    mvariCliPreenchEndCompl = CInt(colConfig.Item(FATCFG_CLIPREENCH_ENDCOMPL).sConteudo)
    mvariCliPreenchContato = CInt(colConfig.Item(FATCFG_CLIPREENCH_CONTATO).sConteudo)
    
    mvariPedidoReservaAutomatica = CInt(colConfig.Item(FATCFG_PEDIDO_RESERVA_AUTOMATICA).sConteudo)
    mvariNFiscalAlocacaoAutomatica = CInt(colConfig.Item(FATCFG_NFISCAL_ALOCA_AUTOMATICA).sConteudo)
    mvariPedidoVendaEditaComissao = CInt(colConfig.Item(FATCFG_PEDIDO_VENDA_EDITA_COMISSAO).sConteudo)
    mvariNFiscalEditaComissao = CInt(colConfig.Item(FATCFG_NFISCAL_EDITA_COMISSAO).sConteudo)
    mvariAglutinaLancamPorDia = CInt(colConfig.Item(FATCFG_AGLUTINA_LANCAM_POR_DIA).sConteudo)
    mvariGeraLoteAutomatico = CInt(colConfig.Item(FATCFG_GERA_LOTE_AUTOMATICO).sConteudo)
    If giFilialEmpresa <> EMPRESA_TODA Then mvariFilialFaturamento = CInt(colConfig.Item(FATCFG_FILIAL_FATURAMENTO).sConteudo)
    mvarsFormatoPrecoUnitario = colConfig.Item(FATCFG_FORMATO_PRECO_UNITARIO).sConteudo
    mvarsCategCustoFixo = colConfig.Item(FATCFG_CATEG_CUSTO_FIXO).sConteudo
    mvarsCategMoeda = colConfig.Item(FATCFG_CATEG_MOEDA).sConteudo
    mvariPedidoValidaEmbalagem = CInt(colConfig.Item(FATCFG_PEDIDO_VALIDA_EMBALAGEM).sConteudo)
    mvariTemVersaoOV = CInt(colConfig.Item(FATCFG_TEM_VERSAO_OV).sConteudo) 'Inserido por Wagner
    mvariBrowseFecha = CInt(colConfig.Item(FATCFG_BROWSE_FECHA).sConteudo) 'Inserido por Wagner
    mvariTipoValidacaoPRJ = CInt(colConfig.Item(FATCFG_TIPO_VALIDACAO_PRJ).sConteudo) 'Inserido por Wagner
    mvariCotarItensOrcVenda = CInt(colConfig.Item(FATCFG_COTAR_ITENS_ORCVENDA).sConteudo)

    mvariTemFaixaReceb = CInt(colConfig.Item(FATCFG_TEM_FAIXA_RECEB).sConteudo)
    mvardPercentMaisReceb = CDbl(colConfig.Item(FATCFG_PERCENT_MAIS_RECEB).sConteudo)
    mvardPercentMenosReceb = CDbl(colConfig.Item(FATCFG_PERCENT_MENOS_RECEB).sConteudo)
    mvariRecebForaFaixa = CInt(colConfig.Item(FATCFG_RECEB_FORA_FAIXA).sConteudo)
    mvariExibePercFatPV = CInt(colConfig.Item(FATCFG_EXIBE_PERCFAT_PV).sConteudo)

    mvariAbateICMSSubstVlrCobr = CInt(colConfig.Item(FATCFG_ABATE_ICMSSUBST_VALOR_COBRANCA).sConteudo)
    mvariFaltaEstObrigaDecisao = CInt(colConfig.Item(FATCFG_FALTA_ESTOQUE_OBRIGA_TOMADA_DECISAO).sConteudo)
    mvariImpedeImpPVComBloq = CInt(colConfig.Item(FATCFG_IMPEDE_IMP_PV_COM_BLOQ).sConteudo)
    mvariVFEUsaTManual = CInt(colConfig.Item(FATCFG_VENDA_FORA_EST_USA_TALAO_MANUAL).sConteudo)
    mvariGeraFatTodasFiliais = CInt(colConfig.Item(FATCFG_GERACAO_FATURA_TODAS_FILIAIS).sConteudo)
    mvariExibeRotasCliente = CInt(colConfig.Item(FATCFG_EXIBE_ROTAS_CLIENTE).sConteudo)
    mvariPossuiIntMapLink = CInt(colConfig.Item(FATCFG_POSSUI_INTEGRACAO_MAPLINK).sConteudo)
    mvariImpDadosInterno = CInt(colConfig.Item(FATCFG_IMPORTACAO_COM_DADOS_INTERNOS).sConteudo)
    mvariPulaLinhaGrade = CInt(colConfig.Item(FATCFG_PULA_LINHA_GRADE).sConteudo)
    mvariImprimePVGravacao = CInt(colConfig.Item(FATCFG_IMPRIME_PV_GRAVACAO).sConteudo)

    mvarlCodCliNFCeSemCPF = CLng(colConfig.Item(FATCFG_CLIENTE_NFCE_SEM_CPF).sConteudo)
    
    mvarlCodCliGov = CLng(colConfig.Item(FATCFG_CLIENTE_NF_TIPO_226_GOVERNO).sConteudo)
    mvarsNatOpGov = colConfig.Item(FATCFG_NATUREZAOP_NF_TIPO_226_GOVERNO).sConteudo
    
    mvariTrazDescPVParaNF = CInt(colConfig.Item(FATCFG_TRAZ_DESCONTO_PV_PARA_NF).sConteudo)
    mvariExibeOPImpPreco = CInt(colConfig.Item(FATCFG_EXIBE_OPCAO_IMPRIMIR_OV_SEM_PRECO).sConteudo)

    mvariInativaCliDiaSemCompra = CInt(colConfig.Item(FATCFG_INATIVA_CLIENTE_DIAS_SEM_COMPRA).sConteudo)
    mvariExibeDataSaida = CInt(colConfig.Item(FATCFG_EXIBE_DATA_SAIDA_NF).sConteudo)
    mvariPedVendaCodBarra = CInt(colConfig.Item(FATCFG_PEDVENDA_CODBARRA).sConteudo)
    mvariBloqEditPrecoPV = CInt(colConfig.Item(FATCFG_BLOQUEIA_EDICAO_PRECO_PV).sConteudo)

    mvariExibeOPImpCodProd = CInt(colConfig.Item(FATCFG_EXIBE_OPCAO_IMPRIMIR_OV_COM_CODPROD).sConteudo)
    mvariExibeOPCalcST = CInt(colConfig.Item(FATCFG_EXIBE_OPCAO_OV_CALCULAR_ST).sConteudo)
    mvariCalcSTPadrao = CInt(colConfig.Item(FATCFG_CALCULAR_ST_OP_PADRAO).sConteudo)

    mvariPrimParcICMSST = CInt(colConfig.Item(FATCFG_PRIM_PARC_ICMSST).sConteudo)
    mvariPrimParcIPI = CInt(colConfig.Item(FATCFG_PRIM_PARC_IPI).sConteudo)
    mvariPrimParcPISST = CInt(colConfig.Item(FATCFG_PRIM_PARC_PISST).sConteudo)
    mvariPrimParcCOFINSST = CInt(colConfig.Item(FATCFG_PRIM_PARC_COFINSST).sConteudo)

    mvariTrataNFDevSimbPorOP = CInt(colConfig.Item(FATCFG_TRATA_NF_DEV_SIMB_BENEF_POR_OP).sConteudo)
    mvariNFDevSimbSugOP = CInt(colConfig.Item(FATCFG_NF_DEV_SIMB_BENEF_SUGERIR_OPS).sConteudo)
    mvariNFDevSimbAlmSobra = CInt(colConfig.Item(FATCFG_NF_DEV_SIMB_BENEF_ALM_SOBRA).sConteudo)
    mvariNFDevSimbUM = CInt(colConfig.Item(FATCFG_NF_DEV_SIMB_UM).sConteudo)

    mvariExibeTabPrecoItensPV = CInt(colConfig.Item(FATCFG_EXIBE_TABPRECO_ITENSPV).sConteudo)
    mvariExibeComissaoItensPV = CInt(colConfig.Item(FATCFG_EXIBE_COMISSAO_ITENSPV).sConteudo)
    mvariExibeQtdeCanceladaItensPV = CInt(colConfig.Item(FATCFG_EXIBE_QTDECANCELADA_ITENSPV).sConteudo)
    mvariAtualizaPlacaPVMapaEntrega = CInt(colConfig.Item(FATCFG_ATUALIZA_PLACAPV_MAPAENTREGA).sConteudo)
    mvariUsaComissaoPorItem = CInt(colConfig.Item(FATCFG_USA_COMISSAO_POR_ITEM).sConteudo)

    Inicializa = SUCESSO
    
    Exit Function
    
Erro_Inicializa:

    Inicializa = Err
    
    Select Case Err
    
        Case 25082
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149728)
        
    End Select
    
    Exit Function

End Function

Public Property Get iChvNFeValidaCNPJ() As Integer
    iChvNFeValidaCNPJ = mvariChvNFeValidaCNPJ
End Property

Public Property Let iChvNFeValidaCNPJ(ByVal vData As Integer)
    mvariChvNFeValidaCNPJ = vData
End Property

Public Property Get lCodCliNFCeSemCPF() As Long
    lCodCliNFCeSemCPF = mvarlCodCliNFCeSemCPF
End Property

Public Property Let lCodCliNFCeSemCPF(ByVal vData As Long)
    mvarlCodCliNFCeSemCPF = vData
End Property

Public Property Get iPRJPrazoValidVist() As Integer
    iPRJPrazoValidVist = mvariPRJPrazoValidVist
End Property

Public Property Let iPRJPrazoValidVist(ByVal vData As Integer)
    mvariPRJPrazoValidVist = vData
End Property

Public Property Get iPRJTipoPrazoValidVist() As Integer
    iPRJTipoPrazoValidVist = mvariPRJTipoPrazoValidVist
End Property

Public Property Let iPRJTipoPrazoValidVist(ByVal vData As Integer)
    mvariPRJTipoPrazoValidVist = vData
End Property

Public Property Get iPRJExibeVistorias() As Integer
    iPRJExibeVistorias = mvariPRJExibeVistorias
End Property

Public Property Let iPRJExibeVistorias(ByVal vData As Integer)
    mvariPRJExibeVistorias = vData
End Property

Public Property Get iNFFaturaApenasQuantReservPV() As Integer
    iNFFaturaApenasQuantReservPV = mvariNFFaturaApenasQuantReservPV
End Property

Public Property Let iNFFaturaApenasQuantReservPV(ByVal vData As Integer)
    mvariNFFaturaApenasQuantReservPV = vData
End Property

Public Property Get iOVBloqEditDescProd() As Integer
    iOVBloqEditDescProd = mvariOVBloqEditDescProd
End Property

Public Property Let iOVBloqEditDescProd(ByVal vData As Integer)
    mvariOVBloqEditDescProd = vData
End Property

Public Property Get iPRJObrigatorio() As Integer
    iPRJObrigatorio = mvariPRJObrigatorio
End Property

Public Property Let iPRJObrigatorio(ByVal vData As Integer)
    mvariPRJObrigatorio = vData
End Property

Public Property Get iParcSTNumDias() As Integer
    iParcSTNumDias = mvariParcSTNumDias
End Property

Public Property Let iParcSTNumDias(ByVal vData As Integer)
    mvariParcSTNumDias = vData
End Property

Public Property Get iParcRecCobraSTNovaParc() As Integer
    iParcRecCobraSTNovaParc = mvariParcRecCobraSTNovaParc
End Property

Public Property Let iParcRecCobraSTNovaParc(ByVal vData As Integer)
    mvariParcRecCobraSTNovaParc = vData
End Property

Public Property Get iZeraQtdeEmNFComplPreco() As Integer
    iZeraQtdeEmNFComplPreco = mvariZeraQtdeEmNFComplPreco
End Property

Public Property Let iZeraQtdeEmNFComplPreco(ByVal vData As Integer)
    mvariZeraQtdeEmNFComplPreco = vData
End Property

Public Property Get iExibeBtnExportaPV() As Integer
    iExibeBtnExportaPV = mvariExibeBtnExportaPV
End Property

Public Property Let iExibeBtnExportaPV(ByVal vData As Integer)
    mvariExibeBtnExportaPV = vData
End Property

Public Property Get iKitVendaExibeComp() As Integer
    iKitVendaExibeComp = mvariKitVendaExibeComp
End Property

Public Property Let iKitVendaExibeComp(ByVal vData As Integer)
    mvariKitVendaExibeComp = vData
End Property

Public Property Get iExibeTabPrecoItensPV() As Integer
    iExibeTabPrecoItensPV = mvariExibeTabPrecoItensPV
End Property

Public Property Let iExibeTabPrecoItensPV(ByVal vData As Integer)
    mvariExibeTabPrecoItensPV = vData
End Property

Public Property Get iExibeComissaoItensPV() As Integer
    iExibeComissaoItensPV = mvariExibeComissaoItensPV
End Property

Public Property Let iExibeComissaoItensPV(ByVal vData As Integer)
    mvariExibeComissaoItensPV = vData
End Property

Public Property Get iExibeQtdeCanceladaItensPV() As Integer
    iExibeQtdeCanceladaItensPV = mvariExibeQtdeCanceladaItensPV
End Property

Public Property Let iExibeQtdeCanceladaItensPV(ByVal vData As Integer)
    mvariExibeQtdeCanceladaItensPV = vData
End Property

Public Property Get iAtualizaPlacaPVMapaEntrega() As Integer
    iAtualizaPlacaPVMapaEntrega = mvariAtualizaPlacaPVMapaEntrega
End Property

Public Property Let iAtualizaPlacaPVMapaEntrega(ByVal vData As Integer)
    mvariAtualizaPlacaPVMapaEntrega = vData
End Property

Public Property Get iUsaComissaoPorItem() As Integer
    iUsaComissaoPorItem = mvariUsaComissaoPorItem
End Property

Public Property Let iUsaComissaoPorItem(ByVal vData As Integer)
    mvariUsaComissaoPorItem = vData
End Property

Public Property Get iUsaParcDePV() As Integer
    iUsaParcDePV = mvariUsaParcDePV
End Property

Public Property Let iUsaParcDePV(ByVal vData As Integer)
    mvariUsaParcDePV = vData
End Property

Public Property Get iOVAcompFiltraVend() As Integer
    iOVAcompFiltraVend = mvariOVAcompFiltraVend
End Property

Public Property Let iOVAcompFiltraVend(ByVal vData As Integer)
    mvariOVAcompFiltraVend = vData
End Property

Public Property Get iOVSemFormPreco() As Integer
    iOVSemFormPreco = mvariOVSemFormPreco
End Property

Public Property Let iOVSemFormPreco(ByVal vData As Integer)
    mvariOVSemFormPreco = vData
End Property

Public Property Get iPesosRecalcula() As Integer
    iPesosRecalcula = mvariPesosRecalcula
End Property

Public Property Let iPesosRecalcula(ByVal vData As Integer)
    mvariPesosRecalcula = vData
End Property

Public Property Get iOVAcompFiltroVendObr() As Integer
    iOVAcompFiltroVendObr = mvariOVAcompFiltroVendObr
End Property

Public Property Let iOVAcompFiltroVendObr(ByVal vData As Integer)
    mvariOVAcompFiltroVendObr = vData
End Property

Public Property Get iOVExibeOrdItens() As Integer
    iOVExibeOrdItens = mvariOVExibeOrdItens
End Property

Public Property Let iOVExibeOrdItens(ByVal vData As Integer)
    mvariOVExibeOrdItens = vData
End Property

Public Property Get iOVExibeTS() As Integer
    iOVExibeTS = mvariOVExibeTS
End Property

Public Property Let iOVExibeTS(ByVal vData As Integer)
    mvariOVExibeTS = vData
End Property

Public Property Get iCliPreenchVend() As Integer
    iCliPreenchVend = mvariCliPreenchVend
End Property

Public Property Let iCliPreenchVend(ByVal vData As Integer)
    mvariCliPreenchVend = vData
End Property

Public Property Get iCliPreenchDtUltVi() As Integer
    iCliPreenchDtUltVi = mvariCliPreenchDtUltVi
End Property

Public Property Let iCliPreenchDtUltVi(ByVal vData As Integer)
    mvariCliPreenchDtUltVi = vData
End Property

Public Property Get iCliPreenchEnd() As Integer
    iCliPreenchEnd = mvariCliPreenchEnd
End Property

Public Property Let iCliPreenchEnd(ByVal vData As Integer)
    mvariCliPreenchEnd = vData
End Property

Public Property Get iCliPreenchEndNum() As Integer
    iCliPreenchEndNum = mvariCliPreenchEndNum
End Property

Public Property Let iCliPreenchEndNum(ByVal vData As Integer)
    mvariCliPreenchEndNum = vData
End Property

Public Property Get iCliPreenchCEP() As Integer
    iCliPreenchCEP = mvariCliPreenchCEP
End Property

Public Property Let iCliPreenchCEP(ByVal vData As Integer)
    mvariCliPreenchCEP = vData
End Property

Public Property Get iCliPreenchTelefone() As Integer
    iCliPreenchTelefone = mvariCliPreenchTelefone
End Property

Public Property Let iCliPreenchTelefone(ByVal vData As Integer)
    mvariCliPreenchTelefone = vData
End Property

Public Property Get iCliPreenchEmail() As Integer
    iCliPreenchEmail = mvariCliPreenchEmail
End Property

Public Property Let iCliPreenchEmail(ByVal vData As Integer)
    mvariCliPreenchEmail = vData
End Property

Public Property Get iCliPreenchContato() As Integer
    iCliPreenchContato = mvariCliPreenchContato
End Property

Public Property Let iCliPreenchContato(ByVal vData As Integer)
    mvariCliPreenchContato = vData
End Property

Public Property Get iCliPreenchEndCompl() As Integer
    iCliPreenchEndCompl = mvariCliPreenchEndCompl
End Property

Public Property Let iCliPreenchEndCompl(ByVal vData As Integer)
    mvariCliPreenchEndCompl = vData
End Property

Public Property Get iNFDevSimbUM() As Integer
    iNFDevSimbUM = mvariNFDevSimbUM
End Property

Public Property Let iNFDevSimbUM(ByVal vData As Integer)
    mvariNFDevSimbUM = vData
End Property

Public Property Get iNFDevSimbAlmSobra() As Integer
    iNFDevSimbAlmSobra = mvariNFDevSimbAlmSobra
End Property

Public Property Let iNFDevSimbAlmSobra(ByVal vData As Integer)
    mvariNFDevSimbAlmSobra = vData
End Property

Public Property Get iTrataNFDevSimbPorOP() As Integer
    iTrataNFDevSimbPorOP = mvariTrataNFDevSimbPorOP
End Property

Public Property Let iTrataNFDevSimbPorOP(ByVal vData As Integer)
    mvariTrataNFDevSimbPorOP = vData
End Property

Public Property Get iNFDevSimbSugOP() As Integer
    iNFDevSimbSugOP = mvariNFDevSimbSugOP
End Property

Public Property Let iNFDevSimbSugOP(ByVal vData As Integer)
    mvariNFDevSimbSugOP = vData
End Property

Public Property Get iExibeOPImpCodProd() As Integer
    iExibeOPImpCodProd = mvariExibeOPImpCodProd
End Property

Public Property Let iExibeOPImpCodProd(ByVal vData As Integer)
    mvariExibeOPImpCodProd = vData
End Property

Public Property Get iExibeOPCalcST() As Integer
    iExibeOPCalcST = mvariExibeOPCalcST
End Property

Public Property Let iExibeOPCalcST(ByVal vData As Integer)
    mvariExibeOPCalcST = vData
End Property

Public Property Get iCalcSTPadrao() As Integer
    iCalcSTPadrao = mvariCalcSTPadrao
End Property

Public Property Let iCalcSTPadrao(ByVal vData As Integer)
    mvariCalcSTPadrao = vData
End Property

Public Property Get sDirImportPVSOC() As String
    sDirImportPVSOC = mvarsDirImportPVSOC
End Property

Public Property Let sDirImportPVSOC(ByVal vData As String)
    mvarsDirImportPVSOC = vData
End Property

Public Property Get iExibeDataSaida() As Integer
    iExibeDataSaida = mvariExibeDataSaida
End Property

Public Property Let iExibeDataSaida(ByVal vData As Integer)
    mvariExibeDataSaida = vData
End Property

Public Property Get iInativaCliDiaSemCompra() As Integer
    iInativaCliDiaSemCompra = mvariInativaCliDiaSemCompra
End Property

Public Property Let iInativaCliDiaSemCompra(ByVal vData As Integer)
    mvariInativaCliDiaSemCompra = vData
End Property

Public Property Get iExibeOPImpPreco() As Integer
    iExibeOPImpPreco = mvariExibeOPImpPreco
End Property

Public Property Let iExibeOPImpPreco(ByVal vData As Integer)
    mvariExibeOPImpPreco = vData
End Property

Public Property Get iTrazDescPVParaNF() As Integer
    iTrazDescPVParaNF = mvariTrazDescPVParaNF
End Property

Public Property Let iTrazDescPVParaNF(ByVal vData As Integer)
    mvariTrazDescPVParaNF = vData
End Property

Public Property Get lCodCliGov() As Long
    lCodCliGov = mvarlCodCliGov
End Property

Public Property Let lCodCliGov(ByVal vData As Long)
    mvarlCodCliGov = vData
End Property

Public Property Get sNatOpGov() As String
    sNatOpGov = mvarsNatOpGov
End Property

Public Property Let sNatOpGov(ByVal vData As String)
    mvarsNatOpGov = vData
End Property

Public Property Get iTemVersaoOV() As Integer
    iTemVersaoOV = mvariTemVersaoOV
End Property

Public Property Let iTemVersaoOV(ByVal vData As Integer)
    mvariTemVersaoOV = vData
End Property

Public Property Get iBrowseFecha() As Integer
   iBrowseFecha = mvariBrowseFecha
End Property

Public Property Let iBrowseFecha(ByVal vData As Integer)
    mvariBrowseFecha = vData
End Property

Public Property Get iTipoValidacaoPRJ() As Integer
   iTipoValidacaoPRJ = mvariTipoValidacaoPRJ
End Property

Public Property Let iTipoValidacaoPRJ(ByVal vData As Integer)
    mvariTipoValidacaoPRJ = vData
End Property

Public Property Get iCotarItensOrcVenda() As Integer
   iCotarItensOrcVenda = mvariCotarItensOrcVenda
End Property

Public Property Let iCotarItensOrcVenda(ByVal vData As Integer)
    mvariCotarItensOrcVenda = vData
End Property

Public Property Get iPVPreenchCondPagto() As Integer
    iPVPreenchCondPagto = mvariPVPreenchCondPagto
End Property

Public Property Let iPVPreenchCondPagto(ByVal vData As Integer)
    mvariPVPreenchCondPagto = vData
End Property

Public Property Get iPVPreenchDataEntrega() As Integer
    iPVPreenchDataEntrega = mvariPVPreenchDataEntrega
End Property

Public Property Let iPVPreenchDataEntrega(ByVal vData As Integer)
    mvariPVPreenchDataEntrega = vData
End Property

Public Property Get iCliPreenchCondPagto() As Integer
    iCliPreenchCondPagto = mvariCliPreenchCondPagto
End Property

Public Property Let iCliPreenchCondPagto(ByVal vData As Integer)
    mvariCliPreenchCondPagto = vData
End Property

Public Property Get iCliPreenchTransp() As Integer
    iCliPreenchTransp = mvariCliPreenchTransp
End Property

Public Property Let iCliPreenchTransp(ByVal vData As Integer)
    mvariCliPreenchTransp = vData
End Property

Public Property Get iCliPreenchCNPJ() As Integer
    iCliPreenchCNPJ = mvariCliPreenchCNPJ
End Property

Public Property Let iCliPreenchCNPJ(ByVal vData As Integer)
    mvariCliPreenchCNPJ = vData
End Property

Public Property Get iCliPreenchInscEst() As Integer
    iCliPreenchInscEst = mvariCliPreenchInscEst
End Property

Public Property Let iCliPreenchInscEst(ByVal vData As Integer)
    mvariCliPreenchInscEst = vData
End Property

Public Property Get iPedidoReservaAutomatica() As Integer
    iPedidoReservaAutomatica = mvariPedidoReservaAutomatica
End Property

Public Property Let iPedidoReservaAutomatica(ByVal vData As Integer)
    mvariPedidoReservaAutomatica = vData
End Property

Public Property Get iNFiscalAlocacaoAutomatica() As Integer
    iNFiscalAlocacaoAutomatica = mvariNFiscalAlocacaoAutomatica
End Property

Public Property Let iNFiscalAlocacaoAutomatica(ByVal vData As Integer)
    mvariNFiscalAlocacaoAutomatica = vData
End Property

Public Property Get iPedidoVendaEditaComissao() As Integer
    iPedidoVendaEditaComissao = mvariPedidoVendaEditaComissao
End Property

Public Property Let iPedidoVendaEditaComissao(ByVal vData As Integer)
    mvariPedidoVendaEditaComissao = vData
End Property

Public Property Get iNFiscalEditaComissao() As Integer
    iNFiscalEditaComissao = mvariNFiscalEditaComissao
End Property

Public Property Let iNFiscalEditaComissao(ByVal vData As Integer)
    mvariNFiscalEditaComissao = vData
End Property

Public Property Get iAglutinaLancamPorDia() As Integer
    iAglutinaLancamPorDia = mvariAglutinaLancamPorDia
End Property

Public Property Let iAglutinaLancamPorDia(ByVal vData As Integer)
    mvariAglutinaLancamPorDia = vData
End Property

Public Property Get iGeraLoteAutomatico() As Integer
    iGeraLoteAutomatico = mvariGeraLoteAutomatico
End Property

Public Property Let iGeraLoteAutomatico(ByVal vData As Integer)
    mvariGeraLoteAutomatico = vData
End Property

Public Property Get iFilialFaturamento() As Integer
    iFilialFaturamento = mvariFilialFaturamento
End Property

Public Property Let iFilialFaturamento(ByVal vData As Integer)
    mvariFilialFaturamento = vData
End Property

Public Property Get sCategMoeda() As String
    sCategMoeda = mvarsCategMoeda
End Property

Public Property Let sCategMoeda(ByVal vData As String)
    mvarsCategMoeda = vData
End Property

Public Property Get sCategCustoFixo() As String
    sCategCustoFixo = mvarsCategCustoFixo
End Property

Public Property Let sCategCustoFixo(ByVal vData As String)
    mvarsCategCustoFixo = vData
End Property

Public Property Get sFormatoPrecoUnitario() As String
    sFormatoPrecoUnitario = mvarsFormatoPrecoUnitario
End Property

Public Property Let sFormatoPrecoUnitario(ByVal vData As String)
    mvarsFormatoPrecoUnitario = vData
End Property

Public Property Get iPedidoValidaEmbalagem() As Integer
    iPedidoValidaEmbalagem = mvariPedidoValidaEmbalagem
End Property

Public Property Let iPedidoValidaEmbalagem(ByVal vData As Integer)
    mvariPedidoValidaEmbalagem = vData
End Property

Public Property Get iAbateICMSSubstVlrCobr() As Integer
    iAbateICMSSubstVlrCobr = mvariAbateICMSSubstVlrCobr
End Property

Public Property Let iAbateICMSSubstVlrCobr(ByVal vData As Integer)
    mvariAbateICMSSubstVlrCobr = vData
End Property

Public Property Get iFaltaEstObrigaDecisao() As Integer
    iFaltaEstObrigaDecisao = mvariFaltaEstObrigaDecisao
End Property

Public Property Let iFaltaEstObrigaDecisao(ByVal vData As Integer)
    mvariFaltaEstObrigaDecisao = vData
End Property

Public Property Get iImpedeImpPVComBloq() As Integer
    iImpedeImpPVComBloq = mvariImpedeImpPVComBloq
End Property

Public Property Let iImpedeImpPVComBloq(ByVal vData As Integer)
    mvariImpedeImpPVComBloq = vData
End Property

Public Property Get iTemFaixaReceb() As Integer
    iTemFaixaReceb = mvariTemFaixaReceb
End Property

Public Property Let iTemFaixaReceb(ByVal vData As Integer)
    mvariTemFaixaReceb = vData
End Property

Public Property Get iRecebForaFaixa() As Integer
    iRecebForaFaixa = mvariRecebForaFaixa
End Property

Public Property Let iRecebForaFaixa(ByVal vData As Integer)
    mvariRecebForaFaixa = vData
End Property

Public Property Get dPercentMaisReceb() As Double
    dPercentMaisReceb = mvardPercentMaisReceb
End Property

Public Property Let dPercentMaisReceb(ByVal vData As Double)
    mvardPercentMaisReceb = vData
End Property

Public Property Get dPercentMenosReceb() As Double
    dPercentMenosReceb = mvardPercentMenosReceb
End Property

Public Property Let dPercentMenosReceb(ByVal vData As Double)
    mvardPercentMenosReceb = vData
End Property

Public Property Get iExibePercFatPV() As Integer
    iExibePercFatPV = mvariExibePercFatPV
End Property

Public Property Let iExibePercFatPV(ByVal vData As Integer)
    mvariExibePercFatPV = vData
End Property

Public Property Get iVFEUsaTManual() As Integer
    iVFEUsaTManual = mvariVFEUsaTManual
End Property

Public Property Let iVFEUsaTManual(ByVal vData As Integer)
    mvariVFEUsaTManual = vData
End Property

Public Property Get iGeraFatTodasFiliais() As Integer
    iGeraFatTodasFiliais = mvariGeraFatTodasFiliais
End Property

Public Property Let iGeraFatTodasFiliais(ByVal vData As Integer)
    mvariGeraFatTodasFiliais = vData
End Property

Public Property Get iExibeRotasCliente() As Integer
    iExibeRotasCliente = mvariExibeRotasCliente
End Property

Public Property Let iExibeRotasCliente(ByVal vData As Integer)
    mvariExibeRotasCliente = vData
End Property

Public Property Get iPossuiIntMapLink() As Integer
    iPossuiIntMapLink = mvariPossuiIntMapLink
End Property

Public Property Let iPossuiIntMapLink(ByVal vData As Integer)
    mvariPossuiIntMapLink = vData
End Property

Public Property Get iImpDadosInterno() As Integer
    iImpDadosInterno = mvariImpDadosInterno
End Property

Public Property Let iImpDadosInterno(ByVal vData As Integer)
    mvariImpDadosInterno = vData
End Property

Public Property Get iPulaLinhaGrade() As Integer
    iPulaLinhaGrade = mvariPulaLinhaGrade
End Property

Public Property Let iPulaLinhaGrade(ByVal vData As Integer)
    mvariPulaLinhaGrade = vData
End Property

Public Property Get iImprimePVGravacao() As Integer
    iImprimePVGravacao = mvariImprimePVGravacao
End Property

Public Property Let iImprimePVGravacao(ByVal vData As Integer)
    mvariImprimePVGravacao = vData
End Property

Public Property Get iPedVendaCodBarra() As Integer
    iPedVendaCodBarra = mvariPedVendaCodBarra
End Property

Public Property Let iPedVendaCodBarra(ByVal vData As Integer)
    mvariPedVendaCodBarra = vData
End Property

Public Property Get iBloqEditPrecoPV() As Integer
    iBloqEditPrecoPV = mvariBloqEditPrecoPV
End Property

Public Property Let iBloqEditPrecoPV(ByVal vData As Integer)
    mvariBloqEditPrecoPV = vData
End Property

Public Property Get iPrimParcICMSST() As Integer
    iPrimParcICMSST = mvariPrimParcICMSST
End Property

Public Property Let iPrimParcICMSST(ByVal vData As Integer)
    mvariPrimParcICMSST = vData
End Property

Public Property Get iPrimParcIPI() As Integer
    iPrimParcIPI = mvariPrimParcIPI
End Property

Public Property Let iPrimParcIPI(ByVal vData As Integer)
    mvariPrimParcIPI = vData
End Property

Public Property Get iPrimParcPISST() As Integer
    iPrimParcPISST = mvariPrimParcPISST
End Property

Public Property Let iPrimParcPISST(ByVal vData As Integer)
    mvariPrimParcPISST = vData
End Property

Public Property Get iPrimParcCOFINSST() As Integer
    iPrimParcCOFINSST = mvariPrimParcCOFINSST
End Property

Public Property Let iPrimParcCOFINSST(ByVal vData As Integer)
    mvariPrimParcCOFINSST = vData
End Property

Function Gravar() As Long
'grava no bd a configuracao atual

Dim lErro As Long, colConfig As New ColFATConfig, objAux As ClassFAT

On Error GoTo Erro_Gravar
    
    Set objAux = Me 'esta linha foi incluida devido ao uso do CF
    
    'chama FATConfig_Grava_Configs
    lErro = CF("FATConfig_Grava_Configs1", objAux)
    If lErro <> SUCESSO Then Error 34513
    
    Gravar = SUCESSO
    
    Exit Function
    
Erro_Gravar:

    Gravar = Err
    
    Select Case Err
    
        Case 34513
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149729)
    
    End Select
        
    Exit Function
    
End Function

Function Gravar_Trans() As Long
'grava no bd a configuracao atual
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long, colConfig As New ColFATConfig

On Error GoTo Erro_Gravar_Trans
    
    Call Preenche_ColConfig(colConfig)
    
    'chama FATConfig_Grava_Configs
    lErro = CF("FATConfig_Grava_Configs_Trans", colConfig)
    If lErro <> SUCESSO Then Error 41819
    
    Gravar_Trans = SUCESSO
    
    Exit Function
    
Erro_Gravar_Trans:

    Gravar_Trans = Err
    
    Select Case Err
    
        Case 41819, 55019
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149730)
    
    End Select
        
    Exit Function
    
End Function

Sub Preenche_ColConfig(colConfig As ColFATConfig)

    colConfig.Add FATCFG_PVPREENCH_CONDPAGTO, EMPRESA_TODA, "", 0, CStr(mvariPVPreenchCondPagto)
    colConfig.Add FATCFG_CLIPREENCH_CONDPAGTO, EMPRESA_TODA, "", 0, CStr(mvariCliPreenchCondPagto)
    colConfig.Add FATCFG_CLIPREENCH_TRANSP, EMPRESA_TODA, "", 0, CStr(mvariCliPreenchTransp)
    colConfig.Add FATCFG_CLIPREENCH_CNPJ, EMPRESA_TODA, "", 0, CStr(mvariCliPreenchCNPJ)
    colConfig.Add FATCFG_CLIPREENCH_INSCEST, EMPRESA_TODA, "", 0, CStr(mvariCliPreenchInscEst)
    
    colConfig.Add FATCFG_PEDIDO_RESERVA_AUTOMATICA, EMPRESA_TODA, "", 0, CStr(mvariPedidoReservaAutomatica)
    colConfig.Add FATCFG_NFISCAL_ALOCA_AUTOMATICA, EMPRESA_TODA, "", 0, CStr(mvariNFiscalAlocacaoAutomatica)
    colConfig.Add FATCFG_PEDIDO_VENDA_EDITA_COMISSAO, EMPRESA_TODA, "", 0, CStr(mvariPedidoVendaEditaComissao)
    colConfig.Add FATCFG_NFISCAL_EDITA_COMISSAO, EMPRESA_TODA, "", 0, CStr(mvariNFiscalEditaComissao)
    colConfig.Add FATCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, CStr(mvariAglutinaLancamPorDia)
    colConfig.Add FATCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, CStr(mvariGeraLoteAutomatico)
    If giFilialEmpresa <> EMPRESA_TODA Then colConfig.Add FATCFG_FILIAL_FATURAMENTO, giFilialEmpresa, "", 0, CStr(mvariFilialFaturamento)
    colConfig.Add FATCFG_FORMATO_PRECO_UNITARIO, EMPRESA_TODA, "", 0, mvarsFormatoPrecoUnitario
    colConfig.Add FATCFG_CATEG_CUSTO_FIXO, EMPRESA_TODA, "", 0, mvarsCategCustoFixo
    colConfig.Add FATCFG_CATEG_MOEDA, EMPRESA_TODA, "", 0, mvarsCategMoeda
    colConfig.Add FATCFG_PEDIDO_VALIDA_EMBALAGEM, EMPRESA_TODA, "", 0, CStr(mvariPedidoValidaEmbalagem)
    colConfig.Add FATCFG_TEM_VERSAO_OV, EMPRESA_TODA, "", 0, CStr(mvariTemVersaoOV)
    colConfig.Add FATCFG_BROWSE_FECHA, EMPRESA_TODA, "", 0, CStr(mvariBrowseFecha)
    colConfig.Add FATCFG_TIPO_VALIDACAO_PRJ, EMPRESA_TODA, "", 0, CStr(mvariTipoValidacaoPRJ)
    colConfig.Add FATCFG_COTAR_ITENS_ORCVENDA, EMPRESA_TODA, "", 0, CStr(mvariCotarItensOrcVenda)

    colConfig.Add FATCFG_TEM_FAIXA_RECEB, EMPRESA_TODA, "", 0, CStr(mvariTemFaixaReceb)
    colConfig.Add FATCFG_PERCENT_MAIS_RECEB, EMPRESA_TODA, "", 0, CStr(mvardPercentMaisReceb)
    colConfig.Add FATCFG_PERCENT_MENOS_RECEB, EMPRESA_TODA, "", 0, CStr(mvardPercentMenosReceb)
    colConfig.Add FATCFG_RECEB_FORA_FAIXA, EMPRESA_TODA, "", 0, CStr(mvariRecebForaFaixa)
    colConfig.Add FATCFG_EXIBE_PERCFAT_PV, EMPRESA_TODA, "", 0, CStr(mvariExibePercFatPV)

    colConfig.Add FATCFG_ABATE_ICMSSUBST_VALOR_COBRANCA, EMPRESA_TODA, "", 0, CStr(mvariAbateICMSSubstVlrCobr)
    colConfig.Add FATCFG_FALTA_ESTOQUE_OBRIGA_TOMADA_DECISAO, EMPRESA_TODA, "", 0, CStr(mvariFaltaEstObrigaDecisao)
    colConfig.Add FATCFG_IMPEDE_IMP_PV_COM_BLOQ, EMPRESA_TODA, "", 0, CStr(mvariImpedeImpPVComBloq)

    colConfig.Add FATCFG_EXIBE_ROTAS_CLIENTE, EMPRESA_TODA, "", 0, CStr(mvariExibeRotasCliente)
    colConfig.Add FATCFG_PEDVENDA_CODBARRA, EMPRESA_TODA, "", 0, CStr(mvariPedVendaCodBarra)

    colConfig.Add FATCFG_CLIENTE_NFCE_SEM_CPF, EMPRESA_TODA, "", 0, CStr(mvarlCodCliNFCeSemCPF)

End Sub

