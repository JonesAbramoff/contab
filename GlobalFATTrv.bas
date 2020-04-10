Attribute VB_Name = "GlobalFATTrv"
Option Explicit

Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags&, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Private Parte_Codigo As String

'PUT BELOW DECLARATIONS IN A .BAS MODULE
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SetTimer& Lib "user32" (ByVal hWnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc&)
Private Declare Function KillTimer& Lib "user32" (ByVal hWnd&, ByVal nIDEvent&)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const EM_SETPASSWORDCHAR = &HCC
Public Const NV_INPUTBOX As Long = &H5000&

Private Const TAMANHO_SEGMENTO_PRODUTO = 10

Public Const TRV_INI_SEQ_BOLETO_GER_FORA_CORP = 5000000

Public Const TRV_OCRCASOTEXTO_TIPO_OBS = 0
Public Const TRV_OCRCASOTEXTO_TIPO_ASSUNTOHIST = 1
Public Const TRV_OCRCASOTEXTO_TIPO_ANOTTEXTO = 2

Public Const TRV_OCRCASOS_LIB_COBERTURA = 1
Public Const TRV_OCRCASOS_LIB_JUDICIAL = 2

Public Const TRV_OCRCASOS_PERDA_CONDENACAO = 1
Public Const TRV_OCRCASOS_PERDA_ACORDO = 2

Public Const TRV_SRV_TIPO_ASSIST = 0
Public Const TRV_SRV_TIPO_SEGURO = 1

Public Const TRV_SRV_TIPO_ASSIST_TEXTO = "Assistência"
Public Const TRV_SRV_TIPO_SEGURO_TEXTO = "Seguro"

Public Const STRING_TRV_MAXIMO = 255

Public Const TRVRELSPARAEXCEL_TIPO_PLANILHA_META = 1
Public Const TRVRELSPARAEXCEL_TIPO_PLANILHA_META_NOMEARQ = "Meta"

Public Const ROTINA_TRVGERACOMIINT_BATCH = 1

Public Const TRV_CLIENTE_FAIXA_MY_DE = 1000000
Public Const TRV_CLIENTE_FAIXA_MY_ATE = 1499999

Public Const TRV_TIPOREL_ATENDENTE = 1
Public Const TRV_TIPOREL_ESTATISTICA = 2

Public Const TRV_EMPRESA_AMBOS = 0
Public Const TRV_EMPRESA_TA = 1
Public Const TRV_EMPRESA_MY = 2

Public Const TRV_TIPO_OCR_INCIDE_BRUTO = 0
Public Const TRV_TIPO_OCR_INCIDE_CMA = 1
Public Const TRV_TIPO_OCR_INCIDE_FAT = 2

Public Const SISTEMA_INTEGRADO_SIGAV = 0
Public Const SISTEMA_INTEGRADO_KOGUT = 1

Public Const FATOR_COD_OCR_IMPORTACAO = 100000

Public Const FATOR_PAX_CLIENTE = 700000

Public Const TRV_EXPORT_VOU_TRANS_CANC_VOU = 1
Public Const TRV_EXPORT_VOU_TRANS_FATURAMENTO = 2
Public Const TRV_EXPORT_VOU_TRANS_CANC_FAT = 3
Public Const TRV_EXPORT_VOU_TRANS_PAGTO = 4
Public Const TRV_EXPORT_VOU_TRANS_CANC_PAGTO = 5
Public Const TRV_EXPORT_VOU_TRANS_ALT_COMI = 6

Public Const TRV_VOU_PAGO = 1
Public Const TRV_VOU_NAO_PAGO = 2
Public Const TRV_VOU_PAGO_E_NAO_PAGO = 3

Public Const TRV_TIPO_VALOR_BASE_LIQ = 1
Public Const TRV_TIPO_VALOR_BASE_BRU = 2
Public Const TRV_TIPO_VALOR_BASE_PER = 3

Public Const TRV_VERSAO_FATURAMENTO = 1
Public Const TRV_VERSAO_COMISSIONAMENTO = 2
Public Const TRV_VERSAO_EMISSAO = 3

Public Const DOCPARAFAT_TIPO_GERACAO_SIGAV = 1
Public Const DOCPARAFAT_TIPO_GERACAO_CORPORATOR = 2

Public Const TRV_TIPO_LIBERACAO_COMISSAO_EMISSAO = 0
Public Const TRV_TIPO_LIBERACAO_COMISSAO_BAIXA = 1
Public Const TRV_TIPO_LIBERACAO_COMISSAO_FAT = 2
Public Const TRV_TIPO_LIBERACAO_COMISSAO_FAT_TITPAG = 3
Public Const TRV_TIPO_LIBERACAO_COMISSAO_FAT_NFPAG = 4

Public Const TRV_TIPO_TRATAMENTO_COMI_VOU_NOVO = 1
Public Const TRV_TIPO_TRATAMENTO_COMI_VOU_VLR_ALTERADO = 2
Public Const TRV_TIPO_TRATAMENTO_COMI_NVL = 3
Public Const TRV_TIPO_TRATAMENTO_COMI_OCR = 4
Public Const TRV_TIPO_TRATAMENTO_COMI_VOU_TRANSF_CARTAO = 5
Public Const TRV_TIPO_TRATAMENTO_COMI_OCR_EXCLUSAO = 6
Public Const TRV_TIPO_TRATAMENTO_COMI_NVL_EXCLUSAO = 7
Public Const TRV_TIPO_TRATAMENTO_COMI_VOU_TRANSF_REP = 8
Public Const TRV_TIPO_TRATAMENTO_COMI_VOU_TRANSF_COR = 9
Public Const TRV_TIPO_TRATAMENTO_COMI_VOU_TRANSF_EMI = 10
Public Const TRV_TIPO_TRATAMENTO_COMI_VOU_TRANSF_CLI = 11
Public Const TRV_TIPO_TRATAMENTO_COMI_CMR_ALT_PERCCOMI = 12
Public Const TRV_TIPO_TRATAMENTO_COMI_CMC_ALT_PERCCOMI = 13
Public Const TRV_TIPO_TRATAMENTO_COMI_OVER_ALT_PERCCOMI = 14
Public Const TRV_TIPO_TRATAMENTO_COMI_CMA_CMCC_ALT_PERCCOMI = 15

Public Const TRV_TIPO_GRAVACAO_VOU_IMPORTACAO = 1
Public Const TRV_TIPO_GRAVACAO_VOU_EXTRACAO = 2
Public Const TRV_TIPO_GRAVACAO_VOU_TELA = 3
Public Const TRV_TIPO_GRAVACAO_ACERTO_VALOR_LIQ = 4

Public Const TRV_TIPO_APORTE_SOBREFATURA = 1
Public Const TRV_TIPO_APORTE_DIRETO = 2
Public Const TRV_TIPO_APORTE_SOBREFATURA_COND = 3
Public Const TRV_TIPO_APORTE_DIRETO_COND = 4

Public Const TRV_VOU_INFO_STATUS_LIBERADO = 1
Public Const TRV_VOU_INFO_STATUS_BLOQUEADO = 2
Public Const TRV_VOU_INFO_STATUS_ANTIGA = 3

Public Const TRV_VOU_INFO_STATUS_LIBERADO_TEXTO = "Liberado"
Public Const TRV_VOU_INFO_STATUS_BLOQUEADO_TEXTO = "Bloqueado"
Public Const TRV_VOU_INFO_STATUS_ANTIGA_TEXTO = "Antigo"

'Importação
Public Const TIPO_NOTA_FISCAL_TRV = 0
Public Const TIPO_TITULO_PAGAR_TRV = 1
Public Const TIPO_TITULO_RECEBER_TRV = 2
Public Const TIPO_CREDITOS_A_PAGAR_TRV = 3
Public Const TIPO_NF_PAGAR_TRV = 4
Public Const TIPO_MOV_ESTOQUE_TRV = 5
Public Const TIPO_OCORRENCIA_TRV = 6

Public Const TRV_CATEGORIA_FATURAMENTO = "FATURAMENTO"
Public Const TRV_CATEGORIA_FATURAMENTO_POR_VOUCHER = "POR VOUCHER"

Public Const TRV_CATEGORIA_FATURAOVER = "Fatura OVER"
Public Const TRV_CATEGORIA_FATURAOVER_MOSTRATODOS = "Mostra Todos"

Public Const TRV_CATEGORIA_NF = "NF"
Public Const TRV_CATEGORIA_NF_AO_PASSAGEIRO = "AO PASSAGEIRO"

Public Const TRV_CATEGORIA_CONDFAT = "Cond. Faturamento"
Public Const TRV_CATEGORIA_CONDFAT_SEMANALMENTE = "Semanal"
Public Const TRV_CATEGORIA_CONDFAT_PASSAGEIRO = "Passageiro"

Public Const TRV_CATEGORIA_RESPONSAVEL = "Responsável"
Public Const TRV_CATEGORIA_RESPONSAVEL_CALLCENTER = "Call Center"
Public Const TRV_CATEGORIA_RESPONSAVEL_PROMOTOR = "Promotor"

Public Const TRV_CATEGORIA_COBRANCA = "Cobrança"
Public Const TRV_CATEGORIA_COBRANCA_NORMAL = "01 - Normal"
Public Const TRV_CATEGORIA_COBRANCA_ACORDO = "02 - Acordo"
Public Const TRV_CATEGORIA_COBRANCA_SERASA = "03 - Serasa"

Public Const TRV_CATEGORIA_ENVIODECOBRANCA = "Envio Cobr. p\ email"
Public Const TRV_CATEGORIA_ENVIODECOBRANCA_SIM = "01 - Sim"
Public Const TRV_CATEGORIA_ENVIODECOBRANCA_NAO = "02 - Não"

Public Const TRV_CATEGORIA_PAGTOOCR = "Pagto OCR"
Public Const TRV_CATEGORIA_PAGTOOCR_JUNTOFAT = "Junto da Fatura"

Public Const TRV_TEMPO_ESPERA_PADRAO_MS = 50
Public Const TRV_NUMTENTARIVAS_LEITURA_VOUCHER = 10

Public Const TRV_TIPO_CLIENTE_PASSAGEIRO = 10

Public Const FATOR_SOMA_COD_EMISSOR = 500000

Public Const TRV_GERACAONF_TITULOS_BAIXADO = 1
Public Const TRV_GERACAONF_TITULOS_EMITIDOS = 2

Public Const STRING_TRV_TIPOOCR_DESCRICAO = 100

Public Const STRING_TRV_OCR_HISTORICO = 255
Public Const STRING_TRV_OCR_OBS = 255
Public Const STRING_TRV_OCR_TIPODOC = 10
Public Const STRING_TRV_OCR_TIPOVOU = 10
Public Const STRING_TRV_OCR_SERIE = 10
Public Const STRING_TRV_OCR_CODGRUPO = 10

Public Const STRING_TRV_TAMANHO_OUTROS = 100

Public Const STRING_TRVTITULOS_TIPODOC = 50

Public Const STRING_TRV_VOU_TITULAR = 100
Public Const STRING_TRV_VOU_MOEADA = 10
Public Const STRING_TRV_VOU_CONTROLE = 100
Public Const STRING_TRV_VOU_PRODUTO = 20
Public Const STRING_TRV_VOU_CIACART = 10
Public Const STRING_TRV_VOU_NUMCCRED = 20
Public Const STRING_TRV_VOU_PAXNOME = 50
Public Const STRING_TRV_VOU_DESTINO = 50
Public Const STRING_TRV_VOU_USUARIOWEB = 250
Public Const STRING_TRV_VOU_NUMAUTO = 20
Public Const STRING_TRV_VOU_VALIDADECC = 10

Public Const STRING_TRV_VOU_INFO_TIPODOC = 10
Public Const STRING_TRV_VOU_INFO_TIPVOU = 10
Public Const STRING_TRV_VOU_INFO_SERIE = 10
Public Const STRING_TRV_VOU_HISTORICO = 100

Public Const STRING_TRV_APORTE_HISTORICO = 100
Public Const STRING_TRV_APORTE_OBS = 255

Public Const STRING_TRV_ACORDO_CONTRATO = 20
Public Const STRING_TRV_ACORDO_OBS = 255
Public Const STRING_TRV_ACORDO_DESC = 255

Public Const STATUS_TRV_OCR_LIBERADO = 1
Public Const STATUS_TRV_OCR_BLOQUEADO = 2
Public Const STATUS_TRV_OCR_FATURADO = 3
Public Const STATUS_TRV_OCR_CANCELADO = 4

Public Const STATUS_TRV_OCR_LIBERADO_TEXTO = "Liberado"
Public Const STATUS_TRV_OCR_BLOQUEADO_TEXTO = "Bloqueado"
Public Const STATUS_TRV_OCR_FATURADO_TEXTO = "Faturado"
Public Const STATUS_TRV_OCR_CANCELADO_TEXTO = "Cancelado"


Public Const STATUS_TRV_VOU_ABERTO = 1
Public Const STATUS_TRV_VOU_CANCELADO = 7

Public Const STATUS_TRV_VOU_ABERTO_TEXTO = "Aberto"
Public Const STATUS_TRV_VOU_CANCELADO_TEXTO = "Cancelado"

Public Const TRV_TIPO_ESTORNO_LANC_NORMAL = 0
Public Const TRV_TIPO_ESTORNO_LANC_ESTORNADO = 1
Public Const TRV_TIPO_ESTORNO_LANC_ESTORNADOR = 2
Public Const TRV_TIPO_ESTORNO_LANC_RELANCAMENTO = 3

Public Const TRV_TIPODOC_OCR = 1
Public Const TRV_TIPODOC_NVL = 2
Public Const TRV_TIPODOC_VOU = 3
Public Const TRV_TIPODOC_CMC = 4
Public Const TRV_TIPODOC_CMR = 5
Public Const TRV_TIPODOC_OVER = 6
Public Const TRV_TIPODOC_CMCC = 7

Public Const TRV_TIPODOC_FAT_OUTROS = 1
Public Const TRV_TIPODOC_FAT_NVL = 2
Public Const TRV_TIPODOC_FAT_OCR = 3
Public Const TRV_TIPODOC_FAT_OVER = 4

Public Const TRV_CLIENTEINFO_TIPO_CLIENTE = 1
Public Const TRV_CLIENTEINFO_TIPO_FORNECEDOR = 2

Public Const TRV_TIPODOC_OCR_TEXTO = "OCR"
Public Const TRV_TIPODOC_NVL_TEXTO = "NVL"
Public Const TRV_TIPODOC_VOU_TEXTO = "VOU"
Public Const TRV_TIPODOC_CMC_TEXTO = "CMC"
Public Const TRV_TIPODOC_CMR_TEXTO = "CMR"
Public Const TRV_TIPODOC_OVER_TEXTO = "OVER"
Public Const TRV_TIPODOC_BRUTO_TEXTO = "BRUTO"
Public Const TRV_TIPODOC_CMA_TEXTO = "CMA"
Public Const TRV_TIPODOC_CMCC_TEXTO = "CMCC"

Public Const FORMAPAGTO_TRV_OCR_FAT = 1
Public Const FORMAPAGTO_TRV_OCR_CRED = 2

Public Const TRV_VALOR_MINIMO_BOLETO = 2.5

Public Const FORMAPAGTO_TRV_APORTE_TIPOPAGTO_DIRETO = 1
Public Const FORMAPAGTO_TRV_APORTE_TIPOPAGTO_COND = 2

Public Const FORMAPAGTO_TRV_OCR_FAT_TEXTO = "Faturamento"
Public Const FORMAPAGTO_TRV_OCR_CRED_TEXTO = "Crédito"

Public Const TRV_FILIAL_BH = 6

Public Const BASE_TRV_APORTE_REAL = 1
Public Const BASE_TRV_APORTE_REALACIMAPREV = 2

Public Const BASE_TRV_APORTE_REAL_TEXTO = "Realizado"
Public Const BASE_TRV_APORTE_REALACIMAPREV_TEXTO = "Realizado / Acima"

Public Const TRV_TIPO_DOC_DESTINO_CREDFORN = 1
Public Const TRV_TIPO_DOC_DESTINO_DEBCLI = 2
Public Const TRV_TIPO_DOC_DESTINO_TITREC = 3
Public Const TRV_TIPO_DOC_DESTINO_TITPAG = 4
Public Const TRV_TIPO_DOC_DESTINO_NFSPAG = 5

Public Const TRV_TIPO_DOC_DESTINO_CREDFORN_TEXTO = "Crédito"
Public Const TRV_TIPO_DOC_DESTINO_DEBCLI_TEXTO = "Débito"
Public Const TRV_TIPO_DOC_DESTINO_TITREC_TEXTO = "Título a Receber"
Public Const TRV_TIPO_DOC_DESTINO_TITPAG_TEXTO = "Título a Pagar"
Public Const TRV_TIPO_DOC_DESTINO_NFSPAG_TEXTO = "NF a Pagar"

Public Const TRV_TIPO_DOC_DESTINO_CREDFORN_TABELA = "CreditosPagForn"
Public Const TRV_TIPO_DOC_DESTINO_DEBCLI_TABELA = "DebitosRecCli"
Public Const TRV_TIPO_DOC_DESTINO_TITREC_TABELA = "TitulosRecTodos"
Public Const TRV_TIPO_DOC_DESTINO_TITPAG_TABELA = "TitulosPagTodos"
Public Const TRV_TIPO_DOC_DESTINO_NFSPAG_TABELA = "NFsPag_Todas"

Public Const TRV_TIPO_DOC_DESTINO_CREDFORN_TELA = "CreditosPagar"
Public Const TRV_TIPO_DOC_DESTINO_DEBCLI_TELA = "DebitosReceb"
Public Const TRV_TIPO_DOC_DESTINO_TITREC_TELA = "TituloReceber_Consulta"
Public Const TRV_TIPO_DOC_DESTINO_TITPAG_TELA = "TituloPagar_Consulta"
Public Const TRV_TIPO_DOC_DESTINO_NFSPAG_TELA = "NFPag_Consulta"

Public Const TRV_TIPO_DOC_DESTINO_TITREC_CLASSE = "ClassTituloReceber"
Public Const TRV_TIPO_DOC_DESTINO_TITPAG_CLASSE = "ClassTituloPagar"
Public Const TRV_TIPO_DOC_DESTINO_DEBCLI_CLASSE = "ClassDebitoRecCli"
Public Const TRV_TIPO_DOC_DESTINO_CREDFORN_CLASSE = "ClassCreditoPagar"
Public Const TRV_TIPO_DOC_DESTINO_NFSPAG_CLASSE = "ClassNFsPag"

Public Const TRVCONFIG_PROX_NUM_TITREC = "PROX_NUM_TITREC"
Public Const TRVCONFIG_PROX_NUM_TITPAG = "PROX_NUM_TITPAG"
Public Const TRVCONFIG_DIRETORIO_FAT_HTML = "DIRETORIO_FAT_HTML"
Public Const TRVCONFIG_DIRETORIO_MODELO_FAT_HTML = "DIRETORIO_MODELO_FAT_HTML"
Public Const TRVCONFIG_PROX_NUM_PASSAGEIRO = "PROX_NUM_PASSAGEIRO"
Public Const TRVCONFIG_DIRETORIO_MODELO_FAT_HTML_CARTAO = "DIRETORIO_MODELO_FAT_HTML_CARTAO"
Public Const TRVCONFIG_NUM_INT_PROX_TRVCLIEMISSORES = "NUM_INT_PROX_TRVCLIEMISSORES"
Public Const TRVCONFIG_NUM_INT_PROX_TRVCLIEMISSORESEXC = "NUM_INT_PROX_TRVCLIEMISSORESEXC"
Public Const TRVCONFIG_NUM_INT_PROX_TRVACORDOCOMISS = "NUM_INT_PROX_TRVACORDOCOMISS"
Public Const TRVCONFIG_NUM_INT_PROX_TRVACORDODIF = "NUM_INT_PROX_TRVACORDODIF"
Public Const TRVCONFIG_DATA_COMIS_CORPORATOR = "DATA_COMIS_CORPORATOR"
Public Const TRVCONFIG_VERSAO_CORPORATOR = "TRV_VERSAO"
Public Const TRVCONFIG_SISTEMA_INTEGRADO = "SISTEMA_INTEGRADO"
Public Const TRVCONFIG_CLIENTE_OVER = "CLIENTE_OVER"
Public Const TRVCONFIG_PERC_COMIS_CLI_OVER = "PERC_COMIS_CLI_OVER"
Public Const TRVCONFIG_TAR_CARTAO_NOVO_CLI_OVER = "TAR_CARTAO_NOVO_CLI_OVER"
Public Const TRVCONFIG_BANCO_PADRAO_PAGTO = "BANCO_PADRAO_PAGTO"
Public Const TRVCONFIG_BANCO_PADRAO_RECBTO = "BANCO_PADRAO_RECBTO"
Public Const TRVCONFIG_DIR_ASSISTENCIA = "DIR_IMPORT_ASSISTENCIA"
Public Const TRVCONFIG_PROX_NUM_FAV_OCRCASO = "PROX_NUM_FAV_OCRCASO"
Public Const TRVCONFIG_CLIENTE_OCR_REEMBOLSO_PADRAO = "CLIENTE_OCR_REEMBOLSO_PADRAO"
Public Const TRVCONFIG_DIR_COBERTURA = "DIR_IMPORT_COBERTURA"
Public Const TRVCONFIG_ASSISTENCIA_LIMITE_DIARIO = "ASSISTENCIA_LIMITE_DIARIO"
Public Const TRVCONFIG_ASSISTENCIA_DATA_INICIO_LIB = "ASSISTENCIA_DATA_INICIO_LIB"
Public Const TRVCONFIG_CALCULA_VALOR_ASSISTENCIA_AUTO = "CALCULA_VALOR_ASSISTENCIA_AUTO"
Public Const TRVCONFIG_VERSAO_EXPORTACAO_VOUCHER = "VERSAO_EXPORTACAO_VOUCHER"
Public Const TRVCONFIG_GERA_LOG = "GERA_LOG"
Public Const TRVCONFIG_DIR_LOG = "DIRETORIO_LOG"
Public Const TRVCONFIG_PRAZO_MIN_PARA_PAGTO = "PRAZO_MIN_PARA_PAGTO"
Public Const TRVCONFIG_PERC_FATOR_DEV_CMCC = "PERC_FATOR_DEV_CMCC"

Public Const VERSAO_EXPORTACAO_VOUCHER_1 = 1
Public Const VERSAO_EXPORTACAO_VOUCHER_2 = 2

Public Const INATIVACAO_AUTOMATICA_TEXTO = "Inativação Automática"
Public Const INATIVACAO_AUTOMATICA_CODIGO = 1

Public Const IMPORTACAO_OCORRENCIA_TEXTO = "Importação Ocr"
Public Const IMPORTACAO_OCORRENCIA_CODIGO = 2

Public Const INATIVACAO_AUTOMATICA_TIPO_NVL_TEXTO = "Inativação"
Public Const INATIVACAO_AUTOMATICA_TIPO_NVL_CODIGO = 1

Public Const INATIVACAO_AUTOMATICA_TIPO_TAR_TEXTO = "Tarifa da administradora"
Public Const INATIVACAO_AUTOMATICA_TIPO_TAR_CODIGO = 2

Public Const INATIVACAO_AUTOMATICA_TIPO_PIS_TEXTO = "PIS"
Public Const INATIVACAO_AUTOMATICA_TIPO_PIS_CODIGO = 3

Public Const INATIVACAO_AUTOMATICA_TIPO_COFINS_TEXTO = "COFINS"
Public Const INATIVACAO_AUTOMATICA_TIPO_COFINS_CODIGO = 4

Public Const INATIVACAO_AUTOMATICA_TIPO_ISS_TEXTO = "ISS"
Public Const INATIVACAO_AUTOMATICA_TIPO_ISS_CODIGO = 5

Public Const IMPORTACAO_OCORRENCIA_TIPO_IMP_TEXTO = "Importação"
Public Const IMPORTACAO_OCORRENCIA_TIPO_IMP_CODIGO = 6

Public Const VENDEDOR_CARGO_PROMOTOR = 1
Public Const VENDEDOR_CARGO_SUPERVISOR = 2
Public Const VENDEDOR_CARGO_GERENTE = 3
Public Const VENDEDOR_CARGO_DIRETOR = 4

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVOcorrencias
    lNumIntDoc As Long
    lNumVou As Long
    sTipoDoc As String
    sSerie As String
    lCodigo As Long
    lCliente As Long
    iFilialCliente As Integer
    dtDataEmissao As Date
    sObservacao As String
    iStatus As Integer
    iOrigem As Integer
    sHistorico As String
    iFormaPagto As Integer
    lNumIntDocDestino As Long
    dValorTotal As Double
    lNumDocDestino As Long
    iTipoDocDestino As Long
    dValorOCRBrutoVou As Double
    dValorOCRCMAVou As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVOcorrenciaDet
    lNumIntDoc As Long
    lNumIntDocOCR As Long
    iTipo As Integer
    dValor As Double
    iSeq As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeImpCoinfoItemFat
    lNumIntArq As Long
    lNumRegArqTexto As Long
    lcodemp As Long
    dtData As Date
    lNumVou As Long
    sMoeda As String
    dValor As Double
    scodest As String
    lNumFat As Long
    dtdataven As Date
    lnumdoc As Long
    sUsuario As String
    dtdatareg As Date
    iMarca As Integer
    sTipoDoc As String
    lrecnsica As Long
    icondpag As Integer
    sSerie As String
    sTipVou As String
    scodgru As String
    sControle As String
    iCartao As Integer
    sCodPro As String
    sNumCCred As String
    sValidCC As String
    iQuantParc As Integer
    dVlCartao As Double
    'lNumAuto As Long
    sNumAuto As String
    sCiaCart As String
    iExportado As Integer
    dtDataExp As Date
    dtDataDep As Date
    iTipoDocCorporator As Integer
    lNumIntDocCorporator As Long
    dtNoCorporatorEm As Date
    iExcluido As Integer
    iComErro As Integer
    iQtdPax As Integer
    sGrupo As String
    idiasantc As Integer
    imaster As Integer
    lnummstr As Long
    lNumIntDoc As Long
    dtDataAtualizacaoMovEst As Date
    lCodigoMovEst As Long
    dtDataAtualizacaoContab As Date
    lKit As Long
    dtDataRecepcao As Date
    lEmissor As Long
    sDestino As String
    dtdatainiciovigencia As Date
    dtdatafimvigencia As Date
    sIdioma As String
    sDestinoVou  As String
    dtarifamoeda As Double
    dCambio As Double
    dtarifareal As Double
    stitularcartao As String
    stitularcartaocpf As String
    spaxcartaofid As String
    spaxsobrenome As String
    spaxnome As String
    dtpaxdatanasc As Date
    spaxsexo As String
    spaxtipodoc As String
    spaxnumdoc As String
    spaxendereco As String
    spaxbairro As String
    spaxcidade As String
    spaxcep As String
    spaxuf As String
    spaxemail As String
    spaxcontato As String
    spaxtelefone1 As String
    spaxtelefone2 As String
    sConvenio As String
    lCodEmpVou As Long
    lNumIntTitulo As Long
    dTarifaUnitaria As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVAportes
    lNumIntDoc As Long
    lCodigo As Long
    lCliente As Long
    iFilialCliente As Integer
    dtDataEmissao As Date
    sObservacao As String
    sHistorico As String
    iTipo As Integer
    iMoeda As Integer
    dPrevValor As Double
    dtPrevDataDe As Date
    dtPrevDataAte As Date
    iProxParcela As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVAportePagtoCond
    lNumIntDoc As Long
    lNumIntDocAporte As Long
    iBase As Integer
    iStatus As Integer
    dtDataPagto As Date
    lNumIntDocDestino As Long
    iFormaPagto As Integer
    iTipoDocDestino As Integer
    dPercentual As Double
    dValor As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVAportePagtoDireto
    lNumIntDoc As Long
    lNumIntDocAporte As Long
    dValor As Double
    dtVencimento As Date
    lNumIntDocDestino As Long
    iFormaPagto As Integer
    iTipoDocDestino As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVAportePagtoFat
    lNumIntDoc As Long
    lNumIntDocAporte As Long
    dValor As Double
    dtValidadeDe As Date
    dtValidadeAte As Date
    dSaldo As Double
    dPercentual As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVAportePagtoFatCond
    lNumIntDoc As Long
    lNumIntDocAporte As Long
    dValor As Double
    dtValidadeDe As Date
    dtValidadeAte As Date
    dPercentual As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVVouchers
    lNumVou As Long
    sTipoDoc As String
    sSerie As String
    lNumIntArq As Long
    lNumRegArqTexto As Long
    lCliente As Long
    lrecnsica As Long
    dValor As Double
    dtData As Date
    iCondPagto As Integer
    sTipVou As String
    sCodGrupo As String
    lNumFatCoinfo As Long
    iTipoDocDestino As Integer
    lNumIntDocDestino As Long
    iCartao As Integer
    lNumIntDocNVL As Long
    iStatus As Integer
    sTitular As String
    sTitularCPF As String
    sProduto As String
    sMoeda As String
    sControle As String
    iPax As Integer
    dValorCambio As Double
    dCambio As Double
    sCiaCart As String
    sNumCCred As String
    sValidadeCC As String
    'lNumAuto As Long
    sNumAuto As String
    iQuantParc As Integer
    idiasantc As Integer
    iKit As Integer
    dValorOcr As Double
    iTemOcr As Integer
    lRepresentante As Long
    lCorrentista As Long
    lEmissor As Long
    dComissaoRep As Double
    dComissaoCorr As Double
    dComissaoEmissor As Double
    dComissaoAg As Double
    dValorBruto As Double
    dValorAporte As Double
    lCodigoAporte As Long
    iParcelaAporte As Integer
    lNumIntDocPagtoAporteFat As Long
    sPassageiroNome As String
    sPassageiroSobreNome As String
    lClienteVou As Long
    lCliPassageiro As Long
    lClienteComissao As Long
    lPromotor As Long
    sDestino As String
    sUsuarioCanc As String
    dValorComissao As Double
    iExtraiInfoSigav As Integer
    dtDataCanc As Date
    dHoraCanc As Double
    dTarifaUnitaria As Double
    iVigencia As Integer
    sUsuarioWeb As String
    dValorBaseComis As Double
    dValorBrutoComOCR As Double
    dValorCMAComOCR As Double
    dValorCMC As Double
    dValorCMCC As Double
    dValorCMR As Double
    dValorCME As Double
    lNumIntDoc As Long
    lNumBoleto As Long
    dtDataVencBoleto As Date
    dValorBoleto As Double
    iTrataBoleto As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVVoucherInfo
    sTipo As String
    sSerie As String
    lNumVou As Long
    dtDataEmissao As Date
    dtDataRecepcao As Date
    lCliente As Long
    lFornEmissor As Long
    sProduto As String
    sDestino As String
    dtDataInicio As Date
    dtDataTermino As Date
    sVigencia As String
    sIdioma As String
    iPax As Integer
    sDestinoVou As String
    iAntc As Integer
    sControle As String
    sConvenio As String
    dtDataPag As Date
    iCartao As Integer
    iPago As Integer
    lNumFat As Long
    lCliPassageiro As Long
    dtDataNasc As Date
    sSexo As String
    sTipoDoc As String
    sCartaoFid As String
    sMoeda As String
    dTarifaUnitaria As Double
    dCambio As Double
    sValor As String
    dTarifaPerc As Double
    dTarifaValorMoeda As Double
    dTarifaValorReal As Double
    dComissaoPerc As Double
    dComissaoValorMoeda As Double
    dComissaoValorReal As Double
    dCartaoPerc As Double
    dCartaoValorMoeda As Double
    dCartaoValorReal As Double
    dOverPerc As Double
    dOverValorMoeda As Double
    dOverValorReal As Double
    dCMRPerc As Double
    dCMRValorMoeda As Double
    dCMRValorReal As Double
    sCia As String
    sValidade As String
    sNumeroCC As String
    sTitular As String
    dValorCartao As Double
    lParcela As Long
    'lAprovacao As Long
    sAprovacao As String
    sPassageiroSobreNome As String
    sPassageiroNome As String
    sPassageiroCGC As String
    sPassageiroEndereco As String
    sPassageiroBairro As String
    sPassageiroCidade As String
    sPassageiroCEP As String
    sPassageiroUF As String
    sPassageiroEmail As String
    sPassageiroContato As String
    sPassageiroTelefone1 As String
    sPassageiroTelefone2 As String
    sGrupo As String
    spaxtipodoc As String
    sTitularCPF As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTitulosRecTRV
    lNumIntDocTitRec As Long
    dValorTarifa As Double
    lNumIntDocNFPagComi As Long
    dValorDeducoes As Double
    dValorComissao As Double
    dValorBruto As Double
End Type


'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeImpCoinfoOcor
    lNumIntArq As Long
    lNumRegArqTexto As Long
    dtData As Date
    dValor As Double
    scodest As String
    sUsuario As String
    dtdatareg As Date
    lrecnsica As Long
    sTipVou As String
    sSerie As String
    lNumVou As Long
    idc As Integer
    lcodemp As Long
    stexto1 As String
    stexto2 As String
    stexto3 As String
    stexto4 As String
    stexto5 As String
    stexto6 As String
    iliberado As Integer
    sControle As String
    lordem As Long
    lnumocorr As Long
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVTitulosExp
    lNumIntDoc As Long
    sUsuario As String
    dtData As Date
    dHora As Double
    iTipoDocOrigem As Integer
    lNumIntDocOrigem As Long
    lNumTitulo As Long
    iExcluido As Integer
    iExportado As Integer
    sUsuarioExportacao As String
    dtDataExportacao As Date
    dHoraExportacao As Double
    sNomeArq As String
    sMotivo As String
    iTemQueContabilizar As Integer
    dValorAporteCred As Double
    dValorAporte As Double
    dValorComissao As Double
    dValorCreditos As Double
    dValorDebitos As Double
    dValorTarifa As Double
    dValorDeducoes As Double
End Type

Type typeCliEmissoresExcTRV
    lNumIntDoc As Long
    lNumIntDocCliEmi As Long
    iSeq As Integer
    sProduto As String
    dPercComissao As Double
End Type

Type typeCliEmissoresTRV
    lNumIntDoc As Long
    lCliente As Long
    iSeq As Integer
    lFornEmissor As Long
    dPercComissao As Double
    dPercCI As Double
    lCargo As Long
    sNumCartao As String
    sCPF As String
End Type

Type typeFiliaisClientesTRV
    lCodCliente As Long
    iCodFilial As Integer
    lRepresentante As Long
    dPercComiRep As Double
    lCorrentista As Long
    dPercComiCorr As Double
    dPercComiAg As Double
    iConsiderarAporte As Integer
    lEmpresaPai As Long
    iFilialNF As Integer
    iFilialFat As Integer
    iFilialEmpresa As Integer
    iFilialCoinfo As Integer
    sGrupo As String
    iCondPagtoCC As Integer
    dPercFatorDevCMCC As Double
End Type

Type typeTRVAcordoTarifaDif
    lNumIntDoc As Long
    lNumIntAcordo As Long
    iSeq As Integer
    sProduto As String
    dPreco As Double
    iMoeda As Integer
    iDestino As Integer
    dPercComissao As Double
End Type

Type typeTRVAcordoComissao
    lNumIntDoc As Long
    lNumIntAcordo As Long
    iSeq As Integer
    sProduto As String
    iDestino As Integer
    dPercComissao As Double
End Type

Type typeTRVExcComissaoCli
    lCliente As Long
    iSeq As Integer
    sProduto As String
    dPercComissao As Double
End Type

Type typeTRVAcordos
    lNumIntDoc As Long
    lNumero As Long
    sContrato As String
    lCliente As Long
    iFilialCliente As Integer
    dtValidadeDe As Date
    dtValidadeAte As Date
    sObservacao As String
    sDescricao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVVoucherInfoN
    lNumIntDoc As Long
    sTipoDoc As String
    sTipVou As String
    sSerie As String
    lNumVou As Long
    iSeq As Integer
    dtData As Date
    iTipoDocDestino As Integer
    lNumIntDocDestino As Long
    lNumIntDocComiInt As Long
    dValor As Double
    sHistorico As String
    lNumTitulo As Long
    iStatus As Integer
    iTipoLiberacao As Integer
    iManual As Integer
    iTipoCliForn As Integer
    iEstorno As Integer
    lNumIntDocLiberacao As Long
    lCliForn As Long
    lNumIntDocOCR As Long
    lNumIntDocEstorno As Long
    iIndireta As Integer
    dtDataRegistro As Date
    dHoraRegistro As Double
    sUsuario As String
End Type

Type typeTRVGerComiIntDet
    lNumIntDoc As Long
    lNumIntDocGerComi As Long
    iSeq As Integer
    lNumIntDocComi As Long
    dValorBase As Double
    dValorComissao As Double
    iVendedor As Integer
    dtDataGeracao As Date
    sNomeReduzidoVendedor As String
End Type

Type typeTRVGerComiInt
    lNumIntDoc As Long
    lCodigo As Long
    dtDataGeracao As Date
    dHoraGeracao As Double
    sUsuario As String
    dtDataPagtoDe As Date
    dtDataPagtoAte As Date
    dtDataEmiDe As Date
    dtDataEmiAte As Date
    objTela As Object
End Type

Type typeTRVRelGerComiInt
    lNumIntRel As Long
    lCodigo As Long
    dtDataGeracao As Date
    dHoraGeracao As Double
    sUsuario As String
    dtDataPagtoDe As Date
    dtDataPagtoAte As Date
    dtDataEmiDe As Date
    dtDataEmiAte As Date
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVOcorrenciaAporte
    lNumIntDocOCR As Long
    iTipoPagtoAporte As Integer
    lNumIntDocPagtoAporte As Long
    lCodigoAporte As Long
    iParcelaAporte As Integer
    dValorAporte As Double
    dValorAporteUSS As Double
    iMoedaAporte As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVVoucherAporte
    sTipVou As String
    sSerie As String
    lNumVou As Long
    iTipoPagtoAporte As Integer
    lNumIntDocPagtoAporte As Long
    lCodigoAporte As Long
    iParcelaAporte As Integer
    dValorAporte As Double
    dValorAporteUSS As Double
    iMoedaAporte As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVClienteCorProd
    lCliente As Long
    iSeq As Integer
    lCorrentista As Long
    sProduto As String
    dPercComis As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVClienteRepProd
    lCliente As Long
    iSeq As Integer
    lRepresentante As Long
    sProduto As String
    dPercComis As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVTiposOcorrencia
    iCodigo As Integer
    sDescricao As String
    iEstornaAporteVou As Integer
    iConsideraComisInt As Integer
    iAlteraComiVou As Integer
    iAlteraCMCC As Integer
    iAlteraCMC As Integer
    iAlteraCMR As Integer
    iAlteraOVER As Integer
    iAlteraCMA As Integer
    iAceitaVlrPositivo As Integer
    iAceitaVlrNegativo As Integer
    iIncideSobre As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVOcrExp
    lNumIntDoc As Long
    sUsuario As String
    dtData As Date
    dHora As Double
    lNumIntDocOCR As Long
    lNumVou As Long
    sTipoDoc As String
    sSerie As String
    lCodigo As Long
    lCliente As Long
    dtDataEmissao As Date
    sObservacao As String
    iStatus As Integer
    iOrigem As Integer
    sHistorico As String
    iFormaPagto As Integer
    dValorTotal As Double
    iExcluido As Integer
    iExportado As Integer
    sUsuarioExportacao As String
    dtDataExportacao As Date
    dHoraExportacao As Double
    sNomeArq As String
    lNumFat As Long
    dtDataFat As Date
    dValorFat As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVVouExp
    lNumIntDoc As Long
    iTransacao As Integer
    sUsuario As String
    dtData As Date
    dHora As Double
    lNumVou As Long
    sTipVou As String
    sSerie As String
    lCliente As Long
    lNumFat As Long
    dtDataFat As Date
    dtDataCanc As Date
    dtDataPag As Date
    iExportado As Integer
    sUsuarioExportacao As String
    dtDataExportacao As Date
    dHoraExportacao As Double
    sNomeArq As String
    lAgenciaComissao As Long
    lEmissorComissao As Long
    lCorrentistaComissao As Long
    lRepresentanteComissao As Long
    dPercComiAg As Double
    dPercComiCor As Double
    dPercComiRep As Double
    dPercComiEmi As Double
    dValorFat As Double
    dValorBrutoComOCR As Double
    dValorCMAComOCR As Double
    dValorCMCC As Double
    dValorCMR As Double
    dValorCMC As Double
    dValorCME As Double
    dValorAporte As Double
    dValorAporteCred As Double
    dValorNeto As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVProdTarifa
    iAno As Integer
    sProduto As String
    iDiasDe As Integer
    iDiasAte As Integer
    iDiario As Integer
    dValor As Double
    dValorAdicional As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVCliDataCallCenter
    lCliente As Long
    dtDataDe As Date
    dtDataAte As Date
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRelTRVAtendCRMSel
    lNumIntRel As Long
    dtFiltroDataDe As Date
    dtFiltroDataAte As Date
    dtDataGer As Date
    dHoraGer As Double
    sUsuGer As String
    iSoCallCenter As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRelTRVAtendCRM
    lNumIntRel As Long
    iSeq As Integer
    iAtendenteCod As Integer
    sAtendenteNome As String
    lTotalCli As Long
    lTotalVou As Long
    lTotalVouCanc As Long
    lTotalPax As Long
    lTotalCliInat As Long
    lTotalCliNovos As Long
    lTotalCliReativ As Long
    lTotalCliInatReativ As Long
    lTotalCliContact As Long
    lTotalContRealizados As Long
    lTotalCliInatContact As Long
    dPercCliReativ As Double
    dPercCliContact As Double
    lMaiorCliCod As Long
    sMaiorCliNome As String
    dMaiorCliVendFat As Double
    dMaiorCliVendLiq As Double
    dMaiorCliVendBruto As Double
    sMaiorProdCod As String
    sMaiorProdDesc As String
    lMaiorProdQtd As Long
    dVendaMediaFat As Double
    dVendaMediaLiq As Double
    dVendaMediaBruto As Double
    dPercDescMedio As Double
    dTotalVendFat As Double
    dTotalVendLiq As Double
    dTotalVendBruto As Double
    dTotalInvestido As Double
    dPercCancMedio As Double
    lTotalContAtivo As Long
    lTotalContReceptivo As Long
    lTotalPDV As Long
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRelTRVDesviosVend
    lNumIntRel As Long
    lSeq As Long
    lCliente As Long
    sNomeCliente As String
    dValorFatNoMes As Double
    dValorBrutoNoMes As Double
    dValorLiqNoMes As Double
    dValorFatMesAnt As Double
    dValorBrutoMesAnt As Double
    dValorLiqMesAnt As Double
    dValorFatMesAnoAnt As Double
    dValorBrutoMesAnoAnt As Double
    dValorLiqMesAnoAnt As Double
    dTotalValorFat As Double
    dTotalValorBruto As Double
    dTotalValorLiq As Double
    dValorFatMedio As Double
    dValorBrutoMedio As Double
    dValorLiqMedio As Double
    dDesvioValorMes As Double
    dDesvioValorAno As Double
    lQtdVouNoMes As Long
    lQtdVouNoMesAnt As Long
    lQtdVouNoMesAnoAnt As Long
    dDesvioQtdMes As Double
    dDesvioQtdAno As Double
    lTotalQtdVou As Long
    dQtdVouMedio As Double
    dtPrimeiraCompra As Date
    dtUltimaCompra As Date
    sRespSetor As String
    sRespFunc As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRelTRVDesviosVendSel
    lNumIntRel As Long
    dtDataGer As Date
    dHoraGer As Double
    sUsuGer As String
    iAno As Integer
    iMes As Integer
    iDesvios As Integer
    dPercDesvMes As Double
    dPercDesvAno As Double
    dMinVendVlr As Double
    iMinVendQtd As Integer
    iTrazerCliNComp As Integer
    iValorBase As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRelTRVEstPeriodoCab
    lNumIntRel As Long
    lCQtdAg As Long
    lPQtdAg As Long
    lCQtdVou As Long
    lPQtdVou As Long
    lCQtdPax As Long
    lPQtdPax As Long
    dCValorFat As Double
    dPValorFat As Double
    dCValorLiq As Double
    dPValorLiq As Double
    dCValorBruto As Double
    dPValorBruto As Double
    dCValorInvestido As Double
    dPValorInvestido As Double
    lCCliNovos As Long
    lPCliNovos As Long
    lCCliReativ As Long
    lPCliReativ As Long
    dCPercDescMedio As Double
    dPPercDescMedio As Double
    sCMaiorProdCod As String
    sPMaiorProdCod As String
    sCMaiorProdDesc As String
    sPMaiorProdDesc As String
    lCMaiorProdQtd As Long
    lPMaiorProdQtd As Long
    dCMaiorProdVlrFat As Double
    dPMaiorProdVlrFat As Double
    dCMaiorProdVlrLiq As Double
    dPMaiorProdVlrLiq As Double
    dCMaiorProdVlrBruto As Double
    dPMaiorProdVlrBruto As Double
    lCContatosCall As Long
    lPContatosCall As Long
    lCContatosCobr As Long
    lPContatosCobr As Long
    lCContatosOutros As Long
    lPContatosOutros As Long
    lCMaiorCliCod As Long
    lPMaiorCliCod As Long
    sCMaiorCliNome As String
    sPMaiorCliNome As String
    dCMaiorCliValorFat As Double
    dPMaiorCliValorFat As Double
    dCMaiorCliValorLiq As Double
    dPMaiorCliValorLiq As Double
    dCMaiorCliValorBruto As Double
    dPMaiorCliValorBruto As Double
    dCPercVouCanc As Double
    dPPercVouCanc As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRelTRVEstPeriodoDet
    lNumIntRel As Long
    lSeq As Long
    iTipo As Integer
    sTipoNome As String
    sTextoLinha As String
    lPQtdVou As Long
    dPPercVou As Double
    lPQtdPax As Long
    dPPercPax As Double
    dPValorFat As Double
    dPPercValorFat As Double
    dPValorLiq As Double
    dPPercValorLiq As Double
    dPValorBruto As Double
    dPPercValorBruto As Double
    dPValorInvestido As Double
    dPPercDescMedio As Double
    lCQtdVou As Long
    dCPercVou As Double
    lCQtdPax As Long
    dCPercPax As Double
    dCValorFat As Double
    dCPercValorFat As Double
    dCValorLiq As Double
    dCPercValorLiq As Double
    dCValorBruto As Double
    dCPercValorBruto As Double
    dCValorInvestido As Double
    dCPercDescMedio As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRelTRVEstPeriodoSel
    lNumIntRel As Long
    dtDataGer As Date
    dHoraGer As Double
    sUsuGer As String
    dtFiltroDataDe As Date
    dtFiltroDataAte As Date
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVVouEmiCI
    sTipVou As String
    sSerie As String
    lNumVou As Long
    lFornEmissor As Long
    dPercCI As Double
    dPercReal As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVOcrCasoArqImport
    lNumIntArq As Long
    dtData As Date
    dHora As Double
    sNomeArq As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVOcrCasoImport
    lNumIntArq As Long
    lSeq As Long
    sCodigo As String
    dtData As Date
    sTCliente As String
    sNome As String
    sSobrenome As String
    sChaveVou As String
    sTipVou As String
    sSerie As String
    lNumVou As Long
    sCidadeOCR As String
    sEstadoOCR As String
    sPaisOCR As String
    sPrestador As String
    sCarater As String
    sGrauSatisfacao As String
    sTelefone As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVOcrCasos
    lNumIntDoc As Long
    sCodigo As String
    sTipVou As String
    sSerie As String
    lNumVou As Long
    spaxnome As String
    sTitularNome As String
    lClienteVou As Long
    dtDataEmissao As Date
    dtDataIda As Date
    dtDataVolta As Date
    sProduto As String
    iQtdPax As Integer
    lEndereco As Long
    dtDataAbertura As Date
    dtDataDocsRec As Date
    dtDataEnvioAnalise As Date
    lCGAnalise As Long
    lCGStatus As Long
    lCGAutorizadoPor As Long
    dValorAutorizadoTotalRS As Double
    dValorAutorizadoTotalUS As Double
    dtDataLimite As Date
    dtDataEnvioFinac As Date
    dtDataProgFinanc As Date
    dtDataPagtoPax As Date
    iJudicial As Integer
    sNumProcesso As String
    iCondenado As Integer
    dValorCondenacao As Double
    sComarca As String
    dtDataFimProcesso As Date
    dtDataPagtoCond As Date
    dValorAutorizadoSeguroRS As Double
    dValorAutorizadoSeguroUS As Double
    dValorAutorizadoAssistRS As Double
    dValorAutorizadoAssistUS As Double
    dCambio As Double
    iAnteciparPagtoSeguro As Integer
    iBanco As Integer
    sAgencia As String
    sContaCorrente As String
    sNomeFavorecido As String
    lCodFornFavorecido As Long
    sFavorecidoCGC As String
    dValorInvoicesTotal As Double
    dValorInvoicesTotalUS As Double
    dValorDespesasTotalRS As Double
    dValorDespesasTotalUS As Double
    lNumIntDocTitPagCobertura As Long
    lNumIntDocTitPagProcesso As Long
    lNumFatCobertura As Long
    lNumFatProcesso As Long
    dtDataPriEvento As Date
    dValorGastosAdvRS As Double
    dtDataIniProcesso As Date
    dProcessoDanoMaterial As Double
    dProcessoDanoMoral As Double
    iProcon As Integer
    iPerdaTipo As Integer
    sNumVouTexto As String
    dValorAutoSegRespTrvRS As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVOcrCasosSrv
    lNumIntDocOcrCaso As Long
    iSeq As Integer
    iAutorizado As Integer
    iSolicitado As Integer
    sDescricao As String
    dValorLimite As Double
    iMoeda As Integer
    iTipo As Integer
    dValorSolicitadoRS As Double
    dValorSolicitadoUS As Double
    dValorAutorizadoRS As Double
    dValorAutorizadoUS As Double
    lCodigoServ As Long
End Type


'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRvOcrCasosInvoices
    lNumIntDocOcrCaso As Long
    iSeq As Integer
    dtDataRecepcao As Date
    dtDataFatura As Date
    sNumero As String
    iMoeda As Integer
    dValorMoeda As Double
    dValorRS As Double
    sObs As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVOcrCasosOutrasFat
    lNumIntDocOcrCaso As Long
    iSeq As Integer
    iConsiderar As Integer
    dtDataRecepcao As Date
    dtDataFatura As Date
    sNumero As String
    dValorUS As Double
    dValorRS As Double
    sDescricao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVOcrCasosHist
    lNumIntDocOcrCaso As Long
    iSeq As Integer
    dtData As Date
    dHora As Double
    dtdatareg As Date
    dHoraReg As Double
    sUsuario As String
    iOrigem As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVOcrCasosTexto
    lNumIntDocOcrCaso As Long
    iTipoPai As Integer
    iSeqPai As Integer
    iSeq As Integer
    sTexto As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVOcrCasosAnotacoes
    lNumIntDocOcrCaso As Long
    iSeq As Integer
    dtData As Date
    dHora As Double
    dtdatareg As Date
    dHoraReg As Double
    sUsuario As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVOcrCasosParcCond
    lNumIntDocOcrCaso As Long
    iSeq As Integer
    dtDataVencimento As Date
    dValor As Double
    dtDataPagto As Date
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVOcrCasosPreRec
    lNumIntDocOcrCaso As Long
    dValor As Double
    dtData As Date
    lNumIntDocTitRecReembolso As Long
    lNumFatTitRecReembolso As Long
    sDescricao As String
    dtDataPagto As Date
End Type

Type typeTRVOcrCasosGastosAdv
    lNumIntDocOcrCaso As Long
    iSeq As Integer
    dtData As Date
    dValor As Double
    sDescricao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRVProdLimites
    sProduto As String
    lCodServ As Long
    sDescProd As String
    dValorLimite As Double
    iMoeda As Integer
    sDescricao1 As String
    sDescricao2 As String
    iTipo As Integer
End Type

Type typeTRVOcrCasosDocs
    lNumIntDocOcrCaso As Long
    lCodigoDoc As Long
    iSeq As Integer
    sDescricao As String
    sObservacao As String
    iRecebido As Integer
    iNecessUsu As Integer
    iNecessSist As Integer
End Type

Public Function Filial_Coinfo_Retorna_Corporator(ByVal lFilialCoinfo As Long) As Integer
'Obtém a filial no corporator com base na filial do Sigav

'São Paulo - Matriz - Código 01
'Av. Ipiranga, 345 sl Centro - 01046-923
'07.139.957/0001-62

'Filial Recife - Código 9969
'Av. Engenheiro Domingos Ferreira, 2391 sala 703 - Boa Viagem 51020-031
'07.139.957/0002-43

'Filial Curitiba - Código 06
'Rua Tibagi, 294 - cjs. 801/802
'07.139.957/0003-24

'Filial Rio de Janeiro - Código 1117
'Avenida Rio Branco, 181 Sala 2005
'07.139.957/0004-05

'Filial Porto Alegre - Código 2842
'Av. Alberto Bins, 392 Sala 1203
'07.139.957/0005-96

'Filial Belo Horizonte - Código 36
'Rua Viçosa, 43 Loja 05 - São Pedro 30330-160
'07.139.957/0006-77

    Select Case lFilialCoinfo
    
        Case 1
            Filial_Coinfo_Retorna_Corporator = 1
        Case 9969
            Filial_Coinfo_Retorna_Corporator = 2
        Case 6
            Filial_Coinfo_Retorna_Corporator = 3
        Case 1117
            Filial_Coinfo_Retorna_Corporator = 4
        Case 2842
            Filial_Coinfo_Retorna_Corporator = 5
        Case 36
            Filial_Coinfo_Retorna_Corporator = 6
    
    End Select

End Function

Public Function Filial_Corporator_Retorna_Coinfo(ByVal iFilialCorporator As Integer) As Long
'Obtém a filial no corporator com base na filial do Sigav

    Select Case iFilialCorporator
    
        Case 1
            Filial_Corporator_Retorna_Coinfo = 1
        Case 2
            Filial_Corporator_Retorna_Coinfo = 9969
        Case 3
            Filial_Corporator_Retorna_Coinfo = 6
        Case 4
            Filial_Corporator_Retorna_Coinfo = 1117
        Case 5
            Filial_Corporator_Retorna_Coinfo = 2842
        Case 6
            Filial_Corporator_Retorna_Coinfo = 36
    
    End Select

End Function

Public Function CondPagto_Corporator_Retorna_Coinfo(ByVal iCondPagtoCorporator As Integer) As Integer

    Select Case iCondPagtoCorporator
    
        Case 1
            CondPagto_Corporator_Retorna_Coinfo = 0
        Case Else
            CondPagto_Corporator_Retorna_Coinfo = iCondPagtoCorporator

    End Select

End Function

Public Function Obter_Dados_Sigav(ByVal objVoucherInfo As ClassTRVVoucherInfo, ByVal sSenha As String) As Long

Dim lErro As Long
Dim lTempoEspera As Long
Dim iIndice As Integer
Dim tVouInfo As typeTRVVoucherInfo
Dim sMensagemDeErro As String
Dim sNumVou As String

On Error GoTo Erro_Obter_Dados_Sigav

    Parte_Codigo = "000001"

    sNumVou = Format(objVoucherInfo.lNumVou, "000000000")
    
    Parte_Codigo = "000002 (Voucher:" & sNumVou & ")"
    
    For iIndice = 1 To TRV_NUMTENTARIVAS_LEITURA_VOUCHER
    
        lTempoEspera = TRV_TEMPO_ESPERA_PADRAO_MS * iIndice
        
        Parte_Codigo = "000003 (Voucher:" & sNumVou & ", Tentativa:" & CStr(iIndice) & ", Tempo de Espera:" & CStr(lTempoEspera) & "ms)"
        
        lErro = Obter_Dados_Sigav2(objVoucherInfo.sTipo, objVoucherInfo.sSerie, sNumVou, lTempoEspera, sMensagemDeErro, tVouInfo, sSenha)
        If lErro = SUCESSO Then Exit For
        
        If lErro = 194105 Then gError 192622
    
    Next

    If lErro <> SUCESSO Then gError 192622
    
    Parte_Codigo = "000004"
    
    Call TRVVoucherInfo_Type_Transfere_Obj(objVoucherInfo, tVouInfo)

    Obter_Dados_Sigav = SUCESSO

    Exit Function

Erro_Obter_Dados_Sigav:

    Obter_Dados_Sigav = gErr

    Select Case gErr
    
        Case 192622
            Call Rotina_Erro(vbOKOnly, sMensagemDeErro & "(" & Parte_Codigo & ")", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192621)

    End Select

    Exit Function
    
End Function

Private Function Envia_Comando_Putty(ByVal objShell As Object, ByVal lTempoEspera As Long, ByVal sComando As String, sMensagemDeErro As String) As Long

Dim lErro As Long

On Error GoTo Erro_Envia_Comando_Putty
    
    Call Wait(lTempoEspera)
    
    lErro = Ativa_Putty(objShell, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192624
    
    If sComando <> "" Then
        Call objShell.SendKeys(sComando, True)
    End If

    Envia_Comando_Putty = SUCESSO

    Exit Function

Erro_Envia_Comando_Putty:

    Envia_Comando_Putty = gErr

    Select Case gErr
    
        Case 192624
    
        Case Else
            sMensagemDeErro = CStr(gErr) & SEPARADOR & Error

    End Select

    Exit Function
    
End Function

Private Function Envia_Comando_Sigav(ByVal objShell As Object, ByVal lTempoEspera As Long, ByVal sComando As String, sMensagemDeErro As String) As Long

Dim lErro As Long

On Error GoTo Erro_Envia_Comando_Sigav
    
    Call Wait(lTempoEspera)
    
    lErro = Ativa_Sigav(objShell, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192624
    
    If sComando <> "" Then
        Call objShell.SendKeys(sComando, True)
    End If

    Envia_Comando_Sigav = SUCESSO

    Exit Function

Erro_Envia_Comando_Sigav:

    Envia_Comando_Sigav = gErr

    Select Case gErr
    
        Case 192624
    
        Case Else
            sMensagemDeErro = CStr(gErr) & SEPARADOR & Error

    End Select

    Exit Function
    
End Function

Private Function Obter_Dados_Sigav2(ByVal sTipo As String, ByVal sSerie As String, ByVal sNumVou As String, ByVal lTempoEsperaMs As Long, sMensagemDeErro As String, tVouInfo As typeTRVVoucherInfo, ByVal sSenha As String) As Long
'Desenvolvido por Wagner Luis

Dim lErro As Long
Dim objShell As Object
Dim sRetornoClip As String
Dim iNumRepeticoesTecla As Integer

On Error GoTo Erro_Obter_Dados_Sigav2

    Parte_Codigo = "000005"

    Set objShell = CreateObject("WScript.Shell")
    
    'Testa para ver se o Sigav Está aberto
    lErro = Ativa_Sigav(objShell, sMensagemDeErro)
    If lErro <> SUCESSO Then
        'Se não conseguiu ativar o Sigav ele deve estar fechado => abre
        lErro = Abre_Sigav(sSenha, lTempoEsperaMs, sMensagemDeErro)
        If lErro <> SUCESSO Then
            'Não consegiu abrir o Sigav => espera 1/2 hora e tenta de novo
            'For iNumRepeticoesTecla = 1 To 1800
                Call Wait(1000)
            'Next
            lErro = Abre_Sigav(sSenha, lTempoEsperaMs, sMensagemDeErro)
            If lErro <> SUCESSO Then gError 192626 'Vai fazer todas as tentativas '194105 'Mesmo assim não conseguiu nem tenta de novo
        End If
    End If
    
    Parte_Codigo = "000006"
    
    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, "^C", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    Parte_Codigo = "000007"
    
    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, "{ENTER}", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626

    Parte_Codigo = "000008"

    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, sSenha, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626

    Parte_Codigo = "000009"

'    For iNumRepeticoesTecla = 1 To 3
'
'        Parte_Codigo = "000010 (Repetição: " & CStr(iNumRepeticoesTecla) & ")"
'
'        lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, "{ENTER}", sMensagemDeErro)
'        If lErro <> SUCESSO Then gError 192626
'
'    Next

    Parte_Codigo = "000010 (Repetição: " & CStr(1) & ")"
    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, "{ENTER}", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    Parte_Codigo = "000010 (Repetição: " & CStr(2) & ")"
    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, "{ENTER}", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    Parte_Codigo = "000010 (Repetição: " & CStr(3) & ")"
    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, IIf(glEmpresa = 1, "TVA", "TVI"), sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    Parte_Codigo = "000010 (Repetição: " & CStr(3) & ")"
    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, "{ENTER}", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    lErro = Envia_Comando_Sigav(objShell, 4 * lTempoEsperaMs, " ", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626

    lErro = Envia_Comando_Sigav(objShell, 4 * lTempoEsperaMs, "+v", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    Parte_Codigo = "000011"
           
    For iNumRepeticoesTecla = 1 To 2

        Parte_Codigo = "000012 (Repetição: " & CStr(iNumRepeticoesTecla) & ")"

        lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, "{ENTER}", sMensagemDeErro)
        If lErro <> SUCESSO Then gError 192626
        
    Next
   
    Parte_Codigo = "000013"
   
    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, sTipo, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626

    Parte_Codigo = "000014"

    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, sSerie, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626

    Parte_Codigo = "000015"

    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, sNumVou, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626

    Parte_Codigo = "000016"

    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, "{ENTER}", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626

    Parte_Codigo = "000017"

    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, "{RIGHT}", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626

    Parte_Codigo = "000018"

    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, "{ENTER}", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    Parte_Codigo = "000019"
    
    lErro = Copiar_Tela(objShell, lTempoEsperaMs, sRetornoClip, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    Parte_Codigo = "000020"
    
    'Verifica se chegou na tela esperada
    lErro = Obter_Dados_Sigav_Valida_Tela(sTipo, sSerie, sNumVou, sRetornoClip, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    Parte_Codigo = "000021"
    
    lErro = Obter_Dados_Sigav_Voucher(sRetornoClip, tVouInfo, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
        
    'Se é um voucher de pago com cartão
    If tVouInfo.iCartao = MARCADO Then
        
        Parte_Codigo = "000022"
        
        'Verifica se chegou na tela esperada
        lErro = Obter_Dados_Sigav_Valida_Tela(sTipo, sSerie, sNumVou, sRetornoClip, sMensagemDeErro)
        If lErro <> SUCESSO Then gError 192626
    
        Parte_Codigo = "000023"
    
        lErro = Obter_Dados_Sigav_Cartao(sRetornoClip, tVouInfo, sMensagemDeErro)
        If lErro <> SUCESSO Then gError 192626
        
        Parte_Codigo = "000024"
        
        lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, " ", sMensagemDeErro)
        If lErro <> SUCESSO Then gError 192626
        
        Parte_Codigo = "000025"
        
        lErro = Copiar_Tela(objShell, lTempoEsperaMs, sRetornoClip, sMensagemDeErro)
        If lErro <> SUCESSO Then gError 192626
        
    End If
        
    If tVouInfo.iAntc = MARCADO Then
    
        'Não está pegando os dados de pagamentos antecipados
    
        lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, " ", sMensagemDeErro)
        If lErro <> SUCESSO Then gError 192626

        lErro = Copiar_Tela(objShell, lTempoEsperaMs, sRetornoClip, sMensagemDeErro)
        If lErro <> SUCESSO Then gError 192626

    End If
    
    Parte_Codigo = "000026"
    
    'Verifica se chegou na tela esperada
    lErro = Obter_Dados_Sigav_Valida_Tela(sTipo, sSerie, sNumVou, sRetornoClip, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    Parte_Codigo = "000027"
    
    lErro = Obter_Dados_Sigav_Passageiro(sRetornoClip, tVouInfo, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
 
    Parte_Codigo = "000028"
 
    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs, " ", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    Parte_Codigo = "000029"
    
    lErro = Copiar_Tela(objShell, lTempoEsperaMs, sRetornoClip, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    Parte_Codigo = "000030"
    
    'Verifica se chegou na tela esperada
    lErro = Obter_Dados_Sigav_Valida_Tela(sTipo, sSerie, sNumVou, sRetornoClip, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    Parte_Codigo = "000031"
    
    lErro = Obter_Dados_Sigav_Valores(sRetornoClip, tVouInfo, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
            
    Parte_Codigo = "000032"
            
    Obter_Dados_Sigav2 = SUCESSO
    
    Exit Function

Erro_Obter_Dados_Sigav2:

    Obter_Dados_Sigav2 = gErr

    Select Case gErr
    
        Case 192626, 194105

        Case Else
            sMensagemDeErro = CStr(gErr) & SEPARADOR & Error
            
    End Select

    Exit Function
       
End Function

Private Function Obter_Dados_Sigav_Valida_Tela(ByVal sTipo As String, ByVal sSerie As String, ByVal sNumVou As String, ByVal sTexto As String, sMensagemErro As String) As Long

Dim iPOS As Integer
Dim sTextoProcurado As String

On Error GoTo Erro_Obter_Dados_Sigav_Valida_Tela
        
    sTextoProcurado = "Tipo    :"
    iPOS = InStr(1, sTexto, sTextoProcurado)
    If iPOS = 0 Then gError 192627

    iPOS = iPOS + Len(sTextoProcurado) + 1
    
    If sTipo <> Mid(sTexto, iPOS, 1) Then gError 192628
    
    iPOS = iPOS + 9
    If sSerie <> Mid(sTexto, iPOS, 1) Then gError 192628

    iPOS = iPOS + 6
    If sNumVou <> Mid(sTexto, iPOS, 9) Then gError 192628
    
    Exit Function

Erro_Obter_Dados_Sigav_Valida_Tela:

    Obter_Dados_Sigav_Valida_Tela = gErr

    Select Case gErr
    
        Case 192627
            sMensagemErro = CStr(gErr) & SEPARADOR & "Voucher da tela difere do voucher procurado."

        Case 192628
            sMensagemErro = CStr(gErr) & SEPARADOR & "Era esperado que o voucher consultado estivesse na tela."

        Case Else
            sMensagemErro = CStr(gErr) & SEPARADOR & Error
            
    End Select

    Exit Function
    
End Function

Private Function Sigav_Extrai_Parte_Numerica(ByVal sTexto As String) As Long

Dim iPOS As Integer
Dim sParteTexto As String

    sTexto = Trim(sTexto)

    iPOS = InStr(1, sTexto, " ")
    
    If iPOS > 0 Then
        sParteTexto = left(sTexto, iPOS - 1)
    Else
        sParteTexto = sTexto
    End If
    
    If IsNumeric(sParteTexto) Then
        Sigav_Extrai_Parte_Numerica = StrParaLongCoinfo(sParteTexto)
    Else
        Sigav_Extrai_Parte_Numerica = 0
    End If

End Function

Private Function Sigav_Converte_Sim_Nao(ByVal sTexto As String) As Integer

    If InStr(1, sTexto, "S") <> 0 Or InStr(1, sTexto, "s") <> 0 Then
        Sigav_Converte_Sim_Nao = MARCADO
    Else
        Sigav_Converte_Sim_Nao = DESMARCADO
    End If

End Function

Private Function Obter_Dados_Sigav_Voucher(ByVal sRetornoClip As String, tVouInfo As typeTRVVoucherInfo, sMensagemErro As String) As Long
'Desenvolvido por Wagner Luis

Dim iPOS As Integer
Dim iPosLinha1 As Integer
Dim sTextoProcurado As String

On Error GoTo Erro_Obter_Dados_Sigav_Voucher
        
    Parte_Codigo = "000033"
        
    sTextoProcurado = "Tipo    :"
    iPOS = InStr(1, sRetornoClip, sTextoProcurado)

    If iPOS = 0 Then gError 192630
    
    Parte_Codigo = "000034"
    
    iPOS = iPOS + Len(sTextoProcurado) + 1
    iPosLinha1 = iPOS
    tVouInfo.sTipo = Trim(Mid(sRetornoClip, iPOS, 1))
    
    Parte_Codigo = "000035"
    
    iPOS = iPOS + 9
    tVouInfo.sSerie = Trim(Mid(sRetornoClip, iPOS, 1))

    Parte_Codigo = "000036"

    iPOS = iPOS + 6
    tVouInfo.lNumVou = StrParaLongCoinfo(Trim(Mid(sRetornoClip, iPOS, 9)))
    
    Parte_Codigo = "000037"
    
    iPOS = iPOS + 20
    tVouInfo.dtDataEmissao = StrParaDate(Mid(sRetornoClip, iPOS, 8))
    
    Parte_Codigo = "000038"
    
    iPOS = iPOS + 24
    tVouInfo.dtDataRecepcao = StrParaDate(Mid(sRetornoClip, iPOS, 8))
    
    Parte_Codigo = "000039"
    
    iPOS = iPosLinha1 + 82
    tVouInfo.lCliente = Sigav_Extrai_Parte_Numerica(Mid(sRetornoClip, iPOS, 25))

    Parte_Codigo = "000040"
    
    iPOS = iPOS + 35
    tVouInfo.lFornEmissor = Sigav_Extrai_Parte_Numerica(Mid(sRetornoClip, iPOS, 32)) + FATOR_SOMA_COD_EMISSOR
    If tVouInfo.lFornEmissor = FATOR_SOMA_COD_EMISSOR Then tVouInfo.lFornEmissor = 0
   
    Parte_Codigo = "000041"
    
    iPOS = iPosLinha1 + 2 * 82
    tVouInfo.sProduto = Trim(Mid(sRetornoClip, iPOS, 5))

    Parte_Codigo = "000042"

    iPOS = iPOS + 35
    tVouInfo.sDestino = Trim(Mid(sRetornoClip, iPOS, 32))

    Parte_Codigo = "000043"

    iPOS = iPosLinha1 + 3 * 82
    tVouInfo.dtDataInicio = StrParaDate(Mid(sRetornoClip, iPOS, 8))

    Parte_Codigo = "000044"

    iPOS = iPOS + 22
    tVouInfo.dtDataTermino = StrParaDate(Mid(sRetornoClip, iPOS, 8))
    
    Parte_Codigo = "000045"
    
    iPOS = iPOS + 19
    tVouInfo.sVigencia = Trim(Mid(sRetornoClip, iPOS, 9))
    
    Parte_Codigo = "000046"
    
    iPOS = iPOS + 18
    tVouInfo.sIdioma = Trim(Mid(sRetornoClip, iPOS, 8))
    
    Parte_Codigo = "000047"
    
    iPOS = iPosLinha1 + 4 * 82
    tVouInfo.iPax = StrParaInt(Mid(sRetornoClip, iPOS, 4))
    
    Parte_Codigo = "000048"
    
    iPOS = iPOS + 22
    tVouInfo.sDestinoVou = Trim(Mid(sRetornoClip, iPOS, 25))
    
    Parte_Codigo = "000049"
    
    iPOS = iPOS + 37
    tVouInfo.iAntc = Sigav_Converte_Sim_Nao(Mid(sRetornoClip, iPOS, 8))
    
    Parte_Codigo = "000050"
    
    iPOS = iPosLinha1 + 5 * 82
    tVouInfo.sControle = Trim(Mid(sRetornoClip, iPOS, 24))
    
    Parte_Codigo = "000051"
    
    iPOS = iPOS + 35
    tVouInfo.sConvenio = Trim(Mid(sRetornoClip, iPOS, 14))
    
    Parte_Codigo = "000052"
    
    iPOS = iPOS + 24
    tVouInfo.dtDataPag = StrParaDate(Mid(sRetornoClip, iPOS, 8))
    
    Parte_Codigo = "000053"
    
    iPOS = iPosLinha1 + 6 * 82
    tVouInfo.iCartao = Sigav_Converte_Sim_Nao(Mid(sRetornoClip, iPOS, 3))
    
    Parte_Codigo = "000054"
    
    iPOS = iPOS + 11
    tVouInfo.sGrupo = Trim(Mid(sRetornoClip, iPOS, 1))
    
    Parte_Codigo = "000055"
    
    iPOS = iPOS + 38
    tVouInfo.iPago = Sigav_Converte_Sim_Nao(Mid(sRetornoClip, iPOS, 3))
    
    Parte_Codigo = "000056"
    
    iPOS = iPOS + 13
    tVouInfo.lNumFat = StrParaLongCoinfo(Trim(Mid(sRetornoClip, iPOS, 6)))
    
    Parte_Codigo = "000057"
    
    Obter_Dados_Sigav_Voucher = SUCESSO
    
    Exit Function

Erro_Obter_Dados_Sigav_Voucher:

    Obter_Dados_Sigav_Voucher = gErr

    Select Case gErr
    
        Case 192630
            sMensagemErro = CStr(gErr) & SEPARADOR & "Era esperado que o voucher consultado estivesse na tela."

        Case Else
            sMensagemErro = CStr(gErr) & SEPARADOR & Error
            
    End Select

    Exit Function
       
End Function
    
Private Function Obter_Dados_Sigav_Passageiro(ByVal sRetornoClip As String, tVouInfo As typeTRVVoucherInfo, sMensagemDeErro As String) As Long
'Desenvolvido por Wagner Luis

Dim iPOS As Integer
Dim iPosLinha1 As Integer
Dim sTextoProcurado As String

On Error GoTo Erro_Obter_Dados_Sigav_Passageiro
    
    Parte_Codigo = "000067"

    sTextoProcurado = "Cidade:"
    iPOS = InStr(1, sRetornoClip, sTextoProcurado)
    
    If iPOS = 0 Then gError 192633
    
    sTextoProcurado = "Fone  :"
    iPOS = InStr(1, sRetornoClip, sTextoProcurado)
    
    If iPOS = 0 Then gError 192633
    
    sTextoProcurado = "Fone..."
    iPOS = InStr(1, sRetornoClip, sTextoProcurado)
    
    If iPOS = 0 Then gError 192633
    
    sTextoProcurado = "UF:"
    iPOS = InStr(1, sRetornoClip, sTextoProcurado)
    
    If iPOS = 0 Then gError 192633

    sTextoProcurado = "Sobrenome.:"
    iPOS = InStr(1, sRetornoClip, sTextoProcurado)
    
    If iPOS = 0 Then
        'MsgBox sRetornoClip
        gError 192633
    End If
    
    Parte_Codigo = "000068"
    
    iPOS = iPOS + Len(sTextoProcurado) + 1
    iPosLinha1 = iPOS
    tVouInfo.sPassageiroSobreNome = Trim(Mid(sRetornoClip, iPOS, 27))

    Parte_Codigo = "000069"

    iPOS = iPOS + 36
    tVouInfo.sPassageiroNome = Trim(Mid(sRetornoClip, iPOS, 29))
    
    Parte_Codigo = "000070"
    
    sTextoProcurado = "Data Nasc.:"
    iPOS = InStr(1, sRetornoClip, sTextoProcurado)
    
    iPOS = iPOS + Len(sTextoProcurado) + 1
   
    'iPOS = (iPosLinha1 + 1 * 82)
    If Len(Trim(Mid(sRetornoClip, iPOS, 10))) > 0 Then
        If IsDate(Trim(Mid(sRetornoClip, iPOS, 10))) = False Then
            MsgBox Trim(Mid(sRetornoClip, iPOS, 10))
            gError 192633
        End If
        tVouInfo.dtDataNasc = StrParaDate(Trim(Mid(sRetornoClip, iPOS, 10)))
    Else
        tVouInfo.dtDataNasc = DATA_NULA
    End If
    
    Parte_Codigo = "000071"
    
    iPOS = iPOS + 18
    'tVouInfo.sIdade = Mid(sRetornoClip, iPos, 9)
    
    iPOS = iPOS + 18
    tVouInfo.sSexo = Trim(Mid(sRetornoClip, iPOS, 29))
    
    Parte_Codigo = "000072"
    
    iPOS = iPosLinha1 + 2 * 82
    tVouInfo.spaxtipodoc = Trim(Mid(sRetornoClip, iPOS, 27))
    
    iPOS = iPOS + 36
    tVouInfo.sPassageiroCGC = Trim(Mid(sRetornoClip, iPOS, 27))
    
    Parte_Codigo = "000073"
    
    iPOS = iPosLinha1 + 3 * 82
    tVouInfo.sPassageiroEndereco = Trim(Mid(sRetornoClip, iPOS, 62))
    
    Parte_Codigo = "000074"
    
    iPOS = iPosLinha1 + 4 * 82
    tVouInfo.sPassageiroBairro = Trim(Mid(sRetornoClip, iPOS, 27))
    
    Parte_Codigo = "000075"
    
    iPOS = iPOS + 36
    tVouInfo.sPassageiroCidade = Trim(Mid(sRetornoClip, iPOS, 20))

    Parte_Codigo = "000076"

    iPOS = iPOS + 25
    tVouInfo.sPassageiroUF = Trim(Mid(sRetornoClip, iPOS, 2))

    Parte_Codigo = "000077"

    iPOS = iPosLinha1 + 5 * 82
    tVouInfo.sPassageiroCEP = Trim(Mid(sRetornoClip, iPOS, 27))
    
    Parte_Codigo = "000078"
    
    iPOS = iPOS + 36
    tVouInfo.sPassageiroTelefone1 = Trim(Mid(sRetornoClip, iPOS, 29))
    
    Parte_Codigo = "000079"
    
    iPOS = iPosLinha1 + 6 * 82
    tVouInfo.sPassageiroEmail = Trim(Mid(sRetornoClip, iPOS, 62))
    
    Parte_Codigo = "000080"
    
    iPOS = iPosLinha1 + 7 * 82
    tVouInfo.sCartaoFid = Trim(Mid(sRetornoClip, iPOS, 62))
    
    Parte_Codigo = "000081"
    
    iPOS = iPosLinha1 + 8 * 82
    tVouInfo.sPassageiroContato = Trim(Mid(sRetornoClip, iPOS, 62))
    
    Parte_Codigo = "000082"
    
    iPOS = iPosLinha1 + 9 * 82
    tVouInfo.sPassageiroTelefone2 = Trim(Mid(sRetornoClip, iPOS, 28))

    Parte_Codigo = "000083"

    Obter_Dados_Sigav_Passageiro = SUCESSO
    
    Exit Function

Erro_Obter_Dados_Sigav_Passageiro:

    Obter_Dados_Sigav_Passageiro = gErr

    Select Case gErr
    
        Case 192633
            sMensagemDeErro = CStr(gErr) & SEPARADOR & "Era esperado que a consulta do voucher estivesse exibindo o passageiro."

        Case Else
            sMensagemDeErro = CStr(gErr) & SEPARADOR & Error
            
    End Select

    Exit Function
       
End Function

Private Function Obter_Dados_Sigav_Valores(ByVal sRetornoClip As String, tVouInfo As typeTRVVoucherInfo, sMensagemErro As String) As Long
'Desenvolvido por Wagner Luis

Dim iPOS As Integer
Dim iPosLinha1 As Integer
Dim sTextoProcurado As String

On Error GoTo Erro_Obter_Dados_Sigav_Valores

    Parte_Codigo = "000084"

    sTextoProcurado = "Liquido TAI"
    iPOS = InStr(1, sRetornoClip, sTextoProcurado)
    
    If iPOS = 0 Then gError 192634

    sTextoProcurado = "Moeda:"
    iPOS = InStr(1, sRetornoClip, sTextoProcurado)
    
    If iPOS = 0 Then gError 192634
    
    Parte_Codigo = "000085"
    
    iPOS = iPOS + Len(sTextoProcurado) + 1
    iPosLinha1 = iPOS
    tVouInfo.sMoeda = Trim(Mid(sRetornoClip, iPOS, 3))

    Parte_Codigo = "000086"

    iPOS = iPOS + 16
    If IsNumeric(Trim(Mid(sRetornoClip, iPOS, 15))) Then
        tVouInfo.dTarifaUnitaria = StrParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 15)))
    Else
        tVouInfo.dTarifaUnitaria = 0
    End If
    
    Parte_Codigo = "000087"
    
    iPOS = iPOS + 24
    If IsNumeric(Trim(Mid(sRetornoClip, iPOS, 11))) Then
        tVouInfo.dCambio = StrParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 11)))
    Else
        tVouInfo.dCambio = 0
    End If
    Parte_Codigo = "000088"
    
    iPOS = iPOS + 19
    tVouInfo.sValor = Trim(Mid(sRetornoClip, iPOS, 11))

    Parte_Codigo = "000089"

    iPOS = iPosLinha1 + (3 * 82) + 15
    If IsNumeric(Trim(Mid(sRetornoClip, iPOS, 7))) Then
        tVouInfo.dTarifaPerc = PercentParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 7)))
    Else
        tVouInfo.dTarifaPerc = 0
    End If
    
    Parte_Codigo = "000090"
    
    iPOS = iPOS + 23
    tVouInfo.dTarifaValorMoeda = StrParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 9)))
    
    Parte_Codigo = "000091"
    
    iPOS = iPOS + 20
    tVouInfo.dTarifaValorReal = StrParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 13)))
    
    Parte_Codigo = "000092"
    
    iPOS = iPosLinha1 + (4 * 82) + 15
    tVouInfo.dComissaoPerc = PercentParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 7)))
    
    Parte_Codigo = "000093"
    
    iPOS = iPOS + 19
    tVouInfo.dComissaoValorMoeda = StrParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 13)))
    
    Parte_Codigo = "000094"
    
    iPOS = iPOS + 20
    tVouInfo.dComissaoValorReal = StrParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 13)))

    Parte_Codigo = "000095"

    iPOS = iPosLinha1 + (5 * 82) + 15
    tVouInfo.dCartaoPerc = PercentParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 7)))
    
    Parte_Codigo = "000096"
    
    iPOS = iPOS + 19
    tVouInfo.dCartaoValorMoeda = StrParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 13)))
    
    Parte_Codigo = "000097"
    
    iPOS = iPOS + 20
    tVouInfo.dCartaoValorReal = StrParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 13)))
    
    Parte_Codigo = "000098"
    
    iPOS = iPosLinha1 + (7 * 82) + 15
    tVouInfo.dOverPerc = PercentParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 7)))
    
    Parte_Codigo = "000099"
    
    iPOS = iPOS + 19
    tVouInfo.dOverValorMoeda = StrParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 13)))
    
    Parte_Codigo = "000100"
    
    iPOS = iPOS + 20
    tVouInfo.dOverValorReal = StrParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 13)))
    
    Parte_Codigo = "000101"
    
    iPOS = iPosLinha1 + (8 * 82) + 15
    tVouInfo.dCMRPerc = PercentParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 7)))

    Parte_Codigo = "000102"

    iPOS = iPOS + 19
    tVouInfo.dCMRValorMoeda = StrParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 13)))

    Parte_Codigo = "000103"

    iPOS = iPOS + 20
    tVouInfo.dCMRValorReal = StrParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 13)))

    Parte_Codigo = "000104"

    Obter_Dados_Sigav_Valores = SUCESSO
    
    Exit Function

Erro_Obter_Dados_Sigav_Valores:

    Obter_Dados_Sigav_Valores = gErr

    Select Case gErr
    
        Case 192634
            sMensagemErro = CStr(gErr) & SEPARADOR & "Era esperado que a consulta do voucher estivesse exibindo os valores."

        Case Else
            sMensagemErro = CStr(gErr) & SEPARADOR & Error
            
    End Select

    Exit Function
       
End Function

Private Function Obter_Dados_Sigav_Cartao(ByVal sRetornoClip As String, tVouInfo As typeTRVVoucherInfo, sMensagemErro As String) As Long
'Desenvolvido por Wagner Luis

Dim iPOS As Integer
Dim iPosLinha1 As Integer
Dim sTextoProcurado As String
Dim sNumero As String

On Error GoTo Erro_Obter_Dados_Sigav_Cartao

    Parte_Codigo = "000058"

    sTextoProcurado = "Numero:"
    iPOS = InStr(1, sRetornoClip, sTextoProcurado)
    
    If iPOS = 0 Then gError 192631

    sTextoProcurado = "Aprovacao:"
    iPOS = InStr(1, sRetornoClip, sTextoProcurado)
    
    If iPOS = 0 Then gError 192631

    sTextoProcurado = "Cia     :"
    iPOS = InStr(1, sRetornoClip, sTextoProcurado)
    
    If iPOS = 0 Then gError 192631
    
    Parte_Codigo = "000059"
    
    iPOS = iPOS + Len(sTextoProcurado) + 1
    iPosLinha1 = iPOS
    tVouInfo.sCia = Trim(Mid(sRetornoClip, iPOS, 15))
    
    Parte_Codigo = "000060"
    
    iPOS = iPOS + 27
    tVouInfo.sValidade = Trim(Mid(sRetornoClip, iPOS, 7))
    
    Parte_Codigo = "000061"
    
    iPOS = iPOS + 20
    tVouInfo.sNumeroCC = Trim(Mid(sRetornoClip, iPOS, 20))

    Parte_Codigo = "000062"

    iPOS = iPosLinha1 + 1 * 82
    tVouInfo.sTitular = Trim(Mid(sRetornoClip, iPOS, 30))

    Parte_Codigo = "000063"
    
    iPOS = iPOS + 50
    Call Formata_String_Numero(Trim(Mid(sRetornoClip, iPOS, 17)), sNumero)
    tVouInfo.sTitularCPF = sNumero

    iPOS = iPosLinha1 + 2 * 82
    If IsNumeric(Trim(Mid(sRetornoClip, iPOS, 15))) Then
        tVouInfo.dValorCartao = StrParaDblCoinfo(Trim(Mid(sRetornoClip, iPOS, 15)))
    Else
        tVouInfo.dValorCartao = 0
    End If

    Parte_Codigo = "000064"

    iPOS = iPOS + 26
    If IsNumeric(Trim(Mid(sRetornoClip, iPOS, 11))) Then
        tVouInfo.lParcela = StrParaLongCoinfo(Trim(Mid(sRetornoClip, iPOS, 11)))
    Else
        tVouInfo.lParcela = 0
    End If

    Parte_Codigo = "000065"

    iPOS = iPOS + 24
'    If IsNumeric(Trim(Mid(sRetornoClip, iPOS, 17))) Then
'        tVouInfo.lAprovacao = StrParaLongCoinfo(Trim(Mid(sRetornoClip, iPOS, 17)))
'    Else
        tVouInfo.sAprovacao = Trim(Mid(sRetornoClip, iPOS, 17))
'    End If

    Parte_Codigo = "000066"

    Obter_Dados_Sigav_Cartao = SUCESSO
    
    Exit Function

Erro_Obter_Dados_Sigav_Cartao:

    Obter_Dados_Sigav_Cartao = gErr

    Select Case gErr
    
        Case 192631
            sMensagemErro = CStr(gErr) & SEPARADOR & "Era esperado que a consulta do voucher estivesse exibindo os dados do cartão."

        Case Else
            sMensagemErro = CStr(gErr) & SEPARADOR & Error
            
    End Select

    Exit Function
       
End Function

Private Function Copiar_Tela(ByVal objShell As Object, ByVal lTempoEsperaMs As Long, sRetornoClip As String, sMensagemErro As String) As Long

Dim lErro As Long

On Error GoTo Erro_Copiar_Tela
    
    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs * 4, "", sMensagemErro)
    If lErro <> SUCESSO Then gError 192626
    
    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs * 2, "% ", sMensagemErro)
    If lErro <> SUCESSO Then gError 192626

    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs * 2, "o", sMensagemErro)
    If lErro <> SUCESSO Then gError 192626
    
    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs * 4, "", sMensagemErro)
    If lErro <> SUCESSO Then gError 192626
        
    sRetornoClip = ClipBoard_GetData
    
    Copiar_Tela = SUCESSO
    
    Exit Function
    
Erro_Copiar_Tela:

    Copiar_Tela = gErr
    
    Select Case gErr
    
        Case 192626
 
        Case Else
            sMensagemErro = gErr & SEPARADOR & Err.Description
            
    End Select

    Exit Function

End Function

Private Function Ativa_Sigav(ByVal objShell As Object, sMensagemDeErro As String) As Long

Dim iCount As Integer
Dim bAbriuSigav As Boolean

On Error GoTo Erro_Ativa_Sigav

    bAbriuSigav = True

    iCount = 0
    Do Until objShell.AppActivate("SIGAV - Sistema")
        iCount = iCount + 1
        If iCount > 1000 Then
            bAbriuSigav = False
            Exit Do 'gError 192625
        End If
    Loop
    
    If Not bAbriuSigav Then
        iCount = 0
        Do Until objShell.AppActivate("SIGAV - Sistema Integrado de Assistenia de Viagem")
            iCount = iCount + 1
            If iCount > 1000 Then Exit Do 'gError 192625
        Loop
        If iCount <= 1000 Then bAbriuSigav = True
    End If

    If Not bAbriuSigav Then gError 192625

    Ativa_Sigav = SUCESSO
    
    Exit Function
    
Erro_Ativa_Sigav:

    Ativa_Sigav = gErr
    
    Select Case gErr
    
        Case 192625
            sMensagemDeErro = CStr(gErr) & SEPARADOR & "Certifique-se que o Sigav esteja executando nessa máquina."
    
        Case Else
            sMensagemDeErro = CStr(gErr) & SEPARADOR & Err.Description
            
    End Select

    Exit Function

End Function

Private Function Ativa_Putty(ByVal objShell As Object, sMensagemDeErro As String) As Long

Dim iCount As Integer

On Error GoTo Erro_Ativa_Putty
    
    iCount = 0
    Do Until objShell.AppActivate("PuTTY Configuration")
        iCount = iCount + 1
        If iCount > 1000 Then gError 192625
    Loop

    Ativa_Putty = SUCESSO
    
    Exit Function
    
Erro_Ativa_Putty:

    Ativa_Putty = gErr
    
    Select Case gErr
    
        Case 192625
            sMensagemDeErro = CStr(gErr) & SEPARADOR & "Certifique-se que o Putty esteja executando nessa máquina."
    
        Case Else
            sMensagemDeErro = CStr(gErr) & SEPARADOR & Err.Description
            
    End Select

    Exit Function

End Function

Private Function Abre_Sigav(ByVal sSenha As String, ByVal lTempoEsperaMs As Long, sMensagemDeErro As String) As Long

Dim lErro As Long
Dim objShell As Object

On Error GoTo Erro_Abre_Sigav

    Set objShell = CreateObject("WScript.Shell")
    
    objShell.Run "C:\SGE\Programa\Putty.exe"
    
    lErro = Envia_Comando_Putty(objShell, lTempoEsperaMs * 20, "%e", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    lErro = Envia_Comando_Putty(objShell, lTempoEsperaMs * 20, "{TAB}", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626

    lErro = Envia_Comando_Putty(objShell, lTempoEsperaMs * 20, "{DOWN}", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626

    lErro = Envia_Comando_Putty(objShell, lTempoEsperaMs * 20, "{DOWN}", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    lErro = Envia_Comando_Putty(objShell, lTempoEsperaMs * 20, "{ENTER}", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626
    
    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs * 60, sSenha, sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626

    lErro = Envia_Comando_Sigav(objShell, lTempoEsperaMs * 20, "{ENTER}", sMensagemDeErro)
    If lErro <> SUCESSO Then gError 192626

    Abre_Sigav = SUCESSO
    
    Exit Function
    
Erro_Abre_Sigav:

    Abre_Sigav = gErr
    
    Select Case gErr
    
        Case 192626
            sMensagemDeErro = CStr(gErr) & SEPARADOR & "Certifique-se que o Putty esteja executando nessa máquina."
    
        Case Else
            sMensagemDeErro = CStr(gErr) & SEPARADOR & Err.Description
            
    End Select

    Exit Function

End Function

Function ClipBoard_GetData()

   Dim hClipMemory As Long
   Dim lpClipMemory As Long
   Dim MyString As String
   Dim RetVal As Long
   Dim sAux As String
   
   sAux = Parte_Codigo
   
   Parte_Codigo = sAux & ".01"

   If OpenClipboard(0&) = 0 Then
      MsgBox "Cannot open Clipboard. Another app. may have it open"
      Exit Function
   End If

   Parte_Codigo = sAux & ".02"

   ' Obtain the handle to the global memory
   ' block that is referencing the text.
   hClipMemory = GetClipboardData(CF_TEXT)
   If IsNull(hClipMemory) Then
      MsgBox "Could not allocate memory"
      GoTo OutOfHere
   End If

   Parte_Codigo = sAux & ".03"

   ' Lock Clipboard memory so we can reference
   ' the actual data string.
   lpClipMemory = GlobalLock(hClipMemory)

   Parte_Codigo = sAux & ".04"

   If Not IsNull(lpClipMemory) Then
      
      Parte_Codigo = sAux & ".05"
      MyString = Space$(MAXSIZE)
      Parte_Codigo = sAux & ".06"
      RetVal = lstrcpy(MyString, lpClipMemory)
      Parte_Codigo = sAux & ".07"
      RetVal = GlobalUnlock(hClipMemory)
      Parte_Codigo = sAux & ".08"

      ' Peel off the null terminating character.
      MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1)
   Else
      MsgBox "Could not lock memory to copy string from."
   End If

OutOfHere:

   Parte_Codigo = sAux & ".09"

   RetVal = CloseClipboard()
   
   Parte_Codigo = sAux & ".10"
   
   ClipBoard_GetData = MyString

End Function

Public Function TRVVoucherInfo_Type_Transfere_Obj(objTRVVoucherInfo As ClassTRVVoucherInfo, tTRVVoucherInfo As typeTRVVoucherInfo) As Long

    objTRVVoucherInfo.sTipo = tTRVVoucherInfo.sTipo
    objTRVVoucherInfo.sSerie = tTRVVoucherInfo.sSerie
    objTRVVoucherInfo.lNumVou = tTRVVoucherInfo.lNumVou
    objTRVVoucherInfo.dtDataEmissao = tTRVVoucherInfo.dtDataEmissao
    objTRVVoucherInfo.dtDataRecepcao = tTRVVoucherInfo.dtDataRecepcao
    objTRVVoucherInfo.lCliente = tTRVVoucherInfo.lCliente
    objTRVVoucherInfo.lFornEmissor = tTRVVoucherInfo.lFornEmissor
    objTRVVoucherInfo.sProduto = tTRVVoucherInfo.sProduto
    objTRVVoucherInfo.sDestino = tTRVVoucherInfo.sDestino
    objTRVVoucherInfo.dtDataInicio = tTRVVoucherInfo.dtDataInicio
    objTRVVoucherInfo.dtDataTermino = tTRVVoucherInfo.dtDataTermino
    objTRVVoucherInfo.sVigencia = tTRVVoucherInfo.sVigencia
    objTRVVoucherInfo.sIdioma = tTRVVoucherInfo.sIdioma
    objTRVVoucherInfo.iPax = tTRVVoucherInfo.iPax
    objTRVVoucherInfo.sDestinoVou = tTRVVoucherInfo.sDestinoVou
    objTRVVoucherInfo.iAntc = tTRVVoucherInfo.iAntc
    objTRVVoucherInfo.sControle = tTRVVoucherInfo.sControle
    objTRVVoucherInfo.sConvenio = tTRVVoucherInfo.sConvenio
    objTRVVoucherInfo.dtDataPag = tTRVVoucherInfo.dtDataPag
    objTRVVoucherInfo.iCartao = tTRVVoucherInfo.iCartao
    objTRVVoucherInfo.iPago = tTRVVoucherInfo.iPago
    objTRVVoucherInfo.lNumFat = tTRVVoucherInfo.lNumFat
    objTRVVoucherInfo.lCliPassageiro = tTRVVoucherInfo.lCliPassageiro
    objTRVVoucherInfo.dtDataNasc = tTRVVoucherInfo.dtDataNasc
    objTRVVoucherInfo.sSexo = tTRVVoucherInfo.sSexo
    objTRVVoucherInfo.sTipoDoc = tTRVVoucherInfo.sTipoDoc
    objTRVVoucherInfo.sCartaoFid = tTRVVoucherInfo.sCartaoFid
    objTRVVoucherInfo.sMoeda = tTRVVoucherInfo.sMoeda
    objTRVVoucherInfo.dTarifaUnitaria = tTRVVoucherInfo.dTarifaUnitaria
    objTRVVoucherInfo.dCambio = tTRVVoucherInfo.dCambio
    objTRVVoucherInfo.sValor = tTRVVoucherInfo.sValor
    objTRVVoucherInfo.dTarifaPerc = tTRVVoucherInfo.dTarifaPerc
    objTRVVoucherInfo.dTarifaValorMoeda = tTRVVoucherInfo.dTarifaValorMoeda
    objTRVVoucherInfo.dTarifaValorReal = tTRVVoucherInfo.dTarifaValorReal
    objTRVVoucherInfo.dComissaoPerc = tTRVVoucherInfo.dComissaoPerc
    objTRVVoucherInfo.dComissaoValorMoeda = tTRVVoucherInfo.dComissaoValorMoeda
    objTRVVoucherInfo.dComissaoValorReal = tTRVVoucherInfo.dComissaoValorReal
    objTRVVoucherInfo.dCartaoPerc = tTRVVoucherInfo.dCartaoPerc
    objTRVVoucherInfo.dCartaoValorMoeda = tTRVVoucherInfo.dCartaoValorMoeda
    objTRVVoucherInfo.dCartaoValorReal = tTRVVoucherInfo.dCartaoValorReal
    objTRVVoucherInfo.dOverPerc = tTRVVoucherInfo.dOverPerc
    objTRVVoucherInfo.dOverValorMoeda = tTRVVoucherInfo.dOverValorMoeda
    objTRVVoucherInfo.dOverValorReal = tTRVVoucherInfo.dOverValorReal
    objTRVVoucherInfo.dCMRPerc = tTRVVoucherInfo.dCMRPerc
    objTRVVoucherInfo.dCMRValorMoeda = tTRVVoucherInfo.dCMRValorMoeda
    objTRVVoucherInfo.dCMRValorReal = tTRVVoucherInfo.dCMRValorReal
    objTRVVoucherInfo.sCia = tTRVVoucherInfo.sCia
    objTRVVoucherInfo.sValidade = tTRVVoucherInfo.sValidade
    objTRVVoucherInfo.sNumeroCC = tTRVVoucherInfo.sNumeroCC
    objTRVVoucherInfo.sTitular = tTRVVoucherInfo.sTitular
    objTRVVoucherInfo.sTitularCPF = tTRVVoucherInfo.sTitularCPF
    objTRVVoucherInfo.dValorCartao = tTRVVoucherInfo.dValorCartao
    objTRVVoucherInfo.lParcela = tTRVVoucherInfo.lParcela
    objTRVVoucherInfo.sAprovacao = tTRVVoucherInfo.sAprovacao
    objTRVVoucherInfo.sPassageiroSobreNome = tTRVVoucherInfo.sPassageiroSobreNome
    objTRVVoucherInfo.sPassageiroNome = tTRVVoucherInfo.sPassageiroNome
    objTRVVoucherInfo.sPassageiroCGC = tTRVVoucherInfo.sPassageiroCGC
    objTRVVoucherInfo.sPassageiroEndereco = tTRVVoucherInfo.sPassageiroEndereco
    objTRVVoucherInfo.sPassageiroBairro = tTRVVoucherInfo.sPassageiroBairro
    objTRVVoucherInfo.sPassageiroCidade = tTRVVoucherInfo.sPassageiroCidade
    objTRVVoucherInfo.sPassageiroCEP = tTRVVoucherInfo.sPassageiroCEP
    objTRVVoucherInfo.sPassageiroUF = tTRVVoucherInfo.sPassageiroUF
    objTRVVoucherInfo.sPassageiroEmail = tTRVVoucherInfo.sPassageiroEmail
    objTRVVoucherInfo.sPassageiroContato = tTRVVoucherInfo.sPassageiroContato
    objTRVVoucherInfo.sPassageiroTelefone1 = tTRVVoucherInfo.sPassageiroTelefone1
    objTRVVoucherInfo.sPassageiroTelefone2 = tTRVVoucherInfo.sPassageiroTelefone2
    objTRVVoucherInfo.sGrupo = tTRVVoucherInfo.sGrupo
    objTRVVoucherInfo.sTitularCPF = tTRVVoucherInfo.sTitularCPF
    
End Function


'PUT THIS SUB IN A .BAS MODULE

Public Sub TimerProc(ByVal hWnd&, ByVal uMsg&, ByVal idEvent&, ByVal dwTime&)

    Dim EditHwnd As Long

    EditHwnd = FindWindowEx(FindWindow("#32770", "Extração de Dados"), 0, "Edit", "")

    Call SendMessage(EditHwnd, EM_SETPASSWORDCHAR, Asc("*"), 0)
    
    KillTimer hWnd, idEvent
    
End Sub

Private Function StrParaDblCoinfo(ByVal sValor As String) As Double
    
    If IsNumeric(sValor) Then
        StrParaDblCoinfo = StrParaDbl(sValor)
    Else
        StrParaDblCoinfo = 0
    End If
    
End Function

Public Function StrParaStrCoinfo(ByVal sValor As String) As String
    
    If sValor <> "?" Then
        StrParaStrCoinfo = sValor
    Else
        StrParaStrCoinfo = ""
    End If
    
End Function

Public Function PercentParaDblCoinfo(ByVal sValor As String) As Double
    
    If IsNumeric(sValor) Then
        PercentParaDblCoinfo = PercentParaDbl(sValor)
    Else
        PercentParaDblCoinfo = 0
    End If
    
End Function

Public Function StrParaLongCoinfo(ByVal sValor As String) As Long
    
    If IsNumeric(sValor) Then
        StrParaLongCoinfo = StrParaLong(sValor)
    Else
        StrParaLongCoinfo = 0
    End If
    
End Function

Function TRVVoucherInfo_Converte_Status(ByVal iStatus As Integer) As String

    Select Case iStatus
    
        Case TRV_VOU_INFO_STATUS_LIBERADO
            TRVVoucherInfo_Converte_Status = TRV_VOU_INFO_STATUS_LIBERADO_TEXTO
        Case TRV_VOU_INFO_STATUS_BLOQUEADO
            TRVVoucherInfo_Converte_Status = TRV_VOU_INFO_STATUS_BLOQUEADO_TEXTO
        Case TRV_VOU_INFO_STATUS_ANTIGA
            TRVVoucherInfo_Converte_Status = TRV_VOU_INFO_STATUS_ANTIGA_TEXTO
    End Select
    
End Function

Public Function giVersaoTRV() As Integer

Dim sConteudo As String
Dim lErro As Long

On Error GoTo Erro_giVersaoTRV

    'Le a versão do sistema
    lErro = CF("TRVConfig_Le", TRVCONFIG_VERSAO_CORPORATOR, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192369
    
    giVersaoTRV = StrParaInt(sConteudo)

    Exit Function

Erro_giVersaoTRV:

    giVersaoTRV = TRV_VERSAO_FATURAMENTO

    Select Case gErr
    
        Case 192369

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192621)

    End Select

    Exit Function
    
End Function

Public Function giBancoPadraoRecbto() As Integer

Dim sConteudo As String
Dim lErro As Long

On Error GoTo Erro_giBancoPadraoRecbto

    lErro = CF("TRVConfig_Le", TRVCONFIG_BANCO_PADRAO_RECBTO, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    giBancoPadraoRecbto = StrParaInt(sConteudo)

    Exit Function

Erro_giBancoPadraoRecbto:

    giBancoPadraoRecbto = 0

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192621)

    End Select

    Exit Function
    
End Function

Public Function giBancoPadraoPagto() As Integer

Dim sConteudo As String
Dim lErro As Long

On Error GoTo Erro_giBancoPadraoPagto

    lErro = CF("TRVConfig_Le", TRVCONFIG_BANCO_PADRAO_PAGTO, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    giBancoPadraoPagto = StrParaInt(sConteudo)

    Exit Function

Erro_giBancoPadraoPagto:

    giBancoPadraoPagto = 0

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192369)

    End Select

    Exit Function
    
End Function

Public Function gdtDataInicioComisCorp() As Date

Dim sConteudo As String
Dim lErro As Long

On Error GoTo Erro_gdtDataInicioComisCorp

    'Le a data de implantação do comissionamento pelo Corporator
    lErro = CF("TRVConfig_Le", TRVCONFIG_DATA_COMIS_CORPORATOR, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192369
    
    gdtDataInicioComisCorp = StrParaDate(sConteudo)

    Exit Function

Erro_gdtDataInicioComisCorp:

    gdtDataInicioComisCorp = DATA_NULA

    Select Case gErr
    
        Case 192369

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192621)

    End Select

    Exit Function
    
End Function

Public Function iSistemaIntegrado() As Integer

Dim sConteudo As String
Dim lErro As Long

On Error GoTo Erro_iSistemaIntegrado

    If iForcaSistemaIntegrado = DESMARCADO Then

        'Le a data de implantação do comissionamento pelo Corporator
        lErro = CF("TRVConfig_Le", TRVCONFIG_SISTEMA_INTEGRADO, EMPRESA_TODA, sConteudo)
        If lErro <> SUCESSO Then gError 192369
        
        iSistemaIntegrado = StrParaInt(sConteudo)
        
    Else
        iSistemaIntegrado = iSistemaIntegradoForcado
    End If

    Exit Function

Erro_iSistemaIntegrado:

    iSistemaIntegrado = SISTEMA_INTEGRADO_SIGAV

    Select Case gErr
    
        Case 192369

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192621)

    End Select

    Exit Function
    
End Function

Public Sub TRVConverte_Prod(sCodPro As String)
    sCodPro = Replace(sCodPro, "/", " ")
    sCodPro = Replace(sCodPro, "\", " ")
    sCodPro = Replace(sCodPro, "-", " ")
    sCodPro = Replace(sCodPro, "_", " ")
    sCodPro = Replace(sCodPro, "&", " ")
    sCodPro = left(sCodPro, TAMANHO_SEGMENTO_PRODUTO)
    sCodPro = sCodPro & String(TAMANHO_SEGMENTO_PRODUTO - Len(sCodPro), 32)
End Sub

Public Function iAssistCalcValorAuto() As Integer

Dim sConteudo As String
Dim lErro As Long

On Error GoTo Erro_iAssistCalcValorAuto

    'Le a data de implantação do comissionamento pelo Corporator
    lErro = CF("TRVConfig_Le", TRVCONFIG_CALCULA_VALOR_ASSISTENCIA_AUTO, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192369
    
    iAssistCalcValorAuto = StrParaInt(sConteudo)

    Exit Function

Erro_iAssistCalcValorAuto:

    iAssistCalcValorAuto = 0

    Select Case gErr
    
        Case 192369

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192621)

    End Select

    Exit Function
    
End Function

Public Function iVersaoVouExp() As Integer

Dim sConteudo As String
Dim lErro As Long

On Error GoTo Erro_iVersaoVouExp

    'Le a data de implantação do comissionamento pelo Corporator
    lErro = CF("TRVConfig_Le", TRVCONFIG_VERSAO_EXPORTACAO_VOUCHER, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192369
    
    iVersaoVouExp = StrParaInt(sConteudo)

    Exit Function

Erro_iVersaoVouExp:

    iVersaoVouExp = 0

    Select Case gErr
    
        Case 192369

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192621)

    End Select

    Exit Function
    
End Function
