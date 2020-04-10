Attribute VB_Name = "GlobalFATTRP"
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

Public Const TRP_TIPO_OCR_INCIDE_BRUTO = 0
Public Const TRP_TIPO_OCR_INCIDE_CMA = 1
Public Const TRP_TIPO_OCR_INCIDE_FAT = 2

Public Const ROTINA_TRPGERACOMIINT_BATCH = 1

Public Const STRING_WEB_GRUPOACESSO = 20
Public Const STRING_WEB_LOGIN = 20
Public Const STRING_WEB_SENHA = 20

Public Const TRP_USUARIO_EMISSOR = 1
Public Const TRP_USUARIO_CLIENTE = 2
Public Const TRP_USUARIO_VENDEDOR = 3


Public Const TRP_GRUPO_INDIVIDUAL = 0
Public Const TRP_GRUPO_FAMILIA = 1
Public Const TRP_GRUPO_GRUPO = 2

Public Const TRP_GRUPO_INDIVIDUAL_TEXTO = "Individual"
Public Const TRP_GRUPO_FAMILIA_TEXTO = "Família"
Public Const TRP_GRUPO_GRUPO_TEXTO = "Grupo"

Public Const TRP_USUARIO_EMISSOR_TEXTO = "Emissor"
Public Const TRP_USUARIO_CLIENTE_TEXTO = "Cliente"
Public Const TRP_USUARIO_VENDEDOR_TEXTO = "Vendedor"

Public Const FAIXA_TRP_TIPO_CLIENTE_EMISSOR_DE = 200001
Public Const FAIXA_TRP_TIPO_CLIENTE_EMISSOR_ATE = 499999

Public Const FAIXA_TRP_TIPO_CLIENTE_PASSAGEIRO_DE = 500001
Public Const FAIXA_TRP_TIPO_CLIENTE_PASSAGEIRO_ATE = 999999

Public Const FAIXA_TRP_TIPO_CLIENTE_OUTROS_DE = 1
Public Const FAIXA_TRP_TIPO_CLIENTE_OUTROS_ATE = 149999

Public Const FAIXA_TRP_TIPO_CLIENTE_FORNECEDORES_DE = 150001
Public Const FAIXA_TRP_TIPO_CLIENTE_FORNECEDORES_ATE = 199999

Public Const TRP_TIPO_CLIENTE_EMISSOR = 1
Public Const TRP_TIPO_CLIENTE_PASSAGEIRO = 2
Public Const TRP_TIPO_CLIENTE_FORNECEDORES = 3
Public Const TRP_TIPO_CLIENTE_OUTROS = 4

Public Const TRP_TIPO_CLIENTE_EMISSOR_TEXTO = "Emissores"
Public Const TRP_TIPO_CLIENTE_PASSAGEIRO_TEXTO = "Passageiros"
Public Const TRP_TIPO_CLIENTE_FORNECEDORES_TEXTO = "Fornecedores"

Public Const TRP_EXPORT_VOU_TRANS_CANC_VOU = 1
Public Const TRP_EXPORT_VOU_TRANS_FATURAMENTO = 2
Public Const TRP_EXPORT_VOU_TRANS_CANC_FAT = 3
Public Const TRP_EXPORT_VOU_TRANS_PAGTO = 4
Public Const TRP_EXPORT_VOU_TRANS_CANC_PAGTO = 5
Public Const TRP_EXPORT_VOU_TRANS_ALT_COMI = 6

Public Const TRP_VOU_PAGO = 1
Public Const TRP_VOU_NAO_PAGO = 2
Public Const TRP_VOU_PAGO_E_NAO_PAGO = 3

Public Const TRP_TIPO_VALOR_BASE_LIQ = 1
Public Const TRP_TIPO_VALOR_BASE_BRU = 2
Public Const TRP_TIPO_VALOR_BASE_PER = 3

Public Const TRP_VERSAO_FATURAMENTO = 1
Public Const TRP_VERSAO_COMISSIONAMENTO = 2
Public Const TRP_VERSAO_EMISSAO = 3

Public Const TRP_TIPO_LIBERACAO_COMISSAO_EMI = 0
Public Const TRP_TIPO_LIBERACAO_COMISSAO_BAIXA = 1
Public Const TRP_TIPO_LIBERACAO_COMISSAO_FAT = 2

Public Const TRP_TIPO_TRATAMENTO_COMI_VOU_NOVO = 1
Public Const TRP_TIPO_TRATAMENTO_COMI_VOU_VLR_ALTERADO = 2
Public Const TRP_TIPO_TRATAMENTO_COMI_NVL = 3
Public Const TRP_TIPO_TRATAMENTO_COMI_OCR = 4
Public Const TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_CARTAO = 5
Public Const TRP_TIPO_TRATAMENTO_COMI_OCR_EXCLUSAO = 6
Public Const TRP_TIPO_TRATAMENTO_COMI_NVL_EXCLUSAO = 7
Public Const TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_REP = 8
Public Const TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_COR = 9
Public Const TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_EMI = 10
Public Const TRP_TIPO_TRATAMENTO_COMI_VOU_TRANSF_CLI = 11
Public Const TRP_TIPO_TRATAMENTO_COMI_CMR_ALT_PERCCOMI = 12
Public Const TRP_TIPO_TRATAMENTO_COMI_CMC_ALT_PERCCOMI = 13
Public Const TRP_TIPO_TRATAMENTO_COMI_OVER_ALT_PERCCOMI = 14
Public Const TRP_TIPO_TRATAMENTO_COMI_CMA_CMCC_ALT_PERCCOMI = 15

Public Const TRP_TIPO_GRAVACAO_VOU_EMISSAO = 1
Public Const TRP_TIPO_GRAVACAO_VOU_MANUTENCAO = 2
Public Const TRP_TIPO_GRAVACAO_VOU_COMISSAO = 3
Public Const TRP_TIPO_GRAVACAO_ACERTO_VALOR_LIQ = 4 'Quando mexe na comissão da agência

Public Const TRP_TIPO_APORTE_SOBREFATURA = 1
Public Const TRP_TIPO_APORTE_DIRETO = 2
Public Const TRP_TIPO_APORTE_SOBREFATURA_COND = 3
Public Const TRP_TIPO_APORTE_DIRETO_COND = 4

Public Const TRP_VOU_INFO_STATUS_LIBERADO = 1
Public Const TRP_VOU_INFO_STATUS_BLOQUEADO = 2
Public Const TRP_VOU_INFO_STATUS_ANTIGA = 3

Public Const TRP_VOU_INFO_STATUS_LIBERADO_TEXTO = "Liberado"
Public Const TRP_VOU_INFO_STATUS_BLOQUEADO_TEXTO = "Bloqueado"
Public Const TRP_VOU_INFO_STATUS_ANTIGA_TEXTO = "Antigo"

'Importação
Public Const TIPO_NOTA_FISCAL_TRP = 0
Public Const TIPO_TITULO_PAGAR_TRP = 1
Public Const TIPO_TITULO_RECEBER_TRP = 2
Public Const TIPO_CREDITOS_A_PAGAR_TRP = 3
Public Const TIPO_NF_PAGAR_TRP = 4
Public Const TIPO_MOV_ESTOQUE_TRP = 5
Public Const TIPO_OCORRENCIA_TRP = 6

Public Const TRP_CATEGORIA_FATURAMENTO = "FATURAMENTO"
Public Const TRP_CATEGORIA_FATURAMENTO_POR_VOUCHER = "POR VOUCHER"

Public Const TRP_CATEGORIA_NF = "NF"
Public Const TRP_CATEGORIA_NF_AO_PASSAGEIRO = "AO PASSAGEIRO"

Public Const TRP_CATEGORIA_LIBCME = "Liberação CME"
Public Const TRP_CATEGORIA_LIBCME_FAT = "Faturamento"
Public Const TRP_CATEGORIA_LIBCME_PAG = "Pagamento"
Public Const TRP_CATEGORIA_LIBCME_EMI = "Emissão"

Public Const TRP_CATEGORIA_CONDFAT = "Cond. Faturamento"
Public Const TRP_CATEGORIA_CONDFAT_SEMANALMENTE = "Semanal"
Public Const TRP_CATEGORIA_CONDFAT_PASSAGEIRO = "Passageiro"

Public Const TRP_CATEGORIA_RESPONSAVEL = "Responsável"
Public Const TRP_CATEGORIA_RESPONSAVEL_CALLCENTER = "Call Center"
Public Const TRP_CATEGORIA_RESPONSAVEL_PROMOTOR = "Promotor"

Public Const TRP_CATEGORIA_COBRANCA = "Cobrança"
Public Const TRP_CATEGORIA_COBRANCA_NORMAL = "01 - Normal"
Public Const TRP_CATEGORIA_COBRANCA_ACORDO = "02 - Acordo"
Public Const TRP_CATEGORIA_COBRANCA_SERASA = "03 - Serasa"

Public Const TRP_CATEGORIA_ENVIODECOBRANCA = "Envio Cobr. p\ email"
Public Const TRP_CATEGORIA_ENVIODECOBRANCA_SIM = "01 - Sim"
Public Const TRP_CATEGORIA_ENVIODECOBRANCA_NAO = "02 - Não"

Public Const TRP_CATEGORIA_PAGTOOCR = "Pagto OCR"
Public Const TRP_CATEGORIA_PAGTOOCR_JUNTOFAT = "Junto da Fatura"

Public Const TRP_GERACAONF_TITULOS_BAIXADO = 1
Public Const TRP_GERACAONF_TITULOS_EMITIDOS = 2

Public Const STRING_TRP_TIPOOCR_DESCRICAO = 100

Public Const STRING_TRP_OCR_HISTORICO = 255
Public Const STRING_TRP_OCR_OBS = 255
Public Const STRING_TRP_OCR_TIPODOC = 10
Public Const STRING_TRP_OCR_TIPOVOU = 10
Public Const STRING_TRP_OCR_SERIE = 10
Public Const STRING_TRP_OCR_CODGRUPO = 10

Public Const STRING_TRP_TAMANHO_OUTROS = 50

Public Const STRING_TRPTITULOS_TIPODOC = 50

Public Const STRING_TRP_VOU_TITULAR = 100
Public Const STRING_TRP_VOU_MOEDA = 10
Public Const STRING_TRP_VOU_CONTROLE = 100
Public Const STRING_TRP_VOU_PRODUTO = 20
Public Const STRING_TRP_VOU_CIACART = 10
Public Const STRING_TRP_VOU_NUMCCRED = 20
Public Const STRING_TRP_VOU_PAXNOME = 100
Public Const STRING_TRP_VOU_DESTINO = 50
Public Const STRING_TRP_VOU_CIAAEREA = 100
Public Const STRING_TRP_VOU_AEROPORTOS = 100
Public Const STRING_TRP_VOU_VALIDADECC = 6

Public Const STRING_TRP_VOU_INFO_TIPODOC = 10
Public Const STRING_TRP_VOU_INFO_TIPVOU = 10
Public Const STRING_TRP_VOU_INFO_SERIE = 10
Public Const STRING_TRP_VOU_HISTORICO = 100

Public Const STRING_TRP_APORTE_HISTORICO = 100
Public Const STRING_TRP_APORTE_OBS = 255

Public Const STRING_TRP_ACORDO_CONTRATO = 20
Public Const STRING_TRP_ACORDO_OBS = 255
Public Const STRING_TRP_ACORDO_DESC = 255

Public Const STATUS_TRP_OCR_LIBERADO = 1
Public Const STATUS_TRP_OCR_BLOQUEADO = 2
Public Const STATUS_TRP_OCR_FATURADO = 3
Public Const STATUS_TRP_OCR_CANCELADO = 4

Public Const STATUS_TRP_OCR_LIBERADO_TEXTO = "Liberado"
Public Const STATUS_TRP_OCR_BLOQUEADO_TEXTO = "Bloqueado"
Public Const STATUS_TRP_OCR_FATURADO_TEXTO = "Faturado"
Public Const STATUS_TRP_OCR_CANCELADO_TEXTO = "Cancelado"

Public Const STATUS_TRP_VOU_ABERTO = 1
Public Const STATUS_TRP_VOU_CANCELADO = 7

Public Const STATUS_TRP_VOU_ABERTO_TEXTO = "Ativo"
Public Const STATUS_TRP_VOU_CANCELADO_TEXTO = "Cancelado"

Public Const TRP_TIPO_ESTORNO_LANC_NORMAL = 0
Public Const TRP_TIPO_ESTORNO_LANC_ESTORNADO = 1
Public Const TRP_TIPO_ESTORNO_LANC_ESTORNADOR = 2
Public Const TRP_TIPO_ESTORNO_LANC_RELANCAMENTO = 3

Public Const TRP_TIPODOC_OCR = 1
Public Const TRP_TIPODOC_NVL = 2
Public Const TRP_TIPODOC_VOU = 3
Public Const TRP_TIPODOC_CMC = 4
Public Const TRP_TIPODOC_CMR = 5
Public Const TRP_TIPODOC_OVER = 6
Public Const TRP_TIPODOC_CMCC = 7

Public Const TRP_TIPODOC_FAT_OUTROS = 1
Public Const TRP_TIPODOC_FAT_NVL = 2
Public Const TRP_TIPODOC_FAT_OCR = 3
Public Const TRP_TIPODOC_FAT_OVER = 4

Public Const TRP_CLIENTEINFO_TIPO_CLIENTE = 1
Public Const TRP_CLIENTEINFO_TIPO_FORNECEDOR = 2

Public Const TRP_TIPODOC_OCR_TEXTO = "OCR"
Public Const TRP_TIPODOC_NVL_TEXTO = "NVL"
Public Const TRP_TIPODOC_VOU_TEXTO = "VOU"
Public Const TRP_TIPODOC_CMC_TEXTO = "CMC"
Public Const TRP_TIPODOC_CMR_TEXTO = "CMR"
Public Const TRP_TIPODOC_OVER_TEXTO = "CME"
Public Const TRP_TIPODOC_BRUTO_TEXTO = "BRUTO"
Public Const TRP_TIPODOC_CMA_TEXTO = "CMA"
Public Const TRP_TIPODOC_CMCC_TEXTO = "CMCC"

Public Const FORMAPAGTO_TRP_OCR_FAT = 1
Public Const FORMAPAGTO_TRP_OCR_CRED = 2

Public Const TRP_VALOR_MINIMO_BOLETO = 2.5

Public Const FORMAPAGTO_TRP_APORTE_TIPOPAGTO_DIRETO = 1
Public Const FORMAPAGTO_TRP_APORTE_TIPOPAGTO_COND = 2

Public Const FORMAPAGTO_TRP_OCR_FAT_TEXTO = "Faturamento"
Public Const FORMAPAGTO_TRP_OCR_CRED_TEXTO = "Crédito"

Public Const BASE_TRP_APORTE_REAL = 1
Public Const BASE_TRP_APORTE_REALACIMAPREV = 2

Public Const BASE_TRP_APORTE_REAL_TEXTO = "Realizado"
Public Const BASE_TRP_APORTE_REALACIMAPREV_TEXTO = "Realizado / Acima"

Public Const TRP_TIPO_DOC_DESTINO_CREDFORN = 1
Public Const TRP_TIPO_DOC_DESTINO_DEBCLI = 2
Public Const TRP_TIPO_DOC_DESTINO_TITREC = 3
Public Const TRP_TIPO_DOC_DESTINO_TITPAG = 4
Public Const TRP_TIPO_DOC_DESTINO_NFSPAG = 5

Public Const TRP_TIPO_DOC_DESTINO_CREDFORN_TEXTO = "Crédito"
Public Const TRP_TIPO_DOC_DESTINO_DEBCLI_TEXTO = "Débito"
Public Const TRP_TIPO_DOC_DESTINO_TITREC_TEXTO = "Título a Receber"
Public Const TRP_TIPO_DOC_DESTINO_TITPAG_TEXTO = "Título a Pagar"
Public Const TRP_TIPO_DOC_DESTINO_NFSPAG_TEXTO = "NF a Pagar"

Public Const TRP_TIPO_DOC_DESTINO_CREDFORN_TABELA = "CreditosPagForn"
Public Const TRP_TIPO_DOC_DESTINO_DEBCLI_TABELA = "DebitosRecCli"
Public Const TRP_TIPO_DOC_DESTINO_TITREC_TABELA = "TitulosRecTodos"
Public Const TRP_TIPO_DOC_DESTINO_TITPAG_TABELA = "TitulosPagTodos"
Public Const TRP_TIPO_DOC_DESTINO_NFSPAG_TABELA = "NFsPag_Todas"

Public Const TRP_TIPO_DOC_DESTINO_CREDFORN_TELA = "CreditosPagar"
Public Const TRP_TIPO_DOC_DESTINO_DEBCLI_TELA = "DebitosReceb"
Public Const TRP_TIPO_DOC_DESTINO_TITREC_TELA = "TituloReceber_Consulta"
Public Const TRP_TIPO_DOC_DESTINO_TITPAG_TELA = "TituloPagar_Consulta"
Public Const TRP_TIPO_DOC_DESTINO_NFSPAG_TELA = "NFPag_Consulta"

Public Const TRP_TIPO_DOC_DESTINO_TITREC_CLASSE = "ClassTituloReceber"
Public Const TRP_TIPO_DOC_DESTINO_TITPAG_CLASSE = "ClassTituloPagar"
Public Const TRP_TIPO_DOC_DESTINO_DEBCLI_CLASSE = "ClassDebitoRecCli"
Public Const TRP_TIPO_DOC_DESTINO_CREDFORN_CLASSE = "ClassCreditoPagar"
Public Const TRP_TIPO_DOC_DESTINO_NFSPAG_CLASSE = "ClassNFsPag"

Public Const TRPCONFIG_PROX_NUM_TITREC = "PROX_NUM_TITREC"
Public Const TRPCONFIG_PROX_NUM_TITPAG = "PROX_NUM_TITPAG"
Public Const TRPCONFIG_DIRETORIO_FAT_HTML = "DIRETORIO_FAT_HTML"
Public Const TRPCONFIG_DIRETORIO_MODELO_FAT_HTML = "DIRETORIO_MODELO_FAT_HTML"
Public Const TRPCONFIG_PROX_NUM_PASSAGEIRO = "PROX_NUM_PASSAGEIRO"
Public Const TRPCONFIG_DIRETORIO_MODELO_FAT_HTML_CARTAO = "DIRETORIO_MODELO_FAT_HTML_CARTAO"
Public Const TRPCONFIG_NUM_INT_PROX_TRPCLIEMISSORES = "NUM_INT_PROX_TRPCLIEMISSORES"
Public Const TRPCONFIG_NUM_INT_PROX_TRPCLIEMISSORESEXC = "NUM_INT_PROX_TRPCLIEMISSORESEXC"
Public Const TRPCONFIG_NUM_INT_PROX_TRPACORDOCOMISS = "NUM_INT_PROX_TRPACORDOCOMISS"
Public Const TRPCONFIG_NUM_INT_PROX_TRPACORDODIF = "NUM_INT_PROX_TRPACORDODIF"
Public Const TRPCONFIG_DATA_COMIS_CORPORATOR = "DATA_COMIS_CORPORATOR"
Public Const TRPCONFIG_VERSAO_CORPORATOR = "TRP_VERSAO"
Public Const TRPCONFIG_SISTEMA_INTEGRADO = "SISTEMA_INTEGRADO"

Public Const INATIVACAO_AUTOMATICA_TEXTO = "Inativação"
Public Const INATIVACAO_AUTOMATICA_CODIGO = 1

Public Const INATIVACAO_AUTOMATICA_TIPO_NVL_TEXTO = "Cancelamento"
Public Const INATIVACAO_AUTOMATICA_TIPO_NVL_CODIGO = 1

Public Const INATIVACAO_AUTOMATICA_TIPO_TAR_TEXTO = "Tarifa"
Public Const INATIVACAO_AUTOMATICA_TIPO_TAR_CODIGO = 2

Public Const INATIVACAO_AUTOMATICA_TIPO_PIS_TEXTO = "PIS"
Public Const INATIVACAO_AUTOMATICA_TIPO_PIS_CODIGO = 3

Public Const INATIVACAO_AUTOMATICA_TIPO_COFINS_TEXTO = "COFINS"
Public Const INATIVACAO_AUTOMATICA_TIPO_COFINS_CODIGO = 4

Public Const INATIVACAO_AUTOMATICA_TIPO_ISS_TEXTO = "ISS"
Public Const INATIVACAO_AUTOMATICA_TIPO_ISS_CODIGO = 5

Public Const VENDEDOR_CARGO_PROMOTOR = 1
Public Const VENDEDOR_CARGO_SUPERVISOR = 2
Public Const VENDEDOR_CARGO_GERENTE = 3
Public Const VENDEDOR_CARGO_DIRETOR = 4

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRPOcorrencias
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
Type typeTRPOcorrenciaDet
    lNumIntDoc As Long
    lNumIntDocOCR As Long
    iTipo As Integer
    dValor As Double
    iSeq As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRPAportes
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
Type typeTRPAportePagtoCond
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
Type typeTRPAportePagtoDireto
    lNumIntDoc As Long
    lNumIntDocAporte As Long
    dValor As Double
    dtVencimento As Date
    lNumIntDocDestino As Long
    iFormaPagto As Integer
    iTipoDocDestino As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRPAportePagtoFat
    lNumIntDoc As Long
    lNumIntDocAporte As Long
    dValor As Double
    dtValidadeDe As Date
    dtValidadeAte As Date
    dSaldo As Double
    dPercentual As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRPAportePagtoFatCond
    lNumIntDoc As Long
    lNumIntDocAporte As Long
    dValor As Double
    dtValidadeDe As Date
    dtValidadeAte As Date
    dPercentual As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRPVouchers
    lNumIntDoc As Long
    lNumVou As Long
    sSerie As String
    lClienteFat As Long
    dValor As Double
    dtData As Date
    sTipVou As String
    lNumFat As Long
    iTipoDocDestino As Integer
    lNumIntDocDestino As Long
    iCartao As Integer
    iStatus As Integer
    lNumIntDocNVL As Long
    dValorAporte As Double
    dValorComissao As Double
    sTitular As String
    sProduto As String
    iPax As Integer
    iMoeda As Integer
    dValorCambio As Double
    dCambio As Double
    sControle As String
    sCiaCart As String
    sNumCCred As String
    lNumAuto As Long
    iQuantParc As Integer
    iDiasAntc As Integer
    iKit As Integer
    dValorOcr As Double
    iTemOcr As Integer
    lRepresentante As Long
    dComissaoRep As Double
    lCorrentista As Long
    dComissaoCorr As Double
    lEmissor As Long
    dComissaoEmissor As Double
    dComissaoAg As Double
    dValorBruto As Double
    sPassageiroNome As String
    sPassageiroSobreNome As String
    lClienteVou As Long
    lCliPassageiro As Long
    dtDataCanc As Date
    dtDataVigenciaDe As Date
    dtDataVigenciaAte As Date
    dHoraCanc As Double
    sUsuarioCanc As String
    lClienteComissao As Long
    sCiaaerea As String
    sAeroportos As String
    lEnderecoPaxTitular As Long
    iTemQueContabilizar As Integer
    iPromotor As Integer
    dComissaoProm As Double
    iDestino As Integer
    sTitularCPF As String
    dTarifaUnitaria As Double
    iVigencia As Integer
    dTarifaUnitariaFolheto As Double
    sPassageiroCGC As String
    dtPassageiroDataNasc As Date
    iGeraComissao As Integer
    iCancelaComissao As Integer
    sUsuarioManut As String
    dtDataUltimaManut As Date
    dHoraUltimaManut As Double
    iIdioma As Integer
    iGrupo As Integer
    dtDataAutoCC As Date
    sValidadeCC As String
    iImprimirValor As Integer
    iCodSegurancaCC As Integer
    dValorBrutoComOCR As Double
    dValorCMAComOCR As Double
    dValorCMC As Double
    dValorCMR As Double
    dValorCMCC As Double
    dValorCME As Double
    sObservacao As String
End Type


'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTitulosRecTRP
    lNumIntDocTitRec As Long
    dValorTarifa As Double
    lNumIntDocNFPagComi As Long
    dValorDeducoes As Double
    dValorComissao As Double
    dValorBruto As Double
End Type


'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRPTitulosExp
    lNumIntDoc As Long
    sUsuario As String
    dtData As Date
    dHora As Double
    iTipoDocOrigem As Integer
    lNumIntDocOrigem As Long
    lNumTitulo As Long
    iExcluido As Integer
    iExportado As Integer
    iTemQueContabilizar As Integer
End Type

Type typeCliEmissoresExcTRP
    lNumIntDoc As Long
    lNumIntDocCliEmi As Long
    iSeq As Integer
    sProduto As String
    dPercComissao As Double
End Type

Type typeCliEmissoresTRP
    lNumIntDoc As Long
    lCliente As Long
    iSeq As Integer
    lFornEmissor As Long
    dPercComissao As Double
    lEmissorSuperior As Long
    lCodigo As Long
End Type

Type typeFiliaisClientesTRP
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
    sGrupo As String
    iCondPagtoCC As Integer
End Type

Type typeTRPAcordoTarifaDif
    lNumIntDoc As Long
    lNumIntAcordoComis As Long
    iDiasDe As Integer
    iDiasAte As Integer
    iDiario As Integer
    dValor As Double
    dValorAdicional As Double
End Type

Type typeTRPAcordoComissao
    lNumIntDoc As Long
    lNumIntAcordo As Long
    iSeq As Integer
    sProduto As String
    iDestino As Integer
    dPercComissao As Double
End Type

Type typeTRPExcComissaoCli
    lCliente As Long
    iSeq As Integer
    sProduto As String
    dPercComissao As Double
End Type

Type typeTRPAcordos
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
Type typeTRPVoucherInfo
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
End Type

Type typeTRPGerComiIntDet
    lNumIntDoc As Long
    lNumIntDocGerComi As Long
    iSeq As Integer
    lNumIntDocComi As Long
    dValorBase As Double
    dValorComissao As Double
    iVendedor As Integer
    dtDataGeracao As Date
    sNomeReduzidoVendedor As String
    dPercComissao As Double
End Type

Type typeTRPGerComiInt
    lNumIntDoc As Long
    dtDataGeracao As Date
    dHoraGeracao As Double
    sUsuario As String
    dtDataEmiAte As Date
    iPrevia As Integer
    sDiretorio As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRPOcorrenciaAporte
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
Type typeTRPVoucherAporte
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
Type typeTRPClienteCorProd
    lCliente As Long
    iSeq As Integer
    lCorrentista As Long
    sProduto As String
    dPercComis As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRPClienteRepProd
    lCliente As Long
    iSeq As Integer
    lRepresentante As Long
    sProduto As String
    dPercComis As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRPTiposOcorrencia
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
Type typeTRPOcrExp
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
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRPVouExp
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
    dtdatafat As Date
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
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRPProdTarifa
    dtVigencia As Date
    sProduto As String
    iDiasDe As Integer
    iDiasAte As Integer
    iDiario As Integer
    dValor As Double
    dValorAdicional As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRPUsuarios
    iTipoUsuario As Integer
    lCodigo As Long
    sLogin As String
    sSenha As String
    iAlteraSenhaProxLog As Integer
    sGrupoAcesso As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRPVouPassageiros
    sTipoVou As String
    sSerie As String
    lNumVou As Long
    iSeq As Integer
    sNome As String
    dtDataNascimento As Date
    sTipoDocumento As String
    sNumeroDocumento As String
    sSexo As String
    dValorPago As Double
    dValorPagoEmi As Double
    iStatus As Integer
    iTitular As Integer
    sPrimeiroNome As String
    sSobreNome As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTRPVouContatos
    sTipoVou As String
    sSerie As String
    lNumVou As Long
    iSeq As Integer
    sNome As String
    sTelefone As String
End Type


Public Function TRPVoucherInfo_Converte_Status(ByVal iStatus As Integer) As String

    Select Case iStatus
    
        Case TRP_VOU_INFO_STATUS_LIBERADO
            TRPVoucherInfo_Converte_Status = TRP_VOU_INFO_STATUS_LIBERADO_TEXTO
        Case TRP_VOU_INFO_STATUS_BLOQUEADO
            TRPVoucherInfo_Converte_Status = TRP_VOU_INFO_STATUS_BLOQUEADO_TEXTO
        Case TRP_VOU_INFO_STATUS_ANTIGA
            TRPVoucherInfo_Converte_Status = TRP_VOU_INFO_STATUS_ANTIGA_TEXTO
    End Select
    
End Function

Public Function gdtDataInicioComisCorp() As Date

Dim sConteudo As String
Dim lErro As Long

On Error GoTo Erro_gdtDataInicioComisCorp

    'Le a data de implantação do comissionamento pelo Corporator
    lErro = CF("TRPConfig_Le", TRPCONFIG_DATA_COMIS_CORPORATOR, EMPRESA_TODA, sConteudo)
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

Public Function TP_Vendedor_Le2(objVendedorMaskEdBox As Object, objVendedor As ClassVendedor, Optional iCria As Integer = 1) As Long
'Lê o Vendedor com Código ou NomeRed em objVendedorMaskEdBox.Text
'Devolve em objVendedor. Coloca código-NomeReduzido no .Text

Dim sVendedor As String
Dim iCodigo As Integer
Dim Vendedor As Object
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo TP_Vendedor_Le2

    Set Vendedor = objVendedorMaskEdBox
    sVendedor = Trim(Vendedor.Text)
    
    'Tenta extrair código de sVendedor
    iCodigo = Codigo_Extrai(sVendedor)
    
    'Se é do tipo código
    If iCodigo > 0 Then
    
        objVendedor.iCodigo = iCodigo
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then Error 25031
        If lErro <> SUCESSO Then Error 25032

        Vendedor.Text = CStr(objVendedor.iCodigo) & SEPARADOR & objVendedor.sNomeReduzido
        
    Else  'Se é do tipo String
            
         objVendedor.sNomeReduzido = sVendedor
         lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
         If lErro <> SUCESSO And lErro <> 25008 Then Error 25029
         If lErro <> SUCESSO Then Error 25030
        
         Vendedor.Text = CStr(objVendedor.iCodigo) & SEPARADOR & sVendedor
    
        
    End If

    TP_Vendedor_Le2 = SUCESSO

    Exit Function

TP_Vendedor_Le2:

    TP_Vendedor_Le2 = Err

    Select Case Err
        
        Case 25029, 25031 'Tratados nas rotinas chamadas

        Case 25030  'Vendedor com NomeReduzido não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Vendedor não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR2", objVendedor.sNomeReduzido)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Vendedores
                    lErro = Chama_Tela("Vendedores", objVendedor)
                End If
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", Err, objVendedor.sNomeReduzido)
            End If
        
        Case 25032  'Vendedor com Código não cadastrado

            If iCria = 1 Then
            
                'Envia aviso que Vendedor não está cadastrado e pergunta se deseja criar
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR1", objVendedor.iCodigo)
    
                If vbMsgRes = vbYes Then
                    'Chama tela de Vendedores
                    lErro = Chama_Tela("Vendedores", objVendedor)
                End If
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", Err, objVendedor.iCodigo)
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153617)

    End Select

End Function
