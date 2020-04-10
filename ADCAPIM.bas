Attribute VB_Name = "ADCAPI"
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function Escape Lib "gdi32" (ByVal hdc As Long, ByVal nEscape As Long, ByVal nCount As Long, ByVal lpInData As String, lpOutData As Any) As Long
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Public Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Public Declare Function LPtoDP Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function DPtoLP Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long

Public Const GETPHYSPAGESIZE = 12

Public Const HORZSIZE = 4           '  Horizontal size in millimeters
Public Const VERTSIZE = 6           '  Vertical size in millimeters

Public Const VERTRES = 10           '  Vertical width in pixels
Public Const HORZRES = 8            '  Horizontal width in pixels

Public Const MM_LOMETRIC = 2

Public Const TIPO_SCRIPT_RESTORE_BD_TESTE = 1

Public Const CLIENTE_LOJA_TRIB = -1

Public Const PLANILHA_MO = 1 'Microsoft Office
Public Const PLANILHA_OO = 2 'Open Office
Public Const PLANILHA_LO = 3 'Libre Office

Public Const ANEXO_TIPO_NF = 1
Public Const ANEXO_TIPO_SAQUE = 2
Public Const ANEXO_TIPO_DEPOSITO = 3
Public Const ANEXO_TIPO_TITREC = 4
Public Const ANEXO_TIPO_TITPAG = 5
Public Const ANEXO_TIPO_OV = 6
Public Const ANEXO_TIPO_PV = 7
Public Const ANEXO_TIPO_NFPAG = 8
Public Const ANEXO_TIPO_OP = 9
Public Const ANEXO_TIPO_PRODENTRADA = 10
Public Const ANEXO_TIPO_REQPROD = 11
Public Const ANEXO_TIPO_ORCSRV = 12
Public Const ANEXO_TIPO_OVHIST = 13
Public Const ANEXO_TIPO_PC = 14
Public Const ANEXO_TIPO_PRJ = 15
Public Const ANEXO_TIPO_VISTPRJ = 16

Public giExeBkp As Integer

Public Const SPED_DOC_NF = 0
Public Const SPED_DOC_CF = 1
Public Const SPED_DOC_CT = 2
Public Const SPED_DOC_NF_INUT = 3

Public Const ACAO_NAO_AVISA = 0
Public Const ACAO_AVISA = 1
Public Const ACAO_ERRO = 2

Public Const USUCONFIG_AVISO_NFE_SEM_AUTO = 1
Public Const USUCONFIG_AVISO_CANC_NFE_SEM_HOM = 2
Public Const USUCONFIG_AVISO_CANC_NFE_HOM = 3

Public Const TIPODOC_INFOADIC_NF = 0
Public Const TIPODOC_INFOADIC_PV = 1
Public Const TIPODOC_INFOADIC_OV = 2
Public Const TIPODOC_INFOADIC_PRJ_PROP = 3
Public Const TIPODOC_INFOADIC_PRJ_CTR = 4
Public Const TIPODOC_INFOADIC_PSRV = 5
Public Const TIPODOC_INFOADIC_OSRV = 6
Public Const TIPODOC_INFOADIC_OVHIST = 7
Public Const TIPODOC_INFOADIC_ITEMPV = 8
Public Const TIPODOC_INFOADIC_ITEMNF = 9
Public Const TIPODOC_INFOADIC_ITEMOV = 10
Public Const TIPODOC_INFOADIC_ITEMOVHIST = 11
Public Const TIPODOC_INFOADIC_PC = 12

Public Const REGIME_TRIBUTARIO_SIMPLES = 1
Public Const REGIME_TRIBUTARIO_SIMPLES_EXCESSO = 2
Public Const REGIME_TRIBUTARIO_NORMAL = 3

Public Const REGIME_TRIBUTARIO_SIMPLES_TEXTO = "Simples Nacional"
Public Const REGIME_TRIBUTARIO_SIMPLES_EXCESSO_TEXTO = "Simples Nacional – excesso de sublimite de receita bruta"
Public Const REGIME_TRIBUTARIO_NORMAL_TEXTO = "Normal"

Public Const STRING_MAXIMO = 255

Public Const TIPO_CAMPO_INTEGER = 1
Public Const TIPO_CAMPO_LONG = 2
Public Const TIPO_CAMPO_DOUBLE = 3
Public Const TIPO_CAMPO_STRING = 4
Public Const TIPO_CAMPO_DATE = 6

Public Const ITEMNF_TIPO_PECA = 1
Public Const ITEMNF_TIPO_SERVICO = 0

Public Const BACKUP_TOKEN_ACAO_LIBERAR = 0
Public Const BACKUP_TOKEN_ACAO_OBTER = 1

Public Const GRADE_LAYOUT_GRADE = 0
Public Const GRADE_LAYOUT_COMBOS = 1

Public Const MENSAGEM_NF_MAXLENGTH = 5000

Public Const PRODUTO_REPETICAO_NAO_PERMITE = 0
Public Const PRODUTO_REPETICAO_AVISA = 1
Public Const PRODUTO_REPETICAO_PERMITE = 2

Public Const STRING_GRUPOEMP_NOMERED = 50
Public Const STRING_GRUPOEMP_DESCRICAO = 250

Public Const VAR_PREENCH_VAZIO = 0              'ainda nao foi preenchida
Public Const VAR_PREENCH_AUTOMATICO = 1         'preenchto segundo calculo do sistema
Public Const VAR_PREENCH_MANUAL = 2             'preenchto pelo usuario

Public Const TIPODOC_TRIB_NF = 0
Public Const TIPODOC_TRIB_PV = 1
Public Const TIPODOC_TRIB_OV = 2
Public Const TIPODOC_TRIB_PRJ_PROP = 3
Public Const TIPODOC_TRIB_PRJ_CTR = 4
Public Const TIPODOC_TRIB_PSRV = 5
Public Const TIPODOC_TRIB_OSRV = 6
Public Const TIPODOC_TRIB_OVHIST = 7
Public Const TIPODOC_TRIB_PRODLOJA = 8

Public Const STRING_SIGNATARIO_CTB = 100
Public Const STRING_CODQUALISIG_CTB = 3

Public Const STRING_SMTP = 50
Public Const STRING_SMTP_USUARIO = 50
Public Const STRING_SMTP_SENHA = 50

Public Const STRING_PRODUTOGENERO_CODIGO = 2
Public Const STRING_PRODUTOGENERO_DESCRICAO = 255
Public Const STRING_ISSQN_CODIGO = 4
Public Const STRING_ISSQN_DESCRICAO = 255
Public Const STRING_CODTRIBMUN = 20

Public Const STRING_CST = 3
Public Const STRING_CSOSN = 5
Public Const STRING_IPI_CLASSE_ENQUADRAMENTO = 5
Public Const STRING_IPI_CODIGO_ENQUADRAMENTO = 3
Public Const STRING_IPI_SELO_CODIGO = 60
Public Const STRING_NATBCCRED = 10

Public Const STRING_DNRC_DESC = 255

Public Const STRING_SPED_CODIGOINST = 10
Public Const STRING_SPED_CODIGOCAD = 50
Public Const STRING_SPED_REGISTRO = 500

Public Const ICMS_OUTROS_DEBITOS = 0
Public Const ICMS_ESTORNO_CREDITOS = 1
Public Const ICMS_OUTROS_CREDITOS = 2
Public Const ICMS_ESTORNO_DEBITOS = 3

Public Const IPI_OUTROS_DEBITOS = 199
Public Const IPI_ESTORNO_CREDITOS = 101
Public Const IPI_OUTROS_CREDITOS = 99
Public Const IPI_ESTORNO_DEBITOS = 1

Public Const SW_NORMAL = 1

Public Const STRING_SPEDFISCAL_PERFIL = 1

Public Const STRING_NFE_NPROT = 15
Public Const STRING_NFE_XMOTIVO = 255

'campos referentes a tela de parametros de relatorios
Public Const REL_OPCOES_ULTIMOS_DIAS = 1
Public Const REL_OPCOES_PROXIMOS_DIAS = 2

Public Const PROC_OLAP_HORARIO_INVALIDO_INICIO = 8
Public Const PROC_OLAP_HORARIO_INVALIDO_FIM = 18

Public Const STRING_SPEDCADASTRO_CODIGOINST = 10
Public Const STRING_SPEDCADASTRO_CODIGOCAD = 50

Public Const CB_VALIDACAO_TIPO_NORMAL = 0
Public Const CB_VALIDACAO_TIPO_BATCH = 1

Public Const EXIBE_BOTOES_FILIAIS = "EXIBE_BOTOES_FILIAIS"

Public Const ERRO_LEITURA_SEM_DADOS = 200000
Public Const ERRO_FUNCAO_COM_ERRO = 200001
Public Const ERRO_SEM_MENSAGEM = 200002

Public Const TIPO_NUMERICO = 0
Public Const TIPO_DATA = 1
Public Const TIPO_HORA = 2
Public Const TIPO_TEXTO = 3
Public Const TIPO_BOOLEANO = 4
Public Const TIPO_INVALIDO = 9

Public Const DATA_IGNORAR = 0
Public Const DATA_ATRAS = 1
Public Const DATA_AFRENTE = 2

Public Const SETA_PARA_CIMA = "^"
Public Const SETA_PARA_BAIXO = "v"

Public Const ORDEM_CRESCENTE = 0
Public Const ORDEM_DECRESCENTE = 1

Public Const STRING_ISS_NUMPROCESSO = 30
Public Const STRING_CIDADE_CODIBGE = 7

'campos referentes a WORKFLOW

Public Const WORKFLOW_ATIVO = 1
Public Const WORKFLOW_INATIVO = 0

Public Const TRANSACAOWFW_PEDIDO_VENDA = 1000
Public Const TRANSACAOWFW_ORCAMENTO_VENDA = 1001
Public Const TRANSACAOWFW_PEDIDO_COMPRA = 1002
Public Const TRANSACAOWFW_ORCAMENTO_SERVICO = 1003
Public Const TRANSACAOWFW_PROJETO = 1010
Public Const TRANSACAOWFW_PRJETAPAS = 1011

Public Const STRING_TRANSACAOWFW_SIGLA = 250
Public Const STRING_TRANSACAOWFW_TRANSACAO = 250
Public Const STRING_TRANSACAOWFW_TRANSACAOTELA = 250
Public Const STRING_TRANSACAOWFW_ORIGEM = 250
Public Const STRING_TRANSACAOWFW_OBSERVACAO = 250

Public Const STRING_REGRAWFW_MODULO = 3
Public Const STRING_REGRAWFW_REGRA = 250
Public Const STRING_REGRAWFW_EMAILPARA = 250
Public Const STRING_REGRAWFW_EMAILASSUNTO = 250
Public Const STRING_REGRAWFW_EMAILMSG = 250
Public Const STRING_REGRAWFW_AVISOMSG = 250
Public Const STRING_REGRAWFW_LOGDOC = 250
Public Const STRING_REGRAWFW_LOGMSG = 250
Public Const STRING_REGRAWFW_RELSEL = 250
Public Const STRING_REGRAWFW_RELANEXO = 250

Public Const STRING_AVISOWFW_MSG = 255

Public Const STRING_REGRAMSG = 255
Public Const STRING_MENSAGEMREGRA = 255


Public Const STRING_MNEMONICOWFW_MNEMONICO = 20
Public Const STRING_MNEMONICOWFW_MNEMONICOCOMBO = 50
Public Const STRING_MNEMONICOWFW_MNEMONICODESC = 255

Public Const AVISOWFW_INTERVALO_MINUTO = 1
Public Const AVISOWFW_INTERVALO_HORA = 2
Public Const AVISOWFW_INTERVALO_DIA = 3
Public Const AVISOWFW_INTERVALO_SEMANA = 4


'indica qual a formula que está sendo executada
Public Const CAMPO_REGRA = 1
Public Const CAMPO_EMAILASSUNTO = 2
Public Const CAMPO_EMAILMSG = 3
Public Const CAMPO_AVISOMSG = 4
Public Const CAMPO_LOGDOC = 5
Public Const CAMPO_LOGMSG = 6
Public Const CAMPO_EMAILPARA = 7
Public Const CAMPO_RELSEL = 8
Public Const CAMPO_RELANEXO = 9

'#####################################
'Inserido por Wagnerv 20/10/2005
Public Const LIMITA_DATA_USO_NAO = 0
Public Const LIMITA_DATA_USO_DATA = 1
Public Const LIMITA_DATA_USO_QTD = 2
Public Const LIMITA_DATA_USO_DATA_QTD = 3
'#####################################

Public Const STRING_VALIDAEXCLUSAO_CODIGO = 50
Public Const STRING_VALIDAEXCLUSAO_TABELA = 50
Public Const STRING_VALIDAEXCLUSAO_CAMPO = 50
Public Const STRING_VALIDAEXCLUSAO_CAMPOLER = 50
Public Const STRING_VALIDAEXCLUSAO_MSGERRO1 = 250
Public Const STRING_VALIDAEXCLUSAO_MSGERRO2 = 250
Public Const STRING_VALIDAEXCLUSAO_MSGERROLER = 250

Public Const VALIDAEXCLUSAO_CODIGO_PRODUTO = "PRODUTO"
Public Const VALIDAEXCLUSAO_CODIGO_CLIENTE = "CLIENTE"
Public Const VALIDAEXCLUSAO_CODIGO_CONTACONTABIL = "CONTA_CONTABIL"
Public Const VALIDAEXCLUSAO_CODIGO_PARCELARECEBER = "PARCELA_RECEBER"
Public Const VALIDAEXCLUSAO_CODIGO_PROJETO = "PROJETO"
Public Const VALIDAEXCLUSAO_CODIGO_ETAPAPRJ = "ETAPAPRJ"
Public Const VALIDAEXCLUSAO_CODIGO_PROPOSTA = "PROPOSTA"
Public Const VALIDAEXCLUSAO_CODIGO_CONTRATOPRJ = "CONTRATOPRJ"
Public Const VALIDAEXCLUSAO_CODIGO_TIPOGARANTIA = "TIPO_GARANTIA"
Public Const VALIDAEXCLUSAO_CODIGO_CERTIFICADOS = "CERTIFICADOS"

Public Const VALIDAEXCLUSAO_GENERO_MASCULINO = 0
Public Const VALIDAEXCLUSAO_GENERO_FEMININO = 1

Public Const CONTROLE_PATH_LOCK = 1
Public Const CONTROLE_BROWSE_POSICAO_ANTIGO = 2
Public Const CONTROLE_TOKEN_BACKUP = 3
Public Const CONTROLE_DIR_XMLS = 101
Public Const CONTROLE_DIR_XMLS_EMP01 = 901 '900 + Código da Empresa
Public Const CONTROLE_DIR_XMLS_EMP99 = 999 '900 + Código da Empresa
Public Const CONTROLE_ID_CLIENTE = 1000
Public Const CONTROLE_VERSAO_PGM_INSTALADA = 1001
Public Const CONTROLE_VERSAO_SGEUPDATE_INSTALADA = 1002 'e SGEUPDATELOJA no BD de ECF
Public Const CONTROLE_VERSAO_RELS_INSTALADA = 1003
Public Const CONTROLE_NOME_PGM_IMPORTXML = 1100
Public Const CONTROLE_ID_CLIENTE_TEMP_ECF = 2000 'BD ECF. Usado somente na instalação para repassar a informação da retaguarda para o ECF
Public Const CONTROLE_DIR_XMLS_EMP01_FIL01 = 90101 '90000 + (Codigo da Empresa) * 100 + Código da Filial Empresa
Public Const CONTROLE_DIR_XMLS_EMP99_FIL99 = 99999 '90000 + (Codigo da Empresa) * 100 + Código da Filial Empresa

Public Const STRING_CONTROLE_CONTEUDO = 255
Public Const STRING_CONTROLE_DESCRICAO = 255

Public Const CHEQUE_APROVADO = 1
Public Const CHEQUE_NAO_APROVADO = 0
Public Const STRING_CHEQUE_APROVADO = "Aprovado"
Public Const STRING_CHEQUE_NAO_APROVADO = "Não Aprovado"

Public Const USA_BALANCA = 1
Public Const NAO_USA_BALANCA = 0
Public Const USA_BALANCA_PARA_ETIQUETA = 2

Public Const CATEGORIAPRODUTO_GENERICO = "Generico"

Public Const NOME_ARQUIVO_ADM = "ADM100.INI"
Public Const STRING_MAX_NOME_ARQUIVO = 255

Public Const ITEM_GRADE = -1

'constantes que preenchem o campo origempedido das tabelas ItensOrdemProducao e ItensOrdemProducaoGrade
Public Const ORIGEM_PEDIDO_ITEM_PV = 0
Public Const ORIGEM_PEDIDO_ITEM_PV_GRADE = 1

Public Const NAO_EXIBE_MSG_ERRO = 1

'serve para indicar se o cliente esta cadastrado no backoffice ou somente na loja
Public Const CLIENTE_CADASTRO_LOJA = 1
Public Const CLIENTE_CADASTRO_BACK = 0

Public Const CAMPO_PREENCH_OPCIONAL = 0
Public Const CAMPO_PREENCH_OBRIGATORIO = 1
Public Const CAMPO_PREENCH_OPCIONAL_AVISO = 2 'avisa que nao está preenchido mas deixa gravar

'constantes referentes a origem das anotacoes
Public Const ANOTACAO_ORIGEM_CLIENTE = 1
Public Const ANOTACAO_ORIGEM_FORNECEDOR = 2
Public Const ANOTACAO_ORIGEM_NFISCAL = 3
Public Const ANOTACAO_ORIGEM_PEDIDOVENDA = 4
Public Const ANOTACAO_ORIGEM_TITPAG = 5
Public Const ANOTACAO_ORIGEM_TITREC = 6
Public Const ANOTACAO_ORIGEM_ORCVENDA = 7 'Wagner
Public Const ANOTACAO_ORIGEM_PEDIDOSERVICO = 8
Public Const ANOTACAO_ORIGEM_ORCSRV = 9
Public Const ANOTACAO_ORIGEM_PRODUTO = 10
Public Const ANOTACAO_ORIGEM_MOVESTOQUE = 11
Public Const ANOTACAO_ORIGEM_INVENTARIOLOTE = 12
Public Const ANOTACAO_ORIGEM_INVENTARIO = 13
Public Const ANOTACAO_ORIGEM_OP = 14

'constantes referentes a tela de anotacoes
Public Const STRING_ANOTACOES_ID = 50
Public Const STRING_ANOTACOES_TITULO = 50
Public Const STRING_ANOTACOESLINHA_TEXTO = 255
Public Const STRING_ORIGEMANOTACOES_DESCRICAO = 50
Public Const STRING_ORIGEMANOTACOES_NOMETABELA = 50

Public Const ERRO_OBJETO_NAO_CADASTRADO = 32999

Public Const AD_SIST_NORMAL = 0
Public Const AD_SIST_RELLIB = 1
Public Const AD_SIST_BATCH = 2


Public Declare Function Sistema_ObterTipoCliente Lib "ADCUSR.DLL" Alias "AD_Sistema_ObterTipoCliente" (ByVal lID_Sistema As Long) As Long

'Constantes Formulas Grid AnaliseLin
Public Const GRID_FORMATO_MOEDA = 1
Public Const GRID_FORMATO_PERCENTAGEM = 2
Public Const GRID_FORMATO_MOEDA_STRING = "Moeda"
Public Const GRID_FORMATO_PERCENTAGEM_STRING = "Percentagem"

Public Const PRODUZIDO_NA_FILIAL = 1
Public Const NAO_PRODUZIDO_NA_FILIAL = 0

Public Const DATA_INICIO_CFOP4 = #1/1/2003#
Public bSGECancelDummy As Boolean

'Moeda Pré-Cadastrada
Public Const MOEDA_REAL = 0
Public Const MOEDA_DOLAR = 1
Public Const MOEDA_EURO = 2
Public Const STRING_NOME_MOEDA = 20
Public Const STRING_SIMBOLO_MOEDA = 10

Public Const LOCALOPERACAO_CAIXA_CENTRAL_BACKOFFICE = 1
Public Const LOCALOPERACAO_BACKOFFICE = 2
Public Const LOCALOPERACAO_CAIXA_CENTRAL = 3
Public Const LOCALOPERACAO_ECF = 4

'utilizada na tabela TabelaConfig
Public Const STRING_NOMETABELA = 50

'flag utilizada nas tabelas config para indicar que vai haver um registro por filial daquela configuracao
Public Const POR_FILIAL = 1

'indica que não é possível replicar o campo do grid
Public Const GRID_CONTEUDO_INVALIDO_PARA_REPLICAR = 94952

'Constantes que guardam nomes de classes
Public Const NOME_CLASSE_ESTADO As String = "ClassEstado"
Public Const NOME_CLASSE_ICMSALIQEXTERNA As String = "ClassICMSAliqExterna"
Public Const NOME_CLASSE_PEDIDOVENDA As String = "ClassPedidoDeVenda"
Public Const NOME_CLASSE_NOTAFISCAL As String = "ClassNFiscal"
Public Const NOME_CLASSE_CTNFISCAL As String = "CTNfiscal"
Public Const NOME_CLASSE_CTNFISCALPEDIDO As String = "CTNFiscalPedido"
Public Const NOME_CLASSE_PRODUTO As String = "ClassProduto"
'****************************************

'Constantes utilizadas como Index para localização de itens em colcolComissoesRegras
'Utilizadas também para selecionar um item da combo DiretoIndireto no Tab de Comissões
Public Const VENDEDOR_DIRETO_STRING = "Direto"
Public Const VENDEDOR_INDIRETO_STRING = "Indireto"
Public Const VENDEDOR_TODOS_STRING = "Todos"
'************************************************************************************

'Mnemônicos genéricos utilizados para cálculo de comissões
Public Const MNEMONICO_COMISSOES_VENDEDOR = "Vendedor"
Public Const MNEMONICO_COMISSOES_REGIAO = "RegiaoVenda"
Public Const MNEMONICO_COMISSOES_CLIENTE = "Cliente"
Public Const MNEMONICO_COMISSOES_FILIALCLI = "FilialCliente"
Public Const MNEMONICO_COMISSOES_ITEMCATPRODUTO = "ItemCatProduto"
Public Const MNEMONICO_COMISSOES_PRODUTO_VLRTOTAL = "Produto_VlrTotal"
Public Const MNEMONICO_COMISSOES_PRODUTO_VLRVENDA = "Produto_VlrVenda"
Public Const MNEMONICO_COMISSOES_PRODUTO_VLRFRETE = "Produto_VlrFrete"
Public Const MNEMONICO_COMISSOES_PRODUTO_VLRSEGURO = "Produto_VlrSeguro"
Public Const MNEMONICO_COMISSOES_PRODUTO_VLROUTRASDESP = "Produto_VlrDespesas"
Public Const MNEMONICO_COMISSOES_PRODUTO_VLRIPI = "Produto_VlrIPI"
Public Const MNEMONICO_COMISSOES_PRODUTO_VLRICMS = "Produto_VlrICMS"
Public Const MNEMONICO_COMISSOES_PRODUTO = "Produto"
Public Const MNEMONICO_COMISSOES_PRECO_VENDA = "Produto_PrecoVenda"
Public Const MNEMONICO_COMISSOES_COMISSAO_ITEM = "Comissao_Item"
Public Const MNEMONICO_COMISSOES_TABPRECO_ITEM = "TabPreco_Item"
Public Const MNEMONICO_COMISSOES_COMISS_TABPRECO_ITEM = "Comis_TabPreco_Item"
Public Const MNEMONICO_COMISSOES_PRODUTO_VLRDESCONTO = "Produto_VlrDesconto"
Public Const MNEMONICO_COMISSOES_PRECO_TABELA = "Produto_PrecoTabela"

Public Const MNEMONICO_COMISSOES_TABELAPRECO_CODIGO = "TabPreco_Cod"
Public Const MNEMONICO_COMISSOES_TABELAPRECO_DESCRICAO = "TabPreco_Desc"
Public Const MNEMONICO_COMISSOES_PRODUTO_PERCDESCONTO = "Produto_PercDesc"
Public Const MNEMONICO_COMISSOES_VENDEDOR_VINCULO_CODIGO = "Vend_Vinculo_Cod"
Public Const MNEMONICO_COMISSOES_VENDEDOR_VINCULO_DESCRICAO = "Vend_Vinculo_Desc"
Public Const MNEMONICO_COMISSOES_ITEMCATCLIENTE = "ItemCatCliente"
Public Const MNEMONICO_COMISSOES_TIPOPRODUTO = "TipoProduto"
Public Const MNEMONICO_COMISSOES_PERCCOMISSCLI = "PercComissCli"
Public Const MNEMONICO_COMISSOES_CANALVENDA = "CanalVenda"
'************************************************************

Public Const MNEMONICO_COMISSOES_TRVTIPOCLIENTE = "TRVTipoCliente"
Public Const MNEMONICO_COMISSOES_TRVPRODUTO = "TRVProduto"
Public Const MNEMONICO_COMISSOES_TRVPRODUTONOMERED = "TRVProdutoNomeRed"
Public Const MNEMONICO_COMISSOES_TRVVENDAANT = "TRVVendaAnt"
Public Const MNEMONICO_COMISSOES_TRVCMA = "TRVTemCMA"
Public Const MNEMONICO_COMISSOES_TRVCMCC = "TRVTemCMCC"
Public Const MNEMONICO_COMISSOES_TRVVENDAANUAL = "TRVVendaAnual"
Public Const MNEMONICO_COMISSOES_TRVPRECONET = "TRVPrecoNet"
Public Const MNEMONICO_COMISSOES_TRVPRODUTOPERCCOMIS = "TRVProdutoPercComis"
Public Const MNEMONICO_COMISSOES_TRVTARIFAALTERADA = "TRVTarifaAlterada"

Public Const MNEMONICO_COMISSOES_TRPTIPOCLIENTE = "TRPTipoCliente"
Public Const MNEMONICO_COMISSOES_TRPPRODUTO = "TRPProduto"
Public Const MNEMONICO_COMISSOES_TRPPRODUTONOMERED = "TRPProdutoNomeRed"
Public Const MNEMONICO_COMISSOES_TRPVENDAANT = "TRPVendaAnt"
Public Const MNEMONICO_COMISSOES_TRPCMA = "TRPTemCMA"
Public Const MNEMONICO_COMISSOES_TRPCMCC = "TRPTemCMCC"
Public Const MNEMONICO_COMISSOES_TRPVENDAANUAL = "TRPVendaAnual"
Public Const MNEMONICO_COMISSOES_TRPPRECONET = "TRPPrecoNet"
Public Const MNEMONICO_COMISSOES_TRPPRODUTOPERCCOMIS = "TRPProdutoPercComis"
Public Const MNEMONICO_COMISSOES_TRPTARIFAALTERADA = "TRPTarifaAlterada"

'Parâemtros de mnemônicos de comissões
Public Const MNEMONICO_PARAMETRO_CATEGORIA As String = "ProdutoCategoria"
'************************************************************

'Constante utilizada na tela de ComissoesRegras
Public Const GRID_PERMITIDO_INCLUIR_NO_MEIO = 0
'*******************************************************

'************************* Nomes de telas ************************************************
Public Const NOME_TELA_GERACAONFISCAL As String = "GeracaoNFiscal"
Public Const NOME_TELA_PEDIDOVENDA As String = "PedidoVenda"
Public Const NOME_TELA_PEDIDOVENDACONSULTA As String = "PedidoVenda_Consulta"
Public Const NOME_TELA_NFISCAL As String = "NFiscal"
Public Const NOME_TELA_NFISCALREM As String = "NFiscalRem"
Public Const NOME_TELA_NFISCALDEV As String = "NFiscalDev"
Public Const NOME_TELA_NFISCALFATURA As String = "NFiscalFatura"
Public Const NOME_TELA_NFISCALPEDIDO As String = "NFiscalPedido"
Public Const NOME_TELA_NFISCALFATURAPEDIDO As String = "NFiscalFaturaPedido"
Public Const NOME_TELA_NFISCALFATPEDSRV As String = "NFiscalFatPedSRV"
Public Const NOME_TELA_CONHECIMENTOFRETEFATURA As String = "ConhecimentoFreteFatura"
Public Const NOME_TELA_COMISSOESCALCULA As String = "ComissoesCalcula"
Public Const NOME_TELA_COMISSOESCALCULALOJA As String = "ComissoesCalculaLoja"
Public Const NOME_TELA_ORCAMENTOVENDA As String = "OrcamentoVenda"
Public Const NOME_TELA_RECEBMATERIALC As String = "RecebMaterialC"
Public Const NOME_TELA_RECEBMATERIALF As String = "RecebMaterialF"
Public Const NOME_TELA_NFISCALENTRADA As String = "NFiscalEntrada"
Public Const NOME_TELA_NFISCALFATENTRADA As String = "NFiscalFatEntrada"
Public Const NOME_TELA_NFISCALENTDEV As String = "NFiscalEntDev"
Public Const NOME_TELA_NFISCALENTREM As String = "NFiscalEntRem"
Public Const NOME_TELA_ORDEMPRODUCAO As String = "OrdemProducao"
Public Const NOME_TELA_PRODUCAOSAIDA As String = "ProducaoSaida"
Public Const NOME_TELA_PRODUCAOENTRADA As String = "ProducaoEntrada"
Public Const NOME_TELA_NFISCALFATURAPEDSRV As String = "NFiscalFatPedSRV"
Public Const NOME_TELA_NFISCALFATURAGARSRV As String = "NFiscalFatGarSRV"
Public Const NOME_TELA_NFISCALSRV As String = "NFiscalSRV"

'*****************************************************************************************

Public Const CLASSE_VENDEDOR = "ClassVendedor"

'Incluído por Luiz Nogueira em 27/10/03
'Usadas para a tabela de RelacionamentoClientes
Public Const RELACIONAMENTOCLIENTES_STATUS_ENCERRADO = 1
'*****************************************************************************************

Public Const DELTA_VALORMONETARIO = 0.009
Public Const DELTA_VALORMONETARIO2 = 0.00001

Public Const QTDE_ESTOQUE_DELTA = 0.0000000001 'durante atualizacao de saldo, se abs(saldo) < QTDE_ESTOQUE_DELTA entao saldo é tornado zero

Public Const QTDE_ESTOQUE_DELTA2 = 0.00001

Public Const STRING_GRADE_CODIGO = 20
Public Const STRING_GRADE_DESCRICAO = 50


Public Const STRING_CATEGORIAPRODUTOITEM_ITEM = 20

Public Const MNEMONICOFPRECO_NAO_E_FUNCAO = 0
Public Const MNEMONICOFPRECO_E_FUNCAO = 1

Public Const MNEMONICOFPRECO_CUSTO_PRODUCAO = "Custo_Produto"
Public Const MNEMONICOFPRECO_CUSTO_PRODUCAO_PROD = "Custo_Produto_Param"
Public Const MNEMONICOFPRECO_CUSTO_KIT_MP = "Custo_Kit_MP"
Public Const MNEMONICOFPRECO_CUSTO_KIT_MP_PARAM = "Custo_Kit_MP_Param"
Public Const MNEMONICOFPRECO_CUSTO_KIT_MP2_PARAM = "Custo_Kit_MP2_Param"
Public Const MNEMONICOFPRECO_CUSTOREP_KIT_MP = "CustoRep_Kit_MP"
Public Const MNEMONICOFPRECO_CUSTOREP_KIT_MP_PARAM = "CustoRep_Kit_MP_P"
Public Const MNEMONICOFPRECO_CUSTO_ULT_ENTRADA = "Custo_Ult_Entrada"

Public Const MNEMONICOFPRECO_VALORCATEGORIA = "ValorCategoria"
Public Const MNEMONICOFPRECO_FRETEFILCLI = "ValorFreteFilCli"

Public Const ICMS_MAIOR_ALIQ_INTERESTADUAL = 0.12

Public Const MNEMONICOFPRECO_CUSTOROTEIROINSUMOSMAQ = "CustoRotInsumos"
Public Const MNEMONICOFPRECO_CUSTOROTEIROMAOOBRA = "CustoRotMaoObra"
Public Const MNEMONICOFPRECO_PRECOPRATICADO = "PrecoPraticado"
Public Const MNEMONICOFPRECO_VALPRES = "ValPres"
Public Const MNEMONICOFPRECO_VALFUT = "ValFut"
Public Const MNEMONICOFPRECO_CUSTOMP = "CustoMP"
Public Const MNEMONICOFPRECO_CUSTOEMB = "CustoEmb"
Public Const MNEMONICOFPRECO_DVVTOTAL = "DVVTotal"
Public Const MNEMONICOFPRECO_ICMSALIQTABPRECO = "ICMSAliqTabPreco"
Public Const MNEMONICOFPRECO_TABELAPRECOPADRAO = "TabelaPrecoPadrao"
Public Const MNEMONICOFPRECO_TABELAPRECO = "TabelaPreco"
Public Const MNEMONICOFPRECO_DEVDUV = "DevDuv"
Public Const MNEMONICOFPRECO_CUSTODIRETO = "CustoDireto"
Public Const MNEMONICOFPRECO_CUSTOFIXO = "CustoFixo"
Public Const MNEMONICOFPRECO_DIASCONDPAGTO = "DiasCondPagto"
Public Const MNEMONICOFPRECO_CONDPAGTOCLIENTE = "CondPagtoCliente"
Public Const MNEMONICOFPRECO_PRECOULTPV = "PrecoUltPV"
Public Const MNEMONICOFPRECO_PRECOULTNF = "PrecoUltNF"
Public Const MNEMONICOFPRECO_COMISSOESPROPCLIPROD = "ComissoesPropCliProd"
Public Const MNEMONICOFPRECO_COMISSOESTERCCLIPROD = "ComissoesTercCliProd"
Public Const MNEMONICOFPRECO_ENCCOMISSOESPROPCLIPROD = "EncSociaisProp"
Public Const MNEMONICOFPRECO_ENCCOMISSOESTERCCLIPROD = "EncSociaisTerc"
Public Const MNEMONICOFPRECO_CLIENTE = "Cliente"
Public Const MNEMONICOFPRECO_FILIALCLI = "FilialCli"
Public Const MNEMONICOFPRECO_PRODUTO = "Produto"
Public Const MNEMONICOFPRECO_QUANTIDADE = "Quantidade"
Public Const MNEMONICOFPRECO_ICMSALIQFILCLI = "ICMSAliqFilCli"
Public Const MNEMONICOFPRECO_PRECOTABELA = "PrecoTabela"
Public Const MNEMONICOFPRECO_PRECOTABELA_PARAM = "PrecoTabela_Param"
Public Const MNEMONICOFPRECO_IPIALIQUOTA = "IPIAliquota"
Public Const MNEMONICOFPRECO_ALIQUOTA_PRODUTO = "Aliquota_Produto"
Public Const MNEMONICOFPRECO_ICMSALIQFILCLIPROD = "ICMSAliqFilCliProd"

'Extensão .TXT para arquivo
Public Const EXTENSAO_ARQUIVO_TXT As String = ".txt"

Public Const CATEGORIA_PRODUTO_PRECO As String = "Preço"

Public Const SM_CYCAPTION = 4       ' Height of caption or title
Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
Public Const SM_CYDLGFRAME = 8
Public Const SM_CXDLGFRAME = 7
Public Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Public Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Public Const FATOR_POLEGADA_PONTO = 72

'constantes para o loja
Public Const CHEQUEPRE_LOCALIZACAO_BACKOFFICE = 0
Public Const CHEQUEPRE_LOCALIZACAO_LOJA = 1
Public Const CHEQUEPRE_LOCALIZACAO_BANCO = 2
Public Const CHEQUEPRE_LOCALIZACAO_CAIXA = 3
Public Const CHEQUEPRE_LOCALIZACAO_EM_TERCEIROS = 4

'Utilizado no Browser
Public Const BROWSE_MAX_REGISTROS_EXPORTAR = 64000
Public Const BROWSE_SUGESTAO_REGISTROS_EXPORTAR = 64000 '1000
Public Const BROWSE_NUMERO_BOTOES = 5
Public Const BROWSE_EXECUTA_BIND = 1
Public Const ESPACO_ENTRE_BOTOES = 100
Public Const LARGURA_BOTOES = 915
Public Const MINIMO_LARGURA_BOTAO = 100

'é uma flag genérica usada nas rotinas de estatistica para indicar se deve cancelar/cadastrar as quantidades computadas nas tabelas estatisticas (SldMesFat, SldMesForn, etc)
Public Const CANCELAMENTO_DOC = 1
Public Const CADASTRAMENTO_DOC = 2

'Constantes detectadas como necessarias na tela
'PlanComissoesInpal
Public Const CODIGO_NAO_PREENCHIDO = 0
Public Const E_ENTRE_ESPACOS = " E "
Public Const OPERADOR_IGUAL = "="
Public Const OPERADOR_SOMA = "+"
Public Const OPERADOR_SUBTRACAO = "-"
Public Const OPERADOR_MAIORIGUAL = ">="
Public Const OPERADOR_MENORIGUAL = "<="
Public Const OPERADOR_DIFERENTE = "<>"
Public Const OPERADOR_MAIOR = ">"
Public Const OPERADOR_MENOR = "<"
Public Const STRING_ESPACO = " "
Public Const STRING_VAZIO = ""
Public Const PERCENTUAL_NAO_PREENCHIDO = 0

'Para Grid Alocacao de NFiscal
Public Const NUM_MAX_ALOCACOES = 1500

Public Const CONCORRENCIA_STATUS_ATIVO = 0
Public Const CONCORRENCIA_STATUS_BAIXADO = 1

Public Const COTACAO_STATUS_ATIVO = 0
Public Const COTACAO_STATUS_BAIXADO = 1

Public Const COTACAOITEMCONCORRENCIA_STATUS_ATIVO = 0
Public Const COTACAOITEMCONCORRENCIA_STATUS_BAIXADO = 1

Public Const COTACAOPRODUTO_STATUS_ATIVO = 0
Public Const COTACAOPRODUTO_STATUS_BAIXADO = 1

Public Const ITEMPEDCOTACAO_STATUS_ATIVO = 0
Public Const ITEMPEDCOTACAO_STATUS_BAIXADO = 1

Public Const ITENSCONCORRENCIA_STATUS_ATIVO = 0
Public Const ITENSCONCORRENCIA_STATUS_BAIXADO = 1

Public Const ITENSCOTACAO_STATUS_ATIVO = 0
Public Const ITENSCOTACAO_STATUS_BAIXADO = 1

Public Const ITENSPEDCOMPRA_STATUS_ATIVO = 0
Public Const ITENSPEDCOMPRA_STATUS_BAIXADO = 1

Public Const ITENSREQCOMPRA_STATUS_ATIVO = 0
Public Const ITENSREQCOMPRA_STATUS_BAIXADO = 1

Public Const PEDIDOCOTACAO_STATUS_ATIVO = 0
Public Const PEDIDOCOTACAO_STATUS_BAIXADO = 1

Public Const PEDIDOCOMPRA_STATUS_ATIVO = 0
Public Const PEDIDOCOMPRA_STATUS_BAIXADO = 1

Public Const REQUISICAOCOMPRA_STATUS_ATIVO = 0
Public Const REQUISICAOCOMPRA_STATUS_BAIXADO = 1


'utilizado pela rotina Conexao_DesativarLocks que ativa/desativa os locks
Public Const DESATIVAR_LOCKS = 1
Public Const REATIVAR_LOCKS = 0

'Utilizado pela função Formata_Campo
Public Const ALINHAMENTO_ESQUERDA = 1
Public Const ALINHAMENTO_DIREITA = 2
Public Const ALINHAMENTO_CENTRALIZADO = 3

'*** EXCEL ***
'Constantes utilizadas na exportação para o excel
Public Const EXCEL_MODO_IMPRESSÃO_GRAFICO = 1 '??? Futuramente: substituir essa constante por EXCEL_MODO_IMPRESSAO
Public Const EXCEL_MODO_IMPRESSAO = 1
Public Const EXCEL_NAO_PARTICIPA_GRAFICO = 0
Public Const EXCEL_PARTICIPA_GRAFICO_X = 1
Public Const EXCEL_PARTICIPA_GRAFICO_Y = 2

Public Const EXCEL_FORMATO_CSV = 1
Public Const EXCEL_FORMATO_XLS = 2

Public Const EXCEL_FORMULA_SUM = -4157
Public Const EXCEL_FORMULA_COUNT = -4113
Public Const EXCEL_FORMULA_MAX = -4136
Public Const EXCEL_FORMULA_MIN = -4139
Public Const EXCEL_FORMULA_AVG = -4106

Public Const EXCEL_FORMULA_SUM_TEXTO = "Somar"
Public Const EXCEL_FORMULA_COUNT_TEXTO = "Contar"
Public Const EXCEL_FORMULA_MAX_TEXTO = "Máximo"
Public Const EXCEL_FORMULA_MIN_TEXTO = "Mínimo"
Public Const EXCEL_FORMULA_AVG_TEXTO = "Média"

Public Const EXCEL_TABDIN_POS_LINHA = 1
Public Const EXCEL_TABDIN_POS_COLUNA = 2
Public Const EXCEL_TABDIN_POS_FILTRO = 3
Public Const EXCEL_TABDIN_POS_VALOR = 4

Public Const EXCEL_TABDIN_POS_LINHA_TEXTO = "Linha"
Public Const EXCEL_TABDIN_POS_COLUNA_TEXTO = "Coluna"
Public Const EXCEL_TABDIN_POS_FILTRO_TEXTO = "Filtro"
Public Const EXCEL_TABDIN_POS_VALOR_TEXTO = "Dados"

Public Const EXCEL_TIPOGRAFICO_PIZZA = -4102
Public Const EXCEL_TIPOGRAFICO_LINHA = -4101
Public Const EXCEL_TIPOGRAFICO_AREA = -4098
Public Const EXCEL_TIPOGRAFICO_COLUNA = -4100

Public Const EXCEL_TIPOGRAFICO_PIZZA_TEXTO = "Pizza"
Public Const EXCEL_TIPOGRAFICO_LINHA_TEXTO = "Linha"
Public Const EXCEL_TIPOGRAFICO_AREA_TEXTO = "Área"
Public Const EXCEL_TIPOGRAFICO_COLUNA_TEXTO = "Coluna"

Public Const EXCEL_TABDIN_TIPO_DATABASE = 1
Public Const EXCEL_TABDIN_TIPO_EXTERNAL = 2
Public Const EXCEL_TABDIN_TIPO_CONSOLIDATION = 3
Public Const EXCEL_TABDIN_TIPO_SCENARIO = 4
Public Const EXCEL_TABDIN_TIPO_PIVOTTABLE = -4148

'Constantes para setagem de configurações do gráfico no excel
'tipos de gráfico
Public Const EXCEL_GRAFICO_COLUMN_CLUSTERED = 51 'Coluna
Public Const EXCEL_GRAFICO_3D_COLUMN_CLUSTERED = 54 'Coluna 3D
Public Const EXCEL_GRAFICO_3D_COLUMN_STACKED = 55
Public Const EXCEL_GRAFICO_3DPIE = -4102 'Torta 3D
Public Const EXCEL_GRAFICO_LINE_MARKERS = 65 'Linha com marcas
Public Const EXCEL_GRAFICO_3D_LINE = -4101 'Linha 3D

Public Const EXCEL_PLANILHA_IMPRIME_GRADE As Boolean = True
Public Const EXCEL_PLANILHA_ORIENTACAO_LANDSCAPE = 2
Public Const EXCEL_LABEL_ORIENTACAO_HORIZONTAL = -4128
Public Const EXCEL_TICKLABEL_POSITION_LOW = -4134
Public Const EXCEL_TICKLABEL_ORIENTATION_UPWARD = -4171
Public Const EXCEL_COLUMNS = 2
Public Const EXCEL_GRAFICO_COM_EIXOS = 1 'Determina que o gráfico possui os eixos X e Y
Public Const EXCEL_GRAFICO_SEM_EIXOS = 2 'Determina que o gráfico não possui os eixos X e Y (Ex.: gráficos do tipo torta)
Public Const EXCEL_NAO_EXIBE_LABELS = -4142 'Determina que o gráfico NÃO exibirá labels (Corresponde à constante xlDataLabelsShowNone da biblioteca do Excel)
Public Const EXCEL_EXIBE_LABELS_VALOR = 2 'Determina que o gráfico exibirá os labels com valores (Corresponde à constante xlDataLabelsShowValue da biblioteca do Excel)
Public Const EXCEL_EXIBE_LABELS_PERCENTUAL = 3 'Determina que o gráfico exibirá os labels com percentuais (Corresponde à constante xlDataLabelsShowPercent da biblioteca do Excel)
Public Const EXCEL_EXIBE_LABELS_LEGENDA = 4 'Determina que o gráfico exibirá os labels com a legenda do gráfico (Corresponde à constante xlDataLabelsShowLabel da biblioteca do Excel)
Public Const EXCEL_EXIBE_LABELS_LEGENDA_PERCENTUAL = 5 'Determina que o gráfico exibirá os labels com a legenda do gráfico e os percentuais (Corresponde à constante xlDataLabelsShowLabelAndPercent da biblioteca do Excel)

'Constantes para posicionamento da legenda do gráfico
Public Const EXCEL_LEGENDA_ABAIXO = -4107
Public Const EXCEL_LEGENDA_ACIMA = -4160
Public Const EXCEL_LEGENDA_ESQUERDA = -4131
Public Const EXCEL_LEGENDA_DIREITA = -4152
Public Const EXCEL_LEGENDA_CORNER = 2
Public Const EXCEL_LEGENDA_NAO_EXIBE = -1

'Constantes com nomes de fontes para serem utilizadas no Excel
Public Const EXCEL_FONTE_TIMES_N_ROMAN As String = "TimesNewRoman "
Public Const EXCEL_FONTE_BOOKMAN As String = "Bookman "
Public Const EXCEL_FONTE_COURIER_NEW As String = "CourierNew "

'Constantes utilizadas para montegem do cabeçalho /  rodapé
Public Const EXCEL_CABECALHO_ESQUERDO = 1 'Cria uma seção esquerda de cabeçalho
Public Const EXCEL_CABECALHO_CENTRAL = 2 'Cria uma seção central de cabeçalho
Public Const EXCEL_CABECALHO_DIREITO = 3 'Cria uma seção direita de cabeçalho
Public Const EXCEL_RODAPE_ESQUERDO = 4 'Cria uma seção esquerda de rodapé
Public Const EXCEL_RODAPE_CENTRAL = 5 'Cria uma seção central de rodapé
Public Const EXCEL_RODAPE_DIREITO = 6 'Cria uma seção direita de rodapé
Public Const EXCEL_CABECALHO_RODAPE_NAO_QUEBRA_LINHA = -1 'Indica que não deve haver quebra de linha ao final da linha atual
Public Const EXCEL_CABECALHO_RODAPE_TAMANHO = 249 'Indica o tamanho máximo de uma string passada como cabeçalho ou rodapé de um documento excel
Public Const EXCEL_CABECALHO_RODAPE_ALINHAMENTO_ESQUERDA As String = " &L" 'Alinha o texto da seção do cabeçalho / rodapé pela esquerda
Public Const EXCEL_CABECALHO_RODAPE_ALINHAMENTO_CENTRAL As String = " &C" 'Centraliza o texto da seção do cabeçalho / rodapé
Public Const EXCEL_CABECALHO_RODAPE_ALINHAMENTO_DIREITA As String = "&R" 'Alinha o texto da seção do cabeçalho / rodapé pela direita
Public Const EXCEL_CABECALHO_RODAPE_NEGRITO As String = " &B" 'Formata o texto em negrito
Public Const EXCEL_CABECALHO_RODAPE_ITALICO As String = " &I" 'Formata o texto em itálico
Public Const EXCEL_CABECALHO_RODAPE_SUBLINHADO As String = " &U" 'Formata o texto com sublinhado
Public Const EXCEL_CABECALHO_RODAPE_PAGINA As String = " &P" 'Exibe a página no cabeçalho / rodapé


'*** EXCEL ***

'codigos de TiposDocInfo
Public Const DOCINFO_NFEEC = 1
Public Const DOCINFO_NFEECNT = 2
Public Const DOCINFO_NFEECP = 3
Public Const DOCINFO_NFEED = 4
Public Const DOCINFO_NFEEDC = 5
Public Const DOCINFO_NFEEDCNT = 6
Public Const DOCINFO_NFEEDD = 7
Public Const DOCINFO_NFEEDV = 8
Public Const DOCINFO_NFEEICM = 9
Public Const DOCINFO_NFEEIPI = 10
Public Const DOCINFO_NFEEN = 11
Public Const DOCINFO_NFEEOC = 12
Public Const DOCINFO_NFEEODC = 13
Public Const DOCINFO_NFEEODF = 14
Public Const DOCINFO_NFEEOF = 15
Public Const DOCINFO_NFEESR = 16
Public Const DOCINFO_NFIEC = 17
Public Const DOCINFO_NFIECNT = 18
Public Const DOCINFO_NFIECP = 19
Public Const DOCINFO_NFIED = 20
Public Const DOCINFO_NFIEDC = 21
Public Const DOCINFO_NFIEDCNT = 22
Public Const DOCINFO_NFIEDD = 23
Public Const DOCINFO_NFIEDV = 24
Public Const DOCINFO_NFIEICM = 25
Public Const DOCINFO_NFIEIPI = 26
Public Const DOCINFO_NFIEN = 27
Public Const DOCINFO_NFIEOC = 28
Public Const DOCINFO_NFIEODC = 29
Public Const DOCINFO_NFIEODF = 30
Public Const DOCINFO_NFIEOF = 31
Public Const DOCINFO_NFIESR = 32
Public Const DOCINFO_NFISC = 33
Public Const DOCINFO_NFISCNT = 34
Public Const DOCINFO_NFISCP = 35
Public Const DOCINFO_NFISD = 36
Public Const DOCINFO_NFISDC = 37
Public Const DOCINFO_NFISDCM = 38
Public Const DOCINFO_NFISDCNT = 39
Public Const DOCINFO_NFISDD = 40
Public Const DOCINFO_NFISICM = 41
Public Const DOCINFO_NFISIPI = 42
Public Const DOCINFO_NFISOC = 43
Public Const DOCINFO_NFISODC = 44
Public Const DOCINFO_NFISODF = 45
Public Const DOCINFO_NFISOF = 46
Public Const DOCINFO_NFISSR = 47
Public Const DOCINFO_NFISV = 48
Public Const DOCINFO_NFISVNE = 49
Public Const DOCINFO_PVN = 50
Public Const DOCINFO_NFISFS = 51
Public Const DOCINFO_NFISFV = 52
Public Const DOCINFO_NFISFIPI = 53
Public Const DOCINFO_NFISFICM = 54
Public Const DOCINFO_NFISFCP = 55
Public Const DOCINFO_NFEEFN = 56
Public Const DOCINFO_NFIEFN = 57
Public Const DOCINFO_NFEEFCP = 58
Public Const DOCINFO_NFEEFICM = 59
Public Const DOCINFO_NFEEFIPI = 60
Public Const DOCINFO_NFIEFCP = 61
Public Const DOCINFO_NFIEFICM = 62
Public Const DOCINFO_NFIEFIPI = 63
Public Const DOCINFO_NRCP = 64
Public Const DOCINFO_NRCC = 65
Public Const DOCINFO_NRFP = 66
Public Const DOCINFO_NRFF = 67
Public Const DOCINFO_NFISFVNE = 68
Public Const DOCINFO_NFISS = 70
Public Const DOCINFO_NFEEBF = 72
Public Const DOCINFO_NFIEBF = 73
Public Const DOCINFO_NFIEFBF = 74
Public Const DOCINFO_NFEEFBF = 75
Public Const DOCINFO_NFEEBEN = 76
Public Const DOCINFO_NFIEEBEN = 77
Public Const DOCINFO_NFEEFBEN = 78
Public Const DOCINFO_NFIEFBEN = 79
Public Const DOCINFO_NRFPCO = 80
Public Const DOCINFO_NRFFCO = 81
Public Const DOCINFO_NFEENCO = 82
Public Const DOCINFO_NFIENCO = 83
Public Const DOCINFO_NFEEFNCO = 84
Public Const DOCINFO_NFIEFNCO = 85
Public Const DOCINFO_NFIEIPICO = 86
Public Const DOCINFO_NFIECPCO = 87
Public Const DOCINFO_NFIEICMCO = 88
Public Const DOCINFO_NFEECPCO = 89
Public Const DOCINFO_NFEEICMCO = 90
Public Const DOCINFO_NFEEIPICO = 91
Public Const DOCINFO_NFIEFIPICO = 92
Public Const DOCINFO_NFIEFICMCO = 93
Public Const DOCINFO_NFIEFCPCO = 94
Public Const DOCINFO_NFEEFICMCO = 95
Public Const DOCINFO_NFEEFCPCO = 96
Public Const DOCINFO_NFEEFIPICO = 97
Public Const DOCINFO_NFISVPV = 98
Public Const DOCINFO_NFISFVPV = 99
Public Const DOCINFO_NFEEBFCOM = 100
Public Const DOCINFO_NFIEBFCOM = 101
Public Const DOCINFO_NFIEFBFCOM = 102
Public Const DOCINFO_NFEEFBFCOM = 103
Public Const DOCINFO_CFECT = 104
Public Const DOCINFO_CFEV = 105
Public Const DOCINFO_NFELUZ = 106
Public Const DOCINFO_NFEFTEL = 107
Public Const DOCINFO_NFES = 108
Public Const DOCINFO_NFEFS = 109
Public Const DOCINFO_NFETEL = 110
Public Const DOCINFO_NFEECS = 111
Public Const DOCINFO_NFIECS = 112
Public Const DOCINFO_NFEEFCS = 113
Public Const DOCINFO_NFIEFCS = 114
Public Const DOCINFO_NFICF = 115
Public Const DOCINFO_NFIFCF = 116
Public Const DOCINFO_NFFISPC = 117
Public Const DOCINFO_NFISPC = 118
Public Const DOCINFO_NFIEIMP = 119
Public Const DOCINFO_NFEE3BF = 120
Public Const DOCINFO_NFIE3BF = 121
Public Const DOCINFO_NFISBF = 122
Public Const DOCINFO_NFISFBF = 123
Public Const DOCINFO_NFISDBF = 124
Public Const DOCINFO_NFISDBFNE = 125
Public Const DOCINFO_NFISDSC = 126
Public Const DOCINFO_NFEEFA = 127
Public Const DOCINFO_NFEEA = 128
Public Const DOCINFO_NFEEDSC = 129
Public Const DOCINFO_NFIEDSC = 130
Public Const DOCINFO_NFISRB = 131
Public Const DOCINFO_NFIEDSB = 132
Public Const DOCINFO_NFEEDSB = 133
Public Const DOCINFO_NFIEDB = 134
Public Const DOCINFO_NFEEDB = 135
Public Const DOCINFO_NFIRS = 136
Public Const DOCINFO_NFERS = 137
Public Const DOCINFO_NFSRS = 138
Public Const DOCINFO_NFIERCICMS = 139
Public Const DOCINFO_NFEERCIMS = 140
Public Const DOCINFO_NFIERCIPI = 141
Public Const DOCINFO_NFEERCIPI = 142
Public Const DOCINFO_NFISRCICMS = 143
Public Const DOCINFO_NFISRCIPI = 144
Public Const DOCINFO_NFEERCP = 145
Public Const DOCINFO_NFIERCP = 146
Public Const DOCINFO_NFISRCP = 147
Public Const DOCINFO_NFISCPDC = 148
Public Const DOCINFO_NFISRMB3 = 149
Public Const DOCINFO_NFISDPV = 150
Public Const DOCINFO_NFISCNTPV = 151
Public Const DOCINFO_NFISCPV = 152
Public Const DOCINFO_NFISSRPV = 153
Public Const DOCINFO_NFISOCPV = 154
Public Const DOCINFO_NFISOFPV = 155
Public Const DOCINFO_NFISRMB3PV = 156
Public Const DOCINFO_NFIERMB3 = 157
Public Const DOCINFO_NFEERMB3 = 158
Public Const DOCINFO_NFEFSME = 159
Public Const DOCINFO_NFIFSME = 160
Public Const DOCINFO_NFESME = 161
Public Const DOCINFO_NFISME = 162
Public Const DOCINFO_NFIEDPC = 163
Public Const DOCINFO_NFEEDPC = 164
Public Const DOCINFO_DDAI = 165
Public Const DOCINFO_FDDAI = 166
Public Const DOCINFO_NFEECCP = 167
Public Const DOCINFO_NFEESAC = 168
Public Const DOCINFO_NFFEESAC = 169
Public Const DOCINFO_CFECC = 170
Public Const DOCINFO_NFISSRF = 171
Public Const DOCINFO_NFICFCICMS = 172
Public Const DOCINFO_NFEESRCOM = 175
Public Const DOCINFO_NFEEFCSCOM = 176
Public Const DOCINFO_NFEFSMECOM = 177
Public Const DOCINFO_NFESMECOM = 178
Public Const DOCINFO_NFIESRCLI = 179
Public Const DOCINFO_NFEESRCLI = 180
Public Const DOCINFO_NFISRVFE = 182
Public Const DOCINFO_NFISFVFE = 183
Public Const DOCINFO_NFIEDVRFE = 184
Public Const DOCINFO_NFESCO = 185
Public Const DOCINFO_NFEFSCO = 186
Public Const DOCINFO_NFIFVEFPV = 187
Public Const DOCINFO_NFIFVEF = 188
Public Const DOCINFO_NFISSRPVEF = 189
Public Const DOCINFO_NFISSREF = 190
Public Const DOCINFO_NFEEFNCO3 = 191
Public Const DOCINFO_NFEEDAR = 192
Public Const DOCINFO_NFEEDVICM = 193
Public Const DOCINFO_CFFECT = 194
Public Const DOCINFO_CFFEV = 195
Public Const DOCINFO_NFEEODFSE = 196
Public Const DOCINFO_NFISOFSE = 197
Public Const DOCINFO_NFISCNTCNT = 198
Public Const DOCINFO_NFIEDCNTCT = 199
Public Const DOCINFO_NFEEDCNTCT = 200
Public Const DOCINFO_NFISOUOU = 201
Public Const DOCINFO_NFIEDOUOU = 202
Public Const DOCINFO_NFEEDOUOU = 203
Public Const DOCINFO_NFISFBFPV = 204
Public Const DOCINFO_NFISSV = 205
Public Const DOCINFO_NFEEFBES = 206
Public Const DOCINFO_NFIEFBES = 207
Public Const DOCINFO_NFFEGAS = 208
Public Const DOCINFO_NFFEAGUA = 209
Public Const DOCINFO_NFISVFE = 210
Public Const DOCINFO_NFFEESACCO = 211
Public Const DOCINFO_FDDAICO = 212
Public Const DOCINFO_NFEEFBENCO = 213
Public Const DOCINFO_NFEEFACO = 214
Public Const DOCINFO_NFFPSRV = 215
Public Const DOCINFO_PSRVN = 216
Public Const DOCINFO_OSRVN = 217
Public Const DOCINFO_NFFGSRV = 218
Public Const DOCINFO_NFFISRS = 219
Public Const DOCINFO_NFEEDVSM = 220
Public Const DOCINFO_NFIEDVSM = 221
Public Const DOCINFO_NFEEFCNT = 222
Public Const DOCINFO_NFSRSF = 223
Public Const DOCINFO_NFFBFRS = 224
Public Const DOCINFO_NFEFBFRS = 225
Public Const DOCINFO_NFISSRGOV = 226
Public Const DOCINFO_NFEEFCNT3 = 227
Public Const DOCINFO_NFISRCICMF = 228
Public Const DOCINFO_NFISRCIPIF = 229
Public Const DOCINFO_NFISRCPF = 230
Public Const DOCINFO_NFISFSPV = 231
Public Const DOCINFO_NFISSPV = 232
Public Const DOCINFO_NFIFVETPV = 233
Public Const DOCINFO_NFEEDSO = 234
Public Const DOCINFO_NFIEDSO = 235
Public Const DOCINFO_NFISRETPV = 236
Public Const DOCINFO_NFEEDVIPI = 237
Public Const DOCINFO_NFSRCO = 238       '   Compra à Ordem - Remessa Adquirente/Destinatário
Public Const DOCINFO_NFECCO = 239       '   Compra p/ entreg. em 3os por conta e ordem
Public Const DOCINFO_NFECCOPC = 240     '   Compra p/ entreg. em 3os por conta e ordem - Pedido
Public Const DOCINFO_NFEFCCO = 241      '   Compra fatura p/ entreg. em 3os por conta e ordem
Public Const DOCINFO_NFEFCCOPC = 242    '   Compra fatura p/ entreg. em 3os por conta e ordem - Pedido
Public Const DOCINFO_NFERICO = 243      '   Remessa p/benef Vendedor/Destinatário-Compra à Ordem (Cadastrada pelo Adquirente)
Public Const DOCINFO_NFSVCO = 244       '   Venda por conta e ordem
Public Const DOCINFO_NFSVCOPV = 245     '   Venda por conta e ordem - Pedido
Public Const DOCINFO_NFSFVCO = 246      '   Venda fatura por conta e ordem
Public Const DOCINFO_NFSFVCOPV = 247    '   Venda fatura por conta e ordem - Pedido
Public Const DOCINFO_NFSRCOC = 248      '   Remessa por conta e ordem para Cliente
Public Const DOCINFO_NFSRCOF = 249      '   Remessa por conta e ordem para Fornecedor
Public Const DOCINFO_NFERICOD = 250     '   Remessa p/benef Vendedor/Destinatário-Compra à Ordem (Cadastrada pelo Destinatário)
Public Const DOCINFO_NFERCOD = 251      '   Remessa Adquirente/Destinatário-Compra à Ordem
Public Const DOCINFO_NFIEIMPSE = 252
Public Const DOCINFO_NFIEIMPPC = 253
Public Const DOCINFO_NFIEDCSM = 254
Public Const DOCINFO_NFEEDCSM = 255
Public Const DOCINFO_NFIEDVICM = 256
Public Const DOCINFO_NFIEIMPCC = 257
Public Const DOCINFO_NFELUZCO = 258
Public Const DOCINFO_NFISFVSEPV = 259
Public Const DOCINFO_NFISFVPAF = 260
Public Const DOCINFO_NFISSAT = 261 'sat-cfe
Public Const DOCINFO_NFFPSRVS = 262
Public Const DOCINFO_NFFPSRVP = 263
Public Const DOCINFO_NFEFTELCO = 264    'Nota Fiscal Fatura Serv Telecom-Compras
Public Const DOCINFO_NFIERICMSC = 265   'NF interna de entrada remessa de compl de ICMS de Cliente
Public Const DOCINFO_NFEERICMSC = 266   'NF externa de entrada remessa de compl. de ICMS de Cliente
Public Const DOCINFO_NFCEPDV = 267 'NFCe - nota fiscal de consumidor eletronica ( vinda do ecf )
Public Const DOCINFO_NFISNFCE = 268 'NFCe - nota fiscal de consumidor eletronica ( gravada diretamente na retaguarda )
Public Const DOCINFO_NFISRAG = 269 'NF interna de saida para remessa de amostra grátis
Public Const DOCINFO_NFISRAGPV = 270 'NF interna de saida para remessa de amostra grátis via pedido de venda
Public Const DOCINFO_NFISCIPIDC = 271 'NF interna de saida para complemento de ipi de devolucao de compra
Public Const DOCINFO_NFISASRVTR = 272 'NF interna de saida anulação serviço transporte
Public Const DOCINFO_TNFISFV = 273 'Teste para cálculo de trib de NF de venda sem cliente
Public Const DOCINFO_NFIEFCNFP = 274 'Contra Nota de Produtor Rural

Public Const DOCINFO_NFSET = 275 'NFSe de Transporte
Public Const DOCINFO_NFSEFT = 276 'NFSe Fatura de Transporte


'Periodicidade para o Cálculo
Public Const PERIODICIDADE_LIVRE = 1
Public Const PERIODICIDADE_SEMANAL = 2
Public Const PERIODICIDADE_DECENDIAL = 3
Public Const PERIODICIDADE_QUINZENAL = 4
Public Const PERIODICIDADE_MENSAL = 5
Public Const PERIODICIDADE_BIMESTRAL = 6
Public Const PERIODICIDADE_TRIMESTRAL = 7
Public Const PERIODICIDADE_QUADRIMESTRAL = 8
Public Const PERIODICIDADE_SEMESTRAL = 9
Public Const PERIODICIDADE_ANUAL = 10

Public Const SQL_ORD_PARAM_ESQ_DIR = 1 'os ?s devem ser passados da esquerda p/a direita do comando
Public Const SQL_ORD_PARAM_EXECUCAO = 2 'os ?s de comandos aninhados (c/em ...EXISTS (SELECT...?...)) devem ser passados antes dos ?s mais externos

'Erro gerado pela função gerror
Public Const ERRO_FUNCAO_GERROR = 4999

Public Const CODUSUARIO_SUPERVISOR As String = "supervisor"

'Indica se o usuário está logado ou não
Public Const USUARIO_LOGADO = 1
Public Const USUARIO_NAO_LOGADO = 0

'Indica se linha do grid está selecionada ou não
Public Const Selecionado = 1
Public Const NAO_SELECIONADO = 0

'Módulo liberado
Public Const MODULO_LIBERADO = 1
Public Const MODULO_NAO_LIBERADO = 0

'Tipo de Acesso
Public Const COM_ACESSO = 1
Public Const SEM_ACESSO = 0

Public Const GRUPO_SUP = "Supervisor"
Public Const REL_ORIGEM_FORPRINT = 0
Public Const REL_ORIGEM_USUARIO = 1

'Tipo de Versao
Public Const VERSAO_LIGHT = 0
Public Const VERSAO_FULL = 1

'Constantes utilizadas para habilitar ou desabilitar telas
Public Const HABILITA_TELA = 1
Public Const DESABILITA_TELA = 0

'Indica que o campo deve participar do teste de integridade
Public Const CAMPO_TESTA_INTEGRIDADE = 1

'Alinhamento dos Campos dos Browsers
Public Const BROWSER_CAMPO_ALINHAMENTO_ESQUERDA = 0
Public Const BROWSER_CAMPO_ALINHAMENTO_DIREITA = 1

Public Const BROWSER_NUM_MAX_CAMPOS = 200

Public Const STRING_CONTEUDO = 255

'tamanhos maximos dos campos das tabelas config.
Public Const STRING_CONFIG_CODIGO = 50
Public Const STRING_CONFIG_DESCRICAO = 150


Public Const STRING_NOME_ARQUIVO = 50
Public Const STRING_CLASSE_OBJETO = 50
Public Const STRING_NOME_OBJETO_MSG = 50

'Strings de Apropriacao
Public Const STRING_PRODUTO_CUSTO_MEDIO = "Custo Médio"
Public Const STRING_PRODUTO_CUSTO_STANDARD = "Custo Standard"
Public Const STRING_PRODUTO_CUSTO_INFORMADO = "Informado"
Public Const STRING_PRODUTO_CUSTO_PRODUCAO = "Custo de Produção"

'Strings referentes as tabelas que farão o funcionamento da combo de opções genérica
Public Const STRING_TELASCAMPOS_NOMETELA = 50
Public Const STRING_TELASCAMPOS_TITULOTELA = 50
Public Const STRING_TELASCAMPOS_NOMECAMPO = 50
Public Const STRING_TELASOPCOESVALORES_NOMEOPCAO = 50
Public Const STRING_TELASOPCOESVALORES_VALORCAMPO = 50

'Constantes de Apropriacao
Public Const PRODUTO_CUSTO_MEDIO = 0
Public Const PRODUTO_CUSTO_STANDARD = 1
Public Const PRODUTO_CUSTO_INFORMADO = 2
Public Const PRODUTO_CUSTO_PRODUCAO = 3

'Constantes de Controle de Estoque
Public Const PRODUTO_RESERVA_ESTOQUE = 1
Public Const PRODUTO_ESTOQUE = 2
Public Const PRODUTO_SEM_ESTOQUE = 3

'Strings de Controle de Estoque
Public Const STRING_PRODUTO_RESERVA_ESTOQUE = "Reserva + Estoque"
Public Const STRING_PRODUTO_ESTOQUE = "Estoque"
Public Const STRING_PRODUTO_SEM_ESTOQUE = "Sem controle de estoque"

'Constantes de Gerencial
Public Const PRODUTO_GERENCIAL = 1
Public Const PRODUTO_FINAL = 0

'Strings de Gerencial
Public Const STRING_PRODUTO_GERENCIAL = "Gerencial"
Public Const STRING_PRODUTO_FINAL = "Final"

'Constantes de Reserva
Public Const RESERVA_MANUT_RESERVA = 0
Public Const RESERVA_PEDIDO = 1

'Strings de Reserva
Public Const STRING_MANUT_RESERVA = "Manut. Reserva"
Public Const STRING_RESERVA_PEDIDO = "Pedido"

'Constante referente a tecla que vai acionar a geracao do proximo codigo para lancamento, nota fiscal, movimento de estoque, etc.
Public Const KEYCODE_PROXIMO_NUMERO = vbKeyF2
Public Const KEYCODE_BROWSER = vbKeyF3
Public Const KEYCODE_BOTAOCONSULTA = vbKeyF5
Public Const KEYCODE_REPETE_LINHA_GRID = vbKeyF6
Public Const KEYCODE_REPETE_CAMPO_GRID = vbKeyF7
Public Const KEYCODE_CODBARRAS = vbKeyF8
Public Const KEYCODE_LOCALIZAGRID = vbKeyF9
Public Const KEYCODE_EXPORTAGRIDEXCEL = vbKeyF10
Public Const KEYCODE_IDIOMAS = vbKeyF11

'Constantes dos tipos de controles do VB
Public Const CONTROLE_OPTIONBUTTON = "OptionButton"
Public Const CONTROLE_CHECKBOX = "CheckBox"
Public Const CONTROLE_MASKEDBOX = "MaskEdBox"
Public Const CONTROLE_COMBOBOX = "ComboBox"

'Constantes das teclas
Public Const TECLA_ENTER = "{Enter}"
Public Const TECLA_ESC = "{Esc}"

Public Const TITULO_TELA_PRINCIPAL = "Corporator - Sistema de Gestão Empresarial"

'Layout p/ Bancos
Public Const NUM_MAX_CARACTERES_NOME_RELATORIO = 8
Public Const NOME_EXTENSAO_RELATORIO = ".TSK" 'extensao do nome dos relatorios


'Constantes que identificam o Tipo do NumIntDocOrigem dos Movimentos de Estoque
Public Const MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL = 1
Public Const MOVEST_TIPONUMINTDOCORIGEM_NFISCAL = 2
Public Const MOVEST_TIPONUMINTDOCORIGEM_INVENTARIO = 3
Public Const MOVEST_TIPONUMINTDOCORIGEM_MOVESTOQUE = 4
Public Const MOVEST_TIPONUMINTDOCORIGEM_ITEMOP = 5
Public Const MOVEST_TIPONUMINTDOCORIGEM_CUPOMFISCAL = 6
Public Const MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE = 7
Public Const MOVEST_TIPONUMINTDOCORIGEM_ITEMOS = 8
Public Const MOVEST_TIPONUMINTDOCORIGEM_ESTORNO_CRED_TRIB_ICMS = 9

Public Const USUARIO_ATIVO = 1
Public Const USUARIO_INATIVO = 0

'Produto
Public Const PRODUTO_COMPRAVEL = 1
Public Const PRODUTO_NAO_COMPRAVEL = 0 '??? Resolver se será apagado
Public Const PRODUTO_PRODUZIVEL = 0
Public Const PRODUTO_VENDAVEL = 1
Public Const PRODUTO_NAO_VENDAVEL = 0
Public Const PRODUTO_PCP_NAOPODE = 0
Public Const PRODUTO_PCP_PODE = 1
Public Const PRODUTO_ATIVO = 0
Public Const PRODUTO_INATIVO = 1

'Código da condição de pagamento a vista
Public Const COD_A_VISTA = 1

'Utilização de Centro de Custo/Lucro no sistema
Public Const CCL_NAO_USA = 0
Public Const CCL_USA_CONTABIL = 1
Public Const CCL_USA_EXTRACONTABIL = 2

Public Const DOC_INICIALIZADO_POR_PERIODO = 1
Public Const DOC_INICIALIZADO_POR_EXERCICIO = 0
Public Const LOTE_INICIALIZADO_POR_PERIODO = 1
Public Const LOTE_INICIALIZADO_POR_EXERCICIO = 0

Public Const SEGMENTO_VAZIO = 1
Public Const SEGMENTO_CHEIO = 0

Public Const SEGMENTOS_RESTANTES_VAZIOS = 1
Public Const SEGMENTOS_RESTANTES_CHEIOS = 0

'Constantes limpeza de tela
Public Const TELA_LIMPA = 1
Public Const TELA_NAO_LIMPA = 0

'Constantes p/ IPI - contribuição
Public Const IPI_NAO_CONTRIBUINTE = 0
Public Const IPI_CONTRIBUINTE = 1
Public Const IPI_CONTRIBUINTE_50 = 2

'Constantes para a tabela ModuloFilEmp
Public Const NAO_CONFIGURADO = 0
Public Const CONFIGURADO = 1

'Tipos para as tabelas Config
Public Const CONFIG_TIPO_TEXTO = 0
Public Const CONFIG_TIPO_INTEIRO = 1
Public Const CONFIG_TIPO_LONG = 2
Public Const CONFIG_TIPO_DOUBLE = 3
Public Const CONFIG_TIPO_DATA = 4

'KIT
Public Const KIT_BASICO = 1
Public Const NAO_KIT_BASICO = 0
Public Const KIT_INTERMEDIARIO = 1
Public Const NAO_KIT_INTERMEDIARIO = 0

Public Const USUARIO_TEM_ACESSO = 1

Public Const ADM_CONFIGURA_PESQUISA = 1
Public Const ADM_CONFIGURA_NORMAL = 0

Public Const BD_ADM = 0
Public Const BD_DICDADOS = 1

'Para informar ao Evento Click de Tipo se obj está sendo
'carregado na Tela. Neste caso o código do Click é bypassado.
Public Const TRAZENDO_DADOS_TELA As Boolean = True

'indica que os netos de um nó de uma árvore(TreeView) já estão carregados
Public Const NETOS_NA_ARVORE As String = "1"

'Produto Gerencial -> Abstrato ou não
Public Const GERENCIAL = 1
Public Const NAO_GERENCIAL = 0

'Número máximo de argumentos
'Public Const MAX_ARGS_BATCH = 10

Public Const CANCELA_BATCH = 1
Public Const STRING_BUFFER_MAX_TEXTO = 256

Public Const NUM_MAX_ITENS_FORMACAOPRECO = 100
Public Const NUM_MAX_MNEMONICOFPRECO = 100


'Titulo da parte de contabilidade nas tabsstrips
Public Const TITULO_CONTABILIZACAO As String = "Contabilização"
Public Const TITULO_CONTABILIZACAO_RESUMIDO As String = "Contab"

Public Const SISTEMA_SGE As String = "SGE"

'Siglas dos modulos tal qual cadastrado na tabela de Modulos
Public Const MODULO_CRM As String = "CRM" ' Incluído por Luiz Nogueira em 13/01/04
Public Const MODULO_PCP As String = "PCP"
Public Const MODULO_FOLHA As String = "FLH"
Public Const MODULO_LOJA As String = "LJ"
Public Const MODULO_CONTABILIDADE As String = "CTB"
Public Const MODULO_CONTASAPAGAR As String = "CP"
Public Const MODULO_CONTASARECEBER As String = "CR"
Public Const MODULO_TESOURARIA As String = "TES"
Public Const MODULO_FATURAMENTO As String = "FAT"
Public Const MODULO_ESTOQUE As String = "EST"
Public Const MODULO_COMPRAS As String = "COM"
Public Const MODULO_LIVROSFISCAIS As String = "FIS"
Public Const MODULO_ADM As String = "ADM"
Public Const MODULO_BATCHCONTASAPAGAR As String = "BCP"
Public Const MODULO_BATCHCONTASARECEBER As String = "BCR"
Public Const MODULO_CUSTOCP As String = "CCP"
Public Const MODULO_CUSTOCR As String = "CCR"
Public Const MODULO_CUSTOEST As String = "CES"
Public Const MODULO_CUSTOFAT As String = "CFT"
Public Const MODULO_CUSTOTES As String = "CTS"
Public Const MODULO_CUSTOCOM As String = "CCM"
Public Const MODULO_QUALIDADE As String = "QUA"
Public Const MODULO_PONTO_DE_VENDA As String = "ECF"
Public Const MODULO_PROJETO As String = "PRJ"
Public Const MODULO_SERVICOS As String = "SRV"

'Origem dos lançamentos de custo
Public Const ORIGEM_CUSTO_ As String = "CST"

'Indicação do uso de Modulo
Public Const USA_MODULO = 1
Public Const NAO_USA_MODULO = 0
Public Const MODULO_ATIVO = 1

'Indica onde é chamada a função Rotina_Grid_Enable
Public Const ROTINA_GRID_CLICK = 1
Public Const ROTINA_GRID_ENTRADA_CELULA = 2
Public Const ROTINA_GRID_ABANDONA_CELULA = 3
Public Const ROTINA_GRID_SCROLL = 4
Public Const ROTINA_GRID_TRATA_TECLA_CAMPO2 = 5
Public Const ROTINA_GRID_TRATA_TECLA = 6

'Status da preparação do relatório para execução
Public Const RELATORIO_OK = 0
Public Const RELATORIO_CANCELA = 1

'Produto vazio ou preenchido
Public Const PRODUTO_VAZIO = 0
Public Const PRODUTO_PREENCHIDO = 1

'Item vazio ou preenchido
Public Const ITEM_VAZIO = 0
Public Const ITEM_PREENCHIDO = 1


'Controle Estoque Produto
Public Const PRODUTO_CONTROLE_RESERVA = 1 'para itens que serao faturados somente apos reserva no almoxarifado
Public Const PRODUTO_CONTROLE_ESTOQUE = 2 'para itens que serao faturados sem necessidade de reserva
Public Const PRODUTO_CONTROLE_SEM_ESTOQUE = 3 'para itens nao "estocaveis". Ex. Servico.


'Tipo do Segmento da Conta
Global Const SEGMENTO_NUMERICO = 1
Global Const SEGMENTO_ALFANUMERICO = 2
Global Const SEGMENTO_ASCII = 3


'Código de Filial para enxergar TODAS as filiais
Public Const EMPRESA_TODA = 0
Public Const FILIAL_EMPRESA = 1
Public Const FILIAL_EMPRESA_TODA = 2

Public Const EMPRESA_TODA_NOME = "<Empresa Toda>"

'A Matriz é uma filial (com código=1)
'Para cadastro de Cliente, Fornecedor, Empresa (versão Light)
Public Const FILIAL_MATRIZ = 1
Public Const MATRIZ = "Matriz"

Public Const ENTER_KEY = 13

Public Const OK = 1
Public Const CANCELA = 0

'Para tabelas que tem campo Ativo
Public Const Ativo = 0
Public Const Inativo = 1

'Para tabelas que tem campo Excluido
Public Const NAO_EXCLUIDO = 0
Public Const EXCLUIDO = 1

'Data nula no VB. A data é armazenada como nulo mas é transformada para o VB na constante abaixo.
Public Const DATA_NULA As Date = #9/7/1822#
Public Const DATA_MAX As Date = #12/31/2100#

'Declara array de variáveis globais tipo variant para BIND das setas
'Estas variáveis são passadas para o Comando_Executar2 como parâmetros depois de inicializadas e dimensionadas
'Precisam estar aqui para poderem ser utilizadas por PROJETOS diferentes e terem
'PERMANENCIA além dos métodos que usam BIND
Public vCampoSelect() As Variant  'Passado para bindar no ComandoExecutar2

'Separador entre conta e descricao de conta na TreeView
Public Const SEPARADOR As String = "-"

'Separador de parâmetros em funções da máquina de expressões (Ex.: Padrões de Contabilização)
Public Const SEPARADOR_FORMULA_FUNCAO As String = ","

'Caracter inicial para KEY de TreeView.
'Tem o segundo para poder ter segundo nível com mesma chave
Public Const KEY_CARACTER As String = "a"
Public Const KEY_CARACTER2 As String = "b"

'Tamanhos de CPF, CGC
Public Const STRING_CPF = 11
Public Const STRING_CGC = 14
Public Const STRING_ID_ESTRANGEIRO = 20
Public Const STRING_RG = 15
Public Const STRING_PREFIXO_CGC = 8


'Strings de DicConfig
Public Const STRING_DICCONFIG_SENHA = 25
Public Const STRING_DICCONFIG_SERIE = 50

Public Const STRING_CODUSUARIO = 10
Public Const STRING_USUARIO_CODIGO = 10
Public Const STRING_USUARIO_NOME = 50
Public Const STRING_USUARIO_NOMEREDUZIDO = 20
Public Const STRING_USUARIO_SENHA = 10
'Public Const STRING_USUARIO_CONEXAO = 255

'Comparacao
Public Const IGUAL = 0
Public Const MAIOR = 1
Public Const MENOR = -1
Public Const DIFERENTE = 1

'Tamanhos de codigos inteiro e long para concatenação nas COMBOS
Public Const TAMANHO_INTEIRO = 4
Public Const TAMANHO_LONG = 5

'Inteiro e Long máximos
Public Const MAXIMO_INTEIRO = 9999
Public Const MAXIMO_LONG = 99999999

'Variáveis que guardam valores da classe AdmSeta
'INDEPENDENTEMENTE da instância (PROJETOS diferentes chamando).
'São acessadas ATRAVÉS de AdmSeta
Public obj_ST_TelaPrincipal As Object  'Tela Principal
Public i_ST_SetaIgnoraClick As Integer 'Flag se é para ignorar o click de seta
Public l_ST_ComandoSeta As Long        'ComandoSeta utilizado no SELECT das setas
Public s_ST_TelaSetaClick As String    'Nome da tela em que houve o último SELECT das setas
Public obj_ST_TelaAtiva As Object      'Última tela sujeita a setas a receber foco
Public i_ST_Ordem_StrCmp As Integer    'Ordem a ser usada em StrComp

'Constante de erro de Chama_Tela1
Public lErro_Chama_Tela1 As Long

'Tipo de tributação ICMS
Public Const TIPOTRIBICMS_SITUACAOTRIBECF_NAO_TRIBUTADO = "N"
Public Const TIPOTRIBICMS_SITUACAOTRIBECF_ISENTA = "I"
Public Const TIPOTRIBICMS_SITUACAOTRIBECF_TRIB_SUBST = "F"
Public Const TIPOTRIBICMS_SITUACAOTRIBECF_INTEGRAL = "T"
Public Const TIPOTRIBISS_SITUACAOTRIBECF_NAO_TRIBUTADO = "NS"
Public Const TIPOTRIBISS_SITUACAOTRIBECF_ISENTA = "IS"
Public Const TIPOTRIBISS_SITUACAOTRIBECF_TRIB_SUBST = "FS"
Public Const TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL = "S"

Public Const TIPOTRIBISS_SITUACAOTRIBECF_NAO_TRIBUTADO_DESC = "ISS Não Tributado"
Public Const TIPOTRIBISS_SITUACAOTRIBECF_ISENTA_DESC = "ISS Isenta"
Public Const TIPOTRIBISS_SITUACAOTRIBECF_TRIB_SUBST_DESC = "ISS Cobrado anteriormente por subst."
Public Const TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL_DESC = "ISS Integral"


Public Const TIPOTRIBICMS_NAO_TRIBUTADO = 0
Public Const TIPOTRIBICMS_ISENTA = 2
Public Const TIPOTRIBICMS_TRIB_SUBST = 8
Public Const TIPOTRIBICMS_INTEGRAL = 1
Public Const TIPOTRIBISS_NAO_TRIBUTADO = 1001
Public Const TIPOTRIBISS_ISENTA = 1002
Public Const TIPOTRIBISS_TRIB_SUBST = 1003
Public Const TIPOTRIBISS_INTEGRAL = 1004


'Tipos de dispositivo de saida para relatorio
Public Const REL_SAIDA_PREVIA = -1
Public Const REL_SAIDA_IMPRESSORA = 0
Public Const REL_SAIDA_VIDEO = 1
Public Const REL_SAIDA_ARQUIVO = 2
Public Const REL_SAIDA_EMAIL = 3

'Definicoes envolvendo relatorios
Public Const STRING_CODIGO_RELATORIO = 50
Public Const STRING_AUTOR_RELATORIO = 50
Public Const STRING_DESC_RELATORIO = 150
Public Const STRING_CODIGO_OPCAO_RELATORIO = 50
Public Const STRING_REL_PARAM = 20
Public Const STRING_REL_PARAM_VALOR = 255

Public Const STRING_FERIADO = 30

Public Const STRING_NOME_ARQ_REIMP = 150

Global Const STRING_CONTA = 20

Global Const STRING_SEGMENTO_TIPO = 20
Global Const STRING_SEGMENTO_DELIMITADOR = 1

Public Const STRING_CPRCONFIG_CONTEUDO = 255

'Para leitura no BD

'FilialEmpresa
Public Const STRING_RAMO_EMPRESA = 80
Public Const STRING_JUCERJA = 20
Public Const STRING_CONTADOR = 50
Public Const STRING_CRC_CONTADOR = 20
Public Const STRING_CPF_CONTADOR = 14
Public Const STRING_CERTIFICADOA1A3 = 50

'Produto
'Public Const STRING_PRODUTO_NOME_REDUZIDO = 20
'Public Const STRING_PRODUTO_MODELO = 40
Public Const STRING_PRODUTO_CODIGO_BARRAS = 15
Public Const STRING_PRODUTO_COR = 20
Public Const STRING_PRODUTO_OBS_FISICA = 500
Public Const STRING_PRODUTO_IPI_CODIGO = 10
Public Const STRING_PRODUTO_IPI_COD_DIPI = 3
Public Const STRING_PRODUTO_ISS_CODIGO = 10
Public Const STRING_PRODUTO_ISS_CODSERV = 20
'Public Const STRING_PRODUTO_REFERENCIA = 20
Public Const STRING_PRODUTO_FIGURA = 80

'??? serao substituidas por STRING_UM_SIGLA
Public Const STRING_PRODUTO_SIGLAUMESTOQUE = 5
Public Const STRING_PRODUTO_SIGLAUMCOMPRA = 5
Public Const STRING_PRODUTO_SIGLAUMVENDA = 5

Public Const STRING_SERIE_RASTREAMENTO = 20

'??? serao substituidas por STRING_PRODUTO
Public Const STRING_PRODUTO_SUBSTITUTO1 = 20
Public Const STRING_PRODUTO_SUBSTITUTO2 = 20

Public Const STRING_PRODUTO = 20
Public Const STRING_PRODUTO_DESCRICAO = 250 '45 Por Daniel em 23/05/2002
Public Const STRING_PRODUTO_GENERO = 2

Public Const STRING_FCI = 36 'ficha de controle de importação
Public Const STRING_NBS = 9
Public Const STRING_CEST = 7

Public Const STRING_DOCORIGEM = 50

'Definição dos campos das tabelas Config
Public Const STRING_CONFIG_CONTEUDO = 255

Public Const ADM_TIPO_BIT = 0
Public Const ADM_TIPO_SMALLINT = 1
Public Const ADM_TIPO_INTEGER = 2
Public Const ADM_TIPO_DOUBLE = 3
Public Const ADM_TIPO_VARCHAR = 4
Public Const ADM_TIPO_LONGVARCHAR = 5
Public Const ADM_TIPO_DATE = 6
Public Const ADM_TIPO_TIME = 7
Public Const ADM_TIPO_TIMESTAMP = 8


'Subtipos usados por telas de browser (String)
Public Const ADM_SUBTIPO_SIM As String = "Sim"
Public Const ADM_SUBTIPO_NAO As String = "Não"
Public Const ADM_PARTICIPA_COMPRAS As String = "Pode ser Comprado"
Public Const ADM_NAO_PARTICIPA_COMPRAS As String = " Não pode ser Comprado"
Public Const ADM_GERENCIAL As String = "Gerencial"
Public Const ADM_ANALITICO As String = "Analitico"
Public Const ADM_TIPO_MATERIAPRIMA As String = "Matéria Prima"
Public Const ADM_TIPO_INTERMEDIARIO As String = "Produto Intermediário"
Public Const ADM_TIPO_EMBALAGEM As String = "Embalagem"
Public Const ADM_TIPO_ACABADO As String = "Produto acabado"
Public Const ADM_TIPO_REVENDA As String = "Produto p/ revenda"
Public Const ADM_TIPO_REPARO  As String = "Produto p/ reparo e etc."
Public Const ADM_TIPO_OUTROS As String = "Outros"
Public Const ADM_TIPO_AJUDACUSTO_FIXA As String = "Fixa"
Public Const ADM_TIPO_AJUDACUSTO_MINIMA As String = "Mínima"

'*****************************************************
'Alteracao Daniel em 11/07/2002
'Constantes de SubTipo
Public Const ADM_TIPO_KIT_SITUACAO_PADRAO As String = "Padrão"
Public Const ADM_TIPO_KIT_SITUACAO_ATIVO As String = "Ativo"
Public Const ADM_TIPO_KIT_SITUACAO_INATIVO As String = "Inativo"
'*****************************************************

'Estes subtipos são utilizados pelas telas de browser
Public Const ADM_SUBTIPO_CONTA = 1
Public Const ADM_SUBTIPO_CCL = 2
Public Const ADM_SUBTIPO_EXERCICIO = 3
Public Const ADM_SUBTIPO_PERIODO = 4
Public Const ADM_SUBTIPO_PERCENTUAL = 5
Public Const ADM_SUBTIPO_PRODUTO = 6
Public Const ADM_SUBTIPO_TIPOCONTACCL = 7
Public Const ADM_SUBTIPO_NATUREZACONTA = 8
Public Const ADM_SUBTIPO_NATUREZAPRODUTO = 9
Public Const ADM_SUBTIPO_SITUACAO = 10
Public Const ADM_SUBTIPO_DESTINACAO = 11
Public Const ADM_SUBTIPO_TIPOMEIOPAGTO = 12
Public Const ADM_SUBTIPO_CONTROLE_ESTOQUE = 13
Public Const ADM_SUBTIPO_GERENCIAL = 14
Public Const ADM_SUBTIPO_TIPODOC_RESERVA = 15
Public Const ADM_SUBTIPO_APROPRIACAO = 16
Public Const ADM_SUBTIPO_TIPORATEIO = 17
Public Const ADM_SUBTIPO_AGLUTINA = 18
Public Const ADM_SUBTIPO_PRECADASTRADO = 19
Public Const ADM_SUBTIPO_SECAO = 20
Public Const ADM_SUBTIPO_STATUS = 21
Public Const ADM_SUBTIPO_FORNECEDOR = 22
Public Const ADM_SUBTIPO_FILIAL = 23
Public Const ADM_SUBTIPO_CONDPAGTO = 24
Public Const ADM_SUBTIPO_COMPRADOR = 25
Public Const ADM_SUBTIPO_HORA = 26
Public Const ADM_SUBTIPO_LIBERADO = 27
Public Const ADM_SUBTIPO_SIMNAO = 28
Public Const ADM_SUBTIPO_FILIALEMPRESA = 29
'Public Const ADM_SUBTIPO_TIPODESTINO = 30
Public Const ADM_SUBTIPO_TIPOFRETE = 31
Public Const ADM_SUBTIPO_URGENTE = 32
Public Const ADM_SUBTIPO_COMPRAS = 33
Public Const ADM_SUBTIPO_NATUREZA = 34
Public Const ADM_SUBTIPO_MATERIAPRIMA = 35
Public Const ADM_SUBTIPO_PRODUTOINTERMD = 36
Public Const ADM_SUBTIPO_EMBALAGEM = 37
Public Const ADM_SUBTIPO_PRODUTOACABADO = 38
Public Const ADM_SUBTIPO_PRODUTOREVENDA = 39
Public Const ADM_SUBTIPO_PRODUTOREPARO = 40
Public Const ADM_SUBTIPO_OUTROS = 41
Public Const ADM_SUBTIPO_TIPOOPERACAO = 42

'implementado em rotinasfat2/classfatformata
Public Const ADM_SUBTIPO_TIPOAJUDACUSTO = 43

'*************************************************
'Alteracao Daniel em 11/07/2002
'Subtipo do Browser de kit
Public Const ADM_SUBTIPO_KIT = 44
'*************************************************
Public Const ADM_SUBTIPO_TIPOTERC = 45
Public Const ADM_SUBTIPO_CODIGOTIPOMOVTOCTACORRENTE1 = 46

'Incluído por Luiz Nogueira em 27/10/03
Public Const ADM_SUBTIPO_RELACIONAMENTOCLIENTE_STATUS = 47
Public Const ADM_SUBTIPO_RELACIONAMENTOCLIENTE_ORIGEM = 48

Public Const ADM_SUBTIPO_TIPOPEDIDO = 49
Public Const ADM_SUBTIPO_POSITIVONEGATIVO_PERCENTUAL = 50

'Incluído por Luiz Nogueira em 06/04/04
Public Const ADM_SUBTIPO_PERCENTUAL2 = 51
Public Const ADM_SUBTIPO_CAIXA_STATUS = 52
Public Const ADM_SUBTIPO_STATUS_PV = 53

Public Const ADM_SUBTIPO_CODIGONATMOVCTA = 54
Public Const ADM_SUBTIPO_RECURSO = 55
Public Const ADM_SUBTIPO_TIPOCARTAO = 56

'##############################################
'Inserido por Wagner
Public Const ADM_SUBTIPO_DATAHEBR = 57
Public Const ADM_SUBTIPO_SEQFAMILIA = 58
Public Const ADM_SUBTIPO_CGC = 59
'##############################################

Public Const ADM_SUBTIPO_PRAZOMEDIO = 60 '54 Código que estava na Hicare
Public Const ADM_SUBTIPO_STATUS_TITPAG = 61 '55 Código que estava na Hicare

Public Const ADM_SUBTIPO_TIPO_DOC_PROJETO = 62
Public Const ADM_SUBTIPO_PROJETO = 63
Public Const ADM_SUBTIPO_QUANTOP_LOTE = 64
Public Const ADM_SUBTIPO_TIPOHARMONIA = 65
Public Const ADM_SUBTIPO_STATUSHARMONIA = 66
Public Const ADM_SUBTIPO_APORTE = 67

Public Const ADM_SUBTIPO_STATUS_OV = 68 'status de orçamento de venda
Public Const ADM_SUBTIPO_STATUS_OV_COMERCIAL = 69 'status de analise de preço de orçamento de venda

Public Const ADM_SUBTIPO_PRECOUNITARIO = 70


Public Const STRING_CNAE_DESCRICAO = 255
Public Const STRING_CNAE_CODIGO = 20

Public Const STRING_BROWSER_OPCAO = 50
Public Const STRING_BROWSEREXCEL_TITULO = 50
Public Const STRING_BROWSEREXCEL_LOCALIZACAOCSV = 255
Public Const STRING_BROWSEREXCEL_ARQUIVO = 50
Public Const STRING_NOME_ROTINABOTAOCONSULTA = 50
Public Const STRING_NOME_ROTINABOTAOEDITA = 50
Public Const STRING_NOME_ROTINABOTAOSELECIONA = 50
Public Const STRING_NOME_INICCOLBROWSE = 50
Public Const STRING_TITULOBROWSER = 50
Public Const STRING_NOME_CLASSEBROWSER = 50
Public Const STRING_NOME_TRATAPARAMETROS = 50
Public Const STRING_NOME_EMPRESA = 50
Public Const STRING_SIGLA_USUARIO = 10
Public Const STRING_USUARIO = 50
Public Const STRING_GRUPO = 10
Public Const STRING_NOME_TELA = 50
Public Const STRING_NOME_TABELA = 50
Public Const STRING_NOME_CAMPO = 50
Public Const STRING_TITULO_CAMPO = 150
Public Const STRING_NOME_INDICE = 150
Public Const STRING_COMANDO_SQL = 4000
Public Const STRING_ORDENACAO_SQL = 200
Public Const STRING_SELECAO_SQL = 250
Public Const STRING_DESCRICAO_CAMPO = 50
Public Const STRING_NOME_BD = 50
Public Const STRING_VALIDACAO_CAMPO = 50
Public Const STRING_VALOR_DEFAULT_CAMPO = 50
Public Const STRING_FORMATACAO_CAMPO = 50
Public Const STRING_TITULO_ENTRADA_DADOS_CAMPO = 50
Public Const STRING_TITULO_GRID_CAMPO = 50
Public Const STRING_CODIGO_ROTINA = 50
Public Const STRING_NOME_ARQ_COMPLETO = 200
Public Const STRING_MODULO_SIGLA = 5
Public Const STRING_MODULO_NOME = 50
Public Const STRING_MODULO_DESCRICAO = 50
Public Const STRING_MODULO_VERSAO = 50
Public Const STRING_INSCR_EST = 18
Public Const STRING_INSCR_MUN = 18

 'Alteracao Daniel
Public Const STRING_INSCR_INSS = 11

Public Const STRING_INSCR_SUF = 9
Public Const STRING_NOME_PROPERTY = 50

'implementacao trocada por property get em admlib.classconstcust
'Public Const STRING_ENDERECO = 40
'Public Const STRING_BAIRRO = 12
'Public Const STRING_CIDADE = 15

Public Const STRING_ESTADO = 2
Public Const STRING_CEP = 8

'implementacao trocada por property get em admlib.classconstcust
'Public Const STRING_TELEFONE = 18
'Public Const STRING_EMAIL = 50
'Public Const STRING_FAX = 18
'Public Const STRING_CONTATO = 50
Public Const STRING_SKYPE = 50
Public Const STRING_RADIO = 50

Public Const STRING_ENDERECO_REFERENCIA = 250
Public Const STRING_ENDERECO_LOGRADOURO = 250
Public Const STRING_ENDERECO_COMPLEMENTO = 250
Public Const STRING_ENDERECO_TIPOLOGRADOURO = 50
Public Const STRING_ENDERECO_TELNUMERO1 = 18

Public Const STRING_ESTADOS_SIGLA = 2
Public Const STRING_PAISES_NOME = 20
Public Const STRING_AGENCIA = 7
Public Const STRING_CONTA_CORRENTE = 14
Public Const STRING_FILIAL_NOME = 50
Public Const STRING_ISS_CODIGO = 10
Public Const STRING_USUARIO_STRINGCONEXAO = 255
Public Const STRING_TITULO_MENU = 50
Public Const STRING_SIGLA_ROTINA = 40
Public Const STRING_NOME_CONTROLE = 50
Public Const STRING_UMEMBALAGEM = 5
Public Const STRING_APROVACAO_CARTAO = 20
Public Const STRING_NUMERO_CARTAO = 20

'Indica que o registro sofreu alteração
Public Const REGISTRO_ALTERADO = 1

'Indica que o registro não existe
Public Const REGISTRO_INEXISTENTE = 2

'Indica que o registro não sofreu alteracao
Public Const REGISTRO_INALTERADO = 3

'Tipo de Operacao
Public Const GRAVACAO = 1
Public Const MODIFICACAO = 0

'Tela Indice
Public Const STRING_TELAINDICE_NOME_EXTERNO = 50

'Tela IndiceCampo
Public Const STRING_TELAINDICECAMPO_NOME_CAMPO = 50

'Tela Segmento
Global Const GRID_PROIBIDO_EXCLUIR = 1
Global Const GRID_PROIBIDO_INCLUIR = 1
Global Const GRID_PROIBIDO_INCLUIR_NOMEIO = 1
Global Const REGISTRO_CANCELADO = 3

'indica se a largura do grid será calcula automaticamente.
Global Const GRID_LARGURA_AUTOMATICA = 1
Global Const GRID_LARGURA_MANUAL = 0

'indica se o checkbox no grid está ativo ou inativo
Public Const GRID_CHECKBOX_ATIVO As String = "1"
Public Const GRID_CHECKBOX_INATIVO As String = "0"

'indica se vai tratar o click da coluna 0 do grid
Public Const GRID_INCLUIR_BOTAO = 1
Public Const GRID_NAO_INCLUIR_BOTAO = 0

'indica se vai permitir visualizar mais uma linha que pode ser encoberta pela barra de scroll horizontal
Public Const GRID_INCLUIR_HSCROLL = 1
Public Const GRID_NAO_INCLUIR_HSCROLL = 0

'índica se deve executar funções do grid ou não
Public Const GRID_EXECUTAR_FUNCAO = 0
Public Const GRID_NAO_EXECUTAR_FUNCAO = 1

'índica se deve executar a rotina que indica a habilitacao dos campos do grid
Public Const GRID_EXECUTAR_ROTINA_ENABLE = 1
Public Const GRID_NAO_EXECUTAR_ROTINA_ENABLE = 0

'Controles do tipo checkbox
Global Const MARCADO = 1
Global Const DESMARCADO = 0

'Controles do tipo Operação
Global Const Cod_Tipo_Imp = 0
Global Const Cod_Tipo_Exp = 1
Global Const Nome_Tipo_Imp = "Importação"
Global Const Nome_Tipo_Exp = "Exportação"
Global Const Nome_Tipo_MercInt = "Mercado Interno"

Global Const POSICAO_FORA_TELA = -20000

Global Const AD_BOOL_TRUE = 1
Global Const AD_BOOL_FALSE = 0

'Global Const STRING_PROJETO = 40
'Global Const STRING_CLASSE = 40

'Nome de Rotina
Global Const NOME_ROTINA = 50
Global Const NOME_PROJETO = 50
Global Const NOME_CLASSE = 50

'Formato para preco unitario de notas externas
Public Const FORMATO_PRECO_UNITARIO_EXTERNO = "#,##0.00######"

'formato p/taxa de conversao de moeda
Public Const FORMATO_TAXA_CONVERSAO_MOEDA = "#,##0.00##"

'Formato para quantidades de Produtos
Public Const FORMATO_ESTOQUE = "#,##0.0###"

'Formato para custos de Produtos
Public Const FORMATO_CUSTO = "#,##0.0000"

'Formato para media de atraso
Public Const FORMATO_MEDIA_ATRASO = "##,##0.#"

'Formato para valores inteiros
Public Const FORMATO_INTEIRO = "###,###,###,##0"

Public Const FORMATO_CPF = "000\.000\.000-00; ; ; "
Public Const FORMATO_CGC = "00\.000\.000\/0000-00; ; ; "

Global Const AUMENTA_DATA = 0
Global Const DIMINUI_DATA = 1

'Largura, Altura no GRID
Global Const ALT_FOLGA = 170 'Folga na altura para mover a Masked Edit
Global Const DIVISAO_LARG = 15 'Largura da linha divisoria
Global Const COL_FOLGA = 19 'Folga em cada coluna
Global Const DIVISAO_ALT = 17 'Altura da linha divisoria
Global Const CELULA_ALT = 225 'Altura de cada célula
    
'Origens
Global Const ORIGEM_CONTABILIDADE = 1
Global Const STRING_ORIGEM = 3
Global Const STRING_ORIGEM_MAIS_UM = STRING_ORIGEM + 1
Global Const STRING_ORIGEM_DESCRICAO = 25


'Estruturas
'Type typeOrigem
'    sOrigem As String * STRING_ORIGEM
'    sDescricao As String * STRING_ORIGEM_DESCRICAO
'End Type

Declare Function Init_Contab_Int Lib "ADCUSR.DLL" Alias "AD_Sistema_InitContab" (ByVal lID_Sistema As Long) As Long
Declare Function Init_Fest_Int Lib "ADCUSR.DLL" Alias "AD_Sistema_InitFest" (ByVal lID_Sistema As Long) As Long

'Rotinas de Gerencia de Arquivo Temporário

' Permite o armazenamento de registros de tamanho fixo
' sequencialmente e a posterior leitura dos mesmos de forma randomica

Global Const ARQ_TEMP_OK = 0
Global Const ARQ_TEMP_ERR_WRITE = 1
Global Const ARQ_TEMP_ERR_SEEK = 2
Global Const ARQ_TEMP_ERR_READ = 3

'Declare Function Rotina1 Lib "ADCUSR.DLL" Alias "AD_RotinaExecutar1" (ByVal IID_DicDados As Long, ByVal sCodigo As String, anyP1 As Any) As Long
'Declare Function Rotina2 Lib "ADCUSR.DLL" Alias "AD_RotinaExecutar2" (ByVal IID_DicDados As Long, ByVal sCodigo As String, anyP1 As Any, anyP2 As Any) As Long
'Declare Function Rotina3 Lib "ADCUSR.DLL" Alias "AD_RotinaExecutar3" (ByVal IID_DicDados As Long, ByVal sCodigo As String, anyP1 As Any, anyP2 As Any, anyP3 As Any) As Long



'Rotinas de Manipulação de Banco de Dados

'Retorno comandos SQL
Global Const AD_SQL_SUCESSO = 0
Global Const AD_SQL_SUCESSO_PARCIAL = 1
Global Const AD_SQL_ERRO = -1
Global Const AD_SQL_SEM_DADOS = 100

Global Const AD_SQL_DRIVER_ODBC = 1

'Arquivo Temporário
Declare Function Arq_Temp_Preparar Lib "ADCRTL.DLL" Alias "FN_Exec_ArqTemp_OnPrepGet" (ByVal lID_Arq_Temp As Long) As Long
Declare Function Arq_Temp_Criar Lib "ADCRTL.DLL" Alias "FN_Exec_ArqTemp_OnCreate" (ByVal iTamanho_Registro As Integer) As Long
' Inicializa tratamento de arq temporario
'
'Parametros:
'(I)iTamanho_Registro: tamanho dos registros que se pretende armazenar/recuperar

Declare Function Arq_Temp_Destruir Lib "ADCRTL.DLL" Alias "FN_Exec_ArqTemp_OnDestroy" (ByVal lID_Arq_Temp As Long) As Long
Declare Function Arq_Temp_Inserir Lib "ADCRTL.DLL" Alias "FN_Exec_ArqTemp_OnInsert" (ByVal lID_Arq_Temp As Long, anyRegistro As Any, lPosicao As Long) As Long
Declare Function Arq_Temp_Ler Lib "ADCRTL.DLL" Alias "FN_Exec_ArqTemp_OnGetDirect" (ByVal lID_Arq_Temp As Long, anyRegistro As Any, lPosicao As Long) As Long

'Rotinas de Gerencia de Sort

'  permite a insercao de chaves de tamanho variavel composta de 1 ou + segmentos,
' que possam ser comparados como texto ou numero long double e sua posterior
' recuperacao de forma ordenada.
' Cada chave incluida deve estar associada a uma referencia de tamanho fixo que
' possa ser usada p/identifica-la na leitura das chaves apos a ordenacao.

Declare Function Sort_Abrir Lib "ADCRTL.DLL" Alias "FN_Sort_Criar" (ByVal iTamanho_Offset As Integer, ByVal iNum_Segmentos As Integer) As Long
Declare Function Sort_Destruir Lib "ADCRTL.DLL" Alias "FN_Sort_Destruir" (ByVal lID_Sort As Long) As Long
'Function Sort_Inserir(ByVal lID_Sort As Long, ByVal lPosicao As Long, vSegmento1 As Variant, Optional vSegmento2 As Variant, Optional vSegmento3 As Variant, Optional vSegmento4 As Variant, Optional vSegmento5 As Variant) As Long
Declare Function Sort_Classificar Lib "ADCRTL.DLL" Alias "FN_Sort_PrepMerge" (ByVal lID_Sort As Long) As Long
Declare Function Sort_Ler Lib "ADCRTL.DLL" Alias "FN_Sort_GetRec" (ByVal lID_Sort As Long, lPosicao As Long) As Long


'''' API do WIndows
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

'****** Inicio Edicao de Telas ***************************

Public Const HWND_TOPMOST = -1

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type RECT
        left As Long
        top As Long
        right As Long
        bottom As Long
End Type

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetDialogBaseUnits Lib "user32" () As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Const GWL_WNDPROC = -4

Public Const SWP_SHOWWINDOW = &H40
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOUSEMOVE = &H200
Public Const WM_NCHITTEST = &H84

'****** Fim Edicao de Telas ***************************

Public Type typeCotacaoMoeda
    dValor As Double
    iMoeda As Integer
    dtData As Date
End Type

Type typeBrowseArquivo
    sNomeTela As String
    sNomeArq As String
    sSelecaoSQL As String
    sProjeto As String
    sClasse As String
    sClasseBrowser As String
    sTrataParametros As String
    sTituloBrowser As String
    sRotinaBotaoEdita As String
    sRotinaBotaoSeleciona As String
    sRotinaBotaoConsulta As String
    iBotaoEdita As Integer
    iBotaoSeleciona As Integer
    iBotaoConsulta As Integer
    sProjetoObjeto As String
    sClasseObjeto As String
    sNomeTelaConsulta As String
    sNomeTelaEdita As String
    iBancoDados As Integer
End Type

Type typeBrowseUsuarioCampo
    sNomeTela As String
    sCodUsuario As String
    sNomeArq As String
    sNome As String
    iPosicaoTela As Integer
    sTitulo As String
    lLargura As Long
End Type

Type typeBrowseCampo
    sNomeTela As String
    sNomeCampo As String
    sNome As String
End Type

Type typeBrowseParamSelecao
    sNomeTela As String
    iOrdem As Integer
    sProjeto As String
    sClasse As String
    sProperty As String
End Type

Type typeBrowseIndice
    sNomeTela As String
    iIndice As Integer
    sNomeIndice As String
    sOrdenacaoSQL As String
    sSelecaoSQL As String
End Type

Type typeBrowseIndiceSegmentos
    sNomeTela As String
    iIndice As Integer
    iPosicaoCampo As Integer
    sNomeCampo As String
End Type

Type typeBrowseUsuarioOrdenacao
    sNomeTela As String
    sCodUsuario As String
    iIndice As Integer
    sSelecaoSQL1 As String
    sSelecaoSQL1Usuario As String
    sNomeIndice As String
End Type

Type typeBrowseOpcao

    sNomeTela As String
    sOpcao As String
    lTopo As Long
    lEsquerda As Long
    lLargura As Long
    lAltura As Long
    
End Type

Type typeBrowseOpcaoCampo
    sNomeTela As String
    sOpcao As String
    sNomeArq As String
    sNome As String
    iPosicaoTela As Integer
    sTitulo As String
    lLargura As Long
End Type

Type typeBrowseOpcaoOrdenacao
    sNomeTela As String
    sOpcao As String
    iIndice As Integer
    sSelecaoSQL1 As String
    sSelecaoSQL1Usuario As String
    sNomeIndice As String
End Type

Type typeGrupoBrowseCampo
    sCodGrupo As String
    sNomeTela As String
    sNomeArq As String
    sNome As String
End Type

Type typeCampos
    sNomeArq As String
    sNome As String
    sNomeBd As String
    sDescricao As String
    iObrigatorio As Integer
    iImexivel As Integer
    iAtivo As Integer
    sValDefault As String
    sValidacao As String
    sFormatacao As String
    iTipo As Integer
    iTamanho As Integer
    iPrecisao As Integer
    iDecimais As Integer
    iTamExibicao As Integer
    sTituloEntradaDados As String
    sTituloGrid As String
    iSubTipo As Integer
    iAlinhamento As Integer
    iTestaIntegridade As Integer
End Type

Type typeModulo
    sSigla As String
    sNome As String
    sDescricao As String
    sVersao As String
    iAtivo As Integer
    sOrigemCusto As String
    sOrigemAglutina As String
End Type

Public Const STRING_NFCE_CSC = 36 'codigo de seguranca do contribuinte
Public Const STRING_NFCE_ID_CSC = 6 'id do codigo de seguranca do contribuinte

Type typeFiliais
    lCodEmpresa As Long
    iCodFilial As Integer
    sNome As String
    sCgc As String
    iAlmoxarifadoPadrao As Integer
    sInscricaoEstadual As String
    iICMSPorEstimativa As Integer
    sInscricaoMunicipal As String
    dISSPercPadrao As Double
    sISSCodigoPadrao As String
    iISSIncluso As Integer
    iContribuinteIPI As Integer
    dIRPercPadrao As Double
    lEndereco As Long
    lEndEntrega As Long
    iInativa As Integer
    sRamoEmpresa As String
    sJucerja As String
    dtDataJucerja As Date
    sContador As String
    sCRCContador As String
    sCPFContador As String
    iTipoTribAtividade As Integer
    iSimplesFederal As Integer
    dSimplesFederalAliq As Double
    dSimplesFederalTeto As Double
    iSuperSimples As Integer
    iPISNaoCumulativo As Integer
    iCOFINSNaoCumulativo As Integer
    iLucroPresumido As Integer
    sCertificadoA1A3 As String
    iNFeAmbiente As Integer
    iRPSAmbiente As Integer
    sCNAE As String
    sSignatarioCTB As String
    sCodQualiSigCTB As String
    sCPFSignatarioCTB As String
    sSpedFiscalPerfil As String
    sNomeReduzido As String
    lEnderecoContador As Long
    sCNPJContador As String
    iPisCofinsRegCumTipo As Integer
    iContribPrevPorRecBruta As Integer
    dAliquotaPadraoContribPrev As Double
    iPisCofinsIndAproCred As Integer
    
    'nfe 3.10
    sNFCECSC As String
    sidNFCECSC As String
    iindSincPadrao As Integer
    iRegimeTrib As Integer
    iRegimeEspecialTrib As Integer

End Type

Type typeSegmento
    iNivel As Integer
    iTipo As Integer
    iTamanho As Integer
    sDelimitador As String
    iPreenchimento As Integer
End Type

Type typeUsuarios
    sCodUsuario As String
    sCodGrupo As String
    sNome As String
    sNomeReduzido As String
    sSenha As String
    dtDataValidade As Date
    iAtivo As Integer
    iWorkFlowAtivo As Integer
    sEmail As String
End Type

Type typeRelTelaCampo
    sCodRel As String
    iSequencial As Integer
    sNome As String
    iTipo As Integer
    iTamanho As Integer
End Type

Type typeRelatorio
    sCodRel As String
    sRotinaAuxiliar As String
    sNomeTsk As String
    sAutor As String
    sDescricao As String
    sUltOpcaoUtilizada As String
    iTipo As Integer
    iOrigem As Integer
    iDispositivoDeSaida As Integer
    sNomeArqReimp As String
    iFilialEmpresa As Integer
End Type

Type typeUsuarioModulo
    sCodUsuario As String
    lCodEmpresa As Long
    iCodFilial As Integer
    dtDataValidade As Date
    sSiglaModulo As String
    sNomeModulo As String
    sVersaoModulo As String
    iMenuModulo As Integer
End Type

Type typeMenuItens
    iIdentificador As Integer
    sTitulo As String
    sSiglaRotina As String
    sNomeTela As String
    sNomeControle As String
    iIndiceControle As Integer
    sNomeControlePai As String
    iIndiceControlePai As Integer
    iIndiceControleNoPai As Integer
    iSeparador As Integer
End Type
    
Type typeModuloFilEmp
    lCodEmpresa As Long
    iCodFilial As Integer
    sSiglaModulo As String
    iConfigurado As Integer
End Type

Type typeDicConfig
    iID As Integer
    iLimiteFiliais As Integer
    iLimiteLogs As Integer
    sSerie As String
    iTipoVersao As Integer
    iLimiteEmpresas As Integer
    dtValidadeDe As Date
    dtValidadeAte As Date
    sSenha As String
    dtDataSenha As Date
End Type


Type typeClientesLimites
    lCodCliente As Long
    iCodFilial As Integer
    sSerie As String
    iTipoVersao As Integer
    sSenha As String
    dtDataSenha As Date
    dtValidadeAte As Date
    sVersao As String
    iLimiteLogs As Integer
    iLimiteFiliais As Integer
    iLimiteEmpresas As Integer
End Type


Type typeObjetosBd

    sClasseObjeto As String
    sNomeArquivo As String
    iTipo As Integer
    sSelecaoSQL As String
    iAvisaSobrePosicao As Integer
    sNomeObjetoMSG As String
    
End Type



'Inicio EdicaoTela Raphael
Public Type typeEdicaoTela

    sNomeTela As String
    sNomeControle As String
    sTitulo As String
    iVisivel As Integer
    iLargura As Integer
    iAltura As Integer
    iEsquerda As Integer
    iTopo As Integer
    iTabStop As Integer
    sContainer As String
    iIndiceContainer As Integer
    iTabIndex As Integer
    sGrupoUsuarios As String 'Inserido por Wagner
    izOrder As Integer 'Inserido por Wagner
    iHabilitado As Integer 'Inserido por Wagner
    
End Type

Public Type typeGrupoUsuario

    sCodGrupo As String
    sDescricao As String
    dtDataValidade As Date
    iLogAtividade As Integer

End Type



'Constantes edicao Tela Raphael

Public Const STRING_EDICAOTELA_CONTAINER = 50
Public Const STRING_EDICAOTELA_TITULO = 50

'Manual Contabilidade
Public Const IDH_CONFIGURACAO_INICIALIZACAO = 1000
Public Const IDH_CONFIGURACAO_CENTRO_CUSTO_LUCRO = 1001
Public Const IDH_CONFIGURACAO_VALORES_INICIAIS = 1002
Public Const IDH_CONFIGURACAO_CONTAS = 1003
Public Const IDH_CONFIG_CAMPOS_GLOBAIS_UTILIZ_CONTABILIZACAO = 1004
Public Const IDH_EXERCICIO = 1005
Public Const IDH_EXERCICIO_FILIAL = 1006
Public Const IDH_SEGMENTOS = 1007
Public Const IDH_HISTORICO_PADRAO = 1008
Public Const IDH_PLANO_CONTAS = 1009
Public Const IDH_CATEGORIA_CONTA = 1010
Public Const IDH_CENTRO_CUSTO_CENTRO_LUCRO = 1011
Public Const IDH_ASSOCIACAO_CONTA_CENTRO_CUSTO_LUCRO_EXTRA = 1012
Public Const IDH_ASSOCIACAO_CONTA_CENTRO_CUSTO_LUCRO_CONTABIL = 1013
Public Const IDH_DOCUMENTO_AUTOMATICO = 1014
Public Const IDH_MANUTENCAO_LOTES = 1015
Public Const IDH_SALDOS_INICIAIS_CENTRO_CUSTO_LUCROS = 1016
Public Const IDH_ORCAMENTO = 1017
Public Const IDH_RATEIO_ON_LINE = 1018
Public Const IDH_RATEIO_OFF_LINE = 1019
Public Const IDH_PADRAO_CONTABILIZACAO = 1020
Public Const IDH_LANCAMENTO_CONSULTA = 1021
Public Const IDH_LANCAMENTOS = 1022
Public Const IDH_RATEIO = 1023
Public Const IDH_EXTORNO_LOTE_CONTABILIZADO = 1024
Public Const IDH_ATUALIZACAO_LOTES = 1026
Public Const IDH_APURACAO_EXERCICIO = 1027
Public Const IDH_FECHAMENTO_DE_EXERCICIO = 1028
Public Const IDH_REABERTURA_EXERCICIO = 1029
Public Const IDH_REPROCESSAMENTO = 1030
Public Const IDH_RATEIO_OFF_LINE_PROCESSAM_BATCH = 1031
Public Const IDH_REL_DMPL_CONFIG = 1032
Public Const IDH_REL_DRE_CONFIG = 1033
Public Const IDH_RELOP_BALANCO_PATR_COMP = 1034
Public Const IDH_RELOP_BALANCO_PATR = 1035
Public Const IDH_RELOP_BALANCO_PATR_PERIODO = 1036
Public Const IDH_RELOP_BALANCO_VERIFIC = 1037
Public Const IDH_RELOP_DEMONST_ORIGENS_APLIC = 1038
Public Const IDH_RELOP_DEMONS_RESULT_EXERCICIO = 1039
Public Const IDH_RELOP_DEMONS_RESULT_PERIODO = 1040
Public Const IDH_RELOP_DESP_PER_CCL = 1041
Public Const IDH_RELOP_DIARIO = 1042
Public Const IDH_RELOP_LANCAMENTO_CCL = 1043
Public Const IDH_RELOP_LANCAMENTO_DATA = 1044
Public Const IDH_RELOP_LANCAMENTO_LOTE = 1045
Public Const IDH_RELOP_LANCAMENTO_PENDENTE = 1047
Public Const IDH_RELOP_LOTE = 1048
Public Const IDH_RELOP_LOTE_PENDENTE = 1049
Public Const IDH_RELOP_MUTACAO_PL = 1050
Public Const IDH_RELOP_ORC_REAL_CCL = 1051
Public Const IDH_RELOP_ORC_REAL = 1052
Public Const IDH_RELOP_PLANO_CONTAS = 1053
Public Const IDH_RELOP_PLAN_SALDOS = 1054
Public Const IDH_RELOP_RAZAO_AUXILIAR = 1055
Public Const IDH_RELOP_RAZAO = 1056
Public Const IDH_APURACAO_PERIODO = 1057
Public Const IDH_CONFIGURACAO_SETUP_ID = 1058
Public Const IDH_CONFIGURACAO_SETUP_CCL = 1059
Public Const IDH_CONFIGURACAO_SETUP_VALORES_INICIAIS = 1060
Public Const IDH_CRIAR_EXERCICIO = 1061
Public Const IDH_LANCAMENTO_EXTORNO = 1062
Public Const IDH_LANCAMENTO_EXTORNO1 = 1063
Public Const IDH_LANCAMENTO_AT = 1064
Public Const IDH_LOTE_ATUALIZA = 1065
Public Const IDH_LOTE_CONSULTA = 1066
Public Const IDH_EXTORNO_LOTE_CONTABILIZADO1 = 1067

'Manual contas a Pagar
Public Const IDH_FORNECEDOR_IDENT = 2000
Public Const IDH_FORNECEDOR_DADOS_FIN = 2001
Public Const IDH_FORNECEDOR_INSCRICOES = 2002
Public Const IDH_FORNECEDOR_ENDERECO = 2003
Public Const IDH_FORNECEDOR_ESTATISTICAS = 2004
Public Const IDH_FILIAL_FORN_IDENT = 2005
Public Const IDH_FILIALFORN_DADOS_FIN = 2006
Public Const IDH_FILIALFORN_INSCRICOES = 2007
Public Const IDH_FILIALFORN_ENDERECO = 2008
Public Const IDH_FILIALFORN_ESTATISTICAS = 2009
Public Const IDH_TIPOS_FORN_IDENT = 2010
Public Const IDH_TIPOS_FORN_DADOS_FIN = 2011
Public Const IDH_PORTADORES = 2012
Public Const IDH_CONDICAO_PAGAMENTO = 2013
Public Const IDH_ADIANTAM_FORNEC_IDENT = 2014
Public Const IDH_ADIANTAM_FORNEC_CONTABILIZACAO = 2015
Public Const IDH_CHEQUES_PAGAR_P1 = 2016
Public Const IDH_IMPRESSAO_CHEQUES_P2 = 2017
Public Const IDH_IMPRESSAO_CHEQUES_P3 = 2018
Public Const IDH_IMPRESSAO_CHEQUES_P4 = 2019
Public Const IDH_DEVOL_CREDITOS_ID = 2020
Public Const IDH_DEVOL_CREDITO_CONTABILIZACAO = 2021
Public Const IDH_NOTAS_FISCAIS_PAGAR_ID = 2022
Public Const IDH_NOTAS_FISCAIS_PAGAR_CONTABILIZACAO = 2023
Public Const IDH_FATURAS_PAGAR_ID = 2024
Public Const IDH_FATURAS_PAGAR_COBRANCA = 2025
Public Const IDH_NOTA_FISCAL_FATURA_PAGAR_ID = 2026
Public Const IDH_NOTA_FISCAL_FATURA_COBRANCA = 2027
Public Const IDH_NOTA_FISCAL_FATURA_CONTABILIZACAO = 2028
Public Const IDH_CONFIRMACAO_COBRANCA = 2029
Public Const IDH_BORDERO_PAGT_P1 = 2030
Public Const IDH_BORDERO_PAGT_P2 = 2031
Public Const IDH_BORDERO_PAGT_P3 = 2032
Public Const IDH_BORDERO_PAGT_P4 = 2033
Public Const IDH_BORDERO_PAGT_P5 = 2034
Public Const IDH_CHEQUE_MANUAL_P1 = 2035
Public Const IDH_CHEQUE_MANUAL_P2 = 2036
Public Const IDH_CHEQUE_MANUAL_P3 = 2037
Public Const IDH_CHEQUE_MANUAL_P4 = 2038
Public Const IDH_CANCELAR_PAGAMENTOS = 2039
Public Const IDH_BAIXA_PARCELAS_PAGAR_TITULO = 2040
Public Const IDH_BAIXA_PARCELAS_PAGAR_PARCELAS = 2041
Public Const IDH_BAIXA_PARCELAS_PAGAR_CONTABILIZACAO = 2042
Public Const IDH_CANCELAR_BAIXA_CONTAS_PAGAR_ID = 2043
Public Const IDH_CANCELAR_BAIXA_CONTAS_PAGAR_CONTABILIZACAO = 2044
Public Const IDH_OUTROS_PAGAMENTOS_ID = 2045
Public Const IDH_OUTROS_PAGAMENTOS_COBRANCA = 2046
Public Const IDH_OUTROS_PAGAMENTOS_CONTABILIZACAO = 2047
Public Const IDH_COMISSOES_PAG = 2048
Public Const IDH_CONFIGURA_CP = 2049
Public Const IDH_EMISSAO_BOLETO_SELECAO = 2050
Public Const IDH_EMISSAO_BOLETO_EMISSAO = 2051
Public Const IDH_GERACAO_ARQICMS = 2052
Public Const IDH_GERACAO_ARQREMCOBR = 2053
Public Const IDH_GERACAO_PLANOCONTA = 2054
Public Const IDH_NOTA_FISCAL_PAGAR_CONS_NF = 2055
Public Const IDH_NOTA_FISCAL_PAGAR_CONS_FATURA = 2056
Public Const IDH_NOTA_FISCAL_PAGAR_CONS_CONTABILIZACAO = 2057
Public Const IDH_PAISES = 2058
Public Const IDH_TITULOPAG_CONS_TITULO = 2059
Public Const IDH_TITULOPAG_CONS_PARCELAS = 2060
Public Const IDH_TITULOPAG_CONS_BAIXA = 2061
Public Const IDH_TITULOPAG_CONS_CONTABILIZACAO = 2062
Public Const IDH_PROCESSA_ARQRETCOBRANCA = 2063

'Manual Contas à Receber
Public Const IDH_ADIANTAM_CLIENTE_ID = 3000
Public Const IDH_ADIANTAM_CLIENTE_CONTABILIZACAO = 3001
Public Const IDH_CHEQUE_PRE_DATADO = 3002
Public Const IDH_BORDERO_DEPOSITO_CHEQUES_PRE_DATADOS = 3003
Public Const IDH_BORDERO_DEPOSITO_CHEQUES_PRE_DATADOS2 = 3004
Public Const IDH_DEVOL_DEB_CLIENTES_ID = 3005
Public Const IDH_DEVOL_DEB_CLIENTES_COMISSSOES = 3006
Public Const IDH_DEVOL_DEB_CLIENTES_CONTABILIZACAO = 3007
Public Const IDH_CARTEIRAS_COBRANCA = 3008
Public Const IDH_BORDERO_COBRANCA = 3009
Public Const IDH_BORDERO_COBRANCA_P2 = 3010
Public Const IDH_BORDERO_COBRANCA_P3 = 3011
Public Const IDH_TITULOS_RECEBER_ID = 3012
Public Const IDH_TITULOS_RECEBER_PARCELAS = 3013
Public Const IDH_TITULOS_RECEBER_CONTABILIZACAO = 3014
Public Const IDH_INSTRUCOES_PARA_TITULO_COBRANCA_ELETRONICA = 3015
Public Const IDH_CANCELAR_BORDERO_COBRANCA = 3016
Public Const IDH_TRANSFERENCIA_CARTEIRA_COBRANCA = 3017
Public Const IDH_BAIXA_TITULOS_RECEBER_RECIBOS = 3018
Public Const IDH_BAIXA_TITULOS_RECEBER_CONTABILIZACAO = 3019
Public Const IDH_BAIXA_PARCELAS_RECEBER_TITULOS = 3020
Public Const IDH_BAIXA_PARCELAS_RECEBER_PARCELAS = 3021
Public Const IDH_BAIXA_PARCELAS_RECEBER_CONTABILIZACAO = 3022
Public Const IDH_CANCELAR_BAIXA_CONTAS_RECEBER_TITULOS = 3023
Public Const IDH_CANCELAR_BAIXA_CONTAS_RECEBER_BAIXAS = 3024
Public Const IDH_REGIOES_VENDA = 3025
Public Const IDH_CATEGORIAS_CLIENTE = 3026
Public Const IDH_CLIENTES_ID = 3027
Public Const IDH_CLIENTES_DADOS_FIN = 3028
Public Const IDH_CLIENTES_INSCRICOES = 3029
Public Const IDH_CLIENTES_ENDERECOS = 3030
Public Const IDH_CLIENTES_VENDAS = 3031
Public Const IDH_CLIENTES_ESTATISTICAS = 3032
Public Const IDH_FILIAIS_CLIENTES_ID = 3033
Public Const IDH_FILIAIS_CLIENTES_INSCRICOES = 3034
Public Const IDH_FILIAIS_CLIENTES_ENDERECOS = 3035
Public Const IDH_FILIAIS_CLIENTES_VENDAS = 3036
Public Const IDH_FILIAIS_CLIENTES_ESTATISTICAS = 3037
Public Const IDH_CONDICOES_PAGAMENTO = 3038
Public Const IDH_TIPOS_VENDEDOR = 3039
Public Const IDH_VENDE_ID = 3040
Public Const IDH_VENDE_COMISSAO = 3041
Public Const IDH_VENDE_ENDERECO = 3042
Public Const IDH_COBRADORES_ID = 3043
Public Const IDH_COBRADORES_CARTEIRAS = 3044
Public Const IDH_COBRADORES_ENDERECOS = 3045
Public Const IDH_PADROES_COBRANCAS = 3046
Public Const IDH_CONFIGURA_CR = 3047
Public Const IDH_TITULOREC_CONS_TITULO = 3048
Public Const IDH_TITULOREC_CONS_PARCELAS = 3049
Public Const IDH_TITULOREC_CONS_BAIXA = 3050
Public Const IDH_TITULOREC_CONS_CONTABILIZACAO = 3051

'RelatoriosCPR
Public Const IDH_RELOP_BAIXASCP = 4000
Public Const IDH_RELOP_BAIXASCR = 4001
Public Const IDH_RELOP_BORDEROPAG = 4002
Public Const IDH_RELOP_CADCLI_L = 4003
Public Const IDH_RELOP_CADFORN_L = 4004
Public Const IDH_RELOP_CADFORN = 4005
Public Const IDH_RELOP_VEND_L = 4006
Public Const IDH_RELOP_VEND = 4007
Public Const IDH_RELOP_CLI_ATRASO = 4008
Public Const IDH_RELOP_COMIS_VEND = 4009
Public Const IDH_RELOP_CONC_PEND = 4010
Public Const IDH_RELOP_DEVEDORES = 4011
Public Const IDH_RELOP_DIRF = 4012
Public Const IDH_RELOP_EXTRATOBAN = 4013
Public Const IDH_RELOP_EXTRATOTES = 4014
Public Const IDH_RELOP_HIST_APLIC = 4015
Public Const IDH_RELOP_ICMS = 4016
Public Const IDH_RELOP_INSSRET = 4017
Public Const IDH_RELOP_IRRF = 4018
Public Const IDH_RELOP_JUROS_REC = 4019
Public Const IDH_RELOP_MOVFIN_DET = 4020
Public Const IDH_RELOP_MOVFIN = 4021
Public Const IDH_RELOP_NFPAG_SEMFAT = 4022
Public Const IDH_RELOP_PAGTOS_CANCELADOS = 4023
Public Const IDH_RELOP_POS_APLIC = 4024
Public Const IDH_RELOP_POSCLI = 4025
Public Const IDH_RELOP_POSCLI_L = 4026
Public Const IDH_RELOP_POSFORN = 4027
Public Const IDH_RELOP_REL_CHEQUES = 4028
Public Const IDH_RELOP_RESUMO_COMIS = 4029
Public Const IDH_RELOP_TITPAG_L = 4030
Public Const IDH_RELOP_TITPAG = 4031
Public Const IDH_RELOP_TITREC_L = 4032
Public Const IDH_RELOP_TIT_REC_MALA = 4033
Public Const IDH_RELOP_TIT_REC = 4034
Public Const IDH_RELOP_TIT_REC_TEL = 4035
Public Const IDH_RELOP_CAD_CLI = 4036
Public Const IDH_RELOP_BORDERO_COBRANCA = 4037

'Manual DIC
Public Const IDH_EMPRESA_ID = 5000
Public Const IDH_EMPRESA_MODULOS = 5001
Public Const IDH_FILIAL_EMPRESA_ID = 5002
Public Const IDH_FILIAL_EMPRESA_MODULOS = 5003
Public Const IDH_FILIAL_EMPRESA_COMPLEM = 5004
Public Const IDH_FILIAL_EMPRESA_ENDERECOS = 5005
Public Const IDH_FILIAL_EMPRESA_TRIBUTACAO = 5006
Public Const IDH_GRUPO = 5007
Public Const IDH_USUARIO_ID = 5008
Public Const IDH_USUARIO_ACESSO_FILIAIS_EMPRESA = 5009
Public Const IDH_ROTINA = 5010
Public Const IDH_TELA = 5011
Public Const IDH_RELATORIOS = 5012
Public Const IDH_GRUPO_ROTINA = 5013
Public Const IDH_GRUPO_TELA = 5014
Public Const IDH_GRUPO_RELATORIO = 5015
Public Const IDH_CAMPOS_PERM_GRUPO_TELA = 5016
Public Const IDH_ROTINA_GRUPO = 5017
Public Const IDH_TELA_GRUPO = 5018
Public Const IDH_RELATORIO_GRUPO = 5019
Public Const IDH_RELCADASTRO = 5020
Public Const IDH_ROTINA_TELA = 5021

'Manual Estoque
'Public Const IDH_SEGMENTOS = 6000
Public Const IDH_ESTOQUE_INICIAL = 6001
Public Const IDH_FORNECEDORES_ID = 6002
Public Const IDH_FORNECEDORES_DADOS_FIN = 6003
Public Const IDH_FORNECEDORES_INSCRICOES = 6004
Public Const IDH_FORNECEDORES_ENDERECOS = 6005
Public Const IDH_FORNECEDORES_ESTATISTICAS = 6006
Public Const IDH_FILIAIS_FORNECEDORES_ID = 6007
Public Const IDH_FILIAIS_FORNECEDORES_DADOS_FIN = 6008
Public Const IDH_FILIAIS_FORNECEDORES_INSCRICOES = 6009
Public Const IDH_FILIAIS_FORNECEDORES_ENDERECO = 6010
Public Const IDH_FILIAIS_FORNECEDORES_ESTATISTICAS = 6011
Public Const IDH_TIPOS_FORNECEDOR_ID = 6012
Public Const IDH_TIPOS_FORNECEDOR_DADOS_FIN = 6013
Public Const IDH_PRODUTOS_ALMOXARIFADO_DADOS_PRINCIPAIS = 6014
Public Const IDH_PRODUTOS_ALMOXARIFADO_SALDOS = 6015
Public Const IDH_PRODUTOS_ALMOXARIFADO_SALDOS_TERCEIROS = 6016
Public Const IDH_KIT = 6017
Public Const IDH_PRODUTO_DADOS_PRINCIPAIS = 6018
Public Const IDH_PRODUTO_CATEGORIA = 6019
Public Const IDH_PRODUTO_COMPLEMENTO = 6020
Public Const IDH_PRODUTO_CARACTERISTICAS_FISICAS = 6021
Public Const IDH_PRODUTO_PRECOS = 6022
Public Const IDH_PRODUTO_UNIDADES_MEDIDA = 6023
Public Const IDH_PRODUTO_TRIBUTACAO = 6024
Public Const IDH_CLASSE_UNIDADE_MEDIDA = 6025
Public Const IDH_ALMOXARIFADO_ID = 6026
Public Const IDH_ALMOXARIFADO_ENDERECO = 6027
Public Const IDH_EXCECOES_ICMS = 6028
Public Const IDH_EXCECOES_IPI = 6029
Public Const IDH_NATUREZA_OPERACAO = 6030
Public Const IDH_SERIE_NOTA_FISCAL = 6031
Public Const IDH_PADROES_TRIBUTACAO_OPERACOES_FORNECEDORES = 6032
Public Const IDH_PADROES_TRIBUTACAO_OPERACOES_CLIENTES = 6033
Public Const IDH_TIPOS_TRIBUTACAO = 6034
Public Const IDH_TIPOS_PRODUTO_DADOS_PRINCIPAIS = 6035
Public Const IDH_TIPOS_PRODUTO_UNIDADES_MEDIDA = 6036
Public Const IDH_TIPOS_PRODUTO_TRIBUTACAO_CONTABILIZACAO = 6037
'Public Const IDH_CENTRO_CUSTO_CENTRO_LUCRO = 6038
'Public Const IDH_CENTRO_CUSTO_CENTRO_LUCRO = 6039
Public Const IDH_MANUTENCAO_LOTE = 6040
Public Const IDH_CATEGORIA_PRODUTO = 6041
'Public Const IDH_CONDICOES_PAGAMENTO = 6042
Public Const IDH_CUSTOS = 6043
Public Const IDH_CUSTO_PRODUCAO = 6044
Public Const IDH_CONTROLE_ESTOQUE = 6045
Public Const IDH_ESTADOS = 6046
Public Const IDH_REQUISICAO_MATERIAL_CONSUMO_MOVIMENTOS = 6047
Public Const IDH_REQUISICAO_MATERIAL_CONSUMO_CONTABILIZACAO = 6048
Public Const IDH_MOVIMENTOS_ESTOQUE_MOVIMENTO = 6049
Public Const IDH_MOVIMENTOS_ESTOQUE_CONTABILIZACAO = 6050
Public Const IDH_TRANSFERENCIA_MOVIMENTO = 6051
Public Const IDH_TRANSFERENCIA_CONTABILIZACAO = 6052
Public Const IDH_RECEBIMENTO_MATERIAL_CLIENTE_DADOS_PRINCIPAIS = 6053
Public Const IDH_RECEBIMENTO_MATERIAL_CLIENTE_ITENS = 6054
Public Const IDH_RECEBIMENTO_MATERIAL_CLIENTE_COMPLEMENTO = 6055
Public Const IDH_RECEBIMENTO_MATERIAL_FORNECEDOR_DADOS_PRINCIPAIS = 6056
Public Const IDH_RECEBIMENTO_MATERIAL_FORNECEDOR_ITENS = 6057
Public Const IDH_RECEBIMENTO_MATERIAL_FORNECEDOR_COMPLEMENTO = 6058
Public Const IDH_RESERVA_PRODUTO = 6059
Public Const IDH_ENTRADA_MATERIAL_PRODUZIDO_MOVIMENTOS = 6060
Public Const IDH_ENTRADA_MATERIAL_PRODUZIDO_CONTABILIZACAO = 6061
Public Const IDH_REQUISICAO_MATERIAL_PRODUCAO_MOVIMENTOS = 6062
Public Const IDH_REQUISICAO_MATERIAL_PRODUCAO_CONTABILIZACAO = 6063
Public Const IDH_INVENTARIO_LANCAMENTOS = 6064
Public Const IDH_INVENTARIO_CONTABILIZACAO = 6065
Public Const IDH_INVENTARIO_LOTES_LANCAMENTOS = 6066
Public Const IDH_INVENTARIO_LOTES_CONTABILIZACAO = 6067
Public Const IDH_EMPENHO = 6068
Public Const IDH_ORDEM_PRODUCAO = 6069
Public Const IDH_ENTRADA_NF_SIMPLES_DADOS_PRINCIPAIS = 6070
Public Const IDH_ENTRADA_NF_SIMPLES_ITENS = 6071
Public Const IDH_ENTRADA_NF_SIMPLES_COMPLEMENTOS = 6072
Public Const IDH_ENTRADA_NF_SIMPLES_TRIBUTACAO = 6073
Public Const IDH_ENTRADA_NF_SIMPLES_CONTABILIZACAO = 6074
Public Const IDH_ENTRADA_NF_FATURA_DADOS_PRINCIPAIS = 6075
Public Const IDH_ENTRADA_NF_FATURA_ITENS = 6076
Public Const IDH_ENTRADA_NF_FATURA_COMPLEMENTOS = 6077
Public Const IDH_ENTRADA_NF_FATURA_PAGAMENTO = 6078
Public Const IDH_ENTRADA_NF_FATURA_TRIBUTACAO = 6079
Public Const IDH_ENTRADA_NF_FATURA_CONTABILIZACAO = 6080
Public Const IDH_ENTRADA_NF_REMESSA_DADOS_PRINCIPAIS = 6081
Public Const IDH_ENTRADA_NF_REMESSA_ITENS = 6082
Public Const IDH_ENTRADA_NF_REMESSA_COMPLEMENTO = 6083
Public Const IDH_ENTRADA_NF_REMESSA_TRIBUTACAO = 6084
Public Const IDH_ENTRADA_NF_REMESSA_CONTABILIZACAO = 6085
Public Const IDH_ENTRADA_NF_DEVOLUCAO_DADOS_PRINCIPAIS = 6086
Public Const IDH_ENTRADA_NF_DEVOLUCAO_ITENS = 6087
Public Const IDH_ENTRADA_NF_DEVOLUCAO_COMPLEMENTO = 6088
Public Const IDH_ENTRADA_NF_DEVOLUCAO_TRIBUTACAO = 6089
Public Const IDH_ENTRADA_NF_DEVOLUCAO_CONTABILIZACAO = 6090
'Public Const IDH_ANALISE_ESTOQUE = 6091
'Public Const IDH_ANALISE_MOVIMENTACOES_ESTOQUE = 6092
'Public Const IDH_BOLETIM_ENTRADA = 6093
'Public Const IDH_CUSTO_REPOSICAO = 6094
'Public Const IDH_ETIQUETAS_INVENTARIO = 6095
'Public Const IDH_LISTA_FALTAS = 6096
'Public Const IDH_KARDEX = 6097
'Public Const IDH_KARDEX_DIA_PARTE_1 = 6098
'Public Const IDH_KARDEX_DIA_PARTE_2 = 6099
'Public Const IDH_RESUMO_KARDEX_DIA = 6100
'Public Const IDH_RELACAO_KITS = 6101
'Public Const IDH_LISTAGEM_INVENTARIO = 6102
'Public Const IDH_RELACOES_MOVIMENTACOES_INTERNAS_PARTE_1 = 6103
'Public Const IDH_RELACOES_MOVIMENTACOES_INTERNAS_PARTE_2 = 6104
'Public Const IDH_RELACOES_MOVIMENTACOES_INTERNAS_PARTE_2 = 6105
Public Const IDH_RELACAO_MOVIMENTOS_ESTOQUE_PARA_ORDEM_PRODUCAO = 6106
Public Const IDH_ORDENS_PRODUCAO_PARTE_1 = 6107
Public Const IDH_ORDENS_PRODUCAO_PARTE_2 = 6108
Public Const IDH_PICK_LIST = 6109
Public Const IDH_PRODUTOS_ATINGIRAM_PONTO_PEDIDO = 6110
Public Const IDH_RELATORIO_PRODUTOS = 6111
Public Const IDH_RELACAO_PRODUTOS_VENDIDOS = 6112
Public Const IDH_RELACAO_CONSUMO_REAL_STANDARD = 6113
Public Const IDH_REQUISICOES_CONSUMO = 6114
Public Const IDH_RESUMO_ENTRADA_SAIDA_VALOR = 6115
Public Const IDH_SALDO_ESTOQUE = 6116
Public Const IDH_UTILIZACAO_PRODUTO = 6117
Public Const IDH_CONSUMO_VENDAS_MENSAIS = 6118
Public Const IDH_DEMONSTRATIVO_APURACAO_INVENTARIO = 6119
'Public Const IDH_SELECIONAR_RELATORIO = 6120
Public Const IDH_RELATORIO_CLASSIFICACAO_ABC = 6121
'Public Const IDH_ATUALIZACAO_LOTES = 6122
Public Const IDH_FECHAMENTO_MES_ESTOQUE = 6123
Public Const IDH_CUSTO_MEDIO_PRODUCAO = 6124
Public Const IDH_CLASSIFICACAO_ABC_CLASSIFICACAO = 6125
Public Const IDH_CLASSIFICACAO_ABC_CURVA_ABC = 6126
Public Const IDH_FORNECEDOR_CONSULTA = 6127
Public Const IDH_CANCELA_NFISCALEST = 6128
Public Const IDH_CONFIGURACAOEST = 6129
Public Const IDH_CONSUMO = 6130
Public Const IDH_CONSUMO_MEDIO = 6131
Public Const IDH_FALTA_ESTOQUE = 6132
Public Const IDH_FECHAMENTO_MES_ESTOQUE1 = 6133
Public Const IDH_FORNECEDOR_PRODUTO_COMPLEMENTO = 6134
Public Const IDH_GERACAO_OP_SELECAO = 6135
Public Const IDH_GERACAO_OP_PV = 6136
Public Const IDH_GERACAO_OP_GERADA = 6137
Public Const IDH_LOTE_EST = 6138
Public Const IDH_LOTE_EST_ATUALIZA = 6139
Public Const IDH_TIPOS_MOV_ESTOQUE = 6140
Public Const IDH_ESTOQUE = 6141
Public Const IDH_CONSUMO_RECALCULO = 6142
Public Const IDH_CUSTO_MEDIO = 6143
Public Const IDH_ESTOQUE_INICIAL1 = 6144
Public Const IDH_ESTOQUE_INICIAL2 = 6145
Public Const IDH_IMPLANTACAO_ESTOQUES = 6146
Public Const IDH_ESTOQUE_DADOS_PRINCIPAIS = 6147
Public Const IDH_ESTOQUE_PONTO_PEDIDO = 6148
Public Const IDH_FORNECEDOR_PRODUTO_ID = 6186
Public Const IDH_RASTROESTOQUEINICIAL = 6187
Public Const IDH_RASTREAMENTO_LOTE = 6188

'RelatoriosEST
Public Const IDH_RELOP_ANALISE_ESTOQUE_L = 7000
Public Const IDH_RELOP_ANALISE_ESTOQUE = 7001
Public Const IDH_RELOP_ANALISE_MOVIMENTO_ESTOQUE = 7002
Public Const IDH_RELOP_BOLETIM_ENTRADA = 7003
Public Const IDH_RELOP_CLASSIFICACAO_ABC = 7004
Public Const IDH_RELOP_CUSTO_REPOSICAO = 7005
Public Const IDH_RELOP_DEM_APURACAO_INVENTARIO_L = 7006
Public Const IDH_RELOP_DEM_APURACAO_INVENTARIO = 7007
Public Const IDH_RELOP_ESTOQUE_PRODUTO = 7008
Public Const IDH_RELOP_ETIQUETA_INVENTARIO = 7009
Public Const IDH_RELOP_FALTAS = 7010
Public Const IDH_RELOP_KARDEX_L = 7011
Public Const IDH_RELOP_KARDEX_DIA_L = 7012
Public Const IDH_RELOP_KARDEX_DIA = 7013
Public Const IDH_RELOP_KARDEX = 7014
Public Const IDH_RELOP_KIT = 7015
Public Const IDH_RELOP_LISTA_INVENTARIO = 7016
Public Const IDH_RELOP_MOVIMENTO_INTERNO = 7017
Public Const IDH_RELOP_MOVIMENTO_ESTOQUE_OP = 7018
Public Const IDH_RELOP_ORDEM_PRODUCAO = 7019
Public Const IDH_RELOP_PICK_LIST = 7020
Public Const IDH_RELOP_PONTO_PEDIDO = 7021
Public Const IDH_RELOP_PRODUTOS = 7022
Public Const IDH_RELOP_PRODUTOS_VENDIDOS = 7023
Public Const IDH_RELOP_CONSUMO_REAL_STANDARD = 7024
Public Const IDH_RELOP_REG_INVENTARIO_MOD7 = 7025
Public Const IDH_RELOP_REG_INVENTARIO_MOD7_L = 7026
Public Const IDH_RELOP_REQUISICAO_CONSUMO = 7027
Public Const IDH_RELOP_RESERVA = 7028
Public Const IDH_RELOP_RESUMO_ENT_SAIDA_VALOR = 7029
Public Const IDH_RELOP_RESUMO_ENT_SAIDA_VALOR_L = 7030
Public Const IDH_RELOP_RESUMO_KARDEX_L = 7031
Public Const IDH_RELOP_RESUMO_KARDEX = 7032
Public Const IDH_RELOP_SALDO_ESTOQUE = 7033
Public Const IDH_RELOP_SALDO_ESTOQUE_L = 7034
Public Const IDH_RELOP_USO_PRODUTO = 7035
Public Const IDH_RELOP_VENDAS_MES = 7036

'Manual Faturamento
'Public Const IDH_EXCECOES_ICMS = 8000
'Public Const IDH_EXCECOES_IPI = 8001
Public Const IDH_TIPOS_BLOQUEIO = 8002
Public Const IDH_CANAIS_VENDA = 8003
Public Const IDH_TRANSPORTADORA_ID = 8004
Public Const IDH_TRANSPORTADORA_ENDERECO = 8005
'Public Const IDH_NATUREZA_OPERACAO = 8006
'Public Const IDH_SERIE_NOTA_FISCAL = 8007
'Public Const IDH_PADROES_TRIBUTACAO_OPERACOES_FORNECEDORES = 8008
Public Const IDH_ITENS_TABELA_PRECO_EMPRESA_TODA = 8009
Public Const IDH_TABELA_PRECOS_CRIACAO = 8010
Public Const IDH_ITENS_TABELA_PRECO = 8011
Public Const IDH_TABELA_PRECOS = 8012
Public Const IDH_TABELA_PRECOS_ALTERACAO = 8013
Public Const IDH_TIPOS_CLIENTE_ID = 8014
Public Const IDH_TIPOS_CLIENTE_DADOS_FIN = 8015
Public Const IDH_TIPOS_CLIENTE_VENDAS = 8016
'Public Const IDH_REGIOES_VENDA = 8017
'Public Const IDH_CATEGORIAS_CLIENTE = 8018
'Public Const IDH_CLIENTES_ID = 8019
'Public Const IDH_CLIENTES_DADOS_FIN = 8020
'Public Const IDH_CLIENTES_INSCRICOES = 8021
'Public Const IDH_CLIENTES_ENDERECOS = 8022
'Public Const IDH_CLIENTES_VENDAS = 8023
'Public Const IDH_CLIENTES_ESTATISTICAS = 8024
'Public Const IDH_FILIAIS_CLIENTES_ID = 8025
'Public Const IDH_FILIAIS_CLIENTES_INSCRICOES = 8026
'Public Const IDH_FILIAIS_CLIENTES_ENDERECOS = 8027
'Public Const IDH_FILIAIS_CLIENTES_VENDAS = 8028
'Public Const IDH_TIPOS_VENDEDOR = 8029
Public Const IDH_VENDEDORES_ID = 8030
Public Const IDH_VENDEDORES_COMISSAO = 8031
Public Const IDH_VENDEDORES_ENDERECO = 8032
'Public Const IDH_CONDICOES_PAGAMENTO = 8033
Public Const IDH_MOVIMENTOS = 8034
Public Const IDH_NF_SAIDA_DADOS_PRINCIPAIS = 8035
Public Const IDH_NF_SAIDA_ITENS = 8036
Public Const IDH_NF_SAIDA_COMPLEMENTO = 8037
Public Const IDH_NF_SAIDA_COMISSOES = 8038
Public Const IDH_NF_SAIDA_ALMOXARIFADO = 8039
Public Const IDH_NF_SAIDA_TRIBUTACAO = 8040
Public Const IDH_NF_SAIDA_CONTABILIZACAO = 8041
Public Const IDH_LOCALIZACAO_PRODUTO = 8042
Public Const IDH_SUBSTITUICAO_PRODUTO = 8043
Public Const IDH_ATUALIZACAO_PRECOS_SELECAO = 8044
Public Const IDH_ATUALIZACAO_PRECOS_METODO_ATUALIZACAO = 8045
Public Const IDH_ALCADA_FAT = 8046
Public Const IDH_ALOCACAO_PRODUTO = 8047
Public Const IDH_ALOCACAO_PRODUTO1 = 8048
Public Const IDH_ALOCACAO_PRODUTO_SAIDA = 8049
Public Const IDH_ALOCACAO_PRODUTO_SAIDA1 = 8050
Public Const IDH_AUTORIZACAO_CREDITO = 8051
Public Const IDH_BAIXA_PEDIDO_SELECAO = 8052
Public Const IDH_BAIXA_PEDIDO_PEDIDOS = 8053
Public Const IDH_CANCELA_NFISCAL_SAIDA = 8054
Public Const IDH_CLIENTE_CONSULTA = 8055
Public Const IDH_COMISSOES = 8056
Public Const IDH_CONFIGURA_FAT = 8057
Public Const IDH_GERACAO_NFISCAL_SELECAO = 8058
Public Const IDH_GERACAO_NFISCAL_PEDIDOS = 8059
Public Const IDH_LIBERACAO_BLOQUEIO_SELECAO = 8060
Public Const IDH_LIBERACAO_BLOQUEIO_BLOQUEIOS = 8061
Public Const IDH_LOCALIZACAO_PRODUTO1 = 8062
Public Const IDH_NF_SAIDA_DEVOLUCAO_DADOS_PRINCIPAIS = 8063
Public Const IDH_NF_SAIDA_DEVOLUCAO_ITENS = 8064
Public Const IDH_NF_SAIDA_DEVOLUCAO_COMPLEMENTO = 8065
Public Const IDH_NF_SAIDA_DEVOLUCAO_TRIBUTACAO = 8066
Public Const IDH_NF_SAIDA_DEVOLUCAO_CONTABILIZACAO = 8067
Public Const IDH_NF_FATURA_SAIDA_DADOS_PRINCIPAIS = 8068
Public Const IDH_NF_FATURA_SAIDA_ITENS = 8069
Public Const IDH_NF_FATURA_SAIDA_COMPLEMENTO = 8070
Public Const IDH_NF_FATURA_SAIDA_COBRANCA = 8071
Public Const IDH_NF_FATURA_SAIDA_COMISSOES = 8072
Public Const IDH_NF_FATURA_SAIDA_ALMOXARIFADO = 8073
Public Const IDH_NF_FATURA_SAIDA_TRIBUTACAO = 8074
Public Const IDH_NF_FATURA_SAIDA_CONTABILIZACAO = 8075
Public Const IDH_NF_REMESSA_SAIDA_DADOS_PRINCIPAIS = 8076
Public Const IDH_NF_REMESSA_SAIDA_ITENS = 8076
Public Const IDH_NF_REMESSA_SAIDA_COMPLEMENTO = 8077
Public Const IDH_NF_REMESSA_SAIDA_TRIBUTACAO = 8078
Public Const IDH_NF_REMESSA_SAIDA_CONTABILIZACAO = 8079
Public Const IDH_NF_FATURA_PEDIDO_DADOS_PRINCIPAIS = 8080
Public Const IDH_NF_FATURA_PEDIDO_ITENS = 8081
Public Const IDH_NF_FATURA_PEDIDO_COMPLEMENTO = 8082
Public Const IDH_NF_FATURA_PEDIDO_COBRANCA = 8083
Public Const IDH_NF_FATURA_PEDIDO_COMISSOES = 8084
Public Const IDH_NF_FATURA_PEDIDO_ALMOXARIFADO = 8085
Public Const IDH_NF_FATURA_PEDIDO_TRIBUTACAO = 8086
Public Const IDH_NF_FATURA_PEDIDO_CONTABILIZACAO = 8087
Public Const IDH_NF_PEDIDO_DADOS_PRINCIPAIS = 8088
Public Const IDH_NF_PEDIDO_ITENS = 8089
Public Const IDH_NF_PEDIDO_COMPLEMENTO = 8090
Public Const IDH_NF_PEDIDO_COMISSOES = 8091
Public Const IDH_NF_PEDIDO_ALMOXARIFADO = 8092
Public Const IDH_NF_PEDIDO_TRIBUTACAO = 8093
Public Const IDH_NF_PEDIDO_CONTABILIZACAO = 8094
Public Const IDH_PEDIDO_VENDA_CONS_DADOS_PRINCIPAIS = 8095
Public Const IDH_PEDIDO_VENDA_CONS_ITENS = 8096
Public Const IDH_PEDIDO_VENDA_CONS_COMPLEMENTO = 8097
Public Const IDH_PEDIDO_VENDA_CONS_COBRANCA = 8098
Public Const IDH_PEDIDO_VENDA_CONS_COMISSOES = 8099
Public Const IDH_PEDIDO_VENDA_CONS_BLOQUEIO = 8100
Public Const IDH_PEDIDO_VENDA_CONS_ALMOXARIFADO = 8101
Public Const IDH_PEDIDO_VENDA_CONS_NOTAS_FISCAIS = 8102
Public Const IDH_PEDIDO_VENDA_CONS_TRIBUTACAO = 8103
Public Const IDH_PEDIDO_VENDA_CONS_CONTABILIZACAO = 8104
Public Const IDH_PEDIDO_VENDA_DADOS_PRINCIPAIS = 8105
Public Const IDH_PEDIDO_VENDA_ITENS = 8106
Public Const IDH_PEDIDO_VENDA_COMPLEMENTO = 8107
Public Const IDH_PEDIDO_VENDA_COBRANCA = 8108
Public Const IDH_PEDIDO_VENDA_COMISSOES = 8109
Public Const IDH_PEDIDO_VENDA_BLOQUEIO = 8110
Public Const IDH_PEDIDO_VENDA_ALMOXARIFADO = 8111
Public Const IDH_PEDIDO_VENDA_TRIBUTACAO = 8112
Public Const IDH_PEDIDO_VENDA_CONTABILIZACAO = 8113
Public Const IDH_PREV_VENDAS = 8114
Public Const IDH_SUBST_PRODUTO_NF = 8115
Public Const IDH_GERACAO_FATURA_SELECAO = 8116
Public Const IDH_GERACAO_FATURA_NF = 8117
Public Const IDH_GERACAO_FATURA_COBRANCA = 8118
Public Const IDH_GERACAO_FATURA_DESCONTOS = 8119
Public Const IDH_GERACAO_FATURA_TRIBUTACAO = 8120
Public Const IDH_GERACAO_FATURA_CONTABILIZACAO = 8121

'RelatoriosFAT
Public Const IDH_RELOP_CONTAS_CORRENTES = 9000
Public Const IDH_RELOP_CONTROLE_IMPRESSAO_FATURAS = 9001
Public Const IDH_RELOP_CONTROLE_IMPRESSAO_NF = 9002
Public Const IDH_RELOP_DESPACHO = 9003
Public Const IDH_RELOP_EMISSAO_DUPLICATAS = 9004
Public Const IDH_RELOP_EMISSAO_FATURAS = 9005
Public Const IDH_RELOP_EMISSAO_NF_FATURA = 9006
Public Const IDH_RELOP_EMISSAO_NF = 9007
Public Const IDH_RELOP_EMISSAO_NOTAS_REC = 9008
Public Const IDH_RELOP_ESTOQUE_VENDAS = 9009
Public Const IDH_RELOP_FAT_CLIENTE = 9010
Public Const IDH_RELOP_FAT_CLIENTE_PRODUTO = 9011
Public Const IDH_RELOP_FAT_CLIENTE_PRODUTO_L = 9012
Public Const IDH_RELOP_FAT_COMISSOES = 9013
Public Const IDH_RELOP_FAT_PICKLIST = 9014
Public Const IDH_RELOP_FAT_VENDEDOR = 9015
Public Const IDH_RELOP_NFISCAL_DEVOLUCAO = 9016
Public Const IDH_RELOP_NFISCAL_TRANSPORT = 9017
Public Const IDH_RELOP_NF = 9018
Public Const IDH_RELOP_PEDIDOS_APTOS_FATURAR = 9019
Public Const IDH_RELOP_PEDIDOS_NAO_ENTREGUES = 9020
Public Const IDH_RELOP_PEDIDOS_PRODUTO = 9021
Public Const IDH_RELOP_PEDIDO_VENDEDOR_CLIENTE = 9022
Public Const IDH_RELOP_PEDIDO_VENDEDOR_PRODUTO = 9023
Public Const IDH_RELOP_PRAZO_PAGTO = 9024
Public Const IDH_RELOP_PRECOS = 9025
Public Const IDH_RELOP_PRE_NOTA = 9026
Public Const IDH_RELOP_PREVISAO_VENDAS = 9027
Public Const IDH_RELOP_REAL_PREVISTO = 9028
Public Const IDH_RELOP_RECIBO_DESPACHO = 9029
Public Const IDH_RELOP_RESUMO_VENDAS = 9030
Public Const IDH_RELOP_TRANSPORTADORAS = 9031
Public Const IDH_RELOP_VISITAS = 9032
Public Const IDH_RELOP_FATPRODUTO = 9033
Public Const IDH_RELOP_PEDIDO_CLIENTE = 9034

'Manual Instalação
Public Const IDH_CONFIGURACAO_EMPRESA = 10000
Public Const IDH_SGE_GERAL = 10001
Public Const IDH_MODULO_CONTABILIDADE_ID = 10002
Public Const IDH_MODULO_CONTABILIDADE_CENTRO_CUSTO_LUCRO = 10003
Public Const IDH_MODULO_CONTABILIDADE_VALORES_INICIAIS = 10004
Public Const IDH_CONFIGURA_SEGMENTOS = 10005
Public Const IDH_MODULO_CONTABILIDADE = 10006
Public Const IDH_MODULO_TESOURARIA = 10007
Public Const IDH_MODULO_CONTAS_PAGAR = 10008
Public Const IDH_MODULO_CONTAS_RECEBER = 10009
Public Const IDH_MODULO_FATURAMENTO = 10010
Public Const IDH_MODULO_FATURAMENTO_CONTINUACAO = 10011
Public Const IDH_MODULO_ESTOQUE = 10012
'Public Const IDH_CONFIGURACAO_EMPRESA = 10011
Public Const IDH_CONFIGURACAO_FILIAL_EMPRESA = 10013
Public Const IDH_CONFIGURACAO_FILIAL_EMPRESA_EST = 10014
Public Const IDH_MODULO_COMPRAS = 10015
Public Const IDH_CONFIGURACAO_FILIAL_EMPRESA_COM = 10016

'Public Const IDH_MODULO_ESTOQUE = 10013

'Manual Tesouraria
Public Const IDH_BANCOS = 11000
Public Const IDH_CAIXA_CONTA_CORRENTE_INTERNA = 11001
Public Const IDH_FAVORECIDOS = 11002
'Public Const IDH_MANUTENCAO_LOTES = 11003
Public Const IDH_TIPOS_APLICACAO = 11004
Public Const IDH_HISTORICO_EXTRATO_CONTA_CORRENTE = 11005
Public Const IDH_SAQUE_ID = 11006
Public Const IDH_SAQUE_CONTABILIZACAO = 11007
Public Const IDH_DEPOSITO_ID = 11008
Public Const IDH_DEPOSITO_CONTABILIZACAO = 11009
Public Const IDH_TRANSFERENCIA_ID = 11010
'Public Const IDH_TRANSFERENCIA_CONTABILIZACAO = 11011
Public Const IDH_APLICACAO_ID = 11012
Public Const IDH_APLICACAO_COMPLEMENTO = 11013
Public Const IDH_APLICACAO_CONTABILIZACAO = 11014
Public Const IDH_RESGATE_ID = 11015
Public Const IDH_RESGATE_COMPLEMENTO = 11016
Public Const IDH_RESGATE_CONTABILIZACAO = 11017
Public Const IDH_CONCILIACAO_BANCARIA_SELECAO = 11018
Public Const IDH_CONCILIACAO_EXTRATO_CNAB = 11019
Public Const IDH_CONCILIACAO_BANCARIA_EXTRATO_PAPEL = 11020
'Public Const IDH_CONCILIACAO_BANCARIA_SELECAO = 11021
Public Const IDH_FLUXO_CAIXA_PAGAMENTO_TIPO_FORNECEDOR = 11022
Public Const IDH_FLUXO_CAIXA_PAGAMENTO_FORNECEDOR = 11023
Public Const IDH_FLUXO_CAIXA_PAGAMENTO_TITULO = 11024
Public Const IDH_FLUXO_CAIXA_RECEBIMENTO_TIPO_CLIENTE = 11025
Public Const IDH_FLUXO_CAIXA_RECEBIMENTOS_TIPO_CLIENTE = 11026
Public Const IDH_FLUXO_CAIXA_RECEBIMENTOS_TITULO = 11027
Public Const IDH_FLUXO_CAIXA_RESGATE_TIPO_APLICACAO = 11028
Public Const IDH_FLUXO_CAIXA_RESGATES = 11029
Public Const IDH_FLUXO_CAIXA_SALDOS_INICIAIS = 11030
Public Const IDH_FLUXO_SINTETICO_PROJETADO = 11031
Public Const IDH_FLUXO_SINTETICO_REVISADO = 11032
Public Const IDH_POSICAO_APLICACOES = 11033
Public Const IDH_MOVIMENTACAO_FINANCEIRA = 11034
Public Const IDH_MOVIMENTACAO_FINANCEIRA_DETALHADA = 11035
Public Const IDH_BANCOSINFO_COBRADOR = 11036
Public Const IDH_BANCOSINFO_CARTEIRA = 11037
Public Const IDH_BANCOSINFO_CODIGOSLANCAMENTO = 11038
Public Const IDH_CARTEIRAS_BANCO = 11039
Public Const IDH_CONFIGURA_TES = 11040
Public Const IDH_FLUXO_CAIXA_ID = 11041
Public Const IDH_FLUXO_CAIXA_SINTETICO = 11042
Public Const IDH_FLUXO_CAIXA_SALDADOSINICIAIS = 11043
Public Const IDH_FLUXO_CAIXA_PAGAMENTOS = 11044
Public Const IDH_FLUXO_CAIXA_RECEBIMENTOS = 11045
Public Const IDH_FLUXO_CAIXA_APLICACOES = 11046
Public Const IDH_FLUXO_CAIXA_COMISSOES = 11047
Public Const IDH_FLUXO_CAIXA_PEDVENDAS = 11048
Public Const IDH_FLUXO_CAIXA_PEDCOMPRAS = 11049
Public Const IDH_FLUXO_CAIXA_CHEQUEPRE = 11050
Public Const IDH_EXTRATO_BANCARIO_CNAB = 11051
Public Const IDH_EXTRATO_BANCARIO_CNAB2 = 11052
Public Const IDH_FLUXO_CAIXA = 11053
Public Const IDH_MENSAGEM = 11054
Public Const IDH_CONCILIAR_EXTRATO_BANCARIO = 11055

'Tributação
Public Const IDH_PADRAO_TRIB_ENTRADA = 12000
Public Const IDH_PADRAO_TRIB_SAIDA = 12001
Public Const IDH_TIPO_TRIBUTACAO = 12002

'Manual para as telas de Browse
Public Const IDH_BROWSE = 13100

'Type utilizado na Funcao Converte_Letra_para_UNC
Private Type NETRESOURCE
   dwScope As Long
   dwType As Long
   dwDisplayType As Long
   dwUsage As Long
   lpLocalName As Long
   lpRemoteName As Long
   lpComment As Long
   lpProvider As Long
End Type

'Constantes utilizadas na Funcao Converte_Letra_para_UNC
Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCE_CONNECTED = &H1
Private Const DRIVE_REMOTE = 4

'Funções API utilizadas na Funcao Converte_Letra_para_UNC
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lphEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, lpBuffer As Any, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long

Public Const NUM_MODULO = 14

Type typeProdutoFilial
    iFilialEmpresa As Integer
    sProduto As String
    iAlmoxarifado As Integer
    iFilialForn As Integer
    lFornecedor As Long
    iVisibilidadeAlmoxarifados As Integer
    dEstoqueSeguranca As Double
    iESAuto As Integer
    dEstoqueMaximo As Double
    iTemPtoPedido As Integer
    dPontoPedido As Double
    iPPAuto As Integer
    sClasseABC As String
    dLoteEconomico As Double
    iIntRessup As Integer
    iTempoRessup As Integer
    iTRAuto As Integer
    dTempoRessupMax As Double
    dConsumoMedio As Double
    iCMAuto As Integer
    dConsumoMedioMax As Double
    iMesesConsumoMedio As Integer
    dQuantPedida As Double
    iTabelaPreco As Integer
    sSituacaoTribECF As String
    sICMSAliquota As String
    dLoteMinimo As Double
    '??? Esse campo será retirado
    dICMSAliquota As Double
    iProdNaFilial As Integer
    dDescontoItem As Double
    dDescontoValor As Double
End Type

Type typeProduto
    sCodigo As String
    iTipo As Integer
    sDescricao As String
    sNomeReduzido As String
    sModelo As String
    iGerencial As Integer
    iNivel As Integer
    sSubstituto1 As String
    sSubstituto2 As String
    iPrazoValidade As Integer
    sCodigoBarras As String
    iEtiquetasCodBarras As Integer
    dPesoLiq As Double
    dPesoBruto As Double
    dComprimento As Double
    dEspessura As Double
    dLargura As Double
    sCor As String
    sObsFisica As String
    iClasseUM As Integer
    sSiglaUMEstoque As String
    sSiglaUMCompra As String
    sSiglaUMVenda As String
    iAtivo As Integer
    iFaturamento As Integer
    iCompras As Integer
    iPCP As Integer
    iKitBasico As Integer
    iKitInt As Integer
    dIPIAliquota As Double
    sIPICodigo As String
    sIPICodDIPI As String
'    dISSAliquota As Double
'    sISSCodigo As String
'    iIRIncide As Integer
    iControleEstoque As Integer
    iICMSAgregaCusto As Integer
    iIPIAgregaCusto As Integer
    iFreteAgregaCusto As Integer
    iApropriacaoCusto As Integer
    sContaContabil As String
    sContaContabilProducao As String
    dResiduo As Double
    iNatureza As Integer
    dCustoReposicao As Double
    iOrigemMercadoria As Integer
    sFCI As String
    iTabelaPreco As Integer
    dPercentMaisQuantCotAnt As Double
    dPercentMenosQuantCotAnt As Double
    iConsideraQuantCotAnt As Integer
    iTemFaixaReceb As Integer
    dPercentMaisReceb As Double
    dPercentMenosReceb As Double
    iRecebForaFaixa As Integer
    iTempoProducao As Integer
    iRastro As Integer
    lHorasMaquina As Long
    dPesoEspecifico As Double
    sReferencia As String
    sFigura As String
    sSituacaoTribECF As String
    iEmbalagem As Integer
    dINSSPercBase As Double
    dICMSAliquota As Double
    sICMSAliquota As String
    iKitVendaComp As Integer
    dPrecoLoja As Double
    dDesconto As Double
    sGrade As String
    dtDataLog As Date
    iUsaBalanca As Integer
    sSerieProx As String
    iSerieParteNum As Integer
    iPack As Integer
    iExTIPI As Integer
    iProdutoEspecifico As Integer
    sGenero As String
    sISSQN As String
    sSiglaUMTrib As String
    sCodServNFe As String
    dPrecoMaxConsumidor As Double
    iTipoServico As Integer
    lFabricante As Long
    iIPIIncide As Integer
    dQtdeEmbBase As Double
    dFatorAjuste As Double
    sNBS As String
    sTruncamento As String
    dQuantEstLoja As Double
    dPercComissao As Double
    dMetaComissao As Double
    sCEST As String
    iProdEmEscalaRelev As Integer
    scProdANVISA As String
End Type

Type typeUsuario
    sCodUsuario As String
    iLote As Integer
    sNome As String
    sNomeReduzido As String
    iAtivo As Integer

End Type

Type typeBrowseUsuario

    sNomeTela As String
    sCodUsuario As String
    lTopo As Long
    lEsquerda As Long
    lLargura As Long
    lAltura As Long
    
End Type

Type typeGrade

    sCodigo As String
    sDescricao As String
    iLayout As Integer
    
End Type

Type typeAnotacoes
    lNumIntDoc As Long
    iTipoDocOrigem As Integer
    sID As String
    sTitulo As String
    dtDataAtualizacao As Date
End Type

Type typeOrigemAnotacoes
    iCodigo As Integer
    sDescricao As String
    sNomeTabela As String
End Type

Type typeCidades
    iCodigo As Integer
    sDescricao As String
    sCodIBGE As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeValidaExclusoes
    sCodigo As String
    sTabela As String
    sCampo As String
    sCampoLer As String
    sMsgErro1 As String
    sMsgErro2 As String
    iTipoCampoLer As Integer
    iTamanhoCampoLer As Integer
    sMsgErroLer As String
    iSubTipoCampoLer As Integer
    iSubTipoCampoProc As Integer
    iGeneroMsgErro As Integer
End Type

Type typeTransacaoWFW
    sSigla As String
    sTransacao As String
    sTransacaoTela As String
    sOrigem As String
    sObservacao As String
    iCodigo As Integer
End Type

Type typeMnemonicoWFW
    sModulo As String
    iTransacao As Integer
    sMnemonico As String
    iTipo As Integer
    iNumParam As Integer
    iParam1 As Integer
    iParam2 As Integer
    iParam3 As Integer
    sNomeGrid As String
    sMnemonicoCombo As String
    sMnemonicoDesc As String
End Type

Type typeRegraWFW
    sModulo As String
    iTransacao As Integer
    iItem As Integer
    sCodUsuario As String
    sRegra As String
    iTipoBloqueio As Integer
    sEmailPara As String
    sEmailAssunto As String
    sEmailMsg As String
    sAvisoMsg As String
    sLogDoc As String
    sLogMsg As String
    dtDataUltExec As Date
    dHoraUltExec As Double
    sRelModulo As String
    sRelNome As String
    sRelOpcao As String
    sBrowseModulo As String
    sBrowseNome As String
    sBrowseOpcao As String
    iRelPorEmail As Integer
    sRelSel As String
    sRelAnexo As String
End Type

Type typeAvisoWFW
    lNumIntDoc As Long
    sMsg As String
    dtData As Date
    dHora As Double
    sUsuario As String
    iTransacao As Integer
    dtDataUltAviso As Date
    dHoraUltAviso As Double
    dIntervalo As Double
    sUsuarioOrig As String
    sTransacaoTela As String
    iUMIntervalo As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeCodServMun
    sCodIBGE As String
    sCodServ As String
    sISSQNBase As String
    iPadraoISSQN As Integer
    sDescricao1 As String
    sDescricao2 As String
    dAliquota As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeCodTribMun
    lCidade As Long
    sProduto As String
    sCodTribMun As String
    dAliquota As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeProdutoCNAE
    sProduto As String
    sCNAE As String
    sISSQN As String
    iLocServCliente As Integer
    iLocIncidImpCliente As Integer
    stpServico As String
    scodAtivEcon As String
End Type

Type typeNBS
    sCodigo As String
    sDescricao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeBackupConfig
    lCodigo As Long
    sDescricao As String
    iHabilitado As Integer
    dtDataInicio As Date
    dHora As Double
    iRepetirDias As Integer
    sDiretorio As String
    iIncluirDataNomeArq As Integer
    dtDataUltBkp As Date
    dtDataProxBkp As Date
    iTransfFTP As Integer
    iCompactar As Integer
    sFTPURL As String
    sFTPUsu As String
    sFTPSenha As String
    sFTPDir As String
    sDirDownload As String
End Type

Type typeTelaUsuario

    sNomeTela As String
    sCodUsuario As String
    lTopo As Long
    lEsquerda As Long
    lLargura As Long
    lAltura As Long
    
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeAnexos
    iTipoDoc As Integer
    lNumIntDoc As Long
    iSeq As Integer
    sArquivo As String
    sDescricao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeIdiomas
    iCodigo As Integer
    sDescricao As String
    sSigla As String
    iPadrao As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeIdiomaTelaControles
    sNomeTabela As String
    sNomeCampo As String
    sNomeTela As String
    sNomeControle As String
    sNomeTelaExibicao As String
    sNomeControleExibicao As String
    iComMaxLen As Integer
    iComMultiLine As Integer
    iEmGrid As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeIdiomaTextos
    sNomeTabela As String
    sNomeCampo As String
    sChaveDocS As String
    lChaveDocL As Long
    iChaveDocI As Integer
    iIdioma As Integer
    iSeq As Integer
    sTexto As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeConfigOutros
    sTabela As String
    sCodigo As String
    sModuloExibicao As String
    sNomeObj As String
    sNomeProperty As String
    sDescricaoGrid As String
    sTipoControle As String
    lCodVlrValidos As Long
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeConfigValoresValidos
    lCodigo As Long
    iSeq As Integer
    sValor As String
    sDescricao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeAvisos
    lCodigo As Long
    dtData As Date
    sAssunto As String
    sLink As String
    iPrioridade As Integer
    iForcaAberturaTela As Integer
End Type


'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeAvisosUsu
    lCodAviso As Long
    sCodUsuario As String
    iLido As Integer
    iExcluido As Integer
End Type



Public Sub gError(lErro As Long)

On Error GoTo Erro_gError

    gErr = lErro
    Error ERRO_FUNCAO_GERROR

    Exit Sub
    
Erro_gError:

    Error ERRO_FUNCAO_GERROR

    Exit Sub

End Sub

Public Function Converte_Letra_para_UNC(DriveLetter As String, sUNC As String) As Long
   
Dim hEnum As Long
Dim NetInfo(1023) As NETRESOURCE
Dim entries As Long
Dim nStatus As Long
Dim LocalName As String
Dim UNCName As String
Dim i As Long
Dim r As Long
Dim lTypedrive As Long
Dim iEncontrou As Long
Dim DriveLetterAux As String

On Error GoTo Erro_Converte_Letra_para_UNC
    
    'Prepara a String
    DriveLetterAux = Mid(DriveLetter, 3)
    DriveLetter = Mid(DriveLetter, 1, 2)
    
    lTypedrive = GetDriveType(DriveLetter)
        
    If lTypedrive = DRIVE_REMOTE Then
        
        iEncontrou = 0
        
        'Begin the enumeration
        nStatus = WNetOpenEnum(RESOURCE_CONNECTED, RESOURCETYPE_ANY, 0&, ByVal 0&, hEnum)

        'Check for success from open enum
        If ((nStatus = 0) And (hEnum <> 0)) Then
            
            ' Set number of entries
            entries = 1024
    
            'Enumerate the resource
            nStatus = WNetEnumResource(hEnum, entries, NetInfo(0), CLng(Len(NetInfo(0))) * 1024)
    
            'Check for success
            If nStatus = 0 Then
                For i = 0 To entries - 1
                    'Get the local name
                    LocalName = ""
                    If NetInfo(i).lpLocalName <> 0 Then
                        LocalName = Space(lstrlen(NetInfo(i).lpLocalName) + 1)
                        r = lstrcpy(LocalName, NetInfo(i).lpLocalName)
                    End If
    
                    ' Strip null character from end
                    If Len(LocalName) <> 0 Then
                        LocalName = left(LocalName, (Len(LocalName) - 1))
                    End If
    
                    If UCase$(LocalName) = UCase$(DriveLetter) Then
                        'Get the remote name
                        UNCName = ""
                        If NetInfo(i).lpRemoteName <> 0 Then
                            UNCName = Space(lstrlen(NetInfo(i).lpRemoteName) + 1)
                            r = lstrcpy(UNCName, NetInfo(i).lpRemoteName)
                        End If
    
                        ' Strip null character from end
                        If Len(UNCName) <> 0 Then
                            UNCName = left(UNCName, (Len(UNCName) - 1))
                        End If
    
                        'Return the UNC path to drive
                        sUNC = UNCName & DriveLetterAux
                        iEncontrou = 1
                        'Exit the loop
                        Exit For
                    End If
                Next i
            End If
        End If
    Else
        sUNC = DriveLetter & DriveLetterAux
        Exit Function
    End If
    
    'End enumeration
    nStatus = WNetCloseEnum(hEnum)
    
    If iEncontrou = 0 Then Error 64070
    
    Converte_Letra_para_UNC = SUCESSO
    
    Exit Function
    
Erro_Converte_Letra_para_UNC:
    
    Converte_Letra_para_UNC = Err
    
    Select Case Err
    
        Case 64070 'Não encontrou ou houve algum problema
        
    End Select
    
    Exit Function
    
End Function

Public Property Let gErr(ByVal vData As Long)
    
    gl_UltimoErro = vData

End Property

Public Property Get gErr() As Long

    If Err = ERRO_FUNCAO_GERROR Then
        gErr = gl_UltimoErro
    Else
        gErr = Err
    End If
        
End Property

Public Function String_Igual(ByVal sTexto1 As String, ByVal sTexto2 As String) As Boolean
    If Trim(UCase(sTexto1)) = Trim(UCase(sTexto2)) Then
        String_Igual = True
    Else
        String_Igual = False
    End If
End Function
