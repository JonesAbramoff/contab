Attribute VB_Name = "GlobalCRFAT"
Option Explicit

Public Const STRING_VERSAONFE_DESCRICAO = 50
Public Const STRING_TIPOFRETE_DESCRICAO = 50

Public Const STRING_RPWSEBCONSSITLOTE_MSG = 200
Public Const STRING_RPSWEBLOTELOG_STATUS = 255


Public Const FORMATO_CODIGO_BARRAS_BCO_TITULO = "#####.#####.#####.######.#####.######.#.#################"

Public Const STRING_SPEDFIS_NUMPROC = 15
Public Const STRING_SPEDFIS_CODIGODA = 1
Public Const STRING_SPEDFIS_NUMDA = 50
Public Const STRING_SPEDFIS_CODAUT = 100
Public Const STRING_SPEDFIS_INDOPER = 1
Public Const STRING_SPEDFIS_INDEMIT = 1
Public Const STRING_SPEDFIS_CODMOD = 2
Public Const STRING_SPEDFIS_SERIE = 4
Public Const STRING_SPEDFIS_ECFFAB = 20

Public Const STRING_NFE_CODVERIFICACAO = 8
Public Const STRING_NFE_NUMGUIA = 12
Public Const STRING_NFE_ISSRETIDO = 1
Public Const STRING_NFE_CHNFE = 44

Public Const TIPOPAGTO_DEPOSITO = 1
Public Const TIPOPAGTO_CHEQUE = 2
Public Const TIPOPAGTO_PREFERENCIAL = 3
Public Const TIPOPAGTO_CARTAO = 4

Public Const EMAIL_PADRAO_TIPO_ENVIO = "EMAIL_PADRAO_TIPO_ENVIO"
Public Const EMAIL_PADRAO_COBREBEMX_ARQLIC = "EMAIL_PADRAO_COBREBEMX_ARQLIC"
Public Const EMAIL_PADRAO_COBREBEMX_LOGO = "EMAIL_PADRAO_COBREBEMX_LOGO"
Public Const EMAIL_PADRAO_COBREBEMX_IMAGENS = "EMAIL_PADRAO_COBREBEMX_IMAGENS"
Public Const EMAIL_PADRAO_COBREBEMX_URLIMGCB = "EMAIL_PADRAO_COBREBEMX_URLIMGCB"
Public Const EMAIL_PADRAO_COBREBEMX_URLLOGO = "EMAIL_PADRAO_COBREBEMX_URLLOGO"
Public Const EMAIL_PADRAO_COBREBEMX_AGENCIA = "EMAIL_PADRAO_COBREBEMX_AGENCIA"
Public Const EMAIL_PADRAO_COBREBEMX_CONTA = "EMAIL_PADRAO_COBREBEMX_CONTA"

Public Const EMAIL_PADRAO_SMTP_HOST = "EMAIL_PADRAO_SMTP_HOST"
Public Const EMAIL_PADRAO_USUARIO = "EMAIL_PADRAO_USUARIO"
Public Const EMAIL_PADRAO_SENHA = "EMAIL_PADRAO_SENHA"

'para parcelas a receber
Public Const TIPO_ACEITE_AMBOS = 0
Public Const TIPO_ACEITE_SEM = 1
Public Const TIPO_ACEITE_COM = 2

Public Const RELACCLI_TIPODOC_PARCELARECEBER = 0
Public Const RELACCLI_TIPODOC_OV = 1
Public Const RELACCLI_TIPODOC_SOLSRV = 2

Public Const CONTRATO_ATIVO = 1
Public Const CONTRATO_INATIVO = 0

Public Const TIPO_RELACIONAMENTO_COBRANCA = 20
Public Const TIPO_RELACIONAMENTO_CONTATO = 21
Public Const TIPO_RELACIONAMENTO_OVACOMP = 22
Public Const TIPO_RELACIONAMENTO_SOLSRV = 23

Public Const PERIODICIDADE_CPR_SEMANAL = 1
Public Const PERIODICIDADE_CPR_QUINZENAL = 2
Public Const PERIODICIDADE_CPR_MENSAL = 3
Public Const PERIODICIDADE_CPR_BIMESTRAL = 4
Public Const PERIODICIDADE_CPR_TRIMESTRAL = 5
Public Const PERIODICIDADE_CPR_QUADRIMESTRAL = 6
Public Const PERIODICIDADE_CPR_SEMESTRAL = 7
Public Const PERIODICIDADE_CPR_ANUAL = 8

Public Const DEVCHEQUE_HISTORICO_CANCELAMENTO = "Devolução do Cheque com Sequencial "
Public Const EXCLUSAO_BORDDESCCHQ_HISTORICO_CANCELAMENTO = "Exclusão Bordero Desconto. Cheque Sequencial "
Public Const EXCLUSAO_BORDCHQPRE_HISTORICO_CANCELAMENTO = "Exclusão Bordero Cheque Pré. Cheque Sequencial """

Public Const STRING_RAZAO_SOCIAL = 40

'Constantes usadas no grids Analise_lin e Dvv
Public Const COLUNAS_GERAIS_ANALISE_LIN = 5
Public Const COLUNAS_GERAIS_DVV = 1

'Numero de campos de Formulas no Grid Analise_lin
Public Const MAX_NUM_FORMULAS_ANALISELIN = 8

'Numero de campos de Formulas no Grid Dvv
Public Const MAX_NUM_FORMULAS_DVV = 3

Public Const STRING_CLASSIFICACAOFISCAL_DESCRICAO = 255

Public Const TRANSFERIDO = 1
Public Const NAO_TRANSFERIDO = 0

Public Const TIPOPLAN_CUSTOIMPORTACAO = 3

Public Const STRING_DVVLIN_DESCRICAO = 50
Public Const STRING_PLANMARGCONTRLIN_DESCRICAO = 255
Public Const STRING_PLANMARGCONTRCOL_TITULO = 50

'#####################################
'Inserido por Wagner
Public Const CONTRATOS_PAGAR = 2
Public Const CONTRATOS_RECEBER = 1
'#####################################
'#######################################################################
'INSERIDO POR WAGNER
'#######################################################################
'Contrato
Public Const STRING_CONTRATO_CODIGO = 10
Public Const STRING_CONTRATO_DESCRICAO = 50
Public Const STRING_CONTRATO_OBSERVACAO = 255
Public Const STRING_CONTRATO_CONTACONTABIL = 20
Public Const STRING_CONTRATO_HISTORICOCONTABIL = 50
Public Const STRING_CONTRATO_CCL = 10

'Itens de Contrato
Public Const STRING_ITENSDECONTRATO_PRODUTO = 20
Public Const STRING_ITENSDECONTRATO_DESCPROD = 250
Public Const STRING_ITENSDECONTRATO_UM = 5
Public Const STRING_ITENSDECONTRATO_OBSERVACAO = 255

Public Const STRING_COBRADO = "Cobrado"
Public Const STRING_NAO_COBRADO = "Não Cobrado"

Public Const COBRADO = 1
Public Const NAO_COBRADO = 2

Public Const STRING_PERIODICIDADE_LIVRE = "Livre"
Public Const STRING_PERIODICIDADE_SEMANAL = "Semanal"
Public Const STRING_PERIODICIDADE_DECENDIAL = "Decendial"
Public Const STRING_PERIODICIDADE_QUINZENAL = "Quinzenal"
Public Const STRING_PERIODICIDADE_MENSAL = "Mensal"
Public Const STRING_PERIODICIDADE_BIMESTRAL = "Bimestral"
Public Const STRING_PERIODICIDADE_TRIMESTRAL = "Trimestral"
Public Const STRING_PERIODICIDADE_QUADRIMESTRAL = "Quadrimestral"
Public Const STRING_PERIODICIDADE_SEMESTRAL = "Semestral"
Public Const STRING_PERIODICIDADE_ANUAL = "Anual"
'#######################################################################
'FIM
'#######################################################################

'p/ gravação de registros envolvidos com a tela MargContr
Public Const MARGCONTR_GRIDDVV = 1
Public Const MARGCONTR_GRIDANALISE = 0

'nº máx de linhas dos grids da tela MargContr
Public Const GRIDDVV_MAX_LINHAS = 100
Public Const GRIDANALISE_MAX_LINHAS = 100

'??? SHirley 21/01/2003
Public Const VERIFICAR_LIMITE_CREDITO = 1

Public Const CLIENTE_ATIVO = 1

'Constante utilizada na tabela ChequePre para indicar se o bordero associado é um bordero de cheque pre ou de desconto.
Public Const TIPO_BORDERO_CHEQUEPRE = 1
Public Const TIPO_BORDERO_DESCONTO = 2
Public Const TIPO_BORDERO_CHEQUEPRE_TEXTO = "Cheque Pré"
Public Const TIPO_BORDERO_DESCONTO_TEXTO = "Desconto"

'tamanho máximo da string
Public Const MAX_REGISTRO_LOG = 1020
Public Const STRING_LOG = 255

'Constante para a Gravação de Log da Tela de VendedorFilial
Public Const ALTERACAO_VENDEDORLOJA_FILIAL = 40

Public Const TITULO_PAGAR_RECIBO_CARRETEIRO = "FC"
Public Const TITULO_PAGAR_DARF = "DARF"

Public Const STRING_FORMACAOPRECO_EXPRESSAO = 255
Public Const STRING_FORMACAOPRECO_TITULO = 255

'Constantes para identificação da VIA DE TRANSPORTE
Public Const TRANSPORTE_AEREO = 1
Public Const TRANSPORTE_AEREO_DESC As String = "Aéreo"
Public Const TRANSPORTE_MARITIMO = 2
Public Const TRANSPORTE_MARITIMO_DESC As String = "Marítimo"
Public Const TRANSPORTE_RODOVIARIO = 3
Public Const TRANSPORTE_RODOVIARIO_DESC As String = "Rodoviário"
Public Const TRANSPORTE_FERROVIARIO = 4
Public Const TRANSPORTE_FERROVIARIO_DESC As String = "Ferroviário"
Public Const TRANSPORTE_FLUVIAL = 5
Public Const TRANSPORTE_FLUVIAL_DESC As String = "Fluvial"
Public Const TRANSPORTE_AEREOFLUVIAL = 6
Public Const TRANSPORTE_AEREOFLUVIAL_DESC As String = "Aéreo / Fluvial"
Public Const TRANSPORTE_RODOVIARIOFERROVIARIO = 7
Public Const TRANSPORTE_RODOVIARIOFERROVIARIO_DESC As String = "Rodoviário / Ferroviário"
Public Const TRANSPORTE_OUTROS_DESC As String = "Outros"
'**************************************************

'Constante que indica se a empresa utiliza regras no calculo de comissoes
Public Const USA_REGRAS = 1

'Constante de Log
Public Const STRING_CONCATENACAO = 255

'Constantes para Log de Cliente
Public Const ALTERACAO_CLIENTE = 55
Public Const INCLUSAO_CLIENTE = 54
Public Const EXCLUSAO_CLIENTE = 55
Public Const INCLUSAO_FILIALCLIENTE = 78
Public Const ALTERACAO_FILIALCLIENTE = 79

'Tabela de FormacaoPreco campo Escopo
Public Const FORMACAO_PRECO_ESCOPO_GERAL = 1
Public Const FORMACAO_PRECO_ESCOPO_CATEGORIA = 2
Public Const FORMACAO_PRECO_ESCOPO_PRODUTO = 3
Public Const FORMACAO_PRECO_ESCOPO_TABPRECO = 4

'Tabela de FormacaoPreco campo Escopo
Public Const MNEMONICOFPRECO_ESCOPO_GERAL = 1
Public Const MNEMONICOFPRECO_ESCOPO_CATEGORIA = 2
Public Const MNEMONICOFPRECO_ESCOPO_PRODUTO = 3
Public Const MNEMONICOFPRECO_ESCOPO_TABPRECO = 4

Public Const CATEGORIAPRODUTO_PRECO = "Preço"

Public Const TIPODOCINFO_CONHECIMENTOFRETE = 115
Public Const TIPODOCINFO_CONHECIMENTOFRETE_FATURA = 116
Public Const TIPODOCINFO_CONHECIMENTOFRETE_COMPLICMS = 172
Public Const TIPODOCINFO_CONHECIMENTOFRETE_FATURA_COMPLICMS = 173

Public Const BLOQUEIO_AUTO_RESP = "BLOQUEIO AUTOMÁTICO"
'Janaina
Public Const DESBLOQUEIO_AUTO_RESP = "DESBLOQUEIO AUTOMÁTICO"
'Janaina

'Public Const STRING_TRANSPORTADORA_NOME = 50
'Public Const STRING_TRANSPORTADORA_NOME_REDUZIDO = 20

'Relatórios de emissão de N.Fiscal
Public Const INTERVALO_MONITORAMENTO_IMPRESSAO_FATURA = 1000 'Milisegundos
Public Const TEMPO_MAX_IMPRESSAO_UMA_FATURA = 7000 'Milisegundos
Public Const RELATORIO_FATURA_LOCKADO = 1
Public Const RELATORIO_FATURA_NAO_LOCKADO = 0
Public Const RELATORIO_FATURA_IMPRIMINDO = 1
Public Const RELATORIO_FATURA_NAO_IMPRIMINDO = 0
Public Const RELATORIO_NF_LOCKADO = 1
Public Const RELATORIO_NF_NAO_LOCKADO = 0
Public Const RELATORIO_NF_IMPRIMINDO = 1
Public Const RELATORIO_NF_NAO_IMPRIMINDO = 0
Public Const INTERVALO_MONITORAMENTO_IMPRESSAO_NF = 1000 'está em milisegundos
Public Const TEMPO_MAX_IMPRESSAO_UMA_NF = 10000 'está em milisegundos

Public Const STRING_CONSULTA = 50
Public Const STRING_CONSULTA_DESCRICAO = 100
Public Const STRING_OBSERVACAO_OBSERVACAO = 255

'Limites versão light
Public Const LIMITE_TRANSP_VLIGHT = 5

Public Const SERIE_VERSAO_LIGHT = "1"

Public Const TRIB_NEM_CREDITO_NEM_DEBITO = 0
Public Const TRIB_GERA_CREDITO = 1
Public Const TRIB_GERA_DEBITO = 2

Public Const STRING_DEB_REC_SIGLA = 4
Public Const STRING_DEB_REC_OBS = 255

'Carteira Cobrança
Public Const CARTCOBR_PARA_BANCOS = 0
Public Const CARTCOBR_PARA_OUTROS = 1
Public Const CARTCOBR_PARA_AMBOS = 2
Public Const CARTCOBR_PARA_EMPRESA = 3 'só vale para a própria Empresa

'para tabela OcorrenciasRemParcRec
Public Const STRING_NUMTITCOBRADOR = 20

'Tipos de Frete
Public Const TIPO_CIF = 0
Public Const TIPO_FOB = 1
Public Const TIPO_FRETE_TERCEIROS = 2
Public Const TIPO_SEM_FRETE = 9

'para tabela CRFATConfig
Public Const STRING_CRFATCONFIG_CODIGO = 50
Public Const STRING_CRFATCONFIG_DESCRICAO = 150
Public Const STRING_CRFATCONFIG_CONTEUDO = 255

'Vias de Transporte - Transportadora
Public Const VIA_TRANSP_AEREO = 1
Public Const VIA_TRANSP_MARITIMO = 2
Public Const VIA_TRANSP_RODOVIARIO = 3
Public Const VIA_TRANSP_FERROVIARIO = 4
Public Const VIA_TRANSP_FLUVIAL = 5
Public Const VIA_TRANSP_AEREO_FLUVIAL = 6
Public Const VIA_TRANSP_RODOVIARIO_FERROVIARIO = 7
Public Const VIA_TRANSP_OUTROS = 8

'Para geração de Conta a Pagar/Receber
Public Const CPR_TITULO_PAGAR = 1
Public Const CPR_TITULO_RECEBER = 2
Public Const CPR_CREDITO_PAGAR = 3
Public Const CPR_DEBITO_RECEBER = 4
Public Const CPR_NF_PAGAR = 5

Public Const TITULO_RECEBER = 2

Public Const COBRADOR_PROPRIA_EMPRESA = 1
Public Const CARTEIRA_CARTEIRA = 1
Public Const CARTEIRA_CHEQUEPRE = 2
Public Const CARTEIRA_JURIDICO = 3
Public Const CARTEIRA_SIMPLES = 4
Public Const CARTEIRA_CAUCIONADA = 5
Public Const CARTEIRA_VINCULADA = 6
Public Const CARTEIRA_DESCONTADA = 7
Public Const CARTEIRA_LOJA = 8

Public Const COBRANCA_NAO_ELETRONICA = 0
Public Const COBRANCA_ELETRONICA = 1

Public Const NUM_MAX_ITENS_CATEGORIA = 500
'Indica o número máximo de Itens para uma Categoria

Public Const NUM_MAXIMO_DESCONTOS = 3
Public Const NUM_MAXIMO_COMISSOES = 10

Public Const OUTROS_RECEBIMENTOS = "OR"

'Periodo de verificação de duplicidade de Notas Fiscais/Notas Recebimento e outros Títulos
Public Const PERIODO_EMISSAO = 180  'periodo em dias

Public Const GRID_CATEGORIA_COL = 1
Public Const GRID_VALOR_COL = 2

'tabela CRFatConfig
Public Const NUM_PROX_ITEM_NOTA_FISCAL = "NUM_PROX_ITEM_NOTA_FISCAL"
Public Const NUM_PROX_NOTA_FISCAL = "NUM_PROX_NOTA_FISCAL"
Public Const NUM_PROX_TABELA_PRECO = "NUM_PROX_TABELA_PRECO"

'Porcent. Juros Máxima
Public Const PORCENTAGEM_JUROS_MAXIMA As Long = 1000

Public Const EMITENTE_EMPRESA = 0
Public Const EMITENTE_CLIENTE = 1
Public Const EMITENTE_FORNECEDOR = 2

Public Const DESTINATARIO_EMPRESA = 0
Public Const DESTINATARIO_CLIENTE = 1
Public Const DESTINATARIO_FORNECEDOR = 2

Public Const DOCINFO_ORIGEM_EMPRESA = 0
Public Const DOCINFO_ORIGEM_CLIENTE = 1
Public Const DOCINFO_ORIGEM_FORNECEDOR = 2

Public Const TIPODOCINFO_TIPO_PV = 0
Public Const TIPODOCINFO_TIPO_NFIE = 1
Public Const TIPODOCINFO_TIPO_NFIS = 2
Public Const TIPODOCINFO_TIPO_NFEXT = 3
Public Const TIPODOCINFO_TIPO_RECFORN = 4
Public Const TIPODOCINFO_TIPO_RECCLI = 5

'Tipos de Documento
Public Const TIPODOC_CREDITOS_A_PAGAR As String = "CP"
Public Const TIPODOC_FATURA_A_PAGAR As String = "FP"
Public Const TIPODOC_FATURA_A_RECEBER As String = "FR"
Public Const TIPODOC_NOTA_CREDITO As String = "NC"
Public Const TIPODOC_NF_DEVOLUCAO As String = "NFDV"
Public Const TIPODOC_NF_FATURA_PAGAR As String = "NFFP"
Public Const TIPODOC_NF_FATURA_RECEBER As String = "NFFR"
Public Const TIPODOC_NF_A_PAGAR As String = "NFP"
Public Const TIPODOC_NF_A_RECEBER As String = "NFR"
Public Const TIPODOC_PAGAMENTO_ANTECIPADO As String = "PA"
Public Const TIPODOC_RECEBIMENTO_ANTECIPADO As String = "RA"
Public Const TIPODOC_CREDITOSPAGFORN As String = "CPF"
Public Const TIPODOC_CREDITOSRECCLI As String = "CRC"
Public Const TIPODOC_NF_FAT_SERVICO_PAGAR As String = "NFSP"
Public Const TIPODOC_NF_FAT_SERVICO_RECEBER As String = "NFSR"
Public Const TIPODOC_LOJA As String = "TL"
Public Const TIPODOC_DEV_CHQ As String = "DCHQ"
Public Const TIPODOC_CONTRATO_REC As String = "CTRR"
Public Const TIPODOC_CONTRATO_PAG As String = "CTRP"
Public Const TIPODOC_FATURA_SERVICO_CR As String = "FSCR"
Public Const TIPODOC_FATURA_SERVICO_CP As String = "FSCP"
Public Const TIPODOC_CARTAO_CRED_DEB As String = "CRT"
Public Const TIPODOC_FATURA_OVER As String = "OVER" 'Travel Ace
Public Const TIPODOC_FATURA_OCR_COBR As String = "OCRC" 'Travel Ace
Public Const TIPODOC_FATURA_OCR_JURI As String = "OCRJ" 'Travel Ace
Public Const TIPODOC_FATURA_OCR_REEM As String = "OCRR" 'Travel Ace
Public Const TIPODOC_PV As String = "PV"
Public Const TIPODOC_PC As String = "PC"
Public Const TIPODOC_IMPORTACAO_PAG As String = "IMPP"
Public Const TIPODOC_IMPORTACAO_REC As String = "IMPR"
Public Const TIPODOC_NFE_FATURA_PAGAR As String = "NFEP"
Public Const TIPODOC_NFE_FATURA_RECEBER As String = "NFER"

Public Const TRIB_ENTRADA_CLI = 0
Public Const TRIB_ENTRADA_FORN = 1
Public Const TRIB_SAIDA_CLI = 2
Public Const TRIB_SAIDA_FORN = 3

Public Const STRING_TABELAPRECO_DESCRICAO = 50

'para tabela NFiscal
'Public Const STRING_NFISCAL_SERIE = 3
Public Const STRING_NFISCAL_MENSAGEM = 250
Public Const STRING_NFISCAL_PLACA = 10
Public Const STRING_NFISCAL_PLACA_UF = 2
Public Const STRING_NFISCAL_VOLUME_ESPECIE = 20
Public Const STRING_NFISCAL_VOLUME_MARCA = 20
Public Const STRING_NFISCAL_VOLUME_NUMERO = 20
Public Const STRING_NUM_PEDIDO_TERC = 20
Public Const STRING_NFISCAL_OBSERVACAO = 40

Public Const STRING_NFISCAL_MOTIVOCANCEL = 50

'Numero maximo de regioes de vendas existentes
Public Const NUM_MAX_REGIAOVENDA = 9999

'para tabela ItemNF
Public Const STRING_ITEMNF_DESCRICAO = 250

Public Const STRING_TIPO_DE_VENDEDOR_DESCRICAO = 50

Public Const STRING_NOSSO_NUMERO = 20
Public Const STRING_OBS_PARC_REC = 255

Public Const STRING_CODIGO_BARRAS_PARC_CPR = 50

Public Const NUMERO_PROXIMO_TIPO_VENDEDOR As String = "NUM_PROX_TIPO_VENDEDOR"

'Número máximo de Clientes, Vendedores
Public Const NUM_MAX_CLIENTES = 99999999
Public Const NUM_MAX_VENDEDORES = 9999

'para tabela de Natureza Operação
Public Const STRING_NATUREZAOP_DESCRICAO = 150
Public Const STRING_NATUREZAOP_CODIGO = 4
Public Const STRING_NATUREZAOP_DESCRNF = 30

'para tabela de Padrão Cobrança
Public Const STRING_PADRAO_COBRANCA_DESCRICAO = 20

'Tipo do Título para Comissões
Public Const TIPO_NF = 0
Public Const TIPO_PARCELA = 1
Public Const TIPO_DEBITO = 2
Public Const TIPO_TITULO_RECEBER = 3
Public Const TIPO_COMISSAO_AVULSA = 4
Public Const TIPO_COMISSAO_LOJA = 5
Public Const TIPO_PV = 6

'Tipos de Desconto (Títulos a Receber)
Public Const VALOR_FIXO = 1
Public Const Percentual = 2
Public Const VALOR_ANT_DIA = 3
Public Const VALOR_ANT_DIA_UTIL = 4
Public Const PERC_ANT_DIA = 5
Public Const PERC_ANT_DIA_UTIL = 6

'Status de títulos, notas fiscais, parcelas e outros

'alterado por tulio 11/09/02, inclusao para o modulo de loja
Public Const STATUS_ATIVO = 0

Public Const STATUS_PENDENTE = 0
Public Const STATUS_LANCADO = 1
Public Const STATUS_BAIXADO = 2
Public Const STATUS_SUSPENSO = 3
Public Const STATUS_ABERTO = 4
Public Const STATUS_EXCLUIDO = 5
Public Const STATUS_LIBERADO = 6
Public Const STATUS_CANCELADO = 7
Public Const STATUS_GRAVADA_IMPORTACAO_DATA_NULA = 8
Public Const STATUS_PREVISAO = 9

Public Const STRING_STATUS_LANCADO = "LANÇADO"
Public Const STRING_STATUS_BAIXADO = "BAIXADO"
Public Const STRING_STATUS_CANCELADO = "CANCELADO"
Public Const STRING_STATUS_PENDENTE = "PENDENTE"
Public Const STRING_STATUS_ABERTO = "ABERTO"
Public Const STRING_STATUS_DENEGADA = "DENEGADA"
Public Const STRING_STATUS_CANCNAOHOM = "CANC NAO HOMOLOG"

Public Const STRING_TITULO_OBSERVACAO = 255

'#########################################
'Inserido por Wagner
Public Const STRING_NFPAG_OBSERVACAO = 255
'#########################################

Public Const STRING_VENDEDOR_NOME_REDUZIDO = 20

Public Const STRING_TIPO_INSTR_COBRANCA_DESCRICAO = 50

Public Const STRING_SIGLA_DOCUMENTO = 4
Public Const STRING_TIPO_DOC_DESCRICAO = 50
Public Const STRING_TIPO_DOC_SIGLA = 4
Public Const STRING_TIPO_DOC_INFO_SIGLA = 10
Public Const STRING_TIPO_DOC_INFO_NOMEREDUZIDO = 40
'Public Const STRING_TIPO_DOC_INFO_NOMETELANFISCAL = 20
Public Const STRING_TIPO_DOC_INFO_TIPODOCCPR = 4
Public Const STRING_TIPO_DOC_INFO_TITULOTELANFISCAL = 100
Public Const STRING_TIPO_DOC_INFO_SIGLANFORIGINAL = 10
'Public Const STRING_TIPO_DOC_INFO_NATOPEXTPADRAO = 3

Public Const NUM_PROX_COMISSAO = "NUM_PROX_COMISSAO"

'para tabela TiposDocInfo
Public Const STRING_TIPODOCINFO_SIGLA = 10
Public Const STRING_TIPODOCINFO_DESCRICAO = 60
Public Const STRING_TIPODOCINFO_NOMEREDUZIDO = 40
Public Const TIPODOCINFO_NAO_FATURAVEL = 0
Public Const TIPODOCINFO_FATURAVEL = 1
Public Const TIPODOCINFO_SEM_COMISSAO = 0
Public Const TIPODOCINFO_COM_COMISSAO = 1
Public Const TIPODOCINFO_NAO_COMPRAS = 0
Public Const TIPODOCINFO_COMPRAS = 1 'indica que o documento participa da estatistica de compras
Public Const TIPODOCINFO_COMPRAS_DEVOLUCAO = 2 'indica que o documento participa da estatistica de compras como devolução

'Public Const STRING_TIPODOCINFO_NOME_TELA = 20
Public Const STRING_TIPODOCINFO_TITULO_TELA = 100

'Campo Complementar
Public Const DOCINFO_NORMAL = 0
Public Const DOCINFO_COMPLEMENTO = 1

'Campo Faturamento
Public Const TIPODOCINFO_FATURAMENTO_NAO = 0
Public Const TIPODOCINFO_FATURAMENTO_SIM = 1
Public Const TIPODOCINFO_FATURAMENTO_DEV = 2

'Campo Rastreavel
Public Const TIPODOCINFO_RASTREAVEL_NAO = 0
Public Const TIPODOCINFO_RASTREAVEL_SIM = 1


'para a tabela de estados
Public Const STRING_ESTADO_SIGLA = 2
Public Const STRING_ESTADO_NOME = 30

'Comentada por Leo em 13/03/02. Passou a ser implementada em AdmLib.ClassConstCust
'Public Const STRING_CLIENTE_RAZAO_SOCIAL = 40

Public Const STRING_CLIENTE_CGC = 14
'Public Const STRING_CLIENTE_NOME_REDUZIDO = 20
'Public Const STRING_CLIENTE_OBSERVACAO = 100
Public Const STRING_CLIENTE_GUIA = 10

Public Const STRING_FILIAL_CLIENTE_NOME = 50

Public Const STRING_TIPO_CLIENTE_DESCRICAO = 50
Public Const STRING_TIPO_CLIENTE_OBS = 100

Public Const STRING_TIPOSDEDESCONTO_DESCRICAO = 50

Public Const STRING_FAIXA_NUM_COB = 20 'Faixa de Numeração
Public Const STRING_NOMENOBANCO = 50 'Nome da carteira fornecido pelo banco
Public Const STRING_CODCARTNOBANCO = 1 'Código da carteira no banco
Public Const STRING_COBRADORES_NOME = 50
Public Const STRING_COBRADOR_NOME_REDUZIDO = 20 'Tamanho do nome reduzido do Cobrador
Public Const STRING_DESCRICAO_CARTCOBR = 50
Public Const STRING_CODIGO_CARTCOBR = 4

Public Const TIPO_COBR_VALIDO_PARA_BORDERO = 1
Public Const NAO_LIQUID_TITULO_OUTRO_BANCO = 0
Public Const LIQUID_TITULO_AMBOS_BANCO = 2
Public Const LIQUID_TITULO_OUTRO_BANCO = 1
Public Const NAO_PERMITE_DEP_OUTRO_BANCO = 0
Public Const PERMITE_DEP_OUTRO_BANCO = 1

'Tabela de Preço default para versão Light
Public Const TABELA_PRECO_DEFAULT = 1

Public Const TABELA_PRECO_TIPO_VENDA = 0
Public Const TABELA_PRECO_TIPO_COMPRA = 1
Public Const TABELA_PRECO_TIPO_ANALISE = 2

Public Const CANCELAMENTO_FILIALFORNFILEMP = 1
Public Const CADASTRAMENTO_FILIALFORNFILEMP = 0

'Por Leo em 20/02/02 ***
'Constantes utilizadas na função NFPag_Grava_BD em ClassCPRGrava
'Servem p/ identificar a origem do documento a ser gravado.
Public Const NFPAG_CONTASAPAGAR = 0 'Oriundo de contas a PAGAR
Public Const NFPAG_NFINSERCAO = 1 'Oriundo de Nota Fiscal Insersão
Public Const NFPAG_NFALTERACAO = 2 'Oriundo de Nota Fiscal Alteração
'Leo até aqui ***

'Retiradas de GlobalFAT.bas e transferidas para GlobalCRFAT.bas em 27/10/03
'****** CAMPOSGENERICOS / CAMPOSGENERICOSVALORES **********
Public Const STRING_CAMPOSGENERICOS_DESCRICAO As String = 50
Public Const STRING_CAMPOSGENERICOS_COMENTARIOS As String = 255
Public Const STRING_CAMPOSGENERICOS_VALIDAEXCLUSAO As String = 50
Public Const STRING_CAMPOSGENERICOSVALORES_VALOR As String = 50
Public Const STRING_CAMPOSGENERICOSVALORES_COMPLEMENTO As String = 50

Public Const CAMPOSGENERICOS_VOLUMEESPECIE = 1
Public Const CAMPOSGENERICOS_VOLUMEMARCA = 2
Public Const CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES = 3
Public Const CAMPOSGENERICOS_DESCRICAORATEIO = 4
Public Const CAMPOSGENERICOS_REQUISITANTE = 5
Public Const CAMPOSGENERICOS_SITUACAO = 6
Public Const CAMPOSGENERICOS_HISTORICO = 7
Public Const CAMPOSGENERICOS_SUBCONTA = 8
Public Const CAMPOSGENERICOS_TIPOPARADA = 9
Public Const CAMPOSGENERICOS_STATUSOV = 10
Public Const CAMPOSGENERICOS_MOTIVOSOV = 11
Public Const CAMPOSGENERICOS_TIPOEMBALAGEM = 12
Public Const CAMPOSGENERICOS_STATUSRELACCLI = 13
Public Const CAMPOSGENERICOS_TIPO_REL_ETIQUETA = 14
Public Const CAMPOSGENERICOS_TRVORIGEM_OCR = 15
Public Const CAMPOSGENERICOS_TRVTIPODET_OCR = 16
Public Const CAMPOSGENERICOS_TRVTIPOAPORTE_OCR = 17
Public Const CAMPOSGENERICOS_TRPORIGEM_OCR = 15
Public Const CAMPOSGENERICOS_TRPTIPODET_OCR = 16
Public Const CAMPOSGENERICOS_TRPTIPOAPORTE_OCR = 17
Public Const CAMPOSGENERICOS_AF_EMPRESAS = 18
Public Const CAMPOSGENERICOS_AF_TIPOAPOS = 19
Public Const CAMPOSGENERICOS_AF_TIPOASSOC = 20
Public Const CAMPOSGENERICOS_OS_STATUSITEM = 21
Public Const CAMPOSGENERICOS_SOLICSRV_STATUSITEM = 22
Public Const CAMPOSGENERICOS_CARGO_VENDEDOR = 23
Public Const CAMPOSGENERICOS_DESTINO_VIAGEM = 24
Public Const CAMPOSGENERICOS_CHAVE_ROTA = 25
Public Const CAMPOSGENERICOS_MEIOS_TRANSP = 26
Public Const CAMPOSGENERICOS_TIPOS_VEICULOS = 27
Public Const CAMPOSGENERICOS_RELACCLI_MOTIVO = 28
Public Const CAMPOSGENERICOS_TIPOOS = 29
Public Const CAMPOSGENERICOS_TRVEMICARGO = 30
Public Const CAMPOSGENERICOS_RELACCLI_SATIS = 31
Public Const CAMPOSGENERICOS_PROD_FABR = 32

Public Const CAMPOSGENERICOS_TRVOCRCASO_ANALISE = 35
Public Const CAMPOSGENERICOS_TRVOCRCASO_STATUS = 36
Public Const CAMPOSGENERICOS_TRVOCRCASO_AUTOPOR = 37

Public Const CAMPOSGENERICOS_PRJ_OBJETIVOS = 38
Public Const CAMPOSGENERICOS_NATCTA_GRUPO = 39

Public Const CAMPOSGENERICOS_KIT_FATOR = 40

Public Const CAMPOSGENERICOS_TIPOSS = 41
Public Const CAMPOSGENERICOS_FASESS = 42
Public Const CAMPOSGENERICOS_PRJ_SEGMENTO = 43

Public Const CAMPOSGENERICOS_PRODUTO_COR = 101
Public Const CAMPOSGENERICOS_PRODUTO_DETALHE_COR = 102
Public Const CAMPOSGENERICOS_TIPO_DESCONTO = 103

'Incluído por Luiz Nogueira em 27/10/03
'Usadas para as tabelas de contatos
Public Const STRING_CONTATOGERAL_CONTATO = 50
Public Const STRING_CONTATOGERAL_SETOR = 50
Public Const STRING_CONTATOGERAL_CARGO = 50
Public Const NUM_MAX_CONTATOS = 200
Public Const STRING_CONTATOGERAL_OUTMEIOCOMUNIC = 250
'***************************************

'Incluído por Luiz Nogueira em 27/10/03
'Usadas para a tabela de RelacionamentoClientes
Public Const RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE = 1
Public Const RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA = 2
Public Const RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE_TEXTO As String = "Receptivo"
Public Const RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA_TEXTO As String = "Ativo"

'Incluído por Luiz Nogueira em 27/10/03
Type typeRelacionamentoClientes
    lCodigo As Long
    iFilialEmpresa As Integer
    iOrigem As Integer
    dtData As Date
    dHora As Double
    lTipo As Long
    lCliente As Long
    iFilialCliente As Integer
    iContato As Integer
    iAtendente As Long
    lRelacionamentoAnt As Long
    sAssunto1 As String
    sAssunto2 As String
    iStatus As Integer
    lNumIntParcRec As Long
    dtDataProxCobr As Date
    dtDataPrevReceb As Date
    iStatusCG As Integer
    dtDataFim As Date
    dHoraFim As Double
    lMotivo As Long
    iTipoDoc As Integer
    lNumIntDocOrigem As Long
    lSatisfacao As Long
End Type

'Incluído por Luiz Nogueira em 27/10/03
Type typeRelacionamentoClientesCons
    vlCodigo As Variant
    viFilialEmpresa As Variant
    viOrigem As Variant
    vdtData As Variant
    vdHora As Variant
    vlTipo As Variant
    vlCliente As Variant
    viFilialCliente As Variant
    viContato As Variant
    viAtendente As Variant
    vlRelacionamentoAnt As Variant
    vsAssunto1 As Variant
    vsAssunto2 As Variant
    viStatus As Variant
    vlCodigoDe As Variant
    vlCodigoAte As Variant
    vdtDataDe As Variant
    vdtDataAte As Variant
    viAtendenteDe As Variant
    viAtendenteAte As Variant
End Type

'Incluído por Luiz Nogueira em 27/10/03
Type typeClienteContatos
    lCliente As Long
    iFilialCliente As Long
    iCodigo As Integer
    sContato As String
    sSetor As String
    sCargo As String
    sTelefone As String
    sFax As String
    sEmail As String
    iPadrao As Integer
    dtDataNasc As Date
    sOutrosMeioComunic As String
End Type

Type typeMvPerCli
    iFilialEmpresa As Integer
    iExercicio As Integer
    lCliente As Long
    iFilial As Integer
    iCodMoeda As Integer
    dSldIni As Double
    dDeb01 As Double
    dCred01 As Double
    dDeb02 As Double
    dCred02 As Double
    dDeb03 As Double
    dCred03 As Double
    dDeb04 As Double
    dCred04 As Double
    dDeb05 As Double
    dCred05 As Double
    dDeb06 As Double
    dCred06 As Double
    dDeb07 As Double
    dCred07 As Double
    dDeb08 As Double
    dCred08 As Double
    dDeb09 As Double
    dCred09 As Double
    dDeb10 As Double
    dCred10 As Double
    dDeb11 As Double
    dCred11 As Double
    dDeb12 As Double
    dCred12 As Double
End Type

Type typeTituloReceber
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    lCliente As Long
    iFilial As Integer
    sSiglaDocumento As String
    lNumTitulo As Long
    dtDataEmissao As Date
    iStatus As Integer
    dSaldo As Double
    iNumParcelas As Integer
    dValor As Double
    dValorIRRF As Double
    dValorISS As Double
    dISSRetido As Double
    dValorINSS As Double
    iINSSRetido As Integer
    dPercJurosDiario As Double
    dPercMulta As Double
    sObservacao As String
    iCondicaoPagto As Integer
    dtDataRegistro As Date
    iEspecie As Integer
    dPISRetido As Double
    dCOFINSRetido As Double
    dCSLLRetido As Double
    iMoeda As Integer
    sCcl As String
    lCpoGenerico1 As Long
    iReajustePeriodicidade As Integer
    dtReajusteBase As Date
    dtReajustadoAte As Date
    sNatureza As String
End Type

Public Const STRING_TRANSPORTADORA_GUIA = 10
Public Const STRING_TRANSPORTADORA_OBS = 250

Type typeTransportadora
    iCodigo As Integer
    sNome As String
    sNomeReduzido As String
    lEndereco As Long
    sCgc As String
    sInscricaoEstadual As String
    sInscricaoMunicipal As String
    iViaTransporte As Integer
    sGuia As String
    dPesoMinimo As Double
    sObservacao As String
    iIEIsento As Integer
    iIENaoContrib As Integer
End Type

Type typeMvDiaCli
    iFilialEmpresa As Integer
    lCliente As Long
    iFilial As Integer
    dtData As Date
    iCodMoeda As Integer
    dSldIni As Double
    dDeb As Double
    dCred As Double
End Type

Type typePadraoCobranca
    iCodigo As Integer
    sDescricao As String
    iInstrucao1 As Integer
    iInstrucao2 As Integer
    dJuros As Double
    iDiasDeProtesto1 As Integer
    iDiasDeProtesto2 As Integer
    iInativo As Integer
End Type

Type typeCarteiraCobrador
    iCobrador As Integer
    iCodCarteiraCobranca As Integer
    sContaContabil As String
    iDesativada As Integer
    iDiasDeRetencao As Integer
    dTaxaCobranca As Double
    dTaxaDesconto As Double
    sContaDuplDescontadas As String
    lQuantidadeAtual As Long
    lQuantidadeAtualBanco As Long
    dSaldoAtual As Double
    dSaldoAtualBanco As Double
    sFaixaNossoNumeroInicial As String
    sFaixaNossoNumeroFinal As String
    sFaixaNossoNumeroProx As String
    sNomeNoBanco As String
    sCodCarteiraNoBanco As String
    iNumCarteiraNoBanco As Integer
    iImprimeBoleta As Integer
    iComRegistro As Integer
    iGeraNossoNumero As Integer
    iFormPreImp As Integer
End Type

Type typeTipoInstrCobr
    iCodigo As Integer
    sDescricao As String
    iRequerDias As Integer
End Type

Type typeTipoDocInfo
    sSigla As String
    sDescricao As String
    iTipoMovtoEstoque As Integer
    iTipoMovtoEstoque2 As Integer
    iTipoMovtoEstoqueConsig As Integer
    iTipoMovtoEstoqueConsig2 As Integer
    iTipoMovtoEstoqueBenef As Integer
    iTipoMovtoEstoqueBenef2 As Integer
    sNaturezaOperacaoPadrao As String
    iInfoContabilizacao As Integer
    sTipoDocCPR As String
    iCodigo As Integer
    sNomeReduzido As String
    sNomeTelaNFiscal As String
    sTituloTelaNFiscal As String
    iFaturavel As Integer
    iComissao As Integer
    iEmitente As Integer
    iDestinatario As Integer
    iComplementar As Integer
    iTipo As Integer
    iOrigem As Integer
    iPadrao As Integer
    iFaturamento As Integer
    iTipoOperacaoTrib As Integer
    sNatOpExtPadrao As String
    sSiglaNFOriginal As String
    iModeloArqICMS As Integer
    iNFFatura As Integer
    iSubTipoContabil As Integer
    iRastreavel As Integer
    iCompras As Integer
    iEscaninhoRastro As Integer
    
    'nfe 3.10
    iModDocFis As Integer
    iModDocFisE As Integer
    iFinalidadeNFe As Integer
    iIndConsumidorFinal As Integer
    iIndPresenca As Integer
    
End Type

Type typeComissaoNF
    lNumIntDoc As Long
    iCodVendedor As Integer
    dValorBase As Double
    dPercentual As Double
    dValor As Double
    dPercentualEmissao As Double
    dValorEmissao As Double
    iIndireta As Integer
    iSeq As Integer
End Type

Type typeInfoComissao
    lNumIntCom As Long
    iTipoTitulo As Integer
    lNumIntDoc As Long
    iCodVendedor As Integer
    dtDataGeracao As Date
    dtDataBaixa As Date
    dPercentual As Double
    dValorBase As Double
    dValor As Double
    iStatus As Integer
    iFilialEmpresa As Integer
    sVendedorNomeRed As String
End Type

Type typeFilialCliente
    lCodCliente As Long
    iCodFilial As Integer
    sNome As String
    sCgc As String
    sInscricaoEstadual As String
    sInscricaoMunicipal As String
    lEndereco As Long
    lEnderecoEntrega As Long
    lEnderecoCobranca As Long
    iCodTransportadora As Integer
    sObservacao As String
    sContaContabil As String
    iVendedor As Integer
    dComissaoVendas As Double
    iRegiao As Integer
    iFreqVisitas As Integer
    dtDataUltVisita As Date
    iCodCobrador As Integer
    iICMSBaseCalculoComIPI As Integer
    lRevendedor As Long
    iTipoFrete As Integer
    sInscricaoSuframa As String
    sRG As String
    iCodFilialLoja As Integer
    iFilialEmpresaLoja As Integer
    lCodClienteLoja As Long
    iAtivo As Integer
    iTransferido As Integer
    lNumIntDocLog As Long
    iQuantLog As Integer
    iCodTranspRedesp As Integer
    iDetPagFrete As Integer
    sGuia As String
    iCodMensagem As Integer 'Inserido por Wagner
    lCodExterno As Long
    iRegimeTributario As Integer
    iIEIsento As Integer
    iIENaoContrib As Integer
    
    'nfe 3.10
    sIdEstrangeiro As String

End Type

Type typeCliente
    iAtivo As Integer
    iTransferido As Integer
    lCodigo As Long
    lCodigoLoja As Long
    sRazaoSocial As String
    sNomeReduzido As String
    iTipo As Integer
    sObservacao As String
    dLimiteCredito As Double
    iCondicaoPagto As Integer
    dDesconto As Double
    iCodPadraoCobranca As Integer
    iCodMensagem As Integer
    iTabelaPreco As Integer
    lNumeroCompras As Long
    dMediaCompra As Double
    dtDataPrimeiraCompra As Date
    dtDataUltimaCompra As Date
    lMediaAtraso As Long
    lMaiorAtraso As Long
    dSaldoTitulos As Double
    dSaldoPedidosLiberados As Double
    dSaldoAtrasados As Double
    dValPagtosAtraso As Double
    dValorAcumuladoCompras As Double
    lNumTitulosProtestados As Long
    dtDataUltimoProtesto As Date
    iNumChequesDevolvidos As Integer
    dtDataUltChequeDevolvido As Date
    dSaldoDuplicatas As Double
    lNumPagamentos As Long
    iProxCodFilial As Integer
    sNome As String
    sCgc As String
    sRG As String
    sInscricaoEstadual As String
    sInscricaoMunicipal As String
    sInscricaoSuframa As String
    lEndereco As Long
    lEnderecoEntrega As Long
    lEnderecoCobranca As Long
    iCodTransportadora As Integer
    sObservacao2 As String
    sContaContabil As String
    iVendedor As Integer
    dComissaoVendas As Double
    iRegiao As Integer
    iFreqVisitas As Integer
    dtDataUltVisita As Date
    iCodCobrador As Integer
    iTipoFrete As Integer
    iFilialEmpresaLoja As Integer
    iFilialEmpresaFilialLoja As Integer
    iCodFilialLoja As Integer
    iCodFilial As Integer
    iCodTranspRedesp As Integer
    iDetPagFrete As Integer
    sGuia As String
    iBloqueado As Integer 'Inserido por Wagner
    sUsuarioCobrador As String
    sUsuRespCallCenter As String
    dPercFatMaiorPV As Double
    
    iIgnoraRecebPadrao As Integer
    iTemFaixaReceb As Integer
    dPercentMaisReceb As Double
    dPercentMenosReceb As Double
    iRecebForaFaixa As Integer
    iRegimeTributario As Integer
    iIEIsento As Integer
    iIENaoContrib As Integer
    
End Type

Type typeEstado
    sSigla As String
    sNome As String
    dICMSAliquotaInterna As Double
    dICMSAliquotaExportacao As Double
    dICMSAliquotaImportacao As Double
    iCodIBGE As Integer
    dICMSPercFCP As Double
    
    dICMSAliquotaInternaAnt As Double
    dICMSAliquotaImportacaoAnt As Double
    dICMSPercFCPAnt As Double
    
    dtDataIniAliqInternaAtual As Date
    dtDataIniAliqImportacaoAtual As Date
    dtDataIniAliqFCPAtual As Date
End Type

Type typeChequePre
    lNumIntCheque As Long
    lNumIntChequeBord As Long
    lNumIntDoc As Long
    lNumIntTituloPag As Long
    lCliente As Long
    iFilial As Integer
    iBanco As Integer
    sAgencia As String
    sContaCorrente As String
    lNumero As Long
    dtDataDeposito As Date
    dValor As Double
    lNumBordero As Long
    sCPFCGC As String
    lNumMovtoCaixa As Long
    lNumMovtoSangria As Long
    iAprovado As Integer
    iNaoEspecificado As Integer
    lNumBorderoLoja As Long
    lNumBorderoLojaBanco As Long
    iFilialEmpresa As Integer
    lSequencialLoja As Long
    lSequencialBack As Long
    iFilialEmpresaLoja As Integer
    iStatus As Integer
    iTipoBordero As Integer
    iCaixa As Integer
    lSequencialCaixa As Long
    lNumIntExt As Long
    lCupomFiscal As Long
    sCarne As String
    iECF As Integer
    lCOO As Long
    iLocalizacao As Integer
    dtDataEmissao As Date
End Type

Type typeTipoDocumento
    sSigla As String
    sDescricao As String
    sDescricaoReduzida As String
    iContabiliza As Integer
    iAumentaValorPagto As Integer
    iEmNFFatPag As Integer
    iEmCreditoPagForn As Integer
    iEmTituloRec As Integer
    iEmDebitosRecCli As Integer
    iClasseDocCPR As Integer
End Type

Public Type typeNFiscal
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    sSerie As String
    lNumNotaFiscal As Long
    lCliente As Long
    iFilialCli As Integer
    iFilialEntrega As Integer
    lFornecedor As Long
    iFilialForn As Integer
    dtDataEmissao As Date
    dtDataSaida As Date
    iFilialPedido As Integer
    lNumPedidoVenda As Long
    sNumPedidoTerc As String
    dValorProdutos As Double
    dValorFrete As Double
    dValorSeguro As Double
    dValorOutrasDespesas As Double
    dValorDesconto As Double
    iCodTransportadora As Integer
    sMensagemNota As String
    iTabelaPreco As Integer
    iTipoNFiscal As Integer
    sNaturezaOp As String
    dPesoLiq As Double
    dPesoBruto As Double
    dtDataVencimento As Date
    lNumIntTrib As Long
    sPlaca As String
    sPlacaUF As String
    lVolumeQuant As Long
    lVolumeEspecie As Long
    lVolumeMarca As Long
    iCanal As Integer
    lNumIntNotaOriginal As Long
    dtDataEntrada As Date
    dValorTotal As Double
    iClasseDocCPR As Integer
    lNumIntDocCPR As Long
    iStatus As Integer
    sVolumeNumero As String
    iFreteRespons As Integer
    dtDataReferencia As Date
    lNumRecebimento As Long
    sObservacao As String
    sCodUsuarioCancel As String
    sMotivoCancel As String
    lClienteBenef As Long
    iFilialCliBenef As Integer
    lFornecedorBenef As Long
    iFilialFornBenef As Integer
    dHoraEntrada As Double
    dHoraSaida As Double
    iCodTranspRedesp As Integer
    iDetPagFrete As Integer
    iSemDataSaida As Integer
    iMoeda As Integer
    dTaxaMoeda As Double
    dVolumeTotal As Double
    sMensagemCorpoNota As String
    iNaoImpCobranca As Integer
    iRecibo As Integer
    lNumNFe As Long
    sCodVerificacaoNFe As String
    iNFe As Integer
    sUFSubstTrib As String
    dValorDescontoTit As Double
    dValorDescontoItens As Double
    dValorItens As Double
    lFornEntTerc As Long
    iFilialFornEntTerc As Integer
    sChvNFe As String
    lCliIntermediario As Long
    iFilialCliIntermediario As Integer
    sSerieNFPOrig As String
    lNumNFPOrig As Long
    sUFRemet As String
    sUFDest As String
End Type

Type typeInfoItemNF
    lNumIntNF As Long
    lNumNotaFiscal As Long
    iItem As Integer
    iFilialEmpresa As Integer
    sProduto As String
    sUnidadeMed As String
    dQuantidade As Double
    dPrecoUnitario As Double
    dPercDesc As Double
    dValorDesconto As Double
    dtDataEntrega As Date
    sDescricaoItem As String
    dValorAbatComissao As Double
    lNumIntPedVenda As Long
    lNumIntItemPedVenda As Long
    lNumIntDoc As Long
    lNumIntTrib As Long
    iAlmoxarifado As Integer
    sAlmoxarifadoNomeRed As String
    iStatus As Integer
    lNumIntDocOrig As Long
 End Type

Type typeItemNF
    lNumIntNF As Long
    lNumNotaFiscal As Long
    iItem As Integer
    sProduto As String
    sUnidadeMed As String
    dQuantidade As Double
    dPrecoUnitario As Double
    dPercDesc As Double
    dValorDesconto As Double
    dtDataEntrega As Date
    sDescricaoItem As String
    dValorAbatComissao As Double
    lNumIntPedVenda As Long
    lNumIntItemPedVenda As Long
    lNumIntDoc As Long
    lNumIntTrib As Long
    iAlmoxarifado As Integer
    sAlmoxarifadoNomeRed As String
    iStatus As Integer
    lNumIntDocOrig As Long
    iControleEstoque As Integer
    sUMEstoque As String
    sCcl As String
    sSerieNF As String
    dPrecoUnitarioMoeda As Double
    iSeqPack As Integer
    dPercentMaisReceb As Double
    dPercentMenosReceb As Double
    iRecebForaFaixa As Integer
    sSerieNFOrig As String
    lNumNFOrig As Long
    iItemNFOrig As Integer
    dComissao As Double
    iTabelaPreco As Integer
End Type

Type typeTipoCliente
    iCodigo As Integer
    sDescricao As String
    dLimiteCredito As Double
    iCondicaoPagto As Integer
    dDesconto As Double
    iCodMensagem As Integer
    iTabelaPreco As Integer
    sObservacao As String
    sContaContabil As String
    iVendedor As Integer
    dComissaoVendas As Double
    iRegiao As Integer
    iFreqVisitas As Integer
    iCodTransportadora As Integer
    iCodCobrador As Integer
    iPadraoCobranca As Integer
End Type

Type typeTributacaoNF
    lNumIntDoc As Long
    sNaturezaOpInterna As String
    iTipoTributacao As Integer
    dIPIBase As Double
    dIPIValor As Double
    dIPICredito As Double
    dICMSBase As Double
    dICMSValor As Double
    dICMSSubstBase As Double
    dICMSSubstValor As Double
    dICMSCredito As Double
    iISSIncluso As Integer
    dISSBase As Double
    dISSAliquota As Double
    dISSValor As Double
    dIRRFBase As Double
    dIRRFAliquota As Double
    dIRRFValor As Double
    dValorINSS As Double
    iINSSRetido As Integer
    dINSSBase As Double
    dINSSDeducoes As Double
    dPISCredito As Double
    dCOFINSCredito As Double
    dPISRetido As Double
    dISSRetido As Double
    dCOFINSRetido As Double
    dCSLLRetido As Double
End Type

Type typeTipoVendedor
    iCodigo As Integer
    sDescricao As String
    dPercComissao As Double
    dPercComissaoBaixa As Double
    dPercComissaoEmissao As Double
    iComissaoSobreTotal As Integer
    iComissaoFrete As Integer
    iComissaoSeguro As Integer
    iComissaoICM As Integer
    iComissaoIPI As Integer
End Type

Type typeTribComplNF
    lNumIntDoc As Long
    iTipo As Integer
    sNaturezaOp As String
    iTipoTributacao As Integer
    iIPITipo As Integer
    'sIPICodProduto As String
    dIPIBaseCalculo As Double
    dIPIPercRedBase As Double
    dIPIAliquota As Double
    dIPIValor As Double
    dIPICredito As Double
    iICMSTipo As Integer
    dICMSBase As Double
    dICMSPercRedBase As Double
    dICMSAliquota As Double
    dICMSValor As Double
    dICMSCredito As Double
    dICMSSubstBase As Double
    dICMSSubstAliquota As Double
    dICMSSubstValor As Double
    dPISCredito As Double
    dCOFINSCredito As Double
End Type

Type typeTribItemNF
    lNumIntNF As Long
    iItem As Integer
    lNumIntDoc As Long
    sNaturezaOp As String
    iTipoTributacao As Integer
    iIPITipo As Integer
    sIPICodProduto As String
    dIPIBaseCalculo As Double
    dIPIPercRedBase As Double
    dIPIAliquota As Double
    dIPIValor As Double
    dIPICredito As Double
    iICMSTipo As Integer
    dICMSBase As Double
    dICMSPercRedBase As Double
    dICMSAliquota As Double
    dICMSValor As Double
    dICMSCredito As Double
    dICMSSubstBase As Double
    dICMSSubstAliquota As Double
    dICMSSubstValor As Double
    dPISCredito As Double
    dCOFINSCredito As Double
    dICMSAliquotaAdicaoDI As Double
    dICMSPercRedBaseAdicaoDI As Double
    dDespImpICMSBase As Double 'valor da base do icms deste item obtido pelo rateio de despesas de importacao
    dDespImpICMSValor As Double 'valor da base do icms deste item obtido pelo rateio de despesas de importacao
    dDespImpICMSCredito As Double 'valor do credito de icms deste item obtido pelo rateio de despesas de importacao
    dICMSSubstPercRedBase As Double
    dICMSSubstPercMVA As Double
    iPISTipo As Integer
    iPISTipoCalculo As Integer
    dPISBase As Double
    dPISAliquota As Double
    dPISAliquotaValor As Double
    dPISQtde As Double
    dPISValor As Double
    iPISSTTipoCalculo As Integer
    dPISSTBase As Double
    dPISSTAliquota As Double
    dPISSTAliquotaValor As Double
    dPISSTQtde As Double
    dPISSTValor As Double
    iCOFINSTipo As Integer
    iCOFINSTipoCalculo As Integer
    dCOFINSBase As Double
    dCOFINSAliquota As Double
    dCOFINSAliquotaValor As Double
    dCOFINSQtde As Double
    dCOFINSValor As Double
    iCOFINSSTTipoCalculo As Integer
    dCOFINSSTBase As Double
    dCOFINSSTAliquota As Double
    dCOFINSSTAliquotaValor As Double
    dCOFINSSTQtde As Double
    dCOFINSSTValor As Double
    sCST As String
    sISSQN As String
    dISSBase As Double
    dISSAliquota As Double
    sISSCidadeIBGE As String
    dISSValor As Double
    dIIBase As Double
    dIIDespAduaneira As Double
    dIIIOF As Double
    dIIValor As Double
    iICMSBaseModalidade As Integer
    iICMSSubstBaseModalidade As Integer
    sIPIEnquadramentoClasse As String
    sIPIEnquadramentoCodigo As String
    sIPICNPJProdutor As String
    sIPISeloCodigo As String
    lIPISeloQtde As Long
    iIPITipoCalculo As Integer
    dIPIUnidadePadraoQtde As Double
    dIPIUnidadePadraoValor As Double
    dValorFreteItem As Double
    dValorSeguroItem As Double
    dValorOutrasDespesasItem As Double
    dValorDescontoItem As Double
    iOrigemMercadoria As Integer
    iExTIPI As Integer
    sGenero As String
    iProdutoEspecifico As Integer
    sEAN As String
    sEANTrib As String
    dQtdTrib As Double
    sUMTrib As String
    dValorUnitTrib As Double
End Type

Type typeParcelaPagar
    dSaldo As Double
    dtDataVencimento As Date
    dtDataVencimentoReal As Date
    dValor As Double
    iBancoCobrador As Integer
    iNumParcela As Integer
    iPortador As Integer
    iProxSeqBaixa As Integer
    iStatus As Integer
    iTipoCobranca As Integer
    lNumIntDoc As Long
    lNumIntTitulo As Long
    sNossoNumero As String
    sCodigoDeBarras As String
    dValorOriginal As Double
    iMotivoDiferenca As Integer
    sCodUsuarioLib As String
    dtDataLib As Date
End Type

Type typeParcelaReceber
    lNumIntDoc As Long
    lNumIntTitulo As Long
    iNumParcela As Integer
    iStatus As Integer
    dtDataVencimento As Date
    dtDataVencimentoReal As Date
    dSaldo As Double
    dValor As Double
    iCobrador As Integer
    iCarteiraCobranca As Integer
    sNumTitCobrador As String
'    lNumIntCheque As Long
    iProxSeqBaixa As Integer
    iProxSeqOcorr As Integer
    iDesconto1Codigo As Integer
    dtDesconto1Ate As Date
    dDesconto1Valor As Double
    iDesconto2Codigo As Integer
    dtDesconto2Ate As Date
    dDesconto2Valor As Double
    iDesconto3Codigo As Integer
    dtDesconto3Ate As Date
    dDesconto3Valor As Double
    iAceite As Integer
    iDescontada As Integer
    '#############################
    'INSERIDO POR WAGNER
    iPrevisao As Integer
    sObservacao As String
    '#############################
    dValorOriginal As Double
    iTipoPagto As Integer
    iCodConta As Integer
    dtDataCredito As Date
    dtDataEmissaoCheque As Date
    iBancoCheque As Integer
    sAgenciaCheque As String
    sContaCorrenteCheque As String
    lNumeroCheque As Long
    dtDataDepositoCheque As Date
    iAdmMeioPagto As Integer
    iParcelamento As Integer
    sNumeroCartao As String
    dtValidadeCartao As Date
    sAprovacaoCartao As String
    dtDataTransacaoCartao As Date
    lIdImpressaoBoleto As Long
End Type

Type typeTipoOcorrRemParcRec
    iFilialEmpresa As Integer
    lNumIntParc As Long
    iNumSeqOcorr As Integer
    iCobrador As Integer
    iCodOcorrencia As Integer
    dtDataRegistro As Date
    dtData As Date
    iTituloVoltaCarteira As Integer
    dtNovaDataVcto As Date
    dJuros As Double
    iInstrucao1 As Integer
    iDiasDeProtesto1 As Integer
    iInstrucao2 As Integer
    iDiasDeProtesto2 As Integer
    lNumBordero As Long
    dValorCobrado As Double
    sNumTitCobrador As String
    lNumIntDoc As Long
End Type

Type TypeTransfCartCobr

    lNumIntParc As Long
    iNumSeqOcorr As Integer
    iCobrador As Integer
    iCarteiraCobranca As Integer
    dtData As Date
    dtDataRegistro As Date

End Type

Type typeDebitosRecCli
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    lCliente As Long
    iFilial As Integer
    sSiglaDocumento As String
    lNumTitulo As Long
    iStatus As Integer
    dtDataEmissao As Date
    dValorTotal As Double
    dSaldo As Double
    dValorSeguro As Double
    dValorFrete As Double
    dOutrasDespesas As Double
    dValorProdutos As Double
    dValorICMS As Double
    dValorICMSSubst As Double
    dValorIPI As Double
    dValorIRRF As Double
    sObservacao As String
    dPISRetido As Double
    dCOFINSRetido As Double
    dCSLLRetido As Double
End Type

Type typeInfoParcRec
    iMarcada As Integer
    sNomeRedCliente As String
    sRazaoSocialCli As String
    iFilialCliente As Integer
    lCliente As Long
    lNumTitulo As Long
    iNumParcela As Integer
    lNumIntParc As Long
    dValor As Double
    'Janaina
    dSaldo As Double
    dValorRecebto As Double
    'Janaina
    dValorOriginal As Double
    dValorJuros As Double
    dValorMulta As Double
    dValorDesconto As Double
    dtVencimento As Date
    iPadraoCobranca As Integer
    sSiglaDocumento As String
    iFilialEmpresa As Integer
    sCobradorNomeRed As String
    sCartCobrDesc As String
    iCobrador As Integer
    iCarteiraCobrador As Integer
    dtDataVencimentoReal As Date
End Type

Type typeCRFATConfig
    sCodigo As String
    iFilialEmpresa As Integer
    sDescricao As String
    iTipo As Integer
    sConteudo As String
End Type

Type typeFilFornFilEmp
    iFilialEmpresa As Integer
    lCodFornecedor As Long
    iCodFilial As Integer
    lNumeroCompras As Long
    dMediaCompra As Double
    dtDataPrimeiraCompra As Date
    dtDataUltimaCompra As Date
    dValorAcumuladoCompras As Double
    dtDataUltDevolucao As Date
    lNumTotalDevolucoes As Long
    dValorAcumuladoDevolucoes As Double
    lMaiorAtraso As Long
    lPedidosEmAberto As Long
    lAtrasoAcumulado As Long
    lPedidosRecebidos As Long
    lItensPedidosRecebidos As Long
End Type

Type typeCobrador
    iCodigo As Integer
    iFilialEmpresa As Integer
    iInativo As Integer
    sNomeReduzido As String
    sNome As String
    lEndereco As Long
    iCodBanco As Integer
    iCobrancaEletronica As Integer
    iCodCCI As Integer
    lCNABProxSeqArqCobr As Long
    lFornecedor As Long
    iFilial As Integer
End Type

Type typeSerie
    iFilialEmpresa As Integer
    sSerie As String
    lProxNumNFiscal As Long
    lProxNumNFiscalEntrada As Long
    lProxNumNFiscalImpressa As Long
    iLockImpressao As Integer
    iImprimindo As Integer
    iTipoFormulario As Integer
    lProxNumRomaneio As Long
    sNomeTsk As String
    iEletronica As Integer
    iModDocFis As Integer
End Type

Type typeConsultas
    sNomeTela As String
    sSigla As String
    sConsulta As String
    iPosicao As Integer
    iNivel As Integer
    sTelaRelacionada As String
    iIconeModulo As Integer
    iIconeConsulta As Integer
    sDescricao As String
    
End Type

'Tipos de Destino para uma Requisição
Public Const TIPO_DESTINO_AUSENTE = -1
Public Const TIPO_DESTINO_EMPRESA = 0
Public Const TIPO_DESTINO_FORNECEDOR = 1

'???? Vai desaparecer
Public Const TIPO_DESTINO_CLIENTE = 1

Type typeImportCli
    lCodCliente As Long
    sRazaoSocial As String
    sNomeReduzido As String
    iTipo As Integer
    sObservacao As String
    dLimiteCredito As Double
    iCondicaoPagto As Integer
    dDesconto As Double
    iCodPadraoCobranca As Integer
    iCodMensagem As Integer
    iTabelaPreco As Integer
    lNumPagamentos As Long
    iProxCodFilial As Integer
    iCodFilial As Integer
    sFilialNome As String
    sFilialCGC As String
    sFilialInscEstadual As String
    sFilialInscMunicipal As String
    iFilialCodTransportadora As Integer
    sFilialObservacao1 As String
    sFilialContaContabil As String
    iFilialVendedor As Integer
    dFilialComissaoVendas As Double
    iFilialRegiao As Integer
    iFilialFreqVisitas As Integer
    dtFilialDataUltVisita As Date
    iFilialCodCobrador As Integer
    iFilialICMSBaseCalculoIPI As Integer
    lFilialRevendedor As Long
    sFilialTipoFrete As String
    sEndereco As String
    sBairro As String
    sCidade As String
    sSiglaEstado As String
    iCodigoPais As Integer
    sCEP As String
    sTelefone1 As String
    sTelefone2 As String
    sEmail As String
    sFax As String
    sContato As String
    sEndereco1 As String
    sBairro1 As String
    sCidade1 As String
    sSiglaEstado1 As String
    iCodigoPais1 As Integer
    sCEP1 As String
    sTelefone11 As String
    sTelefone21 As String
    sEmail1 As String
    sFax1 As String
    sContato1 As String
    sEndereco2 As String
    sBairro2 As String
    sCidade2 As String
    sSiglaEstado2 As String
    iCodigoPais2 As Integer
    sCEP2 As String
    sTelefone12 As String
    sTelefone22 As String
    sEmail2 As String
    sFax2 As String
    sContato2 As String
End Type


Type typeConhecimentoFrete
    lNumIntNFiscal As Long
    dFretePeso As Double
    dFreteValor As Double
    dSEC As Double
    dDespacho As Double
    dPedagio As Double
    dOutrosValores As Double
    dAliquotas As Double
    dBaseCalculo As Double
    dValorTotal As Double
    dValorICMS As Double
    dPesoMercadoria As Double
    dValorMercadoria As Double
    sNotasFiscais As String
    sObservacao As String
    sColeta As String
    sEntrega As String
    sCalculadoAte As String
    sNaturezaCarga As String
    sLocalVeiculo As String
    sRemetente As String
    sEnderecoRemetente As String
    sMunicipioRemetente As String
    sUFRemetente As String
    sCepRemetente As String
    sCGCRemetente As String
    sInscEstadualRemetente As String
    sDestinatario As String
    sEnderecoDestinatario As String
    sMunicipioDestinatario As String
    sUFDestinatario As String
    sCepDestinatario As String
    sCGCDestinatario As String
    sInscEstadualDestinatario As String
    sMarcaVeiculo As String
    iICMSIncluso As Integer
    dValorINSS As Double
    iINSSRetido As Integer
    iIncluiPedagio As Integer
    iImprimeMsgICMS As Integer
End Type

Type typeSldDiaForn
    iFilialEmpresa As Integer
    lFornecedor As Long
    iFilialForn As Integer
    sProduto As String
    dtData As Date
    dQuantCompra As Double
    dValorCompra As Double
End Type

Type typeSldMesForn
    iFilialEmpresa As Integer
    lFornecedor As Long
    iFilialForn As Integer
    iAno As Integer
    sProduto As String
    adQuantCompras(1 To 12) As Double
    adValorCompras(1 To 12) As Double
End Type

'Usado para atualizacao da ClienteHistorico
Type typeClienteHistorico
    lNumIntDoc As Long
    dtDataAtualizacao As Date
    lCodigo As Long
    sCgc As String
    sInscricaoEstadual As String
    sInscricaoMunicipal As String
    sInscricaoSuframa As String
    sRazaoSocial As String
    sEndereco As String
    sBairro As String
    sCidade As String
    sUF As String
    sPais As String
    sCEP As String
    iFilialCliente As Integer
    iAtivo As Integer
    lCodigoLoja As Long
    iFilialEmpresaLoja As Integer
    sNomeReduzido As String
    iTipo As Integer
    dLimiteCredito As Double
    sObservacao As String
    iCondicaoPagto As Integer
    iTabelaPreco As Integer
    sUsuarioCobrador As String
    sUsuRespCallCenter As String
    iCodTransportadora As Integer
    iVendedor As Integer
    sRG As String
    iRegimeTributario As Integer
    sTelefone1 As String
    sTelefone2 As String
    sEmail As String
    sEmail2 As String
    sMensagemNF As String
    dDesconto As Double
    dComissaoVendas As Double
    iRegiao As Integer
    sUsuario As String
    dtDataReg As Date
    dHoraReg As Double
    iCodFilialLoja As Integer
End Type

'Usado para atualizacao de NaturezaOPHistorico
Type typeNaturezaOPHistorico
    lNumIntDoc As Long
    dtDataAtualizacao As Date
    sCodigo As String
    sDescricao As String
End Type

'Usado para atualizacao da TranspHistorico
Type typeTranspHistorico
    lNumIntDoc As Long
    dtDataAtualizacao As Date
    iCodTransp As Integer
    sCgc As String
    sInscricaoEstadual As String
    sInscricaoMunicipal As String
    sRazaoSocial As String
    sEndereco As String
    sBairro As String
    sCidade As String
    sUF As String
    sPais As String
    sCEP As String
End Type

Type typeLog

    lNumIntDoc As Long
    iOperacao As Integer
    sLog1 As String
    sLog2 As String
    sLog3 As String
    sLog4 As String
    dtData As Date
    dHora As Double
    iContinuacao As Integer
    
End Type


'Type da Tela VendedorFilial

Type typeVendedorFilial
    
    iVendedor As Integer
    iFilialEmpresa As Integer

End Type

Type TypeDVVLin
    iLinha As Integer
    sDescricao As String
End Type

Type TypeDVVLinCol
    iLinha As Integer
    iColuna As Integer
    sFormula As String
End Type

Type TypePlanMargContrCol
    iColuna As Integer
    sTitulo As String
    sDescricao As String
End Type

Type TypePlanMargContrLin
    iLinha As Integer
    sDescricao As String
    sFormulaGeral As String
    sFormulaL1 As String
    iFormato As Integer
    iEditavel As Integer
End Type

Type TypePlanMargContrLinCol
    iLinha As Integer
    sFormula As String
    iColuna As Integer
End Type

Type typeClassificacaoFiscal
    sCodigo As String
    sDescricao As String
    dIIAliquota As Double
    dIPIAliquota As Double
    dPISAliquota As Double
    dCOFINSAliquota As Double
    dICMSAliquota As Double
End Type

Type typeContrato
    iFilialEmpresa As Integer
    lNumIntDoc As Long
    sCodigo As String
    sDescricao As String
    iAtivo As Integer
    lCliente As Long
    iFilCli As Integer
    sObservacao As String
    dtDataIniContrato As Date
    dtDataFimContrato As Date
    dtDataRenovContrato As Date
    sContaContabil As String
    sHistoricoContabil As String
    dtDataIniCobrancaPadrao As Date
    iPeriodicidadePadrao As Integer
    iCondPagtoPadrao As Integer
    sCcl As String
    sNaturezaOp As String
    iTipoTributacao As Integer
    iTipo As Integer
    lFornecedor As Long
    iFilialFornecedor As Integer
    iRecibo As Integer
    iNFe As Integer
    sSerie As String
End Type

Type typeItensDeContrato

    lNumIntDoc As Long
    lNumIntContrato As Long
    iSeq As Integer
    iCobrar As Integer
    sProduto As String
    sDescProd As String
    dQuantidade As Double
    sUM As String
    dValor As Double
    iMedicao As Integer
    dtDataIniCobranca As Date
    dtDataProxCobranca As Date
    iPeriodicidade As Integer
    iCondPagto As Integer
    dQtdeFaturada As Double
    dVlrFaturado As Double
    sObservacao As String
    sCcl As String
    dtDataRefIni As Date
    dtDataRefFim As Date
    dSaldo As Double
    dtDataVenctoReal As Date
    iParcela As Integer
    objCondPagto As New ClassCondicaoPagto
    dtDataVencto As Date
    iParcelaExt As Integer
    iQtdeParcelas As Integer
    iUltParcCobrada As Integer

End Type

Type typeItensDeMedicaoContrato
    
    lMedicao As Long
    lNumIntItensContrato As Long
    dQuantidade As Double
    dCusto As Double
    dVlrCobrar As Double
    iStatus As Integer
    dtDataRefIni As Date
    dtDataRefFim As Date
    dtDataCobranca As Date

End Type

Type typeMedicaoContrato

    lCodigo As Long
    dtData As Date
    lNumIntContrato As Long

End Type

'########################################
'Inserido por Wagner 06/07/2006
Type typeFilialContato
    lCodContato As Long
    iCodFilial As Integer
    sNome As String
    sCgc As String
    sInscricaoEstadual As String
    sInscricaoMunicipal As String
    lEndereco As Long
    lEnderecoEntrega As Long
    lEnderecoCobranca As Long
    iCodTransportadora As Integer
    sObservacao As String
    sContaContabil As String
    iVendedor As Integer
    dComissaoVendas As Double
    iRegiao As Integer
    iFreqVisitas As Integer
    dtDataUltVisita As Date
    iCodCobrador As Integer
    iICMSBaseCalculoComIPI As Integer
    lRevendedor As Long
    iTipoFrete As Integer
    sInscricaoSuframa As String
    sRG As String
    iCodFilialLoja As Integer
    iFilialEmpresaLoja As Integer
    lCodContatoLoja As Long
    iAtivo As Integer
    iTransferido As Integer
    lNumIntDocLog As Long
    iQuantLog As Integer
    iCodTranspRedesp As Integer
    iDetPagFrete As Integer
    sGuia As String
    iCodMensagem As Integer
End Type

Type typeContato
    iAtivo As Integer
    iTransferido As Integer
    lCodigo As Long
    lCodigoLoja As Long
    sRazaoSocial As String
    sNomeReduzido As String
    iTipo As Integer
    sObservacao As String
    dLimiteCredito As Double
    iCondicaoPagto As Integer
    dDesconto As Double
    iCodPadraoCobranca As Integer
    iCodMensagem As Integer
    iTabelaPreco As Integer
    lNumeroCompras As Long
    dMediaCompra As Double
    dtDataPrimeiraCompra As Date
    dtDataUltimaCompra As Date
    lMediaAtraso As Long
    lMaiorAtraso As Long
    dSaldoTitulos As Double
    dSaldoPedidosLiberados As Double
    dSaldoAtrasados As Double
    dValPagtosAtraso As Double
    dValorAcumuladoCompras As Double
    lNumTitulosProtestados As Long
    dtDataUltimoProtesto As Date
    iNumChequesDevolvidos As Integer
    dtDataUltChequeDevolvido As Date
    dSaldoDuplicatas As Double
    lNumPagamentos As Long
    iProxCodFilial As Integer
    sNome As String
    sCgc As String
    sRG As String
    sInscricaoEstadual As String
    sInscricaoMunicipal As String
    sInscricaoSuframa As String
    lEndereco As Long
    lEnderecoEntrega As Long
    lEnderecoCobranca As Long
    iCodTransportadora As Integer
    sObservacao2 As String
    sContaContabil As String
    iVendedor As Integer
    dComissaoVendas As Double
    iRegiao As Integer
    iFreqVisitas As Integer
    dtDataUltVisita As Date
    iCodCobrador As Integer
    iTipoFrete As Integer
    iFilialEmpresaLoja As Integer
    iFilialEmpresaFilialLoja As Integer
    iCodFilialLoja As Integer
    iCodFilial As Integer
    iCodTranspRedesp As Integer
    iDetPagFrete As Integer
    sGuia As String
    iBloqueado As Integer '
End Type
'########################################

Type typeItemNFEntrega
    lNumIntDoc As Long
    lNumIntNF As Long
    lNumIntItemNF As Long
    dtDataEntrega As Date
    dQuantidade As Double
    sPedidoCliente As String
End Type

Type typeItensDeContratoCob

    lNumIntItensContrato As Long
    lNumIntItemNFiscal As Long
    dtDataUltCobranca As Date
    dtDataRefIni As Date
    dtDataRefFim As Date
    lNumIntDocCobranca As Long

End Type

Type typeItensMedicaoCob

    lNumIntItensContrato As Long
    lMedicao As Long
    lNumIntItemNFiscal As Long
    dtDataUltCobranca As Date
    dtDataRefIni As Date
    dtDataRefFim As Date
    lNumIntDocCobranca As Long

End Type

Public Const STRING_DI_NUMERO = 50
Public Const STRING_DI_DESCRICAO = 250
Public Const STRING_DI_PROCESSO_TRADING = 50
Public Const STRING_DI_LOCAL_DESEMBARACO = 250
Public Const STRING_DI_AD_COD_FABR = 50
Public Const STRING_DI_COD_EXPORTADOR = 60

Type typeDIInfo
    lNumIntDoc As Long
    sNumero As String
    dtData As Date
    iFilialEmpresa As Integer
    iStatus As Integer
    sDescricao As String
    lFornTrading As Long
    iFilialFornTrading As Integer
    sProcessoTrading As String
    'iMoeda As Integer
    'dTaxaMoeda As Double
    dPesoBrutoKG As Double
    dPesoLiqKG As Double
    dValorMercadoriaMoeda As Double
    dValorFreteInternacMoeda As Double
    dValorSeguroInternacMoeda As Double
    dValorMercadoriaEmReal As Double
    dValorFreteInternacEmReal As Double
    dValorSeguroInternacEmReal As Double

    dIIValor As Double
    dIPIValor As Double
    dPISValor As Double
    dCOFINSValor As Double
    dICMSValor As Double

    dValorDespesas As Double
    
    dtDataDesembaraco As Date
    sUFDesembaraco As String
    sLocalDesembaraco As String
    iMoedaMercadoria As Integer
    iMoedaFrete As Integer
    iMoedaSeguro As Integer
    iMoedaItens As Integer
    iMoeda1 As Integer
    dTaxaMoeda1 As Double
    iMoeda2 As Integer
    dTaxaMoeda2 As Double
    sCodExportador As String
    
    'nfe 3.10
    iViaTransp As Integer
    iIntermedio As Integer
    sCNPJAdquir As String
    sUFAdquir As String
    
End Type

Public Const STRING_DRAWBACK_NUMERO = 11

Type typeAdicaoDI
    lNumIntDoc As Long
    lNumIntDI As Long
    iSeq As Integer 'sequencial da adicao dentro da DI
    sIPICodigo As String
    dValorAduaneiro As Double
    dIIAliquota As Double
    dIPIAliquota As Double
    dPISAliquota As Double
    dCOFINSAliquota As Double
    dICMSAliquota As Double
    dIIValor As Double
    dIPIValor As Double
    dPISValor As Double
    dCOFINSValor As Double
    dICMSValor As Double
    dIPIBase As Double
    dPISBase As Double
    dCOFINSBase As Double
    dICMSBase As Double
    dICMSPercRedBase As Double
    dDespesaAduaneira As Double
    dTaxaSiscomex As Double
    sCodFabricante As String
    sDescricao As String
    sNumDrawback As String
End Type

Type typeItemAdicaoDI
    lNumIntDoc As Long
    lNumIntAdicaoDI As Long
    iAdicao As Integer
    iSeq As Integer 'sequencial do item dentro da adicao
    sProduto As String
    sDescricao As String
    sUM As String
    dQuantidade As Double
    dValorUnitFOBNaMoeda As Double
    dValorUnitFOBEmReal As Double
    dValorUnitCIFNaMoeda As Double
    dValorUnitCIFEmReal As Double
    dValorTotalFOBNaMoeda As Double
    dValorTotalFOBEmReal As Double
    dValorTotalCIFNaMoeda As Double
    dValorTotalCIFEmReal As Double
    iTotalCIFEmRealManual As Integer
    dPesoBruto As Double
    dPesoLiq As Double
    dValorUnitTrib As Double
    dIPIUnidadePadraoValor As Double
End Type

Public Const IMPORTCOMPL_ORIGEM_DI = 0
Public Const IMPORTCOMPL_ORIGEM_NF = 1

Public Const IMPORTCOMPL_TIPO_II = 1
Public Const IMPORTCOMPL_TIPO_PIS = 2
Public Const IMPORTCOMPL_TIPO_COFINS = 3
Public Const IMPORTCOMPL_TIPO_ACERTOS_FISCAIS = 8
Public Const IMPORTCOMPL_TIPO_CUSTO_FINANCEIRO = 9
Public Const IMPORTCOMPL_TIPO_SEGURO_PRODUTO = 10
Public Const IMPORTCOMPL_TIPO_OVERHEAD = 11
Public Const IMPORTCOMPL_TIPO_COMISSAO_FORNEC = 12
Public Const IMPORTCOMPL_TIPO_COMISSAO_FORNEC_MOBIMAX = 13
Public Const IMPORTCOMPL_TIPO_TAXA_DE_LI = 22

Public Const IMPORTCOMPL_TIPORATEIO_PESO = 0
Public Const IMPORTCOMPL_TIPORATEIO_VALOR = 1

Public Const STRING_IMPORTCOMPL_DESCRICAO = 250

Type typeImportCompl
    lNumIntDoc As Long
    iTipoDocOrigem As Integer
    lNumIntDocOrigem As Long
    iSeq As Integer 'para posicionar no grid
    iTipo As Integer
    sDescricao As String
    dValor As Double
    dPerc As Double
    iDias As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTipoImportCompl
    iCodigo As Integer
    sDescReduzida As String
    sDescricao As String
    iIncluiValorProdutos As Integer
    iIncluiBaseICMS As Integer
    iImpressaoNaNF As Integer
    iSeqImpressaoNF As Integer
    iLinhaPadraoNaTela As Integer
    iComAliquota As Integer
    iPodeSerDespAduaneira As Integer
    iTipoRateio As Integer
    iGerencial As Integer
    iAceitaValor As Integer
    iAceitaPerc As Integer
    iAceitaDias As Integer
    iIncluiNoValorAduaneiro As Integer
End Type

Type typeItemAdicaoDIItemNF
    lNumIntItemNF As Long
    lNumIntItemAdicaoDI As Long
    dValorAduaneiro As Double
    dValorII As Double
    dDespImpValorRateado As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeDASAliquotas
    iAno As Integer
    iMes As Integer
    dAliquotaICMS As Double
    dAliquotaICMSServ As Double
    dAliquotaTotal As Double
End Type

Type TypeSpedFisProcessoRef
    lNumIntNF As Long
    sNumProc As String
    iIndProc As Integer
End Type

Type TypeSpedFisArrecadacaoRef
    lNumIntNF As Long
    sCodigoDA As String
    sUF As String
    sNumDA As String
    sCodAut As String
    dValorDA As Double
    dtDataVcto As Date
    dtDataPagto As Date
End Type

Type TypeSpedFisDocRef
    lNumIntNF As Long
    sIndOper As String
    sIndEmit As String
    lCliente As Long
    iFilialCli As Integer
    lFornecedor As Long
    iFilialForn As Integer
    sCodMod As String
    sSerie As String
    iSubSerie As Integer
    lnumdoc As Long
    dtDataDoc As Date
End Type

Type TypeSpedFisCupomRef
    lNumIntNF As Long
    sCodMod As String
    sECFFab As String
    iECFCx As Integer
    lnumdoc As Long
    dtDataDoc As Date
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeVersaoNFe
    iCodigo As Integer
    sDescricao As String
    sColunaCFOP As String
    sProgramaEnvio As String
    iAtivo As Integer
End Type

Type typeTipoFrete
    iCodigo As Integer
    sDescricao As String
    iCodigoNFE As Integer
End Type

Type typeNFeFedScan
    lOcorrencia As Long
    dtDataEntrada As Date
    dHoraEntrada As Double
    dtDataSaida As Date
    dHoraSaida As Double
    sJustificativa As String
    iFilialEmpresa As Integer
End Type

Type typeRetiradaEntrega
    lCodigo As Long
    lEnderecoRet As Long
    lEnderecoEnt As Long
    sCNPJCPFRet As String
    sCNPJCPFEnt As String
    lClienteRet As Long
    lFornecedorRet As Long
    iFilialCliRet As Integer
    iFilialFornRet As Integer
    lClienteEnt As Long
    lFornecedorEnt As Long
    iFilialCliEnt As Integer
    iFilialFornEnt As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeInfoAdicExportacao
    iTipoDoc As Integer
    lNumIntDoc As Long
    sUFEmbarque As String
    sLocalEmbarque As String
    lNumIntDE As Long
    sNumRE As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeInfoAdicCompra
    iTipoDoc As Integer
    lNumIntDoc As Long
    sNotaEmpenho As String
    sPedido As String
    sContrato As String
End Type

Public Const STRING_INFOADIC_ITEM_NUMPEDIDOCOMPRA = 15
Public Const STRING_INFOADIC_ITEM_MSG = 500

Type typeInfoAdicDocItem
    iTipoDoc As Integer
    lNumIntDocItem As Long
    dtDataLimiteFaturamento As Date
    sNumPedidoCompra As String
    lItemPedCompra As Long
    iIncluiValorTotal As Integer
    sMsg As String
    sMsg2 As String
    lNumIntDE As Long
    sNumRE As String
End Type

Type typeInfoAdicDocItemDetExp
    iTipoDoc As Integer
    lNumIntDocItem As Long
    sNumDrawback As String
    sNumRegistExport As String
    sChvNFe As String
    dQuantExport As Double
End Type

Type typeItemPCDI
    lNumIntDI As Long
    iSeq As Integer
    iFilialEmpresa As Integer
    lCodigoPC As Long
    dtDataPC As Date
    sProdutoPC As String
    sDescProdPC As String
    sUMPC As String
    dQuantPC As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeNFSeLinks
    sCodIBGE As String
    sLinkConsulta As String
    sParamConsulta As String
    sLinkVerificacao As String
    sLinkSite As String
    sEmail As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeModelosDocFiscais
    iCodigo As Integer
    iTipo As Integer
    iEletronica As Integer
    sModelo As String
    sDescricao As String
End Type

Public Function Converte_Serie_Eletronica(ByVal sSerie As String, ByVal iEletronica As Integer) As String
    If iEletronica = vbChecked Then
        Converte_Serie_Eletronica = sSerie & "-e"
    Else
        Converte_Serie_Eletronica = sSerie
    End If
End Function

Public Function Desconverte_Serie_Eletronica(ByVal sSerie As String) As String
    If ISSerieEletronica(sSerie) Then
        Desconverte_Serie_Eletronica = Replace(sSerie, "-e", "")
    Else
        Desconverte_Serie_Eletronica = sSerie
    End If
End Function

Public Function ISSerieEletronica(ByVal sSerie As String) As Boolean
    If right(sSerie, 2) = "-e" Then
        ISSerieEletronica = True
    Else
        ISSerieEletronica = False
    End If
End Function

Public Function SerieEletronica(ByVal sSerie As String) As Integer
    If right(sSerie, 2) = "-e" Then
        SerieEletronica = MARCADO
    Else
        SerieEletronica = DESMARCADO
    End If
End Function
