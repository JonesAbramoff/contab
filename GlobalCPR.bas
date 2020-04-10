Attribute VB_Name = "GlobalCPR"
Option Explicit

Public Const FLUXO_NAO_INCLUI_PEDIDOS = 0
Public Const FLUXO_INCLUI_PEDIDOS = 1

Public Const STRING_MNEMONICOCOBREMAIL_TIPO = 20

Public Const STRING_COBRANCAEMAILPADRAO_CC = 255
Public Const STRING_COBRANCAEMAILPADRAO_ASSUNTO = 255
Public Const STRING_COBRANCAEMAILPADRAO_MENSAGEM = 500
Public Const STRING_COBRANCAEMAILPADRAO_DESCRICAO = 50
Public Const STRING_COBRANCAEMAILPADRAO_MODELO = 255
Public Const STRING_COBRANCAEMAILPADRAO_ANEXO = 255
Public Const STRING_COBRANCAEMAILPADRAO_DE = 255
Public Const STRING_COBRANCAEMAILPADRAO_NOMEEXIBICAO = 255
Public Const STRING_COBRANCAEMAILPADRAO_EMAILRESP = 255

Public Const TIPO_COBRANCAEMAILPADRAO_COBRANCA = 1
Public Const TIPO_COBRANCAEMAILPADRAO_AVISO = 2
Public Const TIPO_COBRANCAEMAILPADRAO_AGRADECIMENTO = 3
Public Const TIPO_COBRANCAEMAILPADRAO_AVISO_PAGTO_CP = 4
Public Const TIPO_COBRANCAEMAILPADRAO_COBRANCA_FATURA = 5
Public Const TIPO_COBRANCAEMAILPADRAO_CONTATO_CLIENTE = 6
Public Const TIPO_COBRANCAEMAILPADRAO_NFE = 7

Public Const VENDCARTAO_BROWSER_TIPO_TODASABERTAS = 1
Public Const VENDCARTAO_BROWSER_TIPO_DESSECARTAO = 2
Public Const VENDCARTAO_BROWSER_TIPO_DATAVALOR = 3

Public Const TIPOTELA_EMAIL_COBRANCA = 1
Public Const TIPOTELA_EMAIL_AGRADECIMENTO = 2
Public Const TIPOTELA_EMAIL_COBRANCA_FATURA = 3
Public Const TIPOTELA_EMAIL_AVISO_PAGTO_CP = 4
Public Const TIPOTELA_EMAIL_AVISO_COBRANCA = 5
Public Const TIPOTELA_EMAIL_NFE = 6

Public Const EMAIL_TIPODOC_PARCELASREC = 1
Public Const EMAIL_TIPODOC_PARCELASPAG = 2
Public Const EMAIL_TIPODOC_NFSPAG = 3
Public Const EMAIL_TIPODOC_NF = 4
Public Const EMAIL_TIPODOC_CLIENTE = 5

Public Const STRING_TIPO_COBRANCAEMAILPADRAO_COBRANCA = "CR-Cobrança de Atrasados"
Public Const STRING_TIPO_COBRANCAEMAILPADRAO_AVISO = "CR-Aviso de Cobrança"
Public Const STRING_TIPO_COBRANCAEMAILPADRAO_AGRADECIMENTO = "CR-Agradecimento Pagto"
Public Const STRING_TIPO_COBRANCAEMAILPADRAO_AVISO_PAGTO_CP = "CP-Aviso Pagto"
Public Const STRING_TIPO_COBRANCAEMAILPADRAO_COBRANCA_FATURA = "CP-Cobrança de envio de Fatura"
Public Const STRING_TIPO_COBRANCAEMAILPADRAO_CONTATO_CLIENTE = "CRM-Contato com Cliente"
Public Const STRING_TIPO_COBRANCAEMAILPADRAO_NFE = "FAT-Nota Fiscal Eletrônica Federal"

Public Const STRING_TIPOSDIFPARCREC_DESCRICAO = 150
Public Const STRING_TIPOSDETRETCOBR_DESCRICAO = 150
Public Const STRING_TIPOSMOVRETCOBR_DESCRICAO = 150

Public Const NUM_MAX_TIPOSDIFPARCREC = 9999

Public Const TIPODIF_ACAO_AUTOMATICA = 0
Public Const STRING_TIPODIF_ACAO_AUTOMATICA = "Padrão"
Public Const TIPODIF_ACAO_INFORMATIVA = 1
Public Const STRING_TIPODIF_ACAO_INFORMATIVA = "Informativa"
Public Const TIPODIF_ACAO_SOMA = 2
Public Const STRING_TIPODIF_ACAO_SOMA = "Soma"
Public Const TIPODIF_ACAO_SUBTRAI = 3
Public Const STRING_TIPODIF_ACAO_SUBTRAI = "Subtrai"

'criticas no processamento do reajuste de titulos a receber
Public Const PROCREAJTITREC_CRITICA_REAJUSTE_ATRASADO = 1
Public Const PROCREAJTITREC_CRITICA_NAO_REAJUSTA = 2

'#########################
'Inserido por Wagner
Public Const POSSUI_NATMOVCTA = 1
Public Const NAO_POSSUI_NATMOVCTA = 0
'#########################

Public Const STRING_FILIALCONTATODATA_HISTORICO = 250

'#######################################
'Inserido por Wagner
Public Const TITULOPAG_CHEQUE_STATUS_ABERTO = 0
Public Const TITULOPAG_CHEQUE_STATUS_PAGO = 1
Public Const TITULOPAG_CHEQUE_STATUS_LIQUIDADO = 2

Public Const STRING_TITULOPAG_CHEQUE_STATUS_ABERTO = "Aberto"
Public Const STRING_TITULOPAG_CHEQUE_STATUS_PAGO = "Pago"
Public Const STRING_TITULOPAG_CHEQUE_STATUS_LIQUIDADO = "Liquidado"
'#######################################

Public Const TAMANHO_CADA_INSTRUCAO_BOLETO = 50
Public Const TAMANHO_CAMPO_INSTRUCOES = 250

Public Const RETCOBR_DET_IGNORAR = 0
Public Const RETCOBR_DET_BAIXA = 6
Public Const RETCOBR_DET_CONFIRMADO = 2
Public Const RETCOBR_DET_REJEITADO = 3
Public Const RETCOBR_DET_TARIFAS = 12
Public Const RETCOBR_DET_CUSTAS = 33
Public Const RETCOBR_DET_BAIXA_POR_PROTESTO = 25
Public Const RETCOBR_DET_OUTRAS_BAIXAS = 9


Public Const RETCOBR_TIPO_SEU_NUMERO1 = 1 'geral
Public Const RETCOBR_TIPO_SEU_NUMERO2 = 2 'bb
Public Const RETCOBR_TIPO_SEU_NUMERO3 = 3 'Real
Public Const RETCOBR_TIPO_SEU_NUMERO4 = 4 'NumIntDoc do OcorrRemParcRec

Public Const RETCOBR_CRITICA_SEM_ERRO = 0
Public Const RETCOBR_CRITICA_SEM_PARC = 1
Public Const RETCOBR_CRITICA_VARIAS_PARC = 2
Public Const RETCOBR_CRITICA_ENTRADA_REJEITADA = 3
Public Const RETCOBR_CRITICA_LIQUIDACAO = 6
Public Const RETCOBR_CRITICA_TARIFAS = 12
Public Const RETCOBR_CRITICA_CUSTAS = 33
Public Const RETCOBR_CRITICA_BAIXA_PARCIAL = 100
Public Const RETCOBR_CRITICA_BAIXA_POR_PROTESTO = 25
Public Const RETCOBR_CRITICA_OUTRAS_BAIXAS = 9
Public Const RETCOBR_CRITICA_PARC_BAIXADA = 50
Public Const RETCOBR_CRITICA_PARC_VALOR_DIF = 51
Public Const RETCOBR_CRITICA_PARC_VENC_DIF = 52
Public Const RETCOBR_CRITICA_PAGA_DUPLIC = 53

'COnstantes para as telas BOrderoDescChq
Public Const STRING_CARTEIRACOBRANCA_DESCRICAO = 50
Public Const HIST_SAQUE_CHQ_LJ = "Saque de Cheques Loja"
Public Const HIST_DEP_CHQ_LJ = "Depósito de Cheques Loja"

'constantes para cci de chequepre
Public Const CONTA_CHEQUE_PRE = 1
Public Const CONTA_CHEQUE_PRE_NOMERED = "Cheques Pré"
Public Const CONTA_CHEQUE_PRE_DESCRICAO = "Conta Cheques Pré"

Public Const TIPOCANCELAMENTO_BAIXABAIXA = 1
Public Const TIPOCANCELAMENTO_MOVCCI = 2
Public Const TIPOCONCELAMENTO_DEVCHQ = 3
Public Const TIPOCANCELAMENTO_EXC_BORD_CHEQUE = 4

Public Const STRING_TIPOMOVCCI_NOMERED = 50

Public Const BORDERO_CHEQUEPRE = 1
Public Const BORDERO_DESCONTO = 2
Public Const TIPOMEIOPAGTO_CHEQUE = 2
Public Const CREDITO = 1

'Número da linha de títulos de um grid
Public Const GRID_LINHA_TITULO = 0

'Indicação do grid atual
Public Const NENHUM_GRID_SELECIONADO = 0
Public Const GRID_PARCELAS = 1
Public Const GRID_DEVOLUCOES = 2
Public Const GRID_ADIANTAMENTOS = 3

Public Const NAO_PREENCHIDA = "Não preenchida"

Public Const STRING_HISTORICO_CNAB = 20
Public Const STRING_LANCTOCNAB_HISTORICO = 25
Public Const STRING_MOTIVOSBAIXA_DESCRICAO = 50

Public Const TIPOMOVCCI_CREDITA = 1
Public Const TIPOMOVCCI_DEBITA = 0

Public Const BORDERO_PROCESSADO = 1

Public Const CARTEIRA_SEM_REGISTRO = 0
Public Const CARTEIRA_COM_REGISTRO = 1

Public Const BANCO_IMPRIME_BOLETA = 0
Public Const EMPRESA_IMPRIME_BOLETA = 1

Public Const BANCO_GERA_NOSSONUMERO = 0
Public Const EMPRESA_GERA_NOSSONUMERO = 1
Public Const NUM_NOSSONUMERO_MAX = "99999999"

'Bordero de Pagamento
Public Const BORDERO_NAO_EXCLUIDO = 0
Public Const BORDERO_EXCLUIDO = 1

'RelOpCadCli
Public Const ORD_POR_CODIGO = 0
Public Const ORD_POR_NOME = 1
Public Const ORD_POR_CGCCPF = 2

'Cheques Pag 3
Public Const ATUALIZAR_CHECADO = "1"
Public Const ATUALIZAR_NAO_CHECADO = "0"

'Cheques Pag 2
Public Const EMITIR_CHECADO = "1"
Public Const SELECIONAR_CHECADO = "1"
Public Const EMITIR_NAO_CHECADO = "0"
Public Const SELECIONAR_NAO_CHECADO = "0"

'Cheque Pre
Public Const CHQPRE_MARCADO = "1"
Public Const CHQPRE_NAO_MARCADO = "0"

'Bordero Pag 3, Cheque Pag Avulso 4 e Cheque Pag 4
Public Const ESTADO_ANDAMENTO = 1
Public Const ESTADO_PARADO = 0

'Bordero Pag 2 e Cheque Pag Avulso 2
Public Const PAGO_CHECADO = "1"
Public Const PAGO_NAO_CHECADO = "0"

'Bordero Cobranca 2
Public Const INCLUIR_CHECADO = "1"
Public Const INCLUIR_NAO_CHECADO = "0"

Public Const RECEBIMENTO_EM_DINHEIRO = 1
Public Const PAGAMENTO_EM_DINHEIRO = 1

Public Const INSTR_COBR_REQUER_DIAS = 1
Public Const INSTR_COBR_NAO_TRAZ_CARTEIRA = 0
Public Const INSTR_COBR_TRAZ_CARTEIRA = 1

'Fluxo de Caixa gerado pelo usuario/sistema
Public Const FLUXO_GERADO_PELO_USUARIO = 1
Public Const FLUXO_GERADO_PELO_SISTEMA = 0
Public Const FLUXOSINT_PROJ = 0
Public Const FLUXOSINT_REV = 1

'Fluxo de Caixa
Public Const FLUXOTIPOFORN_TIPOREG_PAGTO = 0
Public Const FLUXOTIPOFORN_TIPOREG_RECEBTO = 1

'valores p/indicar se houve ou nao o credito referente a uma parcela a receber descontada
Public Const PARC_REC_NAO_DESCONTADA = 0
Public Const PARC_REC_DESCONTADA = 1

'Tipos de Registro de FluxoAnalitico
Public Const FLUXOANALITICO_TIPOREG_PAGTO = 0
Public Const FLUXOANALITICO_TIPOREG_RECEBTO = 1
Public Const FLUXOANALITICO_TIPOREG_BORDERO = 2
Public Const FLUXOANALITICO_TIPOREG_APLICACAO = 3
Public Const FLUXOANALITICO_TIPOREG_CHEQUEPRE = 4
Public Const FLUXOANALITICO_TIPOREG_SALDOINI = 5

'tipos de cobranca
Public Const STRING_TIPO_COBRANCA_TODAS As String = "Todas"
Public Const TIPO_COBRANCA_TODAS = 0
Public Const TIPO_COBRANCA_CARTEIRA = 1
Public Const TIPO_COBRANCA_BANCARIA = 2
Public Const TIPO_COBRANCA_DEP_CONTA = 3
Public Const TIPO_COBRANCA_DOC = 4
Public Const TIPO_COBRANCA_OP = 5
Public Const TIPO_COBRANCA_CHEQUE_PRE = 6

Public Const BANCO_BRADESCO = 237

'########################################
'Inserido por Wagner
Public Const BANCO_BB = 1
Public Const BANCO_ITAU = 341
Public Const BANCO_BICBANCO = 320
Public Const BANCO_RURAL = 453
Public Const BANCO_REAL = 356
'########################################

Public Const NUM_MAXIMO_NF_VINCULADA_FATURA = 50

Public Const NUM_MAX_PARC_CHEQUE_MANUAL = 6

Public Const NUM_MAXIMO_PARCELAS_BORDERO = 200

'Número máximo de Fornecedores
Public Const NUM_MAX_FORNECEDORES = 99999999

'Número máximo de Parcelas de um Título
Public Const NUM_MAXIMO_PARCELAS = 200

'Tipo especifico de um codigo na tabela CPRConfig
Public Const NUM_PROX_CLIENTE = "NUM_PROX_CLIENTE"
Public Const NUM_PROX_FAVORECIDO = "NUM_PROX_FAVORECIDO"
Public Const NUM_PROX_HISTPADRAO_MOV = "NUM_PROX_HISTPADRAO_MOV"
Public Const NUM_PROX_PAG_ANTECIPADO = "NUM_PROX_PAG_ANTECIPADO"
Public Const NUM_PROX_REC_ANTECIPADO = "NUM_PROX_REC_ANTECIPADO"
Public Const NUM_PROX_FLUXO = "NUM_PROX_FLUXO"
Public Const NUM_PROX_BANCO = "NUM_PROX_BANCO"
Public Const NUM_PROX_CONTA = "NUM_PROX_CONTA"
Public Const NUM_PROX_TIPO_APLICACAO = "NUM_PROX_TIPO_APLICACAO"

'Motivos de Baixa p/Pagtos
Public Const MOTIVO_PAGAMENTO = 1
Public Const MOTIVO_PAGTO_ANTECIPADO = 2
Public Const MOTIVO_CREDITO_FORNECEDOR = 3
Public Const MOTIVO_CHEQUE_DE_TERCEIROS = 8

'Motivos de Baixa p/Recebtos
Public Const MOTIVO_PERDA = 4
Public Const MOTIVO_RECEBTO_ANTECIPADO = 5
Public Const MOTIVO_DEBITO_CLIENTE = 6
Public Const MOTIVO_RECEBIMENTO = 7
Public Const MOTIVO_CHEQUE_ENVIADO_PARA_TERCEIROS = 9
Public Const MOTIVO_CARTAO_DEBITO_CREDITO = 10

'Tipos Meio Pagamento
Public Const DINHEIRO = 1
Public Const Cheque = 2
Public Const BORDERO = 3

'Tipos de Movimento
Public Const MOVCCI_SAQUE = 0
Public Const MOVCCI_DEPOSITO = 1
Public Const MOVCCI_APLICACAO = 2
Public Const MOVCCI_RESGATE = 3
Public Const MOVCCI_SAIDA_TRANSFERENCIA = 4
Public Const MOVCCI_ENTRADA_TRANSFERENCIA = 5
Public Const MOVCCI_PAGTO_ANTECIPADO = 6
Public Const MOVCCI_RECEB_ANTECIPADO = 7
Public Const MOVCCI_RECEBIMENTO_TITULO = 8
Public Const MOVCCI_PAGTO_TITULO_POR_CHEQUE = 9
Public Const MOVCCI_PAGTO_TITULO_POR_BORDERO = 10
Public Const MOVCCI_CREDITO_RETORNO_COBRANCA = 11
Public Const MOVCCI_PAGTO_TITULO_POR_DINHEIRO = 12
Public Const MOVCCI_PAGTO_CHEQUE_PRE = 13
Public Const MOVCCI_CANC_PAGTO = 14
Public Const MOVCCI_CANC_RECEBTO = 17
Public Const MOVCCI_BORDERO_CHEQUE_PRE = 18
Public Const MOVCCI_BORDERO_CHEQUE_LOJA = 19
Public Const MOVCCI_CRED_RET_COBRANCA = 24
Public Const MOVCCI_DEB_RET_COBRANCA = 25
Public Const MOVCCI_RECEBTO_CANCELA = 26
Public Const MOVCCI_BAIXA_RECEBANTECIPADO = 27
Public Const MOVCCI_BAIXA_DEBITOSRECCLI = 28
Public Const MOVCCI_BAIXA_PAGANTECIPADO = 29
Public Const MOVCCI_BAIXA_CREDITOSPAGFORN = 30
Public Const MOVCCI_DEP_DESCONTO_CHEQUE = 31
Public Const MOVCCI_SAQ_DESCONTO_CHEQUE = 32
Public Const MOVCCI_DEP_DEVOLUCAO_CHEQUE = 33
Public Const MOVCCI_SAQ_DEVOLUCAO_CHEQUE = 34
Public Const MOVCCI_DEP_DIN_LOJA = 35
Public Const MOVCCI_CHEQUE_SANGRIA_LOJA = 36
Public Const MOVCCI_CHEQUEPRE_BACKOFFICE = 37
Public Const MOVCCI_BORDERO_CHEQUE_PRE_SAQ = 38
Public Const MOVCCI_ENVIO_CHEQUES_TERCEIROS = 39
Public Const MOVCCI_DEVOLUCAO_CHEQUES_TERCEIROS = 40
Public Const MOVCCI_CANC_BAIXA_PAGTO_ANTECIPADO = 41
Public Const MOVCCI_EXTRATO_CARTAO_CRED = 42

Public Const HIST_MOVCCI_ENVIO_CHEQUES_TERCEIROS = "Envio de Cheque para Terceiros"
Public Const HIST_MOVCCI_DEVOLUCAO_CHEQUES_TERCEIROS = "Devolução de Cheque em Terceiros"

'Tipos de movimento somente usados para sinalizar como devem ser tratadas estas operacoes em funcao dos totais de aplicacao consolidados
Public Const MOVCCI_APLICACAO_EXCLUSAO = 998
Public Const MOVCCI_RESGATE_EXCLUSAO = 999

'origem da comissao
Public Const COMISSAO_EMISSAO = 1 'a comissao é devida ao vendedor qdo da emissao do titulo
Public Const COMISSAO_BAIXA = 2 'a comissao é devida ao vendedor qdo do pagto do titulo (ou da parcela)
Public Const COMISSAO_AMBOS = 3

'Para objetos como CondicaoPagamento
Public Const Pagamento = 0
Public Const Recebimento = 1

'Tipos de Conciliacao
Public Const EXTRATO_CONCILIADO = 1
Public Const NAO_CONCILIADO = 0
Public Const CONCILIADO_MANUAL = 1
Public Const CONCILIADO_AUTOMATICO = 2

'Numero maximo de movimentos e extratos que são trazidos para a tela de conciliação
Public Const MAX_CONCILIACAO = 500

'Numero maximo de creditos a pagar fornecedor no fluxo de caixa
Public Const MAX_FLUXO_CREDITOSPAGFORN = 1000

'Numero maximo de elementos nos grids de fluxo de caixa
Public Const MAX_FLUXO = 10000

''Origens
'Global Const ORIGEM_CONTAS_PAGAR_RECEBER = "CPR"

''Public Const MODULO_ATIVO = 1
Public Const TIPOMEIOPAGTO_ATIVO = 0
Public Const TIPOMEIOPAGTO_INATIVO = 1
Public Const TIPOMEIOPAGTO_EXIGENUMERO = 1
Public Const FAVORECIDO_ATIVO = 0
Public Const FAVORECIDO_INATIVO = 1
Public Const MOVCONTACORRENTE_EXCLUIDO = 1
Public Const MOVCONTACORRENTE_NAO_EXCLUIDO = 0

'Public Const DEBRECCLI_EXCLUIDO = 3

'Pagamentos Antecipados
Public Const ANTECIPPAG_NAO_EXCLUIDO = 0
Public Const ANTECIPPAG_EXCLUIDO = 1

'Recebimentos Antecipados
Public Const ANTECIPREC_NAO_EXCLUIDO = 0
Public Const ANTECIPREC_EXCLUIDO = 1

'Tipo de Aplicação
Public Const STRING_CODIGO_APLICACAO = 4
Public Const STRING_CARACTER_INICIAL = 1
Public Const CONTA_APLICACAO = "ContaAplicacao"
Public Const CONTA_RECEITA = "ContaReceita"
Public Const CONTA_ANALITICA_ABREV = "A"
Public Const CONTA_SINTETICA_ABREV = "S"
Public Const NUM_PROX_APLICACAO = "NUM_PROX_APLICACAO"
Public Const APLIC_PROX_RESGATE = 1
Public Const TIPOAPLICACAO_INATIVO = 1
Public Const TIPOAPLICACAO_ATIVO = 0
Public Const EXIGE_NUMERO = 1

'Status de Aplicação
Public Const APLICACAO_EXCLUIDA = 0
Public Const APLICACAO_ATIVA = 1
Public Const APLICACAO_RESGATADA = 2

Public Const QUALQUER_PORTADOR = 1
Public Const PRIMEIRA_CONTA = 0
Public Const BANCO_NAO_PREENCHIDO = 0

Public Const TITULO_PAGAR = 1
Public Const SIGLA_OUTROS_PAGAMENTOS = "OP"
Public Const SIGLA_NOTA_CREDITO_PAGAR = "NCP"
Public Const SIGLA_NOTA_DEBITO_RECEBER = "NDR"
Public Const SIGLA_NOTA_ENTRADA_RETORNO = "NER"
Public Const SIGLA_CREDITO_CHEQUE_PRE = "CHQ"

'Resgate
Public Const RESGATE_EXCLUIDO = 0
Public Const RESGATE_ATIVO = 1

Public Const NUM_PROX_MENSAGEM As String = "NUM_PROX_MENSAGEM"
Public Const STRING_CODIGO_MENSAGEM = 4
'Caracter que indica Mensagem
Public Const CARACTER_MENSAGEM As String = "*"

'Tipo específico de um código na tabela CPRConfig
Public Const NUM_PROX_CREDITO_PAGAR = "NUM_PROX_CREDITO_PAGAR"


Public Const STRING_NFSPAG_HISTORICO = 250

'A Exlcuir
Public Const STRING_NOME = 50
Public Const STRING_NOME_REDUZIDO = 15
Public Const STRING_LAYOUT_CHEQUE = 80
Public Const STRING_LAYOUT_BOLETO = 80
Public Const STRING_DESCRICAO = 20
Public Const STRING_GERENTE = 30
'Fim Exclusão

Public Const STRING_TIPO_APLICACAO_DESCRICAO = 50
Public Const STRING_TIPO_OCOR_REM_COBR_DESCRICAO = 50

Public Const STRING_TIPO_FORNECEDOR_CONTA_DESPESA = 20
Public Const STRING_HISTPADRAO_DESCRICAO = 150

Public Const STRING_BANCO_NOME = 50
Public Const STRING_BANCO_NOME_REDUZIDO = 15
Public Const STRING_BANCO_LAYOUT_CHEQUE = 80
Public Const STRING_BANCO_LAYOUT_BOLETO = 80
Public Const STRING_REGIAO_VENDA_DESCRICAO = 20
Public Const STRING_REGIAO_VENDA_GERENTE = 30

Public Const STRING_TIPOINSTRCOBR_DESCRICAO = 50

Public Const STRING_BORDERO_CONVENIO = 20

Public Const STRING_DEBITOSRECCLI_OBSERVACAO = 255
Public Const STRING_CREDITOSPAGFORN_OBSERVACAO = 50

Public Const STRING_TABELA_PRECO_DESCRICAO = 50

Public Const STRING_VENDEDOR_NOME = 50
Public Const STRING_VENDEDOR_MATRICULA = 20
Public Const STRING_VENDEDOR_AGENCIA = 7
Public Const STRING_VENDEDOR_CONTA_CORRENTE = 14

Public Const STRING_TIPO_FORNECEDOR_DESCRICAO = 50

Public Const STRING_CODIGO_TIPOAPLICACAO = 3

'Public Const STRING_TIPO_DOC_SIGLA = 4
'Passou a ser implementada em AdmLib.ClassConstCust
'Public Const STRING_FORNECEDOR_RAZAO_SOC = 40
Public Const STRING_FORNECEDOR_OBS = 255


Public Const STRING_TIPO_FORNECEDOR_OBS = 100
Public Const STRING_FAVORECIDO = 50
Public Const STRING_CODIGO_FAVORECIDO = 4

Public Const STRING_CRED_PAG_SIGLA = 4
Public Const STRING_CRED_PAG_OBS = 50
Public Const STRING_VENDEDOR_NOME_RED = 20

'!!! Colocar nome da tabela antes. Dá MAIS legibilidade
'STRING_<NomeTabela>_NUM_REF_EXTERNA
Public Const STRING_NUMREFEXTERNA = 20

'!!! Nome da tabela deve vir antes do atributo
'Ver exemplo acima. Separar: MOVIMENTO_CONTA
'Não fazer abreviações excessivas ou coladas que
'dificultam a leitura. Quem fez substitui este aqui.
Public Const STRING_HISTORICOMOVCONTA = 150

Public Const STRING_TIPO_MEIO_PAGTO_DESCRICAO = 50

Public Const STRING_CONDICAO_PAGTO_DESCRICAO_REDUZIDA = 30
Public Const STRING_CONDICAO_PAGTO_DESCRICAO = 50

Public Const STRING_COBRADORES_NOME_RED = 20
Public Const STRING_AGENCIA_CCI = 5

Public Const STRING_CONTA_CORRENTE_DESCRICAO = 50
Public Const STRING_CONVENIO_PAGTO = 50
Public Const STRING_CONTA_DIRARQBORDPAGTO = 100

Public Const STRING_DV = 1
Public Const STRING_NUMCONTA = 12
Public Const STRING_EXTRATO_CODIGO_LANCAMENTO = 4
Public Const STRING_EXTRATO_HISTORICO = 25
Public Const STRING_EXTRATO_DOCUMENTO = 20

Public Const STRING_TITULO_SIGLADOCUMENTO = 4

Public Const STRING_PORTADOR_NOME_REDUZIDO = 15
Public Const STRING_PORTADOR_NOME = 50

Public Const STRING_CONTA_CORRENTE_NOME_REDUZIDO = 15
Public Const STRING_CONTA_CORRENTE_INTERNA_CODIGO = 5

Public Const STRING_FLUXOANALITICO_NOME_REDUZIDO = 50
Public Const STRING_FLUXOANALITICO_DESCRICAO = 50
Public Const STRING_FLUXOANALITICO_TITULO = 10
Public Const STRING_FLUXOANALITICO_HISTORICO = 250

Public Const STRING_FLUXO_NOME = 20
Public Const STRING_FLUXO_DESCRICAO = 50

Public Const STRING_FLUXOFORN_NOME_REDUZIDO = 20

Public Const STRING_NOME_REDUZIDOC = 20
'Nome reduzido do cliente

Public Const STRING_TIPOSDECOBRANCA_DESCRICAO = 30
Public Const STRING_PARCELA_PAGAR_NOSSO_NUMERO = 20

'Janaina
Public Const STRING_STRING_TIPOSMOVTOCTACORRENTE_DESCRICAO = 50
'Janaina

Public Const STRING_COD_CARTEIRA_COBR_BANCO = 1 'tamanho do codigo da carteira de cobranca para o banco
Public Const STRING_NOME_CARTEIRA_COBR_BANCO = 50 'tamanho do nome da carteira de cobranca para o banco

'Public Const STRING_FILIAL_CLIENTE_OBS = 100

Public Const STRING_MOV_HISTORICO = 50
Public Const STRING_TIPOAPLIC_DESCRICAO = 50
Public Const STRING_TIPOAPLIC_CONTACONTABAPLIC = 20
Public Const STRING_TIPOAPLIC_CONTARECFINAN = 20
Public Const STRING_TIPOAPLIC_HISTORICO = 50

'tipos de ocorrencia p/remessa de bordero de cobranca
Public Const COBRANCA_OCORR_INC_TITULO = 1
Public Const COBRANCA_OCORR_ALT_VCTO = 6

'Constantes

'Para tabela CPConfig
Public Const STRING_CPCONFIG_CODIGO = 50
Public Const STRING_CPCONFIG_DESCRICAO = 150
Public Const STRING_CPCONFIG_CONTEUDO = 255
'Para tabela CRConfig
Public Const STRING_CRCONFIG_CODIGO = 50
Public Const STRING_CRCONFIG_DESCRICAO = 150
Public Const STRING_CRCONFIG_CONTEUDO = 255
'Para tabela TESConfig
Public Const STRING_TESCONFIG_CODIGO = 50
Public Const STRING_TESCONFIG_DESCRICAO = 150
Public Const STRING_TESCONFIG_CONTEUDO = 255


'Para a tabela BancosInfo
Public Const STRING_BANCOINFO_DESCRICAO = 250
Public Const STRING_BANCOINFO_TEXTO = 50
Public Const STRING_BANCOINFO_VALOR = 50

'GERACAO ARQUIVO ICMS
Public Const MODELO_ARQ_ICMS_NFISCAL = 1
Public Const MODELO_ARQ_ICMS_NFISCAL_ENTRADA = 3
Public Const ROTINA_BACH_GERACAO_ARQ_ICMS = 1
Public Const TITULO_TELABATCH_GERACAO_ARQ_ICMS = "Geração de Arquivo ICMS"

Public Const ROTINA_BACH_ENVIO_DE_EMAIL = 2


Type typeCarteiraCobranca
    iCodigo As Integer
    sDescricao As String
    iValidaPara As Integer
End Type

Type typeBorderoDescChq
    iFilialEmpresa As Integer
    lNumBordero As Long
    iContaCorrente As Integer
    sContaCorrente As String
    dtDataEmissao As Date
    dtDataContabil As Date
    iCobrador As Integer
    sCobrador As String
    iCarteiraCobranca As Integer
    sCarteiraCobranca As String
    dtDataDeposito As Date
    dValorCredito As Double
    iQtdeChequesSel As Integer
    dValorChequesSel As Double
End Type

Type typeMotivosBaixa
    iCodigo As Integer
    sDescricao As String
    iPagamento As Integer
    iRecebimento As Integer
    iPagaComissaoVendas As Integer
    iSubTipo As Integer
End Type

Type typeFilCliFilEmp
    iFilialEmpresa As Integer
    lCodCliente As Long
    iCodFilial As Integer
    lNumeroCompras As Long
    dMediaCompra As Double
    dtDataPrimeiraCompra As Date
    dtDataUltimaCompra As Date
    dValorAcumuladoCompras As Double
End Type

Type typeCPConfig
    sCodigo As String
    iFilialEmpresa As Integer
    sDescricao As String
    iTipo As Integer
    sConteudo As String
End Type

Type typeCRConfig
    sCodigo As String
    iFilialEmpresa As Integer
    sDescricao As String
    iTipo As Integer
    sConteudo As String
End Type

Type typeTESConfig
    sCodigo As String
    iFilialEmpresa As Integer
    sDescricao As String
    iTipo As Integer
    sConteudo As String
End Type

Type typeBanco
    iCodBanco As Integer
    sNome As String
    sNomeReduzido As String
    sLayoutCheque As String
    sLayoutBoleto As String
    dtDataLog As Date
    iAtivo As Integer
    iLayoutCnabConciliacao As Integer
End Type

Type typePais
    iCodigo As Integer
    sNome As String
End Type

Type typePortador
    iCodigo As Integer
    sNome As String
    sNomeReduzido As String
    iInativo As Integer
    iBanco As Integer
End Type

Type typeRegiaoVenda
    iCodigo As Integer
    sDescricao As String
    iCodigoPais As Integer
    sGerente As String
    sUsuarioCobrador As String
    sUsuRespCallCenter As String
End Type

Type typeEndereco
    lCodigo As Long
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
    
    sReferencia As String
    sLogradouro As String
    sComplemento As String
    sTipoLogradouro As String
    sEmail2 As String
    lNumero As Long
    iTelDDD1 As Integer
    iTelDDD2 As Integer
    iFaxDDD As Integer
    sTelNumero1 As String
    sTelNumero2 As String
    sFaxNumero As String
End Type
   
Type typeVendedor
    iCodigo As Integer
    sNome As String
    sNomeReduzido As String
    lEndereco As Long
    iTipo As Integer
    sMatricula As String
    iCodRegiao As Integer
    dSaldoComissao As Double
    dPercComissao As Double
    dPercComissaoBaixa As Double
    dPercComissaoEmissao As Double
    iComissaoSobreTotal As Integer
    iComissaoFrete As Integer
    iComissaoSeguro As Integer
    iComissaoICM As Integer
    iComissaoIPI As Integer
    iBanco As Integer
    sAgencia As String
    sContaCorrente As String
    iAtivo As Integer
    dtDataLog As Date
    iVinculo As Integer
    sCgc As String
    sInscricaoEstadual As String
    sRazaoSocial As String
    iCargo As Integer
    iSuperior As Integer
    sCodUsuario As String
    sRG As String
End Type

Type typeContaCorrenteInt
    iCodigo As Integer
    iFilialEmpresa As Integer
    iChequePre As Integer
    sNomeReduzido As String
    sDescricao As String
    iCodBanco As Integer
    sAgencia As String
    sDVAgencia As String
    sNumConta As String
    sDVNumConta As String
    sDVAgConta As String
    sContato As String
    sConvenioPagto As String
    sTelefone As String
    sFax As String
    dSaldoInicial As Double
    dtDataSaldoInicial As Date
    sContaContabil As String
    lProxSeq As Long
    iNumMenorExtratoNaoConciliado As Integer
    lProxBordero As Long
    iChequeBordero As Integer
    iAtivo As Integer
    dtDataLog As Date
    dRotativo As Double
    sDirArqBordPagto As String
    sContaContabilChqPre As String
    lCNABProxSeqArqCobr As Long
End Type

Type typeFornecedor
    lCodigo As Long
    sRazaoSocial As String
    sNomeReduzido As String
    iTipo As Integer
    sObservacao As String
    iCondicaoPagto As Integer
    dDesconto As Double
    lNumeroCompras As Long
    dMediaCompra As Double
    dtDataPrimeiraCompra As Date
    dtDataUltimaCompra As Date
    dValorAcumuladoCompras As Double
    lMediaAtraso As Long
    lMaiorAtraso As Long
    dSaldoDuplicatas As Double
    dSaldoTitulos As Double
    dValorAcumuladoDevolucoes As Double
    lNumTotalDevolucoes As Long
    dtDataUltDevolucao As Date
    iProxCodFilial As Integer
    sNome As String
    lEndereco As Long
    sCgc As String
    sIdEstrangeiro As String
    sInscricaoEstadual As String
    sInscricaoMunicipal As String
    sContaContabil As String
    sContaFornConsig As String
    iBanco As Integer
    sAgencia As String
    sContaCorrente As String
    sObservacao2 As String
    sInscricaoINSS As String
    iTipoCobranca As Integer
    sContaDespesa As String
    iGeraCredICMS As Integer
    iTipoFrete As Integer
    iAtivo As Integer
    sInscricaoSuframa As String
    iRegimeTributario As Integer
    iIEIsento As Integer
    iIENaoContrib As Integer
    sNatureza As String
End Type

Type typeTipoFornecedor
    iCodigo As Integer
    sDescricao As String
    iCondicaoPagto As Integer
    dDesconto As Double
    sObservacao As String
    sContaDespesa As String
    iHistPadraoDespesa As Integer
End Type

Type typeTipoMeioPagto
    iTipo As Integer
    sDescricao As String
    iExigeNumero As Integer
    iInativo As Integer
End Type

Type typeMovContaCorrente
    iFilialEmpresa As Integer
    lNumMovto As Long
    iCodConta As Integer
    lSequencial As Long
    iTipo As Integer
    iExcluido As Integer
    iTipoMeioPagto As Integer
    lNumero As Long
    dtDataMovimento As Date
    dtDataContabil As Date
    dValor As Double
    sHistorico As String
    iPortador As Integer
    iConciliado As Integer
    iFavorecido As Integer
    sNumRefExterna As String
    lNumRefInterna As Long
    sOrigem As String
    iExercicio As Integer
    iPeriodo As Integer
    iLote As Integer
    lDoc As Long
    sObservacao As String 'Inserido por Wagner
    sNatureza As String
    sCcl As String
End Type

Type typeTiposDeAplicacao
    iCodigo As Integer
    sDescricao As String
    sContaAplicacao As String
    sContaReceita As String
    sHistorico As String
    iInativo As Integer
End Type

Type typeTabelaPreco
    iCodigo As Integer
    sDescricao As String
    dtDataLog As Date
    iAtivo As Integer
    iCargoMinimo As Integer
End Type

Type typeMensagem
    iCodigo As Integer
    sDescricao As String
End Type

Type typeAntecipPag
    lFornecedor As Long
    iFilial As Integer
    lNumMovto As Long
    dSaldoNaoApropriado As Double
    lNumIntPag As Long
    iExcluido As Integer
    iFilialPedCompra As Integer
    lNumPedCompra As Long
End Type

Type typeAntecipRec
    lNumIntRec As Long
    iExcluido As Integer
    lNumMovto As Long
    dSaldoNaoApropriado As Double
    lCliente As Long
    iFilial_Cliente As Integer
End Type

''Type typeNFExterna SUBSTITUIDA por NFsPag
''    dtDataEmissao As Date
''    dtDataVencimento As Date
''    iFilial As Integer
''    lNumIntDoc As Long
''    lNumNotaFiscal As Long
''    lNumIntTituloPagar As Long
''    dOutrasDespesas As Double
''    dValorFrete As Double
''    dValorProdutos As Double
''    dValorSeguro As Double
''    dValorTotal As Double
''    lFornecedor As Long
''    iStatus As Integer
''    dValorIRRF As Double
''    dValorICMS As Double
''    dValorICMSSubst As Double
''    dValorIPI As Double
''    iFilialEmpresa As Integer
''    iCreditoIPI As Integer
''    dSaldo As Double
''    iCreditoICMS As Integer
''End Type

Type typeTituloPagar
    dOutrasDespesas As Double
    dSaldo As Double
    dtDataEmissao As Date
    dtDataRegistro As Date
    dValorFrete As Double
    dValorICMS As Double
    dValorICMSSubst As Double
    dValorINSS As Double
    iINSSRetido As Integer
    dValorIPI As Double
    dValorIRRF As Double
    dValorProdutos As Double
    dValorSeguro As Double
    dValorTotal As Double
    iCreditoICMS As Integer
    iCreditoIPI As Integer
    iFilial As Integer
    iFilialEmpresa As Integer
    iNumParcelas As Integer
    iStatus As Integer
    lFornecedor As Long
    lNumIntDoc As Long
    lNumTitulo As Long
    sObservacao As String
    sSiglaDocumento As String
    iFilialPedCompra As Integer
    lNumPedCompra As Long
    iCondicaoPagto As Integer
    dPISRetido As Double
    dISSRetido As Double
    dCOFINSRetido As Double
    dCSLLRetido As Double
    dTaxaMoeda As Double
    sHistorico As String
    sNatureza As String
    sCcl As String
End Type

Type typeFluxo
    sFluxo As String
    sDescricao As String
    dtDataBase As Date
    dtDataFinal As Date
    dtDataDadosReais As Date
    lFluxoId As Long
    iFilialEmpresa As Integer
    lNumMovCta As Long
    lNumIntBaixaPag As Long
    lNumIntBaixaRec As Long
End Type

Type typeFluxoAnalitico
    lFluxoId As Long
    iTipoReg As Integer
    lFornecedor As Long
    iFilial As Integer
    sSiglaDocumento As String
    iTipoTitulo As Integer
    sTitulo As String
    iItem As Integer
    iNumParcela As Integer
    dtData As Date
    dValor As Double
    iTipo As Integer
    sNomeReduzido As String
    sDescricao As String
    lNumIntDoc As Long
    dtDataReferencia As Date
    iFilialEmpresa As Integer
    sHistorico As String
End Type

Type typeFluxoAplic
    lFluxoId As Long
    lCodigo As Long
    dtDataResgatePrevista As Date
    dSaldoAplicato As Double
    dValorResgatePrevisto As Double
End Type

Type typeFluxoForn
    lFluxoId As Long
    iTipoReg As Integer
    iTipoFornecedor As Integer
    lFornecedor As Long
    dtData As Date
    dTotalSistema As Double
    dTotalAjustado As Double
    dTotalReal As Double
    sNomeReduzido As String
    iUsuario As Integer
End Type

Type typeFluxoSint
    lFluxoId As Long
    dtData As Date
    dRecValorSistema As Double
    dRecValorAjustado As Double
    dRecValorReal As Double
    dPagValorSistema As Double
    dPagValorAjustado As Double
    dPagValorReal As Double
    dTesValorAjustado As Double
    dTesValorReal As Double
    dTesValorSistema As Double
    dSaldoValorSistema As Double
    dSaldoValorAjustado As Double
    dSaldoValorReal As Double
End Type

Type typeFluxoSldIni
    lFluxoId As Long
    sNomeReduzido As String
    iCodConta As Integer
    dSaldoSistema As Double
    dSaldoAjustado As Double
    dSaldoReal As Double
    iUsuario As Integer
End Type

Type typeFluxoTipoAplic
    lFluxoId As Long
    iTipoAplicacao As Integer
    sDescricao As String
    dtData As Date
    dTotalSistema As Double
    dTotalAjustado As Double
    dTotalReal As Double
    iUsuario As Integer
End Type

Type typeFluxoTipoForn
    lFluxoId As Long
    iTipoReg As Integer
    iTipoFornecedor As Integer
    dtData As Date
    dTotalSistema As Double
    dTotalAjustado As Double
    dTotalReal As Double
    sDescricao As String
End Type

Type typeCreditoPagar
    dtDataEmissao As Date
    iFilial As Integer
    lNumIntDoc As Long
    lFornecedor As Long
    iStatus As Integer
    sObservacao As String
    lNumTitulo As Long
    iFilialEmpresa As Integer
    dValorTotal As Double
    dSaldo As Double
    sSiglaDocumento As String
    dValorSeguro As Double
    dValorFrete As Double
    dOutrasDespesas As Double
    dValorProdutos As Double
    dValorICMS As Double
    dValorICMSSubst As Double
    dValorIPI As Double
    dValorIRRF As Double
    iDebitoICMS As Integer
    iDebitoIPI As Integer
    dValorBaixado As Double
    dPISRetido As Double
    dCOFINSRetido As Double
    dCSLLRetido As Double
End Type


Type typeNFsPag
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    lFornecedor As Long
    iFilial As Integer
    lNumNotaFiscal As Long
    dtDataEmissao As Date
    iStatus As Integer
    lNumIntTitPag As Long
    dtDataVencimento As Date
    dValorTotal As Double
    dValorSeguro As Double
    dValorFrete As Double
    dOutrasDespesas As Double
    dValorProdutos As Double
    dValorIRRF As Double
    dValorICMS As Double
    dValorICMSSubst As Double
    iCreditoICMS As Integer
    dValorIPI As Double
    iCreditoIPI As Integer
    dPISRetido As Double
    dISSRetido As Double
    dCOFINSRetido As Double
    dCSLLRetido As Double
    dTaxaMoeda As Double
    sHistorico As String
    iINSSRetido As Integer
    dValorINSS As Double
End Type


Type typeBaixaPagar
    lNumIntBaixa As Long
    sHistorico As String
    iMotivo As Integer
    dtData As Date
    dtDataContabil As Date
    dtDataRegistro As Date
    lNumMovConta As Long
    lNumIntDoc As Long
    iStatus As Integer
    sObservacao As String
    lNumIntCheque As Long
    iFilialEmpresaCheque As Integer
    lNumIntDocBaixaAgrupada As Long
End Type

Type typeBaixaReceber
    dtData As Date
    dtDataContabil As Date
    dtDataRegistro As Date
    iMotivo As Integer
    iStatus As Integer
    lNumIntBaixa As Long
    lNumIntDoc As Long
    lNumMovCta As Long
    sHistorico As String
    lNumIntCheque As Long
    iFilialEmpresaCheque As Integer
    lNumIntDocBaixaAgrupada As Long
End Type

Type typeBaixaParcPag
    lNumIntDoc As Long
    lNumIntBaixa As Long
    lNumIntParcela As Long
    iSequencial As Integer
    dValorMulta As Double
    dValorJuros As Double
    dValorDesconto As Double
    dValorBaixado As Double
    iStatus As Integer
End Type

Type typeBaixaParcRec
    lNumIntDoc As Long
    dValorBaixado As Double
    dValorDesconto As Double
    dValorJuros As Double
    dValorMulta As Double
    iCobrador As Integer
    iSequencial As Integer
    iStatus As Integer
    lNumIntBaixa As Long
    lNumIntParcela As Long
    dValorRecebido As Double
End Type

Type typeAplicacao
    lCodigo As Long
    iFilialEmpresa As Integer
    iStatus As Integer
    dtDataAplicacao As Date
    dValorAplicado As Double
    dtDataBaixa As Date
    iTipoAplicacao As Integer
    lNumMovto As Long
    dSaldoAplicado As Double
    dtDataResgatePrevista As Date
    dValorResgatePrevisto As Double
    dTaxaPrevista As Double
    iProxSeqResgate As Integer
End Type

Type typeResgate
    lCodigoAplicacao As Long
    iSeqResgate As Integer
    lNumMovto As Long
    dValorResgatado As Double
    dRendimentos As Double
    dValorIRRF As Double
    dDescontos As Double
    dSaldoAnterior As Double
    dValorCreditado As Double
    dtDataResgatePrevista As Date
    dValorResgatePrevisto As Double
    dTaxaPrevista As Double
    sHistorico As String
    iTipoMeioPagto As Integer
    lNumero As Long
    sNumRefExterna As String
    dtDataMovimento As Date
    iStatus As Integer
End Type

Type typeFluxoPV
    iFilialEmpresa As Integer
    lPedido As Long
    iNumParcela As Integer
    dtDataVenctoReal As Date
    dValor As Double
    lCliente As Long
    iFilCli As Integer
End Type

Type typeFluxoPC
    iFilialEmpresa As Integer
    lPedido As Long
    iNumParcela As Integer
    dtDataVenctoReal As Date
    dValor As Double
    lFornecedor As Long
    iFilial As Integer
End Type

Type typeFluxoAux
    tCreditoPagar As typeCreditoPagar
    tNFsPag As typeNFsPag
    tParcelaPagar As typeParcelaPagar
    tTituloPagar As typeTituloPagar
    iTipoForn_NFsPag As Integer
    iTipoForn_TitPag As Integer
    sNomeReduzido_NFsPag As String
    sNomeReduzido_TitPag As String
    sDescricaoTipo_NFsPag As String
    sDescricaoTipo_TitPag As String
    iTipoForn_CreditoPagar As Integer
    sNomeReduzido_CreditoPagar As String
    sDescricaoTipo_CreditoPagar As String
    tDebitosRecCli As typeDebitosRecCli
    tNFiscal As typeNFiscal
    tParcelaReceber As typeParcelaReceber
    tTituloReceber As typeTituloReceber
    iTipoForn_NFsRec As Integer
    iTipoForn_TitRec As Integer
    sNomeReduzido_NFsRec As String
    sNomeReduzido_TitRec As String
    sDescricaoTipo_NFsRec As String
    sDescricaoTipo_TitRec As String
    iTipoForn_DebitosRecCli As Integer
    sNomeReduzido_DebitosRecCli As String
    sDescricaoTipo_DebitosRecCli As String
    iDiasRetencao As Integer
    tContrato As typeContrato
    tItemContrato As typeItensDeContrato
    sNomeReduzidoCliForn As String
    sDescricaoTipoCliForn As String
    iTipoCliForn_Ctr As Integer
    tFluxoPV As typeFluxoPV
    sNomeReduzidoCliPV As String
    sDescricaoTipoCliPV As String
    iTipoCliPV As Integer
    tFluxoPC As typeFluxoPC
    sNomeReduzidoFornPC As String
    sDescricaoTipoFornPC As String
    iTipoFornPC As Integer
    objTela As Object
End Type

'Type typeFluxoAux1
'    tDebitosRecCli As typeDebitosRecCli
'    tNFiscal As typeNFiscal
'    tParcelaReceber As typeParcelaReceber
'    tTituloReceber As typeTituloReceber
'    iTipoForn_NFsRec As Integer
'    iTipoForn_TitRec As Integer
'    sNomeReduzido_NFsRec As String
'    sNomeReduzido_TitRec As String
'    sDescricaoTipo_NFsRec As String
'    sDescricaoTipo_TitRec As String
'    iTipoForn_DebitosRecCli As Integer
'    sNomeReduzido_DebitosRecCli As String
'    sDescricaoTipo_DebitosRecCli As String
'    iDiasRetencao As Integer
'End Type

Type typeTipoOcorRemCobr
    iCodigo As Integer
    sDescricao As String
    iRefContabilizacao As Integer
    iTrazParaCarteira As Integer
End Type

Type typeBorderoPag

    lNumIntBordero As Long
    iExcluido As Integer
    iCodConta As Integer
    dtDataEmissao As Date
    lNumero As Long
    sNomeArq As String
    dtDataEnvio As Date
    iTipoDeCobranca As Integer
    iTitOutroBanco As Integer
    iNumArqRemessa As Integer
    dtDataVencimento As Date

End Type

Type typeBorderoCobranca
    
    dTaxaCobranca As Double
    dTaxaDesconto As Double
    dtDataEmissao As Date
    dValor As Double
    dValorDesconto As Double
    iCobrador As Integer
    iCodNossaConta As Integer
    iDiasDeRetencao As Integer
    lNumBordero As Long
    sConvenio As String
    iStatus As Integer
    dtDataCancelamento As Date
    dtDataContabilCancelamento As Date
    iCodCarteiraCobranca As Integer
    sNomeArquivo As String

End Type

'---------------------------- INICIO ARQUIVO ICMS ---------------------------------------

Type typeTipo
    
    dValorFrete As Double
    dValorSeguro As Double
    dValorOutras As Double
    dValorTotal As Double
    dValorDescontoTotal As Double
    dValorDescontoItem As Double
    dICMSAliquota As Double
    dBaseICMS As Double
    dICMSSubstBase As Double
    dValorICMS  As Double
    dValorIPI As Double
    dQuantidade As Double
    dPrecoUnitario As Double
    dtDataEmissao As Date
    iTipoComplemento As Integer
    iItem As Integer
    iColunaNoLivroICMS As Integer
    iColunaNoLivroIPI As Integer
    iModelo As Integer
    iCodigoPais As Integer
    iTipoTribICMS As Integer
    iEmitente As Integer
    iDestinatario As Integer
    iStatus As Integer
    lNumIntDocNF As Long
    lNumeroNF As Long
    sCgc As String
    sInscricaoEstadual As String
    sUnidadeFederacao As String
    sSerie As String
    sNaturezaOp As String
    sIPICodigo As String
    sCodigoProduto As String
    sDescricaoProduto As String
    sUnidMedida As String
    iTipo As Integer
    
End Type

Type typeTipo50
    
    sCgc As String
    sInscricaoEstadual As String
    dtDataEmissao As Date
    sUnidadeFederacao As String
    iModelo As Integer
    sSerie As String
    lNumeroNF As Long
    sCFOP As String
    dValorTotal As Double
    dBaseICMS As Double
    dValorICMS  As Double
    dIsentaNTributada As Double
    dOutras As Double
    dAliquotaICMS As Double
    sSituacao As String
    
End Type

Type typeTipo51
    
    lNumIntDocNF As Long
    sCgc As String
    sInscricaoEstadual As String
    dtDataEmissao As Date
    sUnidadeFederacao As String
    sSerie As String
    lNumeroNF As Long
    sCFOP As String
    dValorTotal As Double
    dValorIPI  As Double
    dIsentaNTributada As Double
    dOutras As Double
    sSituacao As String
    
End Type

Type typeTipo54
    
    sCgc As String
    iModelo As Integer
    sSerie As String
    lNumeroNF As Long
    sCFOP As String
    iItem As Integer
    sCodigoProduto As String
    dQuantidade As Double
    dValorProduto As Double
    dValorDescontos As Double
    dBaseICMS As Double
    dBaseICMSSubst As Double
    dValorIPI As Double
    dAliquotaICMS As Double
    
End Type

Type typeTipo70
    
    sCgc As String
    sInscricaoEstadual As String
    dtDataEmissao As Date
    sUnidadeFederacao As String
    iModelo As Integer
    sSerie As String
    lNumeroNF As Long
    sCFOP As String
    dValorTotal As Double
    dBaseICMS As Double
    dValorICMS  As Double
    dIsentaNTributada As Double
    dOutras As Double
    iFreteRespons As Integer
    sSituacao As String
    
End Type

Type typeTipo75
    
    dtDataInicial As Date
    dtDataFinal As Date
    sCodigoProduto As String
    sCodigoNCM As String
    sDescricaoProduto As String
    sUnidadeMedida As String
    sSituacaoTributaria As String
    dAliquotaIPI As Double
    dAliquotaICMS As Double
    dReducaoBaseICMS As Double
    dBaseICMSSubst As Double
    
End Type
'---------------------- ARQUIVO ICMS FIM ---------------------------------------


Public Type typeLctoExtratoCNAB
    iCodConta As Integer
    iNumExtrato As Integer
    lSeqLcto As Long
    dtData As Date
    dValor As Double
    iCategoria As Integer
    sCodLctoBco As String
    sHistorico As String
    sDocumento As String
    sIncideCPMF As String
    iConciliado As Integer
End Type
    
Public Type typeBancoInfo
     iCodBanco As Integer
     iInfoCodigo As Integer
     sInfoTexto As String
     sInfoDescricao As String
     iInfoNivel As Integer
End Type


Type typeInfoParcPag
    iSeqCheque As Integer
    sNomeRedForn As String
    sRazaoSocialForn As String
    iFilialForn As Integer
    lNumTitulo As Long
    iNumParcela As Integer
    lNumIntParc As Long
    iTipoCobranca As Integer
    sNomeRedPortador As String
    dValorJuros As Double
    dValorMulta As Double
    dValorDesconto As Double
    dValor As Double
    dValorOriginal As Double
    iBancoCobrador As Integer
    dtDataVencimento As Date
    lFornecedor As Long
    iPortador As Integer
    sSiglaDocumento As String
    iFilialEmpresa As Integer
    sContaFilForn As String 'conta contabil da filial do fornecedor
    dtDataEmissao As Date
    iMotivo As Integer
    lNumMovCta As Long
    lNumIntDoc As Long
    lNumIntBaixa As Long
    iSequencial As Integer
End Type

'Usado para manuseio na tela BaixaAntecipDebCliente
'O type possui mais campos que a classe com nome correspondente, pois ele é usado
'tanto para passagem de informações como para leitura de dados
Type typeInfoBaixaAntecipDebCli
    vlNumIntDocumento As Variant
    viExcluido As Variant
    vlNumMovto As Variant
    vdSaldoNaoApropriado As Variant
    vlCliente As Variant
    viFilial_Cliente As Variant
    viFilial As Variant
    viCodConta As Variant
    viTipoMeioPagto As Variant
    vlNumero As Variant
    vdValor As Variant
    vsNomeReduzidoConta As Variant
    vdtDataEmissao As Variant
    vdtDataDeFiltro As Variant
    vdtDataAteFiltro As Variant
    viStatus As Variant
    vsSiglaDocumento As Variant
    vsSiglaDocumentoFiltro As Variant
    vdValorSeguro As Variant
    vdValorFrete As Variant
    vdValorOutrasDespesas As Variant
    vdValorProdutos As Variant
    vdValorICMS As Variant
    vdValorICMSSubst As Variant
    vdValorIPI As Variant
    vdValorIRRF As Variant
    vsObservacao As Variant
    vlNumeroFiltroDe As Variant
    vlNumeroFiltroAte As Variant
    viStatus1 As Variant
End Type

'typeInfoBaixaRecCancelar Criado por Leo em 28/11/01
Type typeInfoBaixaRecCancelar
'Este Type possui campos a mais além da ClassInfoBaixaRecCancelar
'Estes campos foram colocados pois seriam necessários em ClassCprSelect
    dSaldoDebito As Double
    dtDataBaixa As Date
    dtDataCancelamento As Date
    dtDataContabilBaixa As Date
    dtDataEmissaoDebito As Date
    dtDataRegistroBaixa As Date
    dValorBaixado As Double
    dValorBaixadoCanc As Double
    dValorDebito As Double
    dValorDesconto As Double
    dValorDescontoCanc As Double
    dValorJuros As Double
    dValorJurosCanc As Double
    dValorMovCCI As Double
    dValorMulta As Double
    dValorMultaCanc As Double
    dValorParcela As Double
    dValorTipoBaixa As Double
    dValorTotalCanc As Double
    iFilialEmpresa As Integer
    iLinhaGrid As Integer
    iTipoMovCCI As Integer
    iMotivoBaixa As Integer
    iNumParcela As Integer
    iSequencial As Integer
    iStatusBaixaParcRec As Integer
    iStatusBaixaRec As Integer
    lNumDebito As Long
    lNumIntBaixa As Long
    lNumIntDebRecCli As Long
    lNumIntParcela As Long
    lNumIntRecAntecip  As Long
    lNumIntTitulo As Long
    lNumMovCta As Long
    lNumDocumento As Long
    sContaCorrente As String
    sHistoricoBaixa As String
    sHistoricoMovCCI As String
    sSiglaDocumento As String
    iCodConta As Integer
    iTipoMeioPagto As Integer
    dValorPago As Double
    lCliente As Long
    iFilialCliente As Integer
    lTituloInicialFiltro As Long
    lTituloFinalFiltro As Long
    iCodContaFiltro As Integer
    iStatusFiltro As Integer
    iConciliadoFiltro As Integer
    dtDataBaixaInicialFiltro As Date
    dtDataBaixaFinalFiltro As Date
    dtDataVenctoInicialFiltro As Date
    dtDataVenctoFinalFiltro As Date
    lNumIntBaixaParcRec As Long
    iCobrador As Integer
    iCarteiraCobrador As Integer
End Type

Type typeInfoBaixaRecCancelarVar
'Este Type possui campos a mais além da ClassInfoBaixaRecCancelar
'Estes campos foram colocados pois seriam necessários em ClassCprSelect
    vdSaldoDebito As Variant
    vdtDataBaixa As Variant
    vdtDataCredito As Variant
    vdtDataCancelamento As Variant
    vdtDataContabilBaixa As Variant
    vdtDataEmissaoDebito As Variant
    vdtDataRegistroBaixa As Variant
    vdValorBaixado As Variant
    vdValorBaixadoCanc As Variant
    vdValorDebito As Variant
    vdValorDesconto As Variant
    vdValorDescontoCanc As Variant
    vdValorJuros As Variant
    vdValorJurosCanc As Variant
    vdValorMovCCI As Variant
    vdValorMulta As Variant
    vdValorMultaCanc As Variant
    vdValorParcela As Variant
    vdValorTipoBaixa As Variant
    vdValorTotalCanc As Variant
    viFilialEmpresa As Variant
    viLinhaGrid As Variant
    viTipoMovCCI As Variant
    viMotivoBaixa As Variant
    viNumParcela As Variant
    viSequencial As Variant
    viStatusBaixaParcRec As Variant
    viStatusBaixaRec As Variant
    vlNumDebito As Variant
    vlNumIntBaixa As Variant
    vlNumIntDebRecCli As Variant
    vlNumIntParcela As Variant
    vlNumIntRecAntecip  As Variant
    vlNumIntTitulo As Variant
    vlNumMovCta As Variant
    vlNumDocumento As Variant
    vsContaCorrente As Variant
    vsHistoricoBaixa As Variant
    vsHistoricoMovCCI As Variant
    vsSiglaDocumento As Variant
    viCodConta As Variant
    viTipoMeioPagto As Variant
    vdValorPago As Variant
    vlCliente As Variant
    viFilialCliente As Variant
    vlTituloInicialFiltro As Variant
    vlTituloFinalFiltro As Variant
    viCodContaFiltro As Variant
    viStatusFiltro As Variant
    viConciliadoFiltro As Variant
    vdtDataBaixaInicialFiltro As Variant
    vdtDataBaixaFinalFiltro As Variant
    vdtDataVenctoInicialFiltro As Variant
    vdtDataVenctoFinalFiltro As Variant
    vlNumIntBaixaParcRec As Variant
    viCobrador As Variant
    viCarteiraCobrador As Variant
End Type


'TypeBaixaRecCancelar Criado por Leo em 28/11/01
Type typeBaixaRecCancelar
    dtDataBaixaFinal As Date
    dtDataBaixaInicial As Date
    dtDataVenctoFinal As Date
    dtDataVenctoInicial As Date
    dValorBaixasCancelar As Double
    iCtaCorrenteFiltro As Integer
    iFilialCliente As Integer
    iTipoBaixas As Integer
    iUltTipoCancelamento As Integer
    lCliente As Long
    lTituloFinal As Long
    lTituloInicial As Long
End Type
    
'TypeBaixaRecCancelarVar Criado por Leo em 28/11/01
Type typeBaixaRecCancelarVar
    vdtDataBaixaFinal As Variant
    vdtDataBaixaInicial As Variant
    vdtDataVenctoFinal As Variant
    vdtDataVenctoInicial As Variant
    vdValorBaixasCancelar As Variant
    viCtaCorrenteFiltro As Variant
    viFilialCliente As Variant
    viTipoBaixas As Variant
    viUltTipoCancelamento As Variant
    vlCliente As Variant
    vlTituloFinal As Variant
    vlTituloInicial As Variant
End Type

'Usado para manuseio na tela BaixaAntecipCredFornecedor
Type typeInfoBaixaAntecipCredForn
    vlNumIntDocumento As Variant
    viExcluido As Variant
    vlNumMovto As Variant
    vdSaldoNaoApropriado As Variant
    vlFornecedor As Variant
    viFilial_Fornecedor As Variant
    viFilial As Variant
    viCodConta As Variant
    viTipoMeioPagto As Variant
    vlNumero As Variant
    vdValor As Variant
    vsNomeReduzidoConta As Variant
    vdtDataEmissao As Variant
    viStatus As Variant
    vsSiglaDocumento As Variant
    vsSiglaDocumentoFiltro As Variant
    vdValorSeguro As Variant
    vdValorFrete As Variant
    vdValorOutrasDespesas As Variant
    vdValorProdutos As Variant
    vdValorICMS As Variant
    vdValorICMSSubst As Variant
    vdValorIPI As Variant
    vdValorIRRF As Variant
    vsObservacao As Variant
    viFilialPedCompra As Variant
    vlNumPedCompra As Variant
    vdtDataDeFiltro As Variant
    vdtDataAteFiltro As Variant
    vlNumeroDeFiltro As Variant
    vlNumeroAteFiltro As Variant
    viStatus1 As Variant
End Type

'Usado para atualizacao da FornecedorHistorico
Type typeFornecedorHistorico
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
    iFilialForn As Integer
End Type

'Usado para manuseio de registros na atualização de comissões
Type typeComissoesPagVariant
    vlNumIntComissao As Variant
    viCodVendedorIni As Variant
    viCodVendedorFim As Variant
    vdtComisGeradasDe As Variant
    vdtComisGeradasAte As Variant
    viTipo As Variant
    viStatusAnterior As Variant
    viStatusNovo As Variant
    viFilialEmpresa As Variant
End Type

'Usado para devolução de cheques
Type typeDevCheque
    dtData As Date
    dtDataVencimento As Date
    lFornecedor As Long
    iFilial As Integer
    dValorCredito As Double
    lNumIntChqBord As Long
    lNumIntDoc As Long
    lNumIntBaixasParcRecCanc As Long
    lNumIntTituloPag As Long
    lNumIntCheque As Long
End Type

Type typeDetArqCNABPag
    iBancoFavorecido As Integer
    sNomeFavorecido As String
    sAgenciaFavorecido As String
    sCGCFavorecido As String
    dtDataVenctoParcela As Date
    dValorPagto As Double
    sContaFavorecido As String
    sSiglaTitulo As String
    lNumTitulo As Long
    dtDataEmissaoTitulo As Date
    sNossoNumero As String
    sCodigoDeBarras As String 'Guarda o Numero Referênte ao Código da Barras
            
    sEnderecoFavorecido  As String
    sBairroFavorecido  As String
    sCidadeFavorecido  As String
    sEstadoFavorecido  As String
    sCEPFavorecido  As String
    
    lFornecedor As Long
    iFilialForn As Integer
    iNumParcela As Integer
    iSeqBaixaParcPag As Integer
    dValorMulta As Double
    dValorJuros As Double
    dValorDesconto As Double
    dValorTitulo As Double
    
    iTipoCobranca As Integer
    iBancoCobrador As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTiposDifParcRec
    iCodigo As Integer
    sDescricao As String
    sContaContabilCR As String
    sContaContabilRecDesp As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTiposDetRetCobr
    lBanco As Long
    iCodigoMovto As Integer
    iCodigoDetalhe As Integer
    sDescricao As String
    iAcao As Integer
    iAcaoManual As Integer
    iCodTipoDiferenca As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTiposMovRetCobr
    lBanco As Long
    iCodigoMovto As Integer
    sDescricao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeParcelasRecDif
    lNumIntDoc As Long
    lNumIntParc As Long
    iSeq As Integer
    dtDataRegistro As Date
    iCodTipoDif As Integer
    dValorDiferenca As Double
    sObservacao As String
    iNumSeqOcorr As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeBaixasAgrupadas
    lNumIntDoc As Long
    dtDataBaixa As Date
    sUsuario As String
End Type

Type typeFilialContatoData
    lCliente As Long
    iFilial As Integer
    dtData As Date
    iLigar As Integer
    iLigacaoEfetuada As Integer
    sHistorico As String
    sCodUsuario As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeCobrancaEmailPadrao
    lNumIntDoc As Long
    lCodigo As Long
    sDescricao As String
    iAtrasoDe As Integer
    iAtrasoAte As Integer
    sCC As String
    sAssunto As String
    sMensagem As String
    sModelo As String
    sAnexo As String
    iTipo As Integer
    sDe As String
    sNomeExibicao As String
    sUsuarioExclusivo As String
    iConfirmacaoLeitura As Integer
    sEmailResp As String
End Type

Type typeMnemonicoCobrEmail
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

Type typeRelFluxoCaixa
    lNumIntRel As Long
    dtData As Date
    dSaldoInicial As Double
    dEntrada As Double
    dSaida As Double
End Type

Type typeRelFlCxAn
    lNumIntRel As Long
    dtDataVenctoReal As Date
    dSaldoInicial As Double
    dEntrada As Double
    dSaida As Double
    dtDataVencto As Date
    iStatus As Integer
    lNumTitulo As Long
    iNumParcelas As Integer
    iNumParcela As Integer
    sNomeReduzido As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeChequePrePag
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    iContaCorrente As Integer
    lNumero As Long
    dtDataEmissao As Date
    dtDataBomPara As Date
    dtDataDeposito As Date
    dValor As Double
    iStatus As Integer
    sObservacao As String
    sFavorecido As String
    lFornecedor As Long
    iFilial As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeChequePrePagParc
    lNumIntCheque As Long
    lNumIntParcela As Long
    dValorPago As Double
    dValorBaixado As Double
    dJuros As Double
    dMulta As Double
    dDesconto As Double
End Type

Type typePagAntecipBaixado
     lNumIntPag As Long
     dValorBaixado As Double
     sNomeReduzido As String
     lFornecedor As Long
     iFilial_Fornecedor As Integer
     dValor As Double
     dtDataMovimento As Date
     iCodConta As Integer
     iTipoMeioPagto As Integer
     lNumIntBaixa As Long
     lNumMovto As Long
     sContaCorrenteNome As String
End Type

Type typeBaixaPagAntecipadosItem
    lNumIntBaixa As Long
    lNumDocOrigem As Long
    dValor As Double
    iStatus As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeAdmExtFinArqsLidos
    lNumIntDoc As Long
    sNomeArq As String
    dtDataImportacao As Date
    dHoraImportacao As Double
    sUsuario As String
    dtDataAtualizado As Date
    iFilialEmpresa As Integer
    iBandeira As Integer
    iNaoAtualizar As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeAdmExtFinMov
    lNumIntDoc As Long
    lNumIntArq As Long
    sEstabelecimento As String
    iTipo As Integer
    iCodConta As Integer
    dtData As Date
    dValorBruto As Double
    dValorComissao As Double
    dValorRejeitado As Double
    dValorLiq As Double
    iIgnorarErros As Integer
    iFilialEmpresa As Integer
    lNumMovto As Long
    iBandeira As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeAdmExtFinMovDet
    lNumIntDoc As Long
    lNumIntMov As Long
    iTipo As Integer
    dValor As Double
    sNumCartao As String
    dtDataCompra As Date
    iNumParcela As Integer
    sAutorizacao As String
    sRO As String
    sNSU As String
    lNumIntParc As Long
    iCodErro As Integer
    sObservacao As String
    lNumIntBaixaParcRec As Long
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRetPagtoDet
    lNumIntDocRet As Long
    iSeq As Integer
    iLote As Integer
    iSeqLote As Integer
    iTipoMov As Integer
    iCodInstMov As Integer
    iBanco As Integer
    sAgencia As String
    sConta As String
    sNomeFavorecido As String
    sSeuNumero As String
    sNossoNumero As String
    dtDataPagto As Date
    dValorPagto As Double
    dtDataReal As Date
    dValorReal As Double
    sFinalidade As String
    sCodOCR1 As String
    sCodOCR2 As String
    sCodOCR3 As String
    sCodOCR4 As String
    sCodOCR5 As String
    iTipo As Integer
    sCodigoBarras As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRetPagto
    lNumIntDoc As Long
    sNomeArq As String
    dtDataImport As Date
    dHoraImport As Double
    iBanco As Integer
    sAgencia As String
    sConta As String
    dtDataGeracao As Date
    dHoraGeracao As Double
    lSeqArquivo As Long
    sNomeEmpresa As String
    sNomeBanco As String
End Type

Function CondPagto_Extrai(ByVal objCombo As Object) As Integer

Dim iIndice As Integer

On Error GoTo Erro_CondPagto_Extrai

    If Len(Trim(objCombo.Text)) > 0 Then
        If gobjCRFAT.iCondPagtoSemCodigo = 0 Then
            CondPagto_Extrai = Codigo_Extrai(objCombo.Text)
        Else
            If objCombo.ListIndex <> -1 Then
                CondPagto_Extrai = objCombo.ItemData(objCombo.ListIndex)
            Else
                CondPagto_Extrai = 0
                For iIndice = 0 To objCombo.ListCount - 1
                    If UCase(objCombo.Text) = UCase(objCombo.List(iIndice)) Then
                        CondPagto_Extrai = objCombo.ItemData(iIndice)
                        Exit For
                    End If
                Next
            End If
        End If
    Else
        CondPagto_Extrai = 0
    End If
    
    Exit Function

Erro_CondPagto_Extrai:

    CondPagto_Extrai = 0

    Exit Function

End Function

Function CondPagto_Traz(ByVal objCondicaoPagto As ClassCondicaoPagto) As String

On Error GoTo Erro_CondPagto_Traz

    'Preenche campo CondicaoPagamento
    If gobjCRFAT.iCondPagtoSemCodigo = 0 Then
        CondPagto_Traz = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida
    Else
        CondPagto_Traz = objCondicaoPagto.sDescReduzida
    End If
    
    Exit Function

Erro_CondPagto_Traz:

    CondPagto_Traz = ""

    Exit Function

End Function
