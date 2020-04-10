Attribute VB_Name = "ErrosContab"
Option Explicit

'C�digos de Erro - Reservado de 9000 at� 9999
Global Const ERRO_LEITURA_LOTE = 1 'Parametros Filial, origem, exercicio, periodo, lote
'Erro na leitura do lote. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_LOTE_ATUALIZADO = 2
'O lote est� atualizado.
Global Const ERRO_LOCK_LOTE = 3 'Parametros Filial, origem, exercicio, periodo, lote
'N�o foi poss�vel fazer o LOCK do Lote que possui a seguinte chave: Filial=%i, Origem= %s, Exercicio= %i, Periodo=%i, Lote=%i.
Global Const ERRO_LEITURA_PERIODOS = 4 'Parametros Exercicio, Periodo
'Ocorreu um erro na leitura da tabela de Periodos. Exercicio = %i e Periodo = %i.
Global Const ERRO_LOCK_PERIODO = 5 'Parametro Exercicio, Periodo
'N�o conseguiu fazer o lock de um registro da tabela Periodo. Exercicio = %i e Periodo = %i.
Global Const ERRO_LEITURA_LANCAMENTOS = 6 'Parametros Filial, Origem, exercicio, periodo, lote
'Erro na leitura dos Lan�amentos Pendentes do lote que possui chave Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Lote= %i.
Global Const ERRO_LEITURA_MVDIACTA = 7 'Sem parametros
'Erro na leitura da tabela de Movimentos di�rios da Conta.
Public Const ERRO_ATUALIZACAO_MVDIACTA = 8  'Parametros Filial, conta, data
'Ocorreu um erro na atualiza��o da tabela de Saldos Di�rios de Conta. Filial = %i, Conta = %s e Data = %s.
Global Const ERRO_ATUALIZACAO_MVPERCTA = 9 'parametro Filial, Exercicio, Conta
'Ocorreu um erro na atualiza��o da tabela de Saldos Peri�dicos de Conta. Filial = %i, Exercicio = %i e Conta = %s.
Global Const ERRO_ATUALIZACAO_MVPERCCL = 10 'parametros Filial, exercicio, ccl, conta
'Ocorreu um erro na atualiza��o da tabela que guarda os saldos de centro de custo/lucro. Filial = %i, Exercicio = %i, Centro de Custo/Lucro = %s, Conta = %s.
Global Const ERRO_ATUALIZACAO_MVDIACCL = 11 'parametros Filial, ccl, conta, data
'Ocorreu um erro na atualiza��o da tabela de Saldos Di�rios de Centro de Custo/Lucro. Filial = %i, Centro de Custo/Lucro = %s, Conta = %s e Data = %s.
Global Const ERRO_LEITURA_EXERCICIO = 12 'Parametro Exercicio
'Erro na leitura do exercicio %i
Global Const ERRO_LANCAMENTOS_EXERCICIO_FECHADO = 13 'Parametro Exercicio
'O Exerc�cio %i est� fechado. N�o � poss�vel fazer grava��o ou exclus�o de lan�amentos.
Global Const ERRO_LOCACAO_EXERCICIO = 14
'N�o conseguiu fazer o Lock do Exerc�cio.
Global Const ERRO_LEITURA_PERIODO1 = 15 'Sem parametros
'Ocorreu um erro na leitura da tabela de Periodos. Verifique se o periodo existe para os parametros fornecidos.
Global Const ERRO_LEITURA_CONFIGURACAO = 16 'Sem parametros
'Erro na leitura da tabela de Configura��o
Global Const CONTA_SEM_CCL = 17
'Conta n�o possui Centro de Custo/Lucro.
Global Const ERRO_LEITURA_MVPERCTA = 18  'Nao tem parametros
'Ocorreu um erro na leitura da tabela que cont�m os Saldos das Contas (MvPerCta).
Global Const ERRO_INSERCAO_LOTE = 19 'Parametros  Filial, Origem, Exercicio, Periodo, Lote
'Ocorreu um erro na inclus�o de um lote na tabela de Lotes. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_LEITURA_PLANOCONTA = 20 'Sem Parametros
'Erro na leitura da tabela Plano de Contas.
Global Const ERRO_INSERCAO_LANCAMENTO = 21 'Parametros Filial, origem, exercicio, periodo, doc, seq
'Ocorreu um erro no cadastramento do seguinte lan�amento. Filial = %i, Origem = %s, Exercicio=%i, Periodo=%i, Documento=%l, Sequencial=%i.
Global Const ERRO_ATUALIZACAO_LOTE = 22 'Parametro Filial, origem, exercicio, periodo, lote
'Erro de atualiza��o do Lote. Filial= %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_EXERCICIO_NAO_FECHADO = 23
'O Exerc�cio n�o est� fechado.
Global Const ERRO_ATUALIZACAO_EXERCICIOS = 24 'Parametro Exercicio
'Ocorreu um erro na atualiza��o do exercicio %i.
Global Const ERRO_EXERCICIO_NAO_ABERTO = 25 ' Parametro Exercicio
'O Exercicio %s n�o est� aberto. Portanto n�o � capaz de receber lan�amentos.
Global Const ERRO_LEITURA_MVPERCCL = 26  'Sem parametro
'Erro de leitura da tabela MvPerCcl.
Global Const ERRO_EXERCICIO_JA_EXISTE = 27 'Parametro Exercicio
'O exercicio %i j� est� cadastradado.
Global Const ERRO_INSERCAO_EXERCICIO = 28 'Parametro Exercicio
'Erro na inser��o do exercicio %i na tabela Exercicios.
Global Const ERRO_INSERCAO_MVPERCTA = 29 'Parametros: Filial, Exercicio, Conta
'Erro de inser��o na tabela de saldos de conta (MvPerCta). Filial = %i, Exercicio = %i, Conta= %s.
Global Const ERRO_LEITURA_CONTACCL = 30 'Sem Parametro
'Erro de leitura na tabela ContaCcl.
Global Const ERRO_INSERCAO_MVPERCCL = 31 'Parametros Filial, Exercicio, Ccl, Conta
'Ocorreu um erro na inser��o da Filial = %i Exercicio = %i Ccl = %s Conta = %s na tabela de Saldos de Centro de Custo/Lucro(MvPerCcl).
Global Const ERRO_ATUALIZACAO_MVPERCTA1 = 32  'Sem parametro
'Erro de atualiza��o na tabela de saldos de conta.
Global Const ERRO_ATUALIZACAO_MVPERCCL1 = 33 'Sem parametros
'Erro de atualiza��o na tabela de saldos de Centro de Custo/Lucro.
Global Const ERRO_LOCACAO_PERIODO = 34
'N�o conseguiu fazer o Lock do Per�odo.
Global Const ERRO_ATUALIZACAO_PERIODO = 35 'Parametros Exercicio, Periodo
'Ocorreu um erro na atualiza��o da tabela Periodo. Exercicio = %i e Periodo = %i.
Global Const ERRO_UNLOCK_EXERCICIO = 36
'Ocorreu um erro na libera��o do lock do Exerc�cio.
Global Const ERRO_UNLOCK_LOTE = 37
'Ocorreu um erro na libera��o do lock do Lote.
Global Const ERRO_UNLOCK_PERIODO = 38 'Parametros Periodo, Exercicio
'Ocorreu um erro na libera��o do lock do Periodo %s do Exercicio %s
Global Const ERRO_LEITURA_RESULTADO = 39 'Parametro CodigoApuracao
'Ocorreu um erro na leitura da tabela Resultado. Codigo de Apuracao = %l
Global Const ERRO_LEITURA_MVPERCTA1 = 40  'Parametro Filial, Exercicio, Conta
'Ocorreu um erro na leitura da tabela de Saldos de Conta (MvPerCta). Filial=%i, Exercicio=%i, Conta=%s
Global Const ERRO_ATUALIZACAO_PERIODO1 = 41
'Erro de atualiza��o do Per�odo.
Public Const ERRO_INSERCAO_PERIODO = 42 'Parametros Exercicio, Periodo
'Ocorreu um erro ao tentar inserir um registro na tabela de Periodos. Exercicio = %i e Periodo = %i.
Global Const ERRO_LEITURA_LANCAMENTOS1 = 43 'Parametros Exercicio, Periodo
'Erro de leitura na tabela de Lan�amentos. Exerc�cio %s e Per�odo %s.
Global Const ERRO_LOCK_MVPERCTA = 44 'parametro exercicio,conta
'Ocorreu um erro ao tentar executar o "Lock" em um registro da tabela de saldos de conta(MvPerCta). Exercicio = %s, Conta = %s.
Global Const ERRO_UNLOCK_MVPERCTA = 45 'parametro exercicio, conta
'Ocorreu um erro na libera��o do lock em um registro da tabela de saldos de Centro de Custo / Lucro. Exerc�cio = %s e Centro de Custo/Lucro = %s e Conta = %s.
Global Const ERRO_LOCK_MVPERCCL = 46 'parametro Filial, exercicio, ccl, conta
'Ocorreu um erro ao tentar executar o "Lock" em um registro da tabela de saldos de centro de custo/lucro(MvPerCcl). Filial = %i, Exercicio = %i, Centro de Custo/Lucro = %s, Conta = %s.
Global Const ERRO_UNLOCK_MVPERCCL = 47 'parametro exercicio, ccl, conta
'Ocorreu um erro na libera��o do lock em um registro da tabela de saldos de Centro de Custo / Lucro. Exerc�cio %s, Centro de Custo/Lucro %s e Conta %s.
Global Const ERRO_LOCK_MVDIACTA = 48 'parametro conta, data
'Ocorreu um erro ao tentar executar o 'Lock' em um registro da tabela de saldos di�rios de conta. Conta %s e Data %s.
Global Const ERRO_UNLOCK_MVDIACTA = 49 'parametro conta, data
'Ocorreu um erro na libera��o do lock em um registro da tabela de saldos di�rios de conta. Conta %s e Data %s.
Global Const ERRO_LEITURA_MVPERCCL1 = 50  'parametros Filial, exercicio,ccl, conta
'Ocorreu um erro de leitura na tabela de Saldos de Centro de Custo/Lucro (MvPerCcl) para a Filial = %i, exercicio = %i, centro de custo/lucro = %s e conta %s.
Global Const ERRO_LOCK_MVDIACCL = 51 'parametro ccl, conta, data
'Ocorreu um erro ao tentar executar o 'Lock' em um registro da tabela de saldos di�rios de Centro de Custo/Lucro. Centro de Custo/Lucro %s, Conta %s e Data %s.
Global Const ERRO_UNLOCK_MVDIACCL = 52 'parametro ccl, conta, data
'Ocorreu um erro na libera��o do lock em um registro da tabela de saldos di�rios de Centro de Custo/Lucro. Centro de Custo/Lucro %s, Conta %s e Data %s.
Global Const ERRO_LEITURA_MVDIACTA1 = 53 'parametro Filial, conta, data
'Ocorreu um erro na leitura da tabela de Saldos Di�rio de Conta. Filial=%i, Conta=%s e Data=%s.
Global Const ERRO_EXCLUSAO_LANCAMENTO = 54 'parametro Filial, origem, exercicio, periodo, doc, seq
'Ocorreu um erro na exclus�o do lan�amento. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Doc = %l, Seq = %i.
Public Const ERRO_LEITURA_MVDIACCL = 55 'Sem parametro
'Ocorreu um erro na leitura da tabela de Saldos Di�rios de Centro de Custo/Lucro.
Public Const ERRO_LEITURA_MVDIACCL1 = 56 'parametro Filial, ccl, conta, data
'Ocorreu um erro na leitura da tabela de Saldos Di�rios de Centro de Custo/Lucro. Filial = %i, Centro de Custo/Lucro = %s, Conta = %s e Data = %s.
Public Const ERRO_INSERCAO_MVDIACCL = 57 'parametro Filial, ccl, conta, data
'Ocorreu um erro na inser��o de um registro na tabela de Saldos Di�rios de Centro de Custo/Lucro. Filial=%i, Centro de Custo/Lucro=%s, Conta=%s e Data=%s.
Public Const ERRO_INSERCAO_MVDIACTA = 58 'parametro  Filial, conta, data
'Ocorreu um erro na inser��o de um registro na tabela de Saldos Di�rios de Conta. Filial=%i, Conta=%s e Data=%s.
Global Const ERRO_INSERCAO_SORT = 59 'parametro  conta, data
'Erro na inser��o de dados no arquivo de sort. Conta %s e Data %s.
Global Const ERRO_LEITURA_EXERCICIOORIGEM = 60 'parametros Filial, exercicio, periodo, origem
'Ocorreu um erro na leitura da tabela ExercicioOrigem. Filial = %i, Exercicio = %i, Periodo = %i, Origem = %s.
Global Const ERRO_LEITURA_LOTE1 = 61 'Sem parametros
'Ocorreu um erro na leitura da tabela de lotes.
Global Const ERRO_LOCK_EXERCICIOORIGEM = 62 'parametros Filial, exercicio, periodo, origem
'Ocorreu um erro ao tentar fazer um "lock" de um registro da tabela ExercicioOrigem. Filial = %i, Exercicio = %i, Periodo = %i, Origem = %s.
Global Const ERRO_ATUALIZACAO_EXERCICIOORIGEM = 63 'parametro Filial, exercicio, periodo, origem
'Ocorreu um erro na atualiza��o da tabela ExercicioOrigem. Filial = %i, Exercicio = %i, Periodo = %i, Origem = %s.
Global Const ERRO_NUMERO_LOTE_NAO_PREENCHIDO = 64
'N�mero do Lote n�o foi preenchido.
Global Const ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO = 65
'N�mero do Documento n�o foi preenchido.
Global Const ERRO_VALOR_LANCAMENTO_NAO_PREENCHIDO = 66
'Valor do Lan�amento n�o foi preenchido.
Global Const ERRO_LEITURA_LOTE2 = 67 'Parametros origem, exercicio, periodo, lote
'O Lote n�o est� cadastrado - Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i
Global Const ERRO_LEITURA_LOTE_VAZIO = 68
'Tabela de Lotes vazia.
Global Const ERRO_LEITURA_PERIODO = 69 'Parametro Exercicio, Periodo
'Erro de Leitura na Tabela de Periodos
Global Const ERRO_LOTE_NAO_ATUALIZADO = 70 'Parametros Origem, Exercicio, Periodo, Lote
'Erro na atualiza��o do lote. Origem %s, Exerc�cio %s, Per�odo %s e Lote %s.
Global Const ERRO_ATUALIZACAO_EXERCICIO = 71 'Parametro Exercicio
'Erro de atualiza��o do Exercicio %i.
Global Const ERRO_LOTE_EXERCICIO_FECHADO = 72 'Parametro Exercicio
'Lote n�o pode ser criado num exercicio fechado. Exercicio = %i
Global Const ERRO_LOTE_ATUALIZADO_NAO_EDITAVEL = 73 'Parametros Filial, Origem, Exercicio, Periodo, Lote
'Este lote est� atualizado e portanto n�o pode ser editado. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_LOTE_ATUALIZADO_NAO_EXCLUIR = 74 'Parametros Filial, Origem, Exercicio, Periodo, Lote
'Este lote est� atualizado e portanto n�o pode ser removido. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_EXCLUSAO_LOTE = 75 'Parametros Origem, Exercicio, Periodo, Lote
'Houve um erro na exclus�o do lote do banco de dados.
Global Const ERRO_LEITURA_CONTA = 76  'Sem parametros.
'Erro de leitura da tabela Plano de Contas
Global Const ERRO_LEITURA_ORIGEM = 77
'Erro de leitura na tabela Origem.
Global Const ERRO_PLANO_CONTAS_VAZIO = 78
'Tabela de Plano de Contas Vazia.
Global Const ERRO_INSERCAO_LANCAMENTOS = 79 'Sem parametros
'Erro na Inser��o dos Lan�amentos na Tabela de Lan�amentos Pendentes
Global Const ERRO_M_LOTE_LOTE_ATUALIZADO = 80 'Parametros Filial, sOrigem, iExercicio, iPeriodo, iLote
'Este lote j� foi contabilizado, portanto n�o pode ser editado. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_AUSENCIA_LANCAMENTOS_GRAVAR = 81 'Sem parametros
'N�o h� Lan�amentos para Gravar.
Global Const ERRO_TABELA_CCL_VAZIA = 82 'Sem parametro
'Tabela de Centros de Custo e de Lucro Vazia.
Global Const ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA = 83 'Sem Parametro
'Data do Documento n�o preenchida.
Global Const ERRO_LEITURA_PLANOCONTA1 = 84 'Parametro Conta
'A conta %s n�o est� cadastrada
Global Const ERRO_LEITURA_HISTPADRAO = 85 'Parametro HistPadrao
'Erro na leitura do Historico Padrao - Historico = %i
Global Const ERRO_LEITURA_PLANOCONTA2 = 86 'Parametro ContaSimples
'Erro na leitura do Plano de Contas. Conta Simplificada = %l
Global Const ERRO_CONTASIMPLES_JA_UTILIZADA = 87 'Parametros ContaSimples, Conta
'A conta simplificada %s � utilizada na conta %s
Global Const ERRO_LEITURA_LANCAMENTOS2 = 88 'Parametro Conta
'Erro na leitura da tabela de Lan�amentos. Conta = %s
Global Const ERRO_LEITURA_LANPENDENTE = 89 'Parametro Conta
'Erro na leitura da tabela de Lan�amentos Pendentes. Conta = %s
Global Const ERRO_LEITURA_CONTACCL1 = 90 'Parametro Conta
'Erro na leitura da tabela ContaCcl. Conta = %s.
Global Const ERRO_LEITURA_PLANOCONTA3 = 91 'Parametro Conta
'Erro na leitura do Plano de Contas. Conta = %s
Global Const ERRO_CONTAPAI_INEXISTENTE = 92 'Sem Parametro
'A conta em quest�o n�o tem uma conta "pai" dentro da hierarquia do plano de contas.
Global Const ERRO_CONTA_SINTETICA_COM_LANCAMENTOS = 93 'Sem parametro
'A conta n�o pode ser sint�tica pois possui lan�amentos j� contabilizados.
Global Const ERRO_CONTA_SINTETICA_COM_LANC_PEND = 94 'Sem parametro
'A conta n�o pode ser sint�tica pois possui lan�amentos pendentes.
Global Const ERRO_CONTA_SINTETICA_ASSOCIADA_CCL = 95 'Sem parametro
'A conta n�o pode ser sint�tica pois est� associada a centro de custo.
Global Const ERRO_CONTA_ANALITICA_COM_FILHAS = 96 'Sem parametro
'A conta n�o pode ser anal�tica pois possui contas embaixo dela.
Global Const ERRO_DOCUMENTO_JA_LANCADO = 97 'Parametros: lDocumento, sOrigem, iExercicio, iPeriodoLan
'O documento <lDocumento> (Origem: <sOrigem>, Exerc�cio: <iExercicio>, Per�odo: <iPeriodoLan>) j� foi lan�ado.
Global Const ERRO_CONTA_NAO_INFORMADA = 98 'Sem parametro
'A conta n�o foi informada.
Global Const ERRO_LEITURA_LANCAMENTOS_PENDENTES = 99 'Sem parametro
'Erro na Leitura da Tabela de Lan�amentos Pendentes
Global Const ERRO_LEITURA_LANCAMENTOS3 = 100 'Sem parametro
'Erro na Leitura da Tabela de Lan�amentos
Global Const ERRO_MASCARA_CONTA_OBTERNIVEL = 101 'Parametro Conta
'Erro na obten�ao do n�vel da conta. Conta = %s.
Global Const ERRO_INSERCAO_PLANOCONTA = 102 'Parametro Conta
'Erro na inser��o da conta %s na tabela PlanoConta.
Global Const ERRO_LEITURA_EXERCICIOS = 103 'Sem parametro
'Erro de leitura da tabela Exercicios. Verifique se o(s) exercicio(s) existe(m) para os parametros informados.
Global Const ERRO_ATUALIZACAO_PLANOCONTA = 104 'Parametro Conta
'Erro de atualiza��o da conta %s.
Global Const ERRO_LOTE_ATUALIZADO_NAO_RECEBE_LANCAMENTOS = 105 'Par�metros: FilialEmpresa, iLote, iExercicio, iPeriodo, sOrigem
'Lote com chave Filial= %i Lote = %i Exerc�cio = %i Per�odo =  %i Origem = %s est� Atualizado. N�o pode incluir/alterar/excluir Lan�amentos.
Global Const ERRO_DOCUMENTO_NAO_BALANCEADO = 106 'Parametro: lDoc
'O Documento Cont�bil <lDoc> n�o est� balanceado (soma de cr�ditos diferente soma de d�bitos).
Public Const ERRO_LEITURA_LANPENDENTE1 = 107 'Parametros Filial, Origem, Exercicio, Periodo, Lote
'Erro na leitura dos Lan�amentos Pendentes do Lote que possui a chave Filial = %i, Origem = %s, Exericicio = %i, Periodo = %i, Lote = %i
Global Const ERRO_LOTE_COM_LANC_PEND_NAO_EXCLUIR = 108 'Parametros Filial, Origem, Exercicio, Periodo, Lote
'Lote com lan�amentos pendentes n�o pode ser excluido. Filial = %i, Origem = %s, Exericicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_LOCK_PLANOCONTA = 109 'Parametro Conta
'N�o conseguiu fazer o lock da conta %s.
Global Const ERRO_EXCLUSAO_PLANOCONTA = 110 'Parametro Conta
'Houve um erro na exclus�o da conta %s do banco de dados.
Global Const ERRO_EXCLUSAO_CONTA_COM_LANCAMENTOS = 111 'Sem Parametros
'A conta possui lan�amentos contabilizados. Portanto n�o pode ser excluida.
Global Const ERRO_EXCLUSAO_CONTA_COM_LANC_PEND = 112 'Sem Parametros
'A conta possui lan�amentos pendentes. Portanto n�o pode ser excluida.
Global Const ERRO_LOTE_INEXISTENTE = 113  'Parametros Origem, exercicio, periodo, lote
'O lote n�o est� cadastrado.  Origem= %s, Exercicio= %i, Periodo= %i, Lote= %i
Global Const ERRO_EXCLUSAO_CONTACCL = 114 'Parametro Conta, Ccl
'Houve um erro na exclus�o da associa��o da conta %s com o centro de custo/lucro %s.
Global Const ERRO_LEITURA_DOCAUTO = 115 'Sem parametro
'Erro de leitura da tabela de Documentos Autom�ticos.
Global Const ERRO_EXCLUSAO_DOCAUTO = 116 'Parametros documento, sequencial
'Houve um erro na exclus�o do documento autom�tico %l, sequencial %i do banco de dados.
Global Const ERRO_LEITURA_RATEIOON = 117 'Parametro conta.
'Erro de leitura da tabela de Rateios On-Line para a conta %s.
Global Const ERRO_LEITURA_RATEIOOFF = 118 'Parametro conta.
'Erro de leitura da tabela de Rateios Off-Line para a conta %s.
Global Const ERRO_CONTA_NAO_CADASTRADA = 119 'Parametro conta.
'A conta %s n�o est� cadastrada.
Global Const ERRO_EXCLUSAO_CONTA_COM_RATEIOON = 120 'Sem Parametros
'A conta � usada no rateio on-line. Portanto n�o pode ser excluida.
Global Const ERRO_EXCLUSAO_CONTA_COM_RATEIOOFF = 121 'Sem Parametros
'A conta � usada no rateio off-line. Portanto n�o pode ser excluida.
Global Const ERRO_LEITURA_LOTE3 = 122 'Parametros origem, exercicio, lote
'Erro na leitura do lote - Origem = %s, Exercicio = %i, Lote = %i
Global Const ERRO_LEITURA_LOTE4 = 123 'Parametros origem, exercicio, lote
'O Lote n�o est� cadastrado - Origem = %s, Exercicio = %i, Lote = %i
Global Const Erro_Mascara_RetornaContaNoNivel = 124 'Parametros Conta, Nivel
'Erro na obten�ao da conta %s no n�vel %i.
Global Const ERRO_PERIODOS_DIFERENTES = 125 'Par�metros Periodo do Documento, Periodo do Lote
'Per�odo do Documento %i diferente do Per�odo do Lote %i.
Global Const ERRO_LOTE_ATUALIZADO_NAO_SE_EXCLUI = 126 'Par�metros: iLote, iExercicio, iPeriodo, sOrigem
'Lote com chave Lote <iLote>, Exerc�cio <iExercicio>, Per�odo <iPeriodo>, Origem <sOrigem> est� Atualizado. N�o pode ser exclu�do.
Global Const ERRO_LOTE_INEXISTENTE1 = 127 'iLote, iExercicio, iPeriodo, sOrigem
'Nao existe Lote com chave Lote <iLote>, Origem <sOrigem>, Periodo <iPeriodo>, Exerc�cio <iExercicio>.
Global Const ERRO_DOCUMENTO_NAO_EXISTE = 128 'Parametro: lDoc
'N�o existe Documento <lDoc>.
Global Const ERRO_EXERCICIO_FECHADO = 129 'Parametro Exercicio.
'O Exercicio %i encontra-se Fechado. Portanto n�o pode receber novos lan�amentos.
Global Const ERRO_LEITURA_CCL = 130 'Parametro Ccl
'Erro na leitura da tabela de Centros de Custo/Lucro. Centro de Custo/Lucro = %s
Global Const ERRO_LEITURA_CONTACCL2 = 131 'Parametro Ccl
'Erro na leitura da tabela ContaCcl. Centro de Custo/Lucro = %s.
Global Const ERRO_LEITURA_LANCAMENTOS4 = 132 'Parametro Ccl
'Erro na leitura da tabela de Lan�amentos. Centro de Custo/Lucro = %s
Global Const ERRO_LEITURA_LANPENDENTE2 = 133 'Parametro Ccl
'Erro na leitura da tabela de Lan�amentos Pendentes. Centro de Custo/Lucro = %s
Global Const ERRO_CCL_NAO_CADASTRADO = 134 'Parametro Ccl.
'O Centro de Custo/Lucro %s n�o est� cadastrado.
Global Const ERRO_EXCLUSAO_CCL_COM_LANCAMENTOS = 135 'Sem Parametros
'O Centro de Custo/Lucro possui lan�amentos contabilizados. Portanto n�o pode ser excluido.
Global Const ERRO_EXCLUSAO_CCL_COM_LANC_PEND = 136 'Sem Parametros
'O Centro de Custo/Lucro possui lan�amentos pendentes. Portanto n�o pode ser excluido.
Global Const ERRO_CCL_NAO_INFORMADO = 137 'Sem parametro
'O Centro de Custo/Lucro n�o foi informado.
Global Const ERRO_LOCK_CCL = 138 'Parametro Ccl
'N�o conseguiu fazer o lock do centro de custo/lucro %s.
Global Const ERRO_EXCLUSAO_CCL = 139 'Parametro Ccl
'Houve um erro na exclus�o do Centro de Custo/Lucro %s do banco de dados.
Global Const ERRO_ATUALIZACAO_CCL = 140 'Parametro Ccl
'Erro de atualiza��o do Centro de Custo/Lucro %s.
Global Const ERRO_INSERCAO_CCL = 141 'Parametro Ccl
'Erro na inser��o do Centro de Custo/Lucro %s na tabela de centros de custo.
Global Const ERRO_LEITURA_MVPERCTA2 = 142  'Parametro Conta
'Ocorreu um erro de leitura na tabela de Saldos de Conta (MvPerCta) para a conta %s.
Global Const ERRO_EXCLUSAO_CONTA_COM_MOVIMENTO = 143 'Sem Parametros
'A conta possui movimento, portanto n�o pode ser excluida.
Global Const ERRO_EXCLUSAO_MVPERCTA = 144 'Parametro Exercicio, conta
'Houve um erro na exclus�o do saldo de conta (MvPerCta) do Exercicio %i, Conta %s.
Global Const ERRO_LEITURA_MVPERCCL2 = 145  'Parametro Conta
'Ocorreu um erro de leitura na tabela de Saldos de Centro de Custo/Lucro (MvPerCcl) para a conta %s.
Global Const ERRO_EXCLUSAO_MVPERCCL = 146 'Parametro Exercicio, ccl, conta
'Houve um erro na exclus�o do saldo de centro de custo/lucro (MvPerCcl) do Exercicio %i, Ccl %s, Conta %s.
Global Const ERRO_CONTAPAI_ANALITICA = 147 'Sem Parametro
'A conta em quest�o possui uma conta "pai" anal�tica. Contas anal�ticas n�o podem conter contas embaixo dela.
Global Const ERRO_ORIGEM_NAO_PREENCHIDA = 148 'Sem Parametro
'Lan�amentos devem ter origem.
Global Const ERRO_EXCLUSAO_CCL_COM_MOVIMENTACAO = 149 'Sem Parametros
'O Centro de Custo/Lucro possui movimento, portanto n�o pode ser excluido.
Global Const ERRO_EXCLUSAO_CCL_COM_RATEIO = 150 'Sem Parametros
'O Centro de Custo/Lucro possui Rateio associado. Portanto n�o pode ser excluido.
Global Const ERRO_EXCLUSAO_CCL_COM_RATEIOON = 151 'Sem Parametros
'O Centro de Custo/Lucro � usado no rateio on-line. Portanto n�o pode ser excluido.
Global Const ERRO_EXCLUSAO_CCL_COM_RATEIOOFF = 152 'Sem Parametros
'O Centro de Custo/Lucro � usado no rateio off-line. Portanto n�o pode ser excluido.
Global Const ERRO_LEITURA_MVPERCCL3 = 153  'Parametro Ccl
'Ocorreu um erro de leitura na tabela de Saldos de Centro de Custo/Lucro (MvPerCcl) para o Centro de Custo/Lucro %s.
Global Const ERRO_EXCLUSAO_CONTA_COM_LANC_PEND1 = 154 'Parametro Conta
'A conta %s possui lan�amentos pendentes. Portanto n�o pode ser excluida.
Global Const ERRO_EXCLUSAO_CONTA = 155 'Parametro Conta
'Houve um erro na exclus�o da Conta %s do banco de dados.
Global Const ERRO_EXCLUSAO_MVPERCCL1 = 156 'Parametro ccl
'Houve um erro na exclus�o do saldo de centro de custo/lucro (MvPerCcl) do Exercicio %i, Ccl %s, Conta %s.
Global Const ERRO_LEITURA_HISTPADRAO1 = 157 'Sem Parametros
'Erro na leitura da tabela de Hist�rico Padr�o
Global Const ERRO_LOCK_CONFIGURACAO = 158 'Sem Parametros
'N�o conseguiu fazer o lock na tabela de Configura��o.
Global Const ERRO_ATUALIZACAO_CONFIGURACAO = 159 'Sem Parametros
'Erro de atualiza��o da tabela de Configura��o.
Global Const ERRO_HISTPADRAO_NAO_INFORMADO = 160 'Sem parametro
'O C�digo do Hist�rico Padr�o n�o foi informado.
Global Const ERRO_HISTPADRAO_NAO_CADASTRADO = 161 'Parametro Hist�ricoPadr�o
'O Hist�rico Padr�o %i n�o est� cadastrado.
Global Const ERRO_ATUALIZACAO_HISTPADRAO = 162 'Parametro Hist�ricoPadr�o
'Erro de atualiza��o do Hist�rico Padr�o %i.
Global Const ERRO_INSERCAO_HISTPADRAO = 163 'Parametro Hist�ricoPadr�o
'Erro na inser��o do Hist�rico Padr�o %i na tabela Hist�rico Padr�o.
Global Const ERRO_LOCK_HISTPADRAO = 164 'Parametro Hist�ricoPadr�o
'N�o conseguiu fazer o lock do Hist�rico Padr�o %i.
Global Const ERRO_EXCLUSAO_HISTPADRAO = 165 'Parametro Hist�ricoPadr�o
'Houve um erro na exclus�o do Hist�rico Padr�o %i do banco de dados.
Global Const ERRO_HISTPADRAO_PRESENTE_PLANO_CONTAS = 166 'Parametro Hist�ricoPadr�o
'N�o � poss�vel excluir o Hist�rico Padr�o %i que est� presente no Plano de Contas.
Global Const ERRO_COLUNA_GRID_INEXISTENTE = 167 'Parametro Titulo da Coluna
'A coluna cujo t�tulo �: %s n�o foi encontrada no Grid.
Global Const ERRO_LOCK_EXERCICIO = 168 'Parametro Exercicio
'N�o conseguiu fazer o lock do Exercicio %i.
Global Const ERRO_LOTE_PERIODO_FECHADO = 169 'Parametro Exercicio, Periodo
'Lote n�o pode ser criado num periodo fechado. Exercicio = %i, Periodo = %i
Global Const ERRO_CONTA_INATIVA = 170 'Parametro Conta
'Esta conta n�o est� ativa. Conta = %s.
Global Const ERRO_CONTA_NAO_ANALITICA = 171 'Parametro conta
'A conta %s n�o � anal�tica, portanto n�o pode ter lan�amentos associados.
Global Const ERRO_INSERCAO_CONTACCL = 172 'Parametros Conta e Ccl
'Erro na inser��o da associa��o da conta %s com o centro de custo/lucro %s na tabela ContaCcl.
Global Const ERRO_EXCLUSAO_ASSOC_CCLCONTA_COM_MOV = 173 'Parametros Conta e Ccl
'Existe movimento para a associa��o da conta %s com o centro de custo/lucro %s. Portanto n�o pode-se excluir a associa��o.
Global Const ERRO_EXCLUSAO_CONTACCL_COM_RATEIOON = 174 'Parametros Conta Ccl
'A associa��o da Conta %s com o Centro de Custo/Lucro %s � usado no rateio on-line. Portanto n�o pode ser excluida.
Global Const ERRO_EXCLUSAO_CONTACCL_COM_RATEIOOFF = 175 'Parametros Conta Ccl
'A associa��o da Conta %s com o Centro de Custo/Lucro %s � usado no rateio off-line. Portanto n�o pode ser excluido.
Global Const ERRO_LEITURA_CONTACCL3 = 176 'Parametros Conta Ccl
'Erro na leitura da tabela ContaCcl. Conta = %s, Centro de Custo/Lucro = %s.
Global Const ERRO_CONTACCL_NAO_CADASTRADO = 177 'Parametros Conta Ccl
'A associa��o da Conta %s com o Centro de Custo/Lucro %s n�o est� cadastrada.
Global Const ERRO_LOCK_CONTACCL = 178 'Parametros Conta Ccl
'N�o conseguiu fazer o lock da associa��o da Conta %s com o Centro de Custo/Lucro %s.
Global Const ERRO_UNLOCK_PLANOCONTA = 179 'Parametro Conta
'N�o conseguiu fazer liberar o lock da conta %s.
Global Const ERRO_UNLOCK_CONTACCL = 180 'Parametros Conta Ccl
'N�o conseguiu liberar o lock da associa��o da Conta %s com o Centro de Custo/Lucro %s.
Global Const ERRO_DESCRICAO_NAO_INFORMADA = 181 'Sem parametro
'Descri��o do Hist�rico Padr�o n�o foi informada.
Global Const ERRO_LEITURA_LANPENDENTE3 = 182 'Sem parametros
'Erro na leitura da tabela de Lan�amentos Pendentes.
Global Const ERRO_DOC_NAO_CADASTRADO = 183 'Parametros Origem, Exercicio, PeriodoLan, Doc
'O Documento n�o est� cadastrado. Origem = %s Exercicio = %i Periodo = %i Documento = %l
Global Const Erro_Mascara_MascararCcl = 184 'Parametro Ccl
'Erro na formata��o do Centro de Custo/Lucro %s.
Global Const Erro_Mascara_MascararConta = 185 'Parametro Conta
'Erro na formata��o da Conta %s.
Global Const Erro_Mascara_RetornaContaPai = 186 'Parametro Conta
'Erro na fun��o que retorna a conta de nivel imediatamente superior da Conta %s.
Public Const ERRO_DESCRICAO_COM_CARACTER_INICIAL_ERRADO = 187 'Sem parametros
'Descri��o de Hist�rico n�o pode come�ar com este caracter.
Global Const ERRO_CODIGO_HISTPADRAO_INVALIDO = 188 'Parametro codigo do historico padrao(string)
'Ap�s o asterisco deve ser digitado o c�digo de um hist�rico padr�o existente no sistema. O c�digo digitado foi %s.
Global Const ERRO_LEITURA_CCL1 = 189 'Sem Parametros
'Erro na leitura da tabela de Centros de Custo/Lucro.
Global Const ERRO_CONTA_SEG_MEIO_NAO_PREENCHIDOS = 190 'Sem parametro
'Todos os segmentos da conta tem que estar preenchidos. Ex: 1.000.1 est� errado. 1.001.1 est� correto.
Global Const ERRO_CCL_SEG_MEIO_NAO_PREENCHIDOS = 191 'Sem parametro
'Todos os segmentos do centro de custo tem que estar preenchidos. Ex: 1.000.1 est� errado. 1.001.1 est� correto.
Global Const ERRO_CONFIGURACAO_NAO_CADASTRADA = 192 'Sem parametro
'Os dados de configura��o n�o est�o cadastrados
Global Const ERRO_LEITURA_EXERCICIO_DATA = 193 'Parametro Data(String)
'N�o foi encontrado exerc�cio para a data %s.
Global Const ERRO_LEITURA_EXERCICIO1 = 194 'Sem Parametro
'Erro de leitura da tabela de Exercicios.
Global Const ERRO_LEITURA_LOTEPENDENTE = 195 'Parametros FilialEmpresa, origem, exercicio, periodo, lote
'Erro na leitura do lote pendente. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i
Global Const ERRO_LOCK_LOTEPENDENTE = 196 'Parametros FilialEmpresa, origem, exercicio, periodo, lote
'N�o foi poss�vel fazer o "lock" do Lote Pendente que possui a chave: Filial = %i, Origem= %s, Exercicio= %i, Periodo=%i, Lote=%i.
Global Const ERRO_INSERCAO_LOTEPENDENTE = 197 'Parametros Filial, Origem, Exercicio, Periodo, Lote
'Ocorreu um erro ao tentar inserir o lote na tabela de lotes pendentes. Filial=%i, Origem=%s, Exercicio=%i, Periodo=%i, Lote=%i
Global Const ERRO_ATUALIZACAO_LOTEPENDENTE = 198 'Parametro FilialEmpresa, origem, exercicio, periodo, lote
'Erro na atualiza��o do Lote Pendente. Filial=%i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_UNLOCK_LOTEPENDENTE = 199
'N�o conseguiu liberar o lock do lote pendente.
Global Const ERRO_LEITURA_LOTEPENDENTE1 = 200 'Sem parametros
'Erro na leitura da tabela de lotes pendentes.
Global Const ERRO_LEITURA_LOTEPENDENTE2 = 201 'Parametros origem, exercicio, periodo, lote
'O Lote n�o est� cadastrado como lote pendente - Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i
Global Const ERRO_LOTEPENDENTE_INEXISTENTE = 202  'Parametros FilialEmpresa, Origem, exercicio, periodo, lote
'O lote n�o est� cadastrado na tabela de lotes pendentes.  Filial = %i, Origem= %s, Exercicio= %i, Periodo= %i, Lote= %i
Global Const ERRO_EXCLUSAO_LOTEPENDENTE = 203 'Parametros Filial, Origem, Exercicio, Periodo, Lote
'Houve um erro na exclus�o do lote pendente do banco de dados. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i e Lote = %i.
Global Const ERRO_MASCARAR_CONTA = 204 'Parametro Conta
'Erro ao tentar mascarar a conta %s
Global Const ERRO_MASCARAR_CCL = 205 'Parametro Ccl
'Erro ao tentar mascarar o centro de custo/lucro %s
Global Const ERRO_FORMATAR_CONTA = 206 'Parametro Conta
'Erro na formata��o da conta %s
Global Const ERRO_FORMATAR_CCL = 207 'Parametro Ccl
'Erro na formata��o do centro de custo/lucro %s
Global Const ERRO_INICIALIZACAO_TELA = 208 'Sem parametro
'Erro na inicializa��o da tela %s
Global Const ERRO_DOCAUTO_NAO_CADASTRADO = 209 'Sem parametro
'Documento procurado n�o foi encontrado
Global Const ERRO_INSERCAO_DOCAUTO = 210 'Sem parametro
'Erro na atualizacao da tabela de Documento Automatico
Global Const ERRO_AUSENCIA_DOCAUTO_GRAVAR = 211 'Sem parametro
'O Grid est� vazio, h� aus�ncia de dados para gravar
Global Const ERRO_NUMERO_RATEIO_NAO_PREENCHIDO = 212 'Sem Parametro
'O C�digo do Rateio n�o foi digitado
Global Const ERRO_AUSENCIA_RATEIOON_GRAVAR = 213 'Sem Parametro
'N�o existem Lancamentos no Grid para gravar
Global Const ERRO_SOMA_NAO_VALIDA = 214  'Sem Parametro
'A Soma dos Rateios tem que totalizar 100%.
Global Const ERRO_INSERCAO_RATEIOON = 215  'Sem Parametro
'Erro na insercao de Registros na Tabela de Rateios OnLine
Global Const ERRO_RATEIOON_NAO_CADASTRADO = 216  'Parametro Codigo
'N�o existe rateio cadastrado com o codigo %d
Global Const ERRO_EXCLUSAO_RATEIOON = 217  'Parametro Codigo
'Erro na exclusao do Rateio de codigo %d
Global Const ERRO_VALOR_PERCENTUAL = 218   'Sem Parametro
'O valor digitado como percentual de rateio deve estar entre 0 e 100
Global Const ERRO_EXISTENCIA_CONTA = 219 'Sem parametro
'Erro na verificacao da existencia de conta na tabela PlanoConta
Global Const ERRO_LEITURA_SEGMENTO = 221 'Sem parametro
'Erro na leitura da tabela Segmento.
Global Const ERRO_VALOR_FORMATO_NAO_PREENCHIDO = 222 'Sem parametro
'Campo formato n�o preenchido.
Global Const ERRO_VALOR_TIPO_NAO_PREENCHIDO = 223 'Sem parametro
'Campo tipo n�o preenchido.
Global Const ERRO_VALOR_TAMANHO_NAO_PREENCHIDO = 224 'Sem parametro
'Campo tamanho n�o preenchido.
Global Const ERRO_VALOR_PREENCHIMENTO_NAO_PREENCHIDO = 225 'Sem parametro
'Campo preenchimento n�o preenchido.
Global Const ERRO_VALOR_DELIMITADOR_NAO_PREENCHIDO = 226 'Sem parametro
'Campo delimitador n�o preenchido.
Global Const ERRO_SAIDA_DELIMITADOR = 227 'Sem parametro
'O delimitador n�o pode ter mais de um caracter.
Global Const ERRO_SEGMENTO_CONTA_MAIOR_PERMITIDO = 228 'Parametros tamanho do segmento, tamanho total permitido
'O tamanho do segmento da conta %i ultrapassou o tamanho total permitido %i.
Global Const ERRO_SEGMENTO_CCL_MAIOR_PERMITIDO = 229 'Parametros tamanho do segmento, tamanho total permitido
'O tamanho do segmento do centro de custo/lucro %i ultrapassou o tamanho total permitido %i.
Global Const ERRO_VALOR_TAMANHO_INVALIDO = 230
'O tamanho do segmento tem que ser maior do que zero.
Global Const ERRO_MODIFICACAO_CONFIGURACAO = 231 'Parametros TipoConta , Origem , Natureza
'Erro na tentativa de modificar TipoConta para %i , Origem para %s , Natureza para %i na tabela Configuracao.
Global Const ERRO_INSERCAO_EXERCICIOORIGEM = 232  'Parametros Exercicio, Periodo, Origem
'Erro na insercao de Registro na Tabela ExercicioOrigem. Exercicio = %i, Periodo = %i, Origem=%s
Global Const ERRO_EXCLUSAO_EXERCICIO = 233 'Parametro Exercicio
'Houve um erro na exclus�o do exercicio %i da tabela Exercicios.
Global Const ERRO_EXCLUSAO_PERIODO = 234 'Parametro Periodo, Exercicio
'Houve um erro na exclus�o do periodo %i do exercicio %i da tabela de Periodos
Global Const ERRO_LEITURA_EXERCICIOORIGEM1 = 235 'Sem parametros
'Ocorreu um erro na leitura na tabela ExercicioOrigem.
Global Const ERRO_EXCLUSAO_EXERCICIOORIGEM = 236  'Parametros Filial, Exercicio, Periodo, Origem
'Erro na exclus�o de registro na Tabela ExercicioOrigem. Filial = %i, Exercicio = %i, Periodo = %i, Origem=%s
Global Const Erro_Mascara_RetornaCcl = 237 'Parametro Conta
'Erro na fun��o que retorna o centro de custo/lucro associado a conta cont�bil %s.
Global Const ERRO_LOTE_NAO_DESATUALIZADO = 238 'Parametros Filial, Origem, Exercicio, Periodo, Lote
'O Lote (Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i) n�o est� pronto para ser atualizado. Verifique se este lote est� incompleto ou j� foi atualizado.
Global Const ERRO_LOTE_SENDO_ATUALIZADO = 239 'Filial, Origem, Exercicio, Periodo, Lote
'O Lote (Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i) j� est� em processo de atualiza��o por outro programa.
'Global Const ERRO_LEITURA_ORCAMENTO = 240 'Sem parametros
'Erro na leitura da tabela Orcamento
Global Const ERRO_INSERCAO_ORCAMENTO = 241 'Sem parametros
'Erro na insercao de registro na tabela Orcamento
Global Const ERRO_EXCLUSAO_ORCAMENTO_CTB = 242 'Sem parametros
'Erro na exclusao de registro na tabela Orcamento
Global Const ERRO_MASCARA_RETORNACONTAENXUTA = 243 'Parametro Conta
'Erro na formata��o da conta %s.
Global Const ERRO_CONTACREDITO_NAO_DIGITADA = 244
'Conta Credito n�o digitada
Global Const ERRO_CONTAORIGEM_NAO_DIGITADA = 245
'A Conta Origem nao foi digitada
Global Const ERRO_CCLORIGEM_NAO_DIGITADO = 246
'O CclOrigem n�o foi digitado
Global Const ERRO_RATEIOOFF_NAO_CADASTRADO = 247
'Esse Rateio n�o esta cadastrado
Global Const ERRO_EXCLUSAO_RATEIOOFF = 248 'Sem Parametros
'Erro na Exclus�o de registros na tabela RateioOff.
Global Const ERRO_CONTA_JA_UTILIZADA = 249 'Parametro conta
'A conta %s j� foi utiliazada
Global Const ERRO_INSERCAO_RATEIOOFF = 250
'Ocorreu um erro ao tentar inserir dados na tabela RateioOff
Global Const ERRO_AUSENCIA_RATEIOOFF_GRAVAR = 251
'Nao exite nenhuma linha de rateio no grid
Global Const ERRO_EXCLUSAO_RESULTADO = 252 'Parametro CodigoApuracao
'Ocorreu um erro na exclus�o de um registro da tabela Resultado. Codigo de Apuracao = %l
Global Const ERRO_EXERCICIO_POSTERIOR_INEXISTENTE = 253 'Sem Parametro
'Para que ocorra o fechamento de um exercicio � necess�rio que o exercicio seguinte esteja criado. Favor criar o exercicio.
Public Const ERRO_DOC_ATUALIZADO = 254 'Parametros Filial, Origem, Exercicio, PeriodoLan, Doc
'Existe um documento com este n�mero contabilizado. Filial = %i, Origem = %s Exercicio = %i Periodo = %i Documento = %l
Public Const ERRO_DOC_PENDENTE = 255 'Parametros Filial, Origem, Exercicio, PeriodoLan, Doc
'Existe um documento pendente com este n�mero. Filial = %i, Origem = %s Exercicio = %s Periodo = %s Documento = %l
Public Const ERRO_CONTA_COM_MOVIMENTO = 256 'Parametro Conta, Ccl
'N�o foi poss�vel desfazer a associa��o da conta %s com o Centro de Custo %s, pois a conta possui movimenta��o.
Public Const ERRO_EXERCICIOS_FECHADOS = 257
'Todos os Exercicios est�o fechados.
Public Const ERRO_CONTAS_SEM_PREENCHIMENTO = 258
'Nenhumas das Contas inicias est�o Preenchidas
Public Const ERRO_CONTARESULTADO_NAO_PREENCHIDA = 259
'A conta Resultado n�o foi preenchida
Public Const ERRO_EXERCICIO_NAO_SELECIONADO = 260
'O exercicio nao foi selecionado.
Public Const ERRO_CONTARECEITA_INICIAL_MAIOR = 261
'A Conta Receita inicial � maior que a Conta Receita Final
Public Const ERRO_CONTADESPESA_INICIAL_MAIOR = 262
'A Conta Despesa Inicial � maior que a Conta Despesa Final
Public Const ERRO_INTERSECAO_CONJUNTO_CONTAS = 263
'As contas de Receita e despesa se interceptam
Public Const ERRO_INTERSECAO_CONTARESULTADO = 264
'A Conta Resultado faz parte do conjunto de contas envolvidos
Public Const ERRO_PERIODO_INICIAL_NAO_SELECIONADO = 265
'O Periodo Inicial n�o foi selecionado
Public Const ERRO_PERIODO_FINAL_NAO_SELECIONADO = 266
'O Periodo Final n�o foi selecionado
Public Const ERRO_PERIODO_INICIAL_MAIOR = 267
'O Periodo Inicial � maior que o final
Public Const ERRO_CONTACONTRAPARTIDA_VAZIA = 268
'A conta de ContraPartida est� vazia
Public Const ERRO_GRID_VAZIO = 269
'Nao h� contas no grid para Apurar
Public Const ERRO_CONTARESULTADO_VAZIA = 270 'Parametro ilinha
'A conta Resultado na %d linha nao foi preenchida
Public Const ERRO_CONTAINICIAL_VAZIA = 271 'Parametro ilinha
'A conta Inicial na %d linha nao foi informada
Public Const ERRO_CONTAFINAL_VAZIA = 272 'Parametro ilinha
'A conta Final da linha %d n�o foi informada
Public Const ERRO_CONTRAPARTIDA_IGUAL_RESULTADO = 273
'A conta de ContraPartida � igual a conta resultado
Public Const ERRO_INSERCAO_RESULTADO = 274
'Erro na insercao de registros na tabela de resultado
Public Const ERRO_INTERSECAO_CONTACONTRAPARTIDA = 275
'A conta de contraPartida intercepta um cojunto de contas
Public Const ERRO_CONTA_INICIAL_MAIOR = 276
'A conta Inicial � maior que a conta final
Public Const ERRO_EXERCICIO_NAO_PREENCHIDO = 277
'O Exerc�cio n�o foi preenchido
Public Const ERRO_INTERSECAO_CONTAS = 278
'As contas de Ativo e Passivo se interceptam
Public Const ERRO_PASSIVAINICIAL_MAIOR = 279
'conta Passiva Inicial � maior que a conta Passiva Final
Public Const ERRO_ATIVAINICIAL_MAIOR = 280
'Conta Ativa Inicial � maior que a Ativa final
Public Const ERRO_TODOS_EXERCICIOS_FECHADOS = 281
'Todos os exercicios j� estao Fechados
Public Const ERRO_FALTA_LOTE = 282
'Nao existe nenhum lote marcado para ser atualizado.
Public Const ERRO_MODIFICACAO_CONFIG = 283 'Parametro IdAtualizacao
'Erro na tentativa de modificar o campo IdAtualizacao na tabela LotePendente para %i .
Public Const ERRO_LEITURA_LOTE_PENDENTE = 284 'Parametro Origem, Exercicio, Periodo, Lote
'Erro na leitura dos campos Origem = %s , Exercicio = %i , Periodo = %i , Lote = %i da tabela LotePendente
Public Const ERRO_MODIFICACAO_LOTEPENDENTE = 285 'Parametro IdAtualizacao
'Erro na tentativa de modificar o campo IdAtualizacao = %i na tabela LotePendente.
Public Const ERRO_TODOS_EXERCICIOS_ABERTOS = 286  'Sem parametros
'Todos os Exerc�cios est�o abertos.
Public Const ERRO_DATAS_COM_EXERCICIOS_DIFERENTES = 287
'Data Inicial e Final devem estar num mesmo exerc�cio.
Public Const ERRO_LOTE_INICIAL_MAIOR = 288
'Lote inicial n�o pode ser maior que o lote final.
Public Const ERRO_DATA_INICIAL_MAIOR = 289
'Data inicial n�o pode ser maior que a data final.
Public Const ERRO_NOME_RELOP_VAZIO = 290
'O campo Op��o de relat�rio tem que estar preenchido
Public Const ERRO_NOME_RELOP_NAO_SELEC = 291
'N�o existe relat�rio selecionado.
Public Const ERRO_EXERCICIO_VAZIO = 292
'O campo Exerc�cio tem que estar preenchido.
Public Const ERRO_PERIODO_VAZIO = 293
'O campo Per�odo tem que estar preenchido.
Public Const ERRO_CCL_INICIAL_MAIOR = 294
'O Centro de Custo inicial n�o pode ser maior que o Centro de Custo final.
Public Const ERRO_LOTE_FORA_FAIXA = 295
'O lote final deve estar entre 1 e 9999.
Public Const ERRO_PERIODO_INICIAL_VAZIO = 296
'O per�odo inicial tem que estar preenchido.
Public Const ERRO_PERIODO_FINAL_VAZIO = 297
'O per�odo final tem que estar preenchido.
Public Const ERRO_LEITURA_CONTACATEGORIA = 298 'Parametro Codigo
'Erro na leitura da tabela Categoria. Categoria = %i.
Public Const ERRO_NOME_CATEGORIA_NAO_INFORMADO = 299 'Sem parametro
'O nome da categoria n�o foi informado.
Public Const ERRO_CATEGORIA_NAO_CADASTRADA = 300 'Parametro Codigo
'A Categoria %i n�o est� cadastrada.
Public Const ERRO_LOCK_CONTACATEGORIA = 301 'Parametro Codigo
'N�o conseguiu fazer o lock da Categoria %i.
Public Const ERRO_EXCLUSAO_CONTACATEGORIA = 302 'Parametro Codigo
'Houve um erro na exclus�o da Categoria %i do banco de dados.
Public Const ERRO_CATEGORIA_PRESENTE_PLANO_CONTAS = 303 'Parametro Codigo
'N�o � poss�vel excluir a Categoria %i que est� presente no Plano de Contas.
Public Const ERRO_ATUALIZACAO_CONTACATEGORIA = 304 'Parametro Codigo
'Erro de atualiza��o da Categoria %i.
Public Const ERRO_INSERCAO_CONTACATEGORIA = 305 'Parametro Codigo
'Erro na inser��o da Categoria %i na tabela ContaCategoria.
Public Const ERRO_CODIGO_CATEGORIA_NAO_INFORMADO = 306 'Sem parametro
'O codigo da categoria n�o foi informado.
Public Const ERRO_LEITURA_CONTACATEGORIA1 = 307 'Sem Parametros
'Erro na leitura da tabela de ContaCategoria.
Public Const ERRO_CONTA_NIVEL1_CATEGORIA = 308 'Sem Parametro
'A conta � de nivel 1. Favor designar uma categoria para esta conta.
Public Const ERRO_LEITURA_PLANOCONTA4 = 309 'Parametros Categoria, Nivel
'Erro na leitura do Plano de Contas. Categoria = %i Nivel = %i
Public Const ERRO_LEITURA_CONTACATEGORIA2 = 310 'Parametro Nome
'Erro na leitura da tabela Categoria. Categoria = %s.
Public Const ERRO_CATEGORIA_NAO_CADASTRADA1 = 311 'Parametro Nome
'A Categoria %s n�o est� cadastrada.
Public Const ERRO_CONTA_CATEG_NIVEL_NAO_CADASTRADA = 312 'Parametro Codigo da Categoria, Nivel da Conta
'N�o est� cadastrado uma conta da categoria %i no nivel %i.
Public Const ERRO_LEITURA_CTBCONFIG = 313 'Parametro Codigo
'Erro na leitura da tabela CTBConfig. Codigo = %s.
Public Const ERRO_INTERSECAO_CONTARESULTADO_APURACAO = 314
'A Conta Resultado faz parte do conjunto de contas a serem apuradas
Public Const ERRO_INTERSECAO_CONTRAPARTIDA_APURACAO = 315
'A Conta de Contra Partida faz parte do conjunto de contas a serem apuradas
Public Const ERRO_MASCARA_RETORNAULTIMACONTA = 316 'Parametro Conta
'Erro ao tentar retornar a ultima conta do nivel da conta %s
Public Const ERRO_PLANOCONTA_SEM_CATEGORIA_ATIVO = 317 'Sem Parametro
'N�o foi encontrado no plano de contas nenhum grupo designado com a categoria 'Ativo'.
Public Const ERRO_PLANOCONTA_SEM_CATEGORIA_PASSIVO = 318 'Sem Parametro
'N�o foi encontrado no plano de contas nenhum grupo designado com a categoria 'Passivo'.
Public Const ERRO_CONTA_NAO_SELECIONADA = 319 'Sem parametro
'Nenhuma conta foi selecionada. Favor selecionar pelo menos uma conta antes de usar esta fun��o.
Public Const ERRO_CCL_NAO_SELECIONADA = 320 'Sem parametro
'Nenhum Centro de custo/lucro foi selecionado. Favor selecionar um centro de custo/lucro antes de usar esta fun��o.
Public Const ERRO_LEITURA_LANPENDENTE4 = 321 'Parametros Conta e Centro de Custo/Lucro
'Erro na leitura da tabela de Lan�amentos Pendentes. Conta = %s e Centro de Custo/Lucro = %s.
Public Const ERRO_EXCLUSAO_CCLCONTA_COM_LANPENDENTE = 322 'Parametros Conta e Ccl
'Existem lan�amentos pendentes para a associa��o da conta %s com o centro de custo/lucro %s. Portanto n�o pode-se excluir a associa��o.
Public Const ERRO_LEITURA_RATEIOON1 = 323 'Parametros conta e ccl.
'Erro de leitura da tabela de Rateios On-Line para a conta %s e centro de custo/lucro %s.
Public Const ERRO_EXCLUSAO_CCLCONTA_COM_RATEIOON = 324 'Parametros Conta e Ccl
'Existem rateios on-line para a associa��o da conta %s com o centro de custo/lucro %s. Portanto n�o pode-se excluir a associa��o.
Public Const ERRO_PLANOCONTA_SEM_CONTA_SINTETICA = 325 'Sem Parametro
'N�o foi encontrado no plano de contas nenhuma conta sint�tica.
Public Const ERRO_CCL_VAZIO = 326 'Sem parametro
'N�o h� centro de custo/lucro cadastrado.
Public Const ERRO_PLANOCONTA_SEM_CONTA_ANALITICA = 327 'Sem Parametro
'N�o foi encontrado no plano de contas nenhuma conta anal�tica.
Public Const ERRO_CONTACCL_VAZIO = 328 'Sem parametro
'N�o h� associa��o de conta com centro de custo/lucro cadastrada.
Public Const ERRO_CONTA_SEM_CONTACCL = 329 'Parametro Conta
'A conta %s n�o possui nenhuma associa��o com centro de custo/lucro.
Public Const ERRO_ATUALIZACAO_CONTACCL = 330 'Parametros Conta Ccl
'Erro na atualiza��o da tabela que guarda a associa��o da Conta %s com o Centro de Custo/Lucro %s.
Public Const ERRO_SALDOS_INICIAIS_NAO_ALTERAVEIS = 331 'Sem Parametro
'N�o � poss�vel fazer atualiza��o dos saldos iniciais de conta. Verifique se o exerc�cio inicial (exercicio de implanta��o) est� presente e aberto.
Public Const ERRO_CONTA_NAO_ANALITICA_SALDO = 332 'Parametro conta
'A conta %s n�o � anal�tica. Somente as contas anal�ticas podem receber saldos iniciais.
Public Const ERRO_LEITURA_MVPERCCL4 = 333  'parametros ccl, conta
'Ocorreu um erro de leitura na tabela de Saldos de Centro de Custo/Lucro (MvPerCcl). Centro de custo/lucro = %s e Conta = %s.
Public Const ERRO_LOCK_MVPERCCL1 = 334 'parametros  ccl, conta
'Ocorreu um erro ao tentar executar o "Lock" em um registro da tabela de saldos de centro de custo/lucro(MvPerCcl). Centro de Custo/Lucro = %s, Conta = %s.
Public Const ERRO_ATUALIZACAO_MVPERCCL2 = 335 'parametros ccl, conta
'Ocorreu um erro na atualiza��o da tabela que guarda os saldos de centro de custo/lucro. Centro de Custo/Lucro = %s, Conta = %s.
Public Const ERRO_LEITURA_EXERCICIO2 = 336 'Parametro Exercicio
'Erro na leitura do exercicio %i
Public Const ERRO_EXERCICIO_NAO_CADASTRADO = 337 'Parametro Exercicio
'O Exerc�cio %i n�o est� cadastrado.
Public Const ERRO_DATA_INICIO_PERIODO_VAZIA = 338 'parametros Periodo
'A data de inicio do periodo %i est� vazia. Favor preenche-la.
Public Const ERRO_NOME_PERIODO_VAZIO = 339 'Parametro Periodo
'O Nome do per�odo %i est� vazio. Favor preenche-lo.
Public Const ERRO_CONTA_NAO_ANALITICA1 = 340 'Parametro conta
'A conta %s n�o � anal�tica. Somente contas analiticas podem ser utilizadas.
Public Const ERRO_DATAINICIO_EXERCICIO_ALTERADA = 341 'Parametros Data Inicial Antiga, Data Inicial Nova
'A data inicial do exercicio nao pode ser modificada. Data Antiga = %s , Data Nova = %s.
Public Const ERRO_DATAFIM_EXERCICIO_ALTERADA = 342 'Parametros Data Final Antiga, Data Final Nova
'A data final do exercicio nao pode ser modificada. Data Antiga = %s , Data Nova = %s.
Public Const ERRO_NUMERO_PERIODOS_ALTERADO = 343 'Parametros Numero de Periodos Antigo, Numero de Periodos Novo
'O Exercicio possui movimento. O n�mero de periodos n�o pode ser alterado. Num.Periodos Antigo = %i , Num.Periodos Novo = %i.
Public Const ERRO_LEITURA_PERIODO2 = 344 'Parametro Exercicio
'Ocorreu um erro na leitura dos Periodos do Exercicio %i.
Public Const ERRO_DATAINICIO_NOVO_EXERCICIO = 345 'Parametros Data Inicial Correta, Data Inicial Digitada
'A data inicial de um novo exercicio deve ser a data final do ultimo exercicio. Data Correta = %s , Data Digitada = %s
Public Const ERRO_NOME_EXERCICIO_JA_USADO = 346 'Parametro NomeExterno, Exercicio
'O Nome de exerc�cio %s j� foi usado pelo exerc�cio %i. Favor escolher outro nome.
Public Const ERRO_LEITURA_ORCAMENTO1 = 347 'Parametro Exercicio
'Erro na leitura da tabela Orcamento. Exercicio = %i.
Public Const ERRO_LEITURA_LANPENDENTE5 = 348 'Parametro Exercicio
'Erro na leitura da tabela de Lan�amentos Pendentes. Exercicio = %i.
Public Const ERRO_LEITURA_LOTEPENDENTE3 = 349 'Parametro Exercicio
'Erro na leitura da tabela de Lotes Pendentes. Exercicio = %i.
Public Const ERRO_EXERCICIO_NAO_ULTIMO = 350 'Parametro Exercicio
'Somente o ultimo exercicio pode ser excluido. Ultimo Exercicio = %i.
Public Const ERRO_EXERCICIO_COM_MOVIMENTO = 351 'Exercicio
'O Exerc�cio %i possui movimento cont�bil associado ou n�o est� aberto.
Public Const ERRO_EXERCICIO_NAO_ENCONTRADO_TELA = 352 'Exercicio
'O Exercicio %i n�o foi encontrado entre os listados nesta tela. O exercicio pode n�o estar cadastrada ou a tela estar desatualizada.
Public Const ERRO_DATA_FINAL_EXERCICIO_MENOR = 353 'Parametros Data Final e Data Inicial do Exercicio
'A data final do Exercicio %s � menor que a inicial %s.
Public Const ERRO_DATA_INICIAL_EXERCICIO_MAIOR = 354 'Parametros Data Inicial e Data Final do Exercicio
'A data inicial do Exercicio %s � maior que a final %s.
Public Const ERRO_DATA_INICIAL_EXERCICIO_NAO_PREENCHIDA = 355 'Sem parametros
'A data inicial do Exercicio n�o foi preenchida.
Public Const ERRO_DATA_FINAL_EXERCICIO_NAO_PREENCHIDA = 356 'Sem parametros
'A data final do Exercicio n�o foi preenchida.
Public Const ERRO_PERIODICIDADE_INVALIDA = 357 'iPeriodicidade
'A Periodicidade %i � inv�lida.
Public Const ERRO_NUM_PERIODO_INVALIDO = 358 'parametro Maximo de Periodos permitido pelo sistema
'N�mero do per�odo inv�lido. Faixa v�lida: 1 a %i.
Public Const ERRO_TOTAL_PERIODOS_MAIOR_TOTAL_DIAS = 359 'Parametros: Numero de Periodos, Total de Dias do Exercicio
'O n�mero de per�odos requeridos = %i � maior do que o total de dias do Exercicio = %i
Public Const ERRO_NOME_PERIODO_JA_USADO = 360 'Parametros: Nome do Periodo, Periodo
'O Nome de Per�odo %s j� foi utilizado pelo per�odo %i.
Public Const ERRO_DATA_FORA_EXERCICIO = 361 'Parametros Data Inicial do Periodo , Data Inicial do Exercicio e Data Final do Exercicio
'A Data Inicial deste periodo %s n�o est� dentro das faixa abrangida pelo exerc�cio. Data Inicial = %s e Data Final = %s.
Public Const ERRO_DATAINI_PERIODO_MENOR_PERIODO_ANT = 362 'Parametros Data Inicio Periodo e Data Inicio Periodo Anterior
'A data inicial de cada periodo tem que ser maior que a data inicial do periodo anterior. Data Inicio Periodo = %s e Data Inicio Periodo Anterior = %s.
Public Const ERRO_DATAINI_PRIMEIRO_PERIODO = 363
'A data inicial do primeiro per�odo deve ser igual a data de in�cio do exerc�cio. Data Inicial do Primeiro Periodo = %s e Data Inicial do Exercicio = %s
Public Const ERRO_EXERCICIO_SEM_PERIODO = 364 'Parametro Exercicio
'Nao foram especificados periodos para o exerc�cio %i.
Public Const ERRO_PERIODOS_DEMAIS = 365 'Parametros Total de Periodos, Maximo de Periodos do Sistema
'O n�mero de periodos que seriam gerados automaticamente, %i,  ultrapassa o limite do sistema, %i.
Public Const ERRO_STATUS_EXERCICIO_INVALIDO = 366 'Parametro Status do Exercicio
'Status de exerc�cio igual a %i � inv�lido
Public Const ERRO_NOME_EXERCICIO_VAZIO = 367
'O Nome do Exerc�cio n�o foi preenchido.
Public Const ERRO_DATAINI_PERIODO_MAIOR_PERIODO_SEG = 368 'Parametros Data Inicio Periodo e Data Inicio Periodo Seguinte
'A data inicial de cada periodo tem que ser menor do que a data inicial do periodo seguinte. Data Inicio Periodo = %s e Data Inicio Periodo Seguinte = %s.
Public Const ERRO_LANCAMENTOS_PERIODO_FECHADO = 369 'Parametro Periodo, Exercicio
'O Periodo %i do Exerc�cio %i est� fechado. N�o � poss�vel fazer grava��o ou exclus�o de lan�amentos.
Public Const ERRO_LEITURA_PERIODOSFILIAL = 370 'Parametros Filial, Exercicio, Periodo
'Ocorreu um erro na leitura da tabela PeriodosFilial. Filial = %i, Exercicio = %i e Periodo = %i.
Public Const ERRO_LOCK_PERIODOSFILIAL = 371 'Parametros Filial, Exercicio, Periodo
'Ocorreu um erro ao tentar executar o "Lock" em um registro da tabela PeriodosFilial. Filial = %i, Exercicio = %i e Periodo = %i.
Public Const ERRO_LEITURA_EXERCICIOSFILIAL = 372 'Parametros Filial, Exercicio
'Ocorreu um erro na leitura da tabela ExerciciosFilial. Filial = %i e Exercicio = %i.
Public Const ERRO_LOCK_EXERCICIOSFILIAL = 373 'Parametros Filial, Exercicio
'Ocorreu um erro ao tentar executar o "Lock" em um registro da tabela ExerciciosFilial. Filial = %i e Exercicio = %i.
Public Const ERRO_ATUALIZACAO_EXERCICIOSFILIAL = 374 'Parametro Exercicio, Filial
'Ocorreu um erro na atualiza��o do exercicio %i da Filial %i.
Public Const ERRO_LEITURA_PERIODOSFILIAL1 = 375 'Sem Parametros
'Ocorreu um erro na leitura da tabela PeriodosFilial.
Public Const ERRO_EXCLUSAO_PERIODOSFILIAL = 376 'Parametros Filial,Exercicio, Periodo
'Houve um erro na exclus�o de um registro da tabela PeriodosFilial. Filial = %i, Exercicio = %i e Periodo = %i.
Public Const ERRO_LEITURA_PERIODOSFILIAL2 = 377 'Parametros Filial, Exercicio
'Ocorreu um erro na leitura da tabela PeriodosFilial. Filial = %i e Exercicio = %i.
Public Const ERRO_ATUALIZACAO_PERIODOSFILIAL = 378 'Parametros Periodo, Exercicio, Filial
'Ocorreu um erro na atualiza��o do Periodo %i do Exercicio %i da Filial %i.
Public Const ERRO_LEITURA_SALDOINICIALCONTA = 379 'Parametros Filial, Conta
'Ocorreu um erro na leitura da tabela SaldoInicialConta. Filial = %i e Conta = %s.
Public Const ERRO_LEITURA_SALDOINICIALCONTACCL = 380 'Parametros Filial, Conta, Ccl
'Ocorreu um erro na leitura da tabela SaldoInicialContaCcl. Filial = %i, Conta = %s e Centro de Custo/Lucro = %s.
Public Const ERRO_INSERCAO_PERIODOSFILIAL = 381 'Parametros Filial, Exercicio, Periodo
'Ocorreu um erro ao tentar inserir um registro na tabela PeriodosFilial. Filial = %i, Exercicio = %i e Periodo = %i.
Public Const ERRO_INSERCAO_EXERCICIOSFILIAL = 382 'Parametros Filial, Exercicio
'Ocorreu um erro ao tentar inserir um registro na tabela ExerciciosFilial. Filial = %i e Exercicio = %i.
Public Const ERRO_EXERCICIO_FECHADO1 = 383 'Parametro Exercicio.
'O Exercicio %i encontra-se Fechado. Portanto n�o pode ter seus dados alterados.
Public Const ERRO_LEITURA_SALDOINICIALCONTACCL1 = 384 'Parametros Conta, Ccl
'Ocorreu um erro na leitura da tabela SaldoInicialContaCcl. Conta = %s e Centro de Custo/Lucro = %s.
Public Const ERRO_EXCLUSAO_SALDOINICIALCONTACCL = 385 'Parametros Filial, Conta, Ccl
'Houve um erro na exclus�o de um registro da tabela SaldoInicialContaCcl. Filial = %i, Conta = %s e Centro de Custo/Lucro = %s.
Public Const ERRO_LEITURA_SALDOINICIALCONTA1 = 386 'Parametro Conta
'Ocorreu um erro na leitura da tabela SaldoInicialConta. Conta = %s.
Public Const ERRO_EXCLUSAO_SALDOINICIALCONTA = 387 'Parametros Filial, Conta
'Houve um erro na exclus�o de um registro da tabela SaldoInicialConta. Filial = %i e Conta = %s.
Public Const ERRO_ATUALIZACAO_SALDOINICIALCONTA = 388 'Parametros Filial, Conta
'Ocorreu um erro na atualiza��o da tabela SaldoInicialConta. Filial = %i e Conta = %s.
Public Const ERRO_MASCARA_CCL_OBTERNIVEL = 389 'Parametro Ccl
'Erro na obten�ao do n�vel do centro de custo/lucro. Centro de Custo/Lucro = %s.
Public Const Erro_Mascara_RetornaCclNoNivel = 390 'Parametros Ccl, Nivel
'Erro na obten�ao do centro de custo/lucro %s no n�vel %i.
Public Const ERRO_LEITURA_SALDOINICIALCONTACCL2 = 391 'Parametros Filial, Ccl
'Ocorreu um erro na leitura da tabela SaldoInicialContaCcl. Filial = %i e Centro de Custo/Lucro = %s.
Public Const ERRO_ATUALIZACAO_SALDOINICIALCONTACCL = 392 'Parametros Filial, Ccl
'Ocorreu um erro na atualiza��o da tabela SaldoInicialContaCcl. Filial = %i e Centro de Custo/Lucro = %s.
Public Const ERRO_ATUALIZACAO_MVPERCCL3 = 393 'parametros Filial, exercicio, ccl
'Ocorreu um erro na atualiza��o da tabela que guarda os saldos de centro de custo/lucro. Filial= %i, Exercicio = %i, Centro de Custo/Lucro = %s
Public Const ERRO_LOCK_SALDOINICIALCONTACCL = 394 'Parametros Filial, Conta, Ccl
'Ocorreu um erro ao tentar executar o "Lock" em um registro da tabela SaldoInicialContaCcl. Filial = %i, Conta = %s e Centro de Custo/Lucro = %s.
Public Const ERRO_ATUALIZACAO_SALDOINICIALCONTACCL1 = 395 'Parametros Filial, Conta, Ccl
'Ocorreu um erro na atualiza��o da tabela SaldoInicialContaCcl. Filial = %i, Conta = %s e Centro de Custo/Lucro = %s.
Public Const ERRO_LOCK_SALDOINICIALCONTA = 396 'Parametros Filial, Conta
'Ocorreu um erro ao tentar executar o "Lock" em um registro da tabela SaldoInicialConta. Filial = %i e Conta = %s.
Public Const ERRO_INSERCAO_SALDOINICIALCONTA = 397  'Parametros Filial, Conta
'Erro na insercao de Registro na Tabela SaldoInicialConta. Filial = %i e Conta = %s.
Public Const ERRO_INSERCAO_SALDOINICIALCONTACCL = 398  'Parametros Filial, Conta, Ccl
'Erro na insercao de Registro na Tabela SaldoInicialContaCcl. Filial = %i, Conta = %s e Centro de Custo/Lucro = %s.
Public Const ERRO_LEITURA_SALDOINICIALCONTACCL3 = 399 'Parametros Filial, Conta
'Ocorreu um erro na leitura da tabela SaldoInicialContaCcl. Filial = %i e Conta = %s.
Public Const ERRO_LEITURA_SEGMENTO1 = 400 'Parametro Codigo
'Erro na leitura da tabela Segmento. Codigo = %s.
Public Const ERRO_EXCLUSAO_SEGMENTO = 401 'Parametros Codigo, Nivel
'Ocorreu um erro na exclus�o de um registro da tabela Segmento. Codigo = %s e Nivel = %i.
Public Const ERRO_INSERCAO_SEGMENTO = 402  'Parametros Codigo, Nivel
'Erro na insercao de Registro na Tabela Segmento. Codigo = %s e Nivel = %i.
Public Const Erro_Mascara_RetornaCclPai = 403 'Parametro Ccl
'Erro na fun��o que retorna a centro de custo/lucro de nivel imediatamente superior do centro de custo/lucro %s.
Public Const ERRO_LEITURA_SALDOINICIALCONTACCL4 = 404 'Parametros Ccl
'Ocorreu um erro na leitura da tabela SaldoInicialContaCcl. Centro de Custo/Lucro = %s.
Public Const ERRO_EXCLUSAO_ORCAMENTO1 = 405 'Parametros Exercicio, Conta
'Ocorreu um erro na exclusao de um or�amento. Exercicio = %i, Periodo = %i e Conta = %s.
Public Const ERRO_LEITURA_ORCAMENTO2 = 406 'Parametro Conta
'Erro na leitura da tabela Orcamento. Conta = %s.
Public Const ERRO_LEITURA_ORCAMENTO3 = 407 'Parametro Ccl
'Erro na leitura da tabela Orcamento. Centro de Custo/Lucro = %s.
Public Const ERRO_EXCLUSAO_CONTA_COM_MOVIMENTO1 = 408 'Parametro Conta
'A conta %s possui movimento, portanto n�o pode ser excluida.
Public Const ERRO_EXCLUSAO_CCL_COM_MOVIMENTO = 409 'Parametro Ccl
'O Centro de Custo/Lucro %s possui movimento, portanto n�o pode ser excluido.
Public Const ERRO_CCLPAI_INEXISTENTE = 410 'Sem Parametro
'O Centro de Custo/Lucro em quest�o n�o tem Centro de Custo/Lucro "pai".
Public Const ERRO_CCLPAI_ANALITICA = 411 'Sem Parametro
'O Centro de Custo/Lucro em quest�o possui um Centro de Custo/Lucro "pai" anal�tico. Centros de Custo/Lucro anal�ticos n�o podem conter Centros de Custo/Lucro embaixo dele.
Public Const ERRO_CCL_SINTETICA_ASSOCIADA_CONTA = 412 'parametro Ccl
'O Centro de Custo/Lucro %s n�o pode ser sint�tico pois est� associado a alguma conta.
Public Const ERRO_CCL_ANALITICA_COM_FILHOS = 413 'Parametro Ccl
'O Centro de Custo/Lucro %s n�o pode ser anal�tico pois possui Centro de Custo/Lucro embaixo dele.
Public Const ERRO_MASCARA_RETORNACCLENXUTA = 414 'Parametro Ccl
'Erro na formata��o do Centro de Custo/Lucro %s.
Public Const ERRO_CCL_NAO_ANALITICA = 415 'Parametro Ccl
'O Centro de Custo/Lucro %s n�o � anal�tico, portanto n�o pode ter associa��es com contas.
Public Const ERRO_CCL_NAO_ANALITICA1 = 416 'Parametro Ccl
'O Centro de Custo/Lucro %s n�o � anal�tico.
Public Const ERRO_UNLOCK_PERIODOSFILIAL = 417 'Parametros Filial, Exercicio, Periodo
'Ocorreu um erro ao tentar liberar o "Lock" em um registro da tabela PeriodosFilial. Filial = %i, Exercicio = %i e Periodo = %i.
Public Const ERRO_EXERCICIOSFILIAL_NAO_APURADO = 418 'Parametros Filial, Exercicio
'A Filial %i n�o est� com o exercicio %i apurado.
Public Const ERRO_LEITURA_EXERCICIOSFILIAL2 = 419 'Sem Parametro
'Ocorreu um erro na leitura da tabela ExerciciosFilial.
Public Const ERRO_EXCLUSAO_EXERCICIOSFILIAL = 420 'Parametros Filial,Exercicio
'Houve um erro na exclus�o de um registro da tabela ExerciciosFilial. Filial = %i e Exercicio = %i.
Public Const ERRO_PERCENTUAL_INVALIDO = 421
'Valor de Percentual Inv�lido
Public Const ERRO_SOMA_PERCENTUAL_NAO_VALIDA = 422
' O Total dos percentuais nao totalizou 100%
Public Const ERRO_DELIMITADOR_INVALIDO = 423 'Sem Parametros
'O Delimitador digitado n�o � valido.
Public Const ERRO_LEITURA_EXERCICIOSFILIAL1 = 424 'Parametro Exercicio
'Ocorreu um erro na leitura da tabela ExerciciosFilial. Exercicio = %i.
Global Const ERRO_ORCAMENTO_NAO_CADASTRADO = 425 'Parametros Exercicio, Conta.
'Or�amento n�o cadastrado. Exercicio = %i, Conta = %s.
Public Const ERRO_LANCAMENTO_INEXISTENTE = 426 'Par�metros: iFilialEmpresa, sOrigem, iExercicio, iPeriodoLan, lDoc
'N�o existe Lan�amentos cadastrados com Filial Empresa %i, Origem %s, Exerc�cio %i, Per�odo de lan�amento %i e n�mero de documento %l.
Public Const ERRO_LEITURA_EXERCICIOS1 = 427 'Par�metro: sNomeExterno
'Erro de leitura na tabela de Exerc�cios com nome externo %s.
Public Const ERRO_LEITURA_PERIODO3 = 428 'Par�metros: iExercicio, sNomeExterno
'Erro de leitura na tabela de Per�odo.
Public Const ERRO_EXERCICIO_INEXISTENTE = 429 'Par�metro: sNomeExterno
'O Exerc�cio %s n�o est� cadastrado na tabela de Exercicios.
Public Const ERRO_PERIODO_EXERCICIO_INEXISTENTE = 430 'Par�metros: iExercicio, sNomeExterno
'O Per�odo %s do Exerc�cio com c�digo %i n�o est� cadastrado na tabela de Periodo.
Public Const ERRO_MAX_ARGS_BATH = 431 'Sem par�metros
'O n�mero de argumentos ultrapassou o n�mero m�ximo de argumentos
Public Const ERRO_EXERCICIOSFILIAL_INEXISTENTE = 432 'Par�metro: iExercicio, iFilialEmpresa
'O Exerc�cio %i da Filial %i n�o est� cadastrado na tabela de ExerciciosFilial.
Global Const ERRO_TIPOCCL_NAO_INFORMADO = 433 'Sem parametro
'O Tipo do Centro de Custo/Lucro n�o foi informado.
Public Const Erro_MascaraCcl = 434 'Sem Parametro
'Erro na fun��o que retorna a mascara de centro de custo/lucro.
Public Const ERRO_LOTEAPURACAO_JA_UTILIZADO = 435  'Parametros iLote, iExercicio
'O Lote de Apura��o %i do Exercicio %i j� foi utilizado.
Public Const ERRO_LEITURA_PADRAOCONTABITEM = 436 'Parametros sModulo, iTransacao, sModelo
'Ocorreu um erro na leitura da tabela PadraoContabItem. Modulo = %s, Transa��o = %i, Modelo = %s.
Public Const ERRO_LEITURA_PADRAOCONTAB = 437 'Parametros sModulo, iTransacao
'Ocorreu um erro na leitura da tabela PadraoContab. Modulo = %s, Transa��o = %i.
Public Const ERRO_PADRAOCONTAB_SEM_MODELO_PADRAO = 438 'Parametros sModulo, iTransacao
'O M�dulo = %s, Transa��o = %i n�o possui um modelo de contabiliza��o padr�o.
Public Const ERRO_LEITURA_MNEMONICOCTB = 439 'Parametros sModulo, iTransacao
'Ocorreu um erro na leitura da tabela MnemonicoCTB. Modulo = %s, Transa��o = %i.
Public Const ERRO_GRID_NAO_ENCONTRADO = 440 'Parametros sNomeGrid
'O Grid %s n�o foi encontrado.
Public Const ERRO_LEITURA_MNEMONICOCTB1 = 441 'Parametros sModulo, iTransacao, sMnemonicoCombo
'Ocorreu um erro na leitura da tabela MnemonicoCTB. Modulo = %s, Transa��o = %i, Mnemonico = %s.
Public Const ERRO_LEITURA_FORMULAFUNCAO = 442 'Sem Parametros
'Ocorreu um erro na leitura da tabela FormulaFuncao.
Public Const ERRO_LEITURA_FORMULAFUNCAO1 = 443 'Parametro sFuncaoCombo
'Ocorreu um erro na leitura da tabela FormulaFuncao. Fun��o = %s.
Public Const ERRO_LEITURA_FORMULAOPERADOR = 444 'Sem Parametros
'Ocorreu um erro na leitura da tabela FormulaOperador.
Public Const ERRO_LEITURA_FORMULAOPERADOR1 = 445 'Parametro sOperadorCombo
'Ocorreu um erro na leitura da tabela FormulaOperador. Operador = %s.
Public Const ERRO_LEITURA_TRANSACAOCTB = 446 'Parametro sSiglaModulo
'Ocorreu um erro na leitura da tabela TransacaoCTB. Sigla do Modulo = %s.
Public Const ERRO_TIPO_FORMULA_INVALIDA = 447 'Parametro sFormula, sTipoFormula, sTipoEsperado
'O tipo retornado pela formula %s foi %s e deveria ser %s.
Public Const ERRO_LEITURA_PADRAOCONTAB1 = 448 'Parametros sModulo, iTransacao, sModelo
'Ocorreu um erro na leitura da tabela PadraoContab. Modulo = %s, Transa��o = %i, Modelo = %s.
Public Const ERRO_CREDITO_DEBITO_PREENCHIDOS = 449 'Parametro iLinha
'As f�rmulas de cr�dito e d�bito est�o preenchidas na linha %i. Retire uma das duas.
Global Const ERRO_MODELO_NAO_PREENCHIDO = 450
'O Modelo n�o foi preenchido.
Public Const ERRO_ATUALIZACAO_PADRAOCONTAB = 451 'Parametros sModulo, iTransacao, sModelo
'Ocorreu um erro na atualiza��o da tabela PadraoContab. Modulo = %s, Transa��o = %i, Modelo = %s.
Public Const ERRO_LOCK_PADRAOCONTAB = 452 'Parametros sModulo, iTransacao, sModelo
'Ocorreu um erro no lock de um registro da tabela PadraoContab. Modulo = %s, Transa��o = %i, Modelo = %s.
Public Const ERRO_INCLUSAO_PADRAOCONTAB = 453 'Parametros sModulo, iTransacao, sModelo
'Ocorreu um erro na inclus�o de um registro na tabela PadraoContab. Modulo = %s, Transa��o = %i, Modelo = %s.
Public Const ERRO_EXCLUSAO_PADRAOCONTABITEM = 454 'Parametros sModulo, iTransacao, sModelo, iItem
'Ocorreu um erro na exclus�o de um registro da tabela PadraoContabItem. Modulo = %s, Transa��o = %i, Modelo = %s, Item = %i.
Public Const ERRO_PADRAOCONTAB_INEXISTENTE = 455 'Parametros sModulo, iTransacao, sModelo
'Este modelo de contabiliza��o n�o est� cadastrado. Modulo = %s, Transa��o = %i, Modelo = %s.
Public Const ERRO_EXCLUSAO_PADRAOCONTAB = 456 'Parametros sModulo, iTransacao, sModelo
'Ocorreu um erro na exclus�o de um registro da tabela PadraoContab. Modulo = %s, Transa��o = %i, Modelo = %s.
Public Const ERRO_INCLUSAO_PADRAOCONTABITEM = 457  'Parametros sModulo, iTransacao, sModelo, iItem
'Ocorreu um erro na inser��o de um registro na tabela PadraoContabItem. Modulo = %s, Transa��o = %i, Modelo = %s, Item = %i.
Public Const ERRO_MODELO_CONTAB_SEM_PADRAO = 458 'Parametros sModulo, iTransacao.
'N�o h� um modelo padr�o cadastrado. Modulo = %s, Transa��o = %i.
Public Const ERRO_CALCULO_MNEMONICO_INEXISTENTE = 459 'Parametro sMnemonico
'N�o foi encontrada a fun��o que calcula o valor do campo %s.
Public Const ERRO_NAO_HA_LOTE_PENDENTE = 460 'Par�metro: iFilialEmpresa
'N�o existe nenhum Lote Pendente da Filial %i.
Public Const ERRO_VALOR_NAO_PREENCHIDO = 461
'O Valor do Rateio n�o Preenchido
Public Const ERRO_LEITURA_TRANSACAOCTB1 = 462 'Parametros sSiglaModulo, sTransacao
'Ocorreu um erro na leitura da tabela TransacaoCTB. Sigla do Modulo = %s, Transa��o = %s.
Public Const ERRO_DATA_CONTABIL_NAO_PREENCHIDA = 463 'Sem Parametro
'Data Cont�bil n�o preenchida.
Public Const ERRO_DOCUMENTO_CONTABIL_NAO_PREENCHIDO = 464 'Sem Parametro
'O N�mero do Documento Cont�bil n�o foi preenchido.
Public Const ERRO_VALOR_LANCAMENTO_CONTABIL_NAO_PREENCHIDO = 465 'Sem Parametro
'O Valor do Lan�amento Cont�bil n�o foi preenchido.
Public Const ERRO_LANCAMENTOS_CONTABILIZADOS = 466 'Sem parametro
'Aten��o. J� existem lan�amentos atualizados para o documento em quest�o.
Public Const ERRO_ALTERACAO_LANCAMENTO_AGLUTINADO = 467 'Parametros iFilialEmpresa, sOrigem, iexercicio, iperiodo, lDoc, iSeq
'Ocorreu um erro na altera��o do lan�amento aglutinado. Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Doc = %l, Seq = %i.
Public Const ERRO_LEITURA_LANCAMENTOS5 = 468 'Parametros FilialEmpresa, Origem, Data
'Erro na leitura da tabela de Lan�amentos. Filial = %i, Origem = %s, Data = %s.
Public Const ERRO_LEITURA_LOTEPENDENTE4 = 469 'Parametros Filial, origem, exercicio, periodo
'O Lote n�o est� cadastrado como lote pendente ou est� em processo de atualiza��o - Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i
Public Const ERRO_FORMULA_CONTA_NAO_PREENCHIDA = 470 'Sem Parametros
'A f�rmula da conta n�o foi preenchida.
Public Const ERRO_FORMULA_DEBCRE_NAO_PREENCHIDA = 471 'Sem Parametros
'As f�rmulas de d�bito e cr�dito n�o foram preenchidas.
Public Const ERRO_FORMULA_PRODUTO_NAO_PREENCHIDA = 472 'Sem Parametros
'A f�rmula de produton�o foi preenchida.
Public Const ERRO_QUANTIDADE_PRODUTO_DESBALANCEADO = 473 'Parametro sProduto
'Os lan�amentos de custo envolvendo o produto %s est�o desbalanceados.
Public Const ERRO_LANC_CUSTO_CONTA_NAO_INFORMADA = 474 'Sem parametro
'Ocorreu um erro na gera��o dos lan�amentos de custo. A conta de um dos lan�amentos n�o foi preenchida.
Public Const ERRO_LANC_CUSTO_QUANT_NAO_PREENCHIDA = 475 'Sem parametros
'Ocorreu um erro na gera��o dos lan�amentos de custo. Os campos de d�bito e cr�dito de um dos lan�amentos n�o foram preenchidos.
Public Const ERRO_LANC_CUSTO_DEBCRE_PREENCHIDOS = 476 'Sem parametros
'Ocorreu um erro na gera��o dos lan�amentos de custo. Os campos de d�bito e cr�dito de um dos lan�amentos est�o ambos preenchidos.
Public Const ERRO_LEITURA_LANCAMENTOS6 = 477 'Parametros iFilialEmpresa, sOrigem, iexercicio, iperiodo, lDoc, iSeq
'Erro na leitura da tabela de Lan�amentos Cont�beis. Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Doc = %l, Seq = %i.
Public Const ERRO_LOCK_LANCAMENTOS = 478 'Parametros iFilialEmpresa, sOrigem, iexercicio, iperiodo, lDoc, iSeq
'Ocorreu um erro na tentativa de fazer um "lock" em um dos lan�amentos cont�beis. Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Doc = %l, Seq = %i.
Public Const ERRO_UNLOCK_LANCAMENTOS = 479 'Parametros iFilialEmpresa, sOrigem, iexercicio, iperiodo, lDoc, iSeq
'Ocorreu um erro na tentativa de liberar um "lock" em um dos lan�amentos cont�beis. Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Doc = %l, Seq = %i.
Public Const ERRO_PRODUTOSFILIAL_INEXISTENTE = 480   'Parametro: sProduto, iFilial
'O Produto %s da Filial %i n�o est� cadastrado.
Public Const ERRO_MNEMONICO_INEXISTENTE = 481   'Parametro: sMnemonico
'O Campo %s n�o est� cadastrado.
Public Const ERRO_LEITURA_SLDMESEST = 482 'Parametros  iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque (SldMesEst). Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_SLDMESEST_INEXISTENTE = 483 'Parametros  iAno, iFilialEmpresa, sProduto
'N�o existe registro de saldos mensais de estoque (SldMesEst) com os dados a seguir. Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_SLDMESEST_CUSTORPRODUCAO_ZERADO = 484 'Parametros  iAno, iFilialEmpresa, sProduto, iMes
'N�o � poss�vel processar este lote j� que o custo real de produ��o para o produto especificado n�o foi digitado. Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s, Mes=%i.
Public Const ERRO_SLDMESEST_CUSTOMRPRODUCAO_ZERADO = 485 'Parametros  iAno, iFilialEmpresa, sProduto, iMes
'N�o � poss�vel processar este lote j� que o custo m�dio real de produ��o para o produto especificado n�o foi calculado. Existe um programa que calcula estes valores. Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s, Mes=%i.
Public Const ERRO_LEITURA_LANPREPENDENTE = 486 'Sem parametros
'Erro na leitura da tabela de Lan�amentos Pr�-Pendentes.
Public Const ERRO_EXCLUSAO_LANPREPENDENTE = 487 'Sem parametros
'Ocorreu um erro na exclus�o de um lan�amento pr�-pendente.
Public Const ERRO_LOCK_SLDMESEST = 488 'Parametros iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos mensais de estoque (SldMesEst). Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_LANC_CONTA_NAO_INFORMADA = 489 'Sem parametro
'Ocorreu um erro na gera��o dos lan�amentos cont�beis. A conta de um dos lan�amentos n�o foi preenchida.
Public Const ERRO_LANC_QUANT_NAO_PREENCHIDA = 490 'Sem parametros
'Ocorreu um erro na gera��o dos lan�amentos cont�beis. Os campos de d�bito e cr�dito de um dos lan�amentos n�o foram preenchidos.
Public Const ERRO_LANC_DEBCRE_PREENCHIDOS = 491 'Sem parametros
'Ocorreu um erro na gera��o dos lan�amentos cont�beis. Os campos de d�bito e cr�dito de um dos lan�amentos est�o ambos preenchidos.
Public Const ERRO_LEITURA_RATEIOON2 = 492 'Sem Parametros
'Erro de leitura da tabela de Rateios On-Line.
Public Const ERRO_LEITURA_RATEIOOFF1 = 493 'Sem Parametro
'Erro de leitura da tabela de Rateios Off-Line.
Public Const ERRO_RATEIOOFF_CODIGO_NAO_PREENCHIDO = 494
'O C�digo do Rateio n�o foi informado.
Public Const ERRO_RATEIOOFF_BATCH = 495 'Parametro lCodigo
'Ocorreu um erro no processamento do Rateio %l.
Public Const ERRO_ATUALIZACAO_BATCH = 496 'Parametros iFilialEmpresa, sOrigem, iExercicio, iPeriodo, iLote
'Ocorreu um erro na Atualiza��o de um Lote. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Public Const ERRO_RATEIOOFF_NAO_CADASTRADO1 = 497 'Parametro lCodigo
'O Rateio %l n�o est� cadastrado.
Public Const ERRO_TIPO_RATEIOOFF_INVALIDO = 498 'Parametro iTipo
'O Tipo de Rateio Offline %i � inv�lido.
Public Const ERRO_LEITURA_LANCAMENTOS7 = 499 'Parametros Filial, Origem, exercicio, periodo, lote
'Erro na leitura dos Lan�amentos Cont�beis do lote que possui chave Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Lote= %i.
Public Const ERRO_LEITURA_LANCAMENTOS8 = 500 'Parametros iFilialEmpresa, sOrigem, iexercicio, iperiodo, lDoc
'Erro na leitura da tabela de Lan�amentos Cont�beis do Documento que possui chave Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Doc = %l.
Public Const ERRO_LEITURA_CPCONFIG2 = 501 'Par�metros: %s chave %d FilialEmpresa
'Erro na leitura da tabela CPConfig. Codigo = %s Filial = %i.
Public Const ERRO_CPCONFIG_INEXISTENTE = 502 'Par�metros: %s chave %d FilialEmpresa
'N�o foi encontrado registro em CPConfig. Codigo = %s Filial = %i.
Public Const ERRO_ATUALIZACAO_CPCONFIG = 503 'Par�metros: %s chave %d FilialEmpresa
'Erro na grava��o da tabela CPConfig. Codigo = %s Filial = %i.
Public Const ERRO_LEITURA_CPCONFIG = 504 'Sem par�metros
'Erro na leitura da tabela CPConfig.
Public Const ERRO_LEITURA_CRCONFIG2 = 505 'Par�metros: %s chave %d FilialEmpresa
'Erro na leitura da tabela CRConfig. Codigo = %s Filial = %i.
Public Const ERRO_CRCONFIG_INEXISTENTE = 506 'Par�metros: %s chave %d FilialEmpresa
'N�o foi encontrado registro em CRConfig. Codigo = %s Filial = %i.
Public Const ERRO_ATUALIZACAO_CRCONFIG = 507 'Par�metros: %s chave %d FilialEmpresa
'Erro na grava��o da tabela CRConfig. Codigo = %s Filial = %i.
Public Const ERRO_LEITURA_CRCONFIG = 508 'Sem par�metros
'Erro na leitura da tabela CRConfig.
Public Const ERRO_LEITURA_TESCONFIG2 = 509 'Par�metros: %s chave %d FilialEmpresa
'Erro na leitura da tabela TESConfig. Codigo = %s Filial = %i.
Public Const ERRO_TESCONFIG_INEXISTENTE = 510 'Par�metros: %s chave %d FilialEmpresa
'N�o foi encontrado registro em TESConfig. Codigo = %s Filial = %i.
Public Const ERRO_ATUALIZACAO_TESCONFIG = 511 'Par�metros: %s chave %d FilialEmpresa
'Erro na grava��o da tabela TESConfig. Codigo = %s Filial = %i.
Public Const ERRO_LEITURA_TESCONFIG = 512 'Sem par�metros
'Erro na leitura da tabela TESConfig.
Public Const ERRO_CTBCONFIG_CHV_NAO_ENC = 513
'Chave %s n�o encontrada na tabela CTBConfig.
Public Const ERRO_CTBCONFIG_ATUALIZAR_CHV = 514
'Erro na atualiza��o da tabela CTBConfig para a chave %s.
Public Const ERRO_LEITURA_ARQ_CONF_REL_DRE = 515
'Erro na leitura do arquivo de configura��o do relat�rio.
Public Const ERRO_LEITURA_MVPERCTA_PLANOCONTA = 516
'Erro na leitura das tabelas MvPerCta e/ou PlanoConta.
Public Const ERRO_GRAVACAO_ARQ_TEMP_REL_DRE = 517
'Erro na grava��o de arquivo tempor�rio p/relat�rio (RelDRERes).
Public Const ERRO_CONFIG_DR = 518
'A configura��o do demonstrativo de resultados n�o foi feita ou apresenta problemas.
Public Const ERRO_LEITURA_ARQ_CONF_REL_DRP = 519
'Erro na leitura do arquivo de configura��o do relat�rio.
Public Const ERRO_GRAVACAO_ARQ_TEMP_REL_DRP = 520
'Erro na grava��o de arquivo tempor�rio p/relat�rio (RelDRPRes).
Public Const ERRO_RELATORIO_EXECUTANDO = 521
'J� existe um di�logo aberto para a execu��o deste relat�rio
Public Const ERRO_EXERCICIO1_VAZIO = 522
'O campo Exerc�cio1 tem que estar preenchido.
Public Const ERRO_EXERCICIO2_VAZIO = 523
'O campo Exerc�cio2 tem que estar preenchido.
Public Const ERRO_EXERCICIO1_MAIOR = 524
'O Exercicio 1 n�o pode ser maior que o Exercicio 2.
Public Const ERRO_LEITURA_RELDRE = 525 'Sem Parametro
'Erro na leitura da tabela RelDRE.
Public Const ERRO_RELDRE_VAZIA = 526 'Sem Parametro
'Erro a Tabela RelDre est� vazia
Public Const ERRO_DOCUMENTO_INICIAL_MAIOR = 527 'Sem Parametros
'O Documento inicial n�o pode ser maior do que o documento final.
Public Const ERRO_ORIGEM_INICIAL_MAIOR = 528 'Sem Parametros
'A Origem inicial n�o pode ser maior do que a Origem final.
Public Const ERRO_CONTA_PATRIMONIOLIQUIDO_VAZIO = 529 'Sem Parametros
'A Conta de Patrim�nio L�quido tem que estar preenchida.
Public Const ERRO_NO_NAO_SELECIONADO_INSERCAO_FILHO = 530 'Sem parametro.
'Escolha um elemento da �rvore antes de tentar inserir um filho.
Public Const ERRO_FORMULA_NAO_PREENCHIDA = 531 'Parametro: iLinha
'A F�rmula da Linha %i n�o foi preenchida.
Public Const ERRO_OPERADOR_NAO_PREENCHIDO = 532 'Parametro: iLinha
'O Operador "Soma/Subtrai" da Linha %i n�o foi preenchido.
Public Const ERRO_FORMULA_INVALIDA = 533 'Parametro: sFormula, iLinha
'A F�rmula %s utilizada na Linha %i n�o � v�lida.
Public Const ERRO_FORMULA_INVALIDA1 = 534 'Parametro: sFormula
'A F�rmula %s n�o � v�lida.
Public Const ERRO_CONTA_INICIO_NAO_PREENCHIDA = 535 'Parametro: iLinha
'A Conta In�cio da Linha %i n�o foi preenchida.
Public Const ERRO_CONTA_FIM_NAO_PREENCHIDA = 536 'Parametro: iLinha
'A Conta Fim da Linha %i n�o foi preenchida.
Public Const ERRO_LEITURA_RELDRECONTA = 537 'Par�metro Modelo
'Ocorreu um erro na leitura da tabela RelDREConta com o modelo %s .
Public Const ERRO_EXCLUSAO_RELDREFORMULA = 538 'Par�metro Modelo
'Erro na exclus�o do modelo %s da tabela RelDREFormula.
Public Const ERRO_EXCLUSAO_RELDRECONTA = 539 'Par�metro Modelo
'Erro na exclus�o do modelo %s da tabela RelDREConta.
Public Const ERRO_EXCLUSAO_RELDRE = 540 'Par�metro Modelo
'Erro na exclus�o do modelo %s da tabela RelDRE.
Public Const ERRO_MODELO_NAO_INFORMADO = 541
'O modelo precisa ser informado.
Public Const ERRO_NO_NAO_SELECIONADO_REMOVER = 542 'Sem parametro.
'Escolha um elemento da �rvore antes de tentar remover.
Public Const ERRO_LEITURA_RELDRE_MODELO = 543 'Parametro: sModelo
'Ocorreu um erro na leitura do modelo %s tabela RelDRE.
Public Const ERRO_LEITURA_RELDREFORMULA = 544 'Par�metro sModelo
'Ocorreu um erro na leitura da tabela RelDREFormula com o modelo %s .
Public Const ERRO_INCLUSAO_RELDRE = 545 'Par�metro Modelo
'Ocorreu um erro na inclus�o do modelo %s na tabela RelDRE.
Public Const ERRO_INCLUSAO_RELDRECONTA = 546 'Par�metro Modelo
'Ocorreu um erro na inclus�o do modelo %s na tabela RelDREConta.
Public Const ERRO_INCLUSAO_RELDREFORMULA = 547 'Par�metro Modelo
'Ocorreu um erro na inclus�o do modelo %s na tabela RelDREFormula.
Public Const ERRO_NO_NAO_SELECIONADO_MOV_ARV = 548 'Sem parametro.
'Selecione um elemento da �rvore antes de tentar moviment�-lo.
Public Const ERRO_NO_SELECIONADO_NAO_MOV_ACIMA = 549 'Sem parametro.
'O elemento da �rvore selecionado n�o pode ser movimentado para cima.
Public Const ERRO_NO_UTILIZA_NO_EM_FORMULA = 550 'Parametros: sTitulo, sTitulo1
'O elemento da �rvore %s utiliza em sua f�rmula o elemento %s e portanto este n�o pode ser movido.
Public Const ERRO_NO_SELECIONADO_NAO_MOV_ABAIXO = 551 'Sem parametro.
'O elemento da �rvore selecionado n�o pode ser movimentado para baixo.
Public Const ERRO_LEITURA_RATEIOOFF2 = 552 'Parametros conta e ccl.
'Erro de leitura da tabela de Rateios Off-Line para a conta %s e centro de custo/lucro %s.
Public Const ERRO_LEITURA_MNEMONICOCTBVALOR = 553 'Parametro sMnemonico
'Ocorreu um erro na leitura da tabela MnemonicoCTBValor. Mnemonico = %s.
Public Const ERRO_DOCAUTO_NAO_CADASTRADO2 = 554 'Parametro: lDoc
'O  Documento %l n�o foi encontrado.
Public Const ERRO_INCLUSAO_LANPREPENDENTEBAIXADO = 555 'Parametros: iFilialEmpresa, sOrigem, iExercicio, iPeriodoLan, lDoc, iSeq
'Ocorreu um erro na inclus�o de um lan�amento na tabela de lan�amentos pre-pendentes baixados. FilialEmpresa = %i, Origem =%s, Exerc�cio = %i, Periodo = %i, Documento = %l, Sequencial = %i.
Public Const ERRO_EXCLUSAO_LANPREPENDENTE1 = 556 'Parametros: iFilialEmpresa, sOrigem, iExercicio, iPeriodoLan, lDoc, iSeq
'Ocorreu um erro na exclus�o de um lan�amento da tabela de lan�amentos pre-pendentes. FilialEmpresa = %i, Origem =%s, Exerc�cio = %i, Periodo = %i, Documento = %l, Sequencial = %i.
Public Const ERRO_CONTA_NAO_VISIVEL_MODULO = 557 'Sem Parametros
'A conta %s n�o est� vis�vel para o m�dulo em quest�o.
Public Const ERRO_LEITURA_TABELASDEPRECOITENS1 = 558 'Parametros: iCodTabela
'Erro na leitura da tabela TabelasDePrecoItens com C�digo da Tabela %i.
Public Const ERRO_ATUALIZACAO_EXERCICIOORIGEM_BATCH = 559 'sem parametros
'Ocorreu um erro na altera��o de um registro da tabela ExercicioOrigem.
Public Const ERRO_LEITURA_CTBCONFIG2 = 560 'Parametros: Codigo, FilialEmpresa
'Erro na leitura da tabela CTBConfig. Codigo = %s, Filial = %i.
Public Const ERRO_INSERCAO_CTBCONFIG = 561 'Par�metros: Codigo, FilialEmpresa
'Ocorreu um erro na inser��o de um registro na tabela CTBConfig. Codigo = %s, Filial = %i.
Public Const ERRO_ATUALIZACAO_CTBCONFIG = 562 'Par�metros: Codigo, FilialEmpresa
'Erro na grava��o da tabela CTBConfig. Codigo = %s, Filial = %i.
Public Const ERRO_SEGMENTO_CONTA_INVALIDO = 563 'Parametro sCodigo
'Esperava o segmento conta e est� tentando gravar o segmento %s.
Public Const ERRO_SEGMENTO_CCL_INVALIDO = 564 'Parametro sCodigo
'Esperava o segmento centro de custo e est� tentando gravar o segmento %s.
Public Const ERRO_ALTERACAO_SEGMENTO = 565 'Parametros Codigo, Nivel
'Ocorreu um erro na altera��o de um registro da tabela Segmento. Codigo = %s e Nivel = %i.
Public Const ERRO_LANPENDENTE_TRANSACAO_NUMINTDOC = 566 'Parametros iTransacaoLan, lNumIntDocLan, iTransacao, lNumIntDoc
'Foi encontrado um lan�amento cont�bil pendente que n�o pertence ao documento em quest�o. Documento Encontrado: Transa��o = %i , NumIntDoc = %l; Documento em quest�o: Transa��o = %i , NumIntDoc = %l
Public Const ERRO_LANPREPENDENTE_JA_CADASTRADO = 567 'Parametros: iFilialEmpresa, sOrigem , iExercicio, iPeriodoLan , lDoc
'Foi encontrado um documento cadastrado com este mesmo n�mero. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Documento = %l.
Public Const ERRO_ALTERACAO_CONFIGURACAO = 568 'Sem Parametros
'Erro na tentativa de alterar a configura��o da contabilidade.
Public Const ERRO_LEITURA_PADRAOCONTAB_MODPADRAO = 569 'parametros: nome da transacao e sigla do modulo
'Erro na leitura do modelo padr�o para a transa��o %s do m�dulo %s
Public Const ERRO_PADRAOCONTAB_SEMMODPADRAO = 570 'parametros: nome da transacao e sigla do modulo
'N�o h� modelo padr�o para a transa��o %s do m�dulo %s
Public Const ERRO_LOTE_JA_CADASTRADO_OUTRO_PERIODO = 571 'Parametros Filial, origem, exercicio, periodo, lote
'J� existe um lote cadastrado com o mesmo n�mero em outro periodo. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Public Const ERRO_DOC_PENDENTE_OUTRO_PERIODO = 572 'Parametros Filial, Origem, Exercicio, Periodo, Doc
'J� existe um documento pendente com o mesmo n�mero n�mero em outro periodo. Filial = %i, Origem = %s, Exercicio = %s, Periodo = %s, Documento = %l.
Public Const ERRO_CONTRAPARTIDA_NAO_MESMO_LANCAMENTO = 573
'O Lan�amento de Contra Partida tem que ser diferente do lan�amento sendo editado.
Public Const ERRO_CONTRAPARTIDA_LANCAMENTO_INEXISTENTE = 574
'O Lan�amento de Contra Partida tem que ser um lan�amento existente.
Public Const ERRO_LANCAMENTO_CONTRA_PARTIDA_INEXISTENTE = 575 'Parametro iSeq
'O Lan�amento a que a contra partida se refere n�o existe. Sequencial = %i.
Public Const ERRO_LANCAMENTO_CONTRA_PARTIDA_VALOR = 576 'Parametros iSeq, dValorLancamento, dValorContraPartida
'O Valor do Lan�amento n�o � igual ao de sua contra partida. Sequencial = %i, Valor do Lan�amento = %d, Total da Contra-Partida = %d.
Public Const ERRO_HISTORICO_PARAM = 577 'Sem Parametros
'O Hist�rico cont�m parametro(s) que ainda n�o foi(foram) substituido(s).
Public Const ERRO_EXCLUSAO_MNEMONICOCTBVALOR = 578 'Sem Par�metros
'Erro na Exclus�o da Tabela MnemonicoCTBValor
Public Const ERRO_INSERCAO_MNEMONICOCTBVALOR = 579 'Sem Par�metros
'Erro na inser��o de dados na Tabela MnemonicoCTBValor
Public Const ERRO_RATEIOS_NAO_INFORMADOS = 580
'N�o foi informado nenhum Rateio para apura��o.
Public Const ERRO_PERIODOFINAL_MENOR = 581
'O Periodo Final n�o pode ser menor do que o Periodo Inicial.
Public Const ERRO_PERIODOFINAL_MAIOR = 582
'O Periodo Final n�o pode ser maior que o periodo da data de contabiliza��o.
Public Const ERRO_RATEIOS_INEXISTENTES = 583
'N�o existem Rateios dispon�veis para apura��o.
Public Const ERRO_CODIGO_NAO_DIGITADO = 584
'O C�digo do Rateio n�o foi informado.
Public Const ERRO_RELDRE_MODELO_NAO_CADASTRADO = 585 'Sem par�metros: sModelo, iCodigo
'O Modelo %s com C�digo %i do Demonstrativo de Resultado do Exerc�cio n�o est� cadastrado no Banco de Dados.
Public Const ERRO_LOCK_RELDRE = 586 'Par�metros: sModelo, iCodigo
'N�o foi poss�vel fazer o "Lock" na tabela RelDRE com Modelo %s e C�digo %i.
Public Const ERRO_ATUALIZACAO_RELDRE = 587  'Par�metros: sModelo, iCodigo
'Ocorreu um erro na atualiza��o da tabela RelDRE. Modelo %s, C�digo %i.
Public Const ERRO_SEGMENTO_VAZIO = 588 'Sem parametros
'Preencha os segmentos de conta e centro de custo com pelo menos 1 segmento.
Public Const ERRO_SEGMENTO_PRODUTO_MAIOR_PERMITIDO = 589 'Parametros tamanho do segmento, tamanho total permitido
'O tamanho do segmento de produto %i ultrapassou o tamanho total permitido %i.
Public Const ERRO_DOC_PENDENTE_OUTRO_LOTE = 590 'Parametros iLote
'Este documento j� est� cadastrado no lote %i e n�o � poss�vel alterar o lote.
Public Const ERRO_ORIGEM_NAO_PREENCHIDA1 = 591 'Sem par�metros
'O preenchimento de Origem � obrigat�rio.
Public Const ERRO_GRID_DESCONTO_TIPODESCONTO_NAO_PRENCHIDO = 592 'Parametro iLinha
'O campo Tipo Desconto da Linha %i do Grid de Desconto n�o foi preechido.
Public Const ERRO_GRID_DESCONTO_DIAS_NAO_PRENCHIDO = 593 'Parametro iLinha
'O campo Dias da Linha %i do Grid de Desconto n�o foi preechido.
Public Const ERRO_GRID_DESCONTO_PERCENTUAL_NAO_PRENCHIDO = 594 'Parametro iLinha
'O campo Percentual da Linha %i do Grid de Desconto n�o foi preechido.
Public Const ERRO_GRID_DESCONTO_NAO_ORDEM_DECRESCENTE = 595
'Os Campos de Dias e Percentual tem que estar em Ordem Decrescente no Grid de Descontos.
Public Const ERRO_LOTEPENDENTE_NAO_CADASTRADO = 596 'Parametros: iFilialEmpresa, sOrigem, iExercicio, iPeriodo, iLote
'Lote Pendente n�o cadastrado.
Public Const ERRO_CONTA_JA_UTILIZADA_GRID_CONTAS = 597 'Parametro sconta
'A conta %s j� foi utilizada no Grid de Contas.
Public Const ERRO_CONTA_JA_UTILIZADA_GRID_RATEIOS = 598 'Parametro sconta
'A conta %s j� foi utilizada no Grid de Rateios.
Public Const ERRO_CONTA_JA_UTILIZADA_CONTA_CREDITO = 599 'Parametro sconta
'A conta %s j� foi utilizada no Grid de Rateios.
Public Const ERRO_CONTAINICIO_GRIDCONTAS_NAO_INFORMADA = 600 'Parametro iLinha
'A Conta In�cio, localizada na linha %i do grid de Contas, n�o foi informada.
Public Const ERRO_CONTAFIM_GRIDCONTAS_NAO_INFORMADA = 601 'Parametro iLinha
'A Conta Fim, localizada na linha %i do grid de Contas, n�o foi informada.
Public Const ERRO_CONTAFIM_MENOR_CONTAINICIO = 602 'Parametro iLinha, sContaFim, sContaInicio.
'Na linha %i do Grid de Contas, a Conta Fim = %s � menor do que a Conta In�cio = %s.
Public Const ERRO_CONTA_GRID_NAO_PREENCHIDA = 603 'Parametro iLinha.
'A Conta da Linha %i do Grid n�o foi preenchida.
Public Const ERRO_PROXNUM_DATA_NAO_PREENCHIDA = 604 'Sem Parametro
'Para conseguir o pr�ximo n�mero de documento, a data tem que estar preenchida.
Public Const ERRO_DATA_ULT_PERIODO_FORA_EXERCICIO = 605 'Parametros Data Inicial do Periodo , Data Inicial do Exercicio e Data Final do Exercicio
'A Data Inicial do ultimo periodo periodo %s n�o est� dentro da faixa abrangida pelo exerc�cio. Data Inicial = %s e Data Final = %s.
Public Const ERRO_EXERCICIOFILIAL_NAO_CADASTRADO = 606 'Parametros iExercicio, iFilialEmpresa.
'O Exerc�cio %i da Filial %i n�o est� cadastrado.
Public Const ERRO_CODIGO_HISTPADRAO_ZERADO = 607 'Sem parametro
'O codigo do hist�rico padr�o tem que ser um n�mero maior que zero.
Public Const ERRO_CODIGO_CATEGORIA_ZERADO = 608 'Sem parametro
'O codigo da categoria tem que ser um n�mero maior que zero.
Public Const ERRO_LEITURA_SALDOINICIALCONTA2 = 609 'Sem Parametros.
'Ocorreu um erro na leitura da tabela SaldoInicialConta.
Public Const ERRO_LEITURA_SALDOINICIALCONTACCL5 = 610 'Sem Parametros
'Ocorreu um erro na leitura da tabela SaldoInicialContaCcl.
Public Const ERRO_INSERCAO_RATEIOOFFCONTAS = 611
'Ocorreu um erro ao tentar inserir dados na tabela RateioOffContas
Public Const ERRO_EXCLUSAO_RATEIOOFFCONTAS = 612 'Sem Parametros
'Ocorreu um erro na Exclus�o de registros da tabela RateioOffContas.
Public Const ERRO_LEITURA_RATEIOOFFCONTAS = 613 'Parametro lCodigo
'Ocorreu um erro de leitura na tabela de RateioOffContas. Codigo do Rateio = %l.
Public Const ERRO_CONTA_SEG_NUM_CARACTER_INVALIDO = 614 'Sem parametro
'Os segmentos num�ricos da Conta s� podem conter n�meros. Ex: 1.-1.1 est� errado. 1.1.1 est� correto.
Public Const ERRO_CCL_SEG_NUM_CARACTER_INVALIDO = 615 'Sem parametro
'Os segmentos num�ricos do Centro de Custo s� podem conter n�meros. Ex: 1.-1.1 est� errado. 1.1.1 est� correto.
Public Const ERRO_ORIGEM_DIFERENTE_CTB = 616
'N�o pode gravar lan�amentos de outros m�dulos. S� � permitido gerar lan�amentos para a Contabilidade.
Public Const ERRO_ORIGEM_DIFERENTE = 617
'N�o pode gravar, alterar e excluir lotes de outros m�dulos.
Public Const ERRO_AUSENCIA_LANCAMENTOS_PADRAOCONTAB = 618 'Sem parametros
'N�o h� lan�amentos em nenhum dos grids.
Public Const ERRO_INDICE_MAIOR_LINHAS_GRID = 619  'Parametros sMnemonico, iIndice, iLinhasGrid
'O Indice do mnemonico %s ultrapassa o n�mero de linhas do grid. Indice = %i, Linhas do Grid = %i.
Public Const ERRO_DOC_NAO_BALANCEADO = 620 'Parametro: sModelo
'O modelo de contabiliza��o, %s, usado para gerar os lan�amentos cont�beis n�o est�o gerando um documento balanceado, ou seja, o total dos cr�ditos n�o corresponde aos d�bitos.
Public Const ERRO_LINCOL_ZERO_CONTA = 621 'Sem Parametros
'Para que uma c�lula seja do tipo conta ela n�o pode estar posicionada na linha zero ou coluna 0.
Public Const ERRO_LINCOL_ZERO_FORMULA = 622 'Sem Parametros
'Para que uma c�lula seja do tipo f�rmula ela n�o pode estar posicionada na linha zero ou coluna 0.
Public Const ERRO_CEL_UTILIZA_CEL_EM_FORMULA = 623 'Parametros: iLinhaUtiliza, iColunaUtiliza, iLinhaUsada, iColunaUsada
'A c�lula posicionada na linha/coluna (%i, %i) utiliza a c�lula posicionada na linha/coluna (%i, %i).
Public Const ERRO_LEITURA_RELDMPL_MODELO = 624 'Parametro: sModelo
'Ocorreu um erro na leitura do modelo %s na tabela RelDMPL.
Public Const ERRO_LEITURA_RELDMPLFORMULA = 625 'Par�metro sModelo
'Ocorreu um erro na leitura do modelo %s na tabela RelDMPLFormula.
Public Const ERRO_LEITURA_RELDMPLCONTA = 626 'Par�metro sModelo
'Ocorreu um erro na leitura do modelo %s na tabela RelDMPLConta.
Public Const ERRO_EXCLUSAO_RELDMPL = 627 'Par�metro Modelo
'Ocorreu um erro na exclus�o do modelo %s da tabela RelDMPL.
Public Const ERRO_EXCLUSAO_RELDMPLFORMULA = 628 'Par�metro Modelo
'Ocorreu um erro na exclus�o do modelo %s da tabela RelDMPLFormula.
Public Const ERRO_EXCLUSAO_RELDMPLCONTA = 629 'Par�metro Modelo
'Ocorreu um erro na exclus�o do modelo %s da tabela RelDMPLConta.
Public Const ERRO_INCLUSAO_RELDMPL = 630 'Par�metro Modelo
'Ocorreu um erro na inclus�o do modelo %s na tabela RelDMPL.
Public Const ERRO_INCLUSAO_RELDMPLCONTA = 631 'Par�metro Modelo
'Ocorreu um erro na inclus�o do modelo %s na tabela RelDMPLConta.
Public Const ERRO_INCLUSAO_RELDMPLFORMULA = 632 'Par�metro Modelo
'Ocorreu um erro na inclus�o do modelo %s na tabela RelDMPLFormula.
Public Const ERRO_RELDMPL_MODELO_NAO_CADASTRADO = 633 'Par�metros: sModelo, iLinha, iColuna
'O Modelo %s referente � linha/coluna (%i,%i) do Demonstrativo de Muta��o de Patrimonio n�o est� cadastrado no Banco de Dados.
Public Const ERRO_LOCK_RELDMPL = 634 'Par�metros: sModelo, iLinha, iColuna
'N�o foi poss�vel fazer o "Lock" na tabela RelDMPL do registro referente ao Modelo %s , Linha = %i e Coluna = %i.
Public Const ERRO_ATUALIZACAO_RELDMPL = 635  'Par�metros: sModelo, iLinha, iColuna
'Ocorreu um erro na atualiza��o da tabela RelDMPL. Modelo = %s, Linha = %i e Coluna = %i.
Public Const ERRO_CCL_SINTETICA_USADA_EM_MOVESTOQUE = 636 'Parametro: sCcl
'O Centro de Custo/Lucro %s est� sendo utilizado em pelo menos um movimento de estoque.
Public Const ERRO_LIMITE_CCL_VLIGHT = 637 'Parametros : iNumeroMaxCcl
'N�mero m�ximo de Centro de Custo/Lucro anal�tico desta vers�o � %i.
Public Const ERRO_AUSENCIA_CONTAS_CATEGORIA_APURACAO = 638
'N�o foram encontradas contas de grupos de categoria para apura��o.
Public Const ERRO_RATEIOOFF_PROCESSADO = 639 'Parametro: lCodigo
'O rateio offline de c�digo = %l foi processado.
Public Const ERRO_RATEIOOFF_CCL_ZERADO = 640 'Parametro: lCodigo
'O rateio offline de c�digo = %l n�o gerou lan�amento pois o total do centro de custo para as contas em quest�o est� zerado.
Public Const ERRO_NAO_HA_LOTE_PENDENTE1 = 641 'Parametros sOrigem, iExercicio, iPeriodo
'N�o h� lote pendente dispon�vel para o m�dulo %s no Periodo %i do Exercicio %i.
Public Const ERRO_LEITURA_USUARIOLOTE = 642 'Parametro sCodUsuario, sOrigem
'Erro na leitura dos dados do Usuario %s, Origem = %s na tabela UsuarioLote.
Public Const ERRO_ATUALIZACAO_USUARIOLOTE = 643 'Parametro sCodUsuario, sOrigem
'Ocorreu um erro na atualiza��o da tabela UsuarioLote. Usu�rio = %s, Origem = %s.
Public Const ERRO_INSERCAO_USUARIOLOTE = 644 'Parametro sCodUsuario, sOrigem
'Ocorreu um erro na inser��o de um registro na tabela UsuarioLote. Usu�rio = %s, Origem = %s.
Public Const ERRO_LEITURA_LOTEPENDENTE5 = 645 'Parametros Filial, origem, exercicio, periodo
'Ocorreu um erro na leitura da tabela de lotes pendentes. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i.
Public Const ERRO_RECALCULO_AUTOMATICO_SEM_MODELO = 646 'Sem Parametros
'O rec�lculo autom�tico s� pode ser marcado se o modelo estiver preenchido.
Public Const ERRO_LEITURA_SLDMESEST11 = 647 'Parametros  iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque (SldMesEst1). Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_LOCK_SLDMESEST1 = 648 'Parametros iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos mensais de estoque (SldMesEst1). Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_LEITURA_SLDMESEST21 = 649 'Parametros  iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque (SldMesEst2). Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_LOCK_SLDMESEST2 = 650 'Parametros iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos mensais de estoque (SldMesEst2). Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_EXCLUSAO_SLDMESEST11 = 651 'Parametros sProduto, iFilialEmpresa
'Ocorreu um erro na exclus�o de registro da tabela de saldos mensais de estoque (SldMesEst1). Codigo do Produto = %s, Filial = %i.
Public Const ERRO_MNEMONICO_ESCANINHOCUSTOCONSIG_INEX = 652 'Parametros sModulo, iTransacao
'O mnemonico Escaninho_Custo_Consig n�o est� cadastrado para a transa��o em quest�o. Modulo = %s, Transa��o = %i.

'??? jones 28/10

Public Const ERRO_SUBTIPOCONTABIL_NAO_ENCONTRADO = 653 'Par�mtero: objTipoDocInfo.iCodigo, objTipoDocInfo.sDescricao
'N�o foi encontrada transa��o na tabela TransacaoCTB correspondente ao subtipo selecionado.
Public Const ERRO_MNEMONICO_ESCANINHOCUSTOBENEF_INEX = 654 'Parametros sModulo, iTransacao
'O mnemonico Escaninho_Custo_Benef n�o est� cadastrado para a transa��o em quest�o. Modulo = %s, Transa��o = %i.
Public Const ERRO_LEITURA_MVDIACLI1 = 655 'parametro Filial, Cliente, FilialCliente, data
'Ocorreu um erro na leitura da tabela de Saldos Di�rios de Cliente. Filial=%i, Cliente=%l, Filial do Cliente = %i e Data=%s.
Public Const ERRO_INSERCAO_MVDIACLI = 656 'parametro Filial, Cliente, FilialCliente, data
'Ocorreu um erro na inser��o de um registro na tabela de Saldos Di�rios de Cliente. Filial=%i, Cliente=%l, Filial do Cliente = %i e Data=%s.
Public Const ERRO_ATUALIZACAO_MVDIACLI = 657 'parametro Filial, Cliente, FilialCliente, data
'Ocorreu um erro na atualiza��o de um registro na tabela de Saldos Di�rios de Cliente. Filial=%i, Cliente=%l, Filial do Cliente = %i e Data=%s.
Public Const ERRO_LEITURA_MVDIAFORN1 = 658 'parametro Filial, Fornecedor, FilialFornecedor, data
'Ocorreu um erro na leitura da tabela de Saldos Di�rios de Fornecedor. Filial=%i, Fornecedor=%l, Filial do Fornecedor = %i e Data=%s.
Public Const ERRO_INSERCAO_MVDIAFORN = 659 'parametro Filial, Fornecedor, FilialFornecedor, data
'Ocorreu um erro na inser��o de um registro na tabela de Saldos Di�rios de Fornecedor. Filial=%i, Fornecedor=%l, Filial do Fornecedor = %i e Data=%s.
Public Const ERRO_ATUALIZACAO_MVDIAFORN = 660 'parametro Filial, Fornecedor, FilialFornecedor, data
'Ocorreu um erro na atualiza��o de um registro na tabela de Saldos Di�rios de Cliente. Filial=%i, Fornecedor=%l, Filial do Fornecedor = %i e Data=%s.
Public Const ERRO_LEITURA_MVPERCLI = 661 'parametro Filial, Exercicio, Cliente, FilialCliente
'Ocorreu um erro na leitura da tabela de Saldos Mensais de Cliente. Filial=%i, Exercicio = %i, Cliente=%l, Filial do Cliente = %i.
Public Const ERRO_INSERCAO_MVPERCLI = 662 'parametro Filial, Exercicio, Cliente, FilialCliente
'Ocorreu um erro na inser��o de um registro na tabela de Saldos Mensais de Cliente. Filial=%i, Exercicio = %i, Cliente=%l, Filial do Cliente = %i.
Public Const ERRO_ATUALIZACAO_MVPERCLI = 663 'parametro Filial, Exercicio, Cliente, FilialCliente
'Ocorreu um erro na atualiza��o de um registro na tabela de Saldos Mensais de Cliente. Filial=%i, Exercicio = %i, Cliente=%l, Filial do Cliente = %i.
Public Const ERRO_LEITURA_MVPERFORN2 = 664 'parametro Filial, Exercicio, Fornecedor, FilialFornecedor
'Ocorreu um erro na leitura da tabela de Saldos Mensais de Fornecedor. Filial=%i, Exercicio = %i, Fornecedor=%l, Filial do Fornecedor = %i.
Public Const ERRO_INSERCAO_MVPERFORN = 665 'parametro Filial, Exercicio, Fornecedor, FilialFornecedor
'Ocorreu um erro na inser��o de um registro na tabela de Saldos Mensais de Fornecedor. Filial=%i, Exercicio = %i, Fornecedor=%l, Filial do Fornecedor = %i.
Public Const ERRO_ATUALIZACAO_MVPERFORN = 666 'parametro Filial, Exercicio, Fornecedor, FilialFornecedor
'Ocorreu um erro na atualiza��o de um registro na tabela de Saldos Mensais de Fornecedor. Filial=%i, Exercicio = %i, Fornecedor=%l, Filial do Fornecedor = %i.
Public Const ERRO_NAO_HA_EXERCICIO_CONTABIL_ABERTO = 667 'Sem Parametros
'Aten��o! N�o h� exerc�cio cont�bil aberto.
Public Const ERRO_PERIODO_NAO_ENCONTRADO_DATA = 668 'Parametro: Data
'N�o foi encontrado o periodo cont�bil que englobe esta data. Data = %s.
Public Const ERRO_LEITURA_LANCAMENTOS_PENDENTES1 = 669 'Parametros: iFilial, sOrigem, iExercicio, iPeriodo, lDoc, iLote
'Ocorreu um erro na leitura na tabela de Lan�amentos Pendentes. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Doc = %l, Lote = %i.
Public Const ERRO_ALTERACAO_LANPENDENTE = 670 'Parametros: iFilial, sOrigem, iExercicio, iPeriodo, lDoc, iSeq
'Ocorreu um erro na altera��o da tabela de Lan�amentos Pendentes. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Doc = %l, Sequencial = %i.
Public Const ERRO_NUM_LANC_TELA_MAIOR_BD = 671 'Sem Parametros
'O n�mero de lan�amentos da tela � maior do que o encontrado no banco de dados.
Public Const ERRO_NUM_LANC_TELA_MENOR_BD = 672 'Sem Parametros
'O n�mero de lan�amentos da tela � menor do que o encontrado no banco de dados.



''VEIO DE ERROS COM
Public Const ERRO_LEITURA_COMPRASCONFIG = 12005 'Parametro sCodigo
'Erro na leitura de %s na tabela de ComprasConfig.
Public Const ERRO_ATUALIZACAO_COMPRASCONFIG = 12006 'Parametro sCodigo
'Erro na atualizacao de %s na tabela de ComprasConfig.


''VEIO DE ERROS MAT
Public Const ERRO_LEITURA_TABELA_UNIDADESDEMEDIDA = 7328 'Parametros: iClasseUM, sSiglaUM
'Erro na Leitura da Unidade de Medida. Classe=%i e Sigla=%s.
Public Const ERRO_UNIDADE_MEDIDA_NAO_CADASTRADA = 7329 'Parametros: iClasseUM, sSiglaUM
'Unidade de Medida com Classe=%i e Sigla=%s n�o est� cadastrada no Banco de Dados.
Public Const ERRO_MODIFICACAO_UNIDADESDEMEDIDA = 7348 'Sem parametro
'Erro na modifica��o da tabela UnidadesDeMedida.
Public Const ERRO_ESTOQUEMES_INEXISTENTE = 7439 'Parametros iFilialEmpresa, iAno, iMes
'O M�s em quest�o n�o est� aberto. Filial Empresa = %i, Ano = %i, M�s = %i.
Public Const ERRO_LEITURA_ESTOQUEMES = 7441 'Parametros iFilialEmpresa, iAno, iMes
'Ocorreu um erro na leitura da tabela EstoqueMes. FilialEmpresa = %i, Ano = %i, Mes = %i.
Public Const ERRO_QUANTIDADE_NAO_PREENCHIDA = 7447 'Parametro iLinhaGrid
'A Quantidade do �tem %i do Grid n�o foi preenchida.
Public Const ERRO_PRODUTO_NAO_PREENCHIDO = 7523 'Sem par�metros
'O Produto deve estar preenchido.
Public Const ERRO_LOCK_TABELASDEPRECOITENS = 7573
'Erro na tentativa de fazer 'lock' na tabela TabelasDePrecoItens.
Public Const ERRO_EXCLUSAO_TABELASDEPRECOITENS = 7574
'Erro na exclus�o de registro na tabela TabelasDePrecoItens.
Public Const ERRO_ATUALIZACAO_TABELASDEPRECOITENS = 7585
'Erro na atualiza��o da tabela TabelasDePrecoItens
Public Const ERRO_INSERCAO_TABELASDEPRECOITENS = 7586 'Par�metros: iCodTabela, iCodProduto
'Erro na tentativa de inserir registro na tabela TabelasDePrecoItens. Com CodTabela = %i e CodProduto = %s.
Public Const ERRO_NOTA_FISCAL_NAO_CADASTRADA = 7600 'lNumIntDoc
'A Nota Fiscal com o N�mero Interno %l n�o est� cadastrada no Banco de Dados.
Public Const ERRO_LOCK_SERIE = 7601
'Erro na tentativar de fazer "lock" na tabela de S�ries.
Public Const ERRO_LEITURA_NFISCAL1 = 7603 'Par�metros: iTipoNFiscal, lFornecedor, iFilialForn, sSerie, lNumNotaFiscal
'Erro na leitura da tabela NFiscal na Nota Fiscal com Tipo = %i, Fornecedor = %l, Filial = %i, Serie = %s e N�mero = %l.
Public Const ERRO_LEITURA_NFISCALBAIXADA2 = 7636
'Erro na leitura da tabela NFiscalBaixadas
Public Const ERRO_LEITURA_NFISCALBAIXADA = 7641
'Erro na leitura da tabela NFiscalBaixada da Nota Fiscal em quest�o.
Public Const ERRO_SLDMESEST_NAO_CADASTRADO = 7647 'Par�metros: iFilialEmpresa, iAno, sProduto
'Registro da tabela SldMesEst n�o cadastrado. Dados do registro: FilialEmpresa=%i, Ano=%i, Produto=%s.
Public Const ERRO_DATA_NAO_PREENCHIDA = 7726 'Sem par�metros
'O preenchimento da Data � obrigat�rio.
Public Const ERRO_LOTE_NAO_PREENCHIDO = 7769 'Sem parametros
'Preenchimento do lote � obrigat�rio.
Public Const ERRO_LEITURA_ITENSNFISCALBAIXADAS = 7796
'Erro na leitura da tabela dos Itens de N.Fiscal Baixadas nos itens vinculados com a Nota Fiscal em quest�o.
Public Const ERRO_LEITURA_ITENSNFISCAL1 = 7797
'Erro na leitura do Item vinculado com a Nota Fiscal em quest�o na tabela dos Itens de N.Fiscal.
Public Const ERRO_LEITURA_ITENSNFISCALBAIXADAS1 = 7798 'NumIntDoc
'Erro na leitura da tabela dos Itens de N.Fiscal Baixadas nos itens vinculados com a Nota Fiscal em quest�o.
Public Const ERRO_LEITURA_CLIENTES1 = 7914 'Parametro lCodCliente
'Ocorreu um erro na leitura dos dados da tabela de Clientes. C�digo do Cliente = %l.
Public Const ERRO_LEITURA_SLDMESEST3 = 8568  'Parametros:iFilialEmpresa, sProduto
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque (SldMesEst). FilialEmpresa=%i, Produto = %s.
Public Const ERRO_SLDMESEST_NAO_CADASTRADO1 = 8569 'Par�metros: iFilialEmpresa, sProduto
'N�o h� Registro cadastrado na tabela SldMesEst para a FilialEmpresa=%i,  Produto=%s.
Public Const ERRO_LEITURA_TITULOSPAGBAIXADOS = 8590 'Par�metro: lNumIntDoc
'Erro na tentativa de ler registro na tabela TitulosPagBaixados com N�mero Interno %l.



''VEIO DE ERROS FAT
Public Const ERRO_ATUALIZACAO_TABELASDEPRECOITENS1 = 8124 'Parametro: iCodTabela
'Erro na atualiza��o da tabela TabelasDePrecoItens com c�digo da Tabela %i.


''VEIO DE ERROS CRFAT
Public Const ERRO_LEITURA_MVDIACLI = 6077
'Erro de leitura na tabela MvDiaCli.
Public Const ERRO_LEITURA_NFISCAL = 6123
'Erro na leitura da tabela NFiscal da Nota Fiscal em quest�o.
Public Const ERRO_LEITURA_ITENSNFISCAL = 6125 'Par�metro: lNumIntNF
'Erro na leitura do Item vinculado com a Nota Fiscal em quest�o na tabela dos Itens de N.Fiscal.
Public Const ERRO_LEITURA_FORNECEDORES = 6208 'Sem parametros
'Erro na leitura da tabela de Fornecedores.
Public Const ERRO_TABELAPRECO_INEXISTENTE = 6403 'Parametro: iCodigo
'A Tabela de Pre�o com C�digo %i n�o est� cadastrada no Banco de Dados.
Public Const ERRO_LEITURA_TABELASDEPRECO1 = 6404 'Parametro: iCodigo
'Erro na leitura da tabela TabelasDePreco com o C�digo %i.
Public Const ERRO_LOCK_TABELASDEPRECO = 6405 'Parametro: iCodigo
'N�o conseguiu fazer o lock na tabela de TabelasDePreco com C�digo da Tabela %i.



''VEIO DE ERROS CPR
Public Const ERRO_FORNECEDOR_NAO_CADASTRADO = 2034 'Parametro: lCodFornecedor
'O Fornecedor com c�digo %l n�o est� cadastrado no Banco de Dados.
Public Const ERRO_LOCK_FORNECEDORES = 2035 'Parametro Codigo Fornecedor
'Erro na tentativa de fazer "lock" na tabela Fornecedores para C�digo Fornecedor = %l .
Public Const ERRO_FILIALFORNECEDOR_NAO_CADASTRADA = 2300 'Parametro: lFornecedor, iFilial
'A Filial %i do Fornecedor %l nao est� cadastrada no Banco de Dados.
Public Const ERRO_FORNECEDOR_NAO_CADASTRADO1 = 2318 'Parametro: sFornecedorNomeRed
'O Fornecedor %s nao est� cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_PRODUTOSFILIAL1 = 2655 'Par�metros: lCodigo
'Erro na leitura da tabela de ProdutosFilial com Fornecedor %l.
Public Const ERRO_LEITURA_MVPERFORN = 2660 'Par�metros: lCodigo
'Erro na leitura da tabela de MvPerForn com Fornecedor %l.
Public Const ERRO_LEITURA_MVDIAFORN = 2657 'Par�metros: lCodigo
'Erro na leitura da tabela de MvDiaForn com Fornecedor %l.
Public Const ERRO_LOCK_MVDIAFORN = 2658 'Par�metros: lCodigo
'Erro na tentativa de fazer "lock" na tabela MvDiaForn com Fornecedor %l.
Public Const ERRO_EXCLUSAO_MVDIAFORN = 2659 'Par�metros: lCodigo
'Erro na tentativa de excluir um registro da tabela MvDiaForn com Fornecedor %l.
Public Const ERRO_LOCK_MVPERFORN = 2661 'Par�metros: lCodigo
'Erro na tentativa de fazer "lock" na tabela MvPerForn com Fornecedor %l.
Public Const ERRO_EXCLUSAO_MVPERFORN = 2662 'Par�metros: lCodigo
'Erro na tentativa de excluir um registro da tabela MvPerForn com Fornecedor %l.
Public Const ERRO_LOCK_MVDIAFORN1 = 2615 'Par�metros: lCodFornecedor, iCodFilial
'Erro na tentativa de fazer "lock" na tabela MvDiaForn com Fornecedor %l e Filial %i.
Public Const ERRO_EXCLUSAO_MVDIAFORN1 = 2616 'Par�metros: lCodFornecedor, iCodFilial
'Erro na tentativa de excluir um registro da tabela MvDiaForn com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_MVPERFORN1 = 2617 'Par�metros: lCodFornecedor, iCodFilial
'Erro na leitura da tabela de MvPerForn com Fornecedor %l e Filial %i.
Public Const ERRO_LOCK_MVPERFORN1 = 2618 'Par�metros: lCodFornecedor, iCodFilial
'Erro na tentativa de fazer "lock" na tabela MvPerForn com Fornecedor %l e Filial %i.
Public Const ERRO_EXCLUSAO_MVPERFORN1 = 2619 'Par�metros: lCodFornecedor, iCodFilial
'Erro na tentativa de excluir um registro da tabela MvPerForn com Fornecedor %le Filial %i.


''VEIO DE ERROS EST
Public Const ERRO_RECEBIMENTO_VINCULADO_NF = 7125 'Parametros lNumeroRecebimento, lNumeroNFiscal
'O Recebimento %l j� est� vinculado com a Nota Fiscal %l.
Public Const ERRO_CLIENTE_RECEB_DIFERENTE_NF = 7126 'ClienteRecebimento
'O Cliente da Nota Fiscal n�o pode ser diferente do Cliente do Recebimento. Cliente do Recebimento = %l
Public Const ERRO_FORN_RECEB_DIFERENTE_NF = 7127 'FornecedorRecebimento
'O Fornecedor da Nota Fiscal n�o pode ser diferente do Fornecedor do Recebimento. Fornecedor do Recebimento = %l
Public Const ERRO_FILCLIENTE_RECEB_DIFERENTE_NF = 7128 'FilialClienteRecebimento
'A Filial Cliente da Nota Fiscal n�o pode ser diferente da Filial do Cliente do Recebimento. Filial Cliente do Recebimento = %i
Public Const ERRO_FILFORN_RECEB_DIFERENTE_NF = 7129 'FilialForncedorRecebimento
'A Filial Fornecedor da Nota Fiscal n�o pode ser diferente da Filial Fornecedor do Recebimento. Filial Fornecedor do Recebimento = %i
Public Const ERRO_SERIE_RECEB_DIFERENTE_NF = 7130 'Serie Recebimento
'A Serie da Nota Fiscal n�o pode ser diferente da Serie do Recebimento. Serie = %s
Public Const ERRO_DATAENTRADA_RECEB_DIFERENTE_NF = 7131 'Data Entrada
'A Data de Entrada da Nota Fiscal n�o pode ser diferente da Data de Entrada do Recebimento. Data de Entrada do Recebimento = %dt
Public Const ERRO_NUMNF_RECEB_DIFERENTE_NF = 7132 'Serie
'O N�mero da Nota Fiscal n�o pode ser diferente do N�mero da Nota Fiscal do Recebimento. N�mero da Nota Fiscal do Recebimento = %l


''VEIO DE ERROS TRB
Public Const ERRO_LEITURA_TIPO_TRIBUTACAO = 7003 'parametro tipo da tributacao
'Erro na leitura do tipo de tributa��o %d.



'Codigos de Aviso - Reservado de 5100 at� 5199
Global Const AVISO_EXCLUSAO_CONTA_ANALITICA = 5100 'Sem parametros
'A conta que est� sendo excluida � analitica. Confirma a exclus�o?
Global Const AVISO_EXCLUSAO_CONTA_SINTETICA_COM_FILHOS = 5101 'Sem parametros
'A conta que est� sendo excluida � sint�tica e possui contas abaixo dela.
'Ao excluir esta conta, suas "filhas" ser�o tamb�m excluidas.
'Confirma a exclus�o?
Global Const AVISO_EXCLUSAO_CONTA_SINTETICA = 5102 'Sem parametros
'A conta que est� sendo excluida � sint�tica e n�o possui contas abaixo dela.
'Confirma a exclus�o?
Global Const SUBSTITUICAO_DOCUMENTO_PENDENTE = 5103 'Parametros: lDocumento, sOrigem, iExercicio, iPeriodoLan
'O documento pendente %l (Origem: %s, Exerc�cio: %i, Per�odo: %i) j� existe.
'Deseja substitu�-lo?
Global Const AVISO_EXCLUSAO_DOCUMENTO = 5104 'Sem Parametros
'Confirma a exclus�o do documento?
Global Const AVISO_EXCLUSAO_CCL_COM_ASSOCIACOES = 5105 'Parametros: Conta
'O Centro de Custo/Lucro que est� sendo excluida possui associa��o com %s. Estas informa��es ser�o excluidas junto com o Centro de Custo/Lucro.
'Confirma a exclus�o?
Global Const AVISO_EXCLUSAO_CCL = 5106 'Sem parametros
'Confirma a exclus�o do Centro de Custo/Lucro?
Global Const AVISO_ATUALIZACAO_TOTAIS = 5107 'Parametros total de cr�dito, total de d�bito, n�mero de lan�amentos
'Os totais calculados atrav�s da leitura dos lan�amentos, difere do
'dos totais exibidos. Totais Calculados: Cr�dito = %d, D�bito = $d e
'N�mero de lan�amentos = %i. Deseja alterar os valores exibidos para
'que fiquem compat�veis com os lan�amentos?
Global Const AVISO_IGUALDADE_TOTAIS = 5108 'Sem parametros
'Os totais calculados atrav�s da leitura dos lan�amentos s�o iguais
'aos exibidos na sua tela.
Global Const AVISO_EXCLUSAO_CCL_COM_DOCAUTO = 5109 'Sem parametros
'O Centro de Custo/Lucro que est� sendo excluida possui documentos autom�ticos associados. Estas informa��es ser�o excluidas junto com o Centro de Custo/Lucro.
'Confirma a exclus�o?
Global Const EXCLUSAO_HISTPADRAO = 5110 'Sem parametros
'Confirma a exclus�o do Hist�rico Padr�o?
Global Const AVISO_LOTE_INEXISTENTE = 5111 'Parametros Filial, Lote, Origem, Periodo, Exercicio
'Nao existe lote com chave: Filial = %i, Lote = %i, Origem = %s, Periodo = %i, Exerc�cio = %i. Deseja criar?
Global Const AVISO_CONTA_INEXISTENTE = 5112 'Parametro Conta
'A Conta %s n�o est� cadastrada. Deseja cadastr�-la?
Global Const AVISO_EXCLUSAO_CONTA_ANALITICA_COM_ASSOCIACOES = 5113 'Sem parametros
'A conta que est� sendo excluida � anal�tica e possui associa��o com centro de custo. Estas informa��es ser�o excluidas junto com a conta.
'Confirma a exclus�o?
Global Const AVISO_CCL_INEXISTENTE = 5114 'Parametro Ccl
'O Centro de Custo/Lucro %s n�o est� cadastrado. Deseja cadastr�-lo?
Global Const AVISO_CONTACCL_INEXISTENTE = 5115 'Parametros Conta e Ccl
'A associa��o da Conta %s com o Centro de Custo/Lucro %s n�o est� cadastrada. Deseja cadastr�-la?
Global Const AVISO_HISTPADRAO_INEXISTENTE = 5116 'Parametro HistPadrao
'O Historico %i n�o est� cadastrado.  Deseja criar agora?
Global Const AVISO_EXCLUSAO_LOTE = 5117 'Sem parametros
'Confirma a exclus�o do Lote?
Global Const AVISO_EXCLUSAO_RATEIO = 5118 'Sem parametro
'Confirma a exclus�o do Rateio?
Global Const AVISO_ALTERACAO_CONTACCL = 5119 'Sem parametro
'As associa��es selecionadas ser�o feitas. Caso exista alguma associa��o de alguma das contas com Centro de Custo n�o selecionado, esta associa��o ser� exclu�da. Confirma a Altera��o?
Public Const AVISO_EXCLUSAO_CCL_COM_ASSOC_DOCAUTO = 5120 'Sem Parametros
'O Centro de Custo/Lucro que est� sendo excluida possui documentos autom�ticos associados e associa��o com conta. Estas informa��es ser�o excluidas junto com o Centro de Custo/Lucro.
'Confirma a exclus�o?
Public Const AVISO_FECHAMENTO_EXERCICIO_EXECUTADO = 5121 'Parametro Nome do Exercicio
'O Disparo do processo de fechamento do exercicio %s foi feito. O seu processamento
'poder� ser acompanhado por uma tela que surgir� a seguir
Public Const AVISO_NAO_HA_LOTE_DESATUALIZADO = 5122
'Nao existe nenhum lote desatualizado.
Public Const AVISO_REABERTURA_EXERCICIO_EXECUTADA = 5123 'Parametro Nome do Exercicio
'O Disparo do processo de rEABERTURA do exercicio %s foi feito. O seu processamento
'poder� ser acompanhado por uma tela que surgir� a seguir
Public Const AVISO_REPROCESSAMENTO_EXERCICIO_EXECUTADA = 5124
'O Disparo do processo de Reprocessamento do exercicio %s foi feito. O seu processamento
'poder� ser acompanhado por uma tela que surgir� a seguir
Public Const AVISO_EXCLUSAO_ORCAMENTO = 5125
'Confirma a exclus�o do Or�amento?
Public Const AVISO_EXCLUSAO_CONTACATEGORIA = 5126
'Confirma a exclus�o da Categoria?
Public Const AVISO_CONTARESULTADO_NAO_ESPECIFICADA = 5127 'Sem parametro
'A conta de resultado n�o foi espeficada. Deseja espeficar agora?
Public Const AVISO_ULTIMO_PERIODO_MAIOR = 5128
'O �ltimo Per�odo possuir� uma dura��o maior que os demais.
Public Const AVISO_EXCLUSAO_CCL_SINTETICA_COM_FILHOS = 5129 'Sem parametros
'O Centro de Custo/Lucro que est� sendo excluido � sint�tico e possui centros de custo/lucro abaixo dele.
'Ao excluir este centro de custo/lucro, seus "filhos" ser�o tamb�m excluidos.
'Confirma a exclus�o?
Public Const AVISO_EXCLUSAO_CCL_SINTETICA = 5130 'Sem parametros
'O Centro de Custo/Lucro que est� sendo excluido � sint�tico e n�o possui centros de custo/lucro abaixo dele.
'Confirma a exclus�o?
Public Const AVISO_HA_LANCAMENTO_DESATUALIZADO = 5131 'Sem parametros
'Existe(m) lan�amento(s) desatualizado(s) para este exerc�cio. Deseja prosseguir?
Public Const AVISO_EXCLUSAO_PADRAOCONTAB = 5132 'Sem Parametros
'Confirma a exclus�o do modelo de contabiliza��o?
Public Const AVISO_LOTE_ATUALIZANDO = 5133 'Par�metro: iFilialEmpresa
'A Filial %i s� possui Lotes que est�o sendo atualizados. N�o existem lotes pendentes de atualiza��o.
Public Const AVISO_EXCLUSAO_MODELORELDRE = 5134
'Confirma a exclus�o do modelo?
Public Const AVISO_ESTORNO_LANCAMENTO_CANCELADO = 5135 'par�metros sOrigem , iExercicio , iPeriodo , lDoc
'O estorno foi cancelado . Origem= %s, Exercicio= %s, Periodo= %s, Doc = %s.
Public Const AVISO_ESTORNO_LOTE_CANCELADO = 5136 'par�metros : sOrigem , iExercicio, iPeriodo, iLote
'O estorno foi cancelado . Origem= %s, Exercicio= %i, Periodo= %i, Lote= %i
Public Const AVISO_ELEMENTO_TEM_FILHOS = 5137 'Sem parametros
'Este elemento possui 'filhos'. Deseja exclu�-lo assim mesmo ?
Public Const AVISO_CANCELAR_APURACAO_RATEIOS = 5138
'Confirma o cancelamento da apura��o dos rateios ?
Public Const AVISO_EXCLUSAO_EXERCICIO = 5139 'Sem par�metros
'Confirma a exclus�o do Exerc�cio?
Public Const AVISO_ATUALIZACAO_LOTE_TOTAIS_DIFERENTES = 5140 'Parametros iFilialEmpresa, sOrigem, iExercicio, iExercicio, iLote, iNumDocInf, iNumDocAtual, dTotInf, dTotAtual
'Os totais calculados atrav�s da leitura dos lan�amentos, difere do
'dos totais informados para o lote (Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i). N�mero de Documentos (Informado = %i, Cadastrados = %i) e Valores dos Documentos (Informado = %d e Cadastrados = %d).
'Confirma a contabiliza��o deste lote?
Public Const AVISO_LOTE_ATUALIZANDO_GRAVA = 5141 'Par�metro: iLote
'O Lote %i est� sendo atualizado, deseja continuar a altera��o ?
Public Const AVISO_DATA_TROCADA_ULTIMO_PERIODO_ABERTO = 5142 'Parametros: DataAtual, DataFim, DataAtual
'Aten��o. A data do sistema %s foi alterada para %s pois n�o existia periodo cont�bil aberto. Para que o sistema abra com a data %s , favor cadastrar um exercicio que englobe-a.


