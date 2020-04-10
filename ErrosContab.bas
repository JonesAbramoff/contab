Attribute VB_Name = "ErrosContab"
Option Explicit

'Códigos de Erro - Reservado de 9000 até 9999
Global Const ERRO_LEITURA_LOTE = 1 'Parametros Filial, origem, exercicio, periodo, lote
'Erro na leitura do lote. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_LOTE_ATUALIZADO = 2
'O lote está atualizado.
Global Const ERRO_LOCK_LOTE = 3 'Parametros Filial, origem, exercicio, periodo, lote
'Não foi possível fazer o LOCK do Lote que possui a seguinte chave: Filial=%i, Origem= %s, Exercicio= %i, Periodo=%i, Lote=%i.
Global Const ERRO_LEITURA_PERIODOS = 4 'Parametros Exercicio, Periodo
'Ocorreu um erro na leitura da tabela de Periodos. Exercicio = %i e Periodo = %i.
Global Const ERRO_LOCK_PERIODO = 5 'Parametro Exercicio, Periodo
'Não conseguiu fazer o lock de um registro da tabela Periodo. Exercicio = %i e Periodo = %i.
Global Const ERRO_LEITURA_LANCAMENTOS = 6 'Parametros Filial, Origem, exercicio, periodo, lote
'Erro na leitura dos Lançamentos Pendentes do lote que possui chave Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Lote= %i.
Global Const ERRO_LEITURA_MVDIACTA = 7 'Sem parametros
'Erro na leitura da tabela de Movimentos diários da Conta.
Public Const ERRO_ATUALIZACAO_MVDIACTA = 8  'Parametros Filial, conta, data
'Ocorreu um erro na atualização da tabela de Saldos Diários de Conta. Filial = %i, Conta = %s e Data = %s.
Global Const ERRO_ATUALIZACAO_MVPERCTA = 9 'parametro Filial, Exercicio, Conta
'Ocorreu um erro na atualização da tabela de Saldos Periódicos de Conta. Filial = %i, Exercicio = %i e Conta = %s.
Global Const ERRO_ATUALIZACAO_MVPERCCL = 10 'parametros Filial, exercicio, ccl, conta
'Ocorreu um erro na atualização da tabela que guarda os saldos de centro de custo/lucro. Filial = %i, Exercicio = %i, Centro de Custo/Lucro = %s, Conta = %s.
Global Const ERRO_ATUALIZACAO_MVDIACCL = 11 'parametros Filial, ccl, conta, data
'Ocorreu um erro na atualização da tabela de Saldos Diários de Centro de Custo/Lucro. Filial = %i, Centro de Custo/Lucro = %s, Conta = %s e Data = %s.
Global Const ERRO_LEITURA_EXERCICIO = 12 'Parametro Exercicio
'Erro na leitura do exercicio %i
Global Const ERRO_LANCAMENTOS_EXERCICIO_FECHADO = 13 'Parametro Exercicio
'O Exercício %i está fechado. Não é possível fazer gravação ou exclusão de lançamentos.
Global Const ERRO_LOCACAO_EXERCICIO = 14
'Não conseguiu fazer o Lock do Exercício.
Global Const ERRO_LEITURA_PERIODO1 = 15 'Sem parametros
'Ocorreu um erro na leitura da tabela de Periodos. Verifique se o periodo existe para os parametros fornecidos.
Global Const ERRO_LEITURA_CONFIGURACAO = 16 'Sem parametros
'Erro na leitura da tabela de Configuração
Global Const CONTA_SEM_CCL = 17
'Conta não possui Centro de Custo/Lucro.
Global Const ERRO_LEITURA_MVPERCTA = 18  'Nao tem parametros
'Ocorreu um erro na leitura da tabela que contém os Saldos das Contas (MvPerCta).
Global Const ERRO_INSERCAO_LOTE = 19 'Parametros  Filial, Origem, Exercicio, Periodo, Lote
'Ocorreu um erro na inclusão de um lote na tabela de Lotes. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_LEITURA_PLANOCONTA = 20 'Sem Parametros
'Erro na leitura da tabela Plano de Contas.
Global Const ERRO_INSERCAO_LANCAMENTO = 21 'Parametros Filial, origem, exercicio, periodo, doc, seq
'Ocorreu um erro no cadastramento do seguinte lançamento. Filial = %i, Origem = %s, Exercicio=%i, Periodo=%i, Documento=%l, Sequencial=%i.
Global Const ERRO_ATUALIZACAO_LOTE = 22 'Parametro Filial, origem, exercicio, periodo, lote
'Erro de atualização do Lote. Filial= %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_EXERCICIO_NAO_FECHADO = 23
'O Exercício não está fechado.
Global Const ERRO_ATUALIZACAO_EXERCICIOS = 24 'Parametro Exercicio
'Ocorreu um erro na atualização do exercicio %i.
Global Const ERRO_EXERCICIO_NAO_ABERTO = 25 ' Parametro Exercicio
'O Exercicio %s não está aberto. Portanto não é capaz de receber lançamentos.
Global Const ERRO_LEITURA_MVPERCCL = 26  'Sem parametro
'Erro de leitura da tabela MvPerCcl.
Global Const ERRO_EXERCICIO_JA_EXISTE = 27 'Parametro Exercicio
'O exercicio %i já está cadastradado.
Global Const ERRO_INSERCAO_EXERCICIO = 28 'Parametro Exercicio
'Erro na inserção do exercicio %i na tabela Exercicios.
Global Const ERRO_INSERCAO_MVPERCTA = 29 'Parametros: Filial, Exercicio, Conta
'Erro de inserção na tabela de saldos de conta (MvPerCta). Filial = %i, Exercicio = %i, Conta= %s.
Global Const ERRO_LEITURA_CONTACCL = 30 'Sem Parametro
'Erro de leitura na tabela ContaCcl.
Global Const ERRO_INSERCAO_MVPERCCL = 31 'Parametros Filial, Exercicio, Ccl, Conta
'Ocorreu um erro na inserção da Filial = %i Exercicio = %i Ccl = %s Conta = %s na tabela de Saldos de Centro de Custo/Lucro(MvPerCcl).
Global Const ERRO_ATUALIZACAO_MVPERCTA1 = 32  'Sem parametro
'Erro de atualização na tabela de saldos de conta.
Global Const ERRO_ATUALIZACAO_MVPERCCL1 = 33 'Sem parametros
'Erro de atualização na tabela de saldos de Centro de Custo/Lucro.
Global Const ERRO_LOCACAO_PERIODO = 34
'Não conseguiu fazer o Lock do Período.
Global Const ERRO_ATUALIZACAO_PERIODO = 35 'Parametros Exercicio, Periodo
'Ocorreu um erro na atualização da tabela Periodo. Exercicio = %i e Periodo = %i.
Global Const ERRO_UNLOCK_EXERCICIO = 36
'Ocorreu um erro na liberação do lock do Exercício.
Global Const ERRO_UNLOCK_LOTE = 37
'Ocorreu um erro na liberação do lock do Lote.
Global Const ERRO_UNLOCK_PERIODO = 38 'Parametros Periodo, Exercicio
'Ocorreu um erro na liberação do lock do Periodo %s do Exercicio %s
Global Const ERRO_LEITURA_RESULTADO = 39 'Parametro CodigoApuracao
'Ocorreu um erro na leitura da tabela Resultado. Codigo de Apuracao = %l
Global Const ERRO_LEITURA_MVPERCTA1 = 40  'Parametro Filial, Exercicio, Conta
'Ocorreu um erro na leitura da tabela de Saldos de Conta (MvPerCta). Filial=%i, Exercicio=%i, Conta=%s
Global Const ERRO_ATUALIZACAO_PERIODO1 = 41
'Erro de atualização do Período.
Public Const ERRO_INSERCAO_PERIODO = 42 'Parametros Exercicio, Periodo
'Ocorreu um erro ao tentar inserir um registro na tabela de Periodos. Exercicio = %i e Periodo = %i.
Global Const ERRO_LEITURA_LANCAMENTOS1 = 43 'Parametros Exercicio, Periodo
'Erro de leitura na tabela de Lançamentos. Exercício %s e Período %s.
Global Const ERRO_LOCK_MVPERCTA = 44 'parametro exercicio,conta
'Ocorreu um erro ao tentar executar o "Lock" em um registro da tabela de saldos de conta(MvPerCta). Exercicio = %s, Conta = %s.
Global Const ERRO_UNLOCK_MVPERCTA = 45 'parametro exercicio, conta
'Ocorreu um erro na liberação do lock em um registro da tabela de saldos de Centro de Custo / Lucro. Exercício = %s e Centro de Custo/Lucro = %s e Conta = %s.
Global Const ERRO_LOCK_MVPERCCL = 46 'parametro Filial, exercicio, ccl, conta
'Ocorreu um erro ao tentar executar o "Lock" em um registro da tabela de saldos de centro de custo/lucro(MvPerCcl). Filial = %i, Exercicio = %i, Centro de Custo/Lucro = %s, Conta = %s.
Global Const ERRO_UNLOCK_MVPERCCL = 47 'parametro exercicio, ccl, conta
'Ocorreu um erro na liberação do lock em um registro da tabela de saldos de Centro de Custo / Lucro. Exercício %s, Centro de Custo/Lucro %s e Conta %s.
Global Const ERRO_LOCK_MVDIACTA = 48 'parametro conta, data
'Ocorreu um erro ao tentar executar o 'Lock' em um registro da tabela de saldos diários de conta. Conta %s e Data %s.
Global Const ERRO_UNLOCK_MVDIACTA = 49 'parametro conta, data
'Ocorreu um erro na liberação do lock em um registro da tabela de saldos diários de conta. Conta %s e Data %s.
Global Const ERRO_LEITURA_MVPERCCL1 = 50  'parametros Filial, exercicio,ccl, conta
'Ocorreu um erro de leitura na tabela de Saldos de Centro de Custo/Lucro (MvPerCcl) para a Filial = %i, exercicio = %i, centro de custo/lucro = %s e conta %s.
Global Const ERRO_LOCK_MVDIACCL = 51 'parametro ccl, conta, data
'Ocorreu um erro ao tentar executar o 'Lock' em um registro da tabela de saldos diários de Centro de Custo/Lucro. Centro de Custo/Lucro %s, Conta %s e Data %s.
Global Const ERRO_UNLOCK_MVDIACCL = 52 'parametro ccl, conta, data
'Ocorreu um erro na liberação do lock em um registro da tabela de saldos diários de Centro de Custo/Lucro. Centro de Custo/Lucro %s, Conta %s e Data %s.
Global Const ERRO_LEITURA_MVDIACTA1 = 53 'parametro Filial, conta, data
'Ocorreu um erro na leitura da tabela de Saldos Diário de Conta. Filial=%i, Conta=%s e Data=%s.
Global Const ERRO_EXCLUSAO_LANCAMENTO = 54 'parametro Filial, origem, exercicio, periodo, doc, seq
'Ocorreu um erro na exclusão do lançamento. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Doc = %l, Seq = %i.
Public Const ERRO_LEITURA_MVDIACCL = 55 'Sem parametro
'Ocorreu um erro na leitura da tabela de Saldos Diários de Centro de Custo/Lucro.
Public Const ERRO_LEITURA_MVDIACCL1 = 56 'parametro Filial, ccl, conta, data
'Ocorreu um erro na leitura da tabela de Saldos Diários de Centro de Custo/Lucro. Filial = %i, Centro de Custo/Lucro = %s, Conta = %s e Data = %s.
Public Const ERRO_INSERCAO_MVDIACCL = 57 'parametro Filial, ccl, conta, data
'Ocorreu um erro na inserção de um registro na tabela de Saldos Diários de Centro de Custo/Lucro. Filial=%i, Centro de Custo/Lucro=%s, Conta=%s e Data=%s.
Public Const ERRO_INSERCAO_MVDIACTA = 58 'parametro  Filial, conta, data
'Ocorreu um erro na inserção de um registro na tabela de Saldos Diários de Conta. Filial=%i, Conta=%s e Data=%s.
Global Const ERRO_INSERCAO_SORT = 59 'parametro  conta, data
'Erro na inserção de dados no arquivo de sort. Conta %s e Data %s.
Global Const ERRO_LEITURA_EXERCICIOORIGEM = 60 'parametros Filial, exercicio, periodo, origem
'Ocorreu um erro na leitura da tabela ExercicioOrigem. Filial = %i, Exercicio = %i, Periodo = %i, Origem = %s.
Global Const ERRO_LEITURA_LOTE1 = 61 'Sem parametros
'Ocorreu um erro na leitura da tabela de lotes.
Global Const ERRO_LOCK_EXERCICIOORIGEM = 62 'parametros Filial, exercicio, periodo, origem
'Ocorreu um erro ao tentar fazer um "lock" de um registro da tabela ExercicioOrigem. Filial = %i, Exercicio = %i, Periodo = %i, Origem = %s.
Global Const ERRO_ATUALIZACAO_EXERCICIOORIGEM = 63 'parametro Filial, exercicio, periodo, origem
'Ocorreu um erro na atualização da tabela ExercicioOrigem. Filial = %i, Exercicio = %i, Periodo = %i, Origem = %s.
Global Const ERRO_NUMERO_LOTE_NAO_PREENCHIDO = 64
'Número do Lote não foi preenchido.
Global Const ERRO_NUMERO_DOCUMENTO_NAO_PREENCHIDO = 65
'Número do Documento não foi preenchido.
Global Const ERRO_VALOR_LANCAMENTO_NAO_PREENCHIDO = 66
'Valor do Lançamento não foi preenchido.
Global Const ERRO_LEITURA_LOTE2 = 67 'Parametros origem, exercicio, periodo, lote
'O Lote não está cadastrado - Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i
Global Const ERRO_LEITURA_LOTE_VAZIO = 68
'Tabela de Lotes vazia.
Global Const ERRO_LEITURA_PERIODO = 69 'Parametro Exercicio, Periodo
'Erro de Leitura na Tabela de Periodos
Global Const ERRO_LOTE_NAO_ATUALIZADO = 70 'Parametros Origem, Exercicio, Periodo, Lote
'Erro na atualização do lote. Origem %s, Exercício %s, Período %s e Lote %s.
Global Const ERRO_ATUALIZACAO_EXERCICIO = 71 'Parametro Exercicio
'Erro de atualização do Exercicio %i.
Global Const ERRO_LOTE_EXERCICIO_FECHADO = 72 'Parametro Exercicio
'Lote não pode ser criado num exercicio fechado. Exercicio = %i
Global Const ERRO_LOTE_ATUALIZADO_NAO_EDITAVEL = 73 'Parametros Filial, Origem, Exercicio, Periodo, Lote
'Este lote está atualizado e portanto não pode ser editado. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_LOTE_ATUALIZADO_NAO_EXCLUIR = 74 'Parametros Filial, Origem, Exercicio, Periodo, Lote
'Este lote está atualizado e portanto não pode ser removido. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_EXCLUSAO_LOTE = 75 'Parametros Origem, Exercicio, Periodo, Lote
'Houve um erro na exclusão do lote do banco de dados.
Global Const ERRO_LEITURA_CONTA = 76  'Sem parametros.
'Erro de leitura da tabela Plano de Contas
Global Const ERRO_LEITURA_ORIGEM = 77
'Erro de leitura na tabela Origem.
Global Const ERRO_PLANO_CONTAS_VAZIO = 78
'Tabela de Plano de Contas Vazia.
Global Const ERRO_INSERCAO_LANCAMENTOS = 79 'Sem parametros
'Erro na Inserção dos Lançamentos na Tabela de Lançamentos Pendentes
Global Const ERRO_M_LOTE_LOTE_ATUALIZADO = 80 'Parametros Filial, sOrigem, iExercicio, iPeriodo, iLote
'Este lote já foi contabilizado, portanto não pode ser editado. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_AUSENCIA_LANCAMENTOS_GRAVAR = 81 'Sem parametros
'Não há Lançamentos para Gravar.
Global Const ERRO_TABELA_CCL_VAZIA = 82 'Sem parametro
'Tabela de Centros de Custo e de Lucro Vazia.
Global Const ERRO_DATA_DOCUMENTO_NAO_PREENCHIDA = 83 'Sem Parametro
'Data do Documento não preenchida.
Global Const ERRO_LEITURA_PLANOCONTA1 = 84 'Parametro Conta
'A conta %s não está cadastrada
Global Const ERRO_LEITURA_HISTPADRAO = 85 'Parametro HistPadrao
'Erro na leitura do Historico Padrao - Historico = %i
Global Const ERRO_LEITURA_PLANOCONTA2 = 86 'Parametro ContaSimples
'Erro na leitura do Plano de Contas. Conta Simplificada = %l
Global Const ERRO_CONTASIMPLES_JA_UTILIZADA = 87 'Parametros ContaSimples, Conta
'A conta simplificada %s é utilizada na conta %s
Global Const ERRO_LEITURA_LANCAMENTOS2 = 88 'Parametro Conta
'Erro na leitura da tabela de Lançamentos. Conta = %s
Global Const ERRO_LEITURA_LANPENDENTE = 89 'Parametro Conta
'Erro na leitura da tabela de Lançamentos Pendentes. Conta = %s
Global Const ERRO_LEITURA_CONTACCL1 = 90 'Parametro Conta
'Erro na leitura da tabela ContaCcl. Conta = %s.
Global Const ERRO_LEITURA_PLANOCONTA3 = 91 'Parametro Conta
'Erro na leitura do Plano de Contas. Conta = %s
Global Const ERRO_CONTAPAI_INEXISTENTE = 92 'Sem Parametro
'A conta em questão não tem uma conta "pai" dentro da hierarquia do plano de contas.
Global Const ERRO_CONTA_SINTETICA_COM_LANCAMENTOS = 93 'Sem parametro
'A conta não pode ser sintética pois possui lançamentos já contabilizados.
Global Const ERRO_CONTA_SINTETICA_COM_LANC_PEND = 94 'Sem parametro
'A conta não pode ser sintética pois possui lançamentos pendentes.
Global Const ERRO_CONTA_SINTETICA_ASSOCIADA_CCL = 95 'Sem parametro
'A conta não pode ser sintética pois está associada a centro de custo.
Global Const ERRO_CONTA_ANALITICA_COM_FILHAS = 96 'Sem parametro
'A conta não pode ser analítica pois possui contas embaixo dela.
Global Const ERRO_DOCUMENTO_JA_LANCADO = 97 'Parametros: lDocumento, sOrigem, iExercicio, iPeriodoLan
'O documento <lDocumento> (Origem: <sOrigem>, Exercício: <iExercicio>, Período: <iPeriodoLan>) já foi lançado.
Global Const ERRO_CONTA_NAO_INFORMADA = 98 'Sem parametro
'A conta não foi informada.
Global Const ERRO_LEITURA_LANCAMENTOS_PENDENTES = 99 'Sem parametro
'Erro na Leitura da Tabela de Lançamentos Pendentes
Global Const ERRO_LEITURA_LANCAMENTOS3 = 100 'Sem parametro
'Erro na Leitura da Tabela de Lançamentos
Global Const ERRO_MASCARA_CONTA_OBTERNIVEL = 101 'Parametro Conta
'Erro na obtençao do nível da conta. Conta = %s.
Global Const ERRO_INSERCAO_PLANOCONTA = 102 'Parametro Conta
'Erro na inserção da conta %s na tabela PlanoConta.
Global Const ERRO_LEITURA_EXERCICIOS = 103 'Sem parametro
'Erro de leitura da tabela Exercicios. Verifique se o(s) exercicio(s) existe(m) para os parametros informados.
Global Const ERRO_ATUALIZACAO_PLANOCONTA = 104 'Parametro Conta
'Erro de atualização da conta %s.
Global Const ERRO_LOTE_ATUALIZADO_NAO_RECEBE_LANCAMENTOS = 105 'Parâmetros: FilialEmpresa, iLote, iExercicio, iPeriodo, sOrigem
'Lote com chave Filial= %i Lote = %i Exercício = %i Período =  %i Origem = %s está Atualizado. Não pode incluir/alterar/excluir Lançamentos.
Global Const ERRO_DOCUMENTO_NAO_BALANCEADO = 106 'Parametro: lDoc
'O Documento Contábil <lDoc> não está balanceado (soma de créditos diferente soma de débitos).
Public Const ERRO_LEITURA_LANPENDENTE1 = 107 'Parametros Filial, Origem, Exercicio, Periodo, Lote
'Erro na leitura dos Lançamentos Pendentes do Lote que possui a chave Filial = %i, Origem = %s, Exericicio = %i, Periodo = %i, Lote = %i
Global Const ERRO_LOTE_COM_LANC_PEND_NAO_EXCLUIR = 108 'Parametros Filial, Origem, Exercicio, Periodo, Lote
'Lote com lançamentos pendentes não pode ser excluido. Filial = %i, Origem = %s, Exericicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_LOCK_PLANOCONTA = 109 'Parametro Conta
'Não conseguiu fazer o lock da conta %s.
Global Const ERRO_EXCLUSAO_PLANOCONTA = 110 'Parametro Conta
'Houve um erro na exclusão da conta %s do banco de dados.
Global Const ERRO_EXCLUSAO_CONTA_COM_LANCAMENTOS = 111 'Sem Parametros
'A conta possui lançamentos contabilizados. Portanto não pode ser excluida.
Global Const ERRO_EXCLUSAO_CONTA_COM_LANC_PEND = 112 'Sem Parametros
'A conta possui lançamentos pendentes. Portanto não pode ser excluida.
Global Const ERRO_LOTE_INEXISTENTE = 113  'Parametros Origem, exercicio, periodo, lote
'O lote não está cadastrado.  Origem= %s, Exercicio= %i, Periodo= %i, Lote= %i
Global Const ERRO_EXCLUSAO_CONTACCL = 114 'Parametro Conta, Ccl
'Houve um erro na exclusão da associação da conta %s com o centro de custo/lucro %s.
Global Const ERRO_LEITURA_DOCAUTO = 115 'Sem parametro
'Erro de leitura da tabela de Documentos Automáticos.
Global Const ERRO_EXCLUSAO_DOCAUTO = 116 'Parametros documento, sequencial
'Houve um erro na exclusão do documento automático %l, sequencial %i do banco de dados.
Global Const ERRO_LEITURA_RATEIOON = 117 'Parametro conta.
'Erro de leitura da tabela de Rateios On-Line para a conta %s.
Global Const ERRO_LEITURA_RATEIOOFF = 118 'Parametro conta.
'Erro de leitura da tabela de Rateios Off-Line para a conta %s.
Global Const ERRO_CONTA_NAO_CADASTRADA = 119 'Parametro conta.
'A conta %s não está cadastrada.
Global Const ERRO_EXCLUSAO_CONTA_COM_RATEIOON = 120 'Sem Parametros
'A conta é usada no rateio on-line. Portanto não pode ser excluida.
Global Const ERRO_EXCLUSAO_CONTA_COM_RATEIOOFF = 121 'Sem Parametros
'A conta é usada no rateio off-line. Portanto não pode ser excluida.
Global Const ERRO_LEITURA_LOTE3 = 122 'Parametros origem, exercicio, lote
'Erro na leitura do lote - Origem = %s, Exercicio = %i, Lote = %i
Global Const ERRO_LEITURA_LOTE4 = 123 'Parametros origem, exercicio, lote
'O Lote não está cadastrado - Origem = %s, Exercicio = %i, Lote = %i
Global Const Erro_Mascara_RetornaContaNoNivel = 124 'Parametros Conta, Nivel
'Erro na obtençao da conta %s no nível %i.
Global Const ERRO_PERIODOS_DIFERENTES = 125 'Parâmetros Periodo do Documento, Periodo do Lote
'Período do Documento %i diferente do Período do Lote %i.
Global Const ERRO_LOTE_ATUALIZADO_NAO_SE_EXCLUI = 126 'Parâmetros: iLote, iExercicio, iPeriodo, sOrigem
'Lote com chave Lote <iLote>, Exercício <iExercicio>, Período <iPeriodo>, Origem <sOrigem> está Atualizado. Não pode ser excluído.
Global Const ERRO_LOTE_INEXISTENTE1 = 127 'iLote, iExercicio, iPeriodo, sOrigem
'Nao existe Lote com chave Lote <iLote>, Origem <sOrigem>, Periodo <iPeriodo>, Exercício <iExercicio>.
Global Const ERRO_DOCUMENTO_NAO_EXISTE = 128 'Parametro: lDoc
'Não existe Documento <lDoc>.
Global Const ERRO_EXERCICIO_FECHADO = 129 'Parametro Exercicio.
'O Exercicio %i encontra-se Fechado. Portanto não pode receber novos lançamentos.
Global Const ERRO_LEITURA_CCL = 130 'Parametro Ccl
'Erro na leitura da tabela de Centros de Custo/Lucro. Centro de Custo/Lucro = %s
Global Const ERRO_LEITURA_CONTACCL2 = 131 'Parametro Ccl
'Erro na leitura da tabela ContaCcl. Centro de Custo/Lucro = %s.
Global Const ERRO_LEITURA_LANCAMENTOS4 = 132 'Parametro Ccl
'Erro na leitura da tabela de Lançamentos. Centro de Custo/Lucro = %s
Global Const ERRO_LEITURA_LANPENDENTE2 = 133 'Parametro Ccl
'Erro na leitura da tabela de Lançamentos Pendentes. Centro de Custo/Lucro = %s
Global Const ERRO_CCL_NAO_CADASTRADO = 134 'Parametro Ccl.
'O Centro de Custo/Lucro %s não está cadastrado.
Global Const ERRO_EXCLUSAO_CCL_COM_LANCAMENTOS = 135 'Sem Parametros
'O Centro de Custo/Lucro possui lançamentos contabilizados. Portanto não pode ser excluido.
Global Const ERRO_EXCLUSAO_CCL_COM_LANC_PEND = 136 'Sem Parametros
'O Centro de Custo/Lucro possui lançamentos pendentes. Portanto não pode ser excluido.
Global Const ERRO_CCL_NAO_INFORMADO = 137 'Sem parametro
'O Centro de Custo/Lucro não foi informado.
Global Const ERRO_LOCK_CCL = 138 'Parametro Ccl
'Não conseguiu fazer o lock do centro de custo/lucro %s.
Global Const ERRO_EXCLUSAO_CCL = 139 'Parametro Ccl
'Houve um erro na exclusão do Centro de Custo/Lucro %s do banco de dados.
Global Const ERRO_ATUALIZACAO_CCL = 140 'Parametro Ccl
'Erro de atualização do Centro de Custo/Lucro %s.
Global Const ERRO_INSERCAO_CCL = 141 'Parametro Ccl
'Erro na inserção do Centro de Custo/Lucro %s na tabela de centros de custo.
Global Const ERRO_LEITURA_MVPERCTA2 = 142  'Parametro Conta
'Ocorreu um erro de leitura na tabela de Saldos de Conta (MvPerCta) para a conta %s.
Global Const ERRO_EXCLUSAO_CONTA_COM_MOVIMENTO = 143 'Sem Parametros
'A conta possui movimento, portanto não pode ser excluida.
Global Const ERRO_EXCLUSAO_MVPERCTA = 144 'Parametro Exercicio, conta
'Houve um erro na exclusão do saldo de conta (MvPerCta) do Exercicio %i, Conta %s.
Global Const ERRO_LEITURA_MVPERCCL2 = 145  'Parametro Conta
'Ocorreu um erro de leitura na tabela de Saldos de Centro de Custo/Lucro (MvPerCcl) para a conta %s.
Global Const ERRO_EXCLUSAO_MVPERCCL = 146 'Parametro Exercicio, ccl, conta
'Houve um erro na exclusão do saldo de centro de custo/lucro (MvPerCcl) do Exercicio %i, Ccl %s, Conta %s.
Global Const ERRO_CONTAPAI_ANALITICA = 147 'Sem Parametro
'A conta em questão possui uma conta "pai" analítica. Contas analíticas não podem conter contas embaixo dela.
Global Const ERRO_ORIGEM_NAO_PREENCHIDA = 148 'Sem Parametro
'Lançamentos devem ter origem.
Global Const ERRO_EXCLUSAO_CCL_COM_MOVIMENTACAO = 149 'Sem Parametros
'O Centro de Custo/Lucro possui movimento, portanto não pode ser excluido.
Global Const ERRO_EXCLUSAO_CCL_COM_RATEIO = 150 'Sem Parametros
'O Centro de Custo/Lucro possui Rateio associado. Portanto não pode ser excluido.
Global Const ERRO_EXCLUSAO_CCL_COM_RATEIOON = 151 'Sem Parametros
'O Centro de Custo/Lucro é usado no rateio on-line. Portanto não pode ser excluido.
Global Const ERRO_EXCLUSAO_CCL_COM_RATEIOOFF = 152 'Sem Parametros
'O Centro de Custo/Lucro é usado no rateio off-line. Portanto não pode ser excluido.
Global Const ERRO_LEITURA_MVPERCCL3 = 153  'Parametro Ccl
'Ocorreu um erro de leitura na tabela de Saldos de Centro de Custo/Lucro (MvPerCcl) para o Centro de Custo/Lucro %s.
Global Const ERRO_EXCLUSAO_CONTA_COM_LANC_PEND1 = 154 'Parametro Conta
'A conta %s possui lançamentos pendentes. Portanto não pode ser excluida.
Global Const ERRO_EXCLUSAO_CONTA = 155 'Parametro Conta
'Houve um erro na exclusão da Conta %s do banco de dados.
Global Const ERRO_EXCLUSAO_MVPERCCL1 = 156 'Parametro ccl
'Houve um erro na exclusão do saldo de centro de custo/lucro (MvPerCcl) do Exercicio %i, Ccl %s, Conta %s.
Global Const ERRO_LEITURA_HISTPADRAO1 = 157 'Sem Parametros
'Erro na leitura da tabela de Histórico Padrão
Global Const ERRO_LOCK_CONFIGURACAO = 158 'Sem Parametros
'Não conseguiu fazer o lock na tabela de Configuração.
Global Const ERRO_ATUALIZACAO_CONFIGURACAO = 159 'Sem Parametros
'Erro de atualização da tabela de Configuração.
Global Const ERRO_HISTPADRAO_NAO_INFORMADO = 160 'Sem parametro
'O Código do Histórico Padrão não foi informado.
Global Const ERRO_HISTPADRAO_NAO_CADASTRADO = 161 'Parametro HistóricoPadrão
'O Histórico Padrão %i não está cadastrado.
Global Const ERRO_ATUALIZACAO_HISTPADRAO = 162 'Parametro HistóricoPadrão
'Erro de atualização do Histórico Padrão %i.
Global Const ERRO_INSERCAO_HISTPADRAO = 163 'Parametro HistóricoPadrão
'Erro na inserção do Histórico Padrão %i na tabela Histórico Padrão.
Global Const ERRO_LOCK_HISTPADRAO = 164 'Parametro HistóricoPadrão
'Não conseguiu fazer o lock do Histórico Padrão %i.
Global Const ERRO_EXCLUSAO_HISTPADRAO = 165 'Parametro HistóricoPadrão
'Houve um erro na exclusão do Histórico Padrão %i do banco de dados.
Global Const ERRO_HISTPADRAO_PRESENTE_PLANO_CONTAS = 166 'Parametro HistóricoPadrão
'Não é possível excluir o Histórico Padrão %i que está presente no Plano de Contas.
Global Const ERRO_COLUNA_GRID_INEXISTENTE = 167 'Parametro Titulo da Coluna
'A coluna cujo título é: %s não foi encontrada no Grid.
Global Const ERRO_LOCK_EXERCICIO = 168 'Parametro Exercicio
'Não conseguiu fazer o lock do Exercicio %i.
Global Const ERRO_LOTE_PERIODO_FECHADO = 169 'Parametro Exercicio, Periodo
'Lote não pode ser criado num periodo fechado. Exercicio = %i, Periodo = %i
Global Const ERRO_CONTA_INATIVA = 170 'Parametro Conta
'Esta conta não está ativa. Conta = %s.
Global Const ERRO_CONTA_NAO_ANALITICA = 171 'Parametro conta
'A conta %s não é analítica, portanto não pode ter lançamentos associados.
Global Const ERRO_INSERCAO_CONTACCL = 172 'Parametros Conta e Ccl
'Erro na inserção da associação da conta %s com o centro de custo/lucro %s na tabela ContaCcl.
Global Const ERRO_EXCLUSAO_ASSOC_CCLCONTA_COM_MOV = 173 'Parametros Conta e Ccl
'Existe movimento para a associação da conta %s com o centro de custo/lucro %s. Portanto não pode-se excluir a associação.
Global Const ERRO_EXCLUSAO_CONTACCL_COM_RATEIOON = 174 'Parametros Conta Ccl
'A associação da Conta %s com o Centro de Custo/Lucro %s é usado no rateio on-line. Portanto não pode ser excluida.
Global Const ERRO_EXCLUSAO_CONTACCL_COM_RATEIOOFF = 175 'Parametros Conta Ccl
'A associação da Conta %s com o Centro de Custo/Lucro %s é usado no rateio off-line. Portanto não pode ser excluido.
Global Const ERRO_LEITURA_CONTACCL3 = 176 'Parametros Conta Ccl
'Erro na leitura da tabela ContaCcl. Conta = %s, Centro de Custo/Lucro = %s.
Global Const ERRO_CONTACCL_NAO_CADASTRADO = 177 'Parametros Conta Ccl
'A associação da Conta %s com o Centro de Custo/Lucro %s não está cadastrada.
Global Const ERRO_LOCK_CONTACCL = 178 'Parametros Conta Ccl
'Não conseguiu fazer o lock da associação da Conta %s com o Centro de Custo/Lucro %s.
Global Const ERRO_UNLOCK_PLANOCONTA = 179 'Parametro Conta
'Não conseguiu fazer liberar o lock da conta %s.
Global Const ERRO_UNLOCK_CONTACCL = 180 'Parametros Conta Ccl
'Não conseguiu liberar o lock da associação da Conta %s com o Centro de Custo/Lucro %s.
Global Const ERRO_DESCRICAO_NAO_INFORMADA = 181 'Sem parametro
'Descrição do Histórico Padrão não foi informada.
Global Const ERRO_LEITURA_LANPENDENTE3 = 182 'Sem parametros
'Erro na leitura da tabela de Lançamentos Pendentes.
Global Const ERRO_DOC_NAO_CADASTRADO = 183 'Parametros Origem, Exercicio, PeriodoLan, Doc
'O Documento não está cadastrado. Origem = %s Exercicio = %i Periodo = %i Documento = %l
Global Const Erro_Mascara_MascararCcl = 184 'Parametro Ccl
'Erro na formatação do Centro de Custo/Lucro %s.
Global Const Erro_Mascara_MascararConta = 185 'Parametro Conta
'Erro na formatação da Conta %s.
Global Const Erro_Mascara_RetornaContaPai = 186 'Parametro Conta
'Erro na função que retorna a conta de nivel imediatamente superior da Conta %s.
Public Const ERRO_DESCRICAO_COM_CARACTER_INICIAL_ERRADO = 187 'Sem parametros
'Descrição de Histórico não pode começar com este caracter.
Global Const ERRO_CODIGO_HISTPADRAO_INVALIDO = 188 'Parametro codigo do historico padrao(string)
'Após o asterisco deve ser digitado o código de um histórico padrão existente no sistema. O código digitado foi %s.
Global Const ERRO_LEITURA_CCL1 = 189 'Sem Parametros
'Erro na leitura da tabela de Centros de Custo/Lucro.
Global Const ERRO_CONTA_SEG_MEIO_NAO_PREENCHIDOS = 190 'Sem parametro
'Todos os segmentos da conta tem que estar preenchidos. Ex: 1.000.1 está errado. 1.001.1 está correto.
Global Const ERRO_CCL_SEG_MEIO_NAO_PREENCHIDOS = 191 'Sem parametro
'Todos os segmentos do centro de custo tem que estar preenchidos. Ex: 1.000.1 está errado. 1.001.1 está correto.
Global Const ERRO_CONFIGURACAO_NAO_CADASTRADA = 192 'Sem parametro
'Os dados de configuração não estão cadastrados
Global Const ERRO_LEITURA_EXERCICIO_DATA = 193 'Parametro Data(String)
'Não foi encontrado exercício para a data %s.
Global Const ERRO_LEITURA_EXERCICIO1 = 194 'Sem Parametro
'Erro de leitura da tabela de Exercicios.
Global Const ERRO_LEITURA_LOTEPENDENTE = 195 'Parametros FilialEmpresa, origem, exercicio, periodo, lote
'Erro na leitura do lote pendente. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i
Global Const ERRO_LOCK_LOTEPENDENTE = 196 'Parametros FilialEmpresa, origem, exercicio, periodo, lote
'Não foi possível fazer o "lock" do Lote Pendente que possui a chave: Filial = %i, Origem= %s, Exercicio= %i, Periodo=%i, Lote=%i.
Global Const ERRO_INSERCAO_LOTEPENDENTE = 197 'Parametros Filial, Origem, Exercicio, Periodo, Lote
'Ocorreu um erro ao tentar inserir o lote na tabela de lotes pendentes. Filial=%i, Origem=%s, Exercicio=%i, Periodo=%i, Lote=%i
Global Const ERRO_ATUALIZACAO_LOTEPENDENTE = 198 'Parametro FilialEmpresa, origem, exercicio, periodo, lote
'Erro na atualização do Lote Pendente. Filial=%i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Global Const ERRO_UNLOCK_LOTEPENDENTE = 199
'Não conseguiu liberar o lock do lote pendente.
Global Const ERRO_LEITURA_LOTEPENDENTE1 = 200 'Sem parametros
'Erro na leitura da tabela de lotes pendentes.
Global Const ERRO_LEITURA_LOTEPENDENTE2 = 201 'Parametros origem, exercicio, periodo, lote
'O Lote não está cadastrado como lote pendente - Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i
Global Const ERRO_LOTEPENDENTE_INEXISTENTE = 202  'Parametros FilialEmpresa, Origem, exercicio, periodo, lote
'O lote não está cadastrado na tabela de lotes pendentes.  Filial = %i, Origem= %s, Exercicio= %i, Periodo= %i, Lote= %i
Global Const ERRO_EXCLUSAO_LOTEPENDENTE = 203 'Parametros Filial, Origem, Exercicio, Periodo, Lote
'Houve um erro na exclusão do lote pendente do banco de dados. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i e Lote = %i.
Global Const ERRO_MASCARAR_CONTA = 204 'Parametro Conta
'Erro ao tentar mascarar a conta %s
Global Const ERRO_MASCARAR_CCL = 205 'Parametro Ccl
'Erro ao tentar mascarar o centro de custo/lucro %s
Global Const ERRO_FORMATAR_CONTA = 206 'Parametro Conta
'Erro na formatação da conta %s
Global Const ERRO_FORMATAR_CCL = 207 'Parametro Ccl
'Erro na formatação do centro de custo/lucro %s
Global Const ERRO_INICIALIZACAO_TELA = 208 'Sem parametro
'Erro na inicialização da tela %s
Global Const ERRO_DOCAUTO_NAO_CADASTRADO = 209 'Sem parametro
'Documento procurado não foi encontrado
Global Const ERRO_INSERCAO_DOCAUTO = 210 'Sem parametro
'Erro na atualizacao da tabela de Documento Automatico
Global Const ERRO_AUSENCIA_DOCAUTO_GRAVAR = 211 'Sem parametro
'O Grid está vazio, há ausência de dados para gravar
Global Const ERRO_NUMERO_RATEIO_NAO_PREENCHIDO = 212 'Sem Parametro
'O Código do Rateio não foi digitado
Global Const ERRO_AUSENCIA_RATEIOON_GRAVAR = 213 'Sem Parametro
'Não existem Lancamentos no Grid para gravar
Global Const ERRO_SOMA_NAO_VALIDA = 214  'Sem Parametro
'A Soma dos Rateios tem que totalizar 100%.
Global Const ERRO_INSERCAO_RATEIOON = 215  'Sem Parametro
'Erro na insercao de Registros na Tabela de Rateios OnLine
Global Const ERRO_RATEIOON_NAO_CADASTRADO = 216  'Parametro Codigo
'Não existe rateio cadastrado com o codigo %d
Global Const ERRO_EXCLUSAO_RATEIOON = 217  'Parametro Codigo
'Erro na exclusao do Rateio de codigo %d
Global Const ERRO_VALOR_PERCENTUAL = 218   'Sem Parametro
'O valor digitado como percentual de rateio deve estar entre 0 e 100
Global Const ERRO_EXISTENCIA_CONTA = 219 'Sem parametro
'Erro na verificacao da existencia de conta na tabela PlanoConta
Global Const ERRO_LEITURA_SEGMENTO = 221 'Sem parametro
'Erro na leitura da tabela Segmento.
Global Const ERRO_VALOR_FORMATO_NAO_PREENCHIDO = 222 'Sem parametro
'Campo formato não preenchido.
Global Const ERRO_VALOR_TIPO_NAO_PREENCHIDO = 223 'Sem parametro
'Campo tipo não preenchido.
Global Const ERRO_VALOR_TAMANHO_NAO_PREENCHIDO = 224 'Sem parametro
'Campo tamanho não preenchido.
Global Const ERRO_VALOR_PREENCHIMENTO_NAO_PREENCHIDO = 225 'Sem parametro
'Campo preenchimento não preenchido.
Global Const ERRO_VALOR_DELIMITADOR_NAO_PREENCHIDO = 226 'Sem parametro
'Campo delimitador não preenchido.
Global Const ERRO_SAIDA_DELIMITADOR = 227 'Sem parametro
'O delimitador não pode ter mais de um caracter.
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
'Houve um erro na exclusão do exercicio %i da tabela Exercicios.
Global Const ERRO_EXCLUSAO_PERIODO = 234 'Parametro Periodo, Exercicio
'Houve um erro na exclusão do periodo %i do exercicio %i da tabela de Periodos
Global Const ERRO_LEITURA_EXERCICIOORIGEM1 = 235 'Sem parametros
'Ocorreu um erro na leitura na tabela ExercicioOrigem.
Global Const ERRO_EXCLUSAO_EXERCICIOORIGEM = 236  'Parametros Filial, Exercicio, Periodo, Origem
'Erro na exclusão de registro na Tabela ExercicioOrigem. Filial = %i, Exercicio = %i, Periodo = %i, Origem=%s
Global Const Erro_Mascara_RetornaCcl = 237 'Parametro Conta
'Erro na função que retorna o centro de custo/lucro associado a conta contábil %s.
Global Const ERRO_LOTE_NAO_DESATUALIZADO = 238 'Parametros Filial, Origem, Exercicio, Periodo, Lote
'O Lote (Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i) não está pronto para ser atualizado. Verifique se este lote está incompleto ou já foi atualizado.
Global Const ERRO_LOTE_SENDO_ATUALIZADO = 239 'Filial, Origem, Exercicio, Periodo, Lote
'O Lote (Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i) já está em processo de atualização por outro programa.
'Global Const ERRO_LEITURA_ORCAMENTO = 240 'Sem parametros
'Erro na leitura da tabela Orcamento
Global Const ERRO_INSERCAO_ORCAMENTO = 241 'Sem parametros
'Erro na insercao de registro na tabela Orcamento
Global Const ERRO_EXCLUSAO_ORCAMENTO_CTB = 242 'Sem parametros
'Erro na exclusao de registro na tabela Orcamento
Global Const ERRO_MASCARA_RETORNACONTAENXUTA = 243 'Parametro Conta
'Erro na formatação da conta %s.
Global Const ERRO_CONTACREDITO_NAO_DIGITADA = 244
'Conta Credito não digitada
Global Const ERRO_CONTAORIGEM_NAO_DIGITADA = 245
'A Conta Origem nao foi digitada
Global Const ERRO_CCLORIGEM_NAO_DIGITADO = 246
'O CclOrigem não foi digitado
Global Const ERRO_RATEIOOFF_NAO_CADASTRADO = 247
'Esse Rateio não esta cadastrado
Global Const ERRO_EXCLUSAO_RATEIOOFF = 248 'Sem Parametros
'Erro na Exclusão de registros na tabela RateioOff.
Global Const ERRO_CONTA_JA_UTILIZADA = 249 'Parametro conta
'A conta %s já foi utiliazada
Global Const ERRO_INSERCAO_RATEIOOFF = 250
'Ocorreu um erro ao tentar inserir dados na tabela RateioOff
Global Const ERRO_AUSENCIA_RATEIOOFF_GRAVAR = 251
'Nao exite nenhuma linha de rateio no grid
Global Const ERRO_EXCLUSAO_RESULTADO = 252 'Parametro CodigoApuracao
'Ocorreu um erro na exclusão de um registro da tabela Resultado. Codigo de Apuracao = %l
Global Const ERRO_EXERCICIO_POSTERIOR_INEXISTENTE = 253 'Sem Parametro
'Para que ocorra o fechamento de um exercicio é necessário que o exercicio seguinte esteja criado. Favor criar o exercicio.
Public Const ERRO_DOC_ATUALIZADO = 254 'Parametros Filial, Origem, Exercicio, PeriodoLan, Doc
'Existe um documento com este número contabilizado. Filial = %i, Origem = %s Exercicio = %i Periodo = %i Documento = %l
Public Const ERRO_DOC_PENDENTE = 255 'Parametros Filial, Origem, Exercicio, PeriodoLan, Doc
'Existe um documento pendente com este número. Filial = %i, Origem = %s Exercicio = %s Periodo = %s Documento = %l
Public Const ERRO_CONTA_COM_MOVIMENTO = 256 'Parametro Conta, Ccl
'Não foi possível desfazer a associação da conta %s com o Centro de Custo %s, pois a conta possui movimentação.
Public Const ERRO_EXERCICIOS_FECHADOS = 257
'Todos os Exercicios estão fechados.
Public Const ERRO_CONTAS_SEM_PREENCHIMENTO = 258
'Nenhumas das Contas inicias estão Preenchidas
Public Const ERRO_CONTARESULTADO_NAO_PREENCHIDA = 259
'A conta Resultado não foi preenchida
Public Const ERRO_EXERCICIO_NAO_SELECIONADO = 260
'O exercicio nao foi selecionado.
Public Const ERRO_CONTARECEITA_INICIAL_MAIOR = 261
'A Conta Receita inicial é maior que a Conta Receita Final
Public Const ERRO_CONTADESPESA_INICIAL_MAIOR = 262
'A Conta Despesa Inicial é maior que a Conta Despesa Final
Public Const ERRO_INTERSECAO_CONJUNTO_CONTAS = 263
'As contas de Receita e despesa se interceptam
Public Const ERRO_INTERSECAO_CONTARESULTADO = 264
'A Conta Resultado faz parte do conjunto de contas envolvidos
Public Const ERRO_PERIODO_INICIAL_NAO_SELECIONADO = 265
'O Periodo Inicial não foi selecionado
Public Const ERRO_PERIODO_FINAL_NAO_SELECIONADO = 266
'O Periodo Final não foi selecionado
Public Const ERRO_PERIODO_INICIAL_MAIOR = 267
'O Periodo Inicial é maior que o final
Public Const ERRO_CONTACONTRAPARTIDA_VAZIA = 268
'A conta de ContraPartida está vazia
Public Const ERRO_GRID_VAZIO = 269
'Nao há contas no grid para Apurar
Public Const ERRO_CONTARESULTADO_VAZIA = 270 'Parametro ilinha
'A conta Resultado na %d linha nao foi preenchida
Public Const ERRO_CONTAINICIAL_VAZIA = 271 'Parametro ilinha
'A conta Inicial na %d linha nao foi informada
Public Const ERRO_CONTAFINAL_VAZIA = 272 'Parametro ilinha
'A conta Final da linha %d não foi informada
Public Const ERRO_CONTRAPARTIDA_IGUAL_RESULTADO = 273
'A conta de ContraPartida é igual a conta resultado
Public Const ERRO_INSERCAO_RESULTADO = 274
'Erro na insercao de registros na tabela de resultado
Public Const ERRO_INTERSECAO_CONTACONTRAPARTIDA = 275
'A conta de contraPartida intercepta um cojunto de contas
Public Const ERRO_CONTA_INICIAL_MAIOR = 276
'A conta Inicial é maior que a conta final
Public Const ERRO_EXERCICIO_NAO_PREENCHIDO = 277
'O Exercício não foi preenchido
Public Const ERRO_INTERSECAO_CONTAS = 278
'As contas de Ativo e Passivo se interceptam
Public Const ERRO_PASSIVAINICIAL_MAIOR = 279
'conta Passiva Inicial é maior que a conta Passiva Final
Public Const ERRO_ATIVAINICIAL_MAIOR = 280
'Conta Ativa Inicial é maior que a Ativa final
Public Const ERRO_TODOS_EXERCICIOS_FECHADOS = 281
'Todos os exercicios já estao Fechados
Public Const ERRO_FALTA_LOTE = 282
'Nao existe nenhum lote marcado para ser atualizado.
Public Const ERRO_MODIFICACAO_CONFIG = 283 'Parametro IdAtualizacao
'Erro na tentativa de modificar o campo IdAtualizacao na tabela LotePendente para %i .
Public Const ERRO_LEITURA_LOTE_PENDENTE = 284 'Parametro Origem, Exercicio, Periodo, Lote
'Erro na leitura dos campos Origem = %s , Exercicio = %i , Periodo = %i , Lote = %i da tabela LotePendente
Public Const ERRO_MODIFICACAO_LOTEPENDENTE = 285 'Parametro IdAtualizacao
'Erro na tentativa de modificar o campo IdAtualizacao = %i na tabela LotePendente.
Public Const ERRO_TODOS_EXERCICIOS_ABERTOS = 286  'Sem parametros
'Todos os Exercícios estão abertos.
Public Const ERRO_DATAS_COM_EXERCICIOS_DIFERENTES = 287
'Data Inicial e Final devem estar num mesmo exercício.
Public Const ERRO_LOTE_INICIAL_MAIOR = 288
'Lote inicial não pode ser maior que o lote final.
Public Const ERRO_DATA_INICIAL_MAIOR = 289
'Data inicial não pode ser maior que a data final.
Public Const ERRO_NOME_RELOP_VAZIO = 290
'O campo Opção de relatório tem que estar preenchido
Public Const ERRO_NOME_RELOP_NAO_SELEC = 291
'Não existe relatório selecionado.
Public Const ERRO_EXERCICIO_VAZIO = 292
'O campo Exercício tem que estar preenchido.
Public Const ERRO_PERIODO_VAZIO = 293
'O campo Período tem que estar preenchido.
Public Const ERRO_CCL_INICIAL_MAIOR = 294
'O Centro de Custo inicial não pode ser maior que o Centro de Custo final.
Public Const ERRO_LOTE_FORA_FAIXA = 295
'O lote final deve estar entre 1 e 9999.
Public Const ERRO_PERIODO_INICIAL_VAZIO = 296
'O período inicial tem que estar preenchido.
Public Const ERRO_PERIODO_FINAL_VAZIO = 297
'O período final tem que estar preenchido.
Public Const ERRO_LEITURA_CONTACATEGORIA = 298 'Parametro Codigo
'Erro na leitura da tabela Categoria. Categoria = %i.
Public Const ERRO_NOME_CATEGORIA_NAO_INFORMADO = 299 'Sem parametro
'O nome da categoria não foi informado.
Public Const ERRO_CATEGORIA_NAO_CADASTRADA = 300 'Parametro Codigo
'A Categoria %i não está cadastrada.
Public Const ERRO_LOCK_CONTACATEGORIA = 301 'Parametro Codigo
'Não conseguiu fazer o lock da Categoria %i.
Public Const ERRO_EXCLUSAO_CONTACATEGORIA = 302 'Parametro Codigo
'Houve um erro na exclusão da Categoria %i do banco de dados.
Public Const ERRO_CATEGORIA_PRESENTE_PLANO_CONTAS = 303 'Parametro Codigo
'Não é possível excluir a Categoria %i que está presente no Plano de Contas.
Public Const ERRO_ATUALIZACAO_CONTACATEGORIA = 304 'Parametro Codigo
'Erro de atualização da Categoria %i.
Public Const ERRO_INSERCAO_CONTACATEGORIA = 305 'Parametro Codigo
'Erro na inserção da Categoria %i na tabela ContaCategoria.
Public Const ERRO_CODIGO_CATEGORIA_NAO_INFORMADO = 306 'Sem parametro
'O codigo da categoria não foi informado.
Public Const ERRO_LEITURA_CONTACATEGORIA1 = 307 'Sem Parametros
'Erro na leitura da tabela de ContaCategoria.
Public Const ERRO_CONTA_NIVEL1_CATEGORIA = 308 'Sem Parametro
'A conta é de nivel 1. Favor designar uma categoria para esta conta.
Public Const ERRO_LEITURA_PLANOCONTA4 = 309 'Parametros Categoria, Nivel
'Erro na leitura do Plano de Contas. Categoria = %i Nivel = %i
Public Const ERRO_LEITURA_CONTACATEGORIA2 = 310 'Parametro Nome
'Erro na leitura da tabela Categoria. Categoria = %s.
Public Const ERRO_CATEGORIA_NAO_CADASTRADA1 = 311 'Parametro Nome
'A Categoria %s não está cadastrada.
Public Const ERRO_CONTA_CATEG_NIVEL_NAO_CADASTRADA = 312 'Parametro Codigo da Categoria, Nivel da Conta
'Não está cadastrado uma conta da categoria %i no nivel %i.
Public Const ERRO_LEITURA_CTBCONFIG = 313 'Parametro Codigo
'Erro na leitura da tabela CTBConfig. Codigo = %s.
Public Const ERRO_INTERSECAO_CONTARESULTADO_APURACAO = 314
'A Conta Resultado faz parte do conjunto de contas a serem apuradas
Public Const ERRO_INTERSECAO_CONTRAPARTIDA_APURACAO = 315
'A Conta de Contra Partida faz parte do conjunto de contas a serem apuradas
Public Const ERRO_MASCARA_RETORNAULTIMACONTA = 316 'Parametro Conta
'Erro ao tentar retornar a ultima conta do nivel da conta %s
Public Const ERRO_PLANOCONTA_SEM_CATEGORIA_ATIVO = 317 'Sem Parametro
'Não foi encontrado no plano de contas nenhum grupo designado com a categoria 'Ativo'.
Public Const ERRO_PLANOCONTA_SEM_CATEGORIA_PASSIVO = 318 'Sem Parametro
'Não foi encontrado no plano de contas nenhum grupo designado com a categoria 'Passivo'.
Public Const ERRO_CONTA_NAO_SELECIONADA = 319 'Sem parametro
'Nenhuma conta foi selecionada. Favor selecionar pelo menos uma conta antes de usar esta função.
Public Const ERRO_CCL_NAO_SELECIONADA = 320 'Sem parametro
'Nenhum Centro de custo/lucro foi selecionado. Favor selecionar um centro de custo/lucro antes de usar esta função.
Public Const ERRO_LEITURA_LANPENDENTE4 = 321 'Parametros Conta e Centro de Custo/Lucro
'Erro na leitura da tabela de Lançamentos Pendentes. Conta = %s e Centro de Custo/Lucro = %s.
Public Const ERRO_EXCLUSAO_CCLCONTA_COM_LANPENDENTE = 322 'Parametros Conta e Ccl
'Existem lançamentos pendentes para a associação da conta %s com o centro de custo/lucro %s. Portanto não pode-se excluir a associação.
Public Const ERRO_LEITURA_RATEIOON1 = 323 'Parametros conta e ccl.
'Erro de leitura da tabela de Rateios On-Line para a conta %s e centro de custo/lucro %s.
Public Const ERRO_EXCLUSAO_CCLCONTA_COM_RATEIOON = 324 'Parametros Conta e Ccl
'Existem rateios on-line para a associação da conta %s com o centro de custo/lucro %s. Portanto não pode-se excluir a associação.
Public Const ERRO_PLANOCONTA_SEM_CONTA_SINTETICA = 325 'Sem Parametro
'Não foi encontrado no plano de contas nenhuma conta sintética.
Public Const ERRO_CCL_VAZIO = 326 'Sem parametro
'Não há centro de custo/lucro cadastrado.
Public Const ERRO_PLANOCONTA_SEM_CONTA_ANALITICA = 327 'Sem Parametro
'Não foi encontrado no plano de contas nenhuma conta analítica.
Public Const ERRO_CONTACCL_VAZIO = 328 'Sem parametro
'Não há associação de conta com centro de custo/lucro cadastrada.
Public Const ERRO_CONTA_SEM_CONTACCL = 329 'Parametro Conta
'A conta %s não possui nenhuma associação com centro de custo/lucro.
Public Const ERRO_ATUALIZACAO_CONTACCL = 330 'Parametros Conta Ccl
'Erro na atualização da tabela que guarda a associação da Conta %s com o Centro de Custo/Lucro %s.
Public Const ERRO_SALDOS_INICIAIS_NAO_ALTERAVEIS = 331 'Sem Parametro
'Não é possível fazer atualização dos saldos iniciais de conta. Verifique se o exercício inicial (exercicio de implantação) está presente e aberto.
Public Const ERRO_CONTA_NAO_ANALITICA_SALDO = 332 'Parametro conta
'A conta %s não é analítica. Somente as contas analíticas podem receber saldos iniciais.
Public Const ERRO_LEITURA_MVPERCCL4 = 333  'parametros ccl, conta
'Ocorreu um erro de leitura na tabela de Saldos de Centro de Custo/Lucro (MvPerCcl). Centro de custo/lucro = %s e Conta = %s.
Public Const ERRO_LOCK_MVPERCCL1 = 334 'parametros  ccl, conta
'Ocorreu um erro ao tentar executar o "Lock" em um registro da tabela de saldos de centro de custo/lucro(MvPerCcl). Centro de Custo/Lucro = %s, Conta = %s.
Public Const ERRO_ATUALIZACAO_MVPERCCL2 = 335 'parametros ccl, conta
'Ocorreu um erro na atualização da tabela que guarda os saldos de centro de custo/lucro. Centro de Custo/Lucro = %s, Conta = %s.
Public Const ERRO_LEITURA_EXERCICIO2 = 336 'Parametro Exercicio
'Erro na leitura do exercicio %i
Public Const ERRO_EXERCICIO_NAO_CADASTRADO = 337 'Parametro Exercicio
'O Exercício %i não está cadastrado.
Public Const ERRO_DATA_INICIO_PERIODO_VAZIA = 338 'parametros Periodo
'A data de inicio do periodo %i está vazia. Favor preenche-la.
Public Const ERRO_NOME_PERIODO_VAZIO = 339 'Parametro Periodo
'O Nome do período %i está vazio. Favor preenche-lo.
Public Const ERRO_CONTA_NAO_ANALITICA1 = 340 'Parametro conta
'A conta %s não é analítica. Somente contas analiticas podem ser utilizadas.
Public Const ERRO_DATAINICIO_EXERCICIO_ALTERADA = 341 'Parametros Data Inicial Antiga, Data Inicial Nova
'A data inicial do exercicio nao pode ser modificada. Data Antiga = %s , Data Nova = %s.
Public Const ERRO_DATAFIM_EXERCICIO_ALTERADA = 342 'Parametros Data Final Antiga, Data Final Nova
'A data final do exercicio nao pode ser modificada. Data Antiga = %s , Data Nova = %s.
Public Const ERRO_NUMERO_PERIODOS_ALTERADO = 343 'Parametros Numero de Periodos Antigo, Numero de Periodos Novo
'O Exercicio possui movimento. O número de periodos não pode ser alterado. Num.Periodos Antigo = %i , Num.Periodos Novo = %i.
Public Const ERRO_LEITURA_PERIODO2 = 344 'Parametro Exercicio
'Ocorreu um erro na leitura dos Periodos do Exercicio %i.
Public Const ERRO_DATAINICIO_NOVO_EXERCICIO = 345 'Parametros Data Inicial Correta, Data Inicial Digitada
'A data inicial de um novo exercicio deve ser a data final do ultimo exercicio. Data Correta = %s , Data Digitada = %s
Public Const ERRO_NOME_EXERCICIO_JA_USADO = 346 'Parametro NomeExterno, Exercicio
'O Nome de exercício %s já foi usado pelo exercício %i. Favor escolher outro nome.
Public Const ERRO_LEITURA_ORCAMENTO1 = 347 'Parametro Exercicio
'Erro na leitura da tabela Orcamento. Exercicio = %i.
Public Const ERRO_LEITURA_LANPENDENTE5 = 348 'Parametro Exercicio
'Erro na leitura da tabela de Lançamentos Pendentes. Exercicio = %i.
Public Const ERRO_LEITURA_LOTEPENDENTE3 = 349 'Parametro Exercicio
'Erro na leitura da tabela de Lotes Pendentes. Exercicio = %i.
Public Const ERRO_EXERCICIO_NAO_ULTIMO = 350 'Parametro Exercicio
'Somente o ultimo exercicio pode ser excluido. Ultimo Exercicio = %i.
Public Const ERRO_EXERCICIO_COM_MOVIMENTO = 351 'Exercicio
'O Exercício %i possui movimento contábil associado ou não está aberto.
Public Const ERRO_EXERCICIO_NAO_ENCONTRADO_TELA = 352 'Exercicio
'O Exercicio %i não foi encontrado entre os listados nesta tela. O exercicio pode não estar cadastrada ou a tela estar desatualizada.
Public Const ERRO_DATA_FINAL_EXERCICIO_MENOR = 353 'Parametros Data Final e Data Inicial do Exercicio
'A data final do Exercicio %s é menor que a inicial %s.
Public Const ERRO_DATA_INICIAL_EXERCICIO_MAIOR = 354 'Parametros Data Inicial e Data Final do Exercicio
'A data inicial do Exercicio %s é maior que a final %s.
Public Const ERRO_DATA_INICIAL_EXERCICIO_NAO_PREENCHIDA = 355 'Sem parametros
'A data inicial do Exercicio não foi preenchida.
Public Const ERRO_DATA_FINAL_EXERCICIO_NAO_PREENCHIDA = 356 'Sem parametros
'A data final do Exercicio não foi preenchida.
Public Const ERRO_PERIODICIDADE_INVALIDA = 357 'iPeriodicidade
'A Periodicidade %i é inválida.
Public Const ERRO_NUM_PERIODO_INVALIDO = 358 'parametro Maximo de Periodos permitido pelo sistema
'Número do período inválido. Faixa válida: 1 a %i.
Public Const ERRO_TOTAL_PERIODOS_MAIOR_TOTAL_DIAS = 359 'Parametros: Numero de Periodos, Total de Dias do Exercicio
'O número de períodos requeridos = %i é maior do que o total de dias do Exercicio = %i
Public Const ERRO_NOME_PERIODO_JA_USADO = 360 'Parametros: Nome do Periodo, Periodo
'O Nome de Período %s já foi utilizado pelo período %i.
Public Const ERRO_DATA_FORA_EXERCICIO = 361 'Parametros Data Inicial do Periodo , Data Inicial do Exercicio e Data Final do Exercicio
'A Data Inicial deste periodo %s não está dentro das faixa abrangida pelo exercício. Data Inicial = %s e Data Final = %s.
Public Const ERRO_DATAINI_PERIODO_MENOR_PERIODO_ANT = 362 'Parametros Data Inicio Periodo e Data Inicio Periodo Anterior
'A data inicial de cada periodo tem que ser maior que a data inicial do periodo anterior. Data Inicio Periodo = %s e Data Inicio Periodo Anterior = %s.
Public Const ERRO_DATAINI_PRIMEIRO_PERIODO = 363
'A data inicial do primeiro período deve ser igual a data de início do exercício. Data Inicial do Primeiro Periodo = %s e Data Inicial do Exercicio = %s
Public Const ERRO_EXERCICIO_SEM_PERIODO = 364 'Parametro Exercicio
'Nao foram especificados periodos para o exercício %i.
Public Const ERRO_PERIODOS_DEMAIS = 365 'Parametros Total de Periodos, Maximo de Periodos do Sistema
'O número de periodos que seriam gerados automaticamente, %i,  ultrapassa o limite do sistema, %i.
Public Const ERRO_STATUS_EXERCICIO_INVALIDO = 366 'Parametro Status do Exercicio
'Status de exercício igual a %i é inválido
Public Const ERRO_NOME_EXERCICIO_VAZIO = 367
'O Nome do Exercício não foi preenchido.
Public Const ERRO_DATAINI_PERIODO_MAIOR_PERIODO_SEG = 368 'Parametros Data Inicio Periodo e Data Inicio Periodo Seguinte
'A data inicial de cada periodo tem que ser menor do que a data inicial do periodo seguinte. Data Inicio Periodo = %s e Data Inicio Periodo Seguinte = %s.
Public Const ERRO_LANCAMENTOS_PERIODO_FECHADO = 369 'Parametro Periodo, Exercicio
'O Periodo %i do Exercício %i está fechado. Não é possível fazer gravação ou exclusão de lançamentos.
Public Const ERRO_LEITURA_PERIODOSFILIAL = 370 'Parametros Filial, Exercicio, Periodo
'Ocorreu um erro na leitura da tabela PeriodosFilial. Filial = %i, Exercicio = %i e Periodo = %i.
Public Const ERRO_LOCK_PERIODOSFILIAL = 371 'Parametros Filial, Exercicio, Periodo
'Ocorreu um erro ao tentar executar o "Lock" em um registro da tabela PeriodosFilial. Filial = %i, Exercicio = %i e Periodo = %i.
Public Const ERRO_LEITURA_EXERCICIOSFILIAL = 372 'Parametros Filial, Exercicio
'Ocorreu um erro na leitura da tabela ExerciciosFilial. Filial = %i e Exercicio = %i.
Public Const ERRO_LOCK_EXERCICIOSFILIAL = 373 'Parametros Filial, Exercicio
'Ocorreu um erro ao tentar executar o "Lock" em um registro da tabela ExerciciosFilial. Filial = %i e Exercicio = %i.
Public Const ERRO_ATUALIZACAO_EXERCICIOSFILIAL = 374 'Parametro Exercicio, Filial
'Ocorreu um erro na atualização do exercicio %i da Filial %i.
Public Const ERRO_LEITURA_PERIODOSFILIAL1 = 375 'Sem Parametros
'Ocorreu um erro na leitura da tabela PeriodosFilial.
Public Const ERRO_EXCLUSAO_PERIODOSFILIAL = 376 'Parametros Filial,Exercicio, Periodo
'Houve um erro na exclusão de um registro da tabela PeriodosFilial. Filial = %i, Exercicio = %i e Periodo = %i.
Public Const ERRO_LEITURA_PERIODOSFILIAL2 = 377 'Parametros Filial, Exercicio
'Ocorreu um erro na leitura da tabela PeriodosFilial. Filial = %i e Exercicio = %i.
Public Const ERRO_ATUALIZACAO_PERIODOSFILIAL = 378 'Parametros Periodo, Exercicio, Filial
'Ocorreu um erro na atualização do Periodo %i do Exercicio %i da Filial %i.
Public Const ERRO_LEITURA_SALDOINICIALCONTA = 379 'Parametros Filial, Conta
'Ocorreu um erro na leitura da tabela SaldoInicialConta. Filial = %i e Conta = %s.
Public Const ERRO_LEITURA_SALDOINICIALCONTACCL = 380 'Parametros Filial, Conta, Ccl
'Ocorreu um erro na leitura da tabela SaldoInicialContaCcl. Filial = %i, Conta = %s e Centro de Custo/Lucro = %s.
Public Const ERRO_INSERCAO_PERIODOSFILIAL = 381 'Parametros Filial, Exercicio, Periodo
'Ocorreu um erro ao tentar inserir um registro na tabela PeriodosFilial. Filial = %i, Exercicio = %i e Periodo = %i.
Public Const ERRO_INSERCAO_EXERCICIOSFILIAL = 382 'Parametros Filial, Exercicio
'Ocorreu um erro ao tentar inserir um registro na tabela ExerciciosFilial. Filial = %i e Exercicio = %i.
Public Const ERRO_EXERCICIO_FECHADO1 = 383 'Parametro Exercicio.
'O Exercicio %i encontra-se Fechado. Portanto não pode ter seus dados alterados.
Public Const ERRO_LEITURA_SALDOINICIALCONTACCL1 = 384 'Parametros Conta, Ccl
'Ocorreu um erro na leitura da tabela SaldoInicialContaCcl. Conta = %s e Centro de Custo/Lucro = %s.
Public Const ERRO_EXCLUSAO_SALDOINICIALCONTACCL = 385 'Parametros Filial, Conta, Ccl
'Houve um erro na exclusão de um registro da tabela SaldoInicialContaCcl. Filial = %i, Conta = %s e Centro de Custo/Lucro = %s.
Public Const ERRO_LEITURA_SALDOINICIALCONTA1 = 386 'Parametro Conta
'Ocorreu um erro na leitura da tabela SaldoInicialConta. Conta = %s.
Public Const ERRO_EXCLUSAO_SALDOINICIALCONTA = 387 'Parametros Filial, Conta
'Houve um erro na exclusão de um registro da tabela SaldoInicialConta. Filial = %i e Conta = %s.
Public Const ERRO_ATUALIZACAO_SALDOINICIALCONTA = 388 'Parametros Filial, Conta
'Ocorreu um erro na atualização da tabela SaldoInicialConta. Filial = %i e Conta = %s.
Public Const ERRO_MASCARA_CCL_OBTERNIVEL = 389 'Parametro Ccl
'Erro na obtençao do nível do centro de custo/lucro. Centro de Custo/Lucro = %s.
Public Const Erro_Mascara_RetornaCclNoNivel = 390 'Parametros Ccl, Nivel
'Erro na obtençao do centro de custo/lucro %s no nível %i.
Public Const ERRO_LEITURA_SALDOINICIALCONTACCL2 = 391 'Parametros Filial, Ccl
'Ocorreu um erro na leitura da tabela SaldoInicialContaCcl. Filial = %i e Centro de Custo/Lucro = %s.
Public Const ERRO_ATUALIZACAO_SALDOINICIALCONTACCL = 392 'Parametros Filial, Ccl
'Ocorreu um erro na atualização da tabela SaldoInicialContaCcl. Filial = %i e Centro de Custo/Lucro = %s.
Public Const ERRO_ATUALIZACAO_MVPERCCL3 = 393 'parametros Filial, exercicio, ccl
'Ocorreu um erro na atualização da tabela que guarda os saldos de centro de custo/lucro. Filial= %i, Exercicio = %i, Centro de Custo/Lucro = %s
Public Const ERRO_LOCK_SALDOINICIALCONTACCL = 394 'Parametros Filial, Conta, Ccl
'Ocorreu um erro ao tentar executar o "Lock" em um registro da tabela SaldoInicialContaCcl. Filial = %i, Conta = %s e Centro de Custo/Lucro = %s.
Public Const ERRO_ATUALIZACAO_SALDOINICIALCONTACCL1 = 395 'Parametros Filial, Conta, Ccl
'Ocorreu um erro na atualização da tabela SaldoInicialContaCcl. Filial = %i, Conta = %s e Centro de Custo/Lucro = %s.
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
'Ocorreu um erro na exclusão de um registro da tabela Segmento. Codigo = %s e Nivel = %i.
Public Const ERRO_INSERCAO_SEGMENTO = 402  'Parametros Codigo, Nivel
'Erro na insercao de Registro na Tabela Segmento. Codigo = %s e Nivel = %i.
Public Const Erro_Mascara_RetornaCclPai = 403 'Parametro Ccl
'Erro na função que retorna a centro de custo/lucro de nivel imediatamente superior do centro de custo/lucro %s.
Public Const ERRO_LEITURA_SALDOINICIALCONTACCL4 = 404 'Parametros Ccl
'Ocorreu um erro na leitura da tabela SaldoInicialContaCcl. Centro de Custo/Lucro = %s.
Public Const ERRO_EXCLUSAO_ORCAMENTO1 = 405 'Parametros Exercicio, Conta
'Ocorreu um erro na exclusao de um orçamento. Exercicio = %i, Periodo = %i e Conta = %s.
Public Const ERRO_LEITURA_ORCAMENTO2 = 406 'Parametro Conta
'Erro na leitura da tabela Orcamento. Conta = %s.
Public Const ERRO_LEITURA_ORCAMENTO3 = 407 'Parametro Ccl
'Erro na leitura da tabela Orcamento. Centro de Custo/Lucro = %s.
Public Const ERRO_EXCLUSAO_CONTA_COM_MOVIMENTO1 = 408 'Parametro Conta
'A conta %s possui movimento, portanto não pode ser excluida.
Public Const ERRO_EXCLUSAO_CCL_COM_MOVIMENTO = 409 'Parametro Ccl
'O Centro de Custo/Lucro %s possui movimento, portanto não pode ser excluido.
Public Const ERRO_CCLPAI_INEXISTENTE = 410 'Sem Parametro
'O Centro de Custo/Lucro em questão não tem Centro de Custo/Lucro "pai".
Public Const ERRO_CCLPAI_ANALITICA = 411 'Sem Parametro
'O Centro de Custo/Lucro em questão possui um Centro de Custo/Lucro "pai" analítico. Centros de Custo/Lucro analíticos não podem conter Centros de Custo/Lucro embaixo dele.
Public Const ERRO_CCL_SINTETICA_ASSOCIADA_CONTA = 412 'parametro Ccl
'O Centro de Custo/Lucro %s não pode ser sintético pois está associado a alguma conta.
Public Const ERRO_CCL_ANALITICA_COM_FILHOS = 413 'Parametro Ccl
'O Centro de Custo/Lucro %s não pode ser analítico pois possui Centro de Custo/Lucro embaixo dele.
Public Const ERRO_MASCARA_RETORNACCLENXUTA = 414 'Parametro Ccl
'Erro na formatação do Centro de Custo/Lucro %s.
Public Const ERRO_CCL_NAO_ANALITICA = 415 'Parametro Ccl
'O Centro de Custo/Lucro %s não é analítico, portanto não pode ter associações com contas.
Public Const ERRO_CCL_NAO_ANALITICA1 = 416 'Parametro Ccl
'O Centro de Custo/Lucro %s não é analítico.
Public Const ERRO_UNLOCK_PERIODOSFILIAL = 417 'Parametros Filial, Exercicio, Periodo
'Ocorreu um erro ao tentar liberar o "Lock" em um registro da tabela PeriodosFilial. Filial = %i, Exercicio = %i e Periodo = %i.
Public Const ERRO_EXERCICIOSFILIAL_NAO_APURADO = 418 'Parametros Filial, Exercicio
'A Filial %i não está com o exercicio %i apurado.
Public Const ERRO_LEITURA_EXERCICIOSFILIAL2 = 419 'Sem Parametro
'Ocorreu um erro na leitura da tabela ExerciciosFilial.
Public Const ERRO_EXCLUSAO_EXERCICIOSFILIAL = 420 'Parametros Filial,Exercicio
'Houve um erro na exclusão de um registro da tabela ExerciciosFilial. Filial = %i e Exercicio = %i.
Public Const ERRO_PERCENTUAL_INVALIDO = 421
'Valor de Percentual Inválido
Public Const ERRO_SOMA_PERCENTUAL_NAO_VALIDA = 422
' O Total dos percentuais nao totalizou 100%
Public Const ERRO_DELIMITADOR_INVALIDO = 423 'Sem Parametros
'O Delimitador digitado não é valido.
Public Const ERRO_LEITURA_EXERCICIOSFILIAL1 = 424 'Parametro Exercicio
'Ocorreu um erro na leitura da tabela ExerciciosFilial. Exercicio = %i.
Global Const ERRO_ORCAMENTO_NAO_CADASTRADO = 425 'Parametros Exercicio, Conta.
'Orçamento não cadastrado. Exercicio = %i, Conta = %s.
Public Const ERRO_LANCAMENTO_INEXISTENTE = 426 'Parâmetros: iFilialEmpresa, sOrigem, iExercicio, iPeriodoLan, lDoc
'Não existe Lançamentos cadastrados com Filial Empresa %i, Origem %s, Exercício %i, Período de lançamento %i e número de documento %l.
Public Const ERRO_LEITURA_EXERCICIOS1 = 427 'Parâmetro: sNomeExterno
'Erro de leitura na tabela de Exercícios com nome externo %s.
Public Const ERRO_LEITURA_PERIODO3 = 428 'Parâmetros: iExercicio, sNomeExterno
'Erro de leitura na tabela de Período.
Public Const ERRO_EXERCICIO_INEXISTENTE = 429 'Parâmetro: sNomeExterno
'O Exercício %s não está cadastrado na tabela de Exercicios.
Public Const ERRO_PERIODO_EXERCICIO_INEXISTENTE = 430 'Parâmetros: iExercicio, sNomeExterno
'O Período %s do Exercício com código %i não está cadastrado na tabela de Periodo.
Public Const ERRO_MAX_ARGS_BATH = 431 'Sem parâmetros
'O número de argumentos ultrapassou o número máximo de argumentos
Public Const ERRO_EXERCICIOSFILIAL_INEXISTENTE = 432 'Parâmetro: iExercicio, iFilialEmpresa
'O Exercício %i da Filial %i não está cadastrado na tabela de ExerciciosFilial.
Global Const ERRO_TIPOCCL_NAO_INFORMADO = 433 'Sem parametro
'O Tipo do Centro de Custo/Lucro não foi informado.
Public Const Erro_MascaraCcl = 434 'Sem Parametro
'Erro na função que retorna a mascara de centro de custo/lucro.
Public Const ERRO_LOTEAPURACAO_JA_UTILIZADO = 435  'Parametros iLote, iExercicio
'O Lote de Apuração %i do Exercicio %i já foi utilizado.
Public Const ERRO_LEITURA_PADRAOCONTABITEM = 436 'Parametros sModulo, iTransacao, sModelo
'Ocorreu um erro na leitura da tabela PadraoContabItem. Modulo = %s, Transação = %i, Modelo = %s.
Public Const ERRO_LEITURA_PADRAOCONTAB = 437 'Parametros sModulo, iTransacao
'Ocorreu um erro na leitura da tabela PadraoContab. Modulo = %s, Transação = %i.
Public Const ERRO_PADRAOCONTAB_SEM_MODELO_PADRAO = 438 'Parametros sModulo, iTransacao
'O Módulo = %s, Transação = %i não possui um modelo de contabilização padrão.
Public Const ERRO_LEITURA_MNEMONICOCTB = 439 'Parametros sModulo, iTransacao
'Ocorreu um erro na leitura da tabela MnemonicoCTB. Modulo = %s, Transação = %i.
Public Const ERRO_GRID_NAO_ENCONTRADO = 440 'Parametros sNomeGrid
'O Grid %s não foi encontrado.
Public Const ERRO_LEITURA_MNEMONICOCTB1 = 441 'Parametros sModulo, iTransacao, sMnemonicoCombo
'Ocorreu um erro na leitura da tabela MnemonicoCTB. Modulo = %s, Transação = %i, Mnemonico = %s.
Public Const ERRO_LEITURA_FORMULAFUNCAO = 442 'Sem Parametros
'Ocorreu um erro na leitura da tabela FormulaFuncao.
Public Const ERRO_LEITURA_FORMULAFUNCAO1 = 443 'Parametro sFuncaoCombo
'Ocorreu um erro na leitura da tabela FormulaFuncao. Função = %s.
Public Const ERRO_LEITURA_FORMULAOPERADOR = 444 'Sem Parametros
'Ocorreu um erro na leitura da tabela FormulaOperador.
Public Const ERRO_LEITURA_FORMULAOPERADOR1 = 445 'Parametro sOperadorCombo
'Ocorreu um erro na leitura da tabela FormulaOperador. Operador = %s.
Public Const ERRO_LEITURA_TRANSACAOCTB = 446 'Parametro sSiglaModulo
'Ocorreu um erro na leitura da tabela TransacaoCTB. Sigla do Modulo = %s.
Public Const ERRO_TIPO_FORMULA_INVALIDA = 447 'Parametro sFormula, sTipoFormula, sTipoEsperado
'O tipo retornado pela formula %s foi %s e deveria ser %s.
Public Const ERRO_LEITURA_PADRAOCONTAB1 = 448 'Parametros sModulo, iTransacao, sModelo
'Ocorreu um erro na leitura da tabela PadraoContab. Modulo = %s, Transação = %i, Modelo = %s.
Public Const ERRO_CREDITO_DEBITO_PREENCHIDOS = 449 'Parametro iLinha
'As fórmulas de crédito e débito estão preenchidas na linha %i. Retire uma das duas.
Global Const ERRO_MODELO_NAO_PREENCHIDO = 450
'O Modelo não foi preenchido.
Public Const ERRO_ATUALIZACAO_PADRAOCONTAB = 451 'Parametros sModulo, iTransacao, sModelo
'Ocorreu um erro na atualização da tabela PadraoContab. Modulo = %s, Transação = %i, Modelo = %s.
Public Const ERRO_LOCK_PADRAOCONTAB = 452 'Parametros sModulo, iTransacao, sModelo
'Ocorreu um erro no lock de um registro da tabela PadraoContab. Modulo = %s, Transação = %i, Modelo = %s.
Public Const ERRO_INCLUSAO_PADRAOCONTAB = 453 'Parametros sModulo, iTransacao, sModelo
'Ocorreu um erro na inclusão de um registro na tabela PadraoContab. Modulo = %s, Transação = %i, Modelo = %s.
Public Const ERRO_EXCLUSAO_PADRAOCONTABITEM = 454 'Parametros sModulo, iTransacao, sModelo, iItem
'Ocorreu um erro na exclusão de um registro da tabela PadraoContabItem. Modulo = %s, Transação = %i, Modelo = %s, Item = %i.
Public Const ERRO_PADRAOCONTAB_INEXISTENTE = 455 'Parametros sModulo, iTransacao, sModelo
'Este modelo de contabilização não está cadastrado. Modulo = %s, Transação = %i, Modelo = %s.
Public Const ERRO_EXCLUSAO_PADRAOCONTAB = 456 'Parametros sModulo, iTransacao, sModelo
'Ocorreu um erro na exclusão de um registro da tabela PadraoContab. Modulo = %s, Transação = %i, Modelo = %s.
Public Const ERRO_INCLUSAO_PADRAOCONTABITEM = 457  'Parametros sModulo, iTransacao, sModelo, iItem
'Ocorreu um erro na inserção de um registro na tabela PadraoContabItem. Modulo = %s, Transação = %i, Modelo = %s, Item = %i.
Public Const ERRO_MODELO_CONTAB_SEM_PADRAO = 458 'Parametros sModulo, iTransacao.
'Não há um modelo padrão cadastrado. Modulo = %s, Transação = %i.
Public Const ERRO_CALCULO_MNEMONICO_INEXISTENTE = 459 'Parametro sMnemonico
'Não foi encontrada a função que calcula o valor do campo %s.
Public Const ERRO_NAO_HA_LOTE_PENDENTE = 460 'Parâmetro: iFilialEmpresa
'Não existe nenhum Lote Pendente da Filial %i.
Public Const ERRO_VALOR_NAO_PREENCHIDO = 461
'O Valor do Rateio não Preenchido
Public Const ERRO_LEITURA_TRANSACAOCTB1 = 462 'Parametros sSiglaModulo, sTransacao
'Ocorreu um erro na leitura da tabela TransacaoCTB. Sigla do Modulo = %s, Transação = %s.
Public Const ERRO_DATA_CONTABIL_NAO_PREENCHIDA = 463 'Sem Parametro
'Data Contábil não preenchida.
Public Const ERRO_DOCUMENTO_CONTABIL_NAO_PREENCHIDO = 464 'Sem Parametro
'O Número do Documento Contábil não foi preenchido.
Public Const ERRO_VALOR_LANCAMENTO_CONTABIL_NAO_PREENCHIDO = 465 'Sem Parametro
'O Valor do Lançamento Contábil não foi preenchido.
Public Const ERRO_LANCAMENTOS_CONTABILIZADOS = 466 'Sem parametro
'Atenção. Já existem lançamentos atualizados para o documento em questão.
Public Const ERRO_ALTERACAO_LANCAMENTO_AGLUTINADO = 467 'Parametros iFilialEmpresa, sOrigem, iexercicio, iperiodo, lDoc, iSeq
'Ocorreu um erro na alteração do lançamento aglutinado. Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Doc = %l, Seq = %i.
Public Const ERRO_LEITURA_LANCAMENTOS5 = 468 'Parametros FilialEmpresa, Origem, Data
'Erro na leitura da tabela de Lançamentos. Filial = %i, Origem = %s, Data = %s.
Public Const ERRO_LEITURA_LOTEPENDENTE4 = 469 'Parametros Filial, origem, exercicio, periodo
'O Lote não está cadastrado como lote pendente ou está em processo de atualização - Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i
Public Const ERRO_FORMULA_CONTA_NAO_PREENCHIDA = 470 'Sem Parametros
'A fórmula da conta não foi preenchida.
Public Const ERRO_FORMULA_DEBCRE_NAO_PREENCHIDA = 471 'Sem Parametros
'As fórmulas de débito e crédito não foram preenchidas.
Public Const ERRO_FORMULA_PRODUTO_NAO_PREENCHIDA = 472 'Sem Parametros
'A fórmula de produtonão foi preenchida.
Public Const ERRO_QUANTIDADE_PRODUTO_DESBALANCEADO = 473 'Parametro sProduto
'Os lançamentos de custo envolvendo o produto %s estão desbalanceados.
Public Const ERRO_LANC_CUSTO_CONTA_NAO_INFORMADA = 474 'Sem parametro
'Ocorreu um erro na geração dos lançamentos de custo. A conta de um dos lançamentos não foi preenchida.
Public Const ERRO_LANC_CUSTO_QUANT_NAO_PREENCHIDA = 475 'Sem parametros
'Ocorreu um erro na geração dos lançamentos de custo. Os campos de débito e crédito de um dos lançamentos não foram preenchidos.
Public Const ERRO_LANC_CUSTO_DEBCRE_PREENCHIDOS = 476 'Sem parametros
'Ocorreu um erro na geração dos lançamentos de custo. Os campos de débito e crédito de um dos lançamentos estão ambos preenchidos.
Public Const ERRO_LEITURA_LANCAMENTOS6 = 477 'Parametros iFilialEmpresa, sOrigem, iexercicio, iperiodo, lDoc, iSeq
'Erro na leitura da tabela de Lançamentos Contábeis. Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Doc = %l, Seq = %i.
Public Const ERRO_LOCK_LANCAMENTOS = 478 'Parametros iFilialEmpresa, sOrigem, iexercicio, iperiodo, lDoc, iSeq
'Ocorreu um erro na tentativa de fazer um "lock" em um dos lançamentos contábeis. Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Doc = %l, Seq = %i.
Public Const ERRO_UNLOCK_LANCAMENTOS = 479 'Parametros iFilialEmpresa, sOrigem, iexercicio, iperiodo, lDoc, iSeq
'Ocorreu um erro na tentativa de liberar um "lock" em um dos lançamentos contábeis. Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Doc = %l, Seq = %i.
Public Const ERRO_PRODUTOSFILIAL_INEXISTENTE = 480   'Parametro: sProduto, iFilial
'O Produto %s da Filial %i não está cadastrado.
Public Const ERRO_MNEMONICO_INEXISTENTE = 481   'Parametro: sMnemonico
'O Campo %s não está cadastrado.
Public Const ERRO_LEITURA_SLDMESEST = 482 'Parametros  iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque (SldMesEst). Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_SLDMESEST_INEXISTENTE = 483 'Parametros  iAno, iFilialEmpresa, sProduto
'Não existe registro de saldos mensais de estoque (SldMesEst) com os dados a seguir. Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_SLDMESEST_CUSTORPRODUCAO_ZERADO = 484 'Parametros  iAno, iFilialEmpresa, sProduto, iMes
'Não é possível processar este lote já que o custo real de produção para o produto especificado não foi digitado. Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s, Mes=%i.
Public Const ERRO_SLDMESEST_CUSTOMRPRODUCAO_ZERADO = 485 'Parametros  iAno, iFilialEmpresa, sProduto, iMes
'Não é possível processar este lote já que o custo médio real de produção para o produto especificado não foi calculado. Existe um programa que calcula estes valores. Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s, Mes=%i.
Public Const ERRO_LEITURA_LANPREPENDENTE = 486 'Sem parametros
'Erro na leitura da tabela de Lançamentos Pré-Pendentes.
Public Const ERRO_EXCLUSAO_LANPREPENDENTE = 487 'Sem parametros
'Ocorreu um erro na exclusão de um lançamento pré-pendente.
Public Const ERRO_LOCK_SLDMESEST = 488 'Parametros iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos mensais de estoque (SldMesEst). Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_LANC_CONTA_NAO_INFORMADA = 489 'Sem parametro
'Ocorreu um erro na geração dos lançamentos contábeis. A conta de um dos lançamentos não foi preenchida.
Public Const ERRO_LANC_QUANT_NAO_PREENCHIDA = 490 'Sem parametros
'Ocorreu um erro na geração dos lançamentos contábeis. Os campos de débito e crédito de um dos lançamentos não foram preenchidos.
Public Const ERRO_LANC_DEBCRE_PREENCHIDOS = 491 'Sem parametros
'Ocorreu um erro na geração dos lançamentos contábeis. Os campos de débito e crédito de um dos lançamentos estão ambos preenchidos.
Public Const ERRO_LEITURA_RATEIOON2 = 492 'Sem Parametros
'Erro de leitura da tabela de Rateios On-Line.
Public Const ERRO_LEITURA_RATEIOOFF1 = 493 'Sem Parametro
'Erro de leitura da tabela de Rateios Off-Line.
Public Const ERRO_RATEIOOFF_CODIGO_NAO_PREENCHIDO = 494
'O Código do Rateio não foi informado.
Public Const ERRO_RATEIOOFF_BATCH = 495 'Parametro lCodigo
'Ocorreu um erro no processamento do Rateio %l.
Public Const ERRO_ATUALIZACAO_BATCH = 496 'Parametros iFilialEmpresa, sOrigem, iExercicio, iPeriodo, iLote
'Ocorreu um erro na Atualização de um Lote. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Public Const ERRO_RATEIOOFF_NAO_CADASTRADO1 = 497 'Parametro lCodigo
'O Rateio %l não está cadastrado.
Public Const ERRO_TIPO_RATEIOOFF_INVALIDO = 498 'Parametro iTipo
'O Tipo de Rateio Offline %i é inválido.
Public Const ERRO_LEITURA_LANCAMENTOS7 = 499 'Parametros Filial, Origem, exercicio, periodo, lote
'Erro na leitura dos Lançamentos Contábeis do lote que possui chave Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Lote= %i.
Public Const ERRO_LEITURA_LANCAMENTOS8 = 500 'Parametros iFilialEmpresa, sOrigem, iexercicio, iperiodo, lDoc
'Erro na leitura da tabela de Lançamentos Contábeis do Documento que possui chave Filial=%i, Origem= %s, Exercicio= %i, Periodo= %i, Doc = %l.
Public Const ERRO_LEITURA_CPCONFIG2 = 501 'Parâmetros: %s chave %d FilialEmpresa
'Erro na leitura da tabela CPConfig. Codigo = %s Filial = %i.
Public Const ERRO_CPCONFIG_INEXISTENTE = 502 'Parâmetros: %s chave %d FilialEmpresa
'Não foi encontrado registro em CPConfig. Codigo = %s Filial = %i.
Public Const ERRO_ATUALIZACAO_CPCONFIG = 503 'Parâmetros: %s chave %d FilialEmpresa
'Erro na gravação da tabela CPConfig. Codigo = %s Filial = %i.
Public Const ERRO_LEITURA_CPCONFIG = 504 'Sem parâmetros
'Erro na leitura da tabela CPConfig.
Public Const ERRO_LEITURA_CRCONFIG2 = 505 'Parâmetros: %s chave %d FilialEmpresa
'Erro na leitura da tabela CRConfig. Codigo = %s Filial = %i.
Public Const ERRO_CRCONFIG_INEXISTENTE = 506 'Parâmetros: %s chave %d FilialEmpresa
'Não foi encontrado registro em CRConfig. Codigo = %s Filial = %i.
Public Const ERRO_ATUALIZACAO_CRCONFIG = 507 'Parâmetros: %s chave %d FilialEmpresa
'Erro na gravação da tabela CRConfig. Codigo = %s Filial = %i.
Public Const ERRO_LEITURA_CRCONFIG = 508 'Sem parâmetros
'Erro na leitura da tabela CRConfig.
Public Const ERRO_LEITURA_TESCONFIG2 = 509 'Parâmetros: %s chave %d FilialEmpresa
'Erro na leitura da tabela TESConfig. Codigo = %s Filial = %i.
Public Const ERRO_TESCONFIG_INEXISTENTE = 510 'Parâmetros: %s chave %d FilialEmpresa
'Não foi encontrado registro em TESConfig. Codigo = %s Filial = %i.
Public Const ERRO_ATUALIZACAO_TESCONFIG = 511 'Parâmetros: %s chave %d FilialEmpresa
'Erro na gravação da tabela TESConfig. Codigo = %s Filial = %i.
Public Const ERRO_LEITURA_TESCONFIG = 512 'Sem parâmetros
'Erro na leitura da tabela TESConfig.
Public Const ERRO_CTBCONFIG_CHV_NAO_ENC = 513
'Chave %s não encontrada na tabela CTBConfig.
Public Const ERRO_CTBCONFIG_ATUALIZAR_CHV = 514
'Erro na atualização da tabela CTBConfig para a chave %s.
Public Const ERRO_LEITURA_ARQ_CONF_REL_DRE = 515
'Erro na leitura do arquivo de configuração do relatório.
Public Const ERRO_LEITURA_MVPERCTA_PLANOCONTA = 516
'Erro na leitura das tabelas MvPerCta e/ou PlanoConta.
Public Const ERRO_GRAVACAO_ARQ_TEMP_REL_DRE = 517
'Erro na gravação de arquivo temporário p/relatório (RelDRERes).
Public Const ERRO_CONFIG_DR = 518
'A configuração do demonstrativo de resultados não foi feita ou apresenta problemas.
Public Const ERRO_LEITURA_ARQ_CONF_REL_DRP = 519
'Erro na leitura do arquivo de configuração do relatório.
Public Const ERRO_GRAVACAO_ARQ_TEMP_REL_DRP = 520
'Erro na gravação de arquivo temporário p/relatório (RelDRPRes).
Public Const ERRO_RELATORIO_EXECUTANDO = 521
'Já existe um diálogo aberto para a execução deste relatório
Public Const ERRO_EXERCICIO1_VAZIO = 522
'O campo Exercício1 tem que estar preenchido.
Public Const ERRO_EXERCICIO2_VAZIO = 523
'O campo Exercício2 tem que estar preenchido.
Public Const ERRO_EXERCICIO1_MAIOR = 524
'O Exercicio 1 não pode ser maior que o Exercicio 2.
Public Const ERRO_LEITURA_RELDRE = 525 'Sem Parametro
'Erro na leitura da tabela RelDRE.
Public Const ERRO_RELDRE_VAZIA = 526 'Sem Parametro
'Erro a Tabela RelDre está vazia
Public Const ERRO_DOCUMENTO_INICIAL_MAIOR = 527 'Sem Parametros
'O Documento inicial não pode ser maior do que o documento final.
Public Const ERRO_ORIGEM_INICIAL_MAIOR = 528 'Sem Parametros
'A Origem inicial não pode ser maior do que a Origem final.
Public Const ERRO_CONTA_PATRIMONIOLIQUIDO_VAZIO = 529 'Sem Parametros
'A Conta de Patrimônio Líquido tem que estar preenchida.
Public Const ERRO_NO_NAO_SELECIONADO_INSERCAO_FILHO = 530 'Sem parametro.
'Escolha um elemento da árvore antes de tentar inserir um filho.
Public Const ERRO_FORMULA_NAO_PREENCHIDA = 531 'Parametro: iLinha
'A Fórmula da Linha %i não foi preenchida.
Public Const ERRO_OPERADOR_NAO_PREENCHIDO = 532 'Parametro: iLinha
'O Operador "Soma/Subtrai" da Linha %i não foi preenchido.
Public Const ERRO_FORMULA_INVALIDA = 533 'Parametro: sFormula, iLinha
'A Fórmula %s utilizada na Linha %i não é válida.
Public Const ERRO_FORMULA_INVALIDA1 = 534 'Parametro: sFormula
'A Fórmula %s não é válida.
Public Const ERRO_CONTA_INICIO_NAO_PREENCHIDA = 535 'Parametro: iLinha
'A Conta Início da Linha %i não foi preenchida.
Public Const ERRO_CONTA_FIM_NAO_PREENCHIDA = 536 'Parametro: iLinha
'A Conta Fim da Linha %i não foi preenchida.
Public Const ERRO_LEITURA_RELDRECONTA = 537 'Parâmetro Modelo
'Ocorreu um erro na leitura da tabela RelDREConta com o modelo %s .
Public Const ERRO_EXCLUSAO_RELDREFORMULA = 538 'Parâmetro Modelo
'Erro na exclusão do modelo %s da tabela RelDREFormula.
Public Const ERRO_EXCLUSAO_RELDRECONTA = 539 'Parâmetro Modelo
'Erro na exclusão do modelo %s da tabela RelDREConta.
Public Const ERRO_EXCLUSAO_RELDRE = 540 'Parâmetro Modelo
'Erro na exclusão do modelo %s da tabela RelDRE.
Public Const ERRO_MODELO_NAO_INFORMADO = 541
'O modelo precisa ser informado.
Public Const ERRO_NO_NAO_SELECIONADO_REMOVER = 542 'Sem parametro.
'Escolha um elemento da árvore antes de tentar remover.
Public Const ERRO_LEITURA_RELDRE_MODELO = 543 'Parametro: sModelo
'Ocorreu um erro na leitura do modelo %s tabela RelDRE.
Public Const ERRO_LEITURA_RELDREFORMULA = 544 'Parâmetro sModelo
'Ocorreu um erro na leitura da tabela RelDREFormula com o modelo %s .
Public Const ERRO_INCLUSAO_RELDRE = 545 'Parâmetro Modelo
'Ocorreu um erro na inclusão do modelo %s na tabela RelDRE.
Public Const ERRO_INCLUSAO_RELDRECONTA = 546 'Parâmetro Modelo
'Ocorreu um erro na inclusão do modelo %s na tabela RelDREConta.
Public Const ERRO_INCLUSAO_RELDREFORMULA = 547 'Parâmetro Modelo
'Ocorreu um erro na inclusão do modelo %s na tabela RelDREFormula.
Public Const ERRO_NO_NAO_SELECIONADO_MOV_ARV = 548 'Sem parametro.
'Selecione um elemento da árvore antes de tentar movimentá-lo.
Public Const ERRO_NO_SELECIONADO_NAO_MOV_ACIMA = 549 'Sem parametro.
'O elemento da árvore selecionado não pode ser movimentado para cima.
Public Const ERRO_NO_UTILIZA_NO_EM_FORMULA = 550 'Parametros: sTitulo, sTitulo1
'O elemento da árvore %s utiliza em sua fórmula o elemento %s e portanto este não pode ser movido.
Public Const ERRO_NO_SELECIONADO_NAO_MOV_ABAIXO = 551 'Sem parametro.
'O elemento da árvore selecionado não pode ser movimentado para baixo.
Public Const ERRO_LEITURA_RATEIOOFF2 = 552 'Parametros conta e ccl.
'Erro de leitura da tabela de Rateios Off-Line para a conta %s e centro de custo/lucro %s.
Public Const ERRO_LEITURA_MNEMONICOCTBVALOR = 553 'Parametro sMnemonico
'Ocorreu um erro na leitura da tabela MnemonicoCTBValor. Mnemonico = %s.
Public Const ERRO_DOCAUTO_NAO_CADASTRADO2 = 554 'Parametro: lDoc
'O  Documento %l não foi encontrado.
Public Const ERRO_INCLUSAO_LANPREPENDENTEBAIXADO = 555 'Parametros: iFilialEmpresa, sOrigem, iExercicio, iPeriodoLan, lDoc, iSeq
'Ocorreu um erro na inclusão de um lançamento na tabela de lançamentos pre-pendentes baixados. FilialEmpresa = %i, Origem =%s, Exercício = %i, Periodo = %i, Documento = %l, Sequencial = %i.
Public Const ERRO_EXCLUSAO_LANPREPENDENTE1 = 556 'Parametros: iFilialEmpresa, sOrigem, iExercicio, iPeriodoLan, lDoc, iSeq
'Ocorreu um erro na exclusão de um lançamento da tabela de lançamentos pre-pendentes. FilialEmpresa = %i, Origem =%s, Exercício = %i, Periodo = %i, Documento = %l, Sequencial = %i.
Public Const ERRO_CONTA_NAO_VISIVEL_MODULO = 557 'Sem Parametros
'A conta %s não está visível para o módulo em questão.
Public Const ERRO_LEITURA_TABELASDEPRECOITENS1 = 558 'Parametros: iCodTabela
'Erro na leitura da tabela TabelasDePrecoItens com Código da Tabela %i.
Public Const ERRO_ATUALIZACAO_EXERCICIOORIGEM_BATCH = 559 'sem parametros
'Ocorreu um erro na alteração de um registro da tabela ExercicioOrigem.
Public Const ERRO_LEITURA_CTBCONFIG2 = 560 'Parametros: Codigo, FilialEmpresa
'Erro na leitura da tabela CTBConfig. Codigo = %s, Filial = %i.
Public Const ERRO_INSERCAO_CTBCONFIG = 561 'Parâmetros: Codigo, FilialEmpresa
'Ocorreu um erro na inserção de um registro na tabela CTBConfig. Codigo = %s, Filial = %i.
Public Const ERRO_ATUALIZACAO_CTBCONFIG = 562 'Parâmetros: Codigo, FilialEmpresa
'Erro na gravação da tabela CTBConfig. Codigo = %s, Filial = %i.
Public Const ERRO_SEGMENTO_CONTA_INVALIDO = 563 'Parametro sCodigo
'Esperava o segmento conta e está tentando gravar o segmento %s.
Public Const ERRO_SEGMENTO_CCL_INVALIDO = 564 'Parametro sCodigo
'Esperava o segmento centro de custo e está tentando gravar o segmento %s.
Public Const ERRO_ALTERACAO_SEGMENTO = 565 'Parametros Codigo, Nivel
'Ocorreu um erro na alteração de um registro da tabela Segmento. Codigo = %s e Nivel = %i.
Public Const ERRO_LANPENDENTE_TRANSACAO_NUMINTDOC = 566 'Parametros iTransacaoLan, lNumIntDocLan, iTransacao, lNumIntDoc
'Foi encontrado um lançamento contábil pendente que não pertence ao documento em questão. Documento Encontrado: Transação = %i , NumIntDoc = %l; Documento em questão: Transação = %i , NumIntDoc = %l
Public Const ERRO_LANPREPENDENTE_JA_CADASTRADO = 567 'Parametros: iFilialEmpresa, sOrigem , iExercicio, iPeriodoLan , lDoc
'Foi encontrado um documento cadastrado com este mesmo número. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Documento = %l.
Public Const ERRO_ALTERACAO_CONFIGURACAO = 568 'Sem Parametros
'Erro na tentativa de alterar a configuração da contabilidade.
Public Const ERRO_LEITURA_PADRAOCONTAB_MODPADRAO = 569 'parametros: nome da transacao e sigla do modulo
'Erro na leitura do modelo padrão para a transação %s do módulo %s
Public Const ERRO_PADRAOCONTAB_SEMMODPADRAO = 570 'parametros: nome da transacao e sigla do modulo
'Não há modelo padrão para a transação %s do módulo %s
Public Const ERRO_LOTE_JA_CADASTRADO_OUTRO_PERIODO = 571 'Parametros Filial, origem, exercicio, periodo, lote
'Já existe um lote cadastrado com o mesmo número em outro periodo. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i.
Public Const ERRO_DOC_PENDENTE_OUTRO_PERIODO = 572 'Parametros Filial, Origem, Exercicio, Periodo, Doc
'Já existe um documento pendente com o mesmo número número em outro periodo. Filial = %i, Origem = %s, Exercicio = %s, Periodo = %s, Documento = %l.
Public Const ERRO_CONTRAPARTIDA_NAO_MESMO_LANCAMENTO = 573
'O Lançamento de Contra Partida tem que ser diferente do lançamento sendo editado.
Public Const ERRO_CONTRAPARTIDA_LANCAMENTO_INEXISTENTE = 574
'O Lançamento de Contra Partida tem que ser um lançamento existente.
Public Const ERRO_LANCAMENTO_CONTRA_PARTIDA_INEXISTENTE = 575 'Parametro iSeq
'O Lançamento a que a contra partida se refere não existe. Sequencial = %i.
Public Const ERRO_LANCAMENTO_CONTRA_PARTIDA_VALOR = 576 'Parametros iSeq, dValorLancamento, dValorContraPartida
'O Valor do Lançamento não é igual ao de sua contra partida. Sequencial = %i, Valor do Lançamento = %d, Total da Contra-Partida = %d.
Public Const ERRO_HISTORICO_PARAM = 577 'Sem Parametros
'O Histórico contém parametro(s) que ainda não foi(foram) substituido(s).
Public Const ERRO_EXCLUSAO_MNEMONICOCTBVALOR = 578 'Sem Parâmetros
'Erro na Exclusão da Tabela MnemonicoCTBValor
Public Const ERRO_INSERCAO_MNEMONICOCTBVALOR = 579 'Sem Parâmetros
'Erro na inserção de dados na Tabela MnemonicoCTBValor
Public Const ERRO_RATEIOS_NAO_INFORMADOS = 580
'Não foi informado nenhum Rateio para apuração.
Public Const ERRO_PERIODOFINAL_MENOR = 581
'O Periodo Final não pode ser menor do que o Periodo Inicial.
Public Const ERRO_PERIODOFINAL_MAIOR = 582
'O Periodo Final não pode ser maior que o periodo da data de contabilização.
Public Const ERRO_RATEIOS_INEXISTENTES = 583
'Não existem Rateios disponíveis para apuração.
Public Const ERRO_CODIGO_NAO_DIGITADO = 584
'O Código do Rateio não foi informado.
Public Const ERRO_RELDRE_MODELO_NAO_CADASTRADO = 585 'Sem parâmetros: sModelo, iCodigo
'O Modelo %s com Código %i do Demonstrativo de Resultado do Exercício não está cadastrado no Banco de Dados.
Public Const ERRO_LOCK_RELDRE = 586 'Parâmetros: sModelo, iCodigo
'Não foi possível fazer o "Lock" na tabela RelDRE com Modelo %s e Código %i.
Public Const ERRO_ATUALIZACAO_RELDRE = 587  'Parâmetros: sModelo, iCodigo
'Ocorreu um erro na atualização da tabela RelDRE. Modelo %s, Código %i.
Public Const ERRO_SEGMENTO_VAZIO = 588 'Sem parametros
'Preencha os segmentos de conta e centro de custo com pelo menos 1 segmento.
Public Const ERRO_SEGMENTO_PRODUTO_MAIOR_PERMITIDO = 589 'Parametros tamanho do segmento, tamanho total permitido
'O tamanho do segmento de produto %i ultrapassou o tamanho total permitido %i.
Public Const ERRO_DOC_PENDENTE_OUTRO_LOTE = 590 'Parametros iLote
'Este documento já está cadastrado no lote %i e não é possível alterar o lote.
Public Const ERRO_ORIGEM_NAO_PREENCHIDA1 = 591 'Sem parâmetros
'O preenchimento de Origem é obrigatório.
Public Const ERRO_GRID_DESCONTO_TIPODESCONTO_NAO_PRENCHIDO = 592 'Parametro iLinha
'O campo Tipo Desconto da Linha %i do Grid de Desconto não foi preechido.
Public Const ERRO_GRID_DESCONTO_DIAS_NAO_PRENCHIDO = 593 'Parametro iLinha
'O campo Dias da Linha %i do Grid de Desconto não foi preechido.
Public Const ERRO_GRID_DESCONTO_PERCENTUAL_NAO_PRENCHIDO = 594 'Parametro iLinha
'O campo Percentual da Linha %i do Grid de Desconto não foi preechido.
Public Const ERRO_GRID_DESCONTO_NAO_ORDEM_DECRESCENTE = 595
'Os Campos de Dias e Percentual tem que estar em Ordem Decrescente no Grid de Descontos.
Public Const ERRO_LOTEPENDENTE_NAO_CADASTRADO = 596 'Parametros: iFilialEmpresa, sOrigem, iExercicio, iPeriodo, iLote
'Lote Pendente não cadastrado.
Public Const ERRO_CONTA_JA_UTILIZADA_GRID_CONTAS = 597 'Parametro sconta
'A conta %s já foi utilizada no Grid de Contas.
Public Const ERRO_CONTA_JA_UTILIZADA_GRID_RATEIOS = 598 'Parametro sconta
'A conta %s já foi utilizada no Grid de Rateios.
Public Const ERRO_CONTA_JA_UTILIZADA_CONTA_CREDITO = 599 'Parametro sconta
'A conta %s já foi utilizada no Grid de Rateios.
Public Const ERRO_CONTAINICIO_GRIDCONTAS_NAO_INFORMADA = 600 'Parametro iLinha
'A Conta Início, localizada na linha %i do grid de Contas, não foi informada.
Public Const ERRO_CONTAFIM_GRIDCONTAS_NAO_INFORMADA = 601 'Parametro iLinha
'A Conta Fim, localizada na linha %i do grid de Contas, não foi informada.
Public Const ERRO_CONTAFIM_MENOR_CONTAINICIO = 602 'Parametro iLinha, sContaFim, sContaInicio.
'Na linha %i do Grid de Contas, a Conta Fim = %s é menor do que a Conta Início = %s.
Public Const ERRO_CONTA_GRID_NAO_PREENCHIDA = 603 'Parametro iLinha.
'A Conta da Linha %i do Grid não foi preenchida.
Public Const ERRO_PROXNUM_DATA_NAO_PREENCHIDA = 604 'Sem Parametro
'Para conseguir o próximo número de documento, a data tem que estar preenchida.
Public Const ERRO_DATA_ULT_PERIODO_FORA_EXERCICIO = 605 'Parametros Data Inicial do Periodo , Data Inicial do Exercicio e Data Final do Exercicio
'A Data Inicial do ultimo periodo periodo %s não está dentro da faixa abrangida pelo exercício. Data Inicial = %s e Data Final = %s.
Public Const ERRO_EXERCICIOFILIAL_NAO_CADASTRADO = 606 'Parametros iExercicio, iFilialEmpresa.
'O Exercício %i da Filial %i não está cadastrado.
Public Const ERRO_CODIGO_HISTPADRAO_ZERADO = 607 'Sem parametro
'O codigo do histórico padrão tem que ser um número maior que zero.
Public Const ERRO_CODIGO_CATEGORIA_ZERADO = 608 'Sem parametro
'O codigo da categoria tem que ser um número maior que zero.
Public Const ERRO_LEITURA_SALDOINICIALCONTA2 = 609 'Sem Parametros.
'Ocorreu um erro na leitura da tabela SaldoInicialConta.
Public Const ERRO_LEITURA_SALDOINICIALCONTACCL5 = 610 'Sem Parametros
'Ocorreu um erro na leitura da tabela SaldoInicialContaCcl.
Public Const ERRO_INSERCAO_RATEIOOFFCONTAS = 611
'Ocorreu um erro ao tentar inserir dados na tabela RateioOffContas
Public Const ERRO_EXCLUSAO_RATEIOOFFCONTAS = 612 'Sem Parametros
'Ocorreu um erro na Exclusão de registros da tabela RateioOffContas.
Public Const ERRO_LEITURA_RATEIOOFFCONTAS = 613 'Parametro lCodigo
'Ocorreu um erro de leitura na tabela de RateioOffContas. Codigo do Rateio = %l.
Public Const ERRO_CONTA_SEG_NUM_CARACTER_INVALIDO = 614 'Sem parametro
'Os segmentos numéricos da Conta só podem conter números. Ex: 1.-1.1 está errado. 1.1.1 está correto.
Public Const ERRO_CCL_SEG_NUM_CARACTER_INVALIDO = 615 'Sem parametro
'Os segmentos numéricos do Centro de Custo só podem conter números. Ex: 1.-1.1 está errado. 1.1.1 está correto.
Public Const ERRO_ORIGEM_DIFERENTE_CTB = 616
'Não pode gravar lançamentos de outros módulos. Só é permitido gerar lançamentos para a Contabilidade.
Public Const ERRO_ORIGEM_DIFERENTE = 617
'Não pode gravar, alterar e excluir lotes de outros módulos.
Public Const ERRO_AUSENCIA_LANCAMENTOS_PADRAOCONTAB = 618 'Sem parametros
'Não há lançamentos em nenhum dos grids.
Public Const ERRO_INDICE_MAIOR_LINHAS_GRID = 619  'Parametros sMnemonico, iIndice, iLinhasGrid
'O Indice do mnemonico %s ultrapassa o número de linhas do grid. Indice = %i, Linhas do Grid = %i.
Public Const ERRO_DOC_NAO_BALANCEADO = 620 'Parametro: sModelo
'O modelo de contabilização, %s, usado para gerar os lançamentos contábeis não estão gerando um documento balanceado, ou seja, o total dos créditos não corresponde aos débitos.
Public Const ERRO_LINCOL_ZERO_CONTA = 621 'Sem Parametros
'Para que uma célula seja do tipo conta ela não pode estar posicionada na linha zero ou coluna 0.
Public Const ERRO_LINCOL_ZERO_FORMULA = 622 'Sem Parametros
'Para que uma célula seja do tipo fórmula ela não pode estar posicionada na linha zero ou coluna 0.
Public Const ERRO_CEL_UTILIZA_CEL_EM_FORMULA = 623 'Parametros: iLinhaUtiliza, iColunaUtiliza, iLinhaUsada, iColunaUsada
'A célula posicionada na linha/coluna (%i, %i) utiliza a célula posicionada na linha/coluna (%i, %i).
Public Const ERRO_LEITURA_RELDMPL_MODELO = 624 'Parametro: sModelo
'Ocorreu um erro na leitura do modelo %s na tabela RelDMPL.
Public Const ERRO_LEITURA_RELDMPLFORMULA = 625 'Parâmetro sModelo
'Ocorreu um erro na leitura do modelo %s na tabela RelDMPLFormula.
Public Const ERRO_LEITURA_RELDMPLCONTA = 626 'Parâmetro sModelo
'Ocorreu um erro na leitura do modelo %s na tabela RelDMPLConta.
Public Const ERRO_EXCLUSAO_RELDMPL = 627 'Parâmetro Modelo
'Ocorreu um erro na exclusão do modelo %s da tabela RelDMPL.
Public Const ERRO_EXCLUSAO_RELDMPLFORMULA = 628 'Parâmetro Modelo
'Ocorreu um erro na exclusão do modelo %s da tabela RelDMPLFormula.
Public Const ERRO_EXCLUSAO_RELDMPLCONTA = 629 'Parâmetro Modelo
'Ocorreu um erro na exclusão do modelo %s da tabela RelDMPLConta.
Public Const ERRO_INCLUSAO_RELDMPL = 630 'Parâmetro Modelo
'Ocorreu um erro na inclusão do modelo %s na tabela RelDMPL.
Public Const ERRO_INCLUSAO_RELDMPLCONTA = 631 'Parâmetro Modelo
'Ocorreu um erro na inclusão do modelo %s na tabela RelDMPLConta.
Public Const ERRO_INCLUSAO_RELDMPLFORMULA = 632 'Parâmetro Modelo
'Ocorreu um erro na inclusão do modelo %s na tabela RelDMPLFormula.
Public Const ERRO_RELDMPL_MODELO_NAO_CADASTRADO = 633 'Parâmetros: sModelo, iLinha, iColuna
'O Modelo %s referente à linha/coluna (%i,%i) do Demonstrativo de Mutação de Patrimonio não está cadastrado no Banco de Dados.
Public Const ERRO_LOCK_RELDMPL = 634 'Parâmetros: sModelo, iLinha, iColuna
'Não foi possível fazer o "Lock" na tabela RelDMPL do registro referente ao Modelo %s , Linha = %i e Coluna = %i.
Public Const ERRO_ATUALIZACAO_RELDMPL = 635  'Parâmetros: sModelo, iLinha, iColuna
'Ocorreu um erro na atualização da tabela RelDMPL. Modelo = %s, Linha = %i e Coluna = %i.
Public Const ERRO_CCL_SINTETICA_USADA_EM_MOVESTOQUE = 636 'Parametro: sCcl
'O Centro de Custo/Lucro %s está sendo utilizado em pelo menos um movimento de estoque.
Public Const ERRO_LIMITE_CCL_VLIGHT = 637 'Parametros : iNumeroMaxCcl
'Número máximo de Centro de Custo/Lucro analítico desta versão é %i.
Public Const ERRO_AUSENCIA_CONTAS_CATEGORIA_APURACAO = 638
'Não foram encontradas contas de grupos de categoria para apuração.
Public Const ERRO_RATEIOOFF_PROCESSADO = 639 'Parametro: lCodigo
'O rateio offline de código = %l foi processado.
Public Const ERRO_RATEIOOFF_CCL_ZERADO = 640 'Parametro: lCodigo
'O rateio offline de código = %l não gerou lançamento pois o total do centro de custo para as contas em questão está zerado.
Public Const ERRO_NAO_HA_LOTE_PENDENTE1 = 641 'Parametros sOrigem, iExercicio, iPeriodo
'Não há lote pendente disponível para o módulo %s no Periodo %i do Exercicio %i.
Public Const ERRO_LEITURA_USUARIOLOTE = 642 'Parametro sCodUsuario, sOrigem
'Erro na leitura dos dados do Usuario %s, Origem = %s na tabela UsuarioLote.
Public Const ERRO_ATUALIZACAO_USUARIOLOTE = 643 'Parametro sCodUsuario, sOrigem
'Ocorreu um erro na atualização da tabela UsuarioLote. Usuário = %s, Origem = %s.
Public Const ERRO_INSERCAO_USUARIOLOTE = 644 'Parametro sCodUsuario, sOrigem
'Ocorreu um erro na inserção de um registro na tabela UsuarioLote. Usuário = %s, Origem = %s.
Public Const ERRO_LEITURA_LOTEPENDENTE5 = 645 'Parametros Filial, origem, exercicio, periodo
'Ocorreu um erro na leitura da tabela de lotes pendentes. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i.
Public Const ERRO_RECALCULO_AUTOMATICO_SEM_MODELO = 646 'Sem Parametros
'O recálculo automático só pode ser marcado se o modelo estiver preenchido.
Public Const ERRO_LEITURA_SLDMESEST11 = 647 'Parametros  iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque (SldMesEst1). Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_LOCK_SLDMESEST1 = 648 'Parametros iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos mensais de estoque (SldMesEst1). Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_LEITURA_SLDMESEST21 = 649 'Parametros  iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque (SldMesEst2). Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_LOCK_SLDMESEST2 = 650 'Parametros iAno, iFilialEmpresa, sProduto
'Ocorreu um erro na tentativa de fazer 'lock' na tabela de saldos mensais de estoque (SldMesEst2). Ano=%i, FilialEmpresa=%i, Codigo do Produto=%s.
Public Const ERRO_EXCLUSAO_SLDMESEST11 = 651 'Parametros sProduto, iFilialEmpresa
'Ocorreu um erro na exclusão de registro da tabela de saldos mensais de estoque (SldMesEst1). Codigo do Produto = %s, Filial = %i.
Public Const ERRO_MNEMONICO_ESCANINHOCUSTOCONSIG_INEX = 652 'Parametros sModulo, iTransacao
'O mnemonico Escaninho_Custo_Consig não está cadastrado para a transação em questão. Modulo = %s, Transação = %i.

'??? jones 28/10

Public Const ERRO_SUBTIPOCONTABIL_NAO_ENCONTRADO = 653 'Parâmtero: objTipoDocInfo.iCodigo, objTipoDocInfo.sDescricao
'Não foi encontrada transação na tabela TransacaoCTB correspondente ao subtipo selecionado.
Public Const ERRO_MNEMONICO_ESCANINHOCUSTOBENEF_INEX = 654 'Parametros sModulo, iTransacao
'O mnemonico Escaninho_Custo_Benef não está cadastrado para a transação em questão. Modulo = %s, Transação = %i.
Public Const ERRO_LEITURA_MVDIACLI1 = 655 'parametro Filial, Cliente, FilialCliente, data
'Ocorreu um erro na leitura da tabela de Saldos Diários de Cliente. Filial=%i, Cliente=%l, Filial do Cliente = %i e Data=%s.
Public Const ERRO_INSERCAO_MVDIACLI = 656 'parametro Filial, Cliente, FilialCliente, data
'Ocorreu um erro na inserção de um registro na tabela de Saldos Diários de Cliente. Filial=%i, Cliente=%l, Filial do Cliente = %i e Data=%s.
Public Const ERRO_ATUALIZACAO_MVDIACLI = 657 'parametro Filial, Cliente, FilialCliente, data
'Ocorreu um erro na atualização de um registro na tabela de Saldos Diários de Cliente. Filial=%i, Cliente=%l, Filial do Cliente = %i e Data=%s.
Public Const ERRO_LEITURA_MVDIAFORN1 = 658 'parametro Filial, Fornecedor, FilialFornecedor, data
'Ocorreu um erro na leitura da tabela de Saldos Diários de Fornecedor. Filial=%i, Fornecedor=%l, Filial do Fornecedor = %i e Data=%s.
Public Const ERRO_INSERCAO_MVDIAFORN = 659 'parametro Filial, Fornecedor, FilialFornecedor, data
'Ocorreu um erro na inserção de um registro na tabela de Saldos Diários de Fornecedor. Filial=%i, Fornecedor=%l, Filial do Fornecedor = %i e Data=%s.
Public Const ERRO_ATUALIZACAO_MVDIAFORN = 660 'parametro Filial, Fornecedor, FilialFornecedor, data
'Ocorreu um erro na atualização de um registro na tabela de Saldos Diários de Cliente. Filial=%i, Fornecedor=%l, Filial do Fornecedor = %i e Data=%s.
Public Const ERRO_LEITURA_MVPERCLI = 661 'parametro Filial, Exercicio, Cliente, FilialCliente
'Ocorreu um erro na leitura da tabela de Saldos Mensais de Cliente. Filial=%i, Exercicio = %i, Cliente=%l, Filial do Cliente = %i.
Public Const ERRO_INSERCAO_MVPERCLI = 662 'parametro Filial, Exercicio, Cliente, FilialCliente
'Ocorreu um erro na inserção de um registro na tabela de Saldos Mensais de Cliente. Filial=%i, Exercicio = %i, Cliente=%l, Filial do Cliente = %i.
Public Const ERRO_ATUALIZACAO_MVPERCLI = 663 'parametro Filial, Exercicio, Cliente, FilialCliente
'Ocorreu um erro na atualização de um registro na tabela de Saldos Mensais de Cliente. Filial=%i, Exercicio = %i, Cliente=%l, Filial do Cliente = %i.
Public Const ERRO_LEITURA_MVPERFORN2 = 664 'parametro Filial, Exercicio, Fornecedor, FilialFornecedor
'Ocorreu um erro na leitura da tabela de Saldos Mensais de Fornecedor. Filial=%i, Exercicio = %i, Fornecedor=%l, Filial do Fornecedor = %i.
Public Const ERRO_INSERCAO_MVPERFORN = 665 'parametro Filial, Exercicio, Fornecedor, FilialFornecedor
'Ocorreu um erro na inserção de um registro na tabela de Saldos Mensais de Fornecedor. Filial=%i, Exercicio = %i, Fornecedor=%l, Filial do Fornecedor = %i.
Public Const ERRO_ATUALIZACAO_MVPERFORN = 666 'parametro Filial, Exercicio, Fornecedor, FilialFornecedor
'Ocorreu um erro na atualização de um registro na tabela de Saldos Mensais de Fornecedor. Filial=%i, Exercicio = %i, Fornecedor=%l, Filial do Fornecedor = %i.
Public Const ERRO_NAO_HA_EXERCICIO_CONTABIL_ABERTO = 667 'Sem Parametros
'Atenção! Não há exercício contábil aberto.
Public Const ERRO_PERIODO_NAO_ENCONTRADO_DATA = 668 'Parametro: Data
'Não foi encontrado o periodo contábil que englobe esta data. Data = %s.
Public Const ERRO_LEITURA_LANCAMENTOS_PENDENTES1 = 669 'Parametros: iFilial, sOrigem, iExercicio, iPeriodo, lDoc, iLote
'Ocorreu um erro na leitura na tabela de Lançamentos Pendentes. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Doc = %l, Lote = %i.
Public Const ERRO_ALTERACAO_LANPENDENTE = 670 'Parametros: iFilial, sOrigem, iExercicio, iPeriodo, lDoc, iSeq
'Ocorreu um erro na alteração da tabela de Lançamentos Pendentes. Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Doc = %l, Sequencial = %i.
Public Const ERRO_NUM_LANC_TELA_MAIOR_BD = 671 'Sem Parametros
'O número de lançamentos da tela é maior do que o encontrado no banco de dados.
Public Const ERRO_NUM_LANC_TELA_MENOR_BD = 672 'Sem Parametros
'O número de lançamentos da tela é menor do que o encontrado no banco de dados.



''VEIO DE ERROS COM
Public Const ERRO_LEITURA_COMPRASCONFIG = 12005 'Parametro sCodigo
'Erro na leitura de %s na tabela de ComprasConfig.
Public Const ERRO_ATUALIZACAO_COMPRASCONFIG = 12006 'Parametro sCodigo
'Erro na atualizacao de %s na tabela de ComprasConfig.


''VEIO DE ERROS MAT
Public Const ERRO_LEITURA_TABELA_UNIDADESDEMEDIDA = 7328 'Parametros: iClasseUM, sSiglaUM
'Erro na Leitura da Unidade de Medida. Classe=%i e Sigla=%s.
Public Const ERRO_UNIDADE_MEDIDA_NAO_CADASTRADA = 7329 'Parametros: iClasseUM, sSiglaUM
'Unidade de Medida com Classe=%i e Sigla=%s não está cadastrada no Banco de Dados.
Public Const ERRO_MODIFICACAO_UNIDADESDEMEDIDA = 7348 'Sem parametro
'Erro na modificação da tabela UnidadesDeMedida.
Public Const ERRO_ESTOQUEMES_INEXISTENTE = 7439 'Parametros iFilialEmpresa, iAno, iMes
'O Mês em questão não está aberto. Filial Empresa = %i, Ano = %i, Mês = %i.
Public Const ERRO_LEITURA_ESTOQUEMES = 7441 'Parametros iFilialEmpresa, iAno, iMes
'Ocorreu um erro na leitura da tabela EstoqueMes. FilialEmpresa = %i, Ano = %i, Mes = %i.
Public Const ERRO_QUANTIDADE_NAO_PREENCHIDA = 7447 'Parametro iLinhaGrid
'A Quantidade do ítem %i do Grid não foi preenchida.
Public Const ERRO_PRODUTO_NAO_PREENCHIDO = 7523 'Sem parâmetros
'O Produto deve estar preenchido.
Public Const ERRO_LOCK_TABELASDEPRECOITENS = 7573
'Erro na tentativa de fazer 'lock' na tabela TabelasDePrecoItens.
Public Const ERRO_EXCLUSAO_TABELASDEPRECOITENS = 7574
'Erro na exclusão de registro na tabela TabelasDePrecoItens.
Public Const ERRO_ATUALIZACAO_TABELASDEPRECOITENS = 7585
'Erro na atualização da tabela TabelasDePrecoItens
Public Const ERRO_INSERCAO_TABELASDEPRECOITENS = 7586 'Parâmetros: iCodTabela, iCodProduto
'Erro na tentativa de inserir registro na tabela TabelasDePrecoItens. Com CodTabela = %i e CodProduto = %s.
Public Const ERRO_NOTA_FISCAL_NAO_CADASTRADA = 7600 'lNumIntDoc
'A Nota Fiscal com o Número Interno %l não está cadastrada no Banco de Dados.
Public Const ERRO_LOCK_SERIE = 7601
'Erro na tentativar de fazer "lock" na tabela de Séries.
Public Const ERRO_LEITURA_NFISCAL1 = 7603 'Parâmetros: iTipoNFiscal, lFornecedor, iFilialForn, sSerie, lNumNotaFiscal
'Erro na leitura da tabela NFiscal na Nota Fiscal com Tipo = %i, Fornecedor = %l, Filial = %i, Serie = %s e Número = %l.
Public Const ERRO_LEITURA_NFISCALBAIXADA2 = 7636
'Erro na leitura da tabela NFiscalBaixadas
Public Const ERRO_LEITURA_NFISCALBAIXADA = 7641
'Erro na leitura da tabela NFiscalBaixada da Nota Fiscal em questão.
Public Const ERRO_SLDMESEST_NAO_CADASTRADO = 7647 'Parâmetros: iFilialEmpresa, iAno, sProduto
'Registro da tabela SldMesEst não cadastrado. Dados do registro: FilialEmpresa=%i, Ano=%i, Produto=%s.
Public Const ERRO_DATA_NAO_PREENCHIDA = 7726 'Sem parâmetros
'O preenchimento da Data é obrigatório.
Public Const ERRO_LOTE_NAO_PREENCHIDO = 7769 'Sem parametros
'Preenchimento do lote é obrigatório.
Public Const ERRO_LEITURA_ITENSNFISCALBAIXADAS = 7796
'Erro na leitura da tabela dos Itens de N.Fiscal Baixadas nos itens vinculados com a Nota Fiscal em questão.
Public Const ERRO_LEITURA_ITENSNFISCAL1 = 7797
'Erro na leitura do Item vinculado com a Nota Fiscal em questão na tabela dos Itens de N.Fiscal.
Public Const ERRO_LEITURA_ITENSNFISCALBAIXADAS1 = 7798 'NumIntDoc
'Erro na leitura da tabela dos Itens de N.Fiscal Baixadas nos itens vinculados com a Nota Fiscal em questão.
Public Const ERRO_LEITURA_CLIENTES1 = 7914 'Parametro lCodCliente
'Ocorreu um erro na leitura dos dados da tabela de Clientes. Código do Cliente = %l.
Public Const ERRO_LEITURA_SLDMESEST3 = 8568  'Parametros:iFilialEmpresa, sProduto
'Ocorreu um erro na leitura da tabela de saldos mensais de estoque (SldMesEst). FilialEmpresa=%i, Produto = %s.
Public Const ERRO_SLDMESEST_NAO_CADASTRADO1 = 8569 'Parâmetros: iFilialEmpresa, sProduto
'Não há Registro cadastrado na tabela SldMesEst para a FilialEmpresa=%i,  Produto=%s.
Public Const ERRO_LEITURA_TITULOSPAGBAIXADOS = 8590 'Parâmetro: lNumIntDoc
'Erro na tentativa de ler registro na tabela TitulosPagBaixados com Número Interno %l.



''VEIO DE ERROS FAT
Public Const ERRO_ATUALIZACAO_TABELASDEPRECOITENS1 = 8124 'Parametro: iCodTabela
'Erro na atualização da tabela TabelasDePrecoItens com código da Tabela %i.


''VEIO DE ERROS CRFAT
Public Const ERRO_LEITURA_MVDIACLI = 6077
'Erro de leitura na tabela MvDiaCli.
Public Const ERRO_LEITURA_NFISCAL = 6123
'Erro na leitura da tabela NFiscal da Nota Fiscal em questão.
Public Const ERRO_LEITURA_ITENSNFISCAL = 6125 'Parâmetro: lNumIntNF
'Erro na leitura do Item vinculado com a Nota Fiscal em questão na tabela dos Itens de N.Fiscal.
Public Const ERRO_LEITURA_FORNECEDORES = 6208 'Sem parametros
'Erro na leitura da tabela de Fornecedores.
Public Const ERRO_TABELAPRECO_INEXISTENTE = 6403 'Parametro: iCodigo
'A Tabela de Preço com Código %i não está cadastrada no Banco de Dados.
Public Const ERRO_LEITURA_TABELASDEPRECO1 = 6404 'Parametro: iCodigo
'Erro na leitura da tabela TabelasDePreco com o Código %i.
Public Const ERRO_LOCK_TABELASDEPRECO = 6405 'Parametro: iCodigo
'Não conseguiu fazer o lock na tabela de TabelasDePreco com Código da Tabela %i.



''VEIO DE ERROS CPR
Public Const ERRO_FORNECEDOR_NAO_CADASTRADO = 2034 'Parametro: lCodFornecedor
'O Fornecedor com código %l não está cadastrado no Banco de Dados.
Public Const ERRO_LOCK_FORNECEDORES = 2035 'Parametro Codigo Fornecedor
'Erro na tentativa de fazer "lock" na tabela Fornecedores para Código Fornecedor = %l .
Public Const ERRO_FILIALFORNECEDOR_NAO_CADASTRADA = 2300 'Parametro: lFornecedor, iFilial
'A Filial %i do Fornecedor %l nao está cadastrada no Banco de Dados.
Public Const ERRO_FORNECEDOR_NAO_CADASTRADO1 = 2318 'Parametro: sFornecedorNomeRed
'O Fornecedor %s nao está cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_PRODUTOSFILIAL1 = 2655 'Parâmetros: lCodigo
'Erro na leitura da tabela de ProdutosFilial com Fornecedor %l.
Public Const ERRO_LEITURA_MVPERFORN = 2660 'Parâmetros: lCodigo
'Erro na leitura da tabela de MvPerForn com Fornecedor %l.
Public Const ERRO_LEITURA_MVDIAFORN = 2657 'Parâmetros: lCodigo
'Erro na leitura da tabela de MvDiaForn com Fornecedor %l.
Public Const ERRO_LOCK_MVDIAFORN = 2658 'Parâmetros: lCodigo
'Erro na tentativa de fazer "lock" na tabela MvDiaForn com Fornecedor %l.
Public Const ERRO_EXCLUSAO_MVDIAFORN = 2659 'Parâmetros: lCodigo
'Erro na tentativa de excluir um registro da tabela MvDiaForn com Fornecedor %l.
Public Const ERRO_LOCK_MVPERFORN = 2661 'Parâmetros: lCodigo
'Erro na tentativa de fazer "lock" na tabela MvPerForn com Fornecedor %l.
Public Const ERRO_EXCLUSAO_MVPERFORN = 2662 'Parâmetros: lCodigo
'Erro na tentativa de excluir um registro da tabela MvPerForn com Fornecedor %l.
Public Const ERRO_LOCK_MVDIAFORN1 = 2615 'Parâmetros: lCodFornecedor, iCodFilial
'Erro na tentativa de fazer "lock" na tabela MvDiaForn com Fornecedor %l e Filial %i.
Public Const ERRO_EXCLUSAO_MVDIAFORN1 = 2616 'Parâmetros: lCodFornecedor, iCodFilial
'Erro na tentativa de excluir um registro da tabela MvDiaForn com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_MVPERFORN1 = 2617 'Parâmetros: lCodFornecedor, iCodFilial
'Erro na leitura da tabela de MvPerForn com Fornecedor %l e Filial %i.
Public Const ERRO_LOCK_MVPERFORN1 = 2618 'Parâmetros: lCodFornecedor, iCodFilial
'Erro na tentativa de fazer "lock" na tabela MvPerForn com Fornecedor %l e Filial %i.
Public Const ERRO_EXCLUSAO_MVPERFORN1 = 2619 'Parâmetros: lCodFornecedor, iCodFilial
'Erro na tentativa de excluir um registro da tabela MvPerForn com Fornecedor %le Filial %i.


''VEIO DE ERROS EST
Public Const ERRO_RECEBIMENTO_VINCULADO_NF = 7125 'Parametros lNumeroRecebimento, lNumeroNFiscal
'O Recebimento %l já está vinculado com a Nota Fiscal %l.
Public Const ERRO_CLIENTE_RECEB_DIFERENTE_NF = 7126 'ClienteRecebimento
'O Cliente da Nota Fiscal não pode ser diferente do Cliente do Recebimento. Cliente do Recebimento = %l
Public Const ERRO_FORN_RECEB_DIFERENTE_NF = 7127 'FornecedorRecebimento
'O Fornecedor da Nota Fiscal não pode ser diferente do Fornecedor do Recebimento. Fornecedor do Recebimento = %l
Public Const ERRO_FILCLIENTE_RECEB_DIFERENTE_NF = 7128 'FilialClienteRecebimento
'A Filial Cliente da Nota Fiscal não pode ser diferente da Filial do Cliente do Recebimento. Filial Cliente do Recebimento = %i
Public Const ERRO_FILFORN_RECEB_DIFERENTE_NF = 7129 'FilialForncedorRecebimento
'A Filial Fornecedor da Nota Fiscal não pode ser diferente da Filial Fornecedor do Recebimento. Filial Fornecedor do Recebimento = %i
Public Const ERRO_SERIE_RECEB_DIFERENTE_NF = 7130 'Serie Recebimento
'A Serie da Nota Fiscal não pode ser diferente da Serie do Recebimento. Serie = %s
Public Const ERRO_DATAENTRADA_RECEB_DIFERENTE_NF = 7131 'Data Entrada
'A Data de Entrada da Nota Fiscal não pode ser diferente da Data de Entrada do Recebimento. Data de Entrada do Recebimento = %dt
Public Const ERRO_NUMNF_RECEB_DIFERENTE_NF = 7132 'Serie
'O Número da Nota Fiscal não pode ser diferente do Número da Nota Fiscal do Recebimento. Número da Nota Fiscal do Recebimento = %l


''VEIO DE ERROS TRB
Public Const ERRO_LEITURA_TIPO_TRIBUTACAO = 7003 'parametro tipo da tributacao
'Erro na leitura do tipo de tributação %d.



'Codigos de Aviso - Reservado de 5100 até 5199
Global Const AVISO_EXCLUSAO_CONTA_ANALITICA = 5100 'Sem parametros
'A conta que está sendo excluida é analitica. Confirma a exclusão?
Global Const AVISO_EXCLUSAO_CONTA_SINTETICA_COM_FILHOS = 5101 'Sem parametros
'A conta que está sendo excluida é sintética e possui contas abaixo dela.
'Ao excluir esta conta, suas "filhas" serão também excluidas.
'Confirma a exclusão?
Global Const AVISO_EXCLUSAO_CONTA_SINTETICA = 5102 'Sem parametros
'A conta que está sendo excluida é sintética e não possui contas abaixo dela.
'Confirma a exclusão?
Global Const SUBSTITUICAO_DOCUMENTO_PENDENTE = 5103 'Parametros: lDocumento, sOrigem, iExercicio, iPeriodoLan
'O documento pendente %l (Origem: %s, Exercício: %i, Período: %i) já existe.
'Deseja substituí-lo?
Global Const AVISO_EXCLUSAO_DOCUMENTO = 5104 'Sem Parametros
'Confirma a exclusão do documento?
Global Const AVISO_EXCLUSAO_CCL_COM_ASSOCIACOES = 5105 'Parametros: Conta
'O Centro de Custo/Lucro que está sendo excluida possui associação com %s. Estas informações serão excluidas junto com o Centro de Custo/Lucro.
'Confirma a exclusão?
Global Const AVISO_EXCLUSAO_CCL = 5106 'Sem parametros
'Confirma a exclusão do Centro de Custo/Lucro?
Global Const AVISO_ATUALIZACAO_TOTAIS = 5107 'Parametros total de crédito, total de débito, número de lançamentos
'Os totais calculados através da leitura dos lançamentos, difere do
'dos totais exibidos. Totais Calculados: Crédito = %d, Débito = $d e
'Número de lançamentos = %i. Deseja alterar os valores exibidos para
'que fiquem compatíveis com os lançamentos?
Global Const AVISO_IGUALDADE_TOTAIS = 5108 'Sem parametros
'Os totais calculados através da leitura dos lançamentos são iguais
'aos exibidos na sua tela.
Global Const AVISO_EXCLUSAO_CCL_COM_DOCAUTO = 5109 'Sem parametros
'O Centro de Custo/Lucro que está sendo excluida possui documentos automáticos associados. Estas informações serão excluidas junto com o Centro de Custo/Lucro.
'Confirma a exclusão?
Global Const EXCLUSAO_HISTPADRAO = 5110 'Sem parametros
'Confirma a exclusão do Histórico Padrão?
Global Const AVISO_LOTE_INEXISTENTE = 5111 'Parametros Filial, Lote, Origem, Periodo, Exercicio
'Nao existe lote com chave: Filial = %i, Lote = %i, Origem = %s, Periodo = %i, Exercício = %i. Deseja criar?
Global Const AVISO_CONTA_INEXISTENTE = 5112 'Parametro Conta
'A Conta %s não está cadastrada. Deseja cadastrá-la?
Global Const AVISO_EXCLUSAO_CONTA_ANALITICA_COM_ASSOCIACOES = 5113 'Sem parametros
'A conta que está sendo excluida é analítica e possui associação com centro de custo. Estas informações serão excluidas junto com a conta.
'Confirma a exclusão?
Global Const AVISO_CCL_INEXISTENTE = 5114 'Parametro Ccl
'O Centro de Custo/Lucro %s não está cadastrado. Deseja cadastrá-lo?
Global Const AVISO_CONTACCL_INEXISTENTE = 5115 'Parametros Conta e Ccl
'A associação da Conta %s com o Centro de Custo/Lucro %s não está cadastrada. Deseja cadastrá-la?
Global Const AVISO_HISTPADRAO_INEXISTENTE = 5116 'Parametro HistPadrao
'O Historico %i não está cadastrado.  Deseja criar agora?
Global Const AVISO_EXCLUSAO_LOTE = 5117 'Sem parametros
'Confirma a exclusão do Lote?
Global Const AVISO_EXCLUSAO_RATEIO = 5118 'Sem parametro
'Confirma a exclusão do Rateio?
Global Const AVISO_ALTERACAO_CONTACCL = 5119 'Sem parametro
'As associações selecionadas serão feitas. Caso exista alguma associação de alguma das contas com Centro de Custo não selecionado, esta associação será excluída. Confirma a Alteração?
Public Const AVISO_EXCLUSAO_CCL_COM_ASSOC_DOCAUTO = 5120 'Sem Parametros
'O Centro de Custo/Lucro que está sendo excluida possui documentos automáticos associados e associação com conta. Estas informações serão excluidas junto com o Centro de Custo/Lucro.
'Confirma a exclusão?
Public Const AVISO_FECHAMENTO_EXERCICIO_EXECUTADO = 5121 'Parametro Nome do Exercicio
'O Disparo do processo de fechamento do exercicio %s foi feito. O seu processamento
'poderá ser acompanhado por uma tela que surgirá a seguir
Public Const AVISO_NAO_HA_LOTE_DESATUALIZADO = 5122
'Nao existe nenhum lote desatualizado.
Public Const AVISO_REABERTURA_EXERCICIO_EXECUTADA = 5123 'Parametro Nome do Exercicio
'O Disparo do processo de rEABERTURA do exercicio %s foi feito. O seu processamento
'poderá ser acompanhado por uma tela que surgirá a seguir
Public Const AVISO_REPROCESSAMENTO_EXERCICIO_EXECUTADA = 5124
'O Disparo do processo de Reprocessamento do exercicio %s foi feito. O seu processamento
'poderá ser acompanhado por uma tela que surgirá a seguir
Public Const AVISO_EXCLUSAO_ORCAMENTO = 5125
'Confirma a exclusão do Orçamento?
Public Const AVISO_EXCLUSAO_CONTACATEGORIA = 5126
'Confirma a exclusão da Categoria?
Public Const AVISO_CONTARESULTADO_NAO_ESPECIFICADA = 5127 'Sem parametro
'A conta de resultado não foi espeficada. Deseja espeficar agora?
Public Const AVISO_ULTIMO_PERIODO_MAIOR = 5128
'O Último Período possuirá uma duração maior que os demais.
Public Const AVISO_EXCLUSAO_CCL_SINTETICA_COM_FILHOS = 5129 'Sem parametros
'O Centro de Custo/Lucro que está sendo excluido é sintético e possui centros de custo/lucro abaixo dele.
'Ao excluir este centro de custo/lucro, seus "filhos" serão também excluidos.
'Confirma a exclusão?
Public Const AVISO_EXCLUSAO_CCL_SINTETICA = 5130 'Sem parametros
'O Centro de Custo/Lucro que está sendo excluido é sintético e não possui centros de custo/lucro abaixo dele.
'Confirma a exclusão?
Public Const AVISO_HA_LANCAMENTO_DESATUALIZADO = 5131 'Sem parametros
'Existe(m) lançamento(s) desatualizado(s) para este exercício. Deseja prosseguir?
Public Const AVISO_EXCLUSAO_PADRAOCONTAB = 5132 'Sem Parametros
'Confirma a exclusão do modelo de contabilização?
Public Const AVISO_LOTE_ATUALIZANDO = 5133 'Parâmetro: iFilialEmpresa
'A Filial %i só possui Lotes que estão sendo atualizados. Não existem lotes pendentes de atualização.
Public Const AVISO_EXCLUSAO_MODELORELDRE = 5134
'Confirma a exclusão do modelo?
Public Const AVISO_ESTORNO_LANCAMENTO_CANCELADO = 5135 'parâmetros sOrigem , iExercicio , iPeriodo , lDoc
'O estorno foi cancelado . Origem= %s, Exercicio= %s, Periodo= %s, Doc = %s.
Public Const AVISO_ESTORNO_LOTE_CANCELADO = 5136 'parâmetros : sOrigem , iExercicio, iPeriodo, iLote
'O estorno foi cancelado . Origem= %s, Exercicio= %i, Periodo= %i, Lote= %i
Public Const AVISO_ELEMENTO_TEM_FILHOS = 5137 'Sem parametros
'Este elemento possui 'filhos'. Deseja excluí-lo assim mesmo ?
Public Const AVISO_CANCELAR_APURACAO_RATEIOS = 5138
'Confirma o cancelamento da apuração dos rateios ?
Public Const AVISO_EXCLUSAO_EXERCICIO = 5139 'Sem parâmetros
'Confirma a exclusão do Exercício?
Public Const AVISO_ATUALIZACAO_LOTE_TOTAIS_DIFERENTES = 5140 'Parametros iFilialEmpresa, sOrigem, iExercicio, iExercicio, iLote, iNumDocInf, iNumDocAtual, dTotInf, dTotAtual
'Os totais calculados através da leitura dos lançamentos, difere do
'dos totais informados para o lote (Filial = %i, Origem = %s, Exercicio = %i, Periodo = %i, Lote = %i). Número de Documentos (Informado = %i, Cadastrados = %i) e Valores dos Documentos (Informado = %d e Cadastrados = %d).
'Confirma a contabilização deste lote?
Public Const AVISO_LOTE_ATUALIZANDO_GRAVA = 5141 'Parâmetro: iLote
'O Lote %i está sendo atualizado, deseja continuar a alteração ?
Public Const AVISO_DATA_TROCADA_ULTIMO_PERIODO_ABERTO = 5142 'Parametros: DataAtual, DataFim, DataAtual
'Atenção. A data do sistema %s foi alterada para %s pois não existia periodo contábil aberto. Para que o sistema abra com a data %s , favor cadastrar um exercicio que englobe-a.


