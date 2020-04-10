Attribute VB_Name = "ErrosCPR"
Option Explicit

'C�digos de Erro - Reservado de 2000 at� 2999; 16000 at� 16499
Public Const ERRO_LEITURA_BANCOS = 2000 'Parametro iCodigo
' Erro na leitura do Banco %i.
Public Const ERRO_BANCO_NAO_CADASTRADO = 2001 'Parametro iCodigo
' Banco %i n�o est� cadastrado.
Public Const ERRO_INSERCAO_BANCOS = 2002 'Parametro iCodigo
' Erro na inser��o do Banco %i.
Public Const ERRO_ATUALIZACAO_BANCOS = 2003 'Parametro iCodigo
' Erro na atualiza��o do Banco %i.
Public Const ERRO_EXCLUSAO_BANCOS = 2004 'Parametro iCodigo
' Erro na exclus�o do Banco %i.
Public Const ERRO_LOCK_BANCOS = 2005 'Parametro iCodigo
' N�o conseguiu fazer o lock do Banco %i.
Public Const ERRO_NOME_REDUZIDO_BANCO_REPETIDO = 2006 'Sem Parametros
' J� existe Banco com este nome reduzido.
Public Const ERRO_NOME_NAO_PREENCHIDO = 2007 'Sem parametros
' Preenchimento do Nome � obrigat�rio.
Public Const ERRO_NOME_REDUZIDO_NAO_PREENCHIDO = 2008 'Sem parametros
' Preenchimento do Nome Reduzido � obrigat�rio.
Public Const ERRO_LEITURA_CONTASCORRENTESINTERNAS = 2009 'Sem parametros
' Erro na leitura da tabela ContasCorrentesInternas.
Public Const ERRO_BANCO_RELACIONADO_COM_CONTASCORRENTESINTERNAS = 2010 'Sem parametros
' N�o � poss�vel excluir banco relacionado com ContasCorrentesInternas.
Public Const ERRO_INSERCAO_REGIOESVENDAS = 2011 'Parametro iCodigo
' Erro na inser��o da Regi�o de Venda %i.
Public Const ERRO_ATUALIZACAO_REGIOESVENDAS = 2012 'Parametro iCodigo
' Erro na atualiza��o da Regi�o de Venda %i.
Public Const ERRO_EXCLUSAO_REGIOESVENDAS = 2013 'Parametro iCodigo
' Erro na exclus�o da Regi�o de Venda %i.
Public Const ERRO_REGIAO_VENDA_RELACIONADA_COM_VENDEDOR = 2014 'Sem parametros
' N�o � poss�vel excluir Regi�o de Venda relacionada com Vendedor.
Public Const ERRO_REGIAO_VENDA_RELACIONADA_COM_FILIAIS_CLIENTES = 2015 'Sem parametros
' N�o � poss�vel excluir Regi�o de Venda relacionada com Filiais Clientes.
Public Const ERRO_REGIAO_VENDA_RELACIONADA_COM_TIPOSDECLIENTE = 2016 'Sem parametros
' N�o � poss�vel excluir Regi�o de Venda relacionada com Tipos de Cliente.
Public Const ERRO_DESCRICAO_REPETIDA = 2017 'Sem parametros
' J� existe Regi�o de Venda com esta descri��o.
Public Const ERRO_LEITURA_TIPOSDECLIENTE = 2019 'Sem parametros
'Erro na leitura da tabela TipoDeCliente
Public Const ERRO_PAIS_NAO_PREENCHIDO = 2020 'Sem parametros
' Preenchimento do Pa�s � obrigat�rio.
Public Const ERRO_TAMANHO_CGC_CPF = 2022
'O tamanho do campo CGC tem que ser 11 caracteres(CPF) ou 14(CGC).
Public Const ERRO_CODCLIENTE_NAO_PREENCHIDO = 2023
'O c�digo do Cliente n�o foi preenchido.
Public Const ERRO_RAZ_SOC_NAO_PREENCHIDA = 2024
'O Nome n�o foi preenchido.
Public Const ERRO_CREDITO_NEGATIVO = 2025
'O valor do Limite de Credito tem que ser positivo.
Public Const ERRO_FORNECEDOR_SEM_FILIAL = 2027 'Parametro codigo do fornecedor
'O Fornecedor %l n�o est� vinculado a nenhuma filial.
Public Const ERRO_FORNECEDOR_INEXISTENTE = 2028
'O Fornecedor n�o est� cadastrado.
Public Const ERRO_FORNECEDOR_NOME_RED_DUPLICADO = 2029
'Erro na tentativa de cadastrar novo Fornecedor com o Nome Reduzido ja existente.
Public Const ERRO_INSERCAO_FORNECEDORES = 2030
'Erro na tentativa de inserir um novo Fornecedor no Banco de Dados.
Public Const ERRO_INSERCAO_FILIAISFORNECEDORES = 2031
'Erro na tentativa de inserir uma nova Filial de Fornecedor no Banco de Dados.
Public Const ERRO_MODIFICACAO_FORNECEDOR = 2032 'Parametro codigo do Fornecedor
'Erro na tentativa de modificar o Fornecedor %l.
Public Const ERRO_MODIFICACAO_FILIAISFORNECEDORES = 2033 'Parametro Codigo da Filial , Codigo do Fornecedor
'Erro na tentativa de modificar a Filial %i do Fornecedor %l .
Public Const ERRO_FORNECEDOR_REL_RECEB_ANTECIPADO = 2036
'Erro na tentatica de excluir o Fornecedor , pois ele se encontra relacionado com Recebimento Antecipado.
Public Const ERRO_EXCLUSAO_FILIAISFORNECEDORES = 2037 'Parametro Codigo do Fornecedor
'Erro na tentativa de excluir as Filiais do Fornecedor %l.
Public Const ERRO_EXCLUSAO_FORNECEDOR = 2038 'Parametro Codigo do Fornecedor.
'Erro na tentativa de excluir o Fornecedor %l.
Public Const ERRO_LOCK_TIPOSFORNECEDOR = 2039
'Erro na leitura da tabela TiposDeFornecedor.
Public Const ERRO_LEITURA_NOTAS_FISCAIS_EXT = 2041
'Erro na leitura da tabela Notas Fiscais Externas.
Public Const ERRO_LEITURA_CREDITO_PAG_FORN = 2043
'Erro na leitura da tabela Creditos a Pagar Fornecedor.
Public Const ERRO_FORNECEDOR_REL_NOTA_FISCAL_EXT = 2044
'O Forncedor n�o pode ser exclu�do pois se encontra relacionado com Nota Fiscal Externa.
Public Const ERRO_FORNECEDOR_REL_TITULOS_PAGAR = 2045
'O Forncedor n�o pode ser exclu�do pois se encontra relacionado com T�tulos a Pagar.
Public Const ERRO_FORNECEDOR_REL_CREDITOS_PAGAR = 2046
'O Forncedor n�o pode ser exclu�do pois se encontra relacionado com Creditos a Pagar.
Public Const ERRO_CODFORNECEDOR_NAO_PREENCHIDO = 2048
'O C�digo do Fornecedor deve ser preenchido.
Public Const ERRO_FILIALFORNECEDOR_GRAVA_MATRIZ = 2049
'A grava��o da Matriz deve ser feita pela tela de Fornecedores.
Public Const ERRO_FILIAL_FORNECEDOR_EXCLUSAO_MATRIZ = 2050
'A exclus�o da Matriz do Fornecedor deve ser feita pela tela de Fornecedores.
Public Const ERRO_FILIALFORNECEDOR_REL_NFE = 2051
'A Filial Fornecedor n�o pode ser exclu�da pois se encontra relacionada com Nota Fiscal Externa.
Public Const ERRO_FILIALFORNECEDOR_REL_TIT_PAGAR = 2052
'A Filial Fornecedor n�o pode ser exclu�da pois se encontra relacionada com T�tulos a Pagar.
Public Const ERRO_FILIALFORNECEDOR_REL_CREDITOS = 2053
'A Filial Fornecedor n�o pode ser exclu�da pois se encontra relacionada com Cr�ditos a Pagar.
Public Const ERRO_FILIALFORNECEDOR_NOME_DUPLICADO = 2054
'O Nome %s da Filial j� esta sendo utilizado por uma outra.
Public Const ERRO_LEITURA_FAVORECIDOS1 = 2055 'Sem Parametros
'Erro na leitura da tabela de Favorecidos.
Public Const ERRO_FAVORECIDO_NAO_CADASTRADO = 2056 'Parametro Favorecido
'O Favorecido %s n�o est� cadastrado.
Public Const ERRO_CODIGO_FAVORECIDO_NAO_INFORMADO = 2057 'Sem parametro
'O C�digo do Favorecido n�o foi informado.
Public Const ERRO_NOME_FAVORECIDO_NAO_INFORMADO = 2058 'Sem parametro
'Nome do Favorecido n�o foi informado.
Public Const ERRO_LEITURA_FAVORECIDO = 2059 'Parametro Codigo
'Erro na leitura do Favorecidos - Codigo = %s.
Public Const ERRO_ATUALIZACAO_FAVORECIDOS = 2060 'Parametro Favorecido
'Erro de atualiza��o do Favorecido %i.
Public Const ERRO_INSERCAO_FAVORECIDOS = 2061 'Parametro Favorecido
'Erro na inser��o do Favorecido %s na tabela Favorecidos.
Public Const ERRO_ATUALIZACAO_CCIMOVDIA = 2062
'Erro de atualiza��o da tabela de CCIMovDia.
Public Const ERRO_ATUALIZACAO_CCIMOV = 2063
'Erro de atualiza��o da tabela de CCIMov.
Public Const ERRO_ATUALIZACAO_MOVIMENTOSCONTACORRENTE = 2064
'Erro de Atualizacao da Tabela de MovimentosContaCorrente
Public Const ERRO_ATUALIZACAO_CONTASCORRENTESINTERNAS = 2065
'Erro na Atualizacao da Tabela ContasCorrentesInterna
Public Const ERRO_INSERCAO_CCIMOV = 2066
'Erro na Insercao de Registro na Tabela de Saldos Mensais de Conta Corrente.
Public Const ERRO_INSERCAO_CCIMOVDIA = 2067 ''ParaMetros: dtData , icodigo
'Erro na tentativa de insercao de Saldo Di�rio de Conta Corrente. Dia = %s e Conta = %s.
Public Const ERRO_INSERCAO_MOVIMENTOSCONTACORRENTE = 2068 'Parametro: iCodContaCorrente
'Erro na tentativa de inclus�o de Movimento envolvendo a Conta Corrente %i.
Public Const ERRO_LEITURA_CCIMOV = 2069 'Parametros iCodconta ,   iAno
'Erro na leitura da tabela de Saldos Mensais de Conta Corrente. Conta = %s e Ano = %s.
Public Const ERRO_LEITURA_CCIMOVDIA = 2070 'ParaMetros: dtData , icodigo
'Erro de Leitura na tabela de Saldos Di�rios de Conta Corrente. Data = %s e Conta = %s.
Public Const ERRO_LEITURA_TIPOMEIOPAGTO = 2071
'Erro de Leitura na Tabela TipoMeioPagto
Public Const ERRO_LEITURA_MOVIMENTOSCONTACORRENTE = 2072
'Erro na leitura da tabela de Movimentos de Conta Corrente.
Public Const ERRO_LOCK_FAVORECIDOS = 2073
'Erro na tentativa de fazer "lock" na tabela Favorecidos
Public Const ERRO_LOCK_MOVIMENTOSCONTACORRENTE = 2074
'Erro na tentativa de fazer "lock" na tabela MOvimentosContasCorrente
Public Const ERRO_LOCK_TIPOMEIOPAGTO = 2075 'Parametro: iTipoMeioPagto
'Erro na tentativa de fazer "lock" do Tipo de Pagamento %i na tabela TipoMeioPagto.
Public Const ERRO_LOCK_CONTASCORRENTESINTERNAS = 2076 'Parametro: iCodConta
'Erro na tentativa de fazer "lock" na Conta Corrente %i na tabela ContasCorrentesInternas.
Public Const ERRO_EXCLUSAO_MOVIMENTOSCONTACORRENTE = 2077
'Erro na exclusao de Registro da tabela MovimentosContasCorrentes
Public Const ERRO_MOVCONTACORRENTE_CONCILIADO = 2078 'Parametros: icodconta, lsequencial
'O movimento na conta %i e sequencial %l nao pode ser excluido por estar conciliado.
Public Const ERRO_CONTACORRENTE_INEXISTENTE = 2079 'Parametro: iCodConta
'A Conta Corrente %i nao est� cadastrada.
Public Const ERRO_TIPOMEIOPAGTO_INATIVO = 2080 'Parametro: iTipoMeioPagto
'O Tipo de Pagamento %i est� inativo.
Public Const ERRO_TIPO_NAO_SAQUE = 2081 'Parametro: lSequencial
'O movimento %l n�o � do tipo Saque.
Public Const ERRO_FAVORECIDO_INATIVO = 2082 'Parametro: Codigo do Favorecido
'O Favorecido %i n�o est� ativo.
Public Const ERRO_FAVORECIDO_INEXISTENTE = 2083 'Parametro Codigo do Favorecido
'O favorecido %i nao esta cadastrado
Public Const ERRO_DATA_SEM_PREENCHIMENTO = 2084
'A data deve estar preenchida
Public Const ERRO_MOVCONTACORRENTE_EXCLUIDO = 2085 'Parametro: Codconta, sequencial
'O movimento com a conta %i e sequencial %l esta excluido.
'Esse movimento est� excluido
Public Const ERRO_CONTACORRENTE_NAO_INFORMADA = 2086
'A conta Corrente deve ser informada
Public Const ERRO_VALOR_NEGATIVO = 2087
'O valor do saque deve ser positivo
Public Const ERRO_SEQUENCIAL_NAO_PREENCHIDO = 2088
'O Sequencial deve estar preenchido.
Public Const ERRO_TAMANHO_HISTORICOMOVCONTA = 2089
'O Historico deve ter no m�ximo 50 Caracteres
Public Const ERRO_NUMMOVTO_INEXISTENTE = 2090 'Parametro: lNumMovto
'O Movimento <paramentro> nao esta cadastrado
Public Const ERRO_MOVCONTACORRENTE_INEXISTENTE = 2091 'Parametros: CodContaCorrente, Sequencial
'N�o h� movimento cadastrado com a conta corrente %i e o sequencial %l.
Public Const ERRO_TIPOMEIOPAGTO_INEXISTENTE = 2092 'Parametro: iTipoMeioPagto
'O Tipo de Pagmento %i nao est� cadastrado.
Public Const ERRO_DATAMOVIMENTO_MENOR = 2093 'Parametro: dtDataMovimento, dtDataInicialConta, iCodConta
'A data do Movimento %dt � menor que a data inicial %dt da Conta Corrente %i.
Public Const ERRO_CCIMOV_NAO_CADASTRADO = 2094 'Parametros: iCodigo, iAno
'Nao ha movimentos cadastrados para a conta <iCodigo> no Ano de <iAno>
Public Const ERRO_CCIMOVDIA_NAO_CADASTRADO = 2095 'Parametros dtdata, icodigo
'Nao ha movimentos cadastrados no dia <data> para a conta <iCodigo>
Public Const ERRO_LOCK_CCIMOVDIA = 2096
'Erro na tentativa de fazer "lock" em registro da tabela CCIMovDia
Public Const ERRO_LOCK_CCIMOV = 2097
'Erro na tentativa de fazer "lock" em registro da tabela CCIMov
Public Const ERRO_TIPOMEIOPAGTO_NAO_INFORMADO = 2098
'A forma de pagamento n�o foi informada.
Public Const ERRO_TIPOCLIENTE_COD_NAO_PREENCHIDO = 2100
'O C�digo do Tipo de Cliente deve ser preenchido.
Public Const ERRO_TIPOCLIENTE_DESCR_NAO_PREENCHIDA = 2101
'A Descri��o do Tipo de Cliente deve ser preenchida.
Public Const ERRO_INSERCAO_TIPOCLIENTE = 2103
'Erro na tentativa de cadastro na tabela TiposDeClientes no Banco De Dados.
Public Const ERRO_MODIFICACAO_TIPOCLIENTE = 2104
'Erro na tentativa de modificar a tabela TiposDeClientes no Banco De Dados.
Public Const ERRO_FILIAL_DESASSOCIADA_FORNECEDOR = 2108 'Parametro: sCGC
'Filial de Fornecedor com CGC %s desassociada de Fornecedor.
Public Const ERRO_TIPOMOV_DIFERENTE = 2109  'Sem Parametro
'O Tipo do Movimento n�o coincide com o cadastrado.
Public Const ERRO_LEITURA_MOVIMENTOSCONTACORRENTE1 = 2110  'Parametros: iConta e lSequencial
'Erro na leitura da Conta %i e Sequencial %l da tabela MovimentosContaCorrente.
Public Const ERRO_DATA_COM_MOVIMENTOS = 2111 'Parametro: Data Inicial
'A data de Saldo Inicial: <dtdataInicial>, � menor que a data
'de alguns movimento para a conta em questao
Public Const ERRO_EXCLUSAO_CONTASCORRENTESINTERNAS = 2112
'Erro na exlusao de registro da Tabela de Contas Correntes Internas
Public Const ERRO_CHEQUEBORDERO_DIFERENTE_ZERO = 2113 'Parametro: icodconta
'A conta nao pode ser exclu�da pois est�
'sendo usada para emiss�o de cheques ou bordero
Public Const ERRO_AGENCIA_NAO_PREENCHIDA = 2114
'A Ag�ncia deve estar preenchida
Public Const ERRO_CODBANCO_NAO_INFORMADO = 2115
'O codigo do Banco deve ser informado
Public Const ERRO_NUMCONTA_NAO_PREENCHIDO = 2116
'O N�mero da conta deve estar preenchido
Public Const ERRO_INSERCAO_CONTASCORRENTESINTERNAS = 2117 'Parametro: Codigo
'Ocorreu um erro na tentativa de insercao da conta <codconta> na Tabela contascorrentesinternas
Public Const ERRO_MOVIMENTO_NAO_CADASTRADO = 2118 'Sem Parametro
'Movimento n�o cadastrado.
Public Const ERRO_LEITURA_EXTRATO_BCO = 2119 'Sem Parametro
'Erro na leitura da tabela de extratos banc�rios
Public Const ERRO_ALTERACAO_EXTRATO_BCO = 2120 'Sem Parametro
'Erro na alteracao da tabela de extratos banc�rios
Public Const ERRO_INSERCAO_EXTRATO_BCO = 2121 'Sem Parametro
'Erro na inser��o de registro na tabela de extratos banc�rios
Public Const ERRO_LEITURA_LCTO_EXTRATO_BCO = 2122 'Sem Parametro
'Erro na leitura da tabela de lan�amentos de extratos banc�rios
Public Const ERRO_INSERCAO_LCTO_EXTRATO_BCO = 2123 'Sem Parametro
'Erro inser��o de registro na tabela de lan�amentos de extratos banc�rios
Public Const ERRO_LEITURA_TABELA_PRECO = 2124 'Sem parametro
'Erro na leitura da tabela de Tabelas de Pre�o.
Public Const ERRO_TABELA_PRECO_NAO_ENCONTRADA = 2125 'Parametro sTabelaPreco
'Tabela de Pre�o com descri��o %s n�o foi encontrada.
Public Const ERRO_MENSAGEM_NAO_ENCONTRADA = 2126 'Parametro sMensagem
'A Mensagem %s n�o foi encontrada.
Public Const ERRO_COBRADOR_NAO_ENCONTRADO = 2127 'Parametro sCobrador
'O Cobrador %s n�o foi encontrado.
Public Const ERRO_TRANSPORTADORA_NAO_ENCONTRADA = 2128 'Par�metro: sNomeReduzido
'A Transportadora %s n�o foi encontrada.
Public Const ERRO_TIPOMEIOPAGTO_EXIGENUMERO = 2129 'Parametro : tipoMeiopagto
'O Tipo de Pagamento %i exige o preenchimento do campo numero.
Public Const ERRO_LEITURA_TABELA_HISTMOVCTA = 2130 'Sem par�metro
'Erro na leitura da tabela HistPadraoMovConta.
Public Const ERRO_LEITURA_TABELA_HISTMOVCTA1 = 2131 'Parametro Codigo do HistPadrao.
'Erro na leitura da tabela HistPadraoMovConta. Hist�rico = %i.
Public Const ERRO_ATUALIZACAO_HISTMOVCTA = 2132 'Parametro Codigo do Historico
'Erro de atualiza��o do Hist�rico de Movimenta��o de Conta %i.
Public Const ERRO_HISTMOVCTA_NAO_CADASTRADO = 2133 'Parametro Codigo do Hist�rico
'O Hist�rico de Movimenta��o de Conta %i n�o est� cadastrado.
Public Const ERRO_LOCK_HISTMOVCTA = 2134 'Parametro Codigo do Hist�rico
'N�o conseguiu fazer o lock do Hist�rico de Movimenta��o de Conta %i.
Public Const ERRO_EXCLUSAO_HISTMOVCTA = 2135 'Parametro Codigo do Hist�rico
'Houve um erro na exclus�o do Hist�rico de Movimenta��o de Conta %i.
Public Const ERRO_INSERCAO_HISTMOVCTA = 2136 'Par�metro C�digo do HistPadr�o
'Erro na inser��o do Hist�rico Padr�o %i na tabela HistPadraoMovConta.
Public Const ERRO_LEITURA_CONTASCORRENTESINTERNAS1 = 2137 'Parametro: iCodConta
'Erro na tentativa de leitura da Conta Corrente %i na tabela ContasCorrentesInternas.
Public Const ERRO_LEITURA_TIPOMEIOPAGTO1 = 2138 'Parametro iTipoMeioPagto
'Erro na leitura do Tipo de Pagamento %i da tabela TipoMeioPagto.
Public Const ERRO_CCI_IGUAIS = 2139
'As contas de Origem e Destino s�o iguais
Public Const ERRO_CONTADESTINO_NAO_DIGITADA = 2140
'O preenchimento daConta Destino � obrigatorio
Public Const ERRO_TIPOMEIOPAGTO_JA_UTILIZADO = 2141 ' Parametros: iCodConta, iTipoMeioPagto, lNumero
'A Conta %i j� utilizou a Forma de Pagamento %i de Numero %l em outro movimento.
Public Const ERRO_LEITURA_APLICACOES = 2142 'Par�metro C�digo do Tipo de aplica��o
'Erro na leitura da tabela de Aplica��es. Tipo de aplica��o = %s.
Public Const ERRO_LEITURA_TIPOSDEAPLICACAO = 2143 'Par�metro C�digo do Tipo de Aplica��o
'Erro na leitura da tabela de Tipos de Aplica��o. Aplica��o = %s.
Public Const ERRO_ATUALIZACAO_TIPOAPLICACAO = 2144 'Par�metro C�digo do Tipo de Aplica��o
'Erro na atualiza��o do Tipo de aplica��o %s.
Public Const ERRO_INSERCAO_TIPOAPLICACAO = 2145 'Par�metro C�digo do Tipo de Aplica��o
'Ocorreu um erro ao tentar inserir o Tipo de aplica�ao %s na tabela.
Public Const ERRO_EXCLUSAO_TIPOAPLICACAO = 2146 'Par�metro C�digo do Tipo de Aplica��o
'Ocorreu um erro ao tentar excluir o Tipo de aplica��o %s da tabela.
Public Const ERRO_TIPOAPLICACAO_INEXISTENTE = 2147 'Par�metro C�digo do Tipo de Aplica��o
'O Tipo de aplica��o %s n�o est� cadastrado.
Public Const ERRO_LOCK_TIPOAPLICACAO = 2148 'Par�metro C�digo do Tipo de Aplica��o
'Ocorreu um erro ao tentar locar o Tipo de aplica��o %s.
Public Const ERRO_CONTASIGUAIS = 2149 'Par�metro Conta
'As Contas Cont�bil Aplica��o e Receita Financeira s�o iguais : %s.
Public Const ERRO_EXCLUSAO_TIPOAPLICACAO_RELACIONADA = 2150 'Par�metro Descri��o do Tipo de aplica��o
'O Tipo de aplica��o %s n�o pode ser exclu�do pois existem uma ou mais aplica��es deste tipo.
Public Const ERRO_SEQUENCIAL_NAO_INFORMADO = 2151
'O Sequencial n�o foi informado
Public Const ERRO_TIPO_MOVIMENTO_NAO_DEPOSITO = 2152 'Parametros: lsequencial e icodconta
'O Movimento %l na conta %s n�o est� cadastrado como dep�sito.
Public Const ERRO_VALOR_INICIAL_MAIOR = 2153 'Sem Parametros
'O valor inicial n�o pode ser maior que o valor final.
Public Const ERRO_MOVIMENTOS_INEXISTENTES_CONCILIACAO = 2154 'Sem Parametros
'N�o existem movimentos de conta corrente para a sele��o atual.
Public Const ERRO_LEITURA_LCTOSEXTRATOBANCARIO = 2155 'Sem Parametros
'Erro na leitura da tabela de Lan�amentos de Extrato Banc�rio.
Public Const ERRO_LANCEXTRATO_INEXISTENTES_CONCILIACAO = 2156 'Sem Parametros
'N�o existem lan�amentos de extrato de conta corrente para a sele��o atual.
Public Const ERRO_GRIDS_MAIS_UM_ELEMENTO_SELECIONADO = 2157 'Sem parametros
'Ambos os grids cont�m mais de um elemento marcado. Pelo menos um dos grids s� pode ter um elemento selecionado.
Public Const ERRO_GRID_EXTRATO_SEM_SELECAO = 2158 'Sem parametros
'Nenhum lan�amento de extrato foi selecionado no grid.
Public Const ERRO_GRID_MOV_SEM_SELECAO = 2159 'Sem parametros
'Nenhum movimento de conta corrente foi selecionado no grid.
Public Const ERRO_LEITURA_LCTOSEXTRATOBANCARIO1 = 2160 'Parametros CodConta, NumExtrato, SeqLcto
'Erro na leitura da tabela de Lan�amentos de Extrato Banc�rio. Conta Corrente = %i, Extrato = %i, Sequencial = %l.
Public Const ERRO_ATUALIZACAO_LCTOSEXTRATOBANCARIO = 2161 'Parametros CodConta, NumExtrato, SeqLcto
'Erro na atualiza��o da tabela de Lan�amentos de Extrato Banc�rio. Conta Corrente = %i, Extrato = %i, Sequencial = %l.
Public Const ERRO_LEITURA_CONCILIACAOBANCARIA = 2162 'Parametros CodConta, Sequencial do Movto, NumExtrato, Sequencial no Extrato
'Erro na leitura da tabela de Concilia��o Banc�ria. Conta Corrente = %i, Sequencial do Movimento = %l, Extrato = %i, Sequencial no Extrato = %l.
Public Const ERRO_ATUALIZACAO_CONCILIACAOBANCARIA = 2163 'Parametros CodConta, Sequencial do Movto, NumExtrato, Sequencial no Extrato
'Erro na atualiza��o da tabela de Concilia��o Banc�ria. Conta Corrente = %i, Sequencial do Movimento = %l, Extrato = %i, Sequencial no Extrato = %l.
Public Const ERRO_ATUALIZACAO_MOVIMENTOSCONTACORRENTE1 = 2164 'Parametros CodConta, Sequencial
'Erro de Atualizacao da Tabela de MovimentosContaCorrente. Conta Corrente = %i, Sequencial = %l.
Public Const ERRO_INSERCAO_CONCILIACAOBANCARIA = 2165 'Parametros CodConta, Sequencial do Movto, NumExtrato, Sequencial no Extrato
'Erro na inser��o na tabela de Concilia��o Banc�ria. Conta Corrente = %i, Sequencial do Movimento = %l, Extrato = %i, Sequencial no Extrato = %l.
Public Const ERRO_GRID_EXT_JA_CONCILIADO = 2166 'Parametro: Posicao no Grid
'O Lan�amento de Extrato localizado na linha %i do grid j� est� conciliado.
Public Const ERRO_GRID_MOV_JA_CONCILIADO = 2167 'Parametro: Posicao no Grid
'O Movimento de Conta Corrente localizado na linha %i do grid j� est� conciliado.
Public Const ERRO_GRID_EXTRATO_MOV_SEM_SELECAO = 2168 'Sem parametros
'Nenhum lan�amento de extrato ou movimento de conta corrente foi selecionado no grid.
Public Const ERRO_LOCK_LCTOSEXTRATOBANCARIO = 2169 'Parametros CodConta, NumExtrato, SeqLcto
'Erro no lock de um registro da tabela de Lan�amentos de Extrato Banc�rio. Conta Corrente = %i, Extrato = %i, Sequencial = %l.
Public Const ERRO_LOCK_MOVIMENTOSCONTACORRENTE1 = 2170 'Parametros CodConta, Sequencial
'Erro na tentativa de fazer "lock" na tabela MOvimentosContasCorrente. Conta Corrente = %i e Sequencial = %l.
Public Const ERRO_LEITURA_CONCILIACAOBANCARIA1 = 2171 'Parametros CodConta, NumExtrato, Sequencial no Extrato
'Erro na leitura da tabela de Concilia��o Banc�ria. Conta Corrente = %i, Extrato = %i, Sequencial no Extrato = %l.
Public Const ERRO_LOCK_CONCILIACAOBANCARIA = 2172 'Parametros CodConta, Sequencial Mov, NumExtrato, Sequencial no Extrato
'Erro na tentativa de fazer "lock" na tabela de Concilia��o Banc�ria. Conta Corrente = %i, Sequencial do Movimento = %l, Extrato = %i, Sequencial no Extrato = %l.
Public Const ERRO_EXCLUSAO_CONCILIACAOBANCARIA = 2173 'Parametros CodConta, Sequencial Mov, NumExtrato, Sequencial no Extrato
'Erro na exclus�o de um registro da tabela de Concilia��o Banc�ria. Conta Corrente = %s, Sequencial do Movimento = %l, Extrato = %i, Sequencial no Extrato = %l.
Public Const ERRO_LEITURA_CONCILIACAOBANCARIA2 = 2174 'Parametros CodConta, Sequencial do Movto
'Erro na leitura da tabela de Concilia��o Banc�ria. Conta Corrente = %s, Sequencial do Movimento = %s.
Public Const ERRO_GRID_EXT_SEM_SELECAO = 2175 'Sem parametros
'Nenhum lan�amento de extrato foi selecionado no grid.
Public Const ERRO_PESQUISA_GRID_MOV_DATA = 2176 'Parametro Data
'N�o foi encontrado nenhum movimento no grid com a data = %s.
Public Const ERRO_EXT_SEM_MOV_CONCILIADO = 2177 'Parametros CodConta, NumExtrato, Sequencial no Extrato
'O Extrato em quest�o n�o est� associado (conciliado) com nenhum movimento de conta corrente. Conta Corrente = %s, Extrato = %i, Sequencial no Extrato = %l.
Public Const ERRO_MOV_SEM_EXT_CONCILIADO = 2178 'Parametros CodConta, Sequencial
'O Movimento de Conta Corrente em quest�o n�o est� associado (conciliado) com nenhum lan�amento de extrato. Conta Corrente = %i e Sequencial = %l.
Public Const ERRO_PESQUISA_GRID_MOV_VALOR = 2179 'Parametro Valor
'N�o foi encontrado nenhum movimento no grid com o valor = %s.
Public Const ERRO_PESQUISA_GRID_MOV_HISTORICO = 2180 'Parametro Historico
'N�o foi encontrado nenhum movimento no grid com o hist�rico = %s.
Public Const ERRO_PESQUISA_GRID_EXT_DATA = 2181 'Parametro Data
'N�o foi encontrado nenhum lan�amento de extrato no grid com a data = %s.
Public Const ERRO_PESQUISA_GRID_EXT_VALOR = 2182 'Parametro Valor
'N�o foi encontrado nenhum lan�amento de extrato no grid com o valor = %s.
Public Const ERRO_PESQUISA_GRID_EXT_HISTORICO = 2183 'Parametro Historico
'N�o foi encontrado nenhum lan�amento de extrato no grid com o hist�rico = %s.
Public Const ERRO_LEITURA_PARCELAS_PAG = 2184
'A parcela deveria estar baixada ou excluida, pois foi verificado que o saldo do t�tulo � zero.
Public Const ERRO_SEM_PARCELAS_PAG_SEL = 2185
'Ao menos uma Parcela tem que ser selecionada para pagamento.
Public Const ERRO_INSERCAO_TITULOS_PAG = 2186
'Erro na inser��o na tabela de T�tulos a Pagar.
Public Const ERRO_INSERCAO_PARCELAS_PAG = 2187
'Erro na inser��o na tabela de Parcelas a Pagar.
Public Const ERRO_EXCLUSAO_NOTAS_FISCAIS_EXT = 2188
'Erro na exclus�o de um registro da tabela de Notas Fiscais a Pagar.
Public Const ERRO_EXCLUSAO_TITULOS_PAGAR = 2189
'Erro na exclus�o de um registro da tabela de T�tulos a Pagar.
Public Const ERRO_INSERCAO_NOTAS_FISCAIS_EXT_BAIXADAS = 2190
'Erro na inser��o na tabela de Notas Fiscais Baixadas a Pagar.
Public Const ERRO_INSERCAO_BAIXA_PARC_PAG = 2191
'Erro na inser��o na tabela de Parcelas Baixadas a Pagar.
Public Const ERRO_UNLOCK_TITULOS_PAGAR = 2192
'Erro na tentativa de desfazer o "lock" na tabela de T�tulos a Pagar.
Public Const ERRO_UNLOCK_PARCELAS_PAGAR = 2193
'Erro na tentativa de desfazer o "lock" na tabela de Parcelas a Pagar.
Public Const ERRO_MODIFICACAO_PARCELAS_PAGAR = 2194
'Erro de Atualizacao da Tabela de Parcelas a Pagar.
Public Const ERRO_EXCLUSAO_PARCELAS_PAGAR = 2195
'Erro na exclus�o de um registro da tabela de Parcelas a Pagar.
Public Const ERRO_LOCK_TITULOS_PAGAR = 2196
'Erro na tentativa de fazer "lock" na tabela de T�tulos a Pagar.
Public Const ERRO_LOCK_PARCELAS_PAGAR = 2197
'Erro na tentativa de fazer "lock" na tabela de Parcelas a Pagar.
Public Const ERRO_PARCELA_PAGAR_NAO_ABERTA = 2198
'Parcela tem que estar aberta para poder ser baixada
Public Const ERRO_LEITURA_PORTADOR = 2199 'Parametro iCodigo
'Erro na leitura do Portador %i.
Public Const ERRO_PORTADOR_NAO_CADASTRADO = 2200 'Sem parametros
'Portador n�o est� cadastrado.
Public Const ERRO_NOME_REDUZIDO_PORTADOR_REPETIDO = 2201 'Sem parametros
'Nome Reduzido � atributo de outro Portador.
Public Const ERRO_INSERCAO_PORTADOR = 2202 'Parametro iCodigo
'Erro na inser��o do Portador %s.
Public Const ERRO_ATUALIZACAO_PORTADOR = 2203 'Parametro iCodigo
'Erro na atualiza��o do Portador %s.
Public Const ERRO_EXCLUSAO_GERACAO_CHEQUES = 2204
'Erro na exclus�o de um registro da tabela de Gera��o de Cheques.
Public Const ERRO_INSERCAO_GERACAO_CHEQUES = 2205
'Erro na inser��o na tabela de Gera��o de Cheques.
Public Const ERRO_LEITURA_PORTADORES = 2206
'Erro na leitura da tabela Portadores.
Public Const ERRO_LEITURA_BORDERO_PAG = 2207
'Erro na leitura da tabela de Bordero.
Public Const ERRO_MODIFICACAO_BORDERO_PAG = 2208
'Erro de Atualizacao da Tabela de Bordero.
Public Const ERRO_INSERCAO_BORDERO_PAG = 2209
'Erro na inser��o na tabela de Bordero.
Public Const ERRO_LEITURA_PORTADOR2 = 2210 'Sem parametros
'Erro na leitura da tabela Portador
Public Const ERRO_INSERCAO_BAIXAS_PAG = 2211 'Parametro lNumIntBaixa
'Erro de inser��o de registro na tabela BaixasPag. N�mero Interno = %l.
Public Const ERRO_LEITURA_BAIXAS_PAG = 2212 'Sem parametros
'Erro na leitura da tabela BaixasPag.
Public Const ERRO_FLUXO_DATAINI_MAIOR_DATAFIM = 2213 'Parametros DataFinal e DataBase
'A Data Final do fluxo de caixa = %s n�o pode ser menor do que a Data Base = %s.
Public Const ERRO_NOME_FLUXO_VAZIO = 2214
'O Nome do Fluxo n�o foi preenchido.
Public Const ERRO_DATAFINAL_FLUXO_VAZIO = 2215
'A Data Final do Fluxo de Caixa n�o foi preenchida.
Public Const ERRO_LEITURA_FLUXO = 2216 'Parametro Nome do Fluxo de Caixa
'Erro na leitura da tabela de Fluxo de Caixa. Fluxo = %s.
Public Const ERRO_ATUALIZACAO_FLUXO = 2217 'Parametro Nome do Fluxo de Caixa
'Erro na atualiza��o da tabela de Fluxo de Caixa. Fluxo = %s.
Public Const ERRO_MOV_NUMPAGTO_DUPLICADO = 2218 'Par�metros iCodconta, iTipomeioPagto, lNumero
'J� existe um documento com o mesmo n�mero. Conta = %i, Tipo de pagamento = %i e N�mero = %l.
Public Const ERRO_INSERCAO_FLUXO = 2219 'Parametro Nome do Fluxo de Caixa
'Erro na inser��o de um registro na tabela de Fluxo de Caixa. Fluxo = %s.
Public Const ERRO_CREDITOSPAGFORN_NUM_MAX = 2220 'Parametro Numero Maximo de Credito a Pagar que podem ser lidos
'Aten��o! Somente as informa��es oriundas dos primeiros %l registros de Creditos a Pagar ser�o exibidas.
Public Const ERRO_PORTADOR_NAO_CADASTRADO1 = 2221 'Parametro iCodPortador
'Portador %i n�o est� cadastrado.
Public Const ERRO_LOCK_PORTADOR = 2222 'Parametro iCodPortador
'Erro na tentativa de "lock" na tabela Portador, c�digo=%i.
Public Const ERRO_PORTADOR_INATIVO = 2223 'Parametro iCodPortador
'O Portador %i est� inativo.
Public Const ERRO_MOVIMENTO_NAO_CADASTRADO1 = 2224 'Parametros: iCodConta, iTipoMeioPagto, lNumero
'Movimento n�o cadastrado. Conta=%i, Tipo Meio Pagto=%i, N�mero=%l.
Public Const ERRO_MOVIMENTO_NAO_CADASTRADO2 = 2225 'Parametros: iCodConta, lSequencial
'Movimento n�o cadastrado. Conta=%i, Sequencial=%l.
Public Const ERRO_LEITURA_PAGTOANTECIPADO = 2226 'Parametro: lNumeroMovimento
'Erro na leitura da tabela PagtosAntecipados. NumMovto=%l.
Public Const ERRO_PAGTOANTECIPADO_NAO_CADASTRADO = 2227 'Parametro: lNumeroMovimento
'Pagamento Antecipado com NumMovto=%l n�o est� cadastrado.
Public Const ERRO_LEITURA_NFSPAG = 2228
'Erro na leitura da tabela Notas Fiscais a Pagar.
Public Const ERRO_INSERCAO_FLUXOANALITICO = 2230 'Sem Parametro
'Erro na inser��o de um registro na tabela de Fluxo de Caixa Anal�tico.
Public Const ERRO_BANCO_SEM_LAYOUT_CHEQUE = 2231 'Sem Parametro
'banco tem que ter layout de cheque definido
Public Const ERRO_FAVORECIDO_INEXISTENTE1 = 2232 'Parametro: sFavorecido
'O Favorecido %s n�o est� cadastrado
Public Const ERRO_MODIFICACAO_PARCELAS_REC = 2233
'Erro de Atualizacao da Tabela de Parcelas a Receber.
Public Const ERRO_UNLOCK_TITULOS_REC = 2235
'Erro na tentativa de desfazer o "lock" na tabela de T�tulos a Receber.
Public Const ERRO_UNLOCK_PARCELAS_REC = 2236
'Erro na tentativa de desfazer o "lock" na tabela de Parcelas a Receber.
Public Const ERRO_PARCELA_REC_NAO_ABERTA = 2237
'A parcela n�o est� aberta, portanto n�o pode ser baixada.
Public Const ERRO_TITULO_REC_INEXISTENTE = 2238
'O T�tulo n�o est� cadastrado.
Public Const ERRO_PARCELA_REC_INEXISTENTE = 2239
'A Parcela n�o est� cadastrada.
Public Const ERRO_MODIFICACAO_TITULOS_REC = 2240
'Erro de Atualizacao da Tabela de T�tulos a Receber.
Public Const ERRO_INSERCAO_BAIXA_PARC_REC = 2241
'Erro na inser��o de um registro na tabela de Parcelas Baixadas a Receber.
Public Const ERRO_EXCLUSAO_PARCELAS_RECEBER = 2242
'Erro na exclus�o de um registro da tabela de Parcelas a Receber.
Public Const ERRO_EXCLUSAO_NOTAS_FISCAIS_REC = 2243
'Erro na exclus�o de um registro da tabela de Notas Fiscais a Receber.
Public Const ERRO_INSERCAO_NOTAS_FISCAIS_REC_BAIXADAS = 2244
'Erro na inser��o de um registro na tabela de Notas Fiscais Baixadas a Receber.
Public Const ERRO_LEITURA_PARCELAS_REC1 = 2245
'Erro na leitura da tabela de Parcelas a Receber e Titulos a Receber.
Public Const ERRO_LEITURA_APLICACOES1 = 2246 'Sem Parametros
'Erro na leitura da tabela de Aplica��es.
Public Const ERRO_LEITURA_CCI_CCIMOV = 2247 'Sem parametros.
'Erro na tentativa de leitura das tabelas de Conta Corrente e Saldos Mensais de Conta Corrente.
Public Const ERRO_LEITURA_CCIMOVDIA1 = 2248 'Sem parametros.
'Erro de Leitura da Tabela de Saldos Diarios de Conta Corrente.
Public Const ERRO_INSERCAO_FLUXOTIPOAPLIC = 2249 'Sem Parametro
'Erro na inser��o de um registro na tabela de Fluxo de Caixa de Tipos de Aplica��o.
Public Const ERRO_INSERCAO_FLUXOAPLIC = 2250 'Sem Parametro
'Erro na inser��o de um registro na tabela de Fluxo de Caixa de Aplica��es.
Public Const ERRO_LEITURA_APLICACOES2 = 2251 'Parametro: lCodigo
'Erro na leitura da tabela de Aplica��es com C�digo %l.
Public Const ERRO_INSERCAO_FLUXOSALDOSINICIAIS = 2252 'Sem Parametro
'Erro na inser��o de um registro na tabela de Fluxo de Caixa Saldos Iniciais.
Public Const ERRO_LEITURA_FLUXOANALITICO = 2253 'Parametro: FluxoID
'Erro na leitura da tabela de Fluxo de Caixa Analitico. Fluxo = %l.
Public Const ERRO_INSERCAO_FLUXOFORN = 2254 'Sem Parametro
'Erro na inser��o de um registro na tabela de Fluxo de Caixa de Fornecedor (FluxoForn).
Public Const ERRO_INSERCAO_FLUXOTIPOFORN = 2255 'Sem Parametro
'Erro na inser��o de um registro na tabela de Fluxo de Caixa de Tipo de Fornecedor (FluxoTipoForn).
Public Const ERRO_INSERCAO_FLUXOSINTETICO = 2256 'Sem Parametro
'Erro na inser��o de um registro na tabela de Fluxo de Caixa Sintetico.
Public Const ERRO_FLUXO_NAO_CADASTRADO = 2257 'Parametro: Nome do fluxo
'O Fluxo %s n�o est� cadastrado.
Public Const ERRO_FLUXO_NAO_PREENCHIDO = 2258 'Sem parametro
'O Nome do Fluxo n�o foi informado.
Public Const ERRO_LEITURA_TITULOSARECEBER = 2259
'Erro na leitura da tabela de T�tulos a Receber.
Public Const ERRO_LOCK_FLUXO = 2260 'Parametro Nome do Fluxo de Caixa
'Ocorreu um erro ao tentar fazer lock no Fluxo de Caixa %s.
Public Const ERRO_EXCLUSAO_FLUXO = 2261 'Parametro Nome do Fluxo de Caixa
'Erro na exclus�o do Fluxo de Caixa %s.
Public Const ERRO_LEITURA_FLUXOFORN = 2262 'Parametro Nome do Fluxo de Caixa
'Erro na leitura da tabela de Fluxo de Caixa - FluxoForn. Fluxo = %s.
Public Const ERRO_EXCLUSAO_FLUXOFORN = 2263 'Parametro Nome do Fluxo de Caixa
'Erro na exclus�o do Fluxo de Caixa (FluxoForn) %s.
Public Const ERRO_LEITURA_FLUXOTIPOFORN = 2264 'Parametro Nome do Fluxo de Caixa
'Erro na leitura da tabela de Fluxo de Caixa - FluxoTipoForn. Fluxo = %s.
Public Const ERRO_EXCLUSAO_FLUXOTIPOFORN = 2265 'Parametro Nome do Fluxo de Caixa
'Erro na exclus�o do Fluxo de Caixa (FluxoTipoForn) %s.
Public Const ERRO_EXCLUSAO_FLUXOANALITICO = 2266 'Parametro Nome do Fluxo de Caixa
'Erro na exclus�o do Fluxo Anal�tico de Caixa. Fluxo = %s.
Public Const ERRO_LEITURA_FLUXOAPLIC = 2267 'Parametro Nome do Fluxo de Caixa
'Erro na leitura da tabela de Aplica��es de Fluxo de Caixa. Fluxo = %s.
Public Const ERRO_EXCLUSAO_FLUXOAPLIC = 2268 'Parametro Nome do Fluxo de Caixa
'Erro na exclus�o de uma Aplica��o do Fluxo de Caixa %s.
Public Const ERRO_LEITURA_FLUXOTIPOAPLIC = 2269 'Parametro Nome do Fluxo de Caixa
'Erro na leitura da tabela de Tipos de Aplica��o do Fluxo de Caixa. Fluxo = %s.
Public Const ERRO_EXCLUSAO_FLUXOTIPOAPLIC = 2270 'Parametro Nome do Fluxo de Caixa
'Erro na exclus�o de um Tipo de Aplica��o do Fluxo de Caixa  %s.
Public Const ERRO_LEITURA_FLUXOSALDOSINICIAIS = 2271 'Parametro Nome do Fluxo de Caixa
'Erro na leitura da tabela de Saldos Iniciais do Fluxo de Caixa. Fluxo = %s.
Public Const ERRO_EXCLUSAO_FLUXOSALDOSINICIAIS = 2272 'Parametro Nome do Fluxo de Caixa
'Erro na exclus�o dos Saldos Iniciais do Fluxo de Caixa %s.
Public Const ERRO_LEITURA_FLUXOSINTETICO = 2273 'Parametro Nome do Fluxo de Caixa
'Erro na leitura da tabela de Fluxos Sint�ticos. Fluxo = %s.
Public Const ERRO_EXCLUSAO_FLUXOSINTETICO = 2274 'Parametro Nome do Fluxo de Caixa
'Erro na exclus�o de um Fluxo Sint�tico. Fluxo = %s.
Public Const ERRO_LEITURA_FLUXOANALITICO1 = 2275 'Parametro: FluxoID, Data, TipoReg
'Erro na leitura da tabela de Fluxo de Caixa Analitico. Fluxo = %l, Data = %s e Tipo do Registro = %i.
Public Const ERRO_TIPO_NAO_ANTECIPPAG = 2276 'Sem par�metros
'O Movimento n�o � do tipo Pagamento Antecipado.
Public Const ERRO_LEITURA_ANTECIPPAG = 2277 'Par�metro lNumIntPag
'Erro na leitura do Pagamento antecipado %l.
Public Const ERRO_LEITURA_ANTECIPPAG1 = 2278 'Par�metro lNumIntMovto
'Erro na leitura do Pagamento antecipado cujo N�mero de movimento �: %l.
Public Const ERRO_INSERCAO_ANTECIPPAG = 2279 'Par�metro: lNumIntPag
'Ocorreu um erro na tentativa de inser��o do Pagamento antecipado %l na Tabela PagtosAntecipados.
Public Const ERRO_EXCLUSAO_ANTECIPPAG = 2280 'Par�metro: lNumIntPag
'Ocorreu um erro na tentativa de exclus�o do Pagamento antecipado %l.
Public Const ERRO_LOCK_ANTECIPPAG = 2281 'Par�metro lNumIntPag
'Erro na tentativa de fazer o "lock" do Pagamento antecipado %l.
Public Const ERRO_FORNECEDOR_NAO_PREENCHIDO = 2282 'Sem par�metro
'O Fornecedor deve ser preenchido.
Public Const ERRO_EXCLUSAO_MOVIMENTOSCONTACORRENTE1 = 2284 'Par�metro iCodConta + lSequencial
'Erro na exclusao do Movimento. Conta = %i e Sequencial = %l.
Public Const ERRO_ANTECIPPAG_EXCLUIDO = 2285 'Par�metros iCodConta + lSequencial
'O Pagamento antecipado com a Conta %i e Sequencial %l est� exclu�do.
Public Const ERRO_PAGAMENTO_APROPRIADO = 2286 'Par�metro iCodigo
'N�o � poss�vel excluir o Pagamento antecipado %i, pois j� foi apropriado (total ou parcialmente).
Public Const ERRO_FILIAL_NAO_ENCONTRADA = 2287 'Par�metro sFilial
'Filial com descri��o %s n�o foi encontrada.
Public Const ERRO_FORNECEDOR_NAO_COINCIDE = 2288 'Par�metros lFornecedor(da tela) e lFornecedor(da tabela)
'O Fornecedor %l n�o coincide com o Fornecedor %l cadastrado no Pagamento Antecipado
Public Const ERRO_FILIAL_NAO_COINCIDE = 2289 'Par�metros lFilial(da tela) e lFilial(da tabela)
'A Filial %l n�o coincide com a Filial %l cadastrada no Pagamento Antecipado.
Public Const ERRO_ANTECIPPAG_INEXISTENTE = 2290 'Par�metros iCodConta + lSequencial
'O Movimento com a Conta %i e Sequencial %l n�o est� cadastrado.
Public Const ERRO_FILIALFORNECEDOR_INEXISTENTE = 2291 'Parametro: Filial
'A filial %s n�o est� cadastrada.
Public Const ERRO_NUMERO_NAO_INFORMADO = 2293 'Par�metro: iTipo
'Um n�mero de documento tem que ser informado para o Tipo de pagamento %i.
Public Const ERRO_VALOR_MENOR_UM = 2295 'Parametro: dValor
'O Valor %d � menor do que 1.
Public Const ERRO_NFPAG_NAO_CADASTRADA = 2296 'Parametro: lNumIntDoc
'A Nota Fiscal com n�mero interno %l n�o est� cadastrada.
Public Const ERRO_NFPAG_NAO_CADASTRADA1 = 2297 'Parametro: lNumNotaFiscal
'A Nota Fiscal %l n�o est� cadastrada.
Public Const ERRO_DATAVENCIMENTO_MENOR = 2298
'A Data de Vencimento � menor do que a Data de Emiss�o.
Public Const ERRO_LOCK_FILIALFORNECEDOR = 2301 'Parametros: lCodFornecedor, iCodFilial.
' Erro na tentativa de "lock" da tabela FiliaisFornecedores. CodFornecedor=%l, CodFilial=%i.
Public Const ERRO_NF_PENDENTE_MODIFICACAO = 2302 'parametro: lNumNotaFiscal
'N�o � poss�vel modificar a Nota Fiscal %l. Ela faz parte de um Lote Pendente.
Public Const ERRO_NF_PENDENTE_EXCLUSAO = 2303 'parametro: lNumNotaFiscal
'N�o � poss�vel excluir a Nota Fiscal %l. Ela faz parte de um Lote Pendente.
Public Const ERRO_LEITURA_NFSPAGPEND = 2304 'Par�metro: lNumNotaFiscal
'Erro na tentativa de leitura da Nota Fiscal %l na tabela NfsPagPend
Public Const ERRO_NF_FILIALEMPRESA_DIFERENTE = 2305 'Par�metro: lNumNotaFiscal
'N�o � poss�vel modificar a Nota Fiscal %l. Ela pertence a outra Filial da Empresa.
Public Const ERRO_LEITURA_NFSPAGBAIXADAS = 2307 'Par�metro: lNumnotafiscal
'Erro na leitura da Nota Fiscal n�mero %l na tabela NfsPagBaixadas.
Public Const ERRO_NF_BAIXADA_MODIFICACAO = 2308 'Par�metro: lNumNotaFiscal
'N�o � poss�vel modificar a Nota Fiscal %l porque ela est� baixada.
Public Const ERRO_NF_BAIXADA_EXCLUSAO = 2309 'Par�metro: lNumNotaFiscal
'N�o � poss�vel excluir a Nota Fiscal %l porque ela est� baixada.
Public Const ERRO_INSERCAO_NFSPAG = 2310 'Par�metro: lNumNotaFiscal
'Erro na tentativa de inserir a Nota Fiscal n�mero %l na tabela NfsPag.
Public Const ERRO_NF_NAO_INFORMADA = 2311
'O campo n�mero da Nota Fiscal n�o foi preenchido.
Public Const ERRO_VALORTOTAL_NAO_INFORMADO = 2312
'O campo Valor n�o foi preenchido.
Public Const ERRO_VALORPRODUTOS_NAO_INFORMADO = 2313
'O campo Valor dos Produtos n�o foi preenchido.
Public Const ERRO_VALORTOTAL_INVALIDO = 2314 'Parametros: sValorTotal, dValor
'O Valor Total %s n�o � igual � soma dos valores de Frete, Produtos, ICMS Subst, Seguro, IPI, Outras Despesas que � de %d.
Public Const ERRO_DATAEMISSAO_MAIOR = 2315 'Parametros: sDataEmissao, sDataVencimento
'A Data de Emiss�o %s � maior que a Data de Vencimento %s.
Public Const ERRO_LEITURA_FILIAISFORNECEDORES2 = 2319 'Parametro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela FiliaisFornecedores. CodFornecedor = %l e CodFilial = %i.
Public Const ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA = 2320 'Par�metros: sFilialNome
'Filial de Fornecedor %s n�o foi encontrada.
Public Const ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA = 2321 'Parametro: sNomeReduzido
'A Condi��o de Pagamento %s n�o foi encontrada.
Public Const ERRO_CONDICAO_PAGTO_NAO_CADASTRADA1 = 2322 'Parametro sDescReduzida
'A Condi��o de Pagamento %s n�o est� cadastrada no Banco de Dados.
Public Const ERRO_REGIAO_VENDA_NAO_CADASTRADA1 = 2324 'Parametro: sDescricao
'Regi�o de Venda %s n�o est� cadastrada no Banco de Dados.
Public Const ERRO_TIPO_CLIENTE_NAO_PREENCHIDO = 2325 'Sem parametros
'Preenchimento do Tipo de Cliente � obrigat�rio.
Public Const ERRO_SEM_PARCELAS_REC_SEL = 2326
'N�o achou nenhuma parcela a receber dentro dos crit�rios informados.
Public Const ERRO_INSERCAO_BORDERO_COBRANCA = 2327
'Erro na inser��o de Bordero de Cobran�a.
Public Const ERRO_INSERCAO_OCORR_REM_PARC_REC = 2328
'Erro na inser��o de ocorr�ncia de remessa de Bordero de Cobran�a.
Public Const ERRO_ATUALIZACAO_CARTEIRAS_COBRADOR = 2329
'Erro na atualiza��o da tabela Carteiras Cobrador.
Public Const ERRO_PARCELA_NAO_ABERTA = 2330
'A parcela tem que estar aberta.
Public Const ERRO_COBRADOR_JA_DEFINIDO = 2331
'O cobrador tem que estar em aberto.
Public Const ERRO_BORD_COBR_VENCTO_ALTERADO = 2332
'A data de vencimento foi alterada ap�s a sele��o das parcelas.
Public Const ERRO_BORD_COBR_VALOR_ALTERADO = 2333
'O valor a ser cobrado foi alterado ap�s a sele��o das parcelas.
Public Const ERRO_LEITURA_MOVSCCIFLUXO = 2334
'Erro na leitura do Fluxo de Movimento de Conta Corrente.
Public Const ERRO_LEITURA_PAGAMENTOS_PARA_FLUXO = 2335
'Erro na pesquisa de Pagamentos para consulta de Fluxo de Caixa.
Public Const ERRO_LEITURA_PAGTOANTEC_PARA_FLUXO = 2336
'Erro na pesquisa de Pagamentos antecipados para consulta de Fluxo de Caixa.
Public Const ERRO_LEITURA_RECEBTOANTEC_PARA_FLUXO = 2337
'Erro na pesquisa de Recebimentos antecipados para consulta de Fluxo de Caixa.
Public Const ERRO_LEITURA_CREDPAGFORN_PARA_FLUXO = 2338
'Erro na pesquisa de Cr�ditos a Pagar de Fornecedores para consulta de Fluxo de Caixa.
Public Const ERRO_LEITURA_DEBRECCLI_PARA_FLUXO = 2339
'Erro na leitura de Fluxo para D�bitos a Receber de Clientes.
Public Const ERRO_LEITURA_FLUXOTIPOFORN2 = 2340
'Erro na leitura da tabela de Fluxo de Tipo de Fornecedor.
Public Const ERRO_ATUALIZACAO_FLUXOTIPOFORN = 2341
'Erro na tentativa de atualiza��o da tabela de Fluxo de Tipo de Fornecedor.
Public Const ERRO_LEITURA_FLUXOFORN2 = 2342
'Erro na leitura da tabela de Fluxo de Fornecedores.
Public Const ERRO_ATUALIZACAO_FLUXOFORN = 2343
'Erro na tentativa de atualiza��o da tabela de Fluxo de Fornecedores.
Public Const ERRO_LEITURA_RECEBTOS_PARA_FLUXO = 2344
'Erro na pesquisa de Recebimentos para consulta de Fluxo de Caixa.
Public Const ERRO_ATUALIZACAO_FLUXOTIPOAPLIC = 2345
'Erro na tentativa de atualiza��o da tabela de Fluxo de Tipo de Aplica��o.
Public Const ERRO_LEITURA_RESGATESFLUXO = 2346
'Erro na leitura de Resgates de Fluxo.
Public Const ERRO_ATUALIZACAO_FLUXOSINTETICO = 2347
'Erro na tentativa de atualiza��o da tabela de Fluxo Sint�tico.
Public Const ERRO_LEITURA_FLUXOFORN1 = 2348 'Parametros Nome do Fluxo de Caixa, Data, Tipo do Registro
'Erro na leitura da tabela de Fluxo de Caixa - FluxoForn. Fluxo = %s, Data=%s e Tipo do Registro = %i.
Public Const ERRO_TIPO_FORNECEDOR_NAO_CADASTRADO = 2349 'Parametro: iCodigo
'Tipo de Fornecedor com c�digo %i n�o est� cadastrado no BD.
Public Const ERRO_FLUXO_DATA_FORA_FAIXA = 2350 'Parametros Data, DataBase do Fluxo e DataFinal do Fluxo
'A Data em quest�o %s est� fora da faixa abrangida pelo fluxo de caixa. Data Base = %s e Data Final = %s.
Public Const ERRO_PORTADOR_NAO_INFORMADO = 2351
'O Portador deve ser informado.
Public Const ERRO_PROXCHEQUE_NAO_INFORMADO = 2352
'O Pr�ximo cheque deve ser informado.
Public Const ERRO_DATACONTABIL_MENOR_DATAEMISSAO = 2353
'A Data Cont�bil deve ser maior ou igual � Data de Emiss�o.
Public Const ERRO_CONTACORRENTE_NAO_BANCARIA = 2354
'A Conta Corrente deve ser Banc�ria.
Public Const ERRO_TITULO_NAO_PREENCHIDO = 2355 'Sem par�metro
'O T�tulo deve ser preenchido.
Public Const ERRO_PARCELA_NAO_PREENCHIDA = 2356 'Sem par�metro
'A Parcela deve ser preenchida.
Public Const ERRO_FILIALCLIENTE_REL_NF_REC_PEND = 2357 'Parametros: lCodCliente, iCodFilial
'Erro na exclus�o da Filial de Cliente com CodCliente=%l, CodFilial=%i. Est� relacionada com Nota Fiscal a Receber Pendente.
Public Const ERRO_FILIALCLIENTE_REL_NF_REC = 2358 'Parametros: lCodCliente, iCodFilial
'Erro na exclus�o da Filial de Cliente com CodCliente=%l, CodFilial=%i. Est� relacionada com Nota Fiscal a Receber.
Public Const ERRO_FILIALCLIENTE_REL_NF_REC_BAIXADA = 2359 'Parametros: lCodCliente, iCodFilial
'Erro na exclus�o da Filial de Cliente com CodCliente=%l, CodFilial=%i. Est� relacionada com Nota Fiscal a Receber Baixada.
Public Const ERRO_LEITURA_FORNECEDORES_NOMEREDUZIDO = 2360 'Parametro Nome Reduzido
'Erro na leitura da tabela de Forncedores. Nome Reduzido = %s.
Public Const ERRO_LEITURA_FLUXO1 = 2361 'Parametro FluxoId
'Erro na leitura da tabela de Fluxo de Caixa. FluxoId = %l.
Public Const ERRO_FLUXO_NAO_CADASTRADO1 = 2362 'Parametro: FluxoId
'O Fluxo n�o est� cadastrado. FluxoId = %l.
Public Const ERRO_LOCK_FLUXO1 = 2363 'Parametro FluxoId
'Ocorreu um erro ao tentar fazer lock no Fluxo de Caixa. FluxoId = %l.
Public Const ERRO_LEITURA_FLUXOFORN3 = 2364 'Parametros FluxoId, Tipo do Registro, Fornecedor, Data
'Erro na leitura da tabela de Fluxo de Caixa - FluxoForn. FluxoId = %l, Tipo do Registro = %i, C�digo do Fornecedor = %l e Data=%s
Public Const ERRO_EXCLUSAO_FLUXOFORN1 = 2365 'Parametros FluxoId, Tipo do Registro, Fornecedor, Data
'Erro na exclus�o do Fluxo de Caixa (FluxoForn). FluxoId = %l, Tipo do Registro = %i, C�digo do Fornecedor = %l e Data=%s.
Public Const ERRO_LEITURA_FLUXOTIPOFORN1 = 2366 'Parametros FluxoId, Tipo do Registro, Tipo do Fornecedor, Data
'Erro na leitura da tabela de Fluxo de Caixa - FluxoTipoForn. FluxoId = %l, Tipo do Registro = %i, Tipo do Fornecedor = %i e Data=%s.
Public Const ERRO_ATUALIZACAO_FLUXOTIPOFORN1 = 2367 'Parametros FluxoId, Tipo do Registro, Tipo do Fornecedor, Data
'Erro na tentativa de atualiza��o da tabela de Fluxo de Tipo de Fornecedor. FluxoId = %l, Tipo do Registro = %i, Tipo do Fornecedor = %i e Data=%s.
Public Const ERRO_ATUALIZACAO_FLUXOFORN1 = 2368 'Parametros FluxoId, Tipo do Registro, Fornecedor, Data
'Erro na tentativa de atualiza��o da tabela de Fluxo de Caixa - FluxoForn. FluxoId = %l, Tipo do Registro = %i, C�digo do Fornecedor = %l e Data=%s
Public Const ERRO_LEITURA_TIPOSFORNECEDOR1 = 2369 'Parametro iTipoForn
'Erro na leitura da tabela TiposdeFornecedor. Tipo do Fornecedor = %i.
Public Const ERRO_LEITURA_FLUXOSINTETICO1 = 2370 'Parametros FluxoId e Data
'Erro na leitura da tabela de Fluxos Sint�ticos. FluxoId = %l e Data = %s.
Public Const ERRO_ATUALIZACAO_FLUXOSINTETICO1 = 2371 'Parametros FluxoId e Data
'Erro na tentativa de atualiza��o da tabela de Fluxo Sint�tico. FluxoId = %l e Data = %s.
Public Const ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO = 2372 'Parametro: iCodigo
'Condi��o de Pagamento com c�digo %i n�o � de Contas a Pagar.
Public Const ERRO_LEITURA_TIPOSCLIENTE1 = 2373 'Parametro iTipoCli
'Erro na leitura da tabela TiposdeCliente. Tipo do Cliente = %i.
Public Const ERRO_LEITURA_CLIENTES_NOMEREDUZIDO = 2374 'Parametro: sNomeReduzido
'Erro na leitura da tabela de Clientes. Nome Reduzido = %s.
Public Const ERRO_CCI_INSERCAO_EMPRESA_TODA = 2376 'Sem par�metro
'O usu�rio tem que ter selecionado uma filial ao se conectar ao sistema para poder incluir uma conta corrente
Public Const ERRO_CHEQUE_PRE_DIF_VALOR = 2377 'Sem par�metro
'O cheque pr�-datado tem valor diferente do necess�rio para pagar as parcelas associadas � ele
Public Const ERRO_INSERCAO_BAIXAS_REC = 2378 'Sem par�metro
'Erro na inser��o de baixa de t�tulo a receber
Public Const ERRO_INSERCAO_BORDERO_CHEQUE_PRE = 2379 'Sem par�metro
'Erro na inser��o de bordero de cheques pr�-datados
Public Const ERRO_LEITURA_CHEQUES_PRE_BORDERO = 2380 'Sem par�metro
'erro na leitura de cheques pr�-datados para o border�
Public Const ERRO_LEITURA_BORDERO_SEM_CHEQUES_PRE = 2381 'Sem par�metro
'N�o h� cheque pr�-datado a depositar at� a data informada
Public Const ERRO_LEITURA_VENDEDOR1 = 2382 'Parametro: sNomeReduzido
'Erro na leitura do Vendedor com Nome Reduzido %s na tabela de Vendedores.
Public Const ERRO_SEM_COMISSOES_BAIXA = 2384    'Sem par�metro
'N�o h� comiss�es a serem baixadas, verifique os par�metros informados.
Public Const ERRO_GRAVACAO_BAIXA_COMISSAO = 2385    'Sem par�metro
'Erro na grava��o da baixa de uma comiss�o
Public Const ERRO_NFFAT_FILIALEMPRESA_DIFERENTE = 2386 'Par�metro: lNumNotaFiscalFatura
'N�o � poss�vel modificar a Nota Fiscal Fatura %l. Ela pertence a outra Filial da Empresa.
Public Const ERRO_FATURA_FILIALEMPRESA_DIFERENTE = 2387 'Par�metro: lNumFatura
'N�o � poss�vel modificar a Fatura %l. Ela pertence a outra Filial da Empresa.
Public Const ERRO_OUTROPAG_FILIALEMPRESA_DIFERENTE = 2388 'Par�metros: sSiglaDocumento, lNumTitulo
'N�o � poss�vel modificar o T�tulo a Pagar do tipo %s de n�mero %l. Ele pertence a outra Filial da Empresa.
Public Const ERRO_NFFATPAG_NAO_CADASTRADA = 2389 'Par�metro: lNumIntDoc
'A Nota Fiscal Fatura com N�mero Interno %l n�o est� cadastrada
Public Const ERRO_NFFATPAG_NAO_CADASTRADA1 = 2390 'Par�metro: lNumTitulo
'A Nota Fiscal Fatura n�mero %l n�o est� cadastrada
Public Const ERRO_TITULO_NAO_NFFATPAG = 2391 'Par�metro: lNumIntDoc
'O T�tulo de n�mero interno %l n�o � Nota Fiscal Fatura.
Public Const ERRO_DATAVENCIMENTO_PARCELA_MENOR = 2392 'Par�metros: sDataVencimento, sDataEmissao, iParcela
'A Data de Vencimento %s da Parcela %i � menor do que a Data de Emiss�o %s.
Public Const ERRO_AUSENCIA_PARCELAS_GRAVAR = 2393
'N�o existem parcelas no Grid de Parcelas para gravar.
Public Const ERRO_DATAVENCIMENTO_PARCELA_NAO_INFORMADA = 2394 'Par�metro: iParcela
'O campo Data de Vencimento da Parcela %i n�o foi preenchido.
Public Const ERRO_DATAVENCIMENTO_NAO_ORDENADA = 2395 'Sem parametros
'As Datas de Vencimento no Grid n�o est�o ordenadas.
Public Const ERRO_SOMA_PARCELAS_INVALIDA = 2396 'Par�metros: dSomaParcelas, dValorPagar
'A soma das Parcelas %d n�o � igual ao Valor a Pagar %d.
Public Const ERRO_NFFATPAG_BAIXADA_MODIFICACAO = 2397 'Par�metro: lNumNotaFiscal
'N�o � poss�vel modificar Nota Fiscal Fatura n�mero %l porque ela est� baixada.
Public Const ERRO_NFFATPAG_PENDENTE_MODIFICACAO = 2398 'Par�metro: lNumNotaFiscal
'N�o � poss�vel modificar Nota Fiscal Fatura n�mero %l .Ela faz parte de um Lote Pendente.
Public Const ERRO_LEITURA_TITULOSPAG = 2399 'Par�metro: lNumIntDoc
'Erro na tentativa de ler registro na tabela TitulosPag com N�mero Interno %l.
Public Const ERRO_LEITURA_NFFATURA = 2400 'Par�metro: lNumTitulo
'Erro na tentativa de ler Nota Fiscal Fatura %l na tabela TitulosPag.
Public Const ERRO_INSERCAO_NFFATURA = 2401 'Par�metro: lNumTitulo
'Erro na tentativa de inserir a Nota Fiscal Fatura n�mero %l na Tabela TitulosPag.
Public Const ERRO_LEITURA_PARCELASPAG = 2402 'Par�metro: lNumIntTitulo
'Erro na tentativa de ler Parcelas com NumIntTitulo %l na tabela ParcelasPag.
Public Const ERRO_LEITURA_PARCELASPAG1 = 2403 'Par�metros: lNumIntTitulo, iNumParcela
'Erro na tentativa de ler Parcela %i com NumIntTitulo %l na tabela ParcelasPag.
Public Const ERRO_PARCELA_PAGAR_NAO_CADASTRADA = 2404 'Par�metros: iNumParcela
'A Parcela %i deste T�tulo n�o foi encontrada.
Public Const ERRO_LOCK_PARCELASPAG = 2405 'Par�metros: lNumIntTitulo, iNumParcela
'Erro na tentativa de "lock" na Parcela %i com NumIntTitulo %l na tabela ParcelasPag.
Public Const ERRO_ATUALIZACAO_PARCELASPAG = 2406 'Par�metros: lNumIntTitulo, iNumParcela
'Erro na atualiza��o da Parcela %i com NumIntTitulo %l na tabela ParcelasPag.
Public Const ERRO_INSERCAO_PARCELASPAG = 2407 'Par�metro: lNumIntTitulo, iNumParcela
'Erro na tentativa de inser��o para o t�tulo com NumIntTitulo %l da Parcela %i na tabela ParcelasPag.
Public Const ERRO_NFFATPAG_PENDENTE_EXCLUSAO = 2408 'Par�metro: lNumTitulo
'N�o � poss�vel excluir a Nota Fiscal Fatura n�mero %l porque ela faz parte de um Lote Pendente.
Public Const ERRO_NFFATPAG_BAIXADA_EXCLUSAO = 2409 'Par�metro: lNumTitulo
'N�o � poss�vel excluir a Nota Fiscal Fatura n�mero %l porque ela est� baixada.
Public Const ERRO_LOCK_TITULOSPAG = 2410 'Par�metro: lNumIntDoc
'Erro na tentativa de "lock" no T�tulo com NumIntDoc = %l na tabela TitulosPag.
Public Const ERRO_ATUALIZACAO_TITULOSPAG = 2411 'Par�metro: lNumIntDoc
'Erro na atualizacao do T�tulo com N�mero Interno %l na tabela TitulosPag.
Public Const ERRO_TIPOCOBRANCA_NAO_ENCONTRADO = 2412 'Par�metro: sDescricao
'O Tipo de Cobran�a %s n�o foi encontrado.
Public Const ERRO_LEITURA_NFFATPEND = 2413 'Par�metro: lNumTitulo
'Erro na tentativa de leitura da Nota Fiscal Fatura %l na tabela TitulosPagPend.
Public Const ERRO_LEITURA_NFFATBAIXADA = 2414 'Par�metro: lNumTitulo
'Erro na tentativa de leitura da Nota Fiscal Fatura %l na tabela TitulosPagBaixados.
Public Const ERRO_NUMTITULO_NAO_PREENCHIDO = 2415 'Sem parametros
'O campo N�mero n�o foi preenchido
Public Const ERRO_NUM_MAXIMO_PARCELAS_ULTRAPASSADO = 2416 'Par�metros: iNumParcelasBD, iNumMaxParcelas
'O n�mero de parcelas lidas no BD � %i que supera o n�mero m�ximo permitido igual a %i.
Public Const ERRO_TIPOCOBRANCA_NAO_CADASTRADO = 2417 'Par�metro: iCodigo
'O Tipo de Cobran�a com C�digo %i n�o est� cadastrado.
Public Const ERRO_VALORPARCELA_NAO_INFORMADO = 2418 'Par�metro: iParcela
'O Valor da Parcela %i n�o foi informado.
Public Const ERRO_TITULOPAGAR_SEM_PARCELAS = 2419 'Par�metro: lNumIntDoc
'T�tulo a Pagar com n�mero interno %l n�o tem Parcelas associadas.
Public Const ERRO_PAGTO_ANTECIPADO_INEXISTENTE = 2420
'O Pagamento Antecipado n�o est� cadastrado.
Public Const ERRO_MODIFICACAO_PAGTO_ANTECIPADO = 2421
'Erro na tentativa de modificar a tabela de Pagamentos Antecipados.
Public Const ERRO_LEITURA_PAGTO_ANTECIPADO2 = 2422
'Erro na leitura da tabela de Pagamentos Antecipados.
Public Const ERRO_CREDITO_PAG_FORN_INEXISTENTE = 2423
'O Cr�dito n�o est� cadastrado.
Public Const ERRO_MODIFICACAO_CREDITO_PAG_FORN = 2424
'Erro na tentativa de modificar a tabela de Cr�ditos a Pagar.
Public Const ERRO_CREDITO_PAG_FORN_EXCLUIDO = 2425
'Esse Cr�dito est� exclu�do.
Public Const ERRO_LEITURA_BAIXAPAG = 2426
'Erro na leitura da tabela de Baixas a Pagar.
Public Const ERRO_BAIXAPAG_INEXISTENTE = 2427
'A Baixa n�o est� cadastrada.
Public Const ERRO_EXCLUSAO_BAIXAPAG = 2428
'Erro na exclus�o da Baixa a Pagar.
Public Const ERRO_BAIXAPAG_EXCLUIDA = 2429
'A baixa j� havia sido cancelada anteriormente.
Public Const ERRO_PAGTO_ANTECIPADO_EXCLUIDO = 2430
'O pagamento antecipado j� havia sido exclu�do.
Public Const ERRO_BAIXAPARCPAG_INEXISTENTE = 2431
'A Baixa n�o est� cadastrada.
Public Const ERRO_BAIXAPARCPAG_EXCLUIDA = 2432
'A baixa da parcela j� havia sido cancelada anteriormente."
Public Const ERRO_EXCLUSAO_ANTECIPPAG2 = 2433
'Erro na exclus�o do Pagamento Antecipado.
Public Const ERRO_BORDERO_PAGTO_EXCLUIDO = 2434
'Esse Border� de Pagamento est� exclu�do.
Public Const ERRO_EXCLUSAO_BORDERO_PAGTO = 2435
'Erro na exclus�o do Border� de Pagamento.
Public Const ERRO_BORDERO_PAGTO_INEXISTENTE = 2436
'O Border� de Pagamento n�o est� cadastrado.
Public Const ERRO_LEITURA_BORDERO_PAGTO = 2437
'Erro na leitura da tabela de Border� de Pagamento.
Public Const ERRO_ANTECIPPAG_EXCLUIDO1 = 2438 'Parametro: lNumMovto
'O pagamento antecipado associado ao movimento %l de conta corrente j� est� marcado como exclu�do.
Public Const ERRO_ANTECIPPAG_INEXISTENTE2 = 2439 'Parametro: lNumMovto
'Erro na leitura do pagamento antecipado associado ao movimento %l de conta corrente.
Public Const ERRO_MOVIMENTOSCONTACORRENTE_INEXISTENTE = 2440
'O Movimento de Conta Corrente n�o est� cadastrado.
Public Const ERRO_LEITURA_PARCELAS_PAG2 = 2441
'Erro na leitura da tabela de Parcelas a Pagar.
Public Const ERRO_PAGTO_CONCILIADO = 2442
'Um pagamento conciliado n�o pode ser cancelado. Desconcilie-o antes de tentar exclu�-lo.
Public Const ERRO_NUMPAGTO_NAO_INFORMADO = 2443
'O N�mero de Pagamento n�o foi informado.
Public Const ERRO_MEIOPAGTO_NAO_INFORMADO = 2444
'O Meio de Pagamento n�o foi informado.
Public Const ERRO_LEITURA_BAIXAS_MOVTO_CTA = 2445
'Erro na leitura de baixas associadas a um movimento de conta corrente.
Public Const ERRO_LOCK_BAIXAPAG = 2446
'N�o conseguiu fazer o lock da Baixa a Pagar.
Public Const ERRO_UNLOCK_BAIXAPAG = 2447
'Erro na tentativa de desfazer o lock na tabela de Baixas a Pagar.
Public Const ERRO_MODIFICACAO_BAIXAPAG = 2448
'Erro na tentativa de modificar a tabela de Baixas a Pagar.
Public Const ERRO_LEITURA_BAIXAPARCPAG = 2449
'Erro na leitura da tabela de Baixas de Parcelas a Pagar.
Public Const ERRO_UNLOCK_BAIXAPARCPAG = 2450
'Erro na tentativa de desfazer o lock na tabela de Baixas de Parcelas a Pagar.
Public Const ERRO_LOCK_BAIXAPARCPAG = 2451
'N�o conseguiu fazer o lock da Baixa de Parcela a Pagar.
Public Const ERRO_MODIFICACAO_BAIXAPARCPAG = 2452
'Erro na tentativa de modificar a tabela de Baixas de Parcelas a Pagar.
Public Const ERRO_SALDO_PARCELA_MAIOR_QUE_VALOR = 2453
'O saldo da parcela n�o pode ser maior que o seu valor.
Public Const ERRO_TITULO_PAGAR_INEXISTENTE = 2454
'O T�tulo a Pagar n�o est� cadastrado.
Public Const ERRO_PARCELA_PAGAR_INEXISTENTE = 2455
'A Parcela a Pagar n�o est� cadastrada.
Public Const ERRO_EXCLUSAO_TITULOS_PAGAR_BAIXADOS = 2456
'Erro na exclus�o do T�tulo a Pagar Baixado.
Public Const ERRO_EXCLUSAO_PARCELAS_PAGAR_BAIXADAS = 2457
'Erro na exclus�o da Parcela a Pagar Baixada.
Public Const ERRO_TITULO_PAGAR_BAIXADO_INEXISTENTE = 2458
'O T�tulo a Pagar Baixado n�o est� cadastrado.
Public Const ERRO_LEITURA_PARCELAS_PAG_BAIXADA = 2459
'Erro na leitura da tabela de Parcelas a Pagar Baixadas.
Public Const ERRO_PARCELA_PAGAR_BAIXADA_INEXISTENTE = 2460
'A Parcela a Pagar Baixada n�o est� cadastrada.
Public Const ERRO_LOCK_PARCELAS_PAGAR_BAIXADA = 2461
'N�o conseguiu fazer o lock de Parcelas a Pagar Baixadas.
Public Const ERRO_LOCK_TITULOS_PAGAR_BAIXADOS = 2462
'N�o conseguiu fazer o lock de T�tulos a Pagar Baixados.
Public Const ERRO_UNLOCK_PARCELAS_PAGAR_BAIXADA = 2463
'Erro na tentativa de desfazer o lock na tabela de Parcelas a Pagar Baixadas.
Public Const ERRO_UNLOCK_TITULOS_PAGAR_BAIXADOS = 2464
'Erro na tentativa de desfazer o lock na tabela de T�tulos a Pagar Baixados.
Public Const ERRO_MODIFICACAO_TITULOS_PAGAR = 2465
'Erro na tentativa de modificar a tabela de T�tulos a Pagar.
Public Const ERRO_INCLUSAO_CARTEIRAS_COBRADOR = 2468
'Erro na tentativa de inserir a Carteira Cobrador.
Public Const ERRO_EXCLUSAO_CARTEIRAS_COBRADOR = 2469
'Erro na exclus�o da Carteira Cobrador.
Public Const ERRO_EXCLUSAO_COBRADOR = 2470 'Parametro codigo do Cobrador
'Erro na exclus�o do Cobrador.
Public Const ERRO_COBRADOR_INEXISTENTE = 2471 'Parametro codigo do Cobrador
'O Cobrador n�o est� cadastrado.
Public Const ERRO_LEITURA_CARTEIRAS_COBRANCA = 2472
'Erro na leitura da tabela de Carteiras Cobran�a.
Public Const ERRO_MODIFICACAO_CARTEIRAS_COBRANCA = 2473
'Erro na tentativa de modificar a tabela de Carteiras Cobran�a.
Public Const ERRO_MODIFICACAO_COBRADOR = 2474 'Parametro codigo do Cobrador
'Erro na inser��o na tabela de Cobrador , no codigo %i.
Public Const ERRO_INSERCAO_COBRADOR = 2475 'Parametro codigo do Cobrador
'Erro na leitura da tabela de Cobrador , no codigo %i.
Public Const ERRO_EXCLUSAO_CARTEIRAS_COBRANCA = 2476
'Erro na exclus�o da Carteira Cobran�a.
Public Const ERRO_INCLUSAO_CARTEIRAS_COBRANCA = 2477
'Erro na tentativa de inserir a Carteira Cobran�a.
Public Const ERRO_LEITURA_PARCELASREC_SEM_CHEQUEPRE = 2479 'Sem parametros
'Erro na leitura de parcelas � receber para vincula��o com cheque pr�-datado
Public Const ERRO_CHEQUEPRE_INEXISTENTE = 2481 'Parametro: lNumIntCheque
'O Cheque Pre %l n�o existe na tabela de ChequePre.
Public Const ERRO_CLIENTE_CHQPRE_NAO_PREENCHIDO = 2482 'Sem parametro
'O Cliente n�o foi informado.
Public Const ERRO_FILIALCLIENTE_NAO_ENCONTRADA = 2483 'Parametro: Filial.Text
'A Filial %s n�o foi encontrada na tabela de FiliaisClientes.
Public Const ERRO_BANCO_CHQPRE_NAO_PREENCHIDO = 2484
'O preenchimento de Banco � obrigat�rio.
Public Const ERRO_CONTA_CHQPRE_NAO_PREENCHIDA = 2485
'O preenchimento da Conta Corrente � obrigat�rio.
Public Const ERRO_NUMERO_CHQPRE_NAO_PREENCHIDO = 2486
'O preenchimento de N�mero � obrigat�rio.
Public Const ERRO_VALOR_CHQPRE_NAO_PREENCHIDO = 2487
'O preenchimento de Valor � obrigat�rio.
Public Const ERRO_DATADEPOSITO_CHQPRE_NAO_PREENCHIDA = 2488
'O preenchimento de Data de Dep�sito � obrigat�rio.
Public Const ERRO_LEITURA_PARCELASREC_TITULOSREC = 2489 'lNumIntCheque
'Erro na tentativa de ler registro na tabela ParcelasRec E TitulosRec com NumIntCheque = %l
Public Const ERRO_DATADEPOSITO_MAIOR_DATA_CONTABIL = 2493 'Parametro: dtDataDep, dtDataContab
'A Data Cont�bil � maior que a Data de Dep�sito.
Public Const ERRO_TIPO_PARCELA_NAO_INFORMADO = 2494 'Par�metro: iParcela
'O campo Tipo da Parcela %i n�o foi preenchido.
Public Const ERRO_NUMTITULO_PARCELA_NAO_INFORMADO = 2495 'Par�metro: iParcela
'O campo NumTitulo da Parcela %i n�o foi preenchido.
Public Const ERRO_PARCELA_NAO_INFORMADA = 2496 'Par�metro: iParcela
'O campo NumTitulo da Parcela %i n�o foi preenchido.
Public Const ERRO_VALOR_PARCELA_NAO_INFORMADA = 2497 'Par�metro: iParcela
'O campo Valor da Parcela %i n�o foi preenchido.
Public Const ERRO_CHEQUEPRE_NUMBORDERO_DEPOSITADO = 2499 'Parametro: iBanco, sAgencia, sContaCorrente, lNumero
'O ChequePre com Banco %i, Ag�ncia %s, ContaCorrente %s e N�mero %l j� foi depositado.
Public Const ERRO_AGENCIA_CHQPRE_INVALIDA = 2500 'Parametro: sAgencia
'A Ag�ncia %s � inv�lida.
Public Const ERRO_AGENCIA_CHQPRE_NAO_PREENCHIDA = 2501 'Parametro: sAgencia
'O preenchimento da Ag�ncia � obrigat�rio.
Public Const ERRO_CONTA_CHQPRE_INVALIDA = 2502 'Parametro: sConta
'A Conta %s � inv�lida.
Public Const ERRO_FILIAL_CHQPRE_NAO_PREENCHIDA = 2503 'Sem par�metros
'Preenchimento do filial � obrigat�rio.
Public Const ERRO_NUMERO_NAO_E_NUMERICO = 2506 'Par�metro Numero.Text
'O preenchimento de N�mero deve ser num�rico.
Public Const ERRO_TIPO_FORNECEDOR_NAO_ENCONTRADO = 2507 'Parametro sDescricao
'O Tipo de Fornecedor %s n�o foi encontrado.
Public Const ERRO_PERCENTUAL_COMISSAO_NAO_INFORMADO = 2510 'iComissao
'O Percentual da Comissao %i do T�tulo n�o foi informado.
Public Const ERRO_VALORBASE_COMISSAO_NAO_INFORMADO = 2511 'iComissao
'O Valor Base da Comissao %i do T�tulo n�o foi informado.
Public Const ERRO_ALTERACAO_ENDERECO = 2512
'Erro de modificacao na tabela de Enderecos no Banco de Dados.
Public Const ERRO_FILIALCLIENTE_INEXISTENTE = 2513 'Parametro iCodCliente e lCodFilial
'A Filial %i do Cliente com c�digo %l n�o existe.
Public Const ERRO_CODFILIAL_NAO_PREENCHIDO = 2514
'O c�digo da Filial n�o foi preenchido.
Public Const ERRO_NOMEFILIAL_NAO_PREENCHIDO = 2515
'O nome da Filial n�o foi preenchido.
Public Const ERRO_FILIALCLIENTE_EXCLUSAO_MATRIZ = 2516
'A exclus�o da matriz deve ser feita na tela de Clientes.
Public Const ERRO_LOCK_ENDERECOS = 2517
'Erro na tentativa de lockar a tabela de Enderecos.
Public Const ERRO_EXCLUSAO_ENDERECO = 2518 'Parametro lCodigoEndereco
'Erro na tentativa de excluir o Endereco de C�digo = %l .
Public Const ERRO_TIPO_NAO_ANTECIPREC = 2519 'Sem par�metros
'O Movimento n�o � do tipo Recebimento Antecipado.
Public Const ERRO_LEITURA_ANTECIPREC = 2520 'Par�metro lNumIntRec
'Erro na leitura do Recebimento antecipado %l.
Public Const ERRO_LEITURA_ANTECIPREC1 = 2521 'Par�metro lNumIntMovto
'Erro na leitura do Recebimento antecipado associado ao N�mero de movimento: %l na tabela RecebAntecipados.
Public Const ERRO_INSERCAO_ANTECIPREC = 2522 'Par�metro: lNumIntRec
'Ocorreu um erro na tentativa de inser��o do Recebimento antecipado %l na Tabela RecebAntecipados.
Public Const ERRO_EXCLUSAO_ANTECIPREC = 2523 'Par�metro: lNumIntRec
'Ocorreu um erro na tentativa de exclus�o do Recebimento antecipado %l.
Public Const ERRO_LOCK_ANTECIPREC = 2524 'Par�metro lNumIntRec
'Erro na tentativa de fazer o "lock" do Recebimento antecipado %l na tabela RecebAntecipados
Public Const ERRO_ANTECIPREC_EXCLUIDO = 2525 'Par�metros iCodConta + lSequencial
'O Recebimento antecipado com a Conta %i e Sequencial %l est� exclu�do.
Public Const ERRO_RECEBIMENTO_APROPRIADO = 2526 'Par�metro iCodigo
'N�o � poss�vel excluir o Recebimento antecipado %i, pois j� foi apropriado (total ou parcialmente).
Public Const ERRO_CLIENTE_NAO_COINCIDE = 2527 'Par�metros lCliente(da tela) e lCliente(da tabela)
'O Cliente %l n�o coincide com o Cliente %l cadastrado no Recebimento Antecipado.
Public Const ERRO_ANTECIPREC_INEXISTENTE = 2528  'Par�metros iCodConta + lSequencial
'O Movimento com a Conta %i e Sequencial %l n�o est� cadastrado.
Public Const ERRO_TIPOAPLICACAO_INATIVO = 2529  'Parametro: iCodigo
'O Tipo De Aplicacao com C�digo %i est� inativo.
Public Const ERRO_TIPOAPLICACAO_INEXISTENTE1 = 2530  'Parametro: iTipoAplicacao
'N�o existe o tipo de aplicacao %i na tabela de Tipos De Aplicacao.
Public Const ERRO_TIPOAPLICACAO_INEXISTENTE2 = 2531 'Paramento: TipoAplicacao.Text
'O Tipo de Aplicacao %s nao existe.
Public Const ERRO_LEITURA_TIPOSDEAPLICACAO1 = 2532  'Sem parametro
'Erro na leitura da tabela de Tipos De Aplicacoes.
Public Const ERRO_LOCK_TIPOSDEAPLICACAO = 2533 'Parametro: iTipoAplicacao
'Erro no "lock" do Tipo de Aplica��o i% da tabela TiposDeAplicacao.
Public Const ERRO_LOCK_APLICACOES = 2534 'Parametro: lCodigo
'Erro no "lock" da tabela Alpica��es. C�digo = %l.
Public Const ERRO_APLICACAO_INEXISTENTE = 2535 'Parametro: lCodigo
'A Aplica��o %l n�o existe na tabela de Aplica��es.
Public Const ERRO_APLICACAO_EXCLUIDA = 2536 'Parametro: lCodigo
'A Aplicacao com C�digo %l est� exclu�da.
Public Const ERRO_TIPO_NAO_APLICACAO = 2537 'Parametros: lSequencial
'O movimento %l n�o � do tipo Aplica��o.
Public Const ERRO_MOVIMENTO_EXCLUIDO = 2538 'Parametro: lNumMovto
'O Movimento %l est� exclu�do.
Public Const ERRO_TIPOAPLICACAO_NAO_PREENCHIDO = 2539 'Sem parametro
'O preenchimento do Tipo De Aplica��o � obrigat�rio.
Public Const ERRO_DATA_APLICACAO_NAO_PREENCHIDA = 2540 'Sem parametro
'O preenchimento da Data De Aplicacao � obrigat�rio.
Public Const ERRO_CONTACORRENTE_NAO_PREENCHIDA = 2541 'Sem parametro
'O preenchimento de Conta Corrente � obrigat�rio.
Public Const ERRO_VALRESGATE_MENOR_VALAPLICADO = 2542 'Parametro: dValResgPrev, dValAplic
'O Valor do Resgate Previsto %d � menor que o Valor Aplicado %d.
Public Const ERRO_DATARESGPREV_MENOR_DATAAPLIC = 2543 'Parametro: dtDataAplicacao, dtDataAplic
'A Data de Resgate Prevista %dt � menor que a Data de Aplica��o %dt.
Public Const ERRO_VALORAPLICADO_NAO_INFORMADO = 2544
'O valor aplicado n�o foi informado.
Public Const ERRO_DATAAPLICACAO_MENOR = 2545 'Parametro: dtDataAplicacao, dtDataSaldoInicial
'A Data da Aplicacao dt% � menor do que a Data Inicial da Conta dt%.
Public Const ERRO_ATUALIZACAO_APLICACOES = 2546 'Parametro: lCodigo
'Erro na Atualizacao da Aplica��o l% da tabela de Aplica��es.
Public Const ERRO_INSERCAO_APLICACOES = 2547 'Parametro: lCodigo
'Erro na Insercao da Aplica��o %l na tabela Aplica��es.
Public Const ERRO_LEITURA_RESGATES1 = 2548 'Parametro: lCodigoAplicacao
'Erro na leitura dos Resgates da Aplica��o %l.
Public Const ERRO_APLICACAO_RESGATE = 2549 'Parametro: lCodigoAplicacao
'Existe Resgate associado a Aplicacao com C�digo %l.
Public Const ERRO_LEITURA_FAVORECIDOS = 2550 'Parametro: iFavorecido
'Erro na leitura do Favorecido %i da tabela de Favorecidos.
Public Const ERRO_LOCK_FAVORECIDOS1 = 2551 'Parametro iFavorecido
'Erro na tentativa de fazer "lock" do Favorecido i% na tabela Favorecidos.
Public Const ERRO_LEITURA_MOVIMENTOSCONTACORRENTE3 = 2552 'Parametro: lNumMovto
'Erro na leitura do Movimento l% da tabela de Movimentos de Conta Corrente.
Public Const ERRO_LEITURA_MOVIMENTOSCONTACORRENTE2 = 2553 'Parametros: iCodConta, iTipoMeioPagto, lNumero
'Erro na leitura do Movimento da Conta i%, Tipo de Pagamento i% e N�mero l% da tabela de MovimentosContaCorrente.
Public Const ERRO_LOCK_MOVIMENTOSCONTACORRENTE2 = 2554 'Parametro: lNumMovto
'Erro na tentativa de fazer "lock"  no Movimento l% na tabela de MovimentosContasCorrente.
Public Const ERRO_MOVCONTACORRENTE_EXCLUIDO1 = 2555 'Parametro: lNumMovto
'O Movimento l% est� exclu�do.
Public Const ERRO_MOVIMENTO_NAO_CADASTRADO3 = 2556 'Parametro: lNumMovto
'O Movimento %l n�o est� cadastrado.
Public Const ERRO_TIPOMEIOPAGTO_INEXISTENTE1 = 2557 'Parametro: TipoMeioPagto.Text
'O Tipo de Pagmento %s nao est� cadastrado.
Public Const ERRO_ATUALIZACAO_MOVIMENTOSCONTACORRENTE2 = 2558 'Parametro: lNumMovto
'Erro de Atualizacao da Tabela de MovimentosContaCorrente com n�mero do movimento %l.
Public Const ERRO_INSERCAO_MOVIMENTOSCONTACORRENTE2 = 2559 'Parametro: lNumMovto
'Erro na tentativa de inclus�o de Movimento %l.
Public Const ERRO_RESGATE_INEXISTENTE = 2560 'Parametro: lCodigoAplicacao
'O Resgate com C�digo %l n�o existe na tabela de Resgates.
Public Const ERRO_RESGATE_INEXISTENTE1 = 2561 'Parametro: lCodigoAplicacao, iSeqResgate
'O Resgate com C�digo %l e Sequencial %i n�o existe na tabela de Resgates.
Public Const ERRO_RESGATE_EXCLUIDO = 2562 'Parametro: iSeqResgate, lCodigoAplicacao
'O Resgate %i da Aplica��o %l est� exclu�do.
Public Const ERRO_TIPO_NAO_RESGATE = 2563 'Parametro: lSequencial
'O movimento %l n�o � do tipo Resgate.
Public Const ERRO_LEITURA_RESGATES = 2564 'Parametro: lCodigoAplicacao
'Erro de Leitura na Tabela de Resgates com o C�digo %l.
Public Const ERRO_CODIGO_APLICACA0_NAO_PREENCHIDO = 2565 'Sem parametros
'Preenchimento do c�digo da aplica��o � obrigat�rio.
Public Const ERRO_CODIGO_RESGATE_NAO_PREENCHIDO = 2566 'Sem parametros
'Preenchimento do c�digo do resgate � obrigat�rio.
Public Const ERRO_DATA_RESGATE_NAO_PREENCHIDA = 2567 'Sem parametro
'O preenchimento da Data De Resgate � obrigat�rio.
Public Const ERRO_VALOR_CREDITADO_NAO_INFORMADO = 2568 'Sem parametro
'O valor creditado n�o foi informado.
Public Const ERRO_SALDO_ATUAL_NEGATIVO = 2569  'Parametro: SaldoAtual.Caption
'O Saldo Atual est� negativo.
Public Const ERRO_DATARESGPREV_MENOR_DATARESG = 2570 'Parametro: dtDataResgPrev, dtDataResg
'A Data de Resgate Prevista %dt � menor que a Data do Resgate %dt.
Public Const ERRO_VALRESGATE_MENOR_SALATUAL = 2571 'Parametro: dValResgPrev, dSalAtual
'O Valor do Resgate Previsto %d � menor que o Saldo Atual %d.
Public Const ERRO_VALORES_DIFERENTES = 2572 'Parametro: dValCred, dValCredLab
'Os campos de Valor Creditado %d e Valor Creditado Label %d est�o com valores diferentes.
Public Const ERRO_DATA_RESGATE_MENOR = 2573 'Parametro: dtDataMovimento, dtDataSaldoInicial
'A Data de Resgate dt% � menor do que a Data Inicial da Conta dt%.
Public Const ERRO_LOCK_RESGATES = 2574 'Parametro: lCodigoAplicacao, iSeqResgate
'Erro no "lock" da tabela Resgates. C�digo = %l e Sequencial = %i.
Public Const ERRO_ATUALIZACAO_RESGATES = 2575 'Parametro: lCodigoAplicacao, iSequencial
'Erro na Atualizacao do Resgate com C�digo l% e Sequencial %i da tabela de Resgates.
Public Const ERRO_INSERCAO_RESGATES = 2576 'Parametro: lCodigoAplicacao, iSeqResgate
'Erro na Insercao do Resgatec com C�digo %l e Sequencial %i na tabela de Resgates.
Public Const ERRO_LOCK_FLUXOTIPOAPLIC = 2577 'Parametro FluxoId, TipoAplicacao
'Ocorreu um erro ao tentar fazer lock no Fluxo Tipo Aplic %l, no tipo de aplica��o %i.
Public Const ERRO_LEITURA_CONTASCORRENTESINTERNAS_CODIGO = 2578 'Parametro: iCodigo
'Erro na leitura da tabela de ContasCorrentesInternas. Codigo = %s.
Public Const ERRO_ATUALIZACAO_FLUXOSALDOSINICIAIS = 2579 'Parametro Nome do Fluxo de Caixa
'Erro na atualiza��o da tabela de Saldos Iniciais do Fluxo de Caixa. Fluxo = %s.
Public Const ERRO_LOCK_FLUXOSINTETICO = 2580  'Parametro FluxoId
'Ocorreu um erro ao tentar fazer lock no Fluxo Sintetico. FluxoId = %l.
Public Const ERRO_INSERCAO_FLUXOCREDANTECIP = 2581 'Sem Parametro
'Erro na inser��o de um registro na tabela FluxoCredAntecip.
Public Const ERRO_EXCLUSAO_FLUXOCREDANTECIP = 2582 'Parametro Nome do Fluxo de Caixa
'Erro na exclus�o do Fluxo de Caixa (FluxoCredAntecip) %s.
Public Const ERRO_ATUALIZACAO_FLUXOANALITICO = 2583 'Parametro Nome do Fluxo de Caixa
'Erro na tentativa de atualiza��o da tabela de Fluxo Analitico, do Fluxo de Caixa %s.
Public Const ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO = 2584
'O Campo Tipo n�o foi preenchido.
Public Const ERRO_DATA_RESGATE_MENOR1 = 2585 'Parametro: dtDataMovimento, dtDataAplicacao
'A Data de Resgate dt% � menor do que a Data de Aplicacao.
Public Const ERRO_CAMPOS_CREDITO_PAGAR_NAO_PREENCHIDOS = 2586 'Sem par�metros
'Os campos Fornecedor, Filial e Tipo devem estar preenchidos.
Public Const ERRO_LEITURA_CREDITOSPAGFORN1 = 2588 'Par�metro: lNumIntDoc
'Erro na leitura da tabela de CreditosPagForn com n�mero interno do documento %l.
Public Const ERRO_LEITURA_BAIXASPAG1 = 2590 'Par�metro: lNumIntDoc
'Erro na leitura da tabela BaixasPag com n�mero interno do documento %l.
Public Const ERRO_TIPODOC_NAO_E_TIPOCREDITOPAG = 2591 'Par�metro: sSiglaDocumento
'O Tipo de Documento %s n�o � do tipo Devolu��es / Cr�dito.
Public Const ERRO_CREDITOPAGAR_NAO_CADASTRADO = 2594 'Par�metros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Devolu��es / Cr�dito do Fornecedor %l, Filial %i, Tipo de Documento %s, N�mero %l e Data de Emiss�o %dt n�o existe na tabela CreditosPagForn.
Public Const ERRO_CREDITOPAGAR_NAO_CADASTRADO1 = 2595 'Parametro: lNumIntDoc
'Devolu��es / Cr�dito com n�mero interno do documento %l n�o existe na tabela de CreditosPagForn.
Public Const ERRO_LOCK_CREDITOSPAGFORN = 2597 'Par�metros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'N�o conseguiu fazer o "lock", na tabela de CreditosPafForn com Fornecedor %l, Filial %i, Tipo de Documento %s, N�mero %l e Data de Emiss�o dt%.
Public Const ERRO_ALTERACAO_CREDPAG_LANCADO = 2598 'Par�metros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'N�o � poss�vel alterar Devolu��o / Cr�dito, com dados Fornecedor %l, Filial %i, Tipo de Documento %s, N�mero %l e Data de Emiss�o %dt , porque est� lan�ado.
Public Const ERRO_ALTERACAO_CREDPAG_BAIXADO = 2599 'Par�metros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'N�o � poss�vel alterar Devolu��o / Cr�dito com dados Fornecedor %l, Filial %i, Tipo de Documento %s, N�mero %l e Data de Emiss�o %dt porque est� baixado.
Public Const ERRO_EXCLUSAO_CREDPAG_BAIXADO = 2600 'Par�metros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'N�o � possivel excluir Devolu��o / Cr�dito porque est� baixado. Dados da Devolu��o/Cr�dito: Fornecedor %l, Filial %i, Tipo de Documento %s, N�mero %l, Data de Emiss�o dt%.
Public Const ERRO_EXCLUSAO_CREDPAG_VINCULADO_BAIXA = 2601 'Par�metros: lNumIntBaixa, lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmiss�o
'N�o � possivel excluir Devolu��o / Cr�dito porque est� vinculado a Baixa Pagar com n�mero interno %l. Dados do Cr�dito: Fornecedor %l, Filial %i, Tipo de Documento %s, N�mero %l, Data de Emiss�o dt%.
Public Const ERRO_MODIFICAO_CRED_PAG_OUTRA_FILIALEMPRESA = 2602 'Par�metros: iFilialEmpresa, lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'N�o � poss�vel modificar Devolu��es / Cr�dito da Filial Empresa %i. Dados do Documento: c�digo do Fornecedor %l, c�digo da Filial %i, Tipo %s, N�mero %l e Data de Emissao dt%.
Public Const ERRO_EXCLUSAO_CREDITOSPAGFORN = 2603 ''Par�metros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Erro na exclus�o de Devolu��es / Cr�dito com dados Fornecedor %l, Filial %i, Tipo de Documento %s, N�mero %l e Data de Emiss�o %dt, da tabela CreditosPagForn.
Public Const ERRO_LEITURA_NFISCALBAIXADAS2 = 2604 'Par�metro: lCodFornecedor, iCodFilial
'Erro  na leitura da tabela NFiscalBaixadas com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_NFSPAG3 = 2605 'Par�metro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela NFsPag com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_NFSPAGPEND2 = 2606 'Par�metro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela NFsPagPend com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_NFSPAGBAIXADAS2 = 2607 'Par�metro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela NfsPagBaixadas com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_TITULOSPAG2 = 2608 'Par�metro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela TitulosPag com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_TITULOSPAGPEND2 = 2609 'Par�metro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela TitulosPagPend com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_TITULOSPAGBAIXADOS2 = 2610 'Par�metro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela TitulosPagBaixados com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_CREDITOSPAGFORN3 = 2611 'Par�metro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela CreditosPagForn com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_PAGTOSANTECIPADOS2 = 2612 'Par�metro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela PagtosAntecipados com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_PENDENCIASBORDEROPAGTO2 = 2613 'Par�metros: lCodFornecedor, iCodFilial
'Erro na leitura da tabela de PendenciasBorderoPagto com Fornecedor %l e Filial %i.
Public Const ERRO_FILIAL_FORNECEDOR_REL_NFISCAL = 2620 'Par�metros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l est� relacionada com Nota Fiscal com c�digo interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_NFISCALBAIXADA = 2621 'Par�metros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l est� relacionada com Nota Fiscal Baixada com c�digo interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_NFISCALPAG = 2622 'Par�metros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l est� relacionado com Nota Fiscal � Pagar com c�digo interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_NFISCALPAGPEND = 2623 'Par�metros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l est� relacionado com Nota Fiscal � Pagar Pendente com c�digo interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_NFISCALPAGBAIXADA = 2624 'Par�metros: iCodFilial, lCodFornecedor, lCodigo
'A Filial do Fornecedor %l est� relacionado com Nota Fiscal � Pagar Baixada com c�digo interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_TIT_PAGAR = 2625 'Par�metros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l est� relacionado com T�tulo � Pagar com c�digo interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_TIT_PAGAR_PEND = 2626 'Par�metros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l est� relacionado com T�tulo � Pagar Pendente com c�digo interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_TIT_PAGAR_BAIXADO = 2627 'Par�metros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %ii do Fornecedor %l est� relacionado com T�tulo � Pagar Baixado com c�digo interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_CREDITO_PAGAR_FORN = 2628 'Par�metros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l est� relacionado com Cr�dito � Pagar Fornecedor com c�digo interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_PAGTO_ANTECIPADO = 2629 'Par�metros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l est� relacionado com Pagamento Antecipado com c�digo interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_PEND_BORDERO_PAGTO = 2630 'Par�metros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l est� relacionado com Pend�ncia Border� Pagto com c�digo interno %l.
Public Const ERRO_CODIGO_TIPO_FORNECEDOR_NAO_PREENCHIDO = 2631 'Sem parametros
'Preenchimento do Tipo de Fornecedor � obrigat�rio.
Public Const ERRO_LEITURA_NFISCALBAIXADAS1 = 2632 'Par�metro: lCodigo
'Erro  na leitura da tabela NFiscalBaixadas com Fornecedor %l.
Public Const ERRO_FORNECEDOR_REL_NFISCAL = 2633 'Par�metros: lCodFornecedor, lNFiscal
'O Fornecedor %l est� relacionado com Nota Fiscal n�mero %l.
Public Const ERRO_FORNECEDOR_REL_NFISCALBAIXADA = 2634 'Par�metros: lCodFornecedor, lCodNFiscal
'O Fornecedor %l est� relacionado com Nota Fiscal Baixada com c�digo interno %l.
Public Const ERRO_FORNECEDOR_REL_NFISCALPAG = 2635 'Par�metros: lCodFornecedor, lNFiscal
'O Fornecedor %l est� relacionado com Nota Fiscal � Pagar com n�mero %l.
Public Const ERRO_FORNECEDOR_REL_NFISCALPAGPEND = 2636 'Par�metros: lCodFornecedor, lCodNFiscal
'O Fornecedor %l est� relacionado com Nota Fiscal � Pagar Pendente com c�digo interno %l.
Public Const ERRO_FORNECEDOR_REL_NFISCALPAGBAIXADA = 2637 'Par�metros: lCodFornecedor, lNFiscal
'O Fornecedor %l est� relacionado com Nota Fiscal � Pagar Baixada com n�mero %l.
Public Const ERRO_FORNECEDOR_REL_TIT_PAGAR = 2638 'Par�metros: lCodFornecedor, lCodTitPagar
'O Fornecedor %l est� relacionado com T�tulo � Pagar com c�digo interno %l.
Public Const ERRO_FORNECEDOR_REL_TIT_PAGAR_PEND = 2639 'Par�metros: lCodFornecedor, lCodTitPagar
'O Fornecedor %l est� relacionado com T�tulo � Pagar Pendente com c�digo interno %l.
Public Const ERRO_FORNECEDOR_REL_TIT_PAGAR_BAIXADO = 2640 'Par�metros: lCodFornecedor, lCodTitPagar
'O Fornecedor %l est� relacionado com T�tulo � Pagar Baixado com c�digo interno %l.
Public Const ERRO_FORNECEDOR_REL_CREDITO_PAGAR_FORN = 2641 'Par�metros: lCodigo, lCodigo
'O Fornecedor %l est� relacionado com Cr�dito � Pagar Fornecedor com c�digo interno %l.
Public Const ERRO_FORNECEDOR_REL_PAGTO_ANTECIPADO = 2642 'Par�metros: lCodFornecedor, lCodPagtoAntec
'O Fornecedor %l est� relacionado com Pagamento Antecipado com c�digo interno %l.
Public Const ERRO_FORNECEDOR_REL_FORNECEDOR_PRODUTO = 2643 'Par�metros: lCodFornecedor, sProduto
'O Fornecedor %l est� relacionado com o Produto %s.
Public Const ERRO_FORNECEDOR_REL_PRODUTO_FILIAL = 2644 'Par�metros: lCodigo, sProduto
'O Fornecedor %l est� relacionado com o Produto %s da Filial.
Public Const ERRO_FORNECEDOR_REL_PEND_BORDERO_PAGTO = 2645 'Par�metros: lCodFornecedor, lCodPendBordPagto
'O Fornecedor %l est� relacionado com Pend�ncia Border� Pagto com c�digo interno %l.
Public Const ERRO_LEITURA_NFSPAG2 = 2646 'Par�metro: lCodigo
'Erro na leitura da tabela NFsPag com Fornecedor %l.
Public Const ERRO_LEITURA_NFSPAGPEND1 = 2647 'Par�metro: lCodigo
'Erro na leitura da tabela NFsPagPend com Fornecedor %l.
Public Const ERRO_LEITURA_NFSPAGBAIXADAS1 = 2648 'Par�metro: lCodigo
'Erro na leitura da tabela NfsPagBaixadas com Fornecedor %l.
Public Const ERRO_LEITURA_TITULOSPAG1 = 2649 'Par�metro: lCodigo
'Erro na leitura da tabela TitulosPag com Fornecedor %l.
Public Const ERRO_LEITURA_TITULOSPAGPEND1 = 2650 'Par�metro: lCodigo
'Erro na leitura da tabela TitulosPagPend com Fornecedor %l.
Public Const ERRO_LEITURA_TITULOSPAGBAIXADOS1 = 2651 'Par�metro: lCodigo
'Erro na leitura da tabela TitulosPagBaixados com Fornecedor %l.
Public Const ERRO_LEITURA_CREDITOSPAGFORN2 = 2652 'Par�metro: lCodigo
'Erro na leitura da tabela CreditosPagForn com Fornecedor %l.
Public Const ERRO_LEITURA_PAGTOSANTECIPADOS1 = 2653 'Par�metro: lCodigo
'Erro na leitura da tabela PagtosAntecipados com Fornecedor %l.
Public Const ERRO_LEITURA_FORNECEDORPRODUTO2 = 2654 'Par�metros: lCodigo
'Erro na leitura da tabela de FornecedorProduto com Fornecedor %l.
Public Const ERRO_LEITURA_PENDENCIASBORDEROPAGTO1 = 2656 'Par�metros: lCodigo
'Erro na leitura da tabela de PendenciasBorderoPagto com Fornecedor %l.
Public Const ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA3 = 2663 'Par�metro: iCondicaoPagto
'A Condi��o de Pagamento %i n�o est� cadastrada no Banco de Dados.
Public Const ERRO_NFPAGAR_VINCULADA_NFISCAL = 2664 'Par�metro: lNumNotaFiscal
'Nota Fiscal � Pagar com n�mero %l est� vinculada a Nota Fiscal.
Public Const ERRO_TIPODEFORNECEDOR_NAO_CADASTRADO = 2665 'Parametro iCodigo
'Tipo de Fornecedor %i n�o est� cadastrado.
Public Const ERRO_INSERCAO_TIPOSDEFORNECEDOR = 2666 'Parametro iCodigo
'Erro na inser��o do Tipo de Fornecedor %i.
Public Const ERRO_ATUALIZACAO_TIPOSDEFORNECEDOR = 2667 'Parametro iCodigo
'Erro na atualiza��o do Tipo de Fornecedor %i.
Public Const ERRO_EXCLUSAO_TIPOSDEFORNECEDOR = 2668 'Parametro iCodigo
'Erro na exclus�o do Tipo de Fornecedor %i.
Public Const ERRO_LOCK_TIPOSDEFORNECEDOR = 2669 'Parametro iCodigo
'N�o conseguiu fazer o lock do Tipo de Fornecedor %i.
Public Const ERRO_TIPODEFORNECEDOR_RELACIONADO_COM_FORNECEDOR = 2670 'Sem parametros
'N�o � poss�vel excluir Tipo de Fornecedor relacionado com Fornecedor.
Public Const ERRO_DESCRICAO_TIPO_FORNECEDOR_REPETIDA = 2671 'Par�metro: iCodigo
'Tipo de Fornecedor com c�digo %i est� cadastrado e tem esta descri��o.
Public Const ERRO_CAMPOS_APLICACAO_NAO_ALTERAVEIS = 2672
'Os campos n�o ser�o alterados. Para alterar a aplica��o exclua esta e crie uma outra.
Public Const ERRO_DEBITORECCLI_NAO_ENCONTRADO = 2673 'Parametro: lNumIntDoc
'O D�bito a Receber de Cliente com n�mero %l n�o est� cadastrado no Banco de Dados.
Public Const ERRO_NUM_MAXIMO_COMISSOES_ULTRAPASSADO = 2674 'Parametros: iNumParcelasBD, iNumMaxParcelas
'O n�mero de parcelas lidas no BD � %i que supera o n�mero m�ximo permitido igual a %i.
Public Const ERRO_LEITURA_COMISSOES_VENDEDORES = 2675 'lNumIntDoc
'Erro na tentativa de ler registro na tabela Comiss�es e Vendedores com NumIntDoc = %l.
Public Const ERRO_LEITURA_DEBITOSRECCLI1 = 2677 'Parametros: lCliente, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Erro na leitura da tabela DebitosRecCli com Cliente %l, Filial %i, Tipo %s, N�mero do T�tulo %l e Data de Emiss�o %dt.
Public Const ERRO_NAO_E_POSSIVEL_MODIFICAR_DEB_REC_CLI_OUTRA_FILIALEMPRESA = 2679 'Sem parametros
'N�o � poss�vel modificar D�bito a Receber Cliente de outra Filial da Empresa.
Public Const ERRO_DEBITO_REC_CLI_BAIXADO = 2680 'Parametro: lCliente, sVendedorNomeRed
'N�o � poss�vel alterar as comiss�es do D�bito a Receber de Cliente %l. A comiss�o do Vendedor %s est� paga.
Public Const ERRO_DEBITORECCLI_NAO_ENCONTRADO1 = 2681 'Parametros: lCliente, iFilial, sSiglaDocumento, lNumTitulo
'Devolu��o / Cr�dito do Cliente %l, Filial %i, Tipo de Documento %s e N�mero do T�tulo %l n�o est� Cadastrado.
Public Const ERRO_LEITURA_DEBITOSRECCLI2 = 2682 'Parametros: lCliente, iFilial, sSiglaDocumento, lNumTitulo
'Erro na leitura da tabela DebitosRecCli com Cliente %l, Filial %i, Tipo de Documento %s e N�mero do T�tulo %l.
Public Const ERRO_LOCK_DEBITOSRECCLI = 2683  'Parametros: lCliente, iFilial, sSiglaDocumento, lNumTitulo
'N�o conseguiu fazer o lock na tabela de DebitosRecCli com Cliente %l, Filial %i Tipo de Documento %s N�mero do T�tulo %l.
Public Const ERRO_NAO_E_PERMITIDO_EXCLUSAO_DEBRECCLI_VINCULADO_BAIXA = 2684 'Parametro: lNumIntDoc
'N�o � permitida a exclus�o de D�bito a Receber Cliente n�mero %l vinculado a baixa.
Public Const ERRO_EXCLUSAO_DEBITOSRECCLI = 2685 'Parametro: lNumIntDoc
'Erro na exclus�o de D�bito a Receber Cliente n�mero %l da tabela de DebitosRecCli.
Public Const ERRO_CAMPOS_DEBITO_RECEBER_NAO_PREENCHIDOS = 2686 'Sem parametros
'Os campos Cliente, Filial e Tipo devem estar preenchidos.
Public Const ERRO_PORTADOR_NAO_ENCONTRADO = 2690 'Parametro: sPortador
'O Portador %s n�o foi encontrado.
Public Const ERRO_TIPOCOBRANCA_NAO_PREENCHIDO = 2691
'O Tipo de Cobran�a n�o foi informado
Public Const ERRO_TITULO_NAO_CADASTRADO2 = 2692 'Par�metro: sTipo, sFornecedor, sFilial, sNumTitulo, sDataEmissao
'N�o foi encontrado nenhum T�tulo do Tipo %s, com Fornecedor = %s, Filial = %s, N�mero = %s e Data de Emissao %s.
Public Const ERRO_PARCELA_PAG_NAO_ABERTA2 = 2693 'Par�metros: lNumT�tulo, iParcela
'A Parcela i% do T�tulo N�mero %l n�o pode ser modificada porque n�o est� aberta.
Public Const ERRO_LEITURA_FATURASBAIXADAS = 2694 'Par�metro: lNumnotafiscal
'Erro na leitura da Fatura n�mero %l na tabela NfsPagBaixadas.
Public Const ERRO_FATURA_PENDENTE_MODIFICACAO = 2695 'Par�metro: lNumTitulo
'N�o � poss�vel modificar Fatura n�mero %l .Ela faz parte de um Lote Pendente.
Public Const ERRO_FATURA_BAIXADA_MODIFICACAO = 2696 'Par�metro: lNumTitulo
'N�o � poss�vel modificar Fatura n�mero %l porque ela est� baixada.
Public Const ERRO_LEITURA_FATURASPAGBAIXADAS = 2697 'Par�metro: lNumTitulo
'Erro na tentativa de leitura da Fatura %l na tabela TitulosPagBaixados.
Public Const ERRO_TITULO_NAO_FATURAPAGAR = 2698 'Par�metro: lNumTitulo
'O T�tulo com n�mero %l n�o � uma Fatura a Pagar.
Public Const ERRO_FATURAPAG_NAO_CADASTRADA = 2699 'Par�metro: lNumIntDoc
'A Fatura com N�mero Interno = %l n�o est� cadastrada.
Public Const ERRO_FATURAPAG_NAO_CADASTRADA1 = 2700 'Par�metro: lNumTitulo
'A Fatura N�mero %l est� n�o cadastrada ou em Lote Pendente ou Baixada.
Public Const ERRO_SOMA_NFS_SELECIONADAS_INVALIDA = 2701 'Par�metros: sValorTotalNfs, sValorTotal
'A soma dos valores a pagar das Notas Fiscais selecionadas %s n�o � igual ao Valor Total da Fatura %s.
Public Const ERRO_FATURAPAG_PENDENTE_EXCLUSAO = 2702 'lNumTitulo
'N�o � poss�vel excluir a Fatura n�mero %l porque ela faz parte de um Lote Pendente.
Public Const ERRO_FATURAPAG_BAIXADA_EXCLUSAO = 2703 'lNumTitulo
'N�o � poss�vel excluir a Fatura n�mero %l porque ela est� baixada.
Public Const ERRO_FATURAPAG_NAO_CADASTRADA2 = 2704 'Parametro: lNumTitulo
'A Fatura %l n�o est� cadastrada no Banco de Dados.
Public Const ERRO_NF_JA_VINCULADA = 2705 'Par�metro: lNumNotaFiscal
'A Nota Fiscal n�mero %l j� est� vinculada a outra Fatura.
Public Const ERRO_LEITURA_FATURA = 2706 'Par�metro: lNumIntDoc
'Erro na tentativa de ler Fatura na tabela TitulosPag com N�mero Interno %l.
Public Const ERRO_LEITURA_FATURA1 = 2707 'Par�metro: lNumTitulo
'Erro na tentativa de ler Fatura n�mero %l na tabela TitulosPag.
Public Const ERRO_LEITURA_FATURAPEND = 2708 'lNumTitulo
'Erro na leitura da Fatura n�mero %l na Tabela de Titulos Pendentes.
Public Const ERRO_LEITURA_FATURABAIXADA = 2709 'lNumTitulo
'Erro na leitura da Fatura n�mero %l na Tabela de TiTulos Baixados.
Public Const ERRO_LEITURA_TITULOSPAG3 = 2710 'Par�metros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao, sSiglaDocumento
'Erro na leitura da tabela TitulosPag com Fornecedor %l, Filial %i, T�tulo, Data de Emiss�o %dt e Tipo de Documento %s.
Public Const ERRO_INSERCAO_TITULOSPAG = 2711 'Par�metros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao, sSiglaDocumento
'Erro na tentativa de inserir um registro na tabela TitulosPag com Fornecedor %l, Filial %i, T�tulo, Data de Emiss�o %dt e Tipo de Documento %s.
Public Const ERRO_NFFATPAGAR_VINCULADA_NFISCAL = 2712 'Par�metro: lNumNotaFiscal
'Nota Fiscal Fatura � Pagar com n�mero %l est� vinculada � Nota Fiscal.
Public Const ERRO_TIPO_INSTRUCAO_COBRANCA_NAO_ENCONTRADO = 2713 'Parametro iCodigo
'Instru��o de Cobran�a %i n�o foi encontrada.
Public Const ERRO_TIPO_INSTRUCAO_COBRANCA_NAO_CADASTRADA = 2714 'Parametro iCodigo
'Instru��o de Cobran�a %i n�o est� cadastrada.
Public Const ERRO_DIAS_DE_PROTESTO1_NAO_PREENCHIDO = 2715 'Sem Parametros
'Dias para Devolu��o / Protesto de Instru��o Prim�ria deve ser preenchido.
Public Const ERRO_DIAS_DE_PROTESTO2_NAO_PREENCHIDO = 2716 'Sem Parametros
'Dias para Devolu��o / Protesto de Instru��o Secund�ria deve ser preenchido.
Public Const ERRO_LEITURA_PADROES_COBRANCA = 2717 'Sem Parametros
'Erro na leitura da tabela PadroesCobranca
Public Const ERRO_INSTRUCAO_PRIMARIA_NAO_PREENCHIDA = 2718 'Sem Parametros
'O preenchimento da Instru��o Prim�ria � obrigat�rio
Public Const ERRO_LEITURA_TIPOSMOVIMENTO = 2719 'Sem parametro
'Erro na Leitura da Tabela  "TiposMovimento".
Public Const ERRO_CANCELAMENTO_PAG_NAO_SE_APLICA_AO_MOV = 2720 'Sem parametro
'Cancelamento de Pagamento n�o se aplica a este tipo de Movimento.
Public Const ERRO_TIPO_MEIO_PAGAMENTO_NAO_CADASTRADO = 2721 'Sem parametro
'Tipo de Meio De Pagamento n�o est� cadastrado.
Public Const NAO_EXISTE_PAG_PARA_SER_CANCELADO = 2722 'Sem parametro
'N�o existe Pagamento com estes dados para ser cancelado
Public Const ERRO_MENSAGEM_NAO_INFORMADA = 2723 'Sem par�metros
'A Mensagem n�o foi informada.
Public Const ERRO_CODIGO_NAO_INFORMADO = 2724 'Sem par�metros
'O C�digo da Mensagem n�o foi informado
Public Const ERRO_MENSAGEM_COM_CARACTER_INICIAL_ERRADO = 2725 'Sem parametros
'Descri��o de Mensagem n�o pode come�ar com este caracter.
Public Const ERRO_EXCLUSAO_MENSAGEM = 2726 'Parametro C�digo da Mensagem
'Houve um erro na exclus�o da Mensagem %i do banco de dados.
Public Const ERRO_ATUALIZACAO_MENSAGEM = 2727 'Par�metro C�digo da Mensagem
'Erro na atualiza��o da Mensagem %i.
Public Const ERRO_INSERCAO_MENSAGEM = 2728 'Par�metro C�digo da Mensagem
'Erro na Inser��o da Mensagem %i.
Public Const ERRO_LEITURA_MENSAGEM1 = 2729 'Sem Par�metros
'Erro na leitura da Tabela Mensagens
Public Const ERRO_TIPODOC_NAO_ENCONTRADO = 2730 'Parametro: Tipo.Text
'O Tipo De Documento %s n�o est� cadastrado.
Public Const ERRO_LEITURA_CODIGO_PAIS = 2731 'Par�metro C�digo do Pa�s
'Erro na leitura do Pa�s %s.
Public Const ERRO_CODIGO_PAIS_NAO_CADASTRADO = 2732 'Par�metro C�digo do Pa�s
'Pa�s %s n�o est� cadastrado.
Public Const ERRO_LEITURA_OC_COB = 2734 'parametro = codigo do cobrador
'Erro na leitura de ocorr�ncias para o cobrador com c�digo %d
Public Const ERRO_ATUALIZACAO_OC_COB = 2735 'parametro = codigo do cobrador
'Erro na atualiza��o de ocorr�ncias para o cobrador com c�digo %d
Public Const ERRO_ATUALIZACAO_ANTECIPPAG = 2736 'parametros = lNumIntMovto
'Erro na atualiza��o do Pagamento antecipado cujo N�mero de movimento �: %l.
Public Const ERRO_ATUALIZACAO_ANTECIPRECEB = 2737 'parametros = lNumIntMovto
'Erro na atualiza��o do Recebimento antecipado cujo N�mero de movimento �: %l.
Public Const ERRO_SALDO_NEGATIVO = 2738   'Parametros: dValor, dSaldoNaoApropriado
'O valor %d n�o � permitido pois deixar� o saldo negativo %d
Public Const ERRO_COBRADOR_PROPRIA_EMPRESA = 2739 'Par�metro: objCobrador.iCodigo
'O Cobrador %i � da pr�pria empresa.
Public Const ERRO_LEITURA_COBRADOR2 = 2740 'Sem Parametros
'Erro na leitura da tabela de Cobrador.
Public Const ERRO_FILIAL_NAO_ENCONTRADA2 = 2741 'Parametro: iFilial
'A Filial com codigo %i n�o foi encontrada.
Public Const ERRO_LEITURA_OUTROSPAGBAIXADOS = 2742 'Par�metro: sSigla, lNumTitulo
'Erro na tentativa de leitura de Outro Pagamento do Tipo %s N�mero %l na tabela TitulosPagBaixados.
Public Const ERRO_LEITURA_OUTROSPAG1 = 2743 'Par�metro: sSigla, lNumTitulo
'Erro na tentativa de leitura de Outro Pagamento do Tipo %s e N�mero %l na tabela TitulosPag.
Public Const ERRO_LEITURA_OUTROSPAGPEND = 2744 'Par�metro: sSigla, lNumTitulo
'Erro na tentativa de leitura da Outro Pagamento do Tipo %s e N�mero %l na tabela TitulosPagPend.
Public Const ERRO_TITULOPAGAR_NAO_CADASTRADO = 2745 'lNumIntDoc
'O T�tulo a Pagar com N�mero Interno %l n�o est� cadastrado.
Public Const ERRO_TIPO_DOCUMENTO_NAO_OUTROSPAG = 2746 'Parametros: lNumTitulo, sSigla
'A Sigla %s do T�tulo n�mero %l n�o � utilizada em OutrosPag.
Public Const ERRO_TITULO_PENDENTE_MODIFICACAO = 2747 'Par�metro: sSiglaDocumento, lNumTitulo
'N�o � poss�vel modificar T�tulo a pagar do Tipo %s com n�mero %l porque ele faz parte de um Lote Pendente.
Public Const ERRO_TITULO_BAIXADO_MODIFICACAO = 2748 'Par�metro: sSiglaDocumento, lNumTitulo
'N�o � poss�vel modificar T�tulo a pagar do Tipo %s com n�mero %l porque ele est� baixado.
Public Const ERRO_NUMERO_PARCELAS_TITULO_ALTERADO = 2749 'Par�metros: sSigla, lNumTitulo, iNumParcelasTela, iNumParcelasBD
'N�o � poss�vel alterar o n�mero de parcelas de T�tulo a Pagar do Tipo %s com n�mero %l que est� lan�ado. N�mero de Parcelas da Tela: %i. N�mero de Parcelas do BD: %i.
Public Const ERRO_INSERCAO_TITULO_PAGAR = 2750 'Parametros : lNumTitulo, sSigla
'Erro na Tentativa de inserir Titulo n�mero %l do Tipo %s na tabela TitulosPag.
Public Const ERRO_TITULO_PENDENTE_EXCLUSAO = 2751 'Par�metros: 'lNumTitulo
'N�o � poss�vel excluir o T�tulo a Pagar n�mero %l porque ele faz parte de um Lote Pendente.
Public Const ERRO_TITULO_BAIXADO_EXCLUSAO = 2752 'Par�metro: lNumTitulo
'N�o � poss�vel excluir o Titulo a Pagar n�mero %l porque ele est� Baixado
Public Const ERRO_TITULOPAGAR_NAO_CADASTRADO1 = 2753  'Par�metro: lNumTitulo
'O Titulo a Pagar n�mero %l n�o est� cadastrado.
Public Const ERRO_EXCLUSAO_COMISSAO_BAIXADA = 2754 'sem parametros
'N�o pode excluir documento associado a uma comiss�o marcada como "baixada"
Public Const ERRO_LEITURA_BAIXASREC1 = 2755 'Par�metro: lNumIntDoc
'Erro na leitura da tabela BaixasRec com n�mero interno do documento %l.
Public Const ERRO_DEBITOREC_VINCULADO_NFISCAL = 2756 'Par�metro: lNumIntDoc
'D�bito Com Cliente com n�mero interno %l est� vinculado � Nota Fiscal.
Public Const ERRO_LEITURA_PORTADOR1 = 2757 'sem par�metro
'Erro na leitura da tabela Portador.
Public Const ERRO_BANCO_PORTADOR = 2758 'par�metros iBanco
'O Banco %i est� sendo usado na tabela Portador.
Public Const ERRO_NOME_CONTACORRENTEINTERNA_EXISTENTE = 2759 'Par�metros: sNomeReduzido
'J� existe uma Conta Corrente Interna com Nome %s para a Empresa.
Public Const ERRO_MENSAGEM_ASSOCIADA_CLIENTE = 2761 'Par�metro: iCodigo
'N�o � permitida a exclus�o da Mensagem %i porque est� associada com Cliente.
Public Const ERRO_COBRADOR_NAO_CADASTRADO1 = 2764 'Par�metro: sNomeReduzido
'O Cobrador %s n�o est� cadastrado no Banco de Dados.
Public Const ERRO_INSTRUCAO_NAO_SELECIONADA = 2765 'Sem par�metros
'N�o foi poss�vel selecionar a Instru��o.
Public Const ERRO_DIAS_PROTESTO_NAO_PREENCHIDO = 2766 'Sem par�metros
'� obrigat�rio o preenchimento do campo Dias de Protesto para o tipo de Instru��o escolhida.
Public Const ERRO_LEITURA_TIPOSINSTRCOBRANCA1 = 2767 'Par�metro: iCodigo
'Erro na leitura da tabela TiposInstrCobranca com C�digo %i.
Public Const ERRO_LEITURA_TIPOSDEOCORREMCOBR = 2768 'Sem par�metros
'Erro na leitura da tabela TiposDeOcorRemCobr.
Public Const ERRO_LEITURA_TIPOSDEOCORREMCOBR1 = 2769 'Par�metro: iCodOcorrencia
'Erro na leitura da tabela TiposDeOcorRemCobr com C�digo da Ocorr�ncia %i.
Public Const ERRO_NUMERO_NAO_PREENCHIDO = 2770 'Sem par�metros
'Preenchimento do N�mero � obrigat�rio.
Public Const ERRO_INSTRUCAO_NAO_CADASTRADA = 2771 'Par�metro: iCodigo
'A Instru��o com c�digo %i n�o est� cadastrada no Banco de dados.
Public Const ERRO_INSTRUCAO_NAO_CADASTRADA1 = 2772 'Par�metro: Instrucao.Text
'A Instru��o com descri��o %s n�o est� cadastrada no Banco de dados.
Public Const ERRO_OCORRENCIA_NAO_CADASTRADA = 2773 'Par�metro: iCodigo
'A Ocorr�ncia com c�digo %i n�o est� cadastrada no Banco de dados.
Public Const ERRO_OCORRENCIA_NAO_CADASTRADA1 = 2774 'Par�metro: Ocorrencia.Text
'A Ocorr�ncia com descri��o %s n�o est� cadastrada no Banco de dados.
Public Const ERRO_OCORR_REM_COBR_NAO_CADASTRADA = 2775 'Par�metros: iNumSeqOcorr
'A Ocorr�ncia com Sequencial %i desta Parcela n�o est� cadastrada no Banco de dados.
Public Const ERRO_TITULORECEBER_NAO_CADASTRADO2 = 2776 'Par�metros: iFilialEmpresa, lCliente, iFilial, sSiglaDocumento, lNumTitulo
'O Titulo Receber da Filial Empresa %i, Cliente %l , Filial %i, Sigla do Documento %s e N�mero do T�tulo %l n�o est� cadastrado no Banco de Dados.
Public Const ERRO_PARCELAREC_NUMINT_NAO_CADASTRADA = 2777 'Par�metros: lNumIntTitulo, iNumParcela
'A Parcela Receber com n�mero interno do T�tulo %l e n�mero da Parcela %i n�o est� cadastrada no Banco de Dados.
Public Const ERRO_PARCELAREC_NUMINT_BAIXADA = 2778 'Par�metros: lNumIntTitulo, iNumParcela
'A Parcela Receber com n�mero interno do T�tulo %l e n�mero da Parcela %i est� Baixada.
Public Const ERRO_INSTRUCAO_ENVIADA_BANCO = 2779 'Par�metros: lNumIntTitulo, iNumParcela
'A Parcela Receber com n�mero interno do T�tulo %l e n�mero da Parcela %i j� foi enviada ao Banco.
Public Const ERRO_LOCK_OCORRENCIASREMPARCREC = 2780 'Par�metros: lNumIntParc, iNumSeqOcorr
'Erro na tentativa de fazer "lock" na tabela OcorrenciasRemParcRec com n�mero interno da Parcela %l e Sequencial %i.
Public Const ERRO_PARCELAREC_COBR_ELETRONICA = 2781 'Sem par�metros
'A Parcela n�o � do tipo Cobran�a Eletr�nica.
Public Const ERRO_ATUALIZACAO_OCORRENCIASREMPARCREC = 2782 'Par�metros: lNumIntParc, iNumSeqOcorr
' Erro na atualiza��o da Ocorr�ncia com n�mero interno da Parcela %l e Sequencial %i.
Public Const ERRO_INCLUSAO_OCORRENCIASREMPARCREC = 2783 'Par�metro: lNumIntParc, iNumSeqOcorr
'Erro na tentativa de inserir um registro na tabela OcorrenciasRemParcRec com n�mero interno da Parcela %l e Sequencial %i.
Public Const ERRO_PARCELAREC_OCORRENCIA_BORDERO_COBRANCA = 2784 'Par�metros: lNumIntParc, iNumSeqOcorr
'A Ocorr�ncia com n�mero interno da Parcela %l e Sequencial %i faz parte de um Bordero de Cobran�a.
Public Const ERRO_ATUALIZACAO_PARCELAREC_OCORRENCIA_BORDERO_COBRANCA = 2785 'Par�metros: lNumIntParc, iNumSeqOcorr
'N�o � permitido a atualiza��o da Ocorr�ncia com n�mero interno da Parcela %l e Sequencial %i porque faz parte de um Border� de Cobran�a.
Public Const ERRO_ATUALIZACAO_OCORRENCIA_INSTRUCAO = 2786 'Par�metros: lNumParc, iNumSeqOcorr
'N�o � permitido a atualiza��o dos campos Ocorr�ncia e Instru��o1 desta Parcela com Sequencial %i.
Public Const ERRO_PARCELAREC_NAO_CADASTRADA = 2787 'Par�metro: lNumIntParc
'A Parcela com n�mero interno %l n�o est� cadastrada no Banco de Dados.
Public Const ERRO_EXCLUSAO_OCORRENCIASREMPARCREC = 2788 'Par�metros: lNumIntParc, iNumSeqOcorr
'Erro na tentativa de excluir um registro da tabela OcorrenciasRemParcRec com n�mero interno da Parcela %l e Sequencial %i.
Public Const ERRO_OCORRENCIA_COBRANCA_NAO_PREENCHIDA = 2789 'Sem par�metros
'A Ocorr�ncia deve ser preenchida.
Public Const ERRO_INSTRUCAO_COBRANCA_NAO_PREENCHIDA = 2790 'Sem par�metros
'A Instru��o deve ser preenchida.
Public Const ERRO_LEITURA_TIPOSINSTRCOBRANCA2 = 2791 'Par�metro: sDescricao
'Erro na leitura da tabela TiposInstrCobranca com Descri��o %s.
Public Const ERRO_PARCELA_NAO_INFORMADA1 = 2792 'Sem par�metros
'Os campos Cliente, Filial, Tipo, N�mero, Parcela e Sequencial devem ser preenchidos.
Public Const ERRO_INSTRUCAO_COBRANCA_NAO_CADASTRADA = 2793 'Sem par�metros
'A Parcela informada n�o est� cadastrada no Banco de Dados.
Public Const ERRO_ATUALIZACAO_OCORRENCIA = 2794 'Par�metros: lNumParc, iNumSeqOcorr
'Erro na tentativa de atualiza um registro da tabela OcorrenciasRemParcRec desta Parcela com Sequencial %i.
Public Const ERRO_ANTECIPPAG_NAO_CARREGADO = 2795
'O Adiantamento deve ser carregado atrav�s de sua escolha em uma tela de lista. Aperte o bot�o "Adiantamentos a Filial do Fornecedor..." e escolha o Adiantamento a excluir.
Public Const ERRO_ANTECIPRECEB_NAO_CARREGADO = 2796
'O Adiantamento deve ser carregado atrav�s de sua escolha em uma tela de lista. Aperte o bot�o "Adiantamentos � Filial do Cliente..." e escolha o Adiantamento a excluir.
Public Const ERRO_VALOR_CR_MENOR_VALOR_BAIXAR = 2797 'dValorCredito, dValorTotalBaixar
'O Valor do Saldo do Cr�dito, %d,  � menor que o Valor Total a Baixar, %d.
Public Const ERRO_VALOR_PA_MENOR_VALOR_BAIXAR = 2798 'dValorPA, dValorTotalBaixar
'O Valor do Saldo do Pagamento Antecipado, %d,  � menor que o Valor Total a Baixar, %d.
Public Const ERRO_PARCELAS_SUPERIOR_NUM_MAX_PARCELAS_BAIXA = 2799 'Sem par�metros
'O n�mero de parcelas ultrapassou o limite permitido.
Public Const ERRO_CREDITOS_SUPERIOR_NUM_MAX_CREDITOS = 2800 'Sem par�metros
'O n�mero de cr�ditos ultrapassou o limite permitido.
Public Const ERRO_PAGTOSANTECIPADOS_SUPERIOR_NUM_MAX_PAGTOS_ANTECIPADOS = 2801 'Sem par�metros
'O n�mero de pagamentos antecipados ultrapassou o limite permitido.
Public Const ERRO_PORTADOR_NAO_CADASTRADO2 = 2802 'Par�metro: sPortador
'Portador %s n�o est� cadastrado.
Public Const ERRO_LOCK_PAGTOSANTECIPADOS = 2803 'Parametro: lNumMovto
'Erro na tentativa de "lock" da tabela PagtosAntecipados. NumMovto=%l.
Public Const ERRO_ATRIBUTOS_PAGTOANTECIPADO_MUDARAM = 2804 'Parametros: dSaldoNaoApropriadoBD, dSaldoNaoApropriado, lFornecedorBD, lFornecedor, iFilialBD, iFilial
'Atributos de Pagto Antecipado mudaram. Saldo: BD %d Tela %d, Fornecedor: BD %l Tela %l, Filial: BD %i Tela %i.
Public Const ERRO_ATUALIZACAO_PAGTOANTECIPADO_SALDO = 2805 'Parametro: lNumMovto
'Erro na atualiza��o do SaldoNaoApropriado do Pagamento Antecipado com NumMovto=%l.
Public Const ERRO_CREDITOPAGFORN_NAO_CADASTRADO = 2806 'Parametros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Cr�dito Pagar n�o cadastrado. Fornecedor=%l, Filial=%i, SiglaDocumento=%s, NumTitulo=%l, DataEmissao=%dt.
Public Const ERRO_LOCK_CREDITOPAGFORN = 2807 'Parametro: lNumIntDoc
'Erro na tentativa de "lock" da tabela CreditosPagForn. N�mero Interno = %l.
Public Const ERRO_SALDO_CREDITOPAGFORN_MUDOU = 2808 'Parametros: dSaldoBD, dSaldo
'Saldo de Cr�dito Pagar mudou. Saldo no BD: %d. Saldo na Tela: %d.
Public Const ERRO_ATUALIZACAO_CREDITOPAGFORN_SALDO = 2809 'Parametro: lNumIntDoc
'Erro na atualiza��o na tabela CreditosPagForn. N�mero Interno = %l.
Public Const ERRO_TITULOINIC_MAIOR_TITULOFIM = 2810 'Sem par�metro
'T�tulo inicial n�o pode ser maior que o T�tulo final.
Public Const ERRO_DATAEMISSAO_INICIAL_MAIOR = 2811 'Sem par�metro
'Data Emissao Inicial n�o pode ser maior que a Data Emissao Final.
Public Const ERRO_DATAVENCIMENTO_INICIAL_MAIOR = 2812 'Sem par�metro
'Data Vencimento Inicial n�o pode ser maior que a Data Vencimento Final.
Public Const ERRO_VALORBAIXAR_MAIOR_SALDO = 2813 'Sem par�metro
'Valor a Baixar n�o pode ser maior que o Saldo.
Public Const ERRO_VALORBAIXAR_MENOR_IGUAL_DESCONTO = 2814 'Sem par�metro
'Valor a Baixar n�o pode ser menor ou igual ao Desconto.
Public Const ERRO_MULTA_INCOMPATIVEL_DATABAIXA_DATAVENC = 2815 'Sem par�metro
'Multa incompat�vel com Data de Baixa menor ou igual � Data de Vencimento.
Public Const ERRO_LEITURA_CONTASCORRENTESINTERNAS2 = 2816 'Par�metro: sNomeReduzido
'Erro na tentativa de leitura da Conta Corrente %s na tabela ContasCorrentesInternas.
Public Const ERRO_DATA_BAIXA_SEM_PREENCHIMENTO = 2817
'A Data Baixa deve estar preenchida.
Public Const ERRO_VALORBAIXAR_PARCELAS_NAO_INFORMADO = 2818
'O Valor a Baixar das parecelas selecionadas deve estar preenchido.
Public Const ERRO_CONTA_PAGDINHEIRO_NAO_INFORMADA = 2819 'Sem par�metro
'Para Pagamento em Dinheiro a Conta Corrente deve ser informada.
Public Const ERRO_VALOR_PAGDINHEIRO_NAO_INFORMADO = 2820 'Sem par�metro
'Para Pagamento em Dinheiro o Valor do Pagamento deve ser informado.
Public Const ERRO_VALORPAG_DIFERENTE_TOTALPARCELAS = 2821 'Par�metros : sValorPago, dTotalParcelas
'O Valor do Pagamento %s n�o coincide com o Valor a Pagar %d das Parcelas selecionadas.
Public Const ERRO_VALORBAIXAR_PA_NAO_INFORMADO = 2822 'Sem par�metro
'O Valor a Baixar do Pagamento Antecipado deve estar preenchido.
Public Const ERRO_PARCELA_PA_NAO_MARCADA = 2823 'Sem par�metro
'Pelo menos uma Parcela do Pagamento Antecipado deve estar marcada.
Public Const ERRO_VALORUSAR_PA_DIFERENTE_TOTALPARCELAS = 2824 'Par�metros : sValBaixarPA, dTotalParcelas
'O Valor a Usar do Pagamento Antecipado %s n�o coincide com o Valor a Pagar %d das Parcelas selecionadas.
Public Const ERRO_VALORBAIXAR_CR_NAO_INFORMADO = 2825 'Sem par�metro
'O Valor a Baixar do Cr�dito deve estar preenchido.
Public Const ERRO_PARCELA_CR_NAO_MARCADA = 2826 'Sem par�metro
'Pelo menos uma Parcela dos Cr�ditos deve estar marcada.
Public Const ERRO_VALORUSAR_CR_DIFERENTE_TOTALPARCELAS = 2827 'Par�metros : sValBaixarCR, dTotalParcelas
'O Valor a Usar dos Cr�ditos %s n�o coincide com o Valor a Pagar %d das Parcelas selecionadas.
Public Const ERRO_LEITURA_TIPOMEIOPAGTO2 = 2828 'Par�metro sDescricao
'Erro na leitura do Tipo de Pagamento %s da tabela TipoMeioPagto.
Public Const ERRO_TITULO_NAO_MARCADO = 2829 'Sem par�metro
'Pelo menos um T�tulo deve estar marcado para esta opera��o.
Public Const ERRO_TITULOS_NAO_MARCADOS = 2830 'Sem par�metro
'Pelo menos dois T�tulos devem estar marcado para esta opera��o.
Public Const ERRO_DEBITOS_SUPERIOR_NUM_MAX_DEBITOS = 2831 'Sem par�metros
'O n�mero de cr�ditos ultrapassou o limite permitido.
Public Const ERRO_RECEBANTECIPADOS_SUPERIOR_NUM_MAX_RECEB_ANTECIPADOS = 2832 'Sem par�metros
'O n�mero de pagamentos antecipados ultrapassou o limite permitido.
Public Const ERRO_ATRIBUTOS_RECEBANTECIPADO_MUDARAM = 2833 'Parametros: dSaldoNaoApropriadoBD, dSaldoNaoApropriado, lClienteBD, lCliente, iFilialBD, iFilial
'Atributos de Recebimento Antecipado mudaram. Saldo: BD %d Tela %d, Cliente: BD %l Tela %l, Filial: BD %i Tela %i.
Public Const ERRO_ATUALIZACAO_RECEBANTECIPADO_SALDO = 2834 'Parametro: lNumMovto
'Erro na atualiza��o do SaldoNaoApropriado do Recebimento Antecipado com NumMovto=%l.
Public Const ERRO_DEBITORECCLI_NAO_CADASTRADO = 2835 'Parametros: lCliente, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'D�bito Receber n�o cadastrado. Cliente=%l, Filial=%i, SiglaDocumento=%s, NumTitulo=%l, DataEmissao=%dt.
Public Const ERRO_SALDO_DEBITORECCLI_MUDOU = 2836 'Parametros: dSaldoBD, dSaldo
'Saldo de D�bito Receber mudou. Saldo no BD: %d. Saldo na Tela: %d.
Public Const ERRO_ATUALIZACAO_DEBITOSRECCLI_SALDO = 2837 'Parametro: lNumIntDoc
'Erro na atualiza��o na tabela DebitosRecCli. N�mero Interno = %l.
Public Const ERRO_DATA_CREDITO_SEM_PREENCHIMENTO = 2838
'A Data Cr�dito deve estar preenchida.
Public Const ERRO_CONTA_RECDINHEIRO_NAO_INFORMADA = 2839 'Sem par�metro
'Para Recebimento em Dinheiro a Conta Corrente deve ser informada.
Public Const ERRO_VALOR_RECDINHEIRO_NAO_INFORMADO = 2840 'Sem par�metro
'Para Recebimento em Dinheiro o Valor deve ser informado.
Public Const ERRO_VALORREC_DIFERENTE_TOTALPARCELAS = 2841 'Par�metros : sValReceber, dTotalParcelas
'O Valor do Recebimento %s n�o coincide com o Valor a Receber %d das Parcelas selecionadas.
Public Const ERRO_VALORBAIXAR_RA_NAO_INFORMADO = 2842 'Sem par�metro
'O Valor a Baixar do Recebimento Antecipado deve estar preenchido.
Public Const ERRO_PARCELA_RA_NAO_MARCADA = 2843 'Sem par�metro
'Pelo menos uma Parcela do Recebimento Antecipado deve estar marcada.
Public Const ERRO_VALORUSAR_RA_DIFERENTE_TOTALPARCELAS = 2844 'Par�metros : sValorBaixarRA, dTotalParcelas
'O Valor a Usar do Recebimento Antecipado %s n�o coincide com o Valor a Receber %d das Parcelas selecionadas.
Public Const ERRO_VALORBAIXAR_DB_NAO_INFORMADO = 2845 'Sem par�metro
'O Valor a Baixar do D�bito deve estar preenchido.
Public Const ERRO_PARCELA_DB_NAO_MARCADA = 2846 'Sem par�metro
'Pelo menos uma Parcela dos D�bitos deve estar marcada.
Public Const ERRO_VALORUSAR_DB_DIFERENTE_TOTALPARCELAS = 2847 'Par�metros : sValorBaixarDB, dTotalParcelas
'O Valor a Usar dos D�bitos %s n�o coincide com o Valor a Receber %d das Parcelas selecionadas.
Public Const ERRO_RECEBANTECIPADO_NAO_CADASTRADO = 2848 'Parametro: lNumeroMovimento
'Recebimento Antecipado com NumMovto=%l n�o est� cadastrado.
Public Const ERRO_CAMPOS_BAIXAPARCREC_NAO_PREENCHIDOS = 2849 'Sem par�metros
'Para trazer uma Parcela � obrigat�rio que estejam preenchidos os campos Cliente, Filial, Tipo, N�mero, Data Emiss�o , Parcela e Sequencial da tela.
Public Const ERRO_BAIXA_PARCELA_CANCELADA = 2850 'Par�metros: iNumParcela, iSequencialBaixa
'A Baixa da Parcela %i e Sequencial %i j� est� cancelada.
Public Const ERRO_EXCLUSAO_PARCELAS_RECEBER_BAIXADAS = 2851
'Erra na tentativa de excluir registro da tabela de parcelas a receber baixadas.
Public Const ERRO_PARCELA_RECEBER_BAIXADA_INEXISTENTE = 2852
'Parcela a receber baixada n�o encontrada.
Public Const ERRO_UNLOCK_TITULOS_REC_BAIXADOS = 2853
'Erro na tentativa de desfazer o lock na tabela de Titulos a Receber Baixados.
Public Const ERRO_RECEBIMENTO_ANTECIPADO_INEXISTENTE = 2854
'Recebimento antecipado n�o cadastrado.
Public Const ERRO_MODIFICACAO_DEBITO_REC_CLI = 2855
'Erro na tentativa de modificar a tabela de D�bitos a Receber.
Public Const ERRO_DEBITO_REC_CLI_EXCLUIDO = 2856
'Esse D�bito est� exclu�do.
Public Const ERRO_MODIFICACAO_RECEBIMENTO_ANTECIPADO = 2857
'Erro na tentativa de modificar a tabela de Recebimentos Antecipados.
Public Const ERRO_RECEBIMENTO_ANTECIPADO_EXCLUIDO = 2858
'O recebimento antecipado j� havia sido exclu�do.
Public Const ERRO_MODIFICACAO_BAIXAPARCREC = 2859
'Erro na tentativa de modificar a tabela de Baixas de Parcelas a Receber.
Public Const ERRO_EXCLUSAO_TITULOS_RECEBER_BAIXADOS = 2860
'Erro na exclus�o do T�tulo a Receber Baixado.
Public Const ERRO_LOCK_BAIXAPARCREC = 2861
'N�o conseguiu fazer o lock da Baixa de Parcela a Receber.
Public Const ERRO_TITULO_RECEBER_BAIXADO_INEXISTENTE = 2863
'O T�tulo a Receber Baixado n�o est� cadastrado.
Public Const ERRO_LEITURA_BAIXAPARCREC = 2864
'Erro na leitura da tabela de Baixas de Parcelas a Receber
Public Const ERRO_LEITURA_BAIXAREC = 2865
'Erro na leitura da tabela de Baixas a Receber.
Public Const ERRO_UNLOCK_PARCELAS_REC_BAIXADAS = 2866
'Erro na tentativa de desfazer o lock na tabela de Parcelas a Receber Baixadas.
Public Const ERRO_LOCK_TITULOS_REC_BAIXADOS = 2867
'N�o conseguiu fazer o lock de T�tulos a Receber Baixados.
Public Const ERRO_UNLOCK_BAIXAPARCREC = 2868
'Erro na tentativa de desfazer o lock na tabela de Baixas de Parcelas a receber.
Public Const ERRO_BAIXAREC_INEXISTENTE = 2869
'A Baixa a Receber n�o est� cadastrada.
Public Const ERRO_EXCLUSAO_BAIXAREC = 2870
'Erro na exclus�o da Baixa a Receber.
Public Const ERRO_BAIXAPARCREC_INEXISTENTE = 2871
'A Baixa n�o est� cadastrada.
Public Const ERRO_LOCK_PARCELAS_REC_BAIXADAS = 2872
'N�o conseguiu fazer o lock de Parcelas a Receber Baixadas.
Public Const ERRO_BAIXAREC_EXCLUIDA = 2873
'A baixa a receber j� havia sido cancelada anteriormente.
Public Const ERRO_BAIXAPARCREC_EXCLUIDA = 2874
'A baixa a receber da parcela j� havia sido cancelada anteriormente.
Public Const ERRO_PARCELA_JA_EXISTENTE = 2875 'Parametro : iLinha
'A parcela informada � a mesma da %i linha do Grid.
Public Const ERRO_DATA_CREDITO_NAO_PREENCHIDA = 2876
'A Data de Cr�dito deve ser preenchida.
Public Const ERRO_TOTAL_RECEBIDO_NAO_PREENCHIDO = 2877
'O Total Recebido deve ser informado.
Public Const ERRO_VALOR_DESCONTO_PARCELA_SUPERIOR_SOMA_VALOR = 2878
'O Valor do Desconto da parcela na linha %i � superior ao valor recebido.
Public Const ERRO_VALOR_RECEBIDO_PARCELA_NAO_PREENCHIDO = 2879
'O valor recebido na linha %i n�o foi informado.
Public Const ERRO_NUMPARCELA_NAO_INFORMADO = 2880
'O numero da parcela n�o foi informado na linha %i.
Public Const ERRO_QUANTIDADE_INFORMADA_DIFERENTE_GRID = 2881
'A quantidade de Parcelas no Grid � diferente da quantidade informada na tela.
Public Const ERRO_TOTALRECEBIDO_PARCELAS_DIFERENTE = 2882
'A soma dos valores recebidos no Grid � diferente do valor recebido informado a tela.
Public Const ERRO_DATADEPOSITO_MENOR_DATAEMISSAO = 2883
'A Data de deposito deve ser maior ou igual � Data de Emiss�o.
Public Const ERRO_ATUALIZACAO_OCORRCOBR = 2884 'sem parametro
'Erro na tentativa de alterar o registro de uma ocorr�ncia de cobran�a
Public Const ERRO_ALTERACAO_BORDERO_COBRANCA = 2885 'Par�metro: lNumBordero
'Erro na tentativa de alterar o Border� de N�mero %l na Tabela de Border�s de Cobran�a.
Public Const ERRO_LOCK_BORDERO_COBRANCA = 2886 'Par�metro: lNumBordero
'Erro na tentativa de fazer "lock" na tabela BorderosCobranca para N�mero %l .
Public Const ERRO_BORDERO_COBRANCA_NAO_CADASTRADO = 2887 'Par�metro: lNumBordero
'O Bordero de Cobran�a de n�mero %l n�o est� cadastrado no Banco de Dados.
Public Const ERRO_BORDERO_COBRANCA_EXCLUIDO = 2888 'Par�metro: lNumBordero
'Border� de N�mero %l est� exclu�do.
Public Const ERRO_BORDERO_COBRANCA_SEM_OCORRENCIAS = 2889 'Par�metro: lNumBordero
'N�o foram enconradas ocorr�ncias para o Bordero %l no Banco de Dados
Public Const ERRO_PARCELA_RECEBER_BAIXADA = 2891 'Par�metro: lNumIntParc
'A Parcela a Receber com o N�mero Interno %l j� est� baixada.
Public Const ERRO_OCORRENCIA_DIFERENTE_PARCELA = 2892
'O Sequencial da ocorr�ncia n�o corresponde ao indicado na parcela.
Public Const ERRO_SEM_PARCELA_SELECIONADA = 2893
'Nenhuma parcela foi selecionada.
Public Const ERRO_LEITURA_PAGTO_BORDERO = 2894
'Erro na leitura de pagamento efetuado por border�
Public Const ERRO_TITULO_INFORMADO_SEM_FILIAL = 2895 'Sem par�metro
'O T�tulo s� pode ser preenchido se a Filial tiver sido informada.
Public Const ERRO_PARCELA_INFORMADA_SEM_TITULO = 2896 'Sem par�metro
'A Parcela s� pode ser preenchida se o T�tulo tiver sido informado.
Public Const ERRO_VALORPARCELA_MENOR_DESCONTO = 2897 'Sem par�metro
'O desconto ultrapassou o valor total da parcela, com juros e multa.
Public Const ERRO_PORTADORES_DIFERENTES = 2898 'Sem par�metros
'Um mesmo cheque n�o pode pagar T�tulos com Portadores diferentes.
Public Const ERRO_TITULOS_FORN_DIF_NAO_COBRBANC = 2899 'Sem par�metro
'S� � permitido um cheque para fornecedores diferentes quando for cobran�a banc�ria com portador definido.
Public Const ERRO_FILIALCLIENTE_INEXISTENTE1 = 2900 'Par�metros: iCodCliente, sCliente
'A Filial %i do Cliente %s n�o est� cadastrada no Banco de Dados.
Public Const ERRO_CLIENTE_ULTRAPASSOU_NUMMAXPARCELAS = 2901 'Sem par�metros
'O Cliente ultrapassou o n�mero m�ximo de Parcelas.
Public Const ERRO_PARCELA_GRID_NAO_SELECIONADA = 2902
'N�o h� parcela selecionada no grid de parcelas.
Public Const ERRO_SUBIR_PRIM_CHEQUE = 2903 'Sem par�metro
'N�o pode subir o primeiro cheque
Public Const ERRO_DESCER_CHEQUE_ZERO = 2904 'Sem par�metro
'N�o pode descer parcela que n�o esteja marcada para imprimir
Public Const ERRO_TITULOS_DE_UM_CHEQUE = 2905 'Sem par�metro
'Selecione parcelas de um cheque
Public Const ERRO_NAO_PODE_AGRUPAR_SEQ_ZERO = 2906 'Sem par�metro
'N�o pode selecionar t�tulo n�o marcado para emitir para esta opera��o.
Public Const ERRO_TITULOS_TIPO_COBR_DIFERENTE = 2907 'Sem par�metro
'Para esta opera��o os t�tulos devem ter o mesmo tipo de cobran�a.
Public Const ERRO_TITULOS_PORTADOR_DIFERENTE = 2908 'Sem par�metro
'Para esta opera��o os t�tulos devem ter o mesmo portador.
Public Const ERRO_TITULOS_CHEQUE_DIFERENTE = 2909 'Sem par�metro
'Para esta opera��o os t�tulos devem pertencer ao mesmo cheque.
Public Const ERRO_NUMCHEQUE_INVALIDO = 2910 'Par�metros: sCheque e iNumCheque
'O n�mero %s de Cheque � inv�lido. Deve ser maior ou igual a 1 e menor que %i
Public Const ERRO_NUMCHEQUE_NAO_PREENCHIDO = 2911 'Sem par�metro
'O n�mero do Cheque deve ser informado.
Public Const ERRO_NUM_CHEQUES_MARCADOS = 2912 'Sem par�metro
'Deve haver 1 Cheque marcado para esta opera��o.
Public Const ERRO_ESTADO_NAO_PREENCHIDO = 2913 'Sem par�metros
'O preenchimento do Estado correspondente ao Endere�o � obrigat�rio.
Public Const ERRO_COBRADOR_NAO_PERTENCE_FILIAL = 2914 'Parametro: iCobrador, iFilial
'O Cobrador %i n�o pertence � Filial selecionada. Filial: %i.
Public Const ERRO_VENDEDOR_VAZIO = 2915 'Sem parametros
'Nome reduzido do Vendedor selecionado na listbox est� vazio
Public Const ERRO_DATA_FINAL_MAIOR = 2917 'Sem parametros
'Data final n�o pode ser maior que Data de baixa
Public Const ERRO_EXTRATO_NAO_INFORMADO = 2918 'Sem par�metro
'O extrato deve ser informado.
Public Const ERRO_USADA_EM_NAO_PREENCHIDA = 2919 'Sem Parametros
'� obrigat�rio informar onde ser� utilizada esta Condic�o de Pagamento.
Public Const ERRO_VALORES_COMISSAO_BASE = 2920 'Par�metros:Valor Base e Valor da Comiss�o
'Valor Base %d � menor que o Valor da Comiss�o %d
Public Const ERRO_TECLAR_BOTAO_TRAZER = 2921 'Sem Parametros
'Depois de preencher os campos obrigatorios, aperte o Bot�o Trazer.
Public Const ERRO_BANCO_NAO_PREENCHIDO = 2922 'Sem par�metro
'O nome do banco deve ser informado.
Public Const ERRO_ARQUIVO_NAO_PREENCHIDO = 2923 'Sem par�metro
'O nome do arquivo deve ser especificado.
Public Const ERRO_NUM_MAX_NFS_SELEC_EXCEDIDO = 2924 'parametro: %d limite de nfs p/fatura
'N�o pode associar mais de %d notas fiscais a uma fatura
Public Const ERRO_BANCO_NAO_POSITIVO = 2925 'parametro sBanco
'O numero do Banco %s n�o � positivo.
Public Const ERRO_MOVIMENTO_DINHEIRO_SEM_SEQUENCIAL = 2926
'O movimento com Tipo de Pagamento "Dinheiro" deve ser carregado na tela atrav�s da tela de browse.
Public Const ERRO_DATAS_DESCONTOS_DESORDENADAS = 2927 'iParcela
'As datas de descontos para a parcela %i n�o est�o ordenadas.
Public Const ERRO_PARCELAREC_COBR_MANUAL = 2928 'Sem par�metros
'A Parcela n�o � do tipo Cobran�a Manual.
Public Const ERRO_CONTACORRENTE_FILIAL_DIFERENTE = 2929
'Esta Conta Corrente pertence a outra filial.
Public Const ERRO_VALOR_DEPOSITO_NAO_PREENCHIDO = 2930 'Sem Par�metros
'O valor do Dep�sito deve estar preenchido
Public Const ERRO_VALOR_RESGATE_NAO_PREENCHIDO = 2931 'Sem Parametros
'O Valor do Resgate n�o foi preenchido.
Public Const ERRO_SALDO_ZERO = 2932 'Sem Parametros
'N�o pode realizar resgate com o saldo igual a zero.
Public Const ERRO_VALOR_SAQUE_NAO_PREENCHIDO = 2933 'Sem Par�metros
'O valor do Saque deve estar preenchido.
Public Const ERRO_HISTMOVCTA_NAO_CADASTRADO1 = 2934 'Par�metro: Historico.Text
'O Hist�rico de Movimenta��o de Conta %s n�o est� cadastrado.
Public Const ERRO_DATA_BAIXA_INICIAL_MAIOR = 2935
'A data da Baixa Inicial n�o pode ser maior que a Final. 'Sem Parametros
Public Const ERRO_UMA_DATA_NAO_PREENCHIDA = 2936
'Voc� deve preencher pelo menos um dos dois pares de datas. Ou Baixa ou Digita��o.
Public Const ERRO_DATA_DIGITACAO_BAIXA_INICIAL_MAIOR = 2937
'A data da digitacao da baixa Inicial n�o pode ser maior que a Final. 'Sem Parametros
Public Const ERRO_FORNECEDOR_INICIAL_MAIOR = 2940
'O Fornecedor Inicial � maior que o final.'
Public Const ERRO_FORNECEDOR_NAO_CADASTRADO_2 = 2941
'O Fornecedor n�o est� cadastrado.'
Public Const ERRO_VENDEDOR_NAO_CADASTRADO_2 = 2942
'O Vendedor n�o est� cadastrado.'
Public Const ERRO_VALOR_INVALIDO1 = 2943 'Sem parametros
'N�mero de Devedores � invalido
Public Const ERRO_DATA_EMISSAO_INICIAL_MAIOR = 2944
'A data de Emissao Inicial n�o pode ser maior que a Final. 'Sem Parametros
Public Const ERRO_DATA_VENCTO_INICIAL_MAIOR = 2946
'A data do vencimento Inicial n�o pode ser maior que a Final. 'Sem Parametros
Public Const ERRO_CHEQUEDET_GRAVACAO = 2947
'Erro na grava��o de um cheque de pagamento de t�tulo
Public Const ERRO_DEBITO_REC_SALDO_NEGATIVO = 2948
'O saldo do cr�dito do cliente ficaria negativo
Public Const ERRO_LEITURA_VALBAIXAPERDA = 2949 'Sem parametros
'Erro na leitura do valor baixado por perda de uma parcela a receber
Public Const ERRO_CHEQUEPRE_SEM_PARCELA = 2950 'Sem parametros
'N�o encontrou parcela associada ao cheque pr�-datado.
Public Const ERRO_LEITURA_COMISSOES2 = 2951 'Sem parametros
'Erro de leitura de comiss�es de uma parcela
Public Const ERRO_ATUALIZACAO_COMISSOES2 = 2952 'Sem parametros
'Erro na grava��o de comiss�es de uma parcela
Public Const ERRO_LIBERACAO_COMISSAO_STATUS = 2953 'Sem parametros
'O status atual da comiss�o n�o � compat�vel com a sua libera��o.
Public Const ERRO_LEITURA_COBRADORES_BANCO = 2954 'Par�metro: iCodBanco
'Erro na tentativa de ler cobradores associados com o Banco %i.
Public Const ERRO_LEITURA_PARCPAG_BANCO = 2955 'Par�metro: iCodBanco
'Erro na tentativa de ler a tabela ParcelasPag com Banco %i.
Public Const ERRO_LEITURA_PARCPAGBAIXADAS_BANCO = 2956 'Par�metro: iCodBanco
'Erro na tentativa de ler a tabela ParcelasPagBaixadas com Banco %i.
Public Const ERRO_BANCO_ASSOCIADO_PARCELAPAG = 2957 'Par�metros: iCodBanco, lNumIntDoc
'N�o � permitido excluir Banco %i porque est� associado � Parcela com n�mero interno %l.
Public Const ERRO_BANCO_ASSOCIADO_PARCELAPAGBAIXADA = 2958 'Par�metros: iCodBanco, lNumIntDoc
'N�o � permitido excluir Banco %i porque est� associado � Parcela Baixada com n�mero interno %l.
Public Const ERRO_BANCO_ASSOCIADO_COBRADOR = 2959 'Par�metros: iCodBanco, iCobrador
'N�o � permitido excluir Banco %i porque est� associado ao Cobrador %i.
Public Const ERRO_LEITURA_COMISSOES3 = 2960
'Erro na leitura da tabela de Comissoes.
Public Const ERRO_INSERCAO_TESCONFIG = 2961 'Par�metros: Codigo, FilialEmpresa
'Ocorreu um erro na inser��o de um registro na tabela TESConfig. Codigo = %s, Filial = %i.
Public Const ERRO_INSERCAO_FERIADOS_FILIAIS = 2962 'Par�metros: FilialEmpresa
'Erro na inser��o dos feriados na filial %s.
Public Const ERRO_LEITURA_TIPO_OCORR_COBR = 2963 'Parametro: Codigo do Tipo
'Erro na leitura do Tipo de Ocorr�ncia de Cobran�a com c�digo %s.
Public Const ERRO_SALDO_RA_MENOR_TOTALRECEBER = 2964 'sem parametros
'O saldo do adiantamento n�o � suficiente para as baixas selecionadas
Public Const ERRO_SALDO_DB_MENOR_TOTALRECEBER = 2965 'sem parametros
'O saldo do cr�dito/devolu��o n�o � suficiente para as baixas selecionadas
Public Const ERRO_PARCELAREC_TRANSF_CHEQUEPRE = 2966 'sem parametros
'N�o pode transferir parcela com cheque pr�-datado
Public Const ERRO_VIA_TRANSPORTE_NAO_PREENCHIDO = 2967 'Sem Parametros
'Erro � obrigat�rio o preenchimento da Via de Transporte.
Public Const ERRO_TIPO_INICIAL_MAIOR = 2968 'Sem Parametros
'O tipo inicial n�o pode ser maior que o tipo final.
Public Const ERRO_LOCK_BAIXAREC = 2969
'N�o conseguiu fazer o lock da Baixa a Receber.
Public Const ERRO_REGIAO_VENDA_RELACIONADA_COM_PREVVENDA = 2970 'Sem Parametro
'N�o � poss�vel excluir Regi�o de Venda relacionada com PrevVenda.
Public Const ERRO_DEBITO_REC_SALDO_MAIOR_ORIGINAL = 2971 'Sem Parametro
'O saldo do cr�dito do cliente ficaria maior que seu valor original
Public Const ERRO_LOCK_CHEQUEPRE = 2972 'sem parametros
'Erro na tentativa de fazer 'lock' na tabela de ChequesPre
Public Const ERRO_VALOR_CHEQUEPRE_ALTERADO = 2973 'sem parametros
'O valor do cheque pr� datado foi alterado
Public Const ERRO_CHEQUEPRE_JA_DEPOSITADO = 2974 'sem parametros
'O cheque pr� datado j� foi depositado atrav�s de outro border�
Public Const ERRO_DETPAG_CPO_NAO_PREENCHIDO = 2975 'sem parametros
'Preencha os campos Fornecedor, Filial, N�mero do T�tulo e da Parcela
Public Const ERRO_VALORINSS_MAIOR = 2976 'Par�metros: sValorINSS, sValor
'Valor do INSS n�o pode ser maior do que o Valor do Titulo
Public Const ERRO_DATA_CHQPRE_DIFERENTE = 2977 'Parametros: numero da parcela
'A data para dep�sito do cheque pr� datado associado � parcela %s n�o confere com o vencimento da mesma
Public Const ERRO_VENDEDOR_NAO_FORNECIDO = 2978 'Sem Parametros
'N�o h� vendedor selecionado
Public Const EXCLUSAO_RELOPCOMISVEND = 2979 'Sem parametros
'Confirma a exclus�o da Op��o de Relat�rio ?
Public Const ERRO_LINHA_GRID_CATEGCLI_INCOMPLETA = 2980 'sem parametros
'H� uma linha do grid de categorias de clientes que n�o est� completa.
Public Const ERRO_LEITURA_NOTAS_FISCAIS_EXT_BAIXADAS = 2981 'Sem parametro
'Erro na leitura da tabela de notas fiscais baixadas.
Public Const ERRO_EXCLUSAO_NOTAS_FISCAIS_EXT_BAIXADAS = 2982 'Sem parametro
'Erro na exclus�o na tabela de notas fiscais baixadas.
Public Const ERRO_INSERCAO_NOTAS_FISCAIS_EXT = 2983 'Sem parametro
'Erro na inser��o na tabela de notas fiscais baixadas.
Public Const ERRO_PREENCHA_CAMPOS_OBRIGATORIOS = 2984
'Todos os campos devem estar preenchidos.
Public Const ERRO_VALOR_PAGAR_NEGATIVO = 2985 'Sem parametros
'O valor a pagar n�o pode ser negativo
Public Const ERRO_MULTA_JUROS_DATABAIXA_VAZIA = 2986 'Sem parametros
'Para digitar juros ou multa a data da baixa precisa estar prenchida
Public Const ERRO_JUROS_INCOMPATIVEL_DATABAIXA_DATAVENC = 2987 'Sem parametros
'Juros incompat�vel com Data de Baixa menor ou igual � Data de Vencimento.
Public Const ERRO_VALOR_RECEBER_NEGATIVO = 2988 'Sem parametros
'O valor a receber n�o pode ser negativo
Public Const ERRO_PREENCH_CPOS_OBRIG_TELA = 2989 'sem parametros
'Preencha os campos obrigat�rios da tela.
Public Const ERRO_SELECIONE_CHEQUES_NO_GRID = 2990
'Marque na coluna Selecionar do grid os cheques que deseja imprimir
Public Const ERRO_EXCLUSAO_COND_PAGAMENTO_A_VISTA = 2991
'N�o � permitida a exclus�o da condi��o de pagamento � vista.
Public Const ERRO_ALTERACAO_COND_PAGAMENTO_A_VISTA = 2992
'N�o � permitida a altera��o da condi��o de pagamento � vista.
Public Const ERRO_CAMPOS_CANCELARBAIXA_NAO_PREENCHIDOS = 2993 'Sem par�metros
'Os campos Fornecedor, Filial, Tipo, N�mero, Parcela e Sequencial devem estar preenchidos.
Public Const ERRO_VALORES_TOTAL_BASE = 2994 'Sem parametros
'A base de c�lculo da comiss�o n�o pode ser maior que o valor total.
Public Const ERRO_FLUXO_CONTA_NAO_PREENCHIDA = 2995 'Parametros iLinha
'A Conta da linha %i n�o foi preenchida.
Public Const ERRO_CCI_EXISTENTE_GRID = 2996 'Parametros sCCI, iLinha
'A Conta %s j� est� presente na linha %i do grid.
Public Const ERRO_FORNECEDOR_REPETIDO = 2997 'Parametro sFornecedor, iLinhaGrid
'O Fornecedor j� est� presente neste fluxo na linha %i.
Public Const ERRO_CLIENTE_REPETIDO = 2998 'Parametro sCliente, iLinhaGrid
'O Cliente j� est� presente neste fluxo na linha %i.
Public Const ERRO_TIPODEAPLICACAO_CODIGO_REPETIDO = 2999 'Parametros = iAplicacao, iLinha
'O Tipo de Aplica��o %i j� est� presente neste fluxo na linha %i.
Public Const ERRO_EXCLUSAO_OCORR_ALTERACAO_DATA = 16000 'sem parametros
'Uma ocorr�ncia que provocou a altera��o do vencimento da parcela deve ser desfeita pela inclus�o de uma nova ocorr�ncia alterando o vencimento para a data correta.
Public Const ERRO_OUTRA_OCORR_ALTEROU_VCTO = 16001 'sem parametros
'Outra ocorrencia posterior a esta alterou a data de vencimento da parcela.
Public Const ERRO_BAIXAREC_COBR_CART = 16002 'sem parametros
'Para este tipo de baixa a parcela deve estar em carteira
Public Const ERRO_COBRADOR_USADO_TIPOSCLIENTE = 16003 'Parametros: iCodigoCobrador, iCodigoTipoCliente
'O Cobrador %i est� sendo utilizado pelo Tipo de Cliente com o c�digo %i.
Public Const ERRO_COBRADOR_USADO_PARCELASREC = 16004 'Parametros: iCodigoCobrador
'O Cobrador %i est� sendo utilizado por uma Parcela a Receber.
Public Const ERRO_COBRADOR_USADO_PARCELASREC_BAIXADA = 16005 'Parametros: iCodigoCobrador
'O Cobrador %i est� sendo utilizado por uma Parcela a Receber Baixada.
Public Const ERRO_COBRADOR_USADO_TRANSFCARTCOBR = 16006 'Parametros: iCodigoCobrador
'O Cobrador %i foi utilizado na Transfer�ncia de Carteira de Cobran�a.
Public Const ERRO_CARTEIRACOBRADOR_USADO_BORDEROCOBRANCA = 16007 'Parametros: iCodCarteiraCobranca, iCobrador e iNumBordero
'A Carteira %i do Cobrador %i est� sendo utilizado pelo Bordero de Cobran�a n�mero %i.
Public Const ERRO_CARTEIRACOBRADOR_USADO_PARCELAREC = 16008 'Parametros: iCodCarteiraCobranca, iCobrador
'A Carteira %i do Cobrador %i est� sendo utilizado por uma Parcela a Receber.
Public Const ERRO_CARTEIRACOBRADOR_USADO_PARCELAREC_BAIXADAS = 16009 'Parametros: iCodCarteiraCobranca, iCobrador
'A Carteira %i do Cobrador %i est� sendo utilizado por uma Parcela a Receber Baixadas.
Public Const ERRO_SEQUENCIAL_INVALIDO = 16011 'Sem Parametros
'O Sequencial tem que ser um valor inteiro positivo.
Public Const ERRO_OCORR_ALT_DATA_VCTO_SEM_DATA = 16012 'sem parametros
'A nova data de vencimento tem que estar preenchida para este tipo de ocorr�ncia.
Public Const ERRO_OCORR_VCTO_ALT_ERRADA = 16013 'sem parametros
'A nova data de vencimento n�o pode estar preenchida para este tipo de ocorr�ncia.
Public Const ERRO_COMISSOES_BAIXADA_ALT_DEBITOS = 16014 'parametro iCodVendedor
'A comiss�o do vendedor %s, que est� baixada, n�o pode ser exclu�da ou ter seu valor alterado.
Public Const ERRO_CREDITOPAGAR_VINCULADO_NFISCAL = 16015 'parametro: numero do credito
'Cr�dito para com Fornecedor com n�mero %s est� vinculado � Nota Fiscal e portanto n�o pode ser exclu�do neste m�dulo.
Public Const ERRO_COMISSAO_BAIXADA_CANC_BAIXA = 16016 'sem parametros
'N�o pode cancelar uma baixa de parcela que teve a comissao j� baixada (paga).
Public Const ERRO_DATA_DESC_APOS_VCTO = 16017 'parametro: numero da parcela
'A parcela %s tem desconto posterior ao vencimento
Public Const ERRO_LEITURA_ANTECPAG_DATA = 16018 'parametros: data
'Erro n a leitura de adiantamentos a fornecedor na data %s
Public Const ERRO_LEITURA_ANTECREC_DATA = 16019 'parametros: data
'Erro na leitura de adiantamentos de cliente na data %s
Public Const ERRO_LEITURA_BAIXASPARCPAG = 16020 'Sem par�metros
'Erro na leitura da tabela BaixasParcPag.
Public Const ERRO_FILIAL_CENTR_COBR_NAO_SELEC = 16021 'sem parametros
'� preciso selecionar a filial que ir� centralizar a cobran�a ou deix�-la como independente por filial
Public Const ERRO_DATAINICIAL_NAO_PREENCHIDA = 16022
'O preenchimento da data inicial � obrigat�rio.
Public Const ERRO_DATAFINAL_NAO_PREENCHIDA = 16023
'O preenchimento da data final � obrigat�rio.
Public Const ERRO_CONTATO_NAO_PREENCHIDO = 16024
'O preenchimento do contato � obrigat�rio.
Public Const ERRO_TELCONTATO_NAO_PREENCHIDO = 16025
'O preenchimento do telefone de contato � obrigat�rio.
Public Const ERRO_ENDERECO_NAO_PREENCHIDO = 16026 'Sem parametro
'O preenchimento do Endere�o � obrigat�rio.
Public Const ERRO_COMPLEMENTO_NAO_PREENCHIDO = 16027
'O preenchimento do Complemento � obrigat�rio.
Public Const ERRO_DATA_FINAL_DO_MES = 16028 'Parametro : dtDataFinal
'A Data deve ser preenchida com o ultimo dia do m�s. Data Correta: %dt.
Public Const ERRO_DATA_INICIO_DO_MES = 16029 'Parametro : dtDataInicio
'A Data deve ser preenchida com o primeiro dia do m�s. Data Correta: %dt.
Public Const ERRO_NFPAG_SEM_DOCORIGINAL = 16030 'lNumeroNota
'A Nota Fiscal %l n�o possui documento original.
Public Const ERRO_TITULO_PAG_SEM_DOCORIGINAL = 16031 'lNumeroT�tulo
'O T�tulo %l n�o possui Documento Original cadastrado.
Public Const ERRO_TITULO_REC_SEM_DOCORIGINAL = 16032 'Sem Parametros
'Este T�tulo a Receber n�o tem Documento Original.
Public Const ERRO_ATUALIZACAO_NFISCAL2 = 16033 'Sem Parametros
'Erro na atualiza��o da tabela de NFiscal.
Public Const ERRO_BORDERO_COBRANCA_NAO_CADASTRADO_COBRADOR = 16034 'Par�metro: lNumBordero, iCobrador
'O Bordero de Cobran�a de n�mero %l n�o est� cadastrado para o Cobrador %i.
Public Const ERRO_BORDERODE_MAIOR_BORDEROATE = 16035  'Sem Parametros
'O Bordero De n�o pode ser maior que o Bordero at�.
Public Const ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO = 16036 'Par�metros: lCliente, iFilialCliente
'O Endere�o de cobran�a da Filial %i do Cliente %l n�o est� preenchido.
Public Const ERRO_FORMATO_ARQUIVO_INCORRETO = 16037 'Par�metro: sNomeArquivo
'O arquivo de nome "%s" est� em um formato incorreto.
Public Const ERRO_CONTACORRENTE_COBRADOR_NAO_ENCONTRADA = 16038
'A conta corrente associada ao cobrador %i n�o foi encontrada no banco de dados.
Public Const ERRO_OCORRREMPARCREC_NAO_CADASTRADA = 16039
'A Ocorr�ncia com o Numero Interno %l n�o est� cadastrada no Banco de Dados.
Public Const ERRO_ALTERACAO_BORDERO_PAGTO = 16040
'Erro na atualiza��o da tabela de borderos de Pagamento.
Public Const ERRO_BANCO_CCI_DIFERENTE_COBRADOR = 16041 'Par�metros: iCodBancoConta, iCodBancoCobrador
'A conta corrente do Cobrador deve pertencer ao banco do cobrador. Banco Conta: %i, Banco Cobrador: %i.
Public Const ERRO_AGENCIA_CONTA_COBRADOR_NAO_PREENCHIDAS = 16042
'Os dados n�mero da conta e a ag�ncia n�o est�o preenchidos na conta corrente do cobrador.
Public Const ERRO_FAIXA_NOSSONUMERO_INSUFICIENTE = 16043
'A faixa de nosso n�mero dipon�vel n�o � suficiente para a gera��o do arquivo de remessa.
Public Const ERRO_INSERCAO_BORDERO_COBRANCA_RETORNO = 16044
'Erro na inser��o de Bordero de Retorno de Cobran�a.
Public Const ERRO_INSERCAO_OCORR_REM_PARC_RET = 16045
'Erro na inser��o de ocorr�ncia de retorno de remessa de Bordero de Cobran�a.
Public Const ERRO_BORDERO_PAGTO_INEXISTENTE1 = 16046 'Par�metro:lNumIntBordero
'O Border� de Pagamento com n�mero interno %l n�o est� cadastrado.
Public Const ERRO_LOCK_BORDERO_PAGTO = 16047 'Par�metro:lNumIntBordero
'Erro na tentativa de fazer lock no bordero de pagamento de n�mero interno %l.
Public Const ERRO_CARTEIRACOBRADOR_NAO_CADASTRADA1 = 16048 'Par�metros: iCarteira, iCobrador
'A Carteira %i do Cobrador %i n�o est� cadastrada no Bano de Dados.
Public Const ERRO_LEITURA_BANCOSINFO1 = 16049
'Erro de leituta na tabela de informa��es banc�rias CNAB.
Public Const ERRO_INSERCAO_BANCOSINFO = 16050
'Erro na inclus�o de registro na tabela BancosInfo
Public Const ERRO_ALTERACAO_BANCOSINFO = 16051
'Erro na altera��o de registro na tabela BancosInfo
Public Const ERRO_CARTEIRACOBRADORINFO_NAO_CADASTRADA1 = 16052 'Par�metros:  iCodCobrador
'As informa��es banc�rias daS carteiras do cobrador %i n�o est�o cadastradas.
Public Const ERRO_CARTEIRACOBRADORINFO_NAO_CADASTRADA = 16053 'Par�metros:  iCodCarteira, iCodCobrador
'As informa��es banc�rias da carteira %i do cobrador %i n�o est�o cadastradas.
Public Const ERRO_LEITURA_BAIXASEC = 16054 'Sem par�metros
'Erro na leitura da tabela BaixasRec.
Public Const ERRO_BAIXAREC_NAO_ENCONTRADA = 16055 'Par�metros: lNumIntParcRec
'N�o foi poss�vel encontrar em BaixaRec a Data da Baixa relacionada a parcela %l.
Public Const ERRO_LEITURA_PARCELASPAG2 = 16056 'Sem Parametros
'Erro na tentativa de ler a tabela ParcelasPag.
Public Const ERRO_CONTACORRENTE_COBRANCA_NAO_INFORMADA = 16057
'Para a utiliza��o da cobran�a eletr�nica a conta corrente deve ser informada.
Public Const ERRO_NOSSONUMERO_INICIAL_MAIOR = 16058 'Par�metro: iCarteira
'Erro nos dados da Carteira %i. O n�mero inicial � maior que o n�mero final.
Public Const ERRO_PROXNOSSONUMERO_MENOR = 16059 'Par�metro: iCarteira
'Erro nos dados da Carteira %i. O valor do pr�ximo n�mero n�o pode ser inferior ao valor inicial.
Public Const ERRO_PROXNOSSONUMERO_MAIOR = 16060 'Par�metro: iCarteira
'Erro nos dados da Carteira %i. O valor do pr�ximo n�mero n�o pode ser superior ao valor final.
Public Const ERRO_FAIXA_NOSSONUMERO_IMCOMPLETA = 16061 'Par�metro: iCarteira
'Erro nos dados da Carteira %i. Quando a faixa de n�mero for informada os campos dos Numeros Inicial, Final e pr�ximo dever�o ser preenchidos.
Public Const ERRO_CRITERIOS_NAO_SELECIONADOS = 16062 'Sem Par�metros
'Pelo menos um dos crit�rios de concilia��o deve ser escolhido.
Public Const ERRO_EXTRATOS_NAO_ENCONTRADOS = 16063
'Nenhum extrato foi encontrado para concilia��o.
Public Const ERRO_EXTRATO_NAO_ENCONTRADO = 16064 'Par�metro: iNumExtrato
'O Extrato de n�mero %i n�o foi encontrado no Banco de Dados.
Public Const ERRO_TIPOMOVCCI_NAO_CADASTRADO = 16065 'Par�metro: iTipo
'O Tipo de movimento de conta corrente %i n�o esta cadastrado.
Public Const ERRO_LEITURA_CCIMOV1 = 16066 'Parametros iCodconta
'Ocorreu um erro na leitura da tabela de Saldos Mensais de Conta Corrente. Conta = %s.
Public Const ERRO_CONTACORRENTEINTERNA_NAO_CADASTRADA = 16067 'Parametro: iCodConta
'A Conta Corrente %i n�o est� cadastrada.
Public Const ERRO_LEITURA_CCIMOVDIA2 = 16068 'parametros iCodConta.
'Ocorreu um erro na Leitura da Tabela de Saldos Diarios de Conta Corrente. Conta = %i.
Public Const ERRO_DIRETORIO_INVALIDO = 16069 'Par�metro: sDiret�rio
'O Diret�rio %s informado n�o foi encontrado.
Public Const ERRO_AUSENCIA_BORDEROSCOBRANCA = 16070
'N�o existem novos borderos de cobran�a para remessa.
Public Const ERRO_DRIVE_NAO_ACESSIVEL = 16071 'Par�metro:
' %s n�o est� acess�vel.
Public Const ERRO_LEITURA_COBRADORINFO = 16072 'Par�metro: iCodCobrador
'Erro na leitura das informa��es banc�rias do cobrador %i.
Public Const ERRO_EXCLUSAO_COBRADORINFO = 16073 'Par�metro: iCodCobrador
'Erro na tentativa de exclus�o das informa��es banc�rias do cobrador %i.
Public Const ERRO_INSERCAO_COBRADORINFO = 16074 'Par�metro: iCodCobraddor
'Erro na tentativa de inser��o das informa��es banc�rias do cobrador %i.
Public Const ERRO_LEITURA_CARTEIRACOBRADORINFO = 16075 'Par�metro: icodcartera,iCodCobrador
'Erro na leitura das informa��es banc�rias da carteira %i do cobrador %i.
Public Const ERRO_EXCLUSAO_CARTEIRACOBRADORINFO = 16076 'Par�metro: iCarteira, iCodCobrador
'Erro na tentativa de exclus�o das informa��es banc�rias da carteira %i  do cobrador %i.
Public Const ERRO_INSERCAO_CARTEIRASCOBRADORINFO = 16077 'Par�metro: iCarteira, iCodCobrador
'Erro na tentativa de inser��o das informa��es banc�rias da carteira %i do cobrador %i.
Public Const ERRO_LEITURA_TIPOSDELCTOCNAB = 16078
'Erro na eitura da tabela de Tipos de Lancamentos CNAB.
Public Const ERRO_EXCLUSAO_TIPOSDELANCTOCNAB = 16079
'Erro na exclus�o de registro na tabela de Tipos de Lancamentos CNAB
Public Const ERRO_INCLUSAO_TIPOSDELANCTOCNAB = 16080
'Erro na inclus�o de registro na tabela de Tipos de Lancamentos CNAB
Public Const ERRO_CODLANCAMENTO_GRID_NAO_PREENCHIDO = 16081 'Par�metro: iLinha
'O c�digo do lan�amento na linha %i n�o foi informado
Public Const ERRO_DESCLANCAMENTO_GRID_NAO_PREENCHIDO = 16082 'Par�metro: iLinha
'A descri��o do lan�amento na linha %i n�o foi informado
Public Const ERRO_CODIGO_REPETIDO_GRID = 16083 'lCodigo as long
'O C�digo %l j� foi preenchido no Grid.
Public Const ERRO_PAGTOSANTECIPADOS_SUPERIOR_VALOR_PEDCOMPRAS = 16084 'sem parametros
'O total de pagamentos antecipados supera o valor total do Pedido de Compra informado.
Public Const ERRO_FILIAL_FORNECEDOR_REL_PRODUTO = 16085 'iFilialForn, lForncedor, sProduto
'A Filial %s do Fornecedor %s est� relacionada com o Produto %s.
Public Const ERRO_FILIAL_FORNECEDOR_REL_PEDCOMPRA = 16086 'iFilialForn, lForncedor, lPedCOmpra
'A Filial %s do Fornecedor %s est� relacionada com o Pedido de Compra %s.
Public Const ERRO_FILIAL_FORNECEDOR_REL_REQCOMPRA = 16087 'Par�mtros: iCodFilial, lCodFornecedor, lCod
'A Filial %s do Fornecedor %s est� relacionada com a Requisicao de Compra %s.
Public Const ERRO_FILIAL_FORNECEDOR_REL_REQMODELO = 16088 'Par�mtros: iCodFilial, lCodFornecedor, lCod
'A Filial %s do Fornecedor %s est� relacionada com a Requisicao Modelo %s.
Public Const ERRO_FILIAL_FORNECEDOR_REL_COTACAO = 16089 'Par�mtros: iCodFilial, lCodFornecedor, lCod
'A Filial %s do Fornecedor %s est� relacionada com a Cota��o %s.
Public Const ERRO_FILIAL_FORNECEDOR_REL_COTACAOPRODUTO = 16090 'Par�mtros: iCodFilial, lCodFornecedor
'A Filial %s do Fornecedor %s est� relacionada com uma Cota��o de Produto.
Public Const ERRO_FILIAL_FORNECEDOR_REL_PEDCOTACAO = 16091 'Par�mtros: iCodFilial, lCodFornecedor, lCod
'A Filial %s do Fornecedor %s est� relacionada com o Pedido de Cota��o %s.
Public Const ERRO_FILIAL_FORNECEDOR_REL_CONCORRENCIA = 16092 'Par�mtros: iCodFilial, lCodFornecedor, lCod
'A Filial %s do Fornecedor %s est� relacionada com a Concorr�ncia %s.
Public Const ERRO_LEITURA_CCIMOV3 = 16093 'Sem Parametros
'Erro na leitura da tabela de Saldos Mensais de Conta Corrente.
Public Const ERRO_LEITURA_IMPORTFORN = 16094 'Sem parametros
'Erro na leitura da tabela ImportForn.
Public Const ERRO_FORNECEDOR_FINAL_MENOR = 16095 'Sem Parametros
'Fornecedor Final Menor que o Fornecedor Inicial
Public Const EXCLUSAO_RELOPRAZAOCP = 16096 'Sem parametros
'Confirma a exclus�o da Op��o de Relat�rioCP ?
Public Const ERRO_CLIENTE_FINAL_MENOR = 16097 'Sem Parametros
'Cliente Final Menor que o Cliente Inicial
Public Const EXCLUSAO_RELOPRAZAOCR = 16098 'Sem parametros
'Confirma a exclus�o da Op��o de Relat�rioCP ?
Public Const ERRO_SUBTIPOCONTABIL_TIPOBAIXA_NAO_ENCONTRADO = 16099 'Sem par�metros
'N�o foi encontrada transa��o na tabela TransacaoCTB correspondente ao Tipo de Baixa selecionado.
Public Const ERRO_LEITURA_MOTIVOSBAIXA = 16100 'Par�metro: iCodigo
'Ocorreu um erro na leitura da tabela de motivos de baixa de t�tulos a pagar/receber (MotivosBaixa). Motivo = %s
Public Const ERRO_MOTIVOBAIXA_NAO_ENCONTRADO = 16101 'Par�metro: iCodigo
'O motivo de baixa de t�tulos a pagar/receber n�o foi encontrado. Motivo = %s
Public Const ERRO_CREDITO_PAGAR_SEM_DOC_ORIGINAL = 16102 'Par�metro: iNumTitulo
'O t�tulo de cr�dito com fornecedores n�mero %s, n�o possui documento original.
Public Const ERRO_DATAEMISSAO_OBRIGATORIA_DOC_ORIGINAL = 16103 'Sem par�metros
'Para consultar o documento original de um t�tulo, � necess�rio preencher a data de emiss�o.
Public Const ERRO_CREDITOPAGAR_NAO_CADASTRADO2 = 16104 'Par�metros: lNumTitulo, lFornecedor, iFilial, sSiglaDocumento, dtDataEmissao
'O t�tulo de Cr�dito com Fornecedor com os dados abaixo, n�o est� cadastrado.
'N�mero: %s, Fornecedor: %s, Filial: %s, Tipo: %s e Data de Emiss�o: %s .
Public Const ERRO_DEBITO_RECEBER_SEM_DOC_ORIGINAL = 16105 'Par�metro: iNumTitulo
'O t�tulo de d�bito com clientes n�mero %s, n�o possui documento original.
Public Const ERRO_SIGLA_DOCUMENTO_NAO_PREENCHIDO = 16106
'A sigla do documento n�o foi informada
Public Const ERRO_LEITURA_RECEBANTECIPADOS1 = 16107 'Par�metro: lCodCliente, iCodFilial
'Erro na leitura da tabela RecebAntecipados com Cliente %l e Filial %i.
Public Const ERRO_SEM_CHEQUES_SEL = 16108
'Os Par�metros para a sele��o de cheques n�o foram informados
Public Const ERRO_RETCOBR_EXCLUSAO_ERROS = 16109
'Erro na exclus�o de registros de erro do processamento de retorno de cobran�a
Public Const ERRO_RETCOBR_INCLUSAO_ERRO = 16110
'Erro na inclus�o de registro de erro de processamento de retorno de cobran�a
Public Const ERRO_CARTEIRA_COBR_INVALIDA = 16111 'parametros cod carteira no banco, cod cobrador
'N�o est� cadastrada a carteira com c�digo % no cobrador %s
Public Const ERRO_PARCBORDRETCOBR_NENHUMA = 16112 'parametro: "seu numero"
'N�o foi encontrada nenhuma parcela em aberto identificada por %s
Public Const ERRO_PARCBORDRETCOBR_VARIAS = 16113 'parametro: "seu numero"
'Foram encontradas mais de uma parcela em aberto identificadas por %s
Public Const ERRO_LEITURA_PARCREC_RETCOBR = 16114 'sem parametros
'Erro na leitura de parcela para o tratamento do retorno da cobranca
Public Const ERRO_NOSSO_NUMERO_NAO_DEFINIDO = 16115 'Par�metros: iCodBanco, iCodCarteira
'Para gerar um border� de cobran�a, � necess�rio definir o pr�ximo 'Nosso N�mero' para a carteira de cobran�a %s do cobrador %s.
Public Const ERRO_LEITURA_TITULOREC_CHEQUEPRE = 16116 'Par�metros: lNumIntCheque
'Ocorreu um erro na leitura do T�tulo a Receber vinculado ao Cheque Pr�-datado com o n�mero interno %s.
Public Const ERRO_TITULOREC_CHEQUEPRE_NAO_ENCONTRADO = 16117 'Par�metros:
'N�o foi encontrado t�tulo a receber vinculado ao cheque pr�-datado com os seguintes dados: Cliente= %s, N�mero= %s e Valor= %s.
Public Const ERRO_SEM_CHEQUE_SELECIONADO = 16118
'N�o existe cheque selecionado
Public Const ERRO_DATAEMISSAO_NAO_INFORMADA = 16119
'O campo Data de Emiss�o � de preenchimento obrigat�rio
Public Const ERRO_FAVORECIDO_NAO_PREENCHIDO = 16120
'O Favorecido n�o foi informado
Public Const ERRO_CODIGO_NAO_INFORMADO1 = 16121
'O C�digo n�o foi informado
Public Const ERRO_DATA_APLICACAO_NAO_INFORMADA = 16122
'A Data de Aplica��o n�o foi informada
Public Const ERRO_CONTA_CORRENTEORIGEM_NAO_ENCONTRADA = 16123 'Parametro ContaOrigem.Text
'A Conta Corrente de Origem %s n�o foi encontrada - ContaOrigem.Text
Public Const ERRO_CONTACORRENTEORIGEM_NAO_BANCARIA = 16124
'A Conta Corrente de Origem n�o � banc�ria
Public Const ERRO_CONTACORRENTEDESTINO_NAO_ENCONTRADA = 16125 'Parametro ContaDestino.Text
'A Conta Corrente de Destino %s n�o foi encontrada - ContaDestino.Text
Public Const ERRO_CONTACORRENTEDESTINO_NAO_BANCARIA = 16126
'A Conta Corrente de Destino n�o � banc�ria






'C�DIGOS DE AVISO - Reservado de 5200 at� 5399
Public Const AVISO_CONFIRMA_EXCLUSAO_REGIOESVENDAS = 5200 'Parametro iCodigo
'Regi�o de Venda com c�digo %s ser� exclu�da. Cofirma a exclus�o?
Public Const AVISO_CONFIRMA_EXCLUSAO_BANCO = 5201 'Parametro iCodBanco
'Banco com c�digo %s ser� exclu�do. Confirma a exclus�o?
Public Const AVISO_CRIAR_TABELA_PRECO = 5202
'Deseja cadastrar nova Tabela de Pre�o?
Public Const AVISO_CRIAR_CONDICAO_PAGAMENTO = 5203
'Deseja cadastrar nova Condi��o de Pagamento?
Public Const AVISO_CRIAR_MENSAGEM = 5204
'Deseja cadastrar nova Mensagem?
Public Const AVISO_CRIAR_REGIAO = 5205
'Deseja cadastrar nova Regi�o de Venda?
Public Const AVISO_CRIAR_COBRADOR = 5206
'Deseja cadastrar novo Cobrador?
Public Const AVISO_CRIAR_TRANSPORTADORA = 5207
'Deseja cadastrar nova Transportadora?
Public Const AVISO_CRIAR_TIPOCLIENTE = 5208
'Deseja cadastrar novo Tipo de Cliente?
Public Const AVISO_EXCLUIR_CLIENTE = 5209 'Parametro: lCodCliente
'Confirma exclus�o do Cliente com c�digo %l, Matriz e suas Filiais?
Public Const AVISO_EXCLUIR_FILIALCLIENTE = 5210
'Confirma exclus�o de Filial Cliente?
Public Const AVISO_CRIAR_TIPOFORNECEDOR = 5211
'Deseja cadastrar novo Tipo de Fornecedor?
Public Const AVISO_EXCLUIR_FORNECEDOR = 5212 'Parametros: lCodFornecedor
'Confirma exclus�o do Fornecedor com c�digo %l, Matriz e suas Filiais?
Public Const AVISO_CRIAR_FORNECEDOR = 5213
'Deseja cadastrar novo Fornecedor?
Public Const AVISO_EXCLUIR_FILIAL_FORNECEDOR = 5214
'Confirma exclus�o de Filial Fornecedor?
Public Const AVISO_DATA_VALOR_NAO_ALTERAVEIS = 5215
'Os campos Data e Valor n�o ser�o alterados pois o movimento j� foi conciliado. Deseja Prosseguir ?
Public Const AVISO_CODCONTACORRENTE_INEXISTENTE = 5216 'Paramento: iCodigo
'A conta corrente %s n�o existe. Deseja Cri�-la?
Public Const AVISO_TIPOMEIOPAGTO_INEXISTENTE = 5217 'Parametro: iTipoMeioPagto
'O Tipo de Pagamento %s n�o existe. Deseja Cri�-lo?
Public Const AVISO_FAVORECIDO_INEXISTENTE = 5218 'Parametro: iCodigo
'O Favorecido %s n�o existe. Deseja criar novo favorecido?
Public Const AVISO_CONFIRMA_EXCLUSAO_SAQUE = 5219 'Parametros: iCodconta, lSequencial
'Confirma a exclus�o do saque na conta %s , sequencial %s
Public Const AVISO_EXCLUIR_TIPOCLIENTE = 5220
'Confirma exclus�o de Tipo de Cliente?
Public Const AVISO_CRIAR_PADRAOCOBRANCA = 5221
'Padr�o Cobran�a n�o est� cadastrado. Deseja criar?
Public Const AVISO_CRIAR_FORNECEDOR_1 = 5222 'Parametro: sNomeReduzido
'Fornecedor %s n�o est� cadastrado. Deseja criar?
Public Const AVISO_CRIAR_FORNECEDOR_2 = 5223 'Parametro: lCodigo
'Fornecedor com c�digo %s n�o est� cadastrado. Deseja criar?
Public Const AVISO_CRIAR_FORNECEDOR_3 = 5224 'Parametro: sCGC
'Fornecedor com CGC/CPF %s n�o est� cadastrado. Deseja criar?
Public Const AVISO_CODBANCO_INEXISTENTE = 5225 'Parametro: Codigo do Banco
'O Banco com c�digo %s n�o est� cadastrado. Deseja criar?
Public Const AVISO_CONFIRMA_EXCLUSAO_CONTACORRENTE = 5226 'Parametro: codconta
'Confirma a exclus�o da Conta Corrente?
Public Const AVISO_CONFIRMA_EXCLUSAO_HISTMOVCTA = 5227 'Parametro: c�digo
'Confirma a exclus�o do Hist�rico Padr�o de Movimenta��o de Conta %s ?
Public Const AVISO_CRIAR_PADRAO_COBRANCA = 5228
'Deseja cadastrar novo Padr�o Cobran�a ?
Public Const AVISO_CONTAORIGEM_INEXISTENTE = 5229 'Param : conta
'A conta Origem %s nao esta cadastrada. Deseja Cria-la?
Public Const AVISO_CONTADESTINO_INEXISTENTE = 5230 'Parametro: iconta
'A conta Destino %s nao esta cadastrada. Deseja Cri�-la?
Public Const AVISO_CONFIRMA_EXCLUSAO_TIPOAPLICACAO = 5231 'Sem par�metro
'Confirma exclus�o do Tipo de aplica��o ?
Public Const AVISO_CONFIRMA_EXCLUSAO_DEPOSITO = 5232 'Parametro: COdconta, sequencial
'Confirma a exclus�o do deposito da conta %s do sequencial %s
Public Const AVISO_NUM_MOV_ULTRAPASSOU_LIMITE = 5233 'Parametro Limite de Movimentos Exibidos
'O n�mero de movimentos para a condi��o de sele��o atual ultrapassou o limite do sistema. Somente os primeiros %i movimentos ser�o exibidos.
Public Const AVISO_NUM_LANCEXTRATO_ULTRAPASSOU_LIMITE = 5234 'Parametro Limite de Lan�ametos de Extrato Exibidos
'O n�mero de lan�amentos de extrato banc�rio para a condi��o de sele��o atual ultrapassou o limite do sistema. Somente os primeiros %i lan�amentos ser�o exibidos.
Public Const AVISO_CONCILIACAO_TOTAL_MOV_EXT_DIFERENTES = 5235 'Parametros Total de Movimentos e Total de Lancamentos de Extrato
'O total de movimento = %d n�o coincide com o total de extrato = %d. Confirma a concilia��o?
Public Const AVISO_FLUXO_DATAFINAL_INALTERAVEL = 5236 'Sem Parametros
'A Data Final do Fluxo de Caixa n�o ser� alterada. Deseja prosseguir com a grava��o?
Public Const AVISO_CONFIRMA_EXCLUSAO_FLUXO = 5237 'Parametro: Identificador do Fluxo
'Confirma a exclus�o do Fluxo de Caixa %s ?
Public Const AVISO_NUM_FLUXO_PAG_ULTRAPASSOU_LIMITE = 5238 'Data, Parametro Limite de Pagamentos do Fluxo de Caixa
'O n�mero de Pagamentos para a data %s ultrapassou o limite do sistema. Somente os primeiros %l pagamentos ser�o exibidos.
Public Const AVISO_FILIALFORNECEDOR_INEXISTENTE = 5239 'Parametro: Filial
'A Filial %s n�o existe. Deseja cri�-la?
Public Const AVISO_CRIAR_FILIAL_FORNECEDOR = 5240 'Sem par�metros
'Deseja cadastrar nova Filial de Fornecedor?
Public Const AVISO_CONFIRMA_EXCLUSAO_ANTECIPPAG = 5241 'Parametros: iCodconta
'Confirma a exclus�o do Pagamento antecipado na conta %i ?
Public Const AVISO_NFPAG_BAIXADA_MESMO_NUMERO = 5242 'Par�metros :lFornecedor, iFilial, lNumNotaFiscal, dtDataEmissao
'Existe Nota Fiscal Baixada com os dados: C�digo Fornecedor = %l, C�digo Filial = %i, N�mero = %l, Data Emissao = %dt. Deseja prosseguir na inser��o de nova Nota Fiscal com o mesmo n�mero?
Public Const AVISO_NFPAG_PENDENTE_MESMO_NUMERO = 5243 'Par�metros :lFornecedor, iFilial, lNumNotaFiscal, dtDataEmissao
'Existe Nota Fiscal Pendente com os dados: C�digo Fornecedor = %l, C�digo Filial = %i, N�mero = %l, Data Emissao = %dt. Deseja prosseguir na inser��o de nova Nota Fiscal com o mesmo n�mero?
Public Const AVISO_NFPAG_MESMO_NUMERO = 5244 'Par�metros :lFornecedor, iFilial, lNumNotaFiscal, dtDataEmissao
'Existe Nota Fiscal a Pagar com os dados: C�digo Fornecedor = %l, C�digo Filial = %i, N�mero = %l, Data Emissao = %dt. Deseja prosseguir na inser��o de nova Nota Fiscal com o mesmo n�mero?
Public Const AVISO_EXCLUSAO_NFPAG = 5245 'Parametro: lNumNotaFiscal
'Confirma a exclus�o da Nota Fiscal n�mero %l ?
Public Const AVISO_CRIAR_FILIALFORNECEDOR = 5246 'Parametros: sFornecedorNomeRed, iCodFilial
'A Filial %i do Fornecedor %s n�o est� cadastrada. Deseja criar ?
Public Const AVISO_DATAVENCIMENTO_ALTERAVEL = 5247
'Todos os campos, com exce��o da Data de Vencimento, n�o ser�o alterados. Deseja prosseguir?
Public Const AVISO_NUM_FLUXO_RECEB_ULTRAPASSOU_LIMITE = 5248 'Data, Parametro Limite de Recebimentos do Fluxo de Caixa
'O n�mero de Recebimentos para a data %s ultrapassou o limite do sistema. Somente os primeiros %l recebimentos ser�o exibidos.
Public Const AVISO_NUM_FLUXOFORN_ULTRAPASSOU_LIMITE = 5249 'Data, Parametro Limite de Pagamentos por Fornecedor do Fluxo de Caixa
'O n�mero de Fornecedores para a data %s ultrapassou o limite do sistema. Somente os primeiros %l fornecedores ser�o exibidos.
Public Const AVISO_CRIAR_CONDICAOPAGTO = 5250 'Par�metro: iCondicaoPagto
'A Condi��o de Pagamento %i n�o est� cadastrada. Deseja cadastr�-la?
Public Const AVISO_DATAVENCIMENTO_SUSPENSO_ALTERAVEIS = 5251
'Todos os campos com exce��o das Datas de Vencimento, Cobran�a e Suspenso no Grid de Parcelas n�o ser�o alterados. Deseja proseguir na altera��o?
Public Const AVISO_NFFAT_PENDENTE_MESMO_NUMERO = 5252 'Par�metros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Nota Fiscal Fatura Pendente com os dados C�digo do Fornecedor = %l, C�digo da Filial = %i, N�mero = %l, Data Emiss�o = %dt. Deseja prosseguir na inser��o de nova Nota Fiscal Fatura com o mesmo n�mero?
Public Const AVISO_NFFAT_MESMO_NUMERO = 5253 'Par�metros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Nota Fiscal Fatura a Pagar com os dados C�digo do Fornecedor = %l, C�digo da Filial = %i, N�mero = %l, Data Emiss�o = %dt. Deseja prosseguir na inser��o de nova Nota Fiscal Fatura com o mesmo n�mero?
Public Const AVISO_NFFAT_BAIXADA_MESMO_NUMERO = 5254 'Par�metros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Nota Fiscal Fatura Baixada com os dados C�digo do Fornecedor = %l, C�digo da Filial = %i, N�mero = %l, Data Emiss�o = %dt. Deseja prosseguir na inser��o de nova Nota Fiscal Fatura com o mesmo n�mero?
Public Const AVISO_EXCLUSAO_NFFATPAG = 5255 'Par�metro: lNumTitulo
'Confirma a exclus�o da Nota Fiscal n�mero %l ?
Public Const AVISO_CRIAR_FILIALCLIENTE = 5256 'iCodigoFilial,sClienteNomeReduzido
'A Filial %i do Cliente %s n�o existe na tabela de FiliaisClientes. Deseja Criar ?
Public Const AVISO_CONFIRMA_EXCLUSAO_CHEQUEPRE = 5257 'Parametros: lCliente, iFilial, iBanco, sAgencia, sContaCorrente, lNumero
'Confirma a exclus�o do cheque pre com Cliente %l, Filial %i, Banco %i, Ag�ncia %s, ContaCorrente %s e N�mero %l ?
Public Const AVISO_DATAVENCIMENTO_PARCELA_DIFERENTE_DATADEPOSITO = 5258 'Par�metros: dtDataVencimento , sDataDeposito,iParcela
'A Data de Vencimento dt da Parcela %i � diferente da Data de Dep�sito %dt. Deseja prosseguir ?
Public Const AVISO_CONFIRMA_EXCLUSAO_ANTECIPREC = 5259 'Parametro iCodConta, lSequencial
'O Recebimento antecipado com Conta Corrente %i ser� exclu�do. Confirma a exclus�o?
Public Const AVISO_CRIAR_FILIAL_CLIENTE = 5260 'Sem par�metros
'Deseja cadastrar nova Filial de Cliente?
Public Const AVISO_FILIALCLIENTE_INEXISTENTE = 5261 'Parametro: Filial
'A Filial %s n�o existe. Deseja cri�-la ?
Public Const AVISO_CODTIPOAPLICACAO_INEXISTENTE = 5262  'Paramento: iCodigo
'O Tipo De Aplica��o com c�digo %i n�o existe. Deseja Cri�-lo?
Public Const AVISO_CAMPOS_APLICACAO_NAO_ALTERAVEIS = 5263 'Sem parametro
'Os campos, com exce��o dos dados de resgate, n�o ser�o alterados. Deseja Prosseguir?
Public Const AVISO_CONFIRMA_EXCLUSAO_APLICACAO = 5264 'Parametros: lCodigo
'Confirma a exclus�o da aplicacao com C�digo %l?
Public Const AVISO_APLICACAO_MOV_CONCILIADO = 5265 'Parametro: lCodigo
'O movimento de conta corrente associado a aplica��o %l est� conciliado. Deseja prosseguir?
Public Const AVISO_CONFIRMA_EXCLUSAO_RESGATE = 5266 'Parametro: iSequencialResgate, lCodigoAplicacao
'Confirma a exclus�o do Resgate %i da Aplica��o %l ?
Public Const AVISO_CAMPOS_RESGATE_NAO_ALTERAVEIS = 5267 'Sem parametro
'Os campos, com exce��o dos dados Hist�rico, Documento Externo e de Reaplica��o do Saldo Atual, n�o ser�o alterados. Deseja Prosseguir?
Public Const AVISO_NUM_FLUXOTIPOFORN_PAGTO_ULTRAPASSOU_LIMITE = 5268 'Data, Parametro Limite de Pagamentos do FluxoTipoForn
'O n�mero de Pagamentos para a data %s ultrapassou o limite do sistema. Somente os primeiros %l pagamentos ser�o exibidos.
Public Const AVISO_NUM_FLUXOTIPOFORN_RECEBTO_ULTRAPASSOU_LIMITE = 5271 'Data, Parametro Limite de Recebimentos do FluxoTipoForn
'O n�mero de Recebimentos para a data %s ultrapassou o limite do sistema. Somente os primeiros %l recebimentos ser�o exibidos.
Public Const AVISO_NUM_FLUXOCLI_ULTRAPASSOU_LIMITE = 5272 'Data, Parametro Limite de Recebimentos por Cliente do Fluxo de Caixa
'O n�mero de Clientes para a data %s ultrapassou o limite do sistema. Somente os primeiros %l clientes ser�o exibidos.
Public Const AVISO_TIPODEAPLICACAO_CODIGO_NAO_CADASTRADO = 5275 'Parametro: Codigo Tipo de Aplicacao
'O c�digo %s n�o est� cadastrado.
Public Const AVISO_NUM_FLUXOTIPOAPLIC_ULTRAPASSOU_LIMITE = 5276 'Data, Parametro Limite de Tipos de Aplica��o do Fluxo de Caixa
'O n�mero de tipos de aplica��o para a data %s ultrapassou o limite do sistema. Somente os primeiros %l recebimentos ser�o exibidos.
Public Const AVISO_TIPODEAPLICACAO_CODIGO_DEVE_SER_PREENCHIDO = 5277 'Sem parametros
'O c�digo do tipo de aplica��o deve ser preenchido
Public Const AVISO_NUM_FLUXOAPLIC_ULTRAPASSOU_LIMITE = 5279 'Data, Parametro Limite de Aplica��es do FluxoAplic
'O n�mero de Aplica��es para a data %s ultrapassou o limite do sistema. Somente as primeiras %l aplica��es ser�o exibidas.
Public Const AVISO_CONTA_JA_CADASTRADA_FLUXO = 5280 'Sem Parametro
'Conta j� cadastrada neste fluxo
Public Const AVISO_NUM_FLUXOSALDOSINICIAIS_ULTRAPASSOU_LIMITE = 5281 'Fluxo, Parametro Limite de Recebimentos do Fluxo de Caixa
'O n�mero de contas para o fluxo %l ultrapassou o limite do sistema. Somente os primeiros %l recebimentos ser�o exibidos.
Public Const AVISO_FLUXOSALDOINICIAL_CODIGO_DEVE_SER_PREENCHIDO = 5282 'Sem parametros
'O c�digo da conta corrente interna deve ser preenchido
Public Const AVISO_NUM_FLUXOSINTETICO_ULTRAPASSOU_LIMITE = 5283 'Parametro N�mero do fluxo, Limite de opera��es por Fluxo Sint�tico
'O n�mero de opera��es pa6ra o fluxo %l ultrapassou o limite do sistema. Somente as primeiras %l opera��es ser�o exibidas.
Public Const AVISO_CONFIRMA_EXCLUSAO_CREDITOPAGAR = 5284 'Par�metros: sSiglaDocumento, lNumTitulo
'Confirma a exclus�o de Devolu��es / Cr�dito com Tipo %s e N�mero %l ?
Public Const AVISO_CONFIRMA_EXCLUSAO_TIPODEFORNECEDOR = 5286  'Parametro iCodigo
'O Tipo de Fornecedor %i ser� exclu�do. Confirma exclus�o?
Public Const AVISO_NAO_E_PERMITIDO_ALTERACOES_DEBRECCLI_LANCADO = 5288 'Parametro: lNumTitulo
'N�o � permitida altera��es de campos que n�o sejam do Grid de Comiss�es de D�bito / Devolu��o N�mero do T�tulo %l porque est� lan�ado. Deseja prosseguir nas altera��o dos campos alter�veis ?
Public Const AVISO_NAO_E_PERMITIDO_EXCLUSAO_DEBRECCLI_BAIXADO = 5289 'Parametro: lCliente, iFilial, sSiglaDocumento, lNumTitulo
'N�o � possivel excluir Devolu��o / Cr�dito porque est� baixado. Cliente %l, Filial %i, Tipo de Documento %s e N�mero do T�tulo %l.
Public Const AVISO_EXCLUSAO_DEBITORECCLI = 5290 'Par�metro: lNumTitulo
'Confirma a exclus�o do D�bito a Receber N�mero %l ?
Public Const AVISO_CRIAR_COBRADOR1 = 5291 'Par�metro: iCodigo
'O Cobrador %i n�o est� cadastrado. Deseja Cri�-lo?
Public Const AVISO_EXCLUSAO_FATURAPAGAR = 5292 'Par�metro: lNumTitulo
'Confirma exclus�o da Fatura n�mero %l ?
Public Const AVISO_FATURAPAG_PENDENTE_MESMO_NUMERO = 5293 'Par�metros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Fatura Pendente com os dados: C�digo do Fornecedor = %l, C�digo da Filial = %i, N�mero = %l, Data Emiss�o = %dt. Deseja prosseguir na inser��o de nova Fatura com o mesmo n�mero?
Public Const AVISO_FATURAPAG_BAIXADA_MESMO_NUMERO = 5294 'Par�metros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Fatura Baixada com os dados: C�digo do Fornecedor = %l, C�digo da Filial = %i, N�mero = %l, Data Emiss�o = %dt. Deseja prosseguir na inser��o de nova Fatura com o mesmo n�mero?
Public Const AVISO_FATURAPAG_MESMO_NUMERO = 5295 'Par�metros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Fatura a Pagar com os dados: C�digo do Fornecedor = %l, C�digo da Filial = %i, N�mero = %l, Data Emiss�o = %dt. Deseja prosseguir na inser��o de nova Fatura com o mesmo n�mero?
Public Const AVISO_CONFIRMA_EXCLUSAO_PADRAO_COBRANCA = 5296 'Parametro iCodigo
'O Padr�o de Cobran�a %i ser� exclu�do. Confirma exclus�o?
Public Const AVISO_DIAS_PROTESTO = 5297 'Parametro sTipoInstrucao
'Instru��o %s n�o necessita de campo Dias para Devolu��o/Protesto
'que ser� zerado. Deseja prosseguir?
Public Const AVISO_CANCELAR_PAGAMENTO = 5298 'Sem parametro
'Confirma o cancelamento do pagamento?
Public Const AVISO_EXCLUSAO_MENSAGEM = 5299 'Sem par�metros
'Confirma a exclus�o da Mensagem?
Public Const AVISO_TITULO_PENDENTE_MESMO_NUMERO = 5300 'Par�metros: lFornecedor, iFilial, sSigla, lNumTitulo, dtDataEmissao
'No Banco de Dados existe T�tulo a Pagar Pendente com os dados C�digo do Fornecedor = %l, C�digo da Filial = %i, SiglaDocumento = %s, N�mero = %l, Data Emiss�o = %dt. Deseja prosseguir na inser��o de novo T�tulo a Pagar do mesmo Tipo e com o mesmo n�mero?
Public Const AVISO_TITULO_MESMO_NUMERO = 5301 'Par�metros: lFornecedor, iFilial, sSigla, lNumTitulo, dtDataEmissao
'No Banco de Dados existe T�tulo a Pagar com os dados C�digo do Fornecedor = %l, C�digo da Filial = %i, SiglaDocumento = %s, N�mero = %l, Data Emiss�o = %dt. Deseja prosseguir na inser��o de novo T�tulo a Pagar do mesmo Tipo e com o mesmo n�mero?
Public Const AVISO_TITULO_BAIXADO_MESMO_NUMERO = 5302 'Par�metros: lFornecedor, iFilial, sSigla, lNumTitulo, dtDataEmissao
'No Banco de Dados existe T�tulo a Pagar Baixado com os dados C�digo do Fornecedor = %l, C�digo da Filial = %i, SiglaDocumento = %s, N�mero = %l, Data Emiss�o = %dt. Deseja prosseguir na inser��o de novo T�tulo a Pagar do mesmo Tipo e com o mesmo n�mero?
Public Const AVISO_EXCLUSAO_TITULO = 5303 'Par�metro: lNumTitulo
'Confirma a exclus�o do T�tulo n�mero %l e Parcelas associadas ?
Public Const AVISO_CONFIRMA_EXCLUSAO_OCORRENCIA = 5304 'Sem par�metros
'A Ocorr�ncia ser� exclu�da da tabela OcorrenciasRemParcRec. Confirma exclus�o?
Public Const AVISO_CRIAR_PORTADOR = 5305 'Par�metro: iCodigo
'O Portador %i n�o existe, deseja cri�-lo?
Public Const AVISO_CRIAR_FILIALCLIENTE1 = 5306 'Par�metro: sFilial
'A Filial %s n�o existe na tabela de FiliaisClientes. Deseja Criar ?
Public Const AVISO_BORDERO_BD_DIFERENTE = 5307 'Par�metros: lNumBordero, iCobrador, iCodCarteiraCobranca, dtDataEmissao
'O Border� de N�mero %l est� com os dados Cobrador = %i, Carteira = %l, Data de Emiss�o = %dt diferentes da tela. Deseja prosseguir no cancelamento do bordero com esses dados?
Public Const AVISO_BATCH_CANCELADO = 5308
'Atualiza��o cancelada pelo usu�rio.
Public Const AVISO_IMPOSSIVEL_CONTINUAR = 5309
'O processo de atualiza��o ainda est� sendo efetuado.
Public Const AVISO_CRIAR_TIPOAPLICACAO = 5310 'Parametro: iCodigo
'TipoAplica��o com c�digo %i n�o est� cadastrada. Deseja criar?
Public Const AVISO_EXCLUSAO_PAIS = 5311 'Sem par�metros
'Confirma a exclus�o do Pais?
Public Const AVISO_HISTMOVCTA_INEXISTENTE = 5312 'Par�metro: iCodigo
'O Hist�rico de Movimenta��o de Conta %i n�o existe. Deseja Cri�-lo?
Public Const AVISO_EXCLUSAO_RELOPDEVEDORES = 5313 'Sem parametros
'Confirma a exclus�o da Op��o de Relat�rioDevedores ?
Public Const AVISO_EXCLUSAO_RELOPHISTAPLIC = 5314 'Sem parametros
'Confirma a exclus�o da Op��o de Relat�rioHistAplic?
Public Const AVISO_EXCLUSAO_RELOPMOVFIN = 5315 'Sem parametros
'Confirma a exclus�o da Op��o de Relat�rioMovFin ?
Public Const AVISO_EXCLUSAO_RELOPPOSAPLIC = 5316 'Sem parametros
'Confirma a exclus�o da Op��o de Relat�rioPosAplic?
Public Const AVISO_EXCLUSAO_RELOPCOMISVEND = 5317 'Sem parametros
'Confirma a exclus�o da Op��o de Relat�rio ?
Public Const AVISO_EXCLUSAO_RELOPTITRECMALA = 5318 'Sem parametros
'Confirma a exclus�o da Op��o de Relat�rioTitRecMala?
Public Const AVISO_NAO_IMP_CHEQUE = 5319 'sem parametro
'Prossegue mesmo sem ter impresso o cheque ?
Public Const AVISO_PAGTO_ATRASO_DESC = 5320 'parametro: # da linha
'Voc� digitou um valor de desconto para a parcela da linha %s que est� sendo paga em atraso. Deseja voltar para corrigir ?
Public Const AVISO_PAGTO_EM_DIA_MULTA = 5321 'parametro: # da linha
'Voc� digitou um valor de multa ou juros para a parcela da linha %s que est� sendo paga em dia. Deseja voltar para corrigir ?
Public Const AVISO_CHEQUE_NUM_USADO_DATA = 5322 'parametro: data do cheque
'Um cheque com o mesmo n�mero e data %s j� est� registrado no sistema. Deseja voltar para corrigir o n�mero ?
Public Const AVISO_ALTERACAO_VCTO_AFETA_DESC = 5323 'sem parametros
'Esta parcela cont�m descontos por antecipa��o de pagto. Se for necess�rio atualiz�-los poder� ser necess�rio dar baixa no t�tulo e reenvi�-lo ao banco com as informa��es atualizadas. Prossegue assim mesmo ?
Public Const AVISO_DATAVENCIMENTO_NAO_ORDENADA = 5324 'Sem parametros
'As Datas de Vencimento no Grid n�o est�o ordenadas. Deseja prosseguir assim mesmo ?
Public Const AVISO_INCLUSAO_ANTECIPPAG_ADICIONAL = 5325 'parametros: valor
'J� est�o registrados %s em adiantamentos para esta filial de fornecedor nesta data. Deseja registrar um novo adiantamento ?
Public Const AVISO_INCLUSAO_ANTECIPREC_ADICIONAL = 5326 'parametros: valor
'J� est�o registrados %s em adiantamentos desta filial de cliente nesta data. Deseja registrar um novo adiantamento ?
Public Const AVISO_CRIAR_BANCO = 5327 'Parametros: iCodVanco
'O Banco %i n�o existe, deseja cri�-lo?
Public Const AVISO_EXCLUSAO_CARTEIRA = 5328 'Par�metro: iCodCarteira
'Confirma a exclus�o da carteira %i?
Public Const AVISO_TITULOS_PAGAR_CAMPOS_ALTERAVEIS = 5329
'Todos os campos com exce��o das Datas de Vencimento , Tipo de Cobran�a, Suspenso, Banco Cobrador e Portador no Grid de Parcelas n�o ser�o alterados. Deseja proseguir na altera��o?
Public Const AVISO_TITULOS_PAGAR_VINCULADO_ADIANTAMENTO = 5330 'Par�metros: lCodFornecedor, iCodFilial
'O Titulo a pagar em quest�o est� vinculado a um Adiantamento do Fornecedor de c�digo %s e Filial de c�digo %s.
'Deseja prosseguir com a grava��o?
Public Const AVISO_TITULOS_PAGAR_VINCULADO_CREDITO = 5331 'Par�metros: lCodFornecedor, iCodFilial
'O Titulo a pagar em quest�o est� vinculado a um Cr�dito do Fornecedor de c�digo %s e Filial de c�digo %s.
'Deseja prosseguir com a grava��o?
Public Const AVISO_FORNECEDOR_CGC_IGUAL = 5332 'Parametro: sCGC
'J� existe um outro Fornecedor Cadastrado com o CGC %s, deseja continuar a Grava��o?




'!! A EXCLUIR QUANDO SUBSTITUIR FILIALCLIENTE_EXCLUI
Public Const ERRO_FILIALCLIENTE_REL_NFI = 5255
'Filial Cliente relacionada com Nota Fiscal Interna

