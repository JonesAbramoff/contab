Attribute VB_Name = "ErrosCPR"
Option Explicit

'Códigos de Erro - Reservado de 2000 até 2999; 16000 até 16499
Public Const ERRO_LEITURA_BANCOS = 2000 'Parametro iCodigo
' Erro na leitura do Banco %i.
Public Const ERRO_BANCO_NAO_CADASTRADO = 2001 'Parametro iCodigo
' Banco %i não está cadastrado.
Public Const ERRO_INSERCAO_BANCOS = 2002 'Parametro iCodigo
' Erro na inserção do Banco %i.
Public Const ERRO_ATUALIZACAO_BANCOS = 2003 'Parametro iCodigo
' Erro na atualização do Banco %i.
Public Const ERRO_EXCLUSAO_BANCOS = 2004 'Parametro iCodigo
' Erro na exclusão do Banco %i.
Public Const ERRO_LOCK_BANCOS = 2005 'Parametro iCodigo
' Não conseguiu fazer o lock do Banco %i.
Public Const ERRO_NOME_REDUZIDO_BANCO_REPETIDO = 2006 'Sem Parametros
' Já existe Banco com este nome reduzido.
Public Const ERRO_NOME_NAO_PREENCHIDO = 2007 'Sem parametros
' Preenchimento do Nome é obrigatório.
Public Const ERRO_NOME_REDUZIDO_NAO_PREENCHIDO = 2008 'Sem parametros
' Preenchimento do Nome Reduzido é obrigatório.
Public Const ERRO_LEITURA_CONTASCORRENTESINTERNAS = 2009 'Sem parametros
' Erro na leitura da tabela ContasCorrentesInternas.
Public Const ERRO_BANCO_RELACIONADO_COM_CONTASCORRENTESINTERNAS = 2010 'Sem parametros
' Não é possível excluir banco relacionado com ContasCorrentesInternas.
Public Const ERRO_INSERCAO_REGIOESVENDAS = 2011 'Parametro iCodigo
' Erro na inserção da Região de Venda %i.
Public Const ERRO_ATUALIZACAO_REGIOESVENDAS = 2012 'Parametro iCodigo
' Erro na atualização da Região de Venda %i.
Public Const ERRO_EXCLUSAO_REGIOESVENDAS = 2013 'Parametro iCodigo
' Erro na exclusão da Região de Venda %i.
Public Const ERRO_REGIAO_VENDA_RELACIONADA_COM_VENDEDOR = 2014 'Sem parametros
' Não é possível excluir Região de Venda relacionada com Vendedor.
Public Const ERRO_REGIAO_VENDA_RELACIONADA_COM_FILIAIS_CLIENTES = 2015 'Sem parametros
' Não é possível excluir Região de Venda relacionada com Filiais Clientes.
Public Const ERRO_REGIAO_VENDA_RELACIONADA_COM_TIPOSDECLIENTE = 2016 'Sem parametros
' Não é possível excluir Região de Venda relacionada com Tipos de Cliente.
Public Const ERRO_DESCRICAO_REPETIDA = 2017 'Sem parametros
' Já existe Região de Venda com esta descrição.
Public Const ERRO_LEITURA_TIPOSDECLIENTE = 2019 'Sem parametros
'Erro na leitura da tabela TipoDeCliente
Public Const ERRO_PAIS_NAO_PREENCHIDO = 2020 'Sem parametros
' Preenchimento do País é obrigatório.
Public Const ERRO_TAMANHO_CGC_CPF = 2022
'O tamanho do campo CGC tem que ser 11 caracteres(CPF) ou 14(CGC).
Public Const ERRO_CODCLIENTE_NAO_PREENCHIDO = 2023
'O código do Cliente não foi preenchido.
Public Const ERRO_RAZ_SOC_NAO_PREENCHIDA = 2024
'O Nome não foi preenchido.
Public Const ERRO_CREDITO_NEGATIVO = 2025
'O valor do Limite de Credito tem que ser positivo.
Public Const ERRO_FORNECEDOR_SEM_FILIAL = 2027 'Parametro codigo do fornecedor
'O Fornecedor %l não está vinculado a nenhuma filial.
Public Const ERRO_FORNECEDOR_INEXISTENTE = 2028
'O Fornecedor não está cadastrado.
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
'O Forncedor não pode ser excluído pois se encontra relacionado com Nota Fiscal Externa.
Public Const ERRO_FORNECEDOR_REL_TITULOS_PAGAR = 2045
'O Forncedor não pode ser excluído pois se encontra relacionado com Títulos a Pagar.
Public Const ERRO_FORNECEDOR_REL_CREDITOS_PAGAR = 2046
'O Forncedor não pode ser excluído pois se encontra relacionado com Creditos a Pagar.
Public Const ERRO_CODFORNECEDOR_NAO_PREENCHIDO = 2048
'O Código do Fornecedor deve ser preenchido.
Public Const ERRO_FILIALFORNECEDOR_GRAVA_MATRIZ = 2049
'A gravação da Matriz deve ser feita pela tela de Fornecedores.
Public Const ERRO_FILIAL_FORNECEDOR_EXCLUSAO_MATRIZ = 2050
'A exclusão da Matriz do Fornecedor deve ser feita pela tela de Fornecedores.
Public Const ERRO_FILIALFORNECEDOR_REL_NFE = 2051
'A Filial Fornecedor não pode ser excluída pois se encontra relacionada com Nota Fiscal Externa.
Public Const ERRO_FILIALFORNECEDOR_REL_TIT_PAGAR = 2052
'A Filial Fornecedor não pode ser excluída pois se encontra relacionada com Títulos a Pagar.
Public Const ERRO_FILIALFORNECEDOR_REL_CREDITOS = 2053
'A Filial Fornecedor não pode ser excluída pois se encontra relacionada com Créditos a Pagar.
Public Const ERRO_FILIALFORNECEDOR_NOME_DUPLICADO = 2054
'O Nome %s da Filial já esta sendo utilizado por uma outra.
Public Const ERRO_LEITURA_FAVORECIDOS1 = 2055 'Sem Parametros
'Erro na leitura da tabela de Favorecidos.
Public Const ERRO_FAVORECIDO_NAO_CADASTRADO = 2056 'Parametro Favorecido
'O Favorecido %s não está cadastrado.
Public Const ERRO_CODIGO_FAVORECIDO_NAO_INFORMADO = 2057 'Sem parametro
'O Código do Favorecido não foi informado.
Public Const ERRO_NOME_FAVORECIDO_NAO_INFORMADO = 2058 'Sem parametro
'Nome do Favorecido não foi informado.
Public Const ERRO_LEITURA_FAVORECIDO = 2059 'Parametro Codigo
'Erro na leitura do Favorecidos - Codigo = %s.
Public Const ERRO_ATUALIZACAO_FAVORECIDOS = 2060 'Parametro Favorecido
'Erro de atualização do Favorecido %i.
Public Const ERRO_INSERCAO_FAVORECIDOS = 2061 'Parametro Favorecido
'Erro na inserção do Favorecido %s na tabela Favorecidos.
Public Const ERRO_ATUALIZACAO_CCIMOVDIA = 2062
'Erro de atualização da tabela de CCIMovDia.
Public Const ERRO_ATUALIZACAO_CCIMOV = 2063
'Erro de atualização da tabela de CCIMov.
Public Const ERRO_ATUALIZACAO_MOVIMENTOSCONTACORRENTE = 2064
'Erro de Atualizacao da Tabela de MovimentosContaCorrente
Public Const ERRO_ATUALIZACAO_CONTASCORRENTESINTERNAS = 2065
'Erro na Atualizacao da Tabela ContasCorrentesInterna
Public Const ERRO_INSERCAO_CCIMOV = 2066
'Erro na Insercao de Registro na Tabela de Saldos Mensais de Conta Corrente.
Public Const ERRO_INSERCAO_CCIMOVDIA = 2067 ''ParaMetros: dtData , icodigo
'Erro na tentativa de insercao de Saldo Diário de Conta Corrente. Dia = %s e Conta = %s.
Public Const ERRO_INSERCAO_MOVIMENTOSCONTACORRENTE = 2068 'Parametro: iCodContaCorrente
'Erro na tentativa de inclusão de Movimento envolvendo a Conta Corrente %i.
Public Const ERRO_LEITURA_CCIMOV = 2069 'Parametros iCodconta ,   iAno
'Erro na leitura da tabela de Saldos Mensais de Conta Corrente. Conta = %s e Ano = %s.
Public Const ERRO_LEITURA_CCIMOVDIA = 2070 'ParaMetros: dtData , icodigo
'Erro de Leitura na tabela de Saldos Diários de Conta Corrente. Data = %s e Conta = %s.
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
'A Conta Corrente %i nao está cadastrada.
Public Const ERRO_TIPOMEIOPAGTO_INATIVO = 2080 'Parametro: iTipoMeioPagto
'O Tipo de Pagamento %i está inativo.
Public Const ERRO_TIPO_NAO_SAQUE = 2081 'Parametro: lSequencial
'O movimento %l não é do tipo Saque.
Public Const ERRO_FAVORECIDO_INATIVO = 2082 'Parametro: Codigo do Favorecido
'O Favorecido %i não está ativo.
Public Const ERRO_FAVORECIDO_INEXISTENTE = 2083 'Parametro Codigo do Favorecido
'O favorecido %i nao esta cadastrado
Public Const ERRO_DATA_SEM_PREENCHIMENTO = 2084
'A data deve estar preenchida
Public Const ERRO_MOVCONTACORRENTE_EXCLUIDO = 2085 'Parametro: Codconta, sequencial
'O movimento com a conta %i e sequencial %l esta excluido.
'Esse movimento está excluido
Public Const ERRO_CONTACORRENTE_NAO_INFORMADA = 2086
'A conta Corrente deve ser informada
Public Const ERRO_VALOR_NEGATIVO = 2087
'O valor do saque deve ser positivo
Public Const ERRO_SEQUENCIAL_NAO_PREENCHIDO = 2088
'O Sequencial deve estar preenchido.
Public Const ERRO_TAMANHO_HISTORICOMOVCONTA = 2089
'O Historico deve ter no máximo 50 Caracteres
Public Const ERRO_NUMMOVTO_INEXISTENTE = 2090 'Parametro: lNumMovto
'O Movimento <paramentro> nao esta cadastrado
Public Const ERRO_MOVCONTACORRENTE_INEXISTENTE = 2091 'Parametros: CodContaCorrente, Sequencial
'Não há movimento cadastrado com a conta corrente %i e o sequencial %l.
Public Const ERRO_TIPOMEIOPAGTO_INEXISTENTE = 2092 'Parametro: iTipoMeioPagto
'O Tipo de Pagmento %i nao está cadastrado.
Public Const ERRO_DATAMOVIMENTO_MENOR = 2093 'Parametro: dtDataMovimento, dtDataInicialConta, iCodConta
'A data do Movimento %dt é menor que a data inicial %dt da Conta Corrente %i.
Public Const ERRO_CCIMOV_NAO_CADASTRADO = 2094 'Parametros: iCodigo, iAno
'Nao ha movimentos cadastrados para a conta <iCodigo> no Ano de <iAno>
Public Const ERRO_CCIMOVDIA_NAO_CADASTRADO = 2095 'Parametros dtdata, icodigo
'Nao ha movimentos cadastrados no dia <data> para a conta <iCodigo>
Public Const ERRO_LOCK_CCIMOVDIA = 2096
'Erro na tentativa de fazer "lock" em registro da tabela CCIMovDia
Public Const ERRO_LOCK_CCIMOV = 2097
'Erro na tentativa de fazer "lock" em registro da tabela CCIMov
Public Const ERRO_TIPOMEIOPAGTO_NAO_INFORMADO = 2098
'A forma de pagamento não foi informada.
Public Const ERRO_TIPOCLIENTE_COD_NAO_PREENCHIDO = 2100
'O Código do Tipo de Cliente deve ser preenchido.
Public Const ERRO_TIPOCLIENTE_DESCR_NAO_PREENCHIDA = 2101
'A Descrição do Tipo de Cliente deve ser preenchida.
Public Const ERRO_INSERCAO_TIPOCLIENTE = 2103
'Erro na tentativa de cadastro na tabela TiposDeClientes no Banco De Dados.
Public Const ERRO_MODIFICACAO_TIPOCLIENTE = 2104
'Erro na tentativa de modificar a tabela TiposDeClientes no Banco De Dados.
Public Const ERRO_FILIAL_DESASSOCIADA_FORNECEDOR = 2108 'Parametro: sCGC
'Filial de Fornecedor com CGC %s desassociada de Fornecedor.
Public Const ERRO_TIPOMOV_DIFERENTE = 2109  'Sem Parametro
'O Tipo do Movimento não coincide com o cadastrado.
Public Const ERRO_LEITURA_MOVIMENTOSCONTACORRENTE1 = 2110  'Parametros: iConta e lSequencial
'Erro na leitura da Conta %i e Sequencial %l da tabela MovimentosContaCorrente.
Public Const ERRO_DATA_COM_MOVIMENTOS = 2111 'Parametro: Data Inicial
'A data de Saldo Inicial: <dtdataInicial>, é menor que a data
'de alguns movimento para a conta em questao
Public Const ERRO_EXCLUSAO_CONTASCORRENTESINTERNAS = 2112
'Erro na exlusao de registro da Tabela de Contas Correntes Internas
Public Const ERRO_CHEQUEBORDERO_DIFERENTE_ZERO = 2113 'Parametro: icodconta
'A conta nao pode ser excluída pois está
'sendo usada para emissão de cheques ou bordero
Public Const ERRO_AGENCIA_NAO_PREENCHIDA = 2114
'A Agência deve estar preenchida
Public Const ERRO_CODBANCO_NAO_INFORMADO = 2115
'O codigo do Banco deve ser informado
Public Const ERRO_NUMCONTA_NAO_PREENCHIDO = 2116
'O Número da conta deve estar preenchido
Public Const ERRO_INSERCAO_CONTASCORRENTESINTERNAS = 2117 'Parametro: Codigo
'Ocorreu um erro na tentativa de insercao da conta <codconta> na Tabela contascorrentesinternas
Public Const ERRO_MOVIMENTO_NAO_CADASTRADO = 2118 'Sem Parametro
'Movimento não cadastrado.
Public Const ERRO_LEITURA_EXTRATO_BCO = 2119 'Sem Parametro
'Erro na leitura da tabela de extratos bancários
Public Const ERRO_ALTERACAO_EXTRATO_BCO = 2120 'Sem Parametro
'Erro na alteracao da tabela de extratos bancários
Public Const ERRO_INSERCAO_EXTRATO_BCO = 2121 'Sem Parametro
'Erro na inserção de registro na tabela de extratos bancários
Public Const ERRO_LEITURA_LCTO_EXTRATO_BCO = 2122 'Sem Parametro
'Erro na leitura da tabela de lançamentos de extratos bancários
Public Const ERRO_INSERCAO_LCTO_EXTRATO_BCO = 2123 'Sem Parametro
'Erro inserção de registro na tabela de lançamentos de extratos bancários
Public Const ERRO_LEITURA_TABELA_PRECO = 2124 'Sem parametro
'Erro na leitura da tabela de Tabelas de Preço.
Public Const ERRO_TABELA_PRECO_NAO_ENCONTRADA = 2125 'Parametro sTabelaPreco
'Tabela de Preço com descrição %s não foi encontrada.
Public Const ERRO_MENSAGEM_NAO_ENCONTRADA = 2126 'Parametro sMensagem
'A Mensagem %s não foi encontrada.
Public Const ERRO_COBRADOR_NAO_ENCONTRADO = 2127 'Parametro sCobrador
'O Cobrador %s não foi encontrado.
Public Const ERRO_TRANSPORTADORA_NAO_ENCONTRADA = 2128 'Parâmetro: sNomeReduzido
'A Transportadora %s não foi encontrada.
Public Const ERRO_TIPOMEIOPAGTO_EXIGENUMERO = 2129 'Parametro : tipoMeiopagto
'O Tipo de Pagamento %i exige o preenchimento do campo numero.
Public Const ERRO_LEITURA_TABELA_HISTMOVCTA = 2130 'Sem parâmetro
'Erro na leitura da tabela HistPadraoMovConta.
Public Const ERRO_LEITURA_TABELA_HISTMOVCTA1 = 2131 'Parametro Codigo do HistPadrao.
'Erro na leitura da tabela HistPadraoMovConta. Histórico = %i.
Public Const ERRO_ATUALIZACAO_HISTMOVCTA = 2132 'Parametro Codigo do Historico
'Erro de atualização do Histórico de Movimentação de Conta %i.
Public Const ERRO_HISTMOVCTA_NAO_CADASTRADO = 2133 'Parametro Codigo do Histórico
'O Histórico de Movimentação de Conta %i não está cadastrado.
Public Const ERRO_LOCK_HISTMOVCTA = 2134 'Parametro Codigo do Histórico
'Não conseguiu fazer o lock do Histórico de Movimentação de Conta %i.
Public Const ERRO_EXCLUSAO_HISTMOVCTA = 2135 'Parametro Codigo do Histórico
'Houve um erro na exclusão do Histórico de Movimentação de Conta %i.
Public Const ERRO_INSERCAO_HISTMOVCTA = 2136 'Parâmetro Código do HistPadrão
'Erro na inserção do Histórico Padrão %i na tabela HistPadraoMovConta.
Public Const ERRO_LEITURA_CONTASCORRENTESINTERNAS1 = 2137 'Parametro: iCodConta
'Erro na tentativa de leitura da Conta Corrente %i na tabela ContasCorrentesInternas.
Public Const ERRO_LEITURA_TIPOMEIOPAGTO1 = 2138 'Parametro iTipoMeioPagto
'Erro na leitura do Tipo de Pagamento %i da tabela TipoMeioPagto.
Public Const ERRO_CCI_IGUAIS = 2139
'As contas de Origem e Destino são iguais
Public Const ERRO_CONTADESTINO_NAO_DIGITADA = 2140
'O preenchimento daConta Destino é obrigatorio
Public Const ERRO_TIPOMEIOPAGTO_JA_UTILIZADO = 2141 ' Parametros: iCodConta, iTipoMeioPagto, lNumero
'A Conta %i já utilizou a Forma de Pagamento %i de Numero %l em outro movimento.
Public Const ERRO_LEITURA_APLICACOES = 2142 'Parâmetro Código do Tipo de aplicação
'Erro na leitura da tabela de Aplicações. Tipo de aplicação = %s.
Public Const ERRO_LEITURA_TIPOSDEAPLICACAO = 2143 'Parâmetro Código do Tipo de Aplicação
'Erro na leitura da tabela de Tipos de Aplicação. Aplicação = %s.
Public Const ERRO_ATUALIZACAO_TIPOAPLICACAO = 2144 'Parâmetro Código do Tipo de Aplicação
'Erro na atualização do Tipo de aplicação %s.
Public Const ERRO_INSERCAO_TIPOAPLICACAO = 2145 'Parâmetro Código do Tipo de Aplicação
'Ocorreu um erro ao tentar inserir o Tipo de aplicaçao %s na tabela.
Public Const ERRO_EXCLUSAO_TIPOAPLICACAO = 2146 'Parâmetro Código do Tipo de Aplicação
'Ocorreu um erro ao tentar excluir o Tipo de aplicação %s da tabela.
Public Const ERRO_TIPOAPLICACAO_INEXISTENTE = 2147 'Parâmetro Código do Tipo de Aplicação
'O Tipo de aplicação %s não está cadastrado.
Public Const ERRO_LOCK_TIPOAPLICACAO = 2148 'Parâmetro Código do Tipo de Aplicação
'Ocorreu um erro ao tentar locar o Tipo de aplicação %s.
Public Const ERRO_CONTASIGUAIS = 2149 'Parâmetro Conta
'As Contas Contábil Aplicação e Receita Financeira são iguais : %s.
Public Const ERRO_EXCLUSAO_TIPOAPLICACAO_RELACIONADA = 2150 'Parâmetro Descrição do Tipo de aplicação
'O Tipo de aplicação %s não pode ser excluído pois existem uma ou mais aplicações deste tipo.
Public Const ERRO_SEQUENCIAL_NAO_INFORMADO = 2151
'O Sequencial não foi informado
Public Const ERRO_TIPO_MOVIMENTO_NAO_DEPOSITO = 2152 'Parametros: lsequencial e icodconta
'O Movimento %l na conta %s não está cadastrado como depósito.
Public Const ERRO_VALOR_INICIAL_MAIOR = 2153 'Sem Parametros
'O valor inicial não pode ser maior que o valor final.
Public Const ERRO_MOVIMENTOS_INEXISTENTES_CONCILIACAO = 2154 'Sem Parametros
'Não existem movimentos de conta corrente para a seleção atual.
Public Const ERRO_LEITURA_LCTOSEXTRATOBANCARIO = 2155 'Sem Parametros
'Erro na leitura da tabela de Lançamentos de Extrato Bancário.
Public Const ERRO_LANCEXTRATO_INEXISTENTES_CONCILIACAO = 2156 'Sem Parametros
'Não existem lançamentos de extrato de conta corrente para a seleção atual.
Public Const ERRO_GRIDS_MAIS_UM_ELEMENTO_SELECIONADO = 2157 'Sem parametros
'Ambos os grids contém mais de um elemento marcado. Pelo menos um dos grids só pode ter um elemento selecionado.
Public Const ERRO_GRID_EXTRATO_SEM_SELECAO = 2158 'Sem parametros
'Nenhum lançamento de extrato foi selecionado no grid.
Public Const ERRO_GRID_MOV_SEM_SELECAO = 2159 'Sem parametros
'Nenhum movimento de conta corrente foi selecionado no grid.
Public Const ERRO_LEITURA_LCTOSEXTRATOBANCARIO1 = 2160 'Parametros CodConta, NumExtrato, SeqLcto
'Erro na leitura da tabela de Lançamentos de Extrato Bancário. Conta Corrente = %i, Extrato = %i, Sequencial = %l.
Public Const ERRO_ATUALIZACAO_LCTOSEXTRATOBANCARIO = 2161 'Parametros CodConta, NumExtrato, SeqLcto
'Erro na atualização da tabela de Lançamentos de Extrato Bancário. Conta Corrente = %i, Extrato = %i, Sequencial = %l.
Public Const ERRO_LEITURA_CONCILIACAOBANCARIA = 2162 'Parametros CodConta, Sequencial do Movto, NumExtrato, Sequencial no Extrato
'Erro na leitura da tabela de Conciliação Bancária. Conta Corrente = %i, Sequencial do Movimento = %l, Extrato = %i, Sequencial no Extrato = %l.
Public Const ERRO_ATUALIZACAO_CONCILIACAOBANCARIA = 2163 'Parametros CodConta, Sequencial do Movto, NumExtrato, Sequencial no Extrato
'Erro na atualização da tabela de Conciliação Bancária. Conta Corrente = %i, Sequencial do Movimento = %l, Extrato = %i, Sequencial no Extrato = %l.
Public Const ERRO_ATUALIZACAO_MOVIMENTOSCONTACORRENTE1 = 2164 'Parametros CodConta, Sequencial
'Erro de Atualizacao da Tabela de MovimentosContaCorrente. Conta Corrente = %i, Sequencial = %l.
Public Const ERRO_INSERCAO_CONCILIACAOBANCARIA = 2165 'Parametros CodConta, Sequencial do Movto, NumExtrato, Sequencial no Extrato
'Erro na inserção na tabela de Conciliação Bancária. Conta Corrente = %i, Sequencial do Movimento = %l, Extrato = %i, Sequencial no Extrato = %l.
Public Const ERRO_GRID_EXT_JA_CONCILIADO = 2166 'Parametro: Posicao no Grid
'O Lançamento de Extrato localizado na linha %i do grid já está conciliado.
Public Const ERRO_GRID_MOV_JA_CONCILIADO = 2167 'Parametro: Posicao no Grid
'O Movimento de Conta Corrente localizado na linha %i do grid já está conciliado.
Public Const ERRO_GRID_EXTRATO_MOV_SEM_SELECAO = 2168 'Sem parametros
'Nenhum lançamento de extrato ou movimento de conta corrente foi selecionado no grid.
Public Const ERRO_LOCK_LCTOSEXTRATOBANCARIO = 2169 'Parametros CodConta, NumExtrato, SeqLcto
'Erro no lock de um registro da tabela de Lançamentos de Extrato Bancário. Conta Corrente = %i, Extrato = %i, Sequencial = %l.
Public Const ERRO_LOCK_MOVIMENTOSCONTACORRENTE1 = 2170 'Parametros CodConta, Sequencial
'Erro na tentativa de fazer "lock" na tabela MOvimentosContasCorrente. Conta Corrente = %i e Sequencial = %l.
Public Const ERRO_LEITURA_CONCILIACAOBANCARIA1 = 2171 'Parametros CodConta, NumExtrato, Sequencial no Extrato
'Erro na leitura da tabela de Conciliação Bancária. Conta Corrente = %i, Extrato = %i, Sequencial no Extrato = %l.
Public Const ERRO_LOCK_CONCILIACAOBANCARIA = 2172 'Parametros CodConta, Sequencial Mov, NumExtrato, Sequencial no Extrato
'Erro na tentativa de fazer "lock" na tabela de Conciliação Bancária. Conta Corrente = %i, Sequencial do Movimento = %l, Extrato = %i, Sequencial no Extrato = %l.
Public Const ERRO_EXCLUSAO_CONCILIACAOBANCARIA = 2173 'Parametros CodConta, Sequencial Mov, NumExtrato, Sequencial no Extrato
'Erro na exclusão de um registro da tabela de Conciliação Bancária. Conta Corrente = %s, Sequencial do Movimento = %l, Extrato = %i, Sequencial no Extrato = %l.
Public Const ERRO_LEITURA_CONCILIACAOBANCARIA2 = 2174 'Parametros CodConta, Sequencial do Movto
'Erro na leitura da tabela de Conciliação Bancária. Conta Corrente = %s, Sequencial do Movimento = %s.
Public Const ERRO_GRID_EXT_SEM_SELECAO = 2175 'Sem parametros
'Nenhum lançamento de extrato foi selecionado no grid.
Public Const ERRO_PESQUISA_GRID_MOV_DATA = 2176 'Parametro Data
'Não foi encontrado nenhum movimento no grid com a data = %s.
Public Const ERRO_EXT_SEM_MOV_CONCILIADO = 2177 'Parametros CodConta, NumExtrato, Sequencial no Extrato
'O Extrato em questão não está associado (conciliado) com nenhum movimento de conta corrente. Conta Corrente = %s, Extrato = %i, Sequencial no Extrato = %l.
Public Const ERRO_MOV_SEM_EXT_CONCILIADO = 2178 'Parametros CodConta, Sequencial
'O Movimento de Conta Corrente em questão não está associado (conciliado) com nenhum lançamento de extrato. Conta Corrente = %i e Sequencial = %l.
Public Const ERRO_PESQUISA_GRID_MOV_VALOR = 2179 'Parametro Valor
'Não foi encontrado nenhum movimento no grid com o valor = %s.
Public Const ERRO_PESQUISA_GRID_MOV_HISTORICO = 2180 'Parametro Historico
'Não foi encontrado nenhum movimento no grid com o histórico = %s.
Public Const ERRO_PESQUISA_GRID_EXT_DATA = 2181 'Parametro Data
'Não foi encontrado nenhum lançamento de extrato no grid com a data = %s.
Public Const ERRO_PESQUISA_GRID_EXT_VALOR = 2182 'Parametro Valor
'Não foi encontrado nenhum lançamento de extrato no grid com o valor = %s.
Public Const ERRO_PESQUISA_GRID_EXT_HISTORICO = 2183 'Parametro Historico
'Não foi encontrado nenhum lançamento de extrato no grid com o histórico = %s.
Public Const ERRO_LEITURA_PARCELAS_PAG = 2184
'A parcela deveria estar baixada ou excluida, pois foi verificado que o saldo do título é zero.
Public Const ERRO_SEM_PARCELAS_PAG_SEL = 2185
'Ao menos uma Parcela tem que ser selecionada para pagamento.
Public Const ERRO_INSERCAO_TITULOS_PAG = 2186
'Erro na inserção na tabela de Títulos a Pagar.
Public Const ERRO_INSERCAO_PARCELAS_PAG = 2187
'Erro na inserção na tabela de Parcelas a Pagar.
Public Const ERRO_EXCLUSAO_NOTAS_FISCAIS_EXT = 2188
'Erro na exclusão de um registro da tabela de Notas Fiscais a Pagar.
Public Const ERRO_EXCLUSAO_TITULOS_PAGAR = 2189
'Erro na exclusão de um registro da tabela de Títulos a Pagar.
Public Const ERRO_INSERCAO_NOTAS_FISCAIS_EXT_BAIXADAS = 2190
'Erro na inserção na tabela de Notas Fiscais Baixadas a Pagar.
Public Const ERRO_INSERCAO_BAIXA_PARC_PAG = 2191
'Erro na inserção na tabela de Parcelas Baixadas a Pagar.
Public Const ERRO_UNLOCK_TITULOS_PAGAR = 2192
'Erro na tentativa de desfazer o "lock" na tabela de Títulos a Pagar.
Public Const ERRO_UNLOCK_PARCELAS_PAGAR = 2193
'Erro na tentativa de desfazer o "lock" na tabela de Parcelas a Pagar.
Public Const ERRO_MODIFICACAO_PARCELAS_PAGAR = 2194
'Erro de Atualizacao da Tabela de Parcelas a Pagar.
Public Const ERRO_EXCLUSAO_PARCELAS_PAGAR = 2195
'Erro na exclusão de um registro da tabela de Parcelas a Pagar.
Public Const ERRO_LOCK_TITULOS_PAGAR = 2196
'Erro na tentativa de fazer "lock" na tabela de Títulos a Pagar.
Public Const ERRO_LOCK_PARCELAS_PAGAR = 2197
'Erro na tentativa de fazer "lock" na tabela de Parcelas a Pagar.
Public Const ERRO_PARCELA_PAGAR_NAO_ABERTA = 2198
'Parcela tem que estar aberta para poder ser baixada
Public Const ERRO_LEITURA_PORTADOR = 2199 'Parametro iCodigo
'Erro na leitura do Portador %i.
Public Const ERRO_PORTADOR_NAO_CADASTRADO = 2200 'Sem parametros
'Portador não está cadastrado.
Public Const ERRO_NOME_REDUZIDO_PORTADOR_REPETIDO = 2201 'Sem parametros
'Nome Reduzido é atributo de outro Portador.
Public Const ERRO_INSERCAO_PORTADOR = 2202 'Parametro iCodigo
'Erro na inserção do Portador %s.
Public Const ERRO_ATUALIZACAO_PORTADOR = 2203 'Parametro iCodigo
'Erro na atualização do Portador %s.
Public Const ERRO_EXCLUSAO_GERACAO_CHEQUES = 2204
'Erro na exclusão de um registro da tabela de Geração de Cheques.
Public Const ERRO_INSERCAO_GERACAO_CHEQUES = 2205
'Erro na inserção na tabela de Geração de Cheques.
Public Const ERRO_LEITURA_PORTADORES = 2206
'Erro na leitura da tabela Portadores.
Public Const ERRO_LEITURA_BORDERO_PAG = 2207
'Erro na leitura da tabela de Bordero.
Public Const ERRO_MODIFICACAO_BORDERO_PAG = 2208
'Erro de Atualizacao da Tabela de Bordero.
Public Const ERRO_INSERCAO_BORDERO_PAG = 2209
'Erro na inserção na tabela de Bordero.
Public Const ERRO_LEITURA_PORTADOR2 = 2210 'Sem parametros
'Erro na leitura da tabela Portador
Public Const ERRO_INSERCAO_BAIXAS_PAG = 2211 'Parametro lNumIntBaixa
'Erro de inserção de registro na tabela BaixasPag. Número Interno = %l.
Public Const ERRO_LEITURA_BAIXAS_PAG = 2212 'Sem parametros
'Erro na leitura da tabela BaixasPag.
Public Const ERRO_FLUXO_DATAINI_MAIOR_DATAFIM = 2213 'Parametros DataFinal e DataBase
'A Data Final do fluxo de caixa = %s não pode ser menor do que a Data Base = %s.
Public Const ERRO_NOME_FLUXO_VAZIO = 2214
'O Nome do Fluxo não foi preenchido.
Public Const ERRO_DATAFINAL_FLUXO_VAZIO = 2215
'A Data Final do Fluxo de Caixa não foi preenchida.
Public Const ERRO_LEITURA_FLUXO = 2216 'Parametro Nome do Fluxo de Caixa
'Erro na leitura da tabela de Fluxo de Caixa. Fluxo = %s.
Public Const ERRO_ATUALIZACAO_FLUXO = 2217 'Parametro Nome do Fluxo de Caixa
'Erro na atualização da tabela de Fluxo de Caixa. Fluxo = %s.
Public Const ERRO_MOV_NUMPAGTO_DUPLICADO = 2218 'Parâmetros iCodconta, iTipomeioPagto, lNumero
'Já existe um documento com o mesmo número. Conta = %i, Tipo de pagamento = %i e Número = %l.
Public Const ERRO_INSERCAO_FLUXO = 2219 'Parametro Nome do Fluxo de Caixa
'Erro na inserção de um registro na tabela de Fluxo de Caixa. Fluxo = %s.
Public Const ERRO_CREDITOSPAGFORN_NUM_MAX = 2220 'Parametro Numero Maximo de Credito a Pagar que podem ser lidos
'Atenção! Somente as informações oriundas dos primeiros %l registros de Creditos a Pagar serão exibidas.
Public Const ERRO_PORTADOR_NAO_CADASTRADO1 = 2221 'Parametro iCodPortador
'Portador %i não está cadastrado.
Public Const ERRO_LOCK_PORTADOR = 2222 'Parametro iCodPortador
'Erro na tentativa de "lock" na tabela Portador, código=%i.
Public Const ERRO_PORTADOR_INATIVO = 2223 'Parametro iCodPortador
'O Portador %i está inativo.
Public Const ERRO_MOVIMENTO_NAO_CADASTRADO1 = 2224 'Parametros: iCodConta, iTipoMeioPagto, lNumero
'Movimento não cadastrado. Conta=%i, Tipo Meio Pagto=%i, Número=%l.
Public Const ERRO_MOVIMENTO_NAO_CADASTRADO2 = 2225 'Parametros: iCodConta, lSequencial
'Movimento não cadastrado. Conta=%i, Sequencial=%l.
Public Const ERRO_LEITURA_PAGTOANTECIPADO = 2226 'Parametro: lNumeroMovimento
'Erro na leitura da tabela PagtosAntecipados. NumMovto=%l.
Public Const ERRO_PAGTOANTECIPADO_NAO_CADASTRADO = 2227 'Parametro: lNumeroMovimento
'Pagamento Antecipado com NumMovto=%l não está cadastrado.
Public Const ERRO_LEITURA_NFSPAG = 2228
'Erro na leitura da tabela Notas Fiscais a Pagar.
Public Const ERRO_INSERCAO_FLUXOANALITICO = 2230 'Sem Parametro
'Erro na inserção de um registro na tabela de Fluxo de Caixa Analítico.
Public Const ERRO_BANCO_SEM_LAYOUT_CHEQUE = 2231 'Sem Parametro
'banco tem que ter layout de cheque definido
Public Const ERRO_FAVORECIDO_INEXISTENTE1 = 2232 'Parametro: sFavorecido
'O Favorecido %s não está cadastrado
Public Const ERRO_MODIFICACAO_PARCELAS_REC = 2233
'Erro de Atualizacao da Tabela de Parcelas a Receber.
Public Const ERRO_UNLOCK_TITULOS_REC = 2235
'Erro na tentativa de desfazer o "lock" na tabela de Títulos a Receber.
Public Const ERRO_UNLOCK_PARCELAS_REC = 2236
'Erro na tentativa de desfazer o "lock" na tabela de Parcelas a Receber.
Public Const ERRO_PARCELA_REC_NAO_ABERTA = 2237
'A parcela não está aberta, portanto não pode ser baixada.
Public Const ERRO_TITULO_REC_INEXISTENTE = 2238
'O Título não está cadastrado.
Public Const ERRO_PARCELA_REC_INEXISTENTE = 2239
'A Parcela não está cadastrada.
Public Const ERRO_MODIFICACAO_TITULOS_REC = 2240
'Erro de Atualizacao da Tabela de Títulos a Receber.
Public Const ERRO_INSERCAO_BAIXA_PARC_REC = 2241
'Erro na inserção de um registro na tabela de Parcelas Baixadas a Receber.
Public Const ERRO_EXCLUSAO_PARCELAS_RECEBER = 2242
'Erro na exclusão de um registro da tabela de Parcelas a Receber.
Public Const ERRO_EXCLUSAO_NOTAS_FISCAIS_REC = 2243
'Erro na exclusão de um registro da tabela de Notas Fiscais a Receber.
Public Const ERRO_INSERCAO_NOTAS_FISCAIS_REC_BAIXADAS = 2244
'Erro na inserção de um registro na tabela de Notas Fiscais Baixadas a Receber.
Public Const ERRO_LEITURA_PARCELAS_REC1 = 2245
'Erro na leitura da tabela de Parcelas a Receber e Titulos a Receber.
Public Const ERRO_LEITURA_APLICACOES1 = 2246 'Sem Parametros
'Erro na leitura da tabela de Aplicações.
Public Const ERRO_LEITURA_CCI_CCIMOV = 2247 'Sem parametros.
'Erro na tentativa de leitura das tabelas de Conta Corrente e Saldos Mensais de Conta Corrente.
Public Const ERRO_LEITURA_CCIMOVDIA1 = 2248 'Sem parametros.
'Erro de Leitura da Tabela de Saldos Diarios de Conta Corrente.
Public Const ERRO_INSERCAO_FLUXOTIPOAPLIC = 2249 'Sem Parametro
'Erro na inserção de um registro na tabela de Fluxo de Caixa de Tipos de Aplicação.
Public Const ERRO_INSERCAO_FLUXOAPLIC = 2250 'Sem Parametro
'Erro na inserção de um registro na tabela de Fluxo de Caixa de Aplicações.
Public Const ERRO_LEITURA_APLICACOES2 = 2251 'Parametro: lCodigo
'Erro na leitura da tabela de Aplicações com Código %l.
Public Const ERRO_INSERCAO_FLUXOSALDOSINICIAIS = 2252 'Sem Parametro
'Erro na inserção de um registro na tabela de Fluxo de Caixa Saldos Iniciais.
Public Const ERRO_LEITURA_FLUXOANALITICO = 2253 'Parametro: FluxoID
'Erro na leitura da tabela de Fluxo de Caixa Analitico. Fluxo = %l.
Public Const ERRO_INSERCAO_FLUXOFORN = 2254 'Sem Parametro
'Erro na inserção de um registro na tabela de Fluxo de Caixa de Fornecedor (FluxoForn).
Public Const ERRO_INSERCAO_FLUXOTIPOFORN = 2255 'Sem Parametro
'Erro na inserção de um registro na tabela de Fluxo de Caixa de Tipo de Fornecedor (FluxoTipoForn).
Public Const ERRO_INSERCAO_FLUXOSINTETICO = 2256 'Sem Parametro
'Erro na inserção de um registro na tabela de Fluxo de Caixa Sintetico.
Public Const ERRO_FLUXO_NAO_CADASTRADO = 2257 'Parametro: Nome do fluxo
'O Fluxo %s não está cadastrado.
Public Const ERRO_FLUXO_NAO_PREENCHIDO = 2258 'Sem parametro
'O Nome do Fluxo não foi informado.
Public Const ERRO_LEITURA_TITULOSARECEBER = 2259
'Erro na leitura da tabela de Títulos a Receber.
Public Const ERRO_LOCK_FLUXO = 2260 'Parametro Nome do Fluxo de Caixa
'Ocorreu um erro ao tentar fazer lock no Fluxo de Caixa %s.
Public Const ERRO_EXCLUSAO_FLUXO = 2261 'Parametro Nome do Fluxo de Caixa
'Erro na exclusão do Fluxo de Caixa %s.
Public Const ERRO_LEITURA_FLUXOFORN = 2262 'Parametro Nome do Fluxo de Caixa
'Erro na leitura da tabela de Fluxo de Caixa - FluxoForn. Fluxo = %s.
Public Const ERRO_EXCLUSAO_FLUXOFORN = 2263 'Parametro Nome do Fluxo de Caixa
'Erro na exclusão do Fluxo de Caixa (FluxoForn) %s.
Public Const ERRO_LEITURA_FLUXOTIPOFORN = 2264 'Parametro Nome do Fluxo de Caixa
'Erro na leitura da tabela de Fluxo de Caixa - FluxoTipoForn. Fluxo = %s.
Public Const ERRO_EXCLUSAO_FLUXOTIPOFORN = 2265 'Parametro Nome do Fluxo de Caixa
'Erro na exclusão do Fluxo de Caixa (FluxoTipoForn) %s.
Public Const ERRO_EXCLUSAO_FLUXOANALITICO = 2266 'Parametro Nome do Fluxo de Caixa
'Erro na exclusão do Fluxo Analítico de Caixa. Fluxo = %s.
Public Const ERRO_LEITURA_FLUXOAPLIC = 2267 'Parametro Nome do Fluxo de Caixa
'Erro na leitura da tabela de Aplicações de Fluxo de Caixa. Fluxo = %s.
Public Const ERRO_EXCLUSAO_FLUXOAPLIC = 2268 'Parametro Nome do Fluxo de Caixa
'Erro na exclusão de uma Aplicação do Fluxo de Caixa %s.
Public Const ERRO_LEITURA_FLUXOTIPOAPLIC = 2269 'Parametro Nome do Fluxo de Caixa
'Erro na leitura da tabela de Tipos de Aplicação do Fluxo de Caixa. Fluxo = %s.
Public Const ERRO_EXCLUSAO_FLUXOTIPOAPLIC = 2270 'Parametro Nome do Fluxo de Caixa
'Erro na exclusão de um Tipo de Aplicação do Fluxo de Caixa  %s.
Public Const ERRO_LEITURA_FLUXOSALDOSINICIAIS = 2271 'Parametro Nome do Fluxo de Caixa
'Erro na leitura da tabela de Saldos Iniciais do Fluxo de Caixa. Fluxo = %s.
Public Const ERRO_EXCLUSAO_FLUXOSALDOSINICIAIS = 2272 'Parametro Nome do Fluxo de Caixa
'Erro na exclusão dos Saldos Iniciais do Fluxo de Caixa %s.
Public Const ERRO_LEITURA_FLUXOSINTETICO = 2273 'Parametro Nome do Fluxo de Caixa
'Erro na leitura da tabela de Fluxos Sintéticos. Fluxo = %s.
Public Const ERRO_EXCLUSAO_FLUXOSINTETICO = 2274 'Parametro Nome do Fluxo de Caixa
'Erro na exclusão de um Fluxo Sintético. Fluxo = %s.
Public Const ERRO_LEITURA_FLUXOANALITICO1 = 2275 'Parametro: FluxoID, Data, TipoReg
'Erro na leitura da tabela de Fluxo de Caixa Analitico. Fluxo = %l, Data = %s e Tipo do Registro = %i.
Public Const ERRO_TIPO_NAO_ANTECIPPAG = 2276 'Sem parâmetros
'O Movimento não é do tipo Pagamento Antecipado.
Public Const ERRO_LEITURA_ANTECIPPAG = 2277 'Parâmetro lNumIntPag
'Erro na leitura do Pagamento antecipado %l.
Public Const ERRO_LEITURA_ANTECIPPAG1 = 2278 'Parâmetro lNumIntMovto
'Erro na leitura do Pagamento antecipado cujo Número de movimento é: %l.
Public Const ERRO_INSERCAO_ANTECIPPAG = 2279 'Parâmetro: lNumIntPag
'Ocorreu um erro na tentativa de inserção do Pagamento antecipado %l na Tabela PagtosAntecipados.
Public Const ERRO_EXCLUSAO_ANTECIPPAG = 2280 'Parâmetro: lNumIntPag
'Ocorreu um erro na tentativa de exclusão do Pagamento antecipado %l.
Public Const ERRO_LOCK_ANTECIPPAG = 2281 'Parâmetro lNumIntPag
'Erro na tentativa de fazer o "lock" do Pagamento antecipado %l.
Public Const ERRO_FORNECEDOR_NAO_PREENCHIDO = 2282 'Sem parâmetro
'O Fornecedor deve ser preenchido.
Public Const ERRO_EXCLUSAO_MOVIMENTOSCONTACORRENTE1 = 2284 'Parâmetro iCodConta + lSequencial
'Erro na exclusao do Movimento. Conta = %i e Sequencial = %l.
Public Const ERRO_ANTECIPPAG_EXCLUIDO = 2285 'Parâmetros iCodConta + lSequencial
'O Pagamento antecipado com a Conta %i e Sequencial %l está excluído.
Public Const ERRO_PAGAMENTO_APROPRIADO = 2286 'Parâmetro iCodigo
'Não é possível excluir o Pagamento antecipado %i, pois já foi apropriado (total ou parcialmente).
Public Const ERRO_FILIAL_NAO_ENCONTRADA = 2287 'Parâmetro sFilial
'Filial com descrição %s não foi encontrada.
Public Const ERRO_FORNECEDOR_NAO_COINCIDE = 2288 'Parâmetros lFornecedor(da tela) e lFornecedor(da tabela)
'O Fornecedor %l não coincide com o Fornecedor %l cadastrado no Pagamento Antecipado
Public Const ERRO_FILIAL_NAO_COINCIDE = 2289 'Parâmetros lFilial(da tela) e lFilial(da tabela)
'A Filial %l não coincide com a Filial %l cadastrada no Pagamento Antecipado.
Public Const ERRO_ANTECIPPAG_INEXISTENTE = 2290 'Parâmetros iCodConta + lSequencial
'O Movimento com a Conta %i e Sequencial %l não está cadastrado.
Public Const ERRO_FILIALFORNECEDOR_INEXISTENTE = 2291 'Parametro: Filial
'A filial %s não está cadastrada.
Public Const ERRO_NUMERO_NAO_INFORMADO = 2293 'Parâmetro: iTipo
'Um número de documento tem que ser informado para o Tipo de pagamento %i.
Public Const ERRO_VALOR_MENOR_UM = 2295 'Parametro: dValor
'O Valor %d é menor do que 1.
Public Const ERRO_NFPAG_NAO_CADASTRADA = 2296 'Parametro: lNumIntDoc
'A Nota Fiscal com número interno %l não está cadastrada.
Public Const ERRO_NFPAG_NAO_CADASTRADA1 = 2297 'Parametro: lNumNotaFiscal
'A Nota Fiscal %l não está cadastrada.
Public Const ERRO_DATAVENCIMENTO_MENOR = 2298
'A Data de Vencimento é menor do que a Data de Emissão.
Public Const ERRO_LOCK_FILIALFORNECEDOR = 2301 'Parametros: lCodFornecedor, iCodFilial.
' Erro na tentativa de "lock" da tabela FiliaisFornecedores. CodFornecedor=%l, CodFilial=%i.
Public Const ERRO_NF_PENDENTE_MODIFICACAO = 2302 'parametro: lNumNotaFiscal
'Não é possível modificar a Nota Fiscal %l. Ela faz parte de um Lote Pendente.
Public Const ERRO_NF_PENDENTE_EXCLUSAO = 2303 'parametro: lNumNotaFiscal
'Não é possível excluir a Nota Fiscal %l. Ela faz parte de um Lote Pendente.
Public Const ERRO_LEITURA_NFSPAGPEND = 2304 'Parâmetro: lNumNotaFiscal
'Erro na tentativa de leitura da Nota Fiscal %l na tabela NfsPagPend
Public Const ERRO_NF_FILIALEMPRESA_DIFERENTE = 2305 'Parâmetro: lNumNotaFiscal
'Não é possível modificar a Nota Fiscal %l. Ela pertence a outra Filial da Empresa.
Public Const ERRO_LEITURA_NFSPAGBAIXADAS = 2307 'Parâmetro: lNumnotafiscal
'Erro na leitura da Nota Fiscal número %l na tabela NfsPagBaixadas.
Public Const ERRO_NF_BAIXADA_MODIFICACAO = 2308 'Parâmetro: lNumNotaFiscal
'Não é possível modificar a Nota Fiscal %l porque ela está baixada.
Public Const ERRO_NF_BAIXADA_EXCLUSAO = 2309 'Parâmetro: lNumNotaFiscal
'Não é possível excluir a Nota Fiscal %l porque ela está baixada.
Public Const ERRO_INSERCAO_NFSPAG = 2310 'Parâmetro: lNumNotaFiscal
'Erro na tentativa de inserir a Nota Fiscal número %l na tabela NfsPag.
Public Const ERRO_NF_NAO_INFORMADA = 2311
'O campo número da Nota Fiscal não foi preenchido.
Public Const ERRO_VALORTOTAL_NAO_INFORMADO = 2312
'O campo Valor não foi preenchido.
Public Const ERRO_VALORPRODUTOS_NAO_INFORMADO = 2313
'O campo Valor dos Produtos não foi preenchido.
Public Const ERRO_VALORTOTAL_INVALIDO = 2314 'Parametros: sValorTotal, dValor
'O Valor Total %s não é igual à soma dos valores de Frete, Produtos, ICMS Subst, Seguro, IPI, Outras Despesas que é de %d.
Public Const ERRO_DATAEMISSAO_MAIOR = 2315 'Parametros: sDataEmissao, sDataVencimento
'A Data de Emissão %s é maior que a Data de Vencimento %s.
Public Const ERRO_LEITURA_FILIAISFORNECEDORES2 = 2319 'Parametro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela FiliaisFornecedores. CodFornecedor = %l e CodFilial = %i.
Public Const ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA = 2320 'Parâmetros: sFilialNome
'Filial de Fornecedor %s não foi encontrada.
Public Const ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA = 2321 'Parametro: sNomeReduzido
'A Condição de Pagamento %s não foi encontrada.
Public Const ERRO_CONDICAO_PAGTO_NAO_CADASTRADA1 = 2322 'Parametro sDescReduzida
'A Condição de Pagamento %s não está cadastrada no Banco de Dados.
Public Const ERRO_REGIAO_VENDA_NAO_CADASTRADA1 = 2324 'Parametro: sDescricao
'Região de Venda %s não está cadastrada no Banco de Dados.
Public Const ERRO_TIPO_CLIENTE_NAO_PREENCHIDO = 2325 'Sem parametros
'Preenchimento do Tipo de Cliente é obrigatório.
Public Const ERRO_SEM_PARCELAS_REC_SEL = 2326
'Não achou nenhuma parcela a receber dentro dos critérios informados.
Public Const ERRO_INSERCAO_BORDERO_COBRANCA = 2327
'Erro na inserção de Bordero de Cobrança.
Public Const ERRO_INSERCAO_OCORR_REM_PARC_REC = 2328
'Erro na inserção de ocorrência de remessa de Bordero de Cobrança.
Public Const ERRO_ATUALIZACAO_CARTEIRAS_COBRADOR = 2329
'Erro na atualização da tabela Carteiras Cobrador.
Public Const ERRO_PARCELA_NAO_ABERTA = 2330
'A parcela tem que estar aberta.
Public Const ERRO_COBRADOR_JA_DEFINIDO = 2331
'O cobrador tem que estar em aberto.
Public Const ERRO_BORD_COBR_VENCTO_ALTERADO = 2332
'A data de vencimento foi alterada após a seleção das parcelas.
Public Const ERRO_BORD_COBR_VALOR_ALTERADO = 2333
'O valor a ser cobrado foi alterado após a seleção das parcelas.
Public Const ERRO_LEITURA_MOVSCCIFLUXO = 2334
'Erro na leitura do Fluxo de Movimento de Conta Corrente.
Public Const ERRO_LEITURA_PAGAMENTOS_PARA_FLUXO = 2335
'Erro na pesquisa de Pagamentos para consulta de Fluxo de Caixa.
Public Const ERRO_LEITURA_PAGTOANTEC_PARA_FLUXO = 2336
'Erro na pesquisa de Pagamentos antecipados para consulta de Fluxo de Caixa.
Public Const ERRO_LEITURA_RECEBTOANTEC_PARA_FLUXO = 2337
'Erro na pesquisa de Recebimentos antecipados para consulta de Fluxo de Caixa.
Public Const ERRO_LEITURA_CREDPAGFORN_PARA_FLUXO = 2338
'Erro na pesquisa de Créditos a Pagar de Fornecedores para consulta de Fluxo de Caixa.
Public Const ERRO_LEITURA_DEBRECCLI_PARA_FLUXO = 2339
'Erro na leitura de Fluxo para Débitos a Receber de Clientes.
Public Const ERRO_LEITURA_FLUXOTIPOFORN2 = 2340
'Erro na leitura da tabela de Fluxo de Tipo de Fornecedor.
Public Const ERRO_ATUALIZACAO_FLUXOTIPOFORN = 2341
'Erro na tentativa de atualização da tabela de Fluxo de Tipo de Fornecedor.
Public Const ERRO_LEITURA_FLUXOFORN2 = 2342
'Erro na leitura da tabela de Fluxo de Fornecedores.
Public Const ERRO_ATUALIZACAO_FLUXOFORN = 2343
'Erro na tentativa de atualização da tabela de Fluxo de Fornecedores.
Public Const ERRO_LEITURA_RECEBTOS_PARA_FLUXO = 2344
'Erro na pesquisa de Recebimentos para consulta de Fluxo de Caixa.
Public Const ERRO_ATUALIZACAO_FLUXOTIPOAPLIC = 2345
'Erro na tentativa de atualização da tabela de Fluxo de Tipo de Aplicação.
Public Const ERRO_LEITURA_RESGATESFLUXO = 2346
'Erro na leitura de Resgates de Fluxo.
Public Const ERRO_ATUALIZACAO_FLUXOSINTETICO = 2347
'Erro na tentativa de atualização da tabela de Fluxo Sintético.
Public Const ERRO_LEITURA_FLUXOFORN1 = 2348 'Parametros Nome do Fluxo de Caixa, Data, Tipo do Registro
'Erro na leitura da tabela de Fluxo de Caixa - FluxoForn. Fluxo = %s, Data=%s e Tipo do Registro = %i.
Public Const ERRO_TIPO_FORNECEDOR_NAO_CADASTRADO = 2349 'Parametro: iCodigo
'Tipo de Fornecedor com código %i não está cadastrado no BD.
Public Const ERRO_FLUXO_DATA_FORA_FAIXA = 2350 'Parametros Data, DataBase do Fluxo e DataFinal do Fluxo
'A Data em questão %s está fora da faixa abrangida pelo fluxo de caixa. Data Base = %s e Data Final = %s.
Public Const ERRO_PORTADOR_NAO_INFORMADO = 2351
'O Portador deve ser informado.
Public Const ERRO_PROXCHEQUE_NAO_INFORMADO = 2352
'O Próximo cheque deve ser informado.
Public Const ERRO_DATACONTABIL_MENOR_DATAEMISSAO = 2353
'A Data Contábil deve ser maior ou igual à Data de Emissão.
Public Const ERRO_CONTACORRENTE_NAO_BANCARIA = 2354
'A Conta Corrente deve ser Bancária.
Public Const ERRO_TITULO_NAO_PREENCHIDO = 2355 'Sem parâmetro
'O Título deve ser preenchido.
Public Const ERRO_PARCELA_NAO_PREENCHIDA = 2356 'Sem parâmetro
'A Parcela deve ser preenchida.
Public Const ERRO_FILIALCLIENTE_REL_NF_REC_PEND = 2357 'Parametros: lCodCliente, iCodFilial
'Erro na exclusão da Filial de Cliente com CodCliente=%l, CodFilial=%i. Está relacionada com Nota Fiscal a Receber Pendente.
Public Const ERRO_FILIALCLIENTE_REL_NF_REC = 2358 'Parametros: lCodCliente, iCodFilial
'Erro na exclusão da Filial de Cliente com CodCliente=%l, CodFilial=%i. Está relacionada com Nota Fiscal a Receber.
Public Const ERRO_FILIALCLIENTE_REL_NF_REC_BAIXADA = 2359 'Parametros: lCodCliente, iCodFilial
'Erro na exclusão da Filial de Cliente com CodCliente=%l, CodFilial=%i. Está relacionada com Nota Fiscal a Receber Baixada.
Public Const ERRO_LEITURA_FORNECEDORES_NOMEREDUZIDO = 2360 'Parametro Nome Reduzido
'Erro na leitura da tabela de Forncedores. Nome Reduzido = %s.
Public Const ERRO_LEITURA_FLUXO1 = 2361 'Parametro FluxoId
'Erro na leitura da tabela de Fluxo de Caixa. FluxoId = %l.
Public Const ERRO_FLUXO_NAO_CADASTRADO1 = 2362 'Parametro: FluxoId
'O Fluxo não está cadastrado. FluxoId = %l.
Public Const ERRO_LOCK_FLUXO1 = 2363 'Parametro FluxoId
'Ocorreu um erro ao tentar fazer lock no Fluxo de Caixa. FluxoId = %l.
Public Const ERRO_LEITURA_FLUXOFORN3 = 2364 'Parametros FluxoId, Tipo do Registro, Fornecedor, Data
'Erro na leitura da tabela de Fluxo de Caixa - FluxoForn. FluxoId = %l, Tipo do Registro = %i, Código do Fornecedor = %l e Data=%s
Public Const ERRO_EXCLUSAO_FLUXOFORN1 = 2365 'Parametros FluxoId, Tipo do Registro, Fornecedor, Data
'Erro na exclusão do Fluxo de Caixa (FluxoForn). FluxoId = %l, Tipo do Registro = %i, Código do Fornecedor = %l e Data=%s.
Public Const ERRO_LEITURA_FLUXOTIPOFORN1 = 2366 'Parametros FluxoId, Tipo do Registro, Tipo do Fornecedor, Data
'Erro na leitura da tabela de Fluxo de Caixa - FluxoTipoForn. FluxoId = %l, Tipo do Registro = %i, Tipo do Fornecedor = %i e Data=%s.
Public Const ERRO_ATUALIZACAO_FLUXOTIPOFORN1 = 2367 'Parametros FluxoId, Tipo do Registro, Tipo do Fornecedor, Data
'Erro na tentativa de atualização da tabela de Fluxo de Tipo de Fornecedor. FluxoId = %l, Tipo do Registro = %i, Tipo do Fornecedor = %i e Data=%s.
Public Const ERRO_ATUALIZACAO_FLUXOFORN1 = 2368 'Parametros FluxoId, Tipo do Registro, Fornecedor, Data
'Erro na tentativa de atualização da tabela de Fluxo de Caixa - FluxoForn. FluxoId = %l, Tipo do Registro = %i, Código do Fornecedor = %l e Data=%s
Public Const ERRO_LEITURA_TIPOSFORNECEDOR1 = 2369 'Parametro iTipoForn
'Erro na leitura da tabela TiposdeFornecedor. Tipo do Fornecedor = %i.
Public Const ERRO_LEITURA_FLUXOSINTETICO1 = 2370 'Parametros FluxoId e Data
'Erro na leitura da tabela de Fluxos Sintéticos. FluxoId = %l e Data = %s.
Public Const ERRO_ATUALIZACAO_FLUXOSINTETICO1 = 2371 'Parametros FluxoId e Data
'Erro na tentativa de atualização da tabela de Fluxo Sintético. FluxoId = %l e Data = %s.
Public Const ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO = 2372 'Parametro: iCodigo
'Condição de Pagamento com código %i não é de Contas a Pagar.
Public Const ERRO_LEITURA_TIPOSCLIENTE1 = 2373 'Parametro iTipoCli
'Erro na leitura da tabela TiposdeCliente. Tipo do Cliente = %i.
Public Const ERRO_LEITURA_CLIENTES_NOMEREDUZIDO = 2374 'Parametro: sNomeReduzido
'Erro na leitura da tabela de Clientes. Nome Reduzido = %s.
Public Const ERRO_CCI_INSERCAO_EMPRESA_TODA = 2376 'Sem parâmetro
'O usuário tem que ter selecionado uma filial ao se conectar ao sistema para poder incluir uma conta corrente
Public Const ERRO_CHEQUE_PRE_DIF_VALOR = 2377 'Sem parâmetro
'O cheque pré-datado tem valor diferente do necessário para pagar as parcelas associadas à ele
Public Const ERRO_INSERCAO_BAIXAS_REC = 2378 'Sem parâmetro
'Erro na inserção de baixa de título a receber
Public Const ERRO_INSERCAO_BORDERO_CHEQUE_PRE = 2379 'Sem parâmetro
'Erro na inserção de bordero de cheques pré-datados
Public Const ERRO_LEITURA_CHEQUES_PRE_BORDERO = 2380 'Sem parâmetro
'erro na leitura de cheques pré-datados para o borderô
Public Const ERRO_LEITURA_BORDERO_SEM_CHEQUES_PRE = 2381 'Sem parâmetro
'Não há cheque pré-datado a depositar até a data informada
Public Const ERRO_LEITURA_VENDEDOR1 = 2382 'Parametro: sNomeReduzido
'Erro na leitura do Vendedor com Nome Reduzido %s na tabela de Vendedores.
Public Const ERRO_SEM_COMISSOES_BAIXA = 2384    'Sem parâmetro
'Não há comissões a serem baixadas, verifique os parâmetros informados.
Public Const ERRO_GRAVACAO_BAIXA_COMISSAO = 2385    'Sem parâmetro
'Erro na gravação da baixa de uma comissão
Public Const ERRO_NFFAT_FILIALEMPRESA_DIFERENTE = 2386 'Parâmetro: lNumNotaFiscalFatura
'Não é possível modificar a Nota Fiscal Fatura %l. Ela pertence a outra Filial da Empresa.
Public Const ERRO_FATURA_FILIALEMPRESA_DIFERENTE = 2387 'Parâmetro: lNumFatura
'Não é possível modificar a Fatura %l. Ela pertence a outra Filial da Empresa.
Public Const ERRO_OUTROPAG_FILIALEMPRESA_DIFERENTE = 2388 'Parâmetros: sSiglaDocumento, lNumTitulo
'Não é possível modificar o Título a Pagar do tipo %s de número %l. Ele pertence a outra Filial da Empresa.
Public Const ERRO_NFFATPAG_NAO_CADASTRADA = 2389 'Parâmetro: lNumIntDoc
'A Nota Fiscal Fatura com Número Interno %l não está cadastrada
Public Const ERRO_NFFATPAG_NAO_CADASTRADA1 = 2390 'Parâmetro: lNumTitulo
'A Nota Fiscal Fatura número %l não está cadastrada
Public Const ERRO_TITULO_NAO_NFFATPAG = 2391 'Parâmetro: lNumIntDoc
'O Título de número interno %l não é Nota Fiscal Fatura.
Public Const ERRO_DATAVENCIMENTO_PARCELA_MENOR = 2392 'Parâmetros: sDataVencimento, sDataEmissao, iParcela
'A Data de Vencimento %s da Parcela %i é menor do que a Data de Emissão %s.
Public Const ERRO_AUSENCIA_PARCELAS_GRAVAR = 2393
'Não existem parcelas no Grid de Parcelas para gravar.
Public Const ERRO_DATAVENCIMENTO_PARCELA_NAO_INFORMADA = 2394 'Parâmetro: iParcela
'O campo Data de Vencimento da Parcela %i não foi preenchido.
Public Const ERRO_DATAVENCIMENTO_NAO_ORDENADA = 2395 'Sem parametros
'As Datas de Vencimento no Grid não estão ordenadas.
Public Const ERRO_SOMA_PARCELAS_INVALIDA = 2396 'Parâmetros: dSomaParcelas, dValorPagar
'A soma das Parcelas %d não é igual ao Valor a Pagar %d.
Public Const ERRO_NFFATPAG_BAIXADA_MODIFICACAO = 2397 'Parâmetro: lNumNotaFiscal
'Não é possível modificar Nota Fiscal Fatura número %l porque ela está baixada.
Public Const ERRO_NFFATPAG_PENDENTE_MODIFICACAO = 2398 'Parâmetro: lNumNotaFiscal
'Não é possível modificar Nota Fiscal Fatura número %l .Ela faz parte de um Lote Pendente.
Public Const ERRO_LEITURA_TITULOSPAG = 2399 'Parâmetro: lNumIntDoc
'Erro na tentativa de ler registro na tabela TitulosPag com Número Interno %l.
Public Const ERRO_LEITURA_NFFATURA = 2400 'Parâmetro: lNumTitulo
'Erro na tentativa de ler Nota Fiscal Fatura %l na tabela TitulosPag.
Public Const ERRO_INSERCAO_NFFATURA = 2401 'Parâmetro: lNumTitulo
'Erro na tentativa de inserir a Nota Fiscal Fatura número %l na Tabela TitulosPag.
Public Const ERRO_LEITURA_PARCELASPAG = 2402 'Parâmetro: lNumIntTitulo
'Erro na tentativa de ler Parcelas com NumIntTitulo %l na tabela ParcelasPag.
Public Const ERRO_LEITURA_PARCELASPAG1 = 2403 'Parâmetros: lNumIntTitulo, iNumParcela
'Erro na tentativa de ler Parcela %i com NumIntTitulo %l na tabela ParcelasPag.
Public Const ERRO_PARCELA_PAGAR_NAO_CADASTRADA = 2404 'Parâmetros: iNumParcela
'A Parcela %i deste Título não foi encontrada.
Public Const ERRO_LOCK_PARCELASPAG = 2405 'Parâmetros: lNumIntTitulo, iNumParcela
'Erro na tentativa de "lock" na Parcela %i com NumIntTitulo %l na tabela ParcelasPag.
Public Const ERRO_ATUALIZACAO_PARCELASPAG = 2406 'Parâmetros: lNumIntTitulo, iNumParcela
'Erro na atualização da Parcela %i com NumIntTitulo %l na tabela ParcelasPag.
Public Const ERRO_INSERCAO_PARCELASPAG = 2407 'Parâmetro: lNumIntTitulo, iNumParcela
'Erro na tentativa de inserção para o título com NumIntTitulo %l da Parcela %i na tabela ParcelasPag.
Public Const ERRO_NFFATPAG_PENDENTE_EXCLUSAO = 2408 'Parâmetro: lNumTitulo
'Não é possível excluir a Nota Fiscal Fatura número %l porque ela faz parte de um Lote Pendente.
Public Const ERRO_NFFATPAG_BAIXADA_EXCLUSAO = 2409 'Parâmetro: lNumTitulo
'Não é possível excluir a Nota Fiscal Fatura número %l porque ela está baixada.
Public Const ERRO_LOCK_TITULOSPAG = 2410 'Parâmetro: lNumIntDoc
'Erro na tentativa de "lock" no Título com NumIntDoc = %l na tabela TitulosPag.
Public Const ERRO_ATUALIZACAO_TITULOSPAG = 2411 'Parâmetro: lNumIntDoc
'Erro na atualizacao do Título com Número Interno %l na tabela TitulosPag.
Public Const ERRO_TIPOCOBRANCA_NAO_ENCONTRADO = 2412 'Parâmetro: sDescricao
'O Tipo de Cobrança %s não foi encontrado.
Public Const ERRO_LEITURA_NFFATPEND = 2413 'Parâmetro: lNumTitulo
'Erro na tentativa de leitura da Nota Fiscal Fatura %l na tabela TitulosPagPend.
Public Const ERRO_LEITURA_NFFATBAIXADA = 2414 'Parâmetro: lNumTitulo
'Erro na tentativa de leitura da Nota Fiscal Fatura %l na tabela TitulosPagBaixados.
Public Const ERRO_NUMTITULO_NAO_PREENCHIDO = 2415 'Sem parametros
'O campo Número não foi preenchido
Public Const ERRO_NUM_MAXIMO_PARCELAS_ULTRAPASSADO = 2416 'Parâmetros: iNumParcelasBD, iNumMaxParcelas
'O número de parcelas lidas no BD é %i que supera o número máximo permitido igual a %i.
Public Const ERRO_TIPOCOBRANCA_NAO_CADASTRADO = 2417 'Parâmetro: iCodigo
'O Tipo de Cobrança com Código %i não está cadastrado.
Public Const ERRO_VALORPARCELA_NAO_INFORMADO = 2418 'Parâmetro: iParcela
'O Valor da Parcela %i não foi informado.
Public Const ERRO_TITULOPAGAR_SEM_PARCELAS = 2419 'Parâmetro: lNumIntDoc
'Título a Pagar com número interno %l não tem Parcelas associadas.
Public Const ERRO_PAGTO_ANTECIPADO_INEXISTENTE = 2420
'O Pagamento Antecipado não está cadastrado.
Public Const ERRO_MODIFICACAO_PAGTO_ANTECIPADO = 2421
'Erro na tentativa de modificar a tabela de Pagamentos Antecipados.
Public Const ERRO_LEITURA_PAGTO_ANTECIPADO2 = 2422
'Erro na leitura da tabela de Pagamentos Antecipados.
Public Const ERRO_CREDITO_PAG_FORN_INEXISTENTE = 2423
'O Crédito não está cadastrado.
Public Const ERRO_MODIFICACAO_CREDITO_PAG_FORN = 2424
'Erro na tentativa de modificar a tabela de Créditos a Pagar.
Public Const ERRO_CREDITO_PAG_FORN_EXCLUIDO = 2425
'Esse Crédito está excluído.
Public Const ERRO_LEITURA_BAIXAPAG = 2426
'Erro na leitura da tabela de Baixas a Pagar.
Public Const ERRO_BAIXAPAG_INEXISTENTE = 2427
'A Baixa não está cadastrada.
Public Const ERRO_EXCLUSAO_BAIXAPAG = 2428
'Erro na exclusão da Baixa a Pagar.
Public Const ERRO_BAIXAPAG_EXCLUIDA = 2429
'A baixa já havia sido cancelada anteriormente.
Public Const ERRO_PAGTO_ANTECIPADO_EXCLUIDO = 2430
'O pagamento antecipado já havia sido excluído.
Public Const ERRO_BAIXAPARCPAG_INEXISTENTE = 2431
'A Baixa não está cadastrada.
Public Const ERRO_BAIXAPARCPAG_EXCLUIDA = 2432
'A baixa da parcela já havia sido cancelada anteriormente."
Public Const ERRO_EXCLUSAO_ANTECIPPAG2 = 2433
'Erro na exclusão do Pagamento Antecipado.
Public Const ERRO_BORDERO_PAGTO_EXCLUIDO = 2434
'Esse Borderô de Pagamento está excluído.
Public Const ERRO_EXCLUSAO_BORDERO_PAGTO = 2435
'Erro na exclusão do Borderô de Pagamento.
Public Const ERRO_BORDERO_PAGTO_INEXISTENTE = 2436
'O Borderô de Pagamento não está cadastrado.
Public Const ERRO_LEITURA_BORDERO_PAGTO = 2437
'Erro na leitura da tabela de Borderô de Pagamento.
Public Const ERRO_ANTECIPPAG_EXCLUIDO1 = 2438 'Parametro: lNumMovto
'O pagamento antecipado associado ao movimento %l de conta corrente já está marcado como excluído.
Public Const ERRO_ANTECIPPAG_INEXISTENTE2 = 2439 'Parametro: lNumMovto
'Erro na leitura do pagamento antecipado associado ao movimento %l de conta corrente.
Public Const ERRO_MOVIMENTOSCONTACORRENTE_INEXISTENTE = 2440
'O Movimento de Conta Corrente não está cadastrado.
Public Const ERRO_LEITURA_PARCELAS_PAG2 = 2441
'Erro na leitura da tabela de Parcelas a Pagar.
Public Const ERRO_PAGTO_CONCILIADO = 2442
'Um pagamento conciliado não pode ser cancelado. Desconcilie-o antes de tentar excluí-lo.
Public Const ERRO_NUMPAGTO_NAO_INFORMADO = 2443
'O Número de Pagamento não foi informado.
Public Const ERRO_MEIOPAGTO_NAO_INFORMADO = 2444
'O Meio de Pagamento não foi informado.
Public Const ERRO_LEITURA_BAIXAS_MOVTO_CTA = 2445
'Erro na leitura de baixas associadas a um movimento de conta corrente.
Public Const ERRO_LOCK_BAIXAPAG = 2446
'Não conseguiu fazer o lock da Baixa a Pagar.
Public Const ERRO_UNLOCK_BAIXAPAG = 2447
'Erro na tentativa de desfazer o lock na tabela de Baixas a Pagar.
Public Const ERRO_MODIFICACAO_BAIXAPAG = 2448
'Erro na tentativa de modificar a tabela de Baixas a Pagar.
Public Const ERRO_LEITURA_BAIXAPARCPAG = 2449
'Erro na leitura da tabela de Baixas de Parcelas a Pagar.
Public Const ERRO_UNLOCK_BAIXAPARCPAG = 2450
'Erro na tentativa de desfazer o lock na tabela de Baixas de Parcelas a Pagar.
Public Const ERRO_LOCK_BAIXAPARCPAG = 2451
'Não conseguiu fazer o lock da Baixa de Parcela a Pagar.
Public Const ERRO_MODIFICACAO_BAIXAPARCPAG = 2452
'Erro na tentativa de modificar a tabela de Baixas de Parcelas a Pagar.
Public Const ERRO_SALDO_PARCELA_MAIOR_QUE_VALOR = 2453
'O saldo da parcela não pode ser maior que o seu valor.
Public Const ERRO_TITULO_PAGAR_INEXISTENTE = 2454
'O Título a Pagar não está cadastrado.
Public Const ERRO_PARCELA_PAGAR_INEXISTENTE = 2455
'A Parcela a Pagar não está cadastrada.
Public Const ERRO_EXCLUSAO_TITULOS_PAGAR_BAIXADOS = 2456
'Erro na exclusão do Título a Pagar Baixado.
Public Const ERRO_EXCLUSAO_PARCELAS_PAGAR_BAIXADAS = 2457
'Erro na exclusão da Parcela a Pagar Baixada.
Public Const ERRO_TITULO_PAGAR_BAIXADO_INEXISTENTE = 2458
'O Título a Pagar Baixado não está cadastrado.
Public Const ERRO_LEITURA_PARCELAS_PAG_BAIXADA = 2459
'Erro na leitura da tabela de Parcelas a Pagar Baixadas.
Public Const ERRO_PARCELA_PAGAR_BAIXADA_INEXISTENTE = 2460
'A Parcela a Pagar Baixada não está cadastrada.
Public Const ERRO_LOCK_PARCELAS_PAGAR_BAIXADA = 2461
'Não conseguiu fazer o lock de Parcelas a Pagar Baixadas.
Public Const ERRO_LOCK_TITULOS_PAGAR_BAIXADOS = 2462
'Não conseguiu fazer o lock de Títulos a Pagar Baixados.
Public Const ERRO_UNLOCK_PARCELAS_PAGAR_BAIXADA = 2463
'Erro na tentativa de desfazer o lock na tabela de Parcelas a Pagar Baixadas.
Public Const ERRO_UNLOCK_TITULOS_PAGAR_BAIXADOS = 2464
'Erro na tentativa de desfazer o lock na tabela de Títulos a Pagar Baixados.
Public Const ERRO_MODIFICACAO_TITULOS_PAGAR = 2465
'Erro na tentativa de modificar a tabela de Títulos a Pagar.
Public Const ERRO_INCLUSAO_CARTEIRAS_COBRADOR = 2468
'Erro na tentativa de inserir a Carteira Cobrador.
Public Const ERRO_EXCLUSAO_CARTEIRAS_COBRADOR = 2469
'Erro na exclusão da Carteira Cobrador.
Public Const ERRO_EXCLUSAO_COBRADOR = 2470 'Parametro codigo do Cobrador
'Erro na exclusão do Cobrador.
Public Const ERRO_COBRADOR_INEXISTENTE = 2471 'Parametro codigo do Cobrador
'O Cobrador não está cadastrado.
Public Const ERRO_LEITURA_CARTEIRAS_COBRANCA = 2472
'Erro na leitura da tabela de Carteiras Cobrança.
Public Const ERRO_MODIFICACAO_CARTEIRAS_COBRANCA = 2473
'Erro na tentativa de modificar a tabela de Carteiras Cobrança.
Public Const ERRO_MODIFICACAO_COBRADOR = 2474 'Parametro codigo do Cobrador
'Erro na inserção na tabela de Cobrador , no codigo %i.
Public Const ERRO_INSERCAO_COBRADOR = 2475 'Parametro codigo do Cobrador
'Erro na leitura da tabela de Cobrador , no codigo %i.
Public Const ERRO_EXCLUSAO_CARTEIRAS_COBRANCA = 2476
'Erro na exclusão da Carteira Cobrança.
Public Const ERRO_INCLUSAO_CARTEIRAS_COBRANCA = 2477
'Erro na tentativa de inserir a Carteira Cobrança.
Public Const ERRO_LEITURA_PARCELASREC_SEM_CHEQUEPRE = 2479 'Sem parametros
'Erro na leitura de parcelas à receber para vinculação com cheque pré-datado
Public Const ERRO_CHEQUEPRE_INEXISTENTE = 2481 'Parametro: lNumIntCheque
'O Cheque Pre %l não existe na tabela de ChequePre.
Public Const ERRO_CLIENTE_CHQPRE_NAO_PREENCHIDO = 2482 'Sem parametro
'O Cliente não foi informado.
Public Const ERRO_FILIALCLIENTE_NAO_ENCONTRADA = 2483 'Parametro: Filial.Text
'A Filial %s não foi encontrada na tabela de FiliaisClientes.
Public Const ERRO_BANCO_CHQPRE_NAO_PREENCHIDO = 2484
'O preenchimento de Banco é obrigatório.
Public Const ERRO_CONTA_CHQPRE_NAO_PREENCHIDA = 2485
'O preenchimento da Conta Corrente é obrigatório.
Public Const ERRO_NUMERO_CHQPRE_NAO_PREENCHIDO = 2486
'O preenchimento de Número é obrigatório.
Public Const ERRO_VALOR_CHQPRE_NAO_PREENCHIDO = 2487
'O preenchimento de Valor é obrigatório.
Public Const ERRO_DATADEPOSITO_CHQPRE_NAO_PREENCHIDA = 2488
'O preenchimento de Data de Depósito é obrigatório.
Public Const ERRO_LEITURA_PARCELASREC_TITULOSREC = 2489 'lNumIntCheque
'Erro na tentativa de ler registro na tabela ParcelasRec E TitulosRec com NumIntCheque = %l
Public Const ERRO_DATADEPOSITO_MAIOR_DATA_CONTABIL = 2493 'Parametro: dtDataDep, dtDataContab
'A Data Contábil é maior que a Data de Depósito.
Public Const ERRO_TIPO_PARCELA_NAO_INFORMADO = 2494 'Parâmetro: iParcela
'O campo Tipo da Parcela %i não foi preenchido.
Public Const ERRO_NUMTITULO_PARCELA_NAO_INFORMADO = 2495 'Parâmetro: iParcela
'O campo NumTitulo da Parcela %i não foi preenchido.
Public Const ERRO_PARCELA_NAO_INFORMADA = 2496 'Parâmetro: iParcela
'O campo NumTitulo da Parcela %i não foi preenchido.
Public Const ERRO_VALOR_PARCELA_NAO_INFORMADA = 2497 'Parâmetro: iParcela
'O campo Valor da Parcela %i não foi preenchido.
Public Const ERRO_CHEQUEPRE_NUMBORDERO_DEPOSITADO = 2499 'Parametro: iBanco, sAgencia, sContaCorrente, lNumero
'O ChequePre com Banco %i, Agência %s, ContaCorrente %s e Número %l já foi depositado.
Public Const ERRO_AGENCIA_CHQPRE_INVALIDA = 2500 'Parametro: sAgencia
'A Agência %s é inválida.
Public Const ERRO_AGENCIA_CHQPRE_NAO_PREENCHIDA = 2501 'Parametro: sAgencia
'O preenchimento da Agência é obrigatório.
Public Const ERRO_CONTA_CHQPRE_INVALIDA = 2502 'Parametro: sConta
'A Conta %s é inválida.
Public Const ERRO_FILIAL_CHQPRE_NAO_PREENCHIDA = 2503 'Sem parâmetros
'Preenchimento do filial é obrigatório.
Public Const ERRO_NUMERO_NAO_E_NUMERICO = 2506 'Parâmetro Numero.Text
'O preenchimento de Número deve ser numérico.
Public Const ERRO_TIPO_FORNECEDOR_NAO_ENCONTRADO = 2507 'Parametro sDescricao
'O Tipo de Fornecedor %s não foi encontrado.
Public Const ERRO_PERCENTUAL_COMISSAO_NAO_INFORMADO = 2510 'iComissao
'O Percentual da Comissao %i do Título não foi informado.
Public Const ERRO_VALORBASE_COMISSAO_NAO_INFORMADO = 2511 'iComissao
'O Valor Base da Comissao %i do Título não foi informado.
Public Const ERRO_ALTERACAO_ENDERECO = 2512
'Erro de modificacao na tabela de Enderecos no Banco de Dados.
Public Const ERRO_FILIALCLIENTE_INEXISTENTE = 2513 'Parametro iCodCliente e lCodFilial
'A Filial %i do Cliente com código %l não existe.
Public Const ERRO_CODFILIAL_NAO_PREENCHIDO = 2514
'O código da Filial não foi preenchido.
Public Const ERRO_NOMEFILIAL_NAO_PREENCHIDO = 2515
'O nome da Filial não foi preenchido.
Public Const ERRO_FILIALCLIENTE_EXCLUSAO_MATRIZ = 2516
'A exclusão da matriz deve ser feita na tela de Clientes.
Public Const ERRO_LOCK_ENDERECOS = 2517
'Erro na tentativa de lockar a tabela de Enderecos.
Public Const ERRO_EXCLUSAO_ENDERECO = 2518 'Parametro lCodigoEndereco
'Erro na tentativa de excluir o Endereco de Código = %l .
Public Const ERRO_TIPO_NAO_ANTECIPREC = 2519 'Sem parâmetros
'O Movimento não é do tipo Recebimento Antecipado.
Public Const ERRO_LEITURA_ANTECIPREC = 2520 'Parâmetro lNumIntRec
'Erro na leitura do Recebimento antecipado %l.
Public Const ERRO_LEITURA_ANTECIPREC1 = 2521 'Parâmetro lNumIntMovto
'Erro na leitura do Recebimento antecipado associado ao Número de movimento: %l na tabela RecebAntecipados.
Public Const ERRO_INSERCAO_ANTECIPREC = 2522 'Parâmetro: lNumIntRec
'Ocorreu um erro na tentativa de inserção do Recebimento antecipado %l na Tabela RecebAntecipados.
Public Const ERRO_EXCLUSAO_ANTECIPREC = 2523 'Parâmetro: lNumIntRec
'Ocorreu um erro na tentativa de exclusão do Recebimento antecipado %l.
Public Const ERRO_LOCK_ANTECIPREC = 2524 'Parâmetro lNumIntRec
'Erro na tentativa de fazer o "lock" do Recebimento antecipado %l na tabela RecebAntecipados
Public Const ERRO_ANTECIPREC_EXCLUIDO = 2525 'Parâmetros iCodConta + lSequencial
'O Recebimento antecipado com a Conta %i e Sequencial %l está excluído.
Public Const ERRO_RECEBIMENTO_APROPRIADO = 2526 'Parâmetro iCodigo
'Não é possível excluir o Recebimento antecipado %i, pois já foi apropriado (total ou parcialmente).
Public Const ERRO_CLIENTE_NAO_COINCIDE = 2527 'Parâmetros lCliente(da tela) e lCliente(da tabela)
'O Cliente %l não coincide com o Cliente %l cadastrado no Recebimento Antecipado.
Public Const ERRO_ANTECIPREC_INEXISTENTE = 2528  'Parâmetros iCodConta + lSequencial
'O Movimento com a Conta %i e Sequencial %l não está cadastrado.
Public Const ERRO_TIPOAPLICACAO_INATIVO = 2529  'Parametro: iCodigo
'O Tipo De Aplicacao com Código %i está inativo.
Public Const ERRO_TIPOAPLICACAO_INEXISTENTE1 = 2530  'Parametro: iTipoAplicacao
'Não existe o tipo de aplicacao %i na tabela de Tipos De Aplicacao.
Public Const ERRO_TIPOAPLICACAO_INEXISTENTE2 = 2531 'Paramento: TipoAplicacao.Text
'O Tipo de Aplicacao %s nao existe.
Public Const ERRO_LEITURA_TIPOSDEAPLICACAO1 = 2532  'Sem parametro
'Erro na leitura da tabela de Tipos De Aplicacoes.
Public Const ERRO_LOCK_TIPOSDEAPLICACAO = 2533 'Parametro: iTipoAplicacao
'Erro no "lock" do Tipo de Aplicação i% da tabela TiposDeAplicacao.
Public Const ERRO_LOCK_APLICACOES = 2534 'Parametro: lCodigo
'Erro no "lock" da tabela Alpicações. Código = %l.
Public Const ERRO_APLICACAO_INEXISTENTE = 2535 'Parametro: lCodigo
'A Aplicação %l não existe na tabela de Aplicações.
Public Const ERRO_APLICACAO_EXCLUIDA = 2536 'Parametro: lCodigo
'A Aplicacao com Código %l está excluída.
Public Const ERRO_TIPO_NAO_APLICACAO = 2537 'Parametros: lSequencial
'O movimento %l não é do tipo Aplicação.
Public Const ERRO_MOVIMENTO_EXCLUIDO = 2538 'Parametro: lNumMovto
'O Movimento %l está excluído.
Public Const ERRO_TIPOAPLICACAO_NAO_PREENCHIDO = 2539 'Sem parametro
'O preenchimento do Tipo De Aplicação é obrigatório.
Public Const ERRO_DATA_APLICACAO_NAO_PREENCHIDA = 2540 'Sem parametro
'O preenchimento da Data De Aplicacao é obrigatório.
Public Const ERRO_CONTACORRENTE_NAO_PREENCHIDA = 2541 'Sem parametro
'O preenchimento de Conta Corrente é obrigatório.
Public Const ERRO_VALRESGATE_MENOR_VALAPLICADO = 2542 'Parametro: dValResgPrev, dValAplic
'O Valor do Resgate Previsto %d é menor que o Valor Aplicado %d.
Public Const ERRO_DATARESGPREV_MENOR_DATAAPLIC = 2543 'Parametro: dtDataAplicacao, dtDataAplic
'A Data de Resgate Prevista %dt é menor que a Data de Aplicação %dt.
Public Const ERRO_VALORAPLICADO_NAO_INFORMADO = 2544
'O valor aplicado não foi informado.
Public Const ERRO_DATAAPLICACAO_MENOR = 2545 'Parametro: dtDataAplicacao, dtDataSaldoInicial
'A Data da Aplicacao dt% é menor do que a Data Inicial da Conta dt%.
Public Const ERRO_ATUALIZACAO_APLICACOES = 2546 'Parametro: lCodigo
'Erro na Atualizacao da Aplicação l% da tabela de Aplicações.
Public Const ERRO_INSERCAO_APLICACOES = 2547 'Parametro: lCodigo
'Erro na Insercao da Aplicação %l na tabela Aplicações.
Public Const ERRO_LEITURA_RESGATES1 = 2548 'Parametro: lCodigoAplicacao
'Erro na leitura dos Resgates da Aplicação %l.
Public Const ERRO_APLICACAO_RESGATE = 2549 'Parametro: lCodigoAplicacao
'Existe Resgate associado a Aplicacao com Código %l.
Public Const ERRO_LEITURA_FAVORECIDOS = 2550 'Parametro: iFavorecido
'Erro na leitura do Favorecido %i da tabela de Favorecidos.
Public Const ERRO_LOCK_FAVORECIDOS1 = 2551 'Parametro iFavorecido
'Erro na tentativa de fazer "lock" do Favorecido i% na tabela Favorecidos.
Public Const ERRO_LEITURA_MOVIMENTOSCONTACORRENTE3 = 2552 'Parametro: lNumMovto
'Erro na leitura do Movimento l% da tabela de Movimentos de Conta Corrente.
Public Const ERRO_LEITURA_MOVIMENTOSCONTACORRENTE2 = 2553 'Parametros: iCodConta, iTipoMeioPagto, lNumero
'Erro na leitura do Movimento da Conta i%, Tipo de Pagamento i% e Número l% da tabela de MovimentosContaCorrente.
Public Const ERRO_LOCK_MOVIMENTOSCONTACORRENTE2 = 2554 'Parametro: lNumMovto
'Erro na tentativa de fazer "lock"  no Movimento l% na tabela de MovimentosContasCorrente.
Public Const ERRO_MOVCONTACORRENTE_EXCLUIDO1 = 2555 'Parametro: lNumMovto
'O Movimento l% está excluído.
Public Const ERRO_MOVIMENTO_NAO_CADASTRADO3 = 2556 'Parametro: lNumMovto
'O Movimento %l não está cadastrado.
Public Const ERRO_TIPOMEIOPAGTO_INEXISTENTE1 = 2557 'Parametro: TipoMeioPagto.Text
'O Tipo de Pagmento %s nao está cadastrado.
Public Const ERRO_ATUALIZACAO_MOVIMENTOSCONTACORRENTE2 = 2558 'Parametro: lNumMovto
'Erro de Atualizacao da Tabela de MovimentosContaCorrente com número do movimento %l.
Public Const ERRO_INSERCAO_MOVIMENTOSCONTACORRENTE2 = 2559 'Parametro: lNumMovto
'Erro na tentativa de inclusão de Movimento %l.
Public Const ERRO_RESGATE_INEXISTENTE = 2560 'Parametro: lCodigoAplicacao
'O Resgate com Código %l não existe na tabela de Resgates.
Public Const ERRO_RESGATE_INEXISTENTE1 = 2561 'Parametro: lCodigoAplicacao, iSeqResgate
'O Resgate com Código %l e Sequencial %i não existe na tabela de Resgates.
Public Const ERRO_RESGATE_EXCLUIDO = 2562 'Parametro: iSeqResgate, lCodigoAplicacao
'O Resgate %i da Aplicação %l está excluído.
Public Const ERRO_TIPO_NAO_RESGATE = 2563 'Parametro: lSequencial
'O movimento %l não é do tipo Resgate.
Public Const ERRO_LEITURA_RESGATES = 2564 'Parametro: lCodigoAplicacao
'Erro de Leitura na Tabela de Resgates com o Código %l.
Public Const ERRO_CODIGO_APLICACA0_NAO_PREENCHIDO = 2565 'Sem parametros
'Preenchimento do código da aplicação é obrigatório.
Public Const ERRO_CODIGO_RESGATE_NAO_PREENCHIDO = 2566 'Sem parametros
'Preenchimento do código do resgate é obrigatório.
Public Const ERRO_DATA_RESGATE_NAO_PREENCHIDA = 2567 'Sem parametro
'O preenchimento da Data De Resgate é obrigatório.
Public Const ERRO_VALOR_CREDITADO_NAO_INFORMADO = 2568 'Sem parametro
'O valor creditado não foi informado.
Public Const ERRO_SALDO_ATUAL_NEGATIVO = 2569  'Parametro: SaldoAtual.Caption
'O Saldo Atual está negativo.
Public Const ERRO_DATARESGPREV_MENOR_DATARESG = 2570 'Parametro: dtDataResgPrev, dtDataResg
'A Data de Resgate Prevista %dt é menor que a Data do Resgate %dt.
Public Const ERRO_VALRESGATE_MENOR_SALATUAL = 2571 'Parametro: dValResgPrev, dSalAtual
'O Valor do Resgate Previsto %d é menor que o Saldo Atual %d.
Public Const ERRO_VALORES_DIFERENTES = 2572 'Parametro: dValCred, dValCredLab
'Os campos de Valor Creditado %d e Valor Creditado Label %d estão com valores diferentes.
Public Const ERRO_DATA_RESGATE_MENOR = 2573 'Parametro: dtDataMovimento, dtDataSaldoInicial
'A Data de Resgate dt% é menor do que a Data Inicial da Conta dt%.
Public Const ERRO_LOCK_RESGATES = 2574 'Parametro: lCodigoAplicacao, iSeqResgate
'Erro no "lock" da tabela Resgates. Código = %l e Sequencial = %i.
Public Const ERRO_ATUALIZACAO_RESGATES = 2575 'Parametro: lCodigoAplicacao, iSequencial
'Erro na Atualizacao do Resgate com Código l% e Sequencial %i da tabela de Resgates.
Public Const ERRO_INSERCAO_RESGATES = 2576 'Parametro: lCodigoAplicacao, iSeqResgate
'Erro na Insercao do Resgatec com Código %l e Sequencial %i na tabela de Resgates.
Public Const ERRO_LOCK_FLUXOTIPOAPLIC = 2577 'Parametro FluxoId, TipoAplicacao
'Ocorreu um erro ao tentar fazer lock no Fluxo Tipo Aplic %l, no tipo de aplicação %i.
Public Const ERRO_LEITURA_CONTASCORRENTESINTERNAS_CODIGO = 2578 'Parametro: iCodigo
'Erro na leitura da tabela de ContasCorrentesInternas. Codigo = %s.
Public Const ERRO_ATUALIZACAO_FLUXOSALDOSINICIAIS = 2579 'Parametro Nome do Fluxo de Caixa
'Erro na atualização da tabela de Saldos Iniciais do Fluxo de Caixa. Fluxo = %s.
Public Const ERRO_LOCK_FLUXOSINTETICO = 2580  'Parametro FluxoId
'Ocorreu um erro ao tentar fazer lock no Fluxo Sintetico. FluxoId = %l.
Public Const ERRO_INSERCAO_FLUXOCREDANTECIP = 2581 'Sem Parametro
'Erro na inserção de um registro na tabela FluxoCredAntecip.
Public Const ERRO_EXCLUSAO_FLUXOCREDANTECIP = 2582 'Parametro Nome do Fluxo de Caixa
'Erro na exclusão do Fluxo de Caixa (FluxoCredAntecip) %s.
Public Const ERRO_ATUALIZACAO_FLUXOANALITICO = 2583 'Parametro Nome do Fluxo de Caixa
'Erro na tentativa de atualização da tabela de Fluxo Analitico, do Fluxo de Caixa %s.
Public Const ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO = 2584
'O Campo Tipo não foi preenchido.
Public Const ERRO_DATA_RESGATE_MENOR1 = 2585 'Parametro: dtDataMovimento, dtDataAplicacao
'A Data de Resgate dt% é menor do que a Data de Aplicacao.
Public Const ERRO_CAMPOS_CREDITO_PAGAR_NAO_PREENCHIDOS = 2586 'Sem parâmetros
'Os campos Fornecedor, Filial e Tipo devem estar preenchidos.
Public Const ERRO_LEITURA_CREDITOSPAGFORN1 = 2588 'Parâmetro: lNumIntDoc
'Erro na leitura da tabela de CreditosPagForn com número interno do documento %l.
Public Const ERRO_LEITURA_BAIXASPAG1 = 2590 'Parâmetro: lNumIntDoc
'Erro na leitura da tabela BaixasPag com número interno do documento %l.
Public Const ERRO_TIPODOC_NAO_E_TIPOCREDITOPAG = 2591 'Parâmetro: sSiglaDocumento
'O Tipo de Documento %s não é do tipo Devoluções / Crédito.
Public Const ERRO_CREDITOPAGAR_NAO_CADASTRADO = 2594 'Parâmetros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Devoluções / Crédito do Fornecedor %l, Filial %i, Tipo de Documento %s, Número %l e Data de Emissão %dt não existe na tabela CreditosPagForn.
Public Const ERRO_CREDITOPAGAR_NAO_CADASTRADO1 = 2595 'Parametro: lNumIntDoc
'Devoluções / Crédito com número interno do documento %l não existe na tabela de CreditosPagForn.
Public Const ERRO_LOCK_CREDITOSPAGFORN = 2597 'Parâmetros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Não conseguiu fazer o "lock", na tabela de CreditosPafForn com Fornecedor %l, Filial %i, Tipo de Documento %s, Número %l e Data de Emissão dt%.
Public Const ERRO_ALTERACAO_CREDPAG_LANCADO = 2598 'Parâmetros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Não é possível alterar Devolução / Crédito, com dados Fornecedor %l, Filial %i, Tipo de Documento %s, Número %l e Data de Emissão %dt , porque está lançado.
Public Const ERRO_ALTERACAO_CREDPAG_BAIXADO = 2599 'Parâmetros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Não é possível alterar Devolução / Crédito com dados Fornecedor %l, Filial %i, Tipo de Documento %s, Número %l e Data de Emissão %dt porque está baixado.
Public Const ERRO_EXCLUSAO_CREDPAG_BAIXADO = 2600 'Parâmetros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Não é possivel excluir Devolução / Crédito porque está baixado. Dados da Devolução/Crédito: Fornecedor %l, Filial %i, Tipo de Documento %s, Número %l, Data de Emissão dt%.
Public Const ERRO_EXCLUSAO_CREDPAG_VINCULADO_BAIXA = 2601 'Parâmetros: lNumIntBaixa, lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissão
'Não é possivel excluir Devolução / Crédito porque está vinculado a Baixa Pagar com número interno %l. Dados do Crédito: Fornecedor %l, Filial %i, Tipo de Documento %s, Número %l, Data de Emissão dt%.
Public Const ERRO_MODIFICAO_CRED_PAG_OUTRA_FILIALEMPRESA = 2602 'Parâmetros: iFilialEmpresa, lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Não é possível modificar Devoluções / Crédito da Filial Empresa %i. Dados do Documento: código do Fornecedor %l, código da Filial %i, Tipo %s, Número %l e Data de Emissao dt%.
Public Const ERRO_EXCLUSAO_CREDITOSPAGFORN = 2603 ''Parâmetros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Erro na exclusão de Devoluções / Crédito com dados Fornecedor %l, Filial %i, Tipo de Documento %s, Número %l e Data de Emissão %dt, da tabela CreditosPagForn.
Public Const ERRO_LEITURA_NFISCALBAIXADAS2 = 2604 'Parâmetro: lCodFornecedor, iCodFilial
'Erro  na leitura da tabela NFiscalBaixadas com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_NFSPAG3 = 2605 'Parâmetro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela NFsPag com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_NFSPAGPEND2 = 2606 'Parâmetro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela NFsPagPend com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_NFSPAGBAIXADAS2 = 2607 'Parâmetro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela NfsPagBaixadas com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_TITULOSPAG2 = 2608 'Parâmetro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela TitulosPag com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_TITULOSPAGPEND2 = 2609 'Parâmetro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela TitulosPagPend com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_TITULOSPAGBAIXADOS2 = 2610 'Parâmetro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela TitulosPagBaixados com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_CREDITOSPAGFORN3 = 2611 'Parâmetro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela CreditosPagForn com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_PAGTOSANTECIPADOS2 = 2612 'Parâmetro: lCodFornecedor, iCodFilial
'Erro na leitura da tabela PagtosAntecipados com Fornecedor %l e Filial %i.
Public Const ERRO_LEITURA_PENDENCIASBORDEROPAGTO2 = 2613 'Parâmetros: lCodFornecedor, iCodFilial
'Erro na leitura da tabela de PendenciasBorderoPagto com Fornecedor %l e Filial %i.
Public Const ERRO_FILIAL_FORNECEDOR_REL_NFISCAL = 2620 'Parâmetros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l está relacionada com Nota Fiscal com código interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_NFISCALBAIXADA = 2621 'Parâmetros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l está relacionada com Nota Fiscal Baixada com código interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_NFISCALPAG = 2622 'Parâmetros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l está relacionado com Nota Fiscal à Pagar com código interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_NFISCALPAGPEND = 2623 'Parâmetros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l está relacionado com Nota Fiscal à Pagar Pendente com código interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_NFISCALPAGBAIXADA = 2624 'Parâmetros: iCodFilial, lCodFornecedor, lCodigo
'A Filial do Fornecedor %l está relacionado com Nota Fiscal à Pagar Baixada com código interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_TIT_PAGAR = 2625 'Parâmetros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l está relacionado com Título à Pagar com código interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_TIT_PAGAR_PEND = 2626 'Parâmetros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l está relacionado com Título à Pagar Pendente com código interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_TIT_PAGAR_BAIXADO = 2627 'Parâmetros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %ii do Fornecedor %l está relacionado com Título à Pagar Baixado com código interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_CREDITO_PAGAR_FORN = 2628 'Parâmetros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l está relacionado com Crédito à Pagar Fornecedor com código interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_PAGTO_ANTECIPADO = 2629 'Parâmetros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l está relacionado com Pagamento Antecipado com código interno %l.
Public Const ERRO_FILIAL_FORNECEDOR_REL_PEND_BORDERO_PAGTO = 2630 'Parâmetros: iCodFilial, lCodFornecedor, lCodigo
'A Filial %i do Fornecedor %l está relacionado com Pendência Borderô Pagto com código interno %l.
Public Const ERRO_CODIGO_TIPO_FORNECEDOR_NAO_PREENCHIDO = 2631 'Sem parametros
'Preenchimento do Tipo de Fornecedor é obrigatório.
Public Const ERRO_LEITURA_NFISCALBAIXADAS1 = 2632 'Parâmetro: lCodigo
'Erro  na leitura da tabela NFiscalBaixadas com Fornecedor %l.
Public Const ERRO_FORNECEDOR_REL_NFISCAL = 2633 'Parâmetros: lCodFornecedor, lNFiscal
'O Fornecedor %l está relacionado com Nota Fiscal número %l.
Public Const ERRO_FORNECEDOR_REL_NFISCALBAIXADA = 2634 'Parâmetros: lCodFornecedor, lCodNFiscal
'O Fornecedor %l está relacionado com Nota Fiscal Baixada com código interno %l.
Public Const ERRO_FORNECEDOR_REL_NFISCALPAG = 2635 'Parâmetros: lCodFornecedor, lNFiscal
'O Fornecedor %l está relacionado com Nota Fiscal à Pagar com número %l.
Public Const ERRO_FORNECEDOR_REL_NFISCALPAGPEND = 2636 'Parâmetros: lCodFornecedor, lCodNFiscal
'O Fornecedor %l está relacionado com Nota Fiscal à Pagar Pendente com código interno %l.
Public Const ERRO_FORNECEDOR_REL_NFISCALPAGBAIXADA = 2637 'Parâmetros: lCodFornecedor, lNFiscal
'O Fornecedor %l está relacionado com Nota Fiscal à Pagar Baixada com número %l.
Public Const ERRO_FORNECEDOR_REL_TIT_PAGAR = 2638 'Parâmetros: lCodFornecedor, lCodTitPagar
'O Fornecedor %l está relacionado com Título à Pagar com código interno %l.
Public Const ERRO_FORNECEDOR_REL_TIT_PAGAR_PEND = 2639 'Parâmetros: lCodFornecedor, lCodTitPagar
'O Fornecedor %l está relacionado com Título à Pagar Pendente com código interno %l.
Public Const ERRO_FORNECEDOR_REL_TIT_PAGAR_BAIXADO = 2640 'Parâmetros: lCodFornecedor, lCodTitPagar
'O Fornecedor %l está relacionado com Título à Pagar Baixado com código interno %l.
Public Const ERRO_FORNECEDOR_REL_CREDITO_PAGAR_FORN = 2641 'Parâmetros: lCodigo, lCodigo
'O Fornecedor %l está relacionado com Crédito à Pagar Fornecedor com código interno %l.
Public Const ERRO_FORNECEDOR_REL_PAGTO_ANTECIPADO = 2642 'Parâmetros: lCodFornecedor, lCodPagtoAntec
'O Fornecedor %l está relacionado com Pagamento Antecipado com código interno %l.
Public Const ERRO_FORNECEDOR_REL_FORNECEDOR_PRODUTO = 2643 'Parâmetros: lCodFornecedor, sProduto
'O Fornecedor %l está relacionado com o Produto %s.
Public Const ERRO_FORNECEDOR_REL_PRODUTO_FILIAL = 2644 'Parâmetros: lCodigo, sProduto
'O Fornecedor %l está relacionado com o Produto %s da Filial.
Public Const ERRO_FORNECEDOR_REL_PEND_BORDERO_PAGTO = 2645 'Parâmetros: lCodFornecedor, lCodPendBordPagto
'O Fornecedor %l está relacionado com Pendência Borderô Pagto com código interno %l.
Public Const ERRO_LEITURA_NFSPAG2 = 2646 'Parâmetro: lCodigo
'Erro na leitura da tabela NFsPag com Fornecedor %l.
Public Const ERRO_LEITURA_NFSPAGPEND1 = 2647 'Parâmetro: lCodigo
'Erro na leitura da tabela NFsPagPend com Fornecedor %l.
Public Const ERRO_LEITURA_NFSPAGBAIXADAS1 = 2648 'Parâmetro: lCodigo
'Erro na leitura da tabela NfsPagBaixadas com Fornecedor %l.
Public Const ERRO_LEITURA_TITULOSPAG1 = 2649 'Parâmetro: lCodigo
'Erro na leitura da tabela TitulosPag com Fornecedor %l.
Public Const ERRO_LEITURA_TITULOSPAGPEND1 = 2650 'Parâmetro: lCodigo
'Erro na leitura da tabela TitulosPagPend com Fornecedor %l.
Public Const ERRO_LEITURA_TITULOSPAGBAIXADOS1 = 2651 'Parâmetro: lCodigo
'Erro na leitura da tabela TitulosPagBaixados com Fornecedor %l.
Public Const ERRO_LEITURA_CREDITOSPAGFORN2 = 2652 'Parâmetro: lCodigo
'Erro na leitura da tabela CreditosPagForn com Fornecedor %l.
Public Const ERRO_LEITURA_PAGTOSANTECIPADOS1 = 2653 'Parâmetro: lCodigo
'Erro na leitura da tabela PagtosAntecipados com Fornecedor %l.
Public Const ERRO_LEITURA_FORNECEDORPRODUTO2 = 2654 'Parâmetros: lCodigo
'Erro na leitura da tabela de FornecedorProduto com Fornecedor %l.
Public Const ERRO_LEITURA_PENDENCIASBORDEROPAGTO1 = 2656 'Parâmetros: lCodigo
'Erro na leitura da tabela de PendenciasBorderoPagto com Fornecedor %l.
Public Const ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA3 = 2663 'Parâmetro: iCondicaoPagto
'A Condição de Pagamento %i não está cadastrada no Banco de Dados.
Public Const ERRO_NFPAGAR_VINCULADA_NFISCAL = 2664 'Parâmetro: lNumNotaFiscal
'Nota Fiscal à Pagar com número %l está vinculada a Nota Fiscal.
Public Const ERRO_TIPODEFORNECEDOR_NAO_CADASTRADO = 2665 'Parametro iCodigo
'Tipo de Fornecedor %i não está cadastrado.
Public Const ERRO_INSERCAO_TIPOSDEFORNECEDOR = 2666 'Parametro iCodigo
'Erro na inserção do Tipo de Fornecedor %i.
Public Const ERRO_ATUALIZACAO_TIPOSDEFORNECEDOR = 2667 'Parametro iCodigo
'Erro na atualização do Tipo de Fornecedor %i.
Public Const ERRO_EXCLUSAO_TIPOSDEFORNECEDOR = 2668 'Parametro iCodigo
'Erro na exclusão do Tipo de Fornecedor %i.
Public Const ERRO_LOCK_TIPOSDEFORNECEDOR = 2669 'Parametro iCodigo
'Não conseguiu fazer o lock do Tipo de Fornecedor %i.
Public Const ERRO_TIPODEFORNECEDOR_RELACIONADO_COM_FORNECEDOR = 2670 'Sem parametros
'Não é possível excluir Tipo de Fornecedor relacionado com Fornecedor.
Public Const ERRO_DESCRICAO_TIPO_FORNECEDOR_REPETIDA = 2671 'Parâmetro: iCodigo
'Tipo de Fornecedor com código %i está cadastrado e tem esta descrição.
Public Const ERRO_CAMPOS_APLICACAO_NAO_ALTERAVEIS = 2672
'Os campos não serão alterados. Para alterar a aplicação exclua esta e crie uma outra.
Public Const ERRO_DEBITORECCLI_NAO_ENCONTRADO = 2673 'Parametro: lNumIntDoc
'O Débito a Receber de Cliente com número %l não está cadastrado no Banco de Dados.
Public Const ERRO_NUM_MAXIMO_COMISSOES_ULTRAPASSADO = 2674 'Parametros: iNumParcelasBD, iNumMaxParcelas
'O número de parcelas lidas no BD é %i que supera o número máximo permitido igual a %i.
Public Const ERRO_LEITURA_COMISSOES_VENDEDORES = 2675 'lNumIntDoc
'Erro na tentativa de ler registro na tabela Comissões e Vendedores com NumIntDoc = %l.
Public Const ERRO_LEITURA_DEBITOSRECCLI1 = 2677 'Parametros: lCliente, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Erro na leitura da tabela DebitosRecCli com Cliente %l, Filial %i, Tipo %s, Número do Título %l e Data de Emissão %dt.
Public Const ERRO_NAO_E_POSSIVEL_MODIFICAR_DEB_REC_CLI_OUTRA_FILIALEMPRESA = 2679 'Sem parametros
'Não é possível modificar Débito a Receber Cliente de outra Filial da Empresa.
Public Const ERRO_DEBITO_REC_CLI_BAIXADO = 2680 'Parametro: lCliente, sVendedorNomeRed
'Não é possível alterar as comissões do Débito a Receber de Cliente %l. A comissão do Vendedor %s está paga.
Public Const ERRO_DEBITORECCLI_NAO_ENCONTRADO1 = 2681 'Parametros: lCliente, iFilial, sSiglaDocumento, lNumTitulo
'Devolução / Crédito do Cliente %l, Filial %i, Tipo de Documento %s e Número do Título %l não está Cadastrado.
Public Const ERRO_LEITURA_DEBITOSRECCLI2 = 2682 'Parametros: lCliente, iFilial, sSiglaDocumento, lNumTitulo
'Erro na leitura da tabela DebitosRecCli com Cliente %l, Filial %i, Tipo de Documento %s e Número do Título %l.
Public Const ERRO_LOCK_DEBITOSRECCLI = 2683  'Parametros: lCliente, iFilial, sSiglaDocumento, lNumTitulo
'Não conseguiu fazer o lock na tabela de DebitosRecCli com Cliente %l, Filial %i Tipo de Documento %s Número do Título %l.
Public Const ERRO_NAO_E_PERMITIDO_EXCLUSAO_DEBRECCLI_VINCULADO_BAIXA = 2684 'Parametro: lNumIntDoc
'Não é permitida a exclusão de Débito a Receber Cliente número %l vinculado a baixa.
Public Const ERRO_EXCLUSAO_DEBITOSRECCLI = 2685 'Parametro: lNumIntDoc
'Erro na exclusão de Débito a Receber Cliente número %l da tabela de DebitosRecCli.
Public Const ERRO_CAMPOS_DEBITO_RECEBER_NAO_PREENCHIDOS = 2686 'Sem parametros
'Os campos Cliente, Filial e Tipo devem estar preenchidos.
Public Const ERRO_PORTADOR_NAO_ENCONTRADO = 2690 'Parametro: sPortador
'O Portador %s não foi encontrado.
Public Const ERRO_TIPOCOBRANCA_NAO_PREENCHIDO = 2691
'O Tipo de Cobrança não foi informado
Public Const ERRO_TITULO_NAO_CADASTRADO2 = 2692 'Parâmetro: sTipo, sFornecedor, sFilial, sNumTitulo, sDataEmissao
'Não foi encontrado nenhum Título do Tipo %s, com Fornecedor = %s, Filial = %s, Número = %s e Data de Emissao %s.
Public Const ERRO_PARCELA_PAG_NAO_ABERTA2 = 2693 'Parâmetros: lNumTítulo, iParcela
'A Parcela i% do Título Número %l não pode ser modificada porque não está aberta.
Public Const ERRO_LEITURA_FATURASBAIXADAS = 2694 'Parâmetro: lNumnotafiscal
'Erro na leitura da Fatura número %l na tabela NfsPagBaixadas.
Public Const ERRO_FATURA_PENDENTE_MODIFICACAO = 2695 'Parâmetro: lNumTitulo
'Não é possível modificar Fatura número %l .Ela faz parte de um Lote Pendente.
Public Const ERRO_FATURA_BAIXADA_MODIFICACAO = 2696 'Parâmetro: lNumTitulo
'Não é possível modificar Fatura número %l porque ela está baixada.
Public Const ERRO_LEITURA_FATURASPAGBAIXADAS = 2697 'Parâmetro: lNumTitulo
'Erro na tentativa de leitura da Fatura %l na tabela TitulosPagBaixados.
Public Const ERRO_TITULO_NAO_FATURAPAGAR = 2698 'Parâmetro: lNumTitulo
'O Título com número %l não é uma Fatura a Pagar.
Public Const ERRO_FATURAPAG_NAO_CADASTRADA = 2699 'Parâmetro: lNumIntDoc
'A Fatura com Número Interno = %l não está cadastrada.
Public Const ERRO_FATURAPAG_NAO_CADASTRADA1 = 2700 'Parâmetro: lNumTitulo
'A Fatura Número %l está não cadastrada ou em Lote Pendente ou Baixada.
Public Const ERRO_SOMA_NFS_SELECIONADAS_INVALIDA = 2701 'Parâmetros: sValorTotalNfs, sValorTotal
'A soma dos valores a pagar das Notas Fiscais selecionadas %s não é igual ao Valor Total da Fatura %s.
Public Const ERRO_FATURAPAG_PENDENTE_EXCLUSAO = 2702 'lNumTitulo
'Não é possível excluir a Fatura número %l porque ela faz parte de um Lote Pendente.
Public Const ERRO_FATURAPAG_BAIXADA_EXCLUSAO = 2703 'lNumTitulo
'Não é possível excluir a Fatura número %l porque ela está baixada.
Public Const ERRO_FATURAPAG_NAO_CADASTRADA2 = 2704 'Parametro: lNumTitulo
'A Fatura %l não está cadastrada no Banco de Dados.
Public Const ERRO_NF_JA_VINCULADA = 2705 'Parâmetro: lNumNotaFiscal
'A Nota Fiscal número %l já está vinculada a outra Fatura.
Public Const ERRO_LEITURA_FATURA = 2706 'Parâmetro: lNumIntDoc
'Erro na tentativa de ler Fatura na tabela TitulosPag com Número Interno %l.
Public Const ERRO_LEITURA_FATURA1 = 2707 'Parâmetro: lNumTitulo
'Erro na tentativa de ler Fatura número %l na tabela TitulosPag.
Public Const ERRO_LEITURA_FATURAPEND = 2708 'lNumTitulo
'Erro na leitura da Fatura número %l na Tabela de Titulos Pendentes.
Public Const ERRO_LEITURA_FATURABAIXADA = 2709 'lNumTitulo
'Erro na leitura da Fatura número %l na Tabela de TiTulos Baixados.
Public Const ERRO_LEITURA_TITULOSPAG3 = 2710 'Parâmetros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao, sSiglaDocumento
'Erro na leitura da tabela TitulosPag com Fornecedor %l, Filial %i, Título, Data de Emissão %dt e Tipo de Documento %s.
Public Const ERRO_INSERCAO_TITULOSPAG = 2711 'Parâmetros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao, sSiglaDocumento
'Erro na tentativa de inserir um registro na tabela TitulosPag com Fornecedor %l, Filial %i, Título, Data de Emissão %dt e Tipo de Documento %s.
Public Const ERRO_NFFATPAGAR_VINCULADA_NFISCAL = 2712 'Parâmetro: lNumNotaFiscal
'Nota Fiscal Fatura à Pagar com número %l está vinculada à Nota Fiscal.
Public Const ERRO_TIPO_INSTRUCAO_COBRANCA_NAO_ENCONTRADO = 2713 'Parametro iCodigo
'Instrução de Cobrança %i não foi encontrada.
Public Const ERRO_TIPO_INSTRUCAO_COBRANCA_NAO_CADASTRADA = 2714 'Parametro iCodigo
'Instrução de Cobrança %i não está cadastrada.
Public Const ERRO_DIAS_DE_PROTESTO1_NAO_PREENCHIDO = 2715 'Sem Parametros
'Dias para Devolução / Protesto de Instrução Primária deve ser preenchido.
Public Const ERRO_DIAS_DE_PROTESTO2_NAO_PREENCHIDO = 2716 'Sem Parametros
'Dias para Devolução / Protesto de Instrução Secundária deve ser preenchido.
Public Const ERRO_LEITURA_PADROES_COBRANCA = 2717 'Sem Parametros
'Erro na leitura da tabela PadroesCobranca
Public Const ERRO_INSTRUCAO_PRIMARIA_NAO_PREENCHIDA = 2718 'Sem Parametros
'O preenchimento da Instrução Primária é obrigatório
Public Const ERRO_LEITURA_TIPOSMOVIMENTO = 2719 'Sem parametro
'Erro na Leitura da Tabela  "TiposMovimento".
Public Const ERRO_CANCELAMENTO_PAG_NAO_SE_APLICA_AO_MOV = 2720 'Sem parametro
'Cancelamento de Pagamento não se aplica a este tipo de Movimento.
Public Const ERRO_TIPO_MEIO_PAGAMENTO_NAO_CADASTRADO = 2721 'Sem parametro
'Tipo de Meio De Pagamento não está cadastrado.
Public Const NAO_EXISTE_PAG_PARA_SER_CANCELADO = 2722 'Sem parametro
'Não existe Pagamento com estes dados para ser cancelado
Public Const ERRO_MENSAGEM_NAO_INFORMADA = 2723 'Sem parâmetros
'A Mensagem não foi informada.
Public Const ERRO_CODIGO_NAO_INFORMADO = 2724 'Sem parâmetros
'O Código da Mensagem não foi informado
Public Const ERRO_MENSAGEM_COM_CARACTER_INICIAL_ERRADO = 2725 'Sem parametros
'Descrição de Mensagem não pode começar com este caracter.
Public Const ERRO_EXCLUSAO_MENSAGEM = 2726 'Parametro Código da Mensagem
'Houve um erro na exclusão da Mensagem %i do banco de dados.
Public Const ERRO_ATUALIZACAO_MENSAGEM = 2727 'Parâmetro Código da Mensagem
'Erro na atualização da Mensagem %i.
Public Const ERRO_INSERCAO_MENSAGEM = 2728 'Parâmetro Código da Mensagem
'Erro na Inserção da Mensagem %i.
Public Const ERRO_LEITURA_MENSAGEM1 = 2729 'Sem Parâmetros
'Erro na leitura da Tabela Mensagens
Public Const ERRO_TIPODOC_NAO_ENCONTRADO = 2730 'Parametro: Tipo.Text
'O Tipo De Documento %s não está cadastrado.
Public Const ERRO_LEITURA_CODIGO_PAIS = 2731 'Parâmetro Código do País
'Erro na leitura do País %s.
Public Const ERRO_CODIGO_PAIS_NAO_CADASTRADO = 2732 'Parâmetro Código do País
'País %s não está cadastrado.
Public Const ERRO_LEITURA_OC_COB = 2734 'parametro = codigo do cobrador
'Erro na leitura de ocorrências para o cobrador com código %d
Public Const ERRO_ATUALIZACAO_OC_COB = 2735 'parametro = codigo do cobrador
'Erro na atualização de ocorrências para o cobrador com código %d
Public Const ERRO_ATUALIZACAO_ANTECIPPAG = 2736 'parametros = lNumIntMovto
'Erro na atualização do Pagamento antecipado cujo Número de movimento é: %l.
Public Const ERRO_ATUALIZACAO_ANTECIPRECEB = 2737 'parametros = lNumIntMovto
'Erro na atualização do Recebimento antecipado cujo Número de movimento é: %l.
Public Const ERRO_SALDO_NEGATIVO = 2738   'Parametros: dValor, dSaldoNaoApropriado
'O valor %d não é permitido pois deixará o saldo negativo %d
Public Const ERRO_COBRADOR_PROPRIA_EMPRESA = 2739 'Parâmetro: objCobrador.iCodigo
'O Cobrador %i é da própria empresa.
Public Const ERRO_LEITURA_COBRADOR2 = 2740 'Sem Parametros
'Erro na leitura da tabela de Cobrador.
Public Const ERRO_FILIAL_NAO_ENCONTRADA2 = 2741 'Parametro: iFilial
'A Filial com codigo %i não foi encontrada.
Public Const ERRO_LEITURA_OUTROSPAGBAIXADOS = 2742 'Parâmetro: sSigla, lNumTitulo
'Erro na tentativa de leitura de Outro Pagamento do Tipo %s Número %l na tabela TitulosPagBaixados.
Public Const ERRO_LEITURA_OUTROSPAG1 = 2743 'Parâmetro: sSigla, lNumTitulo
'Erro na tentativa de leitura de Outro Pagamento do Tipo %s e Número %l na tabela TitulosPag.
Public Const ERRO_LEITURA_OUTROSPAGPEND = 2744 'Parâmetro: sSigla, lNumTitulo
'Erro na tentativa de leitura da Outro Pagamento do Tipo %s e Número %l na tabela TitulosPagPend.
Public Const ERRO_TITULOPAGAR_NAO_CADASTRADO = 2745 'lNumIntDoc
'O Título a Pagar com Número Interno %l não está cadastrado.
Public Const ERRO_TIPO_DOCUMENTO_NAO_OUTROSPAG = 2746 'Parametros: lNumTitulo, sSigla
'A Sigla %s do Título número %l não é utilizada em OutrosPag.
Public Const ERRO_TITULO_PENDENTE_MODIFICACAO = 2747 'Parâmetro: sSiglaDocumento, lNumTitulo
'Não é possível modificar Título a pagar do Tipo %s com número %l porque ele faz parte de um Lote Pendente.
Public Const ERRO_TITULO_BAIXADO_MODIFICACAO = 2748 'Parâmetro: sSiglaDocumento, lNumTitulo
'Não é possível modificar Título a pagar do Tipo %s com número %l porque ele está baixado.
Public Const ERRO_NUMERO_PARCELAS_TITULO_ALTERADO = 2749 'Parâmetros: sSigla, lNumTitulo, iNumParcelasTela, iNumParcelasBD
'Não é possível alterar o número de parcelas de Título a Pagar do Tipo %s com número %l que está lançado. Número de Parcelas da Tela: %i. Número de Parcelas do BD: %i.
Public Const ERRO_INSERCAO_TITULO_PAGAR = 2750 'Parametros : lNumTitulo, sSigla
'Erro na Tentativa de inserir Titulo número %l do Tipo %s na tabela TitulosPag.
Public Const ERRO_TITULO_PENDENTE_EXCLUSAO = 2751 'Parâmetros: 'lNumTitulo
'Não é possível excluir o Título a Pagar número %l porque ele faz parte de um Lote Pendente.
Public Const ERRO_TITULO_BAIXADO_EXCLUSAO = 2752 'Parâmetro: lNumTitulo
'Não é possível excluir o Titulo a Pagar número %l porque ele está Baixado
Public Const ERRO_TITULOPAGAR_NAO_CADASTRADO1 = 2753  'Parâmetro: lNumTitulo
'O Titulo a Pagar número %l não está cadastrado.
Public Const ERRO_EXCLUSAO_COMISSAO_BAIXADA = 2754 'sem parametros
'Não pode excluir documento associado a uma comissão marcada como "baixada"
Public Const ERRO_LEITURA_BAIXASREC1 = 2755 'Parâmetro: lNumIntDoc
'Erro na leitura da tabela BaixasRec com número interno do documento %l.
Public Const ERRO_DEBITOREC_VINCULADO_NFISCAL = 2756 'Parâmetro: lNumIntDoc
'Débito Com Cliente com número interno %l está vinculado à Nota Fiscal.
Public Const ERRO_LEITURA_PORTADOR1 = 2757 'sem parâmetro
'Erro na leitura da tabela Portador.
Public Const ERRO_BANCO_PORTADOR = 2758 'parâmetros iBanco
'O Banco %i está sendo usado na tabela Portador.
Public Const ERRO_NOME_CONTACORRENTEINTERNA_EXISTENTE = 2759 'Parâmetros: sNomeReduzido
'Já existe uma Conta Corrente Interna com Nome %s para a Empresa.
Public Const ERRO_MENSAGEM_ASSOCIADA_CLIENTE = 2761 'Parâmetro: iCodigo
'Não é permitida a exclusão da Mensagem %i porque está associada com Cliente.
Public Const ERRO_COBRADOR_NAO_CADASTRADO1 = 2764 'Parâmetro: sNomeReduzido
'O Cobrador %s não está cadastrado no Banco de Dados.
Public Const ERRO_INSTRUCAO_NAO_SELECIONADA = 2765 'Sem parâmetros
'Não foi possível selecionar a Instrução.
Public Const ERRO_DIAS_PROTESTO_NAO_PREENCHIDO = 2766 'Sem parâmetros
'É obrigatório o preenchimento do campo Dias de Protesto para o tipo de Instrução escolhida.
Public Const ERRO_LEITURA_TIPOSINSTRCOBRANCA1 = 2767 'Parâmetro: iCodigo
'Erro na leitura da tabela TiposInstrCobranca com Código %i.
Public Const ERRO_LEITURA_TIPOSDEOCORREMCOBR = 2768 'Sem parâmetros
'Erro na leitura da tabela TiposDeOcorRemCobr.
Public Const ERRO_LEITURA_TIPOSDEOCORREMCOBR1 = 2769 'Parâmetro: iCodOcorrencia
'Erro na leitura da tabela TiposDeOcorRemCobr com Código da Ocorrência %i.
Public Const ERRO_NUMERO_NAO_PREENCHIDO = 2770 'Sem parâmetros
'Preenchimento do Número é obrigatório.
Public Const ERRO_INSTRUCAO_NAO_CADASTRADA = 2771 'Parâmetro: iCodigo
'A Instrução com código %i não está cadastrada no Banco de dados.
Public Const ERRO_INSTRUCAO_NAO_CADASTRADA1 = 2772 'Parâmetro: Instrucao.Text
'A Instrução com descrição %s não está cadastrada no Banco de dados.
Public Const ERRO_OCORRENCIA_NAO_CADASTRADA = 2773 'Parâmetro: iCodigo
'A Ocorrência com código %i não está cadastrada no Banco de dados.
Public Const ERRO_OCORRENCIA_NAO_CADASTRADA1 = 2774 'Parâmetro: Ocorrencia.Text
'A Ocorrência com descrição %s não está cadastrada no Banco de dados.
Public Const ERRO_OCORR_REM_COBR_NAO_CADASTRADA = 2775 'Parâmetros: iNumSeqOcorr
'A Ocorrência com Sequencial %i desta Parcela não está cadastrada no Banco de dados.
Public Const ERRO_TITULORECEBER_NAO_CADASTRADO2 = 2776 'Parâmetros: iFilialEmpresa, lCliente, iFilial, sSiglaDocumento, lNumTitulo
'O Titulo Receber da Filial Empresa %i, Cliente %l , Filial %i, Sigla do Documento %s e Número do Título %l não está cadastrado no Banco de Dados.
Public Const ERRO_PARCELAREC_NUMINT_NAO_CADASTRADA = 2777 'Parâmetros: lNumIntTitulo, iNumParcela
'A Parcela Receber com número interno do Título %l e número da Parcela %i não está cadastrada no Banco de Dados.
Public Const ERRO_PARCELAREC_NUMINT_BAIXADA = 2778 'Parâmetros: lNumIntTitulo, iNumParcela
'A Parcela Receber com número interno do Título %l e número da Parcela %i está Baixada.
Public Const ERRO_INSTRUCAO_ENVIADA_BANCO = 2779 'Parâmetros: lNumIntTitulo, iNumParcela
'A Parcela Receber com número interno do Título %l e número da Parcela %i já foi enviada ao Banco.
Public Const ERRO_LOCK_OCORRENCIASREMPARCREC = 2780 'Parâmetros: lNumIntParc, iNumSeqOcorr
'Erro na tentativa de fazer "lock" na tabela OcorrenciasRemParcRec com número interno da Parcela %l e Sequencial %i.
Public Const ERRO_PARCELAREC_COBR_ELETRONICA = 2781 'Sem parâmetros
'A Parcela não é do tipo Cobrança Eletrônica.
Public Const ERRO_ATUALIZACAO_OCORRENCIASREMPARCREC = 2782 'Parâmetros: lNumIntParc, iNumSeqOcorr
' Erro na atualização da Ocorrência com número interno da Parcela %l e Sequencial %i.
Public Const ERRO_INCLUSAO_OCORRENCIASREMPARCREC = 2783 'Parâmetro: lNumIntParc, iNumSeqOcorr
'Erro na tentativa de inserir um registro na tabela OcorrenciasRemParcRec com número interno da Parcela %l e Sequencial %i.
Public Const ERRO_PARCELAREC_OCORRENCIA_BORDERO_COBRANCA = 2784 'Parâmetros: lNumIntParc, iNumSeqOcorr
'A Ocorrência com número interno da Parcela %l e Sequencial %i faz parte de um Bordero de Cobrança.
Public Const ERRO_ATUALIZACAO_PARCELAREC_OCORRENCIA_BORDERO_COBRANCA = 2785 'Parâmetros: lNumIntParc, iNumSeqOcorr
'Não é permitido a atualização da Ocorrência com número interno da Parcela %l e Sequencial %i porque faz parte de um Borderô de Cobrança.
Public Const ERRO_ATUALIZACAO_OCORRENCIA_INSTRUCAO = 2786 'Parâmetros: lNumParc, iNumSeqOcorr
'Não é permitido a atualização dos campos Ocorrência e Instrução1 desta Parcela com Sequencial %i.
Public Const ERRO_PARCELAREC_NAO_CADASTRADA = 2787 'Parâmetro: lNumIntParc
'A Parcela com número interno %l não está cadastrada no Banco de Dados.
Public Const ERRO_EXCLUSAO_OCORRENCIASREMPARCREC = 2788 'Parâmetros: lNumIntParc, iNumSeqOcorr
'Erro na tentativa de excluir um registro da tabela OcorrenciasRemParcRec com número interno da Parcela %l e Sequencial %i.
Public Const ERRO_OCORRENCIA_COBRANCA_NAO_PREENCHIDA = 2789 'Sem parâmetros
'A Ocorrência deve ser preenchida.
Public Const ERRO_INSTRUCAO_COBRANCA_NAO_PREENCHIDA = 2790 'Sem parâmetros
'A Instrução deve ser preenchida.
Public Const ERRO_LEITURA_TIPOSINSTRCOBRANCA2 = 2791 'Parâmetro: sDescricao
'Erro na leitura da tabela TiposInstrCobranca com Descrição %s.
Public Const ERRO_PARCELA_NAO_INFORMADA1 = 2792 'Sem parâmetros
'Os campos Cliente, Filial, Tipo, Número, Parcela e Sequencial devem ser preenchidos.
Public Const ERRO_INSTRUCAO_COBRANCA_NAO_CADASTRADA = 2793 'Sem parâmetros
'A Parcela informada não está cadastrada no Banco de Dados.
Public Const ERRO_ATUALIZACAO_OCORRENCIA = 2794 'Parâmetros: lNumParc, iNumSeqOcorr
'Erro na tentativa de atualiza um registro da tabela OcorrenciasRemParcRec desta Parcela com Sequencial %i.
Public Const ERRO_ANTECIPPAG_NAO_CARREGADO = 2795
'O Adiantamento deve ser carregado através de sua escolha em uma tela de lista. Aperte o botão "Adiantamentos a Filial do Fornecedor..." e escolha o Adiantamento a excluir.
Public Const ERRO_ANTECIPRECEB_NAO_CARREGADO = 2796
'O Adiantamento deve ser carregado através de sua escolha em uma tela de lista. Aperte o botão "Adiantamentos à Filial do Cliente..." e escolha o Adiantamento a excluir.
Public Const ERRO_VALOR_CR_MENOR_VALOR_BAIXAR = 2797 'dValorCredito, dValorTotalBaixar
'O Valor do Saldo do Crédito, %d,  é menor que o Valor Total a Baixar, %d.
Public Const ERRO_VALOR_PA_MENOR_VALOR_BAIXAR = 2798 'dValorPA, dValorTotalBaixar
'O Valor do Saldo do Pagamento Antecipado, %d,  é menor que o Valor Total a Baixar, %d.
Public Const ERRO_PARCELAS_SUPERIOR_NUM_MAX_PARCELAS_BAIXA = 2799 'Sem parâmetros
'O número de parcelas ultrapassou o limite permitido.
Public Const ERRO_CREDITOS_SUPERIOR_NUM_MAX_CREDITOS = 2800 'Sem parâmetros
'O número de créditos ultrapassou o limite permitido.
Public Const ERRO_PAGTOSANTECIPADOS_SUPERIOR_NUM_MAX_PAGTOS_ANTECIPADOS = 2801 'Sem parâmetros
'O número de pagamentos antecipados ultrapassou o limite permitido.
Public Const ERRO_PORTADOR_NAO_CADASTRADO2 = 2802 'Parâmetro: sPortador
'Portador %s não está cadastrado.
Public Const ERRO_LOCK_PAGTOSANTECIPADOS = 2803 'Parametro: lNumMovto
'Erro na tentativa de "lock" da tabela PagtosAntecipados. NumMovto=%l.
Public Const ERRO_ATRIBUTOS_PAGTOANTECIPADO_MUDARAM = 2804 'Parametros: dSaldoNaoApropriadoBD, dSaldoNaoApropriado, lFornecedorBD, lFornecedor, iFilialBD, iFilial
'Atributos de Pagto Antecipado mudaram. Saldo: BD %d Tela %d, Fornecedor: BD %l Tela %l, Filial: BD %i Tela %i.
Public Const ERRO_ATUALIZACAO_PAGTOANTECIPADO_SALDO = 2805 'Parametro: lNumMovto
'Erro na atualização do SaldoNaoApropriado do Pagamento Antecipado com NumMovto=%l.
Public Const ERRO_CREDITOPAGFORN_NAO_CADASTRADO = 2806 'Parametros: lFornecedor, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Crédito Pagar não cadastrado. Fornecedor=%l, Filial=%i, SiglaDocumento=%s, NumTitulo=%l, DataEmissao=%dt.
Public Const ERRO_LOCK_CREDITOPAGFORN = 2807 'Parametro: lNumIntDoc
'Erro na tentativa de "lock" da tabela CreditosPagForn. Número Interno = %l.
Public Const ERRO_SALDO_CREDITOPAGFORN_MUDOU = 2808 'Parametros: dSaldoBD, dSaldo
'Saldo de Crédito Pagar mudou. Saldo no BD: %d. Saldo na Tela: %d.
Public Const ERRO_ATUALIZACAO_CREDITOPAGFORN_SALDO = 2809 'Parametro: lNumIntDoc
'Erro na atualização na tabela CreditosPagForn. Número Interno = %l.
Public Const ERRO_TITULOINIC_MAIOR_TITULOFIM = 2810 'Sem parâmetro
'Título inicial não pode ser maior que o Título final.
Public Const ERRO_DATAEMISSAO_INICIAL_MAIOR = 2811 'Sem parâmetro
'Data Emissao Inicial não pode ser maior que a Data Emissao Final.
Public Const ERRO_DATAVENCIMENTO_INICIAL_MAIOR = 2812 'Sem parâmetro
'Data Vencimento Inicial não pode ser maior que a Data Vencimento Final.
Public Const ERRO_VALORBAIXAR_MAIOR_SALDO = 2813 'Sem parâmetro
'Valor a Baixar não pode ser maior que o Saldo.
Public Const ERRO_VALORBAIXAR_MENOR_IGUAL_DESCONTO = 2814 'Sem parâmetro
'Valor a Baixar não pode ser menor ou igual ao Desconto.
Public Const ERRO_MULTA_INCOMPATIVEL_DATABAIXA_DATAVENC = 2815 'Sem parâmetro
'Multa incompatível com Data de Baixa menor ou igual à Data de Vencimento.
Public Const ERRO_LEITURA_CONTASCORRENTESINTERNAS2 = 2816 'Parâmetro: sNomeReduzido
'Erro na tentativa de leitura da Conta Corrente %s na tabela ContasCorrentesInternas.
Public Const ERRO_DATA_BAIXA_SEM_PREENCHIMENTO = 2817
'A Data Baixa deve estar preenchida.
Public Const ERRO_VALORBAIXAR_PARCELAS_NAO_INFORMADO = 2818
'O Valor a Baixar das parecelas selecionadas deve estar preenchido.
Public Const ERRO_CONTA_PAGDINHEIRO_NAO_INFORMADA = 2819 'Sem parâmetro
'Para Pagamento em Dinheiro a Conta Corrente deve ser informada.
Public Const ERRO_VALOR_PAGDINHEIRO_NAO_INFORMADO = 2820 'Sem parâmetro
'Para Pagamento em Dinheiro o Valor do Pagamento deve ser informado.
Public Const ERRO_VALORPAG_DIFERENTE_TOTALPARCELAS = 2821 'Parâmetros : sValorPago, dTotalParcelas
'O Valor do Pagamento %s não coincide com o Valor a Pagar %d das Parcelas selecionadas.
Public Const ERRO_VALORBAIXAR_PA_NAO_INFORMADO = 2822 'Sem parâmetro
'O Valor a Baixar do Pagamento Antecipado deve estar preenchido.
Public Const ERRO_PARCELA_PA_NAO_MARCADA = 2823 'Sem parâmetro
'Pelo menos uma Parcela do Pagamento Antecipado deve estar marcada.
Public Const ERRO_VALORUSAR_PA_DIFERENTE_TOTALPARCELAS = 2824 'Parâmetros : sValBaixarPA, dTotalParcelas
'O Valor a Usar do Pagamento Antecipado %s não coincide com o Valor a Pagar %d das Parcelas selecionadas.
Public Const ERRO_VALORBAIXAR_CR_NAO_INFORMADO = 2825 'Sem parâmetro
'O Valor a Baixar do Crédito deve estar preenchido.
Public Const ERRO_PARCELA_CR_NAO_MARCADA = 2826 'Sem parâmetro
'Pelo menos uma Parcela dos Créditos deve estar marcada.
Public Const ERRO_VALORUSAR_CR_DIFERENTE_TOTALPARCELAS = 2827 'Parâmetros : sValBaixarCR, dTotalParcelas
'O Valor a Usar dos Créditos %s não coincide com o Valor a Pagar %d das Parcelas selecionadas.
Public Const ERRO_LEITURA_TIPOMEIOPAGTO2 = 2828 'Parâmetro sDescricao
'Erro na leitura do Tipo de Pagamento %s da tabela TipoMeioPagto.
Public Const ERRO_TITULO_NAO_MARCADO = 2829 'Sem parâmetro
'Pelo menos um Título deve estar marcado para esta operação.
Public Const ERRO_TITULOS_NAO_MARCADOS = 2830 'Sem parâmetro
'Pelo menos dois Títulos devem estar marcado para esta operação.
Public Const ERRO_DEBITOS_SUPERIOR_NUM_MAX_DEBITOS = 2831 'Sem parâmetros
'O número de créditos ultrapassou o limite permitido.
Public Const ERRO_RECEBANTECIPADOS_SUPERIOR_NUM_MAX_RECEB_ANTECIPADOS = 2832 'Sem parâmetros
'O número de pagamentos antecipados ultrapassou o limite permitido.
Public Const ERRO_ATRIBUTOS_RECEBANTECIPADO_MUDARAM = 2833 'Parametros: dSaldoNaoApropriadoBD, dSaldoNaoApropriado, lClienteBD, lCliente, iFilialBD, iFilial
'Atributos de Recebimento Antecipado mudaram. Saldo: BD %d Tela %d, Cliente: BD %l Tela %l, Filial: BD %i Tela %i.
Public Const ERRO_ATUALIZACAO_RECEBANTECIPADO_SALDO = 2834 'Parametro: lNumMovto
'Erro na atualização do SaldoNaoApropriado do Recebimento Antecipado com NumMovto=%l.
Public Const ERRO_DEBITORECCLI_NAO_CADASTRADO = 2835 'Parametros: lCliente, iFilial, sSiglaDocumento, lNumTitulo, dtDataEmissao
'Débito Receber não cadastrado. Cliente=%l, Filial=%i, SiglaDocumento=%s, NumTitulo=%l, DataEmissao=%dt.
Public Const ERRO_SALDO_DEBITORECCLI_MUDOU = 2836 'Parametros: dSaldoBD, dSaldo
'Saldo de Débito Receber mudou. Saldo no BD: %d. Saldo na Tela: %d.
Public Const ERRO_ATUALIZACAO_DEBITOSRECCLI_SALDO = 2837 'Parametro: lNumIntDoc
'Erro na atualização na tabela DebitosRecCli. Número Interno = %l.
Public Const ERRO_DATA_CREDITO_SEM_PREENCHIMENTO = 2838
'A Data Crédito deve estar preenchida.
Public Const ERRO_CONTA_RECDINHEIRO_NAO_INFORMADA = 2839 'Sem parâmetro
'Para Recebimento em Dinheiro a Conta Corrente deve ser informada.
Public Const ERRO_VALOR_RECDINHEIRO_NAO_INFORMADO = 2840 'Sem parâmetro
'Para Recebimento em Dinheiro o Valor deve ser informado.
Public Const ERRO_VALORREC_DIFERENTE_TOTALPARCELAS = 2841 'Parâmetros : sValReceber, dTotalParcelas
'O Valor do Recebimento %s não coincide com o Valor a Receber %d das Parcelas selecionadas.
Public Const ERRO_VALORBAIXAR_RA_NAO_INFORMADO = 2842 'Sem parâmetro
'O Valor a Baixar do Recebimento Antecipado deve estar preenchido.
Public Const ERRO_PARCELA_RA_NAO_MARCADA = 2843 'Sem parâmetro
'Pelo menos uma Parcela do Recebimento Antecipado deve estar marcada.
Public Const ERRO_VALORUSAR_RA_DIFERENTE_TOTALPARCELAS = 2844 'Parâmetros : sValorBaixarRA, dTotalParcelas
'O Valor a Usar do Recebimento Antecipado %s não coincide com o Valor a Receber %d das Parcelas selecionadas.
Public Const ERRO_VALORBAIXAR_DB_NAO_INFORMADO = 2845 'Sem parâmetro
'O Valor a Baixar do Débito deve estar preenchido.
Public Const ERRO_PARCELA_DB_NAO_MARCADA = 2846 'Sem parâmetro
'Pelo menos uma Parcela dos Débitos deve estar marcada.
Public Const ERRO_VALORUSAR_DB_DIFERENTE_TOTALPARCELAS = 2847 'Parâmetros : sValorBaixarDB, dTotalParcelas
'O Valor a Usar dos Débitos %s não coincide com o Valor a Receber %d das Parcelas selecionadas.
Public Const ERRO_RECEBANTECIPADO_NAO_CADASTRADO = 2848 'Parametro: lNumeroMovimento
'Recebimento Antecipado com NumMovto=%l não está cadastrado.
Public Const ERRO_CAMPOS_BAIXAPARCREC_NAO_PREENCHIDOS = 2849 'Sem parâmetros
'Para trazer uma Parcela é obrigatório que estejam preenchidos os campos Cliente, Filial, Tipo, Número, Data Emissão , Parcela e Sequencial da tela.
Public Const ERRO_BAIXA_PARCELA_CANCELADA = 2850 'Parâmetros: iNumParcela, iSequencialBaixa
'A Baixa da Parcela %i e Sequencial %i já está cancelada.
Public Const ERRO_EXCLUSAO_PARCELAS_RECEBER_BAIXADAS = 2851
'Erra na tentativa de excluir registro da tabela de parcelas a receber baixadas.
Public Const ERRO_PARCELA_RECEBER_BAIXADA_INEXISTENTE = 2852
'Parcela a receber baixada não encontrada.
Public Const ERRO_UNLOCK_TITULOS_REC_BAIXADOS = 2853
'Erro na tentativa de desfazer o lock na tabela de Titulos a Receber Baixados.
Public Const ERRO_RECEBIMENTO_ANTECIPADO_INEXISTENTE = 2854
'Recebimento antecipado não cadastrado.
Public Const ERRO_MODIFICACAO_DEBITO_REC_CLI = 2855
'Erro na tentativa de modificar a tabela de Débitos a Receber.
Public Const ERRO_DEBITO_REC_CLI_EXCLUIDO = 2856
'Esse Débito está excluído.
Public Const ERRO_MODIFICACAO_RECEBIMENTO_ANTECIPADO = 2857
'Erro na tentativa de modificar a tabela de Recebimentos Antecipados.
Public Const ERRO_RECEBIMENTO_ANTECIPADO_EXCLUIDO = 2858
'O recebimento antecipado já havia sido excluído.
Public Const ERRO_MODIFICACAO_BAIXAPARCREC = 2859
'Erro na tentativa de modificar a tabela de Baixas de Parcelas a Receber.
Public Const ERRO_EXCLUSAO_TITULOS_RECEBER_BAIXADOS = 2860
'Erro na exclusão do Título a Receber Baixado.
Public Const ERRO_LOCK_BAIXAPARCREC = 2861
'Não conseguiu fazer o lock da Baixa de Parcela a Receber.
Public Const ERRO_TITULO_RECEBER_BAIXADO_INEXISTENTE = 2863
'O Título a Receber Baixado não está cadastrado.
Public Const ERRO_LEITURA_BAIXAPARCREC = 2864
'Erro na leitura da tabela de Baixas de Parcelas a Receber
Public Const ERRO_LEITURA_BAIXAREC = 2865
'Erro na leitura da tabela de Baixas a Receber.
Public Const ERRO_UNLOCK_PARCELAS_REC_BAIXADAS = 2866
'Erro na tentativa de desfazer o lock na tabela de Parcelas a Receber Baixadas.
Public Const ERRO_LOCK_TITULOS_REC_BAIXADOS = 2867
'Não conseguiu fazer o lock de Títulos a Receber Baixados.
Public Const ERRO_UNLOCK_BAIXAPARCREC = 2868
'Erro na tentativa de desfazer o lock na tabela de Baixas de Parcelas a receber.
Public Const ERRO_BAIXAREC_INEXISTENTE = 2869
'A Baixa a Receber não está cadastrada.
Public Const ERRO_EXCLUSAO_BAIXAREC = 2870
'Erro na exclusão da Baixa a Receber.
Public Const ERRO_BAIXAPARCREC_INEXISTENTE = 2871
'A Baixa não está cadastrada.
Public Const ERRO_LOCK_PARCELAS_REC_BAIXADAS = 2872
'Não conseguiu fazer o lock de Parcelas a Receber Baixadas.
Public Const ERRO_BAIXAREC_EXCLUIDA = 2873
'A baixa a receber já havia sido cancelada anteriormente.
Public Const ERRO_BAIXAPARCREC_EXCLUIDA = 2874
'A baixa a receber da parcela já havia sido cancelada anteriormente.
Public Const ERRO_PARCELA_JA_EXISTENTE = 2875 'Parametro : iLinha
'A parcela informada é a mesma da %i linha do Grid.
Public Const ERRO_DATA_CREDITO_NAO_PREENCHIDA = 2876
'A Data de Crédito deve ser preenchida.
Public Const ERRO_TOTAL_RECEBIDO_NAO_PREENCHIDO = 2877
'O Total Recebido deve ser informado.
Public Const ERRO_VALOR_DESCONTO_PARCELA_SUPERIOR_SOMA_VALOR = 2878
'O Valor do Desconto da parcela na linha %i é superior ao valor recebido.
Public Const ERRO_VALOR_RECEBIDO_PARCELA_NAO_PREENCHIDO = 2879
'O valor recebido na linha %i não foi informado.
Public Const ERRO_NUMPARCELA_NAO_INFORMADO = 2880
'O numero da parcela não foi informado na linha %i.
Public Const ERRO_QUANTIDADE_INFORMADA_DIFERENTE_GRID = 2881
'A quantidade de Parcelas no Grid é diferente da quantidade informada na tela.
Public Const ERRO_TOTALRECEBIDO_PARCELAS_DIFERENTE = 2882
'A soma dos valores recebidos no Grid é diferente do valor recebido informado a tela.
Public Const ERRO_DATADEPOSITO_MENOR_DATAEMISSAO = 2883
'A Data de deposito deve ser maior ou igual à Data de Emissão.
Public Const ERRO_ATUALIZACAO_OCORRCOBR = 2884 'sem parametro
'Erro na tentativa de alterar o registro de uma ocorrência de cobrança
Public Const ERRO_ALTERACAO_BORDERO_COBRANCA = 2885 'Parâmetro: lNumBordero
'Erro na tentativa de alterar o Borderô de Número %l na Tabela de Borderôs de Cobrança.
Public Const ERRO_LOCK_BORDERO_COBRANCA = 2886 'Parâmetro: lNumBordero
'Erro na tentativa de fazer "lock" na tabela BorderosCobranca para Número %l .
Public Const ERRO_BORDERO_COBRANCA_NAO_CADASTRADO = 2887 'Parâmetro: lNumBordero
'O Bordero de Cobrança de número %l não está cadastrado no Banco de Dados.
Public Const ERRO_BORDERO_COBRANCA_EXCLUIDO = 2888 'Parâmetro: lNumBordero
'Borderô de Número %l está excluído.
Public Const ERRO_BORDERO_COBRANCA_SEM_OCORRENCIAS = 2889 'Parâmetro: lNumBordero
'Não foram enconradas ocorrências para o Bordero %l no Banco de Dados
Public Const ERRO_PARCELA_RECEBER_BAIXADA = 2891 'Parâmetro: lNumIntParc
'A Parcela a Receber com o Número Interno %l já está baixada.
Public Const ERRO_OCORRENCIA_DIFERENTE_PARCELA = 2892
'O Sequencial da ocorrência não corresponde ao indicado na parcela.
Public Const ERRO_SEM_PARCELA_SELECIONADA = 2893
'Nenhuma parcela foi selecionada.
Public Const ERRO_LEITURA_PAGTO_BORDERO = 2894
'Erro na leitura de pagamento efetuado por borderô
Public Const ERRO_TITULO_INFORMADO_SEM_FILIAL = 2895 'Sem parâmetro
'O Título só pode ser preenchido se a Filial tiver sido informada.
Public Const ERRO_PARCELA_INFORMADA_SEM_TITULO = 2896 'Sem parâmetro
'A Parcela só pode ser preenchida se o Título tiver sido informado.
Public Const ERRO_VALORPARCELA_MENOR_DESCONTO = 2897 'Sem parâmetro
'O desconto ultrapassou o valor total da parcela, com juros e multa.
Public Const ERRO_PORTADORES_DIFERENTES = 2898 'Sem parâmetros
'Um mesmo cheque não pode pagar Títulos com Portadores diferentes.
Public Const ERRO_TITULOS_FORN_DIF_NAO_COBRBANC = 2899 'Sem parâmetro
'Só é permitido um cheque para fornecedores diferentes quando for cobrança bancária com portador definido.
Public Const ERRO_FILIALCLIENTE_INEXISTENTE1 = 2900 'Parâmetros: iCodCliente, sCliente
'A Filial %i do Cliente %s não está cadastrada no Banco de Dados.
Public Const ERRO_CLIENTE_ULTRAPASSOU_NUMMAXPARCELAS = 2901 'Sem parâmetros
'O Cliente ultrapassou o número máximo de Parcelas.
Public Const ERRO_PARCELA_GRID_NAO_SELECIONADA = 2902
'Não há parcela selecionada no grid de parcelas.
Public Const ERRO_SUBIR_PRIM_CHEQUE = 2903 'Sem parâmetro
'Não pode subir o primeiro cheque
Public Const ERRO_DESCER_CHEQUE_ZERO = 2904 'Sem parâmetro
'Não pode descer parcela que não esteja marcada para imprimir
Public Const ERRO_TITULOS_DE_UM_CHEQUE = 2905 'Sem parâmetro
'Selecione parcelas de um cheque
Public Const ERRO_NAO_PODE_AGRUPAR_SEQ_ZERO = 2906 'Sem parâmetro
'Não pode selecionar título não marcado para emitir para esta operação.
Public Const ERRO_TITULOS_TIPO_COBR_DIFERENTE = 2907 'Sem parâmetro
'Para esta operação os títulos devem ter o mesmo tipo de cobrança.
Public Const ERRO_TITULOS_PORTADOR_DIFERENTE = 2908 'Sem parâmetro
'Para esta operação os títulos devem ter o mesmo portador.
Public Const ERRO_TITULOS_CHEQUE_DIFERENTE = 2909 'Sem parâmetro
'Para esta operação os títulos devem pertencer ao mesmo cheque.
Public Const ERRO_NUMCHEQUE_INVALIDO = 2910 'Parâmetros: sCheque e iNumCheque
'O número %s de Cheque é inválido. Deve ser maior ou igual a 1 e menor que %i
Public Const ERRO_NUMCHEQUE_NAO_PREENCHIDO = 2911 'Sem parâmetro
'O número do Cheque deve ser informado.
Public Const ERRO_NUM_CHEQUES_MARCADOS = 2912 'Sem parâmetro
'Deve haver 1 Cheque marcado para esta operação.
Public Const ERRO_ESTADO_NAO_PREENCHIDO = 2913 'Sem parâmetros
'O preenchimento do Estado correspondente ao Endereço é obrigatório.
Public Const ERRO_COBRADOR_NAO_PERTENCE_FILIAL = 2914 'Parametro: iCobrador, iFilial
'O Cobrador %i não pertence à Filial selecionada. Filial: %i.
Public Const ERRO_VENDEDOR_VAZIO = 2915 'Sem parametros
'Nome reduzido do Vendedor selecionado na listbox está vazio
Public Const ERRO_DATA_FINAL_MAIOR = 2917 'Sem parametros
'Data final não pode ser maior que Data de baixa
Public Const ERRO_EXTRATO_NAO_INFORMADO = 2918 'Sem parâmetro
'O extrato deve ser informado.
Public Const ERRO_USADA_EM_NAO_PREENCHIDA = 2919 'Sem Parametros
'É obrigatório informar onde será utilizada esta Condicão de Pagamento.
Public Const ERRO_VALORES_COMISSAO_BASE = 2920 'Parâmetros:Valor Base e Valor da Comissão
'Valor Base %d é menor que o Valor da Comissão %d
Public Const ERRO_TECLAR_BOTAO_TRAZER = 2921 'Sem Parametros
'Depois de preencher os campos obrigatorios, aperte o Botão Trazer.
Public Const ERRO_BANCO_NAO_PREENCHIDO = 2922 'Sem parâmetro
'O nome do banco deve ser informado.
Public Const ERRO_ARQUIVO_NAO_PREENCHIDO = 2923 'Sem parâmetro
'O nome do arquivo deve ser especificado.
Public Const ERRO_NUM_MAX_NFS_SELEC_EXCEDIDO = 2924 'parametro: %d limite de nfs p/fatura
'Não pode associar mais de %d notas fiscais a uma fatura
Public Const ERRO_BANCO_NAO_POSITIVO = 2925 'parametro sBanco
'O numero do Banco %s não é positivo.
Public Const ERRO_MOVIMENTO_DINHEIRO_SEM_SEQUENCIAL = 2926
'O movimento com Tipo de Pagamento "Dinheiro" deve ser carregado na tela através da tela de browse.
Public Const ERRO_DATAS_DESCONTOS_DESORDENADAS = 2927 'iParcela
'As datas de descontos para a parcela %i não estão ordenadas.
Public Const ERRO_PARCELAREC_COBR_MANUAL = 2928 'Sem parâmetros
'A Parcela não é do tipo Cobrança Manual.
Public Const ERRO_CONTACORRENTE_FILIAL_DIFERENTE = 2929
'Esta Conta Corrente pertence a outra filial.
Public Const ERRO_VALOR_DEPOSITO_NAO_PREENCHIDO = 2930 'Sem Parâmetros
'O valor do Depósito deve estar preenchido
Public Const ERRO_VALOR_RESGATE_NAO_PREENCHIDO = 2931 'Sem Parametros
'O Valor do Resgate não foi preenchido.
Public Const ERRO_SALDO_ZERO = 2932 'Sem Parametros
'Não pode realizar resgate com o saldo igual a zero.
Public Const ERRO_VALOR_SAQUE_NAO_PREENCHIDO = 2933 'Sem Parâmetros
'O valor do Saque deve estar preenchido.
Public Const ERRO_HISTMOVCTA_NAO_CADASTRADO1 = 2934 'Parâmetro: Historico.Text
'O Histórico de Movimentação de Conta %s não está cadastrado.
Public Const ERRO_DATA_BAIXA_INICIAL_MAIOR = 2935
'A data da Baixa Inicial não pode ser maior que a Final. 'Sem Parametros
Public Const ERRO_UMA_DATA_NAO_PREENCHIDA = 2936
'Você deve preencher pelo menos um dos dois pares de datas. Ou Baixa ou Digitação.
Public Const ERRO_DATA_DIGITACAO_BAIXA_INICIAL_MAIOR = 2937
'A data da digitacao da baixa Inicial não pode ser maior que a Final. 'Sem Parametros
Public Const ERRO_FORNECEDOR_INICIAL_MAIOR = 2940
'O Fornecedor Inicial é maior que o final.'
Public Const ERRO_FORNECEDOR_NAO_CADASTRADO_2 = 2941
'O Fornecedor não está cadastrado.'
Public Const ERRO_VENDEDOR_NAO_CADASTRADO_2 = 2942
'O Vendedor não está cadastrado.'
Public Const ERRO_VALOR_INVALIDO1 = 2943 'Sem parametros
'Número de Devedores é invalido
Public Const ERRO_DATA_EMISSAO_INICIAL_MAIOR = 2944
'A data de Emissao Inicial não pode ser maior que a Final. 'Sem Parametros
Public Const ERRO_DATA_VENCTO_INICIAL_MAIOR = 2946
'A data do vencimento Inicial não pode ser maior que a Final. 'Sem Parametros
Public Const ERRO_CHEQUEDET_GRAVACAO = 2947
'Erro na gravação de um cheque de pagamento de título
Public Const ERRO_DEBITO_REC_SALDO_NEGATIVO = 2948
'O saldo do crédito do cliente ficaria negativo
Public Const ERRO_LEITURA_VALBAIXAPERDA = 2949 'Sem parametros
'Erro na leitura do valor baixado por perda de uma parcela a receber
Public Const ERRO_CHEQUEPRE_SEM_PARCELA = 2950 'Sem parametros
'Não encontrou parcela associada ao cheque pré-datado.
Public Const ERRO_LEITURA_COMISSOES2 = 2951 'Sem parametros
'Erro de leitura de comissões de uma parcela
Public Const ERRO_ATUALIZACAO_COMISSOES2 = 2952 'Sem parametros
'Erro na gravação de comissões de uma parcela
Public Const ERRO_LIBERACAO_COMISSAO_STATUS = 2953 'Sem parametros
'O status atual da comissão não é compatível com a sua liberação.
Public Const ERRO_LEITURA_COBRADORES_BANCO = 2954 'Parâmetro: iCodBanco
'Erro na tentativa de ler cobradores associados com o Banco %i.
Public Const ERRO_LEITURA_PARCPAG_BANCO = 2955 'Parâmetro: iCodBanco
'Erro na tentativa de ler a tabela ParcelasPag com Banco %i.
Public Const ERRO_LEITURA_PARCPAGBAIXADAS_BANCO = 2956 'Parâmetro: iCodBanco
'Erro na tentativa de ler a tabela ParcelasPagBaixadas com Banco %i.
Public Const ERRO_BANCO_ASSOCIADO_PARCELAPAG = 2957 'Parâmetros: iCodBanco, lNumIntDoc
'Não é permitido excluir Banco %i porque está associado à Parcela com número interno %l.
Public Const ERRO_BANCO_ASSOCIADO_PARCELAPAGBAIXADA = 2958 'Parâmetros: iCodBanco, lNumIntDoc
'Não é permitido excluir Banco %i porque está associado à Parcela Baixada com número interno %l.
Public Const ERRO_BANCO_ASSOCIADO_COBRADOR = 2959 'Parâmetros: iCodBanco, iCobrador
'Não é permitido excluir Banco %i porque está associado ao Cobrador %i.
Public Const ERRO_LEITURA_COMISSOES3 = 2960
'Erro na leitura da tabela de Comissoes.
Public Const ERRO_INSERCAO_TESCONFIG = 2961 'Parâmetros: Codigo, FilialEmpresa
'Ocorreu um erro na inserção de um registro na tabela TESConfig. Codigo = %s, Filial = %i.
Public Const ERRO_INSERCAO_FERIADOS_FILIAIS = 2962 'Parâmetros: FilialEmpresa
'Erro na inserção dos feriados na filial %s.
Public Const ERRO_LEITURA_TIPO_OCORR_COBR = 2963 'Parametro: Codigo do Tipo
'Erro na leitura do Tipo de Ocorrência de Cobrança com código %s.
Public Const ERRO_SALDO_RA_MENOR_TOTALRECEBER = 2964 'sem parametros
'O saldo do adiantamento não é suficiente para as baixas selecionadas
Public Const ERRO_SALDO_DB_MENOR_TOTALRECEBER = 2965 'sem parametros
'O saldo do crédito/devolução não é suficiente para as baixas selecionadas
Public Const ERRO_PARCELAREC_TRANSF_CHEQUEPRE = 2966 'sem parametros
'Não pode transferir parcela com cheque pré-datado
Public Const ERRO_VIA_TRANSPORTE_NAO_PREENCHIDO = 2967 'Sem Parametros
'Erro é obrigatório o preenchimento da Via de Transporte.
Public Const ERRO_TIPO_INICIAL_MAIOR = 2968 'Sem Parametros
'O tipo inicial não pode ser maior que o tipo final.
Public Const ERRO_LOCK_BAIXAREC = 2969
'Não conseguiu fazer o lock da Baixa a Receber.
Public Const ERRO_REGIAO_VENDA_RELACIONADA_COM_PREVVENDA = 2970 'Sem Parametro
'Não é possível excluir Região de Venda relacionada com PrevVenda.
Public Const ERRO_DEBITO_REC_SALDO_MAIOR_ORIGINAL = 2971 'Sem Parametro
'O saldo do crédito do cliente ficaria maior que seu valor original
Public Const ERRO_LOCK_CHEQUEPRE = 2972 'sem parametros
'Erro na tentativa de fazer 'lock' na tabela de ChequesPre
Public Const ERRO_VALOR_CHEQUEPRE_ALTERADO = 2973 'sem parametros
'O valor do cheque pré datado foi alterado
Public Const ERRO_CHEQUEPRE_JA_DEPOSITADO = 2974 'sem parametros
'O cheque pré datado já foi depositado através de outro borderô
Public Const ERRO_DETPAG_CPO_NAO_PREENCHIDO = 2975 'sem parametros
'Preencha os campos Fornecedor, Filial, Número do Título e da Parcela
Public Const ERRO_VALORINSS_MAIOR = 2976 'Parâmetros: sValorINSS, sValor
'Valor do INSS não pode ser maior do que o Valor do Titulo
Public Const ERRO_DATA_CHQPRE_DIFERENTE = 2977 'Parametros: numero da parcela
'A data para depósito do cheque pré datado associado à parcela %s não confere com o vencimento da mesma
Public Const ERRO_VENDEDOR_NAO_FORNECIDO = 2978 'Sem Parametros
'Não há vendedor selecionado
Public Const EXCLUSAO_RELOPCOMISVEND = 2979 'Sem parametros
'Confirma a exclusão da Opção de Relatório ?
Public Const ERRO_LINHA_GRID_CATEGCLI_INCOMPLETA = 2980 'sem parametros
'Há uma linha do grid de categorias de clientes que não está completa.
Public Const ERRO_LEITURA_NOTAS_FISCAIS_EXT_BAIXADAS = 2981 'Sem parametro
'Erro na leitura da tabela de notas fiscais baixadas.
Public Const ERRO_EXCLUSAO_NOTAS_FISCAIS_EXT_BAIXADAS = 2982 'Sem parametro
'Erro na exclusão na tabela de notas fiscais baixadas.
Public Const ERRO_INSERCAO_NOTAS_FISCAIS_EXT = 2983 'Sem parametro
'Erro na inserção na tabela de notas fiscais baixadas.
Public Const ERRO_PREENCHA_CAMPOS_OBRIGATORIOS = 2984
'Todos os campos devem estar preenchidos.
Public Const ERRO_VALOR_PAGAR_NEGATIVO = 2985 'Sem parametros
'O valor a pagar não pode ser negativo
Public Const ERRO_MULTA_JUROS_DATABAIXA_VAZIA = 2986 'Sem parametros
'Para digitar juros ou multa a data da baixa precisa estar prenchida
Public Const ERRO_JUROS_INCOMPATIVEL_DATABAIXA_DATAVENC = 2987 'Sem parametros
'Juros incompatível com Data de Baixa menor ou igual à Data de Vencimento.
Public Const ERRO_VALOR_RECEBER_NEGATIVO = 2988 'Sem parametros
'O valor a receber não pode ser negativo
Public Const ERRO_PREENCH_CPOS_OBRIG_TELA = 2989 'sem parametros
'Preencha os campos obrigatórios da tela.
Public Const ERRO_SELECIONE_CHEQUES_NO_GRID = 2990
'Marque na coluna Selecionar do grid os cheques que deseja imprimir
Public Const ERRO_EXCLUSAO_COND_PAGAMENTO_A_VISTA = 2991
'Não é permitida a exclusão da condição de pagamento à vista.
Public Const ERRO_ALTERACAO_COND_PAGAMENTO_A_VISTA = 2992
'Não é permitida a alteração da condição de pagamento à vista.
Public Const ERRO_CAMPOS_CANCELARBAIXA_NAO_PREENCHIDOS = 2993 'Sem parâmetros
'Os campos Fornecedor, Filial, Tipo, Número, Parcela e Sequencial devem estar preenchidos.
Public Const ERRO_VALORES_TOTAL_BASE = 2994 'Sem parametros
'A base de cálculo da comissão não pode ser maior que o valor total.
Public Const ERRO_FLUXO_CONTA_NAO_PREENCHIDA = 2995 'Parametros iLinha
'A Conta da linha %i não foi preenchida.
Public Const ERRO_CCI_EXISTENTE_GRID = 2996 'Parametros sCCI, iLinha
'A Conta %s já está presente na linha %i do grid.
Public Const ERRO_FORNECEDOR_REPETIDO = 2997 'Parametro sFornecedor, iLinhaGrid
'O Fornecedor já está presente neste fluxo na linha %i.
Public Const ERRO_CLIENTE_REPETIDO = 2998 'Parametro sCliente, iLinhaGrid
'O Cliente já está presente neste fluxo na linha %i.
Public Const ERRO_TIPODEAPLICACAO_CODIGO_REPETIDO = 2999 'Parametros = iAplicacao, iLinha
'O Tipo de Aplicação %i já está presente neste fluxo na linha %i.
Public Const ERRO_EXCLUSAO_OCORR_ALTERACAO_DATA = 16000 'sem parametros
'Uma ocorrência que provocou a alteração do vencimento da parcela deve ser desfeita pela inclusão de uma nova ocorrência alterando o vencimento para a data correta.
Public Const ERRO_OUTRA_OCORR_ALTEROU_VCTO = 16001 'sem parametros
'Outra ocorrencia posterior a esta alterou a data de vencimento da parcela.
Public Const ERRO_BAIXAREC_COBR_CART = 16002 'sem parametros
'Para este tipo de baixa a parcela deve estar em carteira
Public Const ERRO_COBRADOR_USADO_TIPOSCLIENTE = 16003 'Parametros: iCodigoCobrador, iCodigoTipoCliente
'O Cobrador %i está sendo utilizado pelo Tipo de Cliente com o código %i.
Public Const ERRO_COBRADOR_USADO_PARCELASREC = 16004 'Parametros: iCodigoCobrador
'O Cobrador %i está sendo utilizado por uma Parcela a Receber.
Public Const ERRO_COBRADOR_USADO_PARCELASREC_BAIXADA = 16005 'Parametros: iCodigoCobrador
'O Cobrador %i está sendo utilizado por uma Parcela a Receber Baixada.
Public Const ERRO_COBRADOR_USADO_TRANSFCARTCOBR = 16006 'Parametros: iCodigoCobrador
'O Cobrador %i foi utilizado na Transferência de Carteira de Cobrança.
Public Const ERRO_CARTEIRACOBRADOR_USADO_BORDEROCOBRANCA = 16007 'Parametros: iCodCarteiraCobranca, iCobrador e iNumBordero
'A Carteira %i do Cobrador %i está sendo utilizado pelo Bordero de Cobrança número %i.
Public Const ERRO_CARTEIRACOBRADOR_USADO_PARCELAREC = 16008 'Parametros: iCodCarteiraCobranca, iCobrador
'A Carteira %i do Cobrador %i está sendo utilizado por uma Parcela a Receber.
Public Const ERRO_CARTEIRACOBRADOR_USADO_PARCELAREC_BAIXADAS = 16009 'Parametros: iCodCarteiraCobranca, iCobrador
'A Carteira %i do Cobrador %i está sendo utilizado por uma Parcela a Receber Baixadas.
Public Const ERRO_SEQUENCIAL_INVALIDO = 16011 'Sem Parametros
'O Sequencial tem que ser um valor inteiro positivo.
Public Const ERRO_OCORR_ALT_DATA_VCTO_SEM_DATA = 16012 'sem parametros
'A nova data de vencimento tem que estar preenchida para este tipo de ocorrência.
Public Const ERRO_OCORR_VCTO_ALT_ERRADA = 16013 'sem parametros
'A nova data de vencimento não pode estar preenchida para este tipo de ocorrência.
Public Const ERRO_COMISSOES_BAIXADA_ALT_DEBITOS = 16014 'parametro iCodVendedor
'A comissão do vendedor %s, que está baixada, não pode ser excluída ou ter seu valor alterado.
Public Const ERRO_CREDITOPAGAR_VINCULADO_NFISCAL = 16015 'parametro: numero do credito
'Crédito para com Fornecedor com número %s está vinculado à Nota Fiscal e portanto não pode ser excluído neste módulo.
Public Const ERRO_COMISSAO_BAIXADA_CANC_BAIXA = 16016 'sem parametros
'Não pode cancelar uma baixa de parcela que teve a comissao já baixada (paga).
Public Const ERRO_DATA_DESC_APOS_VCTO = 16017 'parametro: numero da parcela
'A parcela %s tem desconto posterior ao vencimento
Public Const ERRO_LEITURA_ANTECPAG_DATA = 16018 'parametros: data
'Erro n a leitura de adiantamentos a fornecedor na data %s
Public Const ERRO_LEITURA_ANTECREC_DATA = 16019 'parametros: data
'Erro na leitura de adiantamentos de cliente na data %s
Public Const ERRO_LEITURA_BAIXASPARCPAG = 16020 'Sem parâmetros
'Erro na leitura da tabela BaixasParcPag.
Public Const ERRO_FILIAL_CENTR_COBR_NAO_SELEC = 16021 'sem parametros
'É preciso selecionar a filial que irá centralizar a cobrança ou deixá-la como independente por filial
Public Const ERRO_DATAINICIAL_NAO_PREENCHIDA = 16022
'O preenchimento da data inicial é obrigatório.
Public Const ERRO_DATAFINAL_NAO_PREENCHIDA = 16023
'O preenchimento da data final é obrigatório.
Public Const ERRO_CONTATO_NAO_PREENCHIDO = 16024
'O preenchimento do contato é obrigatório.
Public Const ERRO_TELCONTATO_NAO_PREENCHIDO = 16025
'O preenchimento do telefone de contato é obrigatório.
Public Const ERRO_ENDERECO_NAO_PREENCHIDO = 16026 'Sem parametro
'O preenchimento do Endereço é obrigatório.
Public Const ERRO_COMPLEMENTO_NAO_PREENCHIDO = 16027
'O preenchimento do Complemento é obrigatório.
Public Const ERRO_DATA_FINAL_DO_MES = 16028 'Parametro : dtDataFinal
'A Data deve ser preenchida com o ultimo dia do mês. Data Correta: %dt.
Public Const ERRO_DATA_INICIO_DO_MES = 16029 'Parametro : dtDataInicio
'A Data deve ser preenchida com o primeiro dia do mês. Data Correta: %dt.
Public Const ERRO_NFPAG_SEM_DOCORIGINAL = 16030 'lNumeroNota
'A Nota Fiscal %l não possui documento original.
Public Const ERRO_TITULO_PAG_SEM_DOCORIGINAL = 16031 'lNumeroTítulo
'O Título %l não possui Documento Original cadastrado.
Public Const ERRO_TITULO_REC_SEM_DOCORIGINAL = 16032 'Sem Parametros
'Este Título a Receber não tem Documento Original.
Public Const ERRO_ATUALIZACAO_NFISCAL2 = 16033 'Sem Parametros
'Erro na atualização da tabela de NFiscal.
Public Const ERRO_BORDERO_COBRANCA_NAO_CADASTRADO_COBRADOR = 16034 'Parâmetro: lNumBordero, iCobrador
'O Bordero de Cobrança de número %l não está cadastrado para o Cobrador %i.
Public Const ERRO_BORDERODE_MAIOR_BORDEROATE = 16035  'Sem Parametros
'O Bordero De não pode ser maior que o Bordero até.
Public Const ERRO_ENDERECO_FILIALCLIENTE_NAO_INFORMADO = 16036 'Parâmetros: lCliente, iFilialCliente
'O Endereço de cobrança da Filial %i do Cliente %l não está preenchido.
Public Const ERRO_FORMATO_ARQUIVO_INCORRETO = 16037 'Parâmetro: sNomeArquivo
'O arquivo de nome "%s" está em um formato incorreto.
Public Const ERRO_CONTACORRENTE_COBRADOR_NAO_ENCONTRADA = 16038
'A conta corrente associada ao cobrador %i não foi encontrada no banco de dados.
Public Const ERRO_OCORRREMPARCREC_NAO_CADASTRADA = 16039
'A Ocorrência com o Numero Interno %l não está cadastrada no Banco de Dados.
Public Const ERRO_ALTERACAO_BORDERO_PAGTO = 16040
'Erro na atualização da tabela de borderos de Pagamento.
Public Const ERRO_BANCO_CCI_DIFERENTE_COBRADOR = 16041 'Parâmetros: iCodBancoConta, iCodBancoCobrador
'A conta corrente do Cobrador deve pertencer ao banco do cobrador. Banco Conta: %i, Banco Cobrador: %i.
Public Const ERRO_AGENCIA_CONTA_COBRADOR_NAO_PREENCHIDAS = 16042
'Os dados número da conta e a agência não estão preenchidos na conta corrente do cobrador.
Public Const ERRO_FAIXA_NOSSONUMERO_INSUFICIENTE = 16043
'A faixa de nosso número diponível não é suficiente para a geração do arquivo de remessa.
Public Const ERRO_INSERCAO_BORDERO_COBRANCA_RETORNO = 16044
'Erro na inserção de Bordero de Retorno de Cobrança.
Public Const ERRO_INSERCAO_OCORR_REM_PARC_RET = 16045
'Erro na inserção de ocorrência de retorno de remessa de Bordero de Cobrança.
Public Const ERRO_BORDERO_PAGTO_INEXISTENTE1 = 16046 'Parâmetro:lNumIntBordero
'O Borderô de Pagamento com número interno %l não está cadastrado.
Public Const ERRO_LOCK_BORDERO_PAGTO = 16047 'Parâmetro:lNumIntBordero
'Erro na tentativa de fazer lock no bordero de pagamento de número interno %l.
Public Const ERRO_CARTEIRACOBRADOR_NAO_CADASTRADA1 = 16048 'Parâmetros: iCarteira, iCobrador
'A Carteira %i do Cobrador %i não está cadastrada no Bano de Dados.
Public Const ERRO_LEITURA_BANCOSINFO1 = 16049
'Erro de leituta na tabela de informações bancárias CNAB.
Public Const ERRO_INSERCAO_BANCOSINFO = 16050
'Erro na inclusão de registro na tabela BancosInfo
Public Const ERRO_ALTERACAO_BANCOSINFO = 16051
'Erro na alteração de registro na tabela BancosInfo
Public Const ERRO_CARTEIRACOBRADORINFO_NAO_CADASTRADA1 = 16052 'Parâmetros:  iCodCobrador
'As informações bancárias daS carteiras do cobrador %i não estão cadastradas.
Public Const ERRO_CARTEIRACOBRADORINFO_NAO_CADASTRADA = 16053 'Parâmetros:  iCodCarteira, iCodCobrador
'As informações bancárias da carteira %i do cobrador %i não estão cadastradas.
Public Const ERRO_LEITURA_BAIXASEC = 16054 'Sem parâmetros
'Erro na leitura da tabela BaixasRec.
Public Const ERRO_BAIXAREC_NAO_ENCONTRADA = 16055 'Parâmetros: lNumIntParcRec
'Não foi possível encontrar em BaixaRec a Data da Baixa relacionada a parcela %l.
Public Const ERRO_LEITURA_PARCELASPAG2 = 16056 'Sem Parametros
'Erro na tentativa de ler a tabela ParcelasPag.
Public Const ERRO_CONTACORRENTE_COBRANCA_NAO_INFORMADA = 16057
'Para a utilização da cobrança eletrônica a conta corrente deve ser informada.
Public Const ERRO_NOSSONUMERO_INICIAL_MAIOR = 16058 'Parâmetro: iCarteira
'Erro nos dados da Carteira %i. O número inicial é maior que o número final.
Public Const ERRO_PROXNOSSONUMERO_MENOR = 16059 'Parâmetro: iCarteira
'Erro nos dados da Carteira %i. O valor do próximo número não pode ser inferior ao valor inicial.
Public Const ERRO_PROXNOSSONUMERO_MAIOR = 16060 'Parâmetro: iCarteira
'Erro nos dados da Carteira %i. O valor do próximo número não pode ser superior ao valor final.
Public Const ERRO_FAIXA_NOSSONUMERO_IMCOMPLETA = 16061 'Parâmetro: iCarteira
'Erro nos dados da Carteira %i. Quando a faixa de número for informada os campos dos Numeros Inicial, Final e próximo deverão ser preenchidos.
Public Const ERRO_CRITERIOS_NAO_SELECIONADOS = 16062 'Sem Parâmetros
'Pelo menos um dos critérios de conciliação deve ser escolhido.
Public Const ERRO_EXTRATOS_NAO_ENCONTRADOS = 16063
'Nenhum extrato foi encontrado para conciliação.
Public Const ERRO_EXTRATO_NAO_ENCONTRADO = 16064 'Parâmetro: iNumExtrato
'O Extrato de número %i não foi encontrado no Banco de Dados.
Public Const ERRO_TIPOMOVCCI_NAO_CADASTRADO = 16065 'Parâmetro: iTipo
'O Tipo de movimento de conta corrente %i não esta cadastrado.
Public Const ERRO_LEITURA_CCIMOV1 = 16066 'Parametros iCodconta
'Ocorreu um erro na leitura da tabela de Saldos Mensais de Conta Corrente. Conta = %s.
Public Const ERRO_CONTACORRENTEINTERNA_NAO_CADASTRADA = 16067 'Parametro: iCodConta
'A Conta Corrente %i não está cadastrada.
Public Const ERRO_LEITURA_CCIMOVDIA2 = 16068 'parametros iCodConta.
'Ocorreu um erro na Leitura da Tabela de Saldos Diarios de Conta Corrente. Conta = %i.
Public Const ERRO_DIRETORIO_INVALIDO = 16069 'Parâmetro: sDiretório
'O Diretório %s informado não foi encontrado.
Public Const ERRO_AUSENCIA_BORDEROSCOBRANCA = 16070
'Não existem novos borderos de cobrança para remessa.
Public Const ERRO_DRIVE_NAO_ACESSIVEL = 16071 'Parâmetro:
' %s não está acessível.
Public Const ERRO_LEITURA_COBRADORINFO = 16072 'Parâmetro: iCodCobrador
'Erro na leitura das informações bancárias do cobrador %i.
Public Const ERRO_EXCLUSAO_COBRADORINFO = 16073 'Parâmetro: iCodCobrador
'Erro na tentativa de exclusão das informações bancárias do cobrador %i.
Public Const ERRO_INSERCAO_COBRADORINFO = 16074 'Parâmetro: iCodCobraddor
'Erro na tentativa de inserção das informações bancárias do cobrador %i.
Public Const ERRO_LEITURA_CARTEIRACOBRADORINFO = 16075 'Parâmetro: icodcartera,iCodCobrador
'Erro na leitura das informações bancárias da carteira %i do cobrador %i.
Public Const ERRO_EXCLUSAO_CARTEIRACOBRADORINFO = 16076 'Parâmetro: iCarteira, iCodCobrador
'Erro na tentativa de exclusão das informações bancárias da carteira %i  do cobrador %i.
Public Const ERRO_INSERCAO_CARTEIRASCOBRADORINFO = 16077 'Parâmetro: iCarteira, iCodCobrador
'Erro na tentativa de inserção das informações bancárias da carteira %i do cobrador %i.
Public Const ERRO_LEITURA_TIPOSDELCTOCNAB = 16078
'Erro na eitura da tabela de Tipos de Lancamentos CNAB.
Public Const ERRO_EXCLUSAO_TIPOSDELANCTOCNAB = 16079
'Erro na exclusão de registro na tabela de Tipos de Lancamentos CNAB
Public Const ERRO_INCLUSAO_TIPOSDELANCTOCNAB = 16080
'Erro na inclusão de registro na tabela de Tipos de Lancamentos CNAB
Public Const ERRO_CODLANCAMENTO_GRID_NAO_PREENCHIDO = 16081 'ParÂmetro: iLinha
'O código do lançamento na linha %i não foi informado
Public Const ERRO_DESCLANCAMENTO_GRID_NAO_PREENCHIDO = 16082 'ParÂmetro: iLinha
'A descrição do lançamento na linha %i não foi informado
Public Const ERRO_CODIGO_REPETIDO_GRID = 16083 'lCodigo as long
'O Código %l já foi preenchido no Grid.
Public Const ERRO_PAGTOSANTECIPADOS_SUPERIOR_VALOR_PEDCOMPRAS = 16084 'sem parametros
'O total de pagamentos antecipados supera o valor total do Pedido de Compra informado.
Public Const ERRO_FILIAL_FORNECEDOR_REL_PRODUTO = 16085 'iFilialForn, lForncedor, sProduto
'A Filial %s do Fornecedor %s está relacionada com o Produto %s.
Public Const ERRO_FILIAL_FORNECEDOR_REL_PEDCOMPRA = 16086 'iFilialForn, lForncedor, lPedCOmpra
'A Filial %s do Fornecedor %s está relacionada com o Pedido de Compra %s.
Public Const ERRO_FILIAL_FORNECEDOR_REL_REQCOMPRA = 16087 'Parâmtros: iCodFilial, lCodFornecedor, lCod
'A Filial %s do Fornecedor %s está relacionada com a Requisicao de Compra %s.
Public Const ERRO_FILIAL_FORNECEDOR_REL_REQMODELO = 16088 'Parâmtros: iCodFilial, lCodFornecedor, lCod
'A Filial %s do Fornecedor %s está relacionada com a Requisicao Modelo %s.
Public Const ERRO_FILIAL_FORNECEDOR_REL_COTACAO = 16089 'Parâmtros: iCodFilial, lCodFornecedor, lCod
'A Filial %s do Fornecedor %s está relacionada com a Cotação %s.
Public Const ERRO_FILIAL_FORNECEDOR_REL_COTACAOPRODUTO = 16090 'Parâmtros: iCodFilial, lCodFornecedor
'A Filial %s do Fornecedor %s está relacionada com uma Cotação de Produto.
Public Const ERRO_FILIAL_FORNECEDOR_REL_PEDCOTACAO = 16091 'Parâmtros: iCodFilial, lCodFornecedor, lCod
'A Filial %s do Fornecedor %s está relacionada com o Pedido de Cotação %s.
Public Const ERRO_FILIAL_FORNECEDOR_REL_CONCORRENCIA = 16092 'Parâmtros: iCodFilial, lCodFornecedor, lCod
'A Filial %s do Fornecedor %s está relacionada com a Concorrência %s.
Public Const ERRO_LEITURA_CCIMOV3 = 16093 'Sem Parametros
'Erro na leitura da tabela de Saldos Mensais de Conta Corrente.
Public Const ERRO_LEITURA_IMPORTFORN = 16094 'Sem parametros
'Erro na leitura da tabela ImportForn.
Public Const ERRO_FORNECEDOR_FINAL_MENOR = 16095 'Sem Parametros
'Fornecedor Final Menor que o Fornecedor Inicial
Public Const EXCLUSAO_RELOPRAZAOCP = 16096 'Sem parametros
'Confirma a exclusão da Opção de RelatórioCP ?
Public Const ERRO_CLIENTE_FINAL_MENOR = 16097 'Sem Parametros
'Cliente Final Menor que o Cliente Inicial
Public Const EXCLUSAO_RELOPRAZAOCR = 16098 'Sem parametros
'Confirma a exclusão da Opção de RelatórioCP ?
Public Const ERRO_SUBTIPOCONTABIL_TIPOBAIXA_NAO_ENCONTRADO = 16099 'Sem parâmetros
'Não foi encontrada transação na tabela TransacaoCTB correspondente ao Tipo de Baixa selecionado.
Public Const ERRO_LEITURA_MOTIVOSBAIXA = 16100 'Parâmetro: iCodigo
'Ocorreu um erro na leitura da tabela de motivos de baixa de títulos a pagar/receber (MotivosBaixa). Motivo = %s
Public Const ERRO_MOTIVOBAIXA_NAO_ENCONTRADO = 16101 'Parâmetro: iCodigo
'O motivo de baixa de títulos a pagar/receber não foi encontrado. Motivo = %s
Public Const ERRO_CREDITO_PAGAR_SEM_DOC_ORIGINAL = 16102 'Parâmetro: iNumTitulo
'O título de crédito com fornecedores número %s, não possui documento original.
Public Const ERRO_DATAEMISSAO_OBRIGATORIA_DOC_ORIGINAL = 16103 'Sem parâmetros
'Para consultar o documento original de um título, é necessário preencher a data de emissão.
Public Const ERRO_CREDITOPAGAR_NAO_CADASTRADO2 = 16104 'Parâmetros: lNumTitulo, lFornecedor, iFilial, sSiglaDocumento, dtDataEmissao
'O título de Crédito com Fornecedor com os dados abaixo, não está cadastrado.
'Número: %s, Fornecedor: %s, Filial: %s, Tipo: %s e Data de Emissão: %s .
Public Const ERRO_DEBITO_RECEBER_SEM_DOC_ORIGINAL = 16105 'Parâmetro: iNumTitulo
'O título de débito com clientes número %s, não possui documento original.
Public Const ERRO_SIGLA_DOCUMENTO_NAO_PREENCHIDO = 16106
'A sigla do documento não foi informada
Public Const ERRO_LEITURA_RECEBANTECIPADOS1 = 16107 'Parâmetro: lCodCliente, iCodFilial
'Erro na leitura da tabela RecebAntecipados com Cliente %l e Filial %i.
Public Const ERRO_SEM_CHEQUES_SEL = 16108
'Os Parâmetros para a seleção de cheques não foram informados
Public Const ERRO_RETCOBR_EXCLUSAO_ERROS = 16109
'Erro na exclusão de registros de erro do processamento de retorno de cobrança
Public Const ERRO_RETCOBR_INCLUSAO_ERRO = 16110
'Erro na inclusão de registro de erro de processamento de retorno de cobrança
Public Const ERRO_CARTEIRA_COBR_INVALIDA = 16111 'parametros cod carteira no banco, cod cobrador
'Não está cadastrada a carteira com código % no cobrador %s
Public Const ERRO_PARCBORDRETCOBR_NENHUMA = 16112 'parametro: "seu numero"
'Não foi encontrada nenhuma parcela em aberto identificada por %s
Public Const ERRO_PARCBORDRETCOBR_VARIAS = 16113 'parametro: "seu numero"
'Foram encontradas mais de uma parcela em aberto identificadas por %s
Public Const ERRO_LEITURA_PARCREC_RETCOBR = 16114 'sem parametros
'Erro na leitura de parcela para o tratamento do retorno da cobranca
Public Const ERRO_NOSSO_NUMERO_NAO_DEFINIDO = 16115 'Parâmetros: iCodBanco, iCodCarteira
'Para gerar um borderô de cobrança, é necessário definir o próximo 'Nosso Número' para a carteira de cobrança %s do cobrador %s.
Public Const ERRO_LEITURA_TITULOREC_CHEQUEPRE = 16116 'Parâmetros: lNumIntCheque
'Ocorreu um erro na leitura do Título a Receber vinculado ao Cheque Pré-datado com o número interno %s.
Public Const ERRO_TITULOREC_CHEQUEPRE_NAO_ENCONTRADO = 16117 'Parâmetros:
'Não foi encontrado título a receber vinculado ao cheque pré-datado com os seguintes dados: Cliente= %s, Número= %s e Valor= %s.
Public Const ERRO_SEM_CHEQUE_SELECIONADO = 16118
'Não existe cheque selecionado
Public Const ERRO_DATAEMISSAO_NAO_INFORMADA = 16119
'O campo Data de Emissão é de preenchimento obrigatório
Public Const ERRO_FAVORECIDO_NAO_PREENCHIDO = 16120
'O Favorecido não foi informado
Public Const ERRO_CODIGO_NAO_INFORMADO1 = 16121
'O Código não foi informado
Public Const ERRO_DATA_APLICACAO_NAO_INFORMADA = 16122
'A Data de Aplicação não foi informada
Public Const ERRO_CONTA_CORRENTEORIGEM_NAO_ENCONTRADA = 16123 'Parametro ContaOrigem.Text
'A Conta Corrente de Origem %s não foi encontrada - ContaOrigem.Text
Public Const ERRO_CONTACORRENTEORIGEM_NAO_BANCARIA = 16124
'A Conta Corrente de Origem não é bancária
Public Const ERRO_CONTACORRENTEDESTINO_NAO_ENCONTRADA = 16125 'Parametro ContaDestino.Text
'A Conta Corrente de Destino %s não foi encontrada - ContaDestino.Text
Public Const ERRO_CONTACORRENTEDESTINO_NAO_BANCARIA = 16126
'A Conta Corrente de Destino não é bancária






'CÓDIGOS DE AVISO - Reservado de 5200 até 5399
Public Const AVISO_CONFIRMA_EXCLUSAO_REGIOESVENDAS = 5200 'Parametro iCodigo
'Região de Venda com código %s será excluída. Cofirma a exclusão?
Public Const AVISO_CONFIRMA_EXCLUSAO_BANCO = 5201 'Parametro iCodBanco
'Banco com código %s será excluído. Confirma a exclusão?
Public Const AVISO_CRIAR_TABELA_PRECO = 5202
'Deseja cadastrar nova Tabela de Preço?
Public Const AVISO_CRIAR_CONDICAO_PAGAMENTO = 5203
'Deseja cadastrar nova Condição de Pagamento?
Public Const AVISO_CRIAR_MENSAGEM = 5204
'Deseja cadastrar nova Mensagem?
Public Const AVISO_CRIAR_REGIAO = 5205
'Deseja cadastrar nova Região de Venda?
Public Const AVISO_CRIAR_COBRADOR = 5206
'Deseja cadastrar novo Cobrador?
Public Const AVISO_CRIAR_TRANSPORTADORA = 5207
'Deseja cadastrar nova Transportadora?
Public Const AVISO_CRIAR_TIPOCLIENTE = 5208
'Deseja cadastrar novo Tipo de Cliente?
Public Const AVISO_EXCLUIR_CLIENTE = 5209 'Parametro: lCodCliente
'Confirma exclusão do Cliente com código %l, Matriz e suas Filiais?
Public Const AVISO_EXCLUIR_FILIALCLIENTE = 5210
'Confirma exclusão de Filial Cliente?
Public Const AVISO_CRIAR_TIPOFORNECEDOR = 5211
'Deseja cadastrar novo Tipo de Fornecedor?
Public Const AVISO_EXCLUIR_FORNECEDOR = 5212 'Parametros: lCodFornecedor
'Confirma exclusão do Fornecedor com código %l, Matriz e suas Filiais?
Public Const AVISO_CRIAR_FORNECEDOR = 5213
'Deseja cadastrar novo Fornecedor?
Public Const AVISO_EXCLUIR_FILIAL_FORNECEDOR = 5214
'Confirma exclusão de Filial Fornecedor?
Public Const AVISO_DATA_VALOR_NAO_ALTERAVEIS = 5215
'Os campos Data e Valor não serão alterados pois o movimento já foi conciliado. Deseja Prosseguir ?
Public Const AVISO_CODCONTACORRENTE_INEXISTENTE = 5216 'Paramento: iCodigo
'A conta corrente %s não existe. Deseja Criá-la?
Public Const AVISO_TIPOMEIOPAGTO_INEXISTENTE = 5217 'Parametro: iTipoMeioPagto
'O Tipo de Pagamento %s não existe. Deseja Criá-lo?
Public Const AVISO_FAVORECIDO_INEXISTENTE = 5218 'Parametro: iCodigo
'O Favorecido %s não existe. Deseja criar novo favorecido?
Public Const AVISO_CONFIRMA_EXCLUSAO_SAQUE = 5219 'Parametros: iCodconta, lSequencial
'Confirma a exclusão do saque na conta %s , sequencial %s
Public Const AVISO_EXCLUIR_TIPOCLIENTE = 5220
'Confirma exclusão de Tipo de Cliente?
Public Const AVISO_CRIAR_PADRAOCOBRANCA = 5221
'Padrão Cobrança não está cadastrado. Deseja criar?
Public Const AVISO_CRIAR_FORNECEDOR_1 = 5222 'Parametro: sNomeReduzido
'Fornecedor %s não está cadastrado. Deseja criar?
Public Const AVISO_CRIAR_FORNECEDOR_2 = 5223 'Parametro: lCodigo
'Fornecedor com código %s não está cadastrado. Deseja criar?
Public Const AVISO_CRIAR_FORNECEDOR_3 = 5224 'Parametro: sCGC
'Fornecedor com CGC/CPF %s não está cadastrado. Deseja criar?
Public Const AVISO_CODBANCO_INEXISTENTE = 5225 'Parametro: Codigo do Banco
'O Banco com código %s não está cadastrado. Deseja criar?
Public Const AVISO_CONFIRMA_EXCLUSAO_CONTACORRENTE = 5226 'Parametro: codconta
'Confirma a exclusão da Conta Corrente?
Public Const AVISO_CONFIRMA_EXCLUSAO_HISTMOVCTA = 5227 'Parametro: código
'Confirma a exclusão do Histórico Padrão de Movimentação de Conta %s ?
Public Const AVISO_CRIAR_PADRAO_COBRANCA = 5228
'Deseja cadastrar novo Padrão Cobrança ?
Public Const AVISO_CONTAORIGEM_INEXISTENTE = 5229 'Param : conta
'A conta Origem %s nao esta cadastrada. Deseja Cria-la?
Public Const AVISO_CONTADESTINO_INEXISTENTE = 5230 'Parametro: iconta
'A conta Destino %s nao esta cadastrada. Deseja Criá-la?
Public Const AVISO_CONFIRMA_EXCLUSAO_TIPOAPLICACAO = 5231 'Sem parâmetro
'Confirma exclusão do Tipo de aplicação ?
Public Const AVISO_CONFIRMA_EXCLUSAO_DEPOSITO = 5232 'Parametro: COdconta, sequencial
'Confirma a exclusão do deposito da conta %s do sequencial %s
Public Const AVISO_NUM_MOV_ULTRAPASSOU_LIMITE = 5233 'Parametro Limite de Movimentos Exibidos
'O número de movimentos para a condição de seleção atual ultrapassou o limite do sistema. Somente os primeiros %i movimentos serão exibidos.
Public Const AVISO_NUM_LANCEXTRATO_ULTRAPASSOU_LIMITE = 5234 'Parametro Limite de Lançametos de Extrato Exibidos
'O número de lançamentos de extrato bancário para a condição de seleção atual ultrapassou o limite do sistema. Somente os primeiros %i lançamentos serão exibidos.
Public Const AVISO_CONCILIACAO_TOTAL_MOV_EXT_DIFERENTES = 5235 'Parametros Total de Movimentos e Total de Lancamentos de Extrato
'O total de movimento = %d não coincide com o total de extrato = %d. Confirma a conciliação?
Public Const AVISO_FLUXO_DATAFINAL_INALTERAVEL = 5236 'Sem Parametros
'A Data Final do Fluxo de Caixa não será alterada. Deseja prosseguir com a gravação?
Public Const AVISO_CONFIRMA_EXCLUSAO_FLUXO = 5237 'Parametro: Identificador do Fluxo
'Confirma a exclusão do Fluxo de Caixa %s ?
Public Const AVISO_NUM_FLUXO_PAG_ULTRAPASSOU_LIMITE = 5238 'Data, Parametro Limite de Pagamentos do Fluxo de Caixa
'O número de Pagamentos para a data %s ultrapassou o limite do sistema. Somente os primeiros %l pagamentos serão exibidos.
Public Const AVISO_FILIALFORNECEDOR_INEXISTENTE = 5239 'Parametro: Filial
'A Filial %s não existe. Deseja criá-la?
Public Const AVISO_CRIAR_FILIAL_FORNECEDOR = 5240 'Sem parâmetros
'Deseja cadastrar nova Filial de Fornecedor?
Public Const AVISO_CONFIRMA_EXCLUSAO_ANTECIPPAG = 5241 'Parametros: iCodconta
'Confirma a exclusão do Pagamento antecipado na conta %i ?
Public Const AVISO_NFPAG_BAIXADA_MESMO_NUMERO = 5242 'Parâmetros :lFornecedor, iFilial, lNumNotaFiscal, dtDataEmissao
'Existe Nota Fiscal Baixada com os dados: Código Fornecedor = %l, Código Filial = %i, Número = %l, Data Emissao = %dt. Deseja prosseguir na inserção de nova Nota Fiscal com o mesmo número?
Public Const AVISO_NFPAG_PENDENTE_MESMO_NUMERO = 5243 'Parâmetros :lFornecedor, iFilial, lNumNotaFiscal, dtDataEmissao
'Existe Nota Fiscal Pendente com os dados: Código Fornecedor = %l, Código Filial = %i, Número = %l, Data Emissao = %dt. Deseja prosseguir na inserção de nova Nota Fiscal com o mesmo número?
Public Const AVISO_NFPAG_MESMO_NUMERO = 5244 'Parâmetros :lFornecedor, iFilial, lNumNotaFiscal, dtDataEmissao
'Existe Nota Fiscal a Pagar com os dados: Código Fornecedor = %l, Código Filial = %i, Número = %l, Data Emissao = %dt. Deseja prosseguir na inserção de nova Nota Fiscal com o mesmo número?
Public Const AVISO_EXCLUSAO_NFPAG = 5245 'Parametro: lNumNotaFiscal
'Confirma a exclusão da Nota Fiscal número %l ?
Public Const AVISO_CRIAR_FILIALFORNECEDOR = 5246 'Parametros: sFornecedorNomeRed, iCodFilial
'A Filial %i do Fornecedor %s não está cadastrada. Deseja criar ?
Public Const AVISO_DATAVENCIMENTO_ALTERAVEL = 5247
'Todos os campos, com exceção da Data de Vencimento, não serão alterados. Deseja prosseguir?
Public Const AVISO_NUM_FLUXO_RECEB_ULTRAPASSOU_LIMITE = 5248 'Data, Parametro Limite de Recebimentos do Fluxo de Caixa
'O número de Recebimentos para a data %s ultrapassou o limite do sistema. Somente os primeiros %l recebimentos serão exibidos.
Public Const AVISO_NUM_FLUXOFORN_ULTRAPASSOU_LIMITE = 5249 'Data, Parametro Limite de Pagamentos por Fornecedor do Fluxo de Caixa
'O número de Fornecedores para a data %s ultrapassou o limite do sistema. Somente os primeiros %l fornecedores serão exibidos.
Public Const AVISO_CRIAR_CONDICAOPAGTO = 5250 'Parâmetro: iCondicaoPagto
'A Condição de Pagamento %i não está cadastrada. Deseja cadastrá-la?
Public Const AVISO_DATAVENCIMENTO_SUSPENSO_ALTERAVEIS = 5251
'Todos os campos com exceção das Datas de Vencimento, Cobrança e Suspenso no Grid de Parcelas não serão alterados. Deseja proseguir na alteração?
Public Const AVISO_NFFAT_PENDENTE_MESMO_NUMERO = 5252 'Parâmetros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Nota Fiscal Fatura Pendente com os dados Código do Fornecedor = %l, Código da Filial = %i, Número = %l, Data Emissão = %dt. Deseja prosseguir na inserção de nova Nota Fiscal Fatura com o mesmo número?
Public Const AVISO_NFFAT_MESMO_NUMERO = 5253 'Parâmetros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Nota Fiscal Fatura a Pagar com os dados Código do Fornecedor = %l, Código da Filial = %i, Número = %l, Data Emissão = %dt. Deseja prosseguir na inserção de nova Nota Fiscal Fatura com o mesmo número?
Public Const AVISO_NFFAT_BAIXADA_MESMO_NUMERO = 5254 'Parâmetros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Nota Fiscal Fatura Baixada com os dados Código do Fornecedor = %l, Código da Filial = %i, Número = %l, Data Emissão = %dt. Deseja prosseguir na inserção de nova Nota Fiscal Fatura com o mesmo número?
Public Const AVISO_EXCLUSAO_NFFATPAG = 5255 'Parâmetro: lNumTitulo
'Confirma a exclusão da Nota Fiscal número %l ?
Public Const AVISO_CRIAR_FILIALCLIENTE = 5256 'iCodigoFilial,sClienteNomeReduzido
'A Filial %i do Cliente %s não existe na tabela de FiliaisClientes. Deseja Criar ?
Public Const AVISO_CONFIRMA_EXCLUSAO_CHEQUEPRE = 5257 'Parametros: lCliente, iFilial, iBanco, sAgencia, sContaCorrente, lNumero
'Confirma a exclusão do cheque pre com Cliente %l, Filial %i, Banco %i, Agência %s, ContaCorrente %s e Número %l ?
Public Const AVISO_DATAVENCIMENTO_PARCELA_DIFERENTE_DATADEPOSITO = 5258 'Parâmetros: dtDataVencimento , sDataDeposito,iParcela
'A Data de Vencimento dt da Parcela %i é diferente da Data de Depósito %dt. Deseja prosseguir ?
Public Const AVISO_CONFIRMA_EXCLUSAO_ANTECIPREC = 5259 'Parametro iCodConta, lSequencial
'O Recebimento antecipado com Conta Corrente %i será excluído. Confirma a exclusão?
Public Const AVISO_CRIAR_FILIAL_CLIENTE = 5260 'Sem parâmetros
'Deseja cadastrar nova Filial de Cliente?
Public Const AVISO_FILIALCLIENTE_INEXISTENTE = 5261 'Parametro: Filial
'A Filial %s não existe. Deseja criá-la ?
Public Const AVISO_CODTIPOAPLICACAO_INEXISTENTE = 5262  'Paramento: iCodigo
'O Tipo De Aplicação com código %i não existe. Deseja Criá-lo?
Public Const AVISO_CAMPOS_APLICACAO_NAO_ALTERAVEIS = 5263 'Sem parametro
'Os campos, com exceção dos dados de resgate, não serão alterados. Deseja Prosseguir?
Public Const AVISO_CONFIRMA_EXCLUSAO_APLICACAO = 5264 'Parametros: lCodigo
'Confirma a exclusão da aplicacao com Código %l?
Public Const AVISO_APLICACAO_MOV_CONCILIADO = 5265 'Parametro: lCodigo
'O movimento de conta corrente associado a aplicação %l está conciliado. Deseja prosseguir?
Public Const AVISO_CONFIRMA_EXCLUSAO_RESGATE = 5266 'Parametro: iSequencialResgate, lCodigoAplicacao
'Confirma a exclusão do Resgate %i da Aplicação %l ?
Public Const AVISO_CAMPOS_RESGATE_NAO_ALTERAVEIS = 5267 'Sem parametro
'Os campos, com exceção dos dados Histórico, Documento Externo e de Reaplicação do Saldo Atual, não serão alterados. Deseja Prosseguir?
Public Const AVISO_NUM_FLUXOTIPOFORN_PAGTO_ULTRAPASSOU_LIMITE = 5268 'Data, Parametro Limite de Pagamentos do FluxoTipoForn
'O número de Pagamentos para a data %s ultrapassou o limite do sistema. Somente os primeiros %l pagamentos serão exibidos.
Public Const AVISO_NUM_FLUXOTIPOFORN_RECEBTO_ULTRAPASSOU_LIMITE = 5271 'Data, Parametro Limite de Recebimentos do FluxoTipoForn
'O número de Recebimentos para a data %s ultrapassou o limite do sistema. Somente os primeiros %l recebimentos serão exibidos.
Public Const AVISO_NUM_FLUXOCLI_ULTRAPASSOU_LIMITE = 5272 'Data, Parametro Limite de Recebimentos por Cliente do Fluxo de Caixa
'O número de Clientes para a data %s ultrapassou o limite do sistema. Somente os primeiros %l clientes serão exibidos.
Public Const AVISO_TIPODEAPLICACAO_CODIGO_NAO_CADASTRADO = 5275 'Parametro: Codigo Tipo de Aplicacao
'O código %s não está cadastrado.
Public Const AVISO_NUM_FLUXOTIPOAPLIC_ULTRAPASSOU_LIMITE = 5276 'Data, Parametro Limite de Tipos de Aplicação do Fluxo de Caixa
'O número de tipos de aplicação para a data %s ultrapassou o limite do sistema. Somente os primeiros %l recebimentos serão exibidos.
Public Const AVISO_TIPODEAPLICACAO_CODIGO_DEVE_SER_PREENCHIDO = 5277 'Sem parametros
'O código do tipo de aplicação deve ser preenchido
Public Const AVISO_NUM_FLUXOAPLIC_ULTRAPASSOU_LIMITE = 5279 'Data, Parametro Limite de Aplicações do FluxoAplic
'O número de Aplicações para a data %s ultrapassou o limite do sistema. Somente as primeiras %l aplicações serão exibidas.
Public Const AVISO_CONTA_JA_CADASTRADA_FLUXO = 5280 'Sem Parametro
'Conta já cadastrada neste fluxo
Public Const AVISO_NUM_FLUXOSALDOSINICIAIS_ULTRAPASSOU_LIMITE = 5281 'Fluxo, Parametro Limite de Recebimentos do Fluxo de Caixa
'O número de contas para o fluxo %l ultrapassou o limite do sistema. Somente os primeiros %l recebimentos serão exibidos.
Public Const AVISO_FLUXOSALDOINICIAL_CODIGO_DEVE_SER_PREENCHIDO = 5282 'Sem parametros
'O código da conta corrente interna deve ser preenchido
Public Const AVISO_NUM_FLUXOSINTETICO_ULTRAPASSOU_LIMITE = 5283 'Parametro Número do fluxo, Limite de operações por Fluxo Sintético
'O número de operações pa6ra o fluxo %l ultrapassou o limite do sistema. Somente as primeiras %l operações serão exibidas.
Public Const AVISO_CONFIRMA_EXCLUSAO_CREDITOPAGAR = 5284 'Parâmetros: sSiglaDocumento, lNumTitulo
'Confirma a exclusão de Devoluções / Crédito com Tipo %s e Número %l ?
Public Const AVISO_CONFIRMA_EXCLUSAO_TIPODEFORNECEDOR = 5286  'Parametro iCodigo
'O Tipo de Fornecedor %i será excluído. Confirma exclusão?
Public Const AVISO_NAO_E_PERMITIDO_ALTERACOES_DEBRECCLI_LANCADO = 5288 'Parametro: lNumTitulo
'Não é permitida alterações de campos que não sejam do Grid de Comissões de Débito / Devolução Número do Título %l porque está lançado. Deseja prosseguir nas alteração dos campos alteráveis ?
Public Const AVISO_NAO_E_PERMITIDO_EXCLUSAO_DEBRECCLI_BAIXADO = 5289 'Parametro: lCliente, iFilial, sSiglaDocumento, lNumTitulo
'Não é possivel excluir Devolução / Crédito porque está baixado. Cliente %l, Filial %i, Tipo de Documento %s e Número do Título %l.
Public Const AVISO_EXCLUSAO_DEBITORECCLI = 5290 'Parâmetro: lNumTitulo
'Confirma a exclusão do Débito a Receber Número %l ?
Public Const AVISO_CRIAR_COBRADOR1 = 5291 'Parâmetro: iCodigo
'O Cobrador %i não está cadastrado. Deseja Criá-lo?
Public Const AVISO_EXCLUSAO_FATURAPAGAR = 5292 'Parâmetro: lNumTitulo
'Confirma exclusão da Fatura número %l ?
Public Const AVISO_FATURAPAG_PENDENTE_MESMO_NUMERO = 5293 'Parâmetros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Fatura Pendente com os dados: Código do Fornecedor = %l, Código da Filial = %i, Número = %l, Data Emissão = %dt. Deseja prosseguir na inserção de nova Fatura com o mesmo número?
Public Const AVISO_FATURAPAG_BAIXADA_MESMO_NUMERO = 5294 'Parâmetros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Fatura Baixada com os dados: Código do Fornecedor = %l, Código da Filial = %i, Número = %l, Data Emissão = %dt. Deseja prosseguir na inserção de nova Fatura com o mesmo número?
Public Const AVISO_FATURAPAG_MESMO_NUMERO = 5295 'Parâmetros: lFornecedor, iFilial, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Fatura a Pagar com os dados: Código do Fornecedor = %l, Código da Filial = %i, Número = %l, Data Emissão = %dt. Deseja prosseguir na inserção de nova Fatura com o mesmo número?
Public Const AVISO_CONFIRMA_EXCLUSAO_PADRAO_COBRANCA = 5296 'Parametro iCodigo
'O Padrão de Cobrança %i será excluído. Confirma exclusão?
Public Const AVISO_DIAS_PROTESTO = 5297 'Parametro sTipoInstrucao
'Instrução %s não necessita de campo Dias para Devolução/Protesto
'que será zerado. Deseja prosseguir?
Public Const AVISO_CANCELAR_PAGAMENTO = 5298 'Sem parametro
'Confirma o cancelamento do pagamento?
Public Const AVISO_EXCLUSAO_MENSAGEM = 5299 'Sem parâmetros
'Confirma a exclusão da Mensagem?
Public Const AVISO_TITULO_PENDENTE_MESMO_NUMERO = 5300 'Parâmetros: lFornecedor, iFilial, sSigla, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Título a Pagar Pendente com os dados Código do Fornecedor = %l, Código da Filial = %i, SiglaDocumento = %s, Número = %l, Data Emissão = %dt. Deseja prosseguir na inserção de novo Título a Pagar do mesmo Tipo e com o mesmo número?
Public Const AVISO_TITULO_MESMO_NUMERO = 5301 'Parâmetros: lFornecedor, iFilial, sSigla, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Título a Pagar com os dados Código do Fornecedor = %l, Código da Filial = %i, SiglaDocumento = %s, Número = %l, Data Emissão = %dt. Deseja prosseguir na inserção de novo Título a Pagar do mesmo Tipo e com o mesmo número?
Public Const AVISO_TITULO_BAIXADO_MESMO_NUMERO = 5302 'Parâmetros: lFornecedor, iFilial, sSigla, lNumTitulo, dtDataEmissao
'No Banco de Dados existe Título a Pagar Baixado com os dados Código do Fornecedor = %l, Código da Filial = %i, SiglaDocumento = %s, Número = %l, Data Emissão = %dt. Deseja prosseguir na inserção de novo Título a Pagar do mesmo Tipo e com o mesmo número?
Public Const AVISO_EXCLUSAO_TITULO = 5303 'Parâmetro: lNumTitulo
'Confirma a exclusão do Título número %l e Parcelas associadas ?
Public Const AVISO_CONFIRMA_EXCLUSAO_OCORRENCIA = 5304 'Sem parâmetros
'A Ocorrência será excluída da tabela OcorrenciasRemParcRec. Confirma exclusão?
Public Const AVISO_CRIAR_PORTADOR = 5305 'Parâmetro: iCodigo
'O Portador %i não existe, deseja criá-lo?
Public Const AVISO_CRIAR_FILIALCLIENTE1 = 5306 'Parâmetro: sFilial
'A Filial %s não existe na tabela de FiliaisClientes. Deseja Criar ?
Public Const AVISO_BORDERO_BD_DIFERENTE = 5307 'Parâmetros: lNumBordero, iCobrador, iCodCarteiraCobranca, dtDataEmissao
'O Borderô de Número %l está com os dados Cobrador = %i, Carteira = %l, Data de Emissão = %dt diferentes da tela. Deseja prosseguir no cancelamento do bordero com esses dados?
Public Const AVISO_BATCH_CANCELADO = 5308
'Atualização cancelada pelo usuário.
Public Const AVISO_IMPOSSIVEL_CONTINUAR = 5309
'O processo de atualização ainda está sendo efetuado.
Public Const AVISO_CRIAR_TIPOAPLICACAO = 5310 'Parametro: iCodigo
'TipoAplicação com código %i não está cadastrada. Deseja criar?
Public Const AVISO_EXCLUSAO_PAIS = 5311 'Sem parâmetros
'Confirma a exclusão do Pais?
Public Const AVISO_HISTMOVCTA_INEXISTENTE = 5312 'Parâmetro: iCodigo
'O Histórico de Movimentação de Conta %i não existe. Deseja Criá-lo?
Public Const AVISO_EXCLUSAO_RELOPDEVEDORES = 5313 'Sem parametros
'Confirma a exclusão da Opção de RelatórioDevedores ?
Public Const AVISO_EXCLUSAO_RELOPHISTAPLIC = 5314 'Sem parametros
'Confirma a exclusão da Opção de RelatórioHistAplic?
Public Const AVISO_EXCLUSAO_RELOPMOVFIN = 5315 'Sem parametros
'Confirma a exclusão da Opção de RelatórioMovFin ?
Public Const AVISO_EXCLUSAO_RELOPPOSAPLIC = 5316 'Sem parametros
'Confirma a exclusão da Opção de RelatórioPosAplic?
Public Const AVISO_EXCLUSAO_RELOPCOMISVEND = 5317 'Sem parametros
'Confirma a exclusão da Opção de Relatório ?
Public Const AVISO_EXCLUSAO_RELOPTITRECMALA = 5318 'Sem parametros
'Confirma a exclusão da Opção de RelatórioTitRecMala?
Public Const AVISO_NAO_IMP_CHEQUE = 5319 'sem parametro
'Prossegue mesmo sem ter impresso o cheque ?
Public Const AVISO_PAGTO_ATRASO_DESC = 5320 'parametro: # da linha
'Você digitou um valor de desconto para a parcela da linha %s que está sendo paga em atraso. Deseja voltar para corrigir ?
Public Const AVISO_PAGTO_EM_DIA_MULTA = 5321 'parametro: # da linha
'Você digitou um valor de multa ou juros para a parcela da linha %s que está sendo paga em dia. Deseja voltar para corrigir ?
Public Const AVISO_CHEQUE_NUM_USADO_DATA = 5322 'parametro: data do cheque
'Um cheque com o mesmo número e data %s já está registrado no sistema. Deseja voltar para corrigir o número ?
Public Const AVISO_ALTERACAO_VCTO_AFETA_DESC = 5323 'sem parametros
'Esta parcela contém descontos por antecipação de pagto. Se for necessário atualizá-los poderá ser necessário dar baixa no título e reenviá-lo ao banco com as informações atualizadas. Prossegue assim mesmo ?
Public Const AVISO_DATAVENCIMENTO_NAO_ORDENADA = 5324 'Sem parametros
'As Datas de Vencimento no Grid não estão ordenadas. Deseja prosseguir assim mesmo ?
Public Const AVISO_INCLUSAO_ANTECIPPAG_ADICIONAL = 5325 'parametros: valor
'Já estão registrados %s em adiantamentos para esta filial de fornecedor nesta data. Deseja registrar um novo adiantamento ?
Public Const AVISO_INCLUSAO_ANTECIPREC_ADICIONAL = 5326 'parametros: valor
'Já estão registrados %s em adiantamentos desta filial de cliente nesta data. Deseja registrar um novo adiantamento ?
Public Const AVISO_CRIAR_BANCO = 5327 'Parametros: iCodVanco
'O Banco %i não existe, deseja criá-lo?
Public Const AVISO_EXCLUSAO_CARTEIRA = 5328 'Parâmetro: iCodCarteira
'Confirma a exclusão da carteira %i?
Public Const AVISO_TITULOS_PAGAR_CAMPOS_ALTERAVEIS = 5329
'Todos os campos com exceção das Datas de Vencimento , Tipo de Cobrança, Suspenso, Banco Cobrador e Portador no Grid de Parcelas não serão alterados. Deseja proseguir na alteração?
Public Const AVISO_TITULOS_PAGAR_VINCULADO_ADIANTAMENTO = 5330 'Parâmetros: lCodFornecedor, iCodFilial
'O Titulo a pagar em questão está vinculado a um Adiantamento do Fornecedor de código %s e Filial de código %s.
'Deseja prosseguir com a gravação?
Public Const AVISO_TITULOS_PAGAR_VINCULADO_CREDITO = 5331 'Parâmetros: lCodFornecedor, iCodFilial
'O Titulo a pagar em questão está vinculado a um Crédito do Fornecedor de código %s e Filial de código %s.
'Deseja prosseguir com a gravação?
Public Const AVISO_FORNECEDOR_CGC_IGUAL = 5332 'Parametro: sCGC
'Já existe um outro Fornecedor Cadastrado com o CGC %s, deseja continuar a Gravação?




'!! A EXCLUIR QUANDO SUBSTITUIR FILIALCLIENTE_EXCLUI
Public Const ERRO_FILIALCLIENTE_REL_NFI = 5255
'Filial Cliente relacionada com Nota Fiscal Interna

