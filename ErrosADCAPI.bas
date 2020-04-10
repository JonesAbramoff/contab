Attribute VB_Name = "ErrosADCAPI"
Option Explicit


'Códigos de Erro  RESERVADO de 1000 a 1999
Global Const SUCESSO = 0
Public Const ERRO_LEITURA_TELAINDICE = 1000 'Sem parametro
'Erro na leitura da tabela TelaIndice.
Public Const ERRO_INDICES_DA_TELA_INEXISTENTES = 1001 'parametro sNomeTela
'Não existem índices associados à Tela %s na tabela TelaIndice.
Public Const ERRO_INDICE_DA_TELA_INEXISTENTE = 1002 'parametros sNomeTela (%s1), sNomeIndice (%s2)
'Não existe índice da tela %s1 com nome %s2.
Public Const ERRO_LEITURA_TELAINDICECAMPO = 1003 'Sem parametros
'Erro na leitura da tabela TelaIndiceCampo.
Public Const ERRO_CAMPOS_DO_INDICE_INEXISTENTES = 1004 'parametros sNomeTela, iIndice
'Não existem campos associados ao índice %i da tabela %s.
Public Const ERRO_MENOS_CAMPOS_TELA_QUE_CAMPOS_INDICE = 1005 'sem parametros
'Há menos campos de tela que campos do índice.
Public Const ERRO_NOME_TABELA_VAZIO = 1006 'sem parametros
'Nome da Tabela associada à tela está vazio.
Public Const ERRO_BIND_CAMPO2 = 1007 'sem parametros
'Erro no "bind" de um campo que compõe a expressão de um comando SQL.
Public Const ERRO_INEXISTE_CAMPO_TELA_IGUAL_CAMPO_INDICE = 1008 'parametro sCampo
'Não existe campo de Tela que corresponda ao campo de índice %s.
Public Const ERRO_REFERENCIA_COMANDO_ABERTO = 1009 'sem parametros
'A variável lComando não referencia comando aberto.
'Global Const ERRO_ABERTURA_COMANDO = 1010
'Não conseguiu abrir o comando.
Global Const ERRO_ABERTURA_TRANSACAO = 1011
'Não conseguiu abrir a transação.
Global Const ERRO_INSERCAO_ARQUIVO_TEMPORARIO = 1012
'Erro de inserção no arquivo temporário.
Global Const ERRO_ABERTURA_ARQUIVO_TEMPORARIO = 1013
'Não conseguiu abrir o arquivo temporário.
Global Const ERRO_INSERCAO_ARQUIVO_SORT = 1014
'Erro de inserção no arquivo sort.
Global Const ERRO_ABERTURA_ARQUIVO_SORT = 1015
'Não conseguiu abrir o arquivo sort.
Global Const ERRO_COMMIT = 1016  'Sem parâmetros
'Não foi possível confirmar a transação.
Global Const ERRO_PREPARACAO_ARQUIVO_TEMP = 1017
'Erro na criação de arquivo temporário.
Global Const ERRO_CLASSIFICAR_ARQUIVO_SORT = 1018
'Erro no arquivo sort.
Global Const ERRO_LEITURA_ARQUIVO_TEMP = 1019
'Erro na leitura do arquivo temporário.
Global Const ERRO_FORNECIDO_PELO_VB = 1020
'Erro fornecido pelo VB: %s.
Global Const ERRO_CRIACAO_INDICE = 1021
'Erro na criação de índice.
Global Const ERRO_EXCLUSAO_INDICE = 1022
'Erro na exclusão de índice.
Global Const ERRO_COMMIT_TRANSACAO = 1023
'Não confirmou a transação.
Global Const ERRO_FORMATO_DATA = 1024 'Sem parametros
'Data Inválida.
Global Const ERRO_ATIVACAO_TELA = 1025
'Não conseguiu ativar a tela.
Global Const ERRO_VALOR_INVALIDO = 1026  'Parametro String com um valor monetário
'O valor digitado: %s é inválido. Exemplos de formatos validos: 12560, 12.560, 12.560,35, 12560,35
Global Const ERRO_ROTINA_NAO_DISPONIVEL = 1027 'nao foi possivel acessar a rotina em questão.
'Verifique se ela esta cadastrada nas tabelas Rotinas e GrupoRotinas no dicionário de dados
'Verifique se você tem acesso a esta rotina.
Global Const ERRO_LEITURA_BROWSEUSUARIOCAMPO = 1028 'Sem parametro
'Erro na leitura da tabela BrowseUsuarioCampo.
Global Const ERRO_BROWSE_SEM_COLUNAS = 1029 'Sem parametro
'Atenção. Não foi selecionado nenhum campo para ser exibido nesta tela. Use o configurador para selecionar alguns.
Global Const ERRO_LEITURA_BROWSEUSUARIOORDENACAO = 1030 'Sem parametro
'Erro na leitura da tabela BrowseUsuarioOrdenacao.
Global Const ERRO_LEITURA_BROWSEINDICE = 1031 'Sem parametro
'Erro na leitura da tabela BrowseIndice.
Global Const ERRO_BROWSE_SEM_ORDENACAO = 1032 'Sem parametro
'Atenção. Não há índices de ordenação cadastrados para esta tela.
Global Const ERRO_LEITURA_CAMPOS = 1033 'Sem parametro
'Erro na leitura da tabela de Campos.
Global Const ERRO_CAMPO_NAO_CADASTRADO = 1034 'Parametros Nome, NomeArq
'O campo %s pertencente a tabela %s não está cadastrado na tabela de Campos.
Global Const ERRO_PREPARACAO_COMANDO_SQL = 1035 'Parametro ComandoSQL
'Erro na Preparação do Comando SQL %s.
Global Const ERRO_BIND_CAMPO = 1036 'Parametros NomeArq, Nome
'Erro no "bind" do campo %s da tabela %s.
Global Const ERRO_LEITURA_TABELA = 1037 'Parametro NomeTabela
'Erro na leitura da tabela %s.
Global Const ERRO_EXECUCAO_COMANDO_SQL = 1038 'Parametro ComandoSQL
'Erro na Execução do Comando SQL %s.
Global Const ERRO_TIPO_CAMPO_INVALIDO = 1039 'Parametro Tipo
'O Tipo do Campo é Inválido. Tipo = %i.
Global Const ERRO_BROWSE_EXCEDEU_MAXIMO_COLUNAS = 1040 'Sem parametro
'Atenção. O número máximo de colunas para uma tela de consulta foi ultrapassado. Diminua o número de colunas selecionadas.
Global Const ERRO_LEITURA_GRUPOBROWSECAMPO = 1041 'Sem parametro
'Erro na leitura da tabela GrupoBrowseCampo.
Global Const ERRO_OBTENCAO_CODIGO_USUARIO = 1042 'Sem parametro
'Erro na tentativa de obter o codigo do usuário
Global Const ERRO_OBTENCAO_CODIGO_GRUPO = 1043 'Sem parametro
'Erro na tentativa de obter o codigo do grupo
Global Const ERRO_LEITURA_TELAS = 1044 'Sem parametro
'Erro na leitura da tabela de Telas.
Global Const ERRO_LOCK_TELAS = 1045 'Sem parametro
'Erro na tentativa de fazer "lock" na tabela Telas.
Global Const ERRO_INSERCAO_BROWSEUSUARIOCAMPO = 1046 'Sem parametro
'Erro na inserção de registro na tabela BrowseUsuarioCampo.
Global Const ERRO_EXCLUSAO_BROWSEUSUARIOCAMPO = 1047 'Sem parametro
'Erro na exclusão de um registro da tabela BrowseUsuarioCampo.
Global Const ERRO_ATUALIZACAO_BROWSEUSUARIOORDENACAO = 1048 'Parametros NomeTela, CodUsuario
'Erro na atualização da tabela BrowseUsuarioOrdenacao. Tela = %s,  Usuario = %s.
Global Const ERRO_INSERCAO_BROWSEUSUARIOORDENACAO = 1049 'Parametros NomeTela, CodUsuario
'Erro na inserção na tabela BrowseUsuarioOrdenacao. Tela = %s,  Usuario = %s.
Global Const ERRO_BROWSE_SEM_COLUNAS1 = 1050 'Sem parametro
'Atenção. Não foi selecionado nenhum campos para ser exibido nesta tela. Selecione pelo menos 1 campo.
Global Const ERRO_TELA_NAO_DISPONIVEL = 1051 'Parametro Nome da Tela(String)
'Não foi possível acessar a tela %s.
'Verifique se ela esta cadastrada nas tabelas Telas e GrupoTelas no dicionário de dados
'Verifique se você tem acesso a esta tela.
Global Const ERRO_LEITURA_BROWSEINDICESEGMENTOS = 1052 'Sem parametro
'Erro na leitura da tabela BrowseIndiceSegmentos.
Global Const ERRO_BIND_CAMPO1 = 1053 'Sem Parametros
'Erro no "bind" de um campo que compõe a expressão de seleção de um comando SQL.
Global Const ERRO_BROWSECONFIGURA_SEM_PARAMETRO = 1054 'Sem Parametro
'A tela de configuração de listagem foi chamada sem o parametro requerido.
Global Const ERRO_LEITURA_BROWSEARQUIVO = 1055 'Sem parametro
'Erro na leitura da tabela BrowseArquivo.
Global Const ERRO_TELA_SEM_PARAMETRO = 1056 'Sem parametros
'A tela foi chamada sem o número de parâmetros adequado.
Global Const ERRO_LEITURA_RELATORIOOPCOES = 1057 'Sem parametro
'Erro na leitura da tabela RelatorioOpcoes.
Global Const ERRO_INSERCAO_RELATORIOOPCOES = 1058 'Sem parametro
'Erro na inserção da tabela RelatorioOpcoes.
Global Const ERRO_ATUALIZACAO_RELATORIOOPCOES = 1059 'Sem parametro
'Erro na atualização da tabela RelatorioOpcoes.
Global Const ERRO_LOCK_RELATORIOOPCOES = 1060 'Sem parametro
'Erro na tentativa de fazer "lock" na tabela RelatorioOpcoes.
Global Const ERRO_EXCLUSAO_RELATORIOOPCOES = 1061 'Sem parametro
'Erro na exclusão de um registro da tabela RelatorioOpcoes.
Global Const ERRO_REL_PARAM_NAO_ENCONTRADO = 1062 'parametro eh o nome do parametro
'O parâmetro para relatório não foi encontrado
Public Const ERRO_CPF_NAO_NUMERICO = 1063 'parametro sCpf
'Cpf %s não é numérico.
Public Const ERRO_CPF_MENOR_OU_IGUAL_ZERO = 1064 'parametro sCpf
'Cpf %s é nulo ou negativo.
Public Const ERRO_CPF_NAO_INTEIRO = 1065 'parametro sCpf
'Cpf %s não é inteiro.
Public Const ERRO_CPF_INVALIDO = 1066 'parametro sCpf
'%s é CPF inválido.
Public Const ERRO_CGC_NAO_NUMERICO = 1067 'parametro sCgc
'Cgc %s não é numérico.
Public Const ERRO_CGC_MENOR_OU_IGUAL_ZERO = 1068 'parametro sCgc
'Cgc %s é nulo ou negativo.
Public Const ERRO_CGC_NAO_INTEIRO = 1069 'parametro sCgc
'Cgc %s não é inteiro.
Public Const ERRO_CGC_INVALIDO = 1070 'parametro sCgc
'%s é CGC inválido.
Public Const ERRO_VALOR_NAO_NUMERICO = 1071 'Parametro sValor
'O valor %s tem que ser numérico.
Public Const ERRO_CODIGO_NAO_INTEIRO = 1072 'Parametro sCodigo
'O codigo %s tem que ser inteiro.
Public Const ERRO_VALOR_PORCENTAGEM = 1073 'Parametro dNumero
'O valor %d não está entre 0 e 100.
Public Const ERRO_LEITURA_MODULO = 1074 'Sem parametro
'Erro na leitura da tabela Modulos.
Public Const ERRO_CGC_OVERFLOW = 1075 'parametro sCgc
'Valor %s ultrapassa o limite de CGC.
Public Const ERRO_NUMERO_NAO_INTEIRO = 1076 'parametro sNumero
'O número %s não é inteiro.
Public Const ERRO_NUMERO_NEGATIVO = 1077 'parametro sNumero
'O número %s é negativo.
Public Const ERRO_MODULO_INEXISTENTE = 1078 'parametro Nome
'Não foi encontrado um módulo com este nome: %s.
Public Const ERRO_NUMERO_NAO_POSITIVO = 1079 'parametro sNumero
'O número %s não é positivo.
Public Const ERRO_VALOR_MENOR_QUE_UM = 1080 'parametro sNumero
'O valor %s é menor do que um.
Public Const ERRO_VALOR_NAO_POSITIVO = 1081  'Parametro String com um valor monetário
'O valor digitado tem que ser positivo. Valor = %s.
Public Const ERRO_LEITURA_FILIAL = 1082 'Parametro CodEmpresa.
'Erro na leitura da tabela FiliaisEmpresas. Empresa = %l.
Public Const ERRO_INTEIRO_OVERFLOW = 1083 'Parametro sNumero
'O número %s ultrapassa o limite do tipo Integer.
Public Const ERRO_LONG_OVERFLOW = 1084 'Parametro sNumero
'O número %s ultrapassa o limite do tipo Long.
Public Const ERRO_CPF_MAIOR_QUE_CPFMAXIMO = 1085 'parametro: sCpf
'Cpf %s é maior do que o Cpf máximo.
Public Const ERRO_CGC_MAIOR_QUE_CGCMAXIMO = 1086 'parametro: sCgc
'Cgc %s é maior do que o Cgc máximo.
Public Const ERRO_GRID_LINHA_INEXISTENTE = 1087 'Parametros: Linha a Ser Excluida, Linha Inicial do Grid, Linha Final do Grid
'A linha %i do grid não pode ser excluida. Escolha uma linha do grid existente, entre %i e %i.
Public Const ERRO_LEITURA_FILIALEMPRESA = 1088 'parametro: iFilialEmpresa
'Erro na leitura da filial %i na tabela FiliaisEmpresas.
Public Const ERRO_FILIAL_EMPRESA_NAO_CADASTRADA = 1089 'parametro: iFilialEmpresa
'A filial %i não está cadastrada no Banco de Dados.
Public Const ERRO_LEITURA_USUARIO = 1090 'Parametro sCodUsuario
'Erro na leitura dos dados do Usuario %s na tabela Usuario.
Public Const ERRO_ATUALIZACAO_USUARIO = 1091 'Parametro sCodUsuario
'Ocorreu um erro na atualização da tabela Usuario. Usuário = %s.
Public Const ERRO_INSERCAO_USUARIO = 1092 'Parametro sCodUsuario
'Ocorreu um erro na inserção de um registro na tabela Usuario. Usuário = %s.
Public Const ERRO_LEITURA_TABELA_CONFIG = 1093 'Parametros: sTabelaConfig, sCodigo
'Erro na leitura da tabela %s. Código = %s.
Public Const ERRO_LOCK_TABELA_CONFIG = 1094 'Parametros: sTabelaConfig, sCodigo
'Erro na tentativa de "lock" na tabela %s. Código = %s.
Public Const ERRO_ATUALIZACAO_TABELA_CONFIG = 1095 'Parametros: sTabelaConfig, sCodigo
'Erro na atualização da tabela %s. Código = %s.
Public Const ERRO_INTEIRO_NAO_MES = 1096 'Parametros: iMes
'O valor %i não se refere a um mês válido.
Public Const ERRO_STRING_NAO_MES = 1097 'Parametros: sMes
'A string %s não se refere a um mês válido.
Public Const ERRO_LOCK_PRODUTOS = 1098  'Sem parametro
'Erro na tentativa de fazer 'lock' na tabela de Produtos.
Public Const ERRO_CCL_NAO_USADO = 1099 'Sem parametro
'Esta rotina não pode ser usada, pois o sistema não utiliza Centro de Custo/Lucro.
Public Const ERRO_PRODUTO_INATIVO = 1100 'Parametro: sCodProduto
'Produto %s está inativo.
Public Const ERRO_PRODUTO_GERENCIAL = 1101 'Parametro: sCodProduto
'Produto %s é gerencial.
Public Const ERRO_PRODUTO_INEXISTENTE = 1102   'Parametro: sProduto
'Produto %s não cadastrado
Public Const ERRO_LEITURA_PRODUTOS = 1103 'Parametro: sCodigo
'Erro na leitura do Produto %s.
Public Const ERRO_LEITURA_PRODUTOSFILIAL = 1104 'Parâmetros: giFilialEmpresa, sProduto
'Erro na leitura da tabela de ProdutosFilial com FilialEmpresa %i e Produto %s.
Public Const ERRO_TABELA_PRECO_NAO_CADASTRADA = 1105 'Parâmetros: iCodTabela
'A Tabela de Preco %i não foi encontrada.
Public Const ERRO_CATEGORIA_JA_SELECIONADA = 1106 'Sem parâmetros
'Uma Categoria não pode ser selecionada mais de uma vez.
Public Const ERRO_ATUALIZACAO_CPRCONFIG = 1107 'Parametro sCodigo
'Erro ao atualizar o registro de configuração que possui o codigo %s na tabela CPRConfig.
Public Const ERRO_LEITURA_USUARIOS_DIC = 1108 'Sem Parametros
'Ocorreu um erro na leitura da tabela de Usuarios do Dicionario de Dados.
Public Const ERRO_LEITURA_TABELA_PAISES = 1109 'Sem Parametros
'Erro na leitura da tabela de Paises
Public Const ERRO_ATUALIZACAO_TABELA_PAISES = 1110 'Parâmetro : iCodigo
'Erro na atualização do Pais %i.
Public Const ERRO_INSERCAO_PAISES = 1111 'Parâmetro: iCodigo
'Erro na Inserção do Pais %i.
Public Const ERRO_EXCLUSAO_PAISES = 1112 'Parâmetro: iCodigo
'Erro na exclusao do Pais %i.
Public Const ERRO_PAIS_ASSOCIADO_ENDERECO = 1113 'Parâmetro: iCodigo
'Não é permitida a exclusão do País %i porque está associado ao Endereço.
Public Const ERRO_LOCK_CODIGO_PAIS = 1114 'Parâmetro Código do País
'Erro na tentativa de fazer lock no País %s.
Public Const ERRO_LEITURA_USUARIOS_DIC1 = 1115 'Parametro: sCodUsuario
'Ocorreu um erro na leitura do Usuário %s na tabela de Usuarios do Dicionário de Dados.
Public Const ERRO_USUARIO_NAO_PREENCHIDO = 1116 'Sem Parametros
'O preenchimento do Código do Usuário é obrigatório.
Public Const ERRO_SENHA_NAO_PREENCHIDA = 1117 'Sem Parametros
'O preenchimento da Senha é obrigatória.
Public Const ERRO_USUARIO_NAO_CADASTRADO = 1118 'Parametros: sCodUsuario
'O Usuario %s não está cadastrado.
Public Const ERRO_SENHA_INVALIDA = 1119 'Sem Parametro
'Senha Inválida.
Public Const ERRO_LOCK_USUARIOS_DIC = 1120 'Sem Parametros
'Ocorreu um erro na tentativa de Lock da tabela de Usuarios do Dicionario de Dados.
Public Const ERRO_LONG_OVERFLOW1 = 1121 'Parametro sNumero
'O número %s ultrapassa o valor limite que é 99.999
Public Const ERRO_PAISES_NOME_DUPLICADO = 1122
'Erro na tentativa de cadastrar novo País com o Nome ja existente.
Public Const ERRO_LEITURA_REGIOESVENDAS2 = 1123
'Erro na leitura da Tabela de RegioesVendas.
Public Const ERRO_PAIS_ASSOCIADO_REGIOESVENDAS = 1124 'Parâmetro: iCodigo
'Não é permitida a exclusão do País %i porque está associado a uma Região de Venda.
Public Const ERRO_LEITURA_USUARIOMODULO = 1125 'Parametros: sUsuario, lEmpresa, iFilialEmpresa
'Ocorreu um erro na leitura da visão UsuarioModulo. Usuário = %s, Empresa = %l, Filial da Empresa = %i.
Public Const ERRO_LEITURA_MENUITENS = 1126 'Sem Parametro
'OCorreu um erro na leitura da tabela MenuItens.
Public Const ERRO_LEITURA_USUARIOITENSMENU = 1127 'Parâmetros: sCodUsuario, lCodEmpresa, iCodFilial
'Erro na leitura da tabela UsuarioItensMenu com Usuário %s, Empresa %l e Filial %i.
Public Const ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2 = 1128 'Parametro sFilialEmpresa
'Filial Empresa %s não está cadastrada no Banco de Dados.
Public Const ERRO_TIPO_CAMPO_INVALIDO1 = 1129 'Parametros iLinha, iTipo
'O Tipo do Campo da Linha %i da Pesquisa é Inválido. Tipo = %i.
Public Const ERRO_GRIDSELECAO_SEM_PREENCHIMENTO = 1130 'Parametros iLinha, iColuna
'Na Linha %i do Grid de Pesquisa, a Coluna %i não está preenchida.
Public Const ERRO_GRIDSELECAO_INTEIRO_INVALIDO = 1131 'Parametros iLinha, iColuna, sValor
'Na Linha %i do Grid de Pesquisa, a Coluna %i contém o valor %s que é inválido. Digite um valor inteiro. Exs: 25, 1, 30000, -5126.
Public Const ERRO_GRIDSELECAO_LONG_INVALIDO = 1132 'Parametros iLinha, iColuna, sValor
'Na Linha %i do Grid de Pesquisa, a Coluna %i contém o valor %s que é inválido. Digite um valor longo. Exs: 25, 1, 1250000, -5126000.
Public Const ERRO_GRIDSELECAO_DOUBLE_INVALIDO = 1133 'Parametros iLinha, iColuna, sValor
'Na Linha %i do Grid de Pesquisa, a Coluna %i contém o valor %s que é inválido. Digite um valor duplo. Exs: 25.2, 1.5, 1250000, -5126000, 17, 3.
Public Const ERRO_GRIDSELECAO_DATA_INVALIDA = 1134 'Parametros iLinha, iColuna, sValor
'Na Linha %i do Grid de Pesquisa, a Coluna %i contém o valor %s que é inválido. Digite uma data válida. Exs: 15/01/99, 1/1/2001, 12/7/1998.
Public Const ERRO_OPERADOR_LIKE = 1135 'Parametros iLinha
'Na Linha %i o operador LIKE não pode ser usado pois o campo não é do tipo Texto.
Public Const ERRO_DESCRICAO_FERIADO_NAO_PREENCHIDA = 1136 'Sem parâmetros
'O preenchimento da descrição do Feriado é obrigatório.
Public Const ERRO_DATA_FERIADO_NAO_PREENCHIDA = 1137 'Sem parâmetros
'O preenchimento da Data do Feriado é obrigatório.
Public Const ERRO_LEITURA_FERIADOS1 = 1138 'Sem parâmetros
'Erro na leitura da tabela de Feriados.
Public Const ERRO_LOCK_FERIADOS = 1139 'Parâmetros: dtData, iFilialEmpresa
'Erro na tentativa de fazer "Lock" no Feriado %dt da Filial %i da tabela de Feriados.
Public Const ERRO_ATUALIZACAO_FERIADOS = 1140 'Parâmetros: dtData, iFilialEmpresa
'Erro na tentativa de atualizar o Feriado %dt da Filial %i na tabela de Feriados.
Public Const ERRO_INSERCAO_FERIADOS = 1141 'Parâmetros: dtData, iFilialEmpresa
'Erro na tentativa de inserir o Feriado %dt da Filial %i na tabela de Feriados.
Public Const ERRO_FERIADO_NAO_CADASTRADO = 1142 'Parâmetros: dtData, iFilialEmpresa
'O Feriado %dt da Filial %i não está cadastrado no Banco de Dados.
Public Const ERRO_EXCLUSAO_FERIADO = 1143 'Parâmetros: dtData, iFilialEmpresa
'Erro na tentativa de excluir o Feriado %dt da Filial %i da tabela de Feriados.
Public Const TELA_MODULO_CHAMADA_SEM_PARAMETRO = 1144 'Sem Parametro
'A tela Módulo foi chamada sem a passagem do parametro necessário.
Public Const ERRO_OBTENCAO_MODULO = 1145
'Não conseguiu ler módulo.
Public Const ERRO_OBTENCAO_GRUPO = 1146
'Não conseguiu ler grupo.
Public Const ERRO_LEITURA_RELATORIO = 1147
'Não conseguiu ler relatório.
Public Const ERRO_ITEM_NAO_SELECIONADO = 1148
'Tentativa de excluir item não selecionado.
Public Const ERRO_EXCLUSAO_DE_RELATORIO = 1149
'Erro na exclusão do relatório.
Public Const ERRO_GRAVACAO_RELATORIO = 1150
'Erro na gravação do relatório.
Public Const ERRO_NOME_TSK_NAO_COMECA_LETRA = 1151 'Parametro: NomeTsk.Text
'Nome do Arquivo Tsk não começa por letra.
Public Const ERRO_LEITURA_FILIALEMPRESA1 = 1152 'parametros: lEmpresa, iFilial
'Erro na leitura da tabela FiliaisEmpresas (DIC) com chave CodEmpresa=%l, CodFilial=%i.
Public Const ERRO_TAMANHO_CPF = 1153 'Sem parametros
'O tamanho do campo CPF tem que ser 11 caracteres.
Public Const ERRO_TAMANHO_CGC = 1154 'Sem parametros
'O tamanho do campo CGC tem que ser 14 caracteres.
Public Const ERRO_OBJETO_NAO_TEM_TIPO_ESPERADO = 1155 'Parametro: sNome
'O objeto %s não é do tipo esperado.
Public Const ERRO_PRODUTO_SEG_MEIO_NAO_PREENCHIDOS = 1156 'Sem parametro
'Todos os segmentos do produto tem que estar preenchidos. Ex: 1.000.1 está errado. 1.001.1 está correto.
Public Const ERRO_INSERCAO_ARQCONFIG = 1157 'Sem parametros
'Erro na inserção de registro de configuração do Sistema
Public Const ERRO_CODIGO_NAO_LONG = 1158 'Parametro: sNumero
'O número %s não é Long.
Public Const ERRO_LEITURA_MODULOFILEMP_DIC = 1159 'Sem Parametros
'Ocorreu um erro na leitura da tabela ModuloFilEmp do Dicionário de Dados.
Public Const ERRO_ATUALIZACAO_MODULOFILEMP_DIC = 1160 'Sem Parametros
'Ocorreu um erro na atualização da tabela ModuloFilEmp do Dicionário de Dados.
Public Const ERRO_CONFIG_NAO_CADASTRADO = 1161 'Parametros:  sCodigo, sTabelaConfig
'O Código = %s não está cadastrado na tabela %s.
Public Const ERRO_FILIAL_EMPRESA_NAO_PREENCHIDA = 1162 'Sem parâmetros
'A Filial da Empresa não foi informada.
Public Const ERRO_EMPRESA_SEM_FILIAIS = 1163 'Parametro: sCodEmpresa
'A Empresa %s não possui filiais.
Public Const ERRO_USUARIO_SEM_EMPRESA = 1164 'Parametro: sUsuario
'O Usuário %s não está autorizado a acessar nenhuma empresa.
Public Const ERRO_FECHAR_JANELAS_FILHAS = 1165 'sem parametros
'Precisa fechar as outras janelas da aplicação antes de trocar de Empresa ou Filial.
Public Const ERRO_SENHA_EXPIRADA = 1166 'Sem Parametro.
'A Senha deste usuário está expirada.
Public Const ERRO_EMPRESA_NAO_PREENCHIDA = 1167 'Sem Parametro
'A Empresa não está preenchida.
Public Const ERRO_LEITURA_SLDDIAFAT_PROD = 1168 'Sem parâmetros
'Erro na leitura do Faturamento do Produto.
Public Const ERRO_LEITURA_EMPENHO_PROD = 1169
'Erro na leitura de empenhos de produto.
Public Const ERRO_LEITURA_CADREL = 1170
'Erro na leitura do cadastro de relatórios.
Public Const ERRO_RELATORIO_NAO_CADASTRADO = 1171
'O relatório %s não está cadastrado no dicionário de dados.
Public Const ERRO_ATUALIZACAO_RELATORIO = 1172
'Erro na atualização do relatório %s.
Public Const ERRO_TABELA_VAZIA = 1173 'Parametro:sTabela
'Tabela %s do Banco de Dados está vazia.
Public Const ERRO_CATEGORIAPRODUTOITEM_ICMSEXCECOES = 1174 'Parâmetros: CategoriaProduto e CategoriaProdutoItem
'Categoria Produto %s e Categoria Produto Item %s são usados na tabela ICMSExcecoes.
Public Const ERRO_CATEGORIAPRODUTOITEM_IPIEXCECOES = 1175 'Parâmetros: CategoriaProduto e CategoriaProdutoItem
'Categoria Produto %s e Categoria Produto Item %s são usados na tabela IPIExcecoes.
Public Const ERRO_CATEGORIAPRODUTO_IPIEXCECOES = 1176 'Parâmetros: CategoriaProduto
'Categoria Produto %s é usado na tabela IPIExcecoes.
Public Const ERRO_CATEGORIAPRODUTO_ICMSEXCECOES = 1177 'Parâmetros: CategoriaProduto
'Categoria Produto %s é usado na tabela ICMSExcecoes.
Public Const ERRO_LEITURA_RELTELACAMPOS = 1178
'Erro na leitura da estrutura do registro para impressão de relatório
Public Const ERRO_RELTELA_NUM_CAMPOS = 1179
'A quantidade de campos não confere com a obtida no dicionário de dados
Public Const ERRO_RELTELA_TAM_CAMPO = 1180
'O tamanho do campo não confere com o obtida no dicionário de dados
Public Const ERRO_RELTELA_TIPO_CAMPO = 1181
'O tipo do campo não confere com o obtida no dicionário de dados
Public Const ERRO_VALOR_PORCENTAGEM2 = 1182 'dValor
'O valor %d é menor que -99,99.
Public Const ERRO_VALOR_ZERO = 1183 'Sem Parametros
'O Valor não pode ser zero.
Public Const ERRO_PRODUTO_SEG_NUM_CARACTER_INVALIDO = 1184 'Sem parametro
'Os segmentos numéricos de Produto só podem conter números. Ex: 1.-1.1 está errado. 1.1.1 está correto.
Public Const ERRO_ARQUIVO_NAO_ENCONTRADO = 1185 'Parametros sPathCompleto
'O arquivo %s não foi encontrado.
Public Const ERRO_NOME_ARQUIVO_MAIOR_PERMITIDO = 1186 'Parametro sNomeBaseArquivo
'O nome base do arquivo = %s ultrapassa o tamanho máximo permitido de %i caracteres.
Public Const ERRO_NOME_EXTENSAO_RELATORIO_ERRADO = 1187 ' sNomeExtensaoArquivo, sExtensaoPadrao
'A extensão do nome do arquivo = %s difere da extensão obrigatória = %s.
Public Const ERRO_LEITURA_BROWSEINDICEUSUARIO = 1188 'Sem parametro
'Erro na leitura da tabela BrowseIndiceUsuario.
Public Const ERRO_LEITURA_BROWSEINDICESEGMENTOSUSUARIO = 1189 'Sem parametro
'Erro na leitura da tabela BrowseIndiceSegmentosUsuario.
Public Const ERRO_RELATORIO_ORIGINAL_INALTERAVEL = 1190 'Sem parametro
'Não pode alterar ou excluir um relatório original do Sistema
Public Const ERRO_LEITURA_RELATORIOPARAMETROS = 1191 'Sem parametro
'Erro na leitura da tabela RelatorioParametros.
Public Const ERRO_INSERCAO_RELATORIOPARAMETROS = 1192 'Sem parametro
'Erro na inserção da tabela RelatorioParametros.
Public Const ERRO_ATUALIZACAO_RELATORIOPARAMETROS = 1193 'Sem parametro
'Erro na atualização da tabela RelatorioParametros.
Public Const ERRO_LOCK_RELATORIOPARAMETROS = 1194 'Sem parametro
'Erro na tentativa de fazer "lock" na tabela RelatoriosParametros.
Public Const ERRO_EXCLUSAO_RELATORIOPARAMETROS = 1195 'Sem parametro
'Erro na exclusão de um registro da tabela RelatorioParametros.
Public Const ERRO_INSERCAO_BROWSEINDICEUSUARIO = 1196 'Sem parametro
'Ocorreu um erro na inserção de registro na tabela BrowseIndiceUsuario.
Public Const ERRO_EXCLUSAO_BROWSEINDICEUSUARIO = 1197 'Sem parametro
'Ocorreu um erro na exclusão de um registro da tabela BrowseIndiceUsuario.
Public Const ERRO_CAMPO_NAO_PODE_CONTER_ASPAS = 1198 'Sem Parametros
'Este campo não pode conter aspas no seu interior.
Public Const ERRO_LISTA_ORDENACAO_VAZIA = 1199 'Sem Parametros
'A lista ordenação está vazia. Para gravar uma ordenação é necessário preenche-la com pelo menos um campo.
Public Const ERRO_ORDENACAO_NAO_SELECIONADA = 1200 'Sem Parametros
'Não há nenhum elemento selecionado na combo Ordenação. Selecione uma ordenação antes de exclui-la.
Public Const ERRO_ORDENACAO_JA_CADASTRADA = 1201 'Sem Parametros
'Esta ordenação já foi criada pelo usuário.
Public Const ERRO_ORDENACAO_SISTEMA_JA_CADASTRADA = 1202 'Sem Parametros
'Esta ordenação já foi criada pelo sistema.
Public Const ERRO_LEITURA_TABELA_GERACAOARQICMS = 1203 ' Sem Parametros
'Erro na Leitura da Tabela GeracaoArqICMS.
Public Const ERRO_INSERCAO_TABELA_GERACAOARQICMS = 1204 ' Sem Parametros
'Erro na Tentativa de inserir na Tabela GeracaoArqICMS.
Public Const ERRO_LEITURA_TABELA_GERACAOARQICMSPROD = 1205 ' Sem Parametros
'Erro na Leitura da Tabela GeracaoArqICMSProd.
Public Const ERRO_INSERCAO_TABELA_GERACAOARQICMSPROD = 1206 ' Sem Parametros
'Erro na Tentativa de inserir na Tabela GeracaoArqICMSProd.
Public Const ERRO_LEITURA_DIC_ROTINAS = 1207 'sem parametros
'Erro na leitura de rotinas do dicionário de dados
Public Const ERRO_LEITURA_EDICAOTELA = 1208 'Sem parametro
'Erro na leitura da tabela EdicaoTela.
Public Const ERRO_LEITURA_TABINDEX = 1209 'Sem parametro
'Erro na leitura da tabela TabIndex.
Public Const ERRO_INSERCAO_EDICAOTELA = 1210 'Sem parametro
'Erro na inserção da tabela EdicaoTela.
Public Const ERRO_EXCLUSAO_EDICAOTELA = 1211 'Sem parametro
'Erro na exclusão dos registros da tabela EdicaoTela.
Public Const ERRO_INSERCAO_TABINDEX = 1212 'Sem parametro
'Erro na inserção da tabela TabIndex.
Public Const ERRO_EXCLUSAO_TABINDEX = 1213 'Sem parametro
'Erro na exclusão dos registros da tabela TabIndex.
Public Const ERRO_CONTAINER_INVALIDO = 1214 'Sem Parametros
'Atenção! Container Inválido.
Public Const ERRO_MENUITEM_NAO_CADASTRADO = 1215 'Parametro: sTitulo
'Item de menu com título %s não está cadastrado no Sistema.
Public Const ERRO_VALIDADEATE_NAO_INFORMADA = 1216 'Sem Parâmetros
'A Data de Validade deve ser preenchida.
Public Const ERRO_LEITURA_MODULOCLIENTE = 1217 'ParÂmetro: lCodCliente
'Erro na leitura dos módulos liberados para o cliente %l.
Public Const ERRO_LEITURA_CLIENTESLIMITES = 1218 'Parâmetro:lCodCliente
'Erro na leitura da tabela ClientesLimites para o Cliente %l.
Public Const ERRO_SENHA_NAO_GERADA = 1219 'Sem Parâmetros
'A Senha precisa ser gerada.
Public Const ERRO_FILIALCLIENTE_CGC_NAO_ENCONTRADA = 1220 'Parâmetro: sCGC
'Não foi encontrada nenhuma filial do cliente com o CGC = %s.
Public Const ERRO_LIMITEEMPRESAS_NAO_INFORMADO = 1221
'O número limite de empresas não foi informado.
Public Const ERRO_LIMITELOGS_NAO_INFORMADO = 1222
'O número limite de logs não foi informado.
Public Const ERRO_LIMITEFILIAIS_NAO_INFORMADO = 1223
'O número limite de filiais não foi informado.
Public Const ERRO_ATUALIZACAO_CLIENTESLIMITES = 1224 'Parâmetros: lCodCliente
'Erro na atualização dos limites do Cliente %l.
Public Const ERRO_INSERCAO_CLIENTESLIMITES = 1225 'Parâmetro: lCodCliente
'Erro na criação dos limites para o Cliente %l.
Public Const ERRO_EXCLUSAO_MODULOCLIENTE = 1226
'Erro na tentativa de excluir registro da tabela ModuloCliente.
Public Const ERRO_INSERCAO_MODULOCLIENTE = 1227
'Erro na tentativa de inserir registro na tabela ModuloCliente.
Public Const ERRO_LIMITES_CLIENTE_NAO_CADASTRADOS = 1228 'Parâmetro: lCodCliente
'Os Cliente %l não possui cadastro de limites de sistema.
Public Const ERRO_EXCLUSAO_CLIENTESLIMITES = 1229 'Parâmetros: lCodCliente
'Erro na exclusão de registro em ClientesLimites do Cliente %l.
Public Const ERRO_LIMITEEMPRESAS_MAIOR_LIMITEFILIAIS = 1230
'O Limite de empresas não pode ser maior que o limite de filiais.
Public Const ERRO_CONTROLE_SEM_PAI = 1231 'parametros: nome de controle pai e seu indice
'Não foi encontrado controle de menu identificado pelo nome %s e indice %s
Public Const ERRO_LINHA_GRID_NAO_PREENCHIDA = 1232 'Sem parâmetros
'A linha selecionada não contém dados.
Public Const ERRO_FILIALEMPRESA_NAO_INFORMADO = 1233 'Parametro iIndice
'Atenção. A Filial da linha %i não foi informada.
Public Const ERRO_LEITURA_OBJETOBD = 1234
'Erro de leitura na tabela ObjetosBD
Public Const ERRO_ATUALIZACAO_OBJETOSBD = 1235 'Parametro objObjetoBD.iAvisaSobrePosicao
'Erro na atualização do campo %s na tabela ObjetosBD - 'Parametro objObjetoBD.iAvisaSobrePosicao
Public Const ERRO_CLASSEOBJETO_INEXISTENTE = 1236 'Parametro objObjetoBD.sClasseObjeto
'A classe %s não foi encontrada - objObjetoBD.sClasseObjeto


'FERNANDO
Public Const ERRO_EXCEL_EIXO_X_JA_DEFINIDO = 0 'sem parâmetros
'Um gráfico não pode conter duas colunas participando do eixo X.
Public Const ERRO_GRAFICO_VALORES_A_EXIBIR_NAO_DEFINIDOS = 0 'sem parâmetros
'Não foi definida a origem dos valores que serão exibidos no gráfico: Sistema ou Ajustado.
Public Const ERRO_GRAFICO_VALORES_A_EXIBIR_NAO_DEFINIDOS2 = 0 'sem parâmetros
'Não foi definida a origem dos valores que serão exibidos no gráfico: Ajustado ou Real.
Public Const ERRO_VALORES_COLUNAS_NAO_TRATADOS_GRAFICO = 0 'Parâmetro: iColuna
'Os valores da coluna %s não foram tratados para participarem do gráfico.



'VEIO DE ERROS DIC
Public Const ERRO_EMPRESA_NAO_CADASTRADA = 10006 'Parametro lCódigoEmpresa
'Empresa com código = %l não está cadastrada.
Public Const ERRO_LEITURA_USUARIO1 = 10089 'parametro cod do usuario
'Erro na leitura do Usuário %s.
Public Const ERRO_LOCK_FILIALEMPRESA = 10113 'parametro: cod da filial
'Erro no bloqueio de registro da filial %s
Public Const ERRO_LEITURA_EMPRESA_USUARIO = 10130 'Parametro: sCodUsuario
'Ocorreu um erro na leitura das Empresas acessíveis pelo usuário %s.
Public Const ERRO_CGC_NAO_INFORMADO = 10137 'Sem parametro
'Falta informar código de CGC.
Public Const ERRO_NUMERO_SERIE_NAO_PREENCHIDO = 10161
'O número de série deve ser preenchido.


'VEIO DE ERROS CONTAB
Public Const ERRO_FALTA_DE_DADOS = 220 'Sem parametro
'Pelo menos uma linha do grid deve estar preenchida.


'VEIO DE ERROS CPR
Public Const ERRO_PAIS_NAO_CADASTRADO = 2018 'Parametro: iCodigo
'País %i não está cadastrado.
Public Const ERRO_LEITURA_ENDERECOS = 2021
'Erro na leitura da tabela Enderecos.
Public Const ERRO_FILIAL_NAO_PREENCHIDA = 2283 'Sem parâmetro
'A Filial deve ser preenchida.
Public Const ERRO_PAIS_NAO_CADASTRADO1 = 2323 'Parametro: sNome
'País %s não está cadastrado no Banco de Dados.
Public Const ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA = 2375 'Parametro: sNomeReduzido
'O Nome Reduzido tem que começar por uma letra.
Public Const ERRO_CLIENTE_NAO_INFORMADO = 2945 'Sem parametros
'Cliente não foi informado.


'VEIO DE ERROS MAT
Public Const AVISO_CRIAR_PRODUTO = 5813 'Parametro sCodProduto
'Produto %s não está cadastrado no Banco de Dados. Deseja criar novo Produto?
Public Const ERRO_MASCARA_RETORNAPRODUTOENXUTO = 7304   'Parametro: sProduto
'Erro na formatação do produto %s.
Public Const ERRO_MASCARA_MASCARARPRODUTO = 7305 'Parametro: sProduto
'Erro na formatação do Produto %s.
Public Const ERRO_MASCARA_RETORNAPRODUTOPAI = 7306 'Parametro: sProduto
'Erro na função que retorna o produto de nível imediatamente superior do Produto %s.
Public Const ERRO_LEITURA_PRODUTOS1 = 7307 'Sem parametros
'Erro na leitura da tabela de Produtos.
Public Const ERRO_PRODUTOFILIAL_INEXISTENTE = 7937 'Parametros sCodProduto, iFilialEmpresa
'Produto %s da FilialEmpresa %i não está cadastrado no Banco de Dados.
Public Const ERRO_LINHA_GRID_NAO_SELECIONADA = 8570 'Sem parâmetros
'Uma linha do Grid deve estar selecionada.



'VEIO DE ERROS CRFAT
Public Const ERRO_ESTADO_NAO_CADASTRADO = 6027 'Parametro sSiglaEstado
'O Estado %s não está cadastrado.
Public Const ERRO_CLIENTE_NAO_CADASTRADO = 6048 'Parametro: lCodCliente
'O Cliente com código %l não está cadastrado no Banco de Dados.
Public Const ERRO_LEITURA_FILIAISCLIENTES2 = 6285
'Erro na leitura da tabela de FiliaisCliente.
Public Const ERRO_LEITURA_FERIADOS = 6353
'Erro na leitura da tabela Feriados.
Public Const ERRO_LEITURA_ICMSEXCECOES = 6382 'Sem Parâmetros
'Erro de Leitura na Tabela ICMSExcecoes.
Public Const ERRO_CATEGORIACLIENTEITEM_ICMSEXCECOES = 6384 'Parâmetros: CategoriaCliente e CategoriaClienteItem
'Categoria Cliente %s e Categoria Cliente Item %s são usados na tabela ICMSExcecoes.
Public Const ERRO_CATEGORIACLIENTEITEM_IPIEXCECOES = 6385 'Parâmetros: CategoriaCliente e CategoriaClienteItem
'Categoria Cliente %s e Categoria Cliente Item %s são usados na tabela IPIExcecoes.
Public Const ERRO_LEITURA_FILIAISEMPRESAS = 6429 'Sem parâmetros
'Erro na leitura da tabela FiliaisEmpresa.


'VEIO DE ERROS FAT
Public Const ERRO_LEITURA_IPIEXCECOES = 8035 'Sem parâmetro
'Erro na leitura da tabela de exceções de IPI.
Public Const ERRO_DATAVALIDADE_MENOR = 8120 'Parâmetro sDataValidade
'A Data de Validade %s é menor que a Data Corrente
Public Const ERRO_VALOR_PORCENTAGEM3 = 8266 'Parametro dNumero
'O valor %d não está entre 0 e 99.




'Códigos de Aviso - RESERVADO de 5000 a 5099
Public Const AVISO_CONFIRMA_EXCLUSAO_LINHA_GRID = 5000 'Parametro Linha
'Confirma a exclusão da linha %i?
Public Const AVISO_DESEJA_SALVAR_ALTERACOES = 5001 'Sem parametros
'Deseja salvar as alterações realizadas?
Public Const AVISO_CANCELAR_ATUALIZACAO_LOTES = 5002 'Sem parâmetros
'Confirma o canselamento da atualização de lotes ?
Public Const AVISO_EXCLUSAO_RELOPRAZAO = 5003  'Sem parametros
'Confirma a exclusão da Opção de Relatório ?
Public Const AVISO_CANCELAR_GERACAO_ARQ_ICMS = 5004 'Sem Parametros
'Confirma o cancelamento da Geração do Arquivo ICMS ?
Public Const AVISO_NAO_TORNOU_VISIVEL = 5005 'Sem Parametros
'Não foi possível tornar visível este controle. Verifique se ele está contido em um controle que não está visivel.
Public Const AVISO_LIMITES_ALTERADOS = 5006
'Os dados de limites do sistema foram alterados e não foi criada uma nova senha.
'Deseja prosseguir e perder as alterações efetuadas?
Public Const AVISO_SENHA_GERADA = 5007
'Uma nova senha foi gerada e não foi gravada. Deseja prosseguir e perder a alteração?
Public Const AVISO_EXCLUSAO_CLIENTESLIMITES = 5008
'Confirma a exclusão dos dados de limite do Cliente %l?
Public Const AVISO_ALTERACAO_SERIE = 5009 'Parâmetros: sSerieTela, sSerieBD
'A série informada é diferente da série já cadastrada. Série informada: %s e Série Cadastrada: %s. Deseja prosseguir c\ essa alteração?

'Códigos de Erro  RESERVADO de 11000 a 11199
Public Const ERRO_PERCENTUAL_IGUAL_100 = 11000
'O percentual não pode ser igual a 100%.



'VEIO DE ERROS MAT2
Public Const ERRO_LEITURA_PEDIDOCOMPRA = 11203 'Sem parametro
'Erro na leitura da tabela de Pedido de Compra
Public Const ERRO_LEITURA_FORNECEDORPRODUTOFF1 = 11211 'Parâmetros: lCodFornecedor, iCodFilial, sCodProduto
'Erro na leitura da tabela FilialFornecedorProdutoFF com Fornecedor %l, Filial %i e Produto %s.


'VEIO DE ERROS COM
Public Const ERRO_LEITURA_ITENSREQCOMPRA = 12026 'Sem parâmetro
'Erro na leitura da tabela de Itens de Requisições de Compras.
Public Const ERRO_LEITURA_REQUISICAOCOMPRA = 12125 'Parâmetros: lCodigo
'Erro na leitura da tabela Requisição Compra com Requisição de código %l.


