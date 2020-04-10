Attribute VB_Name = "ErrosDic"

'C�digos de Erro  RESERVADO de 10000 a 10499
Public Const ERRO_LEITURA_EMPRESA = 10000 'Par�metro lC�digoEmpresa
'Erro na leitura da Empresa com c�digo = %l.
Public Const ERRO_LEITURA_EMPRESA1 = 10001 'Sem par�metros
'Erro na leitura da Tabela Empresas.
Public Const ERRO_CODIGO_EMPRESA_NAO_INFORMADO = 10002 'Sem par�metros
'C�digo da Empresa n�o foi informado.
Public Const ERRO_NOME_EMPRESA_NAO_INFORMADO = 10003 'Sem par�metros
'Nome da Empresa n�o foi informado.
Public Const ERRO_ATUALIZACAO_EMPRESA = 10004 'Parametro lC�digoEmpresa
'Erro na atualiza��o da Empresa com c�digo = %l.
Public Const ERRO_INSERCAO_EMPRESA = 10005 'Parametro lC�digoEmpresa
'Erro na inclus�o da Empresa com c�digo = %l.
Public Const ERRO_LOCK_EMPRESA = 10007 'Parametro lC�digoEmpresa
'Erro no lock da Empresa com c�digo = %l.
Public Const ERRO_EXCLUSAO_EMPRESA = 10008 'Parametro lC�digoEmpresa
'Erro na exclus�o da Empresa com c�digo = %l.
Public Const ERRO_LEITURA_GRUPO = 10009 'Sem Parametros
'Erro na leitura da tabela Grupos de Usu�rios.
Public Const ERRO_LEITURA_GRUPO1 = 10010 'Parametro sC�digoGrupo
'Erro na leitura do Grupo com c�digo = %s.
Public Const ERRO_LEITURA_USUEMPGRUPO = 10011 'Sem parametro
'Erro na leitura da tabela UsuEmpGrupo.
Public Const ERRO_DATA_NAO_FUTURA = 10012 'Parametro sData
'Data de validade %s n�o � data futura.
Public Const ERRO_GRUPO_NAO_CADASTRADO = 10013 'Parametro sGrupo
'Grupo com c�digo = %s n�o est� cadastrado.
Public Const ERRO_LOCK_GRUPO = 10014 'Parametro sGrupo
'Erro no lock do Grupo com c�digo = %s.
Public Const ERRO_INSERCAO_USUEMPGRUPO = 10015 'Parametros: sCodUsuario, lCodEmpresa, sCodGrupo
'Erro na inclus�o de registro na tabela UsuEmpGrupo.
'Dados da inclus�o: Usu�rio - <sCodUsuario>, Empresa - %l, Grupo - <sCodGrupo>.
Public Const ERRO_LEITURA_USUEMPGRUPO1 = 10016 'Parametros: sCodUsuario, lCodEmpresa
'Erro na leitura da tabela UsuEmpGrupo.
'Registro com chave: c�digo de Usu�rio - %s, c�digo de Empresa - %l.
Public Const ERRO_USUEMPGRUPO_NAO_CADASTRADO = 10017 'Parametros: sCodUsuario, lCodEmpresa
'Registro da tabela UsuEmpGrupo n�o cadastrado. Chave:
'c�digo Usu�rio = %s e c�digo Empresa = %l.
Public Const ERRO_ATUALIZACAO_USUEMPGRUPO = 10018 'Parametros sCodUsuario, lCodEmpresa
'Erro na atualiza��o da Tabela UsuEmpGrupo no registro com chave:
'c�digo Usu�rio = %s e c�digo Empresa = %l.
Public Const ERRO_CODIGO_GRUPO_NAO_INFORMADO = 10019 'Sem parametro
'C�digo de Grupo n�o foi informado.
Public Const ERRO_CODIGO_USUARIO_NAO_INFORMADO = 10020 'Sem parametro
'C�digo de Usu�rio n�o foi informado.
Public Const ERRO_NOME_USUARIO_NAO_INFORMADO = 10021 'Sem parametro
'Nome de Usu�rio n�o foi informado.
Public Const ERRO_SENHA_USUARIO_NAO_INFORMADA = 10022 'Sem parametro
'Senha do Usu�rio n�o foi informada.
Public Const ERRO_EXCLUSAO_USUARIO = 10023 'Parametro sCodUsuario
'Erro na exclus�o do Usu�rio com c�digo = %s.
Public Const ERRO_EMPRESA_NAO_CADASTRADA1 = 10024 'Parametro sNome
'Empresa com nome = %s n�o est� cadastrada.
Public Const ERRO_EXCLUSAO_USUEMPGRUPO = 10025 'Parametro sCodUsuario
'Erro na exclus�o dos registros da tabela UsuEmpGrupo
'com c�digo de Usu�rio = %s.
Public Const ERRO_MODULO_ROTINA_INEXISTENTE = 10026 'Parametro sRotina
'Aus�ncia de M�dulo que cont�m Rotina %s na Tabela Modulos.
Public Const ERRO_ROTINAS_DO_MODULO_INEXISTENTES = 10027 'Parametro sModulo
'Aus�ncia de Rotinas no M�dulo %s.
Public Const ERRO_LEITURA_ROTINA = 10028 'Sem parametros
'Erro na leitura da tabela Rotinas.
Public Const ERRO_LEITURA_ROTINA1 = 10029 'Parametro sRotina
'Erro na leitura da Rotina com sigla=%s.
Public Const ERRO_ROTINA_NAO_CADASTRADA = 10030 'Parametro sRotina
'Rotina com sigla = %s n�o est� cadastrada.
Public Const ERRO_SIGLA_ROTINA_NAO_INFORMADA = 10031 'Sem parametros
'Sigla de Rotina n�o foi informada.
Public Const ERRO_ATUALIZACAO_ROTINA = 10032 'Parametro sRotina
'Erro na atualiza��o de Rotina com sigla=%s.
Public Const ERRO_MODULO_TELA_INEXISTENTE = 10033 'Parametro sTela
'Aus�ncia de M�dulo que cont�m Tela %s na Tabela Modulos.
Public Const ERRO_TELA_NAO_CADASTRADA = 10034 'Parametro sTela
'Tela %s n�o est� cadastrada.
Public Const ERRO_LEITURA_TELA = 10035 'Sem parametros
'Erro na leitura da tabela Telas.
Public Const ERRO_TELAS_DO_MODULO_INEXISTENTES = 10036 'Parametro sModulo
'Aus�ncia de Telas no M�dulo %s.
Public Const ERRO_LEITURA_TELA1 = 10037 'Parametro sTela
'Erro na leitura da Tela %s.
Public Const ERRO_NOME_TELA_NAO_INFORMADO = 10038 'Sem parametros
'Nome da Tela n�o foi informado.
Public Const ERRO_ATUALIZACAO_TELA = 10039 'Parametro sTela
'Erro na atualiza��o da Tela %s.
Public Const ERRO_GRUPOS_NAO_CADASTRADOS = 10040 'Sem parametro
'N�o existem Grupos cadastrados.
Public Const ERRO_USUARIO_SEM_GRUPO_EMPRESA = 10041 'Parametro sCodigoUsuario
'Usu�rio com c�digo=%s est� cadastrado e n�o est� associado a nenhum Grupo/Empresa.
Public Const ERRO_LEITURA_GRUPO_TELA = 10042 'Sem par�metro
'Erro na leitura da tabela GrupoTela.
Public Const ERRO_AUSENCIA_DADOS_GRID_GRUPOS = 10043 'Sem par�metro
'Aus�ncia de dados no Grid de Grupos.
Public Const ERRO_LEITURA_GRUPO_ROTINAS = 10044 'Sem par�metro
'Erro na leitura da tabela GrupoRotinas.
Public Const ERRO_GRUPOROTINA_ROTINA_NAO_CADASTRADOS = 10045 'Par�metro sSiglaRotina
'Registros de GrupoRotinas associados � rotina %s n�o est�o cadastrados.
Public Const ERRO_FALTA_GRUPO_NA_COLECAO = 10046 'Parametro sCodGrupo
'Falta grupo %s na cole��o colGrupoRotina.
Public Const ERRO_CODIGOS_GRUPO_DIFERENTES = 10047 'Par�metros: sCodGrupo1, sCodGrupo2
'C�digo de grupo %s1 lido no BD n�o corresponde ao c�digo %s2 do Grid.
Public Const ERRO_ATUALIZACAO_GRUPO_ROTINAS = 10048 'Par�metros: sSiglaRotina, sCodGrupo
'Erro na atualiza��o da tabela GrupoRotinas, no registro com chave:
'CodGrupo=<sCodGrupo>, SiglaRotina=<sSiglaRotina>
Public Const ERRO_TELA_NAO_INFORMADA = 10049 'Sem parametros
'Nome de Tela n�o foi informado.
Public Const ERRO_GRUPOTELA_TELA_NAO_CADASTRADOS = 10050 'Par�metro sTela
'Registros de GrupoTela associados � tela %s n�o est�o cadastrados.
Public Const ERRO_FALTA_GRUPO_NA_COLECAO2 = 10051 'Parametro sCodGrupo
'Falta grupo %s na cole��o colGrupoTela.
Public Const ERRO_ATUALIZACAO_GRUPO_TELA = 10052 'Par�metros: sTela, sCodGrupo
'Erro na atualiza��o da tabela GrupoTela, no registro com chave:
'CodGrupo=<sCodGrupo>, NomeTela=<sTela>
Public Const ERRO_GRUPOROTINA_GRUPOMODULO_NAO_CADASTRADOS = 10053 'Par�metros: sGrupo, sModulo
'Registros de GrupoRotinas associados ao Grupo <sGrupo>
'e a rotinas no M�dulo <sModulo> n�o est�o cadastrados.
Public Const ERRO_ROTINAS_INEXISTENTES = 10054 'Sem par�metros
'Tabela Rotinas est� vazia.
Public Const ERRO_GRUPO_NAO_INFORMADO = 10055 'Sem par�metros
'Grupo n�o foi informado.
Public Const ERRO_AUSENCIA_DADOS_GRID_ROTINAS = 10056 'Sem par�metro
'Aus�ncia de dados no Grid de Rotinas.
Public Const ERRO_FALTA_ROTINA_NA_COLECAO = 10057 'Par�metro: sSiglaRotina
'Falta rotina %s na cole��o colGrupoRotina, preenchida do GRID.
Public Const ERRO_SIGLAS_ROTINA_DIFERENTES = 10058 'Par�metros: sSiglaRotina1, sSiglaRotina2
'Sigla de Rotina %s1 lido no BD n�o corresponde � Sigla de Rotina %s2 do Grid.
Public Const ERRO_AUSENCIA_DADOS_GRID_TELAS = 10059 'Sem par�metro
'Aus�ncia de dados no Grid de Telas.
Public Const ERRO_GRUPOTELA_GRUPOMODULO_NAO_CADASTRADOS = 10060 'Par�metros: sGrupo, sModulo
'Registros de GrupoTela associados ao Grupo <sGrupo>
'e a telas no M�dulo <sModulo> n�o est�o cadastrados.
Public Const ERRO_TELAS_INEXISTENTES = 10061 'Sem par�metros
'Tabela Telas est� vazia.
Public Const ERRO_FALTA_TELA_NA_COLECAO = 10062 'Par�metro: sNomeTela
'Falta tela %s na cole��o colGrupoTela, preenchida do GRID.
Public Const ERRO_NOMES_TELA_DIFERENTES = 10063 'Par�metros: sNomeTela1, sNomeTela2
'Nome de Tela %s1 lido no BD n�o corresponde ao Nome de Tela %s2 do Grid.
Public Const ERRO_TIPO_ROTINA_NAO_INFORMADO = 10064 'Sem parametros
'Tipo de Rotina n�o foi informado.
Public Const ERRO_AUSENCIA_DADOS_GRID_MENUITENS = 10065 'Sem parametros
'Aus�ncia de dados no Grid de �tens de Menu.
Public Const ERRO_LEITURA_MENU_ITENS = 10066 'Sem par�metro
'Erro na leitura da tabela MenuItens.
Public Const ERRO_MENUITEM_ROTINAS_MODULO_NAO_CADASTRADOS = 10067 'Par�metros: sTipoRotina, sModulo
'Registros de MenuItens associados a rotinas de Usu�rio do tipo <sTipoRotina> no M�dulo <sModulo> n�o est�o cadastrados.
Public Const ERRO_FALTA_MENUITEM_NA_COLECAO = 10068 'par�metro: sSiglaRotina
'Falta MenuItem correspondente � rotina %s na cole��o colMenuItem, preenchida do GRID.
Public Const ERRO_ATUALIZACAO_MENU_ITENS = 10069 'par�metro: sSiglaRotina
'Erro na atualiza��o da tabela MenuItens, no registro correspondente � rotina %s.
Public Const ERRO_ARQUIVO_NAO_INFORMADO = 10070 'Sem parametros
'Nome de Arquivo n�o foi informado.
Public Const ERRO_CAMPO_NAO_SELECIONADO = 10071 'Sem parametros
'Pelo menos um campo deve ser selecionado.
Public Const ERRO_AUSENCIA_CAMPO_GRUPOBROWSECAMPO = 10072 'Parametros: sGrupo, sTela, sArquivo
'Aus�ncia de nomes de Campos na tabela GrupoBrowseCampo
'associados a Grupo: <sGrupo>, Tela: <sTela>, Arquivo: <sArquivo>.
Public Const ERRO_AUSENCIA_ARQUIVO_BROWSEARQUIVO = 10073 'Parametro: sTela
'Aus�ncia de nomes de Arquivos associados � tela %s na tabela BrowseArquivo.
Public Const ERRO_AUSENCIA_TELAS_BROWSEARQUIVO = 10074 'parametro: sModulo
'Aus�ncia de nomes de Telas de browse do m�dulo %s na tabela BrowseArquivo.
Public Const ERRO_AUSENCIA_CAMPO_CAMPOS = 10075 'parametro: sArquivo
'Aus�ncia de Campos do Arquivo %s na tabela Campos.
Public Const ERRO_EXCLUSAO_GRUPOBROWSECAMPO = 10076 'sem parametros
'Erro na exclus�o de registros da tabela GrupoBrowseCampo.
Public Const ERRO_INCLUSAO_GRUPOBROWSECAMPO = 10077 'parametros: sGrupo, sTela, sArquivo, vCampo
'Erro na inclus�o de registro na tabela GrupoBrowseCampo.
'Dados da inclus�o: Grupo=<sGrupo>, Tela=<sTela>, Arquivo=<sArquivo>, Campo=<vCampo>.
Public Const ERRO_EXCLUSAO_BROWSEUSUARIOCAMPO2 = 10078 'parametro: sCodUsuario
'Erro na exclus�o dos registros da tabela BrowseUsuarioCampo
'com c�digo de Usu�rio = %s.
Public Const ERRO_EXCLUSAO_BROWSEUSUARIOORDENACAO = 10079 'parametro: sCodUsuario
'Erro na exclus�o dos registros da tabela BrowseUsuarioOrdenacao
'com c�digo de Usu�rio = %s.
Public Const ERRO_INSERCAO_GRUPOREL = 10080 'parametros cod grupo e cod do rel
'Erro na tentativa de inserir registros na tabela de Grupos x Relat�rios. Grupo %s e Relat�rio %s.
Public Const ERRO_INSERCAO_GRUPOBROWSECAMPO = 10081 'parametros cod grupo e nome da tela
'Erro na tentativa de inserir registros na tabela de Grupos x Campos das Telas de Browse. Grupo %s e Tela %s.
Public Const ERRO_LEITURA_USU_GRUPO = 10082 'parametro cod do grupo
'Erro na leitura de usu�rio do grupo %s.
Public Const ERRO_LEITURA_USUARIOS_GRUPO = 10083 'parametro cod do grupo
'Erro na leitura dos usu�rios do grupo %s.
Public Const ERRO_ATUALIZACAO_GRUPO = 10084 'parametro cod do grupo
'Erro na tentativa de atualizar o Grupo %s.
Public Const ERRO_INSERCAO_GRUPO = 10085 'parametro cod do grupo
'Erro na tentativa de inserir o Grupo %s.
Public Const ERRO_INSERCAO_GRUPOROTINA = 10086 'parametros cod do grupo e da rotina
'Erro na tentativa de inserir registros na tabela de Grupos x Rotinas. Grupo %s e Rotina %s.
Public Const ERRO_INSERCAO_GRUPOTELA = 10087 'parametros cod do grupo e da tela
'Erro na tentativa de inserir registros na tabela de Grupos x Telas. Grupo %s e Tela %s.
Public Const ERRO_EXCLUSAO_GRUPO = 10088 'parametro cod do grupo
'Erro na exclus�o do Grupo %s.
Public Const ERRO_EXCLUSAO_GRUPOTELA = 10090 'parametro cod do grupo
'Erro na tentativa de excluir registro da tabela de Grupos x Telas. Grupo %s.
Public Const ERRO_LEITURA_USUFILEMP_USU = 10091 'parametro: codusuario
'Erro na leitura de permiss�es de acesso � filiais/Empresas para o usu�rio %s.
Public Const ERRO_AUSENCIA_DADOS_GRID = 10092 'sem parametros
'O grid precisa estar preenchido.
Public Const ERRO_RELATORIOS_DO_MODULO_INEXISTENTES = 10093 'sem parametros
'N�o existem relat�rios associados ao m�dulo
Public Const ERRO_RELATORIO_NAO_INFORMADO = 10094 'sem parametros
'Selecione um relat�rio
Public Const ERRO_LEITURA_GRUPO_RELATORIOS = 10095 'sem parametros
'Erro na leitura de registros da tabela GrupoRelatorios
Public Const ERRO_GRUPORELATORIO_REL_NAO_CADASTRADOS = 10096 'parametro = CodRel
'N�o h� registros na tabela GrupoRelatorios para o relat�rio %s.
Public Const ERRO_LEITURA_MODULOFILEMP = 10097
'Erro na leitura de registros na tabela ModulosFilEmp
Public Const ERRO_INSERCAO_MODULOFILEMP = 10098
'Erro na inser��o de registros na tabela ModulosFilEmp
Public Const ERRO_EXCLUSAO_MODULOFILEMP = 10099
'Erro na exclus�o de registros na tabela ModulosFilEmp
Public Const ERRO_FILIALEMPRESA_INATIVA = 10100 'sem parametros
'N�o se pode alterar ou excluir uma filial inativa
Public Const ERRO_FILIAL_MESMO_NOME = 10101 'sem parametros
'N�o pode haver duas filiais com o mesmo nome.
Public Const ERRO_LIMITE_FILIAISEMPRESA = 10102 'parametro: limite
'N�o pode ultrapassar o limite de filiais que � de %s.
Public Const ERRO_LEITURA_FILIAISEMPRESA = 10103 'sem parametro
'Erro na leitura da tabela de FiliaisEmpresa
Public Const ERRO_LEITURA_DICCONFIG = 10104 'sem parametro
'Erro na leitura de configura��o do dicion�rio de dados
Public Const ERRO_EXCLUSAO_USUFILEMP_USU = 10105 'parametro = codusurio
'Erro na exclus�o dos direitos de acesso � Empresas e Filiais para o usu�rio %s.
Public Const ERRO_INSERCAO_USUFILEMP_USU = 10106 'parametro = codusurio
'Erro na inclus�o dos direitos de acesso � Empresas e Filiais para o usu�rio %s.
Public Const ERRO_EXCLUSAO_BROWSEUSUARIOCAMPO3 = 10107 'sem parametros
'Erro na exclus�o de registros da tabela BrowseUsuarioCampo
Public Const ERRO_ATUALIZACAO_GRUPO_RELATORIOS = 10108 'parametros: sCodRel, sCodGrupo
'Erro na atualiza��o de registro em GrupoRelatorios para o relat�rio %s, grupo %s
Public Const ERRO_CODREL_DIFERENTE = 10109 'parametros: codrel no bd, codrel na colecao
'Altera��es no banco de dados ap�s a carga da tela impedem a atualiza��o dos dados.
Public Const ERRO_GRUPORELATORIO_NAO_CADASTRADO = 10110 'sem parametro
'N�o encontrou registro na tabela GrupoRelatorios
Public Const ERRO_ATUALIZACAO_FILIALEMPRESA = 10111 'parametro: cod da filial
'Erro na atualiza��o de registro da filial %s
Public Const ERRO_INSERCAO_FILIALEMPRESA = 10112 'parametro: cod da filial
'Erro na inser��o de registro da filial %s
Public Const ERRO_EXCLUSAO_USUFILEMP = 10114 'sem parametro
'Erro na exclusao de registros da tabela UsuFilEmp
Public Const ERRO_MODULO_VINCULADO_FILIAL = 10115 'Parametros: sSigla, lCodEmpresa, iCodFilial
'M�dulo %s n�o pode ser desativado pois est� vinculado a FilialEmpresa, Empresa c�digo %l, Filial c�digo %i.
Public Const ERRO_NOMERED_EMPRESA_NAO_INFORMADO = 10116 'Sem par�metros
'Nome Reduzido da Empresa n�o foi informado.
Public Const ERRO_EMPRESA_NOME_JA_EXISTE = 10117 'Parametro: sNome
'Existe Empresa com nome %s no Sistema.
Public Const ERRO_EMPRESA_NOME_RED_JA_EXISTE = 10118 'Parametro: sNomeRed
'Existe Empresa com nome reduzido %s no Sistema.
Public Const ERRO_INSERCAO_MODULO_EMPRESA = 10119 'Parametros: lCodigo, sSigla
'Erro na tentativa de inserir na tabela ModuloEmpresa registro com chave CodEmpresa=%l, SiglaModulo=%s.
Public Const ERRO_INSERCAO_USUFILEMP = 10120 'Sem parametros
'Erro na tentativa de inserir registro na tabela UsuFilEmp.
Public Const ERRO_ALTERACAO_EMPRESA_INATIVA = 10121 'Parametro:lCodigo
'N�o � poss�vel alterar Empresa Inativa %l.
Public Const ERRO_ALTERACAO_NOME = 10122 'Parametro: lCodigo
'N�o � poss�vel alterar Nome da Empresa %l.
Public Const ERRO_LEITURA_MODULO_EMPRESA = 10123 'Parametro:lCodigo
'Erro na leitura da tabela ModuloEmpresa para Empresa com c�digo %l.
Public Const ERRO_ATUALIZACAO_MODULO_EMPRESA = 10124 'Parametros: lCodigo, sSigla
'Erro na atualiza��o na tabela ModuloEmpresa, Empresa com c�digo %l, M�dulo com sigla %s.
Public Const ERRO_LEITURA_FILIAIS_EMPRESAS = 10125 'Par�metro: lCodEmpresa
'Erro na leitura das Filiais da Empresa %l na tabela de Filiais Empresas.
Public Const ERRO_LOCK_FILIAL_EMPRESA = 10126 'Par�metros: lCodFilialEmpresa, lCodEmpresa
'Erro na tentativa de fazer "lock" na Filial %l da Empresa %l na tabela de Filiais Empresas.
Public Const ERRO_EXCLUSAO_FILIAL_EMPRESA = 10127 'Par�metros: lCodFilialEmpresa, lCodEmpresa
'Erro na exclus�o da Filial %l da Empresa %l na tabela de Filiais Empresas.
Public Const ERRO_LEITURA_USUFILEMP = 10128 'Par�metro: lCodEmpresa
'Erro na tentativa de leitura dos usu�rios vinculados � Empresa %l na tabela UsuFilEmp.
Public Const ERRO_LOCK_USUFILEMP = 10129 'Par�metro: lEmpresa
'Erro na tentativa de fazer "lock" na tabela UsuFilEmp onde o C�digo da Empresa = %l.
Public Const ERRO_GRUPORELATORIO_GRUPOMODULO_NAO_CADASTRADOS = 10132 'parametros:sGrupo, sModulo
'N�o h� relat�rios cadastrados para o grupo '%s' m�dulo '%s'.
Public Const ERRO_AUSENCIA_EMPRESAS = 10133 'Sem parametros
'N�o est�o cadastradas Empresas no Sistema.
Public Const ERRO_FILIAL_EMPRESA_NAO_CADASTRADA1 = 10134 'Parametros: lCodigo, iCodFilial
'Filial com c�digo %i da Empresa com c�digo %l n�o est� cadastrada no Sistema.
Public Const ERRO_CODIGO_FILIALEMPRESA_NAO_INFORMADO = 10135 'Sem parametros
'Falta informar C�digo de Filial Empresa.
Public Const ERRO_NOME_FILIALEMPRESA_NAO_INFORMADO = 10136 'Sem parametro
'Falta informar Nome da Filial Empresa.
Public Const ERRO_ESTADO_NAO_INFORMADO_PRINCIPAL = 10138 'Sem parametro
'Falta informar o Estado no Endere�o principal.
Public Const ERRO_ESTADO_NAO_INFORMADO_ENTREGA = 10139 'Sem parametro
'Falta informar o Estado no Endere�o de entrega.
Public Const ERRO_PAIS_NAO_INFORMADO_PRINCIPAL = 10140 'Sem parametro
'Falta informar o Pa�s no Endere�o principal.
Public Const ERRO_PAIS_NAO_INFORMADO_ENTREGA = 10141 'Sem parametro
'Falta informar o Pa�s no Endere�o de entrega.
Public Const ERRO_LEITURA_MODULOFILEMP1 = 10142 'Parametros: sSigla, lCodigo, iFilial
'Erro de leitura na tabela ModuloFilEmp. Chave: SiglaModulo = %s,  CodEmpresa = %l, CodFilial = %i.
Public Const ERRO_LOCK_MODULOFILEMP1 = 10143 'Parametros: sSigla, lCodigo, iFilial
'Erro na de "lock" na tabela ModuloFilEmp. Chave: SiglaModulo = %s,  CodEmpresa = %l, CodFilial = %i.
Public Const ERRO_EXCLUSAO_MODULOFILEMP1 = 10144 'Parametros: sSigla, lCodigo, iFilial
'Erro na exclus�o de registro da tabela ModuloFilEmp. Chave: SiglaModulo = %s,  CodEmpresa = %l, CodFilial = %i.
Public Const ERRO_CODGRUPO_NAO_INFORMADO = 10145
'O c�digo do grupo de usu�rios n�o foi informado.
Public Const ERRO_NOMERED_USUARIO_NAO_INFORMADO = 10146 'Sem parametro
'Nome Reduzido do Usu�rio n�o foi informado.
Public Const ERRO_DATA_FORA_VALIDADE = 10147 'Par�metro: dtData, dtValidadeDe, dtValidadeAte
'A data %dt n�o est� dentro da validade do sistema que vai de %dt at� %dt.
Public Const ERRO_LIMITEEMPRESAS_ATINGIDO = 10148 'Par�metro: iLimiteEmpresas
'N�o � poss�vel criar nova empresa pois o n�mero limite que � de %i j� foi atingido.
Public Const ERRO_LIMITEFILIAIS_ATINGIDO = 10149 'Par�metro: iLimiteFiliais
'N�o � poss�vel criar nova filial pois o n�mero limite que � de %i j� foi atingido.
Public Const ERRO_LIMITELOGS_ATINGIDO = 10150 'Par�metro: iLimiteLogs
'N�o � poss�vel entrar no sistema pois o n�mero limite  de usu�rios logados que � de %i j� foi atingido.
Public Const ERRO_NUMEMPRESAS_MAIOR_LIMITE = 10151 'Par�metro: iNumeroEmpresasBD, iLimiteEmpresa
'O n�mero de empresas no BD que � de %i � superior ao limite a ser implantado que � de %i.
'Exclua as empresas excedentes e tente configurar novamente.
Public Const ERRO_NUMFILIAISEMPRESAS_MAIOR_LIMITE = 10152 'Par�metro: iNumeroFiliaisBD, iLimiteFiliais
'O n�mero de filiais de empresas no BD que � de %i � superior ao limite a ser implantado que � de %i.
'Exclua as filiais excedentes e tente configurar novamente.
Public Const ERRO_ATUALIZACAO_DICCONFIG = 10153
'Erro ao tentar atualizar as informa��es de Configura��o do Sistema.
Public Const ERRO_ATUALIZACAO_MODULOS = 10154
'Erro de atualiza��o na tabela de Modulos.
Public Const ERRO_REGISTRO_DICCONFIG_NAO_ENCONTRADO = 10155
'O Registro de configura��o do sistema n�o foi encontrado.
Public Const ERRO_ODBC_OBTER_HANDLE = 10156 'sem parametros
'Erro na obten��o de recurso para acesso ao banco de dados.
Public Const ERRO_USUARIO_CADASTRADO_COMPRADOR = 10157 'Parametros: sCodUsuario, gsNomeEmpresa, gsNomeFilialEmpresa
'O Usuario sCodUsuario � Comprador da gsNomeEmpresa, Filial gsNomeFilialEmpresa, por isso n�o � possivel Desabilita - lo nesta Filial.
Public Const ERRO_LEITURA_TABELA_COMPRADOR = 10158 'Parametros: gsNomeEmpresa, gsNomeFilialEmpresa
'Erro na Leitura da Tabela de Comprador da Empresa gsNomeEmpresa, Filial gsNomeFilialEmpresa.
Public Const ERRO_LEITURA_VERSAO = 10159 'Sem parametro
'Erro de leitura na tabela de Vers�o.
Public Const ERRO_VERSAO_VIGENTE_AUSENTE = 10160 'Sem parametro
'Aus�ncia de vers�o vigente na tabela de Vers�o.
Public Const ERRO_TRECHO_SENHA_INCOMPLETO = 10162 'Par�metro: iTrecho
'O trecho %i da senha est� incompleto.
Public Const ERRO_NUMERO_SERIE_DIFERENTE_BD = 10163
'O n�mero de s�rie informado � diferente do que consta no Banco de Dados.
Public Const ERRO_DATA_SENHA_BD_MAIOR = 10164
'A data atual � inferior a da ultima configura��o.







'C�digos de Aviso RESERVADO de 10500 a 10999
Public Const EXCLUSAO_EMPRESA = 10500 'Parametros: lCodigoEmpresa , sNomeEmpresa
'A Empresa com c�digo: %l e nome: %s ser� exclu�da. Aten��o: TODOS as suas tabelas ser�o exclu�das.
'Confirma a exclus�o?
Public Const AVISO_GRUPO_INEXISTENTE = 10501 'Parametro: sCodigoGrupo
'O Grupo com c�digo = %s n�o est� cadastrado. Deseja cadastr�-lo?
Public Const EXCLUSAO_USUARIO = 10502 'Parametros: sCodUsuario, sNome
'Confirma a exclus�o de Usu�rio com c�digo=%s e nome=%s ?
Public Const AVISO_EXCLUSAO_GRUPO = 10503 'parametro cod do grupo
'Confirma a exclus�o do Grupo ?
Public Const AVISO_CRIAR_EMPRESA = 10504 'Parametro sCodEmpresa
'Empresa com c�digo %s n�o est� cadastrada no Sistema. Deseja criar?
Public Const AVISO_EXCLUSAO_FILIALEMPRESA = 10505 'Parametros: iCodFilial, lCodEmpresa
'Confirma exclus�o de Filial com c�digo %i da Empresa com c�digo %l ?
Public Const AVISO_EXCLUSAO_EMPRESA = 10506 'Parametros: lCodigo, sNome
'Confirma exclus�o de Empresa com c�digo %l e nome %s ?
Public Const AVISO_EXCLUSAO_USUARIOS = 10507 'Par�metros: sCodGrupo
'O Grupo de c�digo %s possui Usu�rios cadastrados. Confirma exclus�o do Grupo?
Public Const AVISO_EXCLUSAO_FILIAIS = 10508 'Par�metros: lCodigoEmpresa, sNomeEmpresa
'A empresa %s de c�digo %l possui filiais cadastradas. Confirma exclus�o?


