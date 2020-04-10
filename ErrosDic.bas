Attribute VB_Name = "ErrosDic"

'Códigos de Erro  RESERVADO de 10000 a 10499
Public Const ERRO_LEITURA_EMPRESA = 10000 'Parâmetro lCódigoEmpresa
'Erro na leitura da Empresa com código = %l.
Public Const ERRO_LEITURA_EMPRESA1 = 10001 'Sem parâmetros
'Erro na leitura da Tabela Empresas.
Public Const ERRO_CODIGO_EMPRESA_NAO_INFORMADO = 10002 'Sem parâmetros
'Código da Empresa não foi informado.
Public Const ERRO_NOME_EMPRESA_NAO_INFORMADO = 10003 'Sem parâmetros
'Nome da Empresa não foi informado.
Public Const ERRO_ATUALIZACAO_EMPRESA = 10004 'Parametro lCódigoEmpresa
'Erro na atualização da Empresa com código = %l.
Public Const ERRO_INSERCAO_EMPRESA = 10005 'Parametro lCódigoEmpresa
'Erro na inclusão da Empresa com código = %l.
Public Const ERRO_LOCK_EMPRESA = 10007 'Parametro lCódigoEmpresa
'Erro no lock da Empresa com código = %l.
Public Const ERRO_EXCLUSAO_EMPRESA = 10008 'Parametro lCódigoEmpresa
'Erro na exclusão da Empresa com código = %l.
Public Const ERRO_LEITURA_GRUPO = 10009 'Sem Parametros
'Erro na leitura da tabela Grupos de Usuários.
Public Const ERRO_LEITURA_GRUPO1 = 10010 'Parametro sCódigoGrupo
'Erro na leitura do Grupo com código = %s.
Public Const ERRO_LEITURA_USUEMPGRUPO = 10011 'Sem parametro
'Erro na leitura da tabela UsuEmpGrupo.
Public Const ERRO_DATA_NAO_FUTURA = 10012 'Parametro sData
'Data de validade %s não é data futura.
Public Const ERRO_GRUPO_NAO_CADASTRADO = 10013 'Parametro sGrupo
'Grupo com código = %s não está cadastrado.
Public Const ERRO_LOCK_GRUPO = 10014 'Parametro sGrupo
'Erro no lock do Grupo com código = %s.
Public Const ERRO_INSERCAO_USUEMPGRUPO = 10015 'Parametros: sCodUsuario, lCodEmpresa, sCodGrupo
'Erro na inclusão de registro na tabela UsuEmpGrupo.
'Dados da inclusão: Usuário - <sCodUsuario>, Empresa - %l, Grupo - <sCodGrupo>.
Public Const ERRO_LEITURA_USUEMPGRUPO1 = 10016 'Parametros: sCodUsuario, lCodEmpresa
'Erro na leitura da tabela UsuEmpGrupo.
'Registro com chave: código de Usuário - %s, código de Empresa - %l.
Public Const ERRO_USUEMPGRUPO_NAO_CADASTRADO = 10017 'Parametros: sCodUsuario, lCodEmpresa
'Registro da tabela UsuEmpGrupo não cadastrado. Chave:
'código Usuário = %s e código Empresa = %l.
Public Const ERRO_ATUALIZACAO_USUEMPGRUPO = 10018 'Parametros sCodUsuario, lCodEmpresa
'Erro na atualização da Tabela UsuEmpGrupo no registro com chave:
'código Usuário = %s e código Empresa = %l.
Public Const ERRO_CODIGO_GRUPO_NAO_INFORMADO = 10019 'Sem parametro
'Código de Grupo não foi informado.
Public Const ERRO_CODIGO_USUARIO_NAO_INFORMADO = 10020 'Sem parametro
'Código de Usuário não foi informado.
Public Const ERRO_NOME_USUARIO_NAO_INFORMADO = 10021 'Sem parametro
'Nome de Usuário não foi informado.
Public Const ERRO_SENHA_USUARIO_NAO_INFORMADA = 10022 'Sem parametro
'Senha do Usuário não foi informada.
Public Const ERRO_EXCLUSAO_USUARIO = 10023 'Parametro sCodUsuario
'Erro na exclusão do Usuário com código = %s.
Public Const ERRO_EMPRESA_NAO_CADASTRADA1 = 10024 'Parametro sNome
'Empresa com nome = %s não está cadastrada.
Public Const ERRO_EXCLUSAO_USUEMPGRUPO = 10025 'Parametro sCodUsuario
'Erro na exclusão dos registros da tabela UsuEmpGrupo
'com código de Usuário = %s.
Public Const ERRO_MODULO_ROTINA_INEXISTENTE = 10026 'Parametro sRotina
'Ausência de Módulo que contém Rotina %s na Tabela Modulos.
Public Const ERRO_ROTINAS_DO_MODULO_INEXISTENTES = 10027 'Parametro sModulo
'Ausência de Rotinas no Módulo %s.
Public Const ERRO_LEITURA_ROTINA = 10028 'Sem parametros
'Erro na leitura da tabela Rotinas.
Public Const ERRO_LEITURA_ROTINA1 = 10029 'Parametro sRotina
'Erro na leitura da Rotina com sigla=%s.
Public Const ERRO_ROTINA_NAO_CADASTRADA = 10030 'Parametro sRotina
'Rotina com sigla = %s não está cadastrada.
Public Const ERRO_SIGLA_ROTINA_NAO_INFORMADA = 10031 'Sem parametros
'Sigla de Rotina não foi informada.
Public Const ERRO_ATUALIZACAO_ROTINA = 10032 'Parametro sRotina
'Erro na atualização de Rotina com sigla=%s.
Public Const ERRO_MODULO_TELA_INEXISTENTE = 10033 'Parametro sTela
'Ausência de Módulo que contém Tela %s na Tabela Modulos.
Public Const ERRO_TELA_NAO_CADASTRADA = 10034 'Parametro sTela
'Tela %s não está cadastrada.
Public Const ERRO_LEITURA_TELA = 10035 'Sem parametros
'Erro na leitura da tabela Telas.
Public Const ERRO_TELAS_DO_MODULO_INEXISTENTES = 10036 'Parametro sModulo
'Ausência de Telas no Módulo %s.
Public Const ERRO_LEITURA_TELA1 = 10037 'Parametro sTela
'Erro na leitura da Tela %s.
Public Const ERRO_NOME_TELA_NAO_INFORMADO = 10038 'Sem parametros
'Nome da Tela não foi informado.
Public Const ERRO_ATUALIZACAO_TELA = 10039 'Parametro sTela
'Erro na atualização da Tela %s.
Public Const ERRO_GRUPOS_NAO_CADASTRADOS = 10040 'Sem parametro
'Não existem Grupos cadastrados.
Public Const ERRO_USUARIO_SEM_GRUPO_EMPRESA = 10041 'Parametro sCodigoUsuario
'Usuário com código=%s está cadastrado e não está associado a nenhum Grupo/Empresa.
Public Const ERRO_LEITURA_GRUPO_TELA = 10042 'Sem parâmetro
'Erro na leitura da tabela GrupoTela.
Public Const ERRO_AUSENCIA_DADOS_GRID_GRUPOS = 10043 'Sem parâmetro
'Ausência de dados no Grid de Grupos.
Public Const ERRO_LEITURA_GRUPO_ROTINAS = 10044 'Sem parâmetro
'Erro na leitura da tabela GrupoRotinas.
Public Const ERRO_GRUPOROTINA_ROTINA_NAO_CADASTRADOS = 10045 'Parâmetro sSiglaRotina
'Registros de GrupoRotinas associados à rotina %s não estão cadastrados.
Public Const ERRO_FALTA_GRUPO_NA_COLECAO = 10046 'Parametro sCodGrupo
'Falta grupo %s na coleção colGrupoRotina.
Public Const ERRO_CODIGOS_GRUPO_DIFERENTES = 10047 'Parâmetros: sCodGrupo1, sCodGrupo2
'Código de grupo %s1 lido no BD não corresponde ao código %s2 do Grid.
Public Const ERRO_ATUALIZACAO_GRUPO_ROTINAS = 10048 'Parâmetros: sSiglaRotina, sCodGrupo
'Erro na atualização da tabela GrupoRotinas, no registro com chave:
'CodGrupo=<sCodGrupo>, SiglaRotina=<sSiglaRotina>
Public Const ERRO_TELA_NAO_INFORMADA = 10049 'Sem parametros
'Nome de Tela não foi informado.
Public Const ERRO_GRUPOTELA_TELA_NAO_CADASTRADOS = 10050 'Parâmetro sTela
'Registros de GrupoTela associados à tela %s não estão cadastrados.
Public Const ERRO_FALTA_GRUPO_NA_COLECAO2 = 10051 'Parametro sCodGrupo
'Falta grupo %s na coleção colGrupoTela.
Public Const ERRO_ATUALIZACAO_GRUPO_TELA = 10052 'Parâmetros: sTela, sCodGrupo
'Erro na atualização da tabela GrupoTela, no registro com chave:
'CodGrupo=<sCodGrupo>, NomeTela=<sTela>
Public Const ERRO_GRUPOROTINA_GRUPOMODULO_NAO_CADASTRADOS = 10053 'Parâmetros: sGrupo, sModulo
'Registros de GrupoRotinas associados ao Grupo <sGrupo>
'e a rotinas no Módulo <sModulo> não estão cadastrados.
Public Const ERRO_ROTINAS_INEXISTENTES = 10054 'Sem parâmetros
'Tabela Rotinas está vazia.
Public Const ERRO_GRUPO_NAO_INFORMADO = 10055 'Sem parâmetros
'Grupo não foi informado.
Public Const ERRO_AUSENCIA_DADOS_GRID_ROTINAS = 10056 'Sem parâmetro
'Ausência de dados no Grid de Rotinas.
Public Const ERRO_FALTA_ROTINA_NA_COLECAO = 10057 'Parâmetro: sSiglaRotina
'Falta rotina %s na coleção colGrupoRotina, preenchida do GRID.
Public Const ERRO_SIGLAS_ROTINA_DIFERENTES = 10058 'Parâmetros: sSiglaRotina1, sSiglaRotina2
'Sigla de Rotina %s1 lido no BD não corresponde à Sigla de Rotina %s2 do Grid.
Public Const ERRO_AUSENCIA_DADOS_GRID_TELAS = 10059 'Sem parâmetro
'Ausência de dados no Grid de Telas.
Public Const ERRO_GRUPOTELA_GRUPOMODULO_NAO_CADASTRADOS = 10060 'Parâmetros: sGrupo, sModulo
'Registros de GrupoTela associados ao Grupo <sGrupo>
'e a telas no Módulo <sModulo> não estão cadastrados.
Public Const ERRO_TELAS_INEXISTENTES = 10061 'Sem parâmetros
'Tabela Telas está vazia.
Public Const ERRO_FALTA_TELA_NA_COLECAO = 10062 'Parâmetro: sNomeTela
'Falta tela %s na coleção colGrupoTela, preenchida do GRID.
Public Const ERRO_NOMES_TELA_DIFERENTES = 10063 'Parâmetros: sNomeTela1, sNomeTela2
'Nome de Tela %s1 lido no BD não corresponde ao Nome de Tela %s2 do Grid.
Public Const ERRO_TIPO_ROTINA_NAO_INFORMADO = 10064 'Sem parametros
'Tipo de Rotina não foi informado.
Public Const ERRO_AUSENCIA_DADOS_GRID_MENUITENS = 10065 'Sem parametros
'Ausência de dados no Grid de Ítens de Menu.
Public Const ERRO_LEITURA_MENU_ITENS = 10066 'Sem parâmetro
'Erro na leitura da tabela MenuItens.
Public Const ERRO_MENUITEM_ROTINAS_MODULO_NAO_CADASTRADOS = 10067 'Parâmetros: sTipoRotina, sModulo
'Registros de MenuItens associados a rotinas de Usuário do tipo <sTipoRotina> no Módulo <sModulo> não estão cadastrados.
Public Const ERRO_FALTA_MENUITEM_NA_COLECAO = 10068 'parâmetro: sSiglaRotina
'Falta MenuItem correspondente à rotina %s na coleção colMenuItem, preenchida do GRID.
Public Const ERRO_ATUALIZACAO_MENU_ITENS = 10069 'parâmetro: sSiglaRotina
'Erro na atualização da tabela MenuItens, no registro correspondente à rotina %s.
Public Const ERRO_ARQUIVO_NAO_INFORMADO = 10070 'Sem parametros
'Nome de Arquivo não foi informado.
Public Const ERRO_CAMPO_NAO_SELECIONADO = 10071 'Sem parametros
'Pelo menos um campo deve ser selecionado.
Public Const ERRO_AUSENCIA_CAMPO_GRUPOBROWSECAMPO = 10072 'Parametros: sGrupo, sTela, sArquivo
'Ausência de nomes de Campos na tabela GrupoBrowseCampo
'associados a Grupo: <sGrupo>, Tela: <sTela>, Arquivo: <sArquivo>.
Public Const ERRO_AUSENCIA_ARQUIVO_BROWSEARQUIVO = 10073 'Parametro: sTela
'Ausência de nomes de Arquivos associados à tela %s na tabela BrowseArquivo.
Public Const ERRO_AUSENCIA_TELAS_BROWSEARQUIVO = 10074 'parametro: sModulo
'Ausência de nomes de Telas de browse do módulo %s na tabela BrowseArquivo.
Public Const ERRO_AUSENCIA_CAMPO_CAMPOS = 10075 'parametro: sArquivo
'Ausência de Campos do Arquivo %s na tabela Campos.
Public Const ERRO_EXCLUSAO_GRUPOBROWSECAMPO = 10076 'sem parametros
'Erro na exclusão de registros da tabela GrupoBrowseCampo.
Public Const ERRO_INCLUSAO_GRUPOBROWSECAMPO = 10077 'parametros: sGrupo, sTela, sArquivo, vCampo
'Erro na inclusão de registro na tabela GrupoBrowseCampo.
'Dados da inclusão: Grupo=<sGrupo>, Tela=<sTela>, Arquivo=<sArquivo>, Campo=<vCampo>.
Public Const ERRO_EXCLUSAO_BROWSEUSUARIOCAMPO2 = 10078 'parametro: sCodUsuario
'Erro na exclusão dos registros da tabela BrowseUsuarioCampo
'com código de Usuário = %s.
Public Const ERRO_EXCLUSAO_BROWSEUSUARIOORDENACAO = 10079 'parametro: sCodUsuario
'Erro na exclusão dos registros da tabela BrowseUsuarioOrdenacao
'com código de Usuário = %s.
Public Const ERRO_INSERCAO_GRUPOREL = 10080 'parametros cod grupo e cod do rel
'Erro na tentativa de inserir registros na tabela de Grupos x Relatórios. Grupo %s e Relatório %s.
Public Const ERRO_INSERCAO_GRUPOBROWSECAMPO = 10081 'parametros cod grupo e nome da tela
'Erro na tentativa de inserir registros na tabela de Grupos x Campos das Telas de Browse. Grupo %s e Tela %s.
Public Const ERRO_LEITURA_USU_GRUPO = 10082 'parametro cod do grupo
'Erro na leitura de usuário do grupo %s.
Public Const ERRO_LEITURA_USUARIOS_GRUPO = 10083 'parametro cod do grupo
'Erro na leitura dos usuários do grupo %s.
Public Const ERRO_ATUALIZACAO_GRUPO = 10084 'parametro cod do grupo
'Erro na tentativa de atualizar o Grupo %s.
Public Const ERRO_INSERCAO_GRUPO = 10085 'parametro cod do grupo
'Erro na tentativa de inserir o Grupo %s.
Public Const ERRO_INSERCAO_GRUPOROTINA = 10086 'parametros cod do grupo e da rotina
'Erro na tentativa de inserir registros na tabela de Grupos x Rotinas. Grupo %s e Rotina %s.
Public Const ERRO_INSERCAO_GRUPOTELA = 10087 'parametros cod do grupo e da tela
'Erro na tentativa de inserir registros na tabela de Grupos x Telas. Grupo %s e Tela %s.
Public Const ERRO_EXCLUSAO_GRUPO = 10088 'parametro cod do grupo
'Erro na exclusão do Grupo %s.
Public Const ERRO_EXCLUSAO_GRUPOTELA = 10090 'parametro cod do grupo
'Erro na tentativa de excluir registro da tabela de Grupos x Telas. Grupo %s.
Public Const ERRO_LEITURA_USUFILEMP_USU = 10091 'parametro: codusuario
'Erro na leitura de permissões de acesso à filiais/Empresas para o usuário %s.
Public Const ERRO_AUSENCIA_DADOS_GRID = 10092 'sem parametros
'O grid precisa estar preenchido.
Public Const ERRO_RELATORIOS_DO_MODULO_INEXISTENTES = 10093 'sem parametros
'Não existem relatórios associados ao módulo
Public Const ERRO_RELATORIO_NAO_INFORMADO = 10094 'sem parametros
'Selecione um relatório
Public Const ERRO_LEITURA_GRUPO_RELATORIOS = 10095 'sem parametros
'Erro na leitura de registros da tabela GrupoRelatorios
Public Const ERRO_GRUPORELATORIO_REL_NAO_CADASTRADOS = 10096 'parametro = CodRel
'Não há registros na tabela GrupoRelatorios para o relatório %s.
Public Const ERRO_LEITURA_MODULOFILEMP = 10097
'Erro na leitura de registros na tabela ModulosFilEmp
Public Const ERRO_INSERCAO_MODULOFILEMP = 10098
'Erro na inserção de registros na tabela ModulosFilEmp
Public Const ERRO_EXCLUSAO_MODULOFILEMP = 10099
'Erro na exclusão de registros na tabela ModulosFilEmp
Public Const ERRO_FILIALEMPRESA_INATIVA = 10100 'sem parametros
'Não se pode alterar ou excluir uma filial inativa
Public Const ERRO_FILIAL_MESMO_NOME = 10101 'sem parametros
'Não pode haver duas filiais com o mesmo nome.
Public Const ERRO_LIMITE_FILIAISEMPRESA = 10102 'parametro: limite
'Não pode ultrapassar o limite de filiais que é de %s.
Public Const ERRO_LEITURA_FILIAISEMPRESA = 10103 'sem parametro
'Erro na leitura da tabela de FiliaisEmpresa
Public Const ERRO_LEITURA_DICCONFIG = 10104 'sem parametro
'Erro na leitura de configuração do dicionário de dados
Public Const ERRO_EXCLUSAO_USUFILEMP_USU = 10105 'parametro = codusurio
'Erro na exclusão dos direitos de acesso à Empresas e Filiais para o usuário %s.
Public Const ERRO_INSERCAO_USUFILEMP_USU = 10106 'parametro = codusurio
'Erro na inclusão dos direitos de acesso à Empresas e Filiais para o usuário %s.
Public Const ERRO_EXCLUSAO_BROWSEUSUARIOCAMPO3 = 10107 'sem parametros
'Erro na exclusão de registros da tabela BrowseUsuarioCampo
Public Const ERRO_ATUALIZACAO_GRUPO_RELATORIOS = 10108 'parametros: sCodRel, sCodGrupo
'Erro na atualização de registro em GrupoRelatorios para o relatório %s, grupo %s
Public Const ERRO_CODREL_DIFERENTE = 10109 'parametros: codrel no bd, codrel na colecao
'Alterações no banco de dados após a carga da tela impedem a atualização dos dados.
Public Const ERRO_GRUPORELATORIO_NAO_CADASTRADO = 10110 'sem parametro
'Não encontrou registro na tabela GrupoRelatorios
Public Const ERRO_ATUALIZACAO_FILIALEMPRESA = 10111 'parametro: cod da filial
'Erro na atualização de registro da filial %s
Public Const ERRO_INSERCAO_FILIALEMPRESA = 10112 'parametro: cod da filial
'Erro na inserção de registro da filial %s
Public Const ERRO_EXCLUSAO_USUFILEMP = 10114 'sem parametro
'Erro na exclusao de registros da tabela UsuFilEmp
Public Const ERRO_MODULO_VINCULADO_FILIAL = 10115 'Parametros: sSigla, lCodEmpresa, iCodFilial
'Módulo %s não pode ser desativado pois está vinculado a FilialEmpresa, Empresa código %l, Filial código %i.
Public Const ERRO_NOMERED_EMPRESA_NAO_INFORMADO = 10116 'Sem parâmetros
'Nome Reduzido da Empresa não foi informado.
Public Const ERRO_EMPRESA_NOME_JA_EXISTE = 10117 'Parametro: sNome
'Existe Empresa com nome %s no Sistema.
Public Const ERRO_EMPRESA_NOME_RED_JA_EXISTE = 10118 'Parametro: sNomeRed
'Existe Empresa com nome reduzido %s no Sistema.
Public Const ERRO_INSERCAO_MODULO_EMPRESA = 10119 'Parametros: lCodigo, sSigla
'Erro na tentativa de inserir na tabela ModuloEmpresa registro com chave CodEmpresa=%l, SiglaModulo=%s.
Public Const ERRO_INSERCAO_USUFILEMP = 10120 'Sem parametros
'Erro na tentativa de inserir registro na tabela UsuFilEmp.
Public Const ERRO_ALTERACAO_EMPRESA_INATIVA = 10121 'Parametro:lCodigo
'Não é possível alterar Empresa Inativa %l.
Public Const ERRO_ALTERACAO_NOME = 10122 'Parametro: lCodigo
'Não é possível alterar Nome da Empresa %l.
Public Const ERRO_LEITURA_MODULO_EMPRESA = 10123 'Parametro:lCodigo
'Erro na leitura da tabela ModuloEmpresa para Empresa com código %l.
Public Const ERRO_ATUALIZACAO_MODULO_EMPRESA = 10124 'Parametros: lCodigo, sSigla
'Erro na atualização na tabela ModuloEmpresa, Empresa com código %l, Módulo com sigla %s.
Public Const ERRO_LEITURA_FILIAIS_EMPRESAS = 10125 'Parâmetro: lCodEmpresa
'Erro na leitura das Filiais da Empresa %l na tabela de Filiais Empresas.
Public Const ERRO_LOCK_FILIAL_EMPRESA = 10126 'Parâmetros: lCodFilialEmpresa, lCodEmpresa
'Erro na tentativa de fazer "lock" na Filial %l da Empresa %l na tabela de Filiais Empresas.
Public Const ERRO_EXCLUSAO_FILIAL_EMPRESA = 10127 'Parâmetros: lCodFilialEmpresa, lCodEmpresa
'Erro na exclusão da Filial %l da Empresa %l na tabela de Filiais Empresas.
Public Const ERRO_LEITURA_USUFILEMP = 10128 'Parâmetro: lCodEmpresa
'Erro na tentativa de leitura dos usuários vinculados à Empresa %l na tabela UsuFilEmp.
Public Const ERRO_LOCK_USUFILEMP = 10129 'Parâmetro: lEmpresa
'Erro na tentativa de fazer "lock" na tabela UsuFilEmp onde o Código da Empresa = %l.
Public Const ERRO_GRUPORELATORIO_GRUPOMODULO_NAO_CADASTRADOS = 10132 'parametros:sGrupo, sModulo
'Não há relatórios cadastrados para o grupo '%s' módulo '%s'.
Public Const ERRO_AUSENCIA_EMPRESAS = 10133 'Sem parametros
'Não estão cadastradas Empresas no Sistema.
Public Const ERRO_FILIAL_EMPRESA_NAO_CADASTRADA1 = 10134 'Parametros: lCodigo, iCodFilial
'Filial com código %i da Empresa com código %l não está cadastrada no Sistema.
Public Const ERRO_CODIGO_FILIALEMPRESA_NAO_INFORMADO = 10135 'Sem parametros
'Falta informar Código de Filial Empresa.
Public Const ERRO_NOME_FILIALEMPRESA_NAO_INFORMADO = 10136 'Sem parametro
'Falta informar Nome da Filial Empresa.
Public Const ERRO_ESTADO_NAO_INFORMADO_PRINCIPAL = 10138 'Sem parametro
'Falta informar o Estado no Endereço principal.
Public Const ERRO_ESTADO_NAO_INFORMADO_ENTREGA = 10139 'Sem parametro
'Falta informar o Estado no Endereço de entrega.
Public Const ERRO_PAIS_NAO_INFORMADO_PRINCIPAL = 10140 'Sem parametro
'Falta informar o País no Endereço principal.
Public Const ERRO_PAIS_NAO_INFORMADO_ENTREGA = 10141 'Sem parametro
'Falta informar o País no Endereço de entrega.
Public Const ERRO_LEITURA_MODULOFILEMP1 = 10142 'Parametros: sSigla, lCodigo, iFilial
'Erro de leitura na tabela ModuloFilEmp. Chave: SiglaModulo = %s,  CodEmpresa = %l, CodFilial = %i.
Public Const ERRO_LOCK_MODULOFILEMP1 = 10143 'Parametros: sSigla, lCodigo, iFilial
'Erro na de "lock" na tabela ModuloFilEmp. Chave: SiglaModulo = %s,  CodEmpresa = %l, CodFilial = %i.
Public Const ERRO_EXCLUSAO_MODULOFILEMP1 = 10144 'Parametros: sSigla, lCodigo, iFilial
'Erro na exclusão de registro da tabela ModuloFilEmp. Chave: SiglaModulo = %s,  CodEmpresa = %l, CodFilial = %i.
Public Const ERRO_CODGRUPO_NAO_INFORMADO = 10145
'O código do grupo de usuários não foi informado.
Public Const ERRO_NOMERED_USUARIO_NAO_INFORMADO = 10146 'Sem parametro
'Nome Reduzido do Usuário não foi informado.
Public Const ERRO_DATA_FORA_VALIDADE = 10147 'Parâmetro: dtData, dtValidadeDe, dtValidadeAte
'A data %dt não está dentro da validade do sistema que vai de %dt até %dt.
Public Const ERRO_LIMITEEMPRESAS_ATINGIDO = 10148 'Parâmetro: iLimiteEmpresas
'Não é possível criar nova empresa pois o número limite que é de %i já foi atingido.
Public Const ERRO_LIMITEFILIAIS_ATINGIDO = 10149 'Parâmetro: iLimiteFiliais
'Não é possível criar nova filial pois o número limite que é de %i já foi atingido.
Public Const ERRO_LIMITELOGS_ATINGIDO = 10150 'Parâmetro: iLimiteLogs
'Não é possível entrar no sistema pois o número limite  de usuários logados que é de %i já foi atingido.
Public Const ERRO_NUMEMPRESAS_MAIOR_LIMITE = 10151 'Parâmetro: iNumeroEmpresasBD, iLimiteEmpresa
'O número de empresas no BD que é de %i é superior ao limite a ser implantado que é de %i.
'Exclua as empresas excedentes e tente configurar novamente.
Public Const ERRO_NUMFILIAISEMPRESAS_MAIOR_LIMITE = 10152 'Parâmetro: iNumeroFiliaisBD, iLimiteFiliais
'O número de filiais de empresas no BD que é de %i é superior ao limite a ser implantado que é de %i.
'Exclua as filiais excedentes e tente configurar novamente.
Public Const ERRO_ATUALIZACAO_DICCONFIG = 10153
'Erro ao tentar atualizar as informações de Configuração do Sistema.
Public Const ERRO_ATUALIZACAO_MODULOS = 10154
'Erro de atualização na tabela de Modulos.
Public Const ERRO_REGISTRO_DICCONFIG_NAO_ENCONTRADO = 10155
'O Registro de configuração do sistema não foi encontrado.
Public Const ERRO_ODBC_OBTER_HANDLE = 10156 'sem parametros
'Erro na obtenção de recurso para acesso ao banco de dados.
Public Const ERRO_USUARIO_CADASTRADO_COMPRADOR = 10157 'Parametros: sCodUsuario, gsNomeEmpresa, gsNomeFilialEmpresa
'O Usuario sCodUsuario é Comprador da gsNomeEmpresa, Filial gsNomeFilialEmpresa, por isso não é possivel Desabilita - lo nesta Filial.
Public Const ERRO_LEITURA_TABELA_COMPRADOR = 10158 'Parametros: gsNomeEmpresa, gsNomeFilialEmpresa
'Erro na Leitura da Tabela de Comprador da Empresa gsNomeEmpresa, Filial gsNomeFilialEmpresa.
Public Const ERRO_LEITURA_VERSAO = 10159 'Sem parametro
'Erro de leitura na tabela de Versão.
Public Const ERRO_VERSAO_VIGENTE_AUSENTE = 10160 'Sem parametro
'Ausência de versão vigente na tabela de Versão.
Public Const ERRO_TRECHO_SENHA_INCOMPLETO = 10162 'Parâmetro: iTrecho
'O trecho %i da senha está incompleto.
Public Const ERRO_NUMERO_SERIE_DIFERENTE_BD = 10163
'O número de série informado é diferente do que consta no Banco de Dados.
Public Const ERRO_DATA_SENHA_BD_MAIOR = 10164
'A data atual é inferior a da ultima configuração.







'Códigos de Aviso RESERVADO de 10500 a 10999
Public Const EXCLUSAO_EMPRESA = 10500 'Parametros: lCodigoEmpresa , sNomeEmpresa
'A Empresa com código: %l e nome: %s será excluída. Atenção: TODOS as suas tabelas serão excluídas.
'Confirma a exclusão?
Public Const AVISO_GRUPO_INEXISTENTE = 10501 'Parametro: sCodigoGrupo
'O Grupo com código = %s não está cadastrado. Deseja cadastrá-lo?
Public Const EXCLUSAO_USUARIO = 10502 'Parametros: sCodUsuario, sNome
'Confirma a exclusão de Usuário com código=%s e nome=%s ?
Public Const AVISO_EXCLUSAO_GRUPO = 10503 'parametro cod do grupo
'Confirma a exclusão do Grupo ?
Public Const AVISO_CRIAR_EMPRESA = 10504 'Parametro sCodEmpresa
'Empresa com código %s não está cadastrada no Sistema. Deseja criar?
Public Const AVISO_EXCLUSAO_FILIALEMPRESA = 10505 'Parametros: iCodFilial, lCodEmpresa
'Confirma exclusão de Filial com código %i da Empresa com código %l ?
Public Const AVISO_EXCLUSAO_EMPRESA = 10506 'Parametros: lCodigo, sNome
'Confirma exclusão de Empresa com código %l e nome %s ?
Public Const AVISO_EXCLUSAO_USUARIOS = 10507 'Parâmetros: sCodGrupo
'O Grupo de código %s possui Usuários cadastrados. Confirma exclusão do Grupo?
Public Const AVISO_EXCLUSAO_FILIAIS = 10508 'Parâmetros: lCodigoEmpresa, sNomeEmpresa
'A empresa %s de código %l possui filiais cadastradas. Confirma exclusão?


