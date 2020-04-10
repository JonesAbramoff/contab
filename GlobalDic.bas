Attribute VB_Name = "GlobalDic"
Option Explicit

'Variáveis independentes de instâncias
'Suporte a ClassDicDados
Public DicGlob_sRotina As String
Public DicGlob_sGrupo As String
Public DicGlob_sTela As String

'Local para criacao das Global Const deste modulo
Public Const USU_SUP = "Supervisor"

''Data Default para Validade de Usuários
'Public Const DATA_DEFAULT As Date = #12/31/28#

'Tempo inicial (em dias) antes de ser dada senha
Public Const TEMPO_SENHA_INICIAL = 60

'Log Atividade
Public Const LOG_SIM = 1
Public Const LOG_NAO = 0

'Usuario Ativo ou Nao
Public Const ATIVIDADE = 1
Public Const INATIVIDADE = 0

'Número de rotinas de usuário (com substring "_USU_") para cada módulo e tipo de rotina
Public Const USU_ROTINAS = 6

'Versao vigente ou nao
Public Const VERSAO_VIGENTE = 1
Public Const VERSAO_NAO_VIGENTE = 0

'=========================
'Tamanhos de Strings no BD
'=========================

'Empresa
Public Const STRING_STRING_CONEXAO = 256
Public Const STRING_EMPRESA_NOME = 50
Public Const STRING_EMPRESA_CODIGO = 2
Public Const STRING_EMPRESA_NOME_RED = 50
Public Const STRING_EMPRESA_DBOLAP = 50

'Grupo
Public Const STRING_GRUPO_CODIGO = 10
Public Const STRING_GRUPO_DESCRICAO = 50

'Usuario
'Public Const STRING_USUARIO_CODIGO = 10
'Public Const STRING_USUARIO_NOME = 50
'Public Const STRING_USUARIO_NOMEREDUZIDO = 20
'Public Const STRING_USUARIO_SENHA = 10
'Public Const STRING_USUARIO_STRING_CONEXAO = 255

''Módulo
'Public Const STRING_MODULO_NOME = 50

'Incluído por Luiz Nogueira em 27/10/03
'OpcoesTelas
Public Const STRING_OPCOESTELAS_OPCAO = 50

'Rotina
Public Const STRING_ROTINA_SIGLA = 50
Public Const STRING_ROTINA_DESCRICAO = 50
Public Const STRING_ROTINA_PROJ_ORIG = 50
Public Const STRING_ROTINA_CLASS_ORIG = 50
Public Const STRING_ROTINA_PROJ_CUST = 50
Public Const STRING_ROTINA_CLASS_CUST = 50

'Tela
Public Const STRING_TELA_NOME = 50
Public Const STRING_TELA_DESCRICAO = 100
Public Const STRING_TELA_PROJ_ORIG = 50
Public Const STRING_TELA_CLASS_ORIG = 50
Public Const STRING_TELA_PROJ_CUST = 50
Public Const STRING_TELA_CLASS_CUST = 50

'GrupoRotinas
Public Const STRING_GRUPO_ROTINA_SIGLAROTINA = 50
Public Const STRING_GRUPO_ROTINA_CODGRUPO = 10
Public Const STRING_GRUPO_ROTINA_PROJETO = 50
Public Const STRING_GRUPO_ROTINA_CLASSE = 50

'GrupoTela
Public Const STRING_GRUPO_TELA_NOMETELA = 50
Public Const STRING_GRUPO_TELA_CODGRUPO = 10
Public Const STRING_GRUPO_TELA_PROJETO = 50
Public Const STRING_GRUPO_TELA_CLASSE = 50

'MenuItem
Public Const STRING_MENU_ITEM_TITULO = 50
Public Const STRING_MENU_ITEM_SIGLAROTINA = 50
Public Const STRING_MENU_ITEM_NOMETELA = 50

'GrupoBrowseCampo (GBC)
Public Const STRING_GBC_NOME = 50

'BrowseArquivo
Public Const STRING_BROWSEARQUIVO_NOME_TELA = 50
Public Const STRING_BROWSEARQUIVO_NOME_ARQ = 50

'Campos
Public Const STRING_CAMPOS_NOME = 50

'Versao
Public Const STRING_VERSAO_CODIGO = 50

'=============================
'User defined types
'=============================

Type typeDicEmpresa
    lCodigo As Long
    sNome As String
    sNomeReduzido As String
    sStringConexao As String
    iInativa As Integer
End Type

Type typeDicUsuario
    sCodUsuario As String
    sCodGrupo As String
    sNome As String
    sNomeReduzido As String
    sSenha As String
    dtDataValidade As Date
    iAtivo As Integer
    sNomeLogin As String
    sComputador As String
    iLogado As Integer
    sEmail As String
End Type

Type typeRotina
    sSigla As String
    sDescricao As String
    sProjeto_Original As String
    sClasse_Original As String
    sProjeto_Customizado As String
    sClasse_Customizada As String
    iLogAtividade As Integer
End Type

Type typeTela
    sNome As String
    sDescricao As String
    sProjeto_Original As String
    sClasse_Original As String
    sProjeto_Customizado As String
    sClasse_Customizada As String
End Type

Type typeGrupoRotina
    sCodGrupo As String
    sSiglaRotina As String
    iTipoDeAcesso As Integer
    iLogAtividade As Integer
    sProjeto As String
    sClasse As String
End Type

Type typeGrupoTela
    sCodGrupo As String
    sNomeTela As String
    iTipoDeAcesso As Integer
    sProjeto As String
    sClasse As String
End Type

Type typeFilialEmpresa
    lCodEmpresa As Long
    iCodFilial As Integer
    sNomeFilial As String
    sNomeEmpresa As String
    sNomeRedEmpresa As String
End Type

Type typeVersao
    sCodigo As String
    dtData As Date
    iVigente As Integer
End Type

'Incluído por Luiz Nogueira em 27/10/03
Type typeOpcoesTelas
    lCodigo As Long
    sOpcao As String
    sNomeTela As String
    iPadrao As Integer
End Type

'Incluído por Luiz Nogueira em 27/10/03
Type typeOpcoesTelasValores
    lCodOpca As Long
    sNomeControle As String
    sValorCampo As String
End Type

