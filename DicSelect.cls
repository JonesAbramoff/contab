VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DicSelect"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Versao vigente
Private mvarobjVersao As Object

Public Property Set objVersao(ByVal vData As Object)
    Set mvarobjVersao = vData
End Property

Public Property Get objVersao() As Object
    Set objVersao = mvarobjVersao
End Property

Function Empresa_Le(objEmpresa As ClassDicEmpresa) As Long
'Le Empresa ATIVA a partir da chave fornecida em objEmpresa.
'Retorna os dados em objEmpresa.

Dim lComando As Long
Dim lErro As Long
Dim tEmpresa As typeDicEmpresa
    
On Error GoTo Erro_Empresa_Le

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6177
    
    tEmpresa.sNome = String(STRING_EMPRESA_NOME, 0)
    tEmpresa.sNomeReduzido = String(STRING_EMPRESA_NOME_RED, 0)
    tEmpresa.sStringConexao = String(STRING_STRING_CONEXAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Nome, NomeReduzido, StringConexao FROM Empresas WHERE Codigo = ? And Inativa=0", tEmpresa.sNome, tEmpresa.sNomeReduzido, tEmpresa.sStringConexao, objEmpresa.lCodigo)
    If lErro Then Error 6178
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6179
    If lErro <> AD_SQL_SUCESSO Then Error 6180
   
    objEmpresa.sNome = tEmpresa.sNome
    objEmpresa.sNomeReduzido = tEmpresa.sNomeReduzido
    objEmpresa.sStringConexao = tEmpresa.sStringConexao
    
    lErro = Comando_Fechar(lComando)
    
    Empresa_Le = SUCESSO

    Exit Function

Erro_Empresa_Le:

    Empresa_Le = Err

    Select Case Err
    
        Case 6177
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 6178, 6180
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA", Err, objEmpresa.lCodigo)
        
        Case 6179  'Será tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159065)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Empresa_Le_Nome(objEmpresa As ClassDicEmpresa) As Long
'Lê o codigo da Empresa a partir do nome fornecido em objEmpresa
'Retorna os dados em objEmpresa

Dim lComando As Long
Dim lErro As Long
Dim lCodigo As Long
    
On Error GoTo Erro_Empresa_Le_Nome

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6250
    
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Empresas WHERE Nome = ?", lCodigo, objEmpresa.sNome)
    If lErro <> AD_SQL_SUCESSO Then Error 6251
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6290
    
    If lErro = AD_SQL_SEM_DADOS Then Error 6289
   
    objEmpresa.lCodigo = lCodigo
    
    lErro = Comando_Fechar(lComando)
    
    Empresa_Le_Nome = SUCESSO

    Exit Function

Erro_Empresa_Le_Nome:

    Empresa_Le_Nome = Err

    Select Case Err
    
        Case 6250
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 6251, 6290
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA2", Err, objEmpresa.sNome)
        
        Case 6289  'Será tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159066)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Grupo_Le(objGrupo As ClassDicGrupo) As Long
'Le o Grupo a partir de sua chave e retorna os atributos em objGrupo

Dim lComando As Long
Dim lErro As Long, iLogAtividade As Integer
Dim sDescricao As String
Dim dtDataValidade As Date
    
On Error GoTo Erro_Grupo_Le

    sDescricao = String(STRING_GRUPO_DESCRICAO, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6228
    
    lErro = Comando_Executar(lComando, "SELECT Descricao, DataValidade, LogAtividade FROM GruposDeUsuarios WHERE CodGrupo = ?", sDescricao, dtDataValidade, iLogAtividade, objGrupo.sCodGrupo)
    If lErro Then Error 6229
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6230
    If lErro <> AD_SQL_SUCESSO Then Error 6231
   
    'Preenche objGrupo
    objGrupo.sDescricao = sDescricao
    objGrupo.dtDataValidade = dtDataValidade
    objGrupo.iLogAtividade = iLogAtividade
    
    lErro = Comando_Fechar(lComando)
    
    Grupo_Le = SUCESSO

    Exit Function

Erro_Grupo_Le:

    Grupo_Le = Err

    Select Case Err
    
        Case 6228
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 6229, 6231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO1", Err, objGrupo.sCodGrupo)
        
        Case 6230  'Será tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159067)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function DicUsuario_Le(objUsuario As ClassDicUsuario) As Long
'Le o Usuario a partir de sua chave
'Retorna o Usuario em objUsuario

Dim lComando As Long
Dim lErro As Long
Dim tUsuario As typeDicUsuario
    
On Error GoTo Erro_DicUsuario_Le

    tUsuario.sNome = String(STRING_USUARIO_NOME, 0)
    tUsuario.sSenha = String(STRING_USUARIO_SENHA, 0)
    tUsuario.sCodGrupo = String(STRING_GRUPO_CODIGO, 0)
    tUsuario.sNomeReduzido = String(STRING_USUARIO_NOMEREDUZIDO, 0)
    tUsuario.sEmail = String(64, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6221
    
    lErro = Comando_Executar(lComando, "SELECT CodGrupo, Nome, NomeReduzido, Senha, DataValidade, Ativo, Email FROM Usuarios WHERE Usuarios.CodUsuario = ?", tUsuario.sCodGrupo, tUsuario.sNome, tUsuario.sNomeReduzido, tUsuario.sSenha, tUsuario.dtDataValidade, tUsuario.iAtivo, tUsuario.sEmail, objUsuario.sCodUsuario)
    If lErro Then Error 6222
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6223
    If lErro <> AD_SQL_SUCESSO Then Error 6224
   
    'Preenche objUsuario
    objUsuario.sCodGrupo = tUsuario.sCodGrupo
    objUsuario.sNome = tUsuario.sNome
    objUsuario.sNomeReduzido = tUsuario.sNomeReduzido
    objUsuario.sSenha = tUsuario.sSenha
    objUsuario.dtDataValidade = tUsuario.dtDataValidade
    objUsuario.iAtivo = tUsuario.iAtivo
    objUsuario.sEmail = tUsuario.sEmail
    
    lErro = Comando_Fechar(lComando)
    
    DicUsuario_Le = SUCESSO

    Exit Function

Erro_DicUsuario_Le:

    DicUsuario_Le = Err

    Select Case Err
    
        Case 6221
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 6222, 6224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIO", Err, objUsuario.sCodUsuario)
        
        Case 6223  'Será tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159068)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Grupos_Le(colGrupo As Collection) As Long
'Preenche a coleção colGrupo com os códigos de grupos encontrados no BD

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sCodGrupo As String

On Error GoTo Erro_Grupos_Le

    sCodGrupo = String(STRING_GRUPO_CODIGO, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6225

    'seleciona no BD todos os códigos de grupos
    lErro = Comando_Executar(lComando, "SELECT CodGrupo FROM GruposDeUsuarios ORDER BY CodGrupo", sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then Error 6226

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If iFim_de_Arquivo = AD_SQL_SEM_DADOS Then Error 6365  'Não existem Grupos cadastrados
    
    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colGrupo.Add (sCodGrupo)
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6227

    lErro = Comando_Fechar(lComando)

    Grupos_Le = SUCESSO

    Exit Function

Erro_Grupos_Le:

    Grupos_Le = Err

    Select Case Err

        Case 6225
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6226, 6227
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO", Err)
            
        Case 6365  'Será tratado na rotina chamadora
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159069)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Usuarios_Le_Grupo(ByVal sCodGrupo As String, colUsuario As Collection) As Long
'Preenche a coleção colUsuario com os códigos de usuários
'do Grupo sCodGrupo encontrados no BD

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sCodUsuario As String

On Error GoTo Erro_Usuarios_Le_Grupo

    sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6218

    'seleciona no BD todos os códigos de Usuarios pertencentes ao Grupo sCodGrupo
    lErro = Comando_Executar(lComando, "SELECT CodUsuario FROM Usuarios WHERE CodGrupo=? ORDER BY CodUsuario", sCodUsuario, sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then Error 6219

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    
    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colUsuario.Add (sCodUsuario)
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6220

    lErro = Comando_Fechar(lComando)

    Usuarios_Le_Grupo = SUCESSO

    Exit Function

Erro_Usuarios_Le_Grupo:

    Usuarios_Le_Grupo = Err

    Select Case Err

        Case 6218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6219, 6220
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOS_GRUPO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159070)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'''Lê as Empresas e Grupos associados ao Usuário com código sCodUsuario
'''Usuário está cadastrado
'''Retorna na coleção colEmpresasGrupos
''Function Empresas_Grupos_Le(ByVal sCodUsuario As String, colEmpresaGrupo As Collection) As Long
''
''Dim lConexao As Long
''Dim lComando As Long
''Dim lErro As Long
''Dim iFim_de_Arquivo As Integer
''Dim lCodEmpresa As Long
''Dim sCodGrupo As String
''Dim objEmpresaGrupo As ClassDicEmpresaGrupo
''
''On Error GoTo Erro_Empresas_Grupos_Le
''
''    sCodGrupo = string(STRING_GRUPO_CODIGO)
''
''    lConexao = GL_lConexaoDic
''
''    lComando = Comando_AbrirExt(lConexao)
''    If lComando = 0 Then Error 6269
''
''    'seleciona no BD todos os códigos de Empresas e Grupos associados ao Usuário
''    lErro = Comando_Executar(lComando, "SELECT CodEmpresa, CodGrupo FROM UsuEmpGrupo WHERE CodUsuario=?", lCodEmpresa, sCodGrupo, sCodUsuario)
''    If lErro <> AD_SQL_SUCESSO Then Error 6270
''
''    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
''    If iFim_de_Arquivo = AD_SQL_SEM_DADOS Then Error 6367
''
''    'Preenche a coleção
''    Do While iFim_de_Arquivo = AD_SQL_SUCESSO
''
''        'Cria objeto do tipo ClassDicEmpresaGrupo
''        Set objEmpresaGrupo = New ClassDicEmpresaGrupo
''
''        'Preenche o objeto
''        objEmpresaGrupo.lCodEmpresa = lCodEmpresa
''        objEmpresaGrupo.sCodGrupo = sCodGrupo
''
''        'Adiciona objeto à Coleção
''        colEmpresaGrupo.Add objEmpresaGrupo
''        iFim_de_Arquivo = Comando_BuscarProximo(lComando)
''
''    Loop
''
''    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6271
''
''    lErro = Comando_Fechar(lComando)
''
''    Empresas_Grupos_Le = SUCESSO
''
''    Exit Function
''
''Erro_Empresas_Grupos_Le:
''
''    Empresas_Grupos_Le = Err
''
''    Select Case Err
''
''        Case 6269
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 6270, 6271
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUEMPGRUPO", Err)
''            lErro = Comando_Fechar(lComando)
''
''        Case 6367
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_SEM_GRUPO_EMPRESA", Err, sCodUsuario)
''            lErro = Comando_Fechar(lComando)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159071)
''
''    End Select
''
''    Exit Function
''
''End Function

Public Function Modulos_Le(colModulo As Collection) As Long
'Preenche a coleção colModulo com os nomes de módulos LIBERADOS encontrados no BD

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sNome As String
Dim objVersao As New ClassVersao

On Error GoTo Erro_Modulos_Le

    sNome = String(STRING_MODULO_NOME, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6297

    'Lê a versão vigente
    lErro = VersaoVigente_Le(objVersao)
    If lErro <> SUCESSO Then Error 25808
    
    'seleciona no BD todos os nomes de módulos até a versão vigente
    lErro = Comando_Executar(lComando, "SELECT Nome FROM Modulos, Versao WHERE Modulos.Versao = Versao.Codigo AND Versao.Data <= ? AND Modulos.Liberado = ? ORDER BY Nome", sNome, objVersao.dtData, MODULO_LIBERADO)
    If lErro <> AD_SQL_SUCESSO Then Error 6298

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If iFim_de_Arquivo = AD_SQL_SEM_DADOS Then Error 6305
    
    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colModulo.Add (sNome)
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6299

    lErro = Comando_Fechar(lComando)

    Modulos_Le = SUCESSO

    Exit Function

Erro_Modulos_Le:

    Modulos_Le = Err

    Select Case Err

        Case 6297
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6298, 6299
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULO", Err)
            
        Case 6305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODULO_INEXISTENTE", Err)
            
        Case 25808 'tratado na rotina chamada
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159072)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Modulo_Le_Rotina(ByVal sRotina As String, sModulo As String) As Long
'Lê nome de Módulo que contém sRotina (primeiro na ordem alfabética)

Dim lComando As Long
Dim lErro As Long
Dim objVersao As New ClassVersao

On Error GoTo Erro_Modulo_Le_Rotina

    sModulo = String(STRING_MODULO_NOME, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6306

    'Lê a versão vigente
    lErro = VersaoVigente_Le(objVersao)
    If lErro <> SUCESSO Then Error 25811
    
    'pesquisa no BD módulos que contém sRotina
    lErro = Comando_Executar(lComando, "SELECT Nome FROM Modulos, RotinasModulo, Rotinas, Versao WHERE Modulos.Versao = Versao.Codigo AND Versao.Data <= ? AND Modulos.Sigla=RotinasModulo.SiglaModulo AND RotinasModulo.SiglaRotina=Rotinas.Sigla AND Rotinas.Sigla=? ORDER BY Nome", sModulo, objVersao.dtData, sRotina)
    If lErro <> AD_SQL_SUCESSO Then Error 6307

    'Lê o primeiro desses módulos
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6308
    If lErro <> AD_SQL_SUCESSO Then Error 6309

    lErro = Comando_Fechar(lComando)

    Modulo_Le_Rotina = SUCESSO

    Exit Function

Erro_Modulo_Le_Rotina:

    Modulo_Le_Rotina = Err

    Select Case Err

        Case 6306
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6307, 6309
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULO", Err)
            
        Case 6308  'Será tratado na rotina chamadora

        Case 25811 'tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159073)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Rotinas_Le_NomeModulo(ByVal sModulo As String, colRotina As Collection) As Long
'Lê siglas de Rotinas contidas no módulo com NOME sModulo e devolve em colRotina

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sRotina As String  'sigla da Rotina

On Error GoTo Erro_Rotinas_Le_NomeModulo

    sRotina = String(STRING_ROTINA_SIGLA, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6311

    'seleciona no BD todos as siglas de Rotinas pertencentes ao sModulo
    lErro = Comando_Executar(lComando, "SELECT Rotinas.Sigla FROM Rotinas, RotinasModulo, Modulos WHERE Rotinas.Sigla=RotinasModulo.SiglaRotina AND RotinasModulo.SiglaModulo=Modulos.Sigla AND Modulos.Nome=? ORDER BY Rotinas.Sigla", sRotina, sModulo)
    If lErro <> AD_SQL_SUCESSO Then Error 6312

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6313
    
    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colRotina.Add (sRotina)
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6314

    lErro = Comando_Fechar(lComando)

    Rotinas_Le_NomeModulo = SUCESSO

    Exit Function

Erro_Rotinas_Le_NomeModulo:

    Rotinas_Le_NomeModulo = Err

    Select Case Err

        Case 6311
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6312, 6314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTINA", Err)
            
        Case 6313
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ROTINAS_DO_MODULO_INEXISTENTES", Err, sModulo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159074)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function Rotina_Le(objRotina As ClassDicRotina) As Long
'Lê rotina a partir de sua chave, retorna dados em objRotina

Dim lComando As Long
Dim lErro As Long
Dim tRotina As typeRotina
    
On Error GoTo Erro_Rotina_Le

    tRotina.sDescricao = String(STRING_ROTINA_DESCRICAO, 0)
    tRotina.sProjeto_Original = String(STRING_ROTINA_PROJ_ORIG, 0)
    tRotina.sClasse_Original = String(STRING_ROTINA_CLASS_ORIG, 0)
    tRotina.sProjeto_Customizado = String(STRING_ROTINA_PROJ_CUST, 0)
    tRotina.sClasse_Customizada = String(STRING_ROTINA_CLASS_CUST, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6315
    
    lErro = Comando_Executar(lComando, "SELECT Descricao, LogAtividade, projeto_original, classe_original, projeto_customizado, classe_customizado FROM Rotinas WHERE Sigla=?", tRotina.sDescricao, tRotina.iLogAtividade, tRotina.sProjeto_Original, tRotina.sClasse_Original, tRotina.sProjeto_Customizado, tRotina.sClasse_Customizada, objRotina.sSigla)
    If lErro Then Error 6316
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6317
    If lErro <> AD_SQL_SUCESSO Then Error 6318
   
    objRotina.sDescricao = tRotina.sDescricao
    objRotina.sProjeto_Original = tRotina.sProjeto_Original
    objRotina.sClasse_Original = tRotina.sClasse_Original
    objRotina.sProjeto_Customizado = tRotina.sProjeto_Customizado
    objRotina.sClasse_Customizada = tRotina.sClasse_Customizada
    objRotina.iLogAtividade = tRotina.iLogAtividade
    
    lErro = Comando_Fechar(lComando)
    
    Rotina_Le = SUCESSO

    Exit Function

Erro_Rotina_Le:

    Rotina_Le = Err

    Select Case Err
    
        Case 6315
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 6316, 6318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTINA1", Err, objRotina.sSigla)
        
        Case 6317  'Será tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159075)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Modulo_Le_Tela(ByVal sTela As String, sModulo As String) As Long
'Lê nome de Módulo que contém sTela (primeiro na ordem alfabética)

Dim lComando As Long
Dim lErro As Long
Dim objVersao As New ClassVersao

On Error GoTo Erro_Modulo_Le_Tela

    sModulo = String(STRING_MODULO_NOME, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6340

    'Lê a versão vigente
    lErro = VersaoVigente_Le(objVersao)
    If lErro <> SUCESSO Then Error 25810
    
    'pesquisa no BD módulos que contém sTela
    lErro = Comando_Executar(lComando, "SELECT Modulos.Nome FROM Modulos, TelasModulo, Telas, Versao WHERE Modulos.Versao = Versao.Codigo AND Versao.Data <= ? AND Modulos.Sigla=TelasModulo.SiglaModulo AND TelasModulo.NomeTela=Telas.Nome AND Telas.Nome=? ORDER BY Modulos.Nome", sModulo, objVersao.dtData, sTela)
    If lErro <> AD_SQL_SUCESSO Then Error 6341

    'Lê o primeiro desses módulos
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6342 'Não há módulo que contém a tela
    If lErro <> AD_SQL_SUCESSO Then Error 6343

    lErro = Comando_Fechar(lComando)

    Modulo_Le_Tela = SUCESSO

    Exit Function

Erro_Modulo_Le_Tela:

    Modulo_Le_Tela = Err

    Select Case Err

        Case 6340
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6341, 6343
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULO", Err)
            
        Case 6342  'Será tratado na rotina chamadora

        Case 25810 'tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159076)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function Telas_Le_NomeModulo(ByVal sModulo As String, colTela As Collection) As Long
'Lê nomes de Telas contidas no módulo com NOME sModulo e devolve em colTela

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sTela As String  'nome da Tela

On Error GoTo Erro_Telas_Le_NomeModulo

    sTela = String(STRING_TELA_NOME, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6344

    'seleciona no BD todos as siglas de Telas pertencentes ao sModulo
    lErro = Comando_Executar(lComando, "SELECT Telas.Nome FROM Telas, TelasModulo, Modulos WHERE Telas.Nome=TelasModulo.NomeTela AND TelasModulo.SiglaModulo=Modulos.Sigla AND Modulos.Nome=? ORDER BY Telas.Nome", sTela, sModulo)
    If lErro <> AD_SQL_SUCESSO Then Error 6345

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6346
    
    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colTela.Add (sTela)
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6347

    lErro = Comando_Fechar(lComando)

    Telas_Le_NomeModulo = SUCESSO

    Exit Function

Erro_Telas_Le_NomeModulo:

    Telas_Le_NomeModulo = Err

    Select Case Err

        Case 6344
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6345, 6347
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELA", Err)
            
        Case 6346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELAS_DO_MODULO_INEXISTENTES", Err, sModulo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159077)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Tela_Le(objTela As ClassDicTela) As Long
'Lê tela a partir de sua chave, retorna dados em objTela

Dim lComando As Long
Dim lErro As Long
Dim tTela As typeTela
    
On Error GoTo Erro_Tela_Le

    tTela.sDescricao = String(STRING_TELA_DESCRICAO, 0)
    tTela.sProjeto_Original = String(STRING_TELA_PROJ_ORIG, 0)
    tTela.sClasse_Original = String(STRING_TELA_CLASS_ORIG, 0)
    tTela.sProjeto_Customizado = String(STRING_TELA_PROJ_CUST, 0)
    tTela.sClasse_Customizada = String(STRING_TELA_CLASS_CUST, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6348
    
    lErro = Comando_Executar(lComando, "SELECT Descricao, projeto_original, classe_original, projeto_customizado, classe_customizado FROM Telas WHERE Nome=?", tTela.sDescricao, tTela.sProjeto_Original, tTela.sClasse_Original, tTela.sProjeto_Customizado, tTela.sClasse_Customizada, objTela.sNome)
    If lErro Then Error 6349
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6350
    If lErro <> AD_SQL_SUCESSO Then Error 6351
   
    objTela.sDescricao = tTela.sDescricao
    objTela.sProjeto_Original = tTela.sProjeto_Original
    objTela.sClasse_Original = tTela.sClasse_Original
    objTela.sProjeto_Customizado = tTela.sProjeto_Customizado
    objTela.sClasse_Customizada = tTela.sClasse_Customizada
    
    lErro = Comando_Fechar(lComando)
    
    Tela_Le = SUCESSO

    Exit Function

Erro_Tela_Le:

    Tela_Le = Err

    Select Case Err
    
        Case 6348
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 6349, 6351
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELA1", Err, objTela.sNome)
        
        Case 6350  'Será tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159078)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function GrupoRotina_Le_Rotina(ByVal sSiglaRotina As String, colGrupoRotina As colGrupoRotina) As Long
'Lê dados de GrupoRotinas correspondentes a Rotina com sigla sSiglaRotina
'Retorna dados em colGrupoRotina

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim tGrupoRotina As typeGrupoRotina

On Error GoTo Erro_GrupoRotina_Le_Rotina
    
    tGrupoRotina.sCodGrupo = String(STRING_GRUPO_ROTINA_CODGRUPO, 0)
    tGrupoRotina.sProjeto = String(STRING_GRUPO_ROTINA_PROJETO, 0)
    tGrupoRotina.sClasse = String(STRING_GRUPO_ROTINA_CLASSE, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6370

    'seleciona no BD todos os GrupoRotina correspondentes a sSiglaRotina
    lErro = Comando_Executar(lComando, "SELECT CodGrupo, TipoDeAcesso, projeto, classe FROM GrupoRotinas WHERE SiglaRotina=? ORDER BY CodGrupo", tGrupoRotina.sCodGrupo, tGrupoRotina.iTipoDeAcesso, tGrupoRotina.sProjeto, tGrupoRotina.sClasse, sSiglaRotina)
    If lErro <> AD_SQL_SUCESSO Then Error 6371

    'Lê o primeiro registro selecionado
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If iFim_de_Arquivo = AD_SQL_SEM_DADOS Then Error 6372
    
    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colGrupoRotina.Add LOG_NAO, tGrupoRotina.sProjeto, tGrupoRotina.sClasse, tGrupoRotina.sCodGrupo, sSiglaRotina, tGrupoRotina.iTipoDeAcesso
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6373

    lErro = Comando_Fechar(lComando)

    GrupoRotina_Le_Rotina = SUCESSO

    Exit Function

Erro_GrupoRotina_Le_Rotina:

    GrupoRotina_Le_Rotina = Err

    Select Case Err

        Case 6370
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6371, 6373
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_ROTINAS", Err)
            
        Case 6372
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOROTINA_ROTINA_NAO_CADASTRADOS", Err, sSiglaRotina)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159079)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function GrupoTela_Le_Tela(ByVal sTela As String, colGrupoTela As colGrupoTela) As Long
'Lê dados de GrupoTela correspondentes a Tela com sigla sTela
'Retorna dados em colGrupoTela

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim tGrupoTela As typeGrupoTela

On Error GoTo Erro_GrupoTela_Le_Tela
    
    tGrupoTela.sCodGrupo = String(STRING_GRUPO_TELA_CODGRUPO, 0)
    tGrupoTela.sProjeto = String(STRING_GRUPO_TELA_PROJETO, 0)
    tGrupoTela.sClasse = String(STRING_GRUPO_TELA_CLASSE, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6435

    'seleciona no BD todos os GrupoTela correspondentes a sTela
    lErro = Comando_Executar(lComando, "SELECT CodGrupo, TipoDeAcesso, projeto, classe FROM GrupoTela WHERE NomeTela=? ORDER BY CodGrupo", tGrupoTela.sCodGrupo, tGrupoTela.iTipoDeAcesso, tGrupoTela.sProjeto, tGrupoTela.sClasse, sTela)
    If lErro <> AD_SQL_SUCESSO Then Error 6436

    'Lê o primeiro registro selecionado
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If iFim_de_Arquivo = AD_SQL_SEM_DADOS Then Error 6437
    
    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colGrupoTela.Add tGrupoTela.sProjeto, tGrupoTela.sClasse, tGrupoTela.sCodGrupo, tGrupoTela.iTipoDeAcesso, sTela
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6438

    lErro = Comando_Fechar(lComando)

    GrupoTela_Le_Tela = SUCESSO

    Exit Function

Erro_GrupoTela_Le_Tela:

    GrupoTela_Le_Tela = Err

    Select Case Err

        Case 6435
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6436, 6438
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_TELA", Err)
            
        Case 6437
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOTELA_TELA_NAO_CADASTRADOS", Err, sTela)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159080)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function GrupoRotina_Le_GrupoModulo(ByVal sGrupo As String, ByVal sModulo As String, colGrupoRotina As colGrupoRotina) As Long
'Lê dados de GrupoRotinas correspondentes a sGrupo e a rotinas contidas em sModulo
'Retorna dados em colGrupoRotina

Dim lComando As Long
Dim lErro As Long
Dim tGrupoRotina As typeGrupoRotina

On Error GoTo Erro_GrupoRotina_Le_GrupoModulo

    tGrupoRotina.sSiglaRotina = String(STRING_ROTINA_SIGLA, 0)
    tGrupoRotina.sProjeto = String(STRING_GRUPO_ROTINA_PROJETO, 0)
    tGrupoRotina.sClasse = String(STRING_GRUPO_ROTINA_CLASSE, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6458

    'seleciona no BD todos os GrupoRotina correspondentes a sGrupo e rotinas no módulo sModulo
    lErro = Comando_Executar(lComando, "SELECT GrupoRotinas.SiglaRotina, TipoDeAcesso, projeto, classe FROM GrupoRotinas, RotinasModulo, Modulos WHERE CodGrupo=? AND GrupoRotinas.SiglaRotina=RotinasModulo.SiglaRotina AND RotinasModulo.SiglaModulo=Modulos.Sigla AND Modulos.Nome=? ORDER BY GrupoRotinas.SiglaRotina", tGrupoRotina.sSiglaRotina, tGrupoRotina.iTipoDeAcesso, tGrupoRotina.sProjeto, tGrupoRotina.sClasse, sGrupo, sModulo)
    If lErro <> AD_SQL_SUCESSO Then Error 6459

    'Lê o primeiro registro selecionado
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 6461
    If lErro = AD_SQL_SEM_DADOS Then Error 6460

    'Preenche a coleção
    Do While lErro = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colGrupoRotina.Add LOG_NAO, tGrupoRotina.sProjeto, tGrupoRotina.sClasse, sGrupo, tGrupoRotina.sSiglaRotina, tGrupoRotina.iTipoDeAcesso
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6461

    Loop

    lErro = Comando_Fechar(lComando)

    GrupoRotina_Le_GrupoModulo = SUCESSO

    Exit Function

Erro_GrupoRotina_Le_GrupoModulo:

    GrupoRotina_Le_GrupoModulo = Err

    Select Case Err

        Case 6458
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6459, 6461
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_ROTINAS", Err)

        Case 6460
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOROTINA_GRUPOMODULO_NAO_CADASTRADOS", Err, sGrupo, sModulo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159081)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Rotinas_Le_Todas(colRotina As Collection) As Long
'Lê Siglas de todas as Rotinas do BD e devolve em colRotina

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sRotina As String  'sigla da Rotina

On Error GoTo Erro_Rotinas_Le_Todas

    sRotina = String(STRING_ROTINA_SIGLA, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6462

    'seleciona no BD todos as siglas de Rotinas
    lErro = Comando_Executar(lComando, "SELECT Rotinas.Sigla FROM Rotinas ORDER BY Rotinas.Sigla", sRotina)
    If lErro <> AD_SQL_SUCESSO Then Error 6463

    'Lê a primeira sigla de rotina
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6464

    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colRotina.Add (sRotina)
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop

    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6465

    lErro = Comando_Fechar(lComando)

    Rotinas_Le_Todas = SUCESSO

    Exit Function

Erro_Rotinas_Le_Todas:

    Rotinas_Le_Todas = Err

    Select Case Err

        Case 6462
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6463, 6465
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTINA", Err)

        Case 6464
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ROTINAS_INEXISTENTES", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159082)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function GrupoTela_Le_GrupoModulo(ByVal sGrupo As String, ByVal sModulo As String, colGrupoTela As colGrupoTela) As Long
'Lê dados de GrupoTela correspondentes a sGrupo e a telas contidas em sModulo
'Retorna dados em colGrupoTela

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim tGrupoTela As typeGrupoTela

On Error GoTo Erro_GrupoTela_Le_GrupoModulo

    tGrupoTela.sNomeTela = String(STRING_GRUPO_TELA_NOMETELA, 0)
    tGrupoTela.sProjeto = String(STRING_GRUPO_TELA_PROJETO, 0)
    tGrupoTela.sClasse = String(STRING_GRUPO_TELA_CLASSE, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6495

    'seleciona no BD todos os GrupoTela correspondentes a sGrupo e telas no módulo sModulo
    lErro = Comando_Executar(lComando, "SELECT GrupoTela.NomeTela, TipoDeAcesso, projeto, classe FROM GrupoTela, TelasModulo, Modulos WHERE CodGrupo=? AND GrupoTela.NomeTela=TelasModulo.NomeTela AND TelasModulo.SiglaModulo=Modulos.Sigla AND Modulos.Nome=? ORDER BY GrupoTela.NomeTela", tGrupoTela.sNomeTela, tGrupoTela.iTipoDeAcesso, tGrupoTela.sProjeto, tGrupoTela.sClasse, sGrupo, sModulo)
    If lErro <> AD_SQL_SUCESSO Then Error 6496

    'Lê o primeiro registro selecionado
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If iFim_de_Arquivo = AD_SQL_SEM_DADOS Then Error 6497

    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colGrupoTela.Add tGrupoTela.sProjeto, tGrupoTela.sClasse, sGrupo, tGrupoTela.iTipoDeAcesso, tGrupoTela.sNomeTela
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop

    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6498

    lErro = Comando_Fechar(lComando)

    GrupoTela_Le_GrupoModulo = SUCESSO

    Exit Function

Erro_GrupoTela_Le_GrupoModulo:

    GrupoTela_Le_GrupoModulo = Err

    Select Case Err

        Case 6495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6496, 6498
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_TELA", Err)

        Case 6497
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOTELA_GRUPOMODULO_NAO_CADASTRADOS", Err, sGrupo, sModulo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159083)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function Telas_Le_Todas(colTela As Collection) As Long
'Lê Nomes de todas as Telas do BD e devolve em colTela

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sTela As String  'nome da Tela

On Error GoTo Erro_Telas_Le_Todas

    sTela = String(STRING_TELA_NOME, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6499

    'seleciona no BD todos os nomes de Telas
    lErro = Comando_Executar(lComando, "SELECT Telas.Nome FROM Telas ORDER BY Telas.Nome", sTela)
    If lErro <> AD_SQL_SUCESSO Then Error 6500

    'Lê o primeiro nome de tela
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6501

    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colTela.Add (sTela)
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop

    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6502

    lErro = Comando_Fechar(lComando)

    Telas_Le_Todas = SUCESSO

    Exit Function

Erro_Telas_Le_Todas:

    Telas_Le_Todas = Err

    Select Case Err

        Case 6499
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6500, 6502
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELA", Err)

        Case 6501
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELAS_INEXISTENTES", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159084)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'Public Function MenuItem_Le(ByVal sTipoRotina As String, ByVal sModulo As String, colMenuItem As colMenuItem) As Long
''Lê os Itens de Menu com siglas de Rotinas do sTipoRotina contidas em sModulo.
'
'Dim lComando As Long
'Dim lErro As Long
'Dim iFim_de_Arquivo As Integer
'Dim sSiglaRotina As String
'Dim sTitulo As String
'Dim sSubstring As String
'
'On Error GoTo Erro_MenuItem_Le
'
'    'Substring usada na seleção dos ítens de menu
'    sSubstring = "%" & "_USU_" & sTipoRotina & "%"
'
'    sTitulo = String(STRING_MENU_ITEM_TITULO, 0)
'    sSiglaRotina = String(STRING_MENU_ITEM_SIGLAROTINA, 0)
'
'    lComando = Comando_AbrirExt(GL_lConexaoDic)
'    If lComando = 0 Then Error 6540
'
'    'seleciona no BD todos os MenuItens correspondentes a rotinas do sTipoRotina no sModulo
'    lErro = Comando_Executar(lComando, "SELECT Titulo, MenuItens.SiglaRotina FROM MenuItens, RotinasModulo, Modulos WHERE MenuItens.SiglaRotina LIKE ? AND MenuItens.SiglaRotina=RotinasModulo.SiglaRotina AND RotinasModulo.SiglaModulo=Modulos.Sigla AND Modulos.Nome=? ORDER BY MenuItens.SiglaRotina", sTitulo, sSiglaRotina, sSubstring, sModulo)
'    If lErro <> AD_SQL_SUCESSO Then Error 6541
'
'    'Lê o primeiro registro selecionado
'    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
'    If iFim_de_Arquivo = AD_SQL_SEM_DADOS Then Error 6542
'
'    'Preenche a coleção
'    Do While iFim_de_Arquivo = AD_SQL_SUCESSO
'
'        'Adiciona à Coleção
'        colMenuItem.Add sTitulo, sSiglaRotina, "", 0
'        iFim_de_Arquivo = Comando_BuscarProximo(lComando)
'
'    Loop
'
'    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6543
'
'    lErro = Comando_Fechar(lComando)
'
'    MenuItem_Le = SUCESSO
'
'    Exit Function
'
'Erro_MenuItem_Le:
'
'    MenuItem_Le = Err
'
'    Select Case Err
'
'        Case 6540
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
'
'        Case 6541, 6543
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MENU_ITENS", Err)
'
'        Case 6542
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_MENUITEM_ROTINAS_MODULO_NAO_CADASTRADOS", Err, sTipoRotina, sModulo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159085)
'
'    End Select
'
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Public Function GrupoBrowseCampo_Le2(ByVal sGrupo As String, ByVal sTela As String, ByVal sArquivo As String, colCampo As Collection) As Long
'Lê na tabela GrupoBrowseCampo todos os Campos associados
'a sGrupo, sTela, sArquivo. Retorna em colCampo.
     
Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sCampo As String

On Error GoTo Erro_GrupoBrowseCampo_Le2

    sCampo = String(STRING_GBC_NOME, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6556

    'seleciona no BD todos os nomes de Campos associados a sGrupo, sTela, sArquivo
    lErro = Comando_Executar(lComando, "SELECT Nome FROM GrupoBrowseCampo WHERE CodGrupo=? AND NomeTela=? AND NomeArq=? ORDER BY Nome", sCampo, sGrupo, sTela, sArquivo)
    If lErro <> AD_SQL_SUCESSO Then Error 6557

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6558
    
    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colCampo.Add (sCampo)
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6559

    lErro = Comando_Fechar(lComando)

    GrupoBrowseCampo_Le2 = SUCESSO

    Exit Function

Erro_GrupoBrowseCampo_Le2:

    GrupoBrowseCampo_Le2 = Err

    Select Case Err

        Case 6556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6557, 6559
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPOBROWSECAMPO", Err)
            
        Case 6558
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_CAMPO_GRUPOBROWSECAMPO", Err, sGrupo, sTela, sArquivo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159086)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function BrowseArquivo_Le_Arquivos(ByVal sTela As String, colArquivo As Collection) As Long
'Lê na tabela BrowseArquivo todos os nomes de arquivos
'associados a tela sTela. Retorna em colArquivo.

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sArquivo As String

On Error GoTo Erro_BrowseArquivo_Le_Arquivos

    sArquivo = String(STRING_BROWSEARQUIVO_NOME_ARQ, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6560

    'seleciona no BD todos os nomes de Arquivos associados a sTela
    lErro = Comando_Executar(lComando, "SELECT NomeArq FROM BrowseArquivo WHERE NomeTela=? ORDER BY NomeArq", sArquivo, sTela)
    If lErro <> AD_SQL_SUCESSO Then Error 6561

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6562
    
    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colArquivo.Add (sArquivo)
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6563

    lErro = Comando_Fechar(lComando)

    BrowseArquivo_Le_Arquivos = SUCESSO

    Exit Function

Erro_BrowseArquivo_Le_Arquivos:

    BrowseArquivo_Le_Arquivos = Err

    Select Case Err

        Case 6560
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6561, 6563
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEARQUIVO", Err)
            
        Case 6562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ARQUIVO_BROWSEARQUIVO", Err, sTela)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159087)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function BrowseArquivo_Le_Telas(ByVal sModulo As String, colTela As Collection) As Long
'Lê todas os nomes de telas na tabela BrowseArquivo que
'pertencem ao módulo sModulo. Retorna os nomes em colTela.

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sTela As String

On Error GoTo Erro_BrowseArquivo_Le_Telas

    sTela = String(STRING_BROWSEARQUIVO_NOME_TELA, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6564

    'seleciona no BD todos os nomes de Telas de Browse de sModulo
    lErro = Comando_Executar(lComando, "SELECT DISTINCT BrowseArquivo.NomeTela FROM BrowseArquivo, TelasModulo, Modulos WHERE BrowseArquivo.NomeTela=TelasModulo.NomeTela AND TelasModulo.SiglaModulo=Modulos.Sigla AND Modulos.Nome=? ORDER BY BrowseArquivo.NomeTela", sTela, sModulo)
    If lErro <> AD_SQL_SUCESSO Then Error 6565

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6566
    
    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colTela.Add (sTela)
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6567

    lErro = Comando_Fechar(lComando)

    BrowseArquivo_Le_Telas = SUCESSO

    Exit Function

Erro_BrowseArquivo_Le_Telas:

    BrowseArquivo_Le_Telas = Err

    Select Case Err

        Case 6564
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6565, 6567
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEARQUIVO", Err)

        Case 6566
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_TELAS_BROWSEARQUIVO", Err, sModulo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159088)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function Campos_Le3(ByVal sArquivo As String, colCampo As Collection) As Long
'Lê os nomes de todos os campos pertencentes a sArquivo.
'Retorna em colCampo

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sCampo As String

On Error GoTo Erro_Campos_Le3

    sCampo = String(STRING_CAMPOS_NOME, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6568

    'seleciona no BD todos os nomes de Campos de sArquivo
    lErro = Comando_Executar(lComando, "SELECT Nome FROM Campos WHERE NomeArq=? ORDER BY Nome", sCampo, sArquivo)
    If lErro <> AD_SQL_SUCESSO Then Error 6569

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6570
    
    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colCampo.Add (sCampo)
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6571

    lErro = Comando_Fechar(lComando)

    Campos_Le3 = SUCESSO

    Exit Function

Erro_Campos_Le3:

    Campos_Le3 = Err

    Select Case Err

        Case 6568
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6569, 6571
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOS", Err)

        Case 6570
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_CAMPO_CAMPOS", Err, sArquivo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159089)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Usuarios_Le_Todos1(colUsuarios As Collection) As Long
'le todos os usuários ativos contidos na tabela de Usuarios e coloca os dados em colUsuarios

Dim lComando As Long
Dim lErro As Long
Dim tUsuarios As typeDicUsuario
Dim objUsuarios As ClassDicUsuario
    
On Error GoTo Erro_Usuarios_Le_Todos1

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 50185

    tUsuarios.sCodUsuario = String(STRING_SIGLA_USUARIO, 0)
    tUsuarios.sNome = String(STRING_USUARIO, 0)
    tUsuarios.sCodGrupo = String(STRING_GRUPO_CODIGO, 0)
    tUsuarios.sComputador = String(255, 0)
    tUsuarios.sNomeLogin = String(255, 0)
    tUsuarios.sEmail = String(255, 0)

    lErro = Comando_Executar(lComando, "SELECT CodUsuario, Nome, CodGrupo, DataValidade, Computador, NomeLogin, Logado, Email FROM Usuarios WHERE Ativo=?", tUsuarios.sCodUsuario, tUsuarios.sNome, tUsuarios.sCodGrupo, tUsuarios.dtDataValidade, tUsuarios.sComputador, tUsuarios.sNomeLogin, tUsuarios.iLogado, tUsuarios.sEmail, ATIVIDADE)
    If lErro <> AD_SQL_SUCESSO Then Error 50186
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50187
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objUsuarios = New ClassDicUsuario
        
        objUsuarios.sCodUsuario = tUsuarios.sCodUsuario
        objUsuarios.sCodGrupo = tUsuarios.sCodGrupo
        objUsuarios.sNome = tUsuarios.sNome
        objUsuarios.dtDataValidade = tUsuarios.dtDataValidade
        objUsuarios.sComputador = tUsuarios.sComputador
        objUsuarios.sNomeLogin = tUsuarios.sNomeLogin
        objUsuarios.iLogado = tUsuarios.iLogado
        objUsuarios.sEmail = tUsuarios.sEmail
        
        colUsuarios.Add objUsuarios
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50188
            
    Loop
            
    Comando_Fechar (lComando)
    
    Usuarios_Le_Todos1 = SUCESSO
    
    Exit Function
    
Erro_Usuarios_Le_Todos1:

    Usuarios_Le_Todos1 = Err

    Select Case Err
    
        Case 50185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 50186, 50187, 50188
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOS_DIC", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159090)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Empresas_Le_Usuario(sCodUsuario As String, colEmpresas As Collection) As Long
'Le as Empresas que o Usuario com sCodUsuario esta autorizado a utilizar e coloca em colEmpresas

Dim lComando As Long
Dim sNome As String
Dim objEmpresa As ClassDicEmpresa
Dim lCodigo As Long
Dim lErro As Long

On Error GoTo Erro_Empresas_Le_Usuario

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 50178

    sNome = String(STRING_NOME_EMPRESA, 0)

    'Seleciona todas as empresas que o usuario esta autorizado a utilizar
    lErro = Comando_Executar(lComando, "SELECT DISTINCT Nome, Codigo FROM Empresas, UsuFilEmp WHERE Empresas.Codigo=UsuFilEmp.CodEmpresa And UsuFilEmp.CodUsuario=?", sNome, lCodigo, sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 50179
    
    'Busca a primeira empresa
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50182
    
    'nao ha empresa
    If lErro = AD_SQL_SEM_DADOS Then Error 50183
    
    'emquanto houver empresa
    Do While lErro = AD_SQL_SUCESSO
    
        Set objEmpresa = New ClassDicEmpresa
    
        objEmpresa.sNome = sNome
        objEmpresa.lCodigo = lCodigo
    
        'adiciona a empresa a colecao
        colEmpresas.Add objEmpresa
    
        'Busca proxima empresa
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50184
    
    Loop
    
    Call Comando_Fechar(lComando)
    
    Empresas_Le_Usuario = SUCESSO
    
    Exit Function
    
Erro_Empresas_Le_Usuario:

    Empresas_Le_Usuario = Err
    
    Select Case Err
    
        Case 50178
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 50179, 50181, 50184
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA_USUARIO", Err, sCodUsuario)
        
        Case 50183 'não encontrou empresa
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159091)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function FiliaisEmpresa_Le_Usuario(sCodUsuario As String, lCodEmpresa As Long, colFilialEmpresa As Collection) As Long
'Le todas as filiais da empresa com lCodEmpresa selecionada que o usuario com sCodUsuario possa utilizar e coloca em colEmpresa

Dim lComando As Long
Dim sNome As String
Dim objUsuarioEmpresa As ClassUsuarioEmpresa
Dim iCodigo As Integer
Dim lErro As Long

On Error GoTo Erro_FiliaisEmpresa_Le_Usuario

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 50169

    sNome = String(STRING_FILIAL_NOME, 0)

    'seleciona as filiais da empresa em questão que o usuário está autorizado a acessar
    lErro = Comando_Executar(lComando, "SELECT DISTINCT FiliaisEmpresas.Nome, UsuFilEmp.CodFilial FROM UsuFilEmp LEFT JOIN FiliaisEmpresas ON UsuFilEmp.CodEmpresa = FiliaisEmpresas.CodEmpresa AND UsuFilEmp.CodFilial = FiliaisEmpresas.CodFilial WHERE UsuFilEmp.CodEmpresa=? AND UsuFilEmp.CodUsuario=?", sNome, iCodigo, lCodEmpresa, sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 50180

    'le a primeira filial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50171

    'se não houver nenhuma filial para a empresa em questão ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 50172

    'enquanto houverem filiais para a empresa em questão
    Do While lErro = AD_SQL_SUCESSO

        Set objUsuarioEmpresa = New ClassUsuarioEmpresa

        objUsuarioEmpresa.sNomeFilial = sNome
        objUsuarioEmpresa.iCodFilial = iCodigo

        'armazena a filial lida na coleção colFilialEmpresa
        colFilialEmpresa.Add objUsuarioEmpresa

        'lê a próxima filial
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 50174

    Loop

    Call Comando_Fechar(lComando)

    FiliaisEmpresa_Le_Usuario = SUCESSO

    Exit Function

Erro_FiliaisEmpresa_Le_Usuario:

    FiliaisEmpresa_Le_Usuario = Err

    Select Case Err

        Case 50169
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 50171, 50174, 50180
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALEMPRESA", Err, lCodEmpresa, sCodUsuario)

        Case 50172 'Não encontrou filiais associadas a esta empresa

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159092)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function Grupo_Obtem_Usuario(sCodGrupo As String, sCodUsuario As String) As Long
'obtem um usuario de um grupo

Dim lComando As Long
Dim lErro As Long

On Error GoTo Erro_Grupo_Obtem_Usuario
    
    sCodUsuario = String(STRING_USUARIO, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 8226

    'Seleciona as tuplas desejadas no BD, ordenadas pelo Código do grupo
    lErro = Comando_Executar(lComando, "SELECT CodUsuario FROM Usuarios WHERE CodGrupo=? ORDER BY CodUsuario", sCodUsuario, sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then Error 8227

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then
        If lErro = AD_SQL_SEM_DADOS Then
            'Não existe tal Grupo
            Error 8228
        Else
            Error 8229
        End If
    End If

    lErro = Comando_Fechar(lComando)

    Grupo_Obtem_Usuario = SUCESSO

Exit Function

Erro_Grupo_Obtem_Usuario:

    Grupo_Obtem_Usuario = Err

    Select Case Err

        Case 8226
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 8227, 8229
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USU_GRUPO", Err, sCodGrupo)

        Case 8228

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159093)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Empresas_Le(colEmpresa As Collection) As Long
'Lê Empresas e preenche colEmpresa com concatenações do código + nome
'??? acho que deveria usar CodNomes_le e jogar fora esta funcao
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim lComando As Long
Dim lCodigo As Long
Dim sCodigo As String
Dim sNome As String
Dim sEmpresa As String
Dim sEspacos As String
    
On Error GoTo Erro_Empresas_Le

    sNome = String(STRING_EMPRESA_NOME, 0)

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6186
    
    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome FROM Empresas ORDER BY Codigo", lCodigo, sNome)
    If lErro <> AD_SQL_SUCESSO Then Error 6187
    
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
        
    Do Until iFim_de_Arquivo <> AD_SQL_SUCESSO
    
        'Espaços que faltam para completar tamanho STRING_EMPRESA_CODIGO
        sEspacos = Space(STRING_EMPRESA_CODIGO - Len(CStr(lCodigo)))
        
        'Concatena Codigo e Nome da Empresa
        sCodigo = sEspacos & CStr(lCodigo)
        sEmpresa = sCodigo & SEPARADOR & Trim(sNome)
        
        'Adiciona à Coleção
        colEmpresa.Add (sEmpresa)

        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6188
    
    lErro = Comando_Fechar(lComando)
    
    Empresas_Le = SUCESSO
    
    Exit Function

Erro_Empresas_Le:

    Empresas_Le = Err
    
    Select Case Err
                
        Case 6186
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 6187, 6188
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA1", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159094)
        
        End Select
        
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Empresa_Automatica(lCodigo As Long) As Long
'retorna o próximo código de Empresa disponível
'??? como as empresas serao "desativadas" e nao excluidas poderia ter implementacao super simples
Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim lCodigo1 As Long

On Error GoTo Erro_Empresa_Automatica

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6181

    'Pesquisa os códigos das Empresas ordenados
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Empresas ORDER BY Codigo", lCodigo1)
    If lErro <> AD_SQL_SUCESSO Then Error 6182

    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(lComando)
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        lCodigo = 1
        lErro = Comando_Fechar(lComando)
        Empresa_Automatica = SUCESSO
        Exit Function
        
    End If
    
    If lErro <> AD_SQL_SUCESSO Then Error 6183
    
    lCodigo = lCodigo1 + 1
    iFim_de_Arquivo = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6184
    
    'Pesquisa o proximo código de Empresa disponível
    Do While iFim_de_Arquivo <> AD_SQL_SEM_DADOS
              
        If lCodigo < lCodigo1 Then Exit Do
        
        lCodigo = lCodigo1 + 1
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)
        If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6185
         
    Loop
        
    lErro = Comando_Fechar(lComando)
    
    Empresa_Automatica = SUCESSO
        
    Exit Function
        
Erro_Empresa_Automatica:

    Empresa_Automatica = Err

    Select Case Err
    
    Case 6181
        lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
    Case 6182, 6183, 6184, 6185
        lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA1", Err)
        
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159095)
        
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function FiliaisEmpresas_Le_Todas(colFilialEmpresa As Collection) As Long
'Le todas as filiais ATIVAS de todas as empresas e preenche colFilialEmpresa com objetos objUsuarioEmpresa
'ordenadas por nome da empresa, nome da filial.

Dim lComando As Long
Dim sNomeEmp As String, sNomeFil As String
Dim objUsuarioEmpresa As ClassUsuarioEmpresa
Dim lCodEmp As Long, iCodFilial As Integer
Dim lErro As Long

On Error GoTo Erro_FiliaisEmpresas_Le_Todas

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 41800

    sNomeEmp = String(STRING_EMPRESA_NOME, 0)
    sNomeFil = String(STRING_FILIAL_NOME, 0)

    lErro = Comando_Executar(lComando, "SELECT Empresas.Codigo, Empresas.Nome, FiliaisEmpresas.CodFilial, FiliaisEmpresas.Nome FROM Empresas, FiliaisEmpresas WHERE FiliaisEmpresas.CodEmpresa=Empresas.Codigo And Empresas.Inativa=0 And FiliaisEmpresas.Inativa=0 ORDER BY Empresas.Nome, FiliaisEmpresas.Nome", lCodEmp, sNomeEmp, iCodFilial, sNomeFil)
    If lErro <> AD_SQL_SUCESSO Then Error 41801

    'le a primeira filial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41802

    'enquanto houverem filiais para a empresa em questão
    Do While lErro = AD_SQL_SUCESSO

        Set objUsuarioEmpresa = New ClassUsuarioEmpresa

        objUsuarioEmpresa.lCodEmpresa = lCodEmp
        objUsuarioEmpresa.sNomeEmpresa = sNomeEmp
        objUsuarioEmpresa.iCodFilial = iCodFilial
        objUsuarioEmpresa.sNomeFilial = sNomeFil

        'armazena a filial lida na coleção colFilialEmpresa
        colFilialEmpresa.Add objUsuarioEmpresa

        'lê a próxima filial
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41804

    Loop

    Call Comando_Fechar(lComando)

    FiliaisEmpresas_Le_Todas = SUCESSO

    Exit Function

Erro_FiliaisEmpresas_Le_Todas:

    FiliaisEmpresas_Le_Todas = Err

    Select Case Err

        Case 41800
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 41801, 41802, 41804
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESAS", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159096)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function UsuFilEmp_Le_Usuario(sCodUsuario As String, colUsuFilEmp As Collection) As Long
'preenche colecao (com objetos objUsuFilEmp) com todas as empresas-filiais a que o usuario tem acesso

Dim lErro As Long, objUsuFilEmp As ClassUsuFilEmp
Dim lComando As Long, iCodFilial As Integer, lCodEmpresa As Long

On Error GoTo Erro_UsuFilEmp_Le_Usuario

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 32041

    lErro = Comando_Executar(lComando, "SELECT CodEmpresa, CodFilial FROM UsuFilEmp WHERE CodUsuario = ?", lCodEmpresa, iCodFilial, sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 32042

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 32043
    
    'enquanto houverem registros
    Do While lErro = AD_SQL_SUCESSO

        Set objUsuFilEmp = New ClassUsuFilEmp

        objUsuFilEmp.sCodUsuario = sCodUsuario
        objUsuFilEmp.lCodEmpresa = lCodEmpresa
        objUsuFilEmp.iCodFilial = iCodFilial

        Call colUsuFilEmp.Add(objUsuFilEmp)
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 32044
    
    Loop

    Call Comando_Fechar(lComando)

    UsuFilEmp_Le_Usuario = SUCESSO

    Exit Function

Erro_UsuFilEmp_Le_Usuario:

    UsuFilEmp_Le_Usuario = Err

    Select Case Err

        Case 32041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 32042, 32043, 32044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUFILEMP_USU", Err, sCodUsuario)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159097)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function GrupoRelatorio_Le_GrupoModulo(ByVal sGrupo As String, ByVal sModulo As String, colGrupoRelatorio As Collection) As Long
'Lê dados de GrupoRelatorios correspondentes a sGrupo e a relatorios associados a sModulo
'Retorna dados em colGrupoRelatorio

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim objGrupoRelatorio As ClassDicGrupoRelatorio
Dim sCodRel As String, sNomeTsk As String, iTipoDeAcesso As Integer

On Error GoTo Erro_GrupoRelatorio_Le_GrupoModulo

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 32086

    sCodRel = String(STRING_CODIGO_RELATORIO, 0)
    sNomeTsk = String(STRING_NOME_ARQ_COMPLETO, 0)
    
    'seleciona no BD todos os GrupoRelatorio correspondentes a sGrupo e relatorios no módulo sModulo
    lErro = Comando_Executar(lComando, "SELECT GrupoRelatorios.CodRel, TipoDeAcesso, NomeTskCustomizado FROM GrupoRelatorios, RelatoriosModulo, Modulos WHERE CodGrupo=? AND GrupoRelatorios.CodRel=RelatoriosModulo.CodRel AND RelatoriosModulo.Modulo=Modulos.Sigla AND Modulos.Nome=? ORDER BY GrupoRelatorios.CodRel", sCodRel, iTipoDeAcesso, sNomeTsk, sGrupo, sModulo)
    If lErro <> AD_SQL_SUCESSO Then Error 32087

    'Lê o primeiro registro selecionado
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If iFim_de_Arquivo = AD_SQL_SEM_DADOS Then Error 32088

    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        Set objGrupoRelatorio = New ClassDicGrupoRelatorio
        
        objGrupoRelatorio.iTipoDeAcesso = iTipoDeAcesso
        objGrupoRelatorio.sCodGrupo = sGrupo
        objGrupoRelatorio.sCodRel = sCodRel
        objGrupoRelatorio.sNomeTskCustomizado = sNomeTsk
        
        'Adiciona à Coleção
        colGrupoRelatorio.Add objGrupoRelatorio
        
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop

    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 32089

    lErro = Comando_Fechar(lComando)

    GrupoRelatorio_Le_GrupoModulo = SUCESSO

    Exit Function

Erro_GrupoRelatorio_Le_GrupoModulo:

    GrupoRelatorio_Le_GrupoModulo = Err

    Select Case Err

        Case 32086
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 32087, 32089
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_ROTINAS", Err)

        Case 32088
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPORELATORIO_GRUPOMODULO_NAO_CADASTRADOS", Err, sGrupo, sModulo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159098)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function GrupoRelatorio_Le_Relatorio(sCodRel As String, colGrupoRelatorio As Collection) As Long
'Lê dados de GrupoRelatorios correspondentes a este Relatorio
'Retorna dados em colGrupoRelatorio

Dim lComando As Long
Dim lErro As Long, objGrupoRelatorio As ClassDicGrupoRelatorio
Dim iFim_de_Arquivo As Integer, iTipoDeAcesso As Integer
Dim sCodGrupo As String, sNomeTskCustomizado As String

On Error GoTo Erro_GrupoRelatorio_Le_Relatorio
    
    sCodGrupo = String(STRING_GRUPO_CODIGO, 0)
    sNomeTskCustomizado = String(STRING_NOME_ARQ_COMPLETO, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 32120

    'seleciona no BD todos os GrupoRelatorio correspondentes a sCodRel
    lErro = Comando_Executar(lComando, "SELECT CodGrupo, TipoDeAcesso, NomeTskCustomizado FROM GrupoRelatorios WHERE CodRel=? ORDER BY CodGrupo", sCodGrupo, iTipoDeAcesso, sNomeTskCustomizado, sCodRel)
    If lErro <> AD_SQL_SUCESSO Then Error 32121

    'Lê o primeiro registro selecionado
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If iFim_de_Arquivo = AD_SQL_SEM_DADOS Then Error 32122
    
    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        Set objGrupoRelatorio = New ClassDicGrupoRelatorio
        
        objGrupoRelatorio.iTipoDeAcesso = iTipoDeAcesso
        objGrupoRelatorio.sCodGrupo = sCodGrupo
        objGrupoRelatorio.sCodRel = sCodRel
        objGrupoRelatorio.sNomeTskCustomizado = sNomeTskCustomizado
        
        colGrupoRelatorio.Add objGrupoRelatorio
        
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 32123

    lErro = Comando_Fechar(lComando)

    GrupoRelatorio_Le_Relatorio = SUCESSO

    Exit Function

Erro_GrupoRelatorio_Le_Relatorio:

    GrupoRelatorio_Le_Relatorio = Err

    Select Case Err

        Case 32120
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 32121, 32123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_RELATORIOS", Err)
            
        Case 32122
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPORELATORIO_REL_NAO_CADASTRADOS", Err, sCodRel)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159099)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Relatorios_Le_NomeModulo(sNomeModulo As String, colNomesRelatorios As Collection) As Long
'Lê os nomes de Relatorios associados ao Módulo

Dim lComando As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sCodRel As String

On Error GoTo Erro_Relatorios_Le_NomeModulo

    sCodRel = String(STRING_CODIGO_RELATORIO, 0)
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 32116

    'seleciona no BD todos os Relatorios associados ao sNomeModulo
    lErro = Comando_Executar(lComando, "SELECT RelatoriosModulo.CodRel FROM RelatoriosModulo, Modulos WHERE RelatoriosModulo.Modulo=Modulos.Sigla AND Modulos.Nome=? ORDER BY RelatoriosModulo.CodRel", sCodRel, sNomeModulo)
    If lErro <> AD_SQL_SUCESSO Then Error 32117

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 32118
    
    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Adiciona à Coleção
        colNomesRelatorios.Add (sCodRel)
        
        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop
    
    If iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 32119

    lErro = Comando_Fechar(lComando)

    Relatorios_Le_NomeModulo = SUCESSO

    Exit Function

Erro_Relatorios_Le_NomeModulo:

    Relatorios_Le_NomeModulo = Err

    Select Case Err

        Case 32116
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 32117, 32119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTINA", Err)
            
        Case 32118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIOS_DO_MODULO_INEXISTENTES", Err, sNomeModulo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159100)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Modulos_Le2(colModulo As Collection) As Long
'Le todos os modulos liberados contidos na tabela Modulos no dicionario de dados e coloca-os em colModulo

Dim lComando As Long
Dim lErro As Long
Dim objModulo As AdmModulo
Dim tModulo As typeModulo
Dim lConexao As Long
Dim objVersao As New ClassVersao

On Error GoTo Erro_Modulos_Le2

    lConexao = GL_lConexaoDic

    tModulo.sSigla = String(STRING_MODULO_SIGLA, 0)
    tModulo.sNome = String(STRING_MODULO_NOME, 0)
    tModulo.sDescricao = String(STRING_MODULO_DESCRICAO, 0)
    tModulo.sVersao = String(STRING_MODULO_VERSAO, 0)

    'abre comando
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 25463
    
    'Lê a versão vigente
    lErro = VersaoVigente_Le(objVersao)
    If lErro <> SUCESSO Then Error 25809
    
    'Lê todos os registros da tabela Modulos em DIC até a versão vigente
    lErro = Comando_Executar(lComando, "SELECT Sigla, Nome, Descricao, Versao FROM Modulos, Versao WHERE Modulos.Versao = Versao.Codigo AND Versao.Data <= ? AND Modulos.Liberado = ? ORDER BY Sigla ", tModulo.sSigla, tModulo.sNome, tModulo.sDescricao, tModulo.sVersao, objVersao.dtData, MODULO_LIBERADO)
    If lErro <> AD_SQL_SUCESSO Then Error 9642

    'le o primeiro modulo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 9643

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objModulo = New AdmModulo

        objModulo.sSigla = tModulo.sSigla
        objModulo.sNome = tModulo.sNome
        objModulo.sDescricao = tModulo.sDescricao
        objModulo.sVersao = tModulo.sVersao
        
        'coloca o modulo lido na coleção
        colModulo.Add objModulo

        'le o proximo modulo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25465

    Loop

    Call Comando_Fechar(lComando)

    Modulos_Le2 = SUCESSO

    Exit Function

Erro_Modulos_Le2:

    Modulos_Le2 = Err

    Select Case Err

        Case 25463
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 9642, 9643, 25465
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULO", Err)

        Case 25809 'tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159101)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function VersaoVigente_Le(objVersao As ClassVersao) As Long
'Lê versão vigente e devolve em objVersao

Dim lErro As Long
Dim lComando As Long
Dim tVersao As typeVersao
Dim lConexao As Long

On Error GoTo Erro_VersaoVigente_Le

    lConexao = GL_lConexaoDic

    tVersao.sCodigo = String(STRING_VERSAO_CODIGO, 0)

    'abre comandos
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 25802
    
    lErro = Comando_Executar(lComando, "SELECT Codigo, Data FROM Versao WHERE Vigente = ? ", tVersao.sCodigo, tVersao.dtData, VERSAO_VIGENTE)
    If lErro <> AD_SQL_SUCESSO Then Error 25803
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25804
    If lErro <> AD_SQL_SUCESSO Then Error 25805
    
    objVersao.sCodigo = tVersao.sCodigo
    objVersao.dtData = tVersao.dtData
    objVersao.iVigente = VERSAO_VIGENTE
    
    Call Comando_Fechar(lComando)
    
    VersaoVigente_Le = SUCESSO

    Exit Function

Erro_VersaoVigente_Le:

    VersaoVigente_Le = Err

    Select Case Err

        Case 25802
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 25803, 25804
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VERSAO", Err)
        
        Case 25805
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VERSAO_VIGENTE_AUSENTE", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159102)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function



Function Empresas_Le_Todas(colEmpresa As Collection) As Long
'Le todas as Empresas ATIVAS e retorna em colEmpresa.

Dim lComando As Long
Dim lErro As Long
Dim tEmpresa As typeDicEmpresa
Dim objEmpresa As ClassDicEmpresa

On Error GoTo Erro_Empresas_Le_Todas

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6177

    tEmpresa.sNome = String(STRING_EMPRESA_NOME, 0)
    tEmpresa.sNomeReduzido = String(STRING_EMPRESA_NOME_RED, 0)
    tEmpresa.sStringConexao = String(STRING_STRING_CONEXAO, 0)

    lErro = Comando_Executar(lComando, "SELECT Codigo, Nome, NomeReduzido, StringConexao FROM Empresas WHERE Inativa = ? ORDER BY NomeReduzido", tEmpresa.lCodigo, tEmpresa.sNome, tEmpresa.sNomeReduzido, tEmpresa.sStringConexao, Ativo)
    If lErro Then Error 6178

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6180
    If lErro = AD_SQL_SEM_DADOS Then Error 6179 'NÃO HÁ EMPRESAS

    Do While lErro <> AD_SQL_SEM_DADOS

        Set objEmpresa = New ClassDicEmpresa
        
        'Transporta para o obj
        objEmpresa.lCodigo = tEmpresa.lCodigo
        objEmpresa.sNome = tEmpresa.sNome
        objEmpresa.sNomeReduzido = tEmpresa.sNomeReduzido
        objEmpresa.sStringConexao = tEmpresa.sStringConexao
        objEmpresa.iInativa = tEmpresa.iInativa

        'Adiciona à coleção
        colEmpresa.Add objEmpresa

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25471

    Loop

    lErro = Comando_Fechar(lComando)

    Empresas_Le_Todas = SUCESSO

    Exit Function

Erro_Empresas_Le_Todas:

    Empresas_Le_Todas = Err

    Select Case Err

        Case 6177
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6178, 6180
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA", Err, objEmpresa.lCodigo)

        Case 6179  'NÃO HÁ EMPRESAS. Será tratado na rotina chamadora.

        Case 25471 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159103)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Modulos_Le_Empresa1(ByVal lEmpresa As Long, colModulo As Collection) As Long
'Le em DIC todos os modulos LIBERADOS associados a lEmpresa contidos na tabela Modulos no dicionario de dados e coloca-os em colModulo

Dim lComando As Long
Dim lErro As Long
Dim objModulo As AdmModulo
Dim tModulo As typeModulo
Dim lConexao As Long
Dim sSigla As String
Dim objVersao As New ClassVersao

On Error GoTo Erro_Modulos_Le_Empresa1

    lConexao = GL_lConexaoDic

    tModulo.sSigla = String(STRING_MODULO_SIGLA, 0)
    tModulo.sNome = String(STRING_MODULO_NOME, 0)
    tModulo.sDescricao = String(STRING_MODULO_DESCRICAO, 0)
    tModulo.sVersao = String(STRING_MODULO_VERSAO, 0)
    
    'abre conexao para DIC
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 25526
    
    'Lê a versão vigente
    lErro = VersaoVigente_Le(objVersao)
    If lErro <> SUCESSO Then Error 60805
    
    'le todos os registros da tabela Modulos em DIC
    lErro = Comando_Executar(lComando, "SELECT Sigla, Nome, Descricao, Versao, Ativo FROM Modulos, ModuloEmpresa, Versao WHERE Modulos.Sigla=ModuloEmpresa.SiglaModulo And ModuloEmpresa.CodEmpresa = ? AND Modulos.Versao = Versao.Codigo AND Versao.Data <= ? AND Modulos.Liberado = ? ORDER BY Sigla", tModulo.sSigla, tModulo.sNome, tModulo.sDescricao, tModulo.sVersao, tModulo.iAtivo, lEmpresa, objVersao.dtData, MODULO_LIBERADO)
    If lErro <> AD_SQL_SUCESSO Then Error 25527
    
    'le o primeiro modulo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25528
    
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objModulo = New AdmModulo
        
        objModulo.sSigla = tModulo.sSigla
        objModulo.sNome = tModulo.sNome
        objModulo.sDescricao = tModulo.sDescricao
        objModulo.sVersao = tModulo.sVersao
        objModulo.iAtivo = tModulo.iAtivo
        
        'coloca o modulo lido na coleção
        colModulo.Add objModulo
        
        'le o proximo modulo em DIC
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25529

    Loop

    Call Comando_Fechar(lComando)
    
    Modulos_Le_Empresa1 = SUCESSO

    Exit Function

Erro_Modulos_Le_Empresa1:

    Modulos_Le_Empresa1 = Err
    
    Select Case Err

        Case 25526
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25527, 25528, 25529
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULO", Err)
         
        Case 60805
         
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159104)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function FiliaisEmpresas_Le(colFilialEmpresa As Collection) As Long
'Le todas as filiais ATIVAS de todas as empresas e preenche colFilialEmpresa com objetos objFilialEmpresa
'ordenadas por nome reduzido da empresa, nome da filial.

Dim lComando As Long
Dim tFilialEmpresa As typeFilialEmpresa
Dim objFilialEmpresa As ClassFilialEmpresa
Dim lErro As Long

On Error GoTo Erro_FiliaisEmpresas_Le

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 25517

    With tFilialEmpresa
        .sNomeEmpresa = String(STRING_EMPRESA_NOME, 0)
        .sNomeFilial = String(STRING_FILIAL_NOME, 0)
        .sNomeRedEmpresa = String(STRING_EMPRESA_NOME_RED, 0)

        lErro = Comando_Executar(lComando, "SELECT Empresas.Codigo, Empresas.Nome, Empresas.NomeReduzido, FiliaisEmpresas.CodFilial, FiliaisEmpresas.Nome FROM Empresas, FiliaisEmpresas WHERE FiliaisEmpresas.CodEmpresa=Empresas.Codigo And Empresas.Inativa=0 And FiliaisEmpresas.Inativa=0 ORDER BY Empresas.NomeReduzido, FiliaisEmpresas.Nome", .lCodEmpresa, .sNomeEmpresa, .sNomeRedEmpresa, .iCodFilial, .sNomeFilial)
    End With
    
    If lErro <> AD_SQL_SUCESSO Then Error 25518

    'Le a primeira filial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25519

    'Enquanto houver filiais para a empresa em questão
    Do While lErro = AD_SQL_SUCESSO

        Set objFilialEmpresa = New ClassFilialEmpresa

        With tFilialEmpresa
            objFilialEmpresa.lCodEmpresa = .lCodEmpresa
            objFilialEmpresa.sNomeEmpresa = .sNomeEmpresa
            objFilialEmpresa.sNomeRedEmpresa = .sNomeRedEmpresa
            objFilialEmpresa.iCodFilial = .iCodFilial
            objFilialEmpresa.sNomeFilial = .sNomeFilial
        End With

        'Armazena a filial lida na coleção colFilialEmpresa
        colFilialEmpresa.Add objFilialEmpresa

        'Lê a próxima filial
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25520

    Loop

    Call Comando_Fechar(lComando)

    FiliaisEmpresas_Le = SUCESSO

    Exit Function

Erro_FiliaisEmpresas_Le:

    FiliaisEmpresas_Le = Err

    Select Case Err

        Case 25517
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25518, 25519, 25520
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESAS", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159105)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FilialEmpresa_Le2(objFilialEmpresa As AdmFiliais) As Long
'Le FilialEmpresa ATIVA só no DIC a partir da chave passada.

Dim lComando As Long
Dim tFilialEmpresa As typeFilialEmpresa
Dim lErro As Long

On Error GoTo Erro_FilialEmpresa_Le2

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 25577

    With tFilialEmpresa
        .sNomeFilial = String(STRING_FILIAL_NOME, 0)

        lErro = Comando_Executar(lComando, "SELECT Nome FROM FiliaisEmpresas WHERE CodEmpresa = ? AND CodFilial = ? AND Inativa=0", .sNomeFilial, objFilialEmpresa.lCodEmpresa, objFilialEmpresa.iCodFilial)
    End With
    
    If lErro <> AD_SQL_SUCESSO Then Error 25578

    'Busca a Filial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25579
    If lErro = AD_SQL_SEM_DADOS Then Error 25576

    'Armazena Nome da Filial
    objFilialEmpresa.sNome = tFilialEmpresa.sNomeFilial

    Call Comando_Fechar(lComando)

    FilialEmpresa_Le2 = SUCESSO

    Exit Function

Erro_FilialEmpresa_Le2:

    FilialEmpresa_Le2 = Err

    Select Case Err

        Case 25577
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25578, 25579
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESAS", Err)
            
        Case 25576 'Não encontrou Filial. Tratado na rotina chamadora.

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159106)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FilialEmpresa_Automatica(objEmpresa As ClassDicEmpresa, iCodFilial As Integer) As Long
'Calcula próximo código de Filial Empresa p/ ser usado

Dim lComando As Long
Dim lErro As Long
Dim lConexao As Long

On Error GoTo Erro_FilialEmpresa_Automatica

    lConexao = GL_lConexaoDic
    
    'abre conexao para DIC
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 25553
    
    'Seleciona Filiais ligadas a Empresa
    lErro = Comando_Executar(lComando, "SELECT CodFilial FROM FiliaisEmpresas WHERE CodEmpresa = ? ORDER BY CodFilial DESC", iCodFilial, objEmpresa.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 25554
    
    'Le o último código de Filial
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25555
    If lErro = AD_SQL_SEM_DADOS Then iCodFilial = 0
    
    'Retorna próximo código disponível
    iCodFilial = iCodFilial + 1

    Call Comando_Fechar(lComando)
    
    FilialEmpresa_Automatica = SUCESSO

    Exit Function

Erro_FilialEmpresa_Automatica:

    FilialEmpresa_Automatica = Err
    
    Select Case Err

        Case 25553
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25554, 25555
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESAS", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159107)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function ModuloFilEmp_Le_EmpresaFilial(lCodEmpresa As Long, iFilialEmpresa As Integer, colModuloFilEmp As Collection) As Long
'le todos os objetos ModuloFilEmp para a empresa/filial em questãp e coloca-os em colModuloFilEmp

Dim lComando As Long
Dim lErro As Long
Dim tModuloFilEmp As typeModuloFilEmp
Dim objModuloFilEmp As ClassModuloFilEmp
    
On Error GoTo Erro_ModuloFilEmp_Le_EmpresaFilial

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 44853

    tModuloFilEmp.sSiglaModulo = String(STRING_MODULO_SIGLA, 0)

    lErro = Comando_Executar(lComando, "SELECT SiglaModulo, Configurado FROM ModuloFilEmp  WHERE CodEmpresa=? AND CodFilial=?", tModuloFilEmp.sSiglaModulo, tModuloFilEmp.iConfigurado, lCodEmpresa, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 44854
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44855
    
    Do While lErro = AD_SQL_SUCESSO
    
        Set objModuloFilEmp = New ClassModuloFilEmp
    
        objModuloFilEmp.lCodEmpresa = lCodEmpresa
        objModuloFilEmp.iCodFilial = iFilialEmpresa
        objModuloFilEmp.sSiglaModulo = tModuloFilEmp.sSiglaModulo
        objModuloFilEmp.iConfigurado = tModuloFilEmp.iConfigurado
        
        colModuloFilEmp.Add objModuloFilEmp
               
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44856
               
    Loop
               
    Call Comando_Fechar(lComando)
    
    ModuloFilEmp_Le_EmpresaFilial = SUCESSO
    
    Exit Function
    
Erro_ModuloFilEmp_Le_EmpresaFilial:

    ModuloFilEmp_Le_EmpresaFilial = Err

    Select Case Err
    
        Case 44853
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                
        Case 44854, 44855, 44856
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULOFILEMP_DIC", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159108)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function DicConfig_Le(objDicConfig As ClassDicConfig) As Long
'Lê os Dados de configuração do sistema
 
Dim lComando As Long
Dim lErro As Long
Dim tDicConfig As typeDicConfig
Dim lConexao As Long

On Error GoTo Erro_DicConfig_Le

    lConexao = GL_lConexaoDic

    tDicConfig.sSenha = String(STRING_DICCONFIG_SENHA, 0)
    tDicConfig.sSerie = String(STRING_DICCONFIG_SERIE, 0)

    'Abre comando
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 62334
       
    'Lê o registro de configuração do sistema
    lErro = Comando_Executar(lComando, "SELECT ID,LimiteLogs,LimiteFiliais,Serie,TipoVersao,LimiteEmpresas,ValidadeDe,Senha,DataSenha FROM DicConfig", tDicConfig.iID, tDicConfig.iLimiteLogs, tDicConfig.iLimiteFiliais, tDicConfig.sSerie, tDicConfig.iTipoVersao, tDicConfig.iLimiteEmpresas, tDicConfig.dtValidadeDe, tDicConfig.sSenha, tDicConfig.dtDataSenha)
    If lErro <> AD_SQL_SUCESSO Then Error 62335

    'Lê o registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62336
    If lErro <> AD_SQL_SUCESSO Then Error 62337
    
    With tDicConfig
        objDicConfig.dtDataSenha = .dtDataSenha
        objDicConfig.dtValidadeDe = .dtValidadeDe
        objDicConfig.dtValidadeAte = DATA_NULA
        objDicConfig.iID = .iID
        objDicConfig.iLimiteEmpresas = .iLimiteEmpresas
        objDicConfig.iLimiteFiliais = .iLimiteFiliais
        objDicConfig.iLimiteLogs = .iLimiteLogs
        objDicConfig.iTipoVersao = .iTipoVersao
        objDicConfig.sSenha = .sSenha
        objDicConfig.sSerie = .sSerie
    End With

    Call Comando_Fechar(lComando)

    DicConfig_Le = SUCESSO

    Exit Function

Erro_DicConfig_Le:

    DicConfig_Le = Err

    Select Case Err

        Case 62334
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 62335, 62336
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DICCONFIG", Err)

        Case 62337
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGISTRO_DICCONFIG_NAO_ENCONTRADO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159109)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function FiliaisEmpresa_Le_Dados_Empresa(lEmpresa As Long, colFiliais As Collection) As Long
'Lê as Filiais da Empresa passada no SgeDados

Dim lErro As Long
Dim lComando As Long
Dim objFilial As AdmFiliais
Dim tFilialEmpresa As typeFiliais

On Error GoTo Erro_FiliaisEmpresa_Le_Dados_Empresa
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 62352
    
    'Inicializa as strigs
    With tFilialEmpresa
        
        .sCgc = String(STRING_CGC, 0)
        .sContador = String(STRING_CONTADOR, 0)
        .sCPFContador = String(STRING_CPF_CONTADOR, 0)
        .sCRCContador = String(STRING_CRC_CONTADOR, 0)
        .sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        .sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
        .sISSCodigoPadrao = String(STRING_ISS_CODIGO, 0)
        .sJucerja = String(STRING_JUCERJA, 0)
        .sNome = String(STRING_FILIAL_NOME, 0)
        .sRamoEmpresa = String(STRING_RAMO_EMPRESA, 0)
        .sNFCECSC = String(STRING_NFCE_CSC, 0)
        .sidNFCECSC = String(STRING_NFCE_ID_CSC, 0)
        
        'Busca as Filiais para a Empresa passada
        lErro = Comando_Executar(lComando, "SELECT FilialEmpresa,Nome,CGC,Endereco,AlmoxarifadoPadrao,InscricaoEstadual,ICMSPorEstimativa,InscricaoMunicipal,ISSCodigoPadrao,ISSIncluso,ContribuinteIPI,IRPercPadrao,RamoEmpresa,Jucerja,DataJucerja,Contador,CRCContador,CPFContador,TipoTribAtividade, SimplesFederal, SimplesFederalAliq, SimplesFederalTeto, SuperSimples, PISNaoCumulativo, COFINSNaoCumulativo, LucroPresumido, NFCECSC, idNFCECSC, indSincPadrao, RegimeTrib,RegimeEspecialTrib FROM FiliaisEmpresa", .iCodFilial, .sNome, .sCgc, .lEndereco, .iAlmoxarifadoPadrao, .sInscricaoEstadual, .iICMSPorEstimativa, _
            .sInscricaoMunicipal, .sISSCodigoPadrao, .iISSIncluso, .iContribuinteIPI, .dIRPercPadrao, .sRamoEmpresa, .sJucerja, .dtDataJucerja, .sContador, .sCRCContador, .sCPFContador, .iTipoTribAtividade, .iSimplesFederal, .dSimplesFederalAliq, .dSimplesFederalTeto, .iSuperSimples, .iPISNaoCumulativo, .iCOFINSNaoCumulativo, .iLucroPresumido, .sNFCECSC, .sidNFCECSC, .iindSincPadrao, .iRegimeTrib, .iRegimeEspecialTrib)
        If lErro <> AD_SQL_SUCESSO Then Error 62353
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62354
        'Para cada Filial Encontrada
        Do While lErro <> AD_SQL_SEM_DADOS
            
            'Armazena os dados da filial lida
            Set objFilial = New AdmFiliais
            
            objFilial.sCgc = .sCgc
            objFilial.sNome = .sNome
            objFilial.iInativa = .iInativa
            objFilial.sJucerja = .sJucerja
            objFilial.sContador = .sContador
            objFilial.iCodFilial = .iCodFilial
            objFilial.iISSIncluso = .iISSIncluso
            objFilial.lCodEmpresa = .lCodEmpresa
            objFilial.sRamoEmpresa = .sRamoEmpresa
            objFilial.sCPFContador = .sCPFContador
            objFilial.sCRCContador = .sCRCContador
            objFilial.dIRPercPadrao = .dIRPercPadrao
            objFilial.dtDataJucerja = .dtDataJucerja
            objFilial.dISSPercPadrao = .dISSPercPadrao
            objFilial.iContribuinteIPI = .iContribuinteIPI
            objFilial.sISSCodigoPadrao = .sISSCodigoPadrao
            objFilial.iICMSPorEstimativa = .iICMSPorEstimativa
            objFilial.iTipoTribAtividade = .iTipoTribAtividade
            objFilial.sInscricaoEstadual = .sInscricaoEstadual
            objFilial.iAlmoxarifadoPadrao = .iAlmoxarifadoPadrao
            objFilial.sInscricaoMunicipal = .sInscricaoMunicipal
            objFilial.iSimplesFederal = .iSimplesFederal
            objFilial.dSimplesFederalAliq = .dSimplesFederalAliq
            objFilial.dSimplesFederalTeto = .dSimplesFederalTeto
            objFilial.iSuperSimples = .iSuperSimples
            objFilial.iPISNaoCumulativo = .iPISNaoCumulativo
            objFilial.iCOFINSNaoCumulativo = .iCOFINSNaoCumulativo
            objFilial.iLucroPresumido = .iLucroPresumido
            
            'nfe 3.10
            objFilial.sNFCECSC = .sNFCECSC
            objFilial.sidNFCECSC = .sidNFCECSC
            objFilial.iindSincPadrao = .iindSincPadrao
            objFilial.iRegimeTrib = .iRegimeTrib
            objFilial.iRegimeEspecialTrib = .iRegimeEspecialTrib

            'Adiciona na combo de Filiais
            colFiliais.Add objFilial
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62355
            
        Loop
    
    End With
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    FiliaisEmpresa_Le_Dados_Empresa = SUCESSO
    
    Exit Function
    
Erro_FiliaisEmpresa_Le_Dados_Empresa:
    
    FiliaisEmpresa_Le_Dados_Empresa = Err
    
    Select Case Err
    
        Case 62352
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 62353, 62354, 62355
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESA", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159110)
        
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function Senha_Empresa_Decifra(sSenha As String, sCgc As String, sNomeEmpresa, iNumeroLogs As Integer, iNumeroEmpresas As Integer, iNumeroFiliais As Integer, colModulosLib As Collection, dtDataValidade As Date, sTextoSenha As String) As Long

Dim iIndice As Integer
Dim colModulos As New Collection
Dim objModulos As AdmModulo
Dim bModuloLiberado As Boolean
Dim lErro As Long
Dim sCheckSum As String
Dim sTextoAux As String
Dim sTextoAux2 As String
Dim sTextoAux3 As String

On Error GoTo Erro_Senha_Empresa_Decifra
    
    'Desdobra a senha
    Call Desintercala_Texto(10, sSenha, sTextoSenha, sTextoAux)
    Call Desintercala_Texto(7, sTextoSenha, sTextoAux2, sTextoAux3)
    sTextoSenha = sTextoAux2 & sTextoAux3 & sTextoAux
        
    sCgc = HexaParaDec(Mid(sTextoSenha, 1, 4), 4)
    'Lê todos os módulos
    lErro = CF("Modulos_Le_Todos", colModulos)
    If lErro <> SUCESSO Then Error 62333

    sTextoAux = DecParaBin(HexaParaDec(Mid(sTextoSenha, 5, 4), 16), 16)
   
    Set colModulosLib = New Collection
    'Para cada mósulo liberado adiciona o modulo na coleção de módulos
    For Each objModulos In colModulos
'        iIndice = iIndice + 1
'        If Mid(sTextoAux, iIndice, 1) = MODULO_LIBERADO Then colModulosLib.Add objModulos
        colModulosLib.Add objModulos
    Next

    'Recolhe os dados retirados da senha
    sNomeEmpresa = Chr(HexaParaDec(Mid(sTextoSenha, 9, 2), 3)) & Chr(HexaParaDec(Mid(sTextoSenha, 11, 2), 3))
    sTextoAux = "01/" & HexaParaDec(Mid(sTextoSenha, 13, 1), 2) & "/" & "20" & HexaParaDec(Mid(sTextoSenha, 14, 2), 2)
    dtDataValidade = DateAdd("m", 1, CDate(sTextoAux)) - 1
    iNumeroFiliais = HexaParaDec(Mid(sTextoSenha, 16, 2), 3)
    iNumeroLogs = HexaParaDec(Mid(sTextoSenha, 18, 4), 5)
    iNumeroEmpresas = HexaParaDec(Mid(sTextoSenha, 22, 2), 3)
    sCheckSum = HexaParaDec(Mid(sTextoSenha, 24, 2), 3)
    
    Senha_Empresa_Decifra = SUCESSO
    
    Exit Function

Erro_Senha_Empresa_Decifra:

    Senha_Empresa_Decifra = Err
    
    Select Case Err
    
        Case 62333
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159111)
            
    End Select
    
    Exit Function
    
End Function


Function Senha_Empresa_Gera(sCgc As String, sNomeEmpresa As String, iNumeroLogs As Integer, iNumeroEmpresas As Integer, iNumeroFiliais As Integer, colSiglasModulosLib As Collection, dtDataValidade As Date, sSenha As String) As Long
'Gera a senha para o sistema através do CGC, NomeEmpresa, Limites de Logs , empresas e filiais e a Data de validadae.

Dim lErro As Long
Dim asPartesSenha(1 To 7) As String
Dim iIndice As Integer
Dim colModulos As New Collection
Dim objModulos As AdmModulo
Dim bModuloLiberado As Boolean
Dim sTextoAux As String
Dim sCheckSum As String
Dim sTextoSenha As String

On Error GoTo Erro_Senha_Empresa_Gera
    
    For iIndice = LBound(asPartesSenha) To UBound(asPartesSenha)
        asPartesSenha(iIndice) = ""
    Next
    
    'Extrai os 4 primeiros dígitos do CGC
    asPartesSenha(1) = DecParaHexa(left(sCgc, 4), 4)
       
    'Lê todos os módulos
    lErro = CF("Modulos_Le_Todos", colModulos)
    If lErro <> SUCESSO Then Error 62333
    
    'Para cada módulo
    For Each objModulos In colModulos
        
        bModuloLiberado = False
        'MOnta o trecho da senha de liberação dos módulos
        For iIndice = 1 To colSiglasModulosLib.Count
            If objModulos.sSigla = colSiglasModulosLib(iIndice) Then
                bModuloLiberado = True
            End If
        Next
        
        If bModuloLiberado Then
            sTextoAux = sTextoAux & MODULO_LIBERADO
        Else
            sTextoAux = sTextoAux & MODULO_NAO_LIBERADO
        End If
        
    Next
    
    sTextoAux = sTextoAux & FormataCpoNum("", 16 - Len(sTextoAux))
    
    sTextoAux = "1111111111111111" 'Alterado por Wagner 15/04/2008 (Estorou o limite máximo de 16 módulos - Libera tudo)
    
    sTextoAux = BinParaDec(sTextoAux, 5)
    sTextoAux = DecParaHexa(sTextoAux, 4)
    asPartesSenha(2) = sTextoAux
    
    sTextoAux = DecParaHexa(Asc(Mid(sNomeEmpresa, 1, 1)), 2) & DecParaHexa(Asc(Mid(sNomeEmpresa, 2, 1)), 2)
    asPartesSenha(3) = sTextoAux
    
    sTextoAux = DecParaHexa(Month(dtDataValidade), 1) & DecParaHexa(Format(dtDataValidade, "YY"), 2)
    asPartesSenha(4) = sTextoAux
    
    asPartesSenha(5) = DecParaHexa(iNumeroFiliais, 2)
    asPartesSenha(6) = DecParaHexa(iNumeroLogs, 4)
    asPartesSenha(7) = DecParaHexa(iNumeroEmpresas, 2)
    
    sTextoAux = ""
    For iIndice = LBound(asPartesSenha) To UBound(asPartesSenha)
        sTextoAux = sTextoAux & asPartesSenha(iIndice)
    Next
    
    Call Calcula_CheckSum(sTextoAux, sCheckSum)
    
    sTextoAux = sTextoAux & sCheckSum
    
    sTextoSenha = Intercala_Texto(Mid(sTextoAux, 1, 8), Mid(sTextoAux, 9, 7))
    sTextoSenha = Intercala_Texto(sTextoSenha, Mid(sTextoAux, 16, 10))
    
    sSenha = sTextoSenha
    
    Senha_Empresa_Gera = SUCESSO
    
    Exit Function

Erro_Senha_Empresa_Gera:

    Senha_Empresa_Gera = Err
    
    Select Case Err
    
        Case 62333
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159112)
            
    End Select
    
    Exit Function
    
End Function

Sub Desintercala_Texto(iTamanhoTexto2 As Integer, sTextoTotal, sTexto1 As String, sTexto2 As String)
'Desintercala o textototal em sTexto1 e o sTexto2 de acordo com o tam do 2 texto que é passado
'Ex.  "10101010", tam 4 => "1111", "0000"
'     "10101011", tam 3 => "11111", "000"
'     "10101000", tam 5 => "111", "00000"

Dim iTamanhoTotal As Integer
Dim iIndice As Integer
Dim iTamanhoTexto1 As Integer
Dim iTamanhoMenor As Integer
    
    'Deduz o tamanho de texto 1
    iTamanhoTexto1 = Len(sTextoTotal) - iTamanhoTexto2
    
    sTexto1 = ""
    sTexto2 = ""
    
    'Verifica qual dos 2 textos é o menor
    iTamanhoMenor = IIf(iTamanhoTexto1 > iTamanhoTexto2, iTamanhoTexto2, iTamanhoTexto1)
    
    'Desintercala até o tamanho do menor texto
    For iIndice = 1 To iTamanhoMenor
        
        sTexto1 = sTexto1 & Mid(sTextoTotal, 2 * iIndice - 1, 1)
        sTexto2 = sTexto2 & Mid(sTextoTotal, 2 * iIndice, 1)
           
    Next
    
    'Atribui o restante do texto a strign maior
    If iTamanhoMenor = iTamanhoTexto1 Then
        sTexto2 = sTexto2 & Mid(sTextoTotal, (2 * iTamanhoMenor) + 1)
    Else
        sTexto1 = sTexto1 & Mid(sTextoTotal, (2 * iTamanhoMenor) + 1)
    End If

    Exit Sub

End Sub


Function HexaParaDec(vNumero As Variant, iTamanho As Integer) As String
'Faz a conversão de um HexaDecimal para um Decimal devolvendo o número
'formatado no tamanho passado

Dim sNumeroAux As String
Dim sAlgarismo As String
Dim iAlgarismo As String
Dim lNumero As Long
Dim iTamanhoIni As Integer
Dim iIndice As Integer

    sNumeroAux = CStr(Trim(vNumero))
    iTamanhoIni = Len(sNumeroAux)
    lNumero = 0
    
    For iIndice = 1 To iTamanhoIni
        
        sAlgarismo = Mid(sNumeroAux, iIndice, 1)
        
        Select Case Asc(sAlgarismo)
            
            Case Asc("A")
                iAlgarismo = 10
            
            Case Asc("B")
                iAlgarismo = 11
            
            Case Asc("C")
                iAlgarismo = 12
            
            Case Asc("D")
                iAlgarismo = 13
            
            Case Asc("E")
                iAlgarismo = 14
            
            Case Asc("F")
                iAlgarismo = 15
            
            Case Else
                iAlgarismo = CInt(sAlgarismo)
        
        End Select
        
        lNumero = lNumero + (iAlgarismo * (16 ^ (iTamanhoIni - iIndice)))
    
    Next
        
    HexaParaDec = FormataCpoNum(lNumero, iTamanho)
    
    Exit Function
    
End Function


Function DecParaBin(vNumero As Variant, iTamanho As Integer) As String
'Faz a conversão de um Decimal para um Binário devolvento o número
'formatado no tamanho passado

Dim iNumeroAux As String
Dim iIndice As Integer
Dim iResto As Integer
Dim iQuociente As Integer
Dim colRestos As New Collection
Dim sNumero As String

    iNumeroAux = vNumero
    
    iQuociente = iNumeroAux \ 2
    iResto = iNumeroAux Mod 2
    
    colRestos.Add iResto
    
    Do While iQuociente <> 0
        
        iResto = iQuociente Mod 2
        iQuociente = iQuociente \ 2

        colRestos.Add iResto

    Loop
    
    sNumero = ""
    
    For iIndice = colRestos.Count To 1 Step -1
        sNumero = sNumero & colRestos(iIndice)
    Next
        
    If iTamanho > 0 Then sNumero = FormataCpoNum(sNumero, iTamanho)
        
    DecParaBin = sNumero
    
    Exit Function
    
End Function

Function DecParaHexa(vNumero As Variant, iTamanho As Integer) As String
'Faz a conversão de um Decimal para um HexaDecimal devolvendo o número
'formatado no tamanho passado

Dim sNumero As String

    sNumero = CStr(Trim(vNumero))
    sNumero = Hex(sNumero)
            
    DecParaHexa = FormataCpoNum(sNumero, iTamanho)
    
    Exit Function
    
End Function

Function BinParaDec(vNumero As Variant, iTamanho As Integer) As String
'Faz a conversão de um Binário para um Decimal devolvento o número
'formatado no tamanho passado

Dim sNumeroAux As String
Dim sAlgarismo As String
Dim iAlgarismo As String
Dim lNumero As Long
Dim iIndice As Integer
Dim iTamanhoIni As Integer

    sNumeroAux = CStr(Trim(vNumero))
    iTamanhoIni = Len(sNumeroAux)
    lNumero = 0
    
    For iIndice = 1 To iTamanhoIni
        
        sAlgarismo = Mid(sNumeroAux, iIndice, 1)
        
        iAlgarismo = CInt(sAlgarismo)
        
        lNumero = lNumero + (iAlgarismo * (2 ^ (iTamanhoIni - iIndice)))
    
    Next
        
    BinParaDec = FormataCpoNum(lNumero, iTamanho)
    
    Exit Function
    
End Function


Private Sub Calcula_CheckSum(sTexto As String, sCheckSum As String)
'Calcula o conteúdo do campo CheckSum

Dim iTamanhoTexto As Integer
Dim iIndice As Integer
Dim iDigito As Integer
Dim iSoma As Integer
Dim iModulo As Integer
Dim iResto As Integer
        
    'Obtem o tamanho do texto
    iTamanhoTexto = Len(sTexto)
    
    iSoma = 0
    'Para cada digito no texto
    For iIndice = 1 To iTamanhoTexto
        'Transforma o dígito em decimal
        iDigito = HexaParaDec(Mid(sTexto, iIndice, 1), 2)
        
        iModulo = IIf((iIndice Mod 9) = 0, 9, iIndice Mod 9)
       
        iSoma = iSoma + (iDigito * iModulo)
       
    Next
    
    iResto = iSoma Mod 256
    
    'Passa o resultado para Hexadecimal
    sCheckSum = DecParaHexa(iResto, 2)
        
    Exit Sub
    
End Sub

Function Intercala_Texto(sTexto1 As String, sTexto2 As String) As String
'Intercala o sTexto1 e o sTexto2
'Ex. "1111", "0000", Res. = "10101010"
'    "11111", "000", REs. = "10101011"
'    "111", "00000", Res. = "10101000"

Dim sTextoFinal As String
Dim iTamanhoMenor As Integer
Dim iIndice As Integer
    
    'Veriifca o tamanho da menor string
    iTamanhoMenor = IIf(Len(sTexto1) > Len(sTexto2), Len(sTexto2), Len(sTexto1))

    'Intercala até a menor string acabar
    For iIndice = 1 To iTamanhoMenor
        sTextoFinal = sTextoFinal & Mid(sTexto1, iIndice, 1)
        sTextoFinal = sTextoFinal & Mid(sTexto2, iIndice, 1)
    Next
    
    'Intercala o que sobrou
    sTextoFinal = sTextoFinal & Mid(sTexto1, iTamanhoMenor + 1)
    sTextoFinal = sTextoFinal & Mid(sTexto2, iTamanhoMenor + 1)
    
    Intercala_Texto = sTextoFinal
    
    Exit Function

End Function


'Já existe (Criada por Raphael na CLassGeracaoArqIcms)
Private Function FormataCpoNum(vData As Variant, iTam As Integer) As String
'formata campo numerico alinhado-o à direita sem ponto e decimais, com zeros a esquerda

Dim iData As Integer
Dim sData As String

    If Len(vData) = iTam Then

        FormataCpoNum = vData
        Exit Function

    End If

    iData = iTam - Len(vData)
    
    If iData > 0 Then sData = String(iData, "0")

    FormataCpoNum = sData & vData

    Exit Function

End Function

Function Empresa_Le_NomeReduzido(objEmpresa As ClassDicEmpresa) As Long
'Lê o codigo da Empresa a partir do nome reduzido fornecido em objEmpresa
'Retorna os dados em objEmpresa

Dim lComando As Long
Dim lErro As Long
Dim lCodigo As Long
    
On Error GoTo Erro_Empresa_Le_NomeReduzido

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 6250
    
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Empresas WHERE NomeReduzido = ?", lCodigo, objEmpresa.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 92906
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92907
    
    If lErro = AD_SQL_SEM_DADOS Then gError 92908
   
    objEmpresa.lCodigo = lCodigo
    
    Call Comando_Fechar(lComando)
    
    Empresa_Le_NomeReduzido = SUCESSO

    Exit Function

Erro_Empresa_Le_NomeReduzido:

    Empresa_Le_NomeReduzido = gErr

    Select Case gErr
    
        Case 92906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 92907, 92908
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA3", gErr, objEmpresa.sNomeReduzido)
        
        Case 6289  'Será tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159113)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function Usuario_Testa_Atendente(ByVal sCodUsuario As String, Optional ByVal iFilialEmpresa As Integer) As Long
'Verifica se o Usuario passado é atendente na FilialEmpresa passada

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer
Dim iErro As Integer

On Error GoTo Erro_Usuario_Testa_Atendente
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 102785
        
    'Se foi passada uma filialempresa específica
    If iFilialEmpresa > 0 Then
        
        'Le o Usuario na Tabela de Atendentes filtrando por usuário e filialempresa
        lErro = Comando_Executar(lComando, "SELECT Codigo FROM Atendentes WHERE CodUsuario=? AND FilialEmpresa=?", iCodigo, sCodUsuario, iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 102786
        
        iErro = 1
    
    'Senão
    Else
    
        'Le o Usuario na Tabela de Atendentes filtrando apenas por usuário
        lErro = Comando_Executar(lComando, "SELECT Codigo FROM Atendentes WHERE CodUsuario=?", iCodigo, sCodUsuario)
        If lErro <> AD_SQL_SUCESSO Then gError 102789
    
        iErro = 2
        
    End If
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102787
    
    'se encontrou --> Erro o usuario é Atendente
    If lErro = AD_SQL_SUCESSO Then gError 102788
    
    Call Comando_Fechar(lComando)
    
    Usuario_Testa_Atendente = SUCESSO
    
    Exit Function
    
Erro_Usuario_Testa_Atendente:

    Usuario_Testa_Atendente = gErr
    
    Select Case gErr
        
        Case 102785
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 102786, 102787
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ATENDENTES2", gErr, glEmpresa, gsNomeEmpresa)
        
        Case 102788
            
            If iErro = 1 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_VINCULADO_ATENDENTE_FILIAL", gErr, sCodUsuario, iCodigo, glEmpresa, gsNomeEmpresa, giFilialEmpresa)
            Else
                Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_VINCULADO_ATENDENTE", gErr, sCodUsuario, glEmpresa, gsNomeEmpresa)
            End If
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159114)

    End Select
            
    Call Comando_Fechar(lComando)
            
    Exit Function
    
End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function OpcoesTelas_Le_NomeTela(ByVal colOpcoesTelas As Collection, ByVal sNomeTela As String) As Long
'Lê todas as opções para uma tela a partir do nome da tela

Dim lComandoDic As Long
Dim lErro As Long
Dim lCodigo As Long
Dim objOpcoesTelas As ClassOpcoesTelas
Dim tOpcoesTelas As typeOpcoesTelas

On Error GoTo Erro_OpcoesTelas_Le_NomeTela

    'Abre comando no dic
    lComandoDic = Comando_AbrirExt(GL_lConexaoDic)
    If lComandoDic = 0 Then gError 102887
    
    'Inicializa as strings que serão lidas no bd
    tOpcoesTelas.sOpcao = String(STRING_OPCOESTELAS_OPCAO, 0)
    tOpcoesTelas.sNomeTela = String(STRING_TELA_NOME, 0)
    
    'Lê as opções para a tela passada como parâmetros
    lErro = Comando_Executar(lComandoDic, "SELECT Codigo, Opcao, NomeTela, Padrao FROM OpcoesTelas WHERE NomeTela=? ORDER BY Opcao", tOpcoesTelas.lCodigo, tOpcoesTelas.sOpcao, tOpcoesTelas.sNomeTela, tOpcoesTelas.iPadrao, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then gError 102888
    
    lErro = Comando_BuscarPrimeiro(lComandoDic)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102889
    
    'Se não encontrou nenhuma opção => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102891
   
    'Enquantou houverem opções para a tela
    Do While lErro = AD_SQL_SUCESSO
    
        'Instancia um novo obj
        Set objOpcoesTelas = New ClassOpcoesTelas
        
        'Guarda os dados lidos no obj
        objOpcoesTelas.lCodigo = tOpcoesTelas.lCodigo
        objOpcoesTelas.sOpcao = tOpcoesTelas.sOpcao
        objOpcoesTelas.sNomeTela = tOpcoesTelas.sNomeTela
        objOpcoesTelas.iPadrao = tOpcoesTelas.iPadrao
        
        'Adiciona a opção à coleção
        colOpcoesTelas.Add objOpcoesTelas
        
        'Busca a próxima opção lida
        lErro = Comando_BuscarProximo(lComandoDic)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102890
    
    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComandoDic)
    
    OpcoesTelas_Le_NomeTela = SUCESSO

    Exit Function

Erro_OpcoesTelas_Le_NomeTela:

    OpcoesTelas_Le_NomeTela = gErr

    Select Case gErr
    
        Case 102887
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 102888 To 102890
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPCOESTELAS", gErr)
        
        Case 102891 'Não encontrado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159115)

    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComandoDic)
    
    Exit Function

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function Carrega_OpcoesTela(ByVal objTela As Object, Optional ByVal bSelecionaPadrao As Boolean = True) As Long
'Carrega a combo "OpcoesTela" da tela passada como parâmetro
'objTela é usado para obter o nome da tela em questão
'e para obter a combo a ser carregada
'ATENÇÃO: A COMBO OBRIGATORIAMENTE DEVE SE CHAMAR OpcoesTela

Dim lErro As Long
Dim colOpcoesTelas As New Collection
Dim objOpcaoTela As ClassOpcoesTelas
Dim objCombo As Object
Dim lOpcao As Long

On Error GoTo Erro_Carrega_OpcoesTela

    'Lê as opções gravadas para a tela em questão
    lErro = OpcoesTelas_Le_NomeTela(colOpcoesTelas, objTela.Name)
    If lErro <> SUCESSO And lErro <> 102891 Then gError 102892
    
    'Faz objCombo apontar para a combo que será carregada
    Set objCombo = objTela.Controls("OpcoesTela")
    
    'Para cada opção encontrada
    For Each objOpcaoTela In colOpcoesTelas
    
        'joga na combo o nome da opção
        objCombo.AddItem objOpcaoTela.sOpcao
        
        'joga o código da opção no itemdata
        objCombo.ItemData(objCombo.NewIndex) = objOpcaoTela.lCodigo
        
        'Se é para selecionar a opção padrão e a opção carregada é a padrão
        If bSelecionaPadrao And CallByName(objOpcaoTela, "iPadrao", VbGet) = MARCADO Then
        
            'Seleciona a opção
            objCombo.ListIndex = objCombo.NewIndex
            
            'Marca o campo padrão
            objTela.Controls("OpcaoPadrao").Value = MARCADO
        
        End If
    
    Next
    
    Carrega_OpcoesTela = SUCESSO

    Exit Function

Erro_Carrega_OpcoesTela:

    Carrega_OpcoesTela = gErr

    Select Case gErr

        Case 102892, 102893, 102899
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159116)

    End Select

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function OpcoesTelasCampos_Le_NomeTela(ByVal colOpcoesTelasCampos As Collection, ByVal sNomeTela As String) As Long

Dim lErro As Long
Dim lComandoDic As Long
Dim sNomeControle As String
Dim iExecutaValidate As Integer
Dim objOpcoesTelasCampos As ClassOpcoesTelasCampos

On Error GoTo Erro_OpcoesTelasCampos_Le_NomeTela

    'Abre comando no dic
    lComandoDic = Comando_AbrirExt(GL_lConexaoDic)
    If lComandoDic = 0 Then gError 102900
    
    'Esse código foi colocado porque a mesma tela tratará a exportação de diversos
    'documentos e cada um deles terá suas próprias opções, mas como a tela é a mesma
    'em OpcoesTelasCampos será gravado com o nome genérico que vale para todos
    If left(sNomeTela, 14) = "ExportaParaDoc" Then sNomeTela = "ExportaParaDoc"

    'Inicializa a string que será lida do BD
    sNomeControle = String(STRING_NOME_CONTROLE, 0)
    
    'Lê no bd os controles da tela para os quais o usuário pode armazenar opções
    lErro = Comando_Executar(lComandoDic, "SELECT NomeControle, ExecutaValidate FROM OpcoesTelasCampos WHERE NomeTela=? ORDER BY NomeControle", sNomeControle, iExecutaValidate, sNomeTela)
    If lErro <> AD_SQL_SUCESSO Then gError 102901
    
    lErro = Comando_BuscarPrimeiro(lComandoDic)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102902
    
    'Se não encontrou nenhum controle
    If lErro = AD_SQL_SEM_DADOS Then gError 102904
    
    'Enquanto houverem campos para a tela em questão
    Do While lErro = AD_SQL_SUCESSO
    
        'Instancia um novo obj
        Set objOpcoesTelasCampos = New ClassOpcoesTelasCampos
        
        'Guarda os dados lidos no obj
        objOpcoesTelasCampos.sNomeControle = sNomeControle
        objOpcoesTelasCampos.sNomeTela = sNomeTela
        objOpcoesTelasCampos.iExecutaValidate = iExecutaValidate
        
        'adiciona o obj à coleção
        colOpcoesTelasCampos.Add objOpcoesTelasCampos
        
        'busca o próximo nome de controle lido do bd
        lErro = Comando_BuscarProximo(lComandoDic)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102903
    
    Loop
    
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComandoDic)

    OpcoesTelasCampos_Le_NomeTela = SUCESSO

    Exit Function

Erro_OpcoesTelasCampos_Le_NomeTela:

    OpcoesTelasCampos_Le_NomeTela = gErr

    Select Case gErr

        Case 102900
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 102901 To 102903
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPCOESTELASCAMPOS", gErr)
        
        Case 102904 'Não encontrado
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159117)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComandoDic)

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function OpcoesTela_Click(ByVal objTela As Object) As Long

Dim lErro As Long
Dim objCombo As Object
Dim iCodigo As Integer

On Error GoTo Erro_OpcoesTela_Click

    'Instancia a combo OpcoesTela
    Set objCombo = objTela.Controls("OpcoesTela")
    
    'Verifica se foi preenchida a ComboBox
    If Len(Trim(objCombo.Text)) = 0 Then Exit Function

    'Indica que a tela não deve ser atualizada. Isso evita que a função combo_seleciona dispare novamente o click
    Call CallByName(objTela, "iAtualizaTela", VbLet, 0)
    
    'Verifica se existe o item na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(objCombo, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 102931
    
    'Se encontrou a opção cadastrada
    If lErro = SUCESSO Then
    
        'Carrega os dados para a tela
        lErro = Traz_OpcoesTelasValores_Tela(objTela)
        If lErro <> SUCESSO Then gError 102941
    
    End If
    
    'Indica que a tela deve ser atualizada a função OpcoesTela_Click for chamada, pois a função combo_seleciona já foi executada
    Call CallByName(objTela, "iAtualizaTela", VbLet, 1)
    
    OpcoesTela_Click = SUCESSO
    
    Exit Function

Erro_OpcoesTela_Click:

    OpcoesTela_Click = gErr

    Select Case gErr

        Case 102931, 102941
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159118)

    End Select

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function Traz_OpcoesTelasValores_Tela(objTela As Object) As Long

Dim lErro As Long
Dim objOpcaoTela As New ClassOpcoesTelas
Dim objOpcaoTelaValores As New ClassOpcoesTelasValores
Dim objControle As Object
Dim colOpcaoTelaCampos As New Collection
Dim objOpcaoTelaCampos As ClassOpcoesTelasCampos
Dim iItemData As Integer
Dim lItemData As Long

On Error GoTo Erro_Traz_OpcoesTelasValores_Tela

    'Guarda o código da opção que terá os valores lidos
    objOpcaoTela.lCodigo = objTela.Controls("OpcoesTela").ItemData(objTela.Controls("OpcoesTela").ListIndex)
        
    'Lê a opção para verificar se ela é padrão
    lErro = OpcoesTelas_Le(objOpcaoTela)
    If lErro <> SUCESSO And lErro <> 102968 Then gError 102964
    
    'Se não encontrou
    If lErro = 102968 Then gError 102969
    
    'Exibe na tela o valor do campo padrão
    objTela.Controls("OpcaoPadrao").Value = objOpcaoTela.iPadrao
    
    'Lê os valores que serão trazidos para a tela de acordo com a opção selecionada
    lErro = OpcoesTelasValores_Le_CodOpcao(objOpcaoTela)
    If lErro <> SUCESSO And lErro <> 102938 Then gError 102939
    
    'Se não encontrou os valores para a opção em questão => erro
    If lErro = 102938 Then gError 102940
    
    'Lê os campos correspondentes aos valores da tela
    lErro = OpcoesTelasCampos_Le_NomeTela(colOpcaoTelaCampos, objTela.Name)
    If lErro <> SUCESSO And lErro <> 102904 Then gError 102943
    
    'Se não encontrou os campos para a opção em questão => erro
    If lErro = 102904 Then gError 102940
    
    'Para cada valor encontrado
    For Each objOpcaoTelaValores In objOpcaoTela.colOpcoesTelasValores
    
        'Instancia o controle que receberá o valor
        Set objControle = objTela.Controls(objOpcaoTelaValores.sNomeControle)
        
        If TypeName(objControle) = "ComboBox" Then
            If objControle.Style = 2 Then
                If objOpcaoTelaValores.sValorCampo = "" Then
                    objControle.ListIndex = -1
                Else
                    Call Combo_Obtem_ItemData(objControle, objOpcaoTelaValores.sValorCampo, lItemData)
                    Call Combo_Seleciona_ItemData(objControle, lItemData)
                End If
            Else
                'Joga o valor no controle
                objControle = objOpcaoTelaValores.sValorCampo
            End If
        Else
            
            If TypeName(objControle) = "MaskEdBox" Then
                objControle.promptinclude = False
            End If
        
            'Joga o valor no controle
            objControle = objOpcaoTelaValores.sValorCampo
            
            If TypeName(objControle) = "MaskEdBox" Then
                objControle.promptinclude = True
            End If
            
        End If
        
        'Para cada controle da tela cadastrado no BD
        For Each objOpcaoTelaCampos In colOpcaoTelaCampos
            
            'Se o nome do controle na coleção de campos é igual ao nome do controle na coleção de valores => sai da função
            'Isso é feito para encontrar o campo correspondente ao valor que está sendo exibido na tela
            'É preciso encontrar o campo para verificar se ele deve ser validado
            If objOpcaoTelaCampos.sNomeControle = objOpcaoTelaValores.sNomeControle Then Exit For
        
        Next
        
        'Se é para executar o validate do campo
        If objOpcaoTelaCampos.iExecutaValidate = 1 Then
        
            'Executa o validate do controle
            Call CallByName(objTela, objOpcaoTelaValores.sNomeControle & "_Validate", VbMethod, bSGECancelDummy)
        
        End If
    
    Next

    Traz_OpcoesTelasValores_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_OpcoesTelasValores_Tela:

    Traz_OpcoesTelasValores_Tela = gErr
    
    Select Case gErr

        Case 102939, 102942, 102943, 102964
        
        Case 102969
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCOESTELAS_NAO_ENCONTRADA", gErr, objOpcaoTela.lCodigo, objOpcaoTela.sNomeTela)

        Case 102940
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCOESTELASVALORES_NAO_ENCONTRADO", gErr, objTela.Controls("OpcoesTela").Text, objTela.Name)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159119)

    End Select
    
End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function OpcoesTelasValores_Le_CodOpcao(ByVal objOpcaoTela As ClassOpcoesTelas) As Long

Dim lErro As Long
Dim lComandoDic As Long
Dim tOpcoesTelasValores As typeOpcoesTelasValores
Dim objOpcoesTelasValores As ClassOpcoesTelasValores

On Error GoTo Erro_OpcoesTelasValores_Le_CodOpcao

    'Abre comando no dic
    lComandoDic = Comando_AbrirExt(GL_lConexaoDic)
    If lComandoDic = 0 Then gError 102934
    
    'Inicializa as strings
    tOpcoesTelasValores.sNomeControle = String(STRING_NOME_CONTROLE, 0)
    tOpcoesTelasValores.sValorCampo = String(STRING_BUFFER_MAX_TEXTO, 0)
    
    'Lê os valores para opção passada como parâmetro
    lErro = Comando_Executar(lComandoDic, "SELECT NomeControle, ValorCampo FROM OpcoesTelasValores WHERE CodOpcao=? ORDER BY ValorCampo", tOpcoesTelasValores.sNomeControle, tOpcoesTelasValores.sValorCampo, objOpcaoTela.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 102935
    
    lErro = Comando_BuscarPrimeiro(lComandoDic)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102936
    
    'Se não encontrou nenhum valor => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102938
    
    'Enquanto houverem valores
    Do While lErro = AD_SQL_SUCESSO
    
        'Instancia um novo obj
        Set objOpcoesTelasValores = New ClassOpcoesTelasValores
        
        'Guarda os dados lidos no obj
        objOpcoesTelasValores.sNomeControle = tOpcoesTelasValores.sNomeControle
        objOpcoesTelasValores.sValorCampo = tOpcoesTelasValores.sValorCampo
        
        'Guarda o valor na coleção
        objOpcaoTela.colOpcoesTelasValores.Add objOpcoesTelasValores
        
        'Busca o próximo valor encontrado
        lErro = Comando_BuscarProximo(lComandoDic)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102937
    
    Loop
    
    'Fecha comando
    Call Comando_Fechar(lComandoDic)
    
    OpcoesTelasValores_Le_CodOpcao = SUCESSO
    
    Exit Function
    
Erro_OpcoesTelasValores_Le_CodOpcao:

    OpcoesTelasValores_Le_CodOpcao = gErr
    
    Select Case gErr

        Case 102934
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 102935 To 102937
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPCOESTELASVALORES", gErr)
        
        Case 102938 'Não encontrado
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159120)

    End Select
    
    'Fecha comando
    Call Comando_Fechar(lComandoDic)
    
    
End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function OpcoesTelas_Le(ByVal objOpcaoTela As ClassOpcoesTelas) As Long
'Lê a opção com código passado como parâmetro

Dim lErro As Long
Dim lComandoDic As Long
Dim tOpcoesTelas As typeOpcoesTelas

On Error GoTo Erro_OpcoesTelas_Le

    'Abre comando no dic
    lComandoDic = Comando_AbrirExt(GL_lConexaoDic)
    If lComandoDic = 0 Then gError 102965
    
    'Inicializa as strings que serão lidas no bd
    tOpcoesTelas.sOpcao = String(STRING_OPCOESTELAS_OPCAO, 0)
    tOpcoesTelas.sNomeTela = String(STRING_TELA_NOME, 0)
    
    'Lê aopção com código passado como parâmetro
    lErro = Comando_Executar(lComandoDic, "SELECT Codigo, Opcao, NomeTela, Padrao FROM OpcoesTelas WHERE Codigo=?", tOpcoesTelas.lCodigo, tOpcoesTelas.sOpcao, tOpcoesTelas.sNomeTela, tOpcoesTelas.iPadrao, objOpcaoTela.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 102966
    
    lErro = Comando_BuscarPrimeiro(lComandoDic)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102967
    
    'Se não encontrou nenhuma opção => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102968
    
    'Guarda no obj os dados lidos
    objOpcaoTela.lCodigo = tOpcoesTelas.lCodigo
    objOpcaoTela.sOpcao = tOpcoesTelas.sOpcao
    objOpcaoTela.sNomeTela = tOpcoesTelas.sNomeTela
    objOpcaoTela.iPadrao = tOpcoesTelas.iPadrao

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComandoDic)

    OpcoesTelas_Le = SUCESSO

    Exit Function

Erro_OpcoesTelas_Le:

    OpcoesTelas_Le = gErr

    Select Case gErr

        Case 102965
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 102966, 102967
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPCOESTELAS", gErr)
        
        Case 102968 'Não encontrado

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159121)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComandoDic)

End Function

Function Usuario_Le_Login(ByVal sNomeLogin As String, sComputador As String, ByVal objUsuario As ClassDicUsuario) As Long

Dim lComando As Long
Dim lErro As Long
Dim tUsuarios As typeDicUsuario
    
On Error GoTo Erro_Usuario_Le_Login

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 187465

    tUsuarios.sCodUsuario = String(STRING_SIGLA_USUARIO, 0)
    tUsuarios.sNome = String(STRING_USUARIO, 0)
    tUsuarios.sCodGrupo = String(STRING_GRUPO_CODIGO, 0)
    sComputador = String(255, 0)

    lErro = Comando_Executar(lComando, "SELECT CodUsuario, Nome, CodGrupo, DataValidade, Computador FROM Usuarios WHERE NomeLogin = ? ", tUsuarios.sCodUsuario, tUsuarios.sNome, tUsuarios.sCodGrupo, tUsuarios.dtDataValidade, sComputador, sNomeLogin)
    If lErro <> AD_SQL_SUCESSO Then gError 187466
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 187467
    
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_LEITURA_SEM_DADOS
        
    objUsuario.sCodUsuario = tUsuarios.sCodUsuario
    objUsuario.sCodGrupo = tUsuarios.sCodGrupo
    objUsuario.sNome = tUsuarios.sNome
    objUsuario.dtDataValidade = tUsuarios.dtDataValidade
            
    Comando_Fechar (lComando)
    
    Usuario_Le_Login = SUCESSO
    
    Exit Function
    
Erro_Usuario_Le_Login:

    Usuario_Le_Login = gErr

    Select Case gErr
    
        Case 187465
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 187466, 187467
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOS_DIC", gErr)
            
        Case ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187468)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Empresa_Le_DBOlap(sDBOlap As String) As Long

Dim lErro As Long
Dim lComando As Long
    
On Error GoTo Erro_Empresa_Le_DBOlap

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 192420
    
    sDBOlap = String(STRING_EMPRESA_DBOLAP, 0)
    
    lErro = Comando_Executar(lComando, "SELECT DBOlap FROM Empresas WHERE Codigo = ? And Inativa=0", sDBOlap, glEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 192421
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    If lErro <> AD_SQL_SUCESSO Then gError 192422
    
    Call Comando_Fechar(lComando)
    
    Empresa_Le_DBOlap = SUCESSO

    Exit Function

Erro_Empresa_Le_DBOlap:

    Empresa_Le_DBOlap = gErr

    Select Case gErr
    
        Case 192420
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 192421, 192422
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA", gErr, glEmpresa)
        
        Case ERRO_LEITURA_SEM_DADOS  'Será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192423)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Empresa_Le_StringConexao(sStringConexao As String) As Long

Dim lErro As Long
Dim lComando As Long
    
On Error GoTo Erro_Empresa_Le_StringConexao

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 192420
    
    sStringConexao = String(STRING_STRING_CONEXAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT StringConexao FROM Empresas WHERE Codigo = ? And Inativa=0", sStringConexao, glEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 192421
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then gError ERRO_LEITURA_SEM_DADOS
    If lErro <> AD_SQL_SUCESSO Then gError 192422
    
    Call Comando_Fechar(lComando)
    
    Empresa_Le_StringConexao = SUCESSO

    Exit Function

Erro_Empresa_Le_StringConexao:

    Empresa_Le_StringConexao = gErr

    Select Case gErr
    
        Case 192420
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 192421, 192422
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA", gErr, glEmpresa)
        
        Case ERRO_LEITURA_SEM_DADOS  'Será tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192423)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

