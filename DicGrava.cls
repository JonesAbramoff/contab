VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DicGrava"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const RESOURCETYPE_ANY = &H0
Private Const RESOURCE_CONNECTED = &H1
Private Const DRIVE_REMOTE = 4

Private Declare Function CriarDSNAccess Lib "ADINSTAL.DLL" (ByVal sNomeDSN As String, ByVal sNomeMDB As String) As Long

Private Declare Function SQLConfigDataSource Lib "ODBCCP32.dll" (ByVal hWnd As Long, ByVal iTipo As Integer, ByVal sDriver As String, ByVal sAtributos As String) As Long
Private Declare Function SQLAllocEnv Lib "odbc32.dll" (henv As Long) As Integer
Private Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal henv As Long, hdbc As Long) As Integer
Private Declare Function SQLDriverConnect Lib "odbc32.dll" (ByVal hdbc As Long, ByVal hWnd As Long, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal ach_conn_out As String, ByVal maxtamstrout As Integer, cb_conn_out As Integer, ByVal drivercompl As Integer) As Integer
Private Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hdbc As Long) As Integer
Private Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc As Long) As Integer
Private Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal henv As Long) As Integer
Private Declare Function SQLInstallerError Lib "ODBCCP32.dll" (ByVal iErro As Integer, pfErrorCode As Long, ByVal lpszErrorMsg As String, ByVal cbErrorMsgMax As Integer, pcbErrorMsg As Integer) As Integer
Private Declare Function SQLError Lib "odbc32.dll" (ByVal henv As Long, ByVal hdbc As Long, ByVal hstmt As Long, ByVal ssqlstate As String, iNative As Integer, ByVal sMsg As String, ByVal BufferLength As Integer, TextLength As Integer) As Integer

Private Const ODBC_ADD_DSN = 1
Private Const ODBC_REMOVE_DSN = 3
Private Const ODBC_ADD_SYS_DSN = 4

Private Const SQL_DRIVER_NOPROMPT = 0

'******************** Code Start **************************
' This code was originally written by Dev Ashish.
' It is not to be altered or distributed,
' except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code Courtesy of
' Dev Ashish
'
Private Declare Function apiGetComputerName Lib "kernel32" Alias _
    "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function apiGetUserName Lib "advapi32.dll" Alias _
    "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Function fOSMachineName() As String
'Returns the computername
Dim lngLen As Long, lngX As Long
Dim strCompName As String
    lngLen = 16
    strCompName = String$(lngLen, 0)
    lngX = apiGetComputerName(strCompName, lngLen)
    If lngX <> 0 Then
        fOSMachineName = left$(strCompName, lngLen)
    Else
        fOSMachineName = ""
    End If
End Function
'******************** Code End **************************

'******************** Code Start **************************
' This code was originally written by Dev Ashish.
' It is not to be altered or distributed,
' except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code Courtesy of
' Dev Ashish
'
Function fOSUserName() As String
' Returns the network login name
Dim lngLen As Long, lngX As Long
Dim strUserName As String
    strUserName = String$(254, 0)
    lngLen = 255
    lngX = apiGetUserName(strUserName, lngLen)
    If (lngX > 0) Then
        fOSUserName = left$(strUserName, lngLen - 1)
    Else
        fOSUserName = vbNullString
    End If
End Function
'******************** Code End **************************

Function Empresa_Exclui(ByVal lCodigo As Long, ByVal sNome As String) As Long
'Exclui Empresa cujo código é lCodigo do BD.

Dim lErro As Long
Dim alComando(0 To 5) As Long
Dim lTransacao As Long
Dim lCodigo1 As Long
Dim lConexao As Long
Dim iIndice As Integer
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_Empresa_Exclui

    lConexao = GL_lConexaoDic
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(lConexao)
        If alComando(iIndice) = 0 Then Error 6212
    Next
    
    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 6214

    'pesquisa a Empresa em questão no BD
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Empresas WHERE Codigo=?", 0, lCodigo1, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 6215

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro = AD_SQL_SEM_DADOS Then Error 6216
    If lErro <> AD_SQL_SUCESSO Then Error 6217

    'lock da Empresa (usuários terão um lock shared na Empresa ao usarem-na)
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 6218

    'exclui "logicamente" a Empresa do BD
    lErro = Comando_ExecutarPos(alComando(1), "UPDATE Empresas SET Inativa = ?", alComando(0), Inativo)
    If lErro <> AD_SQL_SUCESSO Then Error 6220

    'seleciona as filial empresa ligadas a epresa excluída
    lErro = Comando_ExecutarPos(alComando(2), "SELECT CodFilial FROM FiliaisEmpresas WHERE CodEmpresa = ?", 0, lCodigo1, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 6213
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26566
    
    'se existirem filiais daquela empresa
'    If lErro = AD_SQL_SUCESSO And giTipoVersao = VERSAO_FULL Then
        'Pede confirmação para exclusão das filiais
        vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_FILIAIS", lCodigo, sNome)
        If vbMsgRet = vbNo Then Error 57501
'    Else
'        'Pede confirmação para exclusão ao usuário
'        vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_EMPRESA", lCodigo, sNome)
'        If vbMsgRet = vbNo Then Error 57502
'    End If
    
    'Para cada filial empresa encontrada
    Do While lErro <> AD_SQL_SEM_DADOS
        'Faz o lock exclusivo
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then Error 30465
        
        'exclui "logicamente" a filial empresa do BD
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE FiliaisEmpresas SET Inativa =?", alComando(2), Inativo)
        If lErro <> AD_SQL_SUCESSO Then Error 30472
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30489
    Loop
    
    'Busca os usuários vinculados a empresa
    lErro = Comando_ExecutarPos(alComando(4), "SELECT CodEmpresa FROM UsuFilEmp WHERE CodEmpresa = ?", 0, lCodigo1, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 26510
    
    lErro = Comando_BuscarPrimeiro(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26628
    
    'Para cada vinculo usuário-empresa encontrado
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'Faz o "lock" exclusivo
        lErro = Comando_LockExclusive(alComando(4))
        If lErro <> SUCESSO Then Error 26629
        
        'Exclui o registro de UsuFilEmp
        lErro = Comando_ExecutarPos(alComando(5), "DELETE FROM UsuFilEmp", alComando(4))
        If lErro <> AD_SQL_SUCESSO Then Error 26630
        
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26522
    
    Loop
    
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 6221
    
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    Empresa_Exclui = SUCESSO

    Exit Function

Erro_Empresa_Exclui:

    Empresa_Exclui = Err

    Select Case Err

        Case 6212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6213, 26566, 30489
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAIS_EMPRESAS", Err, lCodigo)

        Case 6214
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 6215, 6217
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA", Err, lCodigo)

        Case 6216
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_NAO_CADASTRADA", Err, lCodigo)

        Case 6218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EMPRESA", Err, lCodigo)

        Case 6220
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_EMPRESA", Err, lCodigo)

        Case 6221
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 26510, 26522, 26628
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUFILEMP", Err, lCodigo)
        
        Case 26629
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_USUFILEMP", Err, lCodigo)
        
        Case 26630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_USUFILEMP", Err, lCodigo)
        
        Case 30465
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAL_EMPRESA", Err, lCodigo1, lCodigo)
        
        Case 30472
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILIAL_EMPRESA", Err, lCodigo1, lCodigo)
        
        Case 57501, 57502
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158987)

    End Select

    lErro = Transacao_RollbackDic
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function Verificar_Integridade_Browses() As Long
'Insere no arquivo "CORRECAO_BROWSE.txt" os erros do browse

Dim lErro As Long
Dim colBrowseUsuarioCampo As New Collection
Dim tCampo As typeCampos
Dim objUsuario As ClassUsuario
Dim objCampo As ClassTelasCampos

On Error GoTo Erro_Verificar_Integridade_Browses

    lErro = Verifica_Sequencial(colBrowseUsuarioCampo)
    If lErro <> SUCESSO Then gError 110131
    
    Open "CORRECAO_BROWSE1.txt" For Output As #1
    
    Print #1, "NOME DA TELA - USUÁRIO"
    
    For Each objUsuario In colBrowseUsuarioCampo
    
        Print #1, objUsuario.sNome & " - " & objUsuario.sCodUsuario
    
    Next
    
    Close
    
    Set colBrowseUsuarioCampo = New Collection
    
    lErro = Identifica_Campos_Invalidos(colBrowseUsuarioCampo)
    If lErro <> SUCESSO Then gError 110132

    Open "CORRECAO_BROWSE2.txt" For Output As #1
    
    Print #1, "TABELAS"
    
    For Each objCampo In colBrowseUsuarioCampo
    
        Print #1, objCampo.sNomeTela
    
    Next
        
    Close

    Verificar_Integridade_Browses = SUCESSO

    Exit Function

Erro_Verificar_Integridade_Browses:

    Verificar_Integridade_Browses = gErr

    Select Case gErr
    
        Case 110131, 110132

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158988)

    End Select
    
    Exit Function
    
End Function
Public Function Identifica_Campos_Invalidos(colBrowseUsuarioCampo As Collection) As Long

Dim lErro As Long
Dim lCount As Long
Dim sWhere As String
Dim sSelect As String
Dim sTabelaAux As String
Dim tCampo As typeCampos
Dim lComandoDic As Long
Dim lComandoSge As Long
Dim lComando As Long
Dim iBanco As Integer
Dim objCampos As ClassTelasCampos

On Error GoTo Erro_Identifica_Campos_Invalidos
    
    'Abre comando
    lComandoDic = Comando_AbrirExt(GL_lConexaoDic)
    If lComandoDic = 0 Then Error 110137
    
    lComandoSge = Comando_Abrir()
    If lComandoSge = 0 Then Error 110141
    
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 110142
    
    'Aloca espaço nas strings
    tCampo.sNomeArq = String(STRING_NOME_ARQUIVO, 0)
    tCampo.sNome = String(STRING_NOME_CAMPO, 0)
    
    'Lê Tabela de Campos (Tabela, Campo, Banco De Dados)
    lErro = Comando_Executar(lComando, "SELECT A.NomeArq, A.Nome, MAX (B.BancoDados) FROM campos A, browsearquivo B " & _
                                       "WHERE  A.NomeArq = B.NomeArq GROUP BY A.NomeArq, A.Nome ORDER By A.NomeArq,A.Nome ", tCampo.sNomeArq, tCampo.sNome, iBanco)
    If lErro <> AD_SQL_SUCESSO Then gError 110138
    
    'Busca primeira tupla
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 110139
    
    sTabelaAux = tCampo.sNomeArq
    sWhere = " 1 = 0 "
        
    Do While lErro <> AD_SQL_SEM_DADOS
        
        'Verifica se mudou de tabela
        If sTabelaAux <> tCampo.sNomeArq Then
        
            'Select que verifica erros de referência browser/Tabela
            sSelect = "SELECT TOP 1 1 FROM " & sTabelaAux & " WHERE " & sWhere

            'Verifica banco que deve ser usado
            If iBanco = BD_DICDADOS Then
                lErro = Comando_Executar(lComandoDic, sSelect, lCount)
            Else
                lErro = Comando_Executar(lComandoSge, sSelect, lCount)
            End If
            
            'Verifica se houve erro
            If lErro <> AD_SQL_SUCESSO Then
            
                Set objCampos = New ClassTelasCampos
            
                objCampos.sNomeTela = sTabelaAux
            
               'Adiciona a coleção de tabelas erradas
                colBrowseUsuarioCampo.Add objCampos
                
            End If
            
            sTabelaAux = tCampo.sNomeArq
            sWhere = " 1 = 0 AND " & tCampo.sNome & " = " & tCampo.sNome
        Else
        
            'Condição do select
            sWhere = sWhere & " AND " & tCampo.sNome & " = " & tCampo.sNome
        
        End If
        
       
        'Busca próxima tupla
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 110139
          
    Loop
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComandoDic)
    Call Comando_Fechar(lComandoSge)
    
    Identifica_Campos_Invalidos = SUCESSO

    Exit Function

Erro_Identifica_Campos_Invalidos:

    Identifica_Campos_Invalidos = gErr

    Select Case gErr
    
        Case 110137, 110141, 110142
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 110138, 110139, 110140
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEUSUARIOCAMPO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158989)

    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComandoDic)
    Call Comando_Fechar(lComandoSge)
   
    Exit Function
    
End Function

Public Function Verifica_Sequencial(colBrowseUsuarioCampo As Collection) As Long
'Retorna uma coleção com os browses com sequencial errado

Dim lErro As Long
Dim sCodUsuario As String
Dim lComando As Long
Dim sNomeTela As String
Dim tBrowseUsuarioCampo As typeBrowseUsuarioCampo
Dim objUsuario As ClassUsuario


On Error GoTo Erro_Verifica_Sequencial

    'Abre comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 110133
    
    'Aloca espaço nas strings
    tBrowseUsuarioCampo.sNomeTela = String(255, 0)
    tBrowseUsuarioCampo.sCodUsuario = String(255, 0)
    
    'Seleciona conjuntos nometela-codusuario NAO possuem os campos
    'com posicao sequencial de 1 a n, sem "pular" nem repetir posicoes.
    lErro = Comando_Executar(lComando, "SELECT NomeTela, Codusuario  FROM browseusuariocampo GROUP BY NomeTela, Codusuario " & _
                                       "HAVING MAX(Posicaotela) - MIN(Posicaotela) + 1 <> Count(distinct(PosicaoTela)) OR " & _
                                       "MAX (PosicaoTela) - MIN(PosicaoTela) + 1 <> COUNT(PosicaoTela) " & _
                                       "ORDER BY NomeTela, codusuario", tBrowseUsuarioCampo.sNomeTela, tBrowseUsuarioCampo.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 110134

    'Busca Primeira tupla
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 110135
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objUsuario = New ClassUsuario
    
        objUsuario.sCodUsuario = tBrowseUsuarioCampo.sCodUsuario
        objUsuario.sNome = tBrowseUsuarioCampo.sNomeTela
    
        'Adiciona a coleção
        colBrowseUsuarioCampo.Add objUsuario
        
        'Busca próxima tupla
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 110136
      
    Loop
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    Verifica_Sequencial = SUCESSO

    Exit Function

Erro_Verifica_Sequencial:

    Verifica_Sequencial = gErr

    Select Case gErr

        Case 110133
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 110134, 110135, 110136
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BROWSEUSUARIOCAMPO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158990)

    End Select
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'''Exclui registros da Tabela UsuEmpGrupo com código do Usuário sCodUsuario
'''Usada na Transação de Exclusão de Usuário
''Function UsuEmpGrupo_Exclui(ByVal sCodUsuario As String) As Long
''
''Dim lConexao As Long
''Dim lComando As Long
''Dim lErro As Long
''
''On Error GoTo Erro_UsuEmpGrupo_Exclui
''
''    lConexao = GL_lConexaoDic
''    lComando = Comando_AbrirExt(lConexao)
''    If lComando = 0 Then Error 6293
''
''    'exclui os registros de UsuEmpGrupo que tem código do Usuário
''    lErro = Comando_Executar(lComando, "DELETE FROM UsuEmpGrupo WHERE CodUsuario=?", sCodUsuario)
''    If lErro <> AD_SQL_SUCESSO Then Error 6294
''
''    lErro = Comando_Fechar(lComando)
''
''    UsuEmpGrupo_Exclui = SUCESSO
''
''    Exit Function
''
''Erro_UsuEmpGrupo_Exclui:
''
''    UsuEmpGrupo_Exclui = Err
''
''    Select Case Err
''
''        Case 6293
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 6294
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_USUEMPGRUPO", Err, sCodUsuario)
''            lErro = Comando_Fechar(lComando)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158991)
''
''    End Select
''
''    Exit Function
''
''End Function

Function Usuario_Exclui(ByVal sCodUsuario As String) As Long

Dim lConexao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sCodUsuario1 As String

On Error GoTo Erro_Usuario_Exclui

    lConexao = GL_lConexaoDic
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 6276
    
    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 6277
    
    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 6278
    
    lErro = Usuario_Exclui_Batch(sCodUsuario, lComando, lComando1)
    If lErro <> SUCESSO Then Error 32051
    
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 6283

    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)

    Usuario_Exclui = SUCESSO
    
    Exit Function
    
Erro_Usuario_Exclui:

    Usuario_Exclui = Err
    
    Select Case Err
    
        Case 6276
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 6277
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 6278
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 6283
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case 32051
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158992)
    
    End Select
    
    lErro = Transacao_RollbackDic
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)
    
    Exit Function

End Function

'''Grava relação Usuário-Grupo na Tabela UsuEmpGrupo
'''Usado na transação de gravar Usuário
''Function UsuEmpGrupo_Grava(ByVal sCodGrupo As String, ByVal sCodUsuario As String) As Long
''
''Dim lConexao As Long
''Dim lComando As Long
''Dim lErro As Long
''
''On Error GoTo Erro_UsuEmpGrupo_Grava
''
''    lConexao = GL_lConexaoDic
''
''    lComando = Comando_AbrirExt(lConexao)
''    If lComando = 0 Then Error 6284
''
''    'insere registro que liga usuário a sCodGrupo
''    lErro = Comando_Executar(lComando, "INSERT INTO UsuEmpGrupo (CodUsuario, CodEmpresa, CodGrupo) VALUES (?,?,?)", sCodUsuario, glEmpresaAtual, sCodGrupo)
''    If lErro <> AD_SQL_SUCESSO Then Error 6285
''
''    lErro = Comando_Fechar(lComando)
''
''    UsuEmpGrupo_Grava = SUCESSO
''
''    Exit Function
''
''Erro_UsuEmpGrupo_Grava:
''
''    UsuEmpGrupo_Grava = Err
''
''    Select Case Err
''
''        Case 6284
''           lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 6285
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_USUEMPGRUPO", Err, sCodUsuario, glEmpresaAtual, sCodGrupo)
''            lErro = Comando_Fechar(lComando)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158993)
''
''    End Select
''
''    Exit Function
''
''End Function
'''Altera (ou grava se for replicação de Usuário) a relação
'''Usuário-Grupo na Tabela UsuEmpGrupo se for necessário
'''retorna MODIFICACAO em iOperacao
'''se for replicacao retorna GRAVACAO em iOperacao
'''Usado na transação de gravar Usuário
''Function UsuEmpGrupo_Altera(ByVal sCodGrupo As String, ByVal sCodUsuario As String, iOperacao As Integer) As Long
''
''Dim objSelect As New DicSelect  'Para usar função de leitura
''Dim lConexao As Long
''Dim lComando As Long
''Dim lComando1 As Long
''Dim lErro As Long
''Dim colEmpresaGrupo As New Collection
''Dim objEmpresaGrupo As ClassDicEmpresaGrupo
''Dim bCoincidencia As Boolean
''Dim sCodUsuario1 As String
''
''On Error GoTo Erro_UsuEmpGrupo_Altera
''
''    lConexao = GL_lConexaoDic
''
''    lComando = Comando_AbrirExt(lConexao)
''    If lComando = 0 Then Error 6263
''
''    lComando1 = Comando_AbrirExt(lConexao)
''    If lComando = 0 Then Error 6264
''
''    'Lê as Empresas e Grupos associados ao Usuário
''    lErro = objSelect.Empresas_Grupos_Le(sCodUsuario, colEmpresaGrupo)
''    If lErro Then Error 6265
''
''    'Compara as Empresas lidas no BD com a EmpresaAtual
''    For Each objEmpresaGrupo In colEmpresaGrupo
''
''        If objEmpresaGrupo.lCodEmpresa = glEmpresaAtual Then
''
''            bCoincidencia = True
''            Exit For
''
''        End If
''
''    Next
''
''    If bCoincidencia Then      'Houve Coincidência de Empresa
''
''        'Operação é modificação
''        iOperacao = MODIFICACAO
''
''        'Grupo da Tela é diferente do Grupo associado no BD
''        If objEmpresaGrupo.sCodGrupo <> sCodGrupo Then
''
''            sCodUsuario1 = String(STRING_USUARIO_CODIGO)
''
''            lErro = Comando_ExecutarPos(lComando, "SELECT CodUsuario FROM UsuEmpGrupo WHERE CodUsuario=? AND CodEmpresa=?", 0, sCodUsuario1, sCodUsuario, glEmpresaAtual)
''            If lErro <> AD_SQL_SUCESSO Then Error 6266
''
''            'Lê o registro em UsuEmpGrupo se estiver no BD
''            lErro = Comando_BuscarPrimeiro(lComando)
''            If lErro = AD_SQL_SEM_DADOS Then Error 6267
''            If lErro <> AD_SQL_SUCESSO Then Error 6286
''
''            lErro = Comando_ExecutarPos(lComando1, "UPDATE UsuEmpGrupo SET CodGrupo=?", lComando, sCodGrupo)
''            If lErro <> AD_SQL_SUCESSO Then Error 6287
''
''        End If
''
''    Else    'Não houve coincidência de Empresa. REPLICAÇÃO.
''
''        'Operacao é Gravação
''        iOperacao = GRAVACAO
''
''        'insere registro que liga usuário a sCodGrupo
''        lErro = UsuEmpGrupo_Grava(sCodGrupo, sCodUsuario)
''        If lErro <> SUCESSO Then Error 6288
''
''    End If
''
''    lErro = Comando_Fechar(lComando)
''
''    UsuEmpGrupo_Altera = SUCESSO
''
''    Exit Function
''
''Erro_UsuEmpGrupo_Altera:
''
''    UsuEmpGrupo_Altera = Err
''
''    Select Case Err
''
''        Case 6263
''           lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 6264
''           lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''           lErro = Comando_Fechar(lComando)
''
''        Case 6265  'Erro tratado na rotina chamada
''            lErro = Comando_Fechar(lComando)
''            lErro = Comando_Fechar(lComando1)
''            lErro = Comando_Fechar(lComando)
''
''        Case 6266, 6286
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUEMPGRUPO1", Err, sCodUsuario, glEmpresaAtual)
''            lErro = Comando_Fechar(lComando1)
''            lErro = Comando_Fechar(lComando)
''
''        Case 6267
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUEMPGRUPO_NAO_CADASTRADO", Err, sCodUsuario, glEmpresaAtual)
''            lErro = Comando_Fechar(lComando1)
''            lErro = Comando_Fechar(lComando)
''
''        Case 6287
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_USUEMPGRUPO", Err, sCodUsuario, glEmpresaAtual)
''            lErro = Comando_Fechar(lComando1)
''            lErro = Comando_Fechar(lComando)
''
''        Case 6288  'Erro tratado na rotina chamada
''            lErro = Comando_Fechar(lComando1)
''            lErro = Comando_Fechar(lComando)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158994)
''
''    End Select
''
''    Exit Function
''
''End Function

Function DicUsuario_Grava(ByVal objUsuario As ClassDicUsuario, colUsuFilEmp As Collection, iOperacao As Integer) As Long
'Grava Usuário fornecido em objUsuario no BD.
'cada item em colUsuFilEmp representa a permissao de acesso a uma filialempresa
'Retorna GRAVACAO ou MODIFICACAO em iOperacao.

Dim lConexao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sCodUsuario As String
Dim sCodGrupo1 As String

On Error GoTo Erro_DicUsuario_Grava

    lConexao = GL_lConexaoDic
    
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 6247
    
    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 6248

    lComando2 = Comando_AbrirExt(lConexao)
    If lComando2 = 0 Then Error 6249

    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 6252
    
    '======================== GRUPO ========================
    
    If UCase(Trim(objUsuario.sCodUsuario)) = "SUPERVISOR" And UCase(Trim(objUsuario.sCodGrupo)) <> "SUPERVISOR" Then Error 6585
    
    sCodGrupo1 = String(STRING_GRUPO_CODIGO, 0)
    
    'Pesquisa sCodGrupo no BD
    lErro = Comando_ExecutarLockado(lComando, "SELECT CodGrupo FROM GruposDeUsuarios WHERE CodGrupo = ?", sCodGrupo1, objUsuario.sCodGrupo)
    If lErro Then Error 6253
    
    'Lê o Grupo, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6254 'Grupo não cadastrado
    If lErro <> AD_SQL_SUCESSO Then Error 6255

    'Locka o Grupo (lock shared)
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 6256

    '========================USUARIO============================
    
    sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    
    'Pesquisa no BD o Usuário em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT CodUsuario FROM Usuarios WHERE CodUsuario=?", 0, sCodUsuario, objUsuario.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 6257

    'Lê o Usuário, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6258

    If lErro = AD_SQL_SEM_DADOS Then
    
        'não encontrou o Usuario ==>  é uma inserção de Usuario
        iOperacao = GRAVACAO
        
        'insere o Usuário no BD
        lErro = Comando_Executar(lComando2, "INSERT INTO Usuarios (CodUsuario, CodGrupo, Nome, NomeReduzido, Senha, DataValidade, Ativo,Email) VALUES (?,?,?,?,?,?,?,?)", objUsuario.sCodUsuario, objUsuario.sCodGrupo, objUsuario.sNome, objUsuario.sNomeReduzido, objUsuario.sSenha, objUsuario.dtDataValidade, objUsuario.iAtivo, objUsuario.sEmail)
        If lErro <> AD_SQL_SUCESSO Then Error 6259
        
        lErro = Usuario_AtualizaBDsDados(objUsuario, colUsuFilEmp, iOperacao)
        If lErro <> SUCESSO Then Error 56834
        
        lErro = UsuFilEmp_Grava_Usuario(objUsuario.sCodUsuario, colUsuFilEmp)
        If lErro <> SUCESSO Then Error 32048
          
    Else    'encontrou o Usuário  - ALTERACAO de Usuário
        
        iOperacao = MODIFICACAO
        
        'Verifica se é comprador nas Filiais Empresa
        lErro = Usuario_Testa_Vinculo_FilialEmpresa(objUsuario, colUsuFilEmp)
        If lErro <> SUCESSO Then Error 58537
        
        lErro = Usuario_AtualizaBDsDados(objUsuario, colUsuFilEmp, iOperacao)
        If lErro <> SUCESSO Then Error 56835
        
        lErro = UsuFilEmp_Exclui_Usuario(objUsuario.sCodUsuario)
        If lErro <> AD_SQL_SUCESSO Then Error 32049
        
        lErro = UsuFilEmp_Grava_Usuario(objUsuario.sCodUsuario, colUsuFilEmp)
        If lErro <> SUCESSO Then Error 32050
        
        'atualiza o Usuário
        lErro = Comando_ExecutarPos(lComando2, "UPDATE Usuarios SET CodGrupo=?,Nome=?,NomeReduzido=?,Senha=?,DataValidade=?,Ativo=?,Email=?", lComando1, objUsuario.sCodGrupo, objUsuario.sNome, objUsuario.sNomeReduzido, objUsuario.sSenha, objUsuario.dtDataValidade, objUsuario.iAtivo, objUsuario.sEmail)
        If lErro <> AD_SQL_SUCESSO Then Error 6261
          
    End If

    '========================== FIM DA TRANSACAO =================

    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 6268

    lErro = Comando_Fechar(lComando2)
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)

    DicUsuario_Grava = SUCESSO

    Exit Function

Erro_DicUsuario_Grava:

    DicUsuario_Grava = Err

    Select Case Err

        Case 6247
           lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6248
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6249
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 6253, 6255
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO1", Err, objUsuario.sCodGrupo)
            
        Case 6254
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPO_NAO_CADASTRADO", Err, objUsuario.sCodGrupo)
          
        Case 6256
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_GRUPO", Err, objUsuario.sCodGrupo)
        
        Case 6257, 6258
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIO", Err, objUsuario.sCodUsuario)
        
        Case 6259
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_USUARIO", Err, objUsuario.sCodUsuario)
           
        Case 6261
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_USUARIO", Err, objUsuario.sCodUsuario)
 
        Case 6268
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 6585
            Call Rotina_Erro(vbOKOnly, "ERRO_USU_SUPERVISOR_GRUPO_NAO_SUPERVISOR", Err)

        Case 32048, 32049, 32050, 56834, 56835, 58537
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158995)

    End Select

    lErro = Transacao_RollbackDic
    lErro = Comando_Fechar(lComando2)
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)
    
    Exit Function

End Function

'Alterada por Luiz Nogueira em 27/10/03
Function Usuario_Testa_Vinculo_FilialEmpresa(objUsuario As ClassDicUsuario, colUsuFilEmp As Collection) As Long
'Testa se o Usuario é Comprador nas Filiais Empresas

Dim lErro As Long
Dim objDicSelect As New DicSelect
Dim colUsuFilEmpBD As New Collection
Dim objUsuFilEmpBD As ClassUsuFilEmp
Dim objUsuFilEmp As ClassUsuFilEmp
Dim objFilialEmpresa As New AdmFiliais
Dim iEncontrou As Integer

On Error GoTo Erro_Usuario_Testa_Vinculo_FilialEmpresa
    
    'Lê em UsuFilEmp todas as Filiais Empresa que este Usuário tem acesso
    lErro = objDicSelect.UsuFilEmp_Le_Usuario(objUsuario.sCodUsuario, colUsuFilEmpBD)
    If lErro <> SUCESSO Then gError 58530
        
    'Verifico se todos os itens da minha coleção colUsuFilEmpBD estão em colUsuFilEmp
    For Each objUsuFilEmpBD In colUsuFilEmpBD
        
        'Zera iEncontrou que é o flag se encontrar
        iEncontrou = 0
        
        'Procura na Colecao que veio da Tela
        For Each objUsuFilEmp In colUsuFilEmp
            
            If (objUsuFilEmp.lCodEmpresa = objUsuFilEmpBD.lCodEmpresa) And (objUsuFilEmp.iCodFilial = objUsuFilEmpBD.iCodFilial) Then
                
                'Seta o Flag que encontrou
                iEncontrou = 1
                Exit For
            
            End If
        Next
       
        'Se não encontrou é porque desmarcou a FilialEmpresa na LIsta
        If iEncontrou = 0 Then
        
            objFilialEmpresa.lCodEmpresa = objUsuFilEmpBD.lCodEmpresa
            objFilialEmpresa.iCodFilial = objUsuFilEmpBD.iCodFilial
            
            'Abre a conexao com o banco de dados da Empresa
            lErro = FilialEmpresa_Conecta_Contab(objFilialEmpresa)
            If lErro <> SUCESSO Then gError 58531
            
            'Testa se o Usuario é Comprador --> Se for não pode desmarcar esta FilialEmpresa
            lErro = Usuario_Testa_Comprador_FilialEmpresa(objUsuario.sCodUsuario, objUsuFilEmpBD.iCodFilial)
            If lErro <> SUCESSO Then gError 58532
        
            'Incluído por Luiz Nogueira em 27/10/03
            'Testa se o Usuario é Atendente --> Se for não pode desmarcar esta FilialEmpresa
            lErro = objDicSelect.Usuario_Testa_Atendente(objUsuario.sCodUsuario, objUsuFilEmpBD.iCodFilial)
            If lErro <> SUCESSO Then gError 102790
        
        End If
        
    Next
    
    Usuario_Testa_Vinculo_FilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_Usuario_Testa_Vinculo_FilialEmpresa:

    Usuario_Testa_Vinculo_FilialEmpresa = gErr
    
    Select Case gErr
        
        Case 58530, 58531, 58532, 102790 'Erros Tratados nas Rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158996)

    End Select
    
    Exit Function
    
End Function

Private Function Usuario_Testa_Comprador_FilialEmpresa(sCodUsuario As String, iFilialEmpresa As Integer) As Long
'Verifica se o Usuario passado é Comprador na FilialEmpresa passada

Dim lErro As Long
Dim lComando As Long
Dim iCodComprador As Integer

On Error GoTo Erro_Usuario_Testa_Comprador_FilialEmpresa
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58533
        
    'Le o Usuario na Tabela de Comprador
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Comprador WHERE FilialEmpresa = ? AND CodUsuario = ?", iCodComprador, iFilialEmpresa, sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 58534
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58535
    
    'se encontrou --> Erro o usuario é Comprador
    If lErro = AD_SQL_SUCESSO Then Error 58536
    
    Call Comando_Fechar(lComando)
    
    Usuario_Testa_Comprador_FilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_Usuario_Testa_Comprador_FilialEmpresa:

    Usuario_Testa_Comprador_FilialEmpresa = Err
    
    Select Case Err
        
        Case 58533
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 58534, 58535
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_COMPRADOR", Err, gsNomeEmpresa, gsNomeFilialEmpresa)
        
        Case 58536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_CADASTRADO_COMPRADOR", Err, sCodUsuario, gsNomeEmpresa, gsNomeFilialEmpresa)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158997)

    End Select
            
    Call Comando_Fechar(lComando)
            
    Exit Function
    
End Function

Private Function Usuario_Testa_Comprador(sCodUsuario As String) As Long
'Verifica se o Usuario passado é Comprador

Dim lErro As Long
Dim lComando As Long
Dim iCodComprador As Integer

On Error GoTo Erro_Usuario_Testa_Comprador
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58533
        
    'Le o Usuario na Tabela de Comprador
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Comprador WHERE CodUsuario = ?", iCodComprador, sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 58534
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58535
    
    'se encontrou --> Erro o usuario é Comprador
    If lErro = AD_SQL_SUCESSO Then Error 58536
    
    Call Comando_Fechar(lComando)
    
    Usuario_Testa_Comprador = SUCESSO
    
    Exit Function
    
Erro_Usuario_Testa_Comprador:

    Usuario_Testa_Comprador = Err
    
    Select Case Err
        
        Case 58533
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 58534, 58535
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_COMPRADOR", Err, gsNomeEmpresa, gsNomeFilialEmpresa)
        
        Case 58536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_CADASTRADO_COMPRADOR", Err, sCodUsuario, gsNomeEmpresa, gsNomeFilialEmpresa)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158998)

    End Select
            
    Call Comando_Fechar(lComando)
            
    Exit Function
    
End Function

Public Function Rotina_Grava(ByVal objRotina As ClassDicRotina) As Long
'Atualiza a Rotina com os dados em objRotina

Dim lConexao As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sSigla As String

On Error GoTo Erro_Rotina_Grava

    sSigla = String(STRING_ROTINA_SIGLA, 0)

    lConexao = GL_lConexaoDic
    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 6323

    lComando2 = Comando_AbrirExt(lConexao)
    If lComando2 = 0 Then Error 6324

    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 6325

    'pesquisa no BD a Rotina em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT Sigla FROM Rotinas WHERE Sigla=?", 0, sSigla, objRotina.sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 6326

    'Lê a Rotina, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro = AD_SQL_SEM_DADOS Then Error 6327
    If lErro <> AD_SQL_SUCESSO Then Error 6328

    'atualiza a Rotina
    lErro = Comando_ExecutarPos(lComando2, "UPDATE Rotinas SET Descricao=?, projeto_customizado=?, classe_customizado=?", lComando1, objRotina.sDescricao, objRotina.sProjeto_Customizado, objRotina.sClasse_Customizada)
    If lErro <> AD_SQL_SUCESSO Then Error 6329

    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 6330

    lErro = Comando_Fechar(lComando2)
    lErro = Comando_Fechar(lComando1)

    Rotina_Grava = SUCESSO

    Exit Function

Erro_Rotina_Grava:

    Rotina_Grava = Err

    Select Case Err

        Case 6323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6324
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6325
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 6326, 6328
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTINA1", Err, objRotina.sSigla)
            
        Case 6327
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ROTINA_NAO_CADASTRADA", Err, objRotina.sSigla)

        Case 6329
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ROTINA", Err, objRotina.sSigla)
 
        Case 6330
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158999)

    End Select

    lErro = Transacao_RollbackDic
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando2)
    
    Exit Function

End Function

Public Function Tela_Grava(ByVal objTela As ClassDicTela) As Long
'Atualiza a Tela com os dados em objTela

Dim lConexao As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sNome As String

On Error GoTo Erro_Tela_Grava

    sNome = String(STRING_TELA_NOME, 0)

    lConexao = GL_lConexaoDic
    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 6357

    lComando2 = Comando_AbrirExt(lConexao)
    If lComando2 = 0 Then Error 6358

    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 6359

    'pesquisa no BD a Tela em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT Nome FROM Telas WHERE Nome=?", 0, sNome, objTela.sNome)
    If lErro <> AD_SQL_SUCESSO Then Error 6360

    'Lê a Tela, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro = AD_SQL_SEM_DADOS Then Error 6361
    If lErro <> AD_SQL_SUCESSO Then Error 6362

    'atualiza a Tela
    lErro = Comando_ExecutarPos(lComando2, "UPDATE Telas SET Descricao=?, projeto_customizado=?, classe_customizado=?", lComando1, objTela.sDescricao, objTela.sProjeto_Customizado, objTela.sClasse_Customizada)
    If lErro <> AD_SQL_SUCESSO Then Error 6363

    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 6364

    lErro = Comando_Fechar(lComando2)
    lErro = Comando_Fechar(lComando1)

    Tela_Grava = SUCESSO

    Exit Function

Erro_Tela_Grava:

    Tela_Grava = Err

    Select Case Err

        Case 6357
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 6360, 6362
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELA1", Err, objTela.sNome)
            
        Case 6361
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_CADASTRADA", Err, objTela.sNome)

        Case 6363
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TELA", Err, objTela.sNome)
 
        Case 6364
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159000)

    End Select

    lErro = Transacao_RollbackDic
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando2)
    
    Exit Function

End Function

Public Function GrupoRotina_Grava(ByVal colGrupoRotina As colGrupoRotina) As Long
'Atualiza registros em GrupoRotinas de varios grupos de uma rotina
'Membros de colGrupoRotina têm a mesma SiglaRotina.

Dim lConexao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sCodGrupo As String
Dim sSiglaRotina As String
Dim iIndice As Integer

On Error GoTo Erro_GrupoRotina_Grava

    lConexao = GL_lConexaoDic
    
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 6387
    
    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 6388

    'Inicia a Transação
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 6389
    
    sSiglaRotina = colGrupoRotina.Item(1).sSiglaRotina
    sCodGrupo = String(STRING_GRUPO_ROTINA_CODGRUPO, 0)
    
    'Pesquisa no BD os registros de GrupoRotinas ligados à sSiglaRotina
    lErro = Comando_ExecutarPos(lComando, "SELECT CodGrupo FROM GrupoRotinas WHERE SiglaRotina=? ORDER BY CodGrupo", 0, sCodGrupo, sSiglaRotina)
    If lErro <> AD_SQL_SUCESSO Then Error 6390

    'Lê registro de GrupoRotinas, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6391
    If lErro <> AD_SQL_SUCESSO Then Error 6392
    
    iIndice = 1  'iIndice é o índice na coleção colGrupoRotina
    
    'Loop de Atualização de GrupoRotina
    Do While lErro = AD_SQL_SUCESSO
               
        'teste se iIndice é válido
        If iIndice > colGrupoRotina.Count Then Error 6393
        
        'testa se CodGrupo lido no BD coincide com CodGrupo na coleção
        If sCodGrupo <> colGrupoRotina.Item(iIndice).sCodGrupo Then Error 6394
        
        'atualiza o registro de GrupoRotinas
        lErro = Comando_ExecutarPos(lComando1, "UPDATE GrupoRotinas SET TipoDeAcesso=?,projeto=?,classe=?", lComando, colGrupoRotina.Item(iIndice).iTipoDeAcesso, colGrupoRotina.Item(iIndice).sProjeto, colGrupoRotina.Item(iIndice).sClasse)
        If lErro <> AD_SQL_SUCESSO Then Error 6395
        
        'vai para o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        iIndice = iIndice + 1
      
    Loop
    
    If lErro <> AD_SQL_SEM_DADOS Then Error 6396
    
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 6397

    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)

    GrupoRotina_Grava = SUCESSO

    Exit Function

Erro_GrupoRotina_Grava:

    GrupoRotina_Grava = Err

    Select Case Err

        Case 6387
           lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
 
        Case 6389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
           
        Case 6390, 6392, 6396
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_ROTINAS", Err)
        
        Case 6391
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOROTINA_ROTINA_NAO_CADASTRADOS", Err, sSiglaRotina)
            
        Case 6393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_GRUPO_NA_COLECAO", Err, sCodGrupo)
            
        Case 6394
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGOS_GRUPO_DIFERENTES", Err, sCodGrupo, colGrupoRotina.Item(iIndice).sCodGrupo)
        
        Case 6395
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_GRUPO_ROTINAS", Err, sSiglaRotina, sCodGrupo)
 
        Case 6397
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159001)

    End Select

    lErro = Transacao_RollbackDic
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)

    Exit Function
    
End Function

Public Function GrupoTela_Grava(ByVal colGrupoTela As colGrupoTela) As Long
'Atualiza registros em GrupoTelas de varios grupos de uma tela
'Os dados de colGrupoTela correspondem a uma única Tela (mesmo NomeTela).

Dim lConexao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sCodGrupo As String
Dim sTela As String
Dim iIndice As Integer

On Error GoTo Erro_GrupoTela_Grava

    lConexao = GL_lConexaoDic
    
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 6424
    
    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 6425

    'Inicia a Transação
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 6426
    
    sTela = colGrupoTela.Item(1).sNomeTela
    sCodGrupo = String(STRING_GRUPO_TELA_CODGRUPO, 0)
    
    'Pesquisa no BD os registros de GrupoTela ligados à sTela
    lErro = Comando_ExecutarPos(lComando, "SELECT CodGrupo FROM GrupoTela WHERE NomeTela=? ORDER BY CodGrupo", 0, sCodGrupo, sTela)
    If lErro <> AD_SQL_SUCESSO Then Error 6427

    'Lê registro de GrupoTela, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6428
    If lErro <> AD_SQL_SUCESSO Then Error 6429
    
    iIndice = 1  'iIndice é o índice na coleção colGrupoTela
    
    'Loop de Atualização de GrupoTela
    Do While lErro = AD_SQL_SUCESSO
               
        'teste se iIndice é válido
        If iIndice > colGrupoTela.Count Then Error 6430
        
        'testa se CodGrupo lido no BD coincide com CodGrupo na coleção
        If sCodGrupo <> colGrupoTela.Item(iIndice).sCodGrupo Then Error 6431
        
        'atualiza o registro de GrupoTela
        lErro = Comando_ExecutarPos(lComando1, "UPDATE GrupoTela SET TipoDeAcesso=?,projeto=?,classe=?", lComando, colGrupoTela.Item(iIndice).iTipoDeAcesso, colGrupoTela.Item(iIndice).sProjeto, colGrupoTela.Item(iIndice).sClasse)
        If lErro <> AD_SQL_SUCESSO Then Error 6432
        
        'vai para o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        iIndice = iIndice + 1
      
    Loop
    
    If lErro <> AD_SQL_SEM_DADOS Then Error 6433
    
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 6434

    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)

    GrupoTela_Grava = SUCESSO

    Exit Function

Erro_GrupoTela_Grava:

    GrupoTela_Grava = Err

    Select Case Err

        Case 6424
           lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6425
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
 
        Case 6426
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
           
        Case 6427, 6429, 6433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_TELA", Err)
        
        Case 6428
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOTELA_TELA_NAO_CADASTRADOS", Err, sTela)
            
        Case 6430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_GRUPO_NA_COLECAO2", Err, sCodGrupo)
            
        Case 6431
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGOS_GRUPO_DIFERENTES", Err, sCodGrupo, colGrupoTela.Item(iIndice).sCodGrupo)
        
        Case 6432
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_GRUPO_TELA", Err, sTela, sCodGrupo)
 
        Case 6434
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159002)

    End Select

    lErro = Transacao_RollbackDic
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)
            
    Exit Function
    
End Function

Public Function GrupoRotina_Grava2(ByVal sModulo As String, ByVal colGrupoRotina As colGrupoRotina) As Long
'Atualiza registros em GrupoRotinas de varias rotinas de um par (grupo,Modulo)
'Membros de colGrupoRotina têm o mesmo CodGrupo e as rotinas pertencem a sModulo.

Dim lConexao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sCodGrupo As String
Dim sSiglaRotina As String
Dim iIndice As Integer

On Error GoTo Erro_GrupoRotina_Grava2

    lConexao = GL_lConexaoDic
    
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 6466
    
    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 6467

    'Inicia a Transação
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 6468
    
    sCodGrupo = colGrupoRotina.Item(1).sCodGrupo
    sSiglaRotina = String(STRING_GRUPO_ROTINA_SIGLAROTINA, 0)
    
    'Pesquisa no BD os registros de GrupoRotinas associados a sCodGrupo e a rotinas em sModulo
    If giSQLTipoOrdParamDic = SQL_ORD_PARAM_ESQ_DIR Then
        lErro = Comando_ExecutarPos(lComando, "SELECT SiglaRotina FROM GrupoRotinas WHERE CodGrupo=? AND SiglaRotina IN (SELECT SiglaRotina FROM RotinasModulo, Modulos WHERE RotinasModulo.SiglaModulo=Modulos.Sigla AND Modulos.Nome=?) ORDER BY SiglaRotina", 0, sSiglaRotina, sCodGrupo, sModulo)
    Else
        lErro = Comando_ExecutarPos(lComando, "SELECT SiglaRotina FROM GrupoRotinas WHERE CodGrupo=? AND SiglaRotina IN (SELECT SiglaRotina FROM RotinasModulo, Modulos WHERE RotinasModulo.SiglaModulo=Modulos.Sigla AND Modulos.Nome=?) ORDER BY SiglaRotina", 0, sSiglaRotina, sModulo, sCodGrupo)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 6469

    'Lê registro de GrupoRotinas, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6470
    If lErro <> AD_SQL_SUCESSO Then Error 6471
    
    iIndice = 1  'iIndice é o índice na coleção colGrupoRotina
    
    'Loop de Atualização de GrupoRotina
    Do While lErro = AD_SQL_SUCESSO
               
        'teste se iIndice é válido
        If iIndice > colGrupoRotina.Count Then Error 6472
        
        'testa se SiglaRotina lida no BD coincide com SiglaRotina na coleção
        If sSiglaRotina <> colGrupoRotina.Item(iIndice).sSiglaRotina Then Error 6473
        
        'atualiza o registro de GrupoRotinas
        lErro = Comando_ExecutarPos(lComando1, "UPDATE GrupoRotinas SET TipoDeAcesso=?,projeto=?,classe=?", lComando, colGrupoRotina.Item(iIndice).iTipoDeAcesso, colGrupoRotina.Item(iIndice).sProjeto, colGrupoRotina.Item(iIndice).sClasse)
        If lErro <> AD_SQL_SUCESSO Then Error 6474
        
        'vai para o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        iIndice = iIndice + 1
      
    Loop
    
    If lErro <> AD_SQL_SEM_DADOS Then Error 6475
    
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 6476

    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)

    GrupoRotina_Grava2 = SUCESSO

    Exit Function

Erro_GrupoRotina_Grava2:

    GrupoRotina_Grava2 = Err

    Select Case Err

        Case 6466
           lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
 
        Case 6468
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
           
        Case 6469, 6471, 6475
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_ROTINAS", Err)
        
        Case 6470
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOROTINA_GRUPOMODULO_NAO_CADASTRADOS", Err, sCodGrupo, sModulo)
            
        Case 6472
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_ROTINA_NA_COLECAO", Err, sSiglaRotina)
            
        Case 6473
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLAS_ROTINA_DIFERENTES", Err, sSiglaRotina, colGrupoRotina.Item(iIndice).sSiglaRotina)
        
        Case 6474
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_GRUPO_ROTINAS", Err, sSiglaRotina, sCodGrupo)
 
        Case 6476
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159003)

    End Select

    lErro = Transacao_RollbackDic
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function GrupoTela_Grava2(ByVal sModulo As String, ByVal colGrupoTela As colGrupoTela) As Long
'Atualiza registros em GrupoTelas de varias rotinas de um par (grupo,Modulo)
'Membros de colGrupoTela têm o mesmo CodGrupo e as telas pertencem a sModulo.

Dim lConexao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sCodGrupo As String
Dim sNomeTela As String
Dim iIndice As Integer

On Error GoTo Erro_GrupoTela_Grava2

    lConexao = GL_lConexaoDic
    
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 6503
    
    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 6504

    'Inicia a Transação
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 6505
    
    sCodGrupo = colGrupoTela.Item(1).sCodGrupo
    sNomeTela = String(STRING_GRUPO_TELA_NOMETELA, 0)
    
    'Pesquisa no BD os registros de GrupoTela associados a sCodGrupo e a telas em sModulo
    If giSQLTipoOrdParamDic = SQL_ORD_PARAM_ESQ_DIR Then
        lErro = Comando_ExecutarPos(lComando, "SELECT NomeTela FROM GrupoTela WHERE CodGrupo=? AND NomeTela IN (SELECT NomeTela FROM TelasModulo, Modulos WHERE TelasModulo.SiglaModulo=Modulos.Sigla AND Modulos.Nome=?) ORDER BY NomeTela", 0, sNomeTela, sCodGrupo, sModulo)
    Else
        lErro = Comando_ExecutarPos(lComando, "SELECT NomeTela FROM GrupoTela WHERE CodGrupo=? AND NomeTela IN (SELECT NomeTela FROM TelasModulo, Modulos WHERE TelasModulo.SiglaModulo=Modulos.Sigla AND Modulos.Nome=?) ORDER BY NomeTela", 0, sNomeTela, sModulo, sCodGrupo)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 6506

    'Lê registro de GrupoTela, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6507
    If lErro <> AD_SQL_SUCESSO Then Error 6508
    
    iIndice = 1  'iIndice é o índice na coleção colGrupoTela
    
    'Loop de Atualização de GrupoTela
    Do While lErro = AD_SQL_SUCESSO
               
        'teste se iIndice é válido
        If iIndice > colGrupoTela.Count Then Error 6509
        
        'testa se NomeTela lida no BD coincide com NomeTela na coleção
        If sNomeTela <> colGrupoTela.Item(iIndice).sNomeTela Then Error 6510
        
        'atualiza o registro de GrupoTela
        lErro = Comando_ExecutarPos(lComando1, "UPDATE GrupoTela SET TipoDeAcesso=?,projeto=?,classe=?", lComando, colGrupoTela.Item(iIndice).iTipoDeAcesso, colGrupoTela.Item(iIndice).sProjeto, colGrupoTela.Item(iIndice).sClasse)
        If lErro <> AD_SQL_SUCESSO Then Error 6511
        
        'vai para o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        iIndice = iIndice + 1
      
    Loop
    
    If lErro <> AD_SQL_SEM_DADOS Then Error 6512
    
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 6513

    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)

    GrupoTela_Grava2 = SUCESSO

    Exit Function

Erro_GrupoTela_Grava2:

    GrupoTela_Grava2 = Err

    Select Case Err

        Case 6503
           lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6504
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
 
        Case 6505
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
           
        Case 6506, 6508, 6512
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_TELA", Err)
        
        Case 6507
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPOTELA_GRUPOMODULO_NAO_CADASTRADOS", Err, sCodGrupo, sModulo)
            
        Case 6509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_TELA_NA_COLECAO", Err, sNomeTela)
            
        Case 6510
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOMES_TELA_DIFERENTES", Err, sNomeTela, colGrupoTela.Item(iIndice).sNomeTela)
        
        Case 6511
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_GRUPO_TELA", Err, sNomeTela, sCodGrupo)
 
        Case 6513
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159004)

    End Select

    lErro = Transacao_RollbackDic
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)

    Exit Function
    
End Function

Public Function MenuItem_Grava(ByVal sModulo As String, ByVal sTipoRotina As String, ByVal colMenuItem As colMenuItem) As Long
'Atualiza dados de colMenuItem na tabela MenuItens.
'Membros de colMenuItem têm siglas de rotinas "_USU_" que pertencem a sModulo.

Dim lConexao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sSiglaRotina As String
Dim sSubstring As String
Dim iIndice As Integer
Dim sTitulo As String

On Error GoTo Erro_MenuItem_Grava

    lConexao = GL_lConexaoDic
    
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 6526
    
    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 6527

    'Inicia a Transação
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 6528
    
    sSiglaRotina = String(STRING_MENU_ITEM_SIGLAROTINA, 0)

    'Substring usada na seleção dos ítens de menu
    sSubstring = "%" & "_USU_" & sTipoRotina & "%"
    
    'Pesquisa no BD os registros de MenuItens correspondentes a rotinas do sTipoRotina em sModulo
    If giSQLTipoOrdParamDic = SQL_ORD_PARAM_ESQ_DIR Then
        lErro = Comando_ExecutarPos(lComando, "SELECT SiglaRotina FROM MenuItens WHERE SiglaRotina LIKE ? AND SiglaRotina IN (SELECT SiglaRotina FROM RotinasModulo, Modulos WHERE RotinasModulo.SiglaModulo=Modulos.Sigla AND Modulos.Nome=?) ORDER BY SiglaRotina", 0, sSiglaRotina, sSubstring, sModulo)
    Else
        lErro = Comando_ExecutarPos(lComando, "SELECT SiglaRotina FROM MenuItens WHERE SiglaRotina LIKE ? AND SiglaRotina IN (SELECT SiglaRotina FROM RotinasModulo, Modulos WHERE RotinasModulo.SiglaModulo=Modulos.Sigla AND Modulos.Nome=?) ORDER BY SiglaRotina", 0, sSiglaRotina, sModulo, sSubstring)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 6529

    'Lê registro de MenuItens, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6530
    If lErro <> AD_SQL_SUCESSO Then Error 6531
    
    iIndice = 1  'iIndice é o índice na coleção colMenuItem
    
    'Loop de Atualização de MenuItens
    Do While lErro = AD_SQL_SUCESSO
               
        'teste se iIndice é válido
        If iIndice > colMenuItem.Count Then Error 6532
        
        'testa se SiglaRotina lida no BD coincide com SiglaRotina na coleção
        If sSiglaRotina <> colMenuItem.Item(iIndice).sSiglaRotina Then Error 6533
        
        'armazena o valor de sTitulo em variável string
        sTitulo = colMenuItem.Item(iIndice).sTitulo
        
        'atualiza o registro de MenuItens
        lErro = Comando_ExecutarPos(lComando1, "UPDATE MenuItens SET Titulo=?", lComando, sTitulo)
        If lErro <> AD_SQL_SUCESSO Then Error 6534
        
        'vai para o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        iIndice = iIndice + 1
      
    Loop
    
    If lErro <> AD_SQL_SEM_DADOS Then Error 6535
    
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 6536

    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)

    MenuItem_Grava = SUCESSO

    Exit Function

Erro_MenuItem_Grava:

    MenuItem_Grava = Err

    Select Case Err

        Case 6526
           lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6527
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
 
        Case 6528
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
           
        Case 6529, 6531, 6535
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MENU_ITENS", Err)
        
        Case 6530
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MENUITEM_ROTINAS_MODULO_NAO_CADASTRADOS", Err, sTipoRotina, sModulo)
            
        Case 6532
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_MENUITEM_NA_COLECAO", Err, sSiglaRotina)
            
        Case 6533
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLAS_ROTINA_DIFERENTES", Err, sSiglaRotina, colMenuItem.Item(iIndice).sSiglaRotina)
        
        Case 6534
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MENU_ITENS", Err, sSiglaRotina)
 
        Case 6536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159005)

    End Select

    lErro = Transacao_RollbackDic
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)

    Exit Function
    
End Function

Public Function GrupoBrowseCampo_Grava(ByVal sGrupo As String, ByVal sTela As String, ByVal sArquivo As String, ByVal colCampo As Collection) As Long
'Exclui todos os registros da tabela GrupoBrowseCampo associados a
'sGrupo, sTela, sArquivo. Inclui os campos de colCampo.

Dim lConexao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sCodGrupo As String
Dim vCampo As Variant

On Error GoTo Erro_GrupoBrowseCampo_Grava

    lConexao = GL_lConexaoDic
    
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 6572
    
    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 6573

    lComando2 = Comando_AbrirExt(lConexao)
    If lComando2 = 0 Then Error 6574

    'Inicia a Transação
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 6575
    
    sCodGrupo = String(STRING_GRUPO_CODIGO, 0)
    
    'Seleciona sGrupo
    lErro = Comando_ExecutarLockado(lComando, "SELECT CodGrupo FROM GruposDeUsuarios WHERE CodGrupo=?", sCodGrupo, sGrupo)
    If lErro <> AD_SQL_SUCESSO Then Error 6576
    
    'Lê o Grupo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6577
    If lErro <> AD_SQL_SUCESSO Then Error 6578
    
    'Faz o lock no Grupo (impede exclusão de Grupo)
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 6579
    
    'Exclui os registros de GrupoBrowseCampo associados a sGrupo, sTela, sArquivo
    lErro = Comando_Executar(lComando1, "DELETE FROM GrupoBrowseCampo WHERE CodGrupo=? AND NomeTela=? AND NomeArq=?", sGrupo, sTela, sArquivo)
    If lErro <> AD_SQL_SUCESSO Then Error 6580
    
    'Loop de inclusão em GrupoBrowseCampo
    For Each vCampo In colCampo
        
        'insere novo registro em GrupoBrowseCampo
        lErro = Comando_Executar(lComando2, "INSERT INTO GrupoBrowseCampo (CodGrupo, NomeTela, NomeArq, Nome) VALUES (?,?,?,?)", sGrupo, sTela, sArquivo, vCampo)
        If lErro <> AD_SQL_SUCESSO Then Error 6581
        
    Next
    
    'faz com que a nivel de usuario seja tratada a alteracao dos campos visiveis p/o grupo de usuarios
    lErro = BrowseUsuarioCampo_GrupoAlterado(sGrupo, sTela, sArquivo, colCampo)
    If lErro <> SUCESSO Then Error 32061
    
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 6582

    lErro = Comando_Fechar(lComando2)
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)

    GrupoBrowseCampo_Grava = SUCESSO

    Exit Function

Erro_GrupoBrowseCampo_Grava:

    GrupoBrowseCampo_Grava = Err

    Select Case Err

        Case 6572
           lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
 
        Case 6574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
 
        Case 6575
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
           
        Case 6576, 6578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO", Err)
            
        Case 6577
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPO_NAO_CADASTRADO", Err, sGrupo)
            
        Case 6579
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_GRUPO", Err, sGrupo)
            
        Case 6580
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GRUPOBROWSECAMPO", Err)
         
        Case 6581
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_GRUPOBROWSECAMPO", Err, sGrupo, sTela, sArquivo, vCampo)
  
        Case 6582
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 32061
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159006)

    End Select
    
    lErro = Transacao_RollbackDic
    lErro = Comando_Fechar(lComando2)
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function BrowseUsuarioCampo_Exclui(ByVal sCodUsuario As String) As Long
'Exclui registros da Tabela BrowseUsuarioCampo com código do Usuário sCodUsuario
'Usada na Transação de Exclusão de Usuário

Dim lComando As Long
Dim lErro As Long

On Error GoTo Erro_BrowseUsuarioCampo_Exclui

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 6585
    
    'exclui os registros de BrowseUsuarioCampo que tem código do Usuário
    lErro = Comando_Executar(lComando, "DELETE FROM BrowseUsuarioCampo WHERE CodUsuario=?", sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 6586

    lErro = Comando_Fechar(lComando)

    BrowseUsuarioCampo_Exclui = SUCESSO
    
    Exit Function
    
Erro_BrowseUsuarioCampo_Exclui:

    BrowseUsuarioCampo_Exclui = Err
    
    Select Case Err
    
        Case 6585
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
         
        Case 6586
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BROWSEUSUARIOCAMPO2", Err, sCodUsuario)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159007)
    
    End Select
    
    lErro = Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Private Function BrowseUsuarioOrdenacao_Exclui(ByVal sCodUsuario As String) As Long
'Exclui registros da Tabela BrowseUsuarioOrdenacao com código do Usuário sCodUsuario
'Usada na Transação de Exclusão de Usuário

Dim lConexao As Long
Dim lComando As Long
Dim lErro As Long

On Error GoTo Erro_BrowseUsuarioOrdenacao_Exclui

    lConexao = GL_lConexaoDic
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 6587
    
    'exclui os registros de BrowseUsuarioOrdenacao que tem código do Usuário
    lErro = Comando_Executar(lComando, "DELETE FROM BrowseUsuarioOrdenacao WHERE CodUsuario=?", sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 6588

    lErro = Comando_Fechar(lComando)

    BrowseUsuarioOrdenacao_Exclui = SUCESSO
    
    Exit Function
    
Erro_BrowseUsuarioOrdenacao_Exclui:

    BrowseUsuarioOrdenacao_Exclui = Err
    
    Select Case Err
    
        Case 6587
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
         
        Case 6588
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BROWSEUSUARIOORDENACAO", Err, sCodUsuario)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159008)
    
    End Select
    
    lErro = Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Grupo_Grava(ByVal objGrupo As ClassDicGrupo, iAlteracao As Integer) As Long

Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim lConexao As Long
Dim sCodGrupo1 As String

On Error GoTo Erro_Grupo_Grava

    sCodGrupo1 = String(STRING_GRUPO_CODIGO, 0)
    
    lConexao = GL_lConexaoDic

    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 8235

    lComando2 = Comando_AbrirExt(lConexao)
    If lComando2 = 0 Then Error 8236

    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 8237

    'Pesquisa no BD o grupo em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT CodGrupo FROM GruposDeUsuarios WHERE CodGrupo=?", 0, sCodGrupo1, objGrupo.sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then Error 8238

    'Le o grupo, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando1)

    'Encontrou o grupo ==> isto é uma alteração
    If lErro = AD_SQL_SUCESSO Then

        'Atualiza o grupo
        lErro = Comando_ExecutarPos(lComando2, "UPDATE GruposDeUsuarios SET Descricao=?, DataValidade=?", lComando1, objGrupo.sDescricao, objGrupo.dtDataValidade)
        If lErro <> AD_SQL_SUCESSO Then Error 8239

        iAlteracao = MODIFICACAO

    Else
        'Não encontrou o grupo ==>  é uma inserção

        'Insere o grupo
        lErro = Comando_Executar(lComando2, "INSERT INTO GruposDeUsuarios (CodGrupo, Descricao, DataValidade) VALUES (?,?,?)", objGrupo.sCodGrupo, objGrupo.sDescricao, objGrupo.dtDataValidade)
        If lErro <> AD_SQL_SUCESSO Then Error 8240

        'Insere em GrupoRotinas e em GrupoTelas todas as combinacoes de Grupos com Rotinas e com CodTelas
        lErro = Insere_Grupo_Combinacoes(objGrupo.sCodGrupo)
        If lErro <> SUCESSO Then Error 8242

        iAlteracao = GRAVACAO

    End If

    'Confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 8241

    lErro = Comando_Fechar(lComando2)
    lErro = Comando_Fechar(lComando1)

    Grupo_Grava = SUCESSO

    Exit Function

Erro_Grupo_Grava:

    Grupo_Grava = Err

    Select Case Err

        Case 8235
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 8236
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 8237
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 8238
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO", Err, objGrupo.sCodGrupo)

        Case 8239
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_GRUPO", Err, objGrupo.sCodGrupo)

        Case 8240
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GRUPO", Err, objGrupo.sCodGrupo)

        Case 8241
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 8242, 8243

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159009)

    End Select

    lErro = Transacao_RollbackDic
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando2)
    
    Exit Function

End Function

Private Function Insere_Grupo_Combinacoes(ByVal sCodGrupo) As Long
'??? pode ser otimizada com INSERT INTO direto do SELECT (Jones 15/04/99)
Dim lComando1 As Long
Dim lComando2 As Long
Dim lErro As Long
Dim lConexao As Long, iFim_de_Arquivo As Integer
Dim sSiglaRot As String, sNomeTela As String, sNomeRelat As String
Dim sNomeArq As String, sNomeCampo As String

On Error GoTo Erro_Insere_Grupo_Combinacoes

    lConexao = GL_lConexaoDic

    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 8244

    lComando2 = Comando_AbrirExt(lConexao)
    If lComando2 = 0 Then Error 8245

    'insere em GrupoRotina
    lErro = Comando_Executar(lComando2, "INSERT INTO GrupoRotinas (CodGrupo, SiglaRotina, TipoDeAcesso, LogAtividade, projeto, classe) SELECT ?, Sigla, ?, ?, ?, ? FROM Rotinas", sCodGrupo, COM_ACESSO, LOG_NAO, "", "")
    If lErro <> AD_SQL_SUCESSO Then Error 8247
            
''    'seleciona no BD todas as rotinas
''    sSiglaRot = String(STRING_ROTINA_SIGLA, 0)
''    lErro = Comando_Executar(lComando1, "SELECT Sigla FROM Rotinas", sSiglaRot)
''    If lErro <> AD_SQL_SUCESSO Then Error 8246
''
''    'le o grupo, se estiver no BD
''    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando1)
''
''    Do While iFim_de_Arquivo = AD_SQL_SUCESSO
''
''        'insere em GrupoRotina
''        lErro = Comando_Executar(lComando2, "INSERT INTO GrupoRotinas (CodGrupo, SiglaRotina, TipoDeAcesso, LogAtividade, projeto, classe) VALUES (?,?,?,?,?,?)", sCodGrupo, sSiglaRot, COM_ACESSO, LOG_NAO, "", "")
''        If lErro <> AD_SQL_SUCESSO Then Error 8247
''
''        iFim_de_Arquivo = Comando_BuscarProximo(lComando1)
''
''        If iFim_de_Arquivo = AD_SQL_ERRO Then Error 8248
''
''    Loop

    'insere em GrupoTela
    lErro = Comando_Executar(lComando2, "INSERT INTO GrupoTela (CodGrupo, NomeTela, TipoDeAcesso, projeto, classe) SELECT ?, Nome, ?, ?, ? FROM Telas", sCodGrupo, COM_ACESSO, "", "")
    If lErro <> AD_SQL_SUCESSO Then Error 8250

''    'seleciona no BD todas as telas
''    sNomeTela = String(STRING_NOME_TELA, 0)
''    lErro = Comando_Executar(lComando1, "SELECT Nome FROM Telas", sNomeTela)
''    If lErro <> AD_SQL_SUCESSO Then Error 8249
''
''    'le o grupo, se estiver no BD
''    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando1)
''
''    Do While iFim_de_Arquivo = AD_SQL_SUCESSO
''
''        'insere em GrupoTela
''        lErro = Comando_Executar(lComando2, "INSERT INTO GrupoTela (CodGrupo, NomeTela, TipoDeAcesso, projeto, classe) VALUES (?,?,?,?,?)", sCodGrupo, sNomeTela, COM_ACESSO, "", "")
''        If lErro <> AD_SQL_SUCESSO Then Error 8250
''
''        iFim_de_Arquivo = Comando_BuscarProximo(lComando1)
''
''        If iFim_de_Arquivo = AD_SQL_ERRO Then Error 8251
''
''    Loop

    'insere em GrupoRelatorios
    lErro = Comando_Executar(lComando2, "INSERT INTO GrupoRelatorios (CodGrupo, CodRel, TipoDeAcesso, NomeTskCustomizado) SELECT ? , CodRel, ?, ? FROM Relatorios", sCodGrupo, COM_ACESSO, "")
    If lErro <> AD_SQL_SUCESSO Then Error 32031

''    'seleciona no BD todos os relatorios
''    sNomeRelat = String(STRING_CODIGO_RELATORIO, 0)
''    lErro = Comando_Executar(lComando1, "SELECT CodRel FROM Relatorios", sNomeRelat)
''    If lErro <> AD_SQL_SUCESSO Then Error 32030
''
''    'le o grupo, se estiver no BD
''    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando1)
''
''    Do While iFim_de_Arquivo = AD_SQL_SUCESSO
''
''        'insere em GrupoRelatorios
''        lErro = Comando_Executar(lComando2, "INSERT INTO GrupoRelatorios (CodGrupo, CodRel, TipoDeAcesso, NomeTskCustomizado) VALUES (?,?,?,?)", sCodGrupo, sNomeRelat, COM_ACESSO, "")
''        If lErro <> AD_SQL_SUCESSO Then Error 32031
''
''        iFim_de_Arquivo = Comando_BuscarProximo(lComando1)
''
''        If iFim_de_Arquivo = AD_SQL_ERRO Then Error 32032
''
''    Loop

    'insere em GrupoBrowseCampo
    lErro = Comando_Executar(lComando2, "INSERT INTO GrupoBrowseCampo (CodGrupo, NomeTela, NomeArq, Nome) SELECT ?, NomeTela, NomeArq, Nome FROM GrupoBrowseCampo WHERE CodGrupo = 'supervisor'", sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then Error 32034

''    'seleciona no BD todos os campos p/browse
''    sNomeTela = String(STRING_NOME_TELA, 0)
''    sNomeArq = String(STRING_BROWSEARQUIVO_NOME_ARQ, 0)
''    sNomeCampo = String(STRING_CAMPOS_NOME, 0)
''    lErro = Comando_Executar(lComando1, "SELECT NomeTela, BrowseArquivo.NomeArq, Nome FROM BrowseArquivo, Campos WHERE BrowseArquivo.NomeArq = Campos.NomeArq", sNomeTela, sNomeArq, sNomeCampo)
''    If lErro <> AD_SQL_SUCESSO Then Error 32033
''
''    'le o grupo, se estiver no BD
''    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando1)
''
''    Do While iFim_de_Arquivo = AD_SQL_SUCESSO
''
''        'insere em GrupoBrowseCampo
''        lErro = Comando_Executar(lComando2, "INSERT INTO GrupoBrowseCampo (CodGrupo, NomeTela, NomeArq, Nome) VALUES (?,?,?,?)", sCodGrupo, sNomeTela, sNomeArq, sNomeCampo)
''        If lErro <> AD_SQL_SUCESSO Then Error 32034
''
''        iFim_de_Arquivo = Comando_BuscarProximo(lComando1)
''
''        If iFim_de_Arquivo = AD_SQL_ERRO Then Error 32035
''
''    Loop
    
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando2)

    Insere_Grupo_Combinacoes = SUCESSO

    Exit Function

Erro_Insere_Grupo_Combinacoes:

    Insere_Grupo_Combinacoes = Err

    Select Case Err

        Case 8244
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 8245
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 8246, 8248
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ROTINA", Err)

        Case 8247
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GRUPOROTINA", Err, sCodGrupo, sSiglaRot)

        Case 8249, 8251
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TELA", Err)

        Case 8250
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GRUPOTELA", Err, sCodGrupo, sNomeTela)

        Case 32030, 32032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELATORIO", Err)

        Case 32031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GRUPOREL", Err, sCodGrupo, sNomeRelat)

        Case 32033, 32035
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOS", Err)

        Case 32034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_GRUPOBROWSECAMPO", Err, sCodGrupo, sNomeTela)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159010)

    End Select

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Public Function Grupo_Exclui(ByVal sCodGrupo As String) As Long
'exclui um grupo e os usuarios associados a ele

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim vbMsgRet As VbMsgBoxResult
Dim colEmpGrupo As New Collection
Dim lConexao As Long
Dim sCodGrupo1 As String
Dim sCodUsuario As String
Dim colUsuarios As New Collection

On Error GoTo Erro_Grupo_Exclui

    sCodGrupo1 = String(STRING_GRUPO_CODIGO, 0)
    
    lConexao = GL_lConexaoDic

    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 8268

    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 8269

    lComando2 = Comando_AbrirExt(lConexao)
    If lComando2 = 0 Then Error 54806
    
    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 8270

    'pesquisa o Grupo em questão no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT CodGrupo FROM GruposDeUsuarios WHERE CodGrupo=?", 0, sCodGrupo1, sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then Error 8271

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then
        If lErro = AD_SQL_SEM_DADOS Then
            Error 8272
        Else
            Error 8273
        End If
    End If

    sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    lErro = Comando_Executar(lComando2, "SELECT CodUsuario FROM Usuarios WHERE CodGrupo=?", sCodUsuario, sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then Error 54804
    
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 54805
    
    If lErro = AD_SQL_SUCESSO Then
        vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_USUARIOS", sCodGrupo)
        If vbMsgRet = vbNo Then Error 54807
    Else
        '??? deveria estar na tela e nao aqui. Na tela poderia testar se tem usuarios ligados ao grupo p/incluir na msg
        vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_GRUPO", sCodGrupo)
        If vbMsgRet = vbNo Then Error 8274
    End If
    
    'lock do grupo
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 8275
    
    'Exclui os usuarios do grupo
    lErro = Grupo_ExcluiUsuarios(sCodGrupo)
    If lErro <> SUCESSO Then Error 32054
    
    'Exclui os registros de GrupoRelatorios
    lErro = Grupo_ExcluiRelatorios(sCodGrupo)
    If lErro <> SUCESSO Then Error 32055

    'Exclui os registros de GrupoBrowseCampo
    lErro = Grupo_ExcluiBrowseCampo(sCodGrupo)
    If lErro <> SUCESSO Then Error 32056

    'Exclui os registros de GrupoRotina
    lErro = Grupo_ExcluiRotina(sCodGrupo)
    If lErro <> SUCESSO Then Error 8278

    'Exclui os registros de GrupoRotina
    lErro = Grupo_ExcluiTela(sCodGrupo)
    If lErro <> SUCESSO Then Error 8279

    'exclui o grupo do BD
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM GruposDeUsuarios", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 8280

    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 8281

    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando2)

    Grupo_Exclui = SUCESSO

    Exit Function

Erro_Grupo_Exclui:

    Grupo_Exclui = Err

    Select Case Err

        Case 8274, 8276, 8277, 8278, 8279, 8298, 32054, 32055, 32056, 54807

        Case 8268, 54806
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 8269
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 8270
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 8271, 8273
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO", Err, sCodGrupo)

        Case 8272
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPO_NAO_CADASTRADO", Err, sCodGrupo)

        Case 8275
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_GRUPO", Err, sCodGrupo)

        Case 8280
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GRUPO", Err, sCodGrupo)

        Case 8281
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 54804, 54805
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOS_DIC", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159011)

    End Select

    lErro = Transacao_RollbackDic
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando2)
    
    Exit Function

End Function

Private Function Grupo_ExcluiRotina(ByVal sCodGrupo As String) As Long
'exclui registros da tabela GrupoRotina de um determinado grupo

Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sCodGrupo1 As String
Dim lComando As Long
Dim lComando1 As Long
Dim lConexao As Long

On Error GoTo Erro_Grupo_ExcluiRotina

    sCodGrupo1 = String(STRING_GRUPO_CODIGO, 0)

    lConexao = GL_lConexaoDic

    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 8299

    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 8300

    sCodGrupo1 = String(STRING_GRUPO_CODIGO, 0)

    'Pesquisa em GrupoRotinas os registros com o código do Grupo
    lErro = Comando_ExecutarPos(lComando, "SELECT CodGrupo FROM GrupoRotinas WHERE CodGrupo=?", 0, sCodGrupo1, sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then Error 8301

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)

    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'exclui os registros pesquisados acima
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM GrupoRotinas", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 8302

        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop

    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)

    Grupo_ExcluiRotina = SUCESSO

    Exit Function

Erro_Grupo_ExcluiRotina:

    Grupo_ExcluiRotina = Err

    Select Case Err

        Case 8299
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 8300
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 8301
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_ROTINAS", Err, sCodGrupo)

        Case 8302
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_ROTINAS", Err, sCodGrupo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159012)

    End Select

    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)
    
    Exit Function

End Function

Private Function Grupo_ExcluiTela(ByVal sCodGrupo As String) As Long
'exclui registros da tabela GrupoTela de um determinado grupo

Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim sCodGrupo1 As String
Dim lComando As Long
Dim lComando1 As Long
Dim lConexao As Long

On Error GoTo Erro_Grupo_ExcluiTela

    sCodGrupo1 = String(STRING_GRUPO_CODIGO, 0)
    
    lConexao = GL_lConexaoDic

    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 8303

    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 8304

    'Pesquisa em GrupoTela os registros com o código do Grupo
    lErro = Comando_ExecutarPos(lComando, "SELECT CodGrupo FROM GrupoTela WHERE CodGrupo=?", 0, sCodGrupo1, sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then Error 8305

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando)

    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'exclui os registros pesquisados acima
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM GrupoTela", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 8306

        iFim_de_Arquivo = Comando_BuscarProximo(lComando)

    Loop

    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)

    Grupo_ExcluiTela = SUCESSO

    Exit Function

Erro_Grupo_ExcluiTela:

    Grupo_ExcluiTela = Err

    Select Case Err

        Case 8303
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 8304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 8305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_TELA", Err, sCodGrupo)

        Case 8306
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GRUPOTELA", Err, sCodGrupo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159013)

    End Select

    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)
    
    Exit Function

End Function

Private Function UsuFilEmp_Exclui_Usuario(sCodUsuario As String) As Long
'exclui os registros da tabela UsuFilEmp para um determinado usuario
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long, lComando As Long

On Error GoTo Erro_UsuFilEmp_Exclui_Usuario

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 32037
    
    lErro = Comando_Executar(lComando, "DELETE FROM UsuFilEmp WHERE CodUsuario = ?", sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 32038
    
    Call Comando_Fechar(lComando)
    
    UsuFilEmp_Exclui_Usuario = SUCESSO
     
    Exit Function
    
Erro_UsuFilEmp_Exclui_Usuario:

    UsuFilEmp_Exclui_Usuario = Err
     
    Select Case Err
          
        Case 32037
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 32038
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_USUFILEMP_USU", Err, sCodUsuario)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159014)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function UsuFilEmp_Grava_Usuario(sCodUsuario As String, colUsuFilEmp As Collection) As Long
'exclui os registros da tabela UsuFilEmp para um determinado usuario
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long, lComando As Long
Dim objUsuFilEmp As ClassUsuFilEmp

On Error GoTo Erro_UsuFilEmp_Grava_Usuario

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 32046
    
    For Each objUsuFilEmp In colUsuFilEmp
    
        lErro = Comando_Executar(lComando, "INSERT INTO UsuFilEmp (CodUsuario,CodEmpresa,CodFilial) VALUES (?,?,?)", sCodUsuario, objUsuFilEmp.lCodEmpresa, objUsuFilEmp.iCodFilial)
        If lErro <> AD_SQL_SUCESSO Then Error 32047
        
    Next
    
    Call Comando_Fechar(lComando)
    
    UsuFilEmp_Grava_Usuario = SUCESSO
     
    Exit Function
    
Erro_UsuFilEmp_Grava_Usuario:

    UsuFilEmp_Grava_Usuario = Err
     
    Select Case Err
          
        Case 32046
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 32047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_USUFILEMP_USU", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159015)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Grupo_ExcluiUsuarios(sCodGrupo As String) As Long
'exclui todos os usuarios do grupo

Dim lErro As Long, iIndice As Integer, iFim_de_Arquivo As Integer
Dim alComando(1 To 3) As Long, sCodUsuario As String

On Error GoTo Erro_Grupo_ExcluiUsuarios

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then Error 32052
        
    Next
        
    'seleciona no BD todos os códigos de Usuarios pertencentes ao Grupo sCodGrupo
    sCodUsuario = String(STRING_USUARIO_CODIGO, 0)
    lErro = Comando_Executar(alComando(1), "SELECT CodUsuario FROM Usuarios WHERE CodGrupo=?", sCodUsuario, sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then Error 6219

    iFim_de_Arquivo = Comando_BuscarPrimeiro(alComando(1))
    
    'Preenche a coleção
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO
        
        lErro = Usuario_Exclui_Batch(sCodUsuario, alComando(2), alComando(3))
        If lErro <> SUCESSO Then Error 32053
        
        iFim_de_Arquivo = Comando_BuscarProximo(alComando(1))

    Loop
    
    For iIndice = LBound(alComando) To UBound(alComando)
    
        Call Comando_Fechar(alComando(iIndice))
        
    Next
    
    Grupo_ExcluiUsuarios = SUCESSO
     
    Exit Function
    
Erro_Grupo_ExcluiUsuarios:

    Grupo_ExcluiUsuarios = Err
     
    Select Case Err
          
        Case 32052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 32053
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159016)
     
    End Select
     
    For iIndice = LBound(alComando) To UBound(alComando)
    
        Call Comando_Fechar(alComando(iIndice))
        
    Next
    
    Exit Function

End Function

Private Function Usuario_Exclui_Batch(ByVal sCodUsuario As String, lComando As Long, lComando1 As Long) As Long
'efetua a exclusao de um usuario dentro de transacao

Dim lErro As Long
Dim sCodUsuario1 As String

On Error GoTo Erro_Usuario_Exclui_Batch

    If UCase(Trim(sCodUsuario)) = "SUPERVISOR" Then Error 6585

    sCodUsuario1 = String(STRING_USUARIO_CODIGO, 0)
    
    'pesquisa o Usuário em questão no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT CodUsuario FROM Usuarios WHERE CodUsuario=?", 0, sCodUsuario1, sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 6279
    
    'lê o Usuário
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 6280
    If lErro <> AD_SQL_SUCESSO Then Error 6281
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 32036
    
    'Exclui registros da Tabela BrowseUsuarioCampo
    lErro = BrowseUsuarioCampo_Exclui(sCodUsuario)
    If lErro <> SUCESSO Then Error 6583
    
    'Exclui registros da Tabela BrowseUsuarioOrdenacao
    lErro = BrowseUsuarioOrdenacao_Exclui(sCodUsuario)
    If lErro <> SUCESSO Then Error 6584
    
    lErro = Usuario_ExcluiBDsDados(sCodUsuario)
    If lErro <> SUCESSO Then Error 56846
    
    'Exclui registros da Tabela UsuFilEmp com código do Usuário
    lErro = UsuFilEmp_Exclui_Usuario(sCodUsuario)
    If lErro Then Error 6292
    
    'exclui o Usuário do BD
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM Usuarios", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 6282
    
    Usuario_Exclui_Batch = SUCESSO
    
    Exit Function
    
Erro_Usuario_Exclui_Batch:

    Usuario_Exclui_Batch = Err
    
    Select Case Err
    
        Case 6279, 6281
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIO", Err, sCodUsuario)
            
        Case 6280
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", Err, sCodUsuario)
        
        Case 6282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_USUARIO", Err, sCodUsuario)
        
        Case 6292, 6583, 6584, 32036, 56846  'Tratado na rotina chamada
        
        Case 6585
            Call Rotina_Erro(vbOKOnly, "ERRO_USU_SUPERVISOR_NAO_PODE_SER_EXC", Err)

            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159017)
    
    End Select
    
    Exit Function

End Function

Private Function Grupo_ExcluiBrowseCampo(sCodGrupo As String) As Long
'exclui todos os registros da tabela GrupoBrowseCampo associados a um grupo

Dim lErro As Long, lComando As Long

On Error GoTo Erro_Grupo_ExcluiBrowseCampo

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 32057
    
    lErro = Comando_Executar(lComando, "DELETE FROM GrupoBrowseCampo WHERE CodGrupo = ?", sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then Error 32058
    
    Call Comando_Fechar(lComando)
    
    Grupo_ExcluiBrowseCampo = SUCESSO
     
    Exit Function
    
Erro_Grupo_ExcluiBrowseCampo:

    Grupo_ExcluiBrowseCampo = Err
     
    Select Case Err
          
        Case 32057
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 32058
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GRUPOBROWSECAMPO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159018)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Grupo_ExcluiRelatorios(sCodGrupo As String) As Long
'exclui todos os registros da tabela GrupoRelatorios associados a um grupo

Dim lErro As Long, lComando As Long

On Error GoTo Erro_Grupo_ExcluiRelatorios

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 32059
    
    lErro = Comando_Executar(lComando, "DELETE FROM GrupoRelatorios WHERE CodGrupo = ?", sCodGrupo)
    If lErro <> AD_SQL_SUCESSO Then Error 32060
    
    Call Comando_Fechar(lComando)
    
    Grupo_ExcluiRelatorios = SUCESSO
     
    Exit Function
    
Erro_Grupo_ExcluiRelatorios:

    Grupo_ExcluiRelatorios = Err
     
    Select Case Err
          
        Case 32059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 32060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_GRUPOBROWSECAMPO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159019)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function BrowseUsuarioCampo_GrupoAlterado(sGrupo, sTela, sArquivo, colCampo) As Long
'faz com que a nivel de usuario seja tratada a alteracao dos campos visiveis p/o grupo de usuarios
'poderia excluir campos nao visiveis e manter outros atributos mas para simplificar vou apenas excluir a
'definicao dos usuarios do grupo

Dim lComando As Long
Dim lErro As Long

On Error GoTo Erro_BrowseUsuarioCampo_GrupoAlterado

    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 32062
    
    'exclui os registros de BrowseUsuarioCampo que tem código do Usuário
    If giSQLTipoOrdParamDic = SQL_ORD_PARAM_ESQ_DIR Then
        lErro = Comando_Executar(lComando, "DELETE FROM BrowseUsuarioCampo WHERE NomeTela = ? AND NomeArq = ? AND CodUsuario IN (SELECT CodUSuario FROM Usuarios WHERE CodGrupo = ?)", sTela, sArquivo, sGrupo)
    Else
        lErro = Comando_Executar(lComando, "DELETE FROM BrowseUsuarioCampo WHERE NomeTela = ? AND NomeArq = ? AND CodUsuario IN (SELECT CodUSuario FROM Usuarios WHERE CodGrupo = ?)", sGrupo, sTela, sArquivo)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 32063

    lErro = Comando_Fechar(lComando)

    BrowseUsuarioCampo_GrupoAlterado = SUCESSO
    
    Exit Function
    
Erro_BrowseUsuarioCampo_GrupoAlterado:

    BrowseUsuarioCampo_GrupoAlterado = Err
    
    Select Case Err
    
        Case 32062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
         
        Case 32063
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_BROWSEUSUARIOCAMPO3", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159020)
    
    End Select
    
    lErro = Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function GrupoRelatorio_Grava2(ByVal sModulo As String, ByVal colGrupoRelatorio As Collection) As Long
'Atualiza registros em GrupoRelatorios de varios relatorios de um par (grupo,Modulo)
'Membros de colGrupoRelatorio têm o mesmo CodGrupo e os relatorios pertencem a sModulo.

Dim lConexao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sCodGrupo As String, sCodGrupoAux As String
Dim sCodRel As String
Dim iIndice As Integer
Dim sNomeTskBD As String, iTipoAcessoBD As Integer, objGrupoRelatorio As ClassDicGrupoRelatorio

On Error GoTo Erro_GrupoRelatorio_Grava2

    lConexao = GL_lConexaoDic
    
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 32064
    
    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 32065

    'Inicia a Transação
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 32066
    
    sCodGrupo = colGrupoRelatorio.Item(1).sCodGrupo
    sCodGrupoAux = String(STRING_GRUPO_CODIGO, 0)
    sCodRel = String(STRING_CODIGO_RELATORIO, 0)
    sNomeTskBD = String(STRING_NOME_ARQ_COMPLETO, 0)
    
    'Pesquisa no BD os registros de GrupoRelatorios associados a sCodGrupo e a relatorios em sModulo
    If giSQLTipoOrdParamDic = SQL_ORD_PARAM_ESQ_DIR Then
        lErro = Comando_ExecutarPos(lComando, "SELECT CodGrupo, CodRel, TipoDeAcesso, NomeTskCustomizado FROM GrupoRelatorios WHERE CodGrupo=? AND CodRel IN (SELECT CodRel FROM RelatoriosModulo, Modulos WHERE RelatoriosModulo.Modulo=Modulos.Sigla AND Modulos.Nome=?) ORDER BY CodRel", 0, sCodGrupoAux, sCodRel, iTipoAcessoBD, sNomeTskBD, sCodGrupo, sModulo)
    Else
        lErro = Comando_ExecutarPos(lComando, "SELECT CodGrupo, CodRel, TipoDeAcesso, NomeTskCustomizado FROM GrupoRelatorios WHERE CodGrupo=? AND CodRel IN (SELECT CodRel FROM RelatoriosModulo, Modulos WHERE RelatoriosModulo.Modulo=Modulos.Sigla AND Modulos.Nome=?) ORDER BY CodRel", 0, sCodGrupoAux, sCodRel, iTipoAcessoBD, sNomeTskBD, sModulo, sCodGrupo)
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 32067

    'Lê registro de GrupoRelatorios, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 32068
    If lErro <> AD_SQL_SUCESSO Then Error 32069
    
    iIndice = 1  'iIndice é o índice na coleção colGrupoRotina
    
    'Loop de Atualização de GrupoRotina
    Do While lErro = AD_SQL_SUCESSO
               
        'teste se iIndice é válido
        If iIndice > colGrupoRelatorio.Count Then Error 32070
        
        Set objGrupoRelatorio = colGrupoRelatorio.Item(iIndice)
        
        'testa se sCodRel lido no BD coincide com sCodRel na coleção
        If sCodRel <> objGrupoRelatorio.sCodRel Then Error 32071
        
        'se trocou o valor de algum atributo
        If objGrupoRelatorio.iTipoDeAcesso <> iTipoAcessoBD Or objGrupoRelatorio.sNomeTskCustomizado <> sNomeTskBD Then
        
            'atualiza o registro de GrupoRelatorios
            lErro = Comando_ExecutarPos(lComando1, "UPDATE GrupoRelatorios SET TipoDeAcesso=?,NomeTskCustomizado=?", lComando, objGrupoRelatorio.iTipoDeAcesso, objGrupoRelatorio.sNomeTskCustomizado)
            If lErro <> AD_SQL_SUCESSO Then Error 32072
        
        End If
        
        'vai para o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        iIndice = iIndice + 1
      
    Loop
    
    If lErro <> AD_SQL_SEM_DADOS Then Error 32073
    
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 32074

    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)

    GrupoRelatorio_Grava2 = SUCESSO

    Exit Function

Erro_GrupoRelatorio_Grava2:

    GrupoRelatorio_Grava2 = Err

    Select Case Err

        Case 32064, 32065
           lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 32066
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
           
        Case 32067, 32069, 32073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_RELATORIOS", Err)
        
        Case 32068
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPORELATORIO_NAO_CADASTRADO", Err)
            
        Case 32070, 32071
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODREL_DIFERENTE", Err)
        
        Case 32072
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_GRUPO_RELATORIOS", Err, sCodRel, sCodGrupo)
 
        Case 32074
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159021)

    End Select

    lErro = Transacao_RollbackDic
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Public Function GrupoRelatorio_Grava(ByVal colGrupoRelatorio As Collection) As Long
'Atualiza registros em GrupoRelatorios de varios grupos de um relatorio
'Membros de colGrupoRelatorio têm o mesmo CodRel.

Dim lConexao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sCodGrupo As String
Dim sCodRel As String
Dim iIndice As Integer
Dim sNomeTskBD As String, iTipoAcessoBD As Integer, objGrupoRelatorio As ClassDicGrupoRelatorio

On Error GoTo Erro_GrupoRelatorio_Grava

    lConexao = GL_lConexaoDic
    
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 32075
    
    lComando1 = Comando_AbrirExt(lConexao)
    If lComando1 = 0 Then Error 32076

    'Inicia a Transação
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 32077
    
    sCodRel = colGrupoRelatorio.Item(1).sCodRel
    sCodGrupo = String(STRING_GRUPO_CODIGO, 0)
    sNomeTskBD = String(STRING_NOME_ARQ_COMPLETO, 0)
    
    'Pesquisa no BD os registros de GrupoRelatorios ligados à sCodRel
    lErro = Comando_ExecutarPos(lComando, "SELECT CodGrupo, TipoDeAcesso, NomeTskCustomizado FROM GrupoRelatorios WHERE CodRel=? ORDER BY CodGrupo", 0, sCodGrupo, iTipoAcessoBD, sNomeTskBD, sCodRel)
    If lErro <> AD_SQL_SUCESSO Then Error 32078

    'Lê registro de GrupoRelatorios, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro = AD_SQL_SEM_DADOS Then Error 32079
    If lErro <> AD_SQL_SUCESSO Then Error 32080
    
    iIndice = 1  'iIndice é o índice na coleção colGrupoRelatorio
    
    'Loop de Atualização de GrupoRelatorios
    Do While lErro = AD_SQL_SUCESSO
               
        'teste se iIndice é válido
        If iIndice > colGrupoRelatorio.Count Then Error 32081
        
        Set objGrupoRelatorio = colGrupoRelatorio.Item(iIndice)
        
        'testa se CodGrupo lido no BD coincide com CodGrupo na coleção
        If sCodGrupo <> objGrupoRelatorio.sCodGrupo Then Error 32082
        
        'se trocou o valor de algum atributo
        If objGrupoRelatorio.iTipoDeAcesso <> iTipoAcessoBD Or objGrupoRelatorio.sNomeTskCustomizado <> sNomeTskBD Then
        
            'atualiza o registro de GrupoRelatorios
            lErro = Comando_ExecutarPos(lComando1, "UPDATE GrupoRelatorios SET TipoDeAcesso=?,NomeTskCustomizado=?", lComando, objGrupoRelatorio.iTipoDeAcesso, objGrupoRelatorio.sNomeTskCustomizado)
            If lErro <> AD_SQL_SUCESSO Then Error 32083
        
        End If
        
        'vai para o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        iIndice = iIndice + 1
      
    Loop
    
    If lErro <> AD_SQL_SEM_DADOS Then Error 32084
    
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 32085

    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)

    GrupoRelatorio_Grava = SUCESSO

    Exit Function

Erro_GrupoRelatorio_Grava:

    GrupoRelatorio_Grava = Err

    Select Case Err

        Case 32075, 32076
           lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 32077
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
           
        Case 32078, 32080, 32084
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_GRUPO_RELATORIOS", Err)
        
        Case 32079
            lErro = Rotina_Erro(vbOKOnly, "ERRO_GRUPORELATORIO_NAO_CADASTRADO", Err)
            
        Case 32081, 32082
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODREL_DIFERENTE", Err)
        
        Case 32083
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_GRUPO_RELATORIOS", Err, sCodRel, sCodGrupo)
 
        Case 32085
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159022)

    End Select

    lErro = Transacao_RollbackDic
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando)

    Exit Function
    
End Function

Public Function FilialEmpresa_Grava(objFilialEmpresa As AdmFiliais) As Long
'inclui ou altera dados de uma filial de uma Empresa, com efeitos no dic e no contab

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_FilialEmpresa_Grava
    
    'abrir transacao no dic
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 41846
    
    'Grava Filial Empresa no BD
    lErro = FilialEmpresa_Grava_EmTrans(objFilialEmpresa)
    If lErro <> SUCESSO Then Error 25936
        
    'Confirma Transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 41849
    
    FilialEmpresa_Grava = SUCESSO
     
    Exit Function
    
Erro_FilialEmpresa_Grava:

    FilialEmpresa_Grava = Err
     
    Select Case Err
          
        Case 25936 'erro tratado pela rotina chamada
        
        Case 41846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 41849
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159023)
     
    End Select
     
    'fazer rollback
    Call Transacao_RollbackDic
     
    Exit Function

End Function
Public Function FilialEmpresa_Grava_EmTrans(objFilialEmpresa As AdmFiliais) As Long
'inclui ou altera dados de uma filial de uma Empresa, com efeitos no dic e no contab

Dim lErro As Long
Dim iIndice As Integer
Dim alComandoDic(1 To 10) As Long

On Error GoTo Erro_FilialEmpresa_Grava_EmTrans

    'abrir comandos no dic
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
    
        alComandoDic(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComandoDic(iIndice) = 0 Then Error 41845
        
    Next
    
    lErro = FilialEmpresa_Grava_Dic(objFilialEmpresa, alComandoDic)
    If lErro <> SUCESSO Then Error 41847
    
    lErro = FilialEmpresa_Grava_Contab(objFilialEmpresa)
    If lErro <> SUCESSO Then Error 41848
        
    'fechar comandos
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    FilialEmpresa_Grava_EmTrans = SUCESSO
     
    Exit Function
    
Erro_FilialEmpresa_Grava_EmTrans:

    FilialEmpresa_Grava_EmTrans = Err
     
    Select Case Err
          
        Case 41845
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 41847, 41848, 62386, 62389, 62415
        
        Case 62387, 62388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESAS", Err)
            
        Case 62390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIMITEFILIAIS_ATINGIDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159024)
     
    End Select
     
    'fechar comandos
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
    
        Call Comando_Fechar(alComandoDic(iIndice))
        
    Next
    
    Exit Function

End Function

Private Function FilialEmpresa_TestaLimite(lComando As Long, lComando1 As Long) As Long
'verifica se pode ser criada mais um filial ou se já atingiu o limite de filiais

Dim lErro As Long, iLimite As Integer, iQtdeFiliaisAtivas As Integer

On Error GoTo Erro_FilialEmpresa_TestaLimite

    lErro = Comando_Executar(lComando, "SELECT LimiteFiliais FROM DicConfig", iLimite)
    If lErro <> AD_SQL_SUCESSO Then Error 41821
            
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41822
    If lErro <> AD_SQL_SUCESSO Then Error 41823
    
    'se nao estiver "ilimitado"
    If iLimite <> -1 Then
    
        lErro = Comando_Executar(lComando1, "SELECT COUNT(*) FROM FiliaisEmpresas WHERE Inativa = ?", iQtdeFiliaisAtivas, Ativo)
        If lErro <> AD_SQL_SUCESSO Then Error 41824
                
        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41825
        If lErro = AD_SQL_SUCESSO Then Error 41826
        
        'se ja existem filiais demais
        If iQtdeFiliaisAtivas >= iLimite Then Error 41827
        
    End If
    
    FilialEmpresa_TestaLimite = SUCESSO
     
    Exit Function
    
Erro_FilialEmpresa_TestaLimite:

    FilialEmpresa_TestaLimite = Err
     
    Select Case Err
          
        Case 41821, 41822, 41823
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DICCONFIG", Err)
                    
        Case 41824, 41825, 41826
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESA", Err)
        
        Case 41827
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIMITE_FILIAISEMPRESA", Err, iLimite)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159025)
     
    End Select
     
    Exit Function

End Function

Private Function FilialEmpresa_Conecta_Contab(objFilialEmpresa As AdmFiliais) As Long
'abre a conexao com o banco de dados da Empresa

Dim lErro As Long, sNomeEmpresa As String, lComandoDic As Long

On Error GoTo Erro_FilialEmpresa_Conecta_Contab

    lComandoDic = Comando_AbrirExt(GL_lConexaoDic)
    If lErro <> AD_SQL_SUCESSO Then Error 41828
    
    'obter nome da empresa
    sNomeEmpresa = String(STRING_EMPRESA_NOME, 0)
    lErro = Comando_Executar(lComandoDic, "SELECT Nome FROM Empresas WHERE Codigo = ?", sNomeEmpresa, objFilialEmpresa.lCodEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 41829
    
    lErro = Comando_BuscarProximo(lComandoDic)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41830
    If lErro = AD_SQL_SEM_DADOS Then Error 41831
    
    'para poder fazer transacao no bd da empresa
    lErro = Sistema_DefEmpresa(sNomeEmpresa, objFilialEmpresa.lCodEmpresa, objFilialEmpresa.sNome, objFilialEmpresa.iCodFilial)
    If lErro <> AD_BOOL_TRUE Then Error 41832
    
    Call Comando_Fechar(lComandoDic)
    
    FilialEmpresa_Conecta_Contab = SUCESSO
     
    Exit Function
    
Erro_FilialEmpresa_Conecta_Contab:

    FilialEmpresa_Conecta_Contab = Err
     
    Select Case Err
          
        Case 41828
           lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 41829, 41830, 41832
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA", Err, objFilialEmpresa.lCodEmpresa)
        
        Case 41831
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159026)
     
    End Select
     
    Call Comando_Fechar(lComandoDic)
    
    Exit Function

End Function

Private Function FilialEmpresa_Grava_Contab(objFilialEmpresa As AdmFiliais) As Long
'grava a parte da filial empresa no Contab (bd fora do dic especifico de cada Empresa)

Dim lNumIntEndereco As Long, lNumIntEnderecoEntrega As Long, lNumIntEnderecoCTB As Long, lTransacao As Long
Dim lErro As Long, alComando(1 To 8) As Long, iIndice As Integer

On Error GoTo Erro_FilialEmpresa_Grava_Contab
   
    'abre a conexao com o banco de dados da Empresa
    lErro = FilialEmpresa_Conecta_Contab(objFilialEmpresa)
    If lErro <> SUCESSO Then Error 41833
    
    'abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 41834
    Next
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 41835
    
    'pesquisa filial no bd
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Endereco, EnderecoEntrega,EnderecoContador FROM FiliaisEmpresa WHERE FilialEmpresa = ?", 0, lNumIntEndereco, lNumIntEnderecoEntrega, lNumIntEnderecoCTB, objFilialEmpresa.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 41836
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41837
    
    'se encontrou é alteracao
    If lErro = AD_SQL_SUCESSO Then
    
        If lNumIntEnderecoCTB = 0 Then
            lErro = CF("Endereco_Automatico", lNumIntEnderecoCTB)
            If lErro <> SUCESSO Then Error 41840
        End If
    
        'altera conteudo do registro na tabela FiliaisEmpresa
        With objFilialEmpresa
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE FiliaisEmpresa SET Nome = ?, CGC = ?, AlmoxarifadoPadrao = ?, InscricaoEstadual = ?, InscricaoMunicipal = ?, ICMSPorEstimativa = ?, ISSPercPadrao = ?, ISSCodigoPadrao = ?, ISSIncluso = ?, ContribuinteIPI = ?, IRPercPadrao = ?, RamoEmpresa = ?, Jucerja = ?, DataJucerja = ?, Contador = ?, CRCContador = ?, CPFContador = ?, TipoTribAtividade = ?, SimplesFederal = ?, SimplesFederalAliq = ?, SimplesFederalTeto = ?, SuperSimples = ?, PISNaoCumulativo = ?, COFINSNaoCumulativo = ?, LucroPresumido = ?, CertificadoA1A3=?, NFeAmbiente=?, CNAE = ?, " & _
            "SignatarioCTB = ?, CPFSignatarioCTB = ?, CodQualiSigCTB = ?, SpedFiscalPerfil = ?,NomeReduzido = ?,CNPJContador = ?, EnderecoContador = ?, RPSAmbiente = ?, NFCECSC = ?, idNFCECSC = ?, indSincPadrao = ?, RegimeTrib = ?, RegimeEspecialTrib = ?", alComando(1), .sNome, .sCgc, .iAlmoxarifadoPadrao, .sInscricaoEstadual, .sInscricaoMunicipal, .iICMSPorEstimativa, .dISSPercPadrao, .sISSCodigoPadrao, .iISSIncluso, .iContribuinteIPI, .dIRPercPadrao, .sRamoEmpresa, .sJucerja, .dtDataJucerja, .sContador, .sCRCContador, .sCPFContador, .iTipoTribAtividade, .iSimplesFederal, .dSimplesFederalAliq, .dSimplesFederalTeto, .iSuperSimples, .iPISNaoCumulativo, _
            .iCOFINSNaoCumulativo, .iLucroPresumido, .sCertificadoA1A3, .iNFeAmbiente, .sCNAE, .sSignatarioCTB, .sCPFSignatarioCTB, .sCodQualiSigCTB, .sSpedFiscalPerfil, .sNomeReduzido, .sCNPJContador, lNumIntEnderecoCTB, .iRPSAmbiente, .sNFCECSC, .sidNFCECSC, .iindSincPadrao, .iRegimeTrib, .iRegimeEspecialTrib)
        End With
        If lErro <> AD_SQL_SUCESSO Then Error 41838
    
        lErro = CF("SpedCadastro_Exclui", objFilialEmpresa.iCodFilial, alComando(3), alComando(4))
        If lErro <> SUCESSO Then Error 41842
            
        lErro = CF("ContribuinteSubstituto_Exclui", objFilialEmpresa.iCodFilial, alComando(7), alComando(8))
        If lErro <> SUCESSO Then Error 41842
            
    Else 'inclusao
    
        lErro = CF("Endereco_Automatico", lNumIntEndereco)
        If lErro <> SUCESSO Then Error 41839
        
        lErro = CF("Endereco_Automatico", lNumIntEnderecoEntrega)
        If lErro <> SUCESSO Then Error 41840
        
        lErro = CF("Endereco_Automatico", lNumIntEnderecoCTB)
        If lErro <> SUCESSO Then Error 41840
        
        With objFilialEmpresa
            lErro = Comando_Executar(alComando(2), "INSERT INTO FiliaisEmpresa (FilialEmpresa, Nome, CGC, Endereco, EnderecoEntrega, AlmoxarifadoPadrao,InscricaoEstadual, InscricaoMunicipal, ICMSPorEstimativa, ISSPercPadrao, ISSCodigoPadrao, ISSIncluso, ContribuinteIPI, IRPercPadrao, RamoEmpresa, Jucerja, DataJucerja, Contador, CRCContador, CPFContador,TipoTribAtividade, SimplesFederal, SimplesFederalAliq, SimplesFederalTeto, SuperSimples, PISNaoCumulativo, COFINSNaoCumulativo, LucroPresumido, CertificadoA1A3, NFeAmbiente, CNAE,SignatarioCTB, CPFSignatarioCTB,CodQualiSigCTB,EnderecoContador,SpedFiscalPerfil,NomeReduzido,CNPJContador, RPSAmbiente, NFCECSC, idNFCECSC, indSincPadrao, RegimeTrib,RegimeEspecialTrib ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                .iCodFilial, .sNome, .sCgc, lNumIntEndereco, lNumIntEnderecoEntrega, .iAlmoxarifadoPadrao, .sInscricaoEstadual, .sInscricaoMunicipal, .iICMSPorEstimativa, .dISSPercPadrao, .sISSCodigoPadrao, .iISSIncluso, .iContribuinteIPI, .dIRPercPadrao, .sRamoEmpresa, .sJucerja, .dtDataJucerja, .sContador, .sCRCContador, .sCPFContador, .iTipoTribAtividade, .iSimplesFederal, .dSimplesFederalAliq, .dSimplesFederalTeto, .iSuperSimples, .iPISNaoCumulativo, .iCOFINSNaoCumulativo, .iLucroPresumido, .sCertificadoA1A3, .iNFeAmbiente, .sCNAE, .sSignatarioCTB, .sCPFSignatarioCTB, .sCodQualiSigCTB, lNumIntEnderecoCTB, .sSpedFiscalPerfil, .sNomeReduzido, .sCNPJContador, .iRPSAmbiente, .sNFCECSC, .sidNFCECSC, .iindSincPadrao, .iRegimeTrib, .iRegimeEspecialTrib)
        End With
        If lErro <> AD_SQL_SUCESSO Then Error 41841
        
        If objFilialEmpresa.iCodFilial <> FILIAL_MATRIZ Then
            lErro = CF("FilialEmpresa_Replica_PreCadastros_Matriz", objFilialEmpresa.iCodFilial)
            If lErro <> SUCESSO Then Error 41840
        End If
    End If
    
    lErro = CF("SpedCadastro_Insere", objFilialEmpresa, alComando(5))
    If lErro <> SUCESSO Then Error 41842
    
    lErro = CF("ContribuinteSubstituto_Insere", objFilialEmpresa, alComando(6))
    If lErro <> SUCESSO Then Error 41842
    
    objFilialEmpresa.objEndereco.lCodigo = lNumIntEndereco
    lErro = CF("Endereco_Grava", objFilialEmpresa.objEndereco, False)
    If lErro <> SUCESSO Then Error 41842
        
    objFilialEmpresa.objEnderecoEntrega.lCodigo = lNumIntEnderecoEntrega
    If objFilialEmpresa.objEnderecoEntrega.sSiglaEstado = "" Then objFilialEmpresa.objEnderecoEntrega.sSiglaEstado = objFilialEmpresa.objEndereco.sSiglaEstado
    lErro = CF("Endereco_Grava", objFilialEmpresa.objEnderecoEntrega, False)
    If lErro <> SUCESSO Then Error 41843
    
    objFilialEmpresa.objEnderecoContador.lCodigo = lNumIntEnderecoCTB
    If objFilialEmpresa.objEnderecoContador.sSiglaEstado = "" Then objFilialEmpresa.objEnderecoContador.sSiglaEstado = objFilialEmpresa.objEndereco.sSiglaEstado
    If objFilialEmpresa.objEnderecoContador.iCodigoPais = 0 Then objFilialEmpresa.objEnderecoContador.iCodigoPais = objFilialEmpresa.objEndereco.iCodigoPais
    lErro = CF("Endereco_Grava", objFilialEmpresa.objEnderecoContador, False)
    If lErro <> SUCESSO Then Error 41842
    
    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 41844
    
    'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    FilialEmpresa_Grava_Contab = SUCESSO
     
    Exit Function
    
Erro_FilialEmpresa_Grava_Contab:

    FilialEmpresa_Grava_Contab = Err
     
    Select Case Err
          
        Case 41833, 41389, 41840, 41842, 41843
        
        Case 41834
           lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 41835
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 41836, 41837
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALEMPRESA", Err, objFilialEmpresa.iCodFilial)
        
        Case 41838
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_FILIALEMPRESA", Err, objFilialEmpresa.iCodFilial)
        
        Case 41841
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FILIALEMPRESA", Err, objFilialEmpresa.iCodFilial)
        
        Case 41844
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159027)
     
    End Select
     
    'Faz rollback na transação
    Call Transacao_Rollback
     
    'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function FilialEmpresa_Grava_Dic(objFilialEmpresa As AdmFiliais, alComandoDic() As Long) As Long
'inclui ou altera dados de uma filial de uma Empresa, com efeitos no dic

Dim lErro As Long, iInativa As Integer, sNomeFilial As String
Dim colModulosTrab As New Collection, vsSiglaModulo As Variant
Dim iIndice As Integer
Dim iTotal As Integer
Dim objDicConfig As New ClassDicConfig
Dim sCgc As String, sNomeEmpresa As String
Dim iNumeroLogs As Integer, iNumeroEmpresas As Integer
Dim iNumeroFiliais As Integer, colModulosLib As New Collection
Dim dtDataValidade As Date, sTextoSenha As String
Dim objDicSelect As New DicSelect

On Error GoTo Erro_FilialEmpresa_Grava_Dic

    'cria cópia da colecao de siglas de modulos
    For Each vsSiglaModulo In objFilialEmpresa.colModulos
        colModulosTrab.Add vsSiglaModulo
    Next
        
    'pesquisar por nome filialempresa com codigo <> do passado
    lErro = Comando_Executar(alComandoDic(1), "SELECT Inativa FROM FiliaisEmpresas WHERE Nome = ? AND CodEmpresa = ? AND CodFilial <> ?", iInativa, objFilialEmpresa.sNome, objFilialEmpresa.lCodEmpresa, objFilialEmpresa.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 41850
    
    'se houver =>erro (nao pode haver duas filiais com mesmo nome)
    lErro = Comando_BuscarProximo(alComandoDic(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41851
    If lErro = AD_SQL_SUCESSO Then Error 41852
        
    'pesquisar filialempresa por codigo
    sNomeFilial = String(STRING_FILIAL_NOME, 0)
    lErro = Comando_ExecutarPos(alComandoDic(2), "SELECT Nome, Inativa FROM FiliaisEmpresas WHERE CodEmpresa = ? AND CodFilial = ?", 0, sNomeFilial, iInativa, objFilialEmpresa.lCodEmpresa, objFilialEmpresa.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 41853
        
    lErro = Comando_BuscarProximo(alComandoDic(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41854
    
    'se houver processa alteracao
    If lErro = AD_SQL_SUCESSO Then
        
        'se a filial está inativa nao pode alterar
        If iInativa = Inativo Then Error 41855
        
        lErro = Comando_LockExclusive(alComandoDic(2))
        If lErro <> AD_SQL_SUCESSO Then Error 41856
        
        'se trocou nome
        If objFilialEmpresa.sNome <> sNomeFilial Then
        
            'altera tabela de filialempresa
            lErro = Comando_ExecutarPos(alComandoDic(3), "UPDATE FiliaisEmpresas SET Nome = ?", alComandoDic(2), objFilialEmpresa.sNome)
            If lErro <> AD_SQL_SUCESSO Then Error 41857
            
        End If
        
        lErro = FilialEmpresa_Grava_Dic1(objFilialEmpresa, alComandoDic(), colModulosTrab)
        If lErro <> SUCESSO Then Error 41858
            
    Else
    
        'verifica se vai ultrapassar o limite de Filiais independente de Empresa (com inativa =0)
        'o limite de filiais é global p/Empresa e está em DicConfig
        lErro = Comando_Executar(alComandoDic(1), "SELECT COUNT(*) FROM FiliaisEmpresas WHERE Inativa = ? ", iTotal, Ativo)
        If lErro <> AD_SQL_SUCESSO Then Error 62387
        
        lErro = Comando_BuscarPrimeiro(alComandoDic(1))
        If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62388
        
        lErro = objDicSelect.DicConfig_Le(objDicConfig)
        If lErro <> SUCESSO Then Error 62386
            
        lErro = objDicSelect.Senha_Empresa_Decifra(objDicConfig.sSenha, sCgc, sNomeEmpresa, iNumeroLogs, iNumeroEmpresas, iNumeroFiliais, colModulosLib, dtDataValidade, sTextoSenha)
        If lErro <> SUCESSO Then Error 62389
        
        objDicConfig.dtValidadeAte = dtDataValidade
        
'        If giTipoVersao = VERSAO_LIGHT Then
'            If iTotal = 0 Then
'                lErro = Gera_Senha_Provisoria(Mid(sCgc, 1, 4), sNomeEmpresa)
'                If lErro <> SUCESSO Then Error 62415
'            End If
'        End If
        
        If iTotal = iNumeroFiliais Then Error 62390
        
        'Insere registro em FilialEmpresa
        lErro = Comando_Executar(alComandoDic(8), "INSERT INTO FiliaisEmpresas (CodEmpresa, CodFilial, Nome, Inativa) VALUES (?,?,?,?)", objFilialEmpresa.lCodEmpresa, objFilialEmpresa.iCodFilial, objFilialEmpresa.sNome, Ativo)
        If lErro <> AD_SQL_SUCESSO Then Error 41860
        
        'em usufilemp insere permissao p/supervisor acessar esta filialempresa
        lErro = Comando_Executar(alComandoDic(9), "INSERT INTO UsuFilEmp (CodUsuario, CodEmpresa, CodFilial) VALUES (?,?,?)", USU_SUP, objFilialEmpresa.lCodEmpresa, objFilialEmpresa.iCodFilial)
        If lErro <> AD_SQL_SUCESSO Then Error 41861
        
    End If
    
    'percorre a colecao copia incluindo registros com status de NAO_CONFIGURADO
    For Each vsSiglaModulo In colModulosTrab
        
        lErro = Comando_Executar(alComandoDic(10), "INSERT INTO ModuloFilEmp (CodEmpresa, CodFilial, SiglaModulo, Configurado) VALUES (?,?,?,?)", objFilialEmpresa.lCodEmpresa, objFilialEmpresa.iCodFilial, vsSiglaModulo, NAO_CONFIGURADO)
        If lErro <> AD_SQL_SUCESSO Then Error 41862
    
    Next
    
    FilialEmpresa_Grava_Dic = SUCESSO
     
    Exit Function
    
Erro_FilialEmpresa_Grava_Dic:

    FilialEmpresa_Grava_Dic = Err
     
    Select Case Err
          
        Case 41850, 41851
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESA", Err)
            
        Case 41852
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_MESMO_NOME", Err)
        
        Case 41853, 41854
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALEMPRESA", Err, objFilialEmpresa.iCodFilial)
        
        Case 41855
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_INATIVA", Err)
        
        Case 41856
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIALEMPRESA", Err, objFilialEmpresa.iCodFilial)
        
        Case 41857
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_FILIALEMPRESA", Err, objFilialEmpresa.iCodFilial)
        
        Case 41858, 41859, 62386, 62389, 62415
        
        Case 41860
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FILIALEMPRESA", Err, objFilialEmpresa.iCodFilial)
        
        Case 41861
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_USUFILEMP_USU", Err, USU_SUP)
        
        Case 41862
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MODULOFILEMP", Err)
        
        Case 62387, 62388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESAS", Err)
            
        Case 62390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIMITEFILIAIS_ATINGIDO", Err, iNumeroFiliais)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159028)
     
    End Select
     
    Exit Function

End Function

Private Function FilialEmpresa_Grava_Dic1(objFilialEmpresa As AdmFiliais, alComandoDic() As Long, colModulosTrab As Collection) As Long
'auxiliar a FilialEmpresa_Grava_Dic
'trata os registros da tabela ModuloFilEmp que já estavam associados a uma filial que está sendo alterada

Dim lErro As Long, sSiglaModulo As String, iIndice As Integer

On Error GoTo Erro_FilialEmpresa_Grava_Dic1

    'le regs de modulofilemp da filemp no bd
    sSiglaModulo = String(STRING_MODULO_SIGLA, 0)
    lErro = Comando_ExecutarPos(alComandoDic(4), "SELECT SiglaModulo FROM ModuloFilEmp WHERE CodEmpresa = ? AND CodFilial = ?", 0, sSiglaModulo, objFilialEmpresa.lCodEmpresa, objFilialEmpresa.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 41863
    
    lErro = Comando_BuscarProximo(alComandoDic(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41864
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        iIndice = FilialEmpresa_Grava_PesqModulo(colModulosTrab, sSiglaModulo)
        'se nao achou modfilemp na colecao
        If iIndice = -1 Then
        
            'exclui do bd
            lErro = Comando_ExecutarPos(alComandoDic(5), "DELETE FROM ModuloFilEmp", alComandoDic(4))
            If lErro <> AD_SQL_SUCESSO Then Error 41865
        
        Else
            
            'exclui da colecao
            colModulosTrab.Remove (iIndice)
            
        End If
        
        lErro = Comando_BuscarProximo(alComandoDic(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41866
    
    Loop

    FilialEmpresa_Grava_Dic1 = SUCESSO
     
    Exit Function
    
Erro_FilialEmpresa_Grava_Dic1:

    FilialEmpresa_Grava_Dic1 = Err
     
    Select Case Err
          
        Case 41863, 41864, 41866
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULOFILEMP", Err)
        
        Case 41865
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MODULOFILEMP", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159029)
     
    End Select
     
    Exit Function

End Function
        
Private Function FilialEmpresa_Grava_PesqModulo(colModulosTrab As Collection, sSiglaModulo As String) As Long
'obtem indice de sigla numa colecao de siglas
'retorna -1 para "nao encontrado"
Dim vsSiglaModulo As Variant, iIndice As Integer, iAchou As Integer

    iAchou = 0
    iIndice = 1
    
    For Each vsSiglaModulo In colModulosTrab
    
        If sSiglaModulo = vsSiglaModulo Then
        
            iAchou = 1
            Exit For
            
        End If
    
        iIndice = iIndice + 1
    
    Next
    
    If iAchou = 0 Then iIndice = -1
    
    FilialEmpresa_Grava_PesqModulo = iIndice
    
End Function

Public Function FilialEmpresa_Exclui(objFilialEmpresa As AdmFiliais) As Long
'exclui uma filial de empresa

Dim lErro As Long
Dim iIndice As Integer
Dim lTransacao As Long

On Error GoTo Erro_FilialEmpresa_Exclui

    'abrir transacao no dic
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 41875
    
    lErro = FilialEmpresa_Exclui_Trans(objFilialEmpresa)
    If lErro <> SUCESSO Then Error 41876
        
    'Confirma Transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 41877
    
    FilialEmpresa_Exclui = SUCESSO
     
    Exit Function
    
Erro_FilialEmpresa_Exclui:

    FilialEmpresa_Exclui = Err
     
    Select Case Err
          
        Case 41875
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 41876
        
        Case 41877
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159030)
     
    End Select
     
    'fazer rollback
    Call Transacao_RollbackDic
        
    Exit Function

End Function

Private Function FilialEmpresa_Exclui_Trans(objFilialEmpresa As AdmFiliais) As Long
'exclui uma filial de empresa
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long, iInativa As Integer
Dim iIndice As Integer, alComandoDic(1 To 3) As Long

On Error GoTo Erro_FilialEmpresa_Exclui_Trans

    'abrir comandos no dic
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
    
        alComandoDic(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComandoDic(iIndice) = 0 Then Error 41867
        
    Next
    
    'seleciona a filial no bd
    lErro = Comando_ExecutarPos(alComandoDic(1), "SELECT Inativa FROM FiliaisEmpresas WHERE CodEmpresa = ? AND CodFilial = ?", 0, iInativa, objFilialEmpresa.lCodEmpresa, objFilialEmpresa.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 41868
    
    lErro = Comando_BuscarProximo(alComandoDic(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 41869
    If lErro <> AD_SQL_SUCESSO Then Error 41870
    
    'se já tiver sido excluida anteriormente
    If iInativa = Inativo Then Error 41871
    
    lErro = Comando_LockExclusive(alComandoDic(1))
    If lErro <> AD_SQL_SUCESSO Then Error 41872
    
    'atualiza o flag de inativa (a exclusao é apenas logica)
    lErro = Comando_ExecutarPos(alComandoDic(2), "UPDATE FiliaisEmpresas SET Inativa = ?", alComandoDic(1), Inativo)
    If lErro <> AD_SQL_SUCESSO Then Error 41873
    
    'exclui todas as permissoes de acesso à filial
    lErro = Comando_Executar(alComandoDic(3), "DELETE FROM UsuFilEmp WHERE CodEmpresa = ? AND CodFilial = ?", objFilialEmpresa.lCodEmpresa, objFilialEmpresa.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 41874
    
    'fechar comandos
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
    
        Call Comando_Fechar(alComandoDic(iIndice))
        
    Next
    
    FilialEmpresa_Exclui_Trans = SUCESSO
     
    Exit Function
    
Erro_FilialEmpresa_Exclui_Trans:

    FilialEmpresa_Exclui_Trans = Err
     
    Select Case Err
          
        Case 41867
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 41868, 41869
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALEMPRESA", Err, objFilialEmpresa.iCodFilial)
        
        Case 41870
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)
        
        Case 41871
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_INATIVA", Err)
        
        Case 41872
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIALEMPRESA", Err, objFilialEmpresa.iCodFilial)
        
        Case 41873
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_FILIALEMPRESA", Err, objFilialEmpresa.iCodFilial)
        
        Case 41874
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_USUFILEMP", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159031)
     
    End Select
     
    'fechar comandos
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
    
        Call Comando_Fechar(alComandoDic(iIndice))
        
    Next
    
    Exit Function

End Function

Function Empresa_Grava(objEmpresa As ClassDicEmpresa, Optional objFilialEmpresa As AdmFiliais) As Long
'Grava Empresa fornecida em objEmpresa no BD

Dim lTransacao As Long
Dim lErro As Long
Dim lConexao As Long

On Error GoTo Erro_Empresa_Grava

    'Abre comandos
    lConexao = GL_lConexaoDic

    'Inicia a Transacao
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 25489
    
    'Grava Empresa no BD
    lErro = Empresa_Grava_EmTrans(objEmpresa)
    If lErro <> SUCESSO Then Error 25935
    
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        'Grava FilialEmpresa no banco de dados
'        lErro = FilialEmpresa_Grava_EmTrans(objFilialEmpresa)
'        If lErro Then Error 25925
'
'    End If
    
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 6206

    Empresa_Grava = SUCESSO

    Exit Function

Erro_Empresa_Grava:

    Empresa_Grava = Err

    Select Case Err

        Case 6206
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 25489
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 25925, 25935 'tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159032)

    End Select

    lErro = Transacao_RollbackDic

    Exit Function

End Function
Function Empresa_Grava_EmTrans(objEmpresa As ClassDicEmpresa) As Long
'Grava Empresa fornecida em objEmpresa no BD.

Dim lComando(1 To 6) As Long
Dim iIndice As Integer
Dim lErro As Long
Dim tEmpresa As typeDicEmpresa
Dim objEmpresaBD As New ClassDicEmpresa
Dim lConexao As Long
Dim iTotal As Integer
Dim objDicConfig As New ClassDicConfig
Dim sCgc As String, sNomeEmpresa As String
Dim iNumeroLogs As Integer, iNumeroEmpresas As Integer
Dim iNumeroFiliais As Integer, colModulosLib As New Collection
Dim dtDataValidade As Date, sTextoSenha As String
Dim objDicSelect As New DicSelect

On Error GoTo Erro_Empresa_Grava_EmTrans

    'Abre comandos
    lConexao = GL_lConexaoDic

    For iIndice = LBound(lComando) To UBound(lComando)
        lComando(iIndice) = Comando_AbrirExt(lConexao)
        If lComando(iIndice) = 0 Then Error 25490
    Next

    
    'Inicializa string
    tEmpresa.sNome = String(STRING_EMPRESA_NOME, 0)
    
    'Verifica se Nome já foi usado
    lErro = Empresa_Nome_Verifica(objEmpresa)
    If lErro <> SUCESSO Then Error 25479

    'Verifica se NomeReduzido já foi usado
    lErro = Empresa_NomeRed_Verifica(objEmpresa)
    If lErro <> SUCESSO Then Error 25480

    'Pesquisa no BD a Empresa em questão
    With tEmpresa
        lErro = Comando_ExecutarPos(lComando(1), "SELECT Nome, Inativa FROM Empresas WHERE Codigo=?", 0, .sNome, .iInativa, objEmpresa.lCodigo)
    End With
    If lErro <> AD_SQL_SUCESSO Then Error 6201

    'Le a Empresa, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6202

    'encontrou o Empresa ==> isto é uma alteração
    If lErro = AD_SQL_SUCESSO Then
    
        'Armazena atributos do BD
        objEmpresaBD.sNome = tEmpresa.sNome
        objEmpresaBD.iInativa = tEmpresa.iInativa
        
        lErro = Empresa_Altera(objEmpresa, objEmpresaBD, lComando())
        If lErro <> SUCESSO Then Error 25504

    Else 'não encontrou o Empresa ==>  é uma inserção de Empresa
        
        lErro = Comando_Executar(lComando(1), "SELECT COUNT(*) FROM Empresas WHERE Inativa = ? ", iTotal, Ativo)
        If lErro <> AD_SQL_SUCESSO Then Error 62382
        
        lErro = Comando_BuscarPrimeiro(lComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62383
        
        If iTotal = 0 Then
            sCgc = "0000"
            
            lErro = Gera_Senha_Provisoria(sCgc, objEmpresa.sNome)
            If lErro <> SUCESSO Then Error 62414
            
        End If
        
        lErro = objDicSelect.DicConfig_Le(objDicConfig)
        If lErro <> SUCESSO Then Error 62381
        
        lErro = objDicSelect.Senha_Empresa_Decifra(objDicConfig.sSenha, sCgc, sNomeEmpresa, iNumeroLogs, iNumeroEmpresas, iNumeroFiliais, colModulosLib, dtDataValidade, sTextoSenha)
        If lErro <> SUCESSO Then Error 62384
        
        objDicConfig.dtValidadeAte = dtDataValidade
        
        If iTotal = iNumeroEmpresas Then Error 62385
            
        lErro = Empresa_Insere(objEmpresa, lComando())
        If lErro <> SUCESSO Then Error 25491

    End If

    'Fecha comandos
    For iIndice = LBound(lComando) To UBound(lComando)
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    Empresa_Grava_EmTrans = SUCESSO

    Exit Function

Erro_Empresa_Grava_EmTrans:

    Empresa_Grava_EmTrans = Err

    Select Case Err

        Case 25479, 25480, 25491, 25504, 62381, 62384, 62414 'tratado na rotina chamada

        Case 25490
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6201, 6202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA", Err, objEmpresa.lCodigo)

        Case 62382, 62383
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA1", Err)
        
        Case 62385
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIMITEEMPRESAS_ATINGIDO", Err, iNumeroEmpresas)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159033)

    End Select

    For iIndice = LBound(lComando) To UBound(lComando)
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function


Private Function Empresa_Altera(objEmpresa As ClassDicEmpresa, objEmpresaBD As ClassDicEmpresa, lComando() As Long) As Long
'Altera no BD Empresa fornecida em objEmpresa.
'CHAMADA EM TRANSACAO

Dim lErro As Long
Dim objModulo As New AdmModulo
Dim vSigla As Variant
Dim tModulo As typeModulo
Dim iAtivoNaTela As Integer
Dim objDicConfig As New ClassDicConfig
Dim objDicSelect As New DicSelect

On Error GoTo Erro_Empresa_Altera

    'Faz lock na Empresa
    lErro = Comando_LockExclusive(lComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 25507
    
    'Se Empresa for inativa não permite alterar
    If objEmpresaBD.iInativa = Inativo Then Error 25499
        
    lErro = objDicSelect.DicConfig_Le(objDicConfig)
    If lErro <> SUCESSO Then Error 62417
        
    'Se Senha já foi recebida verifica se Nome foi alterado
    If objDicConfig.dtDataSenha <> DATA_NULA And objEmpresaBD.sNome <> objEmpresa.sNome Then Error 25500
        
    'Atualiza a Empresa
    lErro = Comando_ExecutarPos(lComando(2), "UPDATE Empresas SET Nome=?, NomeReduzido=?", lComando(1), objEmpresa.sNome, objEmpresa.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then Error 6203
    
    tModulo.sSigla = String(STRING_MODULO_SIGLA, 0)

    'Le Modulos dessa Empresa na tabela de ModuloEmpresa
    lErro = Comando_ExecutarPos(lComando(3), "SELECT SiglaModulo, Ativo FROM ModuloEmpresa WHERE CodEmpresa = ? ORDER BY SiglaModulo", 0, tModulo.sSigla, tModulo.iAtivo, objEmpresa.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 25505

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25506
    
    Do While lErro <> AD_SQL_SEM_DADOS
        
        iAtivoNaTela = 0
        
        'Determina atividade do Módulo
        For Each vSigla In objEmpresa.colModulos
        
            'Testa se Modulo lido no BD foi ativado na Tela
            If tModulo.sSigla = CStr(vSigla) Then
                
                iAtivoNaTela = 1
                Exit For
                
            End If
            
        Next
        
        'Se houver diferença de atividade Tela x BD altera ATIVIDADE no BD
        If iAtivoNaTela <> tModulo.iAtivo Then
            
            objModulo.sSigla = tModulo.sSigla
            'Altera atividade do Módulo na Empresa
            lErro = Empresa_AlteraAtividadeModulo(objEmpresa, objModulo, iAtivoNaTela, lComando())
            If lErro <> SUCESSO Then Error 25588
        
        End If
        
        lErro = Comando_BuscarProximo(lComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25509

    Loop


    Empresa_Altera = SUCESSO

    Exit Function

Erro_Empresa_Altera:

    Empresa_Altera = Err

    Select Case Err

        Case 6203
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_EMPRESA", Err, objEmpresa.lCodigo)

        Case 25499
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_EMPRESA_INATIVA", Err, objEmpresa.lCodigo)
        
        Case 25500
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_NOME", Err, objEmpresa.lCodigo)
        
        Case 25505, 25506, 25509
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULO_EMPRESA", Err, objEmpresa.lCodigo)

        Case 25507
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_EMPRESA", Err, objEmpresa.lCodigo)
        
        Case 25508
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MODULO_EMPRESA", Err, objEmpresa.lCodigo, tModulo.sSigla)
        
        Case 25588, 62417 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159034)

    End Select

    Exit Function

End Function

Private Function Empresa_AlteraAtividadeModulo(objEmpresa As ClassDicEmpresa, objModulo As AdmModulo, iAtivoNaTela As Integer, lComando() As Long) As Long
'Altera atividade de Modulo na Empresa
'CHAMADA EM TRANSACAO

Dim lErro As Long
Dim objModuloFilEmp As New ClassModuloFilEmp
Dim tModuloFilEmp As typeModuloFilEmp

On Error GoTo Erro_Empresa_AlteraAtividadeModulo
    
    'Se estiver desativando o módulo
    If iAtivoNaTela = 0 Then
    
        objModuloFilEmp.sSiglaModulo = objModulo.sSigla
        objModuloFilEmp.lCodEmpresa = objEmpresa.lCodigo
        
'        If giTipoVersao = VERSAO_FULL Then
        
            'Verifica se está vinculado a alguma Filial dessa Empresa que NÃO seja EMPRESA_TODA
            lErro = ModuloFilEmp_Le_ModuloEmpresa(objModuloFilEmp)
            If lErro <> SUCESSO And lErro <> 25516 Then Error 25587
            If lErro = SUCESSO Then Error 25512
    
'        End If
        
        'Pesquisa no BD registro de ModuloFilEmp para EMPRESA_TODA
        With tModuloFilEmp
            lErro = Comando_ExecutarPos(lComando(4), "SELECT CodFilial FROM ModuloFilEmp WHERE SiglaModulo=? AND CodEmpresa=? AND CodFilial=? ", 0, .iCodFilial, objModulo.sSigla, objEmpresa.lCodigo, EMPRESA_TODA)
        End With
        If lErro <> AD_SQL_SUCESSO Then Error 25589
    
        'Tenta ler registro de ModuloFilEmp (EMPRESA_TODA)
        lErro = Comando_BuscarPrimeiro(lComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25590
        
        'Se encontrou
        If lErro = AD_SQL_SUCESSO Then
        
            'Lock de ModuloFilEmp
            lErro = Comando_LockExclusive(lComando(4))
            If lErro <> AD_SQL_SUCESSO Then Error 25591
            
            'Exclui registro da tabela ModuloFilEmp (EMPRESA_TODA)
            lErro = Comando_ExecutarPos(lComando(5), "DELETE FROM ModuloFilEmp", lComando(4))
            If lErro <> AD_SQL_SUCESSO Then Error 25592
        
        End If
    
    Else
    
        lErro = Comando_Executar(lComando(5), "INSERT INTO ModuloFilEmp (CodEmpresa, CodFilial, SiglaModulo, Configurado) VALUES (?,?,?,?)", objEmpresa.lCodigo, EMPRESA_TODA, objModulo.sSigla, NAO_CONFIGURADO)
        If lErro <> AD_SQL_SUCESSO Then Error 55667
    
    End If
    
    'Atualiza ATIVIDADE do Modulo para essa Empresa
    lErro = Comando_ExecutarPos(lComando(6), "UPDATE ModuloEmpresa SET Ativo=?", lComando(3), iAtivoNaTela)
    If lErro <> AD_SQL_SUCESSO Then Error 25508
    
    Empresa_AlteraAtividadeModulo = SUCESSO

    Exit Function

Erro_Empresa_AlteraAtividadeModulo:

    Empresa_AlteraAtividadeModulo = Err

    Select Case Err

        Case 25508
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MODULO_EMPRESA", Err, objEmpresa.lCodigo, objModulo.sSigla)
       
        Case 25512
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MODULO_VINCULADO_FILIAL", Err, objModuloFilEmp.sSiglaModulo, objModuloFilEmp.lCodEmpresa, objModuloFilEmp.iCodFilial)
        
        Case 25587 'Tratado na rotina chamada
        
        Case 25589, 25590
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULOFILEMP1", Err, objModulo.sSigla, objEmpresa.lCodigo, EMPRESA_TODA)
        
        Case 25591
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MODULOFILEMP1", Err, objModulo.sSigla, objEmpresa.lCodigo, EMPRESA_TODA)
        
        Case 25592
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MODULOFILEMP1", Err, objModulo.sSigla, objEmpresa.lCodigo, EMPRESA_TODA)
        
        Case 55667
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MODULOFILEMP", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159035)

    End Select

    Exit Function

End Function

Function ModuloFilEmp_Le_ModuloEmpresa(objModuloFilEmp As ClassModuloFilEmp) As Long
'Tenta ler registro em ModuloFilEmp com siglaModulo e CodEmpresa passados.
'Tenta ler registro que NÃO corresponda a filial EMPRESA_TODA.

Dim lComando As Long
Dim lErro As Long
Dim tModuloFilEmp As typeModuloFilEmp

On Error GoTo Erro_ModuloFilEmp_Le_ModuloEmpresa

    'Abre comandos
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 25513
    
    'Pesquisa no BD registros de ModuloFilEmp vinculados ao Módulo
    With tModuloFilEmp
        lErro = Comando_Executar(lComando, "SELECT CodFilial, Configurado FROM ModuloFilEmp WHERE SiglaModulo=? AND CodEmpresa=? AND CodFilial<>? ORDER BY CodFilial", .iCodFilial, .iConfigurado, objModuloFilEmp.sSiglaModulo, objModuloFilEmp.lCodEmpresa, EMPRESA_TODA)
    End With
    If lErro <> AD_SQL_SUCESSO Then Error 25514

    'Tenta ler registro de ModuloFilEmp
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25515
    If lErro = AD_SQL_SEM_DADOS Then Error 25516

    'Passa os dados lidos para o objeto
    With tModuloFilEmp
        objModuloFilEmp.iCodFilial = .iCodFilial
        objModuloFilEmp.iConfigurado = .iConfigurado
    End With

    'Fecha comando
    lErro = Comando_Fechar(lComando)

    ModuloFilEmp_Le_ModuloEmpresa = SUCESSO

    Exit Function

Erro_ModuloFilEmp_Le_ModuloEmpresa:

    ModuloFilEmp_Le_ModuloEmpresa = Err

    Select Case Err

        Case 25513
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25514, 25515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULOFILEMP", Err)

        Case 25516 'não encontrou registro
        'a ser tratado na rotina chamadora

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159036)

    End Select

    'Fecha comando
    lErro = Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Empresa_Insere(objEmpresa As ClassDicEmpresa, lComando() As Long) As Long
'Insere no BD Empresa fornecida em objEmpresa.
'CHAMADA EM TRANSACAO

Dim iIndice As Integer
Dim lErro As Long
Dim sSigla As String
Dim colModulo As New Collection
Dim objModulo As AdmModulo
Dim vSigla As Variant
Dim iAtivo As Integer
Dim sDataSource As String
Dim sDataBaseFonte As String
Dim sDataBaseDestino As String
Dim sDiretorio As String, sServidorSQL As String
Dim lPosicao, lPosicaoAnterior As Long

On Error GoTo Erro_Empresa_Insere

    'se está instalando em rede
    If Mid$(App.Path, 1, 2) = "\\" Or GetDriveType(Mid$(App.Path, 1, 2)) = DRIVE_REMOTE Then
        
        'obter info do servidor sql a ser usado e do diretorio para onde será copiado o mdf
        sServidorSQL = InputBox("Identifique o servidor de banco de dados", "Servidor SQL", "(local)")

    Else
    
        sServidorSQL = "(local)"
        
    End If
    
    Screen.MousePointer = vbHourglass
    
    sSigla = String(STRING_MODULO_SIGLA, 0)

    'Insere a Empresa
    With objEmpresa
        If sServidorSQL = "(local)" Then
            .sStringConexao = "DSN=SGEDados" & CStr(objEmpresa.lCodigo) & ";UID=sa;PWD=;"
        Else
            .sStringConexao = "DSN=SGEDados" & CStr(objEmpresa.lCodigo)
        End If
        
        .iInativa = 0
        lErro = Comando_Executar(lComando(2), "INSERT INTO Empresas (Codigo, Nome, NomeReduzido, StringConexao, Inativa) VALUES (?,?,?,?,?)", .lCodigo, .sNome, .sNomeReduzido, .sStringConexao, .iInativa)
    End With

    If lErro <> AD_SQL_SUCESSO Then Error 6204

    'Le tabela de Modulos
    lErro = Comando_Executar(lComando(3), "SELECT Sigla FROM Modulos", sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 25492

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25493

    'Guarda os objModulo com siglas em colModulo
    Do While lErro <> AD_SQL_SEM_DADOS

        Set objModulo = New AdmModulo

        objModulo.sSigla = sSigla
        colModulo.Add objModulo
        lErro = Comando_BuscarProximo(lComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25494

    Loop

    'Insere registros em ModuloEmpresa
    For Each objModulo In colModulo

        'Testa se Modulo foi ativado para essa Empresa
        objModulo.iAtivo = 0

        For Each vSigla In objEmpresa.colModulos

            If objModulo.sSigla = CStr(vSigla) Then

                objModulo.iAtivo = 1
                Exit For

            End If

        Next

        'Insere TODOS os módulos em ModuloEmpresa para essa Empresa
        lErro = Comando_Executar(lComando(4), "INSERT INTO ModuloEmpresa (CodEmpresa, SiglaModulo, Ativo) VALUES (?,?,?)", objEmpresa.lCodigo, objModulo.sSigla, objModulo.iAtivo)
        If lErro <> SUCESSO Then Error 25495

    Next

    'Insere registros em ModuloFilEmp com Filial=0 (Empresa Toda).
    'Apenas Modulos que têm atividade na Empresa
    For Each objModulo In colModulo

        If objModulo.iAtivo = 1 Then

            lErro = Comando_Executar(lComando(5), "INSERT INTO ModuloFilEmp (CodEmpresa, CodFilial, SiglaModulo, Configurado) VALUES (?,?,?,?)", objEmpresa.lCodigo, EMPRESA_TODA, objModulo.sSigla, NAO_CONFIGURADO)
            If lErro <> AD_SQL_SUCESSO Then Error 25497

        End If

    Next

    'Insere em UsuFilEmp cruzamento de Filial EmpresaToda com Supervisor
    lErro = Comando_Executar(lComando(6), "INSERT INTO UsuFilEmp (CodEmpresa, CodFilial, CodUsuario) VALUES (?,?,?)", objEmpresa.lCodigo, EMPRESA_TODA, USU_SUP)
    If lErro <> AD_SQL_SUCESSO Then Error 25498
    
''    'Nome do DataSource para essa Empresa
''    sDataSource = "SGEDados" & CStr(objEmpresa.lCodigo)
''
''    'Encontra posicao da última \ no application path
''    lPosicao = 1
''
''    Do While lPosicao <> 0
''        lPosicaoAnterior = lPosicao
''        lPosicao = InStr(lPosicaoAnterior + 1, App.Path, "\")
''    Loop
''
''    'Monta diretorio\nome do DataBaseDestino e do DataBaseFonte
''    sDiretorio = left(App.Path, lPosicaoAnterior) & "Dados\"
''    sDataBaseDestino = sDiretorio & sDataSource & ".mdf"
''    sDataBaseFonte = sDiretorio & "SGEPadrao.mdf"
''
''    'Copia SGEPadrao para SGEDadosi
''    Call FileCopy(sDataBaseFonte, sDataBaseDestino)
''
''    'Cria DataSource
''    Call CriarDSNMSSQL(sDataSource, sDataBaseDestino, sServidorSQL, "Dados da Empresa " & objEmpresa.sNome)
    
    Screen.MousePointer = vbDefault
        
    Empresa_Insere = SUCESSO

    Exit Function

Erro_Empresa_Insere:

    Screen.MousePointer = vbDefault
        
    Empresa_Insere = Err

    Select Case Err

        Case 6204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_EMPRESA", Err, objEmpresa.lCodigo)

        Case 25492, 25493, 25494
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULO", Err)

        Case 25495
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MODULO_EMPRESA", Err, objEmpresa.lCodigo, sSigla)

        Case 25496
        
        Case 25497
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MODULOFILEMP", Err)

        Case 25498
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_USUFILEMP", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159037)

    End Select

    Exit Function

End Function

Private Function Empresa_Nome_Verifica(objEmpresa As ClassDicEmpresa) As Long
'Verifica se já existe esse nome de Empresa no BD.

Dim lErro As Long
Dim lComando As Long
Dim sNome As String

On Error GoTo Erro_Empresa_Nome_Verifica

    'Abre comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 25481

    sNome = String(STRING_EMPRESA_NOME, 0)

    'pesquisa nome no BD
    lErro = Comando_Executar(lComando, "SELECT Nome FROM Empresas WHERE Codigo <> ? AND Nome = ?", sNome, objEmpresa.lCodigo, objEmpresa.sNome)
    If lErro <> AD_SQL_SUCESSO Then Error 25482

    'verifica resultado da pesquisa
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25483

    'Se encontrou o nome, erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 25484

    'Fecha comando
    Call Comando_Fechar(lComando)

    Empresa_Nome_Verifica = SUCESSO

    Exit Function

Erro_Empresa_Nome_Verifica:

    Empresa_Nome_Verifica = Err

    Select Case Err

        Case 25481
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25482, 25483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA1", Err)

        Case 25484
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_NOME_JA_EXISTE", Err, objEmpresa.sNome)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159038)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function Empresa_NomeRed_Verifica(objEmpresa As ClassDicEmpresa) As Long
'Verifica se já existe esse nome reduzido de Empresa no BD.
'CHAMADA EM TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim sNomeRed As String

On Error GoTo Erro_Empresa_NomeRed_Verifica

    'Abre comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then Error 25488

    sNomeRed = String(STRING_EMPRESA_NOME_RED, 0)

    'pesquisa nome no BD
    lErro = Comando_Executar(lComando, "SELECT NomeReduzido FROM Empresas WHERE Codigo <> ? AND NomeReduzido = ?", sNomeRed, objEmpresa.lCodigo, objEmpresa.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then Error 25485

    'verifica resultado da pesquisa
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 25486

    'Se encontrou o nome reduzido, erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 25487

    'Fecha comando
    Call Comando_Fechar(lComando)

    Empresa_NomeRed_Verifica = SUCESSO

    Exit Function

Erro_Empresa_NomeRed_Verifica:

    Empresa_NomeRed_Verifica = Err

    Select Case Err

        Case 25488
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 25485, 25486
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA1", Err)

        Case 25487
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMPRESA_NOME_RED_JA_EXISTE", Err, objEmpresa.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159039)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function Usuario_AtualizaBDsDados1(colEmpresasAntes As Collection, objEmpresa As ClassDicEmpresa, iIndice As Integer) As Long
'retorna indice da Empresa na colecao
'retorna 0(zero) se nao encontrar

Dim lErro As Long, iIndiceTemp As Integer

On Error GoTo Erro_Usuario_AtualizaBDsDados1

    iIndice = 0
    
    For iIndiceTemp = 1 To colEmpresasAntes.Count
    
        If colEmpresasAntes(iIndiceTemp).lCodigo = objEmpresa.lCodigo Then
        
            iIndice = iIndiceTemp
            Exit For
            
        End If
        
    Next
    
    Usuario_AtualizaBDsDados1 = SUCESSO
     
    Exit Function
    
Erro_Usuario_AtualizaBDsDados1:

    Usuario_AtualizaBDsDados1 = Err
     
    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159040)
     
    End Select
     
    Exit Function

End Function

Private Function Usuario_AtualizaBDsDados2(colUsuFilEmp As Collection, objEmpresa As ClassDicEmpresa, iIndice As Integer) As Long
'retorna indice da Empresa na colecao
'retorna 0(zero) se nao encontrar

Dim lErro As Long, iIndiceTemp As Integer

On Error GoTo Erro_Usuario_AtualizaBDsDados2

    iIndice = 0
    
    For iIndiceTemp = 1 To colUsuFilEmp.Count
    
        If colUsuFilEmp(iIndiceTemp).lCodEmpresa = objEmpresa.lCodigo Then
        
            iIndice = iIndiceTemp
            Exit For
            
        End If
        
    Next
    
    Usuario_AtualizaBDsDados2 = SUCESSO
     
    Exit Function
    
Erro_Usuario_AtualizaBDsDados2:

    Usuario_AtualizaBDsDados2 = Err
     
    Select Case Err
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159041)
     
    End Select
     
    Exit Function

End Function

Private Function Usuario_AtualizaBDsDados(objUsuario As ClassDicUsuario, colUsuFilEmp As Collection, iOperacao As Integer) As Long
'atualiza a tabela Usuario dos bds de dados (SGEDados1,...)

Dim lErro As Long, objFilialEmpresa As New AdmFiliais
Dim colTodasEmpresas As New Collection, colEmpresasAntes As New Collection
Dim objEmpresa As ClassDicEmpresa, iIndice As Integer
Dim objUsuFilEmp As ClassUsuFilEmp, X As New DicSelect

On Error GoTo Erro_Usuario_AtualizaBDsDados

    If iOperacao = MODIFICACAO Then
    
        'obter colecao de todas as empresas (ativas)
        lErro = X.Empresas_Le_Todas(colTodasEmpresas)
        If lErro <> SUCESSO Then Error 56815
        
        'obter colecao de empresas onde tinha acesso
        lErro = X.Empresas_Le_Usuario(objUsuario.sCodUsuario, colEmpresasAntes)
        If lErro <> SUCESSO And lErro <> 50183 Then Error 56816
        
        'percorre colecao de todas as empresas
        For Each objEmpresa In colTodasEmpresas
        
            'pesquisa se a empresa tinha acesso
            lErro = Usuario_AtualizaBDsDados1(colEmpresasAntes, objEmpresa, iIndice)
            If lErro <> SUCESSO Then Error 56817
            
            'se tinha acesso
            If iIndice <> 0 Then
            
                objFilialEmpresa.lCodEmpresa = colEmpresasAntes(iIndice).lCodigo
                objFilialEmpresa.iCodFilial = EMPRESA_TODA
                
                lErro = FilialEmpresa_Conecta_Contab(objFilialEmpresa)
                If lErro <> SUCESSO Then Error 56818
                    
                'pesquisa se a empresa continua tendo acesso
                lErro = Usuario_AtualizaBDsDados2(colUsuFilEmp, objEmpresa, iIndice)
                If lErro <> SUCESSO Then Error 56819
                
                'se continua tendo acesso
                If iIndice <> 0 Then
            
                    'melhoria.: testar se foi alterado o nome, nome reduzido ou o status de ativo
                        
                    'atualizar atributos
                    lErro = Usuario_Grava_Contab(objUsuario)
                    If lErro <> SUCESSO Then Error 56820
                            
                Else 'se deixou de ter
                
                    'excluir reg
                    lErro = Usuario_Exclui_Contab(objUsuario.sCodUsuario)
                    If lErro <> SUCESSO Then Error 56821
            
                End If
                
            Else 'se nao tinha acesso
            
                'pesquisa se a empresa passou a ter acesso
                lErro = Usuario_AtualizaBDsDados2(colUsuFilEmp, objEmpresa, iIndice)
                If lErro <> SUCESSO Then Error 56822
                
                'se passou a ter
                If iIndice <> 0 Then
                    
                    objFilialEmpresa.lCodEmpresa = colUsuFilEmp(iIndice).lCodEmpresa
                    objFilialEmpresa.iCodFilial = EMPRESA_TODA
                
                    lErro = FilialEmpresa_Conecta_Contab(objFilialEmpresa)
                    If lErro <> SUCESSO Then Error 56823
                    
                    'incluir reg
                    lErro = Usuario_Grava_Contab(objUsuario)
                    If lErro <> SUCESSO Then Error 56824
                    
                End If
                
            End If
            
        Next
        
    Else
    
        'para cada empresa da colecao de empresas a que o usuario passa a ter acesso
        For Each objUsuFilEmp In colUsuFilEmp

            objFilialEmpresa.lCodEmpresa = objUsuFilEmp.lCodEmpresa
            objFilialEmpresa.iCodFilial = EMPRESA_TODA
            
            lErro = FilialEmpresa_Conecta_Contab(objFilialEmpresa)
            If lErro <> SUCESSO Then Error 56825
            
            'incluir registro no bd de dados correspondente a esta empresa
            lErro = Usuario_Grava_Contab(objUsuario)
            If lErro <> SUCESSO Then Error 56826
            
        Next
        
    End If
    
    Usuario_AtualizaBDsDados = SUCESSO
     
    Exit Function
    
Erro_Usuario_AtualizaBDsDados:

    Usuario_AtualizaBDsDados = Err
     
    Select Case Err
          
        Case 56815 To 56826
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159042)
     
    End Select
     
    Exit Function

End Function

'Alterada por Luiz Nogueira em 27/10/03
Private Function Usuario_ExcluiBDsDados(sCodUsuario As String) As Long
'exclui dos bds de dados (SGEDados1,...) o usuario passado

Dim lErro As Long, objFilialEmpresa As New AdmFiliais
Dim colEmpresas As New Collection, X As New DicSelect
Dim objEmpresa As ClassDicEmpresa, iIndice As Integer
Dim objLiberacaoCredito As Object

On Error GoTo Erro_Usuario_ExcluiBDsDados

    'obter colecao de empresas onde tinha acesso
    lErro = X.Empresas_Le_Usuario(sCodUsuario, colEmpresas)
    If lErro <> SUCESSO And lErro <> 50183 Then gError 56845
        
    'percorre colecao de empresas
    For Each objEmpresa In colEmpresas
    
        objFilialEmpresa.lCodEmpresa = objEmpresa.lCodigo
        objFilialEmpresa.iCodFilial = EMPRESA_TODA
        
        lErro = FilialEmpresa_Conecta_Contab(objFilialEmpresa)
        If lErro <> SUCESSO Then gError 56843
                    
        'Verifica se o Usuario é Comprador
        lErro = Usuario_Testa_Comprador(sCodUsuario)
        If lErro <> SUCESSO Then gError 58538
                
        'Incluído por Luiz Nogueira em 27/10/03
        'Verifica se o Usuario é Atendente
        lErro = X.Usuario_Testa_Atendente(sCodUsuario)
        If lErro <> SUCESSO Then gError 102791
        
        Set objLiberacaoCredito = CreateObject("GlobaisFAT.ClassLiberacaoCredito")
        
        objLiberacaoCredito.sCodUsuario = sCodUsuario
        
        'Exclui a Liberação de Crédito deste Usuario
        lErro = CF("LiberacaoCredito_Exclui", objLiberacaoCredito)
        If lErro <> SUCESSO And lErro <> 58550 Then gError 58539
            
        'Exclui o Usuario
        lErro = Usuario_Exclui_Contab(sCodUsuario)
        If lErro <> SUCESSO Then gError 56844
    
    Next
    
    Usuario_ExcluiBDsDados = SUCESSO
     
    Exit Function
    
Erro_Usuario_ExcluiBDsDados:

    Usuario_ExcluiBDsDados = gErr
     
    Select Case gErr
          
        Case 56843, 56844, 56845, 58538, 58539, 102791 'Tratado na Rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159043)
     
    End Select
     
    Exit Function

End Function

Public Function Usuario_Grava_Contab(ByVal objUsuario As ClassDicUsuario) As Long 'Alterado p/ public por Leo em 01/02/02
'atualiza registro da tabela de usuarios do bd de dados (SGEDados1,...)

Dim lErro As Long, alComando(1 To 4) As Long, iIndice As Integer
Dim lTransacao As Long, iAtivo As Integer, sCodUsuarioAux As String

On Error GoTo Erro_Usuario_Grava_Contab

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 56827
    
    Next
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 56828
    
    sCodUsuarioAux = String(STRING_USUARIO_CODIGO, 0)
    
    lErro = Comando_ExecutarPos(alComando(4), "SELECT CodUsuario FROM Usuario WHERE CodUsuario <> ? AND (Nome = ? OR NomeReduzido = ?)", 0, sCodUsuarioAux, objUsuario.sCodUsuario, objUsuario.sNome, objUsuario.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 213508
    
    lErro = Comando_BuscarProximo(alComando(4))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 213509
    
    If lErro = AD_SQL_SUCESSO Then gError 213510
        
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Ativo FROM Usuario WHERE CodUsuario = ?", 0, iAtivo, objUsuario.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 56829
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 56830
    
    'se encontrou é alteracao
    If lErro = AD_SQL_SUCESSO Then
    
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE Usuario SET Nome = ?, NomeReduzido = ?, Ativo = ?", alComando(1), objUsuario.sNome, objUsuario.sNomeReduzido, objUsuario.iAtivo)
        If lErro <> AD_SQL_SUCESSO Then gError 56831
    
    Else
    
        lErro = Comando_Executar(alComando(3), "INSERT INTO Usuario (CodUsuario, Nome, NomeReduzido, Ativo) VALUES (?,?,?,?)", objUsuario.sCodUsuario, objUsuario.sNome, objUsuario.sNomeReduzido, objUsuario.iAtivo)
        If lErro <> AD_SQL_SUCESSO Then gError 56832
    
    End If
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 56833
    
    For iIndice = LBound(alComando) To UBound(alComando)
    
        Call Comando_Fechar(alComando(iIndice))
    
    Next
    
    Usuario_Grava_Contab = SUCESSO
     
    Exit Function
    
Erro_Usuario_Grava_Contab:

    Usuario_Grava_Contab = gErr
     
    Select Case gErr
          
        Case 56827
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 56828
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 56829, 56830, 213508, 213509
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIO1", gErr, objUsuario.sCodUsuario)
        
        Case 56831
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_USUARIO", gErr, objUsuario.sCodUsuario)
        
        Case 56832
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_USUARIO", gErr, objUsuario.sCodUsuario)
        
        Case 56833
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
            
        Case 213510
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_USUARIO_REPETIDO", gErr, sCodUsuarioAux)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159044)
     
    End Select
     
    'Faz rollback na transação
    Call Transacao_Rollback
     
    For iIndice = LBound(alComando) To UBound(alComando)
    
        Call Comando_Fechar(alComando(iIndice))
    
    Next
    
    Exit Function

End Function

Private Function Usuario_Exclui_Contab(sCodUsuario As String) As Long
'exclui registro da tabela de usuarios do bd de dados (SGEDados1,...)

Dim lErro As Long, alComando(1 To 2) As Long, iIndice As Integer
Dim lTransacao As Long, iAtivo As Integer

On Error GoTo Erro_Usuario_Exclui_Contab

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 56836
    
    Next
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 56837
        
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Ativo FROM Usuario WHERE CodUsuario = ?", 0, iAtivo, sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 56838
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56839
    
    'se nao encontrou
    If lErro <> AD_SQL_SUCESSO Then Error 56840
    
    'exclui
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM Usuario", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 56841
        
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 56842
    
    For iIndice = LBound(alComando) To UBound(alComando)
    
        Call Comando_Fechar(alComando(iIndice))
    
    Next
    
    Usuario_Exclui_Contab = SUCESSO
     
    Exit Function
    
Erro_Usuario_Exclui_Contab:

    Usuario_Exclui_Contab = Err
     
    Select Case Err
          
        Case 56836
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 56837
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 56838, 56839
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIO1", Err, sCodUsuario)
        
        Case 56840
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", Err, sCodUsuario)
        
        Case 56841
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_USUARIO", Err, sCodUsuario)
        
        Case 56842
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159045)
     
    End Select
     
    'Faz rollback na transação
    Call Transacao_Rollback
     
    For iIndice = LBound(alComando) To UBound(alComando)
    
        Call Comando_Fechar(alComando(iIndice))
    
    Next
    
    Exit Function

End Function

Function CriarDSNMSSQL(lpNomeDSN As String, lpNomeMDB As String, lpSQLServer As String, lpDescricaoDSN As String) As Long
'Cria data source p/MS SQL Server 7 fazendo o attachment do bd ao servidor (SQL Server)

Dim rc As Integer, straux As String, straux1 As String
Dim henv As Long, hdbc As Long, cb_conn_out As Integer
Dim ach_conn_out As String
Dim pfErrorCode As Long, lpszErrorMsg As String, cbErrorMsgMax As Integer, pcbErrorMsg As Integer
Dim ssqlstate As String, iNative As Integer, sMsg As String, BufferLength As Integer, TextLength As Integer

On Error GoTo Erro_CriarDSNMSSQL

    rc = SQLAllocEnv(henv)
    If (rc <> 0) Then Error 59344
    
    rc = SQLAllocConnect(henv, hdbc)
    If (rc <> 0) Then Error 59345

    If lpSQLServer = "(local)" Then
        
        straux = "DSN=" & lpNomeDSN & Chr$(0) & "Server=(local)" & Chr$(0) & "Database=" & lpNomeDSN & Chr(0) & "AnsiNPW=yes" & Chr$(0) & "AutoTranslate=no" & Chr$(0) & "QuotedId=yes" & Chr$(0) & "Regional=no" & Chr$(0) & "Description=" & lpDescricaoDSN & Chr$(0) & "Language=English" & Chr$(0) & "Network=dbnmpntw" & Chr$(0) & "Trusted_Connection=no" & Chr$(0) & "QueryLog_On=no" & Chr$(0) & "QueryLogFile=c:\sqlqry.log" & Chr$(0) & "QueryLogTime=30000" & Chr$(0) & "StatsLog_On=no" & Chr$(0) & "StatsLogFile=c:\sqlstats.log" & Chr$(0) & Chr$(0)
        straux1 = "UID=sa;PWD=;DRIVER={SQL Server};DSN=" & lpNomeDSN & ";Server=(local)" & ";Database=" & lpNomeDSN & ";AttachDBFileName=" & lpNomeMDB & ";AnsiNPW=yes;AutoTranslate=no;QuotedId=yes;Regional=no;Description=" & lpDescricaoDSN & ";Language=(default);Network=dbnmpntw;Trusted_Connection=no;QueryLog_On=no;QueryLogFile=c:\sqlqry.log;QueryLogTime=30000;StatsLog_On=no;StatsLogFile=c:\sqlstats.log"
        
    Else
    
        straux = "DSN=" & lpNomeDSN & Chr$(0) & "Server=" & lpSQLServer & Chr$(0) & "Database=" & lpNomeDSN & Chr(0) & "AnsiNPW=yes" & Chr$(0) & "AutoTranslate=no" & Chr$(0) & "QuotedId=yes" & Chr$(0) & "Regional=no" & Chr$(0) & "Description=" & lpDescricaoDSN & Chr$(0) & "Language=English" & Chr$(0) & "Network=dbnmpntw" & Chr$(0) & "Trusted_Connection=yes" & Chr$(0) & "QueryLog_On=no" & Chr$(0) & "QueryLogFile=c:\sqlqry.log" & Chr$(0) & "QueryLogTime=30000" & Chr$(0) & "StatsLog_On=no" & Chr$(0) & "StatsLogFile=c:\sqlstats.log" & Chr$(0) & Chr$(0)
        straux1 = "DRIVER={SQL Server};DSN=" & lpNomeDSN & ";Server=" & lpSQLServer & ";Database=" & lpNomeDSN & ";AttachDBFileName=" & lpNomeMDB & ";AnsiNPW=yes;AutoTranslate=no;QuotedId=yes;Regional=no;Description=" & lpDescricaoDSN & ";Language=(default);Network=dbnmpntw;Trusted_Connection=yes;QueryLog_On=no;QueryLogFile=c:\sqlqry.log;QueryLogTime=30000;StatsLog_On=no;StatsLogFile=c:\sqlstats.log"
        
    End If
    
    rc = SQLConfigDataSource(0, ODBC_REMOVE_DSN, "SQL Server", "DSN" & lpNomeDSN)
    If (rc <> 1) Then
        'pode ocorrer em funcao de ainda nao ter sido criado um data source com este nome
''        cbErrorMsgMax = 2000
''        lpszErrorMsg = String(cbErrorMsgMax, 0)
''        rc = SQLInstallerError(1, pfErrorCode, lpszErrorMsg, cbErrorMsgMax, pcbErrorMsg)
''        MsgBox (lpszErrorMsg)
''        Error 59350
        
    End If

    rc = SQLConfigDataSource(0, ODBC_ADD_DSN, "SQL Server", straux)
    If (rc <> 1) Then
    
        cbErrorMsgMax = 2000
        lpszErrorMsg = String(cbErrorMsgMax, 0)
        rc = SQLInstallerError(1, pfErrorCode, lpszErrorMsg, cbErrorMsgMax, pcbErrorMsg)
        MsgBox (lpszErrorMsg)
        Error 59346
        
    End If

    ach_conn_out = String(2000, 0)
    rc = SQLDriverConnect(hdbc, 0, straux1, Len(straux1), ach_conn_out, 2000, cb_conn_out, SQL_DRIVER_NOPROMPT)
    If (rc <> 0 And rc <> 1) Then
    
        BufferLength = 2000
        sMsg = String(BufferLength, 0)
        ssqlstate = String(200, 0)
        rc = SQLError(henv, hdbc, 0, ssqlstate, iNative, sMsg, BufferLength, TextLength)
        MsgBox (sMsg)
        Error 59347

    End If
    
    Call SQLDisconnect(hdbc)
    Call SQLFreeConnect(hdbc)
    Call SQLFreeEnv(henv)
    
    CriarDSNMSSQL = SUCESSO
    
    Exit Function
    
Erro_CriarDSNMSSQL:

    CriarDSNMSSQL = Err
     
    Select Case Err
          
        Case 59344, 59345
            Call Rotina_Erro(vbOKOnly, "ERRO_ODBC_OBTER_HANDLE", Err)
            
        Case 59346, 59347, 59350
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159046)
     
    End Select
     
    If (hdbc <> 0) Then
    
        Call SQLDisconnect(hdbc)
        Call SQLFreeConnect(hdbc)
        hdbc = 0
        
    End If
    
    If (henv <> 0) Then
    
        Call SQLFreeEnv(henv)
        henv = 0
        
    End If
    
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

Public Function Usuario_Executa_Login(sUsuario As String, sSenha As String) As Long

Dim lComando As Long
Dim lComando2 As Long
Dim lConexao As Long
Dim lTransacao As Long
Dim lErro As Long, lTotalLogs As Long
Dim objDicConfig As New ClassDicConfig
Dim sCgc As String, sNomeEmpresa As String
Dim iNumeroLogs As Integer, iNumeroEmpresas As Integer
Dim iNumeroFiliais As Integer, colModulosLib As New Collection
Dim dtDataValidade As Date, sTextoSenha As String
Dim objDicSelect As New DicSelect, sUsu As String, sComputador As String

On Error GoTo Erro_Usuario_Executa_Login

    lConexao = GL_lConexaoDic
    
    lComando = Comando_AbrirExt(lConexao)
    If lComando = 0 Then Error 62391
    
    lComando2 = Comando_AbrirExt(lConexao)
    If lComando2 = 0 Then Error 62391
    
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 62392
    
    'faz login utilizando o codigo do usuario e a senha
    lErro = Sistema_Login(sUsuario, sSenha)
    If lErro <> AD_BOOL_TRUE Then Error 62398
    
    If sUsuario <> "backup" Then
    
        lErro = objDicSelect.DicConfig_Le(objDicConfig)
        If lErro <> SUCESSO Then Error 62393
        
        lErro = objDicSelect.Senha_Empresa_Decifra(objDicConfig.sSenha, sCgc, sNomeEmpresa, iNumeroLogs, iNumeroEmpresas, iNumeroFiliais, colModulosLib, dtDataValidade, sTextoSenha)
        If lErro <> SUCESSO Then Error 62394
        
        objDicConfig.dtValidadeAte = dtDataValidade
        
        If InStr(fOSMachineName, "JONES-STI") = 0 And InStr(fOSMachineName, "GHEINER-PC") = 0 And InStr(fOSMachineName, "W01-PC") = 0 And InStr(fOSMachineName, "W02-PC") = 0 Then
            
            'If Not (Date >= objDicConfig.dtValidadeDe And Date <= objDicConfig.dtValidadeAte) Then Error 62416
        
            lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM Usuarios WHERE Logado = ? AND CodUsuario <> ? ", lTotalLogs, USUARIO_LOGADO, sUsuario)
            If lErro <> AD_SQL_SUCESSO Then Error 62395
            
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62396
            
            If lTotalLogs >= iNumeroLogs Then Error 62397
            
            If left(UCase(fOSMachineName), 3) = "ASP" And UCase(fOSUserName) <> "DEMO102" And UCase(fOSUserName) <> "DEMO103" And UCase(fOSUserName) <> "DEMO104" And UCase(fOSUserName) <> "LIGHT1" Then
        
                sUsu = String(STRING_USUARIO_CODIGO, 0)
                sComputador = String(255, 0)
                lErro = Comando_Executar(lComando2, "SELECT CodUsuario, Computador FROM Usuarios WHERE Logado = ? AND CodUsuario <> ? AND NomeLogin = ?", sUsu, sComputador, USUARIO_LOGADO, sUsuario, fOSUserName)
                If lErro <> AD_SQL_SUCESSO Then Error 62395
                
                lErro = Comando_BuscarPrimeiro(lComando2)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62396
            
                If lErro = AD_SQL_SUCESSO Then Error 32317
            
            End If
        
        End If
        
        lErro = Usuario_Altera_SituacaoLogin_EmTrans(sUsuario, USUARIO_LOGADO)
        If lErro <> SUCESSO Then Error 62399
    
    End If
    
    lErro = Transacao_CommitDic
    If lErro <> SUCESSO Then Error 62401
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Usuario_Executa_Login = SUCESSO
    
    Exit Function

Erro_Usuario_Executa_Login:

    Usuario_Executa_Login = Err
    
    Select Case Err
    
        Case 32317
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOGIN_EM_USO", Err, fOSUserName, sUsu, sComputador)
            
        Case 62391
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 62392
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 62393, 62394, 62398, 62399
    
        Case 62395, 62396
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIOS_DIC", Err)
    
        Case 62397
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIMITELOGS_ATINGIDO", Err, iNumeroLogs)
        
        Case 62401
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 62416
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_FORA_VALIDADE", Err, Date, objDicConfig.dtValidadeDe, objDicConfig.dtValidadeAte)
        
        Case Err
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159047)
    
    End Select
    
    Call Transacao_RollbackDic
        
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function
    
End Function

Function Usuario_Altera_SituacaoLogin_EmTrans(sUsuario As String, iLogado As Integer)

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iLogadoAux As Integer, sComputador As String, sNomeLogin As String
Dim sComputadorAtual As String, sNomeLoginAtual As String, sModulo As String

On Error GoTo Erro_Usuario_Altera_SituacaoLogin_EmTrans
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComando(iIndice) = 0 Then Error 62402
    Next

    sComputador = String(255, 0)
    sNomeLogin = String(255, 0)
    sModulo = String(255, 0)
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Logado, Computador, NomeLogin, Modulo FROM Usuarios WHERE CodUsuario = ?", 0, iLogadoAux, sComputador, sNomeLogin, sModulo, sUsuario)
    If lErro <> AD_SQL_SUCESSO Then Error 62403
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62404
    If lErro <> AD_SQL_SUCESSO Then Error 62405
    
    'se está saindo
    If iLogado = USUARIO_NAO_LOGADO Then
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Usuarios SET Logado = ?, Computador = ?, NomeLogin = ?, Modulo = ?", alComando(0), iLogado, "", "", gsModulo)
    Else
        sComputadorAtual = fOSMachineName
        sNomeLoginAtual = fOSUserName
        If sComputador <> "" And UCase(sNomeLoginAtual) <> "DEMO102" And UCase(sNomeLoginAtual) <> "DEMO101" And UCase(sNomeLoginAtual) <> "DEMO103" And InStr(sComputadorAtual, "JONES-STI") = 0 And InStr(sComputadorAtual, "GHEINER-PC") = 0 And InStr(sComputadorAtual, "W01-PC") = 0 And InStr(sComputadorAtual, "W02-PC") = 0 Then
            If UCase(sComputador) <> UCase(sComputadorAtual) Or UCase(sNomeLogin) <> UCase(sNomeLoginAtual) Then Error 32317
        End If
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE Usuarios SET Logado = ?, Computador = ?, NomeLogin = ?", alComando(0), iLogado, sComputadorAtual, sNomeLoginAtual)
        gsModulo = IIf(UCase(sNomeLoginAtual) <> "DEMO102", sModulo, "Faturamento")
    End If
    If lErro <> AD_SQL_SUCESSO Then Error 62406

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Usuario_Altera_SituacaoLogin_EmTrans = SUCESSO
    
    Exit Function
    
Erro_Usuario_Altera_SituacaoLogin_EmTrans:

    Usuario_Altera_SituacaoLogin_EmTrans = Err
    
    Select Case Err
        
        Case 62402
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 62403, 62404
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIO", Err, sUsuario)
        
        Case 62405
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_CADASTRADO", Err, sUsuario)
        
        Case 62406
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_USUARIO", Err, sUsuario)
        
        Case 32317
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USUARIO_LOGADO", Err, sUsuario, sComputador, sNomeLogin)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159048)
    
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function DicConfig_Grava(objDicConfig As ClassDicConfig) As Long
'Grava as configuração passadas em DicConfig

Dim lErro As Long
Dim lTransacao As Long
Dim lConexao As Long

On Error GoTo Erro_DicConfig_Grava
    
   
    lConexao = GL_lConexaoDic
    
    'Abre a Transação
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 62356
    
    lErro = DicConfig_Grava_EmTrans(objDicConfig)
    If lErro <> SUCESSO Then Error 62411
    
    'Confira a transacao
    lErro = Transacao_CommitDic
    If lErro <> SUCESSO Then Error 62372
    
    
    DicConfig_Grava = SUCESSO

    Exit Function

Erro_DicConfig_Grava:

    DicConfig_Grava = Err
    
    Select Case Err
    
        Case 62356
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 62372
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
                
        Case 62411
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159049)
    
    End Select
    
    Call Transacao_RollbackDic
        
    Exit Function
    
End Function

Function Usuario_Altera_SituacaoLogin(sUsuario As String, iLogado As Integer)

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim lConexao As Long

On Error GoTo Erro_Usuario_Altera_SituacaoLogin
    
    lConexao = GL_lConexaoDic
    
    lTransacao = Transacao_AbrirDic
    If lTransacao = 0 Then Error 62408

    lErro = Usuario_Altera_SituacaoLogin_EmTrans(sUsuario, iLogado)
    If lErro <> SUCESSO Then Error 62409

    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then Error 62410

    Usuario_Altera_SituacaoLogin = SUCESSO
    
    Exit Function
    
Erro_Usuario_Altera_SituacaoLogin:

    Usuario_Altera_SituacaoLogin = Err
    
    Select Case Err
        
        Case 62408
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 62409
        
        Case 62410
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159050)
    
    End Select

    Call Transacao_RollbackDic
    
    Exit Function

End Function

Function DicConfig_Grava_EmTrans(objDicConfig As ClassDicConfig) As Long
'Grava as configuração passadas em DicConfig

Dim lErro As Long
Dim alComando(0 To 8) As Long
Dim iIndice As Integer
Dim iTotal As Integer
Dim sSigla As String
Dim iLiberado As Integer
Dim lConexao As Long
Dim lCodEmpresa As Long
Dim iFilial As Integer

On Error GoTo Erro_DicConfig_Grava_EmTrans
   
    lConexao = GL_lConexaoDic
    
    'Abre os comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(lConexao)
        If alComando(iIndice) = 0 Then Error 62357
    Next
    
    'BUsca o número de empresas ativas no sistema
    lErro = Comando_Executar(alComando(0), "SELECT COUNT(*) FROM Empresas WHERE Inativa = ?", iTotal, Ativo)
    If lErro <> AD_SQL_SUCESSO Then Error 62358
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62359
    
    'Se o número de empresas ativas for maior que o permitido pela senha, erro.
    If iTotal > objDicConfig.iLimiteEmpresas Then Error 62360
    
    iTotal = 0
    'BUsca o número de filiais empresas ativas no sistema
    lErro = Comando_Executar(alComando(0), "SELECT COUNT(*) FROM FiliaisEmpresas WHERE Inativa =? ", iTotal, Ativo)
    If lErro <> AD_SQL_SUCESSO Then Error 62361
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62362
        
    'Se o número de filiais empresas ativas for maior que o permitido pela senha, erro.
    If iTotal > objDicConfig.iLimiteFiliais Then Error 62363
    
    sSigla = String(STRING_DICCONFIG_SERIE, 0)
    'Lê em DicConfig a configuração
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Serie FROM DicConfig ", 0, sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 62364
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62365
    If lErro <> AD_SQL_SUCESSO Then Error 62366 'Não encontrou
    
    'Atualiza as Configurações de acordo com os dados passados
    lErro = Comando_ExecutarPos(alComando(2), "UPDATE DicConfig SET LimiteLogs =?, LimiteFiliais =?, Serie =?, LimiteEmpresas=?, ValidadeDe=?, Senha=?, DataSenha = ?", alComando(1), objDicConfig.iLimiteLogs, objDicConfig.iLimiteFiliais, objDicConfig.sSerie, objDicConfig.iLimiteEmpresas, objDicConfig.dtValidadeDe, objDicConfig.sSenha, objDicConfig.dtDataSenha)
    If lErro <> AD_SQL_SUCESSO Then Error 62367
    
    'BUsca os módulos do sistema
    sSigla = String(STRING_MODULO_SIGLA, 0)
    lErro = Comando_ExecutarPos(alComando(3), "SELECT Sigla FROM Modulos ", 0, sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 62368
    
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62369
    'Para cada módulo
    Do While lErro <> AD_SQL_SEM_DADOS
        'Verifica se esse módulo está liberado
        iLiberado = MODULO_NAO_LIBERADO
        For iIndice = 1 To objDicConfig.colModulosLib.Count
            If sSigla = objDicConfig.colModulosLib(iIndice).sSigla Then
                iLiberado = MODULO_LIBERADO
                Exit For
            End If
        Next
        'Atualiza a situação de liberação do Módulo
        lErro = Comando_ExecutarPos(alComando(4), "UPDATE Modulos SET Liberado = ?", alComando(3), iLiberado)
        If lErro <> SUCESSO Then Error 62371
        
        If iLiberado = MODULO_NAO_LIBERADO Then
            lErro = Comando_ExecutarPos(alComando(5), "SELECT CodEmpresa,Ativo FROM ModuloEmpresa WHERE SiglaModulo = ?", 0, iLiberado, lCodEmpresa, sSigla)
            If lErro <> AD_SQL_SUCESSO Then Error 62373
            
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62374
            Do While lErro <> AD_SQL_SEM_DADOS
            
                lErro = Comando_ExecutarPos(alComando(6), "UPDATE ModuloEmpresa SET Ativo =?", alComando(5), INATIVIDADE)
                If lErro <> AD_SQL_SUCESSO Then Error 62375
            
                lErro = Comando_BuscarProximo(alComando(5))
                If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 62376
            
            Loop
            lErro = Comando_ExecutarPos(alComando(7), "SELECT Configurado,CodFilial FROM ModuloFilEmp WHERE SiglaModulo = ? AND CodEmpresa = ?", 0, iLiberado, iFilial, sSigla, lCodEmpresa)
            If lErro <> AD_SQL_SUCESSO Then Error 62377
            
            lErro = Comando_BuscarPrimeiro(alComando(7))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62378
            Do While lErro <> AD_SQL_SEM_DADOS
            
                lErro = Comando_ExecutarPos(alComando(8), "DELETE FROM ModuloFilEmp", alComando(7))
                If lErro <> AD_SQL_SUCESSO Then Error 62379
            
                lErro = Comando_BuscarProximo(alComando(7))
                If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 62380
            Loop
        End If
        
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62370
    Loop
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    DicConfig_Grava_EmTrans = SUCESSO

    Exit Function

Erro_DicConfig_Grava_EmTrans:

    DicConfig_Grava_EmTrans = Err
    
    Select Case Err
    
        Case 62356
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 62357
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 62358, 62359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_EMPRESA1", Err)
            
        Case 62360
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMEMPRESAS_MAIOR_LIMITE", Err, iTotal, objDicConfig.iLimiteEmpresas)
        
        Case 62361, 62362
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISEMPRESAS", Err)
    
        Case 62363
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMFILIAISEMPRESAS_MAIOR_LIMITE", Err, iTotal, objDicConfig.iLimiteFiliais)
    
        Case 62364, 62365
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DICCONFIG", Err)
        
        Case 62366
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REGISTRO_DICCONFIG_NAO_ENCONTRADO", Err)
            
        Case 62367
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_DICCONFIG", Err)
        
        Case 62368, 62369, 62370
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULO", Err)
        
        Case 62371
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MODULOS", Err)
            
        Case 62372
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 62373, 62374, 62376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULO_EMPRESA", Err)
        
        Case 62375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MODULO_EMPRESA", Err, lCodEmpresa)
        
        Case 62377, 62378, 62380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MODULOFILEMP", Err)
        
        Case 62379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MODULOFILEMP", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159051)
    
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function Gera_Senha_Provisoria(sCgc As String, sNomeEmpresa As String) As Long
'Gera a senha privisória quan a primeira empresa ativa é cadastrada

Dim lErro As Long
Dim objDicConfig As New ClassDicConfig
Dim sSenha As String
Dim objModulosLib As AdmModulo
Dim colSiglasLib As New Collection
Dim iIndice As Integer, iIndice2 As Integer
Dim objVersao As New ClassVersao
Dim iAchouVersao As Integer
Dim objDicSelect As New DicSelect

On Error GoTo Erro_Gera_Senha_Provisoria
    
    lErro = CF("Modulos_Le_Todos", objDicConfig.colModulosLib)
    If lErro <> SUCESSO Then Error 62419

    lErro = objDicSelect.VersaoVigente_Le(objVersao)
    If lErro <> SUCESSO Then Error 62418

    iAchouVersao = False
    iIndice = 0
    'Libera os módulos que fazem parte das versões lançadas até a vigente
    For Each objModulosLib In objDicConfig.colModulosLib
        iIndice = iIndice + 1
        If objModulosLib.sVersao = objVersao.sCodigo Then iAchouVersao = True
        
        If Not iAchouVersao Or objModulosLib.sVersao = objVersao.sCodigo Then
            colSiglasLib.Add objModulosLib.sSigla
        Else
            Exit For
        End If
        
    Next

    If colSiglasLib.Count <> objDicConfig.colModulosLib.Count Then
    
        For iIndice2 = objDicConfig.colModulosLib.Count To iIndice Step -1
            objDicConfig.colModulosLib.Remove iIndice2
        Next

    End If
    
    'Recolhe os dados da senha provisória
    objDicConfig.dtDataSenha = DATA_NULA
    objDicConfig.dtValidadeDe = Date
    objDicConfig.dtValidadeAte = Date + TEMPO_SENHA_INICIAL
    objDicConfig.iLimiteEmpresas = 2
    objDicConfig.iLimiteFiliais = 2
    objDicConfig.iLimiteLogs = 2
    
    'Gera a senha
    lErro = objDicSelect.Senha_Empresa_Gera(sCgc, Mid(sNomeEmpresa, 1, 2), objDicConfig.iLimiteLogs, objDicConfig.iLimiteEmpresas, objDicConfig.iLimiteFiliais, colSiglasLib, objDicConfig.dtValidadeAte, sSenha)
    If lErro <> SUCESSO Then Error 62412
    
    'Trmina o preenchimento do objDicConfig
    objDicConfig.sSenha = sSenha
    
        
    'Grava em dicConfig as novas configurações
    lErro = DicConfig_Grava_EmTrans(objDicConfig)
    If lErro <> SUCESSO Then Error 62413
    
    Gera_Senha_Provisoria = SUCESSO
    
    Exit Function

Erro_Gera_Senha_Provisoria:

    Gera_Senha_Provisoria = Err
    
    Select Case Err
    
        Case 62412, 62413, 62418, 62419
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159052)
            
    End Select
    
    Exit Function

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function OpcoesTelas_Grava(ByVal objTela As Object) As Long

Dim lTransacaoDic As Long
Dim lErro As Long
Dim objOpcaoTela As New ClassOpcoesTelas
Dim objCombo As Object
Dim iIndice As Integer

On Error GoTo Erro_OpcoesTelas_Grava

    'Transfere os dados da tela para o obj
    lErro = Move_Tela_objOpcaoTela(objTela, objOpcaoTela)
    If lErro <> SUCESSO Then gError 102917
    
    'Inicia a Transacao
    lTransacaoDic = Transacao_AbrirDic
    If lTransacaoDic = 0 Then gError 102918
    
    'Grava Empresa no BD
    lErro = OpcoesTelas_Grava_EmTrans(objOpcaoTela)
    If lErro <> SUCESSO Then gError 102919
    
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 102920
    
    'Instancia objCombo, apontando para a combo OpcoesTela
    Set objCombo = objTela.Controls("OpcoesTela")
    
    'Verifica se o item gravado já está na combo
    For iIndice = 1 To objCombo.ListCount
        If objCombo.List(iIndice - 1) = objOpcaoTela.sOpcao Then Exit For
    Next
    
    'Se não encontrou o item na combo
    If iIndice > objCombo.ListCount Then
        objCombo.AddItem objOpcaoTela.sOpcao
        objCombo.ItemData(objCombo.NewIndex) = objOpcaoTela.lCodigo
    End If
    
    OpcoesTelas_Grava = SUCESSO

    Exit Function

Erro_OpcoesTelas_Grava:

    OpcoesTelas_Grava = gErr

    Select Case gErr

        Case 102917, 102919, 102944
        
        Case 102918
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 102920
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159053)

    End Select

    'Faz o rollback da transação
    Call Transacao_RollbackDic

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function OpcoesTelas_Grava_EmTrans(ByVal objOpcaoTela As ClassOpcoesTelas) As Long

Dim alComandoDic(1) As Long
Dim lErro As Long
Dim lCodigo As Long
Dim iPadrao As Integer
Dim iIndice As Integer

On Error GoTo Erro_OpcoesTelas_Grava_EmTrans

    'Abre comando no dic
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        alComandoDic(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComandoDic(iIndice) = 0 Then gError 102907
    Next
    
    'Se não foi passado o código da opção a ser gravada
    'Significa que é uma opção nova e deve ser inserida no BD
    If objOpcaoTela.lCodigo = 0 Then
    
        'Busca o código da próxima opção a ser gravada
        lErro = Comando_Executar(alComandoDic(0), "SELECT MAX(Codigo) FROM OpcoesTelas", lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 102908
        
        lErro = Comando_BuscarPrimeiro(alComandoDic(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102909
        
        'Guarda o código da opção que será gravada
        objOpcaoTela.lCodigo = lCodigo + 1
        
        'Insere no BD a opção em questão
        lErro = Comando_Executar(alComandoDic(1), "INSERT INTO OpcoesTelas(Codigo,Opcao,NomeTela,Padrao) VALUES(?,?,?,?)", objOpcaoTela.lCodigo, objOpcaoTela.sOpcao, objOpcaoTela.sNomeTela, objOpcaoTela.iPadrao)
        If lErro <> AD_SQL_SUCESSO Then gError 102912
        
        'Desmarca a opção que está no BD gravada como padrão (caso exista)
        lErro = OpcoesTelas_DesmarcaPadrao(objOpcaoTela)
        If lErro <> SUCESSO Then gError 102974
        
    'Senão, significa que é uma opção já existente e apenas a altera no bd
    Else
        
        'Lê a opção em questão no BD
        lErro = Comando_ExecutarPos(alComandoDic(0), "SELECT Padrao FROM OpcoesTelas WHERE Codigo=?", 0, iPadrao, objOpcaoTela.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 102910
        
        lErro = Comando_BuscarPrimeiro(alComandoDic(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102911
        
        'Se não encontrou => erro
        If lErro = AD_SQL_SEM_DADOS Then gError 102913
        
        'Se o campo padrão foi alterado => atualiza o BD
        If objOpcaoTela.iPadrao <> iPadrao Then
        
            'Faz o lock da opção
            lErro = Comando_LockExclusive(alComandoDic(0))
            If lErro <> AD_SQL_SUCESSO Then gError 102914
            
            'Desmarca a opção que está no BD gravada como padrão (caso exista)
            lErro = OpcoesTelas_DesmarcaPadrao(objOpcaoTela)
            If lErro <> SUCESSO Then gError 102973
            
            'Atualiza o campo padrão
            lErro = Comando_ExecutarPos(alComandoDic(1), "UPDATE OpcoesTelas SET Padrao=?", alComandoDic(0), objOpcaoTela.iPadrao)
            If lErro <> AD_SQL_SUCESSO Then gError 102959
            
        End If
        
        'Exclui os valores que estavam armazenados para a opção
        lErro = OpcoesTelasValores_Exclui_EmTrans(objOpcaoTela)
        If lErro <> SUCESSO Then gError 102915
    
    End If
        
    'Insere os valores da opção
    lErro = OpcoesTelasValores_Grava_EmTrans(objOpcaoTela)
    If lErro <> SUCESSO Then gError 102916
    
    'Executa o fechamento do Comando
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    OpcoesTelas_Grava_EmTrans = SUCESSO

    Exit Function

Erro_OpcoesTelas_Grava_EmTrans:

    OpcoesTelas_Grava_EmTrans = gErr

    Select Case gErr

        Case 102915, 102916, 102973, 102974
        
        Case 102907
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 102908 To 102911
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPCOESTELAS", gErr)
            
        Case 102912
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_OPCOESTELAS", gErr)
        
        Case 102959
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_OPCOESTELAS", gErr)
            
        Case 102913
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCOESTELAS_NAO_ENCONTRADA", gErr, objOpcaoTela.sOpcao, objOpcaoTela.sNomeTela)
            
        Case 102914
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_OPCOESTELAS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159054)

    End Select

    'Executa o fechamento do Comando
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next

End Function

'Incluído por Luiz Nogueira em 27/10/03
Private Function Move_Tela_objOpcaoTela(ByVal objTela As Object, ByVal objOpcaoTela As ClassOpcoesTelas) As Long

Dim lErro As Long
Dim objCombo As Object
Dim objControle As Object
Dim objOpcaoTelaValores As ClassOpcoesTelasValores
Dim colOpcoesTelasCampos As New Collection
Dim objOpcoesTelasCampos As ClassOpcoesTelasCampos
Dim iCodigo As Integer

On Error GoTo Erro_Move_Tela_objOpcaoTela

    'Instancia a combo OpcoesTela
    Set objCombo = objTela.Controls("OpcoesTela")
    
    'Guarda no obj o nome da tela e o nome da opção
    objOpcaoTela.sNomeTela = objTela.Name
    objOpcaoTela.sOpcao = objCombo.Text
    
    'Indica se a opção é padrão ou não
    objOpcaoTela.iPadrao = objTela.Controls("OpcaoPadrao").Value
    
    'Se o nome da opção não foi preenchido => erro
    If Len(Trim(objCombo.Text)) = 0 Then gError 102929
    
    'Indica que a tela não deve ser atualizada. Isso evita que a função combo_seleciona dispare novamente o click
    Call CallByName(objTela, "iAtualizaTela", VbLet, 0)
    
    'Verifica se a opção já existe na combo
    lErro = Combo_Seleciona(objCombo, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 102932
    
    'Se encontrou a opção na combo => guarda no obj o código da opção
    If lErro = SUCESSO Then objOpcaoTela.lCodigo = objCombo.ItemData(objCombo.ListIndex)
    
    'Indica que a tela deve ser atualizada a função OpcoesTela_Click for chamada, pois a função combo_seleciona já foi executada
    Call CallByName(objTela, "iAtualizaTela", VbLet, 1)
    
    'Lê os nomes dos controles que deverão ter seus valores gravados
    lErro = CF("OpcoesTelasCampos_Le_NomeTela", colOpcoesTelasCampos, objTela.Name)
    If lErro <> SUCESSO And lErro <> 102904 Then gError 102905
    
    'Se não encontrou nenhum controle para a tela em questão => erro
    If lErro = 102904 Then gError 102906

    'Para cada controle lido da tela que deve ter seu valor gravado
    For Each objOpcoesTelasCampos In colOpcoesTelasCampos
    
        'Instancia o controle em questao
        Set objControle = objTela.Controls(objOpcoesTelasCampos.sNomeControle)
        
        'Instancia um novo obj
        Set objOpcaoTelaValores = New ClassOpcoesTelasValores
            
        'Guarda no obj os dados do controle
        objOpcaoTelaValores.sNomeControle = objOpcoesTelasCampos.sNomeControle
        objOpcaoTelaValores.sValorCampo = objControle
        
        'Adiciona o controle à coleção de valores da opção em questão
        objOpcaoTela.colOpcoesTelasValores.Add objOpcaoTelaValores
    
    Next
    
    Move_Tela_objOpcaoTela = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_objOpcaoTela:

    Move_Tela_objOpcaoTela = gErr
    
    Select Case gErr
    
        Case 102929
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMEOPCAO_OBRIGATORIO", gErr)
            
        Case 102905
        
        Case 102906
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCOESTELASCAMPOS_TELA_NAO_ENCONTRADA", gErr, objTela.Name)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159055)
        
    End Select

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function OpcoesTelasValores_Exclui_EmTrans(ByVal objOpcaoTela As ClassOpcoesTelas) As Long

Dim lErro As Long
Dim alComandoDic(1) As Long
Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_OpcoesTelasValores_Exclui_EmTrans

    'Abre comando no dic
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        alComandoDic(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComandoDic(iIndice) = 0 Then gError 109921
    Next
    
    'Lê os valores para a opção e tela passadas como parâmetro
    lErro = Comando_ExecutarPos(alComandoDic(0), "SELECT CodOpcao FROM OpcoesTelasValores WHERE CodOpcao=?", 0, lCodigo, objOpcaoTela.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 102922
    
    lErro = Comando_BuscarPrimeiro(alComandoDic(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102923
    
    'Se não encontrou valores para a opção em questão
    If lErro = AD_SQL_SEM_DADOS Then gError 102926
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Exclui o valor do BD
        lErro = Comando_ExecutarPos(alComandoDic(1), "DELETE FROM OpcoesTelasValores", alComandoDic(0))
        If lErro <> AD_SQL_SUCESSO Then gError 102925
        
        'Busca o próximo valor encontrado para a opção em questão
        lErro = Comando_BuscarProximo(alComandoDic(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102924
    
    Loop

    'Executa o fechamento do Comando
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next

    OpcoesTelasValores_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_OpcoesTelasValores_Exclui_EmTrans:

    OpcoesTelasValores_Exclui_EmTrans = gErr

    Select Case gErr

        Case 102921
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 102922 To 102924
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPCOESTELASVALORES", gErr)
            
        Case 102925
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_OPCOESTELASVALORES", gErr)
        
        Case 102926
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCOESTELASVALORES_NAO_ENCONTRADO", gErr, objOpcaoTela.sOpcao, objOpcaoTela.sNomeTela)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159056)

    End Select

    'Executa o fechamento do Comando
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function OpcoesTelasValores_Grava_EmTrans(ByVal objOpcaoTela As ClassOpcoesTelas) As Long

Dim lErro As Long
Dim lComandoDic As Long
Dim iIndice As Integer
Dim objOpcaoTelaValor As ClassOpcoesTelasValores

On Error GoTo Erro_OpcoesTelasValores_Grava_EmTrans

    'Abre comando no dic
    lComandoDic = Comando_AbrirExt(GL_lConexaoDic)
    If lComandoDic = 0 Then gError 102927
    
    'Para cada valor dessa opção
    For Each objOpcaoTelaValor In objOpcaoTela.colOpcoesTelasValores
    
        'Insere o valor no BD
        lErro = Comando_Executar(lComandoDic, "INSERT INTO OpcoesTelasValores(CodOpcao,NomeControle,ValorCampo) VALUES(?,?,?)", objOpcaoTela.lCodigo, objOpcaoTelaValor.sNomeControle, objOpcaoTelaValor.sValorCampo)
        If lErro <> AD_SQL_SUCESSO Then gError 102928
    
    Next
    
    'Executa o fechamento do Comando
    Call Comando_Fechar(lComandoDic)

    OpcoesTelasValores_Grava_EmTrans = SUCESSO

    Exit Function

Erro_OpcoesTelasValores_Grava_EmTrans:

    OpcoesTelasValores_Grava_EmTrans = gErr

    Select Case gErr

        Case 102927
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 102928
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_OPCOESTELASVALORES", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159057)

    End Select

    'Executa o fechamento do Comando
    Call Comando_Fechar(lComandoDic)

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function OpcoesTelas_Exclui(ByVal objTela As Object) As Long

Dim lTransacaoDic As Long
Dim lErro As Long
Dim objOpcaoTela As New ClassOpcoesTelas
Dim objCombo As Object
Dim iIndice As Integer
Dim iCodigo As Integer

On Error GoTo Erro_OpcoesTelas_Exclui

    'Instancia a combo OpcoesTela
    Set objCombo = objTela.Controls("OpcoesTela")
    
    'Se o nome da opção não foi preenchido => erro
    If Len(Trim(objCombo.Text)) = 0 Then gError 102947
    
    'Indica que a tela não deve ser atualizada. Isso evita que a função combo_seleciona dispare novamente o click
    Call CallByName(objTela, "iAtualizaTela", VbLet, 0)
    
    'Verifica se a opção já existe na combo
    lErro = Combo_Seleciona(objCombo, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 102948
    
    'Se não encontrou a opção na combo => erro, pois a opção não existe
    If lErro <> SUCESSO Then gError 102949
    
    'Indica que a tela deve ser atualizada a função OpcoesTela_Click for chamada, pois a função combo_seleciona já foi executada
    Call CallByName(objTela, "iAtualizaTela", VbLet, 1)
    
    'Guarda no obj o código da opção
    objOpcaoTela.lCodigo = objCombo.ItemData(objCombo.ListIndex)
    
    'Inicia a Transacao
    lTransacaoDic = Transacao_AbrirDic
    If lTransacaoDic = 0 Then gError 102950
    
    'Grava Empresa no BD
    lErro = OpcoesTelas_Exclui_EmTrans(objOpcaoTela)
    If lErro <> SUCESSO Then gError 102951
    
    'confirma a transação
    lErro = Transacao_CommitDic
    If lErro <> AD_SQL_SUCESSO Then gError 102952
    
    'Instancia objCombo, apontando para a combo OpcoesTela
    Set objCombo = objTela.Controls("OpcoesTela")
    
    'Procura o item excluído na combo
    objCombo.RemoveItem (objCombo.ListIndex)
    
'??? retirar
'    For iIndice = 1 To objCombo.ListCount
'        If objCombo.List(iIndice - 1) = objOpcaoTela.sOpcao Then
'            objCombo.Remove (iIndice)
'            Exit For
'        End If
'    Next
    
    OpcoesTelas_Exclui = SUCESSO

    Exit Function

Erro_OpcoesTelas_Exclui:

    OpcoesTelas_Exclui = gErr

    Select Case gErr

        Case 102948, 102950 To 102952
        
        Case 102947
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMEOPCAO_OBRIGATORIO", gErr)
        
        Case 102949
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCOESTELAS_NAO_ENCONTRADA", gErr, objCombo.Text, objTela.Name)
            
        Case 102918
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 102920
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159058)

    End Select

    'faz o rollback da transação
    Call Transacao_RollbackDic

End Function

'Incluído por Luiz Nogueira em 27/10/03
Private Function OpcoesTelas_Exclui_EmTrans(ByVal objOpcaoTela As ClassOpcoesTelas) As Long
'Exclui a opçao e os valores para a opção passada como parâmetro

Dim lErro As Long
Dim alComandoDic(1) As Long
Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_OpcoesTelas_Exclui_EmTrans

    'Abre comando no dic
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        alComandoDic(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComandoDic(iIndice) = 0 Then gError 102953
    Next
    
    'Lê a opção passada para exclusão
    lErro = Comando_ExecutarPos(alComandoDic(0), "SELECT Codigo FROM OpcoesTelas WHERE Codigo=?", 0, lCodigo, objOpcaoTela.lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 102954
    
    lErro = Comando_BuscarPrimeiro(alComandoDic(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102955
    
    'Se não encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102959
    
    'Lock da opção
    lErro = Comando_LockExclusive(alComandoDic(0))
    If lErro <> AD_SQL_SUCESSO Then gError 102956
    
    'Exclui a opção
    lErro = Comando_ExecutarPos(alComandoDic(1), "DELETE FROM OpcoesTelas", alComandoDic(0))
    If lErro <> AD_SQL_SUCESSO Then gError 102957
    
    'Exclui os valores da opção
    lErro = OpcoesTelasValores_Exclui_EmTrans(objOpcaoTela)
    If lErro <> SUCESSO Then gError 102958

    'Executa o fechamento do Comando
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next

    OpcoesTelas_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_OpcoesTelas_Exclui_EmTrans:

    OpcoesTelas_Exclui_EmTrans = gErr

    Select Case gErr

        Case 102953
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 102954, 102955
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPCOESTELAS", gErr)
            
        Case 102956
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_OPCOESTELAS", gErr)
            
        Case 102959
            Call Rotina_Erro(vbOKOnly, "ERRO_OPCOESTELAS_NAO_ENCONTRADA1", gErr, objOpcaoTela.lCodigo)
            
        Case 102957
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_OPCOESTELAS", gErr)
            
        Case 102958

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159059)

    End Select

    'Executa o fechamento do Comando
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function OpcoesTelas_DesmarcaPadrao(ByVal objOpcaoTela As ClassOpcoesTelas) As Long
'Se está gravando uma opção padrão, verifica se existe outra padrão no BD e, se existir, desmarca

Dim lErro As Long
Dim alComandoDic(1) As Long
Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_OpcoesTelas_DesmarcaPadrao

    'Abre comando no dic
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        alComandoDic(iIndice) = Comando_AbrirExt(GL_lConexaoDic)
        If alComandoDic(iIndice) = 0 Then gError 102972
    Next
    
    'Se está marcando o campo padrão
    If objOpcaoTela.iPadrao = MARCADO Then
    
        'Verifica se existe outra opção para a tela em questão marcada como padrão
        lErro = Comando_ExecutarPos(alComandoDic(0), "SELECT Codigo FROM OpcoesTelas WHERE Codigo<>? AND Padrao=? AND NomeTela=?", 0, lCodigo, objOpcaoTela.lCodigo, MARCADO, objOpcaoTela.sNomeTela)
        If lErro <> AD_SQL_SUCESSO Then gError 102960
        
        lErro = Comando_BuscarPrimeiro(alComandoDic(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102961
        
        'Se encontrou um outro registro marcado como padrão
        If lErro = AD_SQL_SUCESSO Then
        
            'Faz o lock da opção
            lErro = Comando_LockExclusive(alComandoDic(0))
            If lErro <> AD_SQL_SUCESSO Then gError 102962
            
            'Desmarca a opção
            lErro = Comando_ExecutarPos(alComandoDic(1), "UPDATE OpcoesTELAS SET Padrao=?", alComandoDic(0), DESMARCADO)
            If lErro <> AD_SQL_SUCESSO Then gError 102963
        
        End If
    
    End If
    
    'Executa o fechamento do Comando
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next
    
    OpcoesTelas_DesmarcaPadrao = SUCESSO

    Exit Function

Erro_OpcoesTelas_DesmarcaPadrao:

    OpcoesTelas_DesmarcaPadrao = gErr

    Select Case gErr

        Case 102972
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 102960, 102961
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_OPCOESTELAS", gErr)
            
        Case 102962
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_OPCOESTELAS", gErr)
            
        Case 102963
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_OPCOESTELAS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159060)

    End Select

    'Executa o fechamento do Comando
    For iIndice = LBound(alComandoDic) To UBound(alComandoDic)
        Call Comando_Fechar(alComandoDic(iIndice))
    Next

End Function
