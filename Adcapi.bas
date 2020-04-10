Attribute VB_Name = "ADCAPI"

Global GL_lConexao As Long
Global GL_lTransacao As Long

'Rotinas de Gerencia de Arquivo Temporário

Global Const ARQ_TEMP_OK = 0
Global Const ARQ_TEMP_ERR_WRITE = 1
Global Const ARQ_TEMP_ERR_SEEK = 2
Global Const ARQ_TEMP_ERR_READ = 3

Declare Function Arq_Temp_Criar Lib "ADCRTL.DLL" Alias "FN_Exec_ArqTemp_OnCreate" (ByVal iTamanho_Registro As Integer) As Long
Declare Function Arq_Temp_Destruir Lib "ADCRTL.DLL" Alias "FN_Exec_ArqTemp_OnDestroy" (ByVal lID_Arq_Temp As Long) As Long
Declare Function Arq_Temp_Inserir Lib "ADCRTL.DLL" Alias "FN_Exec_ArqTemp_OnInsert" (ByVal lID_Arq_Temp As Long, anyRegistro As Any, lPosicao As Long) As Long
Declare Function Arq_Temp_Ler Lib "ADCRTL.DLL" Alias "FN_Exec_ArqTemp_OnGetDirect" (ByVal lID_Arq_Temp As Long, anyRegistro As Any, lPosicao As Long) As Long
'Informa que já terminou a inserção de registros e irá iniciar a leitura
Declare Function Arq_Temp_Preparar Lib "ADCRTL.DLL" Alias "FN_Exec_ArqTemp_OnPrepGet" (ByVal lID_Arq_Temp As Long) As Long


'Rotinas de Gerencia de Sort
Declare Function Sort_Abrir Lib "ADCRTL.DLL" Alias "FN_Sort_Criar" (ByVal iTamanho_Offset As Integer, ByVal iNum_Segmentos As Integer) As Long
Declare Function Sort_Destruir Lib "ADCRTL.DLL" Alias "FN_Sort_Destruir" (ByVal lID_Sort As Long) As Long
'Function Sort_Inserir Lib "ADCRTL.DLL" Alias "XYZ" (ByVal lID_Sort As Long, ByVal lPosicao As Long, Optional vSegmento1 As Variant, Optional vSegmento2 As Variant, Optional vSegmento3 As Variant, Optional vSegmento4 As Variant, Optional vSegmento5 As Variant, Optional vSegmento6 As Variant, Optional vSegmento7 As Variant, Optional vSegmento8 As Variant, Optional vSegmento9 As Variant, Optional vSegmento10 As Variant) As Long
Declare Function Sort_Classificar Lib "ADCRTL.DLL" Alias "FN_Sort_PrepMerge" (ByVal lID_Sort As Long) As Long
Declare Function Sort_Ler Lib "ADCRTL.DLL" Alias "FN_Sort_GetRec" (ByVal lID_Sort As Long, lPosicao As Long) As Long


'Retorno comandos SQL
Global Const SQL_SUCESSO = 0
Global Const SQL_SUCESSO_PARCIAL = 1
Global Const SQL_SUCESSO_ERRO = -1
Global Const SQL_SEM_DADOS = 100

Global Const AD_SQL_DRIVER_ODBC = 1

'Rotinas de Manipulação de Banco de Dados
'Function Comando_Abrir() As Long
Declare Function Comando_Fechar Lib "ADSQLMN.DLL" Alias "AD_Comando_Fechar" (ByVal lComando As Long) As Long
'Function Comando_Executar(ByVal lComando As Long, sComando_SQL As String, Optional anyParametro1 As Variant, Optional anyParametro2 As Variant, ..., Optional anyParametro30 As Variant) As Long
'Function Comando_ExecutarLockado(ByVal lComando As Long, sComando_SQL As String, Optional anyParametro1 As Variant, Optional anyParametro2 As Variant, ..., Optional anyParametro30 As Variant) As Long
'Function Comando_ExecutarPos(ByVal lComando As Long, sComando_SQL As String, ByVal lComandoSelect As Long, Optional anyParametro1 As Variant, Optional anyParametro2 As Variant, ..., Optional anyParametro30 As Variant) As Long

Declare Function Comando_BuscarPrimeiro Lib "ADSQLMN.DLL" Alias "AD_Comando_BuscarPri" (ByVal lComando As Long) As Long
Declare Function Comando_BuscarProximo Lib "ADSQLMN.DLL" Alias "AD_Comando_BuscarProx" (ByVal lComando As Long) As Long
Declare Function Comando_BuscarUltimo Lib "ADSQLMN.DLL" Alias "AD_Comando_BuscarUlt" (ByVal lComando As Long) As Long
Declare Function Comando_BuscarAnterior Lib "ADSQLMN.DLL" Alias "AD_Comando_BuscarAnt" (ByVal lComando As Long) As Long
Declare Function Comando_BuscarAbsoluto Lib "ADSQLMN.DLL" Alias "AD_Comando_BuscarAbs" (ByVal lComando As Long, ByVal lPosicao As Long) As Long
Declare Function Comando_BuscarRelativo Lib "ADSQLMN.DLL" Alias "AD_Comando_BuscarRel" (ByVal lComando As Long, ByVal lPosicao As Long) As Long

Declare Function Comando_LockShared Lib "ADSQLMN.DLL" Alias "AD_Comando_LockShared" (ByVal lComando As Long) As Long
Declare Function Comando_LockExclusive Lib "ADSQLMN.DLL" Alias "AD_Comando_LockExclusive" (ByVal lComando As Long) As Long
Declare Function Comando_Unlock Lib "ADSQLMN.DLL" Alias "AD_Comando_UnLock" (ByVal lComando As Long) As Long

'escondidas

'para Sort
Declare Function Sort_AddKey Lib "ADCRTL.DLL" Alias "FN_Sort_AddKey" (ByVal lID_Sort As Long, lPosicion As Long, sCampo As Variant) As Long
Declare Function Sort_AddKeySeg Lib "ADCRTL.DLL" Alias "FN_Sort_AddKeySeg" (ByVal lID_Sort As Long, sCampo As Variant) As Long


'para SQL
Declare Function Conexao_AbrirExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Abrir" (ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
Declare Function Conexao_FecharExt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Fechar" (ByVal lConexao As Long) As Long

Declare Function Comando_AbrirExt Lib "ADSQLMN.DLL" Alias "AD_Comando_Abrir" (ByVal lConexao As Long) As Long
Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Declare Function Comando_Preparar_LockadoInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar_Lockado" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Declare Function Transacao_AbrirExt Lib "ADSQLMN.DLL" Alias "AD_Transacao_Abrir" (ByVal lConexao As Long) As Long
Declare Function Transacao_CommitExt Lib "ADSQLMN.DLL" Alias "AD_Transacao_Commit" (ByVal lTransacao As Long) As Long
Declare Function Transacao_RollbackExt Lib "ADSQLMN.DLL" Alias "AD_Transacao_Rollback" (ByVal lTransacao As Long) As Long

Function Comando_Abrir() As Long
    Comando_Abrir = Comando_AbrirExt(GL_lConexao)
End Function

Function Comando_Executar1(ByVal lComando As Long, ByVal sComando_SQL As String, anyParametro1 As Variant) As Long
        Dim ret As Integer
    ret = Comando_PrepararInt(lComando, sComando_SQL)
    ret = Comando_BindVarInt(lComando, anyParametro1)
    ret = Comando_ExecutarInt(lComando)
    If ret = 1 Then ret = 0
    Comando_Executar1 = ret
End Function

Function Comando_Executar2(ByVal lComando As Long, ByVal sComando_SQL As String, anyParametro1 As Variant, anyParametro2 As Variant) As Long
            Dim ret As Integer
    ret = Comando_PrepararInt(lComando, sComando_SQL)
    ret = Comando_BindVarInt(lComando, anyParametro1)
    ret = Comando_BindVarInt(lComando, anyParametro2)
    ret = Comando_ExecutarInt(lComando)
    If ret = 1 Then ret = 0
    Comando_Executar2 = ret
End Function

Function Comando_Executar3(ByVal lComando As Long, ByVal sComando_SQL As String, anyParametro1 As Variant, anyParametro2 As Variant, anyParametro3 As Variant) As Long
            Dim ret As Integer
    ret = Comando_PrepararInt(lComando, sComando_SQL)
    ret = Comando_BindVarInt(lComando, anyParametro1)
    ret = Comando_BindVarInt(lComando, anyParametro2)
    ret = Comando_BindVarInt(lComando, anyParametro3)
    ret = Comando_ExecutarInt(lComando)
    If ret = 1 Then ret = 0
    Comando_Executar3 = ret
End Function

Function Comando_Executar4(ByVal lComando As Long, ByVal sComando_SQL As String, anyParametro1 As Variant, anyParametro2 As Variant, anyParametro3 As Variant, anyParametro4 As Variant) As Long
            Dim ret As Integer
    ret = Comando_PrepararInt(lComando, sComando_SQL)
    ret = Comando_BindVarInt(lComando, anyParametro1)
    ret = Comando_BindVarInt(lComando, anyParametro2)
    ret = Comando_BindVarInt(lComando, anyParametro3)
    ret = Comando_BindVarInt(lComando, anyParametro4)
    ret = Comando_ExecutarInt(lComando)
    If ret = 1 Then ret = 0
    Comando_Executar4 = ret
End Function

Function Conexao_Abrir(ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
    GL_lConexao = Conexao_AbrirExt(driver_sql, lpParamIn, ParamLenIn, lpParamOut, lpParamLenOut)
    Conexao_Abrir = GL_lConexao
End Function

Function Conexao_Fechar() As Long
    Conexao_Fechar = Conexao_FecharExt(GL_lConexao)
    GL_lConexao = 0
End Function

Function Sort_Inserir1(ByVal lID_Sort As Long, ByVal lPosicao As Long, vSegmento1 As Variant) As Long
    Sort_Inserir1 = Sort_AddKey(lID_Sort, lPosicao, vSegmento1)
End Function

Function Sort_Inserir2(ByVal lID_Sort As Long, ByVal lPosicao As Long, vSegmento1 As Variant, vSegmento2 As Variant) As Long
    If (Sort_AddKey(lID_Sort, lPosicao, vSegmento1) = 1) Then
        Sort_Inserir2 = Sort_AddKeySeg(lID_Sort, vSegmento1)
    Else
        Sort_Inserir2 = 0
    End If
End Function

Function Transacao_Abrir() As Long
    GL_lTransacao = Transacao_AbrirExt(GL_lConexao)
    Transacao_Abrir = GL_lTransacao
End Function

Function Transacao_Commit() As Long
    Transacao_Commit = Transacao_CommitExt(GL_lTransacao)
    GL_lTransacao = 0
End Function

Function Transacao_Rollback() As Long
    Transacao_Rollback = Transacao_RollbackExt(GL_lTransacao)
    GL_lTransacao = 0
End Function

