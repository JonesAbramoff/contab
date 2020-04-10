VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmSQL"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'??? pendencia:
    'colocar no padrao de programacao
    'inverter retornos AD_BOOL_TRUE p/SUCESSO

'para SQL
Private Declare Function Conexao_AbrirInt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Abrir" (ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
Private Declare Function Conexao_FecharInt Lib "ADSQLMN.DLL" Alias "AD_Conexao_Fechar" (ByVal lConexao As Long) As Long
Private Declare Function Conexao_ObterInfo Lib "ADSQLMN.DLL" Alias "AD_Conexao_ObterInfo" (ByVal lConexao As Long, ByVal lpParamIn As String, iRet As Integer) As Long
Private Declare Function Conexao_DesativarLocksInt Lib "ADSQLMN.DLL" Alias "AD_Conexao_DesativarLocks" (ByVal lConexao As Long, ByVal bAtivacao As Long) As Long

Private Declare Function Comando_AbrirInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Abrir" (ByVal lConexao As Long) As Long
Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
'Private Declare Function Comando_PrepararParcInt Lib "ADSQLMN.DLL" Alias "AD_Comando_PrepararParc" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_PrepararLockadoInt Lib "ADSQLMN.DLL" Alias "AD_Comando_PrepararLockado" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_PrepararPosInt Lib "ADSQLMN.DLL" Alias "AD_Comando_PrepararPos" (ByVal lComando As Long, ByVal lpSQLStmt As String, ByVal lSelect As Long) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long
Private Declare Function Comando_FecharInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Fechar" (ByVal lComando As Long) As Long
Private Declare Function Comando_DefScrollInt Lib "ADSQLMN.DLL" Alias "AD_Comando_DefScroll" (ByVal lComando As Long, ByVal iComScroll As Integer) As Long
Private Declare Function Comando_BuscarPrimeiroInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BuscarPri" (ByVal lComando As Long) As Long
Private Declare Function Comando_BuscarProximoInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BuscarProx" (ByVal lComando As Long) As Long
Private Declare Function Comando_BuscarUltimoInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BuscarUlt" (ByVal lComando As Long) As Long
Private Declare Function Comando_BuscarAnteriorInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BuscarAnt" (ByVal lComando As Long) As Long
Private Declare Function Comando_BuscarAbsolutoInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BuscarAbs" (ByVal lComando As Long, ByVal lPosicao As Long) As Long
Private Declare Function Comando_BuscarRelativoInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BuscarRel" (ByVal lComando As Long, ByVal lPosicao As Long) As Long
Private Declare Function Comando_LockSharedInt Lib "ADSQLMN.DLL" Alias "AD_Comando_LockShared" (ByVal lComando As Long) As Long
Private Declare Function Comando_LockExclusiveInt Lib "ADSQLMN.DLL" Alias "AD_Comando_LockExclusive" (ByVal lComando As Long) As Long
Private Declare Function Comando_UnlockInt Lib "ADSQLMN.DLL" Alias "AD_Comando_UnLock" (ByVal lComando As Long) As Long
Private Declare Function Comando_ReExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_ReExecutar" (ByVal lComando As Long) As Long

Private Declare Function Transacao_AbrirInt Lib "ADSQLMN.DLL" Alias "AD_Transacao_Abrir" (ByVal lConexao As Long) As Long
Private Declare Function Transacao_CommitInt Lib "ADSQLMN.DLL" Alias "AD_Transacao_Commit" (ByVal lTransacao As Long) As Long
Private Declare Function Transacao_RollbackInt Lib "ADSQLMN.DLL" Alias "AD_Transacao_Rollback" (ByVal lTransacao As Long) As Long

Private Const JONES_280115 = 0

Public Function SQL_Comandos_Abertos() As Integer

Dim iIndice As Integer, iAbertos As Integer

    iAbertos = 0
    
    For iIndice = 1 To giNumCom

        If glCom(iIndice) <> 0 Then

            MsgBox (gsCom(iIndice))
            iAbertos = iAbertos + 1

        End If

    Next
    
    SQL_Comandos_Abertos = iAbertos

End Function

#If JONES_280115 = 1 Then
Private Sub Com_Inclui(ByVal lCom As Long)

Dim iIndice As Integer, bAchouVaga As Boolean

    bAchouVaga = False

    For iIndice = 1 To giNumCom

        If glCom(iIndice) = 0 Then

            glCom(iIndice) = lCom
            bAchouVaga = True
            Exit For

        End If

    Next

    If bAchouVaga = False Then
        If InStr(SistGlob_sNomeEmpresa, "kin") <> 0 Then
            If giNumCom > 1800 Then MsgBox (CStr(giNumCom) & " comandos abertos")
        End If
        giNumCom = giNumCom + 1
        glCom(giNumCom) = lCom
    End If

End Sub

Private Function Com_Exclui(ByVal lCom As Long) As Long

Dim iIndice As Integer, bAchou As Boolean

    Com_Exclui = 0 'sucesso
    
    If lCom <> 0 Then

        bAchou = False
        For iIndice = giNumCom To 1 Step -1
    
            If glCom(iIndice) = lCom Then
    
                glCom(iIndice) = 0
                gsCom(iIndice) = ""
                bAchou = True
                Exit For
    
            End If
    
        Next
    
        If bAchou = False Then
            If InStr(SistGlob_sNomeEmpresa, "kin") <> 0 Or SistGlob_sNomeEmpresa = "Renam e Melo Ltda." Then
                MsgBox ("erro com_exclui")
                Error 12345
            End If
            Com_Exclui = 1
        End If

    End If

End Function

Private Function Com_Testa(ByVal lCom As Long, Optional ByVal sComando_SQL As String = "") As Long

Dim iIndice As Integer, bAchou As Boolean

    Com_Testa = 0
    
    If lCom <> 0 Then
    
        bAchou = False
        For iIndice = giNumCom To 1 Step -1
    
            If glCom(iIndice) = lCom Then
                If sComando_SQL <> "" Then gsCom(iIndice) = sComando_SQL
                bAchou = True
                Exit For
    
            End If
    
        Next
    
        If bAchou = False Then
            If InStr(SistGlob_sNomeEmpresa, "kin") <> 0 Or SistGlob_sNomeEmpresa = "Renam e Melo Ltda." Then
                MsgBox ("erro com_testa")
                Error 12345
            End If
            Com_Testa = 1
        End If

    End If

End Function
#End If

Function Conexao_Suporta_TimeStampDiff(bSuporta As Boolean, lConexao As Long) As Long
'retorna em bSuporta se pode ser usado {fn TIMESTAMPDIFF em comando SQL
Dim lErro As Long, iRet As Integer, X As New ADM

On Error GoTo Erro_Conexao_Suporta_TimeStampDiff

    bSuporta = False
    
    If lConexao = 0 Then Error 59282
    
    lErro = Conexao_ObterInfo(lConexao, "TIMESTAMPDIFF", iRet)
    If lErro <> AD_BOOL_TRUE Then Error 59283
    
    If iRet <> 0 Then bSuporta = True
    
    Conexao_Suporta_TimeStampDiff = SUCESSO
    
    Exit Function
    
Erro_Conexao_Suporta_TimeStampDiff:

    Conexao_Suporta_TimeStampDiff = Err
    
    Select Case Err
          
        Case 59282, 59283
        
        Case Else
            Call X.Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142629)
     
    End Select
     
    Exit Function

End Function

Function Conexao_ObterOrdParam(lConexao As Long, iTipo As Integer) As Long
'retorna em iTipo SQL_ORD_PARAM_ESQ_DIR ou SQL_ORD_PARAM_EXECUCAO

Dim lErro As Long, iRet As Integer, X As New ADM

On Error GoTo Erro_Conexao_ObterOrdParam

    If lConexao = 0 Then Error 59282
    
    lErro = Conexao_ObterInfo(lConexao, "NAMED_PARAMETERS", iRet)
    If lErro <> AD_BOOL_TRUE Then Error 59283
    
    If iRet <> 0 Then
        iTipo = SQL_ORD_PARAM_ESQ_DIR
    Else
        iTipo = SQL_ORD_PARAM_EXECUCAO
    End If
    
    Conexao_ObterOrdParam = SUCESSO
    
    Exit Function
    
Erro_Conexao_ObterOrdParam:

    Conexao_ObterOrdParam = Err
    
    Select Case Err
          
        Case 59282, 59283
        
        Case Else
            Call X.Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142630)
     
    End Select
     
    Exit Function

End Function

'**********************************************************************
'**********************************************************************
' Function Conexao_Abrir (ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
'    Estabelece conexao com bd.
'
'Parametros:
'
'(I)driver_sql: AD_SQL_DRIVER_*. Ver os defines acima.
'(I)lpParamIn: informacoes adicionais para conexao que variam de bd para bd.
'    Ex: p/driver ODBC podemos passar o string de conexao: "DSN=XYZ Corp;UID=Gomez;PWD=Sesame;UIDDBMS=JGomez;PWDDBMS=Shazam;"
'(I)ParamLenIn: tamanho das informacoes referenciadas por lpParamIn.
'(O)lpParamOut: informacoes de conexao na saida. Pode ser um string mais completo que o informado na entrada. Este string mais completo poderia ficar armazenado no cadastro de usuarios criptografado para evitar que nas conexoes seguintes tenha que preencher dialogos p/logon).
'    Pode ser NULL se nao interessar. Neste caso lpParamLenOut nao sera alterado.
'    O tamanho do buffer deve ser >= 255.
'(O)lpParamLenOut: tamanho das informacoes referenciadas por lpParamOut.
'
'Retorno:
'    ponteiro p/conexao (long no VB) ou NULL p/erro.
'Observacao:
'        So trabalharemos com READ_COMMITTED e manual commit.
'        So havera uma transacao por conexao.
'
'Exemplo:
'    Dim ConStrOut As String * 500
'    Dim len_out As Integer
'    Dim lConexao As Long
'
'    lConexao = Conexao_Abrir(AD_SQL_DRIVER_ODBC, "DSN=sdk21-Access", Len("DSN=sdk21-Access") + 1, ByVal ConStrOut, len_out)

'**********************************************************************
'**********************************************************************
'Function Conexao_Fechar () As Long
'    Encerra conexao com bd
'Retorno:
'    AD_BOOL_TRUE p / SUCESSO

'**********************************************************************
'**********************************************************************
'Function Transacao_Abrir () As Long
'    iniciar uma transacao
'Retorno:
'    NULL p/erro

'**********************************************************************
'**********************************************************************
'Function Transacao_Commit () As Long
'    Confirma todas as alteracoes no bd que faziam parte da transacao
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO

'**********************************************************************
'**********************************************************************
'Function Transacao_Rollback () As Long
'    Desfaz todas as alteracoes no bd que faziam parte da transacao
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO
    
'**********************************************************************
'**********************************************************************
'Function Comando_Abrir() As Long
'    Inicializa "comando"
'Retorno:
'    NULL em caso de erro.

Function Comando_Fechar(lComando As Long) As Long
'    encerra um comando
'Parametros:
'    (I)lpComando: retornado em Comando_Abrir
'Retorno:
'    AD_BOOL_TRUE p/sucesso

#If JONES_280115 = 1 Then
    If lComando <> 0 Then Call Com_Exclui(lComando)
#End If

    Comando_Fechar = Comando_FecharInt(lComando)
    lComando = 0

End Function
    
'**********************************************************************
'**********************************************************************
'Function Comando_Executar(ByVal lComando As Long, ByVal sComando_SQL As String, Optional anyParametro1 As Variant, Optional anyParametro2 As Variant, Optional anyParametro3 As Variant, Optional anyParametro4 As Variant, Optional anyParametro5 As Variant, Optional anyParametro6 As Variant, Optional anyParametro7 As Variant, Optional anyParametro8 As Variant, Optional anyParametro9 As Variant, Optional anyParametro10 As Variant, Optional anyParametro11 As Variant, Optional anyParametro12 As Variant, Optional anyParametro13 As Variant, Optional anyParametro14 As Variant, Optional anyParametro15 As Variant) As Long

'**********************************************************************
'**********************************************************************
'Function Comando_ExecutarPos(ByVal lComando As Long, ByVal sComando_SQL As String, ByVal lSelect As Long, Optional anyParametro1 As Variant, Optional anyParametro2 As Variant, Optional anyParametro3 As Variant, Optional anyParametro4 As Variant, Optional anyParametro5 As Variant, Optional anyParametro6 As Variant, Optional anyParametro7 As Variant, Optional anyParametro8 As Variant, Optional anyParametro9 As Variant, Optional anyParametro10 As Variant, Optional anyParametro11 As Variant, Optional anyParametro12 As Variant, Optional anyParametro13 As Variant, Optional anyParametro14 As Variant, Optional anyParametro15 As Variant) As Long

'**********************************************************************
'**********************************************************************
'Function Comando_ExecutarLockado(ByVal lComando As Long, ByVal sComando_SQL As String, Optional anyParametro1 As Variant, Optional anyParametro2 As Variant, Optional anyParametro3 As Variant, Optional anyParametro4 As Variant, Optional anyParametro5 As Variant, Optional anyParametro6 As Variant, Optional anyParametro7 As Variant, Optional anyParametro8 As Variant, Optional anyParametro9 As Variant, Optional anyParametro10 As Variant, Optional anyParametro11 As Variant, Optional anyParametro12 As Variant, Optional anyParametro13 As Variant, Optional anyParametro14 As Variant, Optional anyParametro15 As Variant) As Long

Function Comando_BuscarPrimeiro(ByVal lComando As Long) As Long
'    Posiciona o cursor na primeira linha do resultado atualizando o conteudo dos buffers.
'Parametros:
'    (I)LPCOMANDO lpComando: retornado em AD_Comando_Abrir (referente ao SELECT)
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO, AD_SQL_SEM_DADOS
#If JONES_280115 = 1 Then
    Call Com_Testa(lComando)
#End If
    Comando_BuscarPrimeiro = Comando_BuscarPrimeiroInt(lComando)
    
End Function
    
Function Comando_BuscarProximo(ByVal lComando As Long) As Long
'    Posiciona o cursor na proxima linha do resultado atualizando o conteudo dos buffers.
'Parametros:
'    (I)LPCOMANDO lpComando: retornado em AD_Comando_Abrir (referente ao SELECT)
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO, AD_SQL_SEM_DADOS

#If JONES_280115 = 1 Then
    Call Com_Testa(lComando)
#End If
    Comando_BuscarProximo = Comando_BuscarProximoInt(lComando)
    
End Function
    
Function Comando_BuscarUltimo(ByVal lComando As Long) As Long
'    Posiciona o cursor na ultima linha do resultado atualizando o conteudo dos buffers.
'Parametros:
'    (I)LPCOMANDO lpComando: retornado em AD_Comando_Abrir (referente ao SELECT)
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO, AD_SQL_SEM_DADOS
#If JONES_280115 = 1 Then
    Call Com_Testa(lComando)
#End If
    Comando_BuscarUltimo = Comando_BuscarUltimoInt(lComando)
    
End Function

Function Comando_BuscarAnterior(ByVal lComando As Long) As Long
'    Posiciona o cursor na linha anterior do resultado atualizando o conteudo dos buffers.
'Parametros:
'    (I)LPCOMANDO lpComando: retornado em AD_Comando_Abrir (referente ao SELECT)
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO, AD_SQL_SEM_DADOS
#If JONES_280115 = 1 Then
    Call Com_Testa(lComando)
#End If
    Comando_BuscarAnterior = Comando_BuscarAnteriorInt(lComando)
    
End Function

Function Comando_BuscarAbsoluto(ByVal lComando As Long, ByVal lPosicao As Long) As Long
'    Posiciona o cursor em linha especifica a partir do inicio ou fim do resultado,
'        atualizando o conteudo dos buffers.
'Parametros:
'    (I)LPCOMANDO lpComando: retornado em AD_Comando_Abrir
'    (I)long posicao: numero da linha no resultado . Pode ser negativa como -2 significando a segunda linha do final para o inicio.
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO, AD_SQL_SEM_DADOS
#If JONES_280115 = 1 Then
    Call Com_Testa(lComando)
#End If
    Comando_BuscarAbsoluto = Comando_BuscarAbsolutoInt(lComando, lPosicao)
    
End Function

Function Comando_BuscarRelativo(ByVal lComando As Long, ByVal lPosicao As Long) As Long
'    Posiciona o cursor em linha especifica a partir da posicao corrente,
'        atualizando o conteudo dos buffers.
'Parametros:
'    (I)LPCOMANDO lpComando: retornado em AD_Comando_Abrir
'    (I)long posicao: numero de linhas a avancar ou recuar no resultado.
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO, AD_SQL_SEM_DADOS
#If JONES_280115 = 1 Then
    Call Com_Testa(lComando)
#End If
    Comando_BuscarRelativo = Comando_BuscarRelativoInt(lComando, lPosicao)

End Function

Function Comando_LockShared(ByVal lComando As Long) As Long
'    PARA LEITURAS COMPARTLHADAS
'    Forca o lock do registro correspondente a linha corrente de tal forma que outros locks shared possam ser obtidos,
'    mas nao um lock exclusive. Dessa forma uma "linha" pode ser compartilhada para leitura ficando impedida a
'    sua atualizacao.
'
'Parametros:
'    (I)LPCOMANDO lpComando: retornado em AD_Comando_Abrir (referente ao SELECT)
'
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO
'
'Observacao:
'    Mais de uma linha do resultado pode estar lockada, ou seja a cada novo lock o anterior NAO é desfeito automaticamente.
'    No commit e no fechamento do comando os locks sao liberados.
'    Um statement sera afetado por um commit se tiver sido "aberto" qdo a transacao estava ativa.
'    Se for executar uma leitura que nao necessite impedir atualizacoes nao faca lock.
#If JONES_280115 = 1 Then
    Call Com_Testa(lComando)
#End If
    Comando_LockShared = Comando_LockSharedInt(lComando)
    
End Function

Function Comando_LockExclusive(ByVal lComando As Long) As Long
'    PARA ATUALIZACOES
'    Forca o lock do registro correspondente a linha corrente de tal forma que nenhum outro lock no registro correspondente possa ser obtido.
'    E a forma de lock a ser utilizada na alteracao e delecao de registros.
'
'Parametros:
'    (I)LPCOMANDO lpComando: retornado em AD_Comando_Abrir (referente ao SELECT)
'
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO
'
'Observacao:
'    Mais de uma linha do resultado pode estar lockada, ou seja a cada novo lock o anterior NAO é desfeito automaticamente.
'    No commit e no fechamento do comando os locks sao liberados.
'    Um statement sera afetado por um commit se tiver sido "aberto" qdo a transacao estava ativa.
#If JONES_280115 = 1 Then
    Call Com_Testa(lComando)
#End If
    Comando_LockExclusive = Comando_LockExclusiveInt(lComando)
    
End Function

Function Comando_Unlock(ByVal lComando As Long) As Long
'    Liberacao explicita do lock da linha corrente. Serve p/locks Shared e Exclusive.
'
'Parametros:
'    (I)LPCOMANDO lpComando: retornado em AD_Comando_Abrir (referente ao SELECT)
'
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO
'
'Observacao:
'    Mais de uma linha do resultado pode estar lockada, ou seja a cada novo lock o anterior NAO é desfeito automaticamente.
'    No commit e no fechamento do comando os locks sao liberados.
#If JONES_280115 = 1 Then
    Call Com_Testa(lComando)
#End If
    Comando_Unlock = Comando_UnlockInt(lComando)
    
End Function

Function Comando_ReExecutar(ByVal lComando As Long) As Long
'reexecuta reaproveitando os buffers e parametros passados anteriormente numa chamada anterior a Comando_Executar e suas variacoes.
#If JONES_280115 = 1 Then
    Call Com_Testa(lComando)
#End If
    Comando_ReExecutar = Comando_ReExecutarInt(lComando)
    
End Function

Function Comando_AbrirExt(lConexao As Long) As Long
'**********************************************************************
'**********************************************************************
'    Inicializa "comando"
'Retorno:
'    NULL em caso de erro.
'**********************************************************************
'**********************************************************************
    Comando_AbrirExt = Comando_AbrirInt(lConexao)

#If JONES_280115 = 1 Then
    If Comando_AbrirExt <> 0 Then Call Com_Inclui(Comando_AbrirExt)
#End If

End Function

Function Comando_Abrir() As Long
'**********************************************************************
'**********************************************************************
'    Inicializa "comando"
'Retorno:
'    NULL em caso de erro.
'**********************************************************************
'**********************************************************************
    Comando_Abrir = Comando_AbrirInt(SistGlob_lConexao)
    
#If JONES_280115 = 1 Then
    If Comando_Abrir <> 0 Then Call Com_Inclui(Comando_Abrir)
#End If

End Function

Function Conexao_Abrir(ByVal driver_sql As Integer, ByVal lpParamIn As String, ByVal ParamLenIn As Integer, ByVal lpParamOut As String, lpParamLenOut As Integer) As Long
'**********************************************************************
'**********************************************************************
'    Estabelece conexao com bd.
'
'Parametros:
'
'(I)driver_sql: AD_SQL_DRIVER_*. Ver os defines acima.
'(I)lpParamIn: informacoes adicionais para conexao que variam de bd para bd.
'    Ex: p/driver ODBC podemos passar o string de conexao: "DSN=XYZ Corp;UID=Gomez;PWD=Sesame;UIDDBMS=JGomez;PWDDBMS=Shazam;"
'(I)ParamLenIn: tamanho das informacoes referenciadas por lpParamIn.
'(O)lpParamOut: informacoes de conexao na saida. Pode ser um string mais completo que o informado na entrada. Este string mais completo poderia ficar armazenado no cadastro de usuarios criptografado para evitar que nas conexoes seguintes tenha que preencher dialogos p/logon).
'    Pode ser NULL se nao interessar. Neste caso lpParamLenOut nao sera alterado.
'    O tamanho do buffer deve ser >= 255.
'(O)lpParamLenOut: tamanho das informacoes referenciadas por lpParamOut.
'
'Retorno:
'    ponteiro p/conexao (long no VB) ou NULL p/erro.
'Observacao:
'        So trabalharemos com READ_COMMITTED e manual commit.
'        So havera uma transacao por conexao.
'
'Exemplo:
'    Dim ConStrOut As String * 500
'    Dim len_out As Integer
'    Dim lConexao As Long
'
'    lConexao = Conexao_Abrir(AD_SQL_DRIVER_ODBC, "DSN=sdk21-Access", Len("DSN=sdk21-Access") + 1, ByVal ConStrOut, len_out)
'**********************************************************************
'**********************************************************************
    Conexao_Abrir = Conexao_AbrirInt(driver_sql, lpParamIn, ParamLenIn, lpParamOut, lpParamLenOut)
End Function

Function Conexao_Fechar(lConexao As Long) As Long
'**********************************************************************
'**********************************************************************
'    Encerra conexao com bd
'Retorno:
'    AD_BOOL_TRUE p / SUCESSO
'**********************************************************************
'**********************************************************************
    Conexao_Fechar = Conexao_FecharInt(lConexao)
    
End Function

Function Transacao_AbrirExt(lConexao As Long) As Long
'**********************************************************************
'**********************************************************************
'    iniciar uma transacao
'Retorno:
'    NULL p/erro
'**********************************************************************
'**********************************************************************
Dim lTransacaoNova As Long

    lTransacaoNova = Transacao_AbrirInt(lConexao)
    
    If lTransacaoNova <> 0 And lConexao = SistGlob_lConexaoDic Then
        SistGlob_lTransacaoDic = lTransacaoNova
    End If
    
    Transacao_AbrirExt = lTransacaoNova

End Function

Function Transacao_AbrirDic() As Long
'**********************************************************************
'**********************************************************************
'    iniciar uma transacao
'Retorno:
'    NULL p/erro
'**********************************************************************
'**********************************************************************
Dim lTransacaoNova As Long

    lTransacaoNova = Transacao_AbrirInt(SistGlob_lConexaoDic)
    
    If lTransacaoNova <> 0 Then
        SistGlob_lTransacaoDic = lTransacaoNova
    End If
    
    Transacao_AbrirDic = lTransacaoNova

End Function

Function Transacao_Abrir() As Long
'**********************************************************************
'**********************************************************************
'    iniciar uma transacao
'Retorno:
'    NULL p/erro
'**********************************************************************
'**********************************************************************
Dim lTransacaoNova As Long

    lTransacaoNova = Transacao_AbrirInt(SistGlob_lConexao)
    If lTransacaoNova <> 0 Then
        SistGlob_lTransacao = lTransacaoNova
    End If
    Transacao_Abrir = lTransacaoNova
    
End Function

Function Transacao_CommitExt(lTransacao As Long) As Long
'**********************************************************************
'**********************************************************************
'    Confirma todas as alteracoes no bd que faziam parte da transacao
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO
'**********************************************************************
'**********************************************************************

    Transacao_CommitExt = Transacao_CommitInt(lTransacao)
    If lTransacao = SistGlob_lTransacaoDic Then SistGlob_lTransacaoDic = 0
    
End Function

Function Transacao_CommitDic() As Long
'**********************************************************************
'**********************************************************************
'    Confirma todas as alteracoes no bd que faziam parte da transacao
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO
'**********************************************************************
'**********************************************************************

    Transacao_CommitDic = Transacao_CommitInt(SistGlob_lTransacaoDic)
    SistGlob_lTransacaoDic = 0
    
End Function

Function Transacao_Commit() As Long
'**********************************************************************
'**********************************************************************
'    Confirma todas as alteracoes no bd que faziam parte da transacao
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO
'**********************************************************************
'**********************************************************************

    Transacao_Commit = Transacao_CommitInt(SistGlob_lTransacao)
    SistGlob_lTransacao = 0
    
End Function

Function Transacao_RollbackExt(lTransacao As Long) As Long
'**********************************************************************
'**********************************************************************
'    Desfaz todas as alteracoes no bd que faziam parte da transacao
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO
'**********************************************************************
'**********************************************************************
    Transacao_RollbackExt = Transacao_RollbackInt(lTransacao)
    If lTransacao = SistGlob_lTransacaoDic Then SistGlob_lTransacaoDic = 0

End Function

Function Transacao_RollbackDic() As Long
'**********************************************************************
'**********************************************************************
'    Desfaz todas as alteracoes no bd que faziam parte da transacao
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO
'**********************************************************************
'**********************************************************************
    Transacao_RollbackDic = Transacao_RollbackInt(SistGlob_lTransacaoDic)
    SistGlob_lTransacaoDic = 0

End Function

Function Transacao_Rollback() As Long
'**********************************************************************
'**********************************************************************
'    Desfaz todas as alteracoes no bd que faziam parte da transacao
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO
'**********************************************************************
'**********************************************************************
    Transacao_Rollback = Transacao_RollbackInt(SistGlob_lTransacao)
    SistGlob_lTransacao = 0
    
End Function

Function Comando_Executar(ByVal lComando As Long, ByVal sComando_SQL As String, ParamArray avParam()) As Long
'**********************************************************************
'**********************************************************************
'    Forca a execucao de um comando SQL
'    Os parametros deverao estar representados dentro do comando SQL por "?" e as variaveis associadas devem ser VARIANT
'    dos tipos Integer(2), Long(3), Double(5) e String(8)
'Parametros:
'    (I)lComando: retornado em Comando_Abrir
'    (I)sComandoSQL: comando SQL. Ex.: "SELECT CADASTRO.NOME FROM CADASTRO".
'        Se o comando for igual ao anteriormente passado a "execucao" do comando será imediata, sem passar pela fase de "preparacao".
'    Apos os parametros acima devem ser passados os buffers que guardarao os resultados seguidos dos buffers que conterao os parametros, todos do tipo VARIANT.
'
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO
'    Em caso de erro *lplpConsulta sera NULL
'
'Exemplo:
'status = Comando_Executar( lpComando, "SELECT NOME, ENDERECO FROM CADASTRO WHERE ESTADO=?", nome, end, param_est )
'**********************************************************************
'**********************************************************************
Dim lErro As Long, iIndice As Integer, iIndMin As Integer, iIndMax As Integer
Dim X As New ADM
On Error GoTo saida_com_erro

#If JONES_280115 = 1 Then
    Call Com_Testa(lComando, sComando_SQL)
#End If

    lErro = Comando_PrepararInt(lComando, sComando_SQL)
    If (lErro <> AD_SQL_SUCESSO) Then Error 27620

    iIndMin = LBound(avParam)
    iIndMax = UBound(avParam)
    
    'obs.: vai funcionar mesmo que nao seja passado nenhum parametro pois neste caso UBound é -1
    For iIndice = iIndMin To iIndMax
    
        lErro = Comando_BindVarInt(lComando, avParam(iIndice))
        If (lErro <> AD_SQL_SUCESSO) Then Error 27621
    
    Next
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then Error 27622
    
    Comando_Executar = AD_SQL_SUCESSO
    
    Exit Function

saida_com_erro:
    
    Comando_Executar = AD_SQL_ERRO
    
    Select Case Err

        Case 27620, 27621
            lErro = X.Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", Err, sComando_SQL)

        Case 27622
            'lErro = X.Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", Err, sComando_SQL)
        
        Case Else
            lErro = X.Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142631)

    End Select

    Exit Function
    
End Function

Function Comando_Executar2(ByVal lComando As Long, ByVal sComando_SQL As String, vCampoSelect() As Variant, vValorFiltro() As Variant) As Long
'**********************************************************************
'**********************************************************************
'    Forca a execucao de um comando SQL
'    Os parametros estão representados dentro do comando SQL por "?" e as variaveis associadas são
'    os elementos da ARRAY de Variants vValorFiltro
'Parametros:
'    (I)lComando: lErroornado em Comando_Abrir
'    (I)sComandoSQL: comando SQL. Ex.: "SELECT CADASTRO.NOME FROM CADASTRO".
'        Se o comando for igual ao anteriormente passado a "execucao" do comando será imediata, sem passar pela fase de "preparacao".
'    Apos os parametros acima devem ser passados os buffers que guardarao os resultados seguidos dos buffers que conterao os parametros, todos do tipo VARIANT.
'    Os buffers são os elementos da ARRAY vCampoSelect

'Retorno:
'    lErro
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO
'    Em caso de erro *lplpConsulta sera NULL
'
'Exemplo:
'status = Comando_Executar2( lpComando, "SELECT NOME, ENDERECO FROM CADASTRO WHERE ESTADO=?", vCampoSelect(), vValorFiltro())
'**********************************************************************
'**********************************************************************
Dim iIndice As Integer
Dim lErro As Long
Dim X As New ADM

On Error GoTo Erro_Comando_Executar2

#If JONES_280115 = 1 Then
    Call Com_Testa(lComando, sComando_SQL)
#End If

    lErro = Comando_PrepararInt(lComando, sComando_SQL)
    If (lErro <> AD_SQL_SUCESSO) Then Error 6609
    
    'Faz bind para as variáveis que vão receber os campos do SELECT
    For iIndice = 1 To UBound(vCampoSelect)
        
        lErro = Comando_BindVarInt(lComando, vCampoSelect(iIndice))
        If (lErro <> AD_SQL_SUCESSO) Then Error 6613
    
    Next
    
    If UBound(vValorFiltro) >= 1 Then
    
        'Faz bind para os valores dos filtros na cláusula WHERE
        For iIndice = 1 To UBound(vValorFiltro)
    
            lErro = Comando_BindVarInt(lComando, vValorFiltro(iIndice))
            If (lErro <> AD_SQL_SUCESSO) Then Error 6743
    
        Next
    
    End If
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then Error 6615
    
    Comando_Executar2 = AD_SQL_SUCESSO
    Exit Function

Erro_Comando_Executar2:
    
    Comando_Executar2 = Err

    Select Case Err

        Case 6609
            lErro = X.Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", Err, sComando_SQL)

        Case 6613, 6743
            lErro = X.Rotina_Erro(vbOKOnly, "ERRO_BIND_CAMPO2", Err)

        Case 6615
            'lErro = X.Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", Err, sComando_SQL)
        
        Case Else
            lErro = X.Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142632)

    End Select

    Exit Function

End Function

Function Comando_ExecutarPos(ByVal lComando As Long, ByVal sComando_SQL As String, ByVal lSelect As Long, ParamArray avParam()) As Long
'**********************************************************************
'**********************************************************************
'    Para a execucao de um comando UPDATE, DELETE posicionado (WHERE CURRENT OF)
'    e do SELECT correspondente.
'    Todos os valores devem ser passados por parametro, ex.: "UPDATE Orders SET Quantity=?"
'Parametros:
'    (I)LPCOMANDO lpComando: retornado em AD_Comando_Abrir
'    (I)LPSTR lpComandoSQL: comando SQL. Nao deve conter a clausula WHERE CURRENT OF que sera adicionada pela funcao.
'        Se o comando for igual ao anteriormente passado a "execucao" do comando será imediata, sem passar pela fase de "preparacao".
'    (I)lpSelect: ponteiro retornado em AD_Comando_Abrir no SELECT a ser usado p/posicionar.
'       Qdo estiver executando o proprio SELECT este parametro sera ignorado.
'    Apos os parametros acima devem ser passados os buffers que guardarao os parametros.
'
'Retorno:
'    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO
'
'Exemplo:
'status = AD_Comando_ExecutarPos(lpComando, "UPDATE CADASTRO SET SALARIO=?N", 1, lpSelect, param_est$)
'**********************************************************************
'**********************************************************************
Dim lErro As Long, iIndice As Integer, iIndMin As Integer, iIndMax As Integer
Dim X As New ADM
On Error GoTo saida_com_erro

''    If Mid$(sComando_SQL, 1, 6) = "UPDATE" Or Mid$(sComando_SQL, 1, 6) = "DELETE" Then
''
''        Comando_ExecutarPos = AD_SQL_SUCESSO
''        Exit Function
''
''    End If
''

#If JONES_280115 = 1 Then
    Call Com_Testa(lComando, sComando_SQL)
    Call Com_Testa(lSelect)
#End If

    lErro = Comando_PrepararPosInt(lComando, sComando_SQL, lSelect)
    If (lErro <> AD_SQL_SUCESSO) Then Error 27623

    iIndMin = LBound(avParam)
    iIndMax = UBound(avParam)
    
    'obs.: vai funcionar mesmo que nao seja passado nenhum parametro pois neste caso UBound é -1
    For iIndice = iIndMin To iIndMax
    
        lErro = Comando_BindVarInt(lComando, avParam(iIndice))
        If (lErro <> AD_SQL_SUCESSO) Then Error 27624
    
    Next

    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then Error 27625
    
    Comando_ExecutarPos = AD_SQL_SUCESSO
    
    Exit Function

saida_com_erro:
    
    Comando_ExecutarPos = AD_SQL_ERRO
    
    Select Case Err

        Case 27623, 27624
            lErro = X.Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", Err, sComando_SQL)

        Case 27625
            'lErro = X.Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", Err, sComando_SQL)
        
        Case Else
            lErro = X.Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142633)

    End Select

    Exit Function
    
End Function

Function Comando_ExecutarLockado(ByVal lComando As Long, ByVal sComando_SQL As String, ParamArray avParam()) As Long
'**********************************************************************
'**********************************************************************
' Para a execucao de um comando SQL SELECT em cujas linhas resultado sera utilizado o lock.
' Deve referenciar apenas uma tabela e a qualificacao dos campos nao deve incluir o nome da tabela
'**********************************************************************
'**********************************************************************
Dim lErro As Long, iIndice As Integer, iIndMin As Integer, iIndMax As Integer
Dim X As New ADM
On Error GoTo saida_com_erro

#If JONES_280115 = 1 Then
    Call Com_Testa(lComando, sComando_SQL)
#End If

    lErro = Comando_PrepararLockadoInt(lComando, sComando_SQL)
    If (lErro <> AD_SQL_SUCESSO) Then Error 27626

    iIndMin = LBound(avParam)
    iIndMax = UBound(avParam)
    
    'obs.: vai funcionar mesmo que nao seja passado nenhum parametro pois neste caso UBound é -1
    For iIndice = iIndMin To iIndMax
    
        lErro = Comando_BindVarInt(lComando, avParam(iIndice))
        If (lErro <> AD_SQL_SUCESSO) Then Error 27627
    
    Next
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then Error 27628
    
    Comando_ExecutarLockado = AD_SQL_SUCESSO
    
    Exit Function

saida_com_erro:
    
    Comando_ExecutarLockado = AD_SQL_ERRO
    
    Select Case Err

        Case 27626, 27627
            lErro = X.Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", Err, sComando_SQL)

        Case 27628
            'lErro = X.Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", Err, sComando_SQL)
        
        Case Else
            lErro = X.Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142634)

    End Select

    Exit Function
    
End Function

''Function Comando_ExecutarParc(ByVal lComando As Long, ByVal sComando_SQL As String, ParamArray avParam()) As Long
'''**********************************************************************
'''**********************************************************************
'''    Forca a execucao de um comando SQL em que o result set será lido parcialmente, aos poucos
'''    Os parametros deverao estar representados dentro do comando SQL por "?" e as variaveis associadas devem ser VARIANT
'''    dos tipos Integer(2), Long(3), Double(5) e String(8)
'''Parametros:
'''    (I)lComando: retornado em Comando_Abrir
'''    (I)sComandoSQL: comando SQL. Ex.: "SELECT CADASTRO.NOME FROM CADASTRO".
'''        Se o comando for igual ao anteriormente passado a "execucao" do comando será imediata, sem passar pela fase de "preparacao".
'''    Apos os parametros acima devem ser passados os buffers que guardarao os resultados seguidos dos buffers que conterao os parametros, todos do tipo VARIANT.
'''
'''Retorno:
'''    AD_SQL_SUCESSO , AD_SQL_SUCESSO_PARCIAL, AD_SQL_ERRO
'''    Em caso de erro *lplpConsulta sera NULL
'''
'''Exemplo:
'''status = Comando_Executar( lpComando, "SELECT NOME, ENDERECO FROM CADASTRO WHERE ESTADO=?", nome, end, param_est )
'''**********************************************************************
'''**********************************************************************
''Dim lErro As Long, iIndice As Integer, iIndMin As Integer, iIndMax As Integer
''Dim X As New ADM
''On Error GoTo saida_com_erro
''
''    lErro = Comando_PrepararParcInt(lComando, sComando_SQL)
''    If (lErro <> AD_SQL_SUCESSO) Then Error 59285
''
''    iIndMin = LBound(avParam)
''    iIndMax = UBound(avParam)
''
''    'obs.: vai funcionar mesmo que nao seja passado nenhum parametro pois neste caso UBound é -1
''    For iIndice = iIndMin To iIndMax
''
''        lErro = Comando_BindVarInt(lComando, avParam(iIndice))
''        If (lErro <> AD_SQL_SUCESSO) Then Error 59286
''
''    Next
''
''    lErro = Comando_ExecutarInt(lComando)
''    If (lErro <> AD_SQL_SUCESSO) Then Error 59287
''
''    Comando_ExecutarParc = AD_SQL_SUCESSO
''
''    Exit Function
''
''saida_com_erro:
''
''    Comando_ExecutarParc = AD_SQL_ERRO
''
''    Select Case Err
''
''        Case 59285, 59286
''            lErro = X.Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", Err, sComando_SQL)
''
''        Case 59287
''            lErro = X.Rotina_Erro(vbOKOnly, "ERRO_EXECUCAO_COMANDO_SQL", Err, sComando_SQL)
''
''        Case Else
''            lErro = X.Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142635)
''
''    End Select
''
''    Exit Function
''
''End Function

Function Comando_DefScroll(ByVal lComando As Long, bComScroll As Boolean) As Long
'    Se TRUE força o tipo do cursor p/permitir extendedfetch p/frente/tras/primeiro/ultimo
'    Se FALSE pode deixar ficar readonly forwardonly
'
'Parametros:
'    (I)LPCOMANDO lpComando: retornado em AD_Comando_Abrir (referente ao SELECT)
'    (I)bComScroll
'Retorno:
'    AD_SQL_SUCESSO
#If JONES_280115 = 1 Then
    Call Com_Testa(lComando)
#End If
    Comando_DefScroll = Comando_DefScrollInt(lComando, IIf(bComScroll, 1, 0))
    
End Function
    
Function Conexao_DesativarLocks(lDesativar As Long) As Long
'se lDesativar <> 0 desativa os locks internos do sistema, se zero entao reativa-os
Dim X As New ADM

    Call Conexao_DesativarLocksInt(X.GL_lConexao, lDesativar)
    
    Conexao_DesativarLocks = SUCESSO
    
End Function

Function Conexao_DesativarLocksExt(lConexao As Long, lDesativar As Long) As Long
'se lDesativar <> 0 desativa os locks internos do sistema, se zero entao reativa-os

    Call Conexao_DesativarLocksInt(lConexao, lDesativar)
    
    Conexao_DesativarLocksExt = SUCESSO
    
End Function

